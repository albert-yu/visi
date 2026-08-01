//! Hand-rolled OOXML for native Excel PivotTables. Neither `calamine`
//! (import) nor `rust_xlsxwriter` (export) has any pivot table support, so
//! this module reads/writes the `xl/pivotCache/*` and `xl/pivotTables/*`
//! parts directly by string-templating XML on export and hand-parsing it
//! with `quick_xml` on import -- the same approach `xlsx.rs` already uses
//! for Excel Tables' row-flag XML and for chart import.
//!
//! Export post-processes the zip `rust_xlsxwriter` already produced (it has
//! no hook for injecting arbitrary extra parts): the destination cells are
//! always written as plain computed values first (so any reader sees
//! correct numbers), then this module adds the cache/table parts and
//! wires up the required relationships/content-types so Excel recognizes
//! and can refresh the range as a real PivotTable.
//!
//! Import only reconstructs each `PivotTable`'s source, destination, and
//! row/column/value field assignments (subtotal toggles default back on,
//! and filter selections reset to "all") -- enough for `visi pivot`
//! CLI commands to keep editing a pivot table across process invocations,
//! which is the primary reason this needs to round-trip at all.

use crate::core::engine::{CellRef, ResultData, Sheet};
use crate::core::parser::col_idx_to_letters;
use crate::core::pivot::{
    self, PivotAggregation, PivotAxisItem, PivotField, PivotFilterField, PivotSource, PivotTable,
    PivotValueField, compute_pivot, sorted_distinct_strings,
};
use crate::core::xlsx::{
    get_attr, get_zip_file_content, parse_workbook_rels, parse_workbook_sheets,
};
use std::collections::HashMap;
use std::io::{Read, Write};

const NS_MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL_PIVOT_CACHE_DEF: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";
const REL_PIVOT_CACHE_RECORDS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords";
const REL_PIVOT_TABLE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";

fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

fn a1_cell(row: usize, col: usize) -> String {
    format!("{}{}", col_idx_to_letters(col), row + 1)
}

fn a1_range(r0: usize, c0: usize, r1: usize, c1: usize) -> String {
    format!("{}:{}", a1_cell(r0, c0), a1_cell(r1, c1))
}

struct PivotXmlUnit {
    cache_definition_xml: String,
    cache_records_xml: String,
    pivot_table_xml: String,
    dest_sheet_id: u64,
}

fn is_all_numeric(records: &[Vec<ResultData>], field_idx: usize) -> bool {
    !records.is_empty()
        && records.iter().all(|r| {
            matches!(
                r.get(field_idx),
                Some(ResultData::Integer(_)) | Some(ResultData::Float(_)) | Some(ResultData::None)
            )
        })
}

/// Builds the `<items>` child element for a row/column pivotField: one
/// `<item x="i"/>` per sorted distinct value, plus a trailing
/// `<item t="default"/>` subtotal placeholder when `with_default` is set.
fn build_items_xml(item_count: usize, with_default: bool) -> String {
    let total = item_count + if with_default { 1 } else { 0 };
    let mut s = format!("<items count=\"{}\">", total);
    for idx in 0..item_count {
        s.push_str(&format!("<item x=\"{}\"/>", idx));
    }
    if with_default {
        s.push_str("<item t=\"default\"/>");
    }
    s.push_str("</items>");
    s
}

/// Builds the `<items>` child for a page/filter pivotField: every distinct
/// value, marked `h="1"` (hidden) when it isn't in `selected`.
fn build_filter_items_xml(values: &[String], selected: &Option<Vec<String>>) -> String {
    let mut s = format!("<items count=\"{}\">", values.len());
    for (idx, v) in values.iter().enumerate() {
        let hidden = selected
            .as_ref()
            .is_some_and(|sel| !sel.iter().any(|s| s == v));
        if hidden {
            s.push_str(&format!("<item x=\"{}\" h=\"1\"/>", idx));
        } else {
            s.push_str(&format!("<item x=\"{}\"/>", idx));
        }
    }
    s.push_str("</items>");
    s
}

/// Encodes one axis (row or column) of flattened groups into the
/// `rowItems`/`colItems` `<i>` sequence, applying Excel's leading-field
/// repeat suppression (`r="N"` = "the first N fields are unchanged from
/// the previous item, so aren't repeated here").
fn build_axis_items_xml(
    axis: &[PivotAxisItem],
    field_idxs: &[usize],
    field_items: &HashMap<usize, Vec<String>>,
) -> String {
    let mut out = String::new();
    let mut prev_labels: Vec<Option<String>> = vec![None; field_idxs.len()];
    for item in axis {
        if item.is_grand_total {
            out.push_str("<i t=\"grand\"><x/></i>");
            for l in prev_labels.iter_mut() {
                *l = None;
            }
            continue;
        }
        if field_idxs.is_empty() {
            out.push_str("<i/>");
            continue;
        }
        let last_shown_depth = item
            .labels
            .iter()
            .rposition(|l| l.is_some())
            .unwrap_or(0)
            .min(field_idxs.len() - 1);
        let mut first_diff = last_shown_depth;
        for d in 0..=last_shown_depth {
            let cur = item.labels.get(d).cloned().flatten();
            if cur != prev_labels[d] {
                first_diff = d;
                break;
            }
        }
        let mut x_entries = String::new();
        for d in first_diff..=last_shown_depth {
            let cur = item.labels.get(d).cloned().flatten().unwrap_or_default();
            let items = field_items
                .get(&field_idxs[d])
                .map(|v| v.as_slice())
                .unwrap_or(&[]);
            let x_val = items.iter().position(|v| v == &cur).unwrap_or(0);
            if x_val == 0 {
                x_entries.push_str("<x/>");
            } else {
                x_entries.push_str(&format!("<x v=\"{}\"/>", x_val));
            }
        }
        for (d, cur) in item.labels.iter().enumerate().take(last_shown_depth + 1) {
            prev_labels[d] = cur.clone();
        }
        for l in prev_labels.iter_mut().skip(last_shown_depth + 1) {
            *l = None;
        }
        let type_attr = if item.is_subtotal {
            " t=\"default\""
        } else {
            ""
        };
        if first_diff == 0 {
            out.push_str(&format!("<i{}>{}</i>", type_attr, x_entries));
        } else {
            out.push_str(&format!(
                "<i{} r=\"{}\">{}</i>",
                type_attr, first_diff, x_entries
            ));
        }
    }
    out
}

/// When there are 2+ value fields, Excel places an implicit "Values"
/// pseudo-field as the innermost column level (field index `-2`); this
/// duplicates each column-axis group once per value field so
/// `build_axis_items_xml` can encode that extra level like any other.
fn expand_col_axis_with_values(
    axis: &[PivotAxisItem],
    value_labels: &[String],
) -> Vec<PivotAxisItem> {
    let mut out = Vec::new();
    for item in axis {
        if item.is_grand_total {
            for _ in value_labels {
                out.push(PivotAxisItem {
                    labels: item.labels.clone(),
                    is_subtotal: false,
                    is_grand_total: true,
                });
            }
            continue;
        }
        for vf_label in value_labels {
            let mut labels = item.labels.clone();
            labels.push(Some(vf_label.clone()));
            out.push(PivotAxisItem {
                labels,
                is_subtotal: false,
                is_grand_total: false,
            });
        }
    }
    out
}

const VALUES_SENTINEL: usize = usize::MAX;

fn build_col_items_xml(
    col_axis: &[PivotAxisItem],
    col_field_idxs: &[usize],
    field_items: &HashMap<usize, Vec<String>>,
    value_labels: &[String],
) -> String {
    if value_labels.len() > 1 {
        let expanded_axis = expand_col_axis_with_values(col_axis, value_labels);
        let mut expanded_idxs = col_field_idxs.to_vec();
        expanded_idxs.push(VALUES_SENTINEL);
        let mut expanded_items = field_items.clone();
        expanded_items.insert(VALUES_SENTINEL, value_labels.to_vec());
        build_axis_items_xml(&expanded_axis, &expanded_idxs, &expanded_items)
    } else {
        build_axis_items_xml(col_axis, col_field_idxs, field_items)
    }
}

fn subtotal_token(agg: PivotAggregation) -> &'static str {
    match agg {
        PivotAggregation::Sum => "sum",
        PivotAggregation::Count => "count",
        PivotAggregation::CountNumbers => "countNums",
        PivotAggregation::Average => "average",
        PivotAggregation::Max => "max",
        PivotAggregation::Min => "min",
    }
}

fn build_pivot_xml_unit(
    sheets: &[Sheet],
    pivot: &PivotTable,
    cache_id: usize,
) -> Result<PivotXmlUnit, String> {
    let (src_sheet, col_names, sheet_cols, data_rows) =
        pivot::resolve_source(sheets, &pivot.source)?;

    let mut records: Vec<Vec<ResultData>> = Vec::with_capacity(data_rows.len());
    for &r in &data_rows {
        let mut row = Vec::with_capacity(sheet_cols.len());
        for &c in &sheet_cols {
            row.push(src_sheet.get_result_data(&CellRef::new(r, c)));
        }
        records.push(row);
    }

    let mut cache_fields_xml = String::new();
    for (i, name) in col_names.iter().enumerate() {
        let shared_items_attrs = if is_all_numeric(&records, i) {
            " containsSemiMixedTypes=\"0\" containsString=\"0\" containsNumber=\"1\""
        } else {
            ""
        };
        cache_fields_xml.push_str(&format!(
            "<cacheField name=\"{}\" numFmtId=\"0\"><sharedItems{}/></cacheField>",
            escape_xml(name),
            shared_items_attrs
        ));
    }

    let mut cache_records_xml = String::new();
    for row in &records {
        cache_records_xml.push_str("<r>");
        for val in row {
            match val {
                ResultData::Integer(n) => cache_records_xml.push_str(&format!("<n v=\"{}\"/>", n)),
                ResultData::Float(f) => cache_records_xml.push_str(&format!("<n v=\"{}\"/>", f)),
                ResultData::Boolean(b) => {
                    cache_records_xml.push_str(&format!("<b v=\"{}\"/>", if *b { 1 } else { 0 }))
                }
                ResultData::None => cache_records_xml.push_str("<m/>"),
                other => cache_records_xml
                    .push_str(&format!("<s v=\"{}\"/>", escape_xml(&other.to_string()))),
            }
        }
        cache_records_xml.push_str("</r>");
    }

    let (src_ref, src_sheet_name) = match &pivot.source {
        PivotSource::Table { name } => {
            let (sheet, table) = sheets
                .iter()
                .find_map(|s| s.find_table(name).map(|t| (s, t)))
                .ok_or_else(|| format!("Table '{}' not found", name))?;
            (
                a1_range(
                    table.start_row,
                    table.start_col,
                    table.end_row,
                    table.end_col,
                ),
                sheet.name.clone(),
            )
        }
        PivotSource::Range {
            start_row,
            start_col,
            end_row,
            end_col,
            ..
        } => (
            a1_range(*start_row, *start_col, *end_row, *end_col),
            src_sheet.name.clone(),
        ),
    };

    let cache_definition_xml = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
<pivotCacheDefinition xmlns=\"{ns}\" xmlns:r=\"{ns_r}\" r:id=\"rId1\" refreshOnLoad=\"1\" recordCount=\"{count}\">\
<cacheSource type=\"worksheet\"><worksheetSource ref=\"{src_ref}\" sheet=\"{src_sheet}\"/></cacheSource>\
<cacheFields count=\"{field_count}\">{cache_fields_xml}</cacheFields>\
</pivotCacheDefinition>",
        ns = NS_MAIN,
        ns_r = NS_R,
        count = records.len(),
        src_ref = escape_xml(&src_ref),
        src_sheet = escape_xml(&src_sheet_name),
        field_count = col_names.len(),
        cache_fields_xml = cache_fields_xml,
    );
    let cache_records_xml_doc = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
<pivotCacheRecords xmlns=\"{ns}\" count=\"{count}\">{records}</pivotCacheRecords>",
        ns = NS_MAIN,
        count = records.len(),
        records = cache_records_xml,
    );

    let row_field_idxs: Vec<usize> = pivot
        .row_fields
        .iter()
        .map(|f| pivot::column_index(&col_names, &f.column))
        .collect::<Result<_, _>>()?;
    let col_field_idxs: Vec<usize> = pivot
        .col_fields
        .iter()
        .map(|f| pivot::column_index(&col_names, &f.column))
        .collect::<Result<_, _>>()?;
    let page_field_idxs: Vec<usize> = pivot
        .filter_fields
        .iter()
        .map(|f| pivot::column_index(&col_names, &f.column))
        .collect::<Result<_, _>>()?;
    let data_field_idxs: Vec<usize> = pivot
        .value_fields
        .iter()
        .map(|f| pivot::column_index(&col_names, &f.column))
        .collect::<Result<_, _>>()?;

    let field_items: HashMap<usize, Vec<String>> = row_field_idxs
        .iter()
        .chain(col_field_idxs.iter())
        .chain(page_field_idxs.iter())
        .map(|&idx| {
            let vals: Vec<String> = records.iter().map(|r| pivot::group_key(&r[idx])).collect();
            (idx, sorted_distinct_strings(&vals))
        })
        .collect();

    let value_multiplier = if pivot.value_fields.len() > 1 {
        pivot.value_fields.len()
    } else {
        1
    };

    // pivotFields: one per source column, in order.
    let mut pivot_fields_xml = String::new();
    for i in 0..col_names.len() {
        let mut attrs = String::new();
        let mut items_xml: Option<String> = None;

        if let Some(pos) = row_field_idxs.iter().position(|&x| x == i) {
            attrs.push_str(" axis=\"axisRow\"");
            let subtotal_enabled = pivot.row_fields[pos].subtotal && pos + 1 < row_field_idxs.len();
            let n_items = field_items.get(&i).map(|v| v.len()).unwrap_or(0);
            items_xml = Some(build_items_xml(n_items, subtotal_enabled));
        } else if let Some(pos) = col_field_idxs.iter().position(|&x| x == i) {
            attrs.push_str(" axis=\"axisCol\"");
            let subtotal_enabled = pivot.col_fields[pos].subtotal && pos + 1 < col_field_idxs.len();
            let n_items = field_items.get(&i).map(|v| v.len()).unwrap_or(0);
            items_xml = Some(build_items_xml(n_items, subtotal_enabled));
        } else if let Some(pos) = page_field_idxs.iter().position(|&x| x == i) {
            attrs.push_str(" axis=\"axisPage\" multipleItemSelectionAllowed=\"1\"");
            let values = field_items.get(&i).cloned().unwrap_or_default();
            let selected = &pivot.filter_fields[pos].selected_values;
            items_xml = Some(build_filter_items_xml(&values, selected));
        }
        if data_field_idxs.contains(&i) {
            attrs.push_str(" dataField=\"1\"");
        }
        attrs.push_str(" showAll=\"0\"");

        match items_xml {
            Some(xml) => {
                pivot_fields_xml.push_str(&format!("<pivotField{}>{}</pivotField>", attrs, xml))
            }
            None => pivot_fields_xml.push_str(&format!("<pivotField{}/>", attrs)),
        }
    }

    let mut row_fields_xml = String::new();
    for &idx in &row_field_idxs {
        row_fields_xml.push_str(&format!("<field x=\"{}\"/>", idx));
    }
    let mut col_fields_xml = String::new();
    for &idx in &col_field_idxs {
        col_fields_xml.push_str(&format!("<field x=\"{}\"/>", idx));
    }
    if value_multiplier > 1 {
        col_fields_xml.push_str("<field x=\"-2\"/>");
    }
    let mut page_fields_xml = String::new();
    for &idx in &page_field_idxs {
        page_fields_xml.push_str(&format!("<pageField fld=\"{}\" hier=\"-1\"/>", idx));
    }

    let mut data_fields_xml = String::new();
    for (i, vf) in pivot.value_fields.iter().enumerate() {
        data_fields_xml.push_str(&format!(
            "<dataField name=\"{}\" fld=\"{}\" subtotal=\"{}\" baseField=\"0\" baseItem=\"0\"/>",
            escape_xml(&vf.label()),
            data_field_idxs[i],
            subtotal_token(vf.aggregation),
        ));
    }

    let (
        location_ref,
        first_header_row,
        first_data_row,
        first_data_col,
        row_items_xml,
        col_items_xml,
    ) = if pivot.value_fields.is_empty() {
        let loc = a1_range(
            pivot.dest_row,
            pivot.dest_col,
            pivot.dest_row,
            pivot.dest_col,
        );
        (
            loc,
            0usize,
            0usize,
            0usize,
            "<i/>".to_string(),
            "<i/>".to_string(),
        )
    } else {
        let grid = compute_pivot(sheets, pivot)?;
        // Excel's own `<location>` spans only the row/col header + data
        // grid, not the filter/page-field rows above it (those are
        // documented separately via `rowPageCount`/`colPageCount` below,
        // matching real Excel's convention, verified against an
        // Excel-produced pivotTable1.xml).
        let height = grid.header_rows.len() + grid.body_rows.len();
        let width = grid.width.max(1);
        let grid_start_row = pivot.dest_row + grid.grid_row_offset();
        let loc = a1_range(
            grid_start_row,
            pivot.dest_col,
            grid_start_row + height.saturating_sub(1),
            pivot.dest_col + width.saturating_sub(1),
        );
        let row_label_width = pivot.row_fields.len().max(1);
        let n_header_rows = grid.header_rows.len();
        let row_items = build_axis_items_xml(&grid.row_axis, &row_field_idxs, &field_items);
        let value_labels: Vec<String> = pivot.value_fields.iter().map(|v| v.label()).collect();
        let col_items =
            build_col_items_xml(&grid.col_axis, &col_field_idxs, &field_items, &value_labels);
        (
            loc,
            n_header_rows.saturating_sub(1),
            n_header_rows,
            row_label_width,
            row_items,
            col_items,
        )
    };

    let row_items_count = row_items_xml.matches("<i").count();
    let col_items_count = col_items_xml.matches("<i").count();

    // Matches real Excel: `rowPageCount`/`colPageCount` on `<location>`
    // document how many rows/columns above/left of `ref` are reserved for
    // filter/page fields, only present at all when there are any.
    let page_count_attrs = if pivot.filter_fields.is_empty() {
        String::new()
    } else {
        format!(
            " rowPageCount=\"{}\" colPageCount=\"1\"",
            pivot.filter_fields.len()
        )
    };

    #[allow(clippy::format_in_format_args)]
    let pivot_table_xml = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
<pivotTableDefinition xmlns=\"{ns}\" name=\"{name}\" cacheId=\"{cache_id}\" applyNumberFormats=\"0\" \
applyBorderFormats=\"0\" applyFontFormats=\"0\" applyPatternFormats=\"0\" applyAlignmentFormats=\"0\" \
applyWidthHeightFormats=\"1\" dataCaption=\"Values\" updatedVersion=\"6\" minRefreshableVersion=\"3\" \
useAutoFormatting=\"1\" itemPrintTitles=\"1\" createdVersion=\"6\" indent=\"0\" outline=\"1\" \
outlineData=\"1\" multipleFieldFilters=\"0\" rowGrandTotals=\"{row_grand}\" colGrandTotals=\"{col_grand}\">\
<location ref=\"{location_ref}\" firstHeaderRow=\"{first_header_row}\" firstDataRow=\"{first_data_row}\" firstDataCol=\"{first_data_col}\"{page_count_attrs}/>\
<pivotFields count=\"{n_fields}\">{pivot_fields_xml}</pivotFields>\
{row_fields}{row_items}{col_fields}{col_items}{page_fields}{data_fields}\
<pivotTableStyleInfo name=\"PivotStyleLight16\" showRowHeaders=\"1\" showColHeaders=\"1\" showLastColumn=\"1\"/>\
</pivotTableDefinition>",
        ns = NS_MAIN,
        name = escape_xml(&pivot.name),
        cache_id = cache_id,
        row_grand = if pivot.grand_totals_row { "1" } else { "0" },
        col_grand = if pivot.grand_totals_col { "1" } else { "0" },
        location_ref = escape_xml(&location_ref),
        first_header_row = first_header_row,
        first_data_row = first_data_row,
        first_data_col = first_data_col,
        page_count_attrs = page_count_attrs,
        n_fields = col_names.len(),
        pivot_fields_xml = pivot_fields_xml,
        row_fields = if row_fields_xml.is_empty() {
            String::new()
        } else {
            format!(
                "<rowFields count=\"{}\">{}</rowFields>",
                row_field_idxs.len(),
                row_fields_xml
            )
        },
        row_items = format!(
            "<rowItems count=\"{}\">{}</rowItems>",
            row_items_count, row_items_xml
        ),
        col_fields = if col_fields_xml.is_empty() {
            String::new()
        } else {
            let n = col_field_idxs.len() + usize::from(value_multiplier > 1);
            format!("<colFields count=\"{}\">{}</colFields>", n, col_fields_xml)
        },
        col_items = format!(
            "<colItems count=\"{}\">{}</colItems>",
            col_items_count, col_items_xml
        ),
        page_fields = if page_fields_xml.is_empty() {
            String::new()
        } else {
            format!(
                "<pageFields count=\"{}\">{}</pageFields>",
                page_field_idxs.len(),
                page_fields_xml
            )
        },
        data_fields = if data_fields_xml.is_empty() {
            String::new()
        } else {
            format!(
                "<dataFields count=\"{}\">{}</dataFields>",
                pivot.value_fields.len(),
                data_fields_xml
            )
        },
    );

    Ok(PivotXmlUnit {
        cache_definition_xml,
        cache_records_xml: cache_records_xml_doc,
        pivot_table_xml,
        dest_sheet_id: pivot.dest_sheet_id,
    })
}

fn extract_max_rid(rels_xml: &str) -> usize {
    let mut max_id = 0;
    let mut reader = quick_xml::Reader::from_str(rels_xml);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Eof) => break,
            Ok(quick_xml::events::Event::Start(ref e))
            | Ok(quick_xml::events::Event::Empty(ref e))
                if e.local_name().as_ref() == b"Relationship" =>
            {
                if let Some(id) = get_attr(e, b"Id")
                    && let Some(num) = id.strip_prefix("rId").and_then(|s| s.parse::<usize>().ok())
                {
                    max_id = max_id.max(num);
                }
            }
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    max_id
}

/// Post-processes an already-exported xlsx buffer, adding native PivotTable
/// parts for every entry in `pivots` whose destination sheet still exists.
/// Returns the buffer unchanged if there's nothing to add.
pub fn inject_pivot_tables(
    xlsx_bytes: Vec<u8>,
    sheets: &[Sheet],
    pivots: &[PivotTable],
) -> Result<Vec<u8>, String> {
    let mut units: Vec<PivotXmlUnit> = Vec::new();
    for pivot in pivots {
        if sheets.iter().any(|s| s.id == pivot.dest_sheet_id) {
            units.push(build_pivot_xml_unit(sheets, pivot, units.len())?);
        }
    }
    if units.is_empty() {
        return Ok(xlsx_bytes);
    }

    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(&xlsx_bytes[..]))
        .map_err(|e| format!("Failed to open generated xlsx zip: {}", e))?;

    let workbook_xml =
        get_zip_file_content(&mut archive, "xl/workbook.xml").ok_or("Missing xl/workbook.xml")?;
    let workbook_rels_xml = get_zip_file_content(&mut archive, "xl/_rels/workbook.xml.rels")
        .ok_or("Missing xl/_rels/workbook.xml.rels")?;
    let mut content_types = get_zip_file_content(&mut archive, "[Content_Types].xml")
        .ok_or("Missing [Content_Types].xml")?;

    let rid_to_name = parse_workbook_sheets(&workbook_xml);
    let rid_to_target = parse_workbook_rels(&workbook_rels_xml);
    let mut sheet_name_to_filename: HashMap<String, String> = HashMap::new();
    for (rid, name) in &rid_to_name {
        if let Some(target) = rid_to_target.get(rid) {
            sheet_name_to_filename.insert(name.clone(), target.clone());
        }
    }

    let mut max_rid = extract_max_rid(&workbook_rels_xml);
    let mut new_workbook_rels_xml = workbook_rels_xml.clone();
    let mut pivot_caches_xml = String::new();
    let mut content_type_overrides = String::new();
    let mut new_files: Vec<(String, String)> = Vec::new();
    let mut sheet_rels_edits: HashMap<String, String> = HashMap::new();

    for (cache_id, unit) in units.iter().enumerate() {
        let n = cache_id + 1;
        max_rid += 1;
        let rid = format!("rId{}", max_rid);

        new_files.push((
            format!("xl/pivotCache/pivotCacheDefinition{}.xml", n),
            unit.cache_definition_xml.clone(),
        ));
        new_files.push((
            format!("xl/pivotCache/pivotCacheRecords{}.xml", n),
            unit.cache_records_xml.clone(),
        ));
        new_files.push((
            format!("xl/pivotTables/pivotTable{}.xml", n),
            unit.pivot_table_xml.clone(),
        ));
        new_files.push((
            format!("xl/pivotCache/_rels/pivotCacheDefinition{}.xml.rels", n),
            format!(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\
<Relationship Id=\"rId1\" Type=\"{}\" Target=\"pivotCacheRecords{}.xml\"/></Relationships>",
                REL_PIVOT_CACHE_RECORDS, n
            ),
        ));

        let rel_entry = format!(
            "<Relationship Id=\"{}\" Type=\"{}\" Target=\"pivotCache/pivotCacheDefinition{}.xml\"/>",
            rid, REL_PIVOT_CACHE_DEF, n
        );
        new_workbook_rels_xml = new_workbook_rels_xml.replacen(
            "</Relationships>",
            &format!("{}</Relationships>", rel_entry),
            1,
        );
        pivot_caches_xml.push_str(&format!(
            "<pivotCache cacheId=\"{}\" r:id=\"{}\"/>",
            cache_id, rid
        ));

        content_type_overrides.push_str(&format!(
            "<Override PartName=\"/xl/pivotCache/pivotCacheDefinition{n}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml\"/>\
<Override PartName=\"/xl/pivotCache/pivotCacheRecords{n}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml\"/>\
<Override PartName=\"/xl/pivotTables/pivotTable{n}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml\"/>",
            n = n
        ));

        let sheet_name = sheets
            .iter()
            .find(|s| s.id == unit.dest_sheet_id)
            .map(|s| s.name.clone())
            .ok_or("Pivot table's destination sheet vanished during export")?;
        let target = sheet_name_to_filename.get(&sheet_name).ok_or_else(|| {
            format!(
                "Could not resolve worksheet file for sheet '{}'",
                sheet_name
            )
        })?;
        let basename = target.rsplit('/').next().unwrap_or(target);
        let sheet_rels_path = format!("xl/worksheets/_rels/{}.rels", basename);
        if !sheet_rels_edits.contains_key(&sheet_rels_path) {
            let existing = get_zip_file_content(&mut archive, &sheet_rels_path).unwrap_or_else(|| {
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"></Relationships>"
                    .to_string()
            });
            sheet_rels_edits.insert(sheet_rels_path.clone(), existing);
        }
        let entry = sheet_rels_edits.get_mut(&sheet_rels_path).unwrap();
        let pivot_rel = format!(
            "<Relationship Id=\"rIdPivot{n}\" Type=\"{}\" Target=\"../pivotTables/pivotTable{n}.xml\"/>",
            REL_PIVOT_TABLE,
            n = n
        );
        *entry = entry.replacen(
            "</Relationships>",
            &format!("{}</Relationships>", pivot_rel),
            1,
        );
    }

    let new_workbook_xml = workbook_xml.replacen(
        "</workbook>",
        &format!("<pivotCaches>{}</pivotCaches></workbook>", pivot_caches_xml),
        1,
    );
    content_types = content_types.replacen(
        "</Types>",
        &format!("{}</Types>", content_type_overrides),
        1,
    );

    rewrite_zip_with_pivot_parts(
        &xlsx_bytes,
        content_types,
        new_workbook_xml,
        new_workbook_rels_xml,
        sheet_rels_edits,
        new_files,
    )
}

fn rewrite_zip_with_pivot_parts(
    original: &[u8],
    content_types: String,
    workbook_xml: String,
    workbook_rels_xml: String,
    mut sheet_rels_edits: HashMap<String, String>,
    new_files: Vec<(String, String)>,
) -> Result<Vec<u8>, String> {
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(original))
        .map_err(|e| format!("Failed to re-open generated xlsx zip: {}", e))?;
    let mut writer = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
    let options = zip::write::SimpleFileOptions::default();

    for i in 0..archive.len() {
        let mut file = archive.by_index(i).map_err(|e| e.to_string())?;
        let name = file.name().to_string();
        let mut buf = Vec::new();
        file.read_to_end(&mut buf).map_err(|e| e.to_string())?;
        drop(file);

        writer
            .start_file(&name, options)
            .map_err(|e| e.to_string())?;
        if name == "[Content_Types].xml" {
            writer
                .write_all(content_types.as_bytes())
                .map_err(|e| e.to_string())?;
        } else if name == "xl/workbook.xml" {
            writer
                .write_all(workbook_xml.as_bytes())
                .map_err(|e| e.to_string())?;
        } else if name == "xl/_rels/workbook.xml.rels" {
            writer
                .write_all(workbook_rels_xml.as_bytes())
                .map_err(|e| e.to_string())?;
        } else if let Some(edited) = sheet_rels_edits.remove(&name) {
            writer
                .write_all(edited.as_bytes())
                .map_err(|e| e.to_string())?;
        } else {
            writer.write_all(&buf).map_err(|e| e.to_string())?;
        }
    }

    for (path, content) in sheet_rels_edits {
        writer
            .start_file(&path, options)
            .map_err(|e| e.to_string())?;
        writer
            .write_all(content.as_bytes())
            .map_err(|e| e.to_string())?;
    }
    for (path, content) in new_files {
        writer
            .start_file(&path, options)
            .map_err(|e| e.to_string())?;
        writer
            .write_all(content.as_bytes())
            .map_err(|e| e.to_string())?;
    }

    let cursor = writer.finish().map_err(|e| e.to_string())?;
    Ok(cursor.into_inner())
}

// ---------------------------------------------------------------------------
// Import: reconstruct PivotTable definitions from a workbook's pivot XML.
// ---------------------------------------------------------------------------

enum PivotXmlSection {
    None,
    RowFields,
    ColFields,
}

struct ParsedPivotTable {
    name: String,
    cache_id: String,
    row_grand_totals: bool,
    col_grand_totals: bool,
    location_ref: Option<String>,
    /// Field indices (into the cache's `cacheFields`) used as row/col
    /// fields, in order; the `-2` "Values" pseudo-field sentinel is
    /// filtered out already.
    row_field_x: Vec<usize>,
    col_field_x: Vec<usize>,
    page_field_fld: Vec<usize>,
    /// (source field index, aggregation, display name).
    data_fields: Vec<(usize, PivotAggregation, String)>,
}

fn parse_pivot_table_xml(xml: &str) -> Option<ParsedPivotTable> {
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buf = Vec::new();
    let mut name = String::new();
    let mut cache_id = String::new();
    let mut row_grand_totals = true;
    let mut col_grand_totals = true;
    let mut location_ref = None;
    let mut row_field_x = Vec::new();
    let mut col_field_x = Vec::new();
    let mut page_field_fld = Vec::new();
    let mut data_fields = Vec::new();
    let mut section = PivotXmlSection::None;

    loop {
        let event = match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            Ok(e) => e,
        };
        match event {
            quick_xml::events::Event::Start(ref e) | quick_xml::events::Event::Empty(ref e) => {
                let local = e.name().local_name().into_inner().to_vec();
                match local.as_slice() {
                    b"pivotTableDefinition" => {
                        name = get_attr(e, b"name").unwrap_or_default();
                        cache_id = get_attr(e, b"cacheId").unwrap_or_default();
                        row_grand_totals = get_attr(e, b"rowGrandTotals").as_deref() != Some("0");
                        col_grand_totals = get_attr(e, b"colGrandTotals").as_deref() != Some("0");
                    }
                    b"location" => location_ref = get_attr(e, b"ref"),
                    b"rowFields" => section = PivotXmlSection::RowFields,
                    b"colFields" => section = PivotXmlSection::ColFields,
                    b"field" => {
                        if let Some(x) = get_attr(e, b"x").and_then(|s| s.parse::<i64>().ok())
                            && x >= 0
                        {
                            match section {
                                PivotXmlSection::RowFields => row_field_x.push(x as usize),
                                PivotXmlSection::ColFields => col_field_x.push(x as usize),
                                PivotXmlSection::None => {}
                            }
                        }
                    }
                    b"pageField" => {
                        if let Some(fld) = get_attr(e, b"fld").and_then(|s| s.parse::<usize>().ok())
                        {
                            page_field_fld.push(fld);
                        }
                    }
                    b"dataField" => {
                        if let (Some(fld), Some(dname)) = (
                            get_attr(e, b"fld").and_then(|s| s.parse::<usize>().ok()),
                            get_attr(e, b"name"),
                        ) {
                            let agg = match get_attr(e, b"subtotal").as_deref() {
                                Some("count") => PivotAggregation::Count,
                                Some("countNums") => PivotAggregation::CountNumbers,
                                Some("average") => PivotAggregation::Average,
                                Some("max") => PivotAggregation::Max,
                                Some("min") => PivotAggregation::Min,
                                _ => PivotAggregation::Sum,
                            };
                            data_fields.push((fld, agg, dname));
                        }
                    }
                    _ => {}
                }
            }
            quick_xml::events::Event::End(ref e) => {
                let local = e.name().local_name().into_inner();
                if matches!(local, b"rowFields" | b"colFields") {
                    section = PivotXmlSection::None;
                }
            }
            _ => {}
        }
        buf.clear();
    }

    if name.is_empty() {
        return None;
    }
    Some(ParsedPivotTable {
        name,
        cache_id,
        row_grand_totals,
        col_grand_totals,
        location_ref,
        row_field_x,
        col_field_x,
        page_field_fld,
        data_fields,
    })
}

fn parse_pivot_caches(workbook_xml: &str) -> HashMap<String, String> {
    let mut map = HashMap::new();
    let mut reader = quick_xml::Reader::from_str(workbook_xml);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Eof) => break,
            Ok(quick_xml::events::Event::Start(ref e))
            | Ok(quick_xml::events::Event::Empty(ref e))
                if e.local_name().as_ref() == b"pivotCache" =>
            {
                if let (Some(cache_id), Some(rid)) = (get_attr(e, b"cacheId"), get_attr(e, b"id")) {
                    map.insert(cache_id, rid);
                }
            }
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    map
}

struct ParsedCacheDefinition {
    field_names: Vec<String>,
    source_sheet: String,
    source_ref: String,
}

fn parse_cache_definition_xml(xml: &str) -> Option<ParsedCacheDefinition> {
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buf = Vec::new();
    let mut field_names = Vec::new();
    let mut source_sheet = String::new();
    let mut source_ref = String::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            Ok(quick_xml::events::Event::Start(ref e))
            | Ok(quick_xml::events::Event::Empty(ref e)) => {
                let local = e.name().local_name().into_inner();
                if local == b"cacheField" {
                    field_names.push(get_attr(e, b"name").unwrap_or_default());
                } else if local == b"worksheetSource" {
                    source_sheet = get_attr(e, b"sheet").unwrap_or_default();
                    source_ref = get_attr(e, b"ref").unwrap_or_default();
                }
            }
            _ => {}
        }
        buf.clear();
    }
    if field_names.is_empty() {
        None
    } else {
        Some(ParsedCacheDefinition {
            field_names,
            source_sheet,
            source_ref,
        })
    }
}

/// Parses an A1 cell reference (e.g. "C4") into 0-based (row, col).
fn parse_a1_cell(s: &str) -> Option<(usize, usize)> {
    let col_end = s.find(|c: char| c.is_ascii_digit())?;
    let (col_part, row_part) = s.split_at(col_end);
    if col_part.is_empty() || row_part.is_empty() {
        return None;
    }
    let mut col = 0usize;
    for c in col_part.chars() {
        if !c.is_ascii_alphabetic() {
            return None;
        }
        col = col * 26 + (c.to_ascii_uppercase() as usize - 'A' as usize + 1);
    }
    let row: usize = row_part.parse().ok()?;
    Some((row.checked_sub(1)?, col - 1))
}

/// Parses an A1 range (e.g. "A1:C4", or a single cell "A1") into 0-based
/// `(start_row, start_col, end_row, end_col)`.
pub(crate) fn parse_a1_range(s: &str) -> Option<(usize, usize, usize, usize)> {
    if let Some((start, end)) = s.split_once(':') {
        let (r0, c0) = parse_a1_cell(start)?;
        let (r1, c1) = parse_a1_cell(end)?;
        Some((r0, c0, r1, c1))
    } else {
        let (r, c) = parse_a1_cell(s)?;
        Some((r, c, r, c))
    }
}

/// Reconstructs every `PivotTable` definable from a workbook's pivot XML
/// parts, matching sheets by name against `sheet_id_by_name` (the already
/// -imported sheets, keyed by their possibly de-duplicated import name) and
/// tables by name against `find_table`. Best-effort: silently skips any
/// pivot table whose parts can't be fully resolved rather than failing the
/// whole import.
pub fn import_pivot_tables(
    buffer: &[u8],
    sheet_id_by_name: &HashMap<String, u64>,
    find_matching_table: impl Fn(&str, usize, usize, usize, usize) -> Option<String>,
) -> Vec<PivotTable> {
    let mut result = Vec::new();
    let Ok(mut archive) = zip::ZipArchive::new(std::io::Cursor::new(buffer)) else {
        return result;
    };
    let Some(workbook_xml) = get_zip_file_content(&mut archive, "xl/workbook.xml") else {
        return result;
    };
    let Some(workbook_rels_xml) = get_zip_file_content(&mut archive, "xl/_rels/workbook.xml.rels")
    else {
        return result;
    };
    let cache_id_to_rid = parse_pivot_caches(&workbook_xml);
    let rid_to_target = parse_workbook_rels(&workbook_rels_xml);
    let rid_to_sheet_name = parse_workbook_sheets(&workbook_xml);
    let sheet_rid_to_target = parse_workbook_rels(&workbook_rels_xml);
    let mut sheet_filename_to_name: HashMap<String, String> = HashMap::new();
    for (rid, name) in &rid_to_sheet_name {
        if let Some(target) = sheet_rid_to_target.get(rid) {
            let basename = target.rsplit('/').next().unwrap_or(target).to_string();
            sheet_filename_to_name.insert(basename, name.clone());
        }
    }

    let pivot_table_files: Vec<String> = archive
        .file_names()
        .filter(|n| n.starts_with("xl/pivotTables/") && n.ends_with(".xml"))
        .map(|s| s.to_string())
        .collect();

    for pt_file in pivot_table_files {
        let Some(pt_xml) = get_zip_file_content(&mut archive, &pt_file) else {
            continue;
        };
        let Some(parsed) = parse_pivot_table_xml(&pt_xml) else {
            continue;
        };

        // Destination sheet: whichever worksheet's rels points at this file.
        let pt_basename = pt_file.rsplit('/').next().unwrap_or(&pt_file);
        let worksheet_rels_files: Vec<String> = archive
            .file_names()
            .filter(|n| n.starts_with("xl/worksheets/_rels/") && n.ends_with(".rels"))
            .map(|s| s.to_string())
            .collect();
        let dest_sheet_name = worksheet_rels_files.into_iter().find_map(|rels_path| {
            let content = get_zip_file_content(&mut archive, &rels_path)?;
            if content.contains(pt_basename) {
                let sheet_file = rels_path
                    .strip_prefix("xl/worksheets/_rels/")?
                    .strip_suffix(".rels")?;
                sheet_filename_to_name.get(sheet_file).cloned()
            } else {
                None
            }
        });
        let Some(dest_sheet_name) = dest_sheet_name else {
            continue;
        };
        let Some(&dest_sheet_id) = sheet_id_by_name.get(&dest_sheet_name) else {
            continue;
        };

        let Some(rid) = cache_id_to_rid.get(&parsed.cache_id) else {
            continue;
        };
        let Some(cache_target) = rid_to_target.get(rid) else {
            continue;
        };
        // `cache_target` is relative to `xl/` (e.g. "pivotCache/pivotCacheDefinition1.xml").
        let cache_path = format!("xl/{}", cache_target.trim_start_matches('/'));
        let Some(cache_xml) = get_zip_file_content(&mut archive, &cache_path) else {
            continue;
        };
        let Some(cache_def) = parse_cache_definition_xml(&cache_xml) else {
            continue;
        };

        let Some(&source_sheet_id) = sheet_id_by_name.get(&cache_def.source_sheet) else {
            continue;
        };
        let Some((src_start_row, src_start_col, src_end_row, src_end_col)) =
            parse_a1_range(&cache_def.source_ref)
        else {
            continue;
        };

        let source = match find_matching_table(
            &cache_def.source_sheet,
            src_start_row,
            src_start_col,
            src_end_row,
            src_end_col,
        ) {
            Some(table_name) => PivotSource::Table { name: table_name },
            None => PivotSource::Range {
                sheet_id: source_sheet_id,
                start_row: src_start_row,
                start_col: src_start_col,
                end_row: src_end_row,
                end_col: src_end_col,
            },
        };

        let field_name = |idx: usize| cache_def.field_names.get(idx).cloned().unwrap_or_default();

        let row_fields: Vec<PivotField> = parsed
            .row_field_x
            .iter()
            .map(|&x| PivotField::new(field_name(x)))
            .collect();
        let col_fields: Vec<PivotField> = parsed
            .col_field_x
            .iter()
            .map(|&x| PivotField::new(field_name(x)))
            .collect();
        let filter_fields: Vec<PivotFilterField> = parsed
            .page_field_fld
            .iter()
            .map(|&fld| PivotFilterField::new(field_name(fld)))
            .collect();
        let value_fields: Vec<PivotValueField> = parsed
            .data_fields
            .iter()
            .map(|(fld, agg, dname)| {
                let column = field_name(*fld);
                let default_label = PivotValueField::new(column.clone(), *agg).label();
                let custom_name = if *dname == default_label {
                    None
                } else {
                    Some(dname.clone())
                };
                PivotValueField {
                    column,
                    aggregation: *agg,
                    custom_name,
                }
            })
            .collect();

        // `<location>` spans only the row/col header + data grid, not the
        // filter/page-field rows reserved above it on export (see
        // `build_pivot_xml_unit`), so `dest_row` -- the anchor of the whole
        // visual block, filter rows included -- needs that offset added
        // back in.
        let filter_row_offset = if filter_fields.is_empty() {
            0
        } else {
            filter_fields.len() + 1
        };
        let (dest_row, dest_col, last_end_row, last_end_col) = match &parsed.location_ref {
            Some(loc) => match parse_a1_range(loc) {
                Some((r0, c0, r1, c1)) => {
                    (r0.saturating_sub(filter_row_offset), c0, Some(r1), Some(c1))
                }
                None => (0, 0, None, None),
            },
            None => (0, 0, None, None),
        };

        result.push(PivotTable {
            id: crate::core::engine::generate_unique_id(),
            name: parsed.name,
            source,
            dest_sheet_id,
            dest_row,
            dest_col,
            row_fields,
            col_fields,
            value_fields,
            filter_fields,
            grand_totals_row: parsed.row_grand_totals,
            grand_totals_col: parsed.col_grand_totals,
            last_output_end_row: last_end_row,
            last_output_end_col: last_end_col,
        });
    }

    result
}
