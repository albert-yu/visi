#![allow(missing_docs)]
pub(crate) mod chart;
pub(crate) mod dom;
pub(crate) mod patch;
pub(crate) mod shared_strings;
pub(crate) mod styles;
pub(crate) mod table;
pub(crate) mod workbook;
pub(crate) mod worksheet;
pub(crate) mod xml_util;

use crate::core::opc::OpcPackage;
use crate::core::{CellStyle, CellType, DataColumn, Sheet};
use web_time::Instant;

pub use crate::core::excel::ImportedSheet;
pub type ImportedXlsxData = crate::core::excel::ImportedWorkbookData;

pub fn import_xlsx_data(
    buffer: &[u8],
    existing_sheets: &[Sheet],
    progress_callback: impl FnMut(usize, usize, &str),
) -> crate::Result<ImportedXlsxData> {
    import_xlsx_data_raw(buffer, existing_sheets, progress_callback).map_err(crate::Error::Xlsx)
}

pub(crate) fn import_xlsx_with_source(
    buffer: &[u8],
) -> crate::Result<(ImportedXlsxData, patch::OoxmlSourceState)> {
    let pkg = OpcPackage::from_bytes(buffer)
        .map_err(|e| crate::Error::Xlsx(format!("Failed to open Excel file: {e}")))?;
    let wb = workbook::read_workbook(&pkg).map_err(crate::Error::Xlsx)?;
    let data = import_xlsx_from_parsed(buffer, &pkg, &wb, &[], |_, _, _| {})
        .map_err(crate::Error::Xlsx)?;
    let sheets: Vec<Sheet> = data.0.iter().map(|s| s.sheet.clone()).collect();
    let source =
        patch::OoxmlSourceState::capture(pkg, &wb, &sheets, &data.1, &data.2, data.3.as_ref());
    Ok((data, source))
}

pub(crate) fn import_xlsx_data_raw(
    buffer: &[u8],
    existing_sheets: &[Sheet],
    progress_callback: impl FnMut(usize, usize, &str),
) -> Result<ImportedXlsxData, String> {
    let pkg =
        OpcPackage::from_bytes(buffer).map_err(|e| format!("Failed to open Excel file: {e}"))?;
    let wb = workbook::read_workbook(&pkg)?;
    import_xlsx_from_parsed(buffer, &pkg, &wb, existing_sheets, progress_callback)
}

fn import_xlsx_from_parsed(
    buffer: &[u8],
    pkg: &OpcPackage,
    wb: &workbook::ParsedWorkbook,
    existing_sheets: &[Sheet],
    mut progress_callback: impl FnMut(usize, usize, &str),
) -> Result<ImportedXlsxData, String> {
    let start_total = Instant::now();

    let shared_strings = wb
        .shared_strings_part
        .as_deref()
        .and_then(|part| pkg.get_str(part))
        .map(|xml| shared_strings::parse_shared_strings(&xml))
        .transpose()?
        .unwrap_or_default();

    let parsed_styles = wb
        .styles_part
        .as_deref()
        .and_then(|part| pkg.get_str(part))
        .map(|xml| styles::parse_styles(&xml))
        .unwrap_or_default();

    let total_sheets = wb.sheets.len();
    let mut imported_sheets = Vec::new();
    let mut orig_to_assigned_name: std::collections::HashMap<String, String> =
        std::collections::HashMap::new();

    for (sheet_idx, sheet_meta) in wb.sheets.iter().enumerate() {
        let orig_sheet_name = &sheet_meta.name;
        progress_callback(sheet_idx, total_sheets, orig_sheet_name);

        let Some(sheet_xml) = pkg.get_str(&sheet_meta.part_path) else {
            log::warn!(
                "Skipping worksheet '{}': missing part '{}'",
                orig_sheet_name,
                sheet_meta.part_path
            );
            continue;
        };

        let parsed_ws = worksheet::parse_worksheet(&sheet_xml, &shared_strings, &parsed_styles)?;
        let (rows, cols) = if parsed_ws.rows == 0 || parsed_ws.cols == 0 {
            (10, 5)
        } else {
            (parsed_ws.rows, parsed_ws.cols)
        };

        let mut sheet_name = orig_sheet_name.clone();
        let mut count = 1;
        while existing_sheets.iter().any(|s| s.name == sheet_name)
            || imported_sheets
                .iter()
                .any(|imported: &ImportedSheet| imported.sheet.name == sheet_name)
        {
            sheet_name = format!("{}_{}", orig_sheet_name, count);
            count += 1;
        }

        let mut columns = (0..cols)
            .map(|col_idx| {
                let mut col = DataColumn::new(rows);
                col.id = rand::random::<u64>();
                col.name = crate::core::parser::col_idx_to_letters(col_idx);
                if let Some(&width) = parsed_ws.column_widths.get(&col_idx) {
                    col.width = Some(width);
                }
                if parsed_ws.hidden_cols.contains(&col_idx) {
                    col.hidden = true;
                }
                col
            })
            .collect::<Vec<_>>();

        for cell in parsed_ws.cells {
            if cell.row >= rows || cell.col >= cols {
                continue;
            }
            let col = &mut columns[cell.col];
            if let Some(style_idx) = cell.style_idx
                && let Some(style) = parsed_styles.xf_styles.get(&style_idx)
            {
                col.styles[cell.row] = Some(style.clone());
            }

            if let Some(formula) = cell.formula {
                let init_res = match &cell.value {
                    worksheet::ParsedCellValue::Empty => crate::core::engine::ResultData::None,
                    worksheet::ParsedCellValue::String(s) if s.is_empty() => {
                        crate::core::engine::ResultData::None
                    }
                    other => other.to_result_data(),
                };
                col.data.set(cell.row, init_res);
                col.src[cell.row] = format!("={formula}");
                col.cell_types[cell.row] = cell.value.to_cell_type_and_src().0;
            } else {
                let init_res = cell.value.to_result_data();
                let (cell_type, cell_src) = cell.value.to_cell_type_and_src();
                col.data.set(cell.row, init_res);
                col.src[cell.row] = cell_src;
                col.cell_types[cell.row] = cell_type;
            }
        }

        let mut sheet_row_heights = vec![None; rows];
        for (&row_idx, &height) in &parsed_ws.row_heights {
            if row_idx < sheet_row_heights.len() {
                sheet_row_heights[row_idx] = Some(height);
            }
        }
        let mut sheet_hidden_rows = vec![false; rows];
        for &row_idx in &parsed_ws.hidden_rows {
            if row_idx < sheet_hidden_rows.len() {
                sheet_hidden_rows[row_idx] = true;
            }
        }

        orig_to_assigned_name.insert(orig_sheet_name.clone(), sheet_name.clone());
        let new_sheet = Sheet {
            id: rand::random::<u64>(),
            name: sheet_name,
            columns,
            row_heights: sheet_row_heights,
            hidden_rows: sheet_hidden_rows,
            merged_cells: parsed_ws.merged_cells,
            visibility: workbook::parse_sheet_visibility(sheet_meta.state.as_deref()),
            freeze_pane: parsed_ws.freeze_pane,
            date1904: wb.date1904,
            tables: Vec::new(),
            dependencies: std::collections::HashMap::new(),
            dependencies_rev: std::collections::HashMap::new(),
            uncommitted_actions: Vec::new(),
            locale: crate::core::locale::Locale::default(),
        };
        imported_sheets.push(ImportedSheet { sheet: new_sheet });
    }

    if imported_sheets.is_empty() {
        return Err("No worksheets found in the Excel file".to_string());
    }

    for (orig_sheet_name, parsed) in import_tables_from_opc(&pkg).unwrap_or_default() {
        let Some(assigned_sheet_name) = orig_to_assigned_name.get(&orig_sheet_name) else {
            continue;
        };
        let Some(imported) = imported_sheets
            .iter_mut()
            .find(|imported: &&mut ImportedSheet| &imported.sheet.name == assigned_sheet_name)
        else {
            continue;
        };

        let Some((start_row, start_col, end_row, end_col)) = parsed.bounds else {
            log::warn!(
                "Skipping table '{}' on sheet '{}': missing or unparseable ref",
                parsed.name,
                orig_sheet_name
            );
            continue;
        };

        if end_row >= imported.sheet.row_count() || end_col >= imported.sheet.col_count() {
            imported.sheet.ensure_capacity(end_row, end_col);
        }

        imported.sheet.tables.push(crate::core::table::ExcelTable {
            id: rand::random::<u64>(),
            name: parsed.name,
            sheet_id: imported.sheet.id,
            start_row,
            start_col,
            end_row,
            end_col,
            has_header_row: parsed.has_header_row,
            has_totals_row: parsed.has_totals_row,
            columns: parsed.columns,
            style_name: parsed.style_name,
            has_insert_row: parsed.has_insert_row,
        });
    }

    log::info!(
        "import_xlsx_data finished. Total time: {:.2?}",
        start_total.elapsed()
    );

    let mut imported_charts = Vec::new();
    if let Ok(parsed_charts) = import_charts_from_opc(&pkg) {
        let mut chart_index_by_sheet: std::collections::HashMap<String, usize> =
            std::collections::HashMap::new();
        for parsed in parsed_charts {
            if let Some(imported_sheet) = imported_sheets
                .iter()
                .find(|imported| imported.sheet.name == parsed.sheet_name)
            {
                let data_range = if parsed.info.data_range.is_empty() {
                    format!("{}!A1", imported_sheet.sheet.name)
                } else {
                    parsed.info.data_range.clone()
                };

                let index_in_sheet = chart_index_by_sheet
                    .entry(parsed.sheet_name.clone())
                    .or_insert(0);
                let chart_id = deterministic_chart_id(&parsed.sheet_name, *index_in_sheet);
                *index_in_sheet += 1;

                imported_charts.push(crate::core::chart::Chart {
                    id: chart_id,
                    name: format!("Chart {}", imported_charts.len() + 1),
                    chart_type: parsed.info.chart_type.clone(),
                    data_range,
                    title: parsed.info.title.clone(),
                    xlabel: parsed.info.xlabel.clone(),
                    ylabel: parsed.info.ylabel.clone(),
                    show_legend: parsed.info.show_legend,
                    anchor_row: parsed.anchor.from_row.max(0) as usize,
                    anchor_col: parsed.anchor.from_col.max(0) as usize,
                });
            }
        }
    }

    let sheet_id_by_name: std::collections::HashMap<String, u64> = imported_sheets
        .iter()
        .map(|imported| (imported.sheet.name.clone(), imported.sheet.id))
        .collect();
    let imported_pivots = crate::core::pivot_xlsx::import_pivot_tables(
        buffer,
        &sheet_id_by_name,
        |sheet_name, r0, c0, r1, c1| {
            imported_sheets
                .iter()
                .find(|imported| imported.sheet.name == sheet_name)
                .and_then(|imported| {
                    imported
                        .sheet
                        .tables
                        .iter()
                        .find(|tbl| {
                            (tbl.start_row, tbl.start_col, tbl.end_row, tbl.end_col)
                                == (r0, c0, r1, c1)
                        })
                        .map(|tbl| tbl.name.clone())
                })
        },
    );

    let imported_vba = crate::core::vba_xlsx::import_vba_project(buffer, &sheet_id_by_name)?;

    Ok((
        imported_sheets,
        imported_charts,
        imported_pivots,
        imported_vba,
    ))
}

fn format_result_for_xlsx(res_data: &crate::core::engine::ResultData) -> String {
    match res_data {
        crate::core::engine::ResultData::Integer(i) => i.to_string(),
        crate::core::engine::ResultData::Float(f) => {
            if f.is_nan() || f.is_infinite() {
                "#NUM!".to_string()
            } else {
                f.to_string()
            }
        }
        crate::core::engine::ResultData::Boolean(b) => {
            if *b {
                "TRUE".to_string()
            } else {
                "FALSE".to_string()
            }
        }
        crate::core::engine::ResultData::String(s) => s.clone(),
        crate::core::engine::ResultData::Error(e) => {
            let upper = e.to_uppercase();
            if upper.contains("#DIV/0!")
                || upper.contains("DIVISION BY ZERO")
                || upper.contains("DIV/0")
            {
                "#DIV/0!".to_string()
            } else if upper.contains("#N/A") {
                "#N/A".to_string()
            } else if upper.contains("#REF!") {
                "#REF!".to_string()
            } else if upper.contains("#NUM!") {
                "#NUM!".to_string()
            } else if upper.contains("#NAME?") {
                "#NAME?".to_string()
            } else if upper.contains("#NULL!") {
                "#NULL!".to_string()
            } else if upper.contains("#CALC!") {
                "#CALC!".to_string()
            } else if upper.contains("#SPILL!") {
                "#SPILL!".to_string()
            } else if upper.starts_with('#') {
                e.clone()
            } else {
                "#VALUE!".to_string()
            }
        }
        crate::core::engine::ResultData::None => "".to_string(),
        crate::core::engine::ResultData::List(l) => {
            if let Some(first) = l.first() {
                format_result_for_xlsx(first)
            } else {
                "".to_string()
            }
        }
        _ => res_data.to_string(),
    }
}

fn date_serial_for_export(
    style: Option<&crate::core::CellStyle>,
    col: &crate::core::engine::DataColumn,
    row_idx: usize,
) -> Option<f64> {
    let code = style?.num_format.as_deref()?;
    if !crate::core::date::is_date_code(code) {
        return None;
    }
    match col.data.get(row_idx)? {
        crate::core::engine::ResultData::Float(f) => Some(f),
        crate::core::engine::ResultData::Integer(i) => Some(i as f64),
        _ => None,
    }
}

pub fn export_xlsx_data(
    sheets: &[Sheet],
    charts: &[crate::core::chart::Chart],
    pivots: &[crate::core::pivot::PivotTable],
    vba: Option<&crate::core::vba::VbaProject>,
) -> crate::Result<Vec<u8>> {
    export_xlsx_data_raw(sheets, charts, pivots, vba).map_err(crate::Error::Xlsx)
}

pub(crate) fn export_xlsx_data_raw(
    sheets: &[Sheet],
    charts: &[crate::core::chart::Chart],
    pivots: &[crate::core::pivot::PivotTable],
    vba: Option<&crate::core::vba::VbaProject>,
) -> Result<Vec<u8>, String> {
    use crate::core::opc::{
        CT_CHART, CT_CORE_PROPERTIES, CT_DRAWING, CT_EXTENDED_PROPERTIES, CT_RELS,
        CT_SHARED_STRINGS, CT_STYLES, CT_TABLE, CT_WORKBOOK, CT_WORKSHEET, CT_XML, ContentTypes,
        REL_CHART, REL_CORE_PROPERTIES, REL_DRAWING, REL_EXTENDED_PROPERTIES, REL_OFFICE_DOCUMENT,
        REL_SHARED_STRINGS, REL_STYLES, REL_TABLE, REL_WORKSHEET, Relationship,
    };
    use crate::core::xlsx::xml_util::NamespaceFlavor;

    let flavor = NamespaceFlavor::Transitional;
    let mut pkg = OpcPackage::new();
    let mut ct = ContentTypes::default();
    ct.ensure_default("rels", CT_RELS);
    ct.ensure_default("xml", CT_XML);

    let mut used_names = std::collections::HashSet::new();
    let mut sheet_name_to_ws_name = std::collections::HashMap::new();
    let mut sheet_id_to_worksheet_name = std::collections::HashMap::new();
    let mut ws_names: Vec<String> = Vec::with_capacity(sheets.len());

    for sheet in sheets {
        let mut worksheet_name = if sheet.name.len() > 31 {
            sheet.name[..31].to_string()
        } else {
            sheet.name.clone()
        };
        let mut counter = 1;
        while used_names.contains(&worksheet_name.to_lowercase()) {
            let counter_suffix = format!("_{counter}");
            let max_name_len = 31 - counter_suffix.len();
            let base_name = if sheet.name.len() > max_name_len {
                &sheet.name[..max_name_len]
            } else {
                &sheet.name
            };
            worksheet_name = format!("{base_name}{counter_suffix}");
            counter += 1;
        }
        used_names.insert(worksheet_name.to_lowercase());
        sheet_name_to_ws_name.insert(sheet.name.clone(), worksheet_name.clone());
        sheet_id_to_worksheet_name.insert(sheet.id, worksheet_name.clone());
        ws_names.push(worksheet_name);
    }

    let mut charts_by_ws: std::collections::HashMap<String, Vec<&crate::core::chart::Chart>> =
        std::collections::HashMap::new();
    for ch in charts {
        let ref_sheet = ch.data_range.split('!').next().unwrap_or("").trim();
        let lookup = if ref_sheet.is_empty() {
            sheets.first().map(|s| s.name.as_str()).unwrap_or("")
        } else {
            ref_sheet.trim_matches('\'')
        };
        if let Some(ws_name) = sheet_name_to_ws_name
            .get(lookup)
            .or_else(|| ws_names.first())
            .cloned()
        {
            charts_by_ws.entry(ws_name).or_default().push(ch);
        }
    }

    let mut sst = shared_strings::SharedStringTable::default();
    let mut style_reg = styles::StyleRegistry::default();
    let mut wb_rels: Vec<Relationship> = Vec::new();
    let mut wb_sheets: Vec<(String, String, crate::core::SheetVisibility)> = Vec::new();
    let date1904 = sheets.iter().any(|s| s.date1904);

    let mut next_table_id = 1usize;
    let mut next_drawing_id = 1usize;
    let mut next_chart_id = 1usize;

    for (idx, sheet) in sheets.iter().enumerate() {
        let sheet_num = idx + 1;
        let ws_name = &ws_names[idx];
        let sheet_part = format!("xl/worksheets/sheet{sheet_num}.xml");
        let wb_rid = format!("rId{sheet_num}");
        wb_sheets.push((ws_name.clone(), wb_rid.clone(), sheet.visibility));
        wb_rels.push(Relationship {
            id: wb_rid,
            rel_type: REL_WORKSHEET.to_string(),
            target: format!("worksheets/sheet{sheet_num}.xml"),
            target_mode: None,
        });

        let mut sheet_rels: Vec<Relationship> = Vec::new();
        let mut table_rids: Vec<String> = Vec::new();

        for tbl in &sheet.tables {
            let tbl_id = next_table_id;
            next_table_id += 1;
            let tbl_part = format!("xl/tables/table{tbl_id}.xml");
            let tbl_xml = table::serialize_table_xml(tbl, tbl_id, flavor);
            pkg.set_part(&tbl_part, tbl_xml.into_bytes());
            ct.set_override(&tbl_part, CT_TABLE);

            let rid = format!("rId{}", sheet_rels.len() + 1);
            sheet_rels.push(Relationship {
                id: rid.clone(),
                rel_type: REL_TABLE.to_string(),
                target: format!("../tables/table{tbl_id}.xml"),
                target_mode: None,
            });
            table_rids.push(rid);
        }

        let mut drawing_rid: Option<String> = None;
        if let Some(ws_charts) = charts_by_ws.get(ws_name)
            && !ws_charts.is_empty()
        {
            let drawing_id = next_drawing_id;
            next_drawing_id += 1;
            let drawing_part = format!("xl/drawings/drawing{drawing_id}.xml");
            let mut drawing_rels: Vec<Relationship> = Vec::new();
            let mut anchors: Vec<(usize, usize, String)> = Vec::new();

            for ch in ws_charts {
                let chart_id = next_chart_id;
                next_chart_id += 1;
                let chart_part = format!("xl/charts/chart{chart_id}.xml");
                let chart_xml = chart::serialize_chart_xml(ch, ws_name, flavor);
                pkg.set_part(&chart_part, chart_xml.into_bytes());
                ct.set_override(&chart_part, CT_CHART);

                let c_rid = format!("rId{}", drawing_rels.len() + 1);
                drawing_rels.push(Relationship {
                    id: c_rid.clone(),
                    rel_type: REL_CHART.to_string(),
                    target: format!("../charts/chart{chart_id}.xml"),
                    target_mode: None,
                });
                anchors.push((ch.anchor_row, ch.anchor_col, c_rid));
            }

            let drawing_xml = chart::serialize_drawing_xml(&anchors, flavor);
            pkg.set_part(&drawing_part, drawing_xml.into_bytes());
            pkg.write_relationships(&drawing_part, &drawing_rels);
            ct.set_override(&drawing_part, CT_DRAWING);

            let d_rid = format!("rId{}", sheet_rels.len() + 1);
            sheet_rels.push(Relationship {
                id: d_rid.clone(),
                rel_type: REL_DRAWING.to_string(),
                target: format!("../drawings/drawing{drawing_id}.xml"),
                target_mode: None,
            });
            drawing_rid = Some(d_rid);
        }

        let ws_xml = worksheet::write_worksheet_xml(
            sheet,
            &mut sst,
            &mut style_reg,
            &table_rids,
            drawing_rid.as_deref(),
            flavor,
        );
        pkg.set_part(&sheet_part, ws_xml.into_bytes());
        ct.set_override(&sheet_part, CT_WORKSHEET);
        if !sheet_rels.is_empty() {
            pkg.write_relationships(&sheet_part, &sheet_rels);
        }
    }

    let styles_rid = format!("rId{}", wb_rels.len() + 1);
    wb_rels.push(Relationship {
        id: styles_rid,
        rel_type: REL_STYLES.to_string(),
        target: "styles.xml".to_string(),
        target_mode: None,
    });
    pkg.set_part("xl/styles.xml", style_reg.to_xml(flavor).into_bytes());
    ct.set_override("xl/styles.xml", CT_STYLES);

    if !sst.is_empty() {
        let sst_rid = format!("rId{}", wb_rels.len() + 1);
        wb_rels.push(Relationship {
            id: sst_rid,
            rel_type: REL_SHARED_STRINGS.to_string(),
            target: "sharedStrings.xml".to_string(),
            target_mode: None,
        });
        pkg.set_part("xl/sharedStrings.xml", sst.to_xml(flavor).into_bytes());
        ct.set_override("xl/sharedStrings.xml", CT_SHARED_STRINGS);
    }

    pkg.set_part(
        "xl/workbook.xml",
        workbook::write_workbook_xml(&wb_sheets, date1904, flavor).into_bytes(),
    );
    ct.set_override("xl/workbook.xml", CT_WORKBOOK);
    pkg.write_relationships("xl/workbook.xml", &wb_rels);

    pkg.set_part(
        "docProps/core.xml",
        b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\"><dc:creator>Visi</dc:creator></cp:coreProperties>".to_vec(),
    );
    ct.set_override("docProps/core.xml", CT_CORE_PROPERTIES);

    pkg.set_part(
        "docProps/app.xml",
        b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"><Application>Visi</Application></Properties>".to_vec(),
    );
    ct.set_override("docProps/app.xml", CT_EXTENDED_PROPERTIES);

    pkg.write_relationships(
        "",
        &[
            Relationship {
                id: "rId1".to_string(),
                rel_type: REL_OFFICE_DOCUMENT.to_string(),
                target: "xl/workbook.xml".to_string(),
                target_mode: None,
            },
            Relationship {
                id: "rId2".to_string(),
                rel_type: REL_CORE_PROPERTIES.to_string(),
                target: "docProps/core.xml".to_string(),
                target_mode: None,
            },
            Relationship {
                id: "rId3".to_string(),
                rel_type: REL_EXTENDED_PROPERTIES.to_string(),
                target: "docProps/app.xml".to_string(),
                target_mode: None,
            },
        ],
    );
    pkg.write_content_types(&ct);

    let buffer = pkg.to_bytes()?;

    let buffer = if pivots.is_empty() {
        buffer
    } else {
        crate::core::pivot_xlsx::inject_pivot_tables(buffer, sheets, pivots)?
    };

    crate::core::vba_xlsx::export_vba_project(buffer, vba, &sheet_id_to_worksheet_name)
}

#[allow(dead_code)]
struct ParsedAnchor {
    from_col: isize,
    from_row: isize,
    from_col_off: f32,
    from_row_off: f32,
    to_col: isize,
    to_row: isize,
    to_col_off: f32,
    to_row_off: f32,
    chart_rid: String,
}

struct ParsedChartInfo {
    chart_type: crate::core::chart::ChartType,
    title: Option<String>,
    xlabel: Option<String>,
    ylabel: Option<String>,
    show_legend: bool,
    data_range: String,
}

struct ParsedChartData {
    sheet_name: String,
    anchor: ParsedAnchor,
    info: ParsedChartInfo,
}

fn get_filename(path: &str) -> String {
    path.split('/').next_back().unwrap_or("").to_string()
}

fn clean_excel_formula(f: &str) -> String {
    f.replace('$', "")
}

struct ParsedRange {
    sheet: String,
    start_col: String,
    start_row: usize,
    end_col: String,
    end_row: usize,
}

fn split_col_row(s: &str) -> Option<(String, usize)> {
    let mut col = String::new();
    let mut row_str = String::new();
    for c in s.chars() {
        if c.is_ascii_alphabetic() {
            col.push(c);
        } else if c.is_ascii_digit() {
            row_str.push(c);
        }
    }
    let row = row_str.parse::<usize>().ok()?;
    if col.is_empty() {
        None
    } else {
        Some((col, row))
    }
}

fn parse_range(f: &str) -> Option<ParsedRange> {
    let cleaned = f.replace('$', "");
    let parts: Vec<&str> = cleaned.split('!').collect();
    if parts.len() != 2 {
        return None;
    }
    let sheet = parts[0].to_string();
    let range_part = parts[1];

    if range_part.contains(':') {
        let sub_parts: Vec<&str> = range_part.split(':').collect();
        if sub_parts.len() == 2 {
            let (sc, sr) = split_col_row(sub_parts[0])?;
            let (ec, er) = split_col_row(sub_parts[1])?;
            Some(ParsedRange {
                sheet,
                start_col: sc,
                start_row: sr,
                end_col: ec,
                end_row: er,
            })
        } else {
            None
        }
    } else {
        let (sc, sr) = split_col_row(range_part)?;
        Some(ParsedRange {
            sheet: sheet.clone(),
            start_col: sc.clone(),
            start_row: sr,
            end_col: sc,
            end_row: sr,
        })
    }
}

fn combine_ranges(cat_f: &str, val_f: &str) -> String {
    if let (Some(cat), Some(val)) = (parse_range(cat_f), parse_range(val_f))
        && cat.sheet == val.sheet
        && cat.start_row == val.start_row
        && cat.end_row == val.end_row
    {
        return format!(
            "{}!{}{}:{}{}",
            cat.sheet, cat.start_col, cat.start_row, val.end_col, val.end_row
        );
    }
    clean_excel_formula(val_f)
}

pub(crate) fn get_attr(e: &quick_xml::events::BytesStart, name: &[u8]) -> Option<String> {
    for attr in e.attributes().flatten() {
        let key = attr.key.as_ref();
        let local_name = if let Some(pos) = key.iter().position(|&b| b == b':') {
            &key[pos + 1..]
        } else {
            key
        };
        if local_name == name {
            let raw = String::from_utf8(attr.value.into_owned()).ok()?;
            return Some(
                quick_xml::escape::unescape(&raw)
                    .map(|value| value.into_owned())
                    .unwrap_or(raw),
            );
        }
    }
    None
}

pub(crate) fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

pub(crate) fn escape_xml_text(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
}

pub(crate) fn parse_workbook_sheets(xml: &str) -> std::collections::HashMap<String, String> {
    let mut map = std::collections::HashMap::new();
    let mut reader = quick_xml::reader::Reader::from_str(xml);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Eof) => break,
            Ok(quick_xml::events::Event::Empty(e)) | Ok(quick_xml::events::Event::Start(e)) => {
                let local = e.name().local_name().into_inner();
                if local == b"sheet"
                    && let (Some(name), Some(r_id)) = (get_attr(&e, b"name"), get_attr(&e, b"id"))
                {
                    map.insert(r_id, name);
                }
            }
            _ => {}
        }
        buf.clear();
    }
    map
}

pub(crate) fn parse_workbook_rels(xml: &str) -> std::collections::HashMap<String, String> {
    let mut map = std::collections::HashMap::new();
    let mut reader = quick_xml::reader::Reader::from_str(xml);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Eof) => break,
            Ok(quick_xml::events::Event::Empty(e)) | Ok(quick_xml::events::Event::Start(e)) => {
                let local = e.name().local_name().into_inner();
                if local == b"Relationship"
                    && let (Some(r_id), Some(target)) =
                        (get_attr(&e, b"Id"), get_attr(&e, b"Target"))
                {
                    map.insert(r_id, target);
                }
            }
            _ => {}
        }
        buf.clear();
    }
    map
}

#[allow(dead_code)]
fn parse_sheet_drawing_rels(xml: &str) -> Option<String> {
    let mut reader = quick_xml::reader::Reader::from_str(xml);
    let mut buf = Vec::new();
    let mut drawing_target = None;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Eof) => break,
            Ok(quick_xml::events::Event::Empty(e)) | Ok(quick_xml::events::Event::Start(e)) => {
                let local = e.name().local_name().into_inner();
                if local == b"Relationship"
                    && let Some(ty) = get_attr(&e, b"Type")
                    && ty.contains("relationships/drawing")
                    && let Some(target) = get_attr(&e, b"Target")
                {
                    drawing_target = Some(get_filename(&target));
                    break;
                }
            }
            _ => {}
        }
        buf.clear();
    }
    drawing_target
}

#[allow(dead_code)]
fn parse_sheet_table_rels(xml: &str) -> Vec<String> {
    let mut targets = Vec::new();
    let mut reader = quick_xml::reader::Reader::from_str(xml);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Eof) => break,
            Ok(quick_xml::events::Event::Empty(e)) | Ok(quick_xml::events::Event::Start(e)) => {
                let local = e.name().local_name().into_inner();
                if local == b"Relationship"
                    && let Some(ty) = get_attr(&e, b"Type")
                    && ty.ends_with("relationships/table")
                    && let Some(target) = get_attr(&e, b"Target")
                {
                    let filename = get_filename(&target);
                    if !filename.is_empty() {
                        targets.push(filename);
                    }
                }
            }
            _ => {}
        }
        buf.clear();
    }
    targets
}

#[allow(dead_code)]
fn parse_drawing_chart_rels(xml: &str) -> std::collections::HashMap<String, String> {
    let mut map = std::collections::HashMap::new();
    let mut reader = quick_xml::reader::Reader::from_str(xml);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Eof) => break,
            Ok(quick_xml::events::Event::Empty(e)) | Ok(quick_xml::events::Event::Start(e)) => {
                let local = e.name().local_name().into_inner();
                if local == b"Relationship"
                    && let Some(ty) = get_attr(&e, b"Type")
                    && ty.contains("relationships/chart")
                    && let (Some(r_id), Some(target)) =
                        (get_attr(&e, b"Id"), get_attr(&e, b"Target"))
                {
                    map.insert(r_id, get_filename(&target));
                }
            }
            _ => {}
        }
        buf.clear();
    }
    map
}

fn parse_drawings_xml(xml: &str) -> Vec<ParsedAnchor> {
    let mut anchors = Vec::new();
    let mut reader = quick_xml::reader::Reader::from_str(xml);
    let mut buf = Vec::new();

    let mut in_from = false;
    let mut in_to = false;
    let mut current_tag = Vec::new();

    let mut from_col = 0;
    let mut from_row = 0;
    let mut from_col_off = 0.0f32;
    let mut from_row_off = 0.0f32;

    let mut to_col = 0;
    let mut to_row = 0;
    let mut to_col_off = 0.0f32;
    let mut to_row_off = 0.0f32;

    let mut chart_rid = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Eof) => break,
            Ok(quick_xml::events::Event::Start(e)) => {
                let local = e.name().local_name().into_inner();
                if local == b"from" {
                    in_from = true;
                } else if local == b"to" {
                    in_to = true;
                } else if local == b"chart"
                    && let Some(rid) = get_attr(&e, b"id")
                {
                    chart_rid = rid;
                }
                current_tag = local.to_vec();
            }
            Ok(quick_xml::events::Event::Empty(e)) => {
                let local = e.name().local_name().into_inner();
                if local == b"chart"
                    && let Some(rid) = get_attr(&e, b"id")
                {
                    chart_rid = rid;
                }
            }
            Ok(quick_xml::events::Event::End(e)) => {
                let local = e.name().local_name().into_inner();
                if local == b"from" {
                    in_from = false;
                } else if local == b"to" {
                    in_to = false;
                } else if local == b"twoCellAnchor" {
                    if !chart_rid.is_empty() {
                        anchors.push(ParsedAnchor {
                            from_col,
                            from_row,
                            from_col_off,
                            from_row_off,
                            to_col,
                            to_row,
                            to_col_off,
                            to_row_off,
                            chart_rid: chart_rid.clone(),
                        });
                    }
                    chart_rid.clear();
                }
                current_tag.clear();
            }
            Ok(quick_xml::events::Event::Text(e)) => {
                if let Ok(decoded) = e.decode()
                    && let Ok(text) = quick_xml::escape::unescape(&decoded)
                {
                    let val_str = text.trim();
                    if in_from {
                        if current_tag == b"col" {
                            from_col = val_str.parse().unwrap_or(0);
                        } else if current_tag == b"row" {
                            from_row = val_str.parse().unwrap_or(0);
                        } else if current_tag == b"colOff" {
                            let emu: f32 = val_str.parse().unwrap_or(0.0);
                            from_col_off = emu / 9525.0;
                        } else if current_tag == b"rowOff" {
                            let emu: f32 = val_str.parse().unwrap_or(0.0);
                            from_row_off = emu / 9525.0;
                        }
                    } else if in_to {
                        if current_tag == b"col" {
                            to_col = val_str.parse().unwrap_or(0);
                        } else if current_tag == b"row" {
                            to_row = val_str.parse().unwrap_or(0);
                        } else if current_tag == b"colOff" {
                            let emu: f32 = val_str.parse().unwrap_or(0.0);
                            to_col_off = emu / 9525.0;
                        } else if current_tag == b"rowOff" {
                            let emu: f32 = val_str.parse().unwrap_or(0.0);
                            to_row_off = emu / 9525.0;
                        }
                    }
                }
            }
            _ => {}
        }
        buf.clear();
    }
    anchors
}

fn parse_chart_xml(xml: &str) -> Option<ParsedChartInfo> {
    let mut reader = quick_xml::reader::Reader::from_str(xml);
    let mut buf = Vec::new();

    let mut chart_type = crate::core::chart::ChartType::Column;
    let mut title = None;
    let mut xlabel = None;
    let mut show_legend = false;

    let mut cat_f = String::new();
    let mut val_f = String::new();

    let mut current_tag = Vec::new();

    let mut in_title = false;
    let mut in_cat_ax = false;
    let mut in_val_ax = false;
    let mut in_ser = false;
    let mut in_cat = false;
    let mut in_val = false;

    let mut saw_cat_ax = false;
    let mut val_ax_titles: Vec<String> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Eof) => break,
            Ok(quick_xml::events::Event::Start(e)) => {
                let local = e.name().local_name().into_inner();
                if local == b"lineChart" {
                    chart_type = crate::core::chart::ChartType::Line;
                } else if local == b"barChart" {
                    chart_type = crate::core::chart::ChartType::Bar;
                } else if local == b"colChart" {
                    chart_type = crate::core::chart::ChartType::Column;
                } else if local == b"pieChart" {
                    chart_type = crate::core::chart::ChartType::Pie;
                } else if local == b"scatterChart" {
                    chart_type = crate::core::chart::ChartType::Scatter;
                } else if local == b"areaChart" {
                    chart_type = crate::core::chart::ChartType::Area;
                } else if local == b"title" {
                    in_title = true;
                } else if local == b"catAx" {
                    in_cat_ax = true;
                    saw_cat_ax = true;
                } else if local == b"valAx" {
                    in_val_ax = true;
                } else if local == b"ser" {
                    in_ser = true;
                } else if local == b"cat" || local == b"xVal" {
                    in_cat = true;
                } else if local == b"val" || local == b"yVal" {
                    in_val = true;
                } else if local == b"legend" {
                    show_legend = true;
                }
                current_tag = local.to_vec();
            }
            Ok(quick_xml::events::Event::Empty(e)) => {
                let local = e.name().local_name().into_inner();
                if local == b"legend" {
                    show_legend = true;
                } else if local == b"barDir"
                    && let Some(val) = get_attr(&e, b"val")
                {
                    if val == "col" {
                        chart_type = crate::core::chart::ChartType::Column;
                    } else if val == "bar" {
                        chart_type = crate::core::chart::ChartType::Bar;
                    }
                }
            }
            Ok(quick_xml::events::Event::End(e)) => {
                let local = e.name().local_name().into_inner();
                if local == b"title" {
                    in_title = false;
                } else if local == b"catAx" {
                    in_cat_ax = false;
                } else if local == b"valAx" {
                    in_val_ax = false;
                } else if local == b"ser" {
                    in_ser = false;
                } else if local == b"cat" || local == b"xVal" {
                    in_cat = false;
                } else if local == b"val" || local == b"yVal" {
                    in_val = false;
                }
                current_tag.clear();
            }
            Ok(quick_xml::events::Event::Text(e)) => {
                if let Ok(decoded) = e.decode()
                    && let Ok(text) = quick_xml::escape::unescape(&decoded)
                {
                    let val_str = text.trim();
                    if !val_str.is_empty() {
                        if in_title {
                            if current_tag == b"v" || current_tag == b"t" {
                                if in_cat_ax {
                                    xlabel = Some(val_str.to_string());
                                } else if in_val_ax {
                                    val_ax_titles.push(val_str.to_string());
                                } else {
                                    title = Some(val_str.to_string());
                                }
                            }
                        } else if in_ser && current_tag == b"f" {
                            if in_cat {
                                cat_f = val_str.to_string();
                            } else if in_val {
                                val_f = val_str.to_string();
                            }
                        }
                    }
                }
            }
            _ => {}
        }
        buf.clear();
    }

    let data_range = if !cat_f.is_empty() && !val_f.is_empty() {
        combine_ranges(&cat_f, &val_f)
    } else if !val_f.is_empty() {
        clean_excel_formula(&val_f)
    } else if !cat_f.is_empty() {
        clean_excel_formula(&cat_f)
    } else {
        String::new()
    };

    let mut val_ax_titles = val_ax_titles.into_iter();
    let ylabel = if saw_cat_ax {
        val_ax_titles.next()
    } else {
        if let Some(first) = val_ax_titles.next() {
            xlabel = Some(first);
        }
        val_ax_titles.next()
    };

    Some(ParsedChartInfo {
        chart_type,
        title,
        xlabel,
        ylabel,
        show_legend,
        data_range,
    })
}

pub(crate) fn get_zip_file_content(
    archive: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>,
    name: &str,
) -> Option<String> {
    let mut file = archive.by_name(name).ok()?;
    let mut content = String::new();
    std::io::Read::read_to_string(&mut file, &mut content).ok()?;
    Some(content)
}

struct ParsedTablePart {
    name: String,
    columns: Vec<String>,
    has_header_row: bool,
    has_totals_row: bool,

    bounds: Option<(usize, usize, usize, usize)>,
    style_name: Option<String>,
    has_insert_row: bool,
}

fn parse_table_part_xml(xml: &str) -> Option<ParsedTablePart> {
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buf = Vec::new();
    let mut name = None;
    let mut columns = Vec::new();
    let mut header_row_count = 1usize;
    let mut totals_row_count = 0usize;
    let mut bounds = None;
    let mut style_name = None;
    let mut has_insert_row = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(ref e))
            | Ok(quick_xml::events::Event::Empty(ref e)) => {
                let local_name = e.local_name();
                match local_name.as_ref() {
                    b"table" => {
                        name = get_attr(e, b"displayName");
                        header_row_count = get_attr(e, b"headerRowCount")
                            .and_then(|s| s.parse::<usize>().ok())
                            .unwrap_or(1);
                        totals_row_count = get_attr(e, b"totalsRowCount")
                            .and_then(|s| s.parse::<usize>().ok())
                            .unwrap_or(0);
                        has_insert_row = get_attr(e, b"insertRow")
                            .is_some_and(|s| s == "1" || s.eq_ignore_ascii_case("true"));
                        bounds = get_attr(e, b"ref")
                            .and_then(|s| crate::core::pivot_xlsx::parse_a1_range(&s));
                        if let Some(s) = get_attr(e, b"styleName") {
                            style_name = Some(s);
                        }
                    }
                    b"tableColumn" => {
                        if let Some(col) = get_attr(e, b"name") {
                            columns.push(col);
                        }
                    }
                    b"tableStyleInfo" => {
                        if let Some(s) = get_attr(e, b"name") {
                            style_name = Some(s);
                        }
                    }
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    Some(ParsedTablePart {
        name: name?,
        columns,
        has_header_row: header_row_count != 0,
        has_totals_row: totals_row_count != 0,
        bounds,
        style_name,
        has_insert_row,
    })
}

fn import_tables_from_opc(pkg: &OpcPackage) -> Result<Vec<(String, ParsedTablePart)>, String> {
    let sheet_part_to_name = build_sheet_part_to_name(pkg)?;

    let mut sheet_parts: Vec<String> = sheet_part_to_name.keys().cloned().collect();
    sheet_parts.sort();

    let mut tables = Vec::new();
    for sheet_part in sheet_parts {
        let Some(sheet_name) = sheet_part_to_name.get(&sheet_part).cloned() else {
            continue;
        };
        for rel in pkg.read_relationships(&sheet_part) {
            if !rel.is_external() && crate::core::opc::rel_type_matches(&rel.rel_type, "table") {
                let table_path = rel.resolved_target(&sheet_part);
                if let Some(table_xml) = pkg.get_str(&table_path)
                    && let Some(parsed) = parse_table_part_xml(&table_xml)
                {
                    tables.push((sheet_name.clone(), parsed));
                }
            }
        }
    }
    Ok(tables)
}

const BUILTIN_NUM_FMTS: &[(u32, &str)] = &[
    (0, "General"),
    (1, "0"),
    (2, "0.00"),
    (3, "#,##0"),
    (4, "#,##0.00"),
    (5, "\"$\"#,##0_);(\"$\"#,##0)"),
    (6, "\"$\"#,##0_);[Red](\"$\"#,##0)"),
    (7, "\"$\"#,##0.00_);(\"$\"#,##0.00)"),
    (8, "\"$\"#,##0.00_);[Red](\"$\"#,##0.00)"),
    (9, "0%"),
    (10, "0.00%"),
    (11, "0.00E+00"),
    (12, "# ?/?"),
    (13, "# ??/??"),
    (14, "m/d/yy"),
    (15, "d-mmm-yy"),
    (16, "d-mmm"),
    (17, "mmm-yy"),
    (18, "h:mm AM/PM"),
    (19, "h:mm:ss AM/PM"),
    (20, "h:mm"),
    (21, "h:mm:ss"),
    (22, "m/d/yy h:mm"),
    (37, "#,##0;(#,##0)"),
    (38, "#,##0;[Red](#,##0)"),
    (39, "#,##0.00;(#,##0.00)"),
    (40, "#,##0.00;[Red](#,##0.00)"),
    (41, r#"_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)"#),
    (42, r#"_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)"#),
    (43, r#"_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)"#),
    (
        44,
        r#"_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)"#,
    ),
    (45, "mm:ss"),
    (46, "[h]:mm:ss"),
    (47, "mm:ss.0"),
    (48, "##0.0E+0"),
    (49, "@"),
];

fn builtin_num_fmt_code(id: u32) -> Option<&'static str> {
    BUILTIN_NUM_FMTS
        .iter()
        .find(|(builtin_id, _)| *builtin_id == id)
        .map(|(_, code)| *code)
}

type SheetCellNumberFormats =
    std::collections::HashMap<String, std::collections::HashMap<(usize, usize), String>>;

type SheetCellStyles =
    std::collections::HashMap<String, std::collections::HashMap<(usize, usize), CellStyle>>;

type SheetColumnWidths = std::collections::HashMap<String, std::collections::HashMap<usize, f64>>;

type SheetRowHeights = std::collections::HashMap<String, std::collections::HashMap<usize, f64>>;

fn import_column_widths(pkg: &OpcPackage) -> Result<SheetColumnWidths, String> {
    use std::collections::HashMap;

    let sheet_part_to_name = build_sheet_part_to_name(pkg)?;
    let mut sheet_parts: Vec<String> = sheet_part_to_name.keys().cloned().collect();
    sheet_parts.sort();

    let mut out: SheetColumnWidths = HashMap::new();
    for sheet_part in sheet_parts {
        let Some(sheet_name) = sheet_part_to_name.get(&sheet_part).cloned() else {
            continue;
        };
        let Some(sheet_xml) = pkg.get_str(&sheet_part) else {
            continue;
        };
        let widths = parse_sheet_column_widths(&sheet_xml);
        if !widths.is_empty() {
            out.insert(sheet_name, widths);
        }
    }
    Ok(out)
}

fn parse_sheet_column_widths(xml: &str) -> std::collections::HashMap<usize, f64> {
    let mut out = std::collections::HashMap::new();
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(ref e))
            | Ok(quick_xml::events::Event::Empty(ref e)) => {
                if e.local_name().as_ref() == b"col"
                    && let Some(width) = get_attr(e, b"width").and_then(|s| s.parse::<f64>().ok())
                    && width.is_finite()
                    && width >= 0.0
                {
                    let min = get_attr(e, b"min")
                        .and_then(|s| s.parse::<usize>().ok())
                        .unwrap_or(1)
                        .max(1);
                    let max = get_attr(e, b"max")
                        .and_then(|s| s.parse::<usize>().ok())
                        .unwrap_or(min)
                        .max(min)
                        .min(16_384);
                    for one_based_col in min..=max {
                        out.insert(one_based_col - 1, width);
                    }
                }
            }
            Ok(quick_xml::events::Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    out
}

fn import_row_heights(pkg: &OpcPackage) -> Result<SheetRowHeights, String> {
    use std::collections::HashMap;

    let sheet_part_to_name = build_sheet_part_to_name(pkg)?;
    let mut sheet_parts: Vec<String> = sheet_part_to_name.keys().cloned().collect();
    sheet_parts.sort();

    let mut out: SheetRowHeights = HashMap::new();
    for sheet_part in sheet_parts {
        let Some(sheet_name) = sheet_part_to_name.get(&sheet_part).cloned() else {
            continue;
        };
        let Some(sheet_xml) = pkg.get_str(&sheet_part) else {
            continue;
        };
        let heights = parse_sheet_row_heights(&sheet_xml);
        if !heights.is_empty() {
            out.insert(sheet_name, heights);
        }
    }
    Ok(out)
}

fn parse_sheet_row_heights(xml: &str) -> std::collections::HashMap<usize, f64> {
    let mut out = std::collections::HashMap::new();
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buf = Vec::new();
    let mut next_row = 1usize;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(ref e))
            | Ok(quick_xml::events::Event::Empty(ref e)) => {
                if e.local_name().as_ref() == b"row" {
                    let row = get_attr(e, b"r")
                        .and_then(|s| s.parse::<usize>().ok())
                        .unwrap_or(next_row)
                        .max(1);
                    next_row = row + 1;
                    if let Some(height) = get_attr(e, b"ht").and_then(|s| s.parse::<f64>().ok())
                        && row <= 1_048_576
                        && height.is_finite()
                        && height >= 0.0
                    {
                        out.insert(row - 1, height);
                    }
                }
            }
            Ok(quick_xml::events::Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    out
}

fn import_cell_styles(pkg: &OpcPackage) -> Result<SheetCellStyles, String> {
    use std::collections::HashMap;

    let Some(styles_xml) = pkg.get_str("xl/styles.xml") else {
        return Ok(HashMap::new());
    };
    let xf_to_style = parse_styles_cell_styles(&styles_xml);
    if xf_to_style.is_empty() {
        return Ok(HashMap::new());
    }

    let sheet_part_to_name = build_sheet_part_to_name(pkg)?;
    let mut sheet_parts: Vec<String> = sheet_part_to_name.keys().cloned().collect();
    sheet_parts.sort();

    let mut out: HashMap<String, HashMap<(usize, usize), CellStyle>> = HashMap::new();
    for sheet_part in sheet_parts {
        let Some(sheet_name) = sheet_part_to_name.get(&sheet_part).cloned() else {
            continue;
        };
        let Some(sheet_xml) = pkg.get_str(&sheet_part) else {
            continue;
        };
        let cells = parse_sheet_cell_styles(&sheet_xml, &xf_to_style);
        if !cells.is_empty() {
            out.insert(sheet_name, cells);
        }
    }
    Ok(out)
}

fn parse_styles_cell_styles(xml: &str) -> std::collections::HashMap<u32, CellStyle> {
    use std::collections::HashMap;

    let mut custom_num_formats: HashMap<u32, String> = HashMap::new();
    let mut fonts: Vec<CellStyle> = Vec::new();
    let mut fills: Vec<Option<String>> = Vec::new();
    let mut out: HashMap<u32, CellStyle> = HashMap::new();
    let mut in_fonts = false;
    let mut in_font = false;
    let mut current_font = CellStyle::default();
    let mut in_fills = false;
    let mut in_fill = false;
    let mut current_fill: Option<String> = None;
    let mut in_cell_xfs = false;
    let mut xf_idx = 0u32;

    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(ref e))
            | Ok(quick_xml::events::Event::Empty(ref e)) => match e.local_name().as_ref() {
                b"numFmt" => {
                    if let Some(id) = get_attr(e, b"numFmtId").and_then(|s| s.parse::<u32>().ok())
                        && let Some(code) = get_attr(e, b"formatCode")
                    {
                        custom_num_formats.insert(id, code);
                    }
                }
                b"fonts" => in_fonts = true,
                b"font" if in_fonts => {
                    in_font = true;
                    current_font = CellStyle::default();
                }
                b"b" if in_font => {
                    if style_bool_enabled(e) {
                        current_font.bold = Some(true);
                    }
                }
                b"i" if in_font => {
                    if style_bool_enabled(e) {
                        current_font.italic = Some(true);
                    }
                }
                b"u" if in_font => {
                    if style_bool_enabled(e) {
                        current_font.underline = Some(true);
                    }
                }
                b"color" if in_font => current_font.font_color = style_color_hex(e),
                b"name" if in_font => current_font.font_family = get_attr(e, b"val"),
                b"sz" if in_font => {
                    current_font.font_size = get_attr(e, b"val").and_then(|s| s.parse().ok());
                }
                b"fills" => in_fills = true,
                b"fill" if in_fills => {
                    in_fill = true;
                    current_fill = None;
                }
                b"fgColor" if in_fill => {
                    if let Some(color) = style_color_hex(e) {
                        current_fill = Some(color);
                    }
                }
                b"bgColor" if in_fill => {
                    if current_fill.is_none() {
                        current_fill = style_color_hex(e);
                    }
                }
                b"cellXfs" => in_cell_xfs = true,
                b"xf" if in_cell_xfs => {
                    let mut style = CellStyle::default();
                    if let Some(font_id) =
                        get_attr(e, b"fontId").and_then(|s| s.parse::<usize>().ok())
                        && font_id != 0
                        && let Some(font) = fonts.get(font_id)
                    {
                        style.merge(font);
                    }
                    if let Some(fill_id) =
                        get_attr(e, b"fillId").and_then(|s| s.parse::<usize>().ok())
                        && fill_id > 1
                        && let Some(Some(bg_color)) = fills.get(fill_id)
                    {
                        style.bg_color = Some(bg_color.clone());
                    }
                    if let Some(num_fmt_id) =
                        get_attr(e, b"numFmtId").and_then(|s| s.parse::<u32>().ok())
                    {
                        style.num_format = custom_num_formats
                            .get(&num_fmt_id)
                            .cloned()
                            .or_else(|| builtin_num_fmt_code(num_fmt_id).map(str::to_string))
                            .filter(|code| !code.eq_ignore_ascii_case("General"));
                    }
                    if !style.is_empty() {
                        out.insert(xf_idx, style);
                    }
                    xf_idx += 1;
                }
                _ => {}
            },
            Ok(quick_xml::events::Event::End(ref e)) => match e.local_name().as_ref() {
                b"fonts" => in_fonts = false,
                b"font" if in_font => {
                    fonts.push(current_font.clone());
                    current_font = CellStyle::default();
                    in_font = false;
                }
                b"fills" => in_fills = false,
                b"fill" if in_fill => {
                    fills.push(current_fill.clone());
                    current_fill = None;
                    in_fill = false;
                }
                b"cellXfs" => in_cell_xfs = false,
                _ => {}
            },
            Ok(quick_xml::events::Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    out
}

fn style_bool_enabled(e: &quick_xml::events::BytesStart) -> bool {
    get_attr(e, b"val").is_none_or(|value| {
        value != "0" && !value.eq_ignore_ascii_case("false") && !value.eq_ignore_ascii_case("none")
    })
}

fn style_color_hex(e: &quick_xml::events::BytesStart) -> Option<String> {
    if let Some(rgb) = get_attr(e, b"rgb") {
        let trimmed = rgb.trim().trim_start_matches('#');
        let hex = match trimmed.len() {
            8 => &trimmed[2..],
            6 => trimmed,
            _ => return None,
        };
        return Some(format!("#{}", hex.to_ascii_uppercase()));
    }
    get_attr(e, b"indexed")
        .and_then(|idx| idx.parse::<u32>().ok())
        .and_then(indexed_color_hex)
}

fn indexed_color_hex(index: u32) -> Option<String> {
    let hex = match index {
        0 | 8 => "#000000",
        1 | 9 => "#FFFFFF",
        2 | 10 => "#FF0000",
        3 | 11 => "#00FF00",
        4 | 12 => "#0000FF",
        5 | 13 => "#FFFF00",
        6 | 14 => "#FF00FF",
        7 | 15 => "#00FFFF",
        _ => return None,
    };
    Some(hex.to_string())
}

fn parse_sheet_cell_styles(
    xml: &str,
    xf_to_style: &std::collections::HashMap<u32, CellStyle>,
) -> std::collections::HashMap<(usize, usize), CellStyle> {
    let mut out = std::collections::HashMap::new();
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(ref e))
            | Ok(quick_xml::events::Event::Empty(ref e)) => {
                if e.local_name().as_ref() == b"c"
                    && let Some(style_idx) = get_attr(e, b"s").and_then(|s| s.parse::<u32>().ok())
                    && let Some(style) = xf_to_style.get(&style_idx)
                    && let Some(reference) = get_attr(e, b"r")
                    && let Some(rc) = parse_a1_cell(&reference)
                {
                    out.insert(rc, style.clone());
                }
            }
            Ok(quick_xml::events::Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    out
}

#[allow(dead_code)]
fn import_cell_number_formats(buffer: &[u8]) -> Result<SheetCellNumberFormats, String> {
    use std::collections::HashMap;

    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(buffer))
        .map_err(|e| format!("Failed to open zip: {}", e))?;

    let Some(styles_xml) = get_zip_file_content(&mut archive, "xl/styles.xml") else {
        return Ok(HashMap::new());
    };
    let xf_to_code = parse_styles_num_formats(&styles_xml);
    if xf_to_code.is_empty() {
        return Ok(HashMap::new());
    }

    let sheet_file_to_name = build_sheet_file_to_name(&mut archive)?;
    let mut sheet_files: Vec<String> = sheet_file_to_name.keys().cloned().collect();
    sheet_files.sort();

    let mut out: HashMap<String, HashMap<(usize, usize), String>> = HashMap::new();
    for sheet_file in sheet_files {
        let Some(sheet_name) = sheet_file_to_name.get(&sheet_file).cloned() else {
            continue;
        };
        let path = format!("xl/worksheets/{}", sheet_file);
        let Some(sheet_xml) = get_zip_file_content(&mut archive, &path) else {
            continue;
        };
        let cells = parse_sheet_cell_formats(&sheet_xml, &xf_to_code);
        if !cells.is_empty() {
            out.insert(sheet_name, cells);
        }
    }
    Ok(out)
}

fn parse_styles_num_formats(xml: &str) -> std::collections::HashMap<u32, String> {
    use std::collections::HashMap;

    let mut custom: HashMap<u32, String> = HashMap::new();
    let mut cell_xfs: Vec<u32> = Vec::new();
    let mut in_cell_xfs = false;

    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(ref e))
            | Ok(quick_xml::events::Event::Empty(ref e)) => match e.local_name().as_ref() {
                b"numFmt" => {
                    if let Some(id) = get_attr(e, b"numFmtId").and_then(|s| s.parse::<u32>().ok())
                        && let Some(code) = get_attr(e, b"formatCode")
                    {
                        custom.insert(id, code);
                    }
                }
                b"cellXfs" => in_cell_xfs = true,
                b"xf" if in_cell_xfs => {
                    cell_xfs.push(
                        get_attr(e, b"numFmtId")
                            .and_then(|s| s.parse::<u32>().ok())
                            .unwrap_or(0),
                    );
                }
                _ => {}
            },
            Ok(quick_xml::events::Event::End(ref e)) => {
                if e.local_name().as_ref() == b"cellXfs" {
                    in_cell_xfs = false;
                }
            }
            Ok(quick_xml::events::Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    let mut out = HashMap::new();
    for (xf_idx, num_fmt_id) in cell_xfs.into_iter().enumerate() {
        let code = custom
            .get(&num_fmt_id)
            .cloned()
            .or_else(|| builtin_num_fmt_code(num_fmt_id).map(str::to_string));
        if let Some(code) = code
            && crate::core::date::is_date_code(&code)
        {
            out.insert(xf_idx as u32, code);
        }
    }
    out
}

fn parse_sheet_cell_formats(
    xml: &str,
    xf_to_code: &std::collections::HashMap<u32, String>,
) -> std::collections::HashMap<(usize, usize), String> {
    let mut out = std::collections::HashMap::new();
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(ref e))
            | Ok(quick_xml::events::Event::Empty(ref e)) => {
                if e.local_name().as_ref() == b"c"
                    && let Some(style_idx) = get_attr(e, b"s").and_then(|s| s.parse::<u32>().ok())
                    && let Some(code) = xf_to_code.get(&style_idx)
                    && let Some(reference) = get_attr(e, b"r")
                    && let Some(rc) = parse_a1_cell(&reference)
                {
                    out.insert(rc, code.clone());
                }
            }
            Ok(quick_xml::events::Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    out
}

fn parse_a1_cell(reference: &str) -> Option<(usize, usize)> {
    let split = reference.find(|c: char| c.is_ascii_digit())?;
    let (letters, digits) = reference.split_at(split);
    if letters.is_empty() || !letters.chars().all(|c| c.is_ascii_alphabetic()) {
        return None;
    }
    let mut col = 0usize;
    for c in letters.chars() {
        col = col * 26 + (c.to_ascii_uppercase() as usize - 'A' as usize + 1);
    }
    let row = digits.parse::<usize>().ok()?;
    if row == 0 {
        return None;
    }
    Some((row - 1, col - 1))
}

fn deterministic_chart_id(sheet_name: &str, index_in_sheet: usize) -> u64 {
    use std::collections::hash_map::DefaultHasher;
    use std::hash::{Hash, Hasher};
    let mut hasher = DefaultHasher::new();
    sheet_name.hash(&mut hasher);
    index_in_sheet.hash(&mut hasher);
    hasher.finish() & 0x001F_FFFF_FFFF_FFFF
}

fn build_sheet_file_to_name(
    archive: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>,
) -> Result<std::collections::HashMap<String, String>, String> {
    let workbook_xml = get_zip_file_content(archive, "xl/workbook.xml")
        .ok_or_else(|| "Missing xl/workbook.xml".to_string())?;
    let r_id_to_name = parse_workbook_sheets(&workbook_xml);

    let workbook_rels_xml = get_zip_file_content(archive, "xl/_rels/workbook.xml.rels")
        .ok_or_else(|| "Missing xl/_rels/workbook.xml.rels".to_string())?;
    let r_id_to_target = parse_workbook_rels(&workbook_rels_xml);

    let mut sheet_file_to_name = std::collections::HashMap::new();
    for (r_id, name) in r_id_to_name {
        if let Some(target) = r_id_to_target.get(&r_id) {
            sheet_file_to_name.insert(get_filename(target), name.clone());
        }
    }
    Ok(sheet_file_to_name)
}

fn build_sheet_part_to_name(
    pkg: &OpcPackage,
) -> Result<std::collections::HashMap<String, String>, String> {
    let workbook_part = pkg
        .find_workbook_part()
        .ok_or_else(|| "Missing xl/workbook.xml".to_string())?;
    let workbook_xml = pkg
        .get_str(&workbook_part)
        .ok_or_else(|| format!("Missing {workbook_part}"))?;
    let r_id_to_name = parse_workbook_sheets(&workbook_xml);
    let rels = pkg.read_relationships(&workbook_part);

    let mut sheet_part_to_name = std::collections::HashMap::new();
    for rel in rels {
        if !rel.is_external()
            && let Some(name) = r_id_to_name.get(&rel.id)
        {
            sheet_part_to_name.insert(rel.resolved_target(&workbook_part), name.clone());
        }
    }
    Ok(sheet_part_to_name)
}

fn import_charts_from_opc(pkg: &OpcPackage) -> Result<Vec<ParsedChartData>, String> {
    let sheet_part_to_name = build_sheet_part_to_name(pkg)?;
    let mut parsed_charts = Vec::new();

    for (sheet_part, sheet_name) in &sheet_part_to_name {
        for rel in pkg.read_relationships(sheet_part) {
            if !rel.is_external() && crate::core::opc::rel_type_matches(&rel.rel_type, "drawing") {
                let drawing_part = rel.resolved_target(sheet_part);
                let chart_rid_to_part: std::collections::HashMap<String, String> = pkg
                    .read_relationships(&drawing_part)
                    .into_iter()
                    .filter(|r| {
                        !r.is_external() && crate::core::opc::rel_type_matches(&r.rel_type, "chart")
                    })
                    .map(|r| {
                        let target = r.resolved_target(&drawing_part);
                        (r.id, target)
                    })
                    .collect();

                if let Some(drawing_xml) = pkg.get_str(&drawing_part) {
                    let anchors = parse_drawings_xml(&drawing_xml);
                    for anchor in anchors {
                        if let Some(chart_part) = chart_rid_to_part.get(&anchor.chart_rid)
                            && let Some(chart_xml) = pkg.get_str(chart_part)
                            && let Some(info) = parse_chart_xml(&chart_xml)
                        {
                            parsed_charts.push(ParsedChartData {
                                sheet_name: sheet_name.clone(),
                                anchor,
                                info,
                            });
                        }
                    }
                }
            }
        }
    }

    Ok(parsed_charts)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_xlsx_import_export_cycle() {
        let mut columns = Vec::new();
        let col1 = DataColumn::from_src("A", vec!["10".to_string(), "20".to_string()]);
        columns.push(col1);

        let col2 = DataColumn::from_src("B", vec!["=A1 + A2".to_string(), "abc".to_string()]);
        columns.push(col2);

        let sheet = Sheet {
            id: 1,
            name: "Sheet1".to_string(),
            columns,
            row_heights: Vec::new(),
            hidden_rows: Vec::new(),
            merged_cells: Vec::new(),
            visibility: crate::core::SheetVisibility::Visible,
            freeze_pane: None,
            date1904: false,
            tables: Vec::new(),
            dependencies: std::collections::HashMap::new(),
            dependencies_rev: std::collections::HashMap::new(),
            uncommitted_actions: Vec::new(),
            locale: crate::core::locale::Locale::default(),
        };

        let xlsx_data = export_xlsx_data(&[sheet], &[], &[], None).unwrap();
        assert!(!xlsx_data.is_empty());

        let (imported_sheets, imported_charts, _, _) =
            import_xlsx_data(&xlsx_data, &[], |_, _, _| {}).unwrap();
        assert_eq!(imported_sheets.len(), 1);
        assert_eq!(imported_charts.len(), 0);

        let imported_sheet = &imported_sheets[0].sheet;
        assert_eq!(imported_sheet.name, "Sheet1");
        assert_eq!(imported_sheet.columns.len(), 2);

        assert_eq!(imported_sheet.columns[0].src[0], "10");
        assert_eq!(imported_sheet.columns[0].src[1], "20");
        assert_eq!(imported_sheet.columns[1].src[0], "=A1 + A2");
        assert_eq!(imported_sheet.columns[1].src[1], "abc");
    }

    #[test]
    fn test_xlsx_column_width_survives_round_trip() {
        let mut sheet = Sheet::new(crate::core::SheetInit {
            name: Some("Sheet1".to_string()),
            rows: 1,
            cols: 3,
            ..Default::default()
        });
        sheet.set_column_width(1, Some(22.5));
        sheet.set_cell_src(0, 1, "wide".to_string());

        let xlsx_data = export_xlsx_data(&[sheet], &[], &[], None).unwrap();
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(xlsx_data.as_slice())).unwrap();
        let sheet_xml = get_zip_file_content(&mut archive, "xl/worksheets/sheet1.xml").unwrap();
        assert!(sheet_xml.contains("<cols>"));
        assert!(sheet_xml.contains("min=\"2\""));
        assert!(sheet_xml.contains("max=\"2\""));
        assert!(sheet_xml.contains("width=\"22.5\""));

        let (imported_sheets, _, _, _) = import_xlsx_data(&xlsx_data, &[], |_, _, _| {}).unwrap();
        assert_eq!(imported_sheets[0].sheet.get_column_width(1), Some(22.5));

        let exported_again =
            export_xlsx_data(&[imported_sheets[0].sheet.clone()], &[], &[], None).unwrap();
        let mut archive =
            zip::ZipArchive::new(std::io::Cursor::new(exported_again.as_slice())).unwrap();
        let sheet_xml = get_zip_file_content(&mut archive, "xl/worksheets/sheet1.xml").unwrap();
        assert!(sheet_xml.contains("width=\"22.5\""));
    }

    #[test]
    fn test_xlsx_row_height_survives_round_trip() {
        let mut sheet = Sheet::new(crate::core::SheetInit {
            name: Some("Sheet1".to_string()),
            rows: 3,
            cols: 1,
            ..Default::default()
        });
        sheet.set_row_height(1, Some(31.5));
        sheet.set_cell_src(1, 0, "tall".to_string());

        let xlsx_data = export_xlsx_data(&[sheet], &[], &[], None).unwrap();
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(xlsx_data.as_slice())).unwrap();
        let sheet_xml = get_zip_file_content(&mut archive, "xl/worksheets/sheet1.xml").unwrap();
        assert!(sheet_xml.contains("r=\"2\""));
        assert!(sheet_xml.contains("ht=\"31.5\""));
        assert!(sheet_xml.contains("customHeight=\"1\""));

        let (imported_sheets, _, _, _) = import_xlsx_data(&xlsx_data, &[], |_, _, _| {}).unwrap();
        assert_eq!(imported_sheets[0].sheet.get_row_height(1), Some(31.5));

        let exported_again =
            export_xlsx_data(&[imported_sheets[0].sheet.clone()], &[], &[], None).unwrap();
        let mut archive =
            zip::ZipArchive::new(std::io::Cursor::new(exported_again.as_slice())).unwrap();
        let sheet_xml = get_zip_file_content(&mut archive, "xl/worksheets/sheet1.xml").unwrap();
        assert!(sheet_xml.contains("ht=\"31.5\""));
    }

    #[test]
    fn test_xlsx_cell_style_survives_load_evaluate_save_cycle() {
        let mut sheet = Sheet::new(crate::core::SheetInit {
            name: Some("Sheet1".to_string()),
            rows: 1,
            cols: 3,
            ..Default::default()
        });
        sheet.set_cell_src(0, 0, "5".to_string());
        sheet.set_cell_src(0, 1, "=A1*2".to_string());
        sheet.columns[0].styles[0] = Some(CellStyle {
            font_color: Some("#FF0000".to_string()),
            bg_color: Some("#FFFF00".to_string()),
            bold: Some(true),
            italic: Some(true),
            underline: Some(true),
            font_family: Some("Arial".to_string()),
            font_size: Some(14.0),
            ..Default::default()
        });
        sheet.columns[1].styles[0] = Some(CellStyle {
            font_color: Some("#0000FF".to_string()),
            bg_color: Some("#00FF00".to_string()),
            bold: Some(true),
            ..Default::default()
        });
        sheet.columns[2].styles[0] = Some(CellStyle {
            bg_color: Some("#00FFFF".to_string()),
            ..Default::default()
        });
        sheet.commit(None).unwrap();

        let bytes = export_xlsx_data(&[sheet], &[], &[], None).unwrap();
        let mut wb = crate::core::WorkbookManager::load_bytes(&bytes).unwrap();
        wb.evaluate().unwrap();
        let evaluated = wb.save_bytes().unwrap();
        let reloaded = crate::core::WorkbookManager::load_bytes(&evaluated).unwrap();

        let style_a1 = reloaded.sheets[0].columns[0].styles[0].as_ref().unwrap();
        assert_eq!(style_a1.font_color.as_deref(), Some("#FF0000"));
        assert_eq!(style_a1.bg_color.as_deref(), Some("#FFFF00"));
        assert_eq!(style_a1.bold, Some(true));
        assert_eq!(style_a1.italic, Some(true));
        assert_eq!(style_a1.underline, Some(true));
        assert_eq!(style_a1.font_family.as_deref(), Some("Arial"));
        assert_eq!(style_a1.font_size, Some(14.0));

        let style_b1 = reloaded.sheets[0].columns[1].styles[0].as_ref().unwrap();
        assert_eq!(style_b1.font_color.as_deref(), Some("#0000FF"));
        assert_eq!(style_b1.bg_color.as_deref(), Some("#00FF00"));
        assert_eq!(style_b1.bold, Some(true));

        let style_c1 = reloaded.sheets[0].columns[2].styles[0].as_ref().unwrap();
        assert_eq!(style_c1.bg_color.as_deref(), Some("#00FFFF"));
        assert!(reloaded.sheets[0].columns[2].src[0].is_empty());
    }

    #[test]
    fn test_xlsx_custom_number_format_survives_round_trip_and_display() {
        let mut sheet = Sheet::new(crate::core::SheetInit {
            name: Some("Sheet1".to_string()),
            rows: 1,
            cols: 1,
            ..Default::default()
        });
        sheet.set_cell_src(0, 0, "7.5".to_string());
        sheet.update_cell_style(0, 0, |style| {
            style.num_format = Some("0.0\"x\"".to_string());
        });
        sheet.commit(None).unwrap();

        let bytes = export_xlsx_data(&[sheet], &[], &[], None).unwrap();
        let (imported, _, _, _) = import_xlsx_data(&bytes, &[], |_, _, _| {}).unwrap();
        let imported = &imported[0].sheet;
        assert_eq!(
            imported.columns[0].styles[0]
                .as_ref()
                .and_then(|s| s.num_format.as_deref()),
            Some("0.0\"x\"")
        );
        assert_eq!(
            imported.get_display_string(&crate::core::CellRef::new(0, 0)),
            "7.5x"
        );
    }

    #[test]
    fn test_builtin_date_num_fmt_14_uses_two_digit_year() {
        let styles = r#"
            <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <cellXfs count="1">
                <xf numFmtId="14" applyNumberFormat="1"/>
              </cellXfs>
            </styleSheet>
        "#;
        let formats = parse_styles_num_formats(styles);
        assert_eq!(formats.get(&0).map(String::as_str), Some("m/d/yy"));
    }

    #[test]
    fn test_builtin_and_custom_number_formats_import_as_cell_styles() {
        let styles = r#"
            <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <numFmts count="1">
                <numFmt numFmtId="164" formatCode="0.0&quot;x&quot;"/>
              </numFmts>
              <cellXfs count="3">
                <xf numFmtId="0"/>
                <xf numFmtId="2" applyNumberFormat="1"/>
                <xf numFmtId="164" applyNumberFormat="1"/>
              </cellXfs>
            </styleSheet>
        "#;
        let styles = parse_styles_cell_styles(styles);
        assert!(!styles.contains_key(&0));
        assert_eq!(
            styles.get(&1).and_then(|s| s.num_format.as_deref()),
            Some("0.00")
        );
        assert_eq!(
            styles.get(&2).and_then(|s| s.num_format.as_deref()),
            Some("0.0\"x\"")
        );
    }

    #[test]
    fn test_empty_string_cell_stays_text_not_blank() {
        let mut sheet = Sheet::new(crate::core::SheetInit {
            name: Some("Sheet1".to_string()),
            rows: 2,
            cols: 1,
            ..Default::default()
        });
        sheet.set_cell_with_type(0, 0, String::new(), CellType::String);
        sheet.commit(None).unwrap();

        assert!(
            matches!(
                sheet.get_result_data(&crate::core::CellRef::new(0, 0)),
                crate::core::ResultData::String(ref s) if s.is_empty()
            ),
            "an empty string cell must stay an empty *string*, not become blank"
        );
        assert_eq!(
            sheet.get_cell_type(&crate::core::CellRef::new(0, 0)),
            CellType::String
        );
        assert!(matches!(
            sheet.get_result_data(&crate::core::CellRef::new(1, 0)),
            crate::core::ResultData::None
        ));
    }

    #[test]
    fn test_xlsx_date_notation_survives_round_trip() {
        let mut sheet = Sheet::new(crate::core::SheetInit {
            name: Some("Sheet1".to_string()),
            rows: 3,
            cols: 1,
            ..Default::default()
        });
        for (row, src) in ["6/22/26", "22-Jun-2026", "2026-06-22"].iter().enumerate() {
            sheet.set_cell_src(row, 0, src.to_string());
        }
        sheet.commit(None).unwrap();

        let bytes = export_xlsx_data(&[sheet], &[], &[], None).unwrap();
        let (imported, _, _, _) = import_xlsx_data(&bytes, &[], |_, _, _| {}).unwrap();
        let imported = &imported[0].sheet;

        assert_eq!(imported.columns[0].src[0], "46195");
        for (row, want) in ["6/22/26", "22-Jun-2026", "2026-06-22"].iter().enumerate() {
            assert_eq!(
                imported.get_display_string(&crate::core::CellRef::new(row, 0)),
                *want,
                "row {row} lost its date notation"
            );
        }
    }

    #[test]
    fn test_xlsx_date_looking_text_stays_text() {
        let col = DataColumn::from_src("A", vec!["\"22-Jun\"".to_string()]);
        let sheet = Sheet {
            id: 1,
            name: "Sheet1".to_string(),
            columns: vec![col],
            row_heights: Vec::new(),
            hidden_rows: Vec::new(),
            merged_cells: Vec::new(),
            visibility: crate::core::SheetVisibility::Visible,
            freeze_pane: None,
            date1904: false,
            tables: Vec::new(),
            dependencies: std::collections::HashMap::new(),
            dependencies_rev: std::collections::HashMap::new(),
            uncommitted_actions: Vec::new(),
            locale: crate::core::locale::Locale::default(),
        };

        let bytes = export_xlsx_data(&[sheet], &[], &[], None).unwrap();
        let (imported, _, _, _) = import_xlsx_data(&bytes, &[], |_, _, _| {}).unwrap();
        let mut imported = imported.into_iter().next().unwrap().sheet;
        imported.commit(None).unwrap();

        assert!(matches!(
            imported.get_result_data(&crate::core::CellRef::new(0, 0)),
            crate::core::ResultData::String(ref s) if s == "22-Jun"
        ));
    }

    #[test]
    fn test_xlsx_explicit_cell_type_string_round_trip() {
        let mut sheet = Sheet::new(crate::core::SheetInit {
            name: Some("Sheet1".to_string()),
            rows: 3,
            cols: 1,
            ..Default::default()
        });
        sheet.set_cell_with_type(0, 0, "6/22/26".to_string(), CellType::String);
        sheet.set_cell_with_type(1, 0, "123".to_string(), CellType::String);
        sheet.set_cell_with_type(2, 0, "TRUE".to_string(), CellType::String);
        sheet.commit(None).unwrap();

        let bytes = export_xlsx_data(&[sheet], &[], &[], None).unwrap();
        let (imported, _, _, _) = import_xlsx_data(&bytes, &[], |_, _, _| {}).unwrap();
        let mut imported = imported.into_iter().next().unwrap().sheet;
        imported.commit(None).unwrap();

        assert_eq!(
            imported.get_cell_type(&crate::core::CellRef::new(0, 0)),
            CellType::String
        );
        assert!(matches!(
            imported.get_result_data(&crate::core::CellRef::new(0, 0)),
            crate::core::ResultData::String(ref s) if s == "6/22/26"
        ));
        assert_eq!(
            imported.get_cell_type(&crate::core::CellRef::new(1, 0)),
            CellType::String
        );
        assert!(matches!(
            imported.get_result_data(&crate::core::CellRef::new(1, 0)),
            crate::core::ResultData::String(ref s) if s == "123"
        ));
        assert_eq!(
            imported.get_cell_type(&crate::core::CellRef::new(2, 0)),
            CellType::String
        );
        assert!(matches!(
            imported.get_result_data(&crate::core::CellRef::new(2, 0)),
            crate::core::ResultData::String(ref s) if s == "TRUE"
        ));
    }

    #[test]
    fn test_xlsx_exports_date_format_metadata_for_display_dates_only() {
        let mut sheet = crate::core::engine::tests::create_sheet(&[[
            "6/22/26",
            "=A1+1",
            "=YEAR(A1)",
            "\"6/22/26\"",
        ]]);
        sheet.commit(None).unwrap();

        assert_eq!(
            sheet.get_display_string(&crate::core::CellRef::new(0, 0)),
            "6/22/26"
        );
        assert_eq!(
            sheet.get_display_string(&crate::core::CellRef::new(0, 1)),
            "6/23/26"
        );
        assert_eq!(
            sheet.get_display_string(&crate::core::CellRef::new(0, 2)),
            "2026"
        );
        assert_eq!(
            sheet.get_display_string(&crate::core::CellRef::new(0, 3)),
            "6/22/26"
        );

        let bytes = export_xlsx_data(&[sheet], &[], &[], None).unwrap();
        let formats = import_cell_number_formats(&bytes).unwrap();
        let sheet_formats = formats
            .get("sheet1")
            .expect("date formats should be exported");
        assert_eq!(
            sheet_formats.get(&(0, 0)).map(String::as_str),
            Some("m/d/yy")
        );
        assert_eq!(
            sheet_formats.get(&(0, 1)).map(String::as_str),
            Some("m/d/yy")
        );
        assert!(
            !sheet_formats.contains_key(&(0, 2)),
            "date component formulas should export as ordinary numbers"
        );
        assert!(
            !sheet_formats.contains_key(&(0, 3)),
            "date-looking text should export without a date number format"
        );

        let (imported, _, _, _) = import_xlsx_data(&bytes, &[], |_, _, _| {}).unwrap();
        let imported = &imported[0].sheet;
        assert_eq!(
            imported.get_display_string(&crate::core::CellRef::new(0, 0)),
            "6/22/26"
        );
        assert_eq!(
            imported.get_display_string(&crate::core::CellRef::new(0, 1)),
            "6/23/26"
        );
        assert_eq!(
            imported.get_display_string(&crate::core::CellRef::new(0, 2)),
            "2026"
        );
        assert!(matches!(
            imported.get_result_data(&crate::core::CellRef::new(0, 3)),
            crate::core::ResultData::String(ref s) if s == "6/22/26"
        ));
    }

    #[test]
    fn test_xlsx_import_data_offset_from_column_a_is_not_lost() {
        let col_a = DataColumn::from_src("A", vec![String::new(), String::new()]);

        let col_b = DataColumn::from_src("B", vec!["42".to_string(), "8".to_string()]);

        let col_c = DataColumn::from_src("C", vec!["=B1 + B2".to_string(), String::new()]);

        let sheet = Sheet {
            id: 1,
            name: "Sheet1".to_string(),
            columns: vec![col_a, col_b, col_c],
            row_heights: Vec::new(),
            hidden_rows: Vec::new(),
            merged_cells: Vec::new(),
            visibility: crate::core::SheetVisibility::Visible,
            freeze_pane: None,
            date1904: false,
            tables: Vec::new(),
            dependencies: std::collections::HashMap::new(),
            dependencies_rev: std::collections::HashMap::new(),
            uncommitted_actions: Vec::new(),
            locale: crate::core::locale::Locale::default(),
        };

        let xlsx_data = export_xlsx_data(&[sheet], &[], &[], None).unwrap();

        let (imported_sheets, _, _, _) = import_xlsx_data(&xlsx_data, &[], |_, _, _| {}).unwrap();
        assert_eq!(imported_sheets.len(), 1);
        let imported = &imported_sheets[0].sheet;

        assert_eq!(imported.columns.len(), 3);
        assert!(imported.columns[0].src[0].is_empty());
        assert!(imported.columns[0].src[1].is_empty());
        assert_eq!(imported.columns[1].src[0], "42");
        assert_eq!(imported.columns[1].src[1], "8");
        assert_eq!(imported.columns[2].src[0], "=B1 + B2");
    }

    #[test]
    fn test_xlsx_excel_table_import_export_cycle() {
        let mut sheet = Sheet::new(crate::core::SheetInit {
            name: Some("Sheet1".to_string()),
            rows: 4,
            cols: 2,
            ..Default::default()
        });
        sheet.set_cell_src(0, 0, "Name".to_string());
        sheet.set_cell_src(0, 1, "Amount".to_string());
        sheet.set_cell_src(1, 0, "Widget".to_string());
        sheet.set_cell_src(1, 1, "10".to_string());
        sheet.set_cell_src(2, 0, "Gadget".to_string());
        sheet.set_cell_src(2, 1, "20".to_string());
        sheet.set_cell_src(3, 1, "=SUM(B2:B3)".to_string());
        sheet.commit(None).unwrap();
        sheet
            .add_table("Sales".to_string(), 0, 0, 3, 1, true, true)
            .unwrap();

        let xlsx_data = export_xlsx_data(&[sheet], &[], &[], None).unwrap();
        let (imported_sheets, _, _, _) = import_xlsx_data(&xlsx_data, &[], |_, _, _| {}).unwrap();
        assert_eq!(imported_sheets.len(), 1);

        let imported_sheet = &imported_sheets[0].sheet;
        assert_eq!(imported_sheet.tables.len(), 1);
        let table = &imported_sheet.tables[0];
        assert_eq!(table.name, "Sales");
        assert_eq!(table.columns, vec!["Name", "Amount"]);
        assert_eq!((table.start_row, table.start_col), (0, 0));
        assert_eq!((table.end_row, table.end_col), (3, 1));
        assert!(table.has_header_row);
        assert!(table.has_totals_row);
        assert_eq!(table.data_start_row(), 1);
        assert_eq!(table.data_end_row(), 2);

        let mut imported_sheet = imported_sheet.clone();
        imported_sheet.mark_all_dirty();
        imported_sheet.commit(None).unwrap();
        let (res, _) = imported_sheet.eval("=SUM(Sales[Amount])", None).unwrap();
        assert_eq!(
            match res {
                crate::core::engine::ResultData::Float(f) => Some(f),
                crate::core::engine::ResultData::Integer(i) => Some(i as f64),
                _ => None,
            },
            Some(30.0)
        );
    }

    #[test]
    fn test_xlsx_zero_data_row_table_import_export_cycle() {
        let mut sheet = Sheet::new(crate::core::SheetInit {
            name: Some("Sheet1".to_string()),
            rows: 1,
            cols: 2,
            ..Default::default()
        });
        sheet.set_cell_src(0, 0, "Name".to_string());
        sheet.set_cell_src(0, 1, "Amount".to_string());
        sheet.commit(None).unwrap();
        sheet
            .add_table("Empty".to_string(), 0, 0, 0, 1, true, false)
            .unwrap();

        let xlsx_data = export_xlsx_data(&[sheet], &[], &[], None).unwrap();
        let (imported_sheets, _, _, _) = import_xlsx_data(&xlsx_data, &[], |_, _, _| {}).unwrap();
        assert_eq!(imported_sheets.len(), 1);

        let imported_sheet = &imported_sheets[0].sheet;
        assert_eq!(imported_sheet.tables.len(), 1);
        let table = &imported_sheet.tables[0];
        assert_eq!(table.name, "Empty");
        assert_eq!(table.columns, vec!["Name", "Amount"]);
        assert!(table.has_header_row);
        assert!(!table.has_totals_row);
        assert!(table.data_start_row() > table.data_end_row());
    }

    #[test]
    fn an_insert_row_attribute_is_read_off_the_table_part() {
        let with_flag = r#"<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Hollow" displayName="Hollow" ref="A1:C2" insertRow="1" totalsRowShown="0"><tableColumns count="3"><tableColumn id="1" name="Region"/><tableColumn id="2" name="Product"/><tableColumn id="3" name="Amount"/></tableColumns></table>"#;
        let parsed = parse_table_part_xml(with_flag).expect("parses");
        assert!(parsed.has_insert_row);
        assert_eq!(parsed.bounds, Some((0, 0, 1, 2)));

        let without = with_flag.replace(" insertRow=\"1\"", "");
        let parsed = parse_table_part_xml(&without).expect("parses");
        assert!(!parsed.has_insert_row);
    }

    #[test]
    fn test_xlsx_table_import_with_absolute_relationship_target() {
        let mut sheet = Sheet::new(crate::core::SheetInit {
            name: Some("Sheet1".to_string()),
            rows: 3,
            cols: 2,
            ..Default::default()
        });
        sheet.set_cell_src(0, 0, "Name".to_string());
        sheet.set_cell_src(0, 1, "Amount".to_string());
        sheet.set_cell_src(1, 0, "Widget".to_string());
        sheet.set_cell_src(1, 1, "10".to_string());
        sheet.commit(None).unwrap();
        sheet
            .add_table("Sales".to_string(), 0, 0, 1, 1, true, false)
            .unwrap();

        let xlsx_data = export_xlsx_data(&[sheet], &[], &[], None).unwrap();

        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(&xlsx_data[..])).unwrap();
        let mut writer = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
        let options = zip::write::SimpleFileOptions::default();
        let mut rewrote_a_target = false;
        for i in 0..archive.len() {
            let mut file = archive.by_index(i).unwrap();
            let name = file.name().to_string();
            let mut buf = Vec::new();
            std::io::Read::read_to_end(&mut file, &mut buf).unwrap();
            drop(file);

            if name.starts_with("xl/worksheets/_rels/") && name.ends_with(".rels") {
                let text = String::from_utf8(buf.clone()).unwrap();
                let new_text = text.replace(r#"Target="../tables/"#, r#"Target="/xl/tables/"#);
                if new_text != text {
                    rewrote_a_target = true;
                }
                buf = new_text.into_bytes();
            }

            writer.start_file(&name, options).unwrap();
            std::io::Write::write_all(&mut writer, &buf).unwrap();
        }
        assert!(
            rewrote_a_target,
            "expected to find a relative table relationship target to rewrite"
        );
        let rewritten = writer.finish().unwrap().into_inner();

        let (imported_sheets, _, _, _) = import_xlsx_data(&rewritten, &[], |_, _, _| {}).unwrap();
        assert_eq!(imported_sheets.len(), 1);
        let imported_sheet = &imported_sheets[0].sheet;
        assert_eq!(imported_sheet.tables.len(), 1);
        assert_eq!(imported_sheet.tables[0].name, "Sales");
        assert_eq!(imported_sheet.tables[0].columns, vec!["Name", "Amount"]);
    }

    fn rewrite_table_ref(xlsx_data: &[u8], new_ref_attr: &str) -> Vec<u8> {
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(xlsx_data)).unwrap();
        let mut writer = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
        let options = zip::write::SimpleFileOptions::default();
        let mut rewrote = false;
        for i in 0..archive.len() {
            let mut file = archive.by_index(i).unwrap();
            let name = file.name().to_string();
            let mut buf = Vec::new();
            std::io::Read::read_to_end(&mut file, &mut buf).unwrap();
            drop(file);

            if name.starts_with("xl/tables/") && name.ends_with(".xml") {
                let text = String::from_utf8(buf.clone()).unwrap();
                let new_text = text.replacen(r#" ref="A1:B2""#, new_ref_attr, 1);
                assert_ne!(new_text, text, "expected to rewrite a table ref");
                rewrote = true;
                buf = new_text.into_bytes();
            }

            writer.start_file(&name, options).unwrap();
            std::io::Write::write_all(&mut writer, &buf).unwrap();
        }
        assert!(rewrote, "expected to find a table part to rewrite");
        writer.finish().unwrap().into_inner()
    }

    fn two_row_table_workbook() -> Vec<u8> {
        let mut sheet = Sheet::new(crate::core::SheetInit {
            name: Some("Sheet1".to_string()),
            rows: 2,
            cols: 2,
            ..Default::default()
        });
        sheet.set_cell_src(0, 0, "Name".to_string());
        sheet.set_cell_src(0, 1, "Amount".to_string());
        sheet.set_cell_src(1, 0, "Widget".to_string());
        sheet.set_cell_src(1, 1, "10".to_string());
        sheet.commit(None).unwrap();
        sheet
            .add_table("Sales".to_string(), 0, 0, 1, 1, true, false)
            .unwrap();
        export_xlsx_data(&[sheet], &[], &[], None).unwrap()
    }

    #[test]
    fn test_xlsx_table_with_unparseable_ref_is_skipped_not_misplaced() {
        for bad_ref in [r#" ref="not-a-range""#, ""] {
            let rewritten = rewrite_table_ref(&two_row_table_workbook(), bad_ref);
            let (imported_sheets, _, _, _) =
                import_xlsx_data(&rewritten, &[], |_, _, _| {}).unwrap();
            assert_eq!(imported_sheets.len(), 1);
            let imported_sheet = &imported_sheets[0].sheet;
            assert!(
                imported_sheet.tables.is_empty(),
                "table with ref {bad_ref:?} should be skipped, not placed"
            );
            assert_eq!(imported_sheet.columns[0].src[0], "Name");
            assert_eq!(imported_sheet.columns[0].src[1], "Widget");
        }
    }

    #[test]
    fn test_xlsx_table_ref_survives_round_trip_unchanged() {
        let (imported_sheets, _, _, _) =
            import_xlsx_data(&two_row_table_workbook(), &[], |_, _, _| {}).unwrap();
        let table = &imported_sheets[0].sheet.tables[0];
        assert_eq!(table.name, "Sales");
        assert_eq!(table.columns, vec!["Name", "Amount"]);
        assert_eq!(
            (
                table.start_row,
                table.start_col,
                table.end_row,
                table.end_col
            ),
            (0, 0, 1, 1)
        );
        assert!(table.has_header_row);
        assert!(!table.has_totals_row);
    }

    #[test]
    fn test_xlsx_multiple_tables_on_one_sheet_all_import() {
        let mut sheet = Sheet::new(crate::core::SheetInit {
            name: Some("Sheet1".to_string()),
            rows: 6,
            cols: 2,
            ..Default::default()
        });
        for (row, (a, b)) in [
            ("Name", "Amount"),
            ("Widget", "10"),
            ("", ""),
            ("Region", "Total"),
            ("East", "20"),
            ("West", "30"),
        ]
        .iter()
        .enumerate()
        {
            sheet.set_cell_src(row, 0, a.to_string());
            sheet.set_cell_src(row, 1, b.to_string());
        }
        sheet.commit(None).unwrap();
        sheet
            .add_table("Sales".to_string(), 0, 0, 1, 1, true, false)
            .unwrap();
        sheet
            .add_table("Regions".to_string(), 3, 0, 5, 1, true, false)
            .unwrap();

        let xlsx_data = export_xlsx_data(&[sheet], &[], &[], None).unwrap();
        let (imported_sheets, _, _, _) = import_xlsx_data(&xlsx_data, &[], |_, _, _| {}).unwrap();
        let tables = &imported_sheets[0].sheet.tables;
        assert_eq!(tables.len(), 2);
        let mut names: Vec<&str> = tables.iter().map(|t| t.name.as_str()).collect();
        names.sort();
        assert_eq!(names, vec!["Regions", "Sales"]);
    }

    #[test]
    fn test_xlsx_empty_table_preservation() {
        let mut columns = Vec::new();
        for _ in 0..5 {
            let col = DataColumn::new(10);
            columns.push(col);
        }

        let sheet = Sheet {
            id: 2,
            name: "EmptyTable".to_string(),
            columns,
            row_heights: Vec::new(),
            hidden_rows: Vec::new(),
            merged_cells: Vec::new(),
            visibility: crate::core::SheetVisibility::Visible,
            freeze_pane: None,
            date1904: false,
            tables: Vec::new(),
            dependencies: std::collections::HashMap::new(),
            dependencies_rev: std::collections::HashMap::new(),
            uncommitted_actions: Vec::new(),
            locale: crate::core::locale::Locale::default(),
        };

        let xlsx_data = export_xlsx_data(&[sheet], &[], &[], None).unwrap();
        assert!(!xlsx_data.is_empty());

        let (imported_sheets, _, _, _) = import_xlsx_data(&xlsx_data, &[], |_, _, _| {}).unwrap();
        assert_eq!(imported_sheets.len(), 1);

        let imported_sheet = &imported_sheets[0].sheet;
        assert_eq!(imported_sheet.name, "EmptyTable");
        assert_eq!(imported_sheet.columns.len(), 5);
        assert_eq!(imported_sheet.row_count(), 10);
    }

    #[test]
    fn test_xlsx_chart_import_export_cycle() {
        for chart_type in [
            crate::core::chart::ChartType::Line,
            crate::core::chart::ChartType::Bar,
            crate::core::chart::ChartType::Column,
            crate::core::chart::ChartType::Scatter,
        ] {
            let mut columns = Vec::new();
            let col1 = DataColumn::from_src(
                "A",
                vec!["Cat1".to_string(), "Cat2".to_string(), "Cat3".to_string()],
            );
            columns.push(col1);

            let col2 = DataColumn::from_src(
                "B",
                vec!["10".to_string(), "20".to_string(), "30".to_string()],
            );
            columns.push(col2);

            let sheet = Sheet {
                id: 1,
                name: "Sheet1".to_string(),
                columns,
                row_heights: Vec::new(),
                hidden_rows: Vec::new(),
                merged_cells: Vec::new(),
                visibility: crate::core::SheetVisibility::Visible,
                freeze_pane: None,
                date1904: false,
                tables: Vec::new(),
                dependencies: std::collections::HashMap::new(),
                dependencies_rev: std::collections::HashMap::new(),
                uncommitted_actions: Vec::new(),
                locale: crate::core::locale::Locale::default(),
            };

            let chart = crate::core::chart::Chart {
                id: 101,
                name: "Chart 1".to_string(),
                chart_type: chart_type.clone(),
                data_range: "Sheet1!A1:B3".to_string(),
                title: Some("My Chart Title".to_string()),
                xlabel: Some("X Axis".to_string()),
                ylabel: Some("Y Axis".to_string()),
                show_legend: true,
                anchor_row: 4,
                anchor_col: 2,
            };

            let xlsx_data =
                export_xlsx_data(&[sheet], std::slice::from_ref(&chart), &[], None).unwrap();
            assert!(!xlsx_data.is_empty());

            let (imported_sheets, imported_charts, _, _) =
                import_xlsx_data(&xlsx_data, &[], |_, _, _| {}).unwrap();

            assert_eq!(imported_sheets.len(), 1);
            assert_eq!(imported_charts.len(), 1, "chart_type={:?}", chart_type);

            let imported_chart = &imported_charts[0];
            assert_eq!(
                imported_chart.data_range, "Sheet1!A1:B3",
                "chart_type={:?}",
                chart_type
            );
            assert_eq!(imported_chart.title, Some("My Chart Title".to_string()));
            assert_eq!(
                imported_chart.xlabel,
                Some("X Axis".to_string()),
                "chart_type={:?}",
                chart_type
            );
            assert_eq!(
                imported_chart.ylabel,
                Some("Y Axis".to_string()),
                "chart_type={:?}",
                chart_type
            );
            assert!(imported_chart.show_legend);
            assert_eq!(imported_chart.chart_type, chart_type);
            assert_eq!(
                (imported_chart.anchor_row, imported_chart.anchor_col),
                (4, 2),
                "chart_type={:?}",
                chart_type
            );
        }
    }

    #[test]
    fn test_fuzz_xlsx_spill_reference_round_trip() {
        let mut sheet = Sheet::new(crate::core::SheetInit {
            name: Some("Sheet1".to_string()),
            rows: 3,
            cols: 2,
            ..Default::default()
        });
        sheet.set_cell_src(0, 0, "=SEQUENCE(3)".to_string());
        sheet.set_cell_src(0, 1, "=SUM($A$1#)".to_string());
        sheet.commit(None).unwrap();
        let data = export_xlsx_data(&[sheet], &[], &[], None).unwrap();
        let mut zip = zip::ZipArchive::new(std::io::Cursor::new(&data)).unwrap();
        let mut xml = String::new();
        std::io::Read::read_to_string(
            &mut zip.by_name("xl/worksheets/sheet1.xml").unwrap(),
            &mut xml,
        )
        .unwrap();
        assert!(xml.contains("SUM(_xlfn.ANCHORARRAY($A$1))"), "{xml}");
        assert!(!xml.contains("$A$1#"), "{xml}");
        let (sheets, _, _, _) = import_xlsx_data(&data, &[], |_, _, _| {}).unwrap();
        let mut sheet = sheets.into_iter().next().unwrap().sheet;
        for column in &mut sheet.columns {
            column.mark_dirty(0);
        }
        sheet.commit(None).unwrap();
        let result = sheet.get_result_data(&crate::core::CellRef::new(0, 1));
        assert!(
            matches!(result, crate::core::engine::ResultData::Float(6.0)),
            "Expected imported spill reference to sum to 6, got {result:?}"
        );
    }

    #[test]
    fn test_xlsx_formula_v_caching() {
        let mut columns = Vec::new();
        let mut col1 = DataColumn::from_src(
            "A",
            vec!["10".to_string(), "20".to_string(), "=A1 + A2".to_string()],
        );
        col1.data
            .set(0, crate::core::engine::ResultData::Integer(10));
        col1.data
            .set(1, crate::core::engine::ResultData::Integer(20));
        col1.data
            .set(2, crate::core::engine::ResultData::Integer(30));
        columns.push(col1);

        let sheet = Sheet {
            id: 1,
            name: "Sheet1".to_string(),
            columns,
            row_heights: Vec::new(),
            hidden_rows: Vec::new(),
            merged_cells: Vec::new(),
            visibility: crate::core::SheetVisibility::Visible,
            freeze_pane: None,
            date1904: false,
            tables: Vec::new(),
            dependencies: std::collections::HashMap::new(),
            dependencies_rev: std::collections::HashMap::new(),
            uncommitted_actions: Vec::new(),
            locale: crate::core::locale::Locale::default(),
        };

        let xlsx_data = export_xlsx_data(&[sheet], &[], &[], None).unwrap();

        let cursor = std::io::Cursor::new(xlsx_data);
        let mut zip = zip::ZipArchive::new(cursor).unwrap();
        let mut sheet_file = zip.by_name("xl/worksheets/sheet1.xml").unwrap();
        let mut xml_content = String::new();
        std::io::Read::read_to_string(&mut sheet_file, &mut xml_content).unwrap();

        assert!(xml_content.contains("<v>30</v>"));
    }

    #[test]
    fn test_xlsx_formula_numeric_looking_string_cache_preserves_type() {
        let mut columns = Vec::new();
        let mut col1 = DataColumn::from_src("A", vec!["=BIN2HEX(\"0001\")".to_string()]);
        col1.data
            .set(0, crate::core::engine::ResultData::String("1".to_string()));
        columns.push(col1);

        let sheet = Sheet {
            id: 1,
            name: "Sheet1".to_string(),
            columns,
            row_heights: Vec::new(),
            hidden_rows: Vec::new(),
            merged_cells: Vec::new(),
            visibility: crate::core::SheetVisibility::Visible,
            freeze_pane: None,
            date1904: false,
            tables: Vec::new(),
            dependencies: std::collections::HashMap::new(),
            dependencies_rev: std::collections::HashMap::new(),
            uncommitted_actions: Vec::new(),
            locale: crate::core::locale::Locale::default(),
        };

        let xlsx_data = export_xlsx_data(&[sheet], &[], &[], None).unwrap();

        let cursor = std::io::Cursor::new(xlsx_data);
        let mut zip = zip::ZipArchive::new(cursor).unwrap();
        let mut sheet_file = zip.by_name("xl/worksheets/sheet1.xml").unwrap();
        let mut xml_content = String::new();
        std::io::Read::read_to_string(&mut sheet_file, &mut xml_content).unwrap();

        assert!(xml_content.contains(r#"<c r="A1" t="str"><f>BIN2HEX("0001")</f><v>1</v></c>"#));
    }

    #[test]
    fn test_xlsx_formula_boolean_looking_string_cache_preserves_type() {
        let mut columns = Vec::new();
        let mut col1 = DataColumn::from_src("A", vec!["=T(TRUE)".to_string()]);
        col1.data.set(
            0,
            crate::core::engine::ResultData::String("TRUE".to_string()),
        );
        columns.push(col1);

        let sheet = Sheet {
            id: 1,
            name: "Sheet1".to_string(),
            columns,
            row_heights: Vec::new(),
            hidden_rows: Vec::new(),
            merged_cells: Vec::new(),
            visibility: crate::core::SheetVisibility::Visible,
            freeze_pane: None,
            date1904: false,
            tables: Vec::new(),
            dependencies: std::collections::HashMap::new(),
            dependencies_rev: std::collections::HashMap::new(),
            uncommitted_actions: Vec::new(),
            locale: crate::core::locale::Locale::default(),
        };

        let xlsx_data = export_xlsx_data(&[sheet], &[], &[], None).unwrap();

        let cursor = std::io::Cursor::new(xlsx_data);
        let mut zip = zip::ZipArchive::new(cursor).unwrap();
        let mut sheet_file = zip.by_name("xl/worksheets/sheet1.xml").unwrap();
        let mut xml_content = String::new();
        std::io::Read::read_to_string(&mut sheet_file, &mut xml_content).unwrap();

        assert!(xml_content.contains(r#"<c r="A1" t="str"><f>T(TRUE)</f><v>TRUE</v></c>"#));
    }

    #[test]
    fn test_xlsx_numeric_looking_text_cell_preserves_type() {
        let mut columns = Vec::new();
        let col1 = DataColumn::from_src("A", vec!["\"1\"".to_string()]);
        columns.push(col1);

        let sheet = Sheet {
            id: 1,
            name: "Sheet1".to_string(),
            columns,
            row_heights: Vec::new(),
            hidden_rows: Vec::new(),
            merged_cells: Vec::new(),
            visibility: crate::core::SheetVisibility::Visible,
            freeze_pane: None,
            date1904: false,
            tables: Vec::new(),
            dependencies: std::collections::HashMap::new(),
            dependencies_rev: std::collections::HashMap::new(),
            uncommitted_actions: Vec::new(),
            locale: crate::core::locale::Locale::default(),
        };

        let xlsx_data = export_xlsx_data(&[sheet], &[], &[], None).unwrap();

        let cursor = std::io::Cursor::new(xlsx_data.clone());
        let mut zip = zip::ZipArchive::new(cursor).unwrap();
        let mut sheet_file = zip.by_name("xl/worksheets/sheet1.xml").unwrap();
        let mut xml_content = String::new();
        std::io::Read::read_to_string(&mut sheet_file, &mut xml_content).unwrap();
        assert!(!xml_content.contains("&quot;1&quot;"));

        let (imported_sheets, _, _, _) = import_xlsx_data(&xlsx_data, &[], |_, _, _| {}).unwrap();
        let mut imported_sheet = imported_sheets.into_iter().next().unwrap().sheet;
        imported_sheet.columns[0].mark_dirty(0);
        imported_sheet.commit(None).unwrap();

        match imported_sheet.columns[0].data.get(0) {
            Some(crate::core::engine::ResultData::String(s)) => assert_eq!(s, "1"),
            other => panic!("Expected ResultData::String(\"1\"), got {:?}", other),
        }
    }

    fn emitted_pivot_parts(
        sheets: &[Sheet],
        pivots: &[crate::core::pivot::PivotTable],
    ) -> std::collections::HashMap<String, String> {
        let bytes = export_xlsx_data(sheets, &[], pivots, None).unwrap();
        let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
        let mut out = std::collections::HashMap::new();
        for i in 0..zip.len() {
            let mut f = zip.by_index(i).unwrap();
            let name = f.name().to_string();
            if name.contains("pivot") && !name.contains("_rels") {
                let mut buf = String::new();
                use std::io::Read;
                f.read_to_string(&mut buf).unwrap();
                out.insert(name, buf);
            }
        }
        out
    }

    fn pivot_shape_fixture() -> (Vec<Sheet>, Vec<crate::core::pivot::PivotTable>) {
        use crate::core::pivot::{
            PivotAggregation, PivotField, PivotFilterField, PivotSource, PivotTable,
            PivotValueField,
        };
        let mut sheet = Sheet::new(crate::core::SheetInit {
            name: Some("Data".to_string()),
            rows: 7,
            cols: 3,
            ..Default::default()
        });
        for (c, h) in ["Region", "Product", "Amount"].iter().enumerate() {
            sheet.set_cell_src(0, c, h.to_string());
        }
        let rows = [
            ["East", "Widget", "10"],
            ["East", "Gadget", "5"],
            ["West", "Widget", "30"],
            ["West", "Gadget", "40"],
            ["North", "Doohickey", "7"],
            ["North", "Widget", "3"],
        ];
        for (r, row) in rows.iter().enumerate() {
            for (c, v) in row.iter().enumerate() {
                sheet.set_cell_src(r + 1, c, v.to_string());
            }
        }
        sheet.commit(None).unwrap();
        let sheet_id = sheet.id;
        let pivot = PivotTable {
            id: 1,
            name: "P1".to_string(),
            source: PivotSource::Range {
                sheet_id,
                start_row: 0,
                start_col: 0,
                end_row: 6,
                end_col: 2,
            },
            dest_sheet_id: sheet_id,
            dest_row: 0,
            dest_col: 5,
            row_fields: vec![PivotField::new("Region")],
            col_fields: vec![],
            value_fields: vec![PivotValueField::new("Amount", PivotAggregation::Sum)],
            filter_fields: vec![PivotFilterField::new("Product")],
            grand_totals_row: true,
            grand_totals_col: true,
            last_output_end_row: None,
            last_output_end_col: None,
        };
        (vec![sheet], vec![pivot])
    }

    #[test]
    fn shared_items_carry_the_values_in_first_seen_order() {
        let (sheets, pivots) = pivot_shape_fixture();
        let parts = emitted_pivot_parts(&sheets, &pivots);
        let cache = parts
            .iter()
            .find(|(k, _)| k.contains("pivotCacheDefinition"))
            .map(|(_, v)| v.as_str())
            .expect("a cache definition");
        assert!(
            cache.contains(
                "<sharedItems count=\"3\"><s v=\"East\"/><s v=\"West\"/><s v=\"North\"/></sharedItems>"
            ),
            "Region's shared items must be in first-seen order: {cache}"
        );
        assert!(
            cache.contains(
                "<sharedItems count=\"3\"><s v=\"Widget\"/><s v=\"Gadget\"/><s v=\"Doohickey\"/></sharedItems>"
            ),
            "Product's shared items must be in first-seen order: {cache}"
        );
    }

    #[test]
    fn item_indices_map_display_order_onto_the_cache_order() {
        let (sheets, pivots) = pivot_shape_fixture();
        let parts = emitted_pivot_parts(&sheets, &pivots);
        let table = parts
            .iter()
            .find(|(k, _)| k.contains("pivotTables/"))
            .map(|(_, v)| v.as_str())
            .expect("a pivot table part");
        assert!(
            table.contains(
                "<items count=\"4\"><item x=\"0\"/><item x=\"2\"/><item x=\"1\"/><item t=\"default\"/></items>"
            ),
            "row field items must index the cache order: {table}"
        );
    }

    #[test]
    fn a_page_field_gets_the_all_placeholder_item() {
        let (sheets, mut pivots) = pivot_shape_fixture();
        for selection in [None, Some(vec!["Widget".to_string()])] {
            pivots[0].filter_fields[0].selected_values = selection.clone();
            let parts = emitted_pivot_parts(&sheets, &pivots);
            let table = parts
                .iter()
                .find(|(k, _)| k.contains("pivotTables/"))
                .map(|(_, v)| v.as_str())
                .expect("a pivot table part");
            let page = table
                .split("axis=\"axisPage\"")
                .nth(1)
                .expect("a page field");
            assert!(
                page.starts_with(|_c: char| true) && page.contains("<item t=\"default\"/></items>"),
                "page field items must end with the (All) placeholder ({selection:?}): {table}"
            );
            if selection.is_some() {
                assert!(
                    page.contains("<item h=\"1\""),
                    "a selection hides the unselected items: {table}"
                );
            }
        }
    }

    #[test]
    fn test_xlsx_pivot_table_round_trip() {
        use crate::core::pivot::{
            PivotAggregation, PivotField, PivotSource, PivotTable, PivotValueField,
        };

        let mut sheet = Sheet::new(crate::core::SheetInit {
            name: Some("Data".to_string()),
            rows: 5,
            cols: 3,
            ..Default::default()
        });
        let header = ["Region", "Product", "Amount"];
        for (c, h) in header.iter().enumerate() {
            sheet.set_cell_src(0, c, h.to_string());
        }
        let rows = [
            ["East", "Widget", "10"],
            ["East", "Gadget", "5"],
            ["West", "Widget", "30"],
            ["West", "Gadget", "40"],
        ];
        for (r, row) in rows.iter().enumerate() {
            for (c, v) in row.iter().enumerate() {
                sheet.set_cell_src(r + 1, c, v.to_string());
            }
        }
        sheet.commit(None).unwrap();
        sheet
            .add_table("Sales".to_string(), 0, 0, 4, 2, true, false)
            .unwrap();

        let pivot = PivotTable {
            id: 42,
            name: "Pivot1".to_string(),
            source: PivotSource::Table {
                name: "Sales".to_string(),
            },
            dest_sheet_id: sheet.id,
            dest_row: 0,
            dest_col: 4,
            row_fields: vec![PivotField::new("Region")],
            col_fields: vec![],
            value_fields: vec![PivotValueField::new("Amount", PivotAggregation::Sum)],
            filter_fields: vec![],
            grand_totals_row: true,
            grand_totals_col: true,
            last_output_end_row: None,
            last_output_end_col: None,
        };

        let xlsx_data = export_xlsx_data(
            std::slice::from_ref(&sheet),
            &[],
            std::slice::from_ref(&pivot),
            None,
        )
        .unwrap();

        let cursor = std::io::Cursor::new(xlsx_data.clone());
        let mut zip = zip::ZipArchive::new(cursor).unwrap();
        assert!(zip.by_name("xl/pivotTables/pivotTable1.xml").is_ok());
        assert!(
            zip.by_name("xl/pivotCache/pivotCacheDefinition1.xml")
                .is_ok()
        );
        assert!(zip.by_name("xl/pivotCache/pivotCacheRecords1.xml").is_ok());

        let (imported_sheets, _, imported_pivots, _) =
            import_xlsx_data(&xlsx_data, &[], |_, _, _| {}).unwrap();
        assert_eq!(imported_sheets.len(), 1);
        assert_eq!(imported_pivots.len(), 1);

        let reimported = &imported_pivots[0];
        assert_eq!(reimported.name, "Pivot1");
        assert_eq!(reimported.row_fields.len(), 1);
        assert_eq!(reimported.row_fields[0].column, "Region");
        assert_eq!(reimported.value_fields.len(), 1);
        assert_eq!(reimported.value_fields[0].column, "Amount");
        assert_eq!(
            reimported.value_fields[0].aggregation,
            PivotAggregation::Sum
        );
        match &reimported.source {
            PivotSource::Table { name } => assert_eq!(name, "Sales"),
            other => panic!("Expected PivotSource::Table, got {:?}", other),
        }

        let reimported_sheet = &imported_sheets[0].sheet;
        let grid = crate::core::pivot::compute_pivot(&[reimported_sheet], reimported).unwrap();
        assert_eq!(grid.body_rows.len(), 3);
    }

    #[test]
    fn test_xlsx_pivot_outer_field_subtotal_off_survives_round_trip() {
        use crate::core::pivot::{
            PivotAggregation, PivotField, PivotSource, PivotTable, PivotValueField,
        };

        let mut sheet = Sheet::new(crate::core::SheetInit {
            name: Some("Data".to_string()),
            rows: 5,
            cols: 3,
            ..Default::default()
        });
        let header = ["Region", "Product", "Amount"];
        for (c, h) in header.iter().enumerate() {
            sheet.set_cell_src(0, c, h.to_string());
        }
        let rows = [
            ["East", "Widget", "10"],
            ["East", "Gadget", "5"],
            ["West", "Widget", "30"],
            ["West", "Gadget", "40"],
        ];
        for (r, row) in rows.iter().enumerate() {
            for (c, v) in row.iter().enumerate() {
                sheet.set_cell_src(r + 1, c, v.to_string());
            }
        }
        sheet.commit(None).unwrap();
        sheet
            .add_table("Sales".to_string(), 0, 0, 4, 2, true, false)
            .unwrap();

        let pivot = PivotTable {
            id: 42,
            name: "Pivot1".to_string(),
            source: PivotSource::Table {
                name: "Sales".to_string(),
            },
            dest_sheet_id: sheet.id,
            dest_row: 0,
            dest_col: 4,
            row_fields: vec![PivotField {
                column: "Region".to_string(),
                subtotal: false,
            }],
            col_fields: vec![],
            value_fields: vec![PivotValueField::new("Amount", PivotAggregation::Sum)],
            filter_fields: vec![],
            grand_totals_row: true,
            grand_totals_col: true,
            last_output_end_row: None,
            last_output_end_col: None,
        };

        let xlsx_data = export_xlsx_data(
            std::slice::from_ref(&sheet),
            &[],
            std::slice::from_ref(&pivot),
            None,
        )
        .unwrap();
        let (_, _, imported_pivots, _) = import_xlsx_data(&xlsx_data, &[], |_, _, _| {}).unwrap();
        let reimported = &imported_pivots[0];

        assert_eq!(reimported.row_fields.len(), 1);
        assert!(
            !reimported.row_fields[0].subtotal,
            "a lone (currently-innermost) row field's subtotal:false should still survive the round-trip"
        );
    }

    #[test]
    fn test_xlsx_document_module_codename_survives_sheet_name_truncation() {
        let long_name = "A".repeat(40);
        let sheet = Sheet {
            id: 1,
            name: long_name.clone(),
            columns: Vec::new(),
            row_heights: Vec::new(),
            hidden_rows: Vec::new(),
            merged_cells: Vec::new(),
            visibility: crate::core::SheetVisibility::Visible,
            freeze_pane: None,
            date1904: false,
            tables: Vec::new(),
            dependencies: std::collections::HashMap::new(),
            dependencies_rev: std::collections::HashMap::new(),
            uncommitted_actions: Vec::new(),
            locale: crate::core::locale::Locale::default(),
        };

        let mut project = crate::core::vba::VbaProject::new_empty();
        let prefix_bytes = project.seed_prefix_bytes.clone();
        let module_cookie = project.seed_module_cookie;
        project.modules.push(crate::core::vba::VbaModule {
            name: "SheetCode".to_string(),
            kind: crate::core::vba::VbaModuleKind::Document,
            source: "Attribute VB_Name = \"SheetCode\"\r\n".to_string(),
            bound_sheet_id: Some(1),
            prefix_bytes,
            module_cookie,
            cached_compressed_source: None,
        });

        let xlsx_data = export_xlsx_data(&[sheet], &[], &[], Some(&project)).unwrap();
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(&xlsx_data[..])).unwrap();
        let workbook_xml = get_zip_file_content(&mut archive, "xl/workbook.xml").unwrap();

        let truncated_name = &long_name[..31];
        assert!(
            !workbook_xml.contains(&format!("name=\"{long_name}\"")),
            "test assumption broken: sheet name wasn't actually truncated"
        );
        let needle = format!("name=\"{truncated_name}\" codeName=\"SheetCode\"");
        assert!(
            workbook_xml.contains(&needle),
            "expected codeName attached to the truncated sheet name in: {workbook_xml}"
        );
    }

    #[test]
    fn test_exported_ooxml_passes_opc_and_xsd_sequence_validation() {
        let mut sheet = Sheet::new(crate::core::SheetInit {
            name: Some("Data".to_string()),
            rows: 3,
            cols: 3,
            ..Default::default()
        });
        sheet.set_column_width(0, Some(18.0));
        sheet.set_row_height(0, Some(24.0));
        sheet.set_cell_src(0, 0, "Item".to_string());
        sheet.set_cell_src(0, 1, "Amount".to_string());
        sheet.set_cell_src(1, 0, "Widget".to_string());
        sheet.set_cell_src(1, 1, "42".to_string());
        sheet.set_cell_src(2, 1, "=B2*2".to_string());
        sheet.columns[1].styles[1] = Some(CellStyle {
            bold: Some(true),
            font_color: Some("#FF0000".to_string()),
            bg_color: Some("#FFFF00".to_string()),
            num_format: Some("0.00".to_string()),
            ..Default::default()
        });
        sheet.tables.push(crate::core::table::ExcelTable {
            id: 1,
            name: "Sales".to_string(),
            sheet_id: sheet.id,
            start_row: 0,
            start_col: 0,
            end_row: 1,
            end_col: 1,
            has_header_row: true,
            has_totals_row: false,
            columns: vec!["Item".to_string(), "Amount".to_string()],
            style_name: Some("TableStyleMedium9".to_string()),
            has_insert_row: false,
        });
        sheet.commit(None).unwrap();

        let chart = crate::core::chart::Chart {
            id: 1,
            name: "Chart 1".to_string(),
            chart_type: crate::core::chart::ChartType::Column,
            data_range: "Data!A1:B2".to_string(),
            title: Some("Sales Chart".to_string()),
            xlabel: Some("Item".to_string()),
            ylabel: Some("Amount".to_string()),
            show_legend: true,
            anchor_row: 4,
            anchor_col: 0,
        };

        let bytes = export_xlsx_data(&[sheet], &[chart], &[], None).unwrap();
        let pkg = OpcPackage::from_bytes(&bytes).unwrap();
        assert_eq!(pkg.validate_opc(), Ok(()));

        let wb_doc = dom::XmlDocument::parse(&pkg.get_str("xl/workbook.xml").unwrap()).unwrap();
        assert_eq!(
            wb_doc.validate_child_sequence(dom::WORKBOOK_SCHEMA_ORDER),
            Ok(())
        );
        let st_doc = dom::XmlDocument::parse(&pkg.get_str("xl/styles.xml").unwrap()).unwrap();
        assert_eq!(
            st_doc.validate_child_sequence(dom::STYLESHEET_SCHEMA_ORDER),
            Ok(())
        );
        let ws_doc =
            dom::XmlDocument::parse(&pkg.get_str("xl/worksheets/sheet1.xml").unwrap()).unwrap();
        assert_eq!(
            ws_doc.validate_child_sequence(dom::WORKSHEET_SCHEMA_ORDER),
            Ok(())
        );
        let tbl_doc =
            dom::XmlDocument::parse(&pkg.get_str("xl/tables/table1.xml").unwrap()).unwrap();
        assert_eq!(
            tbl_doc.validate_child_sequence(dom::TABLE_SCHEMA_ORDER),
            Ok(())
        );
    }

    #[test]
    fn test_surgical_dom_patching_preserves_unmodeled_parts_and_xml_extensions() {
        let mut sheet1 = Sheet::new(crate::core::SheetInit {
            name: Some("Sheet1".to_string()),
            rows: 2,
            cols: 2,
            ..Default::default()
        });
        sheet1.set_cell_src(0, 0, "10".to_string());
        sheet1.set_cell_src(0, 1, "=A1*2".to_string());
        sheet1.commit(None).unwrap();

        let mut sheet2 = Sheet::new(crate::core::SheetInit {
            name: Some("Untouched".to_string()),
            rows: 1,
            cols: 1,
            ..Default::default()
        });
        sheet2.set_cell_src(0, 0, "keep".to_string());
        sheet2.commit(None).unwrap();

        let base_bytes = export_xlsx_data(&[sheet1, sheet2], &[], &[], None).unwrap();
        let mut pkg = OpcPackage::from_bytes(&base_bytes).unwrap();

        let theme_bytes = b"<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"CustomTheme\"/>".to_vec();
        pkg.set_part("xl/theme/theme1.xml", theme_bytes.clone());
        let mut ws1_xml = pkg.get_str("xl/worksheets/sheet1.xml").unwrap();
        ws1_xml = ws1_xml.replace(
            "</worksheet>",
            "<pageSetup orientation=\"landscape\"/><extLst><ext uri=\"{9999}\"><customTag flag=\"yes\"/></ext></extLst></worksheet>",
        );
        pkg.set_part("xl/worksheets/sheet1.xml", ws1_xml.into_bytes());
        let untouched_sheet2_bytes = pkg.get_bytes("xl/worksheets/sheet2.xml").unwrap().to_vec();
        let enriched_bytes = pkg.to_bytes().unwrap();

        let mut wb = crate::core::WorkbookManager::load_bytes(&enriched_bytes).unwrap();
        wb.sheets[0].set_cell_src(0, 0, "25".to_string());
        wb.evaluate().unwrap();
        let saved_bytes = wb.save_bytes().unwrap();

        let out_pkg = OpcPackage::from_bytes(&saved_bytes).unwrap();
        assert_eq!(
            out_pkg.get_bytes("xl/theme/theme1.xml"),
            Some(theme_bytes.as_slice())
        );
        assert_eq!(
            out_pkg.get_bytes("xl/worksheets/sheet2.xml"),
            Some(untouched_sheet2_bytes.as_slice())
        );
        let out_ws1 = out_pkg.get_str("xl/worksheets/sheet1.xml").unwrap();
        assert!(out_ws1.contains("orientation=\"landscape\""));
        assert!(out_ws1.contains("<customTag flag=\"yes\"/>"));

        let ws1_doc = dom::XmlDocument::parse(&out_ws1).unwrap();
        assert_eq!(
            ws1_doc.validate_child_sequence(dom::WORKSHEET_SCHEMA_ORDER),
            Ok(())
        );

        let reloaded = crate::core::WorkbookManager::load_bytes(&saved_bytes).unwrap();
        assert_eq!(reloaded.sheets[0].columns[0].src[0], "25");
        assert_eq!(
            reloaded.sheets[0]
                .get_result_data(&crate::core::CellRef::new(0, 1))
                .to_string(),
            "50"
        );
    }

    #[test]
    fn test_phase5_core_sheet_structure_round_trip_and_surgical_patch() {
        use crate::core::{FreezePane, MergedCellRange, SheetVisibility};

        let mut sheet1 = Sheet::new(crate::core::SheetInit {
            name: Some("Main".to_string()),
            rows: 5,
            cols: 5,
            ..Default::default()
        });
        sheet1.set_cell_src(0, 0, "Title".to_string());
        sheet1.add_merged_cells(MergedCellRange {
            start_row: 0,
            start_col: 0,
            end_row: 1,
            end_col: 2,
        });
        sheet1.set_column_hidden(1, true);
        sheet1.set_row_hidden(2, true);
        sheet1.freeze_pane = Some(FreezePane { rows: 1, cols: 2 });
        sheet1.date1904 = true;
        sheet1.commit(None).unwrap();

        let mut sheet2 = Sheet::new(crate::core::SheetInit {
            name: Some("HiddenSheet".to_string()),
            rows: 2,
            cols: 2,
            ..Default::default()
        });
        sheet2.visibility = SheetVisibility::Hidden;
        sheet2.date1904 = true;

        let mut sheet3 = Sheet::new(crate::core::SheetInit {
            name: Some("VeryHiddenSheet".to_string()),
            rows: 2,
            cols: 2,
            ..Default::default()
        });
        sheet3.visibility = SheetVisibility::VeryHidden;
        sheet3.date1904 = true;

        let bytes = export_xlsx_data(&[sheet1, sheet2, sheet3], &[], &[], None).unwrap();
        let pkg = OpcPackage::from_bytes(&bytes).unwrap();
        assert_eq!(pkg.validate_opc(), Ok(()));

        let wb_doc = dom::XmlDocument::parse(&pkg.get_str("xl/workbook.xml").unwrap()).unwrap();
        assert_eq!(
            wb_doc.validate_child_sequence(dom::WORKBOOK_SCHEMA_ORDER),
            Ok(())
        );
        let ws_doc =
            dom::XmlDocument::parse(&pkg.get_str("xl/worksheets/sheet1.xml").unwrap()).unwrap();
        assert_eq!(
            ws_doc.validate_child_sequence(dom::WORKSHEET_SCHEMA_ORDER),
            Ok(())
        );

        let mut wb = crate::core::WorkbookManager::load_bytes(&bytes).unwrap();
        assert!(wb.date1904);
        assert!(wb.sheets[0].date1904);
        assert_eq!(
            wb.sheets[0].merged_cells,
            vec![MergedCellRange {
                start_row: 0,
                start_col: 0,
                end_row: 1,
                end_col: 2,
            }]
        );
        assert!(wb.sheets[0].is_column_hidden(1));
        assert!(!wb.sheets[0].is_column_hidden(0));
        assert!(wb.sheets[0].is_row_hidden(2));
        assert!(!wb.sheets[0].is_row_hidden(0));
        assert_eq!(
            wb.sheets[0].freeze_pane,
            Some(FreezePane { rows: 1, cols: 2 })
        );
        assert_eq!(wb.sheets[0].visibility, SheetVisibility::Visible);
        assert_eq!(wb.sheets[1].visibility, SheetVisibility::Hidden);
        assert_eq!(wb.sheets[2].visibility, SheetVisibility::VeryHidden);

        wb.sheets[0].ensure_capacity(4, 4);
        wb.sheets[0].add_merged_cells(MergedCellRange {
            start_row: 3,
            start_col: 3,
            end_row: 4,
            end_col: 4,
        });
        wb.sheets[0].set_column_hidden(3, true);
        wb.sheets[0].set_row_hidden(4, true);
        wb.sheets[0].freeze_pane = Some(FreezePane { rows: 2, cols: 0 });
        wb.sheets[1].visibility = SheetVisibility::Visible;

        let patched_bytes = wb.save_bytes().unwrap();
        let patched_pkg = OpcPackage::from_bytes(&patched_bytes).unwrap();
        assert_eq!(patched_pkg.validate_opc(), Ok(()));
        let patched_ws_doc =
            dom::XmlDocument::parse(&patched_pkg.get_str("xl/worksheets/sheet1.xml").unwrap())
                .unwrap();
        assert_eq!(
            patched_ws_doc.validate_child_sequence(dom::WORKSHEET_SCHEMA_ORDER),
            Ok(())
        );

        let reloaded = crate::core::WorkbookManager::load_bytes(&patched_bytes).unwrap();
        assert_eq!(reloaded.sheets[0].merged_cells.len(), 2);
        assert!(reloaded.sheets[0].is_column_hidden(3));
        assert!(reloaded.sheets[0].is_row_hidden(4));
        assert_eq!(
            reloaded.sheets[0].freeze_pane,
            Some(FreezePane { rows: 2, cols: 0 })
        );
        assert_eq!(reloaded.sheets[1].visibility, SheetVisibility::Visible);

        let d1904 = crate::core::date::excel_serial_to_date_with_epoch(0.0, true);
        assert_eq!((d1904.year, d1904.month, d1904.day), (1904, 1, 1));
        assert_eq!(
            crate::core::date::date_to_excel_serial_with_epoch(d1904, true),
            0.0
        );
    }

    #[test]
    fn test_row_insert_and_delete_adjusts_merged_cells_and_hidden_rows() {
        use crate::core::MergedCellRange;

        let mut sheet = Sheet::new(crate::core::SheetInit {
            name: Some("Sheet1".to_string()),
            rows: 5,
            cols: 3,
            ..Default::default()
        });
        sheet.add_merged_cells(MergedCellRange {
            start_row: 1,
            start_col: 0,
            end_row: 2,
            end_col: 1,
        });
        sheet.set_row_hidden(2, true);

        sheet.insert_row(0);
        assert_eq!(
            sheet.merged_cells,
            vec![MergedCellRange {
                start_row: 2,
                start_col: 0,
                end_row: 3,
                end_col: 1,
            }]
        );
        assert!(sheet.is_row_hidden(3));
        assert!(!sheet.is_row_hidden(2));

        sheet.delete_row(0);
        assert_eq!(
            sheet.merged_cells,
            vec![MergedCellRange {
                start_row: 1,
                start_col: 0,
                end_row: 2,
                end_col: 1,
            }]
        );
        assert!(sheet.is_row_hidden(2));
    }
}
