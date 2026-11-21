use std::collections::{HashMap, HashSet};

use crate::core::chart::Chart;
use crate::core::engine::{
    CellType, FreezePane, MergedCellRange, ResultData, Sheet, SheetVisibility,
};
use crate::core::opc::{
    CT_CHART, CT_COMMENTS, CT_DRAWING, CT_PIVOT_CACHE_DEF, CT_PIVOT_CACHE_RECORDS, CT_PIVOT_TABLE,
    CT_SHARED_STRINGS, CT_TABLE, CT_VBA_PROJECT, CT_WORKBOOK, CT_WORKBOOK_MACRO_ENABLED,
    CT_WORKSHEET, OpcPackage, REL_CHART, REL_COMMENTS, REL_DRAWING, REL_HYPERLINK,
    REL_PIVOT_CACHE_DEF, REL_PIVOT_CACHE_RECORDS, REL_PIVOT_TABLE, REL_SHARED_STRINGS, REL_TABLE,
    REL_VBA_PROJECT, REL_WORKSHEET, Relationship, next_rid, rel_type_matches, rels_path_for,
};
use crate::core::parser::col_idx_to_letters;
use crate::core::pivot::PivotTable;
use crate::core::style::CellStyle;
use crate::core::table::ExcelTable;
use crate::core::vba::{VbaModuleKind, VbaProject};
use crate::core::xlsx::dom::{
    STYLESHEET_SCHEMA_ORDER, WORKBOOK_SCHEMA_ORDER, WORKSHEET_SCHEMA_ORDER, XmlDocument,
    XmlElement, XmlNode,
};
use crate::core::xlsx::styles::parse_color_argb;
use crate::core::xlsx::xml_util::{NamespaceFlavor, encode_xstring, needs_xml_space_preserve};

#[derive(Debug, Clone)]
struct CellSnapshot {
    src: String,
    res: ResultData,
    cell_type: CellType,
    style: Option<CellStyle>,
}

impl PartialEq for CellSnapshot {
    fn eq(&self, other: &Self) -> bool {
        self.src == other.src
            && self.cell_type == other.cell_type
            && self.style == other.style
            && result_data_eq(&self.res, &other.res)
    }
}

fn result_data_eq(a: &ResultData, b: &ResultData) -> bool {
    match (a, b) {
        (ResultData::None, ResultData::None) => true,
        (ResultData::Integer(x), ResultData::Integer(y)) => x == y,
        (ResultData::Float(x), ResultData::Float(y)) => x.to_bits() == y.to_bits(),
        (ResultData::String(x), ResultData::String(y)) => x == y,
        (ResultData::Boolean(x), ResultData::Boolean(y)) => x == y,
        (ResultData::Error(x), ResultData::Error(y)) => x == y,
        (ResultData::List(x), ResultData::List(y)) => {
            x.len() == y.len() && x.iter().zip(y.iter()).all(|(l, r)| result_data_eq(l, r))
        }
        _ => false,
    }
}

#[derive(Debug, Clone)]
struct SheetSnapshot {
    sheet_id: u64,
    name: String,
    visibility: SheetVisibility,
    part_path: String,
    cells: HashMap<(usize, usize), CellSnapshot>,
    column_widths: HashMap<usize, f64>,
    hidden_cols: HashSet<usize>,
    row_heights: HashMap<usize, f64>,
    hidden_rows: HashSet<usize>,
    merged_cells: Vec<MergedCellRange>,
    freeze_pane: Option<FreezePane>,
    comments: Vec<crate::core::CellComment>,
    hyperlinks: Vec<crate::core::CellHyperlink>,
    data_validations: Vec<crate::core::DataValidationRule>,
    conditional_formats: Vec<crate::core::ConditionalFormatRule>,
    tables: Vec<ExcelTable>,
}

#[derive(Debug, Clone)]
pub(crate) struct OoxmlSourceState {
    package: OpcPackage,
    workbook_part: String,
    flavor: NamespaceFlavor,
    date1904: bool,
    defined_names: Vec<crate::core::DefinedName>,
    shared_strings_part: Option<String>,
    styles_part: Option<String>,
    sheet_snapshots: Vec<SheetSnapshot>,
    charts: Vec<Chart>,
    pivots: Vec<PivotTable>,
    vba: Option<VbaProject>,
}

impl OoxmlSourceState {
    pub(crate) fn capture(
        package: OpcPackage,
        wb: &crate::core::xlsx::workbook::ParsedWorkbook,
        sheets: &[Sheet],
        charts: &[Chart],
        pivots: &[PivotTable],
        vba: Option<&VbaProject>,
    ) -> Self {
        let mut sheet_snapshots = Vec::with_capacity(sheets.len());
        for (sheet, meta) in sheets.iter().zip(wb.sheets.iter()) {
            sheet_snapshots.push(SheetSnapshot {
                sheet_id: sheet.id,
                name: sheet.name.clone(),
                visibility: sheet.visibility,
                part_path: meta.part_path.clone(),
                cells: snapshot_cells(sheet),
                column_widths: snapshot_col_widths(sheet),
                hidden_cols: snapshot_hidden_cols(sheet),
                row_heights: snapshot_row_heights(sheet),
                hidden_rows: snapshot_hidden_rows(sheet),
                merged_cells: sheet.merged_cells.clone(),
                freeze_pane: sheet.freeze_pane,
                comments: sheet.comments.clone(),
                hyperlinks: sheet.hyperlinks.clone(),
                data_validations: sheet.data_validations.clone(),
                conditional_formats: sheet.conditional_formats.clone(),
                tables: sheet.tables.clone(),
            });
        }
        Self {
            package,
            workbook_part: wb.workbook_part.clone(),
            flavor: wb.flavor,
            date1904: wb.date1904,
            defined_names: wb.defined_names.clone(),
            shared_strings_part: wb.shared_strings_part.clone(),
            styles_part: wb.styles_part.clone(),
            sheet_snapshots,
            charts: charts.to_vec(),
            pivots: pivots.to_vec(),
            vba: vba.cloned(),
        }
    }

    pub(crate) fn can_patch_surgically(&self, sheets: &[Sheet]) -> bool {
        !sheets.is_empty()
    }

    pub(crate) fn patch_to_bytes(
        &self,
        sheets: &[Sheet],
        charts: &[Chart],
        pivots: &[PivotTable],
        vba: Option<&VbaProject>,
        defined_names: &[crate::core::DefinedName],
    ) -> Result<Vec<u8>, String> {
        let mut pkg = self.package.clone();

        let snap_by_id: HashMap<u64, (usize, &SheetSnapshot)> = self
            .sheet_snapshots
            .iter()
            .enumerate()
            .map(|(i, s)| (s.sheet_id, (i, s)))
            .collect();
        let active_ids: HashSet<u64> = sheets.iter().map(|s| s.id).collect();
        let topology_changed = sheets.len() != self.sheet_snapshots.len()
            || sheets
                .iter()
                .zip(self.sheet_snapshots.iter())
                .any(|(s, snap)| s.id != snap.sheet_id);

        let mut added_sheet_info: HashMap<u64, (String, String)> = HashMap::new();
        if topology_changed {
            let mut wb_rels = pkg.read_relationships(&self.workbook_part);
            let mut wb_rels_changed = false;
            let mut ct = pkg.read_content_types();
            let mut ct_changed = false;

            for snap in &self.sheet_snapshots {
                if !active_ids.contains(&snap.sheet_id) {
                    remove_sheet_part_and_children(&mut pkg, &mut ct, &snap.part_path);
                    ct_changed = true;
                    let before = wb_rels.len();
                    wb_rels.retain(|r| {
                        !(rel_type_matches(&r.rel_type, "worksheet")
                            && !r.is_external()
                            && r.resolved_target(&self.workbook_part)
                                .eq_ignore_ascii_case(&snap.part_path))
                    });
                    if wb_rels.len() != before {
                        wb_rels_changed = true;
                    }
                }
            }

            for sheet in sheets {
                if !snap_by_id.contains_key(&sheet.id) {
                    let (_, new_part) =
                        allocate_free_numbered_part(&pkg, "xl/worksheets/sheet", ".xml");
                    let base_xml = format!(
                        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"{}\"><sheetData/></worksheet>",
                        self.flavor.spreadsheetml()
                    );
                    pkg.set_part(&new_part, base_xml.into_bytes());
                    ct.set_override(&new_part, CT_WORKSHEET);
                    ct_changed = true;

                    let rid = next_rid(&wb_rels);
                    let target = new_part.trim_start_matches("xl/").to_string();
                    wb_rels.push(Relationship {
                        id: rid.clone(),
                        rel_type: REL_WORKSHEET.to_string(),
                        target,
                        target_mode: None,
                    });
                    wb_rels_changed = true;
                    added_sheet_info.insert(sheet.id, (rid, new_part));
                }
            }

            if wb_rels_changed {
                pkg.write_relationships(&self.workbook_part, &wb_rels);
            }
            if ct_changed {
                pkg.write_content_types(&ct);
            }
        }

        let cur_date1904 = sheets.iter().any(|s| s.date1904);
        let defined_names_changed = defined_names != self.defined_names.as_slice();
        let workbook_meta_changed = topology_changed
            || cur_date1904 != self.date1904
            || defined_names_changed
            || sheets.iter().any(|s| {
                snap_by_id
                    .get(&s.id)
                    .is_some_and(|(_, snap)| s.name != snap.name || s.visibility != snap.visibility)
            });
        if workbook_meta_changed
            && let Some(wb_xml) = pkg.get_str(&self.workbook_part)
            && let Ok(mut doc) = XmlDocument::parse(&wb_xml)
        {
            if cur_date1904 != self.date1904 {
                if cur_date1904 {
                    let wb_pr = doc
                        .root
                        .ensure_child_in_schema_order("workbookPr", WORKBOOK_SCHEMA_ORDER);
                    wb_pr.set_attr("date1904", "1");
                } else if let Some(wb_pr) = doc.root.child_mut("workbookPr") {
                    wb_pr.remove_attr("date1904");
                }
            }
            if defined_names_changed {
                doc.root.remove_children_by_local("definedNames");
                if !defined_names.is_empty() {
                    let mut dns_el = XmlElement::new(doc.root.qualified_child_name("definedNames"));
                    for dn in defined_names {
                        let mut dn_el = XmlElement::new(dns_el.qualified_child_name("definedName"));
                        dn_el.set_attr("name", dn.name.clone());
                        if let Some(local_id) = dn.local_sheet_id {
                            dn_el.set_attr("localSheetId", local_id.to_string());
                        }
                        if dn.hidden {
                            dn_el.set_attr("hidden", "1");
                        }
                        if let Some(ref comment) = dn.comment {
                            dn_el.set_attr("comment", comment.clone());
                        }
                        if !dn.formula.is_empty() {
                            dn_el.children.push(XmlNode::Text(dn.formula.clone()));
                        }
                        dns_el.children.push(XmlNode::Element(dn_el));
                    }
                    doc.root
                        .insert_child_in_schema_order(dns_el, WORKBOOK_SCHEMA_ORDER);
                }
            }
            if let Some(sheets_el) = doc.root.child_mut("sheets") {
                let mut existing_sheet_els: Vec<XmlElement> = Vec::new();
                for node in std::mem::take(&mut sheets_el.children) {
                    if let XmlNode::Element(el) = node
                        && el.local_name() == "sheet"
                    {
                        existing_sheet_els.push(el);
                    }
                }
                let mut max_sheet_id: u32 = existing_sheet_els
                    .iter()
                    .filter_map(|el| el.attr("sheetId").and_then(|v| v.parse::<u32>().ok()))
                    .max()
                    .unwrap_or(0);
                let mut by_snap_idx: HashMap<usize, XmlElement> =
                    existing_sheet_els.into_iter().enumerate().collect();

                for sheet in sheets {
                    let mut el = if let Some(&(snap_idx, _)) = snap_by_id.get(&sheet.id)
                        && let Some(orig) = by_snap_idx.remove(&snap_idx)
                    {
                        orig
                    } else {
                        max_sheet_id += 1;
                        let mut new_el = XmlElement::new(sheets_el.qualified_child_name("sheet"));
                        new_el.set_attr("sheetId", max_sheet_id.to_string());
                        if let Some((rid, _)) = added_sheet_info.get(&sheet.id) {
                            new_el.set_attr("r:id", rid.clone());
                        }
                        new_el
                    };
                    el.set_attr("name", sheet.name.clone());
                    match crate::core::xlsx::workbook::sheet_visibility_attr(sheet.visibility) {
                        Some(state) => el.set_attr("state", state),
                        None => {
                            el.remove_attr("state");
                        }
                    }
                    sheets_el.children.push(XmlNode::Element(el));
                }
            }
            pkg.set_part(&self.workbook_part, doc.to_xml_string().into_bytes());
        }

        let mut sst_patcher: Option<SharedStringsPatcher> = None;
        let mut styles_patcher: Option<StylesPatcher> = None;
        let mut any_formula_changed = defined_names_changed || topology_changed;

        let empty_snap_template = |sheet: &Sheet, part_path: String| SheetSnapshot {
            sheet_id: sheet.id,
            name: sheet.name.clone(),
            visibility: sheet.visibility,
            part_path,
            cells: HashMap::new(),
            column_widths: HashMap::new(),
            hidden_cols: HashSet::new(),
            row_heights: HashMap::new(),
            hidden_rows: HashSet::new(),
            merged_cells: Vec::new(),
            freeze_pane: None,
            comments: Vec::new(),
            hyperlinks: Vec::new(),
            data_validations: Vec::new(),
            conditional_formats: Vec::new(),
            tables: Vec::new(),
        };

        for sheet in sheets {
            let owned_empty;
            let snap: &SheetSnapshot = if let Some(&(_, existing)) = snap_by_id.get(&sheet.id) {
                existing
            } else if let Some((_, new_part)) = added_sheet_info.get(&sheet.id) {
                owned_empty = empty_snap_template(sheet, new_part.clone());
                &owned_empty
            } else {
                continue;
            };
            let cur_cells = snapshot_cells(sheet);
            let cur_widths = snapshot_col_widths(sheet);
            let cur_hidden_cols = snapshot_hidden_cols(sheet);
            let cur_heights = snapshot_row_heights(sheet);
            let cur_hidden_rows = snapshot_hidden_rows(sheet);

            let cells_changed = cur_cells != snap.cells;
            let cols_changed =
                cur_widths != snap.column_widths || cur_hidden_cols != snap.hidden_cols;
            let heights_changed = cur_heights != snap.row_heights;
            let hidden_rows_changed = cur_hidden_rows != snap.hidden_rows;
            let merges_changed = sheet.merged_cells != snap.merged_cells;
            let freeze_changed = sheet.freeze_pane != snap.freeze_pane;
            let comments_changed = sheet.comments != snap.comments;
            let hyperlinks_changed = sheet.hyperlinks != snap.hyperlinks;
            let validations_changed = sheet.data_validations != snap.data_validations;
            let cf_changed = sheet.conditional_formats != snap.conditional_formats;
            let tables_changed = sheet.tables != snap.tables;

            let default_cur_sheet = sheets.first().map(|s| s.name.as_str()).unwrap_or("");
            let default_prev_sheet = self
                .sheet_snapshots
                .first()
                .map(|s| s.name.as_str())
                .unwrap_or("");
            let cur_sheet_charts: Vec<&Chart> = charts
                .iter()
                .filter(|c| chart_target_sheet_name(c, default_cur_sheet) == sheet.name)
                .collect();
            let prev_sheet_charts: Vec<&Chart> = self
                .charts
                .iter()
                .filter(|c| chart_target_sheet_name(c, default_prev_sheet) == snap.name)
                .collect();
            let charts_changed = cur_sheet_charts != prev_sheet_charts;

            if !cells_changed
                && !cols_changed
                && !heights_changed
                && !hidden_rows_changed
                && !merges_changed
                && !freeze_changed
                && !comments_changed
                && !hyperlinks_changed
                && !validations_changed
                && !cf_changed
                && !tables_changed
                && !charts_changed
            {
                continue;
            }

            if comments_changed {
                patch_comments(&mut pkg, &snap.part_path, &sheet.comments, self.flavor);
            }

            let ws_xml_changed = cells_changed
                || cols_changed
                || heights_changed
                || hidden_rows_changed
                || merges_changed
                || freeze_changed
                || hyperlinks_changed
                || validations_changed
                || cf_changed
                || tables_changed
                || charts_changed;

            if !ws_xml_changed {
                continue;
            }

            let Some(ws_xml) = pkg.get_str(&snap.part_path) else {
                continue;
            };
            let mut doc = XmlDocument::parse(&ws_xml)?;

            if freeze_changed {
                patch_freeze_pane(&mut doc.root, sheet.freeze_pane);
            }

            if cols_changed {
                patch_cols(&mut doc.root, &cur_widths, &cur_hidden_cols);
            }

            if merges_changed {
                patch_merge_cells(&mut doc.root, &sheet.merged_cells);
            }

            if cf_changed {
                if styles_patcher.is_none() {
                    styles_patcher = Some(StylesPatcher::load(
                        &pkg,
                        self.styles_part.as_deref(),
                        self.flavor,
                    ));
                }
                patch_conditional_formats(
                    &mut doc.root,
                    &sheet.conditional_formats,
                    styles_patcher.as_mut().unwrap(),
                );
            }

            if validations_changed {
                patch_data_validations(&mut doc.root, &sheet.data_validations);
            }

            if hyperlinks_changed {
                patch_hyperlinks(
                    &mut doc.root,
                    &mut pkg,
                    &snap.part_path,
                    &sheet.hyperlinks,
                    self.flavor,
                );
            }

            if tables_changed {
                patch_tables(
                    &mut doc.root,
                    &mut pkg,
                    &snap.part_path,
                    &sheet.tables,
                    self.flavor,
                );
            }

            if charts_changed {
                patch_charts(
                    &mut doc.root,
                    &mut pkg,
                    &snap.part_path,
                    &sheet.name,
                    &cur_sheet_charts,
                    self.flavor,
                );
            }

            if cells_changed || heights_changed || hidden_rows_changed {
                let mut changed_cells: Vec<(usize, usize, Option<&CellSnapshot>)> = Vec::new();
                for (&rc, cur) in &cur_cells {
                    let prev = snap.cells.get(&rc);
                    if prev != Some(cur) {
                        if cur.src.starts_with('=') || prev.is_some_and(|p| p.src.starts_with('='))
                        {
                            any_formula_changed = true;
                        }
                        changed_cells.push((rc.0, rc.1, Some(cur)));
                    }
                }
                for (&rc, prev) in &snap.cells {
                    if !cur_cells.contains_key(&rc) {
                        if prev.src.starts_with('=') {
                            any_formula_changed = true;
                        }
                        changed_cells.push((rc.0, rc.1, None));
                    }
                }
                changed_cells.sort_by_key(|(r, c, _)| (*r, *c));

                let mut changed_heights: Vec<(usize, Option<f64>)> = Vec::new();
                for (&r, &h) in &cur_heights {
                    if snap.row_heights.get(&r) != Some(&h) {
                        changed_heights.push((r, Some(h)));
                    }
                }
                for &r in snap.row_heights.keys() {
                    if !cur_heights.contains_key(&r) {
                        changed_heights.push((r, None));
                    }
                }
                changed_heights.sort_by_key(|(r, _)| *r);

                let mut changed_hidden_rows: Vec<(usize, bool)> = Vec::new();
                for &r in &cur_hidden_rows {
                    if !snap.hidden_rows.contains(&r) {
                        changed_hidden_rows.push((r, true));
                    }
                }
                for &r in &snap.hidden_rows {
                    if !cur_hidden_rows.contains(&r) {
                        changed_hidden_rows.push((r, false));
                    }
                }
                changed_hidden_rows.sort_by_key(|(r, _)| *r);

                if sst_patcher.is_none() {
                    sst_patcher = Some(SharedStringsPatcher::load(
                        &pkg,
                        self.shared_strings_part.as_deref(),
                        self.flavor,
                    ));
                }
                if styles_patcher.is_none() {
                    styles_patcher = Some(StylesPatcher::load(
                        &pkg,
                        self.styles_part.as_deref(),
                        self.flavor,
                    ));
                }

                patch_sheet_data(
                    &mut doc.root,
                    sheet,
                    &changed_cells,
                    &changed_heights,
                    &changed_hidden_rows,
                    sst_patcher.as_mut().unwrap(),
                    styles_patcher.as_mut().unwrap(),
                );
            }

            pkg.set_part(&snap.part_path, doc.to_xml_string().into_bytes());
        }

        if let Some(sst) = sst_patcher
            && sst.dirty
        {
            let sst_part = self
                .shared_strings_part
                .clone()
                .unwrap_or_else(|| "xl/sharedStrings.xml".to_string());
            let had_part = pkg.has_part(&sst_part);
            pkg.set_part(&sst_part, sst.doc.to_xml_string().into_bytes());
            if !had_part {
                let mut ct = pkg.read_content_types();
                ct.set_override(&sst_part, CT_SHARED_STRINGS);
                pkg.write_content_types(&ct);

                let mut rels = pkg.read_relationships(&self.workbook_part);
                if !rels
                    .iter()
                    .any(|r| rel_type_matches(&r.rel_type, "sharedStrings"))
                {
                    let rid = next_rid(&rels);
                    rels.push(Relationship {
                        id: rid,
                        rel_type: REL_SHARED_STRINGS.to_string(),
                        target: "sharedStrings.xml".to_string(),
                        target_mode: None,
                    });
                    pkg.write_relationships(&self.workbook_part, &rels);
                }
            }
        }

        if let Some(st) = styles_patcher
            && st.dirty
        {
            let styles_part = self
                .styles_part
                .clone()
                .unwrap_or_else(|| "xl/styles.xml".to_string());
            pkg.set_part(&styles_part, st.doc.to_xml_string().into_bytes());
        }

        if any_formula_changed {
            invalidate_calc_chain(&mut pkg, &self.workbook_part);
        }

        let deleted_sheet_ids: HashSet<u64> = self
            .sheet_snapshots
            .iter()
            .map(|s| s.sheet_id)
            .filter(|id| !active_ids.contains(id))
            .collect();
        let renamed_sheet_ids: HashSet<u64> = sheets
            .iter()
            .filter(|s| {
                snap_by_id
                    .get(&s.id)
                    .is_some_and(|(_, snap)| snap.name != s.name)
            })
            .map(|s| s.id)
            .collect();
        let pivots_changed = pivots != self.pivots.as_slice()
            || (!deleted_sheet_ids.is_empty() && (!self.pivots.is_empty() || !pivots.is_empty()))
            || (!renamed_sheet_ids.is_empty() && !pivots.is_empty());
        if pivots_changed {
            let mut sheet_part_by_id: HashMap<u64, String> = HashMap::new();
            for sheet in sheets {
                if let Some(&(_, snap)) = snap_by_id.get(&sheet.id) {
                    sheet_part_by_id.insert(sheet.id, snap.part_path.clone());
                } else if let Some((_, new_part)) = added_sheet_info.get(&sheet.id) {
                    sheet_part_by_id.insert(sheet.id, new_part.clone());
                }
            }
            patch_pivots(
                &mut pkg,
                &self.workbook_part,
                sheets,
                &sheet_part_by_id,
                pivots,
                self.flavor,
            )?;
        }

        if vba != self.vba.as_ref() {
            patch_vba(&mut pkg, &self.workbook_part, sheets, vba)?;
        }

        pkg.to_bytes()
    }
}

fn invalidate_calc_chain(pkg: &mut OpcPackage, workbook_part: &str) {
    let mut calc_chain_parts: Vec<String> = Vec::new();
    let mut rels = pkg.read_relationships(workbook_part);
    let before_rels = rels.len();
    rels.retain(|r| {
        if rel_type_matches(&r.rel_type, "calcChain") {
            if !r.is_external() {
                calc_chain_parts.push(r.resolved_target(workbook_part));
            }
            false
        } else {
            true
        }
    });
    if rels.len() != before_rels {
        pkg.write_relationships(workbook_part, &rels);
    }

    if calc_chain_parts.is_empty() && pkg.has_part("xl/calcChain.xml") {
        calc_chain_parts.push("xl/calcChain.xml".to_string());
    }

    if !calc_chain_parts.is_empty() {
        let mut ct = pkg.read_content_types();
        let mut ct_changed = false;
        for part in &calc_chain_parts {
            pkg.remove_part(part);
            if ct.remove_override(part) {
                ct_changed = true;
            }
        }
        if ct_changed {
            pkg.write_content_types(&ct);
        }
    }
}

fn snapshot_cells(sheet: &Sheet) -> HashMap<(usize, usize), CellSnapshot> {
    let mut out = HashMap::new();
    let rows = sheet.row_count();
    for (col_idx, col) in sheet.columns.iter().enumerate() {
        for row_idx in 0..rows {
            let src = col.src.get(row_idx).cloned().unwrap_or_default();
            let res = col.data.get(row_idx).unwrap_or(ResultData::None);
            let cell_type = col
                .cell_types
                .get(row_idx)
                .copied()
                .unwrap_or(CellType::Empty);
            let style = col.styles.get(row_idx).and_then(|s| s.clone());
            if !src.is_empty()
                || !matches!(res, ResultData::None)
                || cell_type != CellType::Empty
                || style.is_some()
            {
                out.insert(
                    (row_idx, col_idx),
                    CellSnapshot {
                        src,
                        res,
                        cell_type,
                        style,
                    },
                );
            }
        }
    }
    out
}

fn snapshot_col_widths(sheet: &Sheet) -> HashMap<usize, f64> {
    sheet
        .columns
        .iter()
        .enumerate()
        .filter_map(|(idx, col)| {
            col.width
                .filter(|w| w.is_finite() && *w >= 0.0)
                .map(|w| (idx, w))
        })
        .collect()
}

fn snapshot_hidden_cols(sheet: &Sheet) -> HashSet<usize> {
    sheet
        .columns
        .iter()
        .enumerate()
        .filter_map(|(idx, col)| col.hidden.then_some(idx))
        .collect()
}

fn snapshot_row_heights(sheet: &Sheet) -> HashMap<usize, f64> {
    sheet
        .row_heights
        .iter()
        .enumerate()
        .filter_map(|(idx, h)| {
            (*h).filter(|v| v.is_finite() && *v >= 0.0)
                .map(|v| (idx, v))
        })
        .collect()
}

fn snapshot_hidden_rows(sheet: &Sheet) -> HashSet<usize> {
    sheet
        .hidden_rows
        .iter()
        .enumerate()
        .filter_map(|(idx, &hidden)| hidden.then_some(idx))
        .collect()
}

struct SharedStringsPatcher {
    doc: XmlDocument,
    index: HashMap<String, usize>,
    count: usize,
    dirty: bool,
}

impl SharedStringsPatcher {
    fn load(pkg: &OpcPackage, part: Option<&str>, flavor: NamespaceFlavor) -> Self {
        if let Some(p) = part
            && let Some(xml) = pkg.get_str(p)
            && let Ok(doc) = XmlDocument::parse(&xml)
        {
            let existing =
                crate::core::xlsx::shared_strings::parse_shared_strings(&xml).unwrap_or_default();
            let mut index = HashMap::new();
            for (i, s) in existing.iter().enumerate() {
                index.entry(s.clone()).or_insert(i);
            }
            let count = existing.len();
            return Self {
                doc,
                index,
                count,
                dirty: false,
            };
        }
        let mut root = XmlElement::new("sst");
        root.set_attr("xmlns", flavor.spreadsheetml());
        root.set_attr("count", "0");
        root.set_attr("uniqueCount", "0");
        Self {
            doc: XmlDocument::new(root),
            index: HashMap::new(),
            count: 0,
            dirty: false,
        }
    }

    fn intern(&mut self, value: &str) -> usize {
        if let Some(&idx) = self.index.get(value) {
            return idx;
        }
        let idx = self.count;
        self.count += 1;
        self.index.insert(value.to_string(), idx);

        let si_name = self.doc.root.qualified_child_name("si");
        let t_name = self.doc.root.qualified_child_name("t");
        let encoded = encode_xstring(value);
        let mut t_el = XmlElement::new(t_name);
        if needs_xml_space_preserve(&encoded) {
            t_el.set_attr("xml:space", "preserve");
        }
        if !encoded.is_empty() {
            t_el.children.push(XmlNode::Text(encoded));
        }
        let mut si_el = XmlElement::new(si_name);
        si_el.children.push(XmlNode::Element(t_el));
        self.doc.root.children.push(XmlNode::Element(si_el));

        let unique_str = self.count.to_string();
        self.doc.root.set_attr("uniqueCount", unique_str.clone());
        let prev_count = self
            .doc
            .root
            .attr("count")
            .and_then(|s| s.parse::<usize>().ok())
            .unwrap_or(0);
        self.doc
            .root
            .set_attr("count", (prev_count + 1).max(self.count).to_string());
        self.dirty = true;
        idx
    }
}

struct StylesPatcher {
    doc: XmlDocument,
    next_num_fmt_id: u32,
    dirty: bool,
}

impl StylesPatcher {
    fn load(pkg: &OpcPackage, part: Option<&str>, flavor: NamespaceFlavor) -> Self {
        if let Some(p) = part
            && let Some(xml) = pkg.get_str(p)
            && let Ok(doc) = XmlDocument::parse(&xml)
        {
            let mut max_id = 163u32;
            if let Some(num_fmts) = doc.root.child("numFmts") {
                for nf in num_fmts.children_by_local("numFmt") {
                    if let Some(id) = nf.attr("numFmtId").and_then(|s| s.parse::<u32>().ok()) {
                        max_id = max_id.max(id);
                    }
                }
            }
            return Self {
                doc,
                next_num_fmt_id: max_id + 1,
                dirty: false,
            };
        }
        let default_xml = crate::core::xlsx::styles::StyleRegistry::default().to_xml(flavor);
        let doc = XmlDocument::parse(&default_xml).unwrap();
        Self {
            doc,
            next_num_fmt_id: 164,
            dirty: false,
        }
    }

    fn append_style(&mut self, style: &CellStyle) -> u32 {
        if style.is_empty() {
            return 0;
        }
        let num_fmt_id = match style.num_format.as_deref() {
            Some(code) if !code.is_empty() && !code.eq_ignore_ascii_case("General") => {
                if let Some(&(id, _)) = crate::core::xlsx::styles::BUILTIN_NUM_FMTS
                    .iter()
                    .find(|(_, c)| *c == code)
                {
                    id
                } else {
                    let id = self.next_num_fmt_id;
                    self.next_num_fmt_id += 1;
                    let num_fmts = self
                        .doc
                        .root
                        .ensure_child_in_schema_order("numFmts", STYLESHEET_SCHEMA_ORDER);
                    let mut nf = XmlElement::new(num_fmts.qualified_child_name("numFmt"));
                    nf.set_attr("numFmtId", id.to_string());
                    nf.set_attr("formatCode", code);
                    num_fmts.children.push(XmlNode::Element(nf));
                    let cnt = num_fmts.children_by_local("numFmt").count();
                    num_fmts.set_attr("count", cnt.to_string());
                    id
                }
            }
            _ => 0,
        };

        let has_font = style.bold.unwrap_or(false)
            || style.italic.unwrap_or(false)
            || style.underline.unwrap_or(false)
            || style.strikethrough.unwrap_or(false)
            || style.font_color.is_some()
            || style.font_family.is_some()
            || style.font_size.is_some();

        let font_id = if has_font {
            let fonts = self
                .doc
                .root
                .ensure_child_in_schema_order("fonts", STYLESHEET_SCHEMA_ORDER);
            let id = fonts.children_by_local("font").count();
            let mut font_el = XmlElement::new(fonts.qualified_child_name("font"));
            if style.bold.unwrap_or(false) {
                font_el.children.push(XmlNode::Element(XmlElement::new(
                    font_el.qualified_child_name("b"),
                )));
            }
            if style.italic.unwrap_or(false) {
                font_el.children.push(XmlNode::Element(XmlElement::new(
                    font_el.qualified_child_name("i"),
                )));
            }
            if style.strikethrough.unwrap_or(false) {
                font_el.children.push(XmlNode::Element(XmlElement::new(
                    font_el.qualified_child_name("strike"),
                )));
            }
            if style.underline.unwrap_or(false) {
                font_el.children.push(XmlNode::Element(XmlElement::new(
                    font_el.qualified_child_name("u"),
                )));
            }
            let mut sz = XmlElement::new(font_el.qualified_child_name("sz"));
            sz.set_attr("val", style.font_size.unwrap_or(11.0).to_string());
            font_el.children.push(XmlNode::Element(sz));
            if let Some(ref c) = style.font_color
                && let Some(argb) = parse_color_argb(c)
            {
                let mut col_el = XmlElement::new(font_el.qualified_child_name("color"));
                col_el.set_attr("rgb", argb);
                font_el.children.push(XmlNode::Element(col_el));
            }
            let mut name_el = XmlElement::new(font_el.qualified_child_name("name"));
            name_el.set_attr("val", style.font_family.as_deref().unwrap_or("Calibri"));
            font_el.children.push(XmlNode::Element(name_el));
            fonts.children.push(XmlNode::Element(font_el));
            fonts.set_attr("count", (id + 1).to_string());
            id
        } else {
            0
        };

        let fg_argb = style.bg_color.as_deref().and_then(parse_color_argb);
        let fill_id = if fg_argb.is_some()
            || style
                .fill_pattern
                .is_some_and(|p| p != crate::core::FillPattern::None)
        {
            let fills = self
                .doc
                .root
                .ensure_child_in_schema_order("fills", STYLESHEET_SCHEMA_ORDER);
            let id = fills.children_by_local("fill").count();
            let mut fill_el = XmlElement::new(fills.qualified_child_name("fill"));
            let mut pat = XmlElement::new(fill_el.qualified_child_name("patternFill"));
            pat.set_attr(
                "patternType",
                style
                    .fill_pattern
                    .unwrap_or(crate::core::FillPattern::Solid)
                    .as_ooxml(),
            );
            if let Some(argb) = fg_argb {
                let mut fg = XmlElement::new(pat.qualified_child_name("fgColor"));
                fg.set_attr("rgb", argb);
                let mut bg = XmlElement::new(pat.qualified_child_name("bgColor"));
                bg.set_attr("indexed", "64");
                pat.children.push(XmlNode::Element(fg));
                pat.children.push(XmlNode::Element(bg));
            }
            fill_el.children.push(XmlNode::Element(pat));
            fills.children.push(XmlNode::Element(fill_el));
            fills.set_attr("count", (id + 1).to_string());
            id
        } else {
            0
        };

        let border_id = match style.borders.as_ref().filter(|b| !b.is_empty()) {
            Some(b) => {
                let borders = self
                    .doc
                    .root
                    .ensure_child_in_schema_order("borders", STYLESHEET_SCHEMA_ORDER);
                let id = borders.children_by_local("border").count();
                let mut border_el = XmlElement::new(borders.qualified_child_name("border"));
                if b.diagonal_up {
                    border_el.set_attr("diagonalUp", "1");
                }
                if b.diagonal_down {
                    border_el.set_attr("diagonalDown", "1");
                }
                for (tag, side) in [
                    ("left", b.left.as_ref()),
                    ("right", b.right.as_ref()),
                    ("top", b.top.as_ref()),
                    ("bottom", b.bottom.as_ref()),
                    ("diagonal", b.diagonal.as_ref()),
                ] {
                    let mut side_el = XmlElement::new(border_el.qualified_child_name(tag));
                    if let Some(s) = side
                        && s.style != crate::core::BorderStyle::None
                    {
                        side_el.set_attr("style", s.style.as_ooxml());
                        if let Some(argb) = s.color.as_deref().and_then(parse_color_argb) {
                            let mut c_el = XmlElement::new(side_el.qualified_child_name("color"));
                            c_el.set_attr("rgb", argb);
                            side_el.children.push(XmlNode::Element(c_el));
                        }
                    }
                    border_el.children.push(XmlNode::Element(side_el));
                }
                borders.children.push(XmlNode::Element(border_el));
                borders.set_attr("count", (id + 1).to_string());
                id
            }
            None => 0,
        };

        let cell_xfs = self
            .doc
            .root
            .ensure_child_in_schema_order("cellXfs", STYLESHEET_SCHEMA_ORDER);
        let xf_idx = cell_xfs.children_by_local("xf").count() as u32;
        let mut xf = XmlElement::new(cell_xfs.qualified_child_name("xf"));
        xf.set_attr("numFmtId", num_fmt_id.to_string());
        xf.set_attr("fontId", font_id.to_string());
        xf.set_attr("fillId", fill_id.to_string());
        xf.set_attr("borderId", border_id.to_string());
        xf.set_attr("xfId", "0");
        if num_fmt_id != 0 {
            xf.set_attr("applyNumberFormat", "1");
        }
        if font_id != 0 {
            xf.set_attr("applyFont", "1");
        }
        if fill_id != 0 {
            xf.set_attr("applyFill", "1");
        }
        if border_id != 0 {
            xf.set_attr("applyBorder", "1");
        }
        if let Some(ref align) = style.alignment
            && !align.is_empty()
        {
            xf.set_attr("applyAlignment", "1");
            let mut a_el = XmlElement::new(xf.qualified_child_name("alignment"));
            if let Some(h) = align.horizontal {
                a_el.set_attr("horizontal", h.as_ooxml());
            }
            if let Some(v) = align.vertical {
                a_el.set_attr("vertical", v.as_ooxml());
            }
            if let Some(rot) = align.text_rotation {
                a_el.set_attr("textRotation", rot.to_string());
            }
            if let Some(true) = align.wrap_text {
                a_el.set_attr("wrapText", "1");
            }
            if let Some(indent) = align.indent {
                a_el.set_attr("indent", indent.to_string());
            }
            if let Some(true) = align.shrink_to_fit {
                a_el.set_attr("shrinkToFit", "1");
            }
            xf.children.push(XmlNode::Element(a_el));
        }
        cell_xfs.children.push(XmlNode::Element(xf));
        cell_xfs.set_attr("count", (xf_idx + 1).to_string());
        self.dirty = true;
        xf_idx
    }

    fn register_dxf(&mut self, style: &CellStyle) -> u32 {
        let dxfs = self
            .doc
            .root
            .ensure_child_in_schema_order("dxfs", STYLESHEET_SCHEMA_ORDER);
        let dxf_idx = dxfs.children_by_local("dxf").count() as u32;
        let mut dxf_el = XmlElement::new(dxfs.qualified_child_name("dxf"));
        let has_font = style.bold.unwrap_or(false)
            || style.italic.unwrap_or(false)
            || style.underline.unwrap_or(false)
            || style.strikethrough.unwrap_or(false)
            || style.font_color.is_some();
        if has_font {
            let mut font_el = XmlElement::new(dxf_el.qualified_child_name("font"));
            if style.bold.unwrap_or(false) {
                font_el.children.push(XmlNode::Element(XmlElement::new(
                    font_el.qualified_child_name("b"),
                )));
            }
            if style.italic.unwrap_or(false) {
                font_el.children.push(XmlNode::Element(XmlElement::new(
                    font_el.qualified_child_name("i"),
                )));
            }
            if style.strikethrough.unwrap_or(false) {
                font_el.children.push(XmlNode::Element(XmlElement::new(
                    font_el.qualified_child_name("strike"),
                )));
            }
            if style.underline.unwrap_or(false) {
                font_el.children.push(XmlNode::Element(XmlElement::new(
                    font_el.qualified_child_name("u"),
                )));
            }
            if let Some(argb) = style.font_color.as_deref().and_then(parse_color_argb) {
                let mut color_el = XmlElement::new(font_el.qualified_child_name("color"));
                color_el.set_attr("rgb", argb);
                font_el.children.push(XmlNode::Element(color_el));
            }
            dxf_el.children.push(XmlNode::Element(font_el));
        }
        if let Some(argb) = style.bg_color.as_deref().and_then(parse_color_argb) {
            let mut fill_el = XmlElement::new(dxf_el.qualified_child_name("fill"));
            let mut pf_el = XmlElement::new(fill_el.qualified_child_name("patternFill"));
            let pat = style
                .fill_pattern
                .unwrap_or(crate::core::style::FillPattern::Solid)
                .as_ooxml();
            pf_el.set_attr("patternType", pat);
            let mut bg_el = XmlElement::new(pf_el.qualified_child_name("bgColor"));
            bg_el.set_attr("rgb", argb);
            pf_el.children.push(XmlNode::Element(bg_el));
            fill_el.children.push(XmlNode::Element(pf_el));
            dxf_el.children.push(XmlNode::Element(fill_el));
        }
        dxfs.children.push(XmlNode::Element(dxf_el));
        dxfs.set_attr("count", (dxf_idx + 1).to_string());
        self.dirty = true;
        dxf_idx
    }
}

fn patch_conditional_formats(
    ws_root: &mut XmlElement,
    rules: &[crate::core::ConditionalFormatRule],
    styles: &mut StylesPatcher,
) {
    ws_root.remove_children_by_local("conditionalFormatting");
    for cf in rules {
        let mut cf_el = XmlElement::new(ws_root.qualified_child_name("conditionalFormatting"));
        cf_el.set_attr("sqref", cf.sqref.clone());
        let mut rule_el = XmlElement::new(cf_el.qualified_child_name("cfRule"));
        rule_el.set_attr("type", cf.rule_type.clone());
        rule_el.set_attr("priority", cf.priority.to_string());
        if let Some(ref op) = cf.operator {
            rule_el.set_attr("operator", op.clone());
        }
        if let Some(ref style) = cf.style
            && !style.is_empty()
        {
            let dxf_id = styles.register_dxf(style);
            rule_el.set_attr("dxfId", dxf_id.to_string());
        }
        if cf.stop_if_true {
            rule_el.set_attr("stopIfTrue", "1");
        }
        for formula in &cf.formulas {
            let mut f_el = XmlElement::new(rule_el.qualified_child_name("formula"));
            f_el.children.push(XmlNode::Text(formula.clone()));
            rule_el.children.push(XmlNode::Element(f_el));
        }
        cf_el.children.push(XmlNode::Element(rule_el));
        ws_root.insert_child_in_schema_order(cf_el, WORKSHEET_SCHEMA_ORDER);
    }
}

fn patch_data_validations(ws_root: &mut XmlElement, rules: &[crate::core::DataValidationRule]) {
    ws_root.remove_children_by_local("dataValidations");
    if rules.is_empty() {
        return;
    }
    let mut dvs_el = XmlElement::new(ws_root.qualified_child_name("dataValidations"));
    dvs_el.set_attr("count", rules.len().to_string());
    for dv in rules {
        let mut dv_el = XmlElement::new(dvs_el.qualified_child_name("dataValidation"));
        dv_el.set_attr("type", dv.kind.as_ooxml());
        dv_el.set_attr("sqref", dv.sqref.clone());
        if let Some(op) = dv.operator {
            dv_el.set_attr("operator", op.as_ooxml());
        }
        if dv.allow_blank {
            dv_el.set_attr("allowBlank", "1");
        }
        if dv.show_input_message {
            dv_el.set_attr("showInputMessage", "1");
        }
        if dv.show_error_message {
            dv_el.set_attr("showErrorMessage", "1");
        }
        if let Some(ref t) = dv.prompt_title {
            dv_el.set_attr("promptTitle", t.clone());
        }
        if let Some(ref p) = dv.prompt {
            dv_el.set_attr("prompt", p.clone());
        }
        if let Some(ref t) = dv.error_title {
            dv_el.set_attr("errorTitle", t.clone());
        }
        if let Some(ref e) = dv.error {
            dv_el.set_attr("error", e.clone());
        }
        if let Some(ref f1) = dv.formula1 {
            let mut f1_el = XmlElement::new(dv_el.qualified_child_name("formula1"));
            f1_el.children.push(XmlNode::Text(f1.clone()));
            dv_el.children.push(XmlNode::Element(f1_el));
        }
        if let Some(ref f2) = dv.formula2 {
            let mut f2_el = XmlElement::new(dv_el.qualified_child_name("formula2"));
            f2_el.children.push(XmlNode::Text(f2.clone()));
            dv_el.children.push(XmlNode::Element(f2_el));
        }
        dvs_el.children.push(XmlNode::Element(dv_el));
    }
    ws_root.insert_child_in_schema_order(dvs_el, WORKSHEET_SCHEMA_ORDER);
}

fn patch_hyperlinks(
    ws_root: &mut XmlElement,
    pkg: &mut OpcPackage,
    ws_part: &str,
    hyperlinks: &[crate::core::CellHyperlink],
    flavor: NamespaceFlavor,
) {
    ws_root.remove_children_by_local("hyperlinks");
    let mut rels = pkg.read_relationships(ws_part);
    let before_rels = rels.len();
    rels.retain(|r| !rel_type_matches(&r.rel_type, "hyperlink"));
    let mut rels_changed = rels.len() != before_rels;

    if !hyperlinks.is_empty() {
        let mut hls_el = XmlElement::new(ws_root.qualified_child_name("hyperlinks"));
        let mut needs_rel_ns = false;
        for hl in hyperlinks {
            let range_ref = if hl.start_row == hl.end_row && hl.start_col == hl.end_col {
                format!("{}{}", col_idx_to_letters(hl.start_col), hl.start_row + 1)
            } else {
                format!(
                    "{}{}:{}{}",
                    col_idx_to_letters(hl.start_col),
                    hl.start_row + 1,
                    col_idx_to_letters(hl.end_col),
                    hl.end_row + 1
                )
            };
            let mut hl_el = XmlElement::new(hls_el.qualified_child_name("hyperlink"));
            hl_el.set_attr("ref", range_ref);
            if let Some(ref target) = hl.target {
                let rid = next_rid(&rels);
                rels.push(Relationship {
                    id: rid.clone(),
                    rel_type: REL_HYPERLINK.to_string(),
                    target: target.clone(),
                    target_mode: Some("External".to_string()),
                });
                rels_changed = true;
                needs_rel_ns = true;
                hl_el.set_attr("r:id", rid);
            }
            if let Some(ref loc) = hl.location {
                hl_el.set_attr("location", loc.clone());
            }
            if let Some(ref disp) = hl.display {
                hl_el.set_attr("display", disp.clone());
            }
            if let Some(ref tip) = hl.tooltip {
                hl_el.set_attr("tooltip", tip.clone());
            }
            hls_el.children.push(XmlNode::Element(hl_el));
        }
        if needs_rel_ns && ws_root.attr("xmlns:r").is_none() {
            ws_root.set_attr("xmlns:r", flavor.office_rels());
        }
        ws_root.insert_child_in_schema_order(hls_el, WORKSHEET_SCHEMA_ORDER);
    }

    if rels_changed {
        pkg.write_relationships(ws_part, &rels);
    }
}

fn patch_comments(
    pkg: &mut OpcPackage,
    ws_part: &str,
    comments: &[crate::core::CellComment],
    flavor: NamespaceFlavor,
) {
    let mut rels = pkg.read_relationships(ws_part);
    let existing_part = rels
        .iter()
        .find(|r| rel_type_matches(&r.rel_type, "comments") && !r.is_external())
        .map(|r| r.resolved_target(ws_part));

    if comments.is_empty() {
        let before = rels.len();
        rels.retain(|r| !rel_type_matches(&r.rel_type, "comments"));
        if rels.len() != before {
            pkg.write_relationships(ws_part, &rels);
        }
        if let Some(part) = existing_part {
            pkg.remove_part(&part);
            let mut ct = pkg.read_content_types();
            if ct.remove_override(&part) {
                pkg.write_content_types(&ct);
            }
        }
        return;
    }

    let target_part = if let Some(part) = existing_part {
        part
    } else {
        let mut idx = 1usize;
        let candidate = loop {
            let p = format!("xl/comments{idx}.xml");
            if !pkg.has_part(&p) {
                break p;
            }
            idx += 1;
        };
        let rid = next_rid(&rels);
        let rel_target = format!("../{}", candidate.trim_start_matches("xl/"));
        rels.push(Relationship {
            id: rid,
            rel_type: REL_COMMENTS.to_string(),
            target: rel_target,
            target_mode: None,
        });
        pkg.write_relationships(ws_part, &rels);
        let mut ct = pkg.read_content_types();
        ct.set_override(&candidate, CT_COMMENTS);
        pkg.write_content_types(&ct);
        candidate
    };

    let xml = crate::core::xlsx::comments::write_comments_xml(comments, flavor);
    pkg.set_part(target_part, xml.into_bytes());
}

fn remove_sheet_part_and_children(
    pkg: &mut OpcPackage,
    ct: &mut crate::core::opc::ContentTypes,
    ws_part: &str,
) {
    let rels = pkg.read_relationships(ws_part);
    for r in rels {
        if r.is_external() {
            continue;
        }
        let target = r.resolved_target(ws_part);
        if rel_type_matches(&r.rel_type, "comments") || rel_type_matches(&r.rel_type, "table") {
            pkg.remove_part(&target);
            ct.remove_override(&target);
        } else if rel_type_matches(&r.rel_type, "drawing") {
            for drel in pkg.read_relationships(&target) {
                if rel_type_matches(&drel.rel_type, "chart") && !drel.is_external() {
                    let chart_part = drel.resolved_target(&target);
                    pkg.remove_part(&chart_part);
                    ct.remove_override(&chart_part);
                }
            }
            pkg.remove_part(&rels_path_for(&target));
            pkg.remove_part(&target);
            ct.remove_override(&target);
        } else if rel_type_matches(&r.rel_type, "pivotTable") {
            pkg.remove_part(&rels_path_for(&target));
            pkg.remove_part(&target);
            ct.remove_override(&target);
        }
    }
    pkg.remove_part(&rels_path_for(ws_part));
    pkg.remove_part(ws_part);
    ct.remove_override(ws_part);
}

fn chart_target_sheet_name<'a>(ch: &'a Chart, default_sheet: &'a str) -> &'a str {
    let ref_sheet = ch.data_range.split('!').next().unwrap_or("").trim();
    if ref_sheet.is_empty() {
        default_sheet
    } else {
        ref_sheet.trim_matches('\'')
    }
}

fn allocate_free_numbered_part(pkg: &OpcPackage, prefix: &str, suffix: &str) -> (usize, String) {
    let mut id = 1usize;
    loop {
        let candidate = format!("{prefix}{id}{suffix}");
        if !pkg.has_part(&candidate) {
            return (id, candidate);
        }
        id += 1;
    }
}

fn patch_tables(
    ws_root: &mut XmlElement,
    pkg: &mut OpcPackage,
    ws_part: &str,
    tables: &[ExcelTable],
    flavor: NamespaceFlavor,
) {
    ws_root.remove_children_by_local("tableParts");
    let mut rels = pkg.read_relationships(ws_part);
    let mut old_parts: Vec<String> = Vec::new();
    let before_rels = rels.len();
    rels.retain(|r| {
        if rel_type_matches(&r.rel_type, "table") {
            if !r.is_external() {
                old_parts.push(r.resolved_target(ws_part));
            }
            false
        } else {
            true
        }
    });
    let mut rels_changed = rels.len() != before_rels;
    let mut ct = pkg.read_content_types();
    let mut ct_changed = false;

    for part in old_parts {
        pkg.remove_part(&part);
        if ct.remove_override(&part) {
            ct_changed = true;
        }
    }

    if !tables.is_empty() {
        let mut tps_el = XmlElement::new(ws_root.qualified_child_name("tableParts"));
        tps_el.set_attr("count", tables.len().to_string());
        for tbl in tables {
            let (tbl_id, tbl_part) = allocate_free_numbered_part(pkg, "xl/tables/table", ".xml");
            let tbl_xml = crate::core::xlsx::table::serialize_table_xml(tbl, tbl_id, flavor);
            pkg.set_part(&tbl_part, tbl_xml.into_bytes());
            ct.set_override(&tbl_part, CT_TABLE);
            ct_changed = true;

            let rid = next_rid(&rels);
            let rel_target = format!("../{}", tbl_part.trim_start_matches("xl/"));
            rels.push(Relationship {
                id: rid.clone(),
                rel_type: REL_TABLE.to_string(),
                target: rel_target,
                target_mode: None,
            });
            rels_changed = true;

            let mut tp_el = XmlElement::new(tps_el.qualified_child_name("tablePart"));
            tp_el.set_attr("r:id", rid);
            tps_el.children.push(XmlNode::Element(tp_el));
        }
        if ws_root.attr("xmlns:r").is_none() {
            ws_root.set_attr("xmlns:r", flavor.office_rels());
        }
        ws_root.insert_child_in_schema_order(tps_el, WORKSHEET_SCHEMA_ORDER);
    }

    if rels_changed {
        pkg.write_relationships(ws_part, &rels);
    }
    if ct_changed {
        pkg.write_content_types(&ct);
    }
}

fn patch_charts(
    ws_root: &mut XmlElement,
    pkg: &mut OpcPackage,
    ws_part: &str,
    ws_name: &str,
    charts: &[&Chart],
    flavor: NamespaceFlavor,
) {
    ws_root.remove_children_by_local("drawing");
    let mut rels = pkg.read_relationships(ws_part);
    let mut old_drawings: Vec<String> = Vec::new();
    let before_rels = rels.len();
    rels.retain(|r| {
        if rel_type_matches(&r.rel_type, "drawing") {
            if !r.is_external() {
                old_drawings.push(r.resolved_target(ws_part));
            }
            false
        } else {
            true
        }
    });
    let mut rels_changed = rels.len() != before_rels;
    let mut ct = pkg.read_content_types();
    let mut ct_changed = false;

    for drawing_part in old_drawings {
        let drawing_rels = pkg.read_relationships(&drawing_part);
        for drel in drawing_rels {
            if rel_type_matches(&drel.rel_type, "chart") && !drel.is_external() {
                let chart_part = drel.resolved_target(&drawing_part);
                pkg.remove_part(&chart_part);
                if ct.remove_override(&chart_part) {
                    ct_changed = true;
                }
            }
        }
        pkg.remove_part(&rels_path_for(&drawing_part));
        pkg.remove_part(&drawing_part);
        if ct.remove_override(&drawing_part) {
            ct_changed = true;
        }
    }

    if !charts.is_empty() {
        let (_, drawing_part) = allocate_free_numbered_part(pkg, "xl/drawings/drawing", ".xml");
        let mut drawing_rels: Vec<Relationship> = Vec::new();
        let mut anchors: Vec<(usize, usize, String)> = Vec::new();

        for ch in charts {
            let (_, chart_part) = allocate_free_numbered_part(pkg, "xl/charts/chart", ".xml");
            let chart_xml = crate::core::xlsx::chart::serialize_chart_xml(ch, ws_name, flavor);
            pkg.set_part(&chart_part, chart_xml.into_bytes());
            ct.set_override(&chart_part, CT_CHART);

            let c_rid = next_rid(&drawing_rels);
            let chart_target = format!("../{}", chart_part.trim_start_matches("xl/"));
            drawing_rels.push(Relationship {
                id: c_rid.clone(),
                rel_type: REL_CHART.to_string(),
                target: chart_target,
                target_mode: None,
            });
            anchors.push((ch.anchor_row, ch.anchor_col, c_rid));
        }

        let drawing_xml = crate::core::xlsx::chart::serialize_drawing_xml(&anchors, flavor);
        pkg.set_part(&drawing_part, drawing_xml.into_bytes());
        pkg.write_relationships(&drawing_part, &drawing_rels);
        ct.set_override(&drawing_part, CT_DRAWING);
        ct_changed = true;

        let d_rid = next_rid(&rels);
        let drawing_target = format!("../{}", drawing_part.trim_start_matches("xl/"));
        rels.push(Relationship {
            id: d_rid.clone(),
            rel_type: REL_DRAWING.to_string(),
            target: drawing_target,
            target_mode: None,
        });
        rels_changed = true;

        if ws_root.attr("xmlns:r").is_none() {
            ws_root.set_attr("xmlns:r", flavor.office_rels());
        }
        let mut dr_el = XmlElement::new(ws_root.qualified_child_name("drawing"));
        dr_el.set_attr("r:id", d_rid);
        ws_root.insert_child_in_schema_order(dr_el, WORKSHEET_SCHEMA_ORDER);
    }

    if rels_changed {
        pkg.write_relationships(ws_part, &rels);
    }
    if ct_changed {
        pkg.write_content_types(&ct);
    }
}

fn patch_pivots(
    pkg: &mut OpcPackage,
    workbook_part: &str,
    sheets: &[Sheet],
    sheet_part_by_id: &HashMap<u64, String>,
    pivots: &[PivotTable],
    flavor: NamespaceFlavor,
) -> Result<(), String> {
    let sheet_refs: Vec<&Sheet> = sheets.iter().collect();
    let mut units = Vec::new();
    for pivot in pivots {
        if sheets.iter().any(|s| s.id == pivot.dest_sheet_id) {
            units.push(crate::core::pivot_xlsx::build_pivot_xml_unit(
                &sheet_refs,
                pivot,
                units.len(),
                flavor,
            )?);
        }
    }

    let mut ct = pkg.read_content_types();
    let mut ct_changed = false;

    for ws_part in sheet_part_by_id.values() {
        let mut ws_rels = pkg.read_relationships(ws_part);
        let before = ws_rels.len();
        let mut old_pt_parts = Vec::new();
        ws_rels.retain(|r| {
            if rel_type_matches(&r.rel_type, "pivotTable") {
                if !r.is_external() {
                    old_pt_parts.push(r.resolved_target(ws_part));
                }
                false
            } else {
                true
            }
        });
        if ws_rels.len() != before {
            pkg.write_relationships(ws_part, &ws_rels);
        }
        for pt_part in old_pt_parts {
            pkg.remove_part(&rels_path_for(&pt_part));
            pkg.remove_part(&pt_part);
            if ct.remove_override(&pt_part) {
                ct_changed = true;
            }
        }
    }

    let mut wb_rels = pkg.read_relationships(workbook_part);
    let before_wb_rels = wb_rels.len();
    let mut old_pcd_parts = Vec::new();
    wb_rels.retain(|r| {
        if rel_type_matches(&r.rel_type, "pivotCacheDefinition") {
            if !r.is_external() {
                old_pcd_parts.push(r.resolved_target(workbook_part));
            }
            false
        } else {
            true
        }
    });
    let mut wb_rels_changed = wb_rels.len() != before_wb_rels;

    for pcd_part in old_pcd_parts {
        for pcd_rel in pkg.read_relationships(&pcd_part) {
            if rel_type_matches(&pcd_rel.rel_type, "pivotCacheRecords") && !pcd_rel.is_external() {
                let pcr_part = pcd_rel.resolved_target(&pcd_part);
                pkg.remove_part(&pcr_part);
                if ct.remove_override(&pcr_part) {
                    ct_changed = true;
                }
            }
        }
        pkg.remove_part(&rels_path_for(&pcd_part));
        pkg.remove_part(&pcd_part);
        if ct.remove_override(&pcd_part) {
            ct_changed = true;
        }
    }

    let wb_xml = pkg
        .get_str(workbook_part)
        .ok_or_else(|| format!("Missing workbook part '{workbook_part}'"))?;
    let mut wb_doc = XmlDocument::parse(&wb_xml)?;
    wb_doc.root.remove_children_by_local("pivotCaches");

    if !units.is_empty() {
        let mut pivot_caches_el = XmlElement::new(wb_doc.root.qualified_child_name("pivotCaches"));
        for (cache_id, unit) in units.into_iter().enumerate() {
            let (_, pcd_part) =
                allocate_free_numbered_part(pkg, "xl/pivotCache/pivotCacheDefinition", ".xml");
            let (_, pcr_part) =
                allocate_free_numbered_part(pkg, "xl/pivotCache/pivotCacheRecords", ".xml");
            let (_, pt_part) =
                allocate_free_numbered_part(pkg, "xl/pivotTables/pivotTable", ".xml");

            pkg.set_part(&pcd_part, unit.cache_definition_xml.into_bytes());
            pkg.set_part(&pcr_part, unit.cache_records_xml.into_bytes());
            pkg.set_part(&pt_part, unit.pivot_table_xml.into_bytes());

            ct.set_override(&pcd_part, CT_PIVOT_CACHE_DEF);
            ct.set_override(&pcr_part, CT_PIVOT_CACHE_RECORDS);
            ct.set_override(&pt_part, CT_PIVOT_TABLE);
            ct_changed = true;

            let pcr_rel_target = pcr_part.trim_start_matches("xl/pivotCache/").to_string();
            pkg.write_relationships(
                &pcd_part,
                &[Relationship {
                    id: "rId1".to_string(),
                    rel_type: REL_PIVOT_CACHE_RECORDS.to_string(),
                    target: pcr_rel_target,
                    target_mode: None,
                }],
            );

            let pt_to_pcd_target = format!("../{}", pcd_part.trim_start_matches("xl/"));
            pkg.write_relationships(
                &pt_part,
                &[Relationship {
                    id: "rId1".to_string(),
                    rel_type: REL_PIVOT_CACHE_DEF.to_string(),
                    target: pt_to_pcd_target,
                    target_mode: None,
                }],
            );

            let wb_rid = next_rid(&wb_rels);
            wb_rels.push(Relationship {
                id: wb_rid.clone(),
                rel_type: REL_PIVOT_CACHE_DEF.to_string(),
                target: pcd_part.trim_start_matches("xl/").to_string(),
                target_mode: None,
            });
            wb_rels_changed = true;

            let mut pc_el = XmlElement::new(pivot_caches_el.qualified_child_name("pivotCache"));
            pc_el.set_attr("cacheId", cache_id.to_string());
            pc_el.set_attr("r:id", wb_rid);
            pivot_caches_el.children.push(XmlNode::Element(pc_el));

            if let Some(ws_part) = sheet_part_by_id.get(&unit.dest_sheet_id) {
                let mut ws_rels = pkg.read_relationships(ws_part);
                let pt_rid = next_rid(&ws_rels);
                ws_rels.push(Relationship {
                    id: pt_rid,
                    rel_type: REL_PIVOT_TABLE.to_string(),
                    target: format!("../{}", pt_part.trim_start_matches("xl/")),
                    target_mode: None,
                });
                pkg.write_relationships(ws_part, &ws_rels);
            }
        }

        if wb_doc.root.attr("xmlns:r").is_none() {
            wb_doc.root.set_attr("xmlns:r", flavor.office_rels());
        }
        wb_doc
            .root
            .insert_child_in_schema_order(pivot_caches_el, WORKBOOK_SCHEMA_ORDER);
    }

    pkg.set_part(workbook_part, wb_doc.to_xml_string().into_bytes());
    if wb_rels_changed {
        pkg.write_relationships(workbook_part, &wb_rels);
    }
    if ct_changed {
        pkg.write_content_types(&ct);
    }
    Ok(())
}

fn remove_vba_signature_parts(
    pkg: &mut OpcPackage,
    ct: &mut crate::core::opc::ContentTypes,
    vba_part: &str,
) {
    for rel in pkg.read_relationships(vba_part) {
        if !rel.is_external() {
            let target = rel.resolved_target(vba_part);
            pkg.remove_part(&target);
            ct.remove_override(&target);
        }
    }
    pkg.remove_part(&rels_path_for(vba_part));
    if pkg.remove_part("xl/vbaProjectSignature.bin") {
        ct.remove_override("xl/vbaProjectSignature.bin");
    }
}

fn patch_vba(
    pkg: &mut OpcPackage,
    workbook_part: &str,
    sheets: &[Sheet],
    vba: Option<&VbaProject>,
) -> Result<(), String> {
    let mut wb_rels = pkg.read_relationships(workbook_part);
    let mut ct = pkg.read_content_types();

    match vba {
        Some(project) if !project.modules.is_empty() => {
            let new_vba_bin = crate::core::vba_xlsx::build_vba_project_bin(project)?;
            let existing_vba_part = wb_rels
                .iter()
                .find(|r| rel_type_matches(&r.rel_type, "vbaProject") && !r.is_external())
                .map(|r| r.resolved_target(workbook_part));
            let vba_part = existing_vba_part
                .clone()
                .unwrap_or_else(|| "xl/vbaProject.bin".to_string());

            remove_vba_signature_parts(pkg, &mut ct, &vba_part);
            pkg.set_part(&vba_part, new_vba_bin);

            ct.set_override(workbook_part, CT_WORKBOOK_MACRO_ENABLED);
            ct.set_override(&vba_part, CT_VBA_PROJECT);
            pkg.write_content_types(&ct);

            if existing_vba_part.is_none() {
                let rid = next_rid(&wb_rels);
                wb_rels.push(Relationship {
                    id: rid,
                    rel_type: REL_VBA_PROJECT.to_string(),
                    target: vba_part.trim_start_matches("xl/").to_string(),
                    target_mode: None,
                });
                pkg.write_relationships(workbook_part, &wb_rels);
            }

            if let Some(wb_xml) = pkg.get_str(workbook_part)
                && let Ok(mut wb_doc) = XmlDocument::parse(&wb_xml)
            {
                let wb_pr = wb_doc
                    .root
                    .ensure_child_in_schema_order("workbookPr", WORKBOOK_SCHEMA_ORDER);
                if wb_pr.attr("codeName").is_none() {
                    wb_pr.set_attr("codeName", "ThisWorkbook");
                }

                let sheet_name_by_id: HashMap<u64, &str> =
                    sheets.iter().map(|s| (s.id, s.name.as_str())).collect();
                let mut code_name_by_sheet_name: HashMap<&str, &str> = HashMap::new();
                for module in &project.modules {
                    if module.kind == VbaModuleKind::Document
                        && module.name != "ThisWorkbook"
                        && let Some(sheet_id) = module.bound_sheet_id
                        && let Some(&sheet_name) = sheet_name_by_id.get(&sheet_id)
                    {
                        code_name_by_sheet_name.insert(sheet_name, module.name.as_str());
                    }
                }

                if let Some(sheets_el) = wb_doc.root.child_mut("sheets") {
                    for node in &mut sheets_el.children {
                        if let XmlNode::Element(el) = node
                            && el.local_name() == "sheet"
                        {
                            if let Some(name) = el.attr("name")
                                && let Some(&code_name) = code_name_by_sheet_name.get(name)
                            {
                                el.set_attr("codeName", code_name);
                            } else {
                                el.remove_attr("codeName");
                            }
                        }
                    }
                }
                pkg.set_part(workbook_part, wb_doc.to_xml_string().into_bytes());
            }
        }
        _ => {
            let before = wb_rels.len();
            let mut old_vba_parts = Vec::new();
            wb_rels.retain(|r| {
                if rel_type_matches(&r.rel_type, "vbaProject") {
                    if !r.is_external() {
                        old_vba_parts.push(r.resolved_target(workbook_part));
                    }
                    false
                } else {
                    true
                }
            });
            if wb_rels.len() != before {
                pkg.write_relationships(workbook_part, &wb_rels);
            }
            if old_vba_parts.is_empty() && pkg.has_part("xl/vbaProject.bin") {
                old_vba_parts.push("xl/vbaProject.bin".to_string());
            }
            for vba_part in old_vba_parts {
                remove_vba_signature_parts(pkg, &mut ct, &vba_part);
                pkg.remove_part(&vba_part);
                ct.remove_override(&vba_part);
            }

            ct.set_override(workbook_part, CT_WORKBOOK);
            pkg.write_content_types(&ct);

            if let Some(wb_xml) = pkg.get_str(workbook_part)
                && let Ok(mut wb_doc) = XmlDocument::parse(&wb_xml)
            {
                if let Some(wb_pr) = wb_doc.root.child_mut("workbookPr") {
                    wb_pr.remove_attr("codeName");
                }
                if let Some(sheets_el) = wb_doc.root.child_mut("sheets") {
                    for node in &mut sheets_el.children {
                        if let XmlNode::Element(el) = node
                            && el.local_name() == "sheet"
                        {
                            el.remove_attr("codeName");
                        }
                    }
                }
                pkg.set_part(workbook_part, wb_doc.to_xml_string().into_bytes());
            }
        }
    }

    Ok(())
}

fn patch_freeze_pane(ws_root: &mut XmlElement, freeze: Option<FreezePane>) {
    let sheet_views = ws_root.ensure_child_in_schema_order("sheetViews", WORKSHEET_SCHEMA_ORDER);
    let sheet_view = if let Some(sv) = sheet_views.child_mut("sheetView") {
        sv
    } else {
        let mut sv = XmlElement::new(sheet_views.qualified_child_name("sheetView"));
        sv.set_attr("workbookViewId", "0");
        sheet_views.children.push(XmlNode::Element(sv));
        sheet_views.child_mut("sheetView").unwrap()
    };
    sheet_view.remove_children_by_local("pane");
    if let Some(pane) = freeze
        && (pane.rows > 0 || pane.cols > 0)
    {
        let top_left = format!("{}{}", col_idx_to_letters(pane.cols), pane.rows + 1);
        let active_pane = match (pane.rows > 0, pane.cols > 0) {
            (true, true) => "bottomRight",
            (true, false) => "bottomLeft",
            (false, true) => "topRight",
            (false, false) => "topLeft",
        };
        let mut pane_el = XmlElement::new(sheet_view.qualified_child_name("pane"));
        if pane.cols > 0 {
            pane_el.set_attr("xSplit", pane.cols.to_string());
        }
        if pane.rows > 0 {
            pane_el.set_attr("ySplit", pane.rows.to_string());
        }
        pane_el.set_attr("topLeftCell", top_left);
        pane_el.set_attr("activePane", active_pane);
        pane_el.set_attr("state", "frozen");
        sheet_view.children.insert(0, XmlNode::Element(pane_el));
    }
}

fn patch_merge_cells(ws_root: &mut XmlElement, merges: &[MergedCellRange]) {
    ws_root.remove_children_by_local("mergeCells");
    if merges.is_empty() {
        return;
    }
    let mut mc_el = XmlElement::new(ws_root.qualified_child_name("mergeCells"));
    mc_el.set_attr("count", merges.len().to_string());
    for m in merges {
        let mut item = XmlElement::new(mc_el.qualified_child_name("mergeCell"));
        item.set_attr(
            "ref",
            format!(
                "{}{}:{}{}",
                col_idx_to_letters(m.start_col),
                m.start_row + 1,
                col_idx_to_letters(m.end_col),
                m.end_row + 1
            ),
        );
        mc_el.children.push(XmlNode::Element(item));
    }
    ws_root.insert_child_in_schema_order(mc_el, WORKSHEET_SCHEMA_ORDER);
}

fn patch_cols(
    ws_root: &mut XmlElement,
    widths: &HashMap<usize, f64>,
    hidden_cols: &HashSet<usize>,
) {
    ws_root.remove_children_by_local("cols");
    let mut all_cols: Vec<usize> = widths
        .keys()
        .copied()
        .chain(hidden_cols.iter().copied())
        .collect();
    all_cols.sort_unstable();
    all_cols.dedup();
    if all_cols.is_empty() {
        return;
    }
    let mut cols_el = XmlElement::new(ws_root.qualified_child_name("cols"));
    for col_idx in all_cols {
        let mut col_el = XmlElement::new(cols_el.qualified_child_name("col"));
        let one_based = (col_idx + 1).to_string();
        col_el.set_attr("min", one_based.clone());
        col_el.set_attr("max", one_based);
        if let Some(&width) = widths.get(&col_idx) {
            col_el.set_attr("width", width.to_string());
            col_el.set_attr("customWidth", "1");
        }
        if hidden_cols.contains(&col_idx) {
            col_el.set_attr("hidden", "1");
        }
        cols_el.children.push(XmlNode::Element(col_el));
    }
    ws_root.insert_child_in_schema_order(cols_el, WORKSHEET_SCHEMA_ORDER);
}

fn patch_sheet_data(
    ws_root: &mut XmlElement,
    sheet: &Sheet,
    changed_cells: &[(usize, usize, Option<&CellSnapshot>)],
    changed_heights: &[(usize, Option<f64>)],
    changed_hidden_rows: &[(usize, bool)],
    sst: &mut SharedStringsPatcher,
    styles: &mut StylesPatcher,
) {
    let sheet_data = ws_root.ensure_child_in_schema_order("sheetData", WORKSHEET_SCHEMA_ORDER);

    materialize_affected_shared_formulas(sheet_data, changed_cells);

    for &(row_idx, height_opt) in changed_heights {
        let row_1based = row_idx + 1;
        let row_el = ensure_row_element(sheet_data, row_1based);
        match height_opt {
            Some(h) => {
                row_el.set_attr("ht", h.to_string());
                row_el.set_attr("customHeight", "1");
            }
            None => {
                row_el.remove_attr("ht");
                row_el.remove_attr("customHeight");
            }
        }
    }

    for &(row_idx, hidden) in changed_hidden_rows {
        let row_1based = row_idx + 1;
        let row_el = ensure_row_element(sheet_data, row_1based);
        if hidden {
            row_el.set_attr("hidden", "1");
        } else {
            row_el.remove_attr("hidden");
        }
    }

    for &(row_idx, col_idx, cell_opt) in changed_cells {
        let row_1based = row_idx + 1;
        let cell_ref = format!("{}{}", col_idx_to_letters(col_idx), row_1based);

        match cell_opt {
            None => {
                if let Some(row_el) = find_row_element_mut(sheet_data, row_1based) {
                    row_el.children.retain(|node| {
                        !matches!(node, XmlNode::Element(c) if c.local_name() == "c" && c.attr("r") == Some(cell_ref.as_str()))
                    });
                }
            }
            Some(cell) => {
                let col = &sheet.columns[col_idx];
                let row_el = ensure_row_element(sheet_data, row_1based);
                let new_c =
                    build_patched_cell_element(row_el, &cell_ref, cell, col, row_idx, sst, styles);
                upsert_cell_element(row_el, col_idx, new_c);
            }
        }
    }
}

fn materialize_affected_shared_formulas(
    sheet_data: &mut XmlElement,
    changed_cells: &[(usize, usize, Option<&CellSnapshot>)],
) {
    if changed_cells.is_empty() {
        return;
    }
    let changed_set: HashSet<(usize, usize)> =
        changed_cells.iter().map(|&(r, c, _)| (r, c)).collect();

    let mut masters: HashMap<String, (usize, usize, String)> = HashMap::new();
    let mut affected_sis: HashSet<String> = HashSet::new();

    for row_node in &sheet_data.children {
        let XmlNode::Element(row_el) = row_node else {
            continue;
        };
        if row_el.local_name() != "row" {
            continue;
        }
        for cell_node in &row_el.children {
            let XmlNode::Element(c_el) = cell_node else {
                continue;
            };
            if c_el.local_name() != "c" {
                continue;
            }
            let Some((r, c)) = c_el.attr("r").and_then(super::parse_a1_cell) else {
                continue;
            };
            if let Some(f_el) = c_el.child("f")
                && f_el.attr("t") == Some("shared")
                && let Some(si) = f_el.attr("si")
            {
                let body = f_el.text_content();
                if !body.is_empty() {
                    masters.insert(si.to_string(), (r, c, body));
                }
                if changed_set.contains(&(r, c)) {
                    affected_sis.insert(si.to_string());
                }
            }
        }
    }

    if affected_sis.is_empty() {
        return;
    }

    for row_node in &mut sheet_data.children {
        let XmlNode::Element(row_el) = row_node else {
            continue;
        };
        if row_el.local_name() != "row" {
            continue;
        }
        for cell_node in &mut row_el.children {
            let XmlNode::Element(c_el) = cell_node else {
                continue;
            };
            if c_el.local_name() != "c" {
                continue;
            }
            let Some((r, c)) = c_el.attr("r").and_then(super::parse_a1_cell) else {
                continue;
            };
            if changed_set.contains(&(r, c)) {
                continue;
            }
            if let Some(f_el) = c_el.child_mut("f")
                && f_el.attr("t") == Some("shared")
                && let Some(si) = f_el.attr("si").map(str::to_string)
                && affected_sis.contains(&si)
                && let Some((mr, mc, master_formula)) = masters.get(&si)
            {
                let d_row = r as isize - *mr as isize;
                let d_col = c as isize - *mc as isize;
                let expanded =
                    crate::core::xlsx::worksheet::shift_formula_refs(master_formula, d_row, d_col);
                f_el.remove_attr("t");
                f_el.remove_attr("ref");
                f_el.remove_attr("si");
                f_el.children.clear();
                f_el.children.push(XmlNode::Text(expanded));
            }
        }
    }
}

fn find_row_element_mut(sheet_data: &mut XmlElement, row_1based: usize) -> Option<&mut XmlElement> {
    let target = row_1based.to_string();
    sheet_data.children.iter_mut().find_map(|node| match node {
        XmlNode::Element(el)
            if el.local_name() == "row" && el.attr("r") == Some(target.as_str()) =>
        {
            Some(el)
        }
        _ => None,
    })
}

fn ensure_row_element(sheet_data: &mut XmlElement, row_1based: usize) -> &mut XmlElement {
    let mut insert_pos = sheet_data.children.len();
    for (idx, node) in sheet_data.children.iter().enumerate() {
        if let XmlNode::Element(el) = node
            && el.local_name() == "row"
            && let Some(r) = el.attr("r").and_then(|s| s.parse::<usize>().ok())
        {
            if r == row_1based {
                if let XmlNode::Element(ref mut found) = sheet_data.children[idx] {
                    return found;
                }
                unreachable!();
            }
            if r > row_1based {
                insert_pos = idx;
                break;
            }
        }
    }
    let mut row_el = XmlElement::new(sheet_data.qualified_child_name("row"));
    row_el.set_attr("r", row_1based.to_string());
    sheet_data
        .children
        .insert(insert_pos, XmlNode::Element(row_el));
    if let XmlNode::Element(ref mut el) = sheet_data.children[insert_pos] {
        el
    } else {
        unreachable!()
    }
}

fn upsert_cell_element(row_el: &mut XmlElement, target_col: usize, new_c: XmlElement) {
    let mut insert_pos = row_el.children.len();
    for (idx, node) in row_el.children.iter().enumerate() {
        if let XmlNode::Element(el) = node
            && el.local_name() == "c"
            && let Some(r_attr) = el.attr("r")
            && let Some((_, c)) = super::parse_a1_cell(r_attr)
        {
            if c == target_col {
                row_el.children[idx] = XmlNode::Element(new_c);
                return;
            }
            if c > target_col {
                insert_pos = idx;
                break;
            }
        }
    }
    row_el.children.insert(insert_pos, XmlNode::Element(new_c));
}

fn build_patched_cell_element(
    row_el: &XmlElement,
    cell_ref: &str,
    cell: &CellSnapshot,
    col: &crate::core::engine::DataColumn,
    row_idx: usize,
    sst: &mut SharedStringsPatcher,
    styles: &mut StylesPatcher,
) -> XmlElement {
    let mut c_el = XmlElement::new(row_el.qualified_child_name("c"));
    c_el.set_attr("r", cell_ref);

    if let Some(ref style) = cell.style {
        let xf_idx = styles.append_style(style);
        if xf_idx != 0 {
            c_el.set_attr("s", xf_idx.to_string());
        }
    }

    if let Some(formula_src) = cell.src.strip_prefix('=') {
        let encoded = crate::core::xlsx::formula::encode_xlsx_formula(formula_src);
        let mut f_el = XmlElement::new(c_el.qualified_child_name("f"));
        f_el.children.push(XmlNode::Text(encoded));
        c_el.children.push(XmlNode::Element(f_el));

        match &cell.res {
            ResultData::String(s) => {
                c_el.set_attr("t", "str");
                let mut v_el = XmlElement::new(c_el.qualified_child_name("v"));
                v_el.children.push(XmlNode::Text(encode_xstring(s)));
                c_el.children.push(XmlNode::Element(v_el));
            }
            ResultData::Boolean(b) => {
                c_el.set_attr("t", "b");
                let mut v_el = XmlElement::new(c_el.qualified_child_name("v"));
                v_el.children
                    .push(XmlNode::Text(if *b { "1" } else { "0" }.to_string()));
                c_el.children.push(XmlNode::Element(v_el));
            }
            ResultData::Error(e) => {
                c_el.set_attr("t", "e");
                let mut v_el = XmlElement::new(c_el.qualified_child_name("v"));
                v_el.children.push(XmlNode::Text(
                    super::normalize_ooxml_error_literal(e).to_string(),
                ));
                c_el.children.push(XmlNode::Element(v_el));
            }
            ResultData::Integer(i) => {
                let mut v_el = XmlElement::new(c_el.qualified_child_name("v"));
                v_el.children.push(XmlNode::Text(i.to_string()));
                c_el.children.push(XmlNode::Element(v_el));
            }
            ResultData::Float(f) if f.is_finite() => {
                let mut v_el = XmlElement::new(c_el.qualified_child_name("v"));
                v_el.children.push(XmlNode::Text(f.to_string()));
                c_el.children.push(XmlNode::Element(v_el));
            }
            _ => {}
        }
    } else if cell.cell_type == CellType::String {
        let text = if cell.src.starts_with('"') && cell.src.ends_with('"') && cell.src.len() >= 2 {
            &cell.src[1..cell.src.len() - 1]
        } else {
            cell.src.as_str()
        };
        let idx = sst.intern(text);
        c_el.set_attr("t", "s");
        let mut v_el = XmlElement::new(c_el.qualified_child_name("v"));
        v_el.children.push(XmlNode::Text(idx.to_string()));
        c_el.children.push(XmlNode::Element(v_el));
    } else if let Some(serial) = super::date_serial_for_export(cell.style.as_ref(), col, row_idx) {
        let mut v_el = XmlElement::new(c_el.qualified_child_name("v"));
        v_el.children.push(XmlNode::Text(serial.to_string()));
        c_el.children.push(XmlNode::Element(v_el));
    } else if let Ok(num) = cell.src.parse::<f64>()
        && num.is_finite()
    {
        let mut v_el = XmlElement::new(c_el.qualified_child_name("v"));
        v_el.children.push(XmlNode::Text(num.to_string()));
        c_el.children.push(XmlNode::Element(v_el));
    } else if let Ok(b) = cell.src.parse::<bool>() {
        c_el.set_attr("t", "b");
        let mut v_el = XmlElement::new(c_el.qualified_child_name("v"));
        v_el.children
            .push(XmlNode::Text(if b { "1" } else { "0" }.to_string()));
        c_el.children.push(XmlNode::Element(v_el));
    } else if cell.cell_type == CellType::Error && !cell.src.is_empty() {
        c_el.set_attr("t", "e");
        let mut v_el = XmlElement::new(c_el.qualified_child_name("v"));
        v_el.children.push(XmlNode::Text(cell.src.clone()));
        c_el.children.push(XmlNode::Element(v_el));
    } else if !cell.src.is_empty() {
        let text = if cell.src.starts_with('"') && cell.src.ends_with('"') && cell.src.len() >= 2 {
            &cell.src[1..cell.src.len() - 1]
        } else {
            cell.src.as_str()
        };
        let idx = sst.intern(text);
        c_el.set_attr("t", "s");
        let mut v_el = XmlElement::new(c_el.qualified_child_name("v"));
        v_el.children.push(XmlNode::Text(idx.to_string()));
        c_el.children.push(XmlNode::Element(v_el));
    }

    c_el
}
