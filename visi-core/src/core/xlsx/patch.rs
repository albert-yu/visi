use std::collections::{HashMap, HashSet};

use crate::core::chart::Chart;
use crate::core::engine::{
    CellType, FreezePane, MergedCellRange, ResultData, Sheet, SheetVisibility,
};
use crate::core::opc::{
    CT_SHARED_STRINGS, OpcPackage, REL_SHARED_STRINGS, Relationship, next_rid, rel_type_matches,
};
use crate::core::parser::col_idx_to_letters;
use crate::core::pivot::PivotTable;
use crate::core::style::CellStyle;
use crate::core::table::ExcelTable;
use crate::core::vba::VbaProject;
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

    pub(crate) fn can_patch_surgically(
        &self,
        sheets: &[Sheet],
        charts: &[Chart],
        pivots: &[PivotTable],
        vba: Option<&VbaProject>,
    ) -> bool {
        if sheets.len() != self.sheet_snapshots.len() {
            return false;
        }
        for (sheet, snap) in sheets.iter().zip(self.sheet_snapshots.iter()) {
            if sheet.id != snap.sheet_id
                || sheet.tables != snap.tables
                || sheet.comments != snap.comments
                || sheet.hyperlinks != snap.hyperlinks
                || sheet.data_validations != snap.data_validations
                || sheet.conditional_formats != snap.conditional_formats
            {
                return false;
            }
        }
        charts == self.charts.as_slice()
            && pivots == self.pivots.as_slice()
            && vba == self.vba.as_ref()
    }

    pub(crate) fn patch_to_bytes(
        &self,
        sheets: &[Sheet],
        defined_names: &[crate::core::DefinedName],
    ) -> Result<Vec<u8>, String> {
        let mut pkg = self.package.clone();

        let cur_date1904 = sheets.iter().any(|s| s.date1904);
        let defined_names_changed = defined_names != self.defined_names.as_slice();
        let workbook_meta_changed = cur_date1904 != self.date1904
            || defined_names_changed
            || sheets
                .iter()
                .zip(self.sheet_snapshots.iter())
                .any(|(s, snap)| s.name != snap.name || s.visibility != snap.visibility);
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
                let mut sheet_idx = 0usize;
                for node in &mut sheets_el.children {
                    if let XmlNode::Element(el) = node
                        && el.local_name() == "sheet"
                    {
                        if let Some(sheet) = sheets.get(sheet_idx) {
                            el.set_attr("name", sheet.name.clone());
                            match crate::core::xlsx::workbook::sheet_visibility_attr(
                                sheet.visibility,
                            ) {
                                Some(state) => el.set_attr("state", state),
                                None => {
                                    el.remove_attr("state");
                                }
                            }
                        }
                        sheet_idx += 1;
                    }
                }
            }
            pkg.set_part(&self.workbook_part, doc.to_xml_string().into_bytes());
        }

        let mut sst_patcher: Option<SharedStringsPatcher> = None;
        let mut styles_patcher: Option<StylesPatcher> = None;

        for (sheet, snap) in sheets.iter().zip(self.sheet_snapshots.iter()) {
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

            if !cells_changed
                && !cols_changed
                && !heights_changed
                && !hidden_rows_changed
                && !merges_changed
                && !freeze_changed
            {
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

            if cells_changed || heights_changed || hidden_rows_changed {
                let mut changed_cells: Vec<(usize, usize, Option<&CellSnapshot>)> = Vec::new();
                for (&rc, cur) in &cur_cells {
                    if snap.cells.get(&rc) != Some(cur) {
                        changed_cells.push((rc.0, rc.1, Some(cur)));
                    }
                }
                for &rc in snap.cells.keys() {
                    if !cur_cells.contains_key(&rc) {
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

        pkg.to_bytes()
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
