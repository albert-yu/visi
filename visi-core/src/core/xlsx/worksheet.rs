use std::collections::{HashMap, HashSet};

use quick_xml::events::Event;

use crate::core::engine::{CellType, FreezePane, MergedCellRange, ResultData};
use crate::core::parser::{col_idx_to_letters, col_letters_to_idx};
use crate::core::xlsx::styles::ParsedStyles;
use crate::core::xlsx::xml_util::decode_xstring;

#[derive(Debug, Clone)]
pub(crate) enum ParsedCellValue {
    Empty,
    Number(f64),
    DateTime(f64),
    DateTimeIso(String),
    String(String),
    Bool(bool),
    Error(String),
}

impl ParsedCellValue {
    pub(crate) fn to_result_data(&self) -> ResultData {
        match self {
            Self::Empty => ResultData::None,
            Self::Number(f) | Self::DateTime(f) => ResultData::Float(*f),
            Self::DateTimeIso(s) | Self::String(s) => ResultData::String(s.clone()),
            Self::Bool(b) => ResultData::Boolean(*b),
            Self::Error(e) => ResultData::Error(e.clone()),
        }
    }

    pub(crate) fn to_cell_type_and_src(&self) -> (CellType, String) {
        match self {
            Self::Empty => (CellType::Empty, String::new()),
            Self::Number(f) => (CellType::Float, f.to_string()),
            Self::DateTime(f) => (CellType::DateTime, f.to_string()),
            Self::DateTimeIso(s) => (CellType::DateTimeIso, s.clone()),
            Self::String(s) => (CellType::String, s.clone()),
            Self::Bool(b) => (CellType::Bool, b.to_string()),
            Self::Error(e) => (CellType::Error, e.clone()),
        }
    }
}

#[derive(Debug, Clone)]
pub(crate) struct ParsedCell {
    pub(crate) row: usize,
    pub(crate) col: usize,
    pub(crate) value: ParsedCellValue,
    pub(crate) formula: Option<String>,
    pub(crate) style_idx: Option<u32>,
}

#[derive(Debug, Default)]
pub(crate) struct ParsedWorksheet {
    pub(crate) cells: Vec<ParsedCell>,
    pub(crate) column_widths: HashMap<usize, f64>,
    pub(crate) hidden_cols: HashSet<usize>,
    pub(crate) row_heights: HashMap<usize, f64>,
    pub(crate) hidden_rows: HashSet<usize>,
    pub(crate) merged_cells: Vec<MergedCellRange>,
    pub(crate) freeze_pane: Option<FreezePane>,
    pub(crate) rows: usize,
    pub(crate) cols: usize,
}

#[derive(Debug, Clone)]
struct SharedFormulaMaster {
    row: usize,
    col: usize,
    formula: String,
}

pub(crate) fn parse_worksheet(
    xml: &str,
    shared_strings: &[String],
    styles: &ParsedStyles,
) -> Result<ParsedWorksheet, String> {
    let mut out = ParsedWorksheet::default();
    let mut shared_formulas: HashMap<u32, SharedFormulaMaster> = HashMap::new();

    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buf = Vec::new();

    let mut next_row = 0usize;
    let mut current_row = 0usize;
    let mut next_col = 0usize;

    let mut in_cell = false;
    let mut cell_row = 0usize;
    let mut cell_col = 0usize;
    let mut cell_type: Option<String> = None;
    let mut cell_style: Option<u32> = None;
    let mut cell_val_text = String::new();
    let mut cell_has_v = false;
    let mut cell_inline_text = String::new();
    let mut cell_has_is = false;
    let mut cell_formula: Option<String> = None;

    let mut in_v = false;
    let mut in_f = false;
    let mut f_type: Option<String> = None;
    let mut f_si: Option<u32> = None;
    let mut f_text = String::new();

    let mut in_is = false;
    let mut in_is_t = false;
    let mut phonetic_depth = 0usize;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"pane" => {
                    parse_freeze_pane(e, &mut out.freeze_pane);
                }
                b"mergeCell" => {
                    parse_merge_cell(e, &mut out.merged_cells);
                }
                b"col" => {
                    parse_col_props(e, &mut out.column_widths, &mut out.hidden_cols);
                }
                b"row" => {
                    let row_1based = super::get_attr(e, b"r")
                        .and_then(|s| s.parse::<usize>().ok())
                        .unwrap_or(next_row + 1)
                        .max(1);
                    current_row = row_1based - 1;
                    next_row = current_row + 1;
                    next_col = 0;
                    if let Some(height) =
                        super::get_attr(e, b"ht").and_then(|s| s.parse::<f64>().ok())
                        && row_1based <= 1_048_576
                        && height.is_finite()
                        && height >= 0.0
                    {
                        out.row_heights.insert(current_row, height);
                    }
                    if row_1based <= 1_048_576 && is_xml_truthy(super::get_attr(e, b"hidden")) {
                        out.hidden_rows.insert(current_row);
                    }
                }
                b"c" => {
                    in_cell = true;
                    if let Some(r_attr) = super::get_attr(e, b"r")
                        && let Some((r, c)) = super::parse_a1_cell(&r_attr)
                    {
                        cell_row = r;
                        cell_col = c;
                    } else {
                        cell_row = current_row;
                        cell_col = next_col;
                    }
                    next_col = cell_col + 1;
                    cell_type = super::get_attr(e, b"t");
                    cell_style = super::get_attr(e, b"s").and_then(|s| s.parse::<u32>().ok());
                    cell_val_text.clear();
                    cell_has_v = false;
                    cell_inline_text.clear();
                    cell_has_is = false;
                    cell_formula = None;
                }
                b"v" if in_cell => {
                    in_v = true;
                    cell_has_v = true;
                    cell_val_text.clear();
                }
                b"f" if in_cell => {
                    in_f = true;
                    f_type = super::get_attr(e, b"t");
                    f_si = super::get_attr(e, b"si").and_then(|s| s.parse::<u32>().ok());
                    f_text.clear();
                }
                b"is" if in_cell => {
                    in_is = true;
                    cell_has_is = true;
                    cell_inline_text.clear();
                }
                b"rPh" if in_is => {
                    phonetic_depth += 1;
                }
                b"t" if in_is && phonetic_depth == 0 => {
                    in_is_t = true;
                }
                _ => {}
            },
            Ok(Event::Empty(ref e)) => match e.local_name().as_ref() {
                b"pane" => {
                    parse_freeze_pane(e, &mut out.freeze_pane);
                }
                b"mergeCell" => {
                    parse_merge_cell(e, &mut out.merged_cells);
                }
                b"col" => {
                    parse_col_props(e, &mut out.column_widths, &mut out.hidden_cols);
                }
                b"row" => {
                    let row_1based = super::get_attr(e, b"r")
                        .and_then(|s| s.parse::<usize>().ok())
                        .unwrap_or(next_row + 1)
                        .max(1);
                    current_row = row_1based - 1;
                    next_row = current_row + 1;
                    next_col = 0;
                    if let Some(height) =
                        super::get_attr(e, b"ht").and_then(|s| s.parse::<f64>().ok())
                        && row_1based <= 1_048_576
                        && height.is_finite()
                        && height >= 0.0
                    {
                        out.row_heights.insert(current_row, height);
                    }
                    if row_1based <= 1_048_576 && is_xml_truthy(super::get_attr(e, b"hidden")) {
                        out.hidden_rows.insert(current_row);
                    }
                }
                b"c" => {
                    let (r, c) = if let Some(r_attr) = super::get_attr(e, b"r")
                        && let Some(rc) = super::parse_a1_cell(&r_attr)
                    {
                        rc
                    } else {
                        (current_row, next_col)
                    };
                    next_col = c + 1;
                    let s = super::get_attr(e, b"s").and_then(|v| v.parse::<u32>().ok());
                    if s.is_some() {
                        out.rows = out.rows.max(r + 1);
                        out.cols = out.cols.max(c + 1);
                        out.cells.push(ParsedCell {
                            row: r,
                            col: c,
                            value: ParsedCellValue::Empty,
                            formula: None,
                            style_idx: s,
                        });
                    }
                }
                b"f" if in_cell => {
                    let t = super::get_attr(e, b"t");
                    let si = super::get_attr(e, b"si").and_then(|s| s.parse::<u32>().ok());
                    if t.as_deref() == Some("shared")
                        && let Some(si_val) = si
                        && let Some(master) = shared_formulas.get(&si_val)
                    {
                        let d_row = cell_row as isize - master.row as isize;
                        let d_col = cell_col as isize - master.col as isize;
                        cell_formula = Some(shift_formula_refs(&master.formula, d_row, d_col));
                    }
                }
                b"v" if in_cell => {
                    cell_has_v = true;
                    cell_val_text.clear();
                }
                b"is" if in_cell => {
                    cell_has_is = true;
                    cell_inline_text.clear();
                }
                _ => {}
            },
            Ok(Event::Text(e)) => {
                if in_v || in_f || (in_is_t && phonetic_depth == 0) {
                    let decoded = e
                        .decode()
                        .map_err(|err| format!("worksheet XML text decode error: {err}"))?;
                    let unescaped = quick_xml::escape::unescape(&decoded)
                        .map_err(|err| format!("worksheet XML text unescape error: {err}"))?;
                    if in_v {
                        cell_val_text.push_str(&unescaped);
                    } else if in_f {
                        f_text.push_str(&unescaped);
                    } else if in_is_t {
                        cell_inline_text.push_str(&unescaped);
                    }
                }
            }
            Ok(Event::GeneralRef(e)) => {
                if in_v || in_f || (in_is_t && phonetic_depth == 0) {
                    let raw = e
                        .decode()
                        .map_err(|err| format!("worksheet XML entity decode error: {err}"))?;
                    let entity = format!("&{raw};");
                    let resolved = quick_xml::escape::unescape(&entity)
                        .map(|v| v.into_owned())
                        .unwrap_or(entity);
                    if in_v {
                        cell_val_text.push_str(&resolved);
                    } else if in_f {
                        f_text.push_str(&resolved);
                    } else if in_is_t {
                        cell_inline_text.push_str(&resolved);
                    }
                }
            }
            Ok(Event::CData(e)) => {
                if in_v || in_f || (in_is_t && phonetic_depth == 0) {
                    let text = String::from_utf8_lossy(e.as_ref());
                    if in_v {
                        cell_val_text.push_str(&text);
                    } else if in_f {
                        f_text.push_str(&text);
                    } else if in_is_t {
                        cell_inline_text.push_str(&text);
                    }
                }
            }
            Ok(Event::End(ref e)) => match e.local_name().as_ref() {
                b"v" => {
                    in_v = false;
                }
                b"f" => {
                    in_f = false;
                    let formula_body = f_text.strip_prefix('=').unwrap_or(&f_text).to_string();
                    if f_type.as_deref() == Some("shared") {
                        if let Some(si_val) = f_si {
                            if !formula_body.is_empty() {
                                shared_formulas.insert(
                                    si_val,
                                    SharedFormulaMaster {
                                        row: cell_row,
                                        col: cell_col,
                                        formula: formula_body.clone(),
                                    },
                                );
                                cell_formula = Some(formula_body);
                            } else if let Some(master) = shared_formulas.get(&si_val) {
                                let d_row = cell_row as isize - master.row as isize;
                                let d_col = cell_col as isize - master.col as isize;
                                cell_formula =
                                    Some(shift_formula_refs(&master.formula, d_row, d_col));
                            }
                        }
                    } else if !formula_body.is_empty() {
                        cell_formula = Some(formula_body);
                    }
                    f_type = None;
                    f_si = None;
                    f_text.clear();
                }
                b"rPh" if phonetic_depth > 0 => {
                    phonetic_depth -= 1;
                }
                b"t" if in_is => {
                    in_is_t = false;
                }
                b"is" => {
                    in_is = false;
                    in_is_t = false;
                    phonetic_depth = 0;
                }
                b"c" => {
                    in_cell = false;
                    let value = decode_cell_value(
                        cell_type.as_deref(),
                        cell_style,
                        cell_has_v,
                        &cell_val_text,
                        cell_has_is,
                        &cell_inline_text,
                        shared_strings,
                        styles,
                    );
                    if !matches!(value, ParsedCellValue::Empty)
                        || cell_formula.is_some()
                        || cell_style.is_some()
                    {
                        out.rows = out.rows.max(cell_row + 1);
                        out.cols = out.cols.max(cell_col + 1);
                        out.cells.push(ParsedCell {
                            row: cell_row,
                            col: cell_col,
                            value,
                            formula: cell_formula.take(),
                            style_idx: cell_style,
                        });
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(err) => return Err(format!("Failed to parse worksheet XML: {err}")),
            _ => {}
        }
        buf.clear();
    }

    if let Some(&max_w_col) = out.column_widths.keys().max() {
        out.cols = out.cols.max(max_w_col + 1);
    }
    if let Some(&max_h_col) = out.hidden_cols.iter().max() {
        out.cols = out.cols.max(max_h_col + 1);
    }
    if let Some(&max_h_row) = out.row_heights.keys().max() {
        out.rows = out.rows.max(max_h_row + 1);
    }
    if let Some(&max_hid_row) = out.hidden_rows.iter().max() {
        out.rows = out.rows.max(max_hid_row + 1);
    }
    for m in &out.merged_cells {
        out.rows = out.rows.max(m.end_row + 1);
        out.cols = out.cols.max(m.end_col + 1);
    }

    Ok(out)
}

pub(crate) fn write_worksheet_xml(
    sheet: &crate::core::Sheet,
    sst: &mut crate::core::xlsx::shared_strings::SharedStringTable,
    styles: &mut crate::core::xlsx::styles::StyleRegistry,
    table_rids: &[String],
    drawing_rid: Option<&str>,
    flavor: crate::core::xlsx::xml_util::NamespaceFlavor,
) -> String {
    let ns = flavor.spreadsheetml();
    let ns_r = flavor.office_rels();
    let row_count = sheet.row_count();
    let col_count = sheet.columns.len();

    let mut max_row = 0usize;
    let mut max_col = 0usize;
    let mut has_any_cell = false;

    let mut sheet_data_xml = String::new();
    for row_idx in 0..row_count {
        let row_height = sheet
            .row_heights
            .get(row_idx)
            .copied()
            .flatten()
            .filter(|h| h.is_finite() && *h >= 0.0);

        let mut row_cells_xml = String::new();
        for col_idx in 0..col_count {
            let col = &sheet.columns[col_idx];
            let cell_src = col.src.get(row_idx).map(String::as_str).unwrap_or("");
            let cell_type = col
                .cell_types
                .get(row_idx)
                .copied()
                .unwrap_or(CellType::Empty);
            let style_opt = col.styles.get(row_idx).and_then(|s| s.as_ref());
            let xf_idx = style_opt.map(|s| styles.register_style(s)).unwrap_or(0);

            let cell_ref = format!("{}{}", col_idx_to_letters(col_idx), row_idx + 1);
            let style_attr = if xf_idx != 0 {
                format!(" s=\"{xf_idx}\"")
            } else {
                String::new()
            };

            if let Some(formula_src) = cell_src.strip_prefix('=') {
                let encoded_formula =
                    crate::core::parser::encode_xlsx_spill_references(formula_src);
                let res_opt = col.data.get(row_idx);
                let (t_attr, v_xml) = match res_opt {
                    Some(ResultData::String(s)) => (
                        " t=\"str\"",
                        format!(
                            "<v>{}</v>",
                            super::escape_xml_text(&crate::core::xlsx::xml_util::encode_xstring(
                                &s
                            ))
                        ),
                    ),
                    Some(ResultData::Boolean(b)) => {
                        (" t=\"b\"", format!("<v>{}</v>", if b { "1" } else { "0" }))
                    }
                    Some(ResultData::Error(e)) => {
                        (" t=\"e\"", format!("<v>{}</v>", super::escape_xml_text(&e)))
                    }
                    Some(ResultData::Integer(i)) => ("", format!("<v>{i}</v>")),
                    Some(ResultData::Float(f)) if f.is_finite() => ("", format!("<v>{f}</v>")),
                    Some(ResultData::None) | None => ("", String::new()),
                    Some(other) => {
                        let s = super::format_result_for_xlsx(&other);
                        if s.is_empty() {
                            ("", String::new())
                        } else {
                            (
                                " t=\"str\"",
                                format!(
                                    "<v>{}</v>",
                                    super::escape_xml_text(
                                        &crate::core::xlsx::xml_util::encode_xstring(&s)
                                    )
                                ),
                            )
                        }
                    }
                };
                row_cells_xml.push_str(&format!(
                    "<c r=\"{cell_ref}\"{style_attr}{t_attr}><f>{}</f>{v_xml}</c>",
                    super::escape_xml_text(&encoded_formula)
                ));
                has_any_cell = true;
                max_row = max_row.max(row_idx);
                max_col = max_col.max(col_idx);
            } else if cell_type == CellType::String {
                let text = if cell_src.starts_with('"')
                    && cell_src.ends_with('"')
                    && cell_src.len() >= 2
                {
                    &cell_src[1..cell_src.len() - 1]
                } else {
                    cell_src
                };
                let sst_idx = sst.intern(text);
                row_cells_xml.push_str(&format!(
                    "<c r=\"{cell_ref}\"{style_attr} t=\"s\"><v>{sst_idx}</v></c>"
                ));
                has_any_cell = true;
                max_row = max_row.max(row_idx);
                max_col = max_col.max(col_idx);
            } else if let Some(serial) = super::date_serial_for_export(style_opt, col, row_idx) {
                row_cells_xml.push_str(&format!(
                    "<c r=\"{cell_ref}\"{style_attr}><v>{serial}</v></c>"
                ));
                has_any_cell = true;
                max_row = max_row.max(row_idx);
                max_col = max_col.max(col_idx);
            } else if let Ok(val_f64) = cell_src.parse::<f64>()
                && val_f64.is_finite()
            {
                row_cells_xml.push_str(&format!(
                    "<c r=\"{cell_ref}\"{style_attr}><v>{val_f64}</v></c>"
                ));
                has_any_cell = true;
                max_row = max_row.max(row_idx);
                max_col = max_col.max(col_idx);
            } else if let Ok(val_bool) = cell_src.parse::<bool>() {
                row_cells_xml.push_str(&format!(
                    "<c r=\"{cell_ref}\"{style_attr} t=\"b\"><v>{}</v></c>",
                    if val_bool { "1" } else { "0" }
                ));
                has_any_cell = true;
                max_row = max_row.max(row_idx);
                max_col = max_col.max(col_idx);
            } else if cell_type == CellType::Error && !cell_src.is_empty() {
                row_cells_xml.push_str(&format!(
                    "<c r=\"{cell_ref}\"{style_attr} t=\"e\"><v>{}</v></c>",
                    super::escape_xml(cell_src)
                ));
                has_any_cell = true;
                max_row = max_row.max(row_idx);
                max_col = max_col.max(col_idx);
            } else if !cell_src.is_empty() {
                let text = if cell_src.starts_with('"')
                    && cell_src.ends_with('"')
                    && cell_src.len() >= 2
                {
                    &cell_src[1..cell_src.len() - 1]
                } else {
                    cell_src
                };
                let sst_idx = sst.intern(text);
                row_cells_xml.push_str(&format!(
                    "<c r=\"{cell_ref}\"{style_attr} t=\"s\"><v>{sst_idx}</v></c>"
                ));
                has_any_cell = true;
                max_row = max_row.max(row_idx);
                max_col = max_col.max(col_idx);
            } else if xf_idx != 0 {
                row_cells_xml.push_str(&format!("<c r=\"{cell_ref}\"{style_attr}/>"));
                has_any_cell = true;
                max_row = max_row.max(row_idx);
                max_col = max_col.max(col_idx);
            }
        }

        let row_hidden = sheet.hidden_rows.get(row_idx).copied().unwrap_or(false);

        if !row_cells_xml.is_empty() || row_height.is_some() || row_hidden {
            sheet_data_xml.push_str(&format!("<row r=\"{}\"", row_idx + 1));
            if let Some(h) = row_height {
                sheet_data_xml.push_str(&format!(" ht=\"{h}\" customHeight=\"1\""));
            }
            if row_hidden {
                sheet_data_xml.push_str(" hidden=\"1\"");
            }
            if row_cells_xml.is_empty() {
                sheet_data_xml.push_str("/>");
            } else {
                sheet_data_xml.push('>');
                sheet_data_xml.push_str(&row_cells_xml);
                sheet_data_xml.push_str("</row>");
            }
        }
    }

    let dim_ref = if has_any_cell {
        format!("A1:{}{}", col_idx_to_letters(max_col), max_row + 1)
    } else {
        "A1".to_string()
    };

    let mut xml = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"{ns}\" xmlns:r=\"{ns_r}\"><dimension ref=\"{dim_ref}\"/><sheetViews><sheetView workbookViewId=\"0\""
    );
    if let Some(pane) = sheet.freeze_pane
        && (pane.rows > 0 || pane.cols > 0)
    {
        xml.push('>');
        xml.push_str(&format_freeze_pane_xml(pane));
        xml.push_str("</sheetView>");
    } else {
        xml.push_str("/>");
    }
    xml.push_str("</sheetViews><sheetFormatPr defaultRowHeight=\"15\"/>");

    let explicit_cols: Vec<(usize, Option<f64>, bool)> = sheet
        .columns
        .iter()
        .enumerate()
        .filter_map(|(idx, col)| {
            let width = col.width.filter(|w| w.is_finite() && *w >= 0.0);
            if width.is_some() || col.hidden {
                Some((idx, width, col.hidden))
            } else {
                None
            }
        })
        .collect();
    if !explicit_cols.is_empty() {
        xml.push_str("<cols>");
        for (col_idx, width, hidden) in explicit_cols {
            let one_based = col_idx + 1;
            xml.push_str(&format!("<col min=\"{one_based}\" max=\"{one_based}\""));
            if let Some(w) = width {
                xml.push_str(&format!(" width=\"{w}\" customWidth=\"1\""));
            }
            if hidden {
                xml.push_str(" hidden=\"1\"");
            }
            xml.push_str("/>");
        }
        xml.push_str("</cols>");
    }

    if sheet_data_xml.is_empty() {
        xml.push_str("<sheetData/>");
    } else {
        xml.push_str("<sheetData>");
        xml.push_str(&sheet_data_xml);
        xml.push_str("</sheetData>");
    }

    if !sheet.merged_cells.is_empty() {
        xml.push_str(&format!(
            "<mergeCells count=\"{}\">",
            sheet.merged_cells.len()
        ));
        for m in &sheet.merged_cells {
            let range_ref = format!(
                "{}{}:{}{}",
                col_idx_to_letters(m.start_col),
                m.start_row + 1,
                col_idx_to_letters(m.end_col),
                m.end_row + 1
            );
            xml.push_str(&format!("<mergeCell ref=\"{range_ref}\"/>"));
        }
        xml.push_str("</mergeCells>");
    }

    if let Some(r_id) = drawing_rid {
        xml.push_str(&format!("<drawing r:id=\"{}\"/>", super::escape_xml(r_id)));
    }

    if !table_rids.is_empty() {
        xml.push_str(&format!("<tableParts count=\"{}\">", table_rids.len()));
        for r_id in table_rids {
            xml.push_str(&format!(
                "<tablePart r:id=\"{}\"/>",
                super::escape_xml(r_id)
            ));
        }
        xml.push_str("</tableParts>");
    }

    xml.push_str("</worksheet>");
    xml
}

fn is_xml_truthy(val: Option<String>) -> bool {
    val.is_some_and(|s| s == "1" || s.eq_ignore_ascii_case("true"))
}

fn parse_col_props(
    e: &quick_xml::events::BytesStart,
    widths: &mut HashMap<usize, f64>,
    hidden_cols: &mut HashSet<usize>,
) {
    let min = super::get_attr(e, b"min")
        .and_then(|s| s.parse::<usize>().ok())
        .unwrap_or(1)
        .max(1);
    let max = super::get_attr(e, b"max")
        .and_then(|s| s.parse::<usize>().ok())
        .unwrap_or(min)
        .max(min)
        .min(16_384);

    if let Some(width) = super::get_attr(e, b"width").and_then(|s| s.parse::<f64>().ok())
        && width.is_finite()
        && width >= 0.0
    {
        for one_based_col in min..=max {
            widths.insert(one_based_col - 1, width);
        }
    }
    if is_xml_truthy(super::get_attr(e, b"hidden")) {
        for one_based_col in min..=max {
            hidden_cols.insert(one_based_col - 1);
        }
    }
}

fn parse_merge_cell(e: &quick_xml::events::BytesStart, out: &mut Vec<MergedCellRange>) {
    if let Some(ref_str) = super::get_attr(e, b"ref")
        && let Some((start_row, start_col, end_row, end_col)) =
            crate::core::pivot_xlsx::parse_a1_range(&ref_str)
    {
        let range = MergedCellRange {
            start_row,
            start_col,
            end_row,
            end_col,
        };
        if !out.contains(&range) {
            out.push(range);
        }
    }
}

fn parse_freeze_pane(e: &quick_xml::events::BytesStart, out: &mut Option<FreezePane>) {
    let state = super::get_attr(e, b"state");
    if state
        .as_deref()
        .is_some_and(|s| s.eq_ignore_ascii_case("frozen") || s.eq_ignore_ascii_case("frozenSplit"))
    {
        let cols = super::get_attr(e, b"xSplit")
            .and_then(|s| s.parse::<f64>().ok())
            .map(|v| v.max(0.0).round() as usize)
            .unwrap_or(0);
        let rows = super::get_attr(e, b"ySplit")
            .and_then(|s| s.parse::<f64>().ok())
            .map(|v| v.max(0.0).round() as usize)
            .unwrap_or(0);
        if rows > 0 || cols > 0 {
            *out = Some(FreezePane { rows, cols });
        }
    }
}

pub(crate) fn format_freeze_pane_xml(pane: FreezePane) -> String {
    let top_left = format!("{}{}", col_idx_to_letters(pane.cols), pane.rows + 1);
    let active_pane = match (pane.rows > 0, pane.cols > 0) {
        (true, true) => "bottomRight",
        (true, false) => "bottomLeft",
        (false, true) => "topRight",
        (false, false) => "topLeft",
    };
    let mut xml = String::from("<pane");
    if pane.cols > 0 {
        xml.push_str(&format!(" xSplit=\"{}\"", pane.cols));
    }
    if pane.rows > 0 {
        xml.push_str(&format!(" ySplit=\"{}\"", pane.rows));
    }
    xml.push_str(&format!(
        " topLeftCell=\"{top_left}\" activePane=\"{active_pane}\" state=\"frozen\"/>"
    ));
    xml
}

#[allow(clippy::too_many_arguments)]
fn decode_cell_value(
    cell_type: Option<&str>,
    style_idx: Option<u32>,
    has_v: bool,
    v_text: &str,
    has_is: bool,
    is_text: &str,
    shared_strings: &[String],
    styles: &ParsedStyles,
) -> ParsedCellValue {
    match cell_type {
        Some("s") => {
            if !has_v {
                return ParsedCellValue::Empty;
            }
            match v_text.trim().parse::<usize>() {
                Ok(idx) => match shared_strings.get(idx) {
                    Some(s) => ParsedCellValue::String(s.clone()),
                    None => {
                        log::warn!("Shared string index {idx} out of bounds");
                        ParsedCellValue::Empty
                    }
                },
                Err(_) => ParsedCellValue::Empty,
            }
        }
        Some("inlineStr") => {
            if has_is {
                ParsedCellValue::String(decode_xstring(is_text))
            } else {
                ParsedCellValue::Empty
            }
        }
        Some("str") => {
            if has_v {
                ParsedCellValue::String(decode_xstring(v_text))
            } else {
                ParsedCellValue::Empty
            }
        }
        Some("b") => {
            if !has_v {
                return ParsedCellValue::Empty;
            }
            let trimmed = v_text.trim();
            ParsedCellValue::Bool(trimmed == "1" || trimmed.eq_ignore_ascii_case("true"))
        }
        Some("e") => {
            if has_v && !v_text.is_empty() {
                ParsedCellValue::Error(v_text.trim().to_string())
            } else {
                ParsedCellValue::Empty
            }
        }
        Some("d") => {
            if has_v && !v_text.is_empty() {
                ParsedCellValue::DateTimeIso(v_text.trim().to_string())
            } else {
                ParsedCellValue::Empty
            }
        }
        None | Some("n") => {
            if !has_v || v_text.trim().is_empty() {
                return ParsedCellValue::Empty;
            }
            match v_text.trim().parse::<f64>() {
                Ok(num) => {
                    if style_idx.is_some_and(|s| styles.is_date_xf(s)) {
                        ParsedCellValue::DateTime(num)
                    } else {
                        ParsedCellValue::Number(num)
                    }
                }
                Err(_) => ParsedCellValue::Empty,
            }
        }
        Some(_) => {
            if has_v && !v_text.is_empty() {
                ParsedCellValue::String(decode_xstring(v_text))
            } else {
                ParsedCellValue::Empty
            }
        }
    }
}

pub(crate) fn shift_formula_refs(formula: &str, d_row: isize, d_col: isize) -> String {
    if d_row == 0 && d_col == 0 {
        return formula.to_string();
    }
    let bytes = formula.as_bytes();
    let len = bytes.len();
    let mut out = String::with_capacity(len + 8);
    let mut i = 0usize;

    while i < len {
        match bytes[i] {
            b'"' => {
                out.push('"');
                i += 1;
                while i < len {
                    let ch = formula[i..].chars().next().unwrap();
                    out.push(ch);
                    i += ch.len_utf8();
                    if ch == '"' {
                        if i < len && bytes[i] == b'"' {
                            out.push('"');
                            i += 1;
                        } else {
                            break;
                        }
                    }
                }
            }
            b'\'' => {
                out.push('\'');
                i += 1;
                while i < len {
                    let ch = formula[i..].chars().next().unwrap();
                    out.push(ch);
                    i += ch.len_utf8();
                    if ch == '\'' {
                        if i < len && bytes[i] == b'\'' {
                            out.push('\'');
                            i += 1;
                        } else {
                            break;
                        }
                    }
                }
            }
            _ => {
                let prev_ok = i == 0
                    || !bytes[i - 1].is_ascii_alphanumeric()
                        && bytes[i - 1] != b'_'
                        && bytes[i - 1] != b'.';
                if prev_ok
                    && let Some((shifted, consumed)) =
                        try_shift_cell_ref(&formula[i..], d_row, d_col)
                {
                    out.push_str(&shifted);
                    i += consumed;
                    continue;
                }
                let ch = formula[i..].chars().next().unwrap();
                out.push(ch);
                i += ch.len_utf8();
            }
        }
    }

    out
}

fn try_shift_cell_ref(slice: &str, d_row: isize, d_col: isize) -> Option<(String, usize)> {
    let bytes = slice.as_bytes();
    let mut pos = 0usize;

    let col_abs = if pos < bytes.len() && bytes[pos] == b'$' {
        pos += 1;
        true
    } else {
        false
    };

    let col_start = pos;
    while pos < bytes.len() && bytes[pos].is_ascii_alphabetic() {
        pos += 1;
    }
    let col_len = pos - col_start;
    if col_len == 0 || col_len > 3 {
        return None;
    }
    let col_str = &slice[col_start..pos];

    let row_abs = if pos < bytes.len() && bytes[pos] == b'$' {
        pos += 1;
        true
    } else {
        false
    };

    let row_start = pos;
    while pos < bytes.len() && bytes[pos].is_ascii_digit() {
        pos += 1;
    }
    if pos == row_start {
        return None;
    }

    if pos < bytes.len() {
        let next = bytes[pos];
        if next.is_ascii_alphanumeric()
            || next == b'_'
            || next == b'('
            || next == b'!'
            || next == b'.'
        {
            return None;
        }
    }

    let col_idx = col_letters_to_idx(col_str);
    if col_idx >= 16_384 {
        return None;
    }
    let row_1based = slice[row_start..pos].parse::<usize>().ok()?;
    if row_1based == 0 || row_1based > 1_048_576 {
        return None;
    }
    let row_idx = row_1based - 1;

    let new_col = if col_abs {
        col_idx
    } else {
        let shifted = col_idx as isize + d_col;
        if !(0..16_384).contains(&shifted) {
            return Some(("#REF!".to_string(), pos));
        }
        shifted as usize
    };

    let new_row = if row_abs {
        row_idx
    } else {
        let shifted = row_idx as isize + d_row;
        if !(0..1_048_576).contains(&shifted) {
            return Some(("#REF!".to_string(), pos));
        }
        shifted as usize
    };

    let mut result = String::new();
    if col_abs {
        result.push('$');
    }
    result.push_str(&col_idx_to_letters(new_col));
    if row_abs {
        result.push('$');
    }
    result.push_str(&(new_row + 1).to_string());
    Some((result, pos))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn shifts_shared_formula_references_respecting_anchors_and_literals() {
        assert_eq!(
            shift_formula_refs("A1+$B$1+C$2+$D3", 2, 1),
            "B3+$B$1+D$2+$D5"
        );
        assert_eq!(
            shift_formula_refs("SUM(A1:A5)+\"A1\"+Sheet1!B2", 1, 0),
            "SUM(A2:A6)+\"A1\"+Sheet1!B3"
        );
        assert_eq!(shift_formula_refs("LOG10(A1)", 1, 1), "LOG10(B2)");
    }

    #[test]
    fn parses_all_cell_types_and_shared_formulas() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheetData>
                <row r="1">
                    <c r="A1"><v>10</v></c>
                    <c r="B1" t="s"><v>0</v></c>
                    <c r="C1" t="inlineStr"><is><t>inline</t></is></c>
                    <c r="D1" t="b"><v>1</v></c>
                    <c r="E1" t="e"><v>#DIV/0!</v></c>
                    <c r="F1"><f t="shared" ref="F1:F2" si="0">A1*2</f><v>20</v></c>
                </row>
                <row r="2">
                    <c r="A2"><v>15</v></c>
                    <c r="F2"><f t="shared" si="0"/><v>30</v></c>
                </row>
            </sheetData>
        </worksheet>"#;
        let ss = vec!["hello".to_string()];
        let styles = ParsedStyles::default();
        let ws = parse_worksheet(xml, &ss, &styles).unwrap();
        assert_eq!(ws.rows, 2);
        assert_eq!(ws.cols, 6);
        let f2 = ws.cells.iter().find(|c| c.row == 1 && c.col == 5).unwrap();
        assert_eq!(f2.formula.as_deref(), Some("A2*2"));
    }
}
