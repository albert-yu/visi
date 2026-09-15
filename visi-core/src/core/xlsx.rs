use crate::core::{CellStyle, CellType, DataColumn, Sheet};
use calamine::Reader;
use web_time::Instant;

fn result_data_for_calamine(cell_value: &calamine::Data) -> crate::core::engine::ResultData {
    match cell_value {
        calamine::Data::Int(i) => crate::core::engine::ResultData::Integer(*i),
        calamine::Data::Float(f) => crate::core::engine::ResultData::Float(*f),
        calamine::Data::String(s) => crate::core::engine::ResultData::String(s.clone()),
        calamine::Data::Bool(b) => crate::core::engine::ResultData::Boolean(*b),
        calamine::Data::Error(e) => crate::core::engine::ResultData::Error(e.to_string()),
        calamine::Data::DateTime(d) => crate::core::engine::ResultData::Float(d.as_f64()),
        calamine::Data::DateTimeIso(s) | calamine::Data::DurationIso(s) => {
            crate::core::engine::ResultData::String(s.clone())
        }
        calamine::Data::Empty => crate::core::engine::ResultData::None,
    }
}

fn cell_type_and_src_for_calamine(cell_value: &calamine::Data) -> (CellType, String) {
    match cell_value {
        calamine::Data::Empty => (CellType::Empty, String::new()),
        calamine::Data::Int(i) => (CellType::Int, i.to_string()),
        calamine::Data::Float(f) => (CellType::Float, f.to_string()),
        calamine::Data::String(s) => (CellType::String, s.clone()),
        calamine::Data::Bool(b) => (CellType::Bool, b.to_string()),
        calamine::Data::DateTime(d) => (CellType::DateTime, d.to_string()),
        calamine::Data::DateTimeIso(s) => (CellType::DateTimeIso, s.clone()),
        calamine::Data::DurationIso(s) => (CellType::DurationIso, s.clone()),
        calamine::Data::Error(e) => (CellType::Error, e.to_string()),
    }
}

/// [LLM-generated] A worksheet read out of an `.xlsx` file.
#[derive(Debug, serde::Serialize, serde::Deserialize)]
pub struct ImportedSheet {
    /// [LLM-generated] The sheet, with its cells, styles and any Excel Tables on it. Its
    /// values are already committed, so it can be read without evaluating.
    pub sheet: Sheet,
}

/// [LLM-generated] `(imported sheets, charts, pivot tables, VBA project)`.
pub type ImportedXlsxData = (
    Vec<ImportedSheet>,
    Vec<crate::core::chart::Chart>,
    Vec<crate::core::pivot::PivotTable>,
    Option<crate::core::vba::VbaProject>,
);

/// [LLM-generated] Read a `.xlsx` file from memory into sheets, charts, pivot tables and an
/// optional VBA project.
///
/// `existing_sheets` lets the importer keep column ids stable when reloading
/// a workbook it already has in hand; pass `&[]` for a cold load.
/// `progress_callback` is invoked as `(index, total, sheet_name)`.
///
/// # Errors
///
/// Returns [`crate::Error::Xlsx`] if the buffer is not a readable `.xlsx` container
/// or a part of it fails to parse.
pub fn import_xlsx_data(
    buffer: &[u8],
    existing_sheets: &[Sheet],
    progress_callback: impl FnMut(usize, usize, &str),
) -> crate::Result<ImportedXlsxData> {
    import_xlsx_data_raw(buffer, existing_sheets, progress_callback).map_err(crate::Error::Xlsx)
}

pub(crate) fn import_xlsx_data_raw(
    buffer: &[u8],
    existing_sheets: &[Sheet],
    mut progress_callback: impl FnMut(usize, usize, &str),
) -> Result<ImportedXlsxData, String> {
    let start_total = Instant::now();

    let cursor = std::io::Cursor::new(buffer);
    let start_open = Instant::now();
    let mut workbook: calamine::Xlsx<_> = calamine::open_workbook_from_rs(cursor)
        .map_err(|e| format!("Failed to open Excel file: {}", e))?;
    log::info!(
        "Excel open workbook took: {:.2?}, buffer size: {} bytes",
        start_open.elapsed(),
        buffer.len()
    );

    let sheet_names = workbook.sheet_names();
    let total_sheets = sheet_names.len();
    let mut imported_sheets = Vec::new();
    let mut orig_to_assigned_name: std::collections::HashMap<String, String> =
        std::collections::HashMap::new();

    let cell_styles = import_cell_styles(buffer).unwrap_or_default();
    let column_widths = import_column_widths(buffer).unwrap_or_default();
    let row_heights = import_row_heights(buffer).unwrap_or_default();

    for (sheet_idx, orig_sheet_name) in sheet_names.iter().enumerate() {
        progress_callback(sheet_idx, total_sheets, orig_sheet_name);
        let start_sheet = Instant::now();
        let start_range = Instant::now();
        if let Ok(range) = workbook.worksheet_range(orig_sheet_name) {
            let elapsed_range = start_range.elapsed();

            let start_formula = Instant::now();
            let formula_range = workbook.worksheet_formula(orig_sheet_name).ok();
            let elapsed_formula = start_formula.elapsed();

            let ends = [range.end(), formula_range.as_ref().and_then(|fr| fr.end())];
            let (mut rows, mut cols) = ends
                .into_iter()
                .flatten()
                .fold((0usize, 0usize), |acc, (r, c)| {
                    (acc.0.max(r as usize + 1), acc.1.max(c as usize + 1))
                });

            if rows == 0 || cols == 0 {
                rows = 10;
                cols = 5;
            }
            if let Some(sheet_styles) = cell_styles.get(orig_sheet_name) {
                for &(row_idx, col_idx) in sheet_styles.keys() {
                    rows = rows.max(row_idx + 1);
                    cols = cols.max(col_idx + 1);
                }
            }
            if let Some(sheet_widths) = column_widths.get(orig_sheet_name)
                && let Some(max_width_col) = sheet_widths.keys().max()
            {
                cols = cols.max(max_width_col + 1);
            }
            if let Some(sheet_heights) = row_heights.get(orig_sheet_name)
                && let Some(max_height_row) = sheet_heights.keys().max()
            {
                rows = rows.max(max_height_row + 1);
            }

            let mut sheet_name = orig_sheet_name.clone();
            let mut count = 1;
            while existing_sheets.iter().any(|s| s.name == sheet_name)
                || imported_sheets
                    .iter()
                    .any(|imported: &ImportedSheet| imported.sheet.name == sheet_name)
            {
                sheet_name = format!("{}_{}", sheet_name, count);
                count += 1;
            }

            let start_cells = Instant::now();
            let mut columns = (0..cols)
                .map(|col_idx| {
                    let mut col = DataColumn::new(rows);
                    col.id = rand::random::<u64>();
                    col.name = crate::core::parser::col_idx_to_letters(col_idx);
                    if let Some(width) = column_widths
                        .get(orig_sheet_name)
                        .and_then(|sheet_widths| sheet_widths.get(&col_idx))
                    {
                        col.width = Some(*width);
                    }
                    col
                })
                .collect::<Vec<_>>();

            let mut max_cell_lens = (0..cols)
                .map(|col_idx| columns[col_idx].name.len())
                .collect::<Vec<_>>();

            let range_height = range.height();
            let range_width = range.width();
            let matches_exactly = rows == range_height
                && cols == range_width
                && formula_range
                    .as_ref()
                    .is_none_or(|fr| fr.height() == rows && fr.width() == cols);

            if matches_exactly {
                let mut range_rows = range.rows();
                let mut formula_rows = formula_range.as_ref().map(|fr| fr.rows());

                for row_idx in 0..rows {
                    let r_cells = range_rows.next().unwrap();
                    let f_cells = formula_rows.as_mut().map(|fr| fr.next().unwrap());

                    for col_idx in 0..cols {
                        let cell_value = &r_cells[col_idx];
                        let init_res = result_data_for_calamine(cell_value);
                        columns[col_idx].data.set(row_idx, init_res);

                        if let Some(f_cells_slice) = f_cells {
                            let formula = &f_cells_slice[col_idx];
                            if !formula.is_empty() {
                                if matches!(cell_value, calamine::Data::Empty)
                                    || matches!(cell_value, calamine::Data::String(s) if s.is_empty())
                                {
                                    columns[col_idx]
                                        .data
                                        .set(row_idx, crate::core::engine::ResultData::None);
                                }
                                let cell_src = format!("={}", formula);
                                max_cell_lens[col_idx] = max_cell_lens[col_idx].max(cell_src.len());
                                columns[col_idx].src[row_idx] = cell_src;
                                columns[col_idx].cell_types[row_idx] =
                                    cell_type_and_src_for_calamine(cell_value).0;
                                continue;
                            }
                        }

                        let (cell_type, cell_src) = cell_type_and_src_for_calamine(cell_value);
                        max_cell_lens[col_idx] = max_cell_lens[col_idx].max(cell_src.len());
                        columns[col_idx].src[row_idx] = cell_src;
                        columns[col_idx].cell_types[row_idx] = cell_type;
                    }
                }
            } else {
                for row_idx in 0..rows {
                    let row_u32 = row_idx as u32;
                    for col_idx in 0..cols {
                        let col_u32 = col_idx as u32;

                        let cell_value = range
                            .get_value((row_u32, col_u32))
                            .unwrap_or(&calamine::Data::Empty);
                        let init_res = result_data_for_calamine(cell_value);
                        columns[col_idx].data.set(row_idx, init_res);

                        if let Some(ref f_range) = formula_range
                            && let Some(formula) = f_range.get_value((row_u32, col_u32))
                            && !formula.is_empty()
                        {
                            if matches!(cell_value, calamine::Data::Empty)
                                || matches!(cell_value, calamine::Data::String(s) if s.is_empty())
                            {
                                columns[col_idx]
                                    .data
                                    .set(row_idx, crate::core::engine::ResultData::None);
                            }
                            let cell_src = if formula.starts_with('=') {
                                formula.to_string()
                            } else {
                                format!("={}", formula)
                            };
                            max_cell_lens[col_idx] = max_cell_lens[col_idx].max(cell_src.len());
                            columns[col_idx].src[row_idx] = cell_src;
                            columns[col_idx].cell_types[row_idx] =
                                cell_type_and_src_for_calamine(cell_value).0;
                            columns[col_idx].dirty_indices.push(row_idx);
                            continue;
                        }

                        let (cell_type, cell_src) = cell_type_and_src_for_calamine(cell_value);
                        max_cell_lens[col_idx] = max_cell_lens[col_idx].max(cell_src.len());
                        columns[col_idx].src[row_idx] = cell_src;
                        columns[col_idx].cell_types[row_idx] = cell_type;
                    }
                }
            }
            if let Some(sheet_styles) = cell_styles.get(orig_sheet_name) {
                for (&(row_idx, col_idx), style) in sheet_styles {
                    let Some(col) = columns.get_mut(col_idx) else {
                        continue;
                    };
                    if row_idx >= col.styles.len() {
                        continue;
                    }
                    col.styles[row_idx] = Some(style.clone());
                }
            }
            let elapsed_cells = start_cells.elapsed();

            let mut sheet_row_heights = vec![None; rows];
            if let Some(imported_heights) = row_heights.get(orig_sheet_name) {
                for (&row_idx, &height) in imported_heights {
                    if row_idx < sheet_row_heights.len() {
                        sheet_row_heights[row_idx] = Some(height);
                    }
                }
            }

            orig_to_assigned_name.insert(orig_sheet_name.clone(), sheet_name.clone());
            let new_sheet = Sheet {
                id: rand::random::<u64>(),
                name: sheet_name,
                columns,
                row_heights: sheet_row_heights,
                tables: Vec::new(),
                dependencies: std::collections::HashMap::new(),
                dependencies_rev: std::collections::HashMap::new(),
                uncommitted_actions: Vec::new(),
                locale: crate::core::locale::Locale::default(),
            };
            imported_sheets.push(ImportedSheet { sheet: new_sheet });

            log::info!(
                "Worksheet '{}' ({}x{} cells) processed. range_parse: {:.2?}, formula_parse: {:.2?}, cells_convert: {:.2?}, total_sheet: {:.2?}",
                imported_sheets.last().unwrap().sheet.name,
                rows,
                cols,
                elapsed_range,
                elapsed_formula,
                elapsed_cells,
                start_sheet.elapsed()
            );
        }
    }

    if imported_sheets.is_empty() {
        return Err("No worksheets found in the Excel file".to_string());
    }

    for (orig_sheet_name, parsed) in import_tables_from_zip(buffer).unwrap_or_default() {
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
    if let Ok(parsed_charts) = import_charts_from_zip(buffer) {
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

pub(crate) fn parse_xlsx_color(color_str: &str) -> Option<rust_xlsxwriter::Color> {
    let trimmed = color_str.trim().trim_start_matches('#');
    if trimmed.eq_ignore_ascii_case("black") {
        return Some(rust_xlsxwriter::Color::Black);
    }
    if trimmed.eq_ignore_ascii_case("white") {
        return Some(rust_xlsxwriter::Color::White);
    }
    if trimmed.eq_ignore_ascii_case("red") {
        return Some(rust_xlsxwriter::Color::Red);
    }
    if trimmed.eq_ignore_ascii_case("blue") {
        return Some(rust_xlsxwriter::Color::Blue);
    }
    if trimmed.eq_ignore_ascii_case("green") {
        return Some(rust_xlsxwriter::Color::Green);
    }
    if trimmed.eq_ignore_ascii_case("yellow") {
        return Some(rust_xlsxwriter::Color::Yellow);
    }
    if trimmed.eq_ignore_ascii_case("gray") || trimmed.eq_ignore_ascii_case("grey") {
        return Some(rust_xlsxwriter::Color::Gray);
    }
    if trimmed.len() == 6 {
        if let Ok(rgb) = u32::from_str_radix(trimmed, 16) {
            return Some(rust_xlsxwriter::Color::RGB(rgb));
        }
    } else if trimmed.len() == 8
        && let Ok(rgb) = u32::from_str_radix(&trimmed[2..], 16)
    {
        return Some(rust_xlsxwriter::Color::RGB(rgb));
    }
    None
}

pub(crate) fn parse_xlsx_table_style(name: &str) -> rust_xlsxwriter::TableStyle {
    match name.to_lowercase().as_str() {
        "tablestylelight1" | "light1" => rust_xlsxwriter::TableStyle::Light1,
        "tablestylelight2" | "light2" => rust_xlsxwriter::TableStyle::Light2,
        "tablestylelight3" | "light3" => rust_xlsxwriter::TableStyle::Light3,
        "tablestylelight4" | "light4" => rust_xlsxwriter::TableStyle::Light4,
        "tablestylelight5" | "light5" => rust_xlsxwriter::TableStyle::Light5,
        "tablestylelight6" | "light6" => rust_xlsxwriter::TableStyle::Light6,
        "tablestylelight7" | "light7" => rust_xlsxwriter::TableStyle::Light7,
        "tablestylelight8" | "light8" => rust_xlsxwriter::TableStyle::Light8,
        "tablestylelight9" | "light9" => rust_xlsxwriter::TableStyle::Light9,
        "tablestylelight10" | "light10" => rust_xlsxwriter::TableStyle::Light10,
        "tablestylemedium1" | "medium1" => rust_xlsxwriter::TableStyle::Medium1,
        "tablestylemedium2" | "medium2" => rust_xlsxwriter::TableStyle::Medium2,
        "tablestylemedium3" | "medium3" => rust_xlsxwriter::TableStyle::Medium3,
        "tablestylemedium4" | "medium4" => rust_xlsxwriter::TableStyle::Medium4,
        "tablestylemedium5" | "medium5" => rust_xlsxwriter::TableStyle::Medium5,
        "tablestylemedium6" | "medium6" => rust_xlsxwriter::TableStyle::Medium6,
        "tablestylemedium7" | "medium7" => rust_xlsxwriter::TableStyle::Medium7,
        "tablestylemedium8" | "medium8" => rust_xlsxwriter::TableStyle::Medium8,
        "tablestylemedium9" | "medium9" => rust_xlsxwriter::TableStyle::Medium9,
        "tablestylemedium10" | "medium10" => rust_xlsxwriter::TableStyle::Medium10,
        "tablestyledark1" | "dark1" => rust_xlsxwriter::TableStyle::Dark1,
        "tablestyledark2" | "dark2" => rust_xlsxwriter::TableStyle::Dark2,
        "tablestyledark3" | "dark3" => rust_xlsxwriter::TableStyle::Dark3,
        "tablestyledark4" | "dark4" => rust_xlsxwriter::TableStyle::Dark4,
        "tablestyledark5" | "dark5" => rust_xlsxwriter::TableStyle::Dark5,
        "tablestyledark6" | "dark6" => rust_xlsxwriter::TableStyle::Dark6,
        "tablestyledark7" | "dark7" => rust_xlsxwriter::TableStyle::Dark7,
        "tablestyledark8" | "dark8" => rust_xlsxwriter::TableStyle::Dark8,
        "tablestyledark9" | "dark9" => rust_xlsxwriter::TableStyle::Dark9,
        "tablestyledark10" | "dark10" => rust_xlsxwriter::TableStyle::Dark10,
        "tablestyledark11" | "dark11" => rust_xlsxwriter::TableStyle::Dark11,
        _ => rust_xlsxwriter::TableStyle::Medium9,
    }
}

/// [LLM-generated] The numeric serial to export for a date-formatted cell, if this is one.
///
/// A date cell keeps the typed text in `src` and the serial in `data`, so the
/// value -- not the source -- is what Excel needs alongside the `numFmt`.
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

pub(crate) fn build_xlsx_format(style: &crate::core::CellStyle) -> rust_xlsxwriter::Format {
    let mut format = rust_xlsxwriter::Format::new();
    if let Some(font_color) = &style.font_color
        && let Some(color) = parse_xlsx_color(font_color)
    {
        format = format.set_font_color(color);
    }
    if let Some(bg_color) = &style.bg_color
        && let Some(color) = parse_xlsx_color(bg_color)
    {
        format = format.set_background_color(color);
    }
    if let Some(true) = style.bold {
        format = format.set_bold();
    }
    if let Some(true) = style.italic {
        format = format.set_italic();
    }
    if let Some(true) = style.underline {
        format = format.set_underline(rust_xlsxwriter::FormatUnderline::Single);
    }
    if let Some(family) = &style.font_family {
        format = format.set_font_name(family);
    }
    if let Some(size) = style.font_size {
        format = format.set_font_size(size);
    }
    if let Some(code) = &style.num_format {
        format = format.set_num_format(code);
    }
    format
}

/// [LLM-generated] Serialize sheets, charts, pivot tables and an optional VBA project into a
/// `.xlsx` file in memory.
///
/// Formulas are written with their cached results, so readers that do not
/// recalculate (Excel on open, `openpyxl`) still see values.
///
/// # Errors
///
/// Returns [`crate::Error::Xlsx`] if the workbook cannot be serialized.
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
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let mut used_names = std::collections::HashSet::new();
    let mut table_name_to_worksheet_name = std::collections::HashMap::new();
    let mut table_name_to_table = std::collections::HashMap::new();
    let mut sheet_id_to_worksheet_name = std::collections::HashMap::new();
    let mut empty_table_names = std::collections::HashSet::new();
    let mut worksheet_name_to_column_widths = std::collections::HashMap::new();
    let mut worksheet_name_to_row_heights = std::collections::HashMap::new();

    for sheet in sheets {
        table_name_to_table.insert(sheet.name.clone(), sheet);

        let mut worksheet_name = if sheet.name.len() > 31 {
            sheet.name[..31].to_string()
        } else {
            sheet.name.clone()
        };

        let mut counter = 1;
        while used_names.contains(&worksheet_name.to_lowercase()) {
            let counter_suffix = format!("_{}", counter);
            let needed_len = counter_suffix.len();
            let max_name_len = 31 - needed_len;
            let base_name = if sheet.name.len() > max_name_len {
                &sheet.name[..max_name_len]
            } else {
                &sheet.name
            };
            worksheet_name = format!("{}{}", base_name, counter_suffix);
            counter += 1;
        }

        used_names.insert(worksheet_name.to_lowercase());
        table_name_to_worksheet_name.insert(sheet.name.clone(), worksheet_name.clone());
        sheet_id_to_worksheet_name.insert(sheet.id, worksheet_name.clone());
        let explicit_widths = sheet
            .columns
            .iter()
            .enumerate()
            .filter_map(|(idx, col)| {
                col.width
                    .filter(|width| width.is_finite() && *width >= 0.0)
                    .map(|width| (idx, width))
            })
            .collect::<Vec<_>>();
        if !explicit_widths.is_empty() {
            worksheet_name_to_column_widths.insert(worksheet_name.clone(), explicit_widths);
        }
        let explicit_heights = sheet
            .row_heights
            .iter()
            .enumerate()
            .filter_map(|(idx, height)| {
                let value = (*height)?;
                if value.is_finite() && value >= 0.0 {
                    Some((idx, value))
                } else {
                    None
                }
            })
            .collect::<Vec<_>>();
        if !explicit_heights.is_empty() {
            worksheet_name_to_row_heights.insert(worksheet_name.clone(), explicit_heights);
        }

        {
            let worksheet = workbook.add_worksheet();
            worksheet
                .set_name(&worksheet_name)
                .map_err(|e| format!("Failed to set worksheet name: {}", e))?;
            worksheet.set_formula_result_default("");

            for col_idx in 0..sheet.columns.len() {
                let col = &sheet.columns[col_idx];
                for row_idx in 0..sheet.row_count() {
                    let cell_src = col.src.get(row_idx).cloned().unwrap_or_default();
                    let cell_type = col
                        .cell_types
                        .get(row_idx)
                        .copied()
                        .unwrap_or(CellType::Empty);
                    let style_opt = col.styles.get(row_idx).and_then(|s| s.as_ref());
                    let format_opt = style_opt.map(build_xlsx_format);

                    if let Some(formula_src) = cell_src.strip_prefix('=') {
                        let mut formula = rust_xlsxwriter::Formula::new(formula_src);
                        if let Some(res_data) = col.data.get(row_idx) {
                            match res_data {
                                crate::core::engine::ResultData::None => {}
                                _ => {
                                    let res_str = format_result_for_xlsx(&res_data);
                                    formula = formula.set_result(res_str);
                                }
                            }
                        }
                        if let Some(ref rx_format) = format_opt {
                            worksheet
                                .write_formula_with_format(
                                    row_idx as u32,
                                    col_idx as u16,
                                    formula,
                                    rx_format,
                                )
                                .map_err(|e| format!("Failed to write Excel formula: {}", e))?;
                        } else {
                            worksheet
                                .write_formula(row_idx as u32, col_idx as u16, formula)
                                .map_err(|e| format!("Failed to write Excel formula: {}", e))?;
                        }
                    } else if cell_type == CellType::String {
                        let text = if cell_src.starts_with('"')
                            && cell_src.ends_with('"')
                            && cell_src.len() >= 2
                        {
                            &cell_src[1..cell_src.len() - 1]
                        } else {
                            cell_src.as_str()
                        };
                        if let Some(ref rx_format) = format_opt {
                            worksheet
                                .write_string_with_format(
                                    row_idx as u32,
                                    col_idx as u16,
                                    text,
                                    rx_format,
                                )
                                .map_err(|e| format!("Failed to write Excel string: {}", e))?;
                        } else {
                            worksheet
                                .write_string(row_idx as u32, col_idx as u16, text)
                                .map_err(|e| format!("Failed to write Excel string: {}", e))?;
                        }
                    } else if let Some(serial) = date_serial_for_export(style_opt, col, row_idx) {
                        let rx_format = format_opt
                            .as_ref()
                            .expect("a date cell always has a num_format style");
                        worksheet
                            .write_number_with_format(
                                row_idx as u32,
                                col_idx as u16,
                                serial,
                                rx_format,
                            )
                            .map_err(|e| format!("Failed to write Excel date: {}", e))?;
                    } else if let Ok(val_f64) = cell_src.parse::<f64>() {
                        if let Some(ref rx_format) = format_opt {
                            worksheet
                                .write_number_with_format(
                                    row_idx as u32,
                                    col_idx as u16,
                                    val_f64,
                                    rx_format,
                                )
                                .map_err(|e| format!("Failed to write Excel number: {}", e))?;
                        } else {
                            worksheet
                                .write_number(row_idx as u32, col_idx as u16, val_f64)
                                .map_err(|e| format!("Failed to write Excel number: {}", e))?;
                        }
                    } else if let Ok(val_bool) = cell_src.parse::<bool>() {
                        if let Some(ref rx_format) = format_opt {
                            worksheet
                                .write_boolean_with_format(
                                    row_idx as u32,
                                    col_idx as u16,
                                    val_bool,
                                    rx_format,
                                )
                                .map_err(|e| format!("Failed to write Excel boolean: {}", e))?;
                        } else {
                            worksheet
                                .write_boolean(row_idx as u32, col_idx as u16, val_bool)
                                .map_err(|e| format!("Failed to write Excel boolean: {}", e))?;
                        }
                    } else if !cell_src.is_empty() || format_opt.is_some() {
                        let text = if cell_src.starts_with('"')
                            && cell_src.ends_with('"')
                            && cell_src.len() >= 2
                        {
                            &cell_src[1..cell_src.len() - 1]
                        } else {
                            cell_src.as_str()
                        };
                        if let Some(ref rx_format) = format_opt {
                            worksheet
                                .write_string_with_format(
                                    row_idx as u32,
                                    col_idx as u16,
                                    text,
                                    rx_format,
                                )
                                .map_err(|e| format!("Failed to write Excel string: {}", e))?;
                        } else {
                            worksheet
                                .write_string(row_idx as u32, col_idx as u16, text)
                                .map_err(|e| format!("Failed to write Excel string: {}", e))?;
                        }
                    }
                }
            }

            for table in &sheet.tables {
                let rx_columns: Vec<rust_xlsxwriter::TableColumn> = table
                    .columns
                    .iter()
                    .map(|name| rust_xlsxwriter::TableColumn::new().set_header(name))
                    .collect();
                let is_empty_table = table.has_insert_row
                    || (table.has_header_row
                        && table.start_row == table.end_row
                        && !table.has_totals_row);
                if is_empty_table {
                    empty_table_names.insert(table.name.clone());
                }
                let end_row = if is_empty_table {
                    table.start_row + 1
                } else {
                    table.end_row
                };
                let mut rx_table = rust_xlsxwriter::Table::new()
                    .set_name(&table.name)
                    .set_header_row(table.has_header_row)
                    .set_total_row(table.has_totals_row)
                    .set_columns(&rx_columns);
                if let Some(style_name) = &table.style_name {
                    rx_table = rx_table.set_style(parse_xlsx_table_style(style_name));
                }
                worksheet
                    .add_table(
                        table.start_row as u32,
                        table.start_col as u16,
                        end_row as u32,
                        table.end_col as u16,
                        &rx_table,
                    )
                    .map_err(|e| format!("Failed to add Excel table '{}': {}", table.name, e))?;
            }
        }
    }

    for chart in charts {
        let parts: Vec<&str> = chart.data_range.split('!').collect();
        let ref_table_name = parts.first().copied().unwrap_or("").trim();

        let tbl_name = if ref_table_name.is_empty() {
            sheets
                .first()
                .map(|tbl| tbl.name.clone())
                .unwrap_or_default()
        } else {
            ref_table_name.to_string()
        };

        let ws_name = table_name_to_worksheet_name
            .get(&tbl_name)
            .or_else(|| table_name_to_worksheet_name.values().next())
            .cloned();

        if let Some(ws_name) = ws_name
            && let Ok(worksheet) = workbook.worksheet_from_name(&ws_name)
        {
            let rx_chart_type = match chart.chart_type {
                crate::core::chart::ChartType::Column => rust_xlsxwriter::ChartType::Column,
                crate::core::chart::ChartType::Bar => rust_xlsxwriter::ChartType::Bar,
                crate::core::chart::ChartType::Line => rust_xlsxwriter::ChartType::Line,
                crate::core::chart::ChartType::Pie => rust_xlsxwriter::ChartType::Pie,
                crate::core::chart::ChartType::Scatter => rust_xlsxwriter::ChartType::Scatter,
                crate::core::chart::ChartType::Area => rust_xlsxwriter::ChartType::Area,
            };

            let mut rx_chart = rust_xlsxwriter::Chart::new(rx_chart_type);

            let ast = crate::core::parser::parse_excel_formula(&chart.data_range);
            match ast {
                Ok(crate::core::parser::Expr::CellRef { row, col, .. }) => {
                    let col_letter = crate::core::parser::col_idx_to_letters(col);
                    let values_range = format!(
                        "{}!${}${}:${}${}",
                        ws_name,
                        col_letter,
                        row + 1,
                        col_letter,
                        row + 1
                    );
                    rx_chart.add_series().set_values(&values_range);
                }
                Ok(crate::core::parser::Expr::RangeRef {
                    start_row,
                    start_col,
                    end_row,
                    end_col,
                    ..
                }) => {
                    let cols_count = end_col + 1 - start_col;
                    if cols_count >= 2 {
                        let categories_col = crate::core::parser::col_idx_to_letters(start_col);
                        let values_col = crate::core::parser::col_idx_to_letters(start_col + 1);
                        let (categories_range, values_range) = if end_row == usize::MAX {
                            (
                                format!("{}!${}:${}", ws_name, categories_col, categories_col),
                                format!("{}!${}:${}", ws_name, values_col, values_col),
                            )
                        } else {
                            (
                                format!(
                                    "{}!${}${}:${}${}",
                                    ws_name,
                                    categories_col,
                                    start_row + 1,
                                    categories_col,
                                    end_row + 1
                                ),
                                format!(
                                    "{}!${}${}:${}${}",
                                    ws_name,
                                    values_col,
                                    start_row + 1,
                                    values_col,
                                    end_row + 1
                                ),
                            )
                        };
                        rx_chart
                            .add_series()
                            .set_categories(&categories_range)
                            .set_values(&values_range);
                    } else {
                        let col_letter = crate::core::parser::col_idx_to_letters(start_col);
                        let values_range = if end_row == usize::MAX {
                            format!("{}!${}:${}", ws_name, col_letter, col_letter)
                        } else {
                            format!(
                                "{}!${}${}:${}${}",
                                ws_name,
                                col_letter,
                                start_row + 1,
                                col_letter,
                                end_row + 1
                            )
                        };
                        rx_chart.add_series().set_values(&values_range);
                    }
                }
                _ => {
                    rx_chart.add_series().set_values(&chart.data_range);
                }
            }

            if let Some(ref title) = chart.title {
                rx_chart.title().set_name(title);
            }
            let (x_axis_label, y_axis_label) =
                if chart.chart_type == crate::core::chart::ChartType::Bar {
                    (chart.ylabel.as_ref(), chart.xlabel.as_ref())
                } else {
                    (chart.xlabel.as_ref(), chart.ylabel.as_ref())
                };
            if let Some(label) = x_axis_label {
                rx_chart.x_axis().set_name(label);
            }
            if let Some(label) = y_axis_label {
                rx_chart.y_axis().set_name(label);
            }
            if !chart.show_legend {
                rx_chart.legend().set_hidden();
            }

            worksheet
                .insert_chart(chart.anchor_row as u32, chart.anchor_col as u16, &rx_chart)
                .map_err(|e| format!("Failed to insert Excel chart: {}", e))?;
        }
    }

    let buffer = workbook
        .save_to_buffer()
        .map_err(|e| format!("Failed to write XLSX buffer: {}", e))?;

    let buffer = inject_column_widths(buffer, &worksheet_name_to_column_widths)?;
    let buffer = inject_row_heights(buffer, &worksheet_name_to_row_heights)?;

    let formula_string_result_cells = formula_string_result_cells_by_sheet(sheets);
    let buffer = inject_formula_string_result_types(buffer, &formula_string_result_cells)?;

    let buffer = if empty_table_names.is_empty() {
        buffer
    } else {
        inject_empty_table_flags(buffer, &empty_table_names)?
    };

    let buffer = if pivots.is_empty() {
        buffer
    } else {
        crate::core::pivot_xlsx::inject_pivot_tables(buffer, sheets, pivots)?
    };

    crate::core::vba_xlsx::export_vba_project(buffer, vba, &sheet_id_to_worksheet_name)
}

fn formula_string_result_cells_by_sheet(
    sheets: &[Sheet],
) -> Vec<std::collections::HashMap<String, String>> {
    sheets
        .iter()
        .map(|sheet| {
            let mut cells = std::collections::HashMap::new();
            for (col_idx, col) in sheet.columns.iter().enumerate() {
                for row_idx in 0..sheet.row_count() {
                    let is_formula = col.src.get(row_idx).is_some_and(|src| src.starts_with('='));
                    if is_formula
                        && let Some(result_text) = col
                            .data
                            .get(row_idx)
                            .and_then(|res_data| formula_string_result_text(&res_data))
                    {
                        let col_name = crate::core::parser::col_idx_to_letters(col_idx);
                        cells.insert(format!("{}{}", col_name, row_idx + 1), result_text);
                    }
                }
            }
            cells
        })
        .collect()
}

fn formula_string_result_text(res_data: &crate::core::engine::ResultData) -> Option<String> {
    match res_data {
        crate::core::engine::ResultData::String(s) => Some(s.clone()),
        crate::core::engine::ResultData::List(items) => {
            items.first().and_then(formula_string_result_text)
        }
        _ => None,
    }
}

fn inject_formula_string_result_types(
    original: Vec<u8>,
    formula_string_result_cells: &[std::collections::HashMap<String, String>],
) -> Result<Vec<u8>, String> {
    if formula_string_result_cells
        .iter()
        .all(|cells| cells.is_empty())
    {
        return Ok(original);
    }

    use std::io::{Read, Write};
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

        if let Some(sheet_idx) = worksheet_part_index(&name)
            && let Some(cells) = formula_string_result_cells.get(sheet_idx)
            && !cells.is_empty()
            && let Ok(xml) = std::str::from_utf8(&buf)
        {
            let patched = force_formula_cached_strings(xml, cells);
            writer
                .write_all(patched.as_bytes())
                .map_err(|e| e.to_string())?;
            continue;
        }

        writer.write_all(&buf).map_err(|e| e.to_string())?;
    }

    let cursor = writer.finish().map_err(|e| e.to_string())?;
    Ok(cursor.into_inner())
}

fn worksheet_part_index(name: &str) -> Option<usize> {
    let stem = name
        .strip_prefix("xl/worksheets/sheet")?
        .strip_suffix(".xml")?;
    stem.parse::<usize>().ok()?.checked_sub(1)
}

fn force_formula_cached_strings(
    xml: &str,
    formula_string_result_cells: &std::collections::HashMap<String, String>,
) -> String {
    let mut out = String::with_capacity(xml.len() + formula_string_result_cells.len() * 8);
    let mut pos = 0;

    while let Some(rel_start) = xml[pos..].find("<c ") {
        let start = pos + rel_start;
        out.push_str(&xml[pos..start]);

        let Some(rel_end) = xml[start..].find('>') else {
            out.push_str(&xml[start..]);
            return out;
        };
        let end = start + rel_end;
        let open_tag = &xml[start..end];

        if let Some(result_text) = cell_ref_from_open_tag(open_tag)
            .and_then(|cell_ref| formula_string_result_cells.get(cell_ref))
        {
            let Some(rel_close) = xml[end + 1..].find("</c>") else {
                out.push_str(&force_open_tag_type_str(open_tag));
                out.push('>');
                pos = end + 1;
                continue;
            };
            let close = end + 1 + rel_close;
            out.push_str(&force_open_tag_type_str(open_tag));
            out.push('>');
            out.push_str(&force_formula_cached_string_value(
                &xml[end + 1..close],
                result_text,
            ));
            out.push_str("</c>");
            pos = close + 4;
        } else {
            out.push_str(open_tag);
            out.push('>');
            pos = end + 1;
        }
    }

    out.push_str(&xml[pos..]);
    out
}

fn force_formula_cached_string_value(cell_body: &str, result_text: &str) -> String {
    let Some(v_start) = cell_body.find("<v>") else {
        return cell_body.to_string();
    };
    let value_start = v_start + 3;
    let Some(rel_v_end) = cell_body[value_start..].find("</v>") else {
        return cell_body.to_string();
    };
    let value_end = value_start + rel_v_end;
    format!(
        "{}{}{}",
        &cell_body[..value_start],
        escape_xml(result_text),
        &cell_body[value_end..]
    )
}

fn cell_ref_from_open_tag(open_tag: &str) -> Option<&str> {
    let value_start = open_tag
        .find(" r=\"")
        .map(|idx| idx + 4)
        .or_else(|| open_tag.strip_prefix("<c r=\"").map(|_| 6))?;
    let value_end = value_start + open_tag[value_start..].find('"')?;
    Some(&open_tag[value_start..value_end])
}

fn force_open_tag_type_str(open_tag: &str) -> String {
    if let Some(type_attr_start) = open_tag.find(" t=\"") {
        let value_start = type_attr_start + 4;
        if let Some(rel_value_end) = open_tag[value_start..].find('"') {
            let value_end = value_start + rel_value_end;
            return format!("{}str{}", &open_tag[..value_start], &open_tag[value_end..]);
        }
    }

    format!("{} t=\"str\"", open_tag)
}

fn inject_column_widths(
    original: Vec<u8>,
    widths_by_sheet: &std::collections::HashMap<String, Vec<(usize, f64)>>,
) -> Result<Vec<u8>, String> {
    if widths_by_sheet.is_empty() {
        return Ok(original);
    }

    use std::io::{Read, Write};
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(original.as_slice()))
        .map_err(|e| format!("Failed to re-open generated xlsx zip: {}", e))?;
    let sheet_file_to_name = build_sheet_file_to_name(&mut archive)?;
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

        if let Some(sheet_file) = name.strip_prefix("xl/worksheets/")
            && !sheet_file.contains('/')
            && let Some(sheet_name) = sheet_file_to_name.get(sheet_file)
            && let Some(widths) = widths_by_sheet.get(sheet_name)
            && let Ok(xml) = std::str::from_utf8(&buf)
        {
            let cols_xml = build_column_widths_xml(widths);
            let new_xml = replace_or_insert_cols_xml(xml, &cols_xml);
            writer
                .write_all(new_xml.as_bytes())
                .map_err(|e| e.to_string())?;
            continue;
        }

        writer.write_all(&buf).map_err(|e| e.to_string())?;
    }

    let cursor = writer.finish().map_err(|e| e.to_string())?;
    Ok(cursor.into_inner())
}

fn build_column_widths_xml(widths: &[(usize, f64)]) -> String {
    let mut sorted = widths.to_vec();
    sorted.sort_by_key(|(idx, _)| *idx);
    let mut out = String::from("<cols>");
    for (idx, width) in sorted {
        if !width.is_finite() || width < 0.0 || idx >= 16_384 {
            continue;
        }
        let one_based = idx + 1;
        out.push_str(&format!(
            "<col min=\"{}\" max=\"{}\" width=\"{}\" customWidth=\"1\"/>",
            one_based, one_based, width
        ));
    }
    out.push_str("</cols>");
    out
}

fn replace_or_insert_cols_xml(xml: &str, cols_xml: &str) -> String {
    if let Some(cols_start) = xml.find("<cols") {
        if let Some(rel_cols_end) = xml[cols_start..].find("</cols>") {
            let cols_end = cols_start + rel_cols_end + "</cols>".len();
            return format!("{}{}{}", &xml[..cols_start], cols_xml, &xml[cols_end..]);
        }
        if let Some(rel_tag_end) = xml[cols_start..].find("/>") {
            let cols_end = cols_start + rel_tag_end + 2;
            return format!("{}{}{}", &xml[..cols_start], cols_xml, &xml[cols_end..]);
        }
    }

    if let Some(sheet_format_start) = xml.find("<sheetFormatPr") {
        if let Some(rel_empty_end) = xml[sheet_format_start..].find("/>") {
            let insert_at = sheet_format_start + rel_empty_end + 2;
            return format!("{}{}{}", &xml[..insert_at], cols_xml, &xml[insert_at..]);
        }
        if let Some(rel_end) = xml[sheet_format_start..].find("</sheetFormatPr>") {
            let insert_at = sheet_format_start + rel_end + "</sheetFormatPr>".len();
            return format!("{}{}{}", &xml[..insert_at], cols_xml, &xml[insert_at..]);
        }
    }

    if let Some(sheet_data_start) = xml.find("<sheetData") {
        return format!(
            "{}{}{}",
            &xml[..sheet_data_start],
            cols_xml,
            &xml[sheet_data_start..]
        );
    }

    xml.to_string()
}

fn inject_row_heights(
    original: Vec<u8>,
    heights_by_sheet: &std::collections::HashMap<String, Vec<(usize, f64)>>,
) -> Result<Vec<u8>, String> {
    if heights_by_sheet.is_empty() {
        return Ok(original);
    }

    use std::io::{Read, Write};
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(original.as_slice()))
        .map_err(|e| format!("Failed to re-open generated xlsx zip: {}", e))?;
    let sheet_file_to_name = build_sheet_file_to_name(&mut archive)?;
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

        if let Some(sheet_file) = name.strip_prefix("xl/worksheets/")
            && !sheet_file.contains('/')
            && let Some(sheet_name) = sheet_file_to_name.get(sheet_file)
            && let Some(heights) = heights_by_sheet.get(sheet_name)
            && let Ok(xml) = std::str::from_utf8(&buf)
        {
            let new_xml = apply_row_heights_xml(xml, heights);
            writer
                .write_all(new_xml.as_bytes())
                .map_err(|e| e.to_string())?;
            continue;
        }

        writer.write_all(&buf).map_err(|e| e.to_string())?;
    }

    let cursor = writer.finish().map_err(|e| e.to_string())?;
    Ok(cursor.into_inner())
}

fn apply_row_heights_xml(xml: &str, heights: &[(usize, f64)]) -> String {
    let Some(sheet_data_start) = xml.find("<sheetData") else {
        return xml.to_string();
    };
    let mut remaining = heights
        .iter()
        .filter_map(|(idx, height)| {
            if *idx < 1_048_576 && height.is_finite() && *height >= 0.0 {
                Some((idx + 1, *height))
            } else {
                None
            }
        })
        .collect::<Vec<_>>();
    remaining.sort_by_key(|(row, _)| *row);
    remaining.dedup_by_key(|(row, _)| *row);
    if remaining.is_empty() {
        return xml.to_string();
    }

    if let Some(rel_empty_end) = xml[sheet_data_start..].find("/>")
        && xml[sheet_data_start..]
            .find('>')
            .is_some_and(|rel_end| rel_empty_end < rel_end)
    {
        let tag_end = sheet_data_start + rel_empty_end + 2;
        let rows_xml = remaining
            .iter()
            .map(|(row, height)| empty_row_height_xml(*row, *height))
            .collect::<String>();
        return format!(
            "{}<sheetData>{}</sheetData>{}",
            &xml[..sheet_data_start],
            rows_xml,
            &xml[tag_end..]
        );
    }

    let Some(rel_sheet_data_open_end) = xml[sheet_data_start..].find('>') else {
        return xml.to_string();
    };
    let body_start = sheet_data_start + rel_sheet_data_open_end + 1;
    let Some(rel_sheet_data_close) = xml[body_start..].find("</sheetData>") else {
        return xml.to_string();
    };
    let body_end = body_start + rel_sheet_data_close;
    let body = &xml[body_start..body_end];

    let mut out_body = String::with_capacity(body.len() + remaining.len() * 32);
    let mut pos = 0;
    let mut next_height = 0usize;
    while let Some(rel_row_start) = body[pos..].find("<row") {
        let row_start = pos + rel_row_start;
        out_body.push_str(&body[pos..row_start]);
        let Some(rel_open_end) = body[row_start..].find('>') else {
            out_body.push_str(&body[row_start..]);
            pos = body.len();
            break;
        };
        let open_end = row_start + rel_open_end;
        let open_tag = &body[row_start..=open_end];
        let row_num = attr_value(open_tag, "r").and_then(|s| s.parse::<usize>().ok());

        if let Some(row_num) = row_num {
            while next_height < remaining.len() && remaining[next_height].0 < row_num {
                out_body.push_str(&empty_row_height_xml(
                    remaining[next_height].0,
                    remaining[next_height].1,
                ));
                next_height += 1;
            }
        }

        let row_end = if open_tag.ends_with("/>") {
            open_end + 1
        } else if let Some(rel_close) = body[open_end + 1..].find("</row>") {
            open_end + 1 + rel_close + "</row>".len()
        } else {
            open_end + 1
        };
        let mut row_xml = body[row_start..row_end].to_string();
        if let Some(row_num) = row_num
            && next_height < remaining.len()
            && remaining[next_height].0 == row_num
        {
            row_xml = set_row_height_attrs(&row_xml, remaining[next_height].1);
            next_height += 1;
        }
        out_body.push_str(&row_xml);
        pos = row_end;
    }
    out_body.push_str(&body[pos..]);
    while next_height < remaining.len() {
        out_body.push_str(&empty_row_height_xml(
            remaining[next_height].0,
            remaining[next_height].1,
        ));
        next_height += 1;
    }

    format!("{}{}{}", &xml[..body_start], out_body, &xml[body_end..])
}

fn empty_row_height_xml(row: usize, height: f64) -> String {
    format!("<row r=\"{}\" ht=\"{}\" customHeight=\"1\"/>", row, height)
}

fn set_row_height_attrs(row_xml: &str, height: f64) -> String {
    let Some(open_end) = row_xml.find('>') else {
        return row_xml.to_string();
    };
    let open_tag = &row_xml[..open_end];
    let open_tag = set_or_insert_attr(open_tag, "ht", &height.to_string());
    let open_tag = set_or_insert_attr(&open_tag, "customHeight", "1");
    format!("{}{}", open_tag, &row_xml[open_end..])
}

fn set_or_insert_attr(tag: &str, name: &str, value: &str) -> String {
    let needle = format!(" {}=\"", name);
    if let Some(attr_start) = tag.find(&needle) {
        let value_start = attr_start + needle.len();
        if let Some(rel_value_end) = tag[value_start..].find('"') {
            let value_end = value_start + rel_value_end;
            return format!(
                "{}{}{}",
                &tag[..value_start],
                escape_xml(value),
                &tag[value_end..]
            );
        }
    }
    if let Some(prefix) = tag.strip_suffix('/') {
        return format!("{} {}=\"{}\"/", prefix, name, escape_xml(value));
    }
    format!("{} {}=\"{}\"", tag, name, escape_xml(value))
}

fn attr_value<'a>(tag: &'a str, name: &str) -> Option<&'a str> {
    let needle = format!(" {}=\"", name);
    let value_start = tag.find(&needle)? + needle.len();
    let value_end = value_start + tag[value_start..].find('"')?;
    Some(&tag[value_start..value_end])
}

fn inject_empty_table_flags(
    original: Vec<u8>,
    empty_tables: &std::collections::HashSet<String>,
) -> Result<Vec<u8>, String> {
    use std::io::{Read, Write};
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

        if name.starts_with("xl/tables/table")
            && name.ends_with(".xml")
            && let Ok(xml) = std::str::from_utf8(&buf)
        {
            let is_target = empty_tables.iter().any(|t| {
                xml.contains(&format!("displayName=\"{}\"", t))
                    || xml.contains(&format!("name=\"{}\"", t))
            });
            if is_target
                && !xml.contains("insertRow=")
                && let Some(pos) = xml.find("<table ")
            {
                let mut new_xml = String::with_capacity(xml.len() + 16);
                new_xml.push_str(&xml[..pos + 7]);
                new_xml.push_str("insertRow=\"1\" ");
                new_xml.push_str(&xml[pos + 7..]);
                writer
                    .write_all(new_xml.as_bytes())
                    .map_err(|e| e.to_string())?;
                continue;
            }
        }
        writer.write_all(&buf).map_err(|e| e.to_string())?;
    }

    let cursor = writer.finish().map_err(|e| e.to_string())?;
    Ok(cursor.into_inner())
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

/// [LLM-generated] Escapes text for use inside an XML attribute value. Shared by
/// `pivot_xlsx.rs` and `vba_xlsx.rs`.
pub(crate) fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
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

/// [LLM-generated] Collects the table parts a worksheet's `_rels/sheetN.xml.rels` points at,
/// as bare filenames (`table1.xml`). Every table part lives in `xl/tables/`,
/// so taking the basename via `get_filename` sidesteps resolving the
/// `Target` attribute's form entirely: `../tables/table1.xml` (what Excel
/// and `rust_xlsxwriter` emit) and `/xl/tables/table1.xml` (the absolute
/// package path `openpyxl` emits) both reduce to the same name. A worksheet can own several tables, hence a `Vec` -- unlike
/// `parse_sheet_drawing_rels`, which stops at the first match because a
/// worksheet has at most one drawing.
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

/// [LLM-generated] One Excel Table (ListObject) as declared by its own `xl/tables/*.xml`
/// part. Everything here is read straight out of that XML rather than
/// through calamine: `Xlsx::table_by_name` panics outright on a table with a
/// header row and zero data rows (it derives a data-only range whose start
/// is past its end), which is a shape this crate's own export produces. See
/// `import_tables_from_zip`.
struct ParsedTablePart {
    name: String,
    columns: Vec<String>,
    has_header_row: bool,
    has_totals_row: bool,
    /// [LLM-generated] The whole table's bounds from its declared `ref`, header and totals
    /// rows included, 0-based `(start_row, start_col, end_row, end_col)`.
    /// `None` when the `ref` attribute is absent or unparseable, in which
    /// case the table can't be placed and is skipped.
    bounds: Option<(usize, usize, usize, usize)>,
    style_name: Option<String>,
    has_insert_row: bool,
}

/// [LLM-generated] Parses one `xl/tables/tableN.xml` part. Returns `None` if the XML has no
/// `<table>` element with a `displayName`, which is the only field a table
/// can't sensibly be reconstructed without.
///
/// `headerRowCount` defaults to 1 and `totalsRowCount` to 0 when absent,
/// matching both the OOXML default and how most real-world tables are
/// configured -- `rust_xlsxwriter` omits both attributes for an ordinary
/// header-and-data table.
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

/// [LLM-generated] Discovers every Excel Table in the workbook by walking the zip directly,
/// pairing each with the name of the sheet that owns it.
///
/// This deliberately replaces calamine's `load_tables`/`table_by_name`. That
/// API panics ("invalid range bounds") on a table with a header row and zero
/// data rows -- a shape `export_xlsx_data` itself produces.
/// Reading the parts here avoids that issue; calamine is now used only for cell data.
///
/// Within a sheet, tables keep their relationship order. Across sheets the
/// order is irrelevant -- each table is attached to its own sheet -- but the
/// worksheet parts are still walked in sorted order so that iteration
/// doesn't inherit `HashMap`'s randomized ordering.
fn import_tables_from_zip(buffer: &[u8]) -> Result<Vec<(String, ParsedTablePart)>, String> {
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(buffer))
        .map_err(|e| format!("Failed to open zip: {}", e))?;

    let sheet_file_to_name = build_sheet_file_to_name(&mut archive)?;

    let mut sheet_files: Vec<&String> = sheet_file_to_name.keys().collect();
    sheet_files.sort();
    let sheet_files: Vec<String> = sheet_files.into_iter().cloned().collect();

    let mut tables = Vec::new();
    for sheet_file in sheet_files {
        let rels_path = format!("xl/worksheets/_rels/{}.rels", sheet_file);
        let Some(rels_xml) = get_zip_file_content(&mut archive, &rels_path) else {
            continue;
        };
        let Some(sheet_name) = sheet_file_to_name.get(&sheet_file).cloned() else {
            continue;
        };
        for table_file in parse_sheet_table_rels(&rels_xml) {
            let table_path = format!("xl/tables/{}", table_file);
            if let Some(table_xml) = get_zip_file_content(&mut archive, &table_path)
                && let Some(parsed) = parse_table_part_xml(&table_xml)
            {
                tables.push((sheet_name.clone(), parsed));
            }
        }
    }
    Ok(tables)
}

/// [LLM-generated] Built-in OOXML `numFmtId` codes whose display strings can be preserved on import.
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

/// [LLM-generated] Date format codes per cell, keyed by sheet name and then by 0-based
/// `(row, col)`.
type SheetCellNumberFormats =
    std::collections::HashMap<String, std::collections::HashMap<(usize, usize), String>>;

type SheetCellStyles =
    std::collections::HashMap<String, std::collections::HashMap<(usize, usize), CellStyle>>;

type SheetColumnWidths = std::collections::HashMap<String, std::collections::HashMap<usize, f64>>;

type SheetRowHeights = std::collections::HashMap<String, std::collections::HashMap<usize, f64>>;

fn import_column_widths(buffer: &[u8]) -> Result<SheetColumnWidths, String> {
    use std::collections::HashMap;

    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(buffer))
        .map_err(|e| format!("Failed to open zip: {}", e))?;
    let sheet_file_to_name = build_sheet_file_to_name(&mut archive)?;

    let mut sheet_files: Vec<String> = sheet_file_to_name.keys().cloned().collect();
    sheet_files.sort();

    let mut out: SheetColumnWidths = HashMap::new();
    for sheet_file in sheet_files {
        let Some(sheet_name) = sheet_file_to_name.get(&sheet_file).cloned() else {
            continue;
        };
        let path = format!("xl/worksheets/{}", sheet_file);
        let Some(sheet_xml) = get_zip_file_content(&mut archive, &path) else {
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

fn import_row_heights(buffer: &[u8]) -> Result<SheetRowHeights, String> {
    use std::collections::HashMap;

    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(buffer))
        .map_err(|e| format!("Failed to open zip: {}", e))?;
    let sheet_file_to_name = build_sheet_file_to_name(&mut archive)?;

    let mut sheet_files: Vec<String> = sheet_file_to_name.keys().cloned().collect();
    sheet_files.sort();

    let mut out: SheetRowHeights = HashMap::new();
    for sheet_file in sheet_files {
        let Some(sheet_name) = sheet_file_to_name.get(&sheet_file).cloned() else {
            continue;
        };
        let path = format!("xl/worksheets/{}", sheet_file);
        let Some(sheet_xml) = get_zip_file_content(&mut archive, &path) else {
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

fn import_cell_styles(buffer: &[u8]) -> Result<SheetCellStyles, String> {
    use std::collections::HashMap;

    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(buffer))
        .map_err(|e| format!("Failed to open zip: {}", e))?;

    let Some(styles_xml) = get_zip_file_content(&mut archive, "xl/styles.xml") else {
        return Ok(HashMap::new());
    };
    let xf_to_style = parse_styles_cell_styles(&styles_xml);
    if xf_to_style.is_empty() {
        return Ok(HashMap::new());
    }

    let sheet_file_to_name = build_sheet_file_to_name(&mut archive)?;
    let mut sheet_files: Vec<String> = sheet_file_to_name.keys().cloned().collect();
    sheet_files.sort();

    let mut out: HashMap<String, HashMap<(usize, usize), CellStyle>> = HashMap::new();
    for sheet_file in sheet_files {
        let Some(sheet_name) = sheet_file_to_name.get(&sheet_file).cloned() else {
            continue;
        };
        let path = format!("xl/worksheets/{}", sheet_file);
        let Some(sheet_xml) = get_zip_file_content(&mut archive, &path) else {
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

/// [LLM-generated] Maps each cell that carries a date number format to that format's code,
/// keyed by sheet name and then by `(row, col)` -- both 0-based, matching the
/// engine.
///
/// calamine reports a date-formatted cell as `Data::DateTime` but does not
/// expose the format code behind it, and the code is the whole point here: it
/// is what lets `6/22/26` come back as `6/22/26` rather than `2026-06-22`. So
/// this walks the zip directly, joining each worksheet's per-cell style index
/// (`<c s="3">`) through `xl/styles.xml`'s `<cellXfs>` to a `numFmtId`, and
/// then to either a custom `<numFmt>` code or a built-in one.
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

/// [LLM-generated] Resolves `xl/styles.xml` into "cell style index -> date format code",
/// keeping only the entries that denote a date.
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

/// [LLM-generated] Pulls `(row, col) -> date format code` out of one worksheet part, for the
/// cells whose style index resolves to a date format.
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

/// [LLM-generated] `"B3"` -> `(2, 1)`, 0-based. Trailing `$` anchors are not expected in a
/// cell's `r` attribute and are not accepted.
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

/// [LLM-generated] Derives a stable chart id from its sheet name and position within that
/// sheet's charts, so re-importing the same unchanged xlsx always assigns
/// the same id to the same chart. Uses `DefaultHasher`, which (unlike
/// `HashMap`'s default `RandomState`) is not seeded per-process, so this is
/// deterministic across separate CLI invocations of the same binary.
fn deterministic_chart_id(sheet_name: &str, index_in_sheet: usize) -> u64 {
    use std::collections::hash_map::DefaultHasher;
    use std::hash::{Hash, Hasher};
    let mut hasher = DefaultHasher::new();
    sheet_name.hash(&mut hasher);
    index_in_sheet.hash(&mut hasher);
    hasher.finish() & 0x001F_FFFF_FFFF_FFFF
}

/// [LLM-generated] Maps each worksheet part's bare filename (`sheet1.xml`) to the sheet name
/// the workbook declares for it, by joining `xl/workbook.xml`'s
/// `<sheet name= r:id=>` entries against `xl/_rels/workbook.xml.rels`. The
/// part filenames are not in workbook order and carry no reliable
/// relationship to the sheet's position, so this join is the only way to get
/// from a part back to its name. Shared by the chart and table importers.
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

fn import_charts_from_zip(buffer: &[u8]) -> Result<Vec<ParsedChartData>, String> {
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(buffer))
        .map_err(|e| format!("Failed to open zip: {}", e))?;

    let sheet_file_to_name = build_sheet_file_to_name(&mut archive)?;

    let mut parsed_charts = Vec::new();

    for sheet_filename in sheet_file_to_name.keys() {
        let rels_path = format!("xl/worksheets/_rels/{}.rels", sheet_filename);
        if let Some(rels_xml) = get_zip_file_content(&mut archive, &rels_path)
            && let Some(drawing_filename) = parse_sheet_drawing_rels(&rels_xml)
        {
            let sheet_name = sheet_file_to_name.get(sheet_filename).unwrap().clone();

            let drawing_rels_path = format!("xl/drawings/_rels/{}.rels", drawing_filename);
            let chart_rid_to_filename = if let Some(drawing_rels_xml) =
                get_zip_file_content(&mut archive, &drawing_rels_path)
            {
                parse_drawing_chart_rels(&drawing_rels_xml)
            } else {
                std::collections::HashMap::new()
            };

            let drawing_path = format!("xl/drawings/{}", drawing_filename);
            if let Some(drawing_xml) = get_zip_file_content(&mut archive, &drawing_path) {
                let anchors = parse_drawings_xml(&drawing_xml);

                for anchor in anchors {
                    if let Some(chart_filename) = chart_rid_to_filename.get(&anchor.chart_rid) {
                        let chart_path = format!("xl/charts/{}", chart_filename);
                        if let Some(chart_xml) = get_zip_file_content(&mut archive, &chart_path)
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

    /// [LLM-generated] A date cell has to survive as a *date*: the value goes out as a
    /// numeric serial (Excel cannot do date arithmetic on text) while the
    /// notation goes out as the cell's `numFmt` and comes back from it. The
    /// CLI is a fresh process per invocation, so this round trip is the only
    /// thing that makes a date still look like one on the next command.
    /// A worksheet string cell that arrives as the empty string has to stay
    /// a *text* cell rather than becoming blank: Excel reports TYPE 2 and a
    /// text comparison for it, and a fuzz grid containing one disagreed with
    /// Excel in three separate formulas while visi rebuilt it as blank.
    ///
    /// This is reached in practice via whitespace: OOXML strips whitespace-only
    /// `<t>` content that is not marked `xml:space="preserve"` -- which
    /// `openpyxl` omits -- so calamine reports what is left as `String("")`
    /// rather than as `Data::Empty`. Excel strips it the same way and likewise
    /// keeps a text cell. (visi's own writer *does* emit `xml:space`, so a
    /// space survives a visi-to-visi round trip; this covers the other case.)
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

    /// [LLM-generated] Text that merely looks like a date must not become one on import --
    /// Excel handed it over as a string cell, so it stays a string.
    #[test]
    fn test_xlsx_date_looking_text_stays_text() {
        let col = DataColumn::from_src("A", vec!["\"22-Jun\"".to_string()]);
        let sheet = Sheet {
            id: 1,
            name: "Sheet1".to_string(),
            columns: vec![col],
            row_heights: Vec::new(),
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

    /// [LLM-generated] Rewrites the `ref` attribute of every `xl/tables/*.xml` part in an
    /// exported workbook, leaving the rest of the zip byte-for-byte intact.
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

    /// [LLM-generated] The emitted pivot XML, as one string per part, for shape assertions.
    ///
    /// These stand in for a check CI cannot run: the only authority on
    /// whether Excel accepts a pivot part is Excel, and
    /// `fuzz/pivot_filter_probe.py --variant visi` is what actually asks it.
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
}
