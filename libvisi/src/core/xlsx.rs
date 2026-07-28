use crate::core::{DataColumn, Sheet};
use crate::render::Position;
use calamine::Reader;
use web_time::Instant;

#[derive(Debug, serde::Serialize, serde::Deserialize)]
pub struct ImportedSheet {
    pub sheet: Sheet,
    pub pos_preserved: bool,
}

fn encode_sheet_name_to_hex(name: &str) -> String {
    name.as_bytes()
        .iter()
        .map(|b| format!("{:02x}", b))
        .collect()
}

fn decode_hex_sheet_name(hex_str: &str) -> Option<String> {
    let mut bytes = Vec::new();
    for i in (0..hex_str.len()).step_by(2) {
        if i + 2 <= hex_str.len() {
            let byte_str = &hex_str[i..i + 2];
            if let Ok(byte) = u8::from_str_radix(byte_str, 16) {
                bytes.push(byte);
            } else {
                return None;
            }
        } else {
            return None;
        }
    }
    String::from_utf8(bytes).ok()
}

pub fn import_xlsx_data(
    buffer: &[u8],
    target_row: isize,
    target_col: isize,
    existing_sheets: &[Sheet],
    mut progress_callback: impl FnMut(usize, usize, &str),
) -> Result<(Vec<ImportedSheet>, Vec<crate::core::chart::Chart>), String> {
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

    // Parse sheet positions and sizes from defined names
    let start_defined_names = Instant::now();
    let mut sheet_positions: std::collections::HashMap<String, (Option<isize>, Option<isize>)> =
        std::collections::HashMap::new();
    let mut sheet_sizes = std::collections::HashMap::new();
    let mut sheet_z_indexes = std::collections::HashMap::new();
    for (name, formula) in workbook.defined_names() {
        if name.starts_with("GP_POS_X_") {
            let hex_part = &name[9..];
            if let Some(sheet_name) = decode_hex_sheet_name(hex_part) {
                let formula_clean = if formula.starts_with('=') {
                    &formula[1..]
                } else {
                    formula
                };
                if let Ok(x) = formula_clean.parse::<isize>() {
                    let entry = sheet_positions.entry(sheet_name).or_insert((None, None));
                    entry.0 = Some(x);
                }
            }
        } else if name.starts_with("GP_POS_Y_") {
            let hex_part = &name[9..];
            if let Some(sheet_name) = decode_hex_sheet_name(hex_part) {
                let formula_clean = if formula.starts_with('=') {
                    &formula[1..]
                } else {
                    formula
                };
                if let Ok(y) = formula_clean.parse::<isize>() {
                    let entry = sheet_positions.entry(sheet_name).or_insert((None, None));
                    entry.1 = Some(y);
                }
            }
        } else if name.starts_with("GP_SIZE_ROWS_") {
            let hex_part = &name[13..];
            if let Some(sheet_name) = decode_hex_sheet_name(hex_part) {
                let formula_clean = if formula.starts_with('=') {
                    &formula[1..]
                } else {
                    formula
                };
                if let Ok(rows) = formula_clean.parse::<usize>() {
                    let entry = sheet_sizes.entry(sheet_name).or_insert((None, None));
                    entry.0 = Some(rows);
                }
            }
        } else if name.starts_with("GP_SIZE_COLS_") {
            let hex_part = &name[13..];
            if let Some(sheet_name) = decode_hex_sheet_name(hex_part) {
                let formula_clean = if formula.starts_with('=') {
                    &formula[1..]
                } else {
                    formula
                };
                if let Ok(cols) = formula_clean.parse::<usize>() {
                    let entry = sheet_sizes.entry(sheet_name).or_insert((None, None));
                    entry.1 = Some(cols);
                }
            }
        } else if name.starts_with("GP_Z_INDEX_") {
            let hex_part = &name[11..];
            if let Some(sheet_name) = decode_hex_sheet_name(hex_part) {
                let formula_clean = if formula.starts_with('=') {
                    &formula[1..]
                } else {
                    formula
                };
                if let Ok(z) = formula_clean.parse::<i32>() {
                    sheet_z_indexes.insert(sheet_name, z);
                }
            }
        }
    }
    log::info!(
        "Excel defined names parsing took: {:.2?}",
        start_defined_names.elapsed()
    );

    let sheet_names = workbook.sheet_names();
    let total_sheets = sheet_names.len();
    let mut imported_tables = Vec::new();

    let mut current_col = target_col;

    for (sheet_idx, orig_sheet_name) in sheet_names.iter().enumerate() {
        progress_callback(sheet_idx, total_sheets, orig_sheet_name);
        let start_sheet = Instant::now();
        let start_range = Instant::now();
        if let Ok(range) = workbook.worksheet_range(orig_sheet_name) {
            let elapsed_range = start_range.elapsed();
            let mut rows = range.height();
            let mut cols = range.width();

            // Override size if it was preserved in defined names
            if let Some(&(Some(r), Some(c))) = sheet_sizes.get(orig_sheet_name) {
                rows = r;
                cols = c;
            }

            // Fallback for empty worksheets (e.g. from external Excel or empty sheets)
            if rows == 0 || cols == 0 {
                rows = 10;
                cols = 5;
            }

            // Check if position was preserved in defined names
            let mut pos_preserved = false;
            let mut pos_left = current_col;
            let mut pos_top = target_row;

            if let Some(&(Some(x), Some(y))) = sheet_positions.get(orig_sheet_name) {
                pos_left = x;
                pos_top = y;
                pos_preserved = true;
            }

            // Unique sheet name
            let mut sheet_name = orig_sheet_name.clone();
            let mut count = 1;
            while existing_sheets.iter().any(|t| t.name == sheet_name)
                || imported_tables
                    .iter()
                    .any(|t: &ImportedSheet| t.sheet.name == sheet_name)
            {
                sheet_name = format!("{}_{}", sheet_name, count);
                count += 1;
            }

            let start_formula = Instant::now();
            let formula_range = workbook.worksheet_formula(orig_sheet_name).ok();
            let elapsed_formula = start_formula.elapsed();

            let start_cells = Instant::now();
            let mut columns = (0..cols)
                .map(|col_idx| {
                    let mut col = DataColumn::new(rows);
                    col.id = rand::random::<u64>();
                    col.name = crate::core::parser::col_idx_to_letters(col_idx);
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
                    .map_or(true, |fr| fr.height() == rows && fr.width() == cols);

            if matches_exactly {
                let mut range_rows = range.rows();
                let mut formula_rows = formula_range.as_ref().map(|fr| fr.rows());

                for row_idx in 0..rows {
                    let r_cells = range_rows.next().unwrap();
                    let f_cells = formula_rows.as_mut().map(|fr| fr.next().unwrap());

                    for col_idx in 0..cols {
                        if let Some(f_cells_slice) = f_cells {
                            let formula = &f_cells_slice[col_idx];
                            if !formula.is_empty() {
                                let cell_src = format!("={}", formula);
                                max_cell_lens[col_idx] = max_cell_lens[col_idx].max(cell_src.len());
                                columns[col_idx].src[row_idx] = cell_src;
                                continue;
                            }
                        }

                        let cell_value = &r_cells[col_idx];
                        let cell_src = match cell_value {
                            calamine::Data::Empty => String::new(),
                            calamine::Data::String(s) => s.clone(),
                            calamine::Data::Float(f) => f.to_string(),
                            calamine::Data::Int(i) => i.to_string(),
                            calamine::Data::Bool(b) => b.to_string(),
                            calamine::Data::Error(e) => format!("#ERR: {:?}", e),
                            calamine::Data::DateTime(d) => d.to_string(),
                            _ => String::new(),
                        };
                        max_cell_lens[col_idx] = max_cell_lens[col_idx].max(cell_src.len());
                        columns[col_idx].src[row_idx] = cell_src;
                    }
                }
            } else {
                for row_idx in 0..rows {
                    let row_u32 = row_idx as u32;
                    for col_idx in 0..cols {
                        let col_u32 = col_idx as u32;

                        if let Some(ref f_range) = formula_range {
                            if let Some(formula) = f_range.get_value((row_u32, col_u32)) {
                                if !formula.is_empty() {
                                    let cell_src = format!("={}", formula);
                                    max_cell_lens[col_idx] =
                                        max_cell_lens[col_idx].max(cell_src.len());
                                    columns[col_idx].src[row_idx] = cell_src;
                                    continue;
                                }
                            }
                        }

                        let cell_value = range
                            .get_value((row_u32, col_u32))
                            .unwrap_or(&calamine::Data::Empty);
                        let cell_src = match cell_value {
                            calamine::Data::Empty => String::new(),
                            calamine::Data::String(s) => s.clone(),
                            calamine::Data::Float(f) => f.to_string(),
                            calamine::Data::Int(i) => i.to_string(),
                            calamine::Data::Bool(b) => b.to_string(),
                            calamine::Data::Error(e) => format!("#ERR: {:?}", e),
                            calamine::Data::DateTime(d) => d.to_string(),
                            _ => String::new(),
                        };
                        max_cell_lens[col_idx] = max_cell_lens[col_idx].max(cell_src.len());
                        columns[col_idx].src[row_idx] = cell_src;
                    }
                }
            }
            let elapsed_cells = start_cells.elapsed();

            let mut table_width = 0;
            for col_idx in 0..cols {
                let col_width = (max_cell_lens[col_idx] / 4 + 1).max(2);
                table_width += col_width;
            }

            let z_index = sheet_z_indexes.get(orig_sheet_name).copied().unwrap_or(0);

            let new_table = Sheet {
                id: rand::random::<u64>(),
                name: sheet_name,
                columns,
                pos: Position {
                    left: pos_left,
                    top: pos_top,
                },
                z_index,
                dependencies: std::collections::HashMap::new(),
                dependencies_rev: std::collections::HashMap::new(),
                uncommitted_actions: Vec::new(),
                row_tops_cache: std::cell::RefCell::new(None),
            };
            current_col += (table_width + 2) as isize;
            imported_tables.push(ImportedSheet {
                sheet: new_table,
                pos_preserved,
            });

            log::info!(
                "Worksheet '{}' ({}x{} cells) processed. range_parse: {:.2?}, formula_parse: {:.2?}, cells_convert: {:.2?}, total_sheet: {:.2?}",
                imported_tables.last().unwrap().sheet.name,
                rows,
                cols,
                elapsed_range,
                elapsed_formula,
                elapsed_cells,
                start_sheet.elapsed()
            );
        }
    }

    if imported_tables.is_empty() {
        return Err("No worksheets found in the Excel file".to_string());
    }

    log::info!(
        "import_xlsx_data finished. Total time: {:.2?}",
        start_total.elapsed()
    );

    // Import charts from drawings and charts xml in ZIP
    let mut imported_charts = Vec::new();
    if let Ok(parsed_charts) = import_charts_from_zip(buffer) {
        for parsed in parsed_charts {
            if let Some(sheet) = imported_tables
                .iter()
                .find(|t| t.sheet.name == parsed.sheet_name)
            {
                let from_anchor = crate::core::chart::ChartAnchor {
                    col: sheet.sheet.pos.left + parsed.anchor.from_col,
                    row: sheet.sheet.pos.top + parsed.anchor.from_row,
                    col_offset_pixels: parsed.anchor.from_col_off,
                    row_offset_pixels: parsed.anchor.from_row_off,
                };

                let to_anchor = crate::core::chart::ChartAnchor {
                    col: sheet.sheet.pos.left + parsed.anchor.to_col,
                    row: sheet.sheet.pos.top + parsed.anchor.to_row,
                    col_offset_pixels: parsed.anchor.to_col_off,
                    row_offset_pixels: parsed.anchor.to_row_off,
                };

                let data_range = if parsed.info.data_range.is_empty() {
                    format!("{}!A1", sheet.sheet.name)
                } else {
                    parsed.info.data_range.clone()
                };

                imported_charts.push(crate::core::chart::Chart {
                    id: rand::random::<u64>(),
                    name: format!("Chart {}", imported_charts.len() + 1),
                    chart_type: parsed.info.chart_type.clone(),
                    data_range,
                    from_anchor,
                    to_anchor,
                    title: parsed.info.title.clone(),
                    xlabel: parsed.info.xlabel.clone(),
                    ylabel: parsed.info.ylabel.clone(),
                    show_legend: parsed.info.show_legend,
                    z_index: 0,
                });
            }
        }
    }

    Ok((imported_tables, imported_charts))
}

pub fn export_xlsx_data(
    sheets: &[Sheet],
    charts: &[crate::core::chart::Chart],
) -> Result<Vec<u8>, String> {
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let mut used_names = std::collections::HashSet::new();
    let mut table_name_to_worksheet_name = std::collections::HashMap::new();
    let mut table_name_to_table = std::collections::HashMap::new();

    for sheet in sheets {
        table_name_to_table.insert(sheet.name.clone(), sheet);

        // 31-char limit for worksheet name
        let mut worksheet_name = if sheet.name.len() > 31 {
            sheet.name[..31].to_string()
        } else {
            sheet.name.clone()
        };

        // Ensure unique worksheet name (case-insensitive)
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

        {
            let worksheet = workbook.add_worksheet();
            worksheet
                .set_name(&worksheet_name)
                .map_err(|e| format!("Failed to set worksheet name: {}", e))?;

            for col_idx in 0..sheet.columns.len() {
                let col = &sheet.columns[col_idx];
                for row_idx in 0..sheet.row_count() {
                    let cell_src = col.src.get(row_idx).cloned().unwrap_or_default();
                    if cell_src.starts_with('=') {
                        worksheet
                            .write_formula(row_idx as u32, col_idx as u16, &cell_src[1..])
                            .map_err(|e| format!("Failed to write Excel formula: {}", e))?;
                    } else if let Ok(val_f64) = cell_src.parse::<f64>() {
                        worksheet
                            .write_number(row_idx as u32, col_idx as u16, val_f64)
                            .map_err(|e| format!("Failed to write Excel number: {}", e))?;
                    } else if let Ok(val_bool) = cell_src.parse::<bool>() {
                        worksheet
                            .write_boolean(row_idx as u32, col_idx as u16, val_bool)
                            .map_err(|e| format!("Failed to write Excel boolean: {}", e))?;
                    } else if !cell_src.is_empty() {
                        worksheet
                            .write_string(row_idx as u32, col_idx as u16, &cell_src)
                            .map_err(|e| format!("Failed to write Excel string: {}", e))?;
                    }
                }
            }
        }

        // Define position names globally using hex encoding of the final worksheet name
        let hex_name = encode_sheet_name_to_hex(&worksheet_name);
        workbook
            .define_name(
                &format!("GP_POS_X_{}", hex_name),
                &format!("={}", sheet.pos.left),
            )
            .map_err(|e| format!("Failed to define name GP_POS_X: {}", e))?;
        workbook
            .define_name(
                &format!("GP_POS_Y_{}", hex_name),
                &format!("={}", sheet.pos.top),
            )
            .map_err(|e| format!("Failed to define name GP_POS_Y: {}", e))?;
        workbook
            .define_name(
                &format!("GP_SIZE_ROWS_{}", hex_name),
                &format!("={}", sheet.row_count()),
            )
            .map_err(|e| format!("Failed to define name GP_SIZE_ROWS: {}", e))?;
        workbook
            .define_name(
                &format!("GP_SIZE_COLS_{}", hex_name),
                &format!("={}", sheet.columns.len()),
            )
            .map_err(|e| format!("Failed to define name GP_SIZE_COLS: {}", e))?;
        workbook
            .define_name(
                &format!("GP_Z_INDEX_{}", hex_name),
                &format!("={}", sheet.z_index),
            )
            .map_err(|e| format!("Failed to define name GP_Z_INDEX: {}", e))?;
    }

    // Export charts
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

        if let Some(ws_name) = ws_name {
            if let Ok(worksheet) = workbook.worksheet_from_name(&ws_name) {
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
                if let Some(ref xlabel) = chart.xlabel {
                    rx_chart.x_axis().set_name(xlabel);
                }
                if let Some(ref ylabel) = chart.ylabel {
                    rx_chart.y_axis().set_name(ylabel);
                }
                if !chart.show_legend {
                    rx_chart.legend().set_hidden();
                }

                let tile_w = 80.0f32;
                let tile_h = 24.0f32;
                let width_tiles = (chart.to_anchor.col - chart.from_anchor.col) as f32;
                let height_tiles = (chart.to_anchor.row - chart.from_anchor.row) as f32;
                let width_px = (width_tiles * tile_w
                    + (chart.to_anchor.col_offset_pixels - chart.from_anchor.col_offset_pixels))
                    .max(100.0) as u32;
                let height_px = (height_tiles * tile_h
                    + (chart.to_anchor.row_offset_pixels - chart.from_anchor.row_offset_pixels))
                    .max(100.0) as u32;

                rx_chart.set_width(width_px);
                rx_chart.set_height(height_px);

                let table_pos = table_name_to_table
                    .get(&tbl_name)
                    .map(|t| t.pos.clone())
                    .unwrap_or(Position { left: 0, top: 0 });

                let insert_row = (chart.from_anchor.row - table_pos.top).max(0) as u32;
                let insert_col = (chart.from_anchor.col - table_pos.left).max(0) as u16;
                let insert_x_offset = chart.from_anchor.col_offset_pixels.max(0.0) as u32;
                let insert_y_offset = chart.from_anchor.row_offset_pixels.max(0.0) as u32;

                worksheet
                    .insert_chart_with_offset(
                        insert_row,
                        insert_col,
                        &rx_chart,
                        insert_x_offset,
                        insert_y_offset,
                    )
                    .map_err(|e| format!("Failed to insert Excel chart: {}", e))?;
            }
        }
    }

    workbook
        .save_to_buffer()
        .map_err(|e| format!("Failed to write XLSX buffer: {}", e))
}

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
    path.split('/').last().unwrap_or("").to_string()
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
    if let (Some(cat), Some(val)) = (parse_range(cat_f), parse_range(val_f)) {
        if cat.sheet == val.sheet && cat.start_row == val.start_row && cat.end_row == val.end_row {
            return format!(
                "{}!{}{}:{}{}",
                cat.sheet, cat.start_col, cat.start_row, val.end_col, val.end_row
            );
        }
    }
    clean_excel_formula(val_f)
}

fn get_attr(e: &quick_xml::events::BytesStart, name: &[u8]) -> Option<String> {
    for attr_res in e.attributes() {
        if let Ok(attr) = attr_res {
            let key = attr.key.as_ref();
            let local_name = if let Some(pos) = key.iter().position(|&b| b == b':') {
                &key[pos + 1..]
            } else {
                key
            };
            if local_name == name {
                return String::from_utf8(attr.value.into_owned()).ok();
            }
        }
    }
    None
}

fn parse_workbook_sheets(xml: &str) -> std::collections::HashMap<String, String> {
    let mut map = std::collections::HashMap::new();
    let mut reader = quick_xml::reader::Reader::from_str(xml);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Eof) => break,
            Ok(quick_xml::events::Event::Empty(e)) | Ok(quick_xml::events::Event::Start(e)) => {
                let local = e.name().local_name().into_inner();
                if local == b"sheet" {
                    if let (Some(name), Some(r_id)) = (get_attr(&e, b"name"), get_attr(&e, b"id")) {
                        map.insert(r_id, name);
                    }
                }
            }
            _ => {}
        }
        buf.clear();
    }
    map
}

fn parse_workbook_rels(xml: &str) -> std::collections::HashMap<String, String> {
    let mut map = std::collections::HashMap::new();
    let mut reader = quick_xml::reader::Reader::from_str(xml);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Eof) => break,
            Ok(quick_xml::events::Event::Empty(e)) | Ok(quick_xml::events::Event::Start(e)) => {
                let local = e.name().local_name().into_inner();
                if local == b"Relationship" {
                    if let (Some(r_id), Some(target)) =
                        (get_attr(&e, b"Id"), get_attr(&e, b"Target"))
                    {
                        map.insert(r_id, target);
                    }
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
                if local == b"Relationship" {
                    if let Some(ty) = get_attr(&e, b"Type") {
                        if ty.contains("relationships/drawing") {
                            if let Some(target) = get_attr(&e, b"Target") {
                                drawing_target = Some(get_filename(&target));
                                break;
                            }
                        }
                    }
                }
            }
            _ => {}
        }
        buf.clear();
    }
    drawing_target
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
                if local == b"Relationship" {
                    if let Some(ty) = get_attr(&e, b"Type") {
                        if ty.contains("relationships/chart") {
                            if let (Some(r_id), Some(target)) =
                                (get_attr(&e, b"Id"), get_attr(&e, b"Target"))
                            {
                                map.insert(r_id, get_filename(&target));
                            }
                        }
                    }
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
                } else if local == b"chart" {
                    if let Some(rid) = get_attr(&e, b"id") {
                        chart_rid = rid;
                    }
                }
                current_tag = local.to_vec();
            }
            Ok(quick_xml::events::Event::Empty(e)) => {
                let local = e.name().local_name().into_inner();
                if local == b"chart" {
                    if let Some(rid) = get_attr(&e, b"id") {
                        chart_rid = rid;
                    }
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
                if let Ok(text) = e.unescape() {
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
    let mut ylabel = None;
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
                } else if local == b"valAx" {
                    in_val_ax = true;
                } else if local == b"ser" {
                    in_ser = true;
                } else if local == b"cat" {
                    in_cat = true;
                } else if local == b"val" {
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
                } else if local == b"barDir" {
                    if let Some(val) = get_attr(&e, b"val") {
                        if val == "col" {
                            chart_type = crate::core::chart::ChartType::Column;
                        } else if val == "bar" {
                            chart_type = crate::core::chart::ChartType::Bar;
                        }
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
                } else if local == b"cat" {
                    in_cat = false;
                } else if local == b"val" {
                    in_val = false;
                }
                current_tag.clear();
            }
            Ok(quick_xml::events::Event::Text(e)) => {
                if let Ok(text) = e.unescape() {
                    let val_str = text.trim();
                    if !val_str.is_empty() {
                        if in_title {
                            if current_tag == b"v" || current_tag == b"t" {
                                if in_cat_ax {
                                    xlabel = Some(val_str.to_string());
                                } else if in_val_ax {
                                    ylabel = Some(val_str.to_string());
                                } else {
                                    title = Some(val_str.to_string());
                                }
                            }
                        } else if in_ser {
                            if current_tag == b"f" {
                                if in_cat {
                                    cat_f = val_str.to_string();
                                } else if in_val {
                                    val_f = val_str.to_string();
                                }
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

    Some(ParsedChartInfo {
        chart_type,
        title,
        xlabel,
        ylabel,
        show_legend,
        data_range,
    })
}

fn get_zip_file_content(
    archive: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>,
    name: &str,
) -> Option<String> {
    let mut file = archive.by_name(name).ok()?;
    let mut content = String::new();
    std::io::Read::read_to_string(&mut file, &mut content).ok()?;
    Some(content)
}

fn import_charts_from_zip(buffer: &[u8]) -> Result<Vec<ParsedChartData>, String> {
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(buffer))
        .map_err(|e| format!("Failed to open zip: {}", e))?;

    let workbook_xml = get_zip_file_content(&mut archive, "xl/workbook.xml")
        .ok_or_else(|| "Missing xl/workbook.xml".to_string())?;
    let r_id_to_name = parse_workbook_sheets(&workbook_xml);

    let workbook_rels_xml = get_zip_file_content(&mut archive, "xl/_rels/workbook.xml.rels")
        .ok_or_else(|| "Missing xl/_rels/workbook.xml.rels".to_string())?;
    let r_id_to_target = parse_workbook_rels(&workbook_rels_xml);

    let mut sheet_file_to_name = std::collections::HashMap::new();
    for (r_id, name) in r_id_to_name {
        if let Some(target) = r_id_to_target.get(&r_id) {
            let filename = get_filename(target);
            sheet_file_to_name.insert(filename, name.clone());
        }
    }

    let mut parsed_charts = Vec::new();

    for sheet_filename in sheet_file_to_name.keys() {
        let rels_path = format!("xl/worksheets/_rels/{}.rels", sheet_filename);
        if let Some(rels_xml) = get_zip_file_content(&mut archive, &rels_path) {
            if let Some(drawing_filename) = parse_sheet_drawing_rels(&rels_xml) {
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
                            {
                                if let Some(info) = parse_chart_xml(&chart_xml) {
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
        }
    }

    Ok(parsed_charts)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_xlsx_import_export_cycle() {
        // Create a mock sheet
        let mut columns = Vec::new();
        let mut col1 = DataColumn::new(2);
        col1.name = "A".to_string();
        col1.src = vec!["10".to_string(), "20".to_string()].into();
        columns.push(col1);

        let mut col2 = DataColumn::new(2);
        col2.name = "B".to_string();
        col2.src = vec!["=A1 + A2".to_string(), "abc".to_string()].into();
        columns.push(col2);

        let sheet = Sheet {
            id: 1,
            name: "Sheet1".to_string(),
            columns,
            pos: Position { left: 5, top: 8 },
            z_index: 42,
            dependencies: std::collections::HashMap::new(),
            dependencies_rev: std::collections::HashMap::new(),
            uncommitted_actions: Vec::new(),
            row_tops_cache: std::cell::RefCell::new(None),
        };

        // Export it
        let xlsx_data = export_xlsx_data(&[sheet], &[]).unwrap();
        assert!(!xlsx_data.is_empty());

        // Import it back
        let (imported_tables, imported_charts) =
            import_xlsx_data(&xlsx_data, 10, 20, &[], |_, _, _| {}).unwrap();
        assert_eq!(imported_tables.len(), 1);
        assert_eq!(imported_charts.len(), 0);

        let imported_table = &imported_tables[0].sheet;
        assert!(imported_tables[0].pos_preserved);
        assert_eq!(imported_table.name, "Sheet1");
        assert_eq!(imported_table.pos.left, 5);
        assert_eq!(imported_table.pos.top, 8);
        assert_eq!(imported_table.z_index, 42);
        assert_eq!(imported_table.columns.len(), 2);

        // Verify content
        assert_eq!(imported_table.columns[0].src[0], "10");
        assert_eq!(imported_table.columns[0].src[1], "20");
        assert_eq!(imported_table.columns[1].src[0], "=A1 + A2");
        assert_eq!(imported_table.columns[1].src[1], "abc");
    }

    #[test]
    fn test_defined_names_cycle() {
        let mut workbook = rust_xlsxwriter::Workbook::new();

        let worksheet1 = workbook.add_worksheet();
        worksheet1.set_name("SheetOne").unwrap();
        workbook.define_name("SheetOne!GP_POS_X", "=12").unwrap();
        workbook.define_name("SheetOne!GP_POS_Y", "=34").unwrap();

        let worksheet2 = workbook.add_worksheet();
        worksheet2.set_name("SheetTwo").unwrap();
        workbook.define_name("SheetTwo!GP_POS_X", "=56").unwrap();
        workbook.define_name("SheetTwo!GP_POS_Y", "=78").unwrap();

        workbook.define_name("GP_GLOBAL_VAL", "=100").unwrap();

        let xlsx_data = workbook.save_to_buffer().unwrap();

        let cursor = std::io::Cursor::new(xlsx_data);
        let workbook_read: calamine::Xlsx<_> = calamine::open_workbook_from_rs(cursor).unwrap();

        let mut found_pos_x1 = false;
        let mut found_pos_y1 = false;
        let mut found_pos_x2 = false;
        let mut found_pos_y2 = false;

        for (name, formula) in workbook_read.defined_names() {
            if name == "GP_POS_X" && formula == "12" {
                found_pos_x1 = true;
            } else if name == "GP_POS_Y" && formula == "34" {
                found_pos_y1 = true;
            } else if name == "GP_POS_X" && formula == "56" {
                found_pos_x2 = true;
            } else if name == "GP_POS_Y" && formula == "78" {
                found_pos_y2 = true;
            }
        }
        assert!(found_pos_x1);
        assert!(found_pos_y1);
        assert!(found_pos_x2);
        assert!(found_pos_y2);
    }

    #[test]
    fn test_xlsx_empty_table_preservation() {
        // Create an empty sheet (e.g., 5 columns, 10 rows, all empty strings)
        let mut columns = Vec::new();
        for _ in 0..5 {
            let mut col = DataColumn::new(10);
            col.src = vec![String::new(); 10].into();
            columns.push(col);
        }

        let sheet = Sheet {
            id: 2,
            name: "EmptyTable".to_string(),
            columns,
            pos: Position { left: 2, top: 3 },
            z_index: 0,
            dependencies: std::collections::HashMap::new(),
            dependencies_rev: std::collections::HashMap::new(),
            uncommitted_actions: Vec::new(),
            row_tops_cache: std::cell::RefCell::new(None),
        };

        // Export it
        let xlsx_data = export_xlsx_data(&[sheet], &[]).unwrap();
        assert!(!xlsx_data.is_empty());

        // Import it back
        let (imported_tables, _) = import_xlsx_data(&xlsx_data, 10, 20, &[], |_, _, _| {}).unwrap();
        assert_eq!(imported_tables.len(), 1);

        let imported_table = &imported_tables[0].sheet;
        assert!(imported_tables[0].pos_preserved);
        assert_eq!(imported_table.name, "EmptyTable");
        assert_eq!(imported_table.pos.left, 2);
        assert_eq!(imported_table.pos.top, 3);
        assert_eq!(imported_table.columns.len(), 5);
        assert_eq!(imported_table.row_count(), 10);
    }

    #[test]
    fn test_xlsx_chart_import_export_cycle() {
        // Create a sheet with data for the chart to reference
        let mut columns = Vec::new();
        let mut col1 = DataColumn::new(3);
        col1.name = "A".to_string();
        col1.src = vec!["Cat1".to_string(), "Cat2".to_string(), "Cat3".to_string()].into();
        columns.push(col1);

        let mut col2 = DataColumn::new(3);
        col2.name = "B".to_string();
        col2.src = vec!["10".to_string(), "20".to_string(), "30".to_string()].into();
        columns.push(col2);

        let sheet = Sheet {
            id: 1,
            name: "Sheet1".to_string(),
            columns,
            pos: Position { left: 0, top: 0 },
            z_index: 0,
            dependencies: std::collections::HashMap::new(),
            dependencies_rev: std::collections::HashMap::new(),
            uncommitted_actions: Vec::new(),
            row_tops_cache: std::cell::RefCell::new(None),
        };

        // Create a chart referencing that sheet
        let chart = crate::core::chart::Chart {
            id: 101,
            name: "Chart 1".to_string(),
            chart_type: crate::core::chart::ChartType::Line,
            data_range: "Sheet1!A1:B3".to_string(),
            from_anchor: crate::core::chart::ChartAnchor {
                col: 2,
                row: 4,
                col_offset_pixels: 10.0,
                row_offset_pixels: 5.0,
            },
            to_anchor: crate::core::chart::ChartAnchor {
                col: 7,
                row: 15,
                col_offset_pixels: 20.0,
                row_offset_pixels: 15.0,
            },
            title: Some("My Chart Title".to_string()),
            xlabel: Some("X Axis".to_string()),
            ylabel: Some("Y Axis".to_string()),
            show_legend: true,
            z_index: 0,
        };

        // Export sheet + chart
        let xlsx_data = export_xlsx_data(&[sheet], &[chart.clone()]).unwrap();
        assert!(!xlsx_data.is_empty());

        // Import back
        let (imported_tables, imported_charts) =
            import_xlsx_data(&xlsx_data, 10, 20, &[], |_, _, _| {}).unwrap();

        assert_eq!(imported_tables.len(), 1);
        assert_eq!(imported_charts.len(), 1);

        let imported_chart = &imported_charts[0];
        assert_eq!(imported_chart.title, Some("My Chart Title".to_string()));
        assert_eq!(imported_chart.xlabel, Some("X Axis".to_string()));
        assert_eq!(imported_chart.ylabel, Some("Y Axis".to_string()));
        assert_eq!(imported_chart.show_legend, true);
        assert_eq!(
            imported_chart.chart_type,
            crate::core::chart::ChartType::Line
        );

        // Verify the from_anchor coordinates round-trip
        assert_eq!(imported_chart.from_anchor.col, 2);
        assert_eq!(imported_chart.from_anchor.row, 4);

        // Offset conversion might have slight rounding differences, but should be close
        assert!((imported_chart.from_anchor.col_offset_pixels - 10.0).abs() < 1.0);
        assert!((imported_chart.from_anchor.row_offset_pixels - 5.0).abs() < 1.0);
    }
}
