use crate::core::{DataColumn, Sheet};
use calamine::Reader;
use web_time::Instant;

#[derive(Debug, serde::Serialize, serde::Deserialize)]
pub struct ImportedSheet {
    pub sheet: Sheet,
}

/// Quote a text cell's raw string when it would otherwise be re-parsed as a
/// number, boolean, or formula (see `Sheet::commit`'s literal-cell parsing),
/// so a text cell like "1" stays text instead of becoming the number 1.
fn text_cell_src(s: &str) -> String {
    if s.is_empty() {
        return String::new();
    }
    let looks_ambiguous = s.starts_with('=')
        || s.parse::<i64>().is_ok()
        || s.parse::<f64>().is_ok()
        || s.eq_ignore_ascii_case("true")
        || s.eq_ignore_ascii_case("false");
    if looks_ambiguous {
        format!("\"{}\"", s)
    } else {
        s.to_string()
    }
}

/// `(imported sheets, charts, pivot tables, VBA project)`.
type ImportedXlsxData = (
    Vec<ImportedSheet>,
    Vec<crate::core::chart::Chart>,
    Vec<crate::core::pivot::PivotTable>,
    Option<crate::core::vba::VbaProject>,
);

pub fn import_xlsx_data(
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
    let mut imported_tables = Vec::new();
    // Maps each worksheet's original name to the (possibly de-duplicated)
    // name it was actually imported under, so Excel Table definitions --
    // read separately below, keyed by original sheet name -- can be
    // attached to the right imported Sheet.
    let mut orig_to_assigned_name: std::collections::HashMap<String, String> =
        std::collections::HashMap::new();

    for (sheet_idx, orig_sheet_name) in sheet_names.iter().enumerate() {
        progress_callback(sheet_idx, total_sheets, orig_sheet_name);
        let start_sheet = Instant::now();
        let start_range = Instant::now();
        if let Ok(range) = workbook.worksheet_range(orig_sheet_name) {
            let elapsed_range = start_range.elapsed();

            let start_formula = Instant::now();
            let formula_range = workbook.worksheet_formula(orig_sheet_name).ok();
            let elapsed_formula = start_formula.elapsed();

            // calamine's `Range::rows()`/`Index` address cells relative to
            // the used range's own top-left corner, not the sheet's
            // absolute A1 origin -- a sheet whose leftmost/topmost
            // populated cell isn't row 1 / column A (e.g. a lone value at
            // B1) would otherwise get that data silently stored at
            // internal row/col 0, which every later A1 lookup elsewhere
            // (CLI, formulas) treats as meaning row 1 / column A, missing
            // the real column/row entirely. Sizing `rows`/`cols` off the
            // combined absolute `end()` of both ranges pads the sheet with
            // leading empty rows/columns instead, so internal index 0
            // always really is row 1 / column A. `Range::get_value` (used
            // in the slow path below) already takes absolute positions;
            // only `rows()`/`Index` are range-relative, which is why the
            // fast path further down only fires when `matches_exactly`
            // confirms there's no leading offset to account for.
            let ends = [range.end(), formula_range.as_ref().and_then(|fr| fr.end())];
            let (mut rows, mut cols) = ends
                .into_iter()
                .flatten()
                .fold((0usize, 0usize), |acc, (r, c)| {
                    (acc.0.max(r as usize + 1), acc.1.max(c as usize + 1))
                });

            // Fallback for empty worksheets
            if rows == 0 || cols == 0 {
                rows = 10;
                cols = 5;
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
                    .is_none_or(|fr| fr.height() == rows && fr.width() == cols);

            if matches_exactly {
                let mut range_rows = range.rows();
                let mut formula_rows = formula_range.as_ref().map(|fr| fr.rows());

                for row_idx in 0..rows {
                    let r_cells = range_rows.next().unwrap();
                    let f_cells = formula_rows.as_mut().map(|fr| fr.next().unwrap());

                    for col_idx in 0..cols {
                        let cell_value = &r_cells[col_idx];
                        let init_res = match cell_value {
                            calamine::Data::Int(i) => crate::core::engine::ResultData::Integer(*i),
                            calamine::Data::Float(f) => crate::core::engine::ResultData::Float(*f),
                            calamine::Data::String(s) => {
                                crate::core::engine::ResultData::String(s.clone())
                            }
                            calamine::Data::Bool(b) => crate::core::engine::ResultData::Boolean(*b),
                            calamine::Data::Error(e) => {
                                crate::core::engine::ResultData::Error(format!("{:?}", e))
                            }
                            _ => crate::core::engine::ResultData::None,
                        };
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
                                continue;
                            }
                        }

                        let cell_src = match cell_value {
                            calamine::Data::Empty => String::new(),
                            calamine::Data::String(s) => text_cell_src(s),
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

                        let cell_value = range
                            .get_value((row_u32, col_u32))
                            .unwrap_or(&calamine::Data::Empty);
                        let init_res = match cell_value {
                            calamine::Data::Int(i) => crate::core::engine::ResultData::Integer(*i),
                            calamine::Data::Float(f) => crate::core::engine::ResultData::Float(*f),
                            calamine::Data::String(s) => {
                                crate::core::engine::ResultData::String(s.clone())
                            }
                            calamine::Data::Bool(b) => crate::core::engine::ResultData::Boolean(*b),
                            calamine::Data::Error(e) => {
                                crate::core::engine::ResultData::Error(format!("{:?}", e))
                            }
                            _ => crate::core::engine::ResultData::None,
                        };
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
                            columns[col_idx].dirty_indices.push(row_idx);
                            continue;
                        }

                        let cell_src = match cell_value {
                            calamine::Data::Empty => String::new(),
                            calamine::Data::String(s) => text_cell_src(s),
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

            orig_to_assigned_name.insert(orig_sheet_name.clone(), sheet_name.clone());
            let new_table = Sheet {
                id: rand::random::<u64>(),
                name: sheet_name,
                columns,
                tables: Vec::new(),
                dependencies: std::collections::HashMap::new(),
                dependencies_rev: std::collections::HashMap::new(),
                uncommitted_actions: Vec::new(),
            };
            imported_tables.push(ImportedSheet { sheet: new_table });

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

    // Import Excel Table (ListObject) definitions, attaching each one to
    // its owning sheet. Name, sheet, and column names come straight from
    // calamine; the exact header/totals row presence is read directly from
    // the table's own XML since calamine doesn't expose those counts.
    if workbook.load_tables().is_ok() {
        let table_names: Vec<String> = workbook.table_names().into_iter().cloned().collect();
        for table_name in table_names {
            let Ok(table) = workbook.table_by_name(&table_name) else {
                continue;
            };
            let Some(assigned_sheet_name) = orig_to_assigned_name.get(table.sheet_name()) else {
                continue;
            };
            let Some(imported) = imported_tables
                .iter_mut()
                .find(|t: &&mut ImportedSheet| &t.sheet.name == assigned_sheet_name)
            else {
                continue;
            };
            let (has_header_row, has_totals_row, table_ref) =
                read_table_metadata_from_zip(buffer, &table_name);

            // Prefer the table's own declared `ref` (the whole table's
            // bounds, header/totals rows included) over reconstructing
            // position from calamine's `Table::data()` range, which has no
            // `start()`/`end()` at all for a table with zero data rows
            // (`Range::empty()`) -- falling back to the old data-range
            // computation only if the XML `ref` somehow isn't parseable.
            let (start_row, start_col, end_row, end_col) = if let Some(bounds) = table_ref {
                bounds
            } else {
                let Some((data_start_row, data_start_col)) = table.data().start() else {
                    continue;
                };
                let Some((data_end_row, data_end_col)) = table.data().end() else {
                    continue;
                };
                (
                    (data_start_row as usize).saturating_sub(usize::from(has_header_row)),
                    data_start_col as usize,
                    data_end_row as usize + usize::from(has_totals_row),
                    data_end_col as usize,
                )
            };

            // Guard against ranges that don't fit the imported sheet (e.g. a
            // table whose sheet fell back to the default empty-sheet size).
            if end_row >= imported.sheet.row_count() || end_col >= imported.sheet.col_count() {
                continue;
            }

            imported.sheet.tables.push(crate::core::table::ExcelTable {
                id: rand::random::<u64>(),
                name: table_name.clone(),
                sheet_id: imported.sheet.id,
                start_row,
                start_col,
                end_row,
                end_col,
                has_header_row,
                has_totals_row,
                columns: table.columns().to_vec(),
            });
        }
    }

    log::info!(
        "import_xlsx_data finished. Total time: {:.2?}",
        start_total.elapsed()
    );

    // Import charts from drawings and charts xml in ZIP. Chart ids are
    // derived deterministically from (sheet name, position within that
    // sheet's charts) rather than `rand::random()`: the CLI is a fresh
    // process per invocation, so a random id would change on every reload
    // of an unchanged file, making `chart edit --id`/`chart delete --id`
    // unable to find a chart a prior `chart list` just reported.
    let mut imported_charts = Vec::new();
    if let Ok(parsed_charts) = import_charts_from_zip(buffer) {
        let mut chart_index_by_sheet: std::collections::HashMap<String, usize> =
            std::collections::HashMap::new();
        for parsed in parsed_charts {
            if let Some(sheet) = imported_tables
                .iter()
                .find(|t| t.sheet.name == parsed.sheet_name)
            {
                let data_range = if parsed.info.data_range.is_empty() {
                    format!("{}!A1", sheet.sheet.name)
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

    let sheet_id_by_name: std::collections::HashMap<String, u64> = imported_tables
        .iter()
        .map(|t| (t.sheet.name.clone(), t.sheet.id))
        .collect();
    let imported_pivots = crate::core::pivot_xlsx::import_pivot_tables(
        buffer,
        &sheet_id_by_name,
        |sheet_name, r0, c0, r1, c1| {
            imported_tables
                .iter()
                .find(|t| t.sheet.name == sheet_name)
                .and_then(|t| {
                    t.sheet
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
        imported_tables,
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

pub fn export_xlsx_data(
    sheets: &[Sheet],
    charts: &[crate::core::chart::Chart],
    pivots: &[crate::core::pivot::PivotTable],
    vba: Option<&crate::core::vba::VbaProject>,
) -> Result<Vec<u8>, String> {
    let mut workbook = rust_xlsxwriter::Workbook::new();
    let mut used_names = std::collections::HashSet::new();
    let mut table_name_to_worksheet_name = std::collections::HashMap::new();
    let mut table_name_to_table = std::collections::HashMap::new();
    // Sheet id -> the actually-assigned worksheet name (after the 31-char
    // truncation and case-insensitive de-dup below), so
    // `vba_xlsx::export_vba_project` can match a Document module's bound
    // sheet against the same name this function writes into the sheet
    // element, rather than the original (possibly longer/colliding)
    // `sheet.name`.
    let mut sheet_id_to_worksheet_name = std::collections::HashMap::new();

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
        sheet_id_to_worksheet_name.insert(sheet.id, worksheet_name.clone());

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
                        worksheet
                            .write_formula(row_idx as u32, col_idx as u16, formula)
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
                        let text = if cell_src.starts_with('"')
                            && cell_src.ends_with('"')
                            && cell_src.len() >= 2
                        {
                            &cell_src[1..cell_src.len() - 1]
                        } else {
                            cell_src.as_str()
                        };
                        worksheet
                            .write_string(row_idx as u32, col_idx as u16, text)
                            .map_err(|e| format!("Failed to write Excel string: {}", e))?;
                    }
                }
            }

            // Export Excel Tables (ListObjects) defined on this sheet. Cell
            // values/formulas were already written above; a `TableColumn`
            // with no total/formula set leaves the totals row and data cells
            // alone, so `add_table` only adds table metadata (name, header
            // row, autofilter, style) without touching content we already
            // wrote -- except the header row, which it re-writes with the
            // same column name we already used to populate it.
            for table in &sheet.tables {
                let rx_columns: Vec<rust_xlsxwriter::TableColumn> = table
                    .columns
                    .iter()
                    .map(|name| rust_xlsxwriter::TableColumn::new().set_header(name))
                    .collect();
                let rx_table = rust_xlsxwriter::Table::new()
                    .set_name(&table.name)
                    .set_header_row(table.has_header_row)
                    .set_total_row(table.has_totals_row)
                    .set_columns(&rx_columns);
                worksheet
                    .add_table(
                        table.start_row as u32,
                        table.start_col as u16,
                        table.end_row as u32,
                        table.end_col as u16,
                        &rx_table,
                    )
                    .map_err(|e| format!("Failed to add Excel table '{}': {}", table.name, e))?;
            }
        }
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
            // `xlabel`/`ylabel` mean "category axis label"/"value axis
            // label" (matching parse_chart_xml's in_cat_ax/in_val_ax
            // reading on import), not "screen x/y axis". For horizontal
            // Bar charts, rust_xlsxwriter's x_axis()/y_axis() address
            // the *visual* bottom/left axes, which are the value/category
            // axes respectively -- the opposite of every other chart
            // type here, where the category axis is drawn horizontally.
            // Swap the setters here so xlabel always lands on <c:catAx>
            // and ylabel on <c:valAx>, keeping export consistent with
            // what import expects to read back.
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
            return String::from_utf8(attr.value.into_owned()).ok();
        }
    }
    None
}

/// Escapes text for use inside an XML attribute value. Shared by
/// `pivot_xlsx.rs` and `vba_xlsx.rs`, which used to each carry their own
/// (already-drifted: only this version escaped `'`) copy.
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

    // Scatter charts (`<c:scatterChart>`) have no `<c:catAx>` at all -- both
    // of their axes are `<c:valAx>` (there's no category axis for an XY
    // scatter plot), so `in_cat_ax`/`in_val_ax` alone can't tell the first
    // (x) axis's title from the second (y) axis's title the way they can
    // for every other chart type, which has exactly one axis of each kind.
    // Titles found while `in_val_ax` are collected here in document order
    // and only classified into xlabel/ylabel once parsing finishes and it's
    // known whether a `<c:catAx>` ever appeared at all.
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
                    // Scatter charts (`<c:scatterChart>`) encode their two
                    // series ranges as `<c:xVal>`/`<c:yVal>` instead of the
                    // `<c:cat>`/`<c:val>` every other chart type uses --
                    // without this, a Scatter chart's range silently
                    // resolves to empty on every reimport (falls back to
                    // "Sheet!A1" in the caller), corrupting it on the very
                    // next save since the CLI is a fresh process per
                    // invocation.
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
                if let Ok(text) = e.unescape() {
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

    // Classify the collected `<c:valAx>` titles now that it's known whether
    // a `<c:catAx>` ever appeared: with one, this is a normal chart with a
    // single value axis (its title is the y-axis label); without one, it's
    // a Scatter chart whose first/second `<c:valAx>` are the x/y axes
    // respectively, matching the order rust_xlsxwriter's x_axis()/y_axis()
    // write them in on export.
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

/// Reads `headerRowCount`/`totalsRowCount` directly from a table's own XML
/// part (`xl/tables/table*.xml`), matched by its `displayName`. calamine's
/// `Table::data()` already excludes the header row (and, in the common
/// case, the totals row) but doesn't expose these counts directly. Defaults
/// to `(true, false)` -- a header row and no totals row -- if the table's
/// XML can't be found or parsed, matching how most real-world tables are
/// configured.
/// Reads a table's header/totals row flags (calamine doesn't expose these)
/// directly from its own `xl/tables/table*.xml`, along with its declared
/// `ref` range parsed to 0-based `(start_row, start_col, end_row, end_col)`
/// -- the whole table's bounds, header/totals rows included. The `ref` is
/// the authoritative source of truth for a table's position: prefer it over
/// reconstructing bounds from calamine's `Table::data()` range, which is
/// empty (no `start()`/`end()`) for a table with zero data rows.
/// `(has_header_row, has_totals_row, bounds)` where `bounds` is 0-based
/// `(start_row, start_col, end_row, end_col)`.
type TableMetadata = (bool, bool, Option<(usize, usize, usize, usize)>);

fn read_table_metadata_from_zip(buffer: &[u8], table_name: &str) -> TableMetadata {
    let Ok(mut archive) = zip::ZipArchive::new(std::io::Cursor::new(buffer)) else {
        return (true, false, None);
    };
    let table_files: Vec<String> = archive
        .file_names()
        .filter(|n| n.starts_with("xl/tables/") && n.ends_with(".xml"))
        .map(|s| s.to_string())
        .collect();

    for file_name in table_files {
        let Some(xml_content) = get_zip_file_content(&mut archive, &file_name) else {
            continue;
        };
        let mut reader = quick_xml::Reader::from_str(&xml_content);
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(quick_xml::events::Event::Start(ref e))
                | Ok(quick_xml::events::Event::Empty(ref e))
                    if e.local_name().as_ref() == b"table" =>
                {
                    let display_name = get_attr(e, b"displayName").unwrap_or_default();
                    if display_name == table_name {
                        let header_row_count = get_attr(e, b"headerRowCount")
                            .and_then(|s| s.parse::<usize>().ok())
                            .unwrap_or(1);
                        let totals_row_count = get_attr(e, b"totalsRowCount")
                            .and_then(|s| s.parse::<usize>().ok())
                            .unwrap_or(0);
                        let table_ref = get_attr(e, b"ref")
                            .and_then(|s| crate::core::pivot_xlsx::parse_a1_range(&s));
                        return (header_row_count != 0, totals_row_count != 0, table_ref);
                    }
                }
                Ok(quick_xml::events::Event::Eof) | Err(_) => break,
                _ => {}
            }
            buf.clear();
        }
    }
    (true, false, None)
}

/// Derives a stable chart id from its sheet name and position within that
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
    // Cap to JS Number.MAX_SAFE_INTEGER (2^53 - 1), matching
    // generate_unique_id's convention, to avoid serialization precision loss.
    hasher.finish() & 0x001F_FFFF_FFFF_FFFF
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
            tables: Vec::new(),
            dependencies: std::collections::HashMap::new(),
            dependencies_rev: std::collections::HashMap::new(),
            uncommitted_actions: Vec::new(),
        };

        // Export it
        let xlsx_data = export_xlsx_data(&[sheet], &[], &[], None).unwrap();
        assert!(!xlsx_data.is_empty());

        // Import it back
        let (imported_tables, imported_charts, _, _) =
            import_xlsx_data(&xlsx_data, &[], |_, _, _| {}).unwrap();
        assert_eq!(imported_tables.len(), 1);
        assert_eq!(imported_charts.len(), 0);

        let imported_table = &imported_tables[0].sheet;
        assert_eq!(imported_table.name, "Sheet1");
        assert_eq!(imported_table.columns.len(), 2);

        // Verify content
        assert_eq!(imported_table.columns[0].src[0], "10");
        assert_eq!(imported_table.columns[0].src[1], "20");
        assert_eq!(imported_table.columns[1].src[0], "=A1 + A2");
        assert_eq!(imported_table.columns[1].src[1], "abc");
    }

    #[test]
    fn test_xlsx_import_data_offset_from_column_a_is_not_lost() {
        // Column A is entirely blank; real data starts at column B, with a
        // formula in column C referencing it. Regression test for a bug
        // where calamine's used-range addressing (relative to the range's
        // own top-left corner, not the sheet's absolute A1 origin) caused
        // column B's data to be silently stored at internal column 0 and
        // labeled "A" on import, making it unreachable by its true B1-style
        // reference and vanishing from column A entirely.
        let mut col_a = DataColumn::new(2);
        col_a.name = "A".to_string();
        col_a.src = vec![String::new(), String::new()].into();

        let mut col_b = DataColumn::new(2);
        col_b.name = "B".to_string();
        col_b.src = vec!["42".to_string(), "8".to_string()].into();

        let mut col_c = DataColumn::new(2);
        col_c.name = "C".to_string();
        col_c.src = vec!["=B1 + B2".to_string(), String::new()].into();

        let sheet = Sheet {
            id: 1,
            name: "Sheet1".to_string(),
            columns: vec![col_a, col_b, col_c],
            tables: Vec::new(),
            dependencies: std::collections::HashMap::new(),
            dependencies_rev: std::collections::HashMap::new(),
            uncommitted_actions: Vec::new(),
        };

        let xlsx_data = export_xlsx_data(&[sheet], &[], &[], None).unwrap();

        let (imported_tables, _, _, _) = import_xlsx_data(&xlsx_data, &[], |_, _, _| {}).unwrap();
        assert_eq!(imported_tables.len(), 1);
        let imported = &imported_tables[0].sheet;

        // Column A must stay empty and column B must keep its own data --
        // not have it aliased into column A -- and the formula must still
        // read "=B1 + B2" (not silently rewritten to "=A1 + A2").
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
        let (imported_tables, _, _, _) = import_xlsx_data(&xlsx_data, &[], |_, _, _| {}).unwrap();
        assert_eq!(imported_tables.len(), 1);

        let imported_sheet = &imported_tables[0].sheet;
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

        // The structured reference over the re-imported table should still
        // only see the table's own data rows.
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
        // Regression test for a real crash found via libvisi's pivot-table
        // fuzz testing (see vendor/calamine/PATCHES.md): exporting an Excel
        // Table with a header row but zero data rows, then reimporting it,
        // used to panic inside calamine's `Xlsx::table_by_name` ("invalid
        // range bounds") -- unrelated to pivot tables specifically, a plain
        // header-only table triggers it on its own.
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
        let (imported_tables, _, _, _) = import_xlsx_data(&xlsx_data, &[], |_, _, _| {}).unwrap();
        assert_eq!(imported_tables.len(), 1);

        let imported_sheet = &imported_tables[0].sheet;
        assert_eq!(imported_sheet.tables.len(), 1);
        let table = &imported_sheet.tables[0];
        assert_eq!(table.name, "Empty");
        assert_eq!(table.columns, vec!["Name", "Amount"]);
        assert!(table.has_header_row);
        assert!(!table.has_totals_row);
        // Zero data rows: the data range is empty (start past end).
        assert!(table.data_start_row() > table.data_end_row());
    }

    #[test]
    fn test_xlsx_table_import_with_absolute_relationship_target() {
        // Regression test for GitHub issue #16: openpyxl (and other OPC-
        // compliant writers) may emit worksheet->table relationship
        // `Target` attributes as absolute package paths
        // ("/xl/tables/table1.xml") rather than the "../tables/table1.xml"
        // relative form Excel and rust_xlsxwriter always use. Unpatched
        // calamine only special-cases the relative form, so an absolute
        // target silently resolved to zero tables (see
        // vendor/calamine/PATCHES.md). Simulate an openpyxl-authored file
        // by rewriting a normally-exported file's sheet rels to use an
        // absolute target, then confirm the table still imports.
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

        // Rewrite every worksheet rels part's relative table target
        // ("../tables/table1.xml") to the absolute package-path form
        // openpyxl uses ("/xl/tables/table1.xml"), leaving everything else
        // byte-for-byte identical.
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

        let (imported_tables, _, _, _) = import_xlsx_data(&rewritten, &[], |_, _, _| {}).unwrap();
        assert_eq!(imported_tables.len(), 1);
        let imported_sheet = &imported_tables[0].sheet;
        assert_eq!(imported_sheet.tables.len(), 1);
        assert_eq!(imported_sheet.tables[0].name, "Sales");
        assert_eq!(imported_sheet.tables[0].columns, vec!["Name", "Amount"]);
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
            tables: Vec::new(),
            dependencies: std::collections::HashMap::new(),
            dependencies_rev: std::collections::HashMap::new(),
            uncommitted_actions: Vec::new(),
        };

        // Export it
        let xlsx_data = export_xlsx_data(&[sheet], &[], &[], None).unwrap();
        assert!(!xlsx_data.is_empty());

        // Import it back
        let (imported_tables, _, _, _) = import_xlsx_data(&xlsx_data, &[], |_, _, _| {}).unwrap();
        assert_eq!(imported_tables.len(), 1);

        let imported_table = &imported_tables[0].sheet;
        assert_eq!(imported_table.name, "EmptyTable");
        assert_eq!(imported_table.columns.len(), 5);
        assert_eq!(imported_table.row_count(), 10);
    }

    #[test]
    fn test_xlsx_chart_import_export_cycle() {
        // Covers Bar and Column specifically (not just Line) because
        // `parse_chart_xml`'s bar-vs-column disambiguation reads
        // `<c:barDir val="col"|"bar">`, a branch the original single-type
        // (Line) version of this test never exercised.
        for chart_type in [
            crate::core::chart::ChartType::Line,
            crate::core::chart::ChartType::Bar,
            crate::core::chart::ChartType::Column,
            crate::core::chart::ChartType::Scatter,
        ] {
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
                tables: Vec::new(),
                dependencies: std::collections::HashMap::new(),
                dependencies_rev: std::collections::HashMap::new(),
                uncommitted_actions: Vec::new(),
            };

            // Create a chart referencing that sheet
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

            // Export sheet + chart
            let xlsx_data =
                export_xlsx_data(&[sheet], std::slice::from_ref(&chart), &[], None).unwrap();
            assert!(!xlsx_data.is_empty());

            // Import back
            let (imported_tables, imported_charts, _, _) =
                import_xlsx_data(&xlsx_data, &[], |_, _, _| {}).unwrap();

            assert_eq!(imported_tables.len(), 1);
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
        let mut col1 = DataColumn::new(3);
        col1.name = "A".to_string();
        col1.src = vec!["10".to_string(), "20".to_string(), "=A1 + A2".to_string()].into();
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
            tables: Vec::new(),
            dependencies: std::collections::HashMap::new(),
            dependencies_rev: std::collections::HashMap::new(),
            uncommitted_actions: Vec::new(),
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
    fn test_xlsx_numeric_looking_text_cell_preserves_type() {
        // A cell whose text content looks like a number (e.g. imported from a
        // spreadsheet where the user typed "1" into a text-formatted cell)
        // must round-trip as text, not silently become the number 1.
        let mut columns = Vec::new();
        let mut col1 = DataColumn::new(1);
        col1.name = "A".to_string();
        col1.src = vec!["\"1\"".to_string()].into();
        columns.push(col1);

        let sheet = Sheet {
            id: 1,
            name: "Sheet1".to_string(),
            columns,
            tables: Vec::new(),
            dependencies: std::collections::HashMap::new(),
            dependencies_rev: std::collections::HashMap::new(),
            uncommitted_actions: Vec::new(),
        };

        let xlsx_data = export_xlsx_data(&[sheet], &[], &[], None).unwrap();

        // The exported xlsx should contain the plain text "1", not a literal
        // quote-wrapped string.
        let cursor = std::io::Cursor::new(xlsx_data.clone());
        let mut zip = zip::ZipArchive::new(cursor).unwrap();
        let mut sheet_file = zip.by_name("xl/worksheets/sheet1.xml").unwrap();
        let mut xml_content = String::new();
        std::io::Read::read_to_string(&mut sheet_file, &mut xml_content).unwrap();
        assert!(!xml_content.contains("&quot;1&quot;"));

        let (imported_tables, _, _, _) = import_xlsx_data(&xlsx_data, &[], |_, _, _| {}).unwrap();
        let mut imported_sheet = imported_tables.into_iter().next().unwrap().sheet;
        imported_sheet.columns[0].mark_dirty(0);
        imported_sheet.commit(None).unwrap();

        match imported_sheet.columns[0].data.get(0) {
            Some(crate::core::engine::ResultData::String(s)) => assert_eq!(s, "1"),
            other => panic!("Expected ResultData::String(\"1\"), got {:?}", other),
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

        // The exported file must contain real pivot table XML parts.
        let cursor = std::io::Cursor::new(xlsx_data.clone());
        let mut zip = zip::ZipArchive::new(cursor).unwrap();
        assert!(zip.by_name("xl/pivotTables/pivotTable1.xml").is_ok());
        assert!(
            zip.by_name("xl/pivotCache/pivotCacheDefinition1.xml")
                .is_ok()
        );
        assert!(zip.by_name("xl/pivotCache/pivotCacheRecords1.xml").is_ok());

        let (imported_tables, _, imported_pivots, _) =
            import_xlsx_data(&xlsx_data, &[], |_, _, _| {}).unwrap();
        assert_eq!(imported_tables.len(), 1);
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

        // Recomputing from the reimported definition should reproduce the
        // same aggregation (East=15, West=70, Grand Total=85).
        let reimported_sheet = &imported_tables[0].sheet;
        let grid = crate::core::pivot::compute_pivot(&[reimported_sheet], reimported).unwrap();
        assert_eq!(grid.body_rows.len(), 3);
    }

    #[test]
    fn test_xlsx_pivot_outer_field_subtotal_off_survives_round_trip() {
        // Regression test: `import_pivot_tables` used to always construct
        // row/col `PivotField`s via `PivotField::new`, whose `subtotal`
        // defaults to `true`, ignoring the `<item t="default"/>` marker (or
        // lack thereof) `build_pivot_xml_unit` had already written into
        // each `<pivotField>`'s `<items>`. Since every `visi pivot`
        // CLI command round-trips the whole workbook through xlsx (a fresh
        // process per invocation), a `subtotal: false` set at pivot-build
        // time was silently discarded before the next command ran --
        // discovered via fuzz/fuzz_pivot.py, whose Python driver renders
        // Excel's *own* PivotTable via VBA (which does honor the toggle)
        // while `visi` always showed the subtotal row regardless.
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

        // A single row field ("Region") with its subtotal explicitly turned
        // off -- i.e. it's currently the *innermost* (only) field of its
        // axis. This is exactly the state a field is in right after `visi
        // pivot add-field --area row --column Region --no-subtotal` and
        // before any second row field is ever added: the earlier version of
        // this fix suppressed the `<item t="default"/>` marker for
        // whichever field is innermost *at export time*, on the theory that
        // a subtotal has nothing to summarize there anyway -- but that
        // meant a field's `false` was only ever durable once it stopped
        // being innermost, which never survives the CLI's per-invocation
        // xlsx round-trip: each `add-field` call re-imports the file before
        // making its change, so a solo field's `false` was already lost by
        // the time a second field's `add-field` call could read it back.
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
        // Regression test: patch_workbook_code_names used to match a bound
        // Document module's sheet by its original `Sheet::name`, but a
        // worksheet name over 31 chars gets truncated on export -- so the
        // needle it searched for never matched, and the codeName (and thus
        // the module's sheet binding) was silently dropped from the saved
        // file.
        let long_name = "A".repeat(40);
        let sheet = Sheet {
            id: 1,
            name: long_name.clone(),
            columns: Vec::new(),
            tables: Vec::new(),
            dependencies: std::collections::HashMap::new(),
            dependencies_rev: std::collections::HashMap::new(),
            uncommitted_actions: Vec::new(),
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
