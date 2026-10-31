mod chart;
mod formula;
mod function_registry;
mod package;
mod pivot;
mod record;
mod shared_strings;
mod styles;
mod table;
mod vba;
mod workbook;
mod worksheet;

use calamine::{Reader, SheetType, SheetVisible};

use crate::core::engine::{CellType, DataColumn, ResultData, Sheet};
use crate::core::excel::{ImportedSheet, ImportedWorkbookData};

fn result_data(value: &calamine::Data) -> ResultData {
    match value {
        calamine::Data::Int(value) => ResultData::Integer(*value),
        calamine::Data::Float(value) => ResultData::Float(*value),
        calamine::Data::String(value) => ResultData::String(value.clone()),
        calamine::Data::Bool(value) => ResultData::Boolean(*value),
        calamine::Data::Error(value) => ResultData::Error(value.to_string()),
        calamine::Data::DateTime(value) => ResultData::Float(value.as_f64()),
        calamine::Data::DateTimeIso(value) | calamine::Data::DurationIso(value) => {
            ResultData::String(value.clone())
        }
        calamine::Data::Empty => ResultData::None,
    }
}

fn cell_type_and_source(value: &calamine::Data) -> (CellType, String) {
    match value {
        calamine::Data::Empty => (CellType::Empty, String::new()),
        calamine::Data::Int(value) => (CellType::Int, value.to_string()),
        calamine::Data::Float(value) => (CellType::Float, value.to_string()),
        calamine::Data::String(value) => (CellType::String, value.clone()),
        calamine::Data::Bool(value) => (CellType::Bool, value.to_string()),
        calamine::Data::DateTime(value) => (CellType::DateTime, value.to_string()),
        calamine::Data::DateTimeIso(value) => (CellType::DateTimeIso, value.clone()),
        calamine::Data::DurationIso(value) => (CellType::DurationIso, value.clone()),
        calamine::Data::Error(value) => (CellType::Error, value.to_string()),
    }
}

#[doc = "Imports worksheet values and formulas from an XLSB package."]
pub fn import_xlsb_data(
    buffer: &[u8],
    existing_sheets: &[Sheet],
    progress_callback: impl FnMut(usize, usize, &str),
) -> crate::Result<ImportedWorkbookData> {
    match std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        import_xlsb_data_inner(buffer, existing_sheets, progress_callback)
    })) {
        Ok(result) => result,
        Err(payload) => {
            let message = payload
                .downcast_ref::<&str>()
                .copied()
                .or_else(|| payload.downcast_ref::<String>().map(String::as_str))
                .unwrap_or("unknown XLSB parser panic");
            Err(crate::Error::Xlsb(format!(
                "failed to parse malformed XLSB workbook: {message}"
            )))
        }
    }
}

fn import_xlsb_data_inner(
    buffer: &[u8],
    existing_sheets: &[Sheet],
    mut progress_callback: impl FnMut(usize, usize, &str),
) -> crate::Result<ImportedWorkbookData> {
    package::validate_import(buffer).map_err(crate::Error::Xlsb)?;
    let mut layouts = package::read_layouts(buffer).map_err(crate::Error::Xlsb)?;
    let normalized = package::normalize_for_calamine(buffer).map_err(crate::Error::Xlsb)?;
    let cursor = std::io::Cursor::new(normalized);
    let mut workbook: calamine::Xlsb<_> = calamine::open_workbook_from_rs(cursor)
        .map_err(|error| crate::Error::Xlsb(format!("failed to open XLSB workbook: {error}")))?;
    let unsupported_names = workbook
        .defined_names()
        .iter()
        .map(|(name, _)| name.as_str())
        .filter(|name| !name.eq_ignore_ascii_case("_xlfn.SINGLE"))
        .collect::<Vec<_>>();
    if !unsupported_names.is_empty() {
        return Err(crate::Error::Xlsb(format!(
            "defined names are not supported in XLSB input: {}",
            unsupported_names.join(", ")
        )));
    }
    for metadata in workbook.sheets_metadata() {
        if metadata.typ != SheetType::WorkSheet {
            return Err(crate::Error::Xlsb(format!(
                "unsupported XLSB sheet type for '{}'",
                metadata.name
            )));
        }
        if metadata.visible != SheetVisible::Visible {
            return Err(crate::Error::Xlsb(format!(
                "hidden XLSB worksheets are not supported: '{}'",
                metadata.name
            )));
        }
    }
    let sheet_names = workbook.sheet_names();
    let total_sheets = sheet_names.len();
    let mut imported_sheets = Vec::with_capacity(total_sheets);
    for (sheet_index, original_name) in sheet_names.iter().enumerate() {
        progress_callback(sheet_index, total_sheets, original_name);
        let range = workbook.worksheet_range(original_name).map_err(|error| {
            crate::Error::Xlsb(format!(
                "failed to read worksheet '{original_name}': {error}"
            ))
        })?;
        let formula_range = workbook.worksheet_formula(original_name).ok();
        let ends = [
            range.end(),
            formula_range.as_ref().and_then(|value| value.end()),
        ];
        let (mut rows, mut cols) = ends
            .into_iter()
            .flatten()
            .fold((0usize, 0usize), |(rows, cols), (row, col)| {
                (rows.max(row as usize + 1), cols.max(col as usize + 1))
            });
        let layout = layouts.remove(original_name).unwrap_or_default();
        for (_, last, _) in &layout.column_widths {
            cols = cols.max(last + 1);
        }
        for (row, _) in &layout.row_heights {
            rows = rows.max(row + 1);
        }
        if rows == 0 || cols == 0 {
            rows = 10;
            cols = 5;
        }
        let mut name = original_name.clone();
        let mut suffix = 1;
        while existing_sheets.iter().any(|sheet| sheet.name == name)
            || imported_sheets
                .iter()
                .any(|sheet: &ImportedSheet| sheet.sheet.name == name)
        {
            name = format!("{original_name}_{suffix}");
            suffix += 1;
        }
        let mut columns = (0..cols)
            .map(|col| {
                let mut column = DataColumn::new(rows);
                column.name = crate::core::parser::col_idx_to_letters(col);
                column
            })
            .collect::<Vec<_>>();
        for (first, last, width) in layout.column_widths {
            for col in first..=last {
                if let Some(column) = columns.get_mut(col) {
                    column.width = Some(width);
                }
            }
        }
        let mut row_heights = vec![None; rows];
        for (row, height) in layout.row_heights {
            if let Some(slot) = row_heights.get_mut(row) {
                *slot = Some(height);
            }
        }
        for row in 0..rows {
            for (col, column) in columns.iter_mut().enumerate() {
                let coordinate = (row as u32, col as u32);
                let value = range
                    .get_value(coordinate)
                    .unwrap_or(&calamine::Data::Empty);
                column.data.set(row, result_data(value));
                if let Some(formula) = formula_range
                    .as_ref()
                    .and_then(|formulas| formulas.get_value(coordinate))
                    .filter(|formula| !formula.is_empty())
                {
                    if matches!(value, calamine::Data::Empty)
                        || matches!(value, calamine::Data::String(text) if text.is_empty())
                    {
                        column.data.set(row, ResultData::None);
                    }
                    column.src[row] = if formula.starts_with('=') {
                        formula.clone()
                    } else {
                        format!("={formula}")
                    };
                    column.cell_types[row] = cell_type_and_source(value).0;
                    column.dirty_indices.push(row);
                } else {
                    let (cell_type, source) = cell_type_and_source(value);
                    column.src[row] = source;
                    column.cell_types[row] = cell_type;
                }
            }
        }
        imported_sheets.push(ImportedSheet {
            sheet: Sheet {
                id: rand::random::<u64>(),
                name,
                columns,
                row_heights,
                tables: Vec::new(),
                dependencies: std::collections::HashMap::new(),
                dependencies_rev: std::collections::HashMap::new(),
                uncommitted_actions: Vec::new(),
                locale: crate::core::Locale::default(),
            },
        });
    }
    if imported_sheets.is_empty() {
        return Err(crate::Error::Xlsb(
            "no worksheets found in the XLSB workbook".to_string(),
        ));
    }
    Ok((imported_sheets, Vec::new(), Vec::new(), None))
}

#[doc = "Exports workbook sheets as an XLSB package."]
pub fn export_xlsb_data(sheets: &[Sheet]) -> crate::Result<Vec<u8>> {
    package::write(sheets).map_err(crate::Error::Xlsb)
}

#[cfg(test)]
mod tests {
    use crate::core::{CellRef, WorkbookManager};

    #[test]
    fn scalar_workbook_reopens_with_calamine() {
        let mut workbook = WorkbookManager::new_empty().unwrap();
        workbook.set_cell(0, 0, 0, "42".to_string());
        workbook.set_cell(0, 1, 0, "hello".to_string());
        workbook.set_cell(0, 2, 0, "TRUE".to_string());
        workbook.evaluate().unwrap();
        let bytes = workbook
            .save_bytes_as(crate::core::WorkbookFormat::Xlsb)
            .unwrap();
        let loaded = WorkbookManager::load_bytes(&bytes).unwrap();
        assert_eq!(loaded.sheets[0].get_src_str(&CellRef::new(0, 0)), "42");
        assert_eq!(loaded.sheets[0].get_src_str(&CellRef::new(1, 0)), "hello");
        assert_eq!(loaded.sheets[0].get_src_str(&CellRef::new(2, 0)), "true");
    }

    #[test]
    fn layout_round_trip_preserves_column_widths_and_row_heights() {
        let mut workbook = WorkbookManager::new_empty().unwrap();
        workbook.sheets[0].columns[0].width = Some(20.0);
        workbook.sheets[0].row_heights[0] = Some(24.0);
        let bytes = workbook
            .save_bytes_as(crate::core::WorkbookFormat::Xlsb)
            .unwrap();
        let loaded = WorkbookManager::load_bytes(&bytes).unwrap();
        assert_eq!(loaded.sheets[0].columns[0].width, Some(20.0));
        assert_eq!(loaded.sheets[0].row_heights[0], Some(24.0));
    }

    #[test]
    fn formula_workbook_preserves_formula_and_cache() {
        let mut workbook = WorkbookManager::new_empty().unwrap();
        workbook.set_cell(0, 0, 0, "2".to_string());
        workbook.set_cell(0, 0, 1, "=A1+3".to_string());
        workbook.evaluate().unwrap();
        let bytes = workbook
            .save_bytes_as(crate::core::WorkbookFormat::Xlsb)
            .unwrap();
        let loaded = WorkbookManager::load_bytes(&bytes).unwrap();
        assert_eq!(loaded.sheets[0].get_src_str(&CellRef::new(0, 1)), "=$A$1+3");
        assert_eq!(
            loaded.sheets[0].columns[1].value(0).unwrap().to_string(),
            "5"
        );
    }
}
