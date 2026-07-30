use crate::utils::col_idx_to_letters;
use libvisi::core::{
    ExcelTable,
    chart::{Chart, ChartType},
    engine::{Context, DataColumn, ResultData, Sheet, generate_unique_id},
};
use libvisi::{export_xlsx_data, import_xlsx_data};
use std::fs;
use std::io::{self, Read, Write};
use std::path::Path;

pub struct SheetSummary {
    pub name: String,
    pub row_count: usize,
    pub col_count: usize,
    pub formula_count: usize,
}

pub struct WorkbookSummary {
    pub file_name: String,
    pub sheet_count: usize,
    pub chart_count: usize,
    pub sheets: Vec<SheetSummary>,
}

pub struct WorkbookManager {
    pub sheets: Vec<Sheet>,
    pub charts: Vec<Chart>,
}

impl WorkbookManager {
    /// Load Excel workbook from bytes buffer
    pub fn load_bytes(buffer: &[u8]) -> Result<Self, String> {
        let (imported_tables, charts) =
            import_xlsx_data(buffer, &[], |_, _, _| {}).map_err(|e| e.to_string())?;

        let sheets = imported_tables.into_iter().map(|it| it.sheet).collect();
        Ok(Self { sheets, charts })
    }

    /// Load Excel workbook from file path or stdin ("-")
    pub fn load_file(path_str: &str) -> Result<Self, String> {
        let buffer = if path_str == "-" {
            let mut stdin_bytes = Vec::new();
            io::stdin()
                .read_to_end(&mut stdin_bytes)
                .map_err(|e| format!("Failed to read stdin: {}", e))?;
            stdin_bytes
        } else {
            fs::read(path_str).map_err(|e| format!("Failed to read file '{}': {}", path_str, e))?
        };

        Self::load_bytes(&buffer)
    }

    /// Load Excel workbook from file path, or create a new empty workbook if file does not exist
    pub fn load_file_or_create(path_str: &str) -> Result<Self, String> {
        if path_str != "-" && !Path::new(path_str).exists() {
            let mut wb = Self {
                sheets: Vec::new(),
                charts: Vec::new(),
            };
            wb.add_sheet("Sheet1")?;
            Ok(wb)
        } else {
            Self::load_file(path_str)
        }
    }

    /// Save Excel workbook to file path or stdout ("-")
    pub fn save_file(&self, path_str: &str) -> Result<(), String> {
        let bytes = export_xlsx_data(&self.sheets, &self.charts).map_err(|e| e.to_string())?;

        if path_str == "-" {
            io::stdout()
                .write_all(&bytes)
                .map_err(|e| format!("Failed to write to stdout: {}", e))?;
            io::stdout()
                .flush()
                .map_err(|e| format!("Failed to flush stdout: {}", e))?;
        } else {
            let path = Path::new(path_str);
            if let Some(parent) = path.parent() {
                if !parent.as_os_str().is_empty() {
                    fs::create_dir_all(parent)
                        .map_err(|e| format!("Failed to create parent directories: {}", e))?;
                }
            }
            fs::write(path_str, bytes)
                .map_err(|e| format!("Failed to save file to '{}': {}", path_str, e))?;
        }
        Ok(())
    }

    /// Recalculate all formulas in all sheets using libvisi engine
    pub fn evaluate(&mut self) -> Result<(), String> {
        if self.sheets.is_empty() {
            return Ok(());
        }

        // 1. Mark all cells dirty across all sheets
        for sheet in &mut self.sheets {
            sheet.mark_all_dirty();
        }

        // 2. Multi-pass evaluation to resolve cross-sheet formula dependencies
        for _pass in 0..3 {
            for i in 0..self.sheets.len() {
                let (left, right) = self.sheets.split_at_mut(i);
                let (target_sheet, right_tail) = right.split_first_mut().unwrap();

                let mut context = Context::new();
                for s in left.iter() {
                    context.add_table(s.name.clone(), s);
                }
                for s in right_tail.iter() {
                    context.add_table(s.name.clone(), s);
                }

                let _ = target_sheet.commit(Some(&context));
            }
        }

        Ok(())
    }

    /// Find index of sheet by name, or return default index 0 if name is None.
    pub fn find_sheet_index(&self, name_opt: Option<&str>) -> Result<usize, String> {
        if self.sheets.is_empty() {
            return Err("Workbook contains no sheets".to_string());
        }

        match name_opt {
            Some(name) => {
                if let Some(idx) = self
                    .sheets
                    .iter()
                    .position(|s| s.name.eq_ignore_ascii_case(name))
                {
                    Ok(idx)
                } else {
                    let available: Vec<String> =
                        self.sheets.iter().map(|s| s.name.clone()).collect();
                    Err(format!(
                        "Sheet '{}' not found. Available sheets: {}",
                        name,
                        available.join(", ")
                    ))
                }
            }
            None => Ok(0),
        }
    }

    /// Get structural summary of workbook
    pub fn get_summary(&self, file_name: &str) -> WorkbookSummary {
        let sheet_summaries = self
            .sheets
            .iter()
            .map(|sheet| {
                let row_count = sheet.row_count();
                let col_count = sheet.col_count();
                let mut formula_count = 0;

                for col in &sheet.columns {
                    for src in &col.src {
                        if src.starts_with('=') {
                            formula_count += 1;
                        }
                    }
                }

                SheetSummary {
                    name: sheet.name.clone(),
                    row_count,
                    col_count,
                    formula_count,
                }
            })
            .collect();

        WorkbookSummary {
            file_name: file_name.to_string(),
            sheet_count: self.sheets.len(),
            chart_count: self.charts.len(),
            sheets: sheet_summaries,
        }
    }

    /// Ensure sheet bounds can accommodate specified target_row and target_col
    pub fn ensure_capacity(&mut self, sheet_idx: usize, target_row: usize, target_col: usize) {
        if sheet_idx >= self.sheets.len() {
            return;
        }

        let sheet = &mut self.sheets[sheet_idx];
        let current_rows = sheet.row_count();
        let current_cols = sheet.col_count();

        // Add missing columns if target_col >= current_cols
        if target_col >= current_cols {
            let rows_for_new_cols = current_rows.max(target_row + 1).max(1);
            for col_i in current_cols..=target_col {
                let mut new_col = DataColumn::new(rows_for_new_cols);
                new_col.id = generate_unique_id();
                new_col.name = col_idx_to_letters(col_i);
                sheet.columns.push(new_col);
            }
        }

        // Expand all columns if target_row >= current_rows
        let current_rows = sheet.row_count();
        if target_row >= current_rows {
            let needed_rows = target_row + 1;
            for col in &mut sheet.columns {
                while col.src.len() < needed_rows {
                    col.src.push(String::new());
                    col.compiled_src
                        .push(libvisi::core::CompiledFormula::default());
                    col.data.push(ResultData::None);
                }
            }
        }
    }

    /// Update cell source / value at (row, col)
    pub fn set_cell(&mut self, sheet_idx: usize, row: usize, col: usize, value: String) {
        self.ensure_capacity(sheet_idx, row, col);
        let sheet = &mut self.sheets[sheet_idx];
        sheet.set_cell_src(row, col, value);
    }

    /// Insert row at 0-based index
    pub fn insert_row(&mut self, sheet_idx: usize, row_idx: usize) -> Result<(), String> {
        let sheet = &mut self.sheets[sheet_idx];
        sheet.insert_row(row_idx);
        Ok(())
    }

    /// Delete row at 0-based index
    pub fn delete_row(&mut self, sheet_idx: usize, row_idx: usize) -> Result<(), String> {
        let sheet = &mut self.sheets[sheet_idx];
        if row_idx >= sheet.row_count() {
            return Err(format!(
                "Row index {} is out of bounds (sheet has {} rows)",
                row_idx + 1,
                sheet.row_count()
            ));
        }
        sheet.delete_row(row_idx);
        Ok(())
    }

    /// Insert column at 0-based index
    pub fn insert_col(&mut self, sheet_idx: usize, col_idx: usize) -> Result<(), String> {
        let sheet = &mut self.sheets[sheet_idx];
        sheet.insert_col(col_idx);
        Ok(())
    }

    /// Delete column at 0-based index
    pub fn delete_col(&mut self, sheet_idx: usize, col_idx: usize) -> Result<(), String> {
        let sheet = &mut self.sheets[sheet_idx];
        if col_idx >= sheet.col_count() {
            return Err(format!(
                "Column index {} is out of bounds (sheet has {} columns)",
                col_idx + 1,
                sheet.col_count()
            ));
        }
        sheet.delete_col(col_idx);
        Ok(())
    }

    /// Add new sheet with specified name
    pub fn add_sheet(&mut self, name: &str) -> Result<(), String> {
        if self
            .sheets
            .iter()
            .any(|s| s.name.eq_ignore_ascii_case(name))
        {
            return Err(format!("Sheet '{}' already exists", name));
        }

        let mut columns = Vec::new();
        for col_idx in 0..5 {
            let mut col = DataColumn::new(10);
            col.id = generate_unique_id();
            col.name = col_idx_to_letters(col_idx);
            columns.push(col);
        }

        let new_sheet = Sheet {
            id: generate_unique_id(),
            name: name.to_string(),
            columns,
            tables: Vec::new(),
            dependencies: std::collections::HashMap::new(),
            dependencies_rev: std::collections::HashMap::new(),
            uncommitted_actions: Vec::new(),
        };

        self.sheets.push(new_sheet);
        Ok(())
    }

    /// Delete sheet by name
    pub fn delete_sheet(&mut self, name: &str) -> Result<(), String> {
        let idx = self.find_sheet_index(Some(name))?;
        if self.sheets.len() <= 1 {
            return Err("Cannot delete the only sheet in the workbook".to_string());
        }
        self.sheets.remove(idx);
        Ok(())
    }

    /// Rename sheet
    pub fn rename_sheet(&mut self, old_name: &str, new_name: &str) -> Result<(), String> {
        let idx = self.find_sheet_index(Some(old_name))?;
        if self
            .sheets
            .iter()
            .enumerate()
            .any(|(i, s)| i != idx && s.name.eq_ignore_ascii_case(new_name))
        {
            return Err(format!("Sheet name '{}' is already taken", new_name));
        }
        self.sheets[idx].name = new_name.to_string();
        Ok(())
    }

    /// Add chart to workbook
    pub fn add_chart(
        &mut self,
        sheet_name: &str,
        chart_type: ChartType,
        range: String,
        title: Option<String>,
    ) -> Result<u64, String> {
        let _ = self.find_sheet_index(Some(sheet_name))?;
        let id = generate_unique_id();
        let name = format!("Chart {}", self.charts.len() + 1);

        let chart = Chart {
            id,
            name,
            chart_type,
            data_range: range,
            title,
            xlabel: None,
            ylabel: None,
            show_legend: true,
        };

        self.charts.push(chart);
        Ok(id)
    }

    /// Delete chart by u64 ID
    pub fn delete_chart(&mut self, id: u64) -> Result<(), String> {
        if let Some(pos) = self.charts.iter().position(|c| c.id == id) {
            self.charts.remove(pos);
            Ok(())
        } else {
            Err(format!("Chart with ID {} not found", id))
        }
    }

    /// Find the sheet that owns the table with the given name, and the
    /// table itself. Table names are unique across the whole workbook.
    pub fn find_table(&self, name: &str) -> Option<(&Sheet, &ExcelTable)> {
        self.sheets
            .iter()
            .find_map(|s| s.find_table(name).map(|t| (s, t)))
    }

    /// List every table in the workbook, alongside the name of the sheet it
    /// lives on.
    pub fn list_tables(&self) -> Vec<(&str, &ExcelTable)> {
        self.sheets
            .iter()
            .flat_map(|s| s.tables.iter().map(move |t| (s.name.as_str(), t)))
            .collect()
    }

    fn find_table_sheet_index(&self, name: &str) -> Result<usize, String> {
        self.sheets
            .iter()
            .position(|s| s.find_table(name).is_some())
            .ok_or_else(|| format!("Table '{}' not found", name))
    }

    fn table_name_taken(&self, name: &str) -> bool {
        self.sheets
            .iter()
            .any(|s| s.tables.iter().any(|t| t.name.eq_ignore_ascii_case(name)))
    }

    /// Define a new Excel Table over an existing cell range on a sheet.
    /// Table names are unique across the entire workbook (not just the
    /// sheet), matching how Excel itself scopes structured-reference names.
    #[allow(clippy::too_many_arguments)]
    pub fn add_table(
        &mut self,
        sheet_name: Option<&str>,
        name: &str,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
        has_header_row: bool,
        has_totals_row: bool,
    ) -> Result<u64, String> {
        if self.table_name_taken(name) {
            return Err(format!("Table '{}' already exists", name));
        }
        let idx = self.find_sheet_index(sheet_name)?;
        self.sheets[idx].add_table(
            name.to_string(),
            start_row,
            start_col,
            end_row,
            end_col,
            has_header_row,
            has_totals_row,
        )
    }

    /// Delete a table by name (leaves the underlying cell contents alone).
    pub fn delete_table(&mut self, name: &str) -> Result<(), String> {
        let idx = self.find_table_sheet_index(name)?;
        self.sheets[idx].delete_table_by_name(name)
    }

    /// Rename a table.
    pub fn rename_table(&mut self, old_name: &str, new_name: &str) -> Result<(), String> {
        if !old_name.eq_ignore_ascii_case(new_name) && self.table_name_taken(new_name) {
            return Err(format!("Table name '{}' is already taken", new_name));
        }
        let idx = self.find_table_sheet_index(old_name)?;
        self.sheets[idx].rename_table(old_name, new_name)?;
        // Excel updates every structured reference to a renamed table, so a
        // formula like `=SUM(Sales[Amount])` keeps working after "Sales" is
        // renamed; match that instead of silently breaking those formulas.
        self.rewrite_table_references(old_name, Some(new_name), None);
        self.evaluate()
    }

    /// Rewrites every formula in the workbook that structurally references
    /// `table_name` (optionally renaming the table and/or one column),
    /// mirroring how Excel keeps structured references in sync when a Table
    /// or one of its column headers is renamed.
    fn rewrite_table_references(
        &mut self,
        table_name: &str,
        new_table_name: Option<&str>,
        col_rename: Option<(&str, &str)>,
    ) {
        for sheet in &mut self.sheets {
            for col_idx in 0..sheet.columns.len() {
                let row_count = sheet.columns[col_idx].src.len();
                for row_idx in 0..row_count {
                    let src = sheet.columns[col_idx].src[row_idx].clone();
                    if let Some(new_src) = libvisi::core::parser::rewrite_structured_table_reference(
                        &src,
                        table_name,
                        new_table_name,
                        col_rename,
                    ) {
                        sheet.set_cell_src(row_idx, col_idx, new_src);
                    }
                }
            }
        }
    }

    /// Resize a table by moving its bottom-right corner.
    pub fn resize_table(
        &mut self,
        name: &str,
        new_end_row: usize,
        new_end_col: usize,
    ) -> Result<(), String> {
        let idx = self.find_table_sheet_index(name)?;
        self.sheets[idx].resize_table(name, new_end_row, new_end_col)
    }

    /// Rename one column (0-based, relative to the table) of a table.
    pub fn rename_table_column(
        &mut self,
        table_name: &str,
        col_index: usize,
        new_name: &str,
    ) -> Result<(), String> {
        let idx = self.find_table_sheet_index(table_name)?;
        let old_col_name = self.sheets[idx]
            .find_table(table_name)
            .and_then(|t| t.columns.get(col_index).cloned())
            .ok_or_else(|| {
                format!(
                    "Column index {} out of bounds for table '{}'",
                    col_index, table_name
                )
            })?;
        self.sheets[idx].rename_table_column(table_name, col_index, new_name)?;
        // As with rename_table, keep dependent formulas working across the
        // rename instead of leaving them referencing the old column name.
        self.rewrite_table_references(table_name, None, Some((&old_col_name, new_name)));
        self.evaluate()
    }
}
