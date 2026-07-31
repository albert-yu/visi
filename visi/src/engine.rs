use crate::utils::col_idx_to_letters;
use libvisi::core::{
    ExcelTable, PivotAggregation, PivotArea, PivotField, PivotFilterField, PivotGrid, PivotSource,
    PivotTable, PivotValueField,
    chart::{Chart, ChartType},
    compute_pivot,
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
    pub pivot_tables: Vec<PivotTable>,
}

/// Quotes a materialized pivot label that would otherwise be re-parsed as a
/// number, boolean, or formula by `Sheet::commit`'s literal-cell parsing
/// (mirrors `xlsx::text_cell_src`'s treatment of imported text cells).
fn pivot_label_literal(text: &str) -> String {
    if text.is_empty() {
        String::new()
    } else if text.starts_with('=')
        || text.parse::<f64>().is_ok()
        || text.eq_ignore_ascii_case("true")
        || text.eq_ignore_ascii_case("false")
    {
        format!("\"{}\"", text)
    } else {
        text.to_string()
    }
}

/// Renders one aggregated pivot value as literal cell text; errors (e.g.
/// `AVERAGE` over zero numeric records) are written as their Excel error
/// string rather than `ResultData`'s human-readable `"Error: ..."` form.
fn pivot_value_literal(v: &ResultData) -> String {
    match v {
        ResultData::Error(e) => e.clone(),
        other => other.to_string(),
    }
}

fn remove_pivot_field(fields: &mut Vec<PivotField>, column: &str) -> bool {
    let before = fields.len();
    fields.retain(|f| !f.column.eq_ignore_ascii_case(column));
    before != fields.len()
}

impl WorkbookManager {
    /// Load Excel workbook from bytes buffer
    pub fn load_bytes(buffer: &[u8]) -> Result<Self, String> {
        let (imported_tables, charts, pivot_tables) =
            import_xlsx_data(buffer, &[], |_, _, _| {}).map_err(|e| e.to_string())?;

        let sheets = imported_tables.into_iter().map(|it| it.sheet).collect();
        Ok(Self {
            sheets,
            charts,
            pivot_tables,
        })
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
                pivot_tables: Vec::new(),
            };
            wb.add_sheet("Sheet1")?;
            Ok(wb)
        } else {
            Self::load_file(path_str)
        }
    }

    /// Save Excel workbook to file path or stdout ("-")
    pub fn save_file(&self, path_str: &str) -> Result<(), String> {
        let bytes = export_xlsx_data(&self.sheets, &self.charts, &self.pivot_tables)
            .map_err(|e| e.to_string())?;

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

    /// Find a pivot table by name (case-insensitive).
    pub fn find_pivot_table(&self, name: &str) -> Option<&PivotTable> {
        self.pivot_tables
            .iter()
            .find(|p| p.name.eq_ignore_ascii_case(name))
    }

    fn find_pivot_table_index(&self, name: &str) -> Result<usize, String> {
        self.pivot_tables
            .iter()
            .position(|p| p.name.eq_ignore_ascii_case(name))
            .ok_or_else(|| format!("Pivot table '{}' not found", name))
    }

    /// List every pivot table in the workbook.
    pub fn list_pivot_tables(&self) -> &[PivotTable] {
        &self.pivot_tables
    }

    fn pivot_table_name_taken(&self, name: &str) -> bool {
        self.pivot_tables
            .iter()
            .any(|p| p.name.eq_ignore_ascii_case(name))
    }

    /// Defines a new pivot table sourced from an existing Excel Table, with
    /// no fields assigned yet -- mirroring Excel inserting an empty
    /// PivotTable shell that fills in as fields are added to it.
    pub fn add_pivot_table_from_table(
        &mut self,
        name: &str,
        source_table_name: &str,
        dest_sheet_name: Option<&str>,
        dest_row: usize,
        dest_col: usize,
    ) -> Result<u64, String> {
        if self.pivot_table_name_taken(name) {
            return Err(format!("Pivot table '{}' already exists", name));
        }
        self.find_table(source_table_name)
            .ok_or_else(|| format!("Table '{}' not found", source_table_name))?;
        let dest_idx = self.find_sheet_index(dest_sheet_name)?;
        let id = generate_unique_id();
        self.pivot_tables.push(PivotTable {
            id,
            name: name.to_string(),
            source: PivotSource::Table {
                name: source_table_name.to_string(),
            },
            dest_sheet_id: self.sheets[dest_idx].id,
            dest_row,
            dest_col,
            row_fields: Vec::new(),
            col_fields: Vec::new(),
            value_fields: Vec::new(),
            filter_fields: Vec::new(),
            grand_totals_row: true,
            grand_totals_col: true,
            last_output_end_row: None,
            last_output_end_col: None,
        });
        self.refresh_pivot_table(name)?;
        Ok(id)
    }

    /// Defines a new pivot table sourced from a plain cell range (its first
    /// row is treated as column headers), with no fields assigned yet.
    #[allow(clippy::too_many_arguments)]
    pub fn add_pivot_table_from_range(
        &mut self,
        name: &str,
        source_sheet_name: Option<&str>,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
        dest_sheet_name: Option<&str>,
        dest_row: usize,
        dest_col: usize,
    ) -> Result<u64, String> {
        if self.pivot_table_name_taken(name) {
            return Err(format!("Pivot table '{}' already exists", name));
        }
        let src_idx = self.find_sheet_index(source_sheet_name)?;
        let dest_idx = self.find_sheet_index(dest_sheet_name)?;
        let id = generate_unique_id();
        self.pivot_tables.push(PivotTable {
            id,
            name: name.to_string(),
            source: PivotSource::Range {
                sheet_id: self.sheets[src_idx].id,
                start_row,
                start_col,
                end_row,
                end_col,
            },
            dest_sheet_id: self.sheets[dest_idx].id,
            dest_row,
            dest_col,
            row_fields: Vec::new(),
            col_fields: Vec::new(),
            value_fields: Vec::new(),
            filter_fields: Vec::new(),
            grand_totals_row: true,
            grand_totals_col: true,
            last_output_end_row: None,
            last_output_end_col: None,
        });
        self.refresh_pivot_table(name)?;
        Ok(id)
    }

    /// Deletes a pivot table definition and clears its last rendered output
    /// range (leaves the source data untouched).
    pub fn delete_pivot_table(&mut self, name: &str) -> Result<(), String> {
        let idx = self.find_pivot_table_index(name)?;
        let pivot = self.pivot_tables.remove(idx);
        if let (Some(end_row), Some(end_col)) =
            (pivot.last_output_end_row, pivot.last_output_end_col)
            && let Some(sheet_idx) = self.sheets.iter().position(|s| s.id == pivot.dest_sheet_id)
        {
            self.clear_range(sheet_idx, pivot.dest_row, pivot.dest_col, end_row, end_col);
        }
        Ok(())
    }

    /// Renames a pivot table (names are unique workbook-wide, like tables).
    pub fn rename_pivot_table(&mut self, old_name: &str, new_name: &str) -> Result<(), String> {
        if !old_name.eq_ignore_ascii_case(new_name) && self.pivot_table_name_taken(new_name) {
            return Err(format!("Pivot table name '{}' is already taken", new_name));
        }
        let idx = self.find_pivot_table_index(old_name)?;
        self.pivot_tables[idx].name = new_name.to_string();
        Ok(())
    }

    /// Adds a field to one of a pivot table's four areas (Row/Column/
    /// Value/Filter) and immediately refreshes its output, mirroring
    /// Excel's live-updating field list.
    pub fn add_pivot_field(
        &mut self,
        pivot_name: &str,
        area: PivotArea,
        column: &str,
        aggregation: Option<PivotAggregation>,
    ) -> Result<(), String> {
        let idx = self.find_pivot_table_index(pivot_name)?;
        match area {
            PivotArea::Row => self.pivot_tables[idx]
                .row_fields
                .push(PivotField::new(column)),
            PivotArea::Column => self.pivot_tables[idx]
                .col_fields
                .push(PivotField::new(column)),
            PivotArea::Value => {
                let agg = aggregation.unwrap_or(PivotAggregation::Sum);
                self.pivot_tables[idx]
                    .value_fields
                    .push(PivotValueField::new(column, agg));
            }
            PivotArea::Filter => self.pivot_tables[idx]
                .filter_fields
                .push(PivotFilterField::new(column)),
        }
        self.refresh_pivot_table(pivot_name)
    }

    /// Removes a field from one of a pivot table's four areas and
    /// refreshes its output.
    pub fn remove_pivot_field(
        &mut self,
        pivot_name: &str,
        area: PivotArea,
        column: &str,
    ) -> Result<(), String> {
        let idx = self.find_pivot_table_index(pivot_name)?;
        let removed = match area {
            PivotArea::Row => remove_pivot_field(&mut self.pivot_tables[idx].row_fields, column),
            PivotArea::Column => remove_pivot_field(&mut self.pivot_tables[idx].col_fields, column),
            PivotArea::Value => {
                let before = self.pivot_tables[idx].value_fields.len();
                self.pivot_tables[idx]
                    .value_fields
                    .retain(|f| !f.column.eq_ignore_ascii_case(column));
                before != self.pivot_tables[idx].value_fields.len()
            }
            PivotArea::Filter => {
                let before = self.pivot_tables[idx].filter_fields.len();
                self.pivot_tables[idx]
                    .filter_fields
                    .retain(|f| !f.column.eq_ignore_ascii_case(column));
                before != self.pivot_tables[idx].filter_fields.len()
            }
        };
        if !removed {
            return Err(format!(
                "Field '{}' not found in that area of pivot table '{}'",
                column, pivot_name
            ));
        }
        self.refresh_pivot_table(pivot_name)
    }

    /// Restricts (or clears, with `values: None`) a filter field's allowed
    /// values and refreshes the pivot table's output.
    pub fn set_pivot_filter(
        &mut self,
        pivot_name: &str,
        column: &str,
        values: Option<Vec<String>>,
    ) -> Result<(), String> {
        let idx = self.find_pivot_table_index(pivot_name)?;
        let field = self.pivot_tables[idx]
            .filter_fields
            .iter_mut()
            .find(|f| f.column.eq_ignore_ascii_case(column))
            .ok_or_else(|| {
                format!(
                    "Filter field '{}' not found on pivot table '{}'",
                    column, pivot_name
                )
            })?;
        field.selected_values = values;
        self.refresh_pivot_table(pivot_name)
    }

    /// Recomputes a pivot table's aggregation and re-materializes it as
    /// plain values onto its destination sheet. Like Excel, a pivot table
    /// only updates on an explicit refresh, never automatically as its
    /// source data changes.
    pub fn refresh_pivot_table(&mut self, pivot_name: &str) -> Result<(), String> {
        let idx = self.find_pivot_table_index(pivot_name)?;
        let pivot = self.pivot_tables[idx].clone();
        let dest_idx = self
            .sheets
            .iter()
            .position(|s| s.id == pivot.dest_sheet_id)
            .ok_or_else(|| "Pivot table's destination sheet no longer exists".to_string())?;

        let grid: Option<PivotGrid> = if pivot.value_fields.is_empty() {
            None
        } else {
            Some(compute_pivot(&self.sheets, &pivot)?)
        };

        // Clear the previously rendered area first, since a refresh can
        // shrink the grid (fewer groups, a narrower filter, etc).
        if let (Some(old_end_row), Some(old_end_col)) =
            (pivot.last_output_end_row, pivot.last_output_end_col)
        {
            self.clear_range(
                dest_idx,
                pivot.dest_row,
                pivot.dest_col,
                old_end_row,
                old_end_col,
            );
        }

        let new_bounds = grid.as_ref().map(|grid| {
            let height = grid.height();
            let width = grid.width.max(1);
            self.ensure_capacity(
                dest_idx,
                pivot.dest_row + height.saturating_sub(1),
                pivot.dest_col + width.saturating_sub(1),
            );

            let mut r = pivot.dest_row;
            for header in &grid.header_rows {
                for (c, text) in header.iter().enumerate() {
                    self.set_cell(dest_idx, r, pivot.dest_col + c, pivot_label_literal(text));
                }
                r += 1;
            }
            for body in &grid.body_rows {
                for (c, label) in body.row_labels.iter().enumerate() {
                    self.set_cell(dest_idx, r, pivot.dest_col + c, pivot_label_literal(label));
                }
                for (c, val) in body.values.iter().enumerate() {
                    self.set_cell(
                        dest_idx,
                        r,
                        pivot.dest_col + body.row_labels.len() + c,
                        pivot_value_literal(val),
                    );
                }
                r += 1;
            }
            (
                pivot.dest_row + height.saturating_sub(1),
                pivot.dest_col + width.saturating_sub(1),
            )
        });

        self.pivot_tables[idx].last_output_end_row = new_bounds.map(|(r, _)| r);
        self.pivot_tables[idx].last_output_end_col = new_bounds.map(|(_, c)| c);
        self.evaluate()
    }

    /// Blanks every cell in the given rectangular range (inclusive),
    /// clipped to the sheet's current bounds. Used to wipe a pivot table's
    /// previous output before re-rendering a possibly smaller grid.
    fn clear_range(
        &mut self,
        sheet_idx: usize,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
    ) {
        if sheet_idx >= self.sheets.len() {
            return;
        }
        let (row_count, col_count) = {
            let s = &self.sheets[sheet_idx];
            (s.row_count(), s.col_count())
        };
        if row_count == 0 || col_count == 0 {
            return;
        }
        for r in start_row..=end_row.min(row_count - 1) {
            for c in start_col..=end_col.min(col_count - 1) {
                self.sheets[sheet_idx].set_cell_src(r, c, String::new());
            }
        }
    }
}
