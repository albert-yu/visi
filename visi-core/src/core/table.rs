//! Excel Tables (ListObjects): named sub-ranges of a worksheet.
//!
//! Not to be confused with a [`Sheet`], which this codebase informally calls a
//! "table" in places. An [`ExcelTable`] lives *on* a sheet and may cover only
//! part of it.

use serde::{Deserialize, Serialize};

use crate::core::engine::{Sheet, generate_unique_id};

/// A named, rectangular range within a single worksheet, mirroring an Excel
/// Table (a.k.a. `ListObject`): a header row, a body of data rows, and an
/// optional totals row, all with stable per-column names that formulas can
/// reference via structured references (e.g. `Sales[Amount]`).
///
/// This is a distinct concept from a `Sheet`: elsewhere in this codebase a
/// `Sheet` is informally called a "table" (see `Sheet::new`'s default name
/// `"table_1"`), but an `ExcelTable` is a sub-range that lives *on* a sheet,
/// exactly like a real Excel Table can occupy only part of a worksheet.
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
pub struct ExcelTable {
    /// Workbook-unique identifier, stable across renames.
    pub id: u64,
    /// The table's name, as a structured reference spells it. Unique
    /// workbook-wide and matched case-insensitively.
    pub name: String,
    /// The sheet this table occupies part of.
    pub sheet_id: u64,
    /// Topmost row of the range, 0-based -- the header row when there is one.
    pub start_row: usize,
    /// Leftmost column of the range, 0-based.
    pub start_col: usize,
    /// Bottommost row of the range, 0-based and inclusive -- the totals row
    /// when there is one.
    pub end_row: usize,
    /// Rightmost column of the range, 0-based and inclusive.
    pub end_col: usize,
    /// Whether the first row is a header rather than data.
    pub has_header_row: bool,
    /// Whether the last row is a totals row rather than data.
    pub has_totals_row: bool,
    /// Column names, in sheet-column order, one per column in
    /// `start_col..=end_col`. Kept in sync with the header row's cell text
    /// (when `has_header_row` is true) by the CRUD methods in this file.
    pub columns: Vec<String>,
    /// Visual style theme name (e.g. "TableStyleMedium9", "TableStyleLight1", or custom theme)
    #[serde(default)]
    pub style_name: Option<String>,
}

impl ExcelTable {
    /// Sets the table's visual style, or clears it with `None`.
    pub fn set_style_name(&mut self, style_name: Option<String>) {
        self.style_name = style_name;
    }

    /// Total rows in the range, header and totals rows included.
    pub fn row_count(&self) -> usize {
        self.end_row - self.start_row + 1
    }

    /// Columns in the range.
    pub fn col_count(&self) -> usize {
        self.end_col - self.start_col + 1
    }

    /// First row of the table's actual data body (excludes the header row).
    pub fn data_start_row(&self) -> usize {
        self.start_row + usize::from(self.has_header_row)
    }

    /// Last row of the table's actual data body (excludes the totals row).
    /// May be less than `data_start_row()` for a table with no data rows.
    pub fn data_end_row(&self) -> usize {
        self.end_row - usize::from(self.has_totals_row)
    }

    /// The header row's sheet-row index, or `None` if the table has no
    /// header.
    pub fn header_row(&self) -> Option<usize> {
        self.has_header_row.then_some(self.start_row)
    }

    /// The totals row's sheet-row index, or `None` if the table has no
    /// totals row.
    pub fn totals_row(&self) -> Option<usize> {
        self.has_totals_row.then_some(self.end_row)
    }

    /// Index (0-based, relative to the table's own columns) of the column
    /// with the given name, matched case-insensitively as Excel does for
    /// structured references.
    pub fn local_column_index(&self, name: &str) -> Option<usize> {
        self.columns
            .iter()
            .position(|c| c.eq_ignore_ascii_case(name))
    }

    /// Whether this table's range overlaps the given rectangular range.
    pub fn overlaps(
        &self,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
    ) -> bool {
        self.start_row <= end_row
            && start_row <= self.end_row
            && self.start_col <= end_col
            && start_col <= self.end_col
    }
}

fn validate_table_name(name: &str) -> Result<(), String> {
    let trimmed = name.trim();
    if trimmed.is_empty() {
        return Err("Table name cannot be empty".to_string());
    }
    let first = trimmed.chars().next().unwrap();
    if !(first.is_alphabetic() || first == '_') {
        return Err(format!(
            "Table name '{}' must start with a letter or underscore",
            name
        ));
    }
    if !trimmed
        .chars()
        .all(|c| c.is_alphanumeric() || c == '_' || c == '.')
    {
        return Err(format!(
            "Table name '{}' may only contain letters, digits, underscores, and periods",
            name
        ));
    }
    Ok(())
}

fn check_duplicate_column_names(columns: &[String]) -> Result<(), String> {
    let mut seen = std::collections::HashSet::new();
    for c in columns {
        if !seen.insert(c.to_ascii_lowercase()) {
            return Err(format!("Duplicate column name '{}' in table header row", c));
        }
    }
    Ok(())
}

impl Sheet {
    /// Finds a table on this sheet by name, matched case-insensitively as
    /// Excel does.
    pub fn find_table(&self, name: &str) -> Option<&ExcelTable> {
        self.tables
            .iter()
            .find(|t| t.name.eq_ignore_ascii_case(name))
    }

    /// [`Sheet::find_table`], mutably.
    pub fn find_table_mut(&mut self, name: &str) -> Option<&mut ExcelTable> {
        self.tables
            .iter_mut()
            .find(|t| t.name.eq_ignore_ascii_case(name))
    }

    /// Reads the header text for sheet column `col_idx` at `header_row`, if
    /// non-blank; otherwise falls back to a default "ColumnN" name (N is
    /// 1-based within the table).
    fn table_column_header(&self, header_row: usize, col_idx: usize, local_idx: usize) -> String {
        // Prefer the cell's computed value (what a user actually sees) over
        // its raw source text, in case a header cell happens to hold a
        // formula rather than plain text; fall back to raw source for
        // cells that haven't been committed/evaluated yet.
        let computed = self
            .columns
            .get(col_idx)
            .and_then(|c| c.data.get(header_row))
            .map(|d| d.to_string())
            .filter(|s| !s.is_empty());
        computed
            .or_else(|| {
                self.columns
                    .get(col_idx)
                    .and_then(|c| c.src.get(header_row))
                    .map(|s| s.trim().to_string())
                    .filter(|s| !s.is_empty())
            })
            .unwrap_or_else(|| format!("Column{}", local_idx + 1))
    }

    /// Defines a new Excel Table over the rectangular range
    /// `start_row..=end_row` x `start_col..=end_col` (0-based, inclusive)
    /// on this sheet. Column names are read from the header row's existing
    /// cell text when `has_header_row` is true, falling back to "ColumnN"
    /// for blank cells; otherwise every column gets a default "ColumnN"
    /// name.
    #[allow(clippy::too_many_arguments)]
    pub fn add_table(
        &mut self,
        name: String,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
        has_header_row: bool,
        has_totals_row: bool,
    ) -> Result<u64, String> {
        validate_table_name(&name)?;
        if self.find_table(&name).is_some() {
            return Err(format!(
                "Table '{}' already exists on sheet '{}'",
                name, self.name
            ));
        }
        if end_row < start_row || end_col < start_col {
            return Err("Table range end must not precede its start".to_string());
        }
        let (row_count, col_count) = (self.row_count(), self.col_count());
        if end_row >= row_count || end_col >= col_count {
            return Err(format!(
                "Table range exceeds sheet bounds ({} rows x {} cols)",
                row_count, col_count
            ));
        }
        if let Some(existing) = self
            .tables
            .iter()
            .find(|t| t.overlaps(start_row, start_col, end_row, end_col))
        {
            return Err(format!(
                "Table range overlaps existing table '{}' on sheet '{}'",
                existing.name, self.name
            ));
        }

        let columns: Vec<String> = (start_col..=end_col)
            .enumerate()
            .map(|(local_idx, col_idx)| {
                if has_header_row {
                    self.table_column_header(start_row, col_idx, local_idx)
                } else {
                    format!("Column{}", local_idx + 1)
                }
            })
            .collect();
        check_duplicate_column_names(&columns)?;

        let id = generate_unique_id();
        self.tables.push(ExcelTable {
            id,
            name,
            sheet_id: self.id,
            start_row,
            start_col,
            end_row,
            end_col,
            has_header_row,
            has_totals_row,
            columns,
            style_name: None,
        });
        Ok(id)
    }

    /// Removes a table definition from this sheet, leaving the cells it
    /// covered untouched.
    ///
    /// # Errors
    ///
    /// Returns a message if no table on this sheet has that name.
    pub fn delete_table_by_name(&mut self, name: &str) -> Result<(), String> {
        if let Some(pos) = self
            .tables
            .iter()
            .position(|t| t.name.eq_ignore_ascii_case(name))
        {
            self.tables.remove(pos);
            Ok(())
        } else {
            Err(format!(
                "Table '{}' not found on sheet '{}'",
                name, self.name
            ))
        }
    }

    /// Renames a table on this sheet.
    ///
    /// Renaming here does *not* rewrite the formulas that reference the table
    /// -- that cascade is `WorkbookManager::rename_table`'s job, and it is
    /// what keeps `Sales[Amount]` pointing at the renamed table. Prefer that
    /// entry point unless you are rewriting the references yourself.
    ///
    /// # Errors
    ///
    /// Returns a message if the new name is not a valid table name, if
    /// another table on this sheet already has it, or if no table on this
    /// sheet has `old_name`.
    pub fn rename_table(&mut self, old_name: &str, new_name: &str) -> Result<(), String> {
        validate_table_name(new_name)?;
        if self.tables.iter().any(|t| {
            !t.name.eq_ignore_ascii_case(old_name) && t.name.eq_ignore_ascii_case(new_name)
        }) {
            return Err(format!("Table name '{}' is already taken", new_name));
        }
        let sheet_name = self.name.clone();
        let table = self
            .find_table_mut(old_name)
            .ok_or_else(|| format!("Table '{}' not found on sheet '{}'", old_name, sheet_name))?;
        table.name = new_name.to_string();
        Ok(())
    }

    /// Extends or shrinks a table's range by moving its bottom-right corner
    /// to `new_end_row`/`new_end_col` (the top-left corner never moves).
    /// Column names for any newly-included columns are read from the
    /// header row (or default to "ColumnN"); names for columns that
    /// already existed are preserved by position.
    pub fn resize_table(
        &mut self,
        name: &str,
        new_end_row: usize,
        new_end_col: usize,
    ) -> Result<(), String> {
        let (id, start_row, start_col, has_header_row, old_columns) = {
            let table = self
                .find_table(name)
                .ok_or_else(|| format!("Table '{}' not found on sheet '{}'", name, self.name))?;
            (
                table.id,
                table.start_row,
                table.start_col,
                table.has_header_row,
                table.columns.clone(),
            )
        };

        if new_end_row < start_row || new_end_col < start_col {
            return Err("Table range end must not precede its start".to_string());
        }
        let (row_count, col_count) = (self.row_count(), self.col_count());
        if new_end_row >= row_count || new_end_col >= col_count {
            return Err(format!(
                "Table range exceeds sheet bounds ({} rows x {} cols)",
                row_count, col_count
            ));
        }
        if let Some(existing) = self
            .tables
            .iter()
            .find(|t| t.id != id && t.overlaps(start_row, start_col, new_end_row, new_end_col))
        {
            return Err(format!(
                "Resized range would overlap existing table '{}' on sheet '{}'",
                existing.name, self.name
            ));
        }

        let new_columns: Vec<String> = (start_col..=new_end_col)
            .enumerate()
            .map(|(local_idx, col_idx)| {
                old_columns.get(local_idx).cloned().unwrap_or_else(|| {
                    if has_header_row {
                        self.table_column_header(start_row, col_idx, local_idx)
                    } else {
                        format!("Column{}", local_idx + 1)
                    }
                })
            })
            .collect();
        check_duplicate_column_names(&new_columns)?;

        let table = self.find_table_mut(name).unwrap();
        table.end_row = new_end_row;
        table.end_col = new_end_col;
        table.columns = new_columns;
        Ok(())
    }

    /// Renames one column (0-based, relative to the table) of a table,
    /// updating both its stored name and the header row's cell text (if
    /// the table has one).
    pub fn rename_table_column(
        &mut self,
        table_name: &str,
        col_index: usize,
        new_name: &str,
    ) -> Result<(), String> {
        let trimmed = new_name.trim();
        if trimmed.is_empty() {
            return Err("Column name cannot be empty".to_string());
        }

        let (sheet_col, header_row, has_header_row) = {
            let table = self.find_table(table_name).ok_or_else(|| {
                format!("Table '{}' not found on sheet '{}'", table_name, self.name)
            })?;
            if col_index >= table.columns.len() {
                return Err(format!(
                    "Column index {} out of bounds (table '{}' has {} columns)",
                    col_index,
                    table_name,
                    table.columns.len()
                ));
            }
            if table
                .columns
                .iter()
                .enumerate()
                .any(|(i, c)| i != col_index && c.eq_ignore_ascii_case(trimmed))
            {
                return Err(format!(
                    "Table '{}' already has a column named '{}'",
                    table_name, trimmed
                ));
            }
            (
                table.start_col + col_index,
                table.start_row,
                table.has_header_row,
            )
        };

        if let Some(table) = self.find_table_mut(table_name) {
            table.columns[col_index] = trimmed.to_string();
        }
        if has_header_row {
            self.set_cell_src(header_row, sheet_col, trimmed.to_string());
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::core::engine::SheetInit;

    fn sheet_with_data() -> Sheet {
        let mut sheet = Sheet::new(SheetInit {
            name: Some("Sheet1".to_string()),
            rows: 6,
            cols: 3,
            ..Default::default()
        });
        // Row 0: headers, rows 1-4: data, row 5: totals.
        let header = ["Name", "Amount", "Qty"];
        let data = [
            ["Widget", "10", "2"],
            ["Gadget", "20", "3"],
            ["Gizmo", "30", "4"],
            ["Doohickey", "40", "5"],
        ];
        for (c, h) in header.iter().enumerate() {
            sheet.set_cell_src(0, c, h.to_string());
        }
        for (r, row) in data.iter().enumerate() {
            for (c, v) in row.iter().enumerate() {
                sheet.set_cell_src(r + 1, c, v.to_string());
            }
        }
        sheet.set_cell_src(5, 1, "=SUM(B2:B5)".to_string());
        sheet.commit(None).unwrap();
        sheet
    }

    #[test]
    fn test_add_table_reads_headers_and_bounds() {
        let mut sheet = sheet_with_data();
        let id = sheet
            .add_table("Sales".to_string(), 0, 0, 5, 2, true, true)
            .unwrap();
        let table = sheet.find_table("Sales").unwrap();
        assert_eq!(table.id, id);
        assert_eq!(table.columns, vec!["Name", "Amount", "Qty"]);
        assert_eq!(table.data_start_row(), 1);
        assert_eq!(table.data_end_row(), 4);
        assert_eq!(table.header_row(), Some(0));
        assert_eq!(table.totals_row(), Some(5));
    }

    #[test]
    fn test_add_table_no_header_row_uses_default_names() {
        let mut sheet = sheet_with_data();
        sheet
            .add_table("Raw".to_string(), 1, 0, 4, 2, false, false)
            .unwrap();
        let table = sheet.find_table("Raw").unwrap();
        assert_eq!(table.columns, vec!["Column1", "Column2", "Column3"]);
        assert_eq!(table.data_start_row(), 1);
        assert_eq!(table.data_end_row(), 4);
        assert_eq!(table.totals_row(), None);
    }

    #[test]
    fn test_add_table_rejects_duplicate_name() {
        let mut sheet = sheet_with_data();
        sheet
            .add_table("Sales".to_string(), 0, 0, 4, 2, true, false)
            .unwrap();
        let err = sheet
            .add_table("Sales".to_string(), 0, 0, 4, 2, true, false)
            .unwrap_err();
        assert!(err.contains("already exists"));
    }

    #[test]
    fn test_add_table_rejects_invalid_name() {
        let mut sheet = sheet_with_data();
        let err = sheet
            .add_table("1Sales".to_string(), 0, 0, 4, 2, true, false)
            .unwrap_err();
        assert!(err.contains("must start with"));

        let err2 = sheet
            .add_table("Sales Report".to_string(), 0, 0, 4, 2, true, false)
            .unwrap_err();
        assert!(err2.contains("letters, digits"));
    }

    #[test]
    fn test_add_table_rejects_out_of_bounds_range() {
        let mut sheet = sheet_with_data();
        let err = sheet
            .add_table("Sales".to_string(), 0, 0, 10, 2, true, false)
            .unwrap_err();
        assert!(err.contains("exceeds sheet bounds"));
    }

    #[test]
    fn test_add_table_rejects_overlap() {
        let mut sheet = sheet_with_data();
        sheet
            .add_table("Sales".to_string(), 0, 0, 4, 1, true, false)
            .unwrap();
        let err = sheet
            .add_table("Other".to_string(), 0, 1, 4, 2, true, false)
            .unwrap_err();
        assert!(err.contains("overlaps"));
    }

    #[test]
    fn test_delete_and_rename_table() {
        let mut sheet = sheet_with_data();
        sheet
            .add_table("Sales".to_string(), 0, 0, 4, 2, true, false)
            .unwrap();

        sheet.rename_table("Sales", "Revenue").unwrap();
        assert!(sheet.find_table("Sales").is_none());
        assert!(sheet.find_table("Revenue").is_some());

        sheet.delete_table_by_name("Revenue").unwrap();
        assert!(sheet.find_table("Revenue").is_none());

        let err = sheet.delete_table_by_name("Revenue").unwrap_err();
        assert!(err.contains("not found"));
    }

    #[test]
    fn test_resize_table_grows_and_shrinks() {
        let mut sheet = sheet_with_data();
        sheet
            .add_table("Sales".to_string(), 0, 0, 3, 1, true, false)
            .unwrap();
        assert_eq!(sheet.find_table("Sales").unwrap().columns.len(), 2);

        // Grow to include the Qty column and one more row.
        sheet.resize_table("Sales", 4, 2).unwrap();
        let table = sheet.find_table("Sales").unwrap();
        assert_eq!(table.end_row, 4);
        assert_eq!(table.end_col, 2);
        assert_eq!(table.columns, vec!["Name", "Amount", "Qty"]);

        // Shrink back down; existing column names are preserved by position.
        sheet.resize_table("Sales", 3, 0).unwrap();
        let table = sheet.find_table("Sales").unwrap();
        assert_eq!(table.end_row, 3);
        assert_eq!(table.end_col, 0);
        assert_eq!(table.columns, vec!["Name"]);
    }

    #[test]
    fn test_rename_table_column_updates_header_cell() {
        let mut sheet = sheet_with_data();
        sheet
            .add_table("Sales".to_string(), 0, 0, 4, 2, true, false)
            .unwrap();

        sheet.rename_table_column("Sales", 1, "Total").unwrap();
        assert_eq!(
            sheet.find_table("Sales").unwrap().columns,
            vec!["Name", "Total", "Qty"]
        );
        // The header row's actual cell text is kept in sync.
        assert_eq!(sheet.columns[1].src[0], "Total");
    }

    #[test]
    fn test_rename_table_column_rejects_duplicate() {
        let mut sheet = sheet_with_data();
        sheet
            .add_table("Sales".to_string(), 0, 0, 4, 2, true, false)
            .unwrap();
        let err = sheet.rename_table_column("Sales", 1, "Name").unwrap_err();
        assert!(err.contains("already has a column"));
    }
}
