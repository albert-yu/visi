//! Cell and range accessors, text editing, styling, and the row/column
//! structural operations.
//!
//! Split out of the parent module; these are the operations that change a
//! sheet's *shape* or a cell's raw content, as opposed to evaluating it.

use super::super::column::{ColumnPosition, DataColumn};
use super::{CellRef, Direction, ResultData, Sheet, TextCellRef};

/// The word surrounding `char_offset` in `text`, as a half-open range of
/// character indices.
///
/// A "word" is a run of alphanumerics and underscores, a run of whitespace, or
/// a run of punctuation -- so double-clicking in a formula selects a function
/// name or a cell reference rather than the whole line. An offset at the end
/// of the text, or one just past a word onto whitespace, selects the word to
/// its left.
pub fn get_word_boundaries_from_str(text: &str, char_offset: usize) -> (usize, usize) {
    if text.is_empty() {
        return (0, 0);
    }

    let chars: Vec<char> = text.chars().collect();
    let len = chars.len();
    let offset = char_offset.min(len);

    let is_word_char = |c: char| c.is_alphanumeric() || c == '_';

    let (on_idx, on_c) = if offset < len {
        if chars[offset].is_whitespace() && offset > 0 && is_word_char(chars[offset - 1]) {
            (offset - 1, chars[offset - 1])
        } else {
            (offset, chars[offset])
        }
    } else if offset > 0 {
        (offset - 1, chars[offset - 1])
    } else {
        return (0, 0);
    };

    if on_c.is_whitespace() {
        let mut start = on_idx;
        while start > 0 && chars[start - 1].is_whitespace() {
            start -= 1;
        }
        let mut end = on_idx + 1;
        while end < len && chars[end].is_whitespace() {
            end += 1;
        }
        return (start, end);
    }

    if is_word_char(on_c) {
        let mut start = on_idx;
        while start > 0 && is_word_char(chars[start - 1]) {
            start -= 1;
        }
        let mut end = on_idx + 1;
        while end < len && is_word_char(chars[end]) {
            end += 1;
        }
        return (start, end);
    }

    let mut start = on_idx;
    while start > 0 && !is_word_char(chars[start - 1]) && !chars[start - 1].is_whitespace() {
        start -= 1;
    }
    let mut end = on_idx + 1;
    while end < len && !is_word_char(chars[end]) && !chars[end].is_whitespace() {
        end += 1;
    }
    (start, end)
}
impl Sheet {
    /// The computed value of a cell, or [`ResultData::None`] if it is empty
    /// or outside the sheet's allocated grid.
    ///
    /// Reflects the last [`Sheet::commit`]; a cell edited since then still
    /// reads as its old value.
    ///
    /// A date reads back as the plain numeric serial it is. Rendering it in
    /// the notation the cell carries is `Sheet::get_display_string`'s job, and
    /// only its -- do not format a `ResultData` directly if a user will see it.
    pub fn get_result_data(&self, cell: &CellRef) -> ResultData {
        let col = self.columns.get(cell.col);
        if let Some(col) = col {
            col.data.get(cell.row).unwrap_or(ResultData::None)
        } else {
            ResultData::None
        }
    }

    /// The date format a formula should inherit from the cells it reads, if
    /// any -- Excel's "date plus a number is still a date" behavior.
    ///
    /// The rule is deliberately about the *operator*, not about how many
    /// cells the formula touches, because those come apart: `=YEAR(A1)` reads
    /// exactly one date cell and returns a year, which is emphatically not a
    /// date. So only two shapes inherit:
    ///
    /// - a bare reference to a date cell (`=A1`), and
    /// - adding or subtracting a non-date from one (`=A1+1`, `=1+A1`).
    ///
    /// Everything else -- a function call, a product, a difference of two
    /// dates (which is a count of days) -- declines, leaving a plain number.
    pub(super) fn inherited_date_format(&self, ast: &crate::core::parser::Expr) -> Option<String> {
        use crate::core::parser::{Expr, Op};
        match ast {
            Expr::CellRef {
                sheet, row, col, ..
            } if sheet.is_none() => self
                .get_cell_style(*row, *col)
                .and_then(|s| s.num_format.clone())
                .filter(|code| crate::core::date::is_date_code(code)),
            Expr::BinaryOp {
                op: Op::Add | Op::Sub,
                left,
                right,
            } => {
                let left_fmt = self.inherited_date_format(left);
                let right_fmt = self.inherited_date_format(right);
                match (left_fmt, right_fmt) {
                    // Exactly one side is a date: the other is an offset in
                    // days, so the result stays that date's format.
                    (Some(fmt), None) | (None, Some(fmt)) => Some(fmt),
                    // Neither, or both (a day count) -- no date format.
                    _ => None,
                }
            }
            _ => None,
        }
    }

    /// The cell's value as it should be shown, honoring the cell's number
    /// format.
    ///
    /// A date cell holds a plain numeric serial, exactly as in Excel, so
    /// rendering it as a date is a display-time concern: this is the only
    /// place that turns 46195 back into `6/22/26`. Everything that shows a
    /// value to a user should go through here rather than formatting
    /// [`ResultData`] directly, which knows nothing about formats.
    pub fn get_display_string(&self, cell: &CellRef) -> String {
        let value = self.get_result_data(cell);
        let Some(code) = self
            .get_cell_style(cell.row, cell.col)
            .and_then(|s| s.num_format.as_deref())
        else {
            return value.to_string();
        };
        if !crate::core::date::is_date_code(code) {
            return value.to_string();
        }
        // Only a number is a date serial; text and errors render as-is.
        let serial = match value {
            ResultData::Float(f) => f,
            ResultData::Integer(i) => i as f64,
            _ => return value.to_string(),
        };
        if serial < 0.0 {
            return value.to_string();
        }
        crate::core::date::render_date_code(
            crate::core::date::excel_serial_to_date(serial),
            code,
            crate::core::date::StringCase::Title,
        )
    }

    /// Updates the src text of a particular cell but does
    /// not automatically evaluate. Call [`Sheet::commit`] to evaluate
    /// updated cells.
    /// Directly sets the src of a cell and marks it dirty.
    pub fn set_cell_src(&mut self, row: usize, col: usize, src: String) {
        let table_clone = self.clone();
        if let Some(column) = self.columns.get_mut(col)
            && row < column.src.len()
        {
            column.src[row] = src.clone();
            let compiled = crate::core::parser::compile_formula(&src, &[table_clone]);
            column.compiled_src[row] = compiled;
            column.mark_dirty(row);

            self.uncommitted_actions
                .push(crate::core::SheetAction::SetCellSrc {
                    sheet_name: self.name.clone(),
                    col,
                    row,
                    src,
                });
        }
    }

    /// Inserts text into a cell's source at a character offset, as typing
    /// into it would, then recompiles and marks it dirty.
    ///
    /// This is a text edit within one cell, not a range insert; see
    /// [`Sheet::insert_row`] and [`Sheet::insert_col`] for the structural
    /// operations. Out-of-range positions are ignored.
    pub fn insert(&mut self, pos: TextCellRef, input: &str) {
        let TextCellRef {
            row,
            col,
            char_offset,
        } = pos;
        let table_clone = self.clone();
        let existing_col = self.columns.get_mut(col);
        match existing_col {
            Some(existing_column) => {
                existing_column.insert(ColumnPosition { row, char_offset }, input);
                let src = existing_column.src[row].clone();
                let compiled = crate::core::parser::compile_formula(&src, &[table_clone]);
                existing_column.compiled_src[row] = compiled;
                existing_column.mark_dirty(row);
                self.uncommitted_actions
                    .push(crate::core::SheetAction::SetCellSrc {
                        sheet_name: self.name.clone(),
                        col,
                        row,
                        src,
                    });
            }
            None => {
                println!("Warning: column {} does not exist", col)
            }
        }
    }

    /// Delete one before (like backspace)
    pub fn delete_one_before(&mut self, pos: TextCellRef) {
        let char_offset = pos.char_offset;
        let start = if char_offset > 0 {
            TextCellRef {
                row: pos.row,
                col: pos.col,
                char_offset: char_offset - 1,
            }
        } else {
            pos.clone()
        };
        let end = pos;
        self.delete(start, end);
    }

    /// Deletes the text between two positions, recompiling and dirtying every
    /// cell it touches.
    ///
    /// Within a single cell this removes a character range; spanning cells it
    /// truncates the first, clears those in between and trims the last.
    /// Ignored if `end` precedes `start`.
    pub fn delete(&mut self, start: TextCellRef, end: TextCellRef) {
        // Validate positions are in correct order
        if start.col > end.col || (start.col == end.col && start.row > end.row) {
            return;
        }
        let table_clone = self.clone();
        // Handle deletion within a single column
        if start.col == end.col {
            let start_index = start.row;
            let end_index = end.row;

            if let Some(column) = self.columns.get_mut(start.col) {
                // Handle single row deletion
                if start.row == end.row && start_index < column.src.len() {
                    let src = &mut column.src[start_index];
                    let end_offset = std::cmp::min(end.char_offset, src.len());
                    if start.char_offset < end_offset {
                        src.replace_range(start.char_offset..end_offset, "");
                        column.dirty_indices.push(start_index);
                        let updated_src = src.clone();
                        let compiled =
                            crate::core::parser::compile_formula(&updated_src, &[table_clone]);
                        column.compiled_src[start_index] = compiled;
                    }
                }
                // Handle multi-row deletion
                else if start_index < column.len() {
                    // Delete complete rows between start and end
                    if end_index >= start_index {
                        column.drain_rows(start_index..=end_index);
                    }
                }
            }
        } else {
            // Handle multi-column deletion
            for col in start.col..=end.col {
                if let Some(column) = self.columns.get_mut(col) {
                    let start_index = if col == start.col { col } else { 0 };

                    let end_index = if col == end.col {
                        col
                    } else {
                        column.src.len() - 1
                    };

                    if start_index < column.len() {
                        // Delete rows in this column
                        if end_index >= start_index {
                            column.drain_rows(start_index..=end_index);
                        }
                    }
                }
            }
        }
    }

    /// Grows the sheet by one empty row or column on the given side.
    ///
    /// [`Direction::None`] does nothing. Rows are unbounded, but sideways
    /// growth stops once the sheet has 26 columns.
    pub fn extend(&mut self, direction: Direction) {
        if self.columns.is_empty() {
            return;
        }
        let row_count = self.columns[0].src.len();
        const MAX_COLS: usize = 26;
        match direction {
            Direction::Up => {
                for column in &mut self.columns {
                    column.insert_row(0);
                }
                self.uncommitted_actions
                    .push(crate::core::SheetAction::InsertRow {
                        sheet_name: self.name.clone(),
                        index: 0,
                    });
            }
            Direction::Down => {
                for column in &mut self.columns {
                    column.push_row();
                }
                self.uncommitted_actions
                    .push(crate::core::SheetAction::InsertRow {
                        sheet_name: self.name.clone(),
                        index: row_count,
                    });
            }
            Direction::Left => {
                if self.columns.len() < MAX_COLS {
                    self.columns.insert(0, DataColumn::new(row_count));
                    self.uncommitted_actions
                        .push(crate::core::SheetAction::InsertCol {
                            sheet_name: self.name.clone(),
                            index: 0,
                        });
                }
            }
            Direction::Right => {
                if self.columns.len() < MAX_COLS {
                    self.columns.push(DataColumn::new(row_count));
                    self.uncommitted_actions
                        .push(crate::core::SheetAction::InsertCol {
                            sheet_name: self.name.clone(),
                            index: self.columns.len() - 1,
                        });
                }
            }
            Direction::None => {}
        }
    }

    /// Ensure sheet has at least target_row+1 rows and target_col+1 columns
    pub fn ensure_capacity(&mut self, target_row: usize, target_col: usize) {
        let current_rows = self.row_count();
        let needed_rows = target_row + 1;
        let final_rows = current_rows.max(needed_rows);

        while self.columns.len() <= target_col {
            let col_idx = self.columns.len();
            let mut col = DataColumn::new(final_rows);
            col.name = crate::core::parser::col_idx_to_letters(col_idx);
            self.columns.push(col);
        }

        if final_rows > current_rows {
            for col in &mut self.columns {
                col.resize_rows(final_rows);
            }
        }
    }

    /// The style set on a cell, or `None` if it has none.
    ///
    /// This is where a date cell's `num_format` lives -- the notation half of
    /// a date, the value half being the serial in the cell.
    pub fn get_cell_style(&self, row: usize, col: usize) -> Option<&crate::core::CellStyle> {
        self.columns
            .get(col)
            .and_then(|column| column.styles.get(row))
            .and_then(|opt| opt.as_ref())
    }

    /// Replaces a cell's style, growing the sheet if the cell is past its
    /// current bounds. An empty style is stored as no style at all.
    pub fn set_cell_style(&mut self, row: usize, col: usize, style: crate::core::CellStyle) {
        self.ensure_capacity(row, col);
        if let Some(column) = self.columns.get_mut(col)
            && row < column.styles.len()
        {
            if style.is_empty() {
                column.styles[row] = None;
            } else {
                column.styles[row] = Some(style);
            }
        }
    }

    /// Mutates a cell's style in place, starting from the default if it has
    /// none, so one attribute can be changed without disturbing the others.
    ///
    /// Grows the sheet if needed; a style left empty is dropped.
    pub fn update_cell_style<F>(&mut self, row: usize, col: usize, f: F)
    where
        F: FnOnce(&mut crate::core::CellStyle),
    {
        self.ensure_capacity(row, col);
        if let Some(column) = self.columns.get_mut(col)
            && row < column.styles.len()
        {
            let mut current = column.styles[row].clone().unwrap_or_default();
            f(&mut current);
            if current.is_empty() {
                column.styles[row] = None;
            } else {
                column.styles[row] = Some(current);
            }
        }
    }

    /// Removes a cell's style. Unlike the setters, this never grows the sheet.
    pub fn clear_cell_style(&mut self, row: usize, col: usize) {
        if let Some(column) = self.columns.get_mut(col)
            && row < column.styles.len()
        {
            column.styles[row] = None;
        }
    }

    /// Insert a new empty row at the specified index
    /// If index is >= row_count, appends at the end
    pub fn insert_row(&mut self, index: usize) {
        let row_count = self.row_count();
        if index >= row_count {
            // Append at the end
            for column in &mut self.columns {
                column.push_row();
            }
            self.uncommitted_actions
                .push(crate::core::SheetAction::InsertRow {
                    sheet_name: self.name.clone(),
                    index: row_count,
                });
        } else {
            // Insert at the specified index
            for column in &mut self.columns {
                column.insert_row(index);
            }
            self.uncommitted_actions
                .push(crate::core::SheetAction::InsertRow {
                    sheet_name: self.name.clone(),
                    index,
                });
        }
    }

    /// Deletes a row, shifting the rows below it up.
    ///
    /// Removes the entry from all three parallel per-row vectors together,
    /// which is what keeps them the same length, and rebases the dirty queue.
    /// Out-of-range indices are ignored. Everything is marked dirty, since
    /// formulas above the deleted row may refer to it.
    pub fn delete_row(&mut self, index: usize) {
        let row_count = self.row_count();
        if index < row_count {
            for column in &mut self.columns {
                column.remove_row(index);
            }
            self.uncommitted_actions
                .push(crate::core::SheetAction::DeleteRow {
                    sheet_name: self.name.clone(),
                    index,
                });
            self.mark_all_dirty();
        }
    }

    /// Excel's *Insert cells, shift down* over an inclusive column band.
    ///
    /// Unlike [`Sheet::insert_row`] this moves only `first_col..=last_col`,
    /// leaving every other column where it is -- which is what
    /// `ListRows.Add` actually does. Measured: adding a row to a table at
    /// `A1:C4` moves `A8` down to `A9` but leaves `E2` alone.
    ///
    /// Every column keeps the same length: the sheet first grows by `count`
    /// rows, so the rows pushed off the bottom of the band are the blank ones
    /// just added rather than data. Everything moves through `DataColumn`'s
    /// paired operations, so `src` / `data` / `compiled_src` / `styles` stay
    /// aligned.
    ///
    /// Out-of-range bands and a zero `count` are no-ops. Formula references
    /// are *not* rewritten here -- that is
    /// `WorkbookManager::insert_cells_shift_down`'s job, since it spans
    /// sheets.
    pub fn insert_cells_shift_down(
        &mut self,
        row: usize,
        first_col: usize,
        last_col: usize,
        count: usize,
    ) {
        let last_col = last_col.min(self.columns.len().saturating_sub(1));
        if count == 0 || self.columns.is_empty() || first_col > last_col {
            return;
        }
        // Grow every column together first, so the band has somewhere to
        // push into and the sheet stays rectangular throughout.
        for column in &mut self.columns {
            for _ in 0..count {
                column.push_row();
            }
        }
        for column in &mut self.columns[first_col..=last_col] {
            for _ in 0..count {
                column.insert_row(row);
                // Drop the blank row the growth added, so this column ends
                // the same length as the untouched ones.
                column.remove_row(column.len() - 1);
            }
        }
        self.uncommitted_actions
            .push(crate::core::SheetAction::InsertRow {
                sheet_name: self.name.clone(),
                index: row,
            });
        self.mark_all_dirty();
    }

    /// Excel's *Delete cells, shift up* over an inclusive column band; the
    /// inverse of [`Sheet::insert_cells_shift_down`].
    ///
    /// The band's rows below `row` move up and blank rows appear at its
    /// bottom, so the sheet keeps its shape and other columns are untouched.
    pub fn delete_cells_shift_up(
        &mut self,
        row: usize,
        first_col: usize,
        last_col: usize,
        count: usize,
    ) {
        let last_col = last_col.min(self.columns.len().saturating_sub(1));
        if count == 0 || self.columns.is_empty() || first_col > last_col || row >= self.row_count()
        {
            return;
        }
        for column in &mut self.columns[first_col..=last_col] {
            for _ in 0..count {
                if row < column.len() {
                    column.remove_row(row);
                    // Keep the length: the band gains a blank row at the
                    // bottom for each one removed from the middle.
                    column.push_row();
                }
            }
        }
        self.uncommitted_actions
            .push(crate::core::SheetAction::DeleteRow {
                sheet_name: self.name.clone(),
                index: row,
            });
        self.mark_all_dirty();
    }

    /// Deletes a column, shifting the columns to its right left.
    ///
    /// Out-of-range indices are ignored; everything is marked dirty.
    pub fn delete_col(&mut self, index: usize) {
        if index < self.columns.len() {
            self.columns.remove(index);
            self.uncommitted_actions
                .push(crate::core::SheetAction::DeleteCol {
                    sheet_name: self.name.clone(),
                    index,
                });
            self.mark_all_dirty();
        }
    }

    /// Insert a new empty column at the specified index
    /// If index is >= columns.len(), appends at the end
    pub fn insert_col(&mut self, index: usize) {
        let row_count = self.row_count();
        let new_col = DataColumn::new(row_count);
        let col_count = self.columns.len();
        if index >= col_count {
            self.columns.push(new_col);
            self.uncommitted_actions
                .push(crate::core::SheetAction::InsertCol {
                    sheet_name: self.name.clone(),
                    index: col_count,
                });
        } else {
            self.columns.insert(index, new_col);
            self.uncommitted_actions
                .push(crate::core::SheetAction::InsertCol {
                    sheet_name: self.name.clone(),
                    index,
                });
        }
        self.mark_all_dirty();
    }

    /// The sheet's columns.
    ///
    /// Read-only: every column must keep the same number of rows, so growing
    /// or replacing one from outside would desync the sheet. Use
    /// [`Sheet::insert_col`], [`Sheet::delete_col`] and [`Sheet::extend`] to
    /// change the shape.
    pub fn columns(&self) -> &[DataColumn] {
        &self.columns
    }

    /// Allocated rows, taken from the first column -- every column has the
    /// same length.
    pub fn row_count(&self) -> usize {
        self.columns.first().map(|c| c.src.len()).unwrap_or(0)
    }

    /// Allocated columns.
    pub fn col_count(&self) -> usize {
        self.columns.len()
    }
}
