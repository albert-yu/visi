use super::super::column::{ColumnPosition, DataColumn};
use super::{CellRef, CellType, Direction, ResultData, Sheet, TextCellRef};

/// The word surrounding `char_offset` in `text`. The idea
/// is to implement "highlight word on double click".
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
                    (Some(fmt), None) | (None, Some(fmt)) => Some(fmt),
                    _ => None,
                }
            }
            _ => None,
        }
    }

    /// Interpolate the cell's value into a string, with
    /// a specified cell format
    pub fn get_display_string(&self, cell: &CellRef) -> String {
        let value = self.get_result_data(cell);
        let Some(code) = self
            .get_cell_style(cell.row, cell.col)
            .and_then(|s| s.num_format.as_deref())
        else {
            return value.to_string();
        };
        match value {
            ResultData::Float(f) => crate::core::text::format_number_format(f, code)
                .unwrap_or_else(|_| ResultData::Float(f).to_string()),
            ResultData::Integer(i) => crate::core::text::format_number_format(i as f64, code)
                .unwrap_or_else(|_| ResultData::Integer(i).to_string()),
            ResultData::String(s) => crate::core::text::format_text_format(&s, code)
                .unwrap_or(ResultData::String(s).to_string()),
            other => other.to_string(),
        }
    }

    /// Returns the intrinsic data type of a cell.
    pub fn get_cell_type(&self, cell: &CellRef) -> CellType {
        let col = self.columns.get(cell.col);
        if let Some(col) = col {
            col.cell_types
                .get(cell.row)
                .copied()
                .unwrap_or(CellType::Empty)
        } else {
            CellType::Empty
        }
    }

    /// Sets the intrinsic data type of a cell at (row, col).
    pub fn set_cell_type(&mut self, row: usize, col: usize, cell_type: CellType) {
        if let Some(column) = self.columns.get_mut(col)
            && row < column.cell_types.len()
        {
            column.cell_types[row] = cell_type;
            column.mark_dirty(row);
        }
    }

    /// Sets the source text and explicit cell type of a particular cell.
    pub fn set_cell_with_type(&mut self, row: usize, col: usize, src: String, cell_type: CellType) {
        let table_clone = self.clone();
        if let Some(column) = self.columns.get_mut(col)
            && row < column.src.len()
        {
            column.src[row] = src.clone();
            column.cell_types[row] = cell_type;
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

    /// Directly sets the src of a cell and marks it dirty.
    /// Does not automatically evaluate it.
    /// Call [`Sheet::commit`] to evaluate
    /// updated cells.
    pub fn set_cell_src(&mut self, row: usize, col: usize, src: String) {
        let table_clone = self.clone();
        if let Some(column) = self.columns.get_mut(col)
            && row < column.src.len()
        {
            column.src[row] = src.clone();
            column.cell_types[row] = CellType::Empty;
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
    pub fn delete(&mut self, start: TextCellRef, end: TextCellRef) {
        if start.col > end.col || (start.col == end.col && start.row > end.row) {
            return;
        }
        let table_clone = self.clone();
        if start.col == end.col {
            let start_index = start.row;
            let end_index = end.row;

            if let Some(column) = self.columns.get_mut(start.col) {
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
                } else if start_index < column.len() && end_index >= start_index {
                    column.drain_rows(start_index..=end_index);
                }
            }
        } else {
            for col in start.col..=end.col {
                if let Some(column) = self.columns.get_mut(col) {
                    let start_index = if col == start.col { col } else { 0 };

                    let end_index = if col == end.col {
                        col
                    } else {
                        column.src.len() - 1
                    };

                    if start_index < column.len() && end_index >= start_index {
                        column.drain_rows(start_index..=end_index);
                    }
                }
            }
        }
    }

    /// Grows the sheet by one empty row or column on the given side.
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
                self.row_heights.insert(0, None);
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
                self.row_heights.push(None);
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
            self.row_heights.resize(final_rows, None);
        }
    }

    /// The style set on a cell, or `None` if it has none.
    pub fn get_cell_style(&self, row: usize, col: usize) -> Option<&crate::core::CellStyle> {
        self.columns
            .get(col)
            .and_then(|column| column.styles.get(row))
            .and_then(|opt| opt.as_ref())
    }

    /// Replaces a cell's style, growing the sheet if the cell is past its
    /// current bounds
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

    /// Mutates a cell's style in place.
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

    /// Removes a cell's style
    pub fn clear_cell_style(&mut self, row: usize, col: usize) {
        if let Some(column) = self.columns.get_mut(col)
            && row < column.styles.len()
        {
            column.styles[row] = None;
        }
    }

    /// Returns the custom width for the zero-based column, if one is set.
    pub fn get_column_width(&self, col: usize) -> Option<f64> {
        self.columns.get(col).and_then(|column| column.width)
    }

    /// Sets or clears the custom width for an existing zero-based column.
    pub fn set_column_width(&mut self, col: usize, width: Option<f64>) {
        if let Some(column) = self.columns.get_mut(col) {
            column.width = width.filter(|value| value.is_finite() && *value >= 0.0);
        }
    }

    /// Returns the custom height for the zero-based row, if one is set.
    pub fn get_row_height(&self, row: usize) -> Option<f64> {
        self.row_heights.get(row).and_then(|height| *height)
    }

    /// Sets or clears the custom height for an existing zero-based row.
    pub fn set_row_height(&mut self, row: usize, height: Option<f64>) {
        if row >= self.row_count() {
            return;
        }
        let row_count = self.row_count();
        self.row_heights.resize(row_count, None);
        self.row_heights[row] = height.filter(|value| value.is_finite() && *value >= 0.0);
    }

    /// Insert a new empty row at the specified index.
    /// If index is >= row_count, appends at the end
    pub fn insert_row(&mut self, index: usize) {
        let row_count = self.row_count();
        self.row_heights.resize(row_count, None);
        if index >= row_count {
            for column in &mut self.columns {
                column.push_row();
            }
            self.row_heights.push(None);
            self.uncommitted_actions
                .push(crate::core::SheetAction::InsertRow {
                    sheet_name: self.name.clone(),
                    index: row_count,
                });
        } else {
            for column in &mut self.columns {
                column.insert_row(index);
            }
            self.row_heights.insert(index, None);
            self.uncommitted_actions
                .push(crate::core::SheetAction::InsertRow {
                    sheet_name: self.name.clone(),
                    index,
                });
        }
    }

    /// Deletes a row, shifting the rows below it up.
    pub fn delete_row(&mut self, index: usize) {
        let row_count = self.row_count();
        if index < row_count {
            for column in &mut self.columns {
                column.remove_row(index);
            }
            if index < self.row_heights.len() {
                self.row_heights.remove(index);
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
        for column in &mut self.columns {
            for _ in 0..count {
                column.push_row();
            }
        }
        let row_count = self.row_count();
        self.row_heights.resize(row_count, None);
        for column in &mut self.columns[first_col..=last_col] {
            for _ in 0..count {
                column.insert_row(row);
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

    /// The sheet's columns
    pub fn columns(&self) -> &[DataColumn] {
        &self.columns
    }

    /// Allocated rows
    pub fn row_count(&self) -> usize {
        self.columns.first().map(|c| c.src.len()).unwrap_or(0)
    }

    /// Allocated columns
    pub fn col_count(&self) -> usize {
        self.columns.len()
    }
}
