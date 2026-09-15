use serde::{Deserialize, Serialize};

/// [LLM-generated] A single edit made to a workbook, recorded so a host can observe or replay
/// it.
///
/// `Sheet` appends one of these to `Sheet::uncommitted_actions` for each edit.
/// Sheets are named rather than held by id, since an action is meant to
/// survive being written down and applied elsewhere.
///
/// "Table" in the variant names means a *sheet*, following this codebase's
/// older informal naming -- not an `ExcelTable`.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum SheetAction {
    /// [LLM-generated] A cell's raw text was replaced.
    SetCellSrc {
        /// [LLM-generated] Sheet the cell is on.
        sheet_name: String,
        /// [LLM-generated] Column index, 0-based.
        col: usize,
        /// [LLM-generated] Row index, 0-based.
        row: usize,
        /// [LLM-generated] The new text.
        src: String,
    },
    /// [LLM-generated] A column was renamed.
    UpdateColName {
        /// [LLM-generated] Sheet the column is on.
        sheet_name: String,
        /// [LLM-generated] Column index, 0-based.
        col: usize,
        /// [LLM-generated] The new name.
        name: String,
    },
    /// [LLM-generated] A sheet was renamed.
    UpdateTableName {
        /// [LLM-generated] The sheet's previous name.
        old_name: String,
        /// [LLM-generated] Its new name.
        new_name: String,
    },
    /// [LLM-generated] An empty row was inserted.
    InsertRow {
        /// [LLM-generated] Sheet the row was inserted on.
        sheet_name: String,
        /// [LLM-generated] Position it was inserted at, 0-based.
        index: usize,
    },
    /// [LLM-generated] A row was deleted.
    DeleteRow {
        /// [LLM-generated] Sheet the row was deleted from.
        sheet_name: String,
        /// [LLM-generated] Position it occupied, 0-based.
        index: usize,
    },
    /// [LLM-generated] An empty column was inserted.
    InsertCol {
        /// [LLM-generated] Sheet the column was inserted on.
        sheet_name: String,
        /// [LLM-generated] Position it was inserted at, 0-based.
        index: usize,
    },
    /// [LLM-generated] A column was deleted.
    DeleteCol {
        /// [LLM-generated] Sheet the column was deleted from.
        sheet_name: String,
        /// [LLM-generated] Position it occupied, 0-based.
        index: usize,
    },
    /// [LLM-generated] A sheet was added to the workbook.
    AddTable {
        /// [LLM-generated] The sheet, in full.
        sheet: Box<crate::core::Sheet>,
    },
    /// [LLM-generated] A sheet was removed from the workbook.
    DeleteTable {
        /// [LLM-generated] Name of the sheet that was removed.
        sheet_name: String,
    },
}
