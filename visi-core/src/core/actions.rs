//! The edit log a `Sheet` records as it is changed.

use serde::{Deserialize, Serialize};

/// A single edit made to a workbook, recorded so a host can observe or replay
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
    /// A cell's raw text was replaced.
    SetCellSrc {
        /// Sheet the cell is on.
        sheet_name: String,
        /// Column index, 0-based.
        col: usize,
        /// Row index, 0-based.
        row: usize,
        /// The new text.
        src: String,
    },
    /// A column was renamed.
    UpdateColName {
        /// Sheet the column is on.
        sheet_name: String,
        /// Column index, 0-based.
        col: usize,
        /// The new name.
        name: String,
    },
    /// A sheet was renamed.
    UpdateTableName {
        /// The sheet's previous name.
        old_name: String,
        /// Its new name.
        new_name: String,
    },
    /// An empty row was inserted.
    InsertRow {
        /// Sheet the row was inserted on.
        sheet_name: String,
        /// Position it was inserted at, 0-based.
        index: usize,
    },
    /// A row was deleted.
    DeleteRow {
        /// Sheet the row was deleted from.
        sheet_name: String,
        /// Position it occupied, 0-based.
        index: usize,
    },
    /// An empty column was inserted.
    InsertCol {
        /// Sheet the column was inserted on.
        sheet_name: String,
        /// Position it was inserted at, 0-based.
        index: usize,
    },
    /// A column was deleted.
    DeleteCol {
        /// Sheet the column was deleted from.
        sheet_name: String,
        /// Position it occupied, 0-based.
        index: usize,
    },
    /// A sheet was added to the workbook.
    AddTable {
        /// The sheet, in full.
        sheet: crate::core::Sheet,
    },
    /// A sheet was removed from the workbook.
    DeleteTable {
        /// Name of the sheet that was removed.
        sheet_name: String,
    },
}
