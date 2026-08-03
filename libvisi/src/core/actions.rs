use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum SheetAction {
    SetCellSrc {
        sheet_name: String,
        col: usize,
        row: usize,
        src: String,
    },
    UpdateColName {
        sheet_name: String,
        col: usize,
        name: String,
    },
    UpdateTableName {
        old_name: String,
        new_name: String,
    },
    InsertRow {
        sheet_name: String,
        index: usize,
    },
    DeleteRow {
        sheet_name: String,
        index: usize,
    },
    InsertCol {
        sheet_name: String,
        index: usize,
    },
    DeleteCol {
        sheet_name: String,
        index: usize,
    },
    AddTable {
        sheet: crate::core::Sheet,
    },
    DeleteTable {
        sheet_name: String,
    },
}
