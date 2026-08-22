//! The spreadsheet engine: sheets, cells, columns and values.

pub mod bitmask;
pub mod cell;
pub mod column;
pub mod result_data;
pub mod sheet;

#[cfg(test)]
pub(crate) mod tests;

pub use bitmask::Bitmask;
pub use cell::{
    CellRef, Dependency, EngineError, EvalError, RefType, TextCellRef, generate_unique_id,
};
pub use column::{ColumnData, DataColumn};
pub use result_data::ResultData;
pub use sheet::{Context, Direction, Sheet, SheetInit, get_word_boundaries_from_str};
