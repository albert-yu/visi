#[doc = "Compact boolean bitmask storage used by engine data structures."]
pub mod bitmask;
#[doc = "Cell references, value types, dependency metadata, and engine errors."]
pub mod cell;
#[doc = "Column-oriented storage for sheet data."]
pub mod column;
#[doc = "Evaluation result values produced by formulas and cells."]
pub mod result_data;
#[doc = "Sheets, workbook context, and sheet-level evaluation helpers."]
pub mod sheet;

#[cfg(test)]
pub(crate) mod tests;

pub use bitmask::Bitmask;
pub use cell::{
    CellRef, CellType, Dependency, EngineError, EvalError, RefType, TextCellRef, generate_unique_id,
};
pub use column::{ColumnData, DataColumn};
pub use result_data::ResultData;
pub use sheet::{Context, Direction, Sheet, SheetInit, get_word_boundaries_from_str};
