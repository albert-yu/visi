//! Cell coordinates, dependency edges, and the engine's error types.

use serde::{Deserialize, Serialize};

/// A random 53-bit identifier for a sheet or column.
///
/// Capped to `2^53 - 1` so it survives a round trip through a JSON number,
/// which is what a JavaScript host would deserialize it as. Falls back to the
/// wall clock if the system random source is unavailable.
pub fn generate_unique_id() -> u64 {
    let mut buf = [0u8; 8];
    let val = if getrandom::getrandom(&mut buf).is_err() {
        let now = web_time::SystemTime::now()
            .duration_since(web_time::SystemTime::UNIX_EPOCH)
            .map(|d| d.as_nanos())
            .unwrap_or(0);
        now as u64
    } else {
        u64::from_le_bytes(buf)
    };
    // Cap to JS Number.MAX_SAFE_INTEGER (2^53 - 1) to prevent serialization precision loss
    val & 0x001F_FFFF_FFFF_FFFF
}

/// For either a column or row
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub enum RefType {
    /// Written without a `$`, so it shifts when the formula is filled or
    /// copied.
    Relative,
    /// Written with a `$`, so it stays put.
    Absolute,
}

impl std::fmt::Display for RefType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            RefType::Relative => write!(f, ""),
            RefType::Absolute => write!(f, "$"),
        }
    }
}

/// A cell's position, plus whether it was written as absolute.
///
/// Coordinates are 0-based, as everywhere inside the engine; `A1` is
/// `CellRef::new(0, 0)`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub struct CellRef {
    /// Row index, 0-based.
    pub row: usize,
    /// Column index, 0-based.
    pub col: usize,
    /// Whether the row was written with a `$`.
    pub row_ref_type: RefType,
    /// Whether the column was written with a `$`.
    pub col_ref_type: RefType,
}

impl std::fmt::Display for CellRef {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(
            f,
            "CELL({}{}, {}{})",
            self.row_ref_type, self.row, self.col_ref_type, self.col
        )
    }
}

impl CellRef {
    /// A relative reference to `(row, col)`, 0-based.
    pub fn new(row: usize, col: usize) -> CellRef {
        Self {
            row,
            col,
            row_ref_type: RefType::Relative,
            col_ref_type: RefType::Relative,
        }
    }
}

/// Something a formula reads, and therefore an edge in the recalculation
/// graph.
///
/// The local/remote split is load-bearing. `Sheet::commit` propagates through
/// the `Local` variants only -- a sheet cannot reach into its neighbors, so a
/// remote edge it finds is recorded but not followed. Chasing those is
/// `WorkbookManager::evaluate`'s job, which marks every sheet dirty and runs a
/// fixed number of passes over the workbook; a cross-sheet chain deeper than
/// that number of hops will not have converged when it stops.
///
/// Remote variants key on the sheet *name* rather than its id, since that is
/// what a formula's text carries and what `Context` is indexed by.
#[derive(Debug, Clone, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub enum Dependency {
    /// A cell on the same sheet.
    Local(CellRef),
    /// A whole column on the same sheet, by 0-based position.
    LocalColumn(usize),
    /// A cell on another sheet.
    Remote {
        /// Name of the sheet the cell is on.
        sheet: String,
        /// The cell, on that sheet.
        cell: CellRef,
    },
    /// A whole column on another sheet.
    RemoteColumn {
        /// Name of the sheet the column is on.
        sheet: String,
        /// Column index, 0-based.
        col: usize,
    },
}

/// A caret position: a cell plus an offset within its source text, for the
/// text-editing operations `Sheet::insert` and `Sheet::delete`.
#[derive(Debug, Clone, Default)]
pub struct TextCellRef {
    /// Row index, 0-based.
    pub row: usize,
    /// Column index, 0-based.
    pub col: usize,
    /// Offset into the cell's source text, in characters rather than bytes.
    pub char_offset: usize,
}

/// A formula that could not be evaluated at all.
///
/// Distinct from an Excel error value: `=1/0` evaluates successfully to
/// `ResultData::Error("#DIV/0!")`, whereas this is for text that never became
/// a computable formula.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum EvalError {
    /// The formula could not be parsed, or named something unrecognized. The
    /// string is the message, which for some failures is an Excel error code.
    UnknownFunction(String),
}

impl std::error::Error for EvalError {}

impl std::fmt::Display for EvalError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            EvalError::UnknownFunction(func) => write!(f, "{}", func),
        }
    }
}

/// What the engine's evaluation entry points return on failure.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum EngineError {
    /// A formula could not be evaluated.
    EvalError(EvalError),
}

impl std::error::Error for EngineError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            EngineError::EvalError(err) => Some(err),
        }
    }
}

impl std::fmt::Display for EngineError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            EngineError::EvalError(err) => write!(f, "{}", err),
        }
    }
}

impl From<EvalError> for EngineError {
    fn from(err: EvalError) -> Self {
        EngineError::EvalError(err)
    }
}
