use serde::{Deserialize, Serialize};

/// [LLM-generated] A random 53-bit identifier for a sheet or column.
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
    val & 0x001F_FFFF_FFFF_FFFF
}

/// [LLM-generated] The intrinsic data type of a cell, mirroring calamine worksheet value variants.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize, Default)]
pub enum CellType {
    /// [LLM-generated] Empty cell (`calamine::Data::Empty`).
    #[default]
    Empty,
    /// [LLM-generated] Signed integer (`calamine::Data::Int`).
    Int,
    /// [LLM-generated] Floating point number (`calamine::Data::Float`).
    Float,
    /// [LLM-generated] String (`calamine::Data::String`).
    String,
    /// [LLM-generated] Boolean (`calamine::Data::Bool`).
    Bool,
    /// [LLM-generated] Date/time serial identified by calamine from workbook formatting.
    DateTime,
    /// [LLM-generated] ISO 8601 date/time (`calamine::Data::DateTimeIso`, OpenXML `t="d"`).
    DateTimeIso,
    /// [LLM-generated] ISO 8601 duration (`calamine::Data::DurationIso`).
    DurationIso,
    /// [LLM-generated] Error cell (`calamine::Data::Error`, OpenXML `t="e"`).
    Error,
}

impl CellType {
    /// [LLM-generated] Whether this cell type is explicitly a string/text cell.
    pub fn is_string(&self) -> bool {
        matches!(self, CellType::String)
    }

    /// [LLM-generated] Stable lowercase spelling used by CLI and JSON output.
    pub fn as_str(&self) -> &'static str {
        match self {
            CellType::Empty => "empty",
            CellType::Int => "int",
            CellType::Float => "float",
            CellType::String => "string",
            CellType::Bool => "bool",
            CellType::DateTime => "date-time",
            CellType::DateTimeIso => "date-time-iso",
            CellType::DurationIso => "duration-iso",
            CellType::Error => "error",
        }
    }
}

/// [LLM-generated] For either a column or row
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub enum RefType {
    /// [LLM-generated] Written without a `$`, so it shifts when the formula is filled or
    /// copied.
    Relative,
    /// [LLM-generated] Written with a `$`, so it stays put.
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

/// [LLM-generated] A cell's position, plus whether it was written as absolute.
///
/// Coordinates are 0-based, as everywhere inside the engine; `A1` is
/// `CellRef::new(0, 0)`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub struct CellRef {
    /// [LLM-generated] Row index, 0-based.
    pub row: usize,
    /// [LLM-generated] Column index, 0-based.
    pub col: usize,
    /// [LLM-generated] Whether the row was written with a `$`.
    pub row_ref_type: RefType,
    /// [LLM-generated] Whether the column was written with a `$`.
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
    /// [LLM-generated] A relative reference to `(row, col)`, 0-based.
    pub fn new(row: usize, col: usize) -> CellRef {
        Self {
            row,
            col,
            row_ref_type: RefType::Relative,
            col_ref_type: RefType::Relative,
        }
    }
}

/// [LLM-generated] Something a formula reads, and therefore an edge in the recalculation
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
    /// [LLM-generated] A cell on the same sheet.
    Local(CellRef),
    /// [LLM-generated] A whole column on the same sheet, by 0-based position.
    LocalColumn(usize),
    /// [LLM-generated] A cell on another sheet.
    Remote {
        /// [LLM-generated] Name of the sheet the cell is on.
        sheet: String,
        /// [LLM-generated] The cell, on that sheet.
        cell: CellRef,
    },
    /// [LLM-generated] A whole column on another sheet.
    RemoteColumn {
        /// [LLM-generated] Name of the sheet the column is on.
        sheet: String,
        /// [LLM-generated] Column index, 0-based.
        col: usize,
    },
}

/// [LLM-generated] A caret position: a cell plus an offset within its source text, for the
/// text-editing operations `Sheet::insert` and `Sheet::delete`.
#[derive(Debug, Clone, Default)]
pub struct TextCellRef {
    /// [LLM-generated] Row index, 0-based.
    pub row: usize,
    /// [LLM-generated] Column index, 0-based.
    pub col: usize,
    /// [LLM-generated] Offset into the cell's source text, in characters rather than bytes.
    pub char_offset: usize,
}

/// [LLM-generated] A formula that could not be evaluated at all.
///
/// Distinct from an Excel error value: `=1/0` evaluates successfully to
/// `ResultData::Error("#DIV/0!")`, whereas this is for text that never became
/// a computable formula.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum EvalError {
    /// [LLM-generated] The formula could not be parsed, or named something unrecognized. The
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

/// [LLM-generated] What the engine's evaluation entry points return on failure.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum EngineError {
    /// [LLM-generated] A formula could not be evaluated.
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
