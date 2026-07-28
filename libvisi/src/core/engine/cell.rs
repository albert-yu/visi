use serde::{Deserialize, Serialize};

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
    Relative,
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

/// Cell reference for indexing
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub struct CellRef {
    pub row: usize,
    pub col: usize,
    pub row_ref_type: RefType,
    pub col_ref_type: RefType,
}

impl std::fmt::Display for CellRef {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(
            f,
            "CELL({}{}, {}{})",
            self.row_ref_type.to_string(),
            self.row,
            self.col_ref_type.to_string(),
            self.col
        )
    }
}

impl CellRef {
    pub fn new(row: usize, col: usize) -> CellRef {
        Self {
            row,
            col,
            row_ref_type: RefType::Relative,
            col_ref_type: RefType::Relative,
        }
    }
}

// Cell padding constants (must match ui.rs)
pub const PADDING_X: f32 = 0.004; // Horizontal padding in world coordinates
pub const PADDING_Y: f32 = 0.003; // Vertical padding in world coordinates

#[derive(Debug, Clone, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub enum Dependency {
    Local(CellRef),
    LocalColumn(usize),
    Remote { sheet: String, cell: CellRef },
    RemoteColumn { sheet: String, col: usize },
}

#[derive(Debug, Clone)]
pub struct TextCellRef {
    pub row: usize,
    pub col: usize,
    pub char_offset: usize,
}

impl Default for TextCellRef {
    fn default() -> Self {
        Self {
            row: 0,
            col: 0,
            char_offset: 0,
        }
    }
}

#[derive(Debug)]
pub enum EvalError {
    UnknownFunction(String),
}

impl std::fmt::Display for EvalError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            EvalError::UnknownFunction(func) => write!(f, "Unknown function: {}", func),
        }
    }
}

#[derive(Debug)]
pub enum EngineError {
    EvalError(EvalError),
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
