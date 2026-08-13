use crate::core::RefType;
use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SheetSection {
    Data,
    Headers,
    Totals,
    All,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum FormulaPart {
    Text(String),
    SheetReference {
        sheet_id: u64,
        row: usize,
        col: usize,
        row_ref_type: RefType,
        col_ref_type: RefType,
    },
    ColumnReference {
        sheet_id: u64,
        col_id: u64,
    },
    StructuredReference {
        sheet_id: u64,
        col_id: Option<u64>,
        is_this_row: bool,
        section: SheetSection,
    },
    RangeReference {
        sheet_id: u64,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
        start_row_ref_type: RefType,
        start_col_ref_type: RefType,
        end_row_ref_type: RefType,
        end_col_ref_type: RefType,
    },
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CompiledFormula {
    pub parts: Vec<FormulaPart>,
}

impl CompiledFormula {
    /// Creates a plain formula from a raw string, without any parsed references.
    /// Useful as a default constructor or fallback.
    pub fn plain(text: String) -> Self {
        Self {
            parts: vec![FormulaPart::Text(text)],
        }
    }

    /// Checks if the formula is empty
    #[allow(dead_code)]
    pub fn is_empty(&self) -> bool {
        self.parts.is_empty()
            || (self.parts.len() == 1
                && match &self.parts[0] {
                    FormulaPart::Text(s) => s.is_empty(),
                    _ => false,
                })
    }
}

impl Default for CompiledFormula {
    fn default() -> Self {
        Self::plain(String::new())
    }
}
