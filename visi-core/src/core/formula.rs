use crate::core::RefType;
use serde::{Deserialize, Serialize};

/// [LLM-generated] Which part of an Excel Table a structured reference selects, as in
/// `Sales[#Headers]`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SheetSection {
    /// [LLM-generated] The body rows, excluding header and totals. The default.
    Data,
    /// [LLM-generated] The header row.
    Headers,
    /// [LLM-generated] The totals row.
    Totals,
    /// [LLM-generated] Header, data and totals together.
    All,
}

/// [LLM-generated] One piece of a compiled formula: either literal text or a reference held by
/// id.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum FormulaPart {
    /// [LLM-generated] A literal stretch of the formula -- operators, function names,
    /// constants -- copied through unchanged.
    Text(String),
    /// [LLM-generated] A single cell, as in `Sheet2!$A1`.
    SheetReference {
        /// [LLM-generated] Sheet the cell is on.
        sheet_id: u64,
        /// [LLM-generated] Row, 0-based.
        row: usize,
        /// [LLM-generated] Column, 0-based.
        col: usize,
        /// [LLM-generated] Whether the row was written with a `$`.
        row_ref_type: RefType,
        /// [LLM-generated] Whether the column was written with a `$`.
        col_ref_type: RefType,
    },
    /// [LLM-generated] A whole column, held by column id so a column rename survives.
    ColumnReference {
        /// [LLM-generated] Sheet the column is on.
        sheet_id: u64,
        /// [LLM-generated] The column's identifier, not its position.
        col_id: u64,
    },
    /// [LLM-generated] An Excel Table structured reference, as in `Sales[Amount]` or
    /// `[@Amount]`.
    StructuredReference {
        /// [LLM-generated] Sheet the reference resolves against.
        sheet_id: u64,
        /// [LLM-generated] The referenced column, or `None` for a whole-table reference.
        col_id: Option<u64>,
        /// [LLM-generated] `true` for the `[@Amount]` form, which means the current row.
        is_this_row: bool,
        /// [LLM-generated] Which part of the table is selected.
        section: SheetSection,
    },
    /// [LLM-generated] A rectangular range, as in `Sheet2!A1:$B$10`.
    RangeReference {
        /// [LLM-generated] Sheet the range is on.
        sheet_id: u64,
        /// [LLM-generated] First row, 0-based.
        start_row: usize,
        /// [LLM-generated] First column, 0-based.
        start_col: usize,
        /// [LLM-generated] Last row, 0-based and inclusive.
        end_row: usize,
        /// [LLM-generated] Last column, 0-based and inclusive.
        end_col: usize,
        /// [LLM-generated] Whether the start row was written with a `$`.
        start_row_ref_type: RefType,
        /// [LLM-generated] Whether the start column was written with a `$`.
        start_col_ref_type: RefType,
        /// [LLM-generated] Whether the end row was written with a `$`.
        end_row_ref_type: RefType,
        /// [LLM-generated] Whether the end column was written with a `$`.
        end_col_ref_type: RefType,
    },
}

/// [LLM-generated] A formula split into literal text and id-held references.
///
/// Cached per cell in `DataColumn::compiled_src`, and rendered back to A1 text
/// on demand by `parser::serialize_formula`.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CompiledFormula {
    /// [LLM-generated] The pieces, in the order they appear in the formula text.
    pub parts: Vec<FormulaPart>,
}

impl CompiledFormula {
    /// [LLM-generated] Creates a plain formula from a raw string, without any parsed references.
    /// Useful as a default constructor or fallback.
    pub fn plain(text: String) -> Self {
        Self {
            parts: vec![FormulaPart::Text(text)],
        }
    }

    /// [LLM-generated] Checks if the formula is empty
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
