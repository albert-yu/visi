//! Formula text with its references resolved to ids.
//!
//! `parser::compile_formula` turns formula text into a [`CompiledFormula`]:
//! the literal stretches stay text, but every reference becomes a `sheet_id`
//! or `col_id` rather than a name. `parser::serialize_formula` renders it back
//! to A1 text using whatever the names are *now*, which is what makes renaming
//! a sheet, an Excel Table or a table column non-destructive -- nothing has to
//! find and rewrite the formulas that mention it.
//!
//! This is not the evaluation form. Evaluating goes through
//! `parser::parse_excel_formula`, which produces an AST; `Sheet::commit`
//! compiles, re-serializes, and then evaluates the re-serialized text.

use crate::core::RefType;
use serde::{Deserialize, Serialize};

/// Which part of an Excel Table a structured reference selects, as in
/// `Sales[#Headers]`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SheetSection {
    /// The body rows, excluding header and totals. The default.
    Data,
    /// The header row.
    Headers,
    /// The totals row.
    Totals,
    /// Header, data and totals together.
    All,
}

/// One piece of a compiled formula: either literal text or a reference held by
/// id.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum FormulaPart {
    /// A literal stretch of the formula -- operators, function names,
    /// constants -- copied through unchanged.
    Text(String),
    /// A single cell, as in `Sheet2!$A1`.
    SheetReference {
        /// Sheet the cell is on.
        sheet_id: u64,
        /// Row, 0-based.
        row: usize,
        /// Column, 0-based.
        col: usize,
        /// Whether the row was written with a `$`.
        row_ref_type: RefType,
        /// Whether the column was written with a `$`.
        col_ref_type: RefType,
    },
    /// A whole column, held by column id so a column rename survives.
    ColumnReference {
        /// Sheet the column is on.
        sheet_id: u64,
        /// The column's identifier, not its position.
        col_id: u64,
    },
    /// An Excel Table structured reference, as in `Sales[Amount]` or
    /// `[@Amount]`.
    StructuredReference {
        /// Sheet the reference resolves against.
        sheet_id: u64,
        /// The referenced column, or `None` for a whole-table reference.
        col_id: Option<u64>,
        /// `true` for the `[@Amount]` form, which means the current row.
        is_this_row: bool,
        /// Which part of the table is selected.
        section: SheetSection,
    },
    /// A rectangular range, as in `Sheet2!A1:$B$10`.
    RangeReference {
        /// Sheet the range is on.
        sheet_id: u64,
        /// First row, 0-based.
        start_row: usize,
        /// First column, 0-based.
        start_col: usize,
        /// Last row, 0-based and inclusive.
        end_row: usize,
        /// Last column, 0-based and inclusive.
        end_col: usize,
        /// Whether the start row was written with a `$`.
        start_row_ref_type: RefType,
        /// Whether the start column was written with a `$`.
        start_col_ref_type: RefType,
        /// Whether the end row was written with a `$`.
        end_row_ref_type: RefType,
        /// Whether the end column was written with a `$`.
        end_col_ref_type: RefType,
    },
}

/// A formula split into literal text and id-held references.
///
/// Cached per cell in `DataColumn::compiled_src`, and rendered back to A1 text
/// on demand by `parser::serialize_formula`.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CompiledFormula {
    /// The pieces, in the order they appear in the formula text.
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
