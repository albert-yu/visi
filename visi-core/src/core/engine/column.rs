//! Column storage: the typed value column and the per-column bundle of
//! parallel vectors that a `Sheet` is made of.

use crate::core::CompiledFormula;
use crate::core::SharedVec;
use serde::{Deserialize, Serialize};

use super::bitmask::Bitmask;
use super::cell::generate_unique_id;
use super::result_data::ResultData;

/// A column of computed values, stored in whichever representation fits what
/// it currently holds.
///
/// A column starts out as `Integer` and widens as needed: writing a float
/// promotes it to `Float`, and writing anything that is neither demotes it to
/// `Any`. It never narrows back. The two numeric representations keep a
/// separate validity [`Bitmask`] so a blank cell is distinct from a zero.
///
/// This is a storage detail of [`DataColumn`], exposed for reading. The
/// operations that change a column's length are crate-private, since they
/// would desync it from the sibling vectors it must stay aligned with -- go
/// through `Sheet` to edit cells.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum ColumnData {
    /// All-integer, or integer-and-blank.
    Integer {
        /// Which positions hold a value rather than a blank.
        validity: Bitmask,
        /// The values. Positions marked invalid hold a placeholder.
        values: SharedVec<i64>,
    },
    /// Numeric with at least one non-integer, or integer-and-blank promoted.
    Float {
        /// Which positions hold a value rather than a blank.
        validity: Bitmask,
        /// The values. Positions marked invalid hold a placeholder.
        values: SharedVec<f64>,
    },
    /// Mixed: anything the numeric representations cannot hold -- text,
    /// booleans, errors.
    Any(SharedVec<ResultData>),
}

impl ColumnData {
    pub(crate) fn new(size: usize) -> Self {
        Self::Integer {
            validity: Bitmask::with_size(size),
            values: vec![0; size].into(),
        }
    }

    /// How many rows the column holds.
    pub fn len(&self) -> usize {
        match self {
            Self::Integer { validity, .. } => validity.len,
            Self::Float { validity, .. } => validity.len,
            Self::Any(v) => v.len(),
        }
    }

    /// Whether the column holds no rows at all.
    pub fn is_empty(&self) -> bool {
        self.len() == 0
    }

    pub(crate) fn push(&mut self, value: ResultData) {
        let index = self.len();
        self.insert(index, value);
    }

    /// The value at `index`, or `None` if that is past the end.
    ///
    /// A blank within the column's range reads as
    /// `Some(ResultData::None)`, which is what distinguishes it from an
    /// out-of-range index.
    pub fn get(&self, index: usize) -> Option<ResultData> {
        if index >= self.len() {
            return None;
        }
        match self {
            Self::Integer { validity, values } => {
                if validity.get(index) {
                    Some(ResultData::Integer(values[index]))
                } else {
                    Some(ResultData::None)
                }
            }
            Self::Float { validity, values } => {
                if validity.get(index) {
                    Some(ResultData::Float(values[index]))
                } else {
                    Some(ResultData::None)
                }
            }
            Self::Any(v) => Some(v[index].clone()),
        }
    }

    pub(crate) fn demote_to_any(&mut self) {
        let len = self.len();
        let mut any = Vec::with_capacity(len);
        for i in 0..len {
            any.push(self.get(i).unwrap());
        }
        *self = Self::Any(any.into());
    }

    pub(crate) fn promote_to_float(&mut self) {
        if let Self::Integer { validity, values } = self {
            let float_values = values.iter().map(|&i| i as f64).collect();
            *self = Self::Float {
                validity: validity.clone(),
                values: float_values,
            };
        }
    }

    pub(crate) fn resize(&mut self, size: usize) {
        match self {
            Self::Integer { validity, values } => {
                values.resize(size, 0);
                *validity = Bitmask::with_size(size);
            }
            Self::Float { validity, values } => {
                values.resize(size, 0.0);
                *validity = Bitmask::with_size(size);
            }
            Self::Any(v) => {
                v.resize(size, ResultData::None);
            }
        }
    }

    pub(crate) fn set(&mut self, index: usize, value: ResultData) {
        if index >= self.len() {
            return;
        }
        match self {
            Self::Integer { validity, values } => match value {
                ResultData::Integer(i) => {
                    validity.set(index, true);
                    values[index] = i;
                }
                ResultData::Float(f) => {
                    self.promote_to_float();
                    self.set(index, ResultData::Float(f));
                }
                ResultData::None => {
                    validity.set(index, false);
                    values[index] = 0;
                }
                _ => {
                    self.demote_to_any();
                    if let Self::Any(v) = self {
                        v[index] = value;
                    }
                }
            },
            Self::Float { validity, values } => match value {
                ResultData::Float(f) => {
                    validity.set(index, true);
                    values[index] = f;
                }
                ResultData::Integer(i) => {
                    validity.set(index, true);
                    values[index] = i as f64;
                }
                ResultData::None => {
                    validity.set(index, false);
                    values[index] = 0.0;
                }
                _ => {
                    self.demote_to_any();
                    if let Self::Any(v) = self {
                        v[index] = value;
                    }
                }
            },
            Self::Any(v) => {
                v[index] = value;
            }
        }
    }

    pub(crate) fn insert(&mut self, index: usize, value: ResultData) {
        match self {
            Self::Integer { validity, values } => match value {
                ResultData::Integer(i) => {
                    validity.insert(index, true);
                    values.insert(index, i);
                }
                ResultData::Float(f) => {
                    self.promote_to_float();
                    self.insert(index, ResultData::Float(f));
                }
                ResultData::None => {
                    validity.insert(index, false);
                    values.insert(index, 0);
                }
                _ => {
                    self.demote_to_any();
                    if let Self::Any(v) = self {
                        v.insert(index, value);
                    }
                }
            },
            Self::Float { validity, values } => match value {
                ResultData::Float(f) => {
                    validity.insert(index, true);
                    values.insert(index, f);
                }
                ResultData::Integer(i) => {
                    validity.insert(index, true);
                    values.insert(index, i as f64);
                }
                ResultData::None => {
                    validity.insert(index, false);
                    values.insert(index, 0.0);
                }
                _ => {
                    self.demote_to_any();
                    if let Self::Any(v) = self {
                        v.insert(index, value);
                    }
                }
            },
            Self::Any(v) => {
                v.insert(index, value);
            }
        }
    }

    pub(crate) fn remove(&mut self, index: usize) {
        match self {
            Self::Integer { validity, values } => {
                validity.remove(index);
                values.remove(index);
            }
            Self::Float { validity, values } => {
                validity.remove(index);
                values.remove(index);
            }
            Self::Any(v) => {
                v.remove(index);
            }
        }
    }

    pub(crate) fn drain<R: std::ops::RangeBounds<usize> + Clone>(&mut self, range: R) {
        match self {
            Self::Integer { validity, values } => {
                validity.drain(range.clone());
                values.drain(range);
            }
            Self::Float { validity, values } => {
                validity.drain(range.clone());
                values.drain(range);
            }
            Self::Any(v) => {
                v.drain(range);
            }
        }
    }
}

impl Default for ColumnData {
    fn default() -> Self {
        Self::Integer {
            validity: Bitmask::with_size(0),
            values: SharedVec::new(),
        }
    }
}

/// One column of a sheet: the raw text, the computed values, the compiled
/// formulas and the styles, as parallel per-row vectors.
///
/// # Invariant
///
/// `src`, `data`, `compiled_src` and `styles` must all stay the same length --
/// row `r` of the column is entry `r` of each. Nothing enforces this; the
/// row and column insert/delete paths in `Sheet` maintain it by hand, and
/// `Sheet::setup_after_deserialization` restores it after a load, since only
/// `src` and `styles` are persisted. Mutating one of these vectors directly
/// will break it.
///
/// `dirty_indices` is not part of that invariant -- it is a queue of rows
/// awaiting recomputation, and is emptied by `Sheet::commit`.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DataColumn {
    /// Identifier, stable across renames and repositioning. Compiled formulas
    /// refer to a column by this rather than by name or position.
    #[serde(default = "generate_unique_id")]
    pub id: u64,
    /// Display name, empty unless one was set.
    #[serde(default)]
    pub name: String,
    /// The computed values. Rebuilt on load, so not persisted.
    #[serde(skip, default)]
    pub data: ColumnData,
    /// The raw text of each cell, exactly as typed. The only representation
    /// that is persisted, and the one everything else is rebuilt from.
    pub src: SharedVec<String>,
    /// Cached compile output for each cell. Rebuilt on load.
    #[serde(skip, default)]
    pub compiled_src: SharedVec<CompiledFormula>,
    /// Rows awaiting recomputation. Drained by `Sheet::commit`.
    #[serde(skip, default)]
    pub dirty_indices: SharedVec<usize>,
    /// Per-cell styling, `None` where a cell has none. Carries a date cell's
    /// number format.
    #[serde(default)]
    pub styles: SharedVec<Option<crate::core::CellStyle>>,
}

pub(crate) struct ColumnPosition {
    pub row: usize,
    pub char_offset: usize,
}

impl DataColumn {
    /// A column of `size` empty rows, with every parallel vector sized to
    /// match and a freshly generated id.
    pub fn new(size: usize) -> Self {
        Self {
            id: generate_unique_id(),
            name: String::new(),
            data: ColumnData::new(size),
            src: vec![String::new(); size].into(),
            compiled_src: vec![CompiledFormula::default(); size].into(),
            dirty_indices: SharedVec::new(),
            styles: vec![None; size].into(),
        }
    }

    pub(crate) fn mark_dirty(&mut self, row: usize) {
        if !self.dirty_indices.contains(&row) {
            self.dirty_indices.push(row);
        }
    }

    /// Row is absolutely referenced
    pub(crate) fn insert(&mut self, position: ColumnPosition, input: &str) {
        let ColumnPosition { row, char_offset } = position;
        let index = row;
        if index < self.src.len() {
            if self.src[index].is_empty() {
                self.src[index].push_str(input);
            } else {
                self.src[index].insert_str(char_offset, input);
            }
        } else if index == self.src.len() {
            self.src.push(input.to_string());
            self.compiled_src.push(CompiledFormula::default());
            self.data.push(ResultData::None);
            self.styles.push(None);
        } else {
            let mut i = self.src.len();
            while i < index {
                self.src.push(String::new());
                self.compiled_src.push(CompiledFormula::default());
                self.data.push(ResultData::None);
                self.styles.push(None);
                i += 1;
            }
            self.src.push(input.to_string());
            self.compiled_src.push(CompiledFormula::default());
            self.data.push(ResultData::None);
            self.styles.push(None);
        }
    }
}
