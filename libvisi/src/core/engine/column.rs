use crate::core::CompiledFormula;
use crate::core::SharedVec;
use serde::{Deserialize, Serialize};

use super::bitmask::Bitmask;
use super::cell::generate_unique_id;
use super::result_data::ResultData;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum ColumnData {
    Integer {
        validity: Bitmask,
        values: SharedVec<i64>,
    },
    Float {
        validity: Bitmask,
        values: SharedVec<f64>,
    },
    Any(SharedVec<ResultData>),
}

impl ColumnData {
    pub fn new(size: usize) -> Self {
        Self::Integer {
            validity: Bitmask::with_size(size),
            values: vec![0; size].into(),
        }
    }

    pub fn len(&self) -> usize {
        match self {
            Self::Integer { validity, .. } => validity.len,
            Self::Float { validity, .. } => validity.len,
            Self::Any(v) => v.len(),
        }
    }

    pub fn push(&mut self, value: ResultData) {
        let index = self.len();
        self.insert(index, value);
    }

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

    pub fn demote_to_any(&mut self) {
        let len = self.len();
        let mut any = Vec::with_capacity(len);
        for i in 0..len {
            any.push(self.get(i).unwrap());
        }
        *self = Self::Any(any.into());
    }

    pub fn promote_to_float(&mut self) {
        if let Self::Integer { validity, values } = self {
            let float_values = values.iter().map(|&i| i as f64).collect();
            *self = Self::Float {
                validity: validity.clone(),
                values: float_values,
            };
        }
    }

    pub fn resize(&mut self, size: usize) {
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

    pub fn set(&mut self, index: usize, value: ResultData) {
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

    pub fn insert(&mut self, index: usize, value: ResultData) {
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

    pub fn remove(&mut self, index: usize) {
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

    pub fn drain<R: std::ops::RangeBounds<usize> + Clone>(&mut self, range: R) {
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

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DataColumn {
    #[serde(default = "generate_unique_id")]
    pub id: u64,
    #[serde(default)]
    pub name: String,
    #[serde(skip, default)]
    pub data: ColumnData,
    /// 1-1 to data
    pub src: SharedVec<String>,
    #[serde(skip, default)]
    pub compiled_src: SharedVec<CompiledFormula>,
    #[serde(skip, default)]
    pub dirty_indices: SharedVec<usize>,
}

pub struct ColumnPosition {
    pub row: usize,
    pub char_offset: usize,
}

impl DataColumn {
    pub fn new(size: usize) -> Self {
        Self {
            id: generate_unique_id(),
            name: String::new(),
            data: ColumnData::new(size),
            src: vec![String::new(); size].into(),
            compiled_src: vec![CompiledFormula::default(); size].into(),
            dirty_indices: SharedVec::new(),
        }
    }

    pub fn mark_dirty(&mut self, row: usize) {
        if !self.dirty_indices.contains(&row) {
            self.dirty_indices.push(row);
        }
    }

    /// Row is absolutely referenced
    pub fn insert(&mut self, position: ColumnPosition, input: &str) {
        let ColumnPosition { row, char_offset } = position;
        let index = row;
        if index < self.src.len() {
            if self.src[index].len() == 0 {
                self.src[index].push_str(input);
            } else {
                self.src[index].insert_str(char_offset, input);
            }
        } else if index == self.src.len() {
            self.src.push(input.to_string());
            self.compiled_src.push(CompiledFormula::default());
            self.data.push(ResultData::None);
        } else {
            let mut i = self.src.len();
            while i < index {
                self.src.push(String::new());
                self.compiled_src.push(CompiledFormula::default());
                self.data.push(ResultData::None);
                i += 1;
            }
            self.src.push(input.to_string());
            self.compiled_src.push(CompiledFormula::default());
            self.data.push(ResultData::None);
        }
    }
}
