use crate::core::CompiledFormula;
use crate::core::SharedVec;
use crate::render::UnitValues;
use serde::{Deserialize, Serialize};

use super::bitmask::Bitmask;
use super::cell::{CellRef, PADDING_X, generate_unique_id};
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
        self.dirty_indices.push(index);
    }

    /// Returns the minimum area required
    /// to fit the content of the src
    /// (width, height)
    pub fn min_src_dims(&self, units: UnitValues, index: usize) -> (usize, usize) {
        let (char_w, _) = units.char_dims;
        let (tile_w, _) = units.tile_dims;
        let row_str = self.src[index].to_string();
        let lines = row_str
            .split('\n')
            .map(String::from)
            .collect::<Vec<String>>();
        let char_count = lines.iter().map(|line| line.len()).max().unwrap_or(0);
        let cells_needed_width = get_cells_needed_for_text(tile_w, char_w, char_count);
        (cells_needed_width, lines.len())
    }

    pub fn min_data_height(&self, index: usize) -> usize {
        let val = self.data.get(index).unwrap_or(ResultData::None);
        match val {
            ResultData::None => {
                if let Some(src_str) = self.src.get(index) {
                    if src_str.is_empty() {
                        1
                    } else {
                        let mut nl_count = 0;
                        for b in src_str.bytes() {
                            if b == b'\n' {
                                nl_count += 1;
                            }
                        }
                        nl_count + 1
                    }
                } else {
                    1
                }
            }
            ResultData::String(ref s) => {
                let mut nl_count = 0;
                for b in s.bytes() {
                    if b == b'\n' {
                        nl_count += 1;
                    }
                }
                nl_count + 1
            }
            ResultData::Integer(_)
            | ResultData::Float(_)
            | ResultData::Boolean(_)
            | ResultData::Error(_)
            | ResultData::Plot { .. } => 1,
            _ => {
                let s = val.to_string();
                let mut nl_count = 0;
                for b in s.bytes() {
                    if b == b'\n' {
                        nl_count += 1;
                    }
                }
                nl_count + 1
            }
        }
    }

    /// Returns the minimum area required
    /// to fit the content of the data (evaluated result)
    /// (width, height)
    pub fn min_data_dims(&self, units: UnitValues, index: usize) -> (usize, usize) {
        let (char_w, _) = units.char_dims;
        let (tile_w, _) = units.tile_dims;

        let val = self.data.get(index).unwrap_or(ResultData::None);

        let (char_count, line_count, is_link) = match val {
            ResultData::None => {
                if let Some(src_str) = self.src.get(index) {
                    if src_str.is_empty() {
                        (0, 1, false)
                    } else {
                        // Check if it's a link first
                        let trimmed = src_str.trim();
                        let is_link = (trimmed.starts_with("http://")
                            || trimmed.starts_with("https://"))
                            && !trimmed.contains('\n');

                        let target_str = if is_link {
                            trimmed
                                .split("://")
                                .nth(1)
                                .and_then(|s| s.split('/').next())
                                .unwrap_or(trimmed)
                        } else {
                            src_str.as_str()
                        };

                        let mut max_len = 0;
                        let mut cur_len = 0;
                        let mut nl_count = 0;
                        for b in target_str.bytes() {
                            if b == b'\n' {
                                max_len = max_len.max(cur_len);
                                cur_len = 0;
                                nl_count += 1;
                            } else {
                                cur_len += 1;
                            }
                        }
                        max_len = max_len.max(cur_len);
                        (max_len, nl_count + 1, is_link)
                    }
                } else {
                    (0, 1, false)
                }
            }
            ResultData::String(ref s) => {
                let trimmed = s.trim();
                let is_link = (trimmed.starts_with("http://") || trimmed.starts_with("https://"))
                    && !trimmed.contains('\n');

                let target_str = if is_link {
                    trimmed
                        .split("://")
                        .nth(1)
                        .and_then(|s| s.split('/').next())
                        .unwrap_or(trimmed)
                } else {
                    s.as_str()
                };

                let mut max_len = 0;
                let mut cur_len = 0;
                let mut nl_count = 0;
                for b in target_str.bytes() {
                    if b == b'\n' {
                        max_len = max_len.max(cur_len);
                        cur_len = 0;
                        nl_count += 1;
                    } else {
                        cur_len += 1;
                    }
                }
                max_len = max_len.max(cur_len);
                (max_len, nl_count + 1, is_link)
            }
            ResultData::Integer(i) => {
                let s = i.to_string();
                (s.len(), 1, false)
            }
            ResultData::Float(f) => {
                let s = f.to_string();
                (s.len(), 1, false)
            }
            ResultData::Boolean(b) => (if b { 4 } else { 5 }, 1, false),
            ResultData::Error(ref e) => (e.len() + 6, 1, false),
            _ => {
                let s = val.to_string();
                let mut max_len = 0;
                let mut cur_len = 0;
                let mut nl_count = 0;
                for b in s.bytes() {
                    if b == b'\n' {
                        max_len = max_len.max(cur_len);
                        cur_len = 0;
                        nl_count += 1;
                    } else {
                        cur_len += 1;
                    }
                }
                max_len = max_len.max(cur_len);
                (max_len, nl_count + 1, false)
            }
        };

        let h_padding = if is_link {
            5.0 * PADDING_X
        } else {
            2.0 * PADDING_X
        };
        let text_width = char_count as f32 * char_w + h_padding;
        let cells_needed_width = (text_width / tile_w).ceil() as usize;

        (cells_needed_width.max(1), line_count)
    }

    /// Minimum dimensions needed for the column header
    pub fn min_header_dims(&self, units: UnitValues, col_idx: usize) -> (usize, usize) {
        let (char_w, _) = units.char_dims;
        let (tile_w, _) = units.tile_dims;

        let header_str = if !self.name.is_empty() {
            format!("[{}] {}", col_idx, self.name)
        } else {
            col_idx.to_string()
        };

        let lines = header_str
            .split('\n')
            .map(String::from)
            .collect::<Vec<String>>();
        let char_count = lines.iter().map(|line| line.len()).max().unwrap_or(0);
        let text_width = char_count as f32 * char_w + 2.0 * PADDING_X;
        let cells_needed_width = (text_width / tile_w).ceil() as usize;
        (cells_needed_width.max(1), lines.len())
    }

    // Returns the column's width, accounting for active cell and header
    pub fn width(&self, units: UnitValues, active_cell: Option<CellRef>, col_idx: usize) -> usize {
        let mut max_width = 0;

        // Consider the header width
        let (header_w, _) = self.min_header_dims(units, col_idx);
        max_width = max_width.max(header_w);

        let scan_limit = self.src.len().min(1000);
        for i in 0..scan_limit {
            let is_active = match active_cell {
                Some(cell_ref) => cell_ref.row == i && cell_ref.col == col_idx,
                None => false,
            };
            let (width, _) = if is_active {
                self.min_src_dims(units, i)
            } else {
                let mut dims = self.min_data_dims(units, i);
                if let Some(plot_data) = self.data.get(i) {
                    if let Some((w, _)) = plot_data.plot_cell_dims() {
                        dims.0 = dims.0.max(w);
                    }
                }
                dims
            };
            max_width = max_width.max(width);
        }

        // Always check the active cell if it is in this column but beyond the scan limit
        if let Some(cell_ref) = active_cell {
            if cell_ref.col == col_idx
                && cell_ref.row >= scan_limit
                && cell_ref.row < self.src.len()
            {
                let (width, _) = self.min_src_dims(units, cell_ref.row);
                max_width = max_width.max(width);
            }
        }

        max_width
    }
}

fn get_cells_needed_for_text(tile_w: f32, char_w: f32, char_count: usize) -> usize {
    let text_width = char_count as f32 * char_w + 2.0 * PADDING_X;
    let cells_needed_width = (text_width / tile_w).ceil() as usize;
    cells_needed_width.max(1)
}
