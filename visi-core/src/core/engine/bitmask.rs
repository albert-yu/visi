//! A packed bit-per-row validity mask.

use crate::core::SharedVec;
use serde::{Deserialize, Serialize};

/// One bit per row, recording which entries of a numeric [`ColumnData`] hold a
/// value rather than a blank.
///
/// A set bit means the value at that index is real; a clear bit means the cell
/// is empty and the underlying slot holds a placeholder. Keeping this separate
/// is what lets a numeric column stay unboxed and still tell a blank cell
/// apart from a zero.
///
/// Read-only from outside the crate -- the mutators are crate-private, since
/// changing a mask's length independently of the column it belongs to would
/// desync the two.
///
/// [`ColumnData`]: crate::core::ColumnData
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Bitmask {
    data: SharedVec<u8>,
    /// How many bits are in use, which is the row count of the column this
    /// mask belongs to. Not the capacity of the backing bytes.
    pub len: usize,
}

impl Bitmask {
    pub(crate) fn with_size(size: usize) -> Self {
        Self {
            data: vec![0; size.div_ceil(8)].into(),
            len: size,
        }
    }
    pub(crate) fn push(&mut self, value: bool) {
        let byte_idx = self.len / 8;
        let bit_idx = self.len % 8;
        if byte_idx >= self.data.len() {
            self.data.push(0);
        }
        if value {
            self.data[byte_idx] |= 1 << bit_idx;
        } else {
            self.data[byte_idx] &= !(1 << bit_idx);
        }
        self.len += 1;
    }
    /// Whether the entry at `index` holds a value. `false` for an index at or
    /// past [`Bitmask::len`], so an out-of-range read is indistinguishable
    /// from a blank.
    pub fn get(&self, index: usize) -> bool {
        if index >= self.len {
            return false;
        }
        let byte_idx = index / 8;
        let bit_idx = index % 8;
        (self.data[byte_idx] & (1 << bit_idx)) != 0
    }
    pub(crate) fn set(&mut self, index: usize, value: bool) {
        if index >= self.len {
            return;
        }
        let byte_idx = index / 8;
        let bit_idx = index % 8;
        if value {
            self.data[byte_idx] |= 1 << bit_idx;
        } else {
            self.data[byte_idx] &= !(1 << bit_idx);
        }
    }
    pub(crate) fn insert(&mut self, index: usize, value: bool) {
        if index >= self.len {
            self.push(value);
            return;
        }
        self.push(false);
        for i in (index + 1..self.len).rev() {
            let prev = self.get(i - 1);
            self.set(i, prev);
        }
        self.set(index, value);
    }
    pub(crate) fn remove(&mut self, index: usize) {
        if index >= self.len {
            return;
        }
        for i in index..self.len - 1 {
            let next = self.get(i + 1);
            self.set(i, next);
        }
        self.len -= 1;
        let required_bytes = self.len.div_ceil(8);
        if self.data.len() > required_bytes {
            self.data.pop();
        }
    }
    pub(crate) fn drain<R: std::ops::RangeBounds<usize> + Clone>(&mut self, range: R) {
        let start = match range.start_bound() {
            std::ops::Bound::Included(&n) => n,
            std::ops::Bound::Excluded(&n) => n + 1,
            std::ops::Bound::Unbounded => 0,
        };
        let end = match range.end_bound() {
            std::ops::Bound::Included(&n) => n + 1,
            std::ops::Bound::Excluded(&n) => n,
            std::ops::Bound::Unbounded => self.len,
        };
        let start = start.min(self.len);
        let end = end.min(self.len);
        if start >= end {
            return;
        }
        let count = end - start;
        for i in end..self.len {
            let val = self.get(i);
            self.set(i - count, val);
        }
        self.len -= count;
        self.data.truncate(self.len.div_ceil(8));
    }
}
