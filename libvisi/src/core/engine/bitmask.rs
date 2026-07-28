use crate::core::SharedVec;
use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Bitmask {
    data: SharedVec<u8>,
    pub len: usize,
}

impl Bitmask {
    pub fn with_size(size: usize) -> Self {
        Self {
            data: vec![0; (size + 7) / 8].into(),
            len: size,
        }
    }
    pub fn push(&mut self, value: bool) {
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
    pub fn get(&self, index: usize) -> bool {
        if index >= self.len {
            return false;
        }
        let byte_idx = index / 8;
        let bit_idx = index % 8;
        (self.data[byte_idx] & (1 << bit_idx)) != 0
    }
    pub fn set(&mut self, index: usize, value: bool) {
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
    pub fn insert(&mut self, index: usize, value: bool) {
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
    pub fn remove(&mut self, index: usize) {
        if index >= self.len {
            return;
        }
        for i in index..self.len - 1 {
            let next = self.get(i + 1);
            self.set(i, next);
        }
        self.len -= 1;
        let required_bytes = (self.len + 7) / 8;
        if self.data.len() > required_bytes {
            self.data.pop();
        }
    }
    pub fn drain<R: std::ops::RangeBounds<usize> + Clone>(&mut self, range: R) {
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
        self.data.truncate((self.len + 7) / 8);
    }
}
