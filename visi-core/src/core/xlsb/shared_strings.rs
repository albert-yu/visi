use std::collections::HashMap;

use super::record::{write_empty_record, write_record};

#[derive(Default)]
pub(crate) struct SharedStrings {
    values: Vec<String>,
    indices: HashMap<String, u32>,
    total: u32,
}

impl SharedStrings {
    pub(crate) fn intern(&mut self, value: &str) -> u32 {
        self.total = self.total.saturating_add(1);
        if let Some(index) = self.indices.get(value) {
            return *index;
        }
        let index = self.values.len() as u32;
        self.values.push(value.to_string());
        self.indices.insert(value.to_string(), index);
        index
    }

    pub(crate) fn encode(&self) -> Vec<u8> {
        let mut out = Vec::new();
        let mut header = [0u8; 8];
        header[..4].copy_from_slice(&self.total.to_le_bytes());
        header[4..].copy_from_slice(&(self.values.len() as u32).to_le_bytes());
        write_record(0x009f, &header, &mut out);
        for value in &self.values {
            let encoded = value.encode_utf16().collect::<Vec<_>>();
            let mut payload = Vec::with_capacity(5 + encoded.len() * 2);
            payload.push(0);
            payload.extend_from_slice(&(encoded.len() as u32).to_le_bytes());
            for unit in encoded {
                payload.extend_from_slice(&unit.to_le_bytes());
            }
            write_record(0x0013, &payload, &mut out);
        }
        write_empty_record(0x00a0, &mut out);
        out
    }
}
