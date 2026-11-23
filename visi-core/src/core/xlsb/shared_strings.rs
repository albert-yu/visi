use std::collections::HashMap;

use super::package::read_varint;
use super::record::{write_empty_record, write_record};

#[derive(Debug, Clone, Default)]
pub(crate) struct SharedStrings {
    raw_items: Vec<Vec<u8>>,
    indices: HashMap<String, u32>,
    total: u32,
    pub(crate) dirty: bool,
}

impl SharedStrings {
    pub(crate) fn parse(data: &[u8]) -> Result<Self, String> {
        let mut sst = Self::default();
        let mut offset = 0usize;
        while offset < data.len() {
            let id = read_varint(data, &mut offset)
                .ok_or_else(|| "truncated XLSB sharedStrings record id".to_string())?;
            let len = read_varint(data, &mut offset)
                .ok_or_else(|| "truncated XLSB sharedStrings record length".to_string())?
                as usize;
            let end = offset
                .checked_add(len)
                .filter(|end| *end <= data.len())
                .ok_or_else(|| {
                    "XLSB sharedStrings record extends past the part boundary".to_string()
                })?;
            let payload = &data[offset..end];
            offset = end;
            match id {
                0x009f if payload.len() >= 8 => {
                    sst.total = u32::from_le_bytes(payload[0..4].try_into().unwrap());
                }
                0x0013 if payload.len() >= 5 => {
                    let index = sst.raw_items.len() as u32;
                    let flags = payload[0];
                    if flags == 0 {
                        let cch = u32::from_le_bytes(payload[1..5].try_into().unwrap()) as usize;
                        let str_end = 5usize.saturating_add(cch.saturating_mul(2));
                        if str_end <= payload.len() {
                            let units = payload[5..str_end]
                                .as_chunks::<2>()
                                .0
                                .iter()
                                .map(|b| u16::from_le_bytes(*b))
                                .collect::<Vec<_>>();
                            if let Ok(decoded) = String::from_utf16(&units) {
                                sst.indices.entry(decoded).or_insert(index);
                            }
                        }
                    }
                    sst.raw_items.push(payload.to_vec());
                }
                0x00a0 => break,
                _ => {}
            }
        }
        Ok(sst)
    }

    pub(crate) fn intern(&mut self, value: &str) -> u32 {
        self.total = self.total.saturating_add(1);
        self.dirty = true;
        if let Some(index) = self.indices.get(value) {
            return *index;
        }
        let index = self.raw_items.len() as u32;
        let encoded = value.encode_utf16().collect::<Vec<_>>();
        let mut payload = Vec::with_capacity(5 + encoded.len() * 2);
        payload.push(0);
        payload.extend_from_slice(&(encoded.len() as u32).to_le_bytes());
        for unit in encoded {
            payload.extend_from_slice(&unit.to_le_bytes());
        }
        self.raw_items.push(payload);
        self.indices.insert(value.to_string(), index);
        index
    }

    pub(crate) fn encode(&self) -> Vec<u8> {
        let mut out = Vec::new();
        let mut header = [0u8; 8];
        header[..4].copy_from_slice(&self.total.to_le_bytes());
        header[4..].copy_from_slice(&(self.raw_items.len() as u32).to_le_bytes());
        write_record(0x009f, &header, &mut out);
        for payload in &self.raw_items {
            write_record(0x0013, payload, &mut out);
        }
        write_empty_record(0x00a0, &mut out);
        out
    }
}
