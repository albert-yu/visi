//! Feeds arbitrary bytes as a raw `vbaProject.bin` (CFB container) into the
//! VBA import pipeline: CFB parsing, the `dir` stream's OVBA decompression,
//! `PROJECTMODULES` record walking, and per-module stream decompression.
//! This is the full path a malicious or merely corrupt `.xlsm` someone
//! imports would exercise, so -- like `ovba_decompress` -- this target only
//! hunts for panics / unbounded allocation, not an output property.

#![no_main]

use libfuzzer_sys::fuzz_target;
use std::collections::HashMap;

fuzz_target!(|data: &[u8]| {
    let sheet_id_by_code_name: HashMap<String, u64> = HashMap::new();
    let _ =
        libvisi::core::vba_xlsx::parse_vba_project_from_cfb_bytes(data.to_vec(), &sheet_id_by_code_name);
});
