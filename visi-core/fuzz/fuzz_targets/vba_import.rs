#![no_main]

use libfuzzer_sys::fuzz_target;
use std::collections::HashMap;

fuzz_target!(|data: &[u8]| {
    let sheet_id_by_code_name: HashMap<String, u64> = HashMap::new();
    let _ =
        visi_core::core::vba_xlsx::parse_vba_project_from_cfb_bytes(data.to_vec(), &sheet_id_by_code_name);
});
