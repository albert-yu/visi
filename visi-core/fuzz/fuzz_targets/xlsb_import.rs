#![no_main]

use libfuzzer_sys::fuzz_target;

fuzz_target!(|data: &[u8]| {
    let _ = visi_core::core::xlsb::import_xlsb_data(data, &[], |_, _, _| {});
});
