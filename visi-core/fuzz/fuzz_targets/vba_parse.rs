#![no_main]

use libfuzzer_sys::fuzz_target;

fuzz_target!(|data: &[u8]| {
    // Only valid UTF-8 can reach the parser: VBA module streams are decoded
    // to `String` well before this point.
    if let Ok(src) = std::str::from_utf8(data) {
        let _ = visi_core::core::vba::parser::parse_module(src);
    }
});
