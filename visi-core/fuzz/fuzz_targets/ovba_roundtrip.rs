#![no_main]

use libfuzzer_sys::fuzz_target;

fuzz_target!(|data: &[u8]| {
    if let Ok(compressed) = visi_core::core::ovba::compress(data) {
        let decompressed = visi_core::core::ovba::decompress(&compressed)
            .expect("compressed data must always decompress");
        assert_eq!(decompressed, data);
    }
});
