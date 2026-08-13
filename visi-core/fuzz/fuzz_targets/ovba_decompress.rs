//! Feeds arbitrary bytes into the MS-OVBA "Compressed Container" decoder.
//! `decompress` is the only place in this codebase that parses a
//! completely untrusted byte stream (a `dir` or module stream pulled out of
//! someone else's `.xlsm`), so this target exists purely to find panics /
//! unbounded allocation, not to check any output property.

#![no_main]

use libfuzzer_sys::fuzz_target;

fuzz_target!(|data: &[u8]| {
    let _ = visi_core::core::ovba::decompress(data);
});
