//! Feeds arbitrary text into the VBA lexer and parser.
//!
//! Unlike the other VBA targets here, the input is not a binary this codebase
//! wrote and then read back -- it is source someone typed, or pasted, or
//! generated. `visi macro check` and `visi macro add` both put user text
//! straight into this parser, so the never-panics property matters more here
//! than anywhere else in the VBA code.
//!
//! Checks only that it terminates without panicking or allocating without
//! bound. Whether the verdict *agrees with Excel* is a different question,
//! answered by `fuzz/fuzz_vba_parse.py`.

#![no_main]

use libfuzzer_sys::fuzz_target;

fuzz_target!(|data: &[u8]| {
    // Only valid UTF-8 can reach the parser: VBA module streams are decoded
    // to `String` well before this point.
    if let Ok(src) = std::str::from_utf8(data) {
        let _ = visi_core::core::vba::parser::parse_module(src);
    }
});
