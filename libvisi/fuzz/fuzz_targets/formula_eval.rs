//! Feeds arbitrary bytes as a formula string through the *full* formula
//! pipeline -- `compile_formula` (name/ref resolution against real sheet
//! IDs), `serialize_formula`, `parse_excel_formula` (the recursive-descent
//! `Expr` parser), and `evaluate_ast`/`evaluate_function` -- against a
//! small pre-populated sheet, via `Sheet::commit` exactly as a real edit
//! would. Unlike `../../fuzz/fuzz_excel.py` (differential: compares
//! `visi`'s output against real Excel's), this only hunts for panics,
//! unbounded allocation, and stack overflow / infinite loops on
//! adversarial formula text -- e.g. deeply nested parentheses/function
//! calls stressing the recursive-descent parser and recursive AST
//! evaluation, or a malformed structured/range reference. No output
//! property is checked.

#![no_main]

use libfuzzer_sys::fuzz_target;
use libvisi::core::{Sheet, SheetInit};

fuzz_target!(|data: &[u8]| {
    let Ok(text) = std::str::from_utf8(data) else {
        return;
    };
    if text.is_empty() || text.len() > 8192 {
        return;
    }

    let mut sheet = Sheet::new(SheetInit {
        id: None,
        name: Some("Sheet1".to_string()),
        rows: 20,
        cols: 20,
    });
    sheet.set_cell_src(0, 0, "1".to_string());
    sheet.set_cell_src(1, 0, "2".to_string());
    sheet.set_cell_src(2, 0, "hello".to_string());
    sheet.set_cell_src(0, 1, "3.5".to_string());
    sheet.set_cell_src(1, 1, "TRUE".to_string());
    let _ = sheet.commit(None);

    let src = if text.starts_with('=') {
        text.to_string()
    } else {
        format!("={text}")
    };
    sheet.set_cell_src(10, 10, src);
    let _ = sheet.commit(None);
});
