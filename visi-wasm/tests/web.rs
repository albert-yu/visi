#![cfg(target_arch = "wasm32")]

use visi_wasm::{Workbook, check_syntax, col_idx_to_letters, col_letters_to_idx, is_cell_error};
use wasm_bindgen_test::*;

#[wasm_bindgen_test]
fn workbook_eval_and_roundtrip() {
    let mut wb = Workbook::new(None).unwrap();
    wb.set_cell(0, 0, "21".to_string(), None, None).unwrap();
    wb.set_cell(0, 1, "=A1*2".to_string(), None, None).unwrap();
    wb.evaluate().unwrap();
    assert_eq!(wb.get_cell(0, 1, None).unwrap().as_f64(), Some(42.0));

    let again = wb.roundtrip().unwrap();
    assert_eq!(again.get_cell(0, 1, None).unwrap().as_f64(), Some(42.0));
}

#[wasm_bindgen_test]
fn cell_error_and_helpers() {
    let mut wb = Workbook::new(None).unwrap();
    wb.set_cell(0, 0, "=1/0".to_string(), None, None).unwrap();
    wb.evaluate().unwrap();
    let v = wb.get_cell(0, 0, None).unwrap();
    assert!(is_cell_error(&v));

    assert_eq!(col_idx_to_letters(0), "A");
    assert_eq!(col_letters_to_idx("A"), 0);
    assert_eq!(
        check_syntax("Sub Foo()\nEnd Sub\n").unwrap(),
        vec!["Foo".to_string()]
    );
}
