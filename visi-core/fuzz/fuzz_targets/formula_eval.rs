#![no_main]

use libfuzzer_sys::fuzz_target;
use visi_core::core::{Sheet, SheetInit};

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
