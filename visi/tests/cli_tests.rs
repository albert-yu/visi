use std::fs;
use visi::engine::WorkbookManager;
use visi::utils::{parse_cell_ref, parse_range_ref};

#[test]
fn test_workbook_create_and_formula_evaluation() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("test_formulas_eval.xlsx");
    let file_str = file_path.to_str().unwrap();

    let mut initial_sheet = libvisi::core::Sheet::new(libvisi::core::SheetInit {
        id: None,
        name: Some("Sheet1".to_string()),
        rows: 3,
        cols: 2,
    });

    initial_sheet.set_cell_src(0, 0, "10".to_string());
    initial_sheet.set_cell_src(1, 0, "20".to_string());
    initial_sheet.set_cell_src(0, 1, "=A1 + A2".to_string());
    initial_sheet.set_cell_src(1, 1, "=SUM(A1:A2)".to_string());

    let bytes = libvisi::export_xlsx_data(&[initial_sheet], &[]).unwrap();
    fs::write(&file_path, bytes).unwrap();

    // Load with WorkbookManager and evaluate
    let mut wb = WorkbookManager::load_file(file_str).unwrap();
    wb.evaluate().unwrap();

    let sheet = &wb.sheets[0];
    let b1_val = sheet.get_result_data(&libvisi::core::CellRef::new(0, 1));
    let b2_val = sheet.get_result_data(&libvisi::core::CellRef::new(1, 1));

    assert_eq!(b1_val.to_string(), "30");
    assert_eq!(b2_val.to_string(), "30");

    // Update cell A1 to 50
    wb.set_cell(0, 0, 0, "50".to_string());
    wb.evaluate().unwrap();

    let b1_val_updated = wb.sheets[0].get_result_data(&libvisi::core::CellRef::new(0, 1));
    let b2_val_updated = wb.sheets[0].get_result_data(&libvisi::core::CellRef::new(1, 1));

    assert_eq!(b1_val_updated.to_string(), "70");
    assert_eq!(b2_val_updated.to_string(), "70");

    // Save file
    let out_path = temp_dir.join("output_eval.xlsx");
    wb.save_file(out_path.to_str().unwrap()).unwrap();
    assert!(out_path.exists());
    let _ = fs::remove_file(file_path);
    let _ = fs::remove_file(out_path);
}

#[test]
fn test_coordinate_parsing() {
    let (sheet, row, col) = parse_cell_ref("Sheet2!D10").unwrap();
    assert_eq!(sheet, Some("Sheet2".to_string()));
    assert_eq!(row, 9);
    assert_eq!(col, 3);

    let (sheet, s_row, s_col, e_row, e_col) = parse_range_ref("A1:B5").unwrap();
    assert_eq!(sheet, None);
    assert_eq!((s_row, s_col), (0, 0));
    assert_eq!((e_row, e_col), (4, 1));
}
