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

    let bytes = libvisi::export_xlsx_data(&[initial_sheet], &[], &[]).unwrap();
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
fn test_workbook_table_crud_and_evaluation() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("test_table_crud.xlsx");
    let file_str = file_path.to_str().unwrap();

    // Start from an empty workbook, exactly like `visi set` does for a
    // nonexistent file, then populate it via the same WorkbookManager API
    // the CLI itself calls into.
    let mut wb = WorkbookManager::load_file_or_create(file_str).unwrap();
    wb.set_cell(0, 0, 0, "Name".to_string());
    wb.set_cell(0, 0, 1, "Amount".to_string());
    wb.set_cell(0, 1, 0, "Widget".to_string());
    wb.set_cell(0, 1, 1, "10".to_string());
    wb.set_cell(0, 2, 0, "Gadget".to_string());
    wb.set_cell(0, 2, 1, "20".to_string());
    wb.evaluate().unwrap();

    let table_id = wb
        .add_table(None, "Sales", 0, 0, 2, 1, true, false)
        .unwrap();
    assert!(wb.find_table("Sales").is_some());
    assert_eq!(wb.list_tables().len(), 1);

    wb.set_cell(0, 0, 2, "=SUM(Sales[Amount])".to_string());
    wb.evaluate().unwrap();
    let total = wb.sheets[0].get_result_data(&libvisi::core::CellRef::new(0, 2));
    assert_eq!(total.to_string(), "30");

    // Rename a column and the table itself. Like Excel, both cascade into
    // existing formula text -- the SUM formula above (still untouched by
    // this test) must keep resolving and computing the same result after
    // both renames, without ever being rewritten by hand.
    wb.rename_table_column("Sales", 1, "Total").unwrap();
    assert_eq!(
        wb.find_table("Sales").unwrap().1.columns,
        vec!["Name", "Total"]
    );
    let src_after_col_rename = wb.sheets[0]
        .get_src(&libvisi::core::CellRef::new(0, 2))
        .cloned();
    assert_eq!(src_after_col_rename.as_deref(), Some("=SUM(Sales[Total])"));
    let total_after_col_rename = wb.sheets[0].get_result_data(&libvisi::core::CellRef::new(0, 2));
    assert_eq!(total_after_col_rename.to_string(), "30");

    wb.rename_table("Sales", "Revenue").unwrap();
    assert!(wb.find_table("Sales").is_none());
    assert!(wb.find_table("Revenue").is_some());

    let src_after_rename = wb.sheets[0]
        .get_src(&libvisi::core::CellRef::new(0, 2))
        .cloned();
    assert_eq!(src_after_rename.as_deref(), Some("=SUM(Revenue[Total])"));
    let total_after_rename = wb.sheets[0].get_result_data(&libvisi::core::CellRef::new(0, 2));
    assert_eq!(total_after_rename.to_string(), "30");

    // Save and reload: the table definition and the structured-reference
    // formula must both survive the xlsx round trip.
    wb.save_file(file_str).unwrap();
    let mut reloaded = WorkbookManager::load_file(file_str).unwrap();
    assert_eq!(reloaded.list_tables().len(), 1);
    let (sheet, table) = reloaded.find_table("Revenue").unwrap();
    assert_eq!(sheet.name, "Sheet1");
    assert_eq!(table.columns, vec!["Name", "Total"]);

    reloaded.evaluate().unwrap();
    let total_after_reload = reloaded.sheets[0].get_result_data(&libvisi::core::CellRef::new(0, 2));
    assert_eq!(total_after_reload.to_string(), "30");

    reloaded.delete_table("Revenue").unwrap();
    assert!(reloaded.list_tables().is_empty());

    let _ = table_id;
    let _ = fs::remove_file(file_path);
}

#[test]
fn test_workbook_pivot_crud_and_computation() {
    use libvisi::core::{PivotAggregation, PivotArea};

    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("test_pivot_crud.xlsx");
    let file_str = file_path.to_str().unwrap();

    let mut wb = WorkbookManager::load_file_or_create(file_str).unwrap();
    let data = [
        ["Region", "Product", "Amount"],
        ["East", "Widget", "10"],
        ["East", "Gadget", "5"],
        ["West", "Widget", "30"],
        ["West", "Gadget", "40"],
    ];
    for (r, row) in data.iter().enumerate() {
        for (c, v) in row.iter().enumerate() {
            wb.set_cell(0, r, c, v.to_string());
        }
    }
    wb.evaluate().unwrap();
    wb.add_table(None, "Sales", 0, 0, 4, 2, true, false)
        .unwrap();

    let pivot_id = wb
        .add_pivot_table_from_table("SalesPivot", "Sales", None, 0, 4)
        .unwrap();
    assert!(wb.find_pivot_table("SalesPivot").is_some());

    wb.add_pivot_field("SalesPivot", PivotArea::Row, "Region", None)
        .unwrap();
    wb.add_pivot_field(
        "SalesPivot",
        PivotArea::Value,
        "Amount",
        Some(PivotAggregation::Sum),
    )
    .unwrap();

    // East = 10+5 = 15, West = 30+40 = 70, Grand Total = 85, materialized
    // as plain values at the pivot's destination (column E, 0-based col 4).
    let east = wb.sheets[0].get_result_data(&libvisi::core::CellRef::new(1, 5));
    let west = wb.sheets[0].get_result_data(&libvisi::core::CellRef::new(2, 5));
    let grand_total = wb.sheets[0].get_result_data(&libvisi::core::CellRef::new(3, 5));
    assert_eq!(east.to_string(), "15");
    assert_eq!(west.to_string(), "70");
    assert_eq!(grand_total.to_string(), "85");

    // Save and reload: like a table, a pivot table's definition (source,
    // destination, row/value fields) must survive the xlsx round trip so a
    // later CLI invocation can keep editing it.
    wb.save_file(file_str).unwrap();
    let mut reloaded = WorkbookManager::load_file(file_str).unwrap();
    let reloaded_pivot = reloaded.find_pivot_table("SalesPivot").unwrap().clone();
    assert_eq!(reloaded_pivot.row_fields.len(), 1);
    assert_eq!(reloaded_pivot.row_fields[0].column, "Region");
    assert_eq!(reloaded_pivot.value_fields.len(), 1);
    assert_eq!(reloaded_pivot.value_fields[0].column, "Amount");

    // Adding a second value field after reload should still refresh
    // correctly against the (re-resolved) source table.
    reloaded
        .add_pivot_field(
            "SalesPivot",
            PivotArea::Value,
            "Amount",
            Some(PivotAggregation::Count),
        )
        .unwrap();
    let east_count = reloaded.sheets[0].get_result_data(&libvisi::core::CellRef::new(2, 6));
    assert_eq!(east_count.to_string(), "2");

    reloaded.delete_pivot_table("SalesPivot").unwrap();
    assert!(reloaded.find_pivot_table("SalesPivot").is_none());

    let _ = pivot_id;
    let _ = fs::remove_file(file_path);
}

#[test]
fn test_pivot_filter_field_materializes_as_header_row_above_grid() {
    use libvisi::core::{CellRef, PivotAggregation, PivotArea};

    // Verified against real Excel: a filter/page field always renders as
    // its own row above the row/col header grid ("FieldName" | "(All)" or
    // "(Multiple Items)"), followed by one blank spacer row -- visi's
    // pivot output previously never rendered these at all, leaving every
    // pivot with a filter field completely cell-misaligned vs. Excel.
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("test_pivot_filter_header_row.xlsx");
    let file_str = file_path.to_str().unwrap();

    let mut wb = WorkbookManager::load_file_or_create(file_str).unwrap();
    let data = [
        ["Region", "Product", "Amount"],
        ["East", "Widget", "10"],
        ["East", "Gadget", "5"],
        ["West", "Widget", "30"],
        ["West", "Gadget", "40"],
    ];
    for (r, row) in data.iter().enumerate() {
        for (c, v) in row.iter().enumerate() {
            wb.set_cell(0, r, c, v.to_string());
        }
    }
    wb.evaluate().unwrap();
    wb.add_table(None, "Sales", 0, 0, 4, 2, true, false)
        .unwrap();

    // Destination anchored at (row 0, col 4) = E1.
    wb.add_pivot_table_from_table("SalesPivot", "Sales", None, 0, 4)
        .unwrap();
    wb.add_pivot_field("SalesPivot", PivotArea::Row, "Region", None)
        .unwrap();
    wb.add_pivot_field(
        "SalesPivot",
        PivotArea::Value,
        "Amount",
        Some(PivotAggregation::Sum),
    )
    .unwrap();
    wb.add_pivot_field("SalesPivot", PivotArea::Filter, "Product", None)
        .unwrap();

    // Row 0 (E1:F1): filter field name + "(All)" (no selection applied yet).
    assert_eq!(
        wb.sheets[0]
            .get_result_data(&CellRef::new(0, 4))
            .to_string(),
        "Product"
    );
    assert_eq!(
        wb.sheets[0]
            .get_result_data(&CellRef::new(0, 5))
            .to_string(),
        "(All)"
    );
    // Row 1 is the blank spacer row.
    assert_eq!(
        wb.sheets[0]
            .get_result_data(&CellRef::new(1, 4))
            .to_string(),
        ""
    );
    // The row/col header grid now starts at row 2, not row 0.
    assert_eq!(
        wb.sheets[0]
            .get_result_data(&CellRef::new(2, 4))
            .to_string(),
        "Region"
    );
    assert_eq!(
        wb.sheets[0]
            .get_result_data(&CellRef::new(3, 4))
            .to_string(),
        "East"
    );
    assert_eq!(
        wb.sheets[0]
            .get_result_data(&CellRef::new(3, 5))
            .to_string(),
        "15"
    );

    // Selecting a strict subset switches the state to "(Multiple Items)"
    // (matches real Excel: even a single selected value shows this, never
    // the value's own name).
    wb.set_pivot_filter("SalesPivot", "Product", Some(vec!["Widget".to_string()]))
        .unwrap();
    assert_eq!(
        wb.sheets[0]
            .get_result_data(&CellRef::new(0, 5))
            .to_string(),
        "(Multiple Items)"
    );

    // The pivot definition's `dest_row`/`dest_col` must still be the
    // anchor of the whole visual block (E1), not shifted down to where the
    // grid itself begins -- otherwise every refresh after a reload would
    // drift further down.
    wb.save_file(file_str).unwrap();
    let reloaded = WorkbookManager::load_file(file_str).unwrap();
    let reloaded_pivot = reloaded.find_pivot_table("SalesPivot").unwrap();
    assert_eq!(reloaded_pivot.dest_row, 0);
    assert_eq!(reloaded_pivot.dest_col, 4);

    let _ = fs::remove_file(file_path);
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
