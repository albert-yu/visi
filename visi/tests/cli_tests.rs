use clap::Parser;
use std::fs;
use visi::cli::{ChartSubcommands, Cli, Commands, PivotSubcommands};
use visi::engine::WorkbookManager;
use visi::utils::{parse_cell_ref, parse_range_ref};

#[test]
fn test_chart_edit_parses_setters_and_clear_flags() {
    let cli = Cli::try_parse_from([
        "visi",
        "chart",
        "edit",
        "data.xlsx",
        "--id",
        "5",
        "--name",
        "New",
        "--chart-type",
        "bar",
        "--range",
        "Sheet1!A1:C5",
        "--title",
        "T",
        "--xlabel",
        "X",
        "--show-legend",
        "--anchor",
        "D5",
        "-i",
    ])
    .expect("clap should parse a full chart edit invocation");

    let Commands::Chart(chart_args) = cli.command else {
        panic!("expected Commands::Chart");
    };
    let ChartSubcommands::Edit(edit_args) = chart_args.command else {
        panic!("expected ChartSubcommands::Edit");
    };
    assert_eq!(edit_args.id, 5);
    assert_eq!(edit_args.name.as_deref(), Some("New"));
    assert_eq!(edit_args.range.as_deref(), Some("Sheet1!A1:C5"));
    assert_eq!(edit_args.title.as_deref(), Some("T"));
    assert_eq!(edit_args.xlabel.as_deref(), Some("X"));
    assert!(edit_args.show_legend);
    assert_eq!(edit_args.anchor.as_deref(), Some("D5"));
    assert!(edit_args.in_place);
}

#[test]
fn test_chart_edit_title_and_clear_title_are_mutually_exclusive() {
    let result = Cli::try_parse_from([
        "visi",
        "chart",
        "edit",
        "data.xlsx",
        "--id",
        "5",
        "--title",
        "T",
        "--clear-title",
    ]);
    assert!(result.is_err());
}

#[test]
fn test_chart_edit_show_legend_and_hide_legend_are_mutually_exclusive() {
    let result = Cli::try_parse_from([
        "visi",
        "chart",
        "edit",
        "data.xlsx",
        "--id",
        "5",
        "--show-legend",
        "--hide-legend",
    ]);
    assert!(result.is_err());
}

#[test]
fn test_pivot_filter_values_accepts_leading_hyphen_values() {
    // Regression test: `visi pivot filter --values -7,3,12` used to fail
    // clap parsing entirely -- a numeric-looking column's filter values can
    // legitimately start with "-" (e.g. a negative Amount), but without
    // `allow_hyphen_values` clap mistook the leading "-7" for an unknown
    // flag ("error: unexpected argument '-7' found") instead of treating it
    // as the value for `--values`. Found via fuzz/fuzz_pivot.py, whose
    // NumStr source column includes negative-number-looking text.
    let cli = Cli::try_parse_from([
        "visi",
        "pivot",
        "filter",
        "data.xlsx",
        "--name",
        "P1",
        "--column",
        "Num",
        "--values",
        "-7,3,12",
    ])
    .expect("clap should accept hyphen-leading filter values");

    let Commands::Pivot(pivot_args) = cli.command else {
        panic!("expected Commands::Pivot");
    };
    let PivotSubcommands::Filter(filter_args) = pivot_args.command else {
        panic!("expected PivotSubcommands::Filter");
    };
    assert_eq!(filter_args.values, vec!["-7", "3", "12"]);
}

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

    let bytes = libvisi::export_xlsx_data(&[initial_sheet], &[], &[], None).unwrap();
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
fn test_workbook_vba_crud_and_roundtrip() {
    use libvisi::core::VbaModuleKind;

    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("test_vba_crud.xlsm");
    let file_str = file_path.to_str().unwrap();

    let mut wb = WorkbookManager::load_file_or_create(file_str).unwrap();
    assert!(!wb.has_vba_project());
    assert!(wb.list_vba_modules().is_empty());

    wb.add_vba_module(
        "Module1".to_string(),
        VbaModuleKind::Standard,
        "Attribute VB_Name = \"Module1\"\r\nSub Foo()\r\nEnd Sub\r\n".to_string(),
        None,
    )
    .unwrap();
    assert!(wb.has_vba_project());
    assert_eq!(wb.list_vba_modules().len(), 1);

    // Duplicate names are rejected, case-insensitively.
    assert!(
        wb.add_vba_module(
            "module1".to_string(),
            VbaModuleKind::Standard,
            String::new(),
            None,
        )
        .is_err()
    );

    let sheet1_id = wb.sheets[0].id;
    wb.add_vba_module(
        "Sheet1".to_string(),
        VbaModuleKind::Document,
        "Attribute VB_Name = \"Sheet1\"\r\n".to_string(),
        Some(sheet1_id),
    )
    .unwrap();
    assert_eq!(wb.list_vba_modules().len(), 2);

    wb.rename_vba_module("Module1", "Helpers").unwrap();
    assert!(
        wb.vba_project
            .as_ref()
            .unwrap()
            .find_module("Module1")
            .is_none()
    );
    assert!(
        wb.vba_project
            .as_ref()
            .unwrap()
            .find_module("Helpers")
            .is_some()
    );

    wb.set_vba_module_source(
        "Helpers",
        "Attribute VB_Name = \"Helpers\"\r\nSub Bar()\r\nEnd Sub\r\n".to_string(),
    )
    .unwrap();
    assert_eq!(
        wb.vba_project
            .as_ref()
            .unwrap()
            .find_module("Helpers")
            .unwrap()
            .source,
        "Attribute VB_Name = \"Helpers\"\r\nSub Bar()\r\nEnd Sub\r\n"
    );

    // Save and reload: module list, kinds, sources, and the document
    // module's sheet binding must all survive the xlsx round trip.
    wb.save_file(file_str).unwrap();
    let reloaded = WorkbookManager::load_file(file_str).unwrap();
    let project = reloaded
        .vba_project
        .as_ref()
        .expect("vba project should survive reload");
    assert_eq!(project.modules.len(), 2);

    let helpers = project
        .find_module("Helpers")
        .expect("Helpers module should survive reload");
    assert_eq!(helpers.kind, VbaModuleKind::Standard);
    assert_eq!(
        helpers.source,
        "Attribute VB_Name = \"Helpers\"\r\nSub Bar()\r\nEnd Sub\r\n"
    );

    let sheet1_module = project
        .find_module("Sheet1")
        .expect("Sheet1 document module should survive reload");
    assert_eq!(sheet1_module.kind, VbaModuleKind::Document);
    let reloaded_sheet1_id = reloaded.sheets[0].id;
    assert_eq!(sheet1_module.bound_sheet_id, Some(reloaded_sheet1_id));

    // Continue editing after reload, exactly like a later CLI invocation
    // would: remove a module and confirm only it disappears.
    let mut reloaded = reloaded;
    reloaded.remove_vba_module("Helpers").unwrap();
    assert_eq!(reloaded.list_vba_modules().len(), 1);
    assert!(
        reloaded
            .vba_project
            .as_ref()
            .unwrap()
            .find_module("Sheet1")
            .is_some()
    );

    reloaded.save_file(file_str).unwrap();
    let final_reload = WorkbookManager::load_file(file_str).unwrap();
    assert_eq!(final_reload.list_vba_modules().len(), 1);

    let _ = fs::remove_file(file_path);
}

#[test]
fn test_workbook_vba_this_workbook_module_needs_no_sheet_binding() {
    use libvisi::core::VbaModuleKind;

    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("test_vba_this_workbook.xlsm");
    let file_str = file_path.to_str().unwrap();

    let mut wb = WorkbookManager::load_file_or_create(file_str).unwrap();

    // A ThisWorkbook document module -- like real Excel's own always-present
    // one -- isn't tied to any particular sheet, so bound_sheet_id: None
    // must be accepted rather than rejected as "Document modules require a
    // bound sheet".
    wb.add_vba_module(
        "ThisWorkbook".to_string(),
        VbaModuleKind::Document,
        "Attribute VB_Name = \"ThisWorkbook\"\r\n".to_string(),
        None,
    )
    .unwrap();
    assert_eq!(
        wb.vba_project
            .as_ref()
            .unwrap()
            .find_module("ThisWorkbook")
            .unwrap()
            .bound_sheet_id,
        None
    );

    // Regression test: adding ThisWorkbook used to force a caller to pass
    // some --sheet, which got stored as ThisWorkbook's bound_sheet_id and
    // then made that sheet's *real* code-behind module add fail with "That
    // sheet already has a bound document module".
    let sheet1_id = wb.sheets[0].id;
    wb.add_vba_module(
        "Sheet1".to_string(),
        VbaModuleKind::Document,
        "Attribute VB_Name = \"Sheet1\"\r\n".to_string(),
        Some(sheet1_id),
    )
    .unwrap();
    assert_eq!(
        wb.vba_project
            .as_ref()
            .unwrap()
            .find_module("Sheet1")
            .unwrap()
            .bound_sheet_id,
        Some(sheet1_id)
    );

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
        .add_pivot_table_from_table("SalesPivot", "Sales", None, 0, 4, true, true)
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
fn test_workbook_chart_edit_and_round_trip() {
    use libvisi::core::chart::ChartType;

    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("test_chart_edit.xlsx");
    let file_str = file_path.to_str().unwrap();

    let mut wb = WorkbookManager::load_file_or_create(file_str).unwrap();
    let data = [["Cat1", "10"], ["Cat2", "20"], ["Cat3", "30"]];
    for (r, row) in data.iter().enumerate() {
        for (c, v) in row.iter().enumerate() {
            wb.set_cell(0, r, c, v.to_string());
        }
    }
    wb.evaluate().unwrap();

    let chart_id = wb
        .add_chart(
            "Sheet1",
            ChartType::Line,
            "Sheet1!A1:B3".to_string(),
            Some("Orig".to_string()),
            Some((3, 1)),
        )
        .unwrap();

    // Exercise a genuine mix of set / clear / leave-unchanged fields.
    wb.edit_chart(
        chart_id,
        Some("Renamed".to_string()),
        Some(ChartType::Bar),
        Some("Sheet1!A1:B3".to_string()),
        Some(Some("New Title".to_string())),
        Some(Some("X".to_string())),
        None, // leave ylabel unchanged (stays None)
        Some(false),
        Some((5, 2)),
    )
    .unwrap();

    {
        let chart = wb.charts.iter().find(|c| c.id == chart_id).unwrap();
        assert_eq!(chart.name, "Renamed");
        assert_eq!(chart.chart_type, ChartType::Bar);
        assert_eq!(chart.title, Some("New Title".to_string()));
        assert_eq!(chart.xlabel, Some("X".to_string()));
        assert_eq!(chart.ylabel, None);
        assert!(!chart.show_legend);
        assert_eq!((chart.anchor_row, chart.anchor_col), (5, 2));
    }

    // Clear the title in a second edit call.
    wb.edit_chart(
        chart_id,
        None,
        None,
        None,
        Some(None),
        None,
        None,
        None,
        None,
    )
    .unwrap();
    assert_eq!(
        wb.charts.iter().find(|c| c.id == chart_id).unwrap().title,
        None
    );

    // Save + reload: chart_type/data_range/title/xlabel/ylabel/show_legend/
    // anchor must survive the xlsx round trip. `id`/`name` do NOT round-trip
    // today (xlsx import always regenerates both), so they're deliberately
    // not asserted on here.
    wb.save_file(file_str).unwrap();
    let reloaded = WorkbookManager::load_file(file_str).unwrap();
    assert_eq!(reloaded.charts.len(), 1);
    let reloaded_chart = &reloaded.charts[0];
    assert_eq!(reloaded_chart.chart_type, ChartType::Bar);
    assert_eq!(reloaded_chart.data_range, "Sheet1!A1:B3");
    assert_eq!(reloaded_chart.title, None);
    assert_eq!(reloaded_chart.xlabel, Some("X".to_string()));
    assert!(!reloaded_chart.show_legend);
    assert_eq!(
        (reloaded_chart.anchor_row, reloaded_chart.anchor_col),
        (5, 2)
    );

    // Editing a nonexistent id must error.
    let mut wb2 = reloaded;
    assert!(
        wb2.edit_chart(
            999999,
            Some("X".into()),
            None,
            None,
            None,
            None,
            None,
            None,
            None
        )
        .is_err()
    );

    let _ = fs::remove_file(file_path);
}

#[test]
fn test_pivot_field_area_reassignment_evicts_from_previous_area() {
    use libvisi::core::{PivotAggregation, PivotArea};

    // Matches real Excel's PivotField.Orientation semantics (verified via
    // the win32com fuzz driver): a field can only live in one of
    // Row/Column/Filter at a time, so re-adding it to a different area
    // moves it rather than leaving it in both. Previously visi's
    // `add_pivot_field` just pushed onto the target area's Vec, so a field
    // used as e.g. both a column field and a filter field would stay
    // grouped by in the column area *and* filtered -- a config real Excel
    // can't represent, since assigning the page/filter orientation there
    // relocates the field out of the column area. Discovered via
    // fuzz/fuzz_pivot.py's differential fuzzer (iteration mismatched by 8
    // cells because visi kept a two-level column grouping Excel had
    // collapsed to one level).
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("test_pivot_field_area_reassignment.xlsx");
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
    wb.add_pivot_table_from_table("SalesPivot", "Sales", None, 0, 4, true, true)
        .unwrap();

    wb.add_pivot_field("SalesPivot", PivotArea::Column, "Region", None)
        .unwrap();
    wb.add_pivot_field(
        "SalesPivot",
        PivotArea::Value,
        "Amount",
        Some(PivotAggregation::Sum),
    )
    .unwrap();
    // Re-assign "Region" to the filter area: it should vanish from
    // col_fields, not appear in both.
    wb.add_pivot_field("SalesPivot", PivotArea::Filter, "Region", None)
        .unwrap();

    let pivot = wb.find_pivot_table("SalesPivot").unwrap();
    assert!(pivot.col_fields.is_empty());
    assert_eq!(pivot.filter_fields.len(), 1);
    assert_eq!(pivot.filter_fields[0].column, "Region");

    // Value fields are the exception: Excel allows the same source column
    // to be summarized multiple ways at once, so adding "Amount" to Value
    // again (a different aggregation) must not evict the existing one, and
    // adding a Row/Column/Filter field must not evict any value field.
    wb.add_pivot_field(
        "SalesPivot",
        PivotArea::Value,
        "Amount",
        Some(PivotAggregation::Average),
    )
    .unwrap();
    wb.add_pivot_field("SalesPivot", PivotArea::Row, "Product", None)
        .unwrap();
    let pivot = wb.find_pivot_table("SalesPivot").unwrap();
    assert_eq!(pivot.value_fields.len(), 2);
    assert_eq!(pivot.row_fields.len(), 1);

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
    wb.add_pivot_table_from_table("SalesPivot", "Sales", None, 0, 4, true, true)
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
    // The row/col header grid now starts at row 2, not row 0. The outermost
    // (here, only) row field's caption is Excel's literal "Row Labels" text,
    // not the field's own name (matches real Excel's compact-form display).
    assert_eq!(
        wb.sheets[0]
            .get_result_data(&CellRef::new(2, 4))
            .to_string(),
        "Row Labels"
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
