use std::ffi::OsStr;
use std::fs;
use visi::cli::{ChartSubcommands, Cli, Commands, MacroSubcommands, PivotSubcommands};
use visi::engine::{WorkbookFile, WorkbookManager};
use visi::utils::{parse_cell_ref, parse_range_ref};

fn try_parse<'a>(args: &'a [&'a str]) -> Result<Cli, usage::Error<'static, 'a>> {
    let os_strs: Vec<&OsStr> = args.iter().copied().map(OsStr::new).collect();
    Cli::try_parse_from(&os_strs)
}

#[test]
fn test_chart_edit_parses_setters_and_clear_flags() {
    let cli = try_parse(&[
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
    .expect("should parse a full chart edit invocation");

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
    let result = try_parse(&[
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
    let result = try_parse(&[
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
    // A numeric-looking column's filter values can legitimately start with
    // "-" (e.g. a negative Amount); `allow_hyphen_values` ensures leading
    // "-7" is treated as the value for `--values` rather than an unknown flag.
    let cli = try_parse(&[
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
    .expect("should accept hyphen-leading filter values");

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

    let mut initial_sheet = visi_core::core::Sheet::new(visi_core::core::SheetInit {
        id: None,
        name: Some("Sheet1".to_string()),
        rows: 3,
        cols: 2,
    });

    initial_sheet.set_cell_src(0, 0, "10".to_string());
    initial_sheet.set_cell_src(1, 0, "20".to_string());
    initial_sheet.set_cell_src(0, 1, "=A1 + A2".to_string());
    initial_sheet.set_cell_src(1, 1, "=SUM(A1:A2)".to_string());

    let bytes = visi_core::export_xlsx_data(&[initial_sheet], &[], &[], None).unwrap();
    fs::write(&file_path, bytes).unwrap();

    let mut wb = WorkbookManager::load_file(file_str).unwrap();
    wb.evaluate().unwrap();

    let sheet = &wb.sheets[0];
    let b1_val = sheet.get_result_data(&visi_core::core::CellRef::new(0, 1));
    let b2_val = sheet.get_result_data(&visi_core::core::CellRef::new(1, 1));

    assert_eq!(b1_val.to_string(), "30");
    assert_eq!(b2_val.to_string(), "30");

    wb.set_cell(0, 0, 0, "50".to_string());
    wb.evaluate().unwrap();

    let b1_val_updated = wb.sheets[0].get_result_data(&visi_core::core::CellRef::new(0, 1));
    let b2_val_updated = wb.sheets[0].get_result_data(&visi_core::core::CellRef::new(1, 1));

    assert_eq!(b1_val_updated.to_string(), "70");
    assert_eq!(b2_val_updated.to_string(), "70");

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
    let total = wb.sheets[0].get_result_data(&visi_core::core::CellRef::new(0, 2));
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
        .get_src(&visi_core::core::CellRef::new(0, 2))
        .cloned();
    assert_eq!(src_after_col_rename.as_deref(), Some("=SUM(Sales[Total])"));
    let total_after_col_rename = wb.sheets[0].get_result_data(&visi_core::core::CellRef::new(0, 2));
    assert_eq!(total_after_col_rename.to_string(), "30");

    wb.rename_table("Sales", "Revenue").unwrap();
    assert!(wb.find_table("Sales").is_none());
    assert!(wb.find_table("Revenue").is_some());

    let src_after_rename = wb.sheets[0]
        .get_src(&visi_core::core::CellRef::new(0, 2))
        .cloned();
    assert_eq!(src_after_rename.as_deref(), Some("=SUM(Revenue[Total])"));
    let total_after_rename = wb.sheets[0].get_result_data(&visi_core::core::CellRef::new(0, 2));
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
    let total_after_reload =
        reloaded.sheets[0].get_result_data(&visi_core::core::CellRef::new(0, 2));
    assert_eq!(total_after_reload.to_string(), "30");

    reloaded.delete_table("Revenue").unwrap();
    assert!(reloaded.list_tables().is_empty());

    let _ = table_id;
    let _ = fs::remove_file(file_path);
}

#[test]
fn test_workbook_vba_crud_and_roundtrip() {
    use visi_core::core::VbaModuleKind;

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
    use visi_core::core::VbaModuleKind;

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

    // Adding ThisWorkbook does not set a bound_sheet_id, so that Sheet1's
    // real code-behind module add succeeds.
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
    use visi_core::core::{PivotAggregation, PivotArea};

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
    let east = wb.sheets[0].get_result_data(&visi_core::core::CellRef::new(1, 5));
    let west = wb.sheets[0].get_result_data(&visi_core::core::CellRef::new(2, 5));
    let grand_total = wb.sheets[0].get_result_data(&visi_core::core::CellRef::new(3, 5));
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
    let east_count = reloaded.sheets[0].get_result_data(&visi_core::core::CellRef::new(2, 6));
    assert_eq!(east_count.to_string(), "2");

    reloaded.delete_pivot_table("SalesPivot").unwrap();
    assert!(reloaded.find_pivot_table("SalesPivot").is_none());

    let _ = pivot_id;
    let _ = fs::remove_file(file_path);
}

#[test]
fn test_getpivotdata_formula_resolves_against_rendered_pivot() {
    use visi_core::core::{PivotAggregation, PivotArea};

    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("test_getpivotdata.xlsx");
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

    // Destination anchored at column E (0-based col 4): "Row Labels" header
    // at row 0, East/West/Grand Total rows follow, values land in column F.
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

    // A formula elsewhere on the sheet, pointing at a cell inside the
    // pivot's rendered output (F2, the East row's value cell), should
    // resolve GETPIVOTDATA against the same pivot table without needing
    // its name.
    wb.set_cell(
        0,
        0,
        8,
        "=GETPIVOTDATA(\"Amount\", F2, \"Region\", \"East\")".to_string(),
    );
    wb.set_cell(0, 1, 8, "=GETPIVOTDATA(\"Amount\", F2)".to_string());
    wb.set_cell(
        0,
        2,
        8,
        "=GETPIVOTDATA(\"Amount\", F2, \"Region\", \"North\")".to_string(),
    );
    wb.evaluate().unwrap();

    let east = wb.sheets[0].get_result_data(&visi_core::core::CellRef::new(0, 8));
    assert_eq!(east.to_string(), "15");

    // No field/item criteria at all means the grand total.
    let grand_total = wb.sheets[0].get_result_data(&visi_core::core::CellRef::new(1, 8));
    assert_eq!(grand_total.to_string(), "85");

    // An item that doesn't exist in the pivot is a #REF! error, matching
    // real Excel's GETPIVOTDATA.
    let bad_item = wb.sheets[0].get_result_data(&visi_core::core::CellRef::new(2, 8));
    assert!(matches!(bad_item, visi_core::core::ResultData::Error(ref e) if e == "#REF!"));

    let _ = fs::remove_file(file_path);
}

#[test]
fn test_workbook_chart_edit_and_round_trip() {
    use visi_core::core::chart::ChartType;

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
    use visi_core::core::{PivotAggregation, PivotArea};

    // Matches real Excel's PivotField.Orientation semantics (verified via
    // the win32com fuzz driver): a field can only live in one of
    // Row/Column/Filter at a time, so re-adding it to a different area
    // moves it rather than leaving it in both.
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
    use visi_core::core::{CellRef, PivotAggregation, PivotArea};

    // Verified against real Excel: a filter/page field always renders as
    // its own row above the row/col header grid ("FieldName" | "(All)" or
    // "(Multiple Items)"), followed by one blank spacer row.
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
fn test_sheet_function_reports_real_ordinal_across_workbook_manager_evaluate() {
    // SHEET() reports the sheet's 1-based position when evaluated across
    // WorkbookManager::evaluate. Exercised through WorkbookManager::evaluate
    // itself, not just a hand-built Context, since that's the actual
    // production code path that populates it.
    let mut wb = WorkbookManager {
        sheets: Vec::new(),
        charts: Vec::new(),
        pivot_tables: Vec::new(),
        vba_project: None,
        locale: visi_core::core::Locale::default(),
    };
    wb.add_sheet("First").unwrap();
    wb.add_sheet("Second").unwrap();
    wb.add_sheet("Third").unwrap();

    wb.sheets[0].set_cell_src(0, 0, "=SHEET()".to_string());
    wb.sheets[2].set_cell_src(0, 0, "=SHEET()".to_string());
    // A cross-sheet reference from sheet 1 reports the *referenced*
    // sheet's ordinal.
    wb.sheets[0].set_cell_src(0, 1, "=SHEET(Third!A1)".to_string());

    wb.evaluate().unwrap();

    assert_eq!(
        wb.sheets[0]
            .get_result_data(&visi_core::core::CellRef::new(0, 0))
            .to_string(),
        "1"
    );
    assert_eq!(
        wb.sheets[2]
            .get_result_data(&visi_core::core::CellRef::new(0, 0))
            .to_string(),
        "3"
    );
    assert_eq!(
        wb.sheets[0]
            .get_result_data(&visi_core::core::CellRef::new(0, 1))
            .to_string(),
        "3"
    );
}

#[test]
fn test_evaluate_resolves_a_two_hop_cross_sheet_dependency_chain() {
    // WorkbookManager::evaluate re-marks dirty cells across multiple
    // passes so that a cross-sheet chain more than one hop deep
    // (First's formula -> Second's formula -> First's formula again)
    // resolves correctly.
    let mut wb = WorkbookManager {
        sheets: Vec::new(),
        charts: Vec::new(),
        pivot_tables: Vec::new(),
        vba_project: None,
        locale: visi_core::core::Locale::default(),
    };
    wb.add_sheet("First").unwrap();
    wb.add_sheet("Second").unwrap();

    wb.sheets[0].set_cell_src(0, 0, "10".to_string());
    wb.sheets[0].set_cell_src(0, 1, "=A1*2".to_string());
    wb.sheets[1].set_cell_src(0, 0, "=First!B1+1".to_string());
    // This is the cell that needs a *second* pass: it depends on
    // Second!A1, which itself only becomes correct after First!B1 is
    // computed earlier in the very same first pass.
    wb.sheets[0].set_cell_src(0, 2, "=Second!A1*3".to_string());

    wb.evaluate().unwrap();

    assert_eq!(
        wb.sheets[0]
            .get_result_data(&visi_core::core::CellRef::new(0, 1))
            .to_string(),
        "20"
    );
    assert_eq!(
        wb.sheets[1]
            .get_result_data(&visi_core::core::CellRef::new(0, 0))
            .to_string(),
        "21"
    );
    assert_eq!(
        wb.sheets[0]
            .get_result_data(&visi_core::core::CellRef::new(0, 2))
            .to_string(),
        "63"
    );
}

#[test]
fn test_cross_sheet_circular_reference_terminates_without_hanging() {
    // Cross-sheet circular reference test: cross-sheet cycle evaluation is
    // naturally bounded by the fixed pass count.
    let mut wb = WorkbookManager {
        sheets: Vec::new(),
        charts: Vec::new(),
        pivot_tables: Vec::new(),
        vba_project: None,
        locale: visi_core::core::Locale::default(),
    };
    wb.add_sheet("First").unwrap();
    wb.add_sheet("Second").unwrap();
    wb.sheets[0].set_cell_src(0, 0, "=Second!A1+1".to_string());
    wb.sheets[1].set_cell_src(0, 0, "=First!A1+1".to_string());

    let start = std::time::Instant::now();
    wb.evaluate().unwrap();
    assert!(
        start.elapsed().as_secs() < 5,
        "a cross-sheet cycle must not hang"
    );

    for (sheet_idx, label) in [(0, "First!A1"), (1, "Second!A1")] {
        match wb.sheets[sheet_idx].get_result_data(&visi_core::core::CellRef::new(0, 0)) {
            visi_core::core::ResultData::Float(f) => assert!(f.is_finite(), "{label} not finite"),
            visi_core::core::ResultData::Integer(_) => {}
            other => panic!("expected a finite numeric result for {label}, got {other:?}"),
        }
    }
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

#[test]
fn test_cell_style_setting_and_xlsx_round_trip() {
    use visi_core::core::CellStyle;

    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("test_cell_styles.xlsx");
    let file_str = file_path.to_str().unwrap();

    let mut wb = WorkbookManager::load_file_or_create(file_str).unwrap();

    wb.set_cell(0, 0, 0, "Styled Header".to_string());
    // 0-based (row, col): A1.
    wb.set_cell_style(
        None,
        0,
        0,
        CellStyle {
            font_color: Some("#FF0000".to_string()),
            bg_color: Some("#FFFF00".to_string()),
            bold: Some(true),
            italic: Some(true),
            font_family: Some("Arial".to_string()),
            font_size: Some(14.0),
            ..Default::default()
        },
    )
    .unwrap();

    // Set range style on B1:B3, as 0-based (start_row, start_col, end_row, end_col).
    wb.set_range_style(
        None,
        0,
        1,
        2,
        1,
        CellStyle {
            bold: Some(true),
            font_color: Some("blue".to_string()),
            ..Default::default()
        },
    )
    .unwrap();

    let style_a1 = wb.get_cell_style(None, 0, 0).unwrap().unwrap();
    assert_eq!(style_a1.font_color, Some("#FF0000".to_string()));
    assert_eq!(style_a1.bg_color, Some("#FFFF00".to_string()));
    assert_eq!(style_a1.bold, Some(true));
    assert_eq!(style_a1.italic, Some(true));
    assert_eq!(style_a1.font_family, Some("Arial".to_string()));
    assert_eq!(style_a1.font_size, Some(14.0));

    let style_b2 = wb.get_cell_style(None, 1, 1).unwrap().unwrap();
    assert_eq!(style_b2.bold, Some(true));
    assert_eq!(style_b2.font_color, Some("blue".to_string()));

    wb.save_file(file_str).unwrap();

    let _ = fs::remove_file(file_path);
}

#[test]
fn test_style_cell_ref_sheet_prefix_overrides_sheet_flag() {
    // `visi style cell --cell 'Sheet2!C3' --sheet Sheet1` must style C3 on
    // *Sheet2*: an explicit prefix in the reference beats the --sheet flag.
    use visi_core::core::CellStyle;

    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("test_style_sheet_prefix.xlsx");
    let file_str = file_path.to_str().unwrap();
    let _ = fs::remove_file(&file_path);

    let mut wb = WorkbookManager::load_file_or_create(file_str).unwrap();
    wb.add_sheet("Sheet2").unwrap();

    // What main.rs does for `--cell 'Sheet2!C3' --sheet Sheet1`.
    let (specified_sheet, row, col) = visi::utils::parse_cell_ref("Sheet2!C3").unwrap();
    assert_eq!(specified_sheet.as_deref(), Some("Sheet2"));
    assert_eq!((row, col), (2, 2));
    let sheet = specified_sheet.as_deref().or(Some("Sheet1"));

    wb.set_cell_style(
        sheet,
        row,
        col,
        CellStyle {
            bold: Some(true),
            ..Default::default()
        },
    )
    .unwrap();

    // Applied to Sheet2, not the --sheet flag's Sheet1.
    assert_eq!(
        wb.get_cell_style(Some("Sheet2"), 2, 2).unwrap(),
        Some(CellStyle {
            bold: Some(true),
            ..Default::default()
        })
    );
    assert_eq!(wb.get_cell_style(Some("Sheet1"), 2, 2).unwrap(), None);

    let _ = fs::remove_file(file_path);
}

#[test]
fn test_table_style_theme_setting_and_xlsx_round_trip() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("test_table_style.xlsx");
    let file_str = file_path.to_str().unwrap();

    let mut wb = WorkbookManager::load_file_or_create(file_str).unwrap();

    wb.set_cell(0, 0, 0, "ID".to_string());
    wb.set_cell(0, 0, 1, "Name".to_string());
    wb.set_cell(0, 1, 0, "1".to_string());
    wb.set_cell(0, 1, 1, "Alice".to_string());

    wb.add_table(None, "SalesTable", 0, 0, 1, 1, true, false)
        .unwrap();
    wb.set_table_style("SalesTable", "TableStyleMedium9")
        .unwrap();

    assert_eq!(
        wb.get_table_style("SalesTable").unwrap(),
        Some("TableStyleMedium9".to_string())
    );

    wb.save_file(file_str).unwrap();
    let reloaded = WorkbookManager::load_file(file_str).unwrap();

    assert_eq!(
        reloaded.get_table_style("SalesTable").unwrap(),
        Some("TableStyleMedium9".to_string())
    );

    let _ = fs::remove_file(file_path);
}

/// Syntax checking through the same API `visi macro check` calls, including
/// the round trip that matters: a module's source only survives to be checked
/// in a later invocation because it goes through `vbaProject.bin`.
#[test]
fn test_vba_syntax_check_through_a_real_roundtrip() {
    use visi_core::core::{VbaModuleKind, check_syntax};

    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("test_vba_check.xlsm");
    let file_str = file_path.to_str().unwrap();

    let good = "Attribute VB_Name = \"Good\"\n\
                Public Sub Alpha()\n    Dim x As Long\n    x = -2 ^ 2\nEnd Sub\n\
                Public Function Beta() As String\n    Beta = \"b\"\nEnd Function\n";
    // Valid enough for `macro add` to store -- nothing validates syntax on
    // the way in, which is exactly the gap `macro check` fills.
    let bad = "Attribute VB_Name = \"Bad\"\n\
               Public Sub Broken()\n    If x Then\nEnd Sub\n";

    let mut wb = WorkbookManager::new_empty().unwrap();
    wb.add_vba_module(
        "Good".to_string(),
        VbaModuleKind::Standard,
        good.to_string(),
        None,
    )
    .unwrap();
    wb.add_vba_module(
        "Bad".to_string(),
        VbaModuleKind::Standard,
        bad.to_string(),
        None,
    )
    .unwrap();
    wb.save_file(file_str).unwrap();

    let reloaded = WorkbookManager::load_file(file_str).unwrap();
    let modules = reloaded.list_vba_modules();
    assert_eq!(modules.len(), 2);

    let good_mod = modules.iter().find(|m| m.name == "Good").unwrap();
    let syntax = good_mod.check_syntax().expect("valid module should parse");
    assert_eq!(syntax.procedures, vec!["Alpha", "Beta"]);

    let bad_mod = modules.iter().find(|m| m.name == "Bad").unwrap();
    match bad_mod.check_syntax() {
        Err(visi_core::Error::VbaSyntax {
            module,
            line,
            column,
            ..
        }) => {
            // The module name is what makes a multi-module report readable.
            assert_eq!(module.as_deref(), Some("Bad"));
            // Line 3 -- the unclosed `If` itself, matching how VBA words and
            // places the error, rather than the perfectly correct `End Sub`
            // on line 4 that merely arrived where `End If` was due.
            assert_eq!(line, 3);
            assert!(column >= 1);
        }
        other => panic!("expected a syntax error, got {other:?}"),
    }

    // The bare-source entry point carries no module name, since there is none.
    assert!(check_syntax(good).is_ok());
    match check_syntax(bad) {
        Err(visi_core::Error::VbaSyntax { module, .. }) => assert!(module.is_none()),
        other => panic!("expected a syntax error, got {other:?}"),
    }

    let _ = fs::remove_file(file_path);
}

/// A workbook-bound macro run, through the same `WorkbookManager::run_macro`
/// the CLI handler calls, over a real `.xlsm` round trip.
///
/// The round trip is the point rather than incidental. The CLI is a fresh
/// process per invocation, so a macro only survives to be *run* in a later
/// command because it went out through `vbaProject.bin` and came back --
/// exactly the property `test_vba_syntax_check_through_a_real_roundtrip`
/// covers for checking, now that running can change the file too.
#[test]
fn test_vba_macro_run_reads_and_writes_a_real_workbook() {
    use visi_core::core::{CellRef, VbaModuleKind};

    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("test_vba_run.xlsm");
    let out_path = temp_dir.join("test_vba_run_out.xlsm");
    let (file_str, out_str) = (file_path.to_str().unwrap(), out_path.to_str().unwrap());

    let source = "Attribute VB_Name = \"Demo\"\n\
        Public Function Total() As Variant\n\
        \x20   Dim ws As Worksheet\n\
        \x20   Set ws = ThisWorkbook.Worksheets(\"Sheet1\")\n\
        \x20   Dim c As Range, running As Double\n\
        \x20   For Each c In ws.Range(\"A1:A3\")\n\
        \x20       running = running + c.Value\n\
        \x20   Next c\n\
        \x20   ws.Range(\"C1\").Value = running\n\
        \x20   ws.Range(\"C2\").Formula = \"=C1*2\"\n\
        \x20   Total = ws.Range(\"C2\").Value\n\
        End Function\n\
        Public Function JustLooks() As Variant\n\
        \x20   JustLooks = ThisWorkbook.Worksheets.Count\n\
        End Function\n";

    let mut wb = WorkbookManager::new_empty().unwrap();
    for (row, v) in [1, 2, 3].into_iter().enumerate() {
        wb.set_cell(0, row, 0, v.to_string());
    }
    wb.evaluate().unwrap();
    wb.add_vba_module(
        "Demo".to_string(),
        VbaModuleKind::Standard,
        source.to_string(),
        None,
    )
    .unwrap();
    wb.save_file(file_str).unwrap();

    // A second process would see exactly this: the module read back out of
    // the binary VBA part, with no in-memory state carried over.
    let mut reloaded = WorkbookManager::load_file(file_str).unwrap();

    // A read-only macro reports no mutation, which is what lets the CLI
    // accept it without an --output.
    let looked = reloaded.run_macro(None, "JustLooks", &[]).unwrap();
    assert_eq!(looked.value.as_deref(), Some("1"));
    assert!(!looked.mutated);

    // The writing one sees its own formula recalculated, exactly as Excel in
    // automatic mode would.
    let total = reloaded.run_macro(None, "Total", &[]).unwrap();
    assert_eq!(total.value.as_deref(), Some("12"));
    assert!(total.mutated);

    reloaded.save_file(out_str).unwrap();
    let saved = WorkbookManager::load_file(out_str).unwrap();
    assert_eq!(saved.sheets[0].get_display_string(&CellRef::new(0, 2)), "6");
    assert_eq!(
        saved.sheets[0].get_display_string(&CellRef::new(1, 2)),
        "12"
    );

    let _ = fs::remove_file(file_path);
    let _ = fs::remove_file(out_path);
}

/// `visi macro run` takes the same write flags as every other write command.
#[test]
fn test_macro_run_parses_output_and_in_place() {
    let cli = try_parse(&[
        "visi",
        "macro",
        "run",
        "book.xlsm",
        "--name",
        "Go",
        "--module",
        "M",
        "-a",
        "1",
        "-a",
        "two",
        "--output",
        "done.xlsm",
        "--json",
    ])
    .expect("should parse a macro run with a write target");

    let Commands::Macro(macro_args) = cli.command else {
        panic!("expected Commands::Macro");
    };
    let MacroSubcommands::Run(run_args) = macro_args.command else {
        panic!("expected MacroSubcommands::Run");
    };
    assert_eq!(run_args.name, "Go");
    assert_eq!(run_args.module.as_deref(), Some("M"));
    assert_eq!(run_args.args, vec!["1", "two"]);
    assert_eq!(run_args.output.as_deref(), Some("done.xlsm"));
    assert!(!run_args.in_place);
    assert!(run_args.json);

    // --output and --in-place are mutually exclusive, as they are everywhere
    // else in the CLI.
    assert!(
        try_parse(&[
            "visi",
            "macro",
            "run",
            "book.xlsm",
            "--name",
            "Go",
            "--output",
            "a.xlsm",
            "-i",
        ])
        .is_err()
    );
}

// --------------------------------------------------------------------------
// Structural edits: what a row or column insert/delete does to everything
// that holds a coordinate.
//
// Every expectation below is Excel's own behaviour. The cases that look
// arbitrary -- a `$` not pinning a reference, an insert at a range's first
// row moving the range while an insert one row lower grows it -- are the ones
// most worth having written down.
// --------------------------------------------------------------------------

/// A workbook with one sheet of the given cell sources, round-tripped through
/// a real `.xlsx` so the test exercises import and export too.
fn workbook_from(
    name: &str,
    cells: &[(usize, usize, &str)],
    rows: usize,
    cols: usize,
) -> WorkbookManager {
    let mut sheet = visi_core::core::Sheet::new(visi_core::core::SheetInit {
        id: None,
        name: Some("Sheet1".to_string()),
        rows,
        cols,
    });
    for (row, col, src) in cells {
        sheet.set_cell_src(*row, *col, (*src).to_string());
    }
    let bytes = visi_core::export_xlsx_data(&[sheet], &[], &[], None).unwrap();
    let path = std::env::temp_dir().join(name);
    fs::write(&path, bytes).unwrap();
    let wb = WorkbookManager::load_file(path.to_str().unwrap()).unwrap();
    let _ = fs::remove_file(&path);
    wb
}

/// The raw source text of a cell -- what a reference shift actually rewrites.
/// The computed value can agree by coincidence; the text cannot.
fn cell_src(wb: &WorkbookManager, sheet: usize, row: usize, col: usize) -> String {
    wb.sheets[sheet].columns()[col]
        .src(row)
        .unwrap_or_default()
        .to_string()
}

fn cell_value(wb: &WorkbookManager, sheet: usize, row: usize, col: usize) -> String {
    wb.sheets[sheet]
        .get_result_data(&visi_core::core::CellRef::new(row, col))
        .to_string()
}

#[test]
fn test_insert_row_shifts_references_the_way_excel_does() {
    // C1 =A3, C2 =SUM(A1:A3), C3 =SUM(A3:A4), C4 =$A$3*10
    let mut wb = workbook_from(
        "test_insert_row_shift.xlsx",
        &[
            (0, 0, "1"),
            (1, 0, "2"),
            (2, 0, "3"),
            (0, 2, "=A3"),
            (1, 2, "=SUM(A1:A3)"),
            (2, 2, "=SUM(A3:A4)"),
            (3, 2, "=$A$3*10"),
        ],
        4,
        3,
    );
    wb.evaluate().unwrap();

    // Insert above everything. The formulas move down a row *and* their
    // references move down a row.
    wb.insert_row(0, 0).unwrap();

    assert_eq!(cell_src(&wb, 0, 1, 2), "=A4");
    // The whole range was at or below the insert, so it moves rather than
    // growing.
    assert_eq!(cell_src(&wb, 0, 2, 2), "=SUM(A2:A4)");
    assert_eq!(cell_src(&wb, 0, 3, 2), "=SUM(A4:A5)");
    // A `$` governs fill and copy, not a structural edit: $A$3 shifts too.
    assert_eq!(cell_src(&wb, 0, 4, 2), "=$A$4*10");

    // And the values are the ones they had before the insert.
    assert_eq!(cell_value(&wb, 0, 1, 2), "3");
    assert_eq!(cell_value(&wb, 0, 2, 2), "6");
    assert_eq!(cell_value(&wb, 0, 4, 2), "30");
}

#[test]
fn test_insert_row_inside_a_range_grows_it_instead_of_moving_it() {
    let mut wb = workbook_from(
        "test_insert_row_grows.xlsx",
        &[(0, 0, "1"), (1, 0, "2"), (2, 0, "3"), (0, 2, "=SUM(A1:A3)")],
        3,
        3,
    );
    wb.evaluate().unwrap();

    // Insert at row 2 (0-based 1), which is inside A1:A3.
    wb.insert_row(0, 1).unwrap();
    assert_eq!(cell_src(&wb, 0, 0, 2), "=SUM(A1:A4)");
    // The new row is empty, so the sum is unchanged -- the range grew to keep
    // covering the same three values.
    assert_eq!(cell_value(&wb, 0, 0, 2), "6");
}

#[test]
fn test_delete_row_breaks_only_the_reference_that_pointed_at_it() {
    let mut wb = workbook_from(
        "test_delete_row_ref.xlsx",
        &[
            (0, 0, "1"),
            (1, 0, "2"),
            (2, 0, "3"),
            (0, 2, "=A3"),
            (1, 2, "=SUM(A1:A3)"),
            (2, 2, "=A3+1"),
        ],
        3,
        3,
    );
    wb.evaluate().unwrap();

    wb.delete_row(0, 2).unwrap();

    // The reference is replaced, not the formula, so the rest still
    // evaluates -- to #REF!, by ordinary error propagation.
    assert_eq!(cell_src(&wb, 0, 0, 2), "=#REF!");
    assert_eq!(cell_value(&wb, 0, 0, 2), "Error: #REF!");
    // A range that only partly overlapped the deleted row shrinks.
    assert_eq!(cell_src(&wb, 0, 1, 2), "=SUM(A1:A2)");
    assert_eq!(cell_value(&wb, 0, 1, 2), "3");
    // C3 held `=A3+1` and was itself in the deleted row, so it is gone.
    assert_eq!(wb.sheets[0].row_count(), 2);
}

#[test]
fn test_a_ref_error_survives_an_xlsx_round_trip() {
    let mut wb = workbook_from(
        "test_ref_error_roundtrip.xlsx",
        &[(2, 0, "3"), (0, 2, "=A3"), (1, 2, "=A3+1")],
        3,
        3,
    );
    wb.evaluate().unwrap();
    wb.delete_row(0, 2).unwrap();

    let out = std::env::temp_dir().join("test_ref_error_roundtrip_out.xlsx");
    wb.save_file(out.to_str().unwrap()).unwrap();
    let mut reloaded = WorkbookManager::load_file(out.to_str().unwrap()).unwrap();
    reloaded.evaluate().unwrap();
    let _ = fs::remove_file(&out);

    // A `#REF!` written into a formula has to lex on the way back in, or the
    // reload turns a broken reference into a broken *formula*.
    assert_eq!(cell_src(&reloaded, 0, 0, 2), "=#REF!");
    assert_eq!(cell_value(&reloaded, 0, 0, 2), "Error: #REF!");
    assert_eq!(cell_value(&reloaded, 0, 1, 2), "Error: #REF!");
}

#[test]
fn test_column_edits_shift_references_including_whole_column_ones() {
    let mut wb = workbook_from(
        "test_col_shift.xlsx",
        &[
            (0, 0, "1"),
            (0, 1, "2"),
            (0, 2, "3"),
            (0, 4, "=C1"),
            (1, 4, "=SUM(A1:C1)"),
            (2, 4, "=SUM(B:B)"),
        ],
        3,
        5,
    );
    wb.evaluate().unwrap();

    // Insert a column at B. A whole-column reference is held by column id, so
    // it has to be re-rendered *after* the edit or it would keep the letter
    // `B`, which now names a different column.
    wb.insert_col(0, 1).unwrap();
    assert_eq!(cell_src(&wb, 0, 0, 5), "=D1");
    assert_eq!(cell_src(&wb, 0, 1, 5), "=SUM(A1:D1)");
    assert_eq!(cell_src(&wb, 0, 2, 5), "=SUM(C:C)");
    assert_eq!(cell_value(&wb, 0, 2, 5), "2");
}

#[test]
fn test_delete_column_breaks_the_references_that_named_it() {
    let mut wb = workbook_from(
        "test_col_delete.xlsx",
        &[
            (0, 0, "1"),
            (0, 1, "2"),
            (0, 2, "3"),
            (0, 4, "=B1"),
            (1, 4, "=SUM(A1:C1)"),
            (2, 4, "=SUM(B:B)"),
        ],
        3,
        5,
    );
    wb.evaluate().unwrap();

    wb.delete_col(0, 1).unwrap();
    assert_eq!(cell_src(&wb, 0, 0, 3), "=#REF!");
    assert_eq!(cell_src(&wb, 0, 1, 3), "=SUM(A1:B1)");
    // The whole column it named is gone, and nothing about the remaining
    // grid stands in for it.
    assert_eq!(cell_src(&wb, 0, 2, 3), "=SUM(#REF!)");
    assert_eq!(cell_value(&wb, 0, 2, 3), "Error: #REF!");
}

#[test]
fn test_a_structural_edit_shifts_references_from_other_sheets_only() {
    let mut sheet1 = visi_core::core::Sheet::new(visi_core::core::SheetInit {
        id: None,
        name: Some("Sheet1".to_string()),
        rows: 6,
        cols: 1,
    });
    sheet1.set_cell_src(0, 0, "=Data!A3".to_string());
    sheet1.set_cell_src(1, 0, "=SUM(Data!A1:A3)".to_string());
    sheet1.set_cell_src(2, 0, "=SUM(Sheet1!A5:A6)".to_string());

    let mut data = visi_core::core::Sheet::new(visi_core::core::SheetInit {
        id: None,
        name: Some("Data".to_string()),
        rows: 3,
        cols: 1,
    });
    data.set_cell_src(0, 0, "10".to_string());
    data.set_cell_src(1, 0, "20".to_string());
    data.set_cell_src(2, 0, "30".to_string());

    let bytes = visi_core::export_xlsx_data(&[sheet1, data], &[], &[], None).unwrap();
    let path = std::env::temp_dir().join("test_cross_sheet_shift.xlsx");
    fs::write(&path, bytes).unwrap();
    let mut wb = WorkbookManager::load_file(path.to_str().unwrap()).unwrap();
    let _ = fs::remove_file(&path);
    wb.evaluate().unwrap();
    assert_eq!(cell_value(&wb, 0, 0, 0), "30");

    let data_idx = wb.sheets.iter().position(|s| s.name == "Data").unwrap();
    wb.insert_row(data_idx, 0).unwrap();

    // References *into* the edited sheet move.
    assert_eq!(cell_src(&wb, 0, 0, 0), "=Data!A4");
    assert_eq!(cell_src(&wb, 0, 1, 0), "=SUM(Data!A2:A4)");
    // A reference to the sheet that was *not* edited must not.
    assert_eq!(cell_src(&wb, 0, 2, 0), "=SUM(Sheet1!A5:A6)");
    assert_eq!(cell_value(&wb, 0, 0, 0), "30");
    assert_eq!(cell_value(&wb, 0, 1, 0), "60");
}

#[test]
fn test_a_table_extent_follows_a_structural_edit() {
    let mut wb = workbook_from(
        "test_table_extent_shift.xlsx",
        &[
            (0, 0, "Region"),
            (0, 1, "Amount"),
            (1, 0, "East"),
            (1, 1, "10"),
            (2, 0, "West"),
            (2, 1, "20"),
            (3, 0, "North"),
            (3, 1, "30"),
            (0, 3, "=SUM(Sales[Amount])"),
        ],
        4,
        4,
    );
    wb.add_table(None, "Sales", 0, 0, 3, 1, true, false)
        .unwrap();
    wb.evaluate().unwrap();
    assert_eq!(cell_value(&wb, 0, 0, 3), "60");

    // A row inserted inside the table body grows the table, so the structured
    // reference keeps covering the same data.
    wb.insert_row(0, 2).unwrap();
    let table = wb.sheets[0].find_table("Sales").unwrap();
    assert_eq!((table.start_row, table.end_row), (0, 4));
    assert_eq!(cell_value(&wb, 0, 0, 3), "60");

    // A column inserted inside it widens it, and the column names stay lined
    // up with the sheet columns they describe.
    wb.insert_col(0, 1).unwrap();
    let table = wb.sheets[0].find_table("Sales").unwrap();
    assert_eq!((table.start_col, table.end_col), (0, 2));
    assert_eq!(table.columns.len(), 3);
    assert_eq!(table.columns[0], "Region");
    assert_eq!(table.columns[2], "Amount");
    assert_eq!(cell_value(&wb, 0, 0, 4), "60");
}

#[test]
fn test_deleting_every_row_of_a_table_removes_the_table() {
    let mut wb = workbook_from(
        "test_table_fully_deleted.xlsx",
        &[(2, 0, "Region"), (2, 1, "Amount")],
        3,
        2,
    );
    wb.add_table(None, "Solo", 2, 0, 2, 1, true, false).unwrap();
    wb.delete_row(0, 2).unwrap();
    // Nothing is left for the table to describe, so it goes rather than
    // lingering over cells that are now something else.
    assert!(wb.sheets[0].find_table("Solo").is_none());
}

#[test]
fn test_a_pivot_source_range_and_destination_follow_a_structural_edit() {
    use visi_core::core::{PivotAggregation, PivotArea, PivotSource};

    let mut wb = workbook_from(
        "test_pivot_range_shift.xlsx",
        &[
            (0, 0, "Region"),
            (0, 1, "Amount"),
            (1, 0, "East"),
            (1, 1, "10"),
            (2, 0, "West"),
            (2, 1, "20"),
            (3, 0, "East"),
            (3, 1, "30"),
        ],
        4,
        6,
    );
    wb.evaluate().unwrap();
    wb.add_pivot_table_from_range("P1", None, 0, 0, 3, 1, None, 0, 4, true, true)
        .unwrap();
    wb.add_pivot_field("P1", PivotArea::Row, "Region", None)
        .unwrap();
    wb.add_pivot_field(
        "P1",
        PivotArea::Value,
        "Amount",
        Some(PivotAggregation::Sum),
    )
    .unwrap();
    wb.refresh_pivot_table("P1").unwrap();
    assert_eq!(cell_value(&wb, 0, 1, 5), "40");

    // A row inserted above everything slides the source range and the
    // destination corner down together.
    wb.insert_row(0, 0).unwrap();

    let pivot = wb.find_pivot_table("P1").unwrap();
    let PivotSource::Range {
        start_row, end_row, ..
    } = pivot.source
    else {
        panic!("expected a range source");
    };
    assert_eq!((start_row, end_row), (1, 4));
    assert_eq!(pivot.dest_row, 1);

    // Nothing recomputes a pivot table on its own -- as in Excel -- so the
    // check that matters is that an explicit refresh still finds the data and
    // writes the grid where the destination moved to.
    wb.refresh_pivot_table("P1").unwrap();
    assert_eq!(cell_value(&wb, 0, 2, 5), "40");
    assert_eq!(cell_value(&wb, 0, 3, 5), "20");
}

#[test]
fn test_a_whole_column_range_does_not_overflow_on_a_row_edit() {
    // `A:C` compiles to a range whose end row is the unbounded sentinel.
    // Shifting it as an ordinary span overflows; it must be screened out.
    let mut wb = workbook_from(
        "test_unbounded_range_shift.xlsx",
        &[(0, 0, "1"), (0, 1, "2"), (0, 2, "3"), (0, 4, "=SUM(A:C)")],
        2,
        5,
    );
    wb.evaluate().unwrap();
    assert_eq!(cell_value(&wb, 0, 0, 4), "6");

    // A row edit leaves it alone: it already covers every row.
    wb.insert_row(0, 0).unwrap();
    assert_eq!(cell_src(&wb, 0, 1, 4), "=SUM(A:C)");
    assert_eq!(cell_value(&wb, 0, 1, 4), "6");
    wb.delete_row(0, 0).unwrap();
    assert_eq!(cell_src(&wb, 0, 0, 4), "=SUM(A:C)");

    // A column edit still moves its column bounds.
    wb.insert_col(0, 0).unwrap();
    assert_eq!(cell_src(&wb, 0, 0, 5), "=SUM(B:D)");
    assert_eq!(cell_value(&wb, 0, 0, 5), "6");
}

// --------------------------------------------------------------------------
// Insert / delete cells over a column band -- Excel's "shift down / up",
// which is what `ListRows.Add` actually is rather than a row insert.
//
// Measured with `fuzz/band_insert_probe.py`; the reference rules are the
// ordinary ones, applied only where the band covers the reference's columns.
// --------------------------------------------------------------------------

#[test]
fn test_band_insert_moves_only_the_bands_own_columns() {
    // The case that forced this operation to exist: adding a row to a table
    // at A1:C4 moves what is below it in A:C, and leaves what is beside it
    // alone. Both measured against Excel.
    let mut wb = workbook_from(
        "test_band_insert_moves.xlsx",
        &[(0, 0, "1"), (1, 0, "2"), (7, 0, "BELOW"), (1, 4, "BESIDE")],
        9,
        6,
    );
    wb.evaluate().unwrap();

    // Insert one row at row 4 (0-based) across A:C only.
    wb.insert_cells_shift_down(0, 4, 0, 2, 1).unwrap();

    // A8 "BELOW" was inside the band and below the insert: it moved to A9.
    assert_eq!(cell_value(&wb, 0, 8, 0), "BELOW");
    assert_eq!(cell_value(&wb, 0, 7, 0), "");
    // E2 "BESIDE" was outside the band: it did not move, even though it is
    // below the insert row in its own column.
    assert_eq!(cell_value(&wb, 0, 1, 4), "BESIDE");
    // The sheet stays rectangular -- every column the same length.
    let lengths: Vec<usize> = wb.sheets[0].columns().iter().map(|c| c.len()).collect();
    assert!(
        lengths.windows(2).all(|w| w[0] == w[1]),
        "columns desynced: {lengths:?}"
    );
}

#[test]
fn test_band_insert_shifts_formulas_only_inside_the_band() {
    let mut wb = workbook_from(
        "test_band_insert_formulas.xlsx",
        &[
            (4, 0, "5"),
            (4, 4, "50"),
            // H1..H5, all outside the A:C band so they survive to be read.
            (0, 7, "=A5"),
            (1, 7, "=E5"),
            (2, 7, "=SUM(A5:A6)"),
            (3, 7, "=SUM(A5:E5)"),
            (4, 7, "=SUM(A1:A6)"),
        ],
        10,
        8,
    );
    wb.evaluate().unwrap();

    wb.insert_cells_shift_down(0, 1, 0, 2, 1).unwrap();

    // Inside the band, below the insert: shifts.
    assert_eq!(cell_src(&wb, 0, 0, 7), "=A6");
    // Outside the band: untouched.
    assert_eq!(cell_src(&wb, 0, 1, 7), "=E5");
    // Wholly inside: shifts.
    assert_eq!(cell_src(&wb, 0, 2, 7), "=SUM(A6:A7)");
    // Straddling the band's edge: unchanged. It cannot both shift and not
    // shift, and Excel resolves that by leaving it alone.
    assert_eq!(cell_src(&wb, 0, 3, 7), "=SUM(A5:E5)");
    // Spanning the insert point, inside the band: grows rather than moves.
    assert_eq!(cell_src(&wb, 0, 4, 7), "=SUM(A1:A7)");
}

#[test]
fn test_band_delete_is_the_inverse_of_band_insert() {
    let mut wb = workbook_from(
        "test_band_delete.xlsx",
        &[(0, 0, "a"), (1, 0, "b"), (2, 0, "c"), (1, 4, "BESIDE")],
        5,
        6,
    );
    wb.evaluate().unwrap();

    wb.delete_cells_shift_up(0, 0, 0, 2, 1).unwrap();

    // Column A moved up; E did not.
    assert_eq!(cell_value(&wb, 0, 0, 0), "b");
    assert_eq!(cell_value(&wb, 0, 1, 0), "c");
    assert_eq!(cell_value(&wb, 0, 1, 4), "BESIDE");
    let lengths: Vec<usize> = wb.sheets[0].columns().iter().map(|c| c.len()).collect();
    assert!(
        lengths.windows(2).all(|w| w[0] == w[1]),
        "columns desynced: {lengths:?}"
    );
}

#[test]
fn test_a_band_insert_keeps_styles_with_their_cells() {
    // The invariant that has been broken before: `src`, `data`,
    // `compiled_src` and `styles` have to move together, or a style ends up
    // on the wrong row.
    let mut wb = workbook_from(
        "test_band_insert_styles.xlsx",
        &[(3, 0, "styled"), (3, 4, "beside")],
        6,
        6,
    );
    wb.sheets[0].update_cell_style(3, 0, |s| s.bold = Some(true));
    wb.sheets[0].update_cell_style(3, 4, |s| s.italic = Some(true));
    wb.evaluate().unwrap();

    wb.insert_cells_shift_down(0, 1, 0, 2, 1).unwrap();

    // The styled cell in the band moved down with its style.
    assert_eq!(cell_value(&wb, 0, 4, 0), "styled");
    assert_eq!(
        wb.sheets[0].get_cell_style(4, 0).and_then(|s| s.bold),
        Some(true)
    );
    assert!(wb.sheets[0].get_cell_style(3, 0).is_none());
    // The one outside the band stayed put, style and all.
    assert_eq!(
        wb.sheets[0].get_cell_style(3, 4).and_then(|s| s.italic),
        Some(true)
    );
}

#[test]
fn test_a_band_insert_grows_a_table_it_passes_through() {
    let mut wb = workbook_from(
        "test_band_insert_table.xlsx",
        &[
            (0, 0, "Region"),
            (0, 1, "Amount"),
            (1, 0, "East"),
            (1, 1, "10"),
            (2, 0, "West"),
            (2, 1, "20"),
            (0, 4, "=SUM(Sales[Amount])"),
        ],
        6,
        6,
    );
    wb.add_table(None, "Sales", 0, 0, 2, 1, true, false)
        .unwrap();
    wb.evaluate().unwrap();
    assert_eq!(cell_value(&wb, 0, 0, 4), "30");

    // Insert inside the table's own columns and rows: the table grows.
    wb.insert_cells_shift_down(0, 2, 0, 1, 1).unwrap();
    let table = wb.sheets[0].find_table("Sales").unwrap();
    assert_eq!((table.start_row, table.end_row), (0, 3));
    // The structured reference still covers the same data.
    assert_eq!(cell_value(&wb, 0, 0, 4), "30");
}

#[test]
fn test_a_band_insert_outside_a_tables_columns_leaves_it_alone() {
    let mut wb = workbook_from(
        "test_band_insert_other_table.xlsx",
        &[
            (0, 3, "Region"),
            (0, 4, "Amount"),
            (1, 3, "East"),
            (1, 4, "10"),
        ],
        6,
        6,
    );
    wb.add_table(None, "Side", 0, 3, 1, 4, true, false).unwrap();
    wb.evaluate().unwrap();

    // The band is A:C; the table lives in D:E and must not move.
    wb.insert_cells_shift_down(0, 0, 0, 2, 1).unwrap();
    let table = wb.sheets[0].find_table("Side").unwrap();
    assert_eq!((table.start_row, table.end_row), (0, 1));
    assert_eq!((table.start_col, table.end_col), (3, 4));
}

#[test]
fn test_cli_set_cell_type_flag_parsing() {
    let cli = try_parse(&[
        "visi",
        "set",
        "data.xlsx",
        "--cell",
        "A1",
        "--value",
        "12345",
        "--type",
        "string",
        "-i",
    ])
    .expect("should parse set with --type string");

    let Commands::Set(set_args) = cli.command else {
        panic!("expected Commands::Set");
    };
    assert_eq!(set_args.cell, vec!["A1"]);
    assert_eq!(set_args.value, vec!["12345"]);
    assert_eq!(set_args.cell_type, Some(visi::cli::CellTypeArg::String));

    let cli_alias = try_parse(&[
        "visi",
        "set",
        "data.xlsx",
        "--cell",
        "B2",
        "--cell-type",
        "number",
        "-i",
    ])
    .expect("should parse set with --cell-type alias");

    let Commands::Set(set_args2) = cli_alias.command else {
        panic!("expected Commands::Set");
    };
    assert_eq!(set_args2.cell, vec!["B2"]);
    assert_eq!(set_args2.cell_type, Some(visi::cli::CellTypeArg::Number));
}

#[test]
fn test_workbook_set_cell_type_and_roundtrip() {
    let temp_dir = std::env::temp_dir();
    let file_path = temp_dir.join("test_cell_type_roundtrip.xlsx");
    let file_str = file_path.to_str().unwrap();

    let mut wb = WorkbookManager::new_empty().unwrap();
    wb.set_cell_with_type(
        0,
        0,
        0,
        "12345".to_string(),
        visi_core::core::CellType::String,
    );
    wb.set_cell(0, 0, 1, "12345".to_string());
    wb.set_cell_with_type(
        0,
        0,
        2,
        "TRUE".to_string(),
        visi_core::core::CellType::Boolean,
    );
    wb.set_cell(0, 0, 3, "999".to_string());
    wb.set_cell_type(0, 0, 3, visi_core::core::CellType::String);

    wb.evaluate().unwrap();

    assert_eq!(wb.get_cell_type(0, 0, 0), visi_core::core::CellType::String);
    assert!(matches!(
        wb.sheets[0].get_result_data(&visi_core::core::CellRef::new(0, 0)),
        visi_core::core::ResultData::String(ref s) if s == "12345"
    ));

    assert!(matches!(
        wb.sheets[0].get_result_data(&visi_core::core::CellRef::new(0, 1)),
        visi_core::core::ResultData::Integer(12345)
    ));

    assert_eq!(wb.get_cell_type(0, 0, 3), visi_core::core::CellType::String);
    assert!(matches!(
        wb.sheets[0].get_result_data(&visi_core::core::CellRef::new(0, 3)),
        visi_core::core::ResultData::String(ref s) if s == "999"
    ));

    wb.save_file(file_str).unwrap();

    let loaded = WorkbookManager::load_file(file_str).unwrap();
    assert_eq!(
        loaded.get_cell_type(0, 0, 0),
        visi_core::core::CellType::String
    );
    assert!(matches!(
        loaded.sheets[0].get_result_data(&visi_core::core::CellRef::new(0, 0)),
        visi_core::core::ResultData::String(ref s) if s == "12345"
    ));
    assert!(matches!(
        loaded.sheets[0].get_result_data(&visi_core::core::CellRef::new(0, 1)),
        visi_core::core::ResultData::Integer(12345) | visi_core::core::ResultData::Float(_)
    ));

    let _ = fs::remove_file(&file_path);
}

#[test]
fn test_cli_global_locale_flag_parses() {
    let cli = try_parse(&["visi", "--locale", "de-DE", "eval", "data.xlsx"])
        .expect("should parse global --locale flag");
    assert_eq!(cli.locale.as_deref(), Some("de-DE"));
}

#[test]
fn test_cli_workbook_locale_date_evaluation() {
    let mut wb = WorkbookManager::new_empty().unwrap();
    wb.set_locale(visi_core::core::Locale::de_de());
    wb.set_cell(0, 0, 0, "22.06.2026".to_string());
    wb.set_cell(0, 0, 1, "=A1+1".to_string());
    wb.evaluate().unwrap();

    assert_eq!(
        wb.sheets[0]
            .get_result_data(&visi_core::core::CellRef::new(0, 0))
            .to_string(),
        "46195"
    );
    assert_eq!(
        wb.sheets[0]
            .get_result_data(&visi_core::core::CellRef::new(0, 1))
            .to_string(),
        "46196"
    );
}
