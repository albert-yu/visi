use std::env;
use std::fs;
use std::path::{Path, PathBuf};

use visi_core::core::chart::ChartType;
use visi_core::core::{
    ArrayFormulaEntry, ArrayFormulaKind, BorderSide, BorderStyle, CellAlignment, CellBorders,
    CellComment, CellHyperlink, CellStyle, ConditionalFormatRule, DataValidationKind,
    DataValidationOperator, DataValidationRule, DefinedName, FillPattern, FreezePane,
    HorizontalAlignment, MergedCellRange, PivotAggregation, PivotArea, SheetVisibility,
    VbaModuleKind, VerticalAlignment, WorkbookManager,
};

fn write_fixture(dir: &Path, name: &str, bytes: &[u8]) {
    let path = dir.join(name);
    fs::write(&path, bytes).unwrap_or_else(|e| panic!("failed writing {}: {e}", path.display()));
}

fn emit_values_and_types(dir: &Path) {
    let mut mgr = WorkbookManager::new_empty().unwrap();
    mgr.set_cell(0, 0, 0, "42".to_string());
    mgr.set_cell(0, 0, 1, "-3.14159265".to_string());
    mgr.set_cell(0, 0, 2, "1.23e12".to_string());
    mgr.set_cell(0, 1, 0, "plain text".to_string());
    mgr.set_cell(0, 1, 1, "  leading and trailing  ".to_string());
    mgr.set_cell(0, 1, 2, "xml & <tags> \"quotes\" 'apos'".to_string());
    mgr.set_cell(0, 2, 0, "=TRUE()".to_string());
    mgr.set_cell(0, 2, 1, "=FALSE()".to_string());
    mgr.set_cell(0, 3, 0, "=1/0".to_string());
    mgr.set_cell(0, 3, 1, "=NA()".to_string());
    mgr.set_cell(0, 3, 2, "=SQRT(-1)".to_string());
    mgr.set_cell(0, 3, 3, "=\"a\"+1".to_string());
    mgr.set_cell(0, 3, 4, "=NONEXISTENT_FN_XYZ()".to_string());
    mgr.set_cell(0, 4, 0, "2026-11-20".to_string());
    mgr.evaluate().unwrap();

    let fresh = mgr.save_bytes().unwrap();
    write_fixture(dir, "01_values_and_types.xlsx", &fresh);

    let mut loaded = WorkbookManager::load_bytes(&fresh).unwrap();
    loaded.set_cell(0, 0, 0, "100".to_string());
    loaded.set_cell(0, 1, 3, "\tindented\nline2".to_string());
    loaded.evaluate().unwrap();
    write_fixture(
        dir,
        "01_values_and_types_patched.xlsx",
        &loaded.save_bytes().unwrap(),
    );
}

fn emit_styles_and_formatting(dir: &Path) {
    let mut mgr = WorkbookManager::new_empty().unwrap();
    mgr.set_cell(0, 0, 0, "Header".to_string());
    mgr.set_cell(0, 1, 0, "12345.678".to_string());
    mgr.set_cell(0, 2, 0, "0.456".to_string());
    mgr.set_cell(0, 3, 0, "Wrapped & Rotated".to_string());
    mgr.set_cell(0, 4, 0, "Pattern Fill".to_string());

    mgr.sheets[0].set_cell_style(
        0,
        0,
        CellStyle {
            bold: Some(true),
            italic: Some(true),
            underline: Some(true),
            strikethrough: Some(true),
            font_family: Some("Arial".to_string()),
            font_size: Some(14.0),
            font_color: Some("#FFFFFF".to_string()),
            bg_color: Some("#1F4E78".to_string()),
            fill_pattern: Some(FillPattern::Solid),
            borders: Some(CellBorders {
                left: Some(BorderSide {
                    style: BorderStyle::Thin,
                    color: Some("#000000".to_string()),
                }),
                right: Some(BorderSide {
                    style: BorderStyle::Medium,
                    color: Some("#FF0000".to_string()),
                }),
                top: Some(BorderSide {
                    style: BorderStyle::Double,
                    color: Some("#00FF00".to_string()),
                }),
                bottom: Some(BorderSide {
                    style: BorderStyle::Thick,
                    color: Some("#0000FF".to_string()),
                }),
                diagonal: Some(BorderSide {
                    style: BorderStyle::Dashed,
                    color: Some("#808080".to_string()),
                }),
                diagonal_up: true,
                diagonal_down: false,
            }),
            alignment: Some(CellAlignment {
                horizontal: Some(HorizontalAlignment::Center),
                vertical: Some(VerticalAlignment::Center),
                wrap_text: Some(true),
                text_rotation: Some(0),
                indent: Some(1),
                shrink_to_fit: Some(false),
            }),
            ..Default::default()
        },
    );

    mgr.sheets[0].set_cell_style(
        1,
        0,
        CellStyle {
            num_format: Some("#,##0.00".to_string()),
            alignment: Some(CellAlignment {
                horizontal: Some(HorizontalAlignment::Right),
                ..Default::default()
            }),
            ..Default::default()
        },
    );

    mgr.sheets[0].set_cell_style(
        2,
        0,
        CellStyle {
            num_format: Some("0.0%".to_string()),
            ..Default::default()
        },
    );

    mgr.sheets[0].set_cell_style(
        3,
        0,
        CellStyle {
            alignment: Some(CellAlignment {
                horizontal: Some(HorizontalAlignment::Left),
                vertical: Some(VerticalAlignment::Top),
                wrap_text: Some(true),
                text_rotation: Some(45),
                shrink_to_fit: Some(true),
                ..Default::default()
            }),
            ..Default::default()
        },
    );

    mgr.sheets[0].set_cell_style(
        4,
        0,
        CellStyle {
            bg_color: Some("#D9E1F2".to_string()),
            fill_pattern: Some(FillPattern::DarkHorizontal),
            ..Default::default()
        },
    );

    mgr.evaluate().unwrap();
    let fresh = mgr.save_bytes().unwrap();
    write_fixture(dir, "02_styles_and_formatting.xlsx", &fresh);

    let mut loaded = WorkbookManager::load_bytes(&fresh).unwrap();
    loaded.set_cell(0, 5, 0, "Patched Style".to_string());
    loaded.sheets[0].set_cell_style(
        5,
        0,
        CellStyle {
            bold: Some(true),
            bg_color: Some("#FFE699".to_string()),
            borders: Some(CellBorders {
                bottom: Some(BorderSide {
                    style: BorderStyle::Dotted,
                    color: Some("#333333".to_string()),
                }),
                ..Default::default()
            }),
            ..Default::default()
        },
    );
    loaded.evaluate().unwrap();
    write_fixture(
        dir,
        "02_styles_and_formatting_patched.xlsx",
        &loaded.save_bytes().unwrap(),
    );
}

fn emit_sheet_structure(dir: &Path) {
    let mut mgr = WorkbookManager::new_empty().unwrap();
    mgr.add_sheet("HiddenSheet").unwrap();
    mgr.add_sheet("VeryHiddenSheet").unwrap();
    mgr.add_sheet("ColFreeze").unwrap();

    mgr.set_cell(0, 0, 0, "Title Banner".to_string());
    mgr.set_cell(0, 2, 0, "Row 3".to_string());
    mgr.set_cell(0, 3, 2, "Merged Block".to_string());
    mgr.sheets[0].merged_cells.push(MergedCellRange {
        start_row: 0,
        start_col: 0,
        end_row: 0,
        end_col: 3,
    });
    mgr.sheets[0].merged_cells.push(MergedCellRange {
        start_row: 3,
        start_col: 2,
        end_row: 4,
        end_col: 3,
    });
    mgr.sheets[0].set_column_width(0, Some(24.0));
    mgr.sheets[0].set_column_hidden(1, true);
    mgr.sheets[0].set_row_height(0, Some(28.5));
    mgr.sheets[0].set_row_hidden(2, true);
    mgr.sheets[0].freeze_pane = Some(FreezePane { rows: 1, cols: 1 });
    mgr.sheets[0].date1904 = true;

    mgr.sheets[1].visibility = SheetVisibility::Hidden;
    mgr.sheets[1].date1904 = true;
    mgr.sheets[1].freeze_pane = Some(FreezePane { rows: 2, cols: 0 });
    mgr.set_cell(1, 0, 0, "secret".to_string());

    mgr.sheets[2].visibility = SheetVisibility::VeryHidden;
    mgr.sheets[2].date1904 = true;
    mgr.set_cell(2, 0, 0, "very secret".to_string());

    mgr.sheets[3].date1904 = true;
    mgr.sheets[3].freeze_pane = Some(FreezePane { rows: 0, cols: 2 });
    mgr.set_cell(3, 0, 0, "col freeze".to_string());

    mgr.evaluate().unwrap();
    let fresh = mgr.save_bytes().unwrap();
    write_fixture(dir, "03_sheet_structure.xlsx", &fresh);

    let mut loaded = WorkbookManager::load_bytes(&fresh).unwrap();
    loaded.set_cell(0, 5, 0, "added row".to_string());
    loaded.sheets[0].set_row_hidden(5, true);
    loaded.evaluate().unwrap();
    write_fixture(
        dir,
        "03_sheet_structure_patched.xlsx",
        &loaded.save_bytes().unwrap(),
    );
}

fn emit_formulas_and_arrays(dir: &Path) {
    let mut mgr = WorkbookManager::new_empty().unwrap();
    mgr.add_sheet("Lookup").unwrap();

    mgr.set_cell(0, 0, 0, "10".to_string());
    mgr.set_cell(0, 1, 0, "20".to_string());
    mgr.set_cell(0, 2, 0, "30".to_string());
    mgr.set_cell(0, 0, 1, "=SUM(A1:A3)".to_string());
    mgr.set_cell(0, 1, 1, "=XLOOKUP(20, A1:A3, Lookup!A1:A3)".to_string());
    mgr.set_cell(0, 2, 1, "=LET(x, A1*2, x + TaxRate)".to_string());
    mgr.set_cell(0, 3, 1, "=LocalBonus + A1".to_string());
    mgr.set_cell(0, 0, 3, "=SEQUENCE(3, 1, 1, 1)".to_string());
    mgr.sheets[0].array_formulas.push(ArrayFormulaEntry {
        start_row: 0,
        start_col: 3,
        end_row: 2,
        end_col: 3,
        kind: ArrayFormulaKind::DynamicArray,
    });
    mgr.set_cell(0, 0, 5, "=A1:A3*2".to_string());
    mgr.sheets[0].array_formulas.push(ArrayFormulaEntry {
        start_row: 0,
        start_col: 5,
        end_row: 2,
        end_col: 5,
        kind: ArrayFormulaKind::LegacyCse,
    });

    mgr.set_cell(1, 0, 0, "Alpha".to_string());
    mgr.set_cell(1, 1, 0, "Beta".to_string());
    mgr.set_cell(1, 2, 0, "Gamma".to_string());

    mgr.add_defined_name(DefinedName {
        name: "TaxRate".to_string(),
        formula: "0.08".to_string(),
        local_sheet_id: None,
        hidden: false,
        comment: Some("Global tax rate".to_string()),
    });
    mgr.add_defined_name(DefinedName {
        name: "LocalBonus".to_string(),
        formula: "5".to_string(),
        local_sheet_id: Some(0),
        hidden: true,
        comment: None,
    });

    mgr.evaluate().unwrap();
    let fresh = mgr.save_bytes().unwrap();
    write_fixture(dir, "04_formulas_and_arrays.xlsx", &fresh);

    let mut loaded = WorkbookManager::load_bytes(&fresh).unwrap();
    loaded.set_cell(0, 0, 0, "15".to_string());
    loaded.evaluate().unwrap();
    write_fixture(
        dir,
        "04_formulas_and_arrays_patched.xlsx",
        &loaded.save_bytes().unwrap(),
    );
}

fn emit_comments_hyperlinks_validation_cf(dir: &Path) {
    let mut mgr = WorkbookManager::new_empty().unwrap();
    mgr.add_sheet("Target").unwrap();

    mgr.set_cell(0, 0, 0, "Visit Docs".to_string());
    mgr.set_cell(0, 1, 0, "Go to Target".to_string());
    mgr.set_cell(0, 0, 1, "50".to_string());
    mgr.set_cell(0, 1, 1, "150".to_string());
    mgr.set_cell(0, 2, 1, "High".to_string());
    mgr.set_cell(1, 0, 0, "Arrived".to_string());

    mgr.sheets[0].comments.push(CellComment {
        row: 0,
        col: 1,
        author: "Alice".to_string(),
        text: "First note".to_string(),
    });
    mgr.sheets[0].comments.push(CellComment {
        row: 1,
        col: 1,
        author: "Bob".to_string(),
        text: "Line 1\nLine 2".to_string(),
    });

    mgr.sheets[0].hyperlinks.push(CellHyperlink {
        start_row: 0,
        start_col: 0,
        end_row: 0,
        end_col: 0,
        target: Some("https://example.com/docs".to_string()),
        location: None,
        display: Some("Visit Docs".to_string()),
        tooltip: Some("Open external docs".to_string()),
    });
    mgr.sheets[0].hyperlinks.push(CellHyperlink {
        start_row: 1,
        start_col: 0,
        end_row: 1,
        end_col: 0,
        target: None,
        location: Some("Target!A1".to_string()),
        display: Some("Go to Target".to_string()),
        tooltip: None,
    });

    mgr.sheets[0].data_validations.push(DataValidationRule {
        sqref: "B1:B2".to_string(),
        kind: DataValidationKind::Whole,
        operator: Some(DataValidationOperator::Between),
        allow_blank: true,
        formula1: Some("0".to_string()),
        formula2: Some("500".to_string()),
        prompt_title: Some("Score".to_string()),
        prompt: Some("Enter 0..500".to_string()),
        error_title: Some("Invalid".to_string()),
        error: Some("Out of range".to_string()),
        show_input_message: true,
        show_error_message: true,
    });
    mgr.sheets[0].data_validations.push(DataValidationRule {
        sqref: "B3".to_string(),
        kind: DataValidationKind::List,
        operator: None,
        allow_blank: false,
        formula1: Some("\"Low,Medium,High\"".to_string()),
        formula2: None,
        ..Default::default()
    });

    mgr.sheets[0]
        .conditional_formats
        .push(ConditionalFormatRule {
            sqref: "B1:B2".to_string(),
            rule_type: "cellIs".to_string(),
            operator: Some("greaterThan".to_string()),
            priority: 1,
            stop_if_true: true,
            formulas: vec!["100".to_string()],
            style: Some(CellStyle {
                bold: Some(true),
                font_color: Some("#9C0006".to_string()),
                bg_color: Some("#FFC7CE".to_string()),
                ..Default::default()
            }),
        });
    mgr.sheets[0]
        .conditional_formats
        .push(ConditionalFormatRule {
            sqref: "B1:B2".to_string(),
            rule_type: "expression".to_string(),
            operator: None,
            priority: 2,
            stop_if_true: false,
            formulas: vec!["B1>0".to_string()],
            style: Some(CellStyle {
                italic: Some(true),
                ..Default::default()
            }),
        });

    mgr.evaluate().unwrap();
    let fresh = mgr.save_bytes().unwrap();
    write_fixture(dir, "05_comments_hyperlinks_validation_cf.xlsx", &fresh);

    let mut loaded = WorkbookManager::load_bytes(&fresh).unwrap();
    loaded.set_cell(0, 0, 1, "200".to_string());
    loaded.sheets[0].comments.push(CellComment {
        row: 2,
        col: 1,
        author: "Alice".to_string(),
        text: "Added on patch".to_string(),
    });
    loaded.evaluate().unwrap();
    write_fixture(
        dir,
        "05_comments_hyperlinks_validation_cf_patched.xlsx",
        &loaded.save_bytes().unwrap(),
    );
}

fn emit_tables(dir: &Path) {
    let mut mgr = WorkbookManager::new_empty().unwrap();
    let rows = [
        ("Item", "Qty", "Price"),
        ("Widget", "3", "10"),
        ("Gadget", "5", "20"),
    ];
    for (r, (a, b, c)) in rows.iter().enumerate() {
        mgr.set_cell(0, r, 0, (*a).to_string());
        mgr.set_cell(0, r, 1, (*b).to_string());
        mgr.set_cell(0, r, 2, (*c).to_string());
    }
    mgr.add_table(None, "Sales", 0, 0, 2, 2, true, false)
        .unwrap();
    mgr.set_table_style("Sales", "TableStyleMedium9").unwrap();
    mgr.set_cell(0, 0, 4, "=SUM(Sales[Price])".to_string());
    mgr.evaluate().unwrap();

    let fresh = mgr.save_bytes().unwrap();
    write_fixture(dir, "06_tables.xlsx", &fresh);

    let mut loaded = WorkbookManager::load_bytes(&fresh).unwrap();
    loaded.insert_row(0, 2).unwrap();
    loaded.set_cell(0, 2, 0, "Bolt".to_string());
    loaded.set_cell(0, 2, 1, "7".to_string());
    loaded.set_cell(0, 2, 2, "5".to_string());
    loaded.evaluate().unwrap();
    write_fixture(dir, "06_tables_patched.xlsx", &loaded.save_bytes().unwrap());
}

fn emit_pivots(dir: &Path) {
    let mut mgr = WorkbookManager::new_empty().unwrap();
    let rows = [
        ("Region", "Quarter", "Category", "Sales", "Units"),
        ("East", "Q1", "A", "120", "4"),
        ("West", "Q1", "B", "200", "6"),
        ("East", "Q2", "A", "80", "2"),
        ("West", "Q2", "B", "150", "5"),
    ];
    for (r, row) in rows.iter().enumerate() {
        mgr.set_cell(0, r, 0, row.0.to_string());
        mgr.set_cell(0, r, 1, row.1.to_string());
        mgr.set_cell(0, r, 2, row.2.to_string());
        mgr.set_cell(0, r, 3, row.3.to_string());
        mgr.set_cell(0, r, 4, row.4.to_string());
    }
    mgr.evaluate().unwrap();

    mgr.add_pivot_table_from_range("RangePivot", None, 0, 0, 4, 4, None, 0, 7, true, true)
        .unwrap();
    mgr.add_pivot_field("RangePivot", PivotArea::Row, "Region", None)
        .unwrap();
    mgr.add_pivot_field("RangePivot", PivotArea::Column, "Quarter", None)
        .unwrap();
    mgr.add_pivot_field("RangePivot", PivotArea::Filter, "Category", None)
        .unwrap();
    mgr.add_pivot_field(
        "RangePivot",
        PivotArea::Value,
        "Sales",
        Some(PivotAggregation::Sum),
    )
    .unwrap();
    mgr.add_pivot_field(
        "RangePivot",
        PivotArea::Value,
        "Units",
        Some(PivotAggregation::Average),
    )
    .unwrap();
    mgr.refresh_pivot_table("RangePivot").unwrap();

    let fresh = mgr.save_bytes().unwrap();
    write_fixture(dir, "07_pivots.xlsx", &fresh);

    let mut loaded = WorkbookManager::load_bytes(&fresh).unwrap();
    loaded.set_cell(0, 1, 3, "300".to_string());
    loaded.evaluate().unwrap();
    loaded.refresh_pivot_table("RangePivot").unwrap();
    write_fixture(dir, "07_pivots_patched.xlsx", &loaded.save_bytes().unwrap());
}

fn emit_charts(dir: &Path) {
    let mut mgr = WorkbookManager::new_empty().unwrap();
    let rows = [
        ("Month", "Rev", "Cost"),
        ("Jan", "100", "60"),
        ("Feb", "140", "70"),
        ("Mar", "180", "90"),
    ];
    for (r, (m, rev, cost)) in rows.iter().enumerate() {
        mgr.set_cell(0, r, 0, (*m).to_string());
        mgr.set_cell(0, r, 1, (*rev).to_string());
        mgr.set_cell(0, r, 2, (*cost).to_string());
    }
    mgr.evaluate().unwrap();

    let kinds = [
        (ChartType::Column, "Column Chart", (0, 5)),
        (ChartType::Bar, "Bar Chart", (16, 5)),
        (ChartType::Line, "Line Chart", (32, 5)),
        (ChartType::Pie, "Pie Chart", (0, 14)),
        (ChartType::Scatter, "Scatter Chart", (16, 14)),
        (ChartType::Area, "Area Chart", (32, 14)),
    ];
    for (ct, title, anchor) in kinds {
        mgr.add_chart(
            "Sheet1",
            ct,
            "Sheet1!A1:C4".to_string(),
            Some(title.to_string()),
            Some(anchor),
        )
        .unwrap();
    }

    let fresh = mgr.save_bytes().unwrap();
    write_fixture(dir, "08_charts.xlsx", &fresh);

    let mut loaded = WorkbookManager::load_bytes(&fresh).unwrap();
    let first_id = loaded.charts[0].id;
    loaded
        .edit_chart(
            first_id,
            None,
            None,
            None,
            Some(Some("Updated Column".to_string())),
            Some(Some("Months".to_string())),
            Some(Some("USD".to_string())),
            Some(true),
            None,
        )
        .unwrap();
    write_fixture(dir, "08_charts_patched.xlsx", &loaded.save_bytes().unwrap());
}

fn emit_structural_edits(dir: &Path) {
    let mut mgr = WorkbookManager::new_empty().unwrap();
    mgr.add_sheet("Data").unwrap();
    mgr.add_sheet("Scratch").unwrap();
    mgr.set_cell(0, 0, 0, "=SUM(Data!A1:A3)".to_string());
    mgr.set_cell(1, 0, 0, "10".to_string());
    mgr.set_cell(1, 1, 0, "20".to_string());
    mgr.set_cell(1, 2, 0, "30".to_string());
    mgr.evaluate().unwrap();
    let base = mgr.save_bytes().unwrap();

    let mut loaded = WorkbookManager::load_bytes(&base).unwrap();
    loaded.insert_row(1, 1).unwrap();
    loaded.set_cell(1, 1, 0, "15".to_string());
    loaded.insert_col(1, 0).unwrap();
    loaded.delete_col(1, 0).unwrap();
    loaded.rename_sheet("Data", "Metrics").unwrap();
    loaded.delete_sheet("Scratch").unwrap();
    loaded.evaluate().unwrap();
    write_fixture(
        dir,
        "09_structural_edits_patched.xlsx",
        &loaded.save_bytes().unwrap(),
    );
}

fn emit_vba_macros(dir: &Path) {
    let mut mgr = WorkbookManager::new_empty().unwrap();
    mgr.set_cell(0, 0, 0, "10".to_string());
    let src = "Attribute VB_Name = \"Mod1\"\nPublic Sub DoubleA1()\n    Range(\"B1\").Value = Range(\"A1\").Value * 2\nEnd Sub\n";
    mgr.add_vba_module(
        "Mod1".to_string(),
        VbaModuleKind::Standard,
        src.to_string(),
        None,
    )
    .unwrap();
    mgr.add_vba_module(
        "ThisWorkbook".to_string(),
        VbaModuleKind::Document,
        "Attribute VB_Name = \"ThisWorkbook\"\n".to_string(),
        None,
    )
    .unwrap();
    let fresh = mgr.save_bytes().unwrap();
    write_fixture(dir, "10_vba_macros.xlsm", &fresh);

    let mut loaded = WorkbookManager::load_bytes(&fresh).unwrap();
    loaded.run_macro(None, "DoubleA1", &[]).unwrap();
    loaded.evaluate().unwrap();
    write_fixture(
        dir,
        "10_vba_macros_patched.xlsm",
        &loaded.save_bytes().unwrap(),
    );
}

fn main() {
    let out_dir = env::args()
        .nth(1)
        .map(PathBuf::from)
        .expect("usage: emit_ooxml_fixtures <output-dir>");
    fs::create_dir_all(&out_dir).unwrap();

    emit_values_and_types(&out_dir);
    emit_styles_and_formatting(&out_dir);
    emit_sheet_structure(&out_dir);
    emit_formulas_and_arrays(&out_dir);
    emit_comments_hyperlinks_validation_cf(&out_dir);
    emit_tables(&out_dir);
    emit_pivots(&out_dir);
    emit_charts(&out_dir);
    emit_structural_edits(&out_dir);
    emit_vba_macros(&out_dir);
}
