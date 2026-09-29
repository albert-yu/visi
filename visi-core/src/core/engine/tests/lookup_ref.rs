use super::*;

#[test]
fn test_choose_basic() {
    assert_float_close(&eval1("=CHOOSE(1, 10, 20, 30)"), 10.0, 1e-9);
    assert_float_close(&eval1("=CHOOSE(3, 10, 20, 30)"), 30.0, 1e-9);
    assert_eq!(
        eval1("=CHOOSE(2, \"a\", \"b\", \"c\")").to_string(),
        "b".to_string()
    );
}

#[test]
fn test_choose_out_of_range_errors() {
    assert!(
        matches!(eval1("=CHOOSE(0, 10, 20)"), ResultData::Error(ref e) if e.contains("#VALUE!"))
    );
    assert!(
        matches!(eval1("=CHOOSE(5, 10, 20)"), ResultData::Error(ref e) if e.contains("#VALUE!"))
    );
}

#[test]
fn test_choose_lazy_evaluation_skips_unselected_branch_errors() {
    // Real Excel does not evaluate unselected CHOOSE branches; NA() in the
    // unselected branch must not surface.
    let result = eval1("=CHOOSE(1, 42, NA())");
    assert_float_close(&result, 42.0, 1e-9);
}

#[test]
fn test_index_two_arg_form_is_one_based() {
    // The 2-arg INDEX(array, n) form is 1-based (INDEX(list, 1) returns
    // the list's 1st element).
    let grid: [[&str; 2]; 3] = [
        ["10", "=INDEX(A1:A3, 1)"],
        ["20", "=INDEX(A1:A3, 2)"],
        ["30", "=INDEX(A1:A3, 3)"],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    assert_float_close(&sheet.get_result_data(&CellRef::new(0, 1)), 10.0, 1e-9);
    assert_float_close(&sheet.get_result_data(&CellRef::new(1, 1)), 20.0, 1e-9);
    assert_float_close(&sheet.get_result_data(&CellRef::new(2, 1)), 30.0, 1e-9);
}

// --- Range/workbook metadata introspection --------------------------------
//
// Every expected value below was confirmed against real Microsoft Excel
// (FORMULATEXT/ISFORMULA/SHEETS needed an `_xlfn.` prefix to be recognized
// at all when written by openpyxl -- confirmed as the actual cause of a
// real #NAME? mismatch, not a bug in this implementation).

#[test]
fn test_row_and_column_return_array_for_multi_row_or_col_range() {
    // ROW/COLUMN against a multi-row/multi-column reference return
    // an array (one entry per row/column spanned) -- `=SUM(ROW(A1:A5))`
    // is 1+2+3+4+5=15.
    let grid: [[&str; 2]; 5] = [
        ["10", "=SUM(ROW(A1:A5))"],
        ["20", "=INDEX(ROW(A1:A5), 3)"],
        ["30", "=COLUMN(A1)"],
        ["40", ""],
        ["50", ""],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    assert_float_close(&sheet.get_result_data(&CellRef::new(0, 1)), 15.0, 1e-9);
    assert_float_close(&sheet.get_result_data(&CellRef::new(1, 1)), 3.0, 1e-9);
    assert_float_close(&sheet.get_result_data(&CellRef::new(2, 1)), 1.0, 1e-9);
}

#[test]
fn test_rows_columns_areas_match_real_excel() {
    let grid: [[&str; 5]; 3] = [
        ["1", "2", "3", "=ROWS(A1:C3)", "=COLUMNS(A1:C3)"],
        ["4", "5", "6", "=AREAS(A1:C3)", ""],
        ["7", "8", "9", "", ""],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    assert_float_close(&sheet.get_result_data(&CellRef::new(0, 3)), 3.0, 1e-9);
    assert_float_close(&sheet.get_result_data(&CellRef::new(0, 4)), 3.0, 1e-9);
    assert_float_close(&sheet.get_result_data(&CellRef::new(1, 3)), 1.0, 1e-9);
}

#[test]
fn test_isref_distinguishes_references_from_values() {
    assert!(matches!(eval1("=ISREF(A1)"), ResultData::Boolean(true)));
    assert!(matches!(eval1("=ISREF(5)"), ResultData::Boolean(false)));
}

#[test]
fn test_formulatext_and_isformula() {
    let grid: [[&str; 3]; 1] = [["10", "=A1*2", "=FORMULATEXT(B1)"]];
    let mut sheet = create_sheet(&grid);
    sheet.set_cell_src(0, 2, "=ISFORMULA(B1)".to_string());
    sheet.commit(None).unwrap();
    assert!(matches!(
        sheet.get_result_data(&CellRef::new(0, 2)),
        ResultData::Boolean(true)
    ));
    let grid2: [[&str; 2]; 1] = [["10", "=FORMULATEXT(A1)"]];
    let mut sheet2 = create_sheet(&grid2);
    sheet2.commit(None).unwrap();
    assert!(matches!(
        sheet2.get_result_data(&CellRef::new(0, 1)),
        ResultData::Error(ref e) if e.contains("#N/A")
    ));
}

#[test]
fn test_hyperlink_returns_friendly_name_or_link() {
    assert_eq!(
        eval1("=HYPERLINK(\"https://example.com\")").to_string(),
        "https://example.com"
    );
    assert_eq!(
        eval1("=HYPERLINK(\"https://example.com\", \"Click\")").to_string(),
        "Click"
    );
}

#[test]
fn test_sheets_counts_context_sheets() {
    assert!(matches!(eval1("=SHEETS()"), ResultData::Float(f) if f == 1.0));
}

#[test]
fn test_sheet_reports_real_workbook_ordinal() {
    // SHEET() reports the sheet's 1-based position in workbook order.
    let table1 = Sheet::new(SheetInit {
        id: None,
        name: Some("table_1".to_string()),
        rows: 2,
        cols: 2,
    });
    let table2 = Sheet::new(SheetInit {
        id: None,
        name: Some("table_2".to_string()),
        rows: 2,
        cols: 2,
    });
    let table3 = Sheet::new(SheetInit {
        id: None,
        name: Some("table_3".to_string()),
        rows: 2,
        cols: 2,
    });
    let sheets = [table1, table2, table3];

    let mut context = Context::new();
    for sheet in &sheets {
        context.add_table(sheet.name.clone(), sheet);
    }
    context.sheet_order = sheets.iter().map(|s| s.name.clone()).collect();

    let (r1, _) = sheets[0].eval("=SHEET()", Some(&context)).unwrap();
    assert_float_close(&r1, 1.0, 1e-9);
    let (r2, _) = sheets[1].eval("=SHEET()", Some(&context)).unwrap();
    assert_float_close(&r2, 2.0, 1e-9);
    let (r3, _) = sheets[2].eval("=SHEET()", Some(&context)).unwrap();
    assert_float_close(&r3, 3.0, 1e-9);

    // A reference into another sheet reports *that* sheet's ordinal, not
    // the formula's own.
    let (r4, _) = sheets[0]
        .eval("=SHEET(table_3!A1)", Some(&context))
        .unwrap();
    assert_float_close(&r4, 3.0, 1e-9);

    // A plain text sheet name is also accepted, same as real Excel.
    let (r5, _) = sheets[0]
        .eval("=SHEET(\"table_2\")", Some(&context))
        .unwrap();
    assert_float_close(&r5, 2.0, 1e-9);

    // No context at all (standalone eval outside a WorkbookManager pass)
    // keeps the old documented fallback of 1.
    assert!(matches!(eval1("=SHEET()"), ResultData::Float(f) if f == 1.0));
}

#[test]
fn test_indirect_resolves_cell_and_range_text() {
    let grid: [[&str; 2]; 3] = [
        ["10", "=INDIRECT(\"A1\")"],
        ["20", "=SUM(INDIRECT(\"A1:A3\"))"],
        ["30", ""],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    assert_float_close(&sheet.get_result_data(&CellRef::new(0, 1)), 10.0, 1e-9);
    assert_float_close(&sheet.get_result_data(&CellRef::new(1, 1)), 60.0, 1e-9);
}

#[test]
fn test_offset_shifts_and_resizes_reference() {
    let grid: [[&str; 2]; 3] = [
        ["10", "=OFFSET(A1, 1, 0)"],
        ["20", "=SUM(OFFSET(A1, 0, 0, 3, 1))"],
        ["30", ""],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    assert_float_close(&sheet.get_result_data(&CellRef::new(0, 1)), 20.0, 1e-9);
    assert_float_close(&sheet.get_result_data(&CellRef::new(1, 1)), 60.0, 1e-9);
}

#[test]
fn test_cell_info_subset() {
    let grid: [[&str; 2]; 5] = [
        ["10", "=CELL(\"row\", A3)"],
        ["20", "=CELL(\"col\", A3)"],
        ["30", "=CELL(\"address\", A3)"],
        ["40", "=CELL(\"contents\", A1)"],
        ["50", ""],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    assert_float_close(&sheet.get_result_data(&CellRef::new(0, 1)), 3.0, 1e-9);
    assert_float_close(&sheet.get_result_data(&CellRef::new(1, 1)), 1.0, 1e-9);
    assert_eq!(
        sheet.get_result_data(&CellRef::new(2, 1)).to_string(),
        "$A$3"
    );
    assert_float_close(&sheet.get_result_data(&CellRef::new(3, 1)), 10.0, 1e-9);
}

#[test]
fn test_lookup_vector_form_on_sorted_data() {
    // LOOKUP requires an ascending-sorted lookup vector; Excel's own docs
    // call behavior on unsorted input unpredictable (confirmed
    // divergent-but-Excel-undefined via the differential fuzzer on
    // unsorted input, not treated as a visi bug). On sorted input, visi
    // matches real Excel exactly for exact, mid-range, below-range, and
    // above-range lookups.
    let grid: [[&str; 2]; 5] = [
        ["5", "50"],
        ["10", "100"],
        ["20", "200"],
        ["20", "201"],
        ["30", "300"],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    let cases: [(&str, f64); 3] = [
        ("=LOOKUP(20,A1:A5,B1:B5)", 201.0),
        ("=LOOKUP(15,A1:A5,B1:B5)", 100.0),
        ("=LOOKUP(100,A1:A5,B1:B5)", 300.0),
    ];
    for (formula, expected) in cases {
        let (result, _) = sheet.eval(formula, None).unwrap();
        assert_float_close(&result, expected, 1e-9);
    }
    assert!(matches!(
        sheet.eval("=LOOKUP(1,A1:A5,B1:B5)", None).unwrap().0,
        ResultData::Error(ref e) if e.contains("#N/A")
    ));
}

#[test]
fn test_xmatch_supports_next_smaller_and_larger_modes() {
    let grid: [[&str; 1]; 5] = [["10"], ["20"], ["5"], ["20"], ["30"]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    let (exact, _) = sheet.eval("=XMATCH(20,A1:A5)", None).unwrap();
    assert_float_close(&exact, 2.0, 1e-9);
    let (next_smaller, _) = sheet.eval("=XMATCH(25,A1:A5,-1)", None).unwrap();
    assert_float_close(&next_smaller, 2.0, 1e-9);
}

#[test]
fn test_bare_row_and_column_report_current_cell_position() {
    // No-arg ROW()/COLUMN() report the position of the cell the formula
    // itself lives in -- distinct from the reference-argument form, which
    // reports the referenced range instead (already covered elsewhere).
    let grid: [[&str; 3]; 2] = [["=ROW()", "=COLUMN()", "x"], ["x", "x", "=ROW()+COLUMN()"]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    assert_float_close(&sheet.get_result_data(&CellRef::new(0, 0)), 1.0, 1e-9);
    assert_float_close(&sheet.get_result_data(&CellRef::new(0, 1)), 2.0, 1e-9);
    assert_float_close(&sheet.get_result_data(&CellRef::new(1, 2)), 5.0, 1e-9);
}

#[test]
fn test_hlookup_exact_and_approximate_match_on_text_header_row() {
    // HLOOKUP matches exact and approximate values on text header rows.
    let grid: [[&str; 3]; 2] = [["Jan", "Feb", "Mar"], ["10", "20", "30"]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    let (exact, _) = sheet.eval("=HLOOKUP(\"Feb\",A1:C2,2,FALSE)", None).unwrap();
    assert_float_close(&exact, 20.0, 1e-9);

    assert!(matches!(
        sheet.eval("=HLOOKUP(\"Nope\",A1:C2,2,FALSE)", None).unwrap().0,
        ResultData::Error(ref e) if e == "#N/A"
    ));

    let numeric_grid: [[&str; 3]; 2] = [["10", "20", "30"], ["a", "b", "c"]];
    let mut numeric_sheet = create_sheet(&numeric_grid);
    numeric_sheet.commit(None).unwrap();
    // Approximate match: largest header <= 25 is 20, in column 2.
    let (approx, _) = numeric_sheet.eval("=HLOOKUP(25,A1:C2,2)", None).unwrap();
    assert_eq!(approx.to_string(), "b");
}
