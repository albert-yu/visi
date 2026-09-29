use super::*;

// --- Dynamic array reshaping/lookup batch ---
//
// Each test below builds its 2D input with `SEQUENCE(rows, cols)` so it
// doesn't need a cell grid, then pulls values back out with `INDEX`/`SUM`.
// HSTACK, VSTACK, UNIQUE, SORT, XMATCH, and FILTER (with a genuine boolean
// helper range rather than a broadcast comparison -- this engine's
// comparison operators don't broadcast across ranges, a separate,
// pre-existing, out-of-scope limitation) were confirmed byte-for-byte
// against real Microsoft Excel via the differential fuzzer. WRAPROWS,
// WRAPCOLS, CHOOSEROWS, CHOOSECOLS, DROP, TAKE, EXPAND, TOCOL, TOROW,
// SORTBY, TRIMRANGE, and LOOKUP (on genuinely sorted input) were likewise
// confirmed against real Excel. TRANSPOSE could not be verified against
// real Excel: every authoring variant tried (bare, `_xlfn.`,
// `_xlfn._xlws.`, standalone, and nested inside SUM/INDEX) gave `#VALUE!`
// in real Excel when the formula was written by openpyxl rather than
// Excel itself, even though TRANSPOSE predates dynamic arrays entirely --
// this points at an openpyxl authoring limitation (TRANSPOSE has always
// required the legacy CSE `t="array"` formula flag, which openpyxl's plain
// string assignment doesn't produce) rather than a bug in visi's TRANSPOSE
// logic, which is hand-verified correct below.

#[test]
fn test_transpose_swaps_rows_and_cols() {
    // SEQUENCE(2,3) = [[1,2,3],[4,5,6]]; transposed = [[1,4],[2,5],[3,6]].
    assert_float_close(
        &eval_formula("=INDEX(TRANSPOSE(SEQUENCE(2,3)),3,1)"),
        3.0,
        1e-9,
    );
    assert_float_close(
        &eval_formula("=INDEX(TRANSPOSE(SEQUENCE(2,3)),1,2)"),
        4.0,
        1e-9,
    );
    assert_float_close(&eval_formula("=SUM(TRANSPOSE(SEQUENCE(2,3)))"), 21.0, 1e-9);
}

#[test]
fn test_index_recovers_shape_of_nested_reshape_function() {
    // INDEX's 3-arg row/col form recovers the 2D shape of array-producing
    // function calls (such as EXPAND).
    let grid: [[&str; 3]; 2] = [
        ["1", "2", "=INDEX(EXPAND(A1:B2,3,3,0),3,3)"],
        ["4", "5", ""],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    assert_float_close(&sheet.get_result_data(&CellRef::new(0, 2)), 0.0, 1e-9);
}

#[test]
fn test_hstack_vstack_combine_arrays() {
    // HSTACK(SEQUENCE(2,1), SEQUENCE(2,1)) side-by-side: [[1,1],[2,2]].
    assert_float_close(
        &eval_formula("=INDEX(HSTACK(SEQUENCE(2,1),SEQUENCE(2,1)),2,1)"),
        2.0,
        1e-9,
    );
    assert_float_close(
        &eval_formula("=INDEX(HSTACK(SEQUENCE(2,1),SEQUENCE(2,1)),2,2)"),
        2.0,
        1e-9,
    );
    // VSTACK(SEQUENCE(1,2), SEQUENCE(1,2)) stacked: [[1,2],[1,2]].
    assert_float_close(
        &eval_formula("=INDEX(VSTACK(SEQUENCE(1,2),SEQUENCE(1,2)),2,2)"),
        2.0,
        1e-9,
    );
    assert_float_close(
        &eval_formula("=SUM(HSTACK(SEQUENCE(2,1),SEQUENCE(2,1)))"),
        6.0,
        1e-9,
    );
}

#[test]
fn test_chooserows_chosecols_select_by_index() {
    // SEQUENCE(3,3) = [[1,2,3],[4,5,6],[7,8,9]].
    assert_float_close(
        &eval_formula("=INDEX(CHOOSEROWS(SEQUENCE(3,3),2),1,1)"),
        4.0,
        1e-9,
    );
    assert_float_close(
        &eval_formula("=INDEX(CHOOSEROWS(SEQUENCE(3,3),-1),1,1)"),
        7.0,
        1e-9,
    );
    assert_float_close(
        &eval_formula("=INDEX(CHOOSECOLS(SEQUENCE(3,3),2),1,1)"),
        2.0,
        1e-9,
    );
}

#[test]
fn test_drop_take_slice_from_either_end() {
    assert_float_close(&eval_formula("=SUM(DROP(SEQUENCE(3,3),1))"), 39.0, 1e-9);
    assert_float_close(&eval_formula("=SUM(DROP(SEQUENCE(3,3),-1))"), 21.0, 1e-9);
    assert_float_close(&eval_formula("=SUM(TAKE(SEQUENCE(3,3),2))"), 21.0, 1e-9);
    assert_float_close(&eval_formula("=SUM(TAKE(SEQUENCE(3,3),-1))"), 24.0, 1e-9);
}

#[test]
fn test_expand_pads_with_given_value() {
    assert_float_close(
        &eval_formula("=INDEX(EXPAND(SEQUENCE(2,2),3,3,0),3,3)"),
        0.0,
        1e-9,
    );
    assert_float_close(
        &eval_formula("=INDEX(EXPAND(SEQUENCE(2,2),3,3,0),1,1)"),
        1.0,
        1e-9,
    );
}

#[test]
fn test_tocol_torow_flatten() {
    assert_float_close(&eval_formula("=SUM(TOCOL(SEQUENCE(3,3)))"), 45.0, 1e-9);
    assert_float_close(&eval_formula("=SUM(TOROW(SEQUENCE(3,3)))"), 45.0, 1e-9);
    assert_float_close(&eval_formula("=INDEX(TOCOL(SEQUENCE(2,2)),3)"), 3.0, 1e-9);
}

#[test]
fn test_wraprows_wrapcols_reshape_flat_sequence() {
    // Confirmed against real Excel: WRAPROWS(SEQUENCE(7),3,0) wraps
    // [1..7] into rows of 3, padding the last row with 0; WRAPCOLS does
    // the same column-major.
    assert_float_close(
        &eval_formula("=INDEX(WRAPROWS(SEQUENCE(7),3,0),3,1)"),
        7.0,
        1e-9,
    );
    assert_float_close(&eval_formula("=SUM(WRAPROWS(SEQUENCE(7),3,0))"), 28.0, 1e-9);
    assert_float_close(
        &eval_formula("=INDEX(WRAPCOLS(SEQUENCE(7),3,0),1,3)"),
        7.0,
        1e-9,
    );
    assert_float_close(&eval_formula("=SUM(WRAPCOLS(SEQUENCE(7),3,0))"), 28.0, 1e-9);
}

#[test]
fn test_unique_sort_sortby_filter_trimrange() {
    // UNIQUE(A1:A5) -> {10,20,5,30}; SUM should drop the duplicate 20.
    let grid_u: [[&str; 3]; 5] = [
        ["10", "", "=SUM(UNIQUE(A1:A5))"],
        ["20", "", ""],
        ["5", "", ""],
        ["20", "", ""],
        ["30", "", ""],
    ];
    let mut sheet_u = create_sheet(&grid_u);
    sheet_u.commit(None).unwrap();
    assert_float_close(&sheet_u.get_result_data(&CellRef::new(0, 2)), 65.0, 1e-9);

    // SORT(A1:A5, 1, -1) descending -> first element should be the max, 30.
    let grid_s: [[&str; 3]; 5] = [
        ["10", "", "=INDEX(SORT(A1:A5,1,-1),1)"],
        ["20", "", ""],
        ["5", "", ""],
        ["20", "", ""],
        ["30", "", ""],
    ];
    let mut sheet_s = create_sheet(&grid_s);
    sheet_s.commit(None).unwrap();
    assert_float_close(&sheet_s.get_result_data(&CellRef::new(0, 2)), 30.0, 1e-9);

    // SORTBY(A1:A5, B1:B5, -1): sort A by B descending; B's max (50) is row1 (A=10).
    let grid_sb: [[&str; 3]; 5] = [
        ["10", "50", "=INDEX(SORTBY(A1:A5,B1:B5,-1),1)"],
        ["20", "0", ""],
        ["5", "1", ""],
        ["20", "1", ""],
        ["30", "1", ""],
    ];
    let mut sheet_sb = create_sheet(&grid_sb);
    sheet_sb.commit(None).unwrap();
    assert_float_close(&sheet_sb.get_result_data(&CellRef::new(0, 2)), 10.0, 1e-9);

    // FILTER(A1:A5, B1:B5) with a genuine boolean helper range (not a
    // broadcast comparison -- this engine's comparison operators don't
    // broadcast across a range, a separate pre-existing limitation).
    let grid_f: [[&str; 3]; 5] = [
        ["10", "0", "=SUM(FILTER(A1:A5,B1:B5))"],
        ["20", "1", ""],
        ["5", "0", ""],
        ["20", "1", ""],
        ["30", "1", ""],
    ];
    let mut sheet_f = create_sheet(&grid_f);
    sheet_f.commit(None).unwrap();
    assert_float_close(&sheet_f.get_result_data(&CellRef::new(0, 2)), 70.0, 1e-9);

    // TRIMRANGE(A1:A5) with no blank/error padding is a pass-through.
    assert_float_close(&eval_formula("=SUM(TRIMRANGE(SEQUENCE(3,3)))"), 45.0, 1e-9);
}

#[test]
fn test_fuzz_unique_distinguishes_numeric_text_from_numbers() {
    // Harvested from fuzz/fuzz_excel.py seed 993170: UNIQUE keeps text "3"
    // distinct from numeric 3. SUM then ignores the text and sums 3 + 10.
    let grid = [
        ["\"3\"", "=SUM(UNIQUE(A1:A5))"],
        ["OUiVTqS", ""],
        ["", ""],
        ["3", ""],
        ["10", ""],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    assert_float_close(&sheet.get_result_data(&CellRef::new(0, 1)), 13.0, 1e-9);
}

#[test]
fn test_sort_and_sortby_always_place_blanks_last() {
    // Both SORT and SORTBY always place blanks last, regardless of sort direction.
    let grid: [[&str; 4]; 5] = [
        ["-215.8", "10", "", "=INDEX(SORT(A1:A5,1,-1),1)"],
        ["", "20", "4", "=INDEX(SORTBY(B1:B5,C1:C5,-1),1)"],
        ["-100", "30", "3", ""],
        ["-240.97", "40", "2", ""],
        ["-88", "50", "1", ""],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    // SORT(A1:A5,1,-1): A2 is blank, sorted last regardless of direction,
    // so the largest real number (-88) is first.
    assert_float_close(&sheet.get_result_data(&CellRef::new(0, 3)), -88.0, 1e-9);
    // SORTBY(B1:B5,C1:C5,-1): C1 is blank, sorted last regardless of
    // direction, so the row with C's largest real value (row2, C=4, B=20)
    // comes first.
    assert_float_close(&sheet.get_result_data(&CellRef::new(1, 3)), 20.0, 1e-9);
}

#[test]
fn test_randarray_respects_shape_bounds_and_whole_number_flag() {
    match eval_formula("=RANDARRAY(2,3,10,20,TRUE)") {
        ResultData::List(rows) => {
            assert_eq!(rows.len(), 2);
            for row in &rows {
                match row {
                    ResultData::List(cells) => {
                        assert_eq!(cells.len(), 3);
                        for cell in cells {
                            let v = match cell {
                                ResultData::Float(f) => *f,
                                ResultData::Integer(i) => *i as f64,
                                other => panic!("expected numeric cell, got {other:?}"),
                            };
                            assert!((10.0..=20.0).contains(&v), "{v} out of [10,20]");
                            assert_eq!(v.fract(), 0.0, "expected a whole number, got {v}");
                        }
                    }
                    other => panic!("expected a row List, got {other:?}"),
                }
            }
        }
        other => panic!("expected a List of rows, got {other:?}"),
    }
}

#[test]
fn test_implicit_intersection_operator_on_ranges() {
    let grid: [[&str; 4]; 3] = [
        ["10", "20", "=@A:A", ""],
        ["30", "40", "=@A:A", ""],
        ["50", "=@A1:B1", "=@A1:B2", ""],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    assert_float_close(&sheet.get_result_data(&CellRef::new(0, 2)), 10.0, 1e-9);
    assert_float_close(&sheet.get_result_data(&CellRef::new(1, 2)), 30.0, 1e-9);
    assert_float_close(&sheet.get_result_data(&CellRef::new(2, 1)), 20.0, 1e-9);
    assert!(matches!(
        sheet.get_result_data(&CellRef::new(2, 2)),
        ResultData::Error(ref e) if e == "#VALUE!"
    ));
}

#[test]
fn test_fuzz_xlsx_anchorarray_reads_dynamic_array_anchor() {
    let grid = [["=SEQUENCE(3)", "=SUM(_xlfn.ANCHORARRAY(A1))"]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    assert_float_close(&sheet.get_result_data(&CellRef::new(0, 1)), 6.0, 1e-9);
}

#[test]
fn test_fuzz_xlsx_single_uses_formula_row() {
    let grid = [["10", ""], ["20", "=_xlfn.SINGLE(A1:A2)"]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    assert_float_close(&sheet.get_result_data(&CellRef::new(1, 1)), 20.0, 1e-9);
}

#[test]
fn test_spill_operator_reads_dynamic_array_anchor() {
    let grid: [[&str; 3]; 2] = [
        ["=SEQUENCE(3)", "=SUM(A1#)", "=INDEX(A1#,2)"],
        ["5", "=A2#", ""],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    assert_float_close(&sheet.get_result_data(&CellRef::new(0, 1)), 6.0, 1e-9);
    assert_float_close(&sheet.get_result_data(&CellRef::new(0, 2)), 2.0, 1e-9);
    assert!(matches!(
        sheet.get_result_data(&CellRef::new(1, 1)),
        ResultData::Error(ref e) if e == "#REF!"
    ));
}

#[test]
fn test_fuzz_anchorarray_scalar_string_formula_sum_ignores_text() {
    let grid = [
        ["=(IF(1 > 0, \"hello\", \"world\"))", ""],
        ["=SUM((_xlfn.ANCHORARRAY($A$1)))", ""],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    assert_float_close(&sheet.get_result_data(&CellRef::new(1, 0)), 0.0, 1e-9);
}

#[test]
fn test_fuzz_anchorarray_scalar_number_formula() {
    let grid = [
        ["=42", ""],
        ["=_xlfn.ANCHORARRAY(A1)", "=SUM(_xlfn.ANCHORARRAY(A1))"],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    assert_float_close(&sheet.get_result_data(&CellRef::new(1, 0)), 42.0, 1e-9);
    assert_float_close(&sheet.get_result_data(&CellRef::new(1, 1)), 42.0, 1e-9);
}
