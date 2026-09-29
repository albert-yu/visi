use super::*;

// --- LAMBDA family (LAMBDA, MAP, REDUCE, SCAN, BYROW, BYCOL, MAKEARRAY,
// ISOMITTED) --------------------------------------------------------------
//
// MAP and REDUCE results were confirmed against real Microsoft Excel.
// BYROW/BYCOL/MAKEARRAY, a bare uninvoked LAMBDA, and even SCAN could not
// be reliably confirmed that way: any dynamic-array-spilling formula --
// even a bare `=SEQUENCE(3)` with no LAMBDA involved at all -- breaks this
// environment's Excel AppleScript automation bridge intermittently
// (confirmed directly, and not specific to any one function; the same
// formula sometimes succeeds and sometimes doesn't across repeat runs).
// Their expected values below are independently verified instead, either
// by hand-calculated arithmetic or against Microsoft's own documented
// SCAN example (SCAN(0, {1,2,3}, LAMBDA(a,v,a+v)) => {1,3,6}).

#[test]
fn test_lambda_bare_is_uncallable() {
    // The parser has no `(expr)(args)` immediate-invocation syntax (that
    // would require calling an arbitrary sub-expression, not just a bare
    // identifier), so an uninvoked, unnamed LAMBDA can't produce a value,
    // matching Excel's own #CALC! for this case.
    assert!(matches!(
        eval1("=LAMBDA(x, x*2)"),
        ResultData::Error(ref e) if e.contains("#CALC!")
    ));
}

#[test]
fn test_isomitted_best_effort() {
    // Best-effort implementation (see the doc comment in evaluate_function
    // above `ISOMITTED`'s dispatch): every lambda invocation path here
    // always supplies exactly as many values as declared parameters, so a
    // declared, in-scope parameter is never actually omitted -- this only
    // exercises the "identifier not found in scope at all" case.
    assert!(matches!(
        eval1("=INDEX(MAP(1, LAMBDA(x, ISOMITTED(x))), 1)"),
        ResultData::Boolean(false)
    ));
    assert!(matches!(
        eval1("=ISOMITTED(some_undeclared_name)"),
        ResultData::Boolean(true)
    ));
}

#[test]
fn test_map_single_and_multiple_arrays_match_real_excel() {
    let grid: [[&str; 3]; 3] = [
        ["10", "1", "=INDEX(MAP(A1:A3, LAMBDA(x, x*2)), 2)"],
        ["20", "2", "=INDEX(MAP(A1:A3, B1:B3, LAMBDA(x,y, x+y)), 3)"],
        ["30", "3", ""],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    assert_float_close(&sheet.get_result_data(&CellRef::new(0, 2)), 40.0, 1e-9); // 20*2
    assert_float_close(&sheet.get_result_data(&CellRef::new(1, 2)), 33.0, 1e-9); // 30+3
}

#[test]
fn test_reduce_and_scan_match_real_excel() {
    let grid: [[&str; 2]; 3] = [
        ["1", "=REDUCE(0, A1:A3, LAMBDA(acc,v, acc+v))"],
        ["2", "=INDEX(SCAN(0, A1:A3, LAMBDA(acc,v, acc+v)), 3)"],
        ["3", ""],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    assert_float_close(&sheet.get_result_data(&CellRef::new(0, 1)), 6.0, 1e-9);
    assert_float_close(&sheet.get_result_data(&CellRef::new(1, 1)), 6.0, 1e-9);
}

#[test]
fn test_reduce_two_arg_form_seeds_from_first_element() {
    // Real Excel's REDUCE always takes 3 arguments, but initial_value is
    // documented as optional; since this parser has no syntax to express
    // "omitted argument" (no leading/trailing empty comma support), a
    // plain 2-argument REDUCE(array, lambda) is accepted as that
    // omitted-initial-value form: the array's own first element seeds the
    // accumulator, and the rest are folded in.
    let grid: [[&str; 2]; 3] = [
        ["1", "=REDUCE(A1:A3, LAMBDA(acc,v, acc+v))"],
        ["2", ""],
        ["3", ""],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    assert_float_close(&sheet.get_result_data(&CellRef::new(0, 1)), 6.0, 1e-9);
}

#[test]
fn test_byrow_and_bycol_sum_hand_verified() {
    // Data: [[1,2,3],[4,5,6],[7,8,9]]. Row sums: 6, 15, 24.
    // Column sums: 12, 15, 18. Not verifiable against real Excel here (see
    // module doc comment above) -- hand-verified arithmetic instead.
    let grid: [[&str; 5]; 3] = [
        ["1", "2", "3", "", ""],
        ["4", "5", "6", "", ""],
        ["7", "8", "9", "", ""],
    ];
    let mut sheet = create_sheet(&grid);
    for i in 0..3 {
        sheet.set_cell_src(
            i,
            3,
            format!("=INDEX(BYROW(A1:C3, LAMBDA(zz, SUM(zz))), {})", i + 1),
        );
        sheet.set_cell_src(
            i,
            4,
            format!("=INDEX(BYCOL(A1:C3, LAMBDA(zz, SUM(zz))), {})", i + 1),
        );
    }
    sheet.commit(None).unwrap();
    for (i, expected_row) in [6.0, 15.0, 24.0].iter().enumerate() {
        assert_float_close(
            &sheet.get_result_data(&CellRef::new(i, 3)),
            *expected_row,
            1e-9,
        );
    }
    for (i, expected_col) in [12.0, 15.0, 18.0].iter().enumerate() {
        assert_float_close(
            &sheet.get_result_data(&CellRef::new(i, 4)),
            *expected_col,
            1e-9,
        );
    }
}

#[test]
fn test_makearray_builds_row_major_flat_array_hand_verified() {
    // MAKEARRAY(2, 3, LAMBDA(r,c, r*10+c)) should build
    // [[11,12,13],[21,22,23]] flattened row-major: [11,12,13,21,22,23].
    let grid: [[&str; 1]; 6] = [
        ["=INDEX(MAKEARRAY(2, 3, LAMBDA(rr,cc, rr*10+cc)), 1)"],
        ["=INDEX(MAKEARRAY(2, 3, LAMBDA(rr,cc, rr*10+cc)), 2)"],
        ["=INDEX(MAKEARRAY(2, 3, LAMBDA(rr,cc, rr*10+cc)), 3)"],
        ["=INDEX(MAKEARRAY(2, 3, LAMBDA(rr,cc, rr*10+cc)), 4)"],
        ["=INDEX(MAKEARRAY(2, 3, LAMBDA(rr,cc, rr*10+cc)), 5)"],
        ["=INDEX(MAKEARRAY(2, 3, LAMBDA(rr,cc, rr*10+cc)), 6)"],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    for (i, expected) in [11.0, 12.0, 13.0, 21.0, 22.0, 23.0].iter().enumerate() {
        assert_float_close(&sheet.get_result_data(&CellRef::new(i, 0)), *expected, 1e-9);
    }
}

#[test]
fn test_let_binds_names_in_sequence_and_rejects_duplicate_names() {
    assert_float_close(&eval1("=LET(x, 5, x * 2)"), 10.0, 1e-9);
    // Later pairs can reference earlier ones in the same LET.
    assert_float_close(&eval1("=LET(x, 5, y, x + 1, x + y)"), 11.0, 1e-9);
    assert!(matches!(
        eval1("=LET(x, 1, x, 2, x)"),
        ResultData::Error(ref e) if e == "#VALUE!"
    ));
}
