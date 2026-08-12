use super::*;

#[test]
fn test_date_and_time_functions() {
    let grid = [[
        "=DATE(2024, 8, 3)",
        "=YEAR(DATE(2024, 8, 3))",
        "=MONTH(DATE(2024, 8, 3))",
        "=DAY(DATE(2024, 8, 3))",
        "=TIME(12, 30, 0)",
        "=HOUR(0.5)",
    ]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    let r1 = sheet.get_result_data(&CellRef::new(0, 0));
    assert!(matches!(r1, ResultData::Float(v) if (v - 45507.0).abs() < 10.0));

    let r2 = sheet.get_result_data(&CellRef::new(0, 1));
    assert!(matches!(r2, ResultData::Float(v) if (v - 2024.0).abs() < 1e-6));

    let r3 = sheet.get_result_data(&CellRef::new(0, 2));
    assert!(matches!(r3, ResultData::Float(v) if (v - 8.0).abs() < 1e-6));

    let r4 = sheet.get_result_data(&CellRef::new(0, 3));
    assert!(matches!(r4, ResultData::Float(v) if (v - 3.0).abs() < 1e-6));

    let r5 = sheet.get_result_data(&CellRef::new(0, 4));
    assert!(matches!(r5, ResultData::Float(v) if (v - 0.52083333).abs() < 1e-4));

    let r6 = sheet.get_result_data(&CellRef::new(0, 5));
    assert!(matches!(r6, ResultData::Float(v) if (v - 12.0).abs() < 1e-6));
}

#[test]
fn test_engineering_functions() {
    let grid = [[
        "=BIN2DEC(\"1010\")",
        "=DEC2HEX(255)",
        "=BITAND(6, 3)",
        "=DELTA(5, 5)",
        "=GESTEP(10, 5)",
        "=CONVERT(1, \"km\", \"m\")",
    ]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    let r1 = sheet.get_result_data(&CellRef::new(0, 0));
    assert!(matches!(r1, ResultData::Float(v) if (v - 10.0).abs() < 1e-6));

    let r2 = sheet.get_result_data(&CellRef::new(0, 1));
    assert!(matches!(r2, ResultData::String(ref s) if s == "FF"));

    let r3 = sheet.get_result_data(&CellRef::new(0, 2));
    assert!(matches!(r3, ResultData::Float(v) if (v - 2.0).abs() < 1e-6));

    let r4 = sheet.get_result_data(&CellRef::new(0, 3));
    assert!(matches!(r4, ResultData::Float(v) if (v - 1.0).abs() < 1e-6));

    let r5 = sheet.get_result_data(&CellRef::new(0, 4));
    assert!(matches!(r5, ResultData::Float(v) if (v - 1.0).abs() < 1e-6));

    let r6 = sheet.get_result_data(&CellRef::new(0, 5));
    assert!(matches!(r6, ResultData::Float(v) if (v - 1000.0).abs() < 1e-6));
}

#[test]
fn test_information_logical_lookup_web_functions() {
    let grid = [[
        "=ISEVEN(4)",
        "=ISODD(5)",
        "=TYPE(100)",
        "=XOR(TRUE, FALSE)",
        "=ADDRESS(1, 1)",
        "=ENCODEURL(\"hello world\")",
    ]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    let r1 = sheet.get_result_data(&CellRef::new(0, 0));
    assert!(matches!(r1, ResultData::Boolean(true)));

    let r2 = sheet.get_result_data(&CellRef::new(0, 1));
    assert!(matches!(r2, ResultData::Boolean(true)));

    let r3 = sheet.get_result_data(&CellRef::new(0, 2));
    assert!(matches!(r3, ResultData::Float(v) if (v - 1.0).abs() < 1e-6));

    let r4 = sheet.get_result_data(&CellRef::new(0, 3));
    assert!(matches!(r4, ResultData::Boolean(true)));

    let r5 = sheet.get_result_data(&CellRef::new(0, 4));
    assert!(matches!(r5, ResultData::String(ref s) if s == "$A$1"));

    let r6 = sheet.get_result_data(&CellRef::new(0, 5));
    assert!(matches!(r6, ResultData::String(ref s) if s == "hello%20world"));
}

#[test]
fn test_datevalue_and_timevalue_parse_common_formats() {
    // DATEVALUE and TIMEVALUE were both wired to core::text::value (the
    // generic VALUE() numeric-string parser), which doesn't understand
    // date/time strings at all -- both always returned #VALUE! for any
    // real date or time string. Found via differential fuzzing (every
    // DATEVALUE/TIMEVALUE call mismatched real Excel on every run).
    let grid = [[
        "=DATEVALUE(\"2000-01-01\")",
        "=DATEVALUE(\"1/1/2000\")",
        "=TIMEVALUE(\"12:00:00\")",
        "=TIMEVALUE(\"6:00 AM\")",
    ]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    let r1 = sheet.get_result_data(&CellRef::new(0, 0));
    assert!(
        matches!(r1, ResultData::Float(v) if (v - 36526.0).abs() < 1e-9),
        "{r1:?}"
    );

    let r2 = sheet.get_result_data(&CellRef::new(0, 1));
    assert!(
        matches!(r2, ResultData::Float(v) if (v - 36526.0).abs() < 1e-9),
        "{r2:?}"
    );

    let r3 = sheet.get_result_data(&CellRef::new(0, 2));
    assert!(
        matches!(r3, ResultData::Float(v) if (v - 0.5).abs() < 1e-9),
        "{r3:?}"
    );

    let r4 = sheet.get_result_data(&CellRef::new(0, 3));
    assert!(
        matches!(r4, ResultData::Float(v) if (v - 0.25).abs() < 1e-9),
        "{r4:?}"
    );
}

#[test]
fn test_address_mixed_reference_types_not_swapped() {
    // abs_num 2 ("absolute row; relative column", e.g. "A$1") and 3
    // ("relative row; absolute column", e.g. "$A1") were swapped in
    // address_fn's match arms.
    let grid = [["=ADDRESS(1, 17, 2)", "=ADDRESS(1, 17, 3)"]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    let r1 = sheet.get_result_data(&CellRef::new(0, 0));
    assert!(
        matches!(r1, ResultData::String(ref s) if s == "Q$1"),
        "{r1:?}"
    );

    let r2 = sheet.get_result_data(&CellRef::new(0, 1));
    assert!(
        matches!(r2, ResultData::String(ref s) if s == "$Q1"),
        "{r2:?}"
    );
}

#[test]
fn test_error_type_and_ifna_receive_the_actual_error() {
    // ERROR.TYPE and IFNA both need to see the error value in their
    // argument, not have it propagate past them before they get a
    // chance to inspect it -- the generic "any error in an argument
    // makes the whole call error" short-circuit didn't exclude them, so
    // ERROR.TYPE(1/0) itself evaluated to #DIV/0! instead of 2, and
    // IFNA's second argument was evaluated eagerly (not lazily, unlike
    // IFERROR) so IFNA(5, 1/0) incorrectly propagated #DIV/0! instead of
    // returning 5 without ever needing the second argument.
    let grid = [["=ERROR.TYPE(1/0)", "=IFNA(5, 1/0)", "=IFNA(NA(), 99)"]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    let r1 = sheet.get_result_data(&CellRef::new(0, 0));
    assert!(
        matches!(r1, ResultData::Float(v) if (v - 2.0).abs() < 1e-9),
        "{r1:?}"
    );

    let r2 = sheet.get_result_data(&CellRef::new(0, 1));
    assert!(
        matches!(r2, ResultData::Float(v) if (v - 5.0).abs() < 1e-9),
        "{r2:?}"
    );

    let r3 = sheet.get_result_data(&CellRef::new(0, 2));
    assert!(
        matches!(r3, ResultData::Float(v) if (v - 99.0).abs() < 1e-9),
        "{r3:?}"
    );
}

#[test]
fn test_datedif_counts_only_completed_intervals() {
    // DATEDIF used to compute each unit independently -- plain `m2 - m1`
    // for "M"/"YM", plain `d2 - d1` for "MD", and `(end - start) % 365`
    // for "YD" -- which overcounts by a month whenever the end day of
    // month hasn't reached the start's, and could even go negative
    // ("MD" reported -9 for one of these pairs). DATEDIF counts
    // *completed* intervals, so a short final month has to borrow the
    // length of the month preceding the end date. Every expected value
    // below was read out of real Excel.
    for (start, end, unit, expected) in [
        (40314.0, 45719.0, "M", 177.0),
        (42733.0, 45935.0, "M", 105.0),
        (40237.0, 44396.0, "MD", 21.0),
        (41714.0, 45088.0, "YM", 2.0),
        (42314.0, 45737.0, "YM", 4.0),
        (40269.0, 44942.0, "YM", 9.0),
        (41102.0, 43892.0, "YD", 234.0),
        (40792.0, 44087.0, "YD", 7.0),
    ] {
        let got = crate::core::date_fn::datedif(start, end, unit);
        assert_eq!(
            got,
            Ok(expected),
            "DATEDIF({start}, {end}, \"{unit}\") = {got:?}, want {expected}"
        );
    }
}

#[test]
fn test_besselj_stays_accurate_where_excel_does_not() {
    // Differential fuzzing flagged BESSELJ as a mismatch against real
    // Excel for larger arguments -- but arbitrating the two against
    // 60-significant-digit reference values (Decimal evaluation of the
    // same ascending series) showed *visi* is the accurate one and Excel
    // is not: at x = 9.59, order 1, visi's relative error is ~1e-13 while
    // Excel's is ~1.3e-5. The fuzz generator therefore caps its Bessel
    // arguments below where Excel degrades; this test pins visi's own
    // accuracy so that cap can never quietly mask a real regression here.
    for (x, n, expected) in [
        (9.59_f64, 1.0_f64, 0.141754162508486556734783214599_f64),
        (8.72, 2.0, 0.079558608902434556482112499322),
        // Excel is already off by 3.8e-7 here, at a modest argument --
        // its accuracy degrades with the *order* as well as x, which is
        // why no argument cap makes it usable as a reference.
        (2.95, 3.0, 0.300141005800689674499808380416),
    ] {
        let got = crate::core::engineering::besselj(x, n).expect("BESSELJ computes");
        let rel = (got - expected).abs() / expected.abs();
        assert!(
            rel < 1e-12,
            "BESSELJ({x}, {n}) = {got}, want {expected} (relative error {rel:e})"
        );
    }
}

#[test]
fn test_present_but_non_numeric_optional_argument_is_value_error() {
    // An *optional* numeric argument that is absent falls back to its
    // default, but one that is present and non-numeric is #VALUE!. These
    // were conflated by the `.and_then(to_f64).unwrap_or(default)` shape,
    // so LOG silently computed base 10 for a text base and MOD silently
    // treated a text operand as 0.
    let sheet = Sheet::new(SheetInit::default());
    for f in ["=LOG(3.14, \"E\")", "=MOD(5, \"E\")", "=MOD(\"E\", 5)"] {
        let got = sheet.eval(f, None).unwrap().0;
        assert!(
            matches!(got, ResultData::Error(ref e) if e == "#VALUE!"),
            "{f} = {got:?}, want #VALUE!"
        );
    }
    // An omitted base still defaults to 10.
    let got = sheet.eval("=LOG(1000)", None).unwrap().0;
    assert!(
        matches!(got, ResultData::Float(v) if (v - 3.0).abs() < 1e-9),
        "{got:?}"
    );
    // And a zero divisor is #DIV/0!.
    let got = sheet.eval("=MOD(5, 0)", None).unwrap().0;
    assert!(
        matches!(got, ResultData::Error(ref e) if e == "#DIV/0!"),
        "{got:?}"
    );
}
