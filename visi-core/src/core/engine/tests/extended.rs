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
fn test_datevalue_text_cell_refs_report_days_between() {
    // The NVDA put implied-volatility workbook stores its date inputs as text
    // cells (imported source is quoted to keep them from becoming typed date
    // serials), then computes C9 as `DATEVALUE(B9)-DATEVALUE(B5)`.
    let grid = [
        ["", "", ""],
        ["", "", ""],
        ["", "", ""],
        ["", "", ""],
        ["", "\"2026-08-12\"", ""],
        ["", "", ""],
        ["", "", ""],
        ["", "Expiration Date", "DTE (Days)"],
        ["", "\"2026-08-21\"", "=DATEVALUE(B9)-DATEVALUE(B5)"],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    let got = sheet.get_result_data(&CellRef::new(8, 2));
    assert!(
        matches!(got, ResultData::Float(v) if (v - 9.0).abs() < 1e-9),
        "C9 = {got:?}, want 9"
    );
}

#[test]
fn test_datevalue_and_timevalue_reject_typed_date_time_cells() {
    // Entering an unquoted date/time-looking value into Excel stores a numeric
    // serial, not text. DATEVALUE/TIMEVALUE accept text only, so references to
    // those typed cells return #VALUE! even though VALUE(A1) and arithmetic can
    // still use the serial. Measured via Excel for Mac AppleScript automation.
    let grid = [[
        "2026-08-12",     // A1
        "=DATEVALUE(A1)", // B1
        "12:00:00",       // C1
        "=TIMEVALUE(C1)", // D1
    ]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    for col in [1, 3] {
        let got = sheet.get_result_data(&CellRef::new(0, col));
        assert!(
            matches!(got, ResultData::Error(ref e) if e == "#VALUE!"),
            "col {col}: {got:?}"
        );
    }
}

#[test]
fn test_fuzz_datevalue_and_timevalue_reject_numeric_serials() {
    // DATEVALUE/TIMEVALUE parse text only. Excel returns #VALUE! for both
    // literal numbers and numeric cell references; accepting serials here was
    // found by differential fuzzing against real Excel on seed 354657.
    let grid = [[
        "46195",                // A1
        "=DATEVALUE(46195)",    // B1
        "=DATEVALUE(A1)",       // C1
        "=TIMEVALUE(46195.25)", // D1
        "=TIMEVALUE(A1)",       // E1
    ]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    for col in 1..=4 {
        let got = sheet.get_result_data(&CellRef::new(0, col));
        assert!(
            matches!(got, ResultData::Error(ref e) if e == "#VALUE!"),
            "col {col}: {got:?}"
        );
    }
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
// The expected values are the 30-digit reference evaluations the comment
// below describes. They are kept at full width on purpose: the digits past
// f64 are what make them checkable against the oracle, and truncating to
// the shortest round-tripping literal would hide where they came from.
#[allow(clippy::excessive_precision)]
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

#[test]
fn test_forecast_ets_reproduces_excel_on_well_posed_series() {
    // The ETS family used to be hardcoded stubs (SEASONALITY always 1,
    // STAT always 0.5, CONFINT always 0, and FORECAST.ETS falling back to
    // a straight linear fit). It is now a real AAA Holt-Winters model.
    //
    // Excel's alpha/beta/gamma come out of a proprietary optimizer that an
    // independent implementation cannot be expected to reproduce digit for
    // digit on noisy data. What *is* checkable -- and what these cases
    // cover -- is a series the model fits perfectly, where the forecast is
    // the same for any sane parameter triple and Excel's own answer is the
    // exact continuation. Every expectation below was read from real Excel.
    let t8: Vec<f64> = (1..=8).map(|i| i as f64).collect();
    let linear: Vec<f64> = vec![10.0, 12.0, 14.0, 16.0, 18.0, 20.0, 22.0, 24.0];

    assert_eq!(
        crate::core::ets::detect_period(&linear),
        0,
        "no seasonality"
    );
    let m = crate::core::ets::prepare(&linear, &t8, 1.0, true).expect("fits");
    for (h, want) in [(1usize, 26.0), (2, 28.0), (3, 30.0)] {
        let got = m.forecast(h);
        assert!((got - want).abs() < 1e-9, "h={h}: got {got}, want {want}");
    }
    // A perfect fit leaves no error and no prediction interval.
    for stat in [4usize, 5, 6, 7] {
        assert!(m.stat(stat).expect("stat").abs() < 1e-9, "stat {stat}");
    }
    assert!(m.confint(1, 0.95).expect("confint").abs() < 1e-9);
    // Excel reports exactly these for the degenerate (perfect-fit) case.
    assert!((m.alpha - 0.9).abs() < 1e-12, "alpha = {}", m.alpha);
    assert!((m.beta - 0.001).abs() < 1e-12, "beta = {}", m.beta);

    // Trend plus a repeating period-4 season: Excel detects 4 and forecasts
    // the exact continuation, 18 then 28.
    let t16: Vec<f64> = (1..=16).map(|i| i as f64).collect();
    let seasonal: Vec<f64> = vec![
        10.0, 20.0, 15.0, 5.0, 12.0, 22.0, 17.0, 7.0, 14.0, 24.0, 19.0, 9.0, 16.0, 26.0, 21.0, 11.0,
    ];
    assert_eq!(crate::core::ets::detect_period(&seasonal), 4, "period 4");
    let m2 = crate::core::ets::prepare(&seasonal, &t16, 1.0, true).expect("fits");
    for (h, want) in [(1usize, 18.0), (2, 28.0)] {
        let got = m2.forecast(h);
        assert!(
            (got - want).abs() < 1e-9,
            "seasonal h={h}: got {got}, want {want}"
        );
    }

    // A zigzag with no trend is period 2 (Excel agrees).
    let zig: Vec<f64> = vec![10.0, 14.0, 11.0, 17.0, 15.0, 20.0, 18.0, 24.0, 21.0, 27.0];
    assert_eq!(crate::core::ets::detect_period(&zig), 2);
}

#[test]
fn test_forecast_ets_timeline_validation() {
    let vals = vec![1.0, 2.0, 3.0, 4.0];
    // Irregular spacing has no constant step.
    let ragged = vec![1.0, 2.0, 4.5, 9.0];
    assert_eq!(
        crate::core::ets::build_series(&vals, &ragged, true).err(),
        Some("#NUM!".to_string())
    );
    // Mismatched lengths.
    assert_eq!(
        crate::core::ets::build_series(&vals, &[1.0, 2.0], true).err(),
        Some("#N/A".to_string())
    );
    // A gap is interpolated, not rejected: 1, 2, _, 4 over t = 1,2,4.
    let s = crate::core::ets::build_series(&[1.0, 2.0, 4.0], &[1.0, 2.0, 4.0], true)
        .expect("gap is completed");
    assert_eq!(s.values.len(), 4);
    assert!((s.values[2] - 3.0).abs() < 1e-9, "{:?}", s.values);
    assert!((s.step - 1.0).abs() < 1e-12);

    // A target at or before the end of the timeline has no horizon.
    assert_eq!(
        crate::core::ets::horizon(1.0, 1.0, 4, 4.0).err(),
        Some("#NUM!".to_string())
    );
    assert_eq!(crate::core::ets::horizon(1.0, 1.0, 4, 6.0), Ok(2));
}
