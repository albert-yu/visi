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
    let grid = [["2026-08-12", "=DATEVALUE(A1)", "12:00:00", "=TIMEVALUE(C1)"]];
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
    let grid = [[
        "46195",
        "=DATEVALUE(46195)",
        "=DATEVALUE(A1)",
        "=TIMEVALUE(46195.25)",
        "=TIMEVALUE(A1)",
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
#[allow(clippy::excessive_precision)]
fn test_besselj_stays_accurate_where_excel_does_not() {
    for (x, n, expected) in [
        (9.59_f64, 1.0_f64, 0.141754162508486556734783214599_f64),
        (8.72, 2.0, 0.079558608902434556482112499322),
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
    let sheet = Sheet::new(SheetInit::default());
    for f in ["=LOG(3.14, \"E\")", "=MOD(5, \"E\")", "=MOD(\"E\", 5)"] {
        let got = sheet.eval(f, None).unwrap().0;
        assert!(
            matches!(got, ResultData::Error(ref e) if e == "#VALUE!"),
            "{f} = {got:?}, want #VALUE!"
        );
    }
    let got = sheet.eval("=LOG(1000)", None).unwrap().0;
    assert!(
        matches!(got, ResultData::Float(v) if (v - 3.0).abs() < 1e-9),
        "{got:?}"
    );
    let got = sheet.eval("=MOD(5, 0)", None).unwrap().0;
    assert!(
        matches!(got, ResultData::Error(ref e) if e == "#DIV/0!"),
        "{got:?}"
    );
}

#[test]
fn test_forecast_ets_reproduces_excel_on_well_posed_series() {
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
    for stat in [4usize, 5, 6, 7] {
        assert!(m.stat(stat).expect("stat").abs() < 1e-9, "stat {stat}");
    }
    assert!(m.confint(1, 0.95).expect("confint").abs() < 1e-9);
    assert!((m.alpha - 0.9).abs() < 1e-12, "alpha = {}", m.alpha);
    assert!((m.beta - 0.001).abs() < 1e-12, "beta = {}", m.beta);

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

    let zig: Vec<f64> = vec![10.0, 14.0, 11.0, 17.0, 15.0, 20.0, 18.0, 24.0, 21.0, 27.0];
    assert_eq!(crate::core::ets::detect_period(&zig), 2);
}

#[test]
fn test_forecast_ets_timeline_validation() {
    let vals = vec![1.0, 2.0, 3.0, 4.0];
    let ragged = vec![1.0, 2.0, 4.5, 9.0];
    assert_eq!(
        crate::core::ets::build_series(&vals, &ragged, true).err(),
        Some("#NUM!".to_string())
    );
    assert_eq!(
        crate::core::ets::build_series(&vals, &[1.0, 2.0], true).err(),
        Some("#N/A".to_string())
    );
    let s = crate::core::ets::build_series(&[1.0, 2.0, 4.0], &[1.0, 2.0, 4.0], true)
        .expect("gap is completed");
    assert_eq!(s.values.len(), 4);
    assert!((s.values[2] - 3.0).abs() < 1e-9, "{:?}", s.values);
    assert!((s.step - 1.0).abs() < 1e-12);

    assert_eq!(
        crate::core::ets::horizon(1.0, 1.0, 4, 4.0).err(),
        Some("#NUM!".to_string())
    );
    assert_eq!(crate::core::ets::horizon(1.0, 1.0, 4, 6.0), Ok(2));
}

#[test]
fn test_date_serial_matches_real_excel_reference_values() {
    // ymd_to_serial's Excel-epoch offset was off by one for every date
    // after the fictitious Feb 29, 1900 (serial 60) -- found via
    // differential fuzzing against real Excel while validating the new
    // bond-pricing functions below, which all depend on serial-date
    // arithmetic. These are well-known, independently verifiable
    // Excel serial numbers (e.g. 25569 = Unix epoch, a widely cited
    // Excel/Unix-timestamp conversion constant), confirmed again here
    // directly against real Microsoft Excel via the differential fuzzer.
    assert_float_close(&eval_formula("=DATE(1900,1,1)"), 1.0, 1e-9);
    assert_float_close(&eval_formula("=DATE(1970,1,1)"), 25569.0, 1e-9);
    assert_float_close(&eval_formula("=DATE(2000,1,1)"), 36526.0, 1e-9);
    assert_float_close(&eval_formula("=DATE(2007,11,21)"), 39407.0, 1e-9);
    assert_float_close(&eval_formula("=DATE(2021,1,1)"), 44197.0, 1e-9);
}

#[test]
fn test_date_functions_match_documented_excel_examples() {
    // Jan 1, 2024 (serial 45292) is a Monday. WORKDAY skips both weekend
    // days landing on Mon 1/8/2024 (serial 45299) five working days later
    // (Tue-Fri, then Mon).
    assert_float_close(&eval_formula("=WORKDAY(DATE(2024,1,1),5)"), 45299.0, 1e-9);
    // Inclusive of both endpoints, excluding the Sat/Sun in between:
    // Jan 1, 2, 3, 4, 5, 8 = 6 working days.
    assert_float_close(
        &eval_formula("=NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,8))"),
        6.0,
        1e-9,
    );
    // EOMONTH(0) is the same month's last day; EOMONTH(1) rolls into
    // Feb 2024, a leap year (serial 45351 = Feb 29), so the last day is
    // the 29th.
    assert_float_close(
        &eval_formula("=DAY(EOMONTH(DATE(2024,1,15),0))"),
        31.0,
        1e-9,
    );
    assert_float_close(&eval_formula("=EOMONTH(DATE(2024,1,15),1)"), 45351.0, 1e-9);
    // WEEKNUM with the default return type (week starts Sunday): Jan 1,
    // 2024 (a Monday) is always week 1; the first Sunday (Jan 7) starts
    // week 2.
    assert_float_close(&eval_formula("=WEEKNUM(DATE(2024,1,1))"), 1.0, 1e-9);
    assert_float_close(&eval_formula("=WEEKNUM(DATE(2024,1,7))"), 2.0, 1e-9);
    // Microsoft's own DAYS360 documentation example (US/NASD method).
    assert_float_close(
        &eval_formula("=DAYS360(DATE(2011,1,30),DATE(2011,2,1))"),
        1.0,
        1e-9,
    );
}

#[test]
fn test_besselk_bessely_match_known_reference_values() {
    // BESSELK/BESSELY match known reference values (K_n/Y_n diverge as x -> 0
    // while I_n/J_n stay finite there). Expected values are well-known
    // constants (Abramowitz & Stegun tables).
    assert_float_close(&eval_formula("=BESSELK(1,0)"), 0.4210244382, 1e-8);
    assert_float_close(&eval_formula("=BESSELK(1,1)"), 0.6019072301, 1e-8);
    assert_float_close(&eval_formula("=BESSELY(1,0)"), 0.0882569642, 1e-8);
    assert_float_close(&eval_formula("=BESSELY(1,1)"), -0.7812128213, 1e-8);
    // Sanity check the still-correct BESSELI/BESSELJ weren't disturbed.
    assert_float_close(&eval_formula("=BESSELI(1,0)"), 1.2660658778, 1e-8);
    assert_float_close(&eval_formula("=BESSELJ(1,0)"), 0.7651976866, 1e-8);
}

#[test]
fn test_complex_number_functions_round_trip() {
    assert_eq!(eval_formula("=COMPLEX(3,4)").to_string(), "3+4i");
    assert_float_close(&eval_formula("=IMABS(\"3+4i\")"), 5.0, 1e-9);
    assert_float_close(&eval_formula("=IMREAL(\"3+4i\")"), 3.0, 1e-9);
    assert_float_close(&eval_formula("=IMAGINARY(\"3+4i\")"), 4.0, 1e-9);
    assert_eq!(eval_formula("=IMCONJUGATE(\"3+4i\")").to_string(), "3-4i");
    assert_eq!(
        eval_formula("=IMSUM(\"3+4i\",\"1-2i\")").to_string(),
        "4+2i"
    );
    assert_eq!(
        eval_formula("=IMSUB(\"3+4i\",\"1-2i\")").to_string(),
        "2+6i"
    );
    // (3+4i)(1-2i) = 3-6i+4i-8i^2 = 3-2i+8 = 11-2i
    assert_eq!(
        eval_formula("=IMPRODUCT(\"3+4i\",\"1-2i\")").to_string(),
        "11-2i"
    );
    // (3+4i)/(1-2i) = (3+4i)(1+2i)/5 = (3+6i+4i-8)/5 = (-5+10i)/5 = -1+2i
    assert_eq!(
        eval_formula("=IMDIV(\"3+4i\",\"1-2i\")").to_string(),
        "-1+2i"
    );
}

#[test]
fn test_cube_webservice_image_report_unavailable_connections_not_echo_stub_args() {
    // None has a local data source this engine can serve (a live OLAP cube
    // connection, actual network access, real image decoding); the error
    // codes match what real Excel shows once its equivalent live connection/
    // resource is unavailable (#N/A for the CUBE* family, mirroring
    // RTD/STOCKHISTORY just below; #VALUE! for WEBSERVICE/IMAGE, per
    // Microsoft's own docs).
    for formula in [
        "=CUBEKPIMEMBER(\"conn\",\"kpi\",1)",
        "=CUBEMEMBER(\"conn\",\"member\")",
        "=CUBEMEMBERPROPERTY(\"conn\",\"member\",\"prop\")",
        "=CUBERANKEDMEMBER(\"conn\",\"set\",1)",
        "=CUBESET(\"conn\",\"set\")",
        "=CUBESETCOUNT(\"set\")",
        "=CUBEVALUE(\"conn\",\"member\")",
    ] {
        assert!(
            matches!(eval_formula(formula), ResultData::Error(ref e) if e == "#N/A"),
            "expected #N/A for {formula}"
        );
    }
    assert!(matches!(
        eval_formula("=WEBSERVICE(\"https://example.com\")"),
        ResultData::Error(ref e) if e == "#VALUE!"
    ));
    assert!(matches!(
        eval_formula("=IMAGE(\"https://example.com/pic.png\")"),
        ResultData::Error(ref e) if e == "#VALUE!"
    ));
}

#[test]
fn test_stockhistory_and_rtd_report_unavailable_data_source() {
    // Neither has a local data source this engine can serve (a live
    // Microsoft stock-data cloud connection, a registered Windows COM RTD
    // server) -- #N/A matches real Excel's own display once its
    // equivalent live connection is unavailable.
    assert!(matches!(
        eval_formula("=STOCKHISTORY(\"MSFT\",\"2024-01-01\",\"2024-01-31\")"),
        ResultData::Error(ref e) if e == "#N/A"
    ));
    assert!(matches!(
        eval_formula("=RTD(\"prog.id\",\"server\",\"topic\")"),
        ResultData::Error(ref e) if e == "#N/A"
    ));
}
