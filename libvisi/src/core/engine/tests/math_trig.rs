use super::*;

#[test]
fn test_trig_and_hyperbolic_functions() {
    let grid = [[
        "=DEGREES(PI())",
        "=RADIANS(180)",
        "=SINH(0)",
        "=COSH(0)",
        "=TANH(0)",
        "=SQRTPI(4)",
    ]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    let r1 = sheet.get_result_data(&CellRef::new(0, 0));
    assert!(matches!(r1, ResultData::Float(v) if (v - 180.0).abs() < 1e-6));

    let r2 = sheet.get_result_data(&CellRef::new(0, 1));
    assert!(matches!(r2, ResultData::Float(v) if (v - std::f64::consts::PI).abs() < 1e-6));

    let r3 = sheet.get_result_data(&CellRef::new(0, 2));
    assert!(matches!(r3, ResultData::Float(v) if v.abs() < 1e-6));

    let r4 = sheet.get_result_data(&CellRef::new(0, 3));
    assert!(matches!(r4, ResultData::Float(v) if (v - 1.0).abs() < 1e-6));

    let r5 = sheet.get_result_data(&CellRef::new(0, 4));
    assert!(matches!(r5, ResultData::Float(v) if v.abs() < 1e-6));

    let r6 = sheet.get_result_data(&CellRef::new(0, 5));
    assert!(
        matches!(r6, ResultData::Float(v) if (v - (4.0 * std::f64::consts::PI).sqrt()).abs() < 1e-6)
    );
}

#[test]
fn test_rounding_and_integers() {
    let grid = [[
        "=EVEN(3)",
        "=ODD(4)",
        "=MROUND(10, 3)",
        "=QUOTIENT(10, 3)",
        "=SIGN(-5)",
        "=TRUNC(3.14159, 2)",
    ]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    let r1 = sheet.get_result_data(&CellRef::new(0, 0));
    assert!(matches!(r1, ResultData::Float(v) if (v - 4.0).abs() < 1e-6));

    let r2 = sheet.get_result_data(&CellRef::new(0, 1));
    assert!(matches!(r2, ResultData::Float(v) if (v - 5.0).abs() < 1e-6));

    let r3 = sheet.get_result_data(&CellRef::new(0, 2));
    assert!(matches!(r3, ResultData::Float(v) if (v - 9.0).abs() < 1e-6));

    let r4 = sheet.get_result_data(&CellRef::new(0, 3));
    assert!(matches!(r4, ResultData::Float(v) if (v - 3.0).abs() < 1e-6));

    let r5 = sheet.get_result_data(&CellRef::new(0, 4));
    assert!(matches!(r5, ResultData::Float(v) if (v + 1.0).abs() < 1e-6));

    let r6 = sheet.get_result_data(&CellRef::new(0, 5));
    #[allow(clippy::approx_constant)]
    let expected = 3.14; // TRUNC(3.14159, 2), not an approximation of PI
    assert!(matches!(r6, ResultData::Float(v) if (v - expected).abs() < 1e-6));
}

#[test]
fn test_base_conversions_and_roman() {
    let grid = [[
        "=BASE(255, 16)",
        "=DECIMAL(\"FF\", 16)",
        "=ARABIC(\"MCMXCIX\")",
        "=ROMAN(1999)",
    ]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    let r1 = sheet.get_result_data(&CellRef::new(0, 0));
    assert!(matches!(r1, ResultData::String(ref s) if s == "FF"));

    let r2 = sheet.get_result_data(&CellRef::new(0, 1));
    assert!(matches!(r2, ResultData::Float(v) if (v - 255.0).abs() < 1e-6));

    let r3 = sheet.get_result_data(&CellRef::new(0, 2));
    assert!(matches!(r3, ResultData::Float(v) if (v - 1999.0).abs() < 1e-6));

    let r4 = sheet.get_result_data(&CellRef::new(0, 3));
    assert!(matches!(r4, ResultData::String(ref s) if s == "MCMXCIX"));
}

#[test]
fn test_combinatorics_and_factors() {
    let grid = [[
        "=COMBIN(5, 2)",
        "=COMBINA(5, 2)",
        "=FACT(5)",
        "=FACTDOUBLE(5)",
        "=GCD(12, 18, 24)",
        "=LCM(4, 6)",
    ]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    let r1 = sheet.get_result_data(&CellRef::new(0, 0));
    assert!(matches!(r1, ResultData::Float(v) if (v - 10.0).abs() < 1e-6));

    let r2 = sheet.get_result_data(&CellRef::new(0, 1));
    assert!(matches!(r2, ResultData::Float(v) if (v - 15.0).abs() < 1e-6));

    let r3 = sheet.get_result_data(&CellRef::new(0, 2));
    assert!(matches!(r3, ResultData::Float(v) if (v - 120.0).abs() < 1e-6));

    let r4 = sheet.get_result_data(&CellRef::new(0, 3));
    assert!(matches!(r4, ResultData::Float(v) if (v - 15.0).abs() < 1e-6));

    let r5 = sheet.get_result_data(&CellRef::new(0, 4));
    assert!(matches!(r5, ResultData::Float(v) if (v - 6.0).abs() < 1e-6));

    let r6 = sheet.get_result_data(&CellRef::new(0, 5));
    assert!(matches!(r6, ResultData::Float(v) if (v - 12.0).abs() < 1e-6));
}

#[test]
fn test_array_and_matrix_functions() {
    let grid = [
        ["1", "2", "0", "0"],
        ["3", "4", "0", "0"],
        [
            "=SUMPRODUCT(A1:B1, A2:B2)",
            "=SUMSQ(A1:B2)",
            "=POWER(2, 10)",
            "=LOG(1000, 10)",
        ],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    // SUMPRODUCT([1, 2], [3, 4]) = 1*3 + 2*4 = 11
    let r1 = sheet.get_result_data(&CellRef::new(2, 0));
    assert!(matches!(r1, ResultData::Float(v) if (v - 11.0).abs() < 1e-6));

    // SUMSQ(1, 2, 3, 4) = 1 + 4 + 9 + 16 = 30
    let r2 = sheet.get_result_data(&CellRef::new(2, 1));
    assert!(matches!(r2, ResultData::Float(v) if (v - 30.0).abs() < 1e-6));

    // POWER(2, 10) = 1024
    let r3 = sheet.get_result_data(&CellRef::new(2, 2));
    assert!(matches!(r3, ResultData::Float(v) if (v - 1024.0).abs() < 1e-6));

    // LOG(1000, 10) = 3
    let r4 = sheet.get_result_data(&CellRef::new(2, 3));
    assert!(matches!(r4, ResultData::Float(v) if (v - 3.0).abs() < 1e-6));
}

#[test]
fn test_ceiling_floor_honor_significance_argument() {
    // Legacy 2-arg CEILING/FLOOR were completely ignoring their second
    // (significance) argument and just calling f64::ceil()/floor() --
    // e.g. CEILING(63.55, 5) returned 64 (plain ceil), not a multiple of
    // 5 at all, when it should round up to the nearest multiple of 5
    // (65). Found via differential fuzzing against real Excel.
    let grid = [["=CEILING(63.55, 5)", "=FLOOR(16.34, 10)"]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    let r1 = sheet.get_result_data(&CellRef::new(0, 0));
    assert!(matches!(r1, ResultData::Float(v) if (v - 65.0).abs() < 1e-9));

    let r2 = sheet.get_result_data(&CellRef::new(0, 1));
    assert!(matches!(r2, ResultData::Float(v) if (v - 10.0).abs() < 1e-9));
}

#[test]
fn test_gcd_lcm_error_on_non_numeric_argument() {
    // Unlike SUM/AVERAGE-style aggregates, real Excel's GCD/LCM don't
    // silently ignore a non-numeric cell in a range argument -- they
    // return #VALUE!. visi used to flatten with the same lenient logic
    // SUM uses, silently dropping the text cell and computing GCD/LCM
    // over whatever numbers were left.
    let grid = [["\"not a number\"", "6", "=GCD(A1:B1)", "=LCM(A1:B1)"]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    let r1 = sheet.get_result_data(&CellRef::new(0, 2));
    assert!(matches!(r1, ResultData::Error(ref e) if e == "#VALUE!"));

    let r2 = sheet.get_result_data(&CellRef::new(0, 3));
    assert!(matches!(r2, ResultData::Error(ref e) if e == "#VALUE!"));
}

#[test]
fn test_mmult_error_on_non_numeric_cell() {
    // MMULT used to coerce a non-numeric cell in either operand to 0
    // (via to_f64(..).unwrap_or(0.0)) instead of propagating #VALUE! the
    // way real Excel does.
    let grid = [
        ["1", "2"],
        ["3", "4"],
        ["\"x\"", "5"],
        ["6", "7"],
        ["=INDEX(MMULT(A1:B2, A3:B4), 1, 1)", ""],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    let r = sheet.get_result_data(&CellRef::new(4, 0));
    assert!(matches!(r, ResultData::Error(ref e) if e == "#VALUE!"));
}

#[test]
fn test_complex_number_trig_exp_log_functions_are_actually_computed() {
    // IMCOS/IMCOSH/IMCOT/IMCSC/IMCSCH/IMEXP/IMLN/IMLOG10/IMLOG2/IMPOWER/
    // IMSEC/IMSECH/IMSIN/IMSINH/IMSQRT/IMTAN were all a single stub arm
    // that just echoed the input string back unchanged (`Ok(ResultData::
    // String(t))`), so e.g. IMCOS("0") returned the text "0" by
    // coincidence, not because it computed cos(0)=1. Found via
    // differential fuzzing (every one of these functions mismatched
    // against real Excel on every run).
    let grid = [[
        "=IMCOS(0)",
        "=IMSIN(0)",
        "=IMTAN(0)",
        "=IMEXP(0)",
        "=IMLN(1)",
        "=IMSQRT(-1)",
        "=IMPOWER(2, 2)",
    ]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    let r1 = sheet.get_result_data(&CellRef::new(0, 0));
    assert!(
        matches!(r1, ResultData::String(ref s) if s == "1"),
        "IMCOS(0) = {r1:?}"
    );

    let r2 = sheet.get_result_data(&CellRef::new(0, 1));
    assert!(
        matches!(r2, ResultData::String(ref s) if s == "0"),
        "IMSIN(0) = {r2:?}"
    );

    let r3 = sheet.get_result_data(&CellRef::new(0, 2));
    assert!(
        matches!(r3, ResultData::String(ref s) if s == "0"),
        "IMTAN(0) = {r3:?}"
    );

    let r4 = sheet.get_result_data(&CellRef::new(0, 3));
    assert!(
        matches!(r4, ResultData::String(ref s) if s == "1"),
        "IMEXP(0) = {r4:?}"
    );

    let r5 = sheet.get_result_data(&CellRef::new(0, 4));
    assert!(
        matches!(r5, ResultData::String(ref s) if s == "0"),
        "IMLN(1) = {r5:?}"
    );

    // sqrt(-1) = i. Real Excel reports this as "6.12323399573677E-17+i",
    // not a clean "i": the polar form's angle is f64's rounded pi/2, and
    // cos(pi/2) in f64 is ~6.12e-17 rather than exactly 0. visi matches
    // that verbatim on purpose -- an earlier version snapped the
    // negligible component to zero, which reads "more correct" but
    // disagrees with the thing this engine is trying to be compatible
    // with (confirmed by probing real Excel directly).
    let r6 = sheet.get_result_data(&CellRef::new(0, 5));
    assert!(
        matches!(r6, ResultData::String(ref s) if s == "6.12323399573677E-17+i"),
        "IMSQRT(-1) = {r6:?}"
    );

    // 2^2 = 4 (a real number stays a plain real, no spurious "+0i")
    let r7 = sheet.get_result_data(&CellRef::new(0, 6));
    assert!(
        matches!(r7, ResultData::String(ref s) if s == "4"),
        "IMPOWER(2,2) = {r7:?}"
    );
}

#[test]
fn test_complex_number_formatting_uses_excel_precision_not_raw_f64() {
    // format_complex used to interpolate the real/imaginary f64 parts
    // directly (`format!("{}", c.re)`), which prints full f64 precision
    // (e.g. 0.1 + 0.2 as raw f64 addition prints "0.30000000000000004")
    // instead of Excel's 15-significant-digit display rules. Every IM*
    // result with a non-exact float component mismatched real Excel by a
    // handful of ULPs in the last few digits.
    let grid = [["=IMSUM(0.1, 0.2)"]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    let r = sheet.get_result_data(&CellRef::new(0, 0));
    assert!(
        matches!(r, ResultData::String(ref s) if s == "0.3"),
        "IMSUM(0.1,0.2) = {r:?}"
    );
}

#[test]
fn test_complex_tan_cot_stay_precise_for_large_imaginary_parts() {
    // IMTAN/IMCOT used to be computed as the complex quotient
    // sin(z)/cos(z), which loses most of its significant digits once
    // |Im z| grows: both operands pick up components of order
    // cosh(Im z) (already ~550 by Im z = 7) and the result's real part is
    // the tiny residual left after those large nearly-equal terms cancel.
    // Against real Excel that showed up as agreement to only ~10
    // significant digits. Now computed from the double-angle identities,
    // where every intermediate is the same magnitude as the result.
    //
    // Reference values are verbatim real-Excel output. The tolerance is
    // per-component relative 1e-14 rather than string equality: the last
    // displayed digit can still differ by one ulp depending on the exact
    // order f64 operations happen in, which is not something either
    // engine can control, while the bug this guards against was four
    // orders of magnitude larger than that.
    fn parts(s: &str) -> (f64, f64) {
        let c = crate::core::engineering::parse_complex(s).expect("parses as complex");
        (c.re, c.im)
    }
    let grid = [[
        "=IMTAN(\"-1-7i\")",
        "=IMTAN(\"2+6i\")",
        "=IMTAN(\"9-2i\")",
        "=IMCOT(\"-7-7j\")",
        "=IMCOT(\"9-2j\")",
        "=IMCOT(\"-6+8i\")",
    ]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    for (col, expected) in [
        (0, "-1.51221489579179E-06-1.00000069207519i"),
        (1, "-9.29998518085996E-06+1.00000803223943i"),
        (2, "-0.0268511331123278-0.975735876254188i"),
        (3, "-0.0000016474373058602+1.00000022740052j"),
        (4, "-0.0281818376780689+1.02409198026131j"),
        (5, "1.20766677090395E-07-1.00000018992652i"),
    ] {
        let got = sheet.get_result_data(&CellRef::new(0, col));
        let ResultData::String(ref got_s) = got else {
            panic!("col {col}: expected a complex string, got {got:?}");
        };
        let (gr, gi) = parts(got_s);
        let (er, ei) = parts(expected);
        let close = |a: f64, b: f64| (a - b).abs() <= 1e-14 * b.abs().max(1e-300);
        assert!(
            close(gr, er) && close(gi, ei),
            "col {col}: got {got_s}, want {expected}"
        );
        // The i/j suffix must survive verbatim.
        assert_eq!(
            got_s.chars().last(),
            expected.chars().last(),
            "col {col}: suffix changed ({got_s} vs {expected})"
        );
    }
}

#[test]
fn test_roman_concise_forms_match_excel() {
    // ROMAN's `form` argument was ignored entirely, so all five forms
    // rendered as classic notation. Every expectation below is verbatim
    // real-Excel output.
    for (n, form, expected) in [
        (990.0, 0.0, "CMXC"),
        (990.0, 1.0, "LMXL"),
        (990.0, 2.0, "XM"),
        (990.0, 3.0, "XM"),
        (990.0, 4.0, "XM"),
        (1481.0, 0.0, "MCDLXXXI"),
        (1481.0, 1.0, "MLDXXXI"),
        (1481.0, 4.0, "MLDXXXI"),
        (1999.0, 0.0, "MCMXCIX"),
        (1999.0, 1.0, "MLMVLIV"),
        (1999.0, 2.0, "MXMIX"),
        (1999.0, 3.0, "MVMIV"),
        (1999.0, 4.0, "MIM"),
        (499.0, 0.0, "CDXCIX"),
        (499.0, 1.0, "LDVLIV"),
        (499.0, 2.0, "XDIX"),
        (499.0, 3.0, "VDIV"),
        (499.0, 4.0, "ID"),
        (45.0, 0.0, "XLV"),
        (45.0, 1.0, "VL"),
    ] {
        let got = crate::core::math_trig::roman(n, Some(form));
        assert_eq!(
            got,
            Ok(expected.to_string()),
            "ROMAN({n}, {form}) = {got:?}, want {expected:?}"
        );
    }
}
