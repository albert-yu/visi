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
fn test_fuzz_quotient_zero_does_not_keep_negative_sign_in_atan2() {
    // Harvested from fuzz/fuzz_excel.py seed 567480: QUOTIENT(PI(), -37)
    // displays as zero and behaves as +0 in ATAN2's quadrant choice.
    match eval_one("=ATAN2(RADIANS(-45), QUOTIENT(PI(), -37))") {
        ResultData::Float(v) => assert!((v - std::f64::consts::PI).abs() < 1e-12, "got {v}"),
        other => panic!("expected pi, got {other:?}"),
    }
}

#[test]
fn test_fuzz_power_type_checks_base_before_exponent_error() {
    // Harvested from fuzz/fuzz_excel.py seeds 61472 and 148208: POWER checks
    // its base's type before propagating a later exponent error, so this is
    // #VALUE!, not #N/A.
    match eval_one("=POWER(\"C\", NA())") {
        ResultData::Error(e) => assert_eq!(e, "#VALUE!"),
        other => panic!("expected #VALUE!, got {other:?}"),
    }
}

#[test]
fn test_fuzz_power_negative_base_rejects_huge_exponent() {
    // Harvested from fuzz/fuzz_excel.py seed 151238: SINH(-424.13) is an
    // enormous negative integer-valued double, but Excel still rejects a
    // negative POWER base once the exponent is outside its supported range.
    match eval_one("=POWER(-95, SINH(-424.13))") {
        ResultData::Error(e) => assert_eq!(e, "#NUM!"),
        other => panic!("expected #NUM!, got {other:?}"),
    }
}

#[test]
fn test_fuzz_sin_cos_tan_and_reciprocals_num_error_past_2_pow_27() {
    // Measured directly against real Windows Excel: SIN/COS/TAN refuse an
    // argument at or beyond 2^27 (134217728) radians with #NUM! -- past
    // that magnitude a double can no longer resolve which multiple of
    // 2*pi the value is near, so any answer would be numerically
    // meaningless. 2^27 - 1 still computes; 2^27 does not. CSC/SEC/COT
    // inherit the same boundary since they're built on SIN/COS/TAN.
    // fuzz/fuzz_excel.py seed 676008 hit this via CSC(F4^47), where
    // F4^47 is on the order of 1e101: visi returned a plain float, real
    // Excel #NUM!.
    let grid = [[
        "=SIN(134217727)",
        "=SIN(134217728)",
        "=COS(134217728)",
        "=TAN(134217728)",
        "=CSC(134217728)",
        "=SEC(134217728)",
        "=COT(134217728)",
    ]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    match sheet.get_result_data(&CellRef::new(0, 0)) {
        ResultData::Float(_) => {}
        other => panic!("SIN(2^27 - 1): expected a number, got {other:?}"),
    }
    for (col, name) in [
        (1, "SIN"),
        (2, "COS"),
        (3, "TAN"),
        (4, "CSC"),
        (5, "SEC"),
        (6, "COT"),
    ] {
        match sheet.get_result_data(&CellRef::new(0, col)) {
            ResultData::Error(e) => assert_eq!(e, "#NUM!", "{name}(2^27)"),
            other => panic!("{name}(2^27): expected #NUM!, got {other:?}"),
        }
    }
}

#[test]
fn test_fuzz_coth_does_not_overflow_to_nan_for_a_large_argument() {
    // Harvested from fuzz/fuzz_excel.py seed 711993: COTH(47692.3), where
    // 47692.3 is VARPA's own ordinary, finite result -- not an extreme
    // input. COTH's old implementation, `x.cosh() / x.sinh()`, has both
    // sides overflow to `f64::INFINITY` well before |x| gets anywhere near
    // where coth itself misbehaves (~710), leaving `inf / inf = NaN`,
    // which this engine's NaN guard turns into `#NUM!` -- for an `x`
    // where real Excel returns a perfectly good answer near +-1 (COTH of
    // any large positive number this size is `1`, confirmed against real
    // Excel). `tanh` saturates to +-1 directly with no such overflow.
    assert_eq!(num("=COTH(47692.3)"), 1.0);
    // Ordinary arguments are untouched.
    assert!((num("=COTH(1)") - 1.3130352854993312).abs() < 1e-9);
}

#[test]
fn test_fuzz_gcd_first_arg_type_error_wins_over_later_arg_error() {
    // Harvested from fuzz/fuzz_excel.py seed 751310:
    // GCD(AND(J1>0, Sheet1[[#Headers],[C]]<100), CORREL(F3:J4, F2:G2)) --
    // the AND(...) result is a boolean (GCD rejects booleans outright,
    // `GCD(TRUE, 8)` is `#VALUE!`) and CORREL's mismatched-size ranges
    // (1x3 vs 4x5) are `#N/A`. Same first-argument-wins shape as LOG/ATAN2
    // above: GCD/LCM walk their arguments in order and reject the first
    // non-numeric one, so the boolean should win with `#VALUE!`, not
    // CORREL's `#N/A` (measured via win32com: `GCD(TRUE, NA())` is
    // `#VALUE!` in real Excel, matching `GCD(TRUE, 8)`).
    match eval_one("=GCD(TRUE, NA())") {
        ResultData::Error(e) => assert_eq!(e, "#VALUE!"),
        other => panic!("expected #VALUE!, got {other:?}"),
    }
    match eval_one("=LCM(TRUE, NA())") {
        ResultData::Error(e) => assert_eq!(e, "#VALUE!"),
        other => panic!("expected #VALUE!, got {other:?}"),
    }
}

#[test]
fn test_fuzz_isna_ifna_see_a_hard_err_from_a_nested_first_arg_error() {
    // Harvested from fuzz/fuzz_excel.py, several seeds all shaped like
    // ISNA(LOG(<errors #N/A>, 25)) / IFNA(ATAN2(<errors #N/A>, x), y):
    // once LOG/ATAN2 are exempted from the generic "first error found in
    // any argument" pre-check (so their own first-argument type check can
    // run instead, see the LOG/ATAN2 exemption above), a first argument
    // that is *itself already an error* propagates out of
    // `to_f64_arg(...)?` as a hard `Err`, not the `Ok(ResultData::Error(_))`
    // every ordinary function argument normally produces. ISNA's and
    // IFNA's own hand-rolled early-return branches only checked for
    // `Ok(ResultData::Error(_))`, so that hard `Err` fell through to their
    // catch-alls (`_ => false` / a bare `?` that just propagated the `Err`
    // further) instead of being recognized as the `#N/A` it actually was.
    match eval_one("=ISNA(LOG(NA(), 25))") {
        ResultData::Boolean(b) => assert!(b),
        other => panic!("expected TRUE, got {other:?}"),
    }
    match eval_one("=ISNA(ATAN2(NA(), 5))") {
        ResultData::Boolean(b) => assert!(b),
        other => panic!("expected TRUE, got {other:?}"),
    }
    assert_eq!(num("=IFNA(LOG(NA(), 25), 999)"), 999.0);
    assert_eq!(num("=IFNA(ATAN2(NA(), 5), 999)"), 999.0);
    // A non-#N/A hard error from the same path still surfaces as itself,
    // not swallowed into FALSE / silently replaced.
    match eval_one("=ISNA(LOG(\"C\", 25))") {
        ResultData::Boolean(b) => assert!(!b),
        other => panic!("expected FALSE, got {other:?}"),
    }
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
fn test_gcd_lcm_treats_scalar_blank_cell_as_omitted_not_zero() {
    // A blank scalar reference passed to GCD/LCM is dropped, not coerced to 0 --
    // real Excel gives LCM(1, <blank>) = 1 (as if LCM(1)), not
    // LCM(1, 0) = 0. Measured with fuzz/fuzz_excel.py seed 308076,
    // where LCM(1, I2) (I2 blank) came back visi=0, Excel=1.
    let grid = [["1", "", "=LCM(A1, B1)", "=GCD(A1, B1)"]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    let lcm = sheet.get_result_data(&CellRef::new(0, 2));
    assert!(
        matches!(lcm, ResultData::Float(v) if (v - 1.0).abs() < 1e-9),
        "LCM(1, <blank>) should be 1, got {lcm:?}"
    );

    let gcd = sheet.get_result_data(&CellRef::new(0, 3));
    assert!(
        matches!(gcd, ResultData::Float(v) if (v - 1.0).abs() < 1e-9),
        "GCD(1, <blank>) should be 1, got {gcd:?}"
    );
}

#[test]
fn test_fuzz_lcm_range_blank_cells_count_as_zero() {
    // Harvested from fuzz/fuzz_excel.py seed 747962. Unlike a blank scalar
    // reference, a blank inside a range argument participates as zero, so
    // the LCM of the range is zero even when the non-blank cells alone
    // would have produced 40568.
    let grid = [["461.4064", "88", "=LCM(A1:B2)"], ["", "", "=LCM(A1:B1)"]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    let range_with_blanks = sheet.get_result_data(&CellRef::new(0, 2));
    assert!(
        matches!(range_with_blanks, ResultData::Float(v) if v.abs() < 1e-9),
        "LCM(range containing blanks) should be 0, got {range_with_blanks:?}"
    );

    let no_blanks = sheet.get_result_data(&CellRef::new(1, 2));
    assert!(
        matches!(no_blanks, ResultData::Float(v) if (v - 40568.0).abs() < 1e-9),
        "LCM(non-blank range) should be 40568, got {no_blanks:?}"
    );
}

#[test]
fn test_fuzz_gcd_lcm_one_cell_blank_range_is_missing_operand() {
    // Harvested from fuzz/fuzz_excel.py seed 579827: a one-cell blank range is
    // a missing operand (#VALUE!), while the existing multi-cell range rule
    // above still counts blanks as zero.
    let grid = [
        ["", "=GCD(A1:A1)", "=LCM(A1:A1)", "=GCD(A1:A2)"],
        ["", "", "", ""],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    for col in 1..=2 {
        match sheet.get_result_data(&CellRef::new(0, col)) {
            ResultData::Error(e) => assert_eq!(e, "#VALUE!", "column {col}"),
            other => panic!("expected #VALUE! in column {col}, got {other:?}"),
        }
    }
    assert!(matches!(
        sheet.get_result_data(&CellRef::new(0, 3)),
        ResultData::Float(v) if v.abs() < 1e-9
    ));
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

fn eval_one(source: &str) -> ResultData {
    let sheet = Sheet::new(SheetInit::default());
    sheet.eval(source, None).unwrap().0
}

fn num(source: &str) -> f64 {
    match eval_one(source) {
        ResultData::Float(f) => f,
        ResultData::Integer(i) => i as f64,
        other => panic!("expected a number for {source}, got {other:?}"),
    }
}

#[test]
fn test_days360_us_method_pulls_february_month_ends_to_the_30th() {
    // Serial 40602 is 2011-02-28, a February month-end; 42543 is
    // 2016-06-24. The US (NASD) method moves that start day to the 30th
    // and so counts two days fewer than the European method. Real Excel:
    //   DAYS360(40602, 42543, FALSE) = 1912   (also the default)
    //   DAYS360(40602, 42543, TRUE)  = 1914
    // Before the February rule was implemented both spellings returned
    // 1914, because the two methods only diverge on a month-end February.
    assert_eq!(num("=DAYS360(40602, 42543, FALSE)"), 1912.0);
    assert_eq!(num("=DAYS360(40602, 42543)"), 1912.0);
    assert_eq!(num("=DAYS360(40602, 42543, TRUE)"), 1914.0);
}

#[test]
fn test_supplied_blank_argument_is_zero_not_the_default() {
    // An *omitted* optional argument takes its default; an argument that
    // is supplied but blank is 0. Real Excel draws the line sharply:
    //   LOG(100)        = 2         base defaults to 10
    //   LOG(100, blank) = #NUM!     base is 0
    //   LOG(1, blank)   = #NUM!     still #NUM!, not 0
    //   MROUND(10, blank) = 0
    // Z90 is empty in a fresh sheet.
    assert_eq!(num("=LOG(100)"), 2.0);
    assert_eq!(num("=MROUND(10, Z90)"), 0.0);
    for src in ["=LOG(100, Z90)", "=LOG(1, Z90)"] {
        match eval_one(src) {
            ResultData::Error(e) => assert_eq!(e, "#NUM!", "for {src}"),
            other => panic!("expected #NUM! for {src}, got {other:?}"),
        }
    }
    // Base 1 stays #DIV/0! rather than #NUM! -- log(n)/log(1) divides by 0.
    match eval_one("=LOG(1, 1)") {
        ResultData::Error(e) => assert_eq!(e, "#DIV/0!"),
        other => panic!("expected #DIV/0!, got {other:?}"),
    }
}

#[test]
fn test_power_of_zero_to_the_zero_is_num_error() {
    // Excel declines to pick a value for 0^0 -- both spellings are #NUM!.
    // The `^` operator already did this; the POWER function returned 1,
    // so POWER(<blank>, <blank>) quietly evaluated to 1 and turned
    // OR(POWER(blank, blank) > 0, ...) into TRUE where Excel says #NUM!.
    for src in ["=POWER(0, 0)", "=0^0", "=POWER(Y90, Y91)"] {
        match eval_one(src) {
            ResultData::Error(e) => assert_eq!(e, "#NUM!", "for {src}"),
            other => panic!("expected #NUM! for {src}, got {other:?}"),
        }
    }
    // The neighbouring domain rules are unchanged.
    assert_eq!(num("=POWER(0, 2)"), 0.0);
    assert_eq!(num("=POWER(2, 0)"), 1.0);
    match eval_one("=POWER(0, -1)") {
        ResultData::Error(e) => assert_eq!(e, "#DIV/0!"),
        other => panic!("expected #DIV/0!, got {other:?}"),
    }
}

#[test]
fn test_sumproduct_treats_non_numeric_entries_as_zero() {
    // Real Excel: SUMPRODUCT(2, "abc") = 0 and SUMPRODUCT(A1:B1, A2:B2)
    // with a text cell in the second array = 3. Non-numeric entries count
    // as zero *and keep their slot*, so the arrays stay the same length
    // and the remaining terms still line up. Dropping them instead made
    // the first case #VALUE! -- one array of length 1 against one of
    // length 0 -- which then showed up as TYPE(...) = 16 rather than 1.
    assert_eq!(num("=SUMPRODUCT(2, \"abc\")"), 0.0);
    assert_eq!(num("=TYPE(SUMPRODUCT(2, \"abc\"))"), 1.0);

    let mut sheet = create_sheet(&[["1", "2", "=SUMPRODUCT(A1:B1, A2:B2)"], ["3", "=\"x\"", ""]]);
    sheet.commit(None).unwrap();
    match sheet.get_result_data(&CellRef::new(0, 2)) {
        ResultData::Float(f) => assert!((f - 3.0).abs() < 1e-9, "got {f}"),
        ResultData::Integer(i) => assert_eq!(i, 3),
        other => panic!("expected 3, got {other:?}"),
    }
}

#[test]
fn test_only_a_few_numeric_functions_reject_booleans() {
    // Excel's coercion is not uniform here, and the split does not follow
    // from anything about the functions -- it had to be probed one at a
    // time. These four answer #VALUE! to a boolean:
    for src in [
        "=FACTDOUBLE(TRUE)",
        "=SQRTPI(TRUE)",
        "=ERF(TRUE)",
        "=ERFC(TRUE)",
        "=BIN2DEC(TRUE)",
    ] {
        match eval_one(src) {
            ResultData::Error(e) => assert_eq!(e, "#VALUE!", "for {src}"),
            other => panic!("expected #VALUE! for {src}, got {other:?}"),
        }
    }

    // ... while their neighbours take TRUE as 1 without complaint. All of
    // these values are real Excel's.
    assert_eq!(num("=FACTDOUBLE(6)"), 48.0);
    assert_eq!(num("=FACT(TRUE)"), 1.0);
    assert_eq!(num("=SQRT(TRUE)"), 1.0);
    assert_eq!(num("=SIGN(TRUE)"), 1.0);
    assert_eq!(num("=INT(TRUE)"), 1.0);
    assert_eq!(num("=EVEN(TRUE)"), 2.0);
    assert_eq!(num("=ODD(TRUE)"), 1.0);
    assert_eq!(num("=LN(TRUE)"), 0.0);
    assert_eq!(num("=LOG10(TRUE)"), 0.0);
    assert_eq!(num("=GAMMALN(TRUE)"), 0.0);
    assert!((num("=EXP(TRUE)") - std::f64::consts::E).abs() < 1e-15);
    assert!((num("=DEGREES(TRUE)") - 57.29577951308232).abs() < 1e-13);
}

#[test]
fn test_days360_and_yearfrac_use_different_thirty_360_rules() {
    // Excel's DAYS360 function and its YEARFRAC basis 0 (the NASD
    // convention the bond functions share) genuinely disagree, which is
    // why visi implements them separately. Two rules differ:
    //
    //  - When *both* ends are February month-ends, YEARFRAC pulls the end
    //    date to the 30th and DAYS360 does not.
    //  - DAYS360's "end date on the 31st comes back to the 30th" rule
    //    tests the *adjusted* start day, YEARFRAC's the original -- so a
    //    February month-end start triggers it for one and not the other.
    //
    // Every pair below is a real-Excel value; the ones where the two
    // columns differ are exactly the cases that separate the rules.
    let cases: [(&str, &str, f64, f64); 12] = [
        ("DATE(2003,2,28)", "DATE(2005,2,28)", 718.0, 720.0),
        ("DATE(2004,2,29)", "DATE(2008,2,29)", 1439.0, 1440.0),
        ("DATE(2004,2,29)", "DATE(2005,2,28)", 358.0, 360.0),
        ("DATE(2003,2,28)", "DATE(2004,2,29)", 359.0, 360.0),
        ("DATE(2003,2,28)", "DATE(2005,3,31)", 750.0, 751.0),
        ("DATE(2003,1,31)", "DATE(2005,2,28)", 748.0, 748.0),
        ("DATE(2003,2,28)", "DATE(2005,2,27)", 717.0, 717.0),
        ("DATE(2003,3,31)", "DATE(2005,2,28)", 688.0, 688.0),
        ("DATE(2003,1,31)", "DATE(2005,3,31)", 780.0, 780.0),
        ("DATE(2003,3,15)", "DATE(2005,5,31)", 796.0, 796.0),
        ("DATE(2003,1,30)", "DATE(2005,3,31)", 780.0, 780.0),
        ("DATE(2003,4,30)", "DATE(2005,2,28)", 658.0, 658.0),
    ];
    for (start, end, days360, yearfrac) in cases {
        let d = num(&format!("=DAYS360({start}, {end}, FALSE)"));
        assert!(
            (d - days360).abs() < 1e-9,
            "DAYS360({start}, {end}) expected {days360}, got {d}"
        );
        let y = num(&format!("=YEARFRAC({start}, {end}, 0) * 360"));
        assert!(
            (y - yearfrac).abs() < 1e-6,
            "YEARFRAC({start}, {end}, 0) * 360 expected {yearfrac}, got {y}"
        );
    }
}

#[test]
fn test_oddlprice_treats_a_month_end_coupon_date_as_the_30th() {
    // On basis 0 the two ODDLPRICE spans that *end at a coupon date* --
    // the quasi-coupon period length and last-interest-to-maturity -- pull
    // a month-end end date back to the 30th, February's included. The
    // spans ending at the settlement date use the plain NASD count, so the
    // same date pair counts differently depending on its role.
    //
    // Every expected value is real Excel's. The pairs below were chosen to
    // separate the rule from the plain European one: a leap-year 28 Feb is
    // *not* a month end and must not be adjusted, while 29 Feb is.
    let cases: [(&str, &str, &str, f64, f64); 8] = [
        // last_interest, settlement, maturity, basis, expected
        (
            "DATE(2017,12,27)",
            "DATE(2018,1,4)",
            "DATE(2018,2,28)",
            0.0,
            100.40414916157073,
        ),
        (
            "DATE(2017,12,27)",
            "DATE(2018,1,4)",
            "DATE(2018,1,31)",
            0.0,
            100.17445487014778,
        ),
        (
            "DATE(2017,12,27)",
            "DATE(2018,1,4)",
            "DATE(2018,3,31)",
            0.0,
            100.59075984089594,
        ),
        // 2016 is a leap year: 28 Feb is not the month end, 29 Feb is.
        (
            "DATE(2015,12,27)",
            "DATE(2016,1,4)",
            "DATE(2016,2,28)",
            0.0,
            100.37619967431928,
        ),
        (
            "DATE(2015,12,27)",
            "DATE(2016,1,4)",
            "DATE(2016,2,29)",
            0.0,
            100.39711327585337,
        ),
        // A month-end *last_interest* shortens the quasi-coupon period.
        (
            "DATE(2018,2,28)",
            "DATE(2018,3,10)",
            "DATE(2018,4,30)",
            0.0,
            100.35615901837147,
        ),
        // Basis 4 stays the plain European count -- 29 Feb is not adjusted
        // there, which is what makes it differ from basis 0 above.
        (
            "DATE(2015,12,27)",
            "DATE(2016,1,4)",
            "DATE(2016,2,29)",
            4.0,
            100.38313951056004,
        ),
        (
            "DATE(2015,12,27)",
            "DATE(2016,1,4)",
            "DATE(2016,1,31)",
            4.0,
            100.18148895627503,
        ),
    ];
    for (last_interest, settlement, maturity, basis, expected) in cases {
        let got = num(&format!(
            "=ODDLPRICE({settlement}, {maturity}, {last_interest}, 0.0505, 0.0253, 100, 4, {basis})"
        ));
        assert!(
            (got - expected).abs() < 1e-9,
            "ODDLPRICE(li={last_interest}, mat={maturity}, basis={basis}) \
             expected {expected}, got {got}"
        );
    }
}

#[test]
fn test_mod_reports_num_once_the_quotient_stops_being_meaningful() {
    // Excel gives up on MOD once the quotient is large enough that
    // `n - d * INT(n / d)` is noise, rather than returning a number built
    // out of it. The cutoff is on the *quotient*: MOD over a huge dividend
    // is fine as long as the divisor is huge too.
    //
    // All expected values are real Excel's.
    for src in [
        "=MOD(POWER(28, 31), 3)",
        "=MOD(10000000000000, 3)",
        "=MOD(-1000000000000000, 3)",
    ] {
        match eval_one(src) {
            ResultData::Error(e) => assert_eq!(e, "#NUM!", "for {src}"),
            other => panic!("expected #NUM! for {src}, got {other:?}"),
        }
    }
    assert_eq!(num("=MOD(1000000000000, 3)"), 1.0);
    assert_eq!(num("=MOD(1000000000000000, 10000000)"), 0.0);
    assert_eq!(num("=MOD(1000000000000000, 1000000)"), 0.0);
    // 2^40 * 3 is inside the limit, 2^41 * 3 is past it.
    assert_eq!(num("=MOD(3298534883328, 3)"), 0.0);
    match eval_one("=MOD(6597069766656, 3)") {
        ResultData::Error(e) => assert_eq!(e, "#NUM!"),
        other => panic!("expected #NUM!, got {other:?}"),
    }
    // Ordinary MOD is untouched, including its sign convention.
    assert_eq!(num("=MOD(10, 3)"), 1.0);
    assert_eq!(num("=MOD(-10, 3)"), 2.0);
    assert_eq!(num("=MOD(10, -3)"), -2.0);
    assert_eq!(num("=MOD(TRUE, 2)"), 1.0);
}

#[test]
fn test_fuzz_mod_stays_exact_at_an_integer_quotient_boundary() {
    // Harvested from fuzz/fuzz_excel.py seed 550442: MOD(-47, 47 / -13).
    // True mathematics (47 and -13 taken as exact integers, not the
    // rounded double `47.0/-13.0` computed first) makes
    // `-47 / (47/-13) = 13` exactly, so `INT(quotient)` should be 13 and
    // the remainder exactly 0 -- and it stays exact even carried through
    // an actual `f64` division of the two doubles: `-47.0 / (47.0/-13.0)`
    // evaluates to precisely `13.0` in IEEE 754 double precision, no
    // rounding residue at all. Real Excel returns `-3.615384615384615`
    // (the divisor itself) instead, as if `INT(quotient)` had come out to
    // 12 rather than the mathematically exact 13 -- a real Excel
    // precision loss at this boundary, not a rounding convention
    // difference (see "docs/excel-discrepancies.md" section 15, which
    // this is another instance of). visi's 0 is correct and is left
    // alone.
    assert_eq!(num("=MOD(-47, (47 / -13))"), 0.0);
}

#[test]
fn test_fuzz_mod_tiny_power_against_negative_divisor_is_not_zero() {
    // Harvested from fuzz/fuzz_excel.py seed 747962:
    // MOD((-5 ^ -16), (-44 * 85)). The true dividend is a tiny positive
    // number and the divisor is -3740, so Excel's documented
    // n - d*INT(n/d) rule gives a remainder infinitesimally above -3740,
    // not 0. This is the same Excel precision loss documented in
    // docs/excel-discrepancies.md section 15.
    assert!((num("=MOD((-5 ^ -16), (-44 * 85))") + 3740.0).abs() < 1e-9);
}

#[test]
// `exact` is a reference value carried at more digits than f64 holds, so it
// can be read against the oracle that produced it; see the comment on it.
#[allow(clippy::excessive_precision)]
fn test_coupdaysnc_and_acoth_precision() {
    // COUPDAYSNC's span ends at a coupon date, so on basis 0 a month-end
    // coupon is pulled back to the 30th -- the same rule ODDLPRICE's
    // coupon-ended spans use. Settlement 2011-08-28 against a 2013-02-28
    // maturity has its next coupon on 2011-08-31, and real Excel counts 2
    // days there, not the 3 the plain NASD rule gives. COUPDAYBS is
    // unaffected: its span ends at the settlement date.
    let s = "DATE(2011,8,28)";
    let m = "EDATE(DATE(2011,8,28),18)";
    assert_eq!(num(&format!("=COUPDAYSNC({s}, {m}, 2, 0)")), 2.0);
    assert_eq!(num(&format!("=COUPDAYBS({s}, {m}, 2, 0)")), 178.0);
    assert_eq!(num(&format!("=COUPDAYS({s}, {m}, 2, 0)")), 180.0);
    // Other bases and a non-month-end coupon are unchanged.
    assert_eq!(
        num("=COUPDAYSNC(DATE(2003,12,21), EDATE(DATE(2003,12,21),108), 1, 1)"),
        366.0
    );
    assert_eq!(
        num("=COUPDAYSNC(DATE(2017,9,22), EDATE(DATE(2017,9,22),36), 2, 3)"),
        181.0
    );
    assert_eq!(
        num("=COUPDAYSNC(DATE(2011,8,15), EDATE(DATE(2011,8,15),18), 2, 0)"),
        180.0
    );

    // ACOTH via atanh(1/x). The 0.5 * ln((x+1)/(x-1)) form loses its
    // significant digits for large |x|: the true ACOTH(-165) is
    // -0.006060680266172405095..., and it returned
    // -0.006060680266172425 -- wrong from the 15th digit, which is
    // precisely where Excel's display lands.
    // Tolerance is relative: the old form was off by 3.3e-15 relative,
    // this one by under 2e-16.
    let exact = -0.0060606802661724050957;
    let got = num("=ACOTH(-165)");
    assert!(
        ((got - exact) / exact).abs() < 1e-15,
        "ACOTH(-165) expected {exact}, got {got}"
    );
    // Small |x| was never the problem, so this one is only a guard that the
    // rewrite didn't break the ordinary case. It is checked at the same
    // relative tolerance as the case above rather than an absolute 1e-16:
    // one ulp here is ~1.1e-16, so a sub-ulp bound demands a bit-exact
    // `atanh` and Apple's libm and glibc are entitled to differ by an ulp.
    // They do -- that assertion passed on macOS and failed on Linux.
    let acoth2 = 0.54930614433405484570;
    let got2 = num("=ACOTH(2)");
    assert!(
        ((got2 - acoth2) / acoth2).abs() < 1e-15,
        "ACOTH(2) expected {acoth2}, got {got2}"
    );
}

#[test]
fn test_iseven_and_isodd_past_the_i64_range() {
    // Parity is decided in f64, not through an i64 cast: that cast
    // saturates at i64::MAX (odd) for anything past ~9.2e18, so
    // ISEVEN(19^24) came out FALSE. Excel works from the double it holds --
    // 19^24 is odd mathematically, but its f64 is a multiple of a large
    // power of two, and Excel answers TRUE.
    assert!(matches!(
        eval_one("=ISEVEN(INT(19^24))"),
        ResultData::Boolean(true)
    ));
    assert!(matches!(
        eval_one("=ISODD(INT(19^24))"),
        ResultData::Boolean(false)
    ));
    // Ordinary parity is unchanged, including negatives and truncation.
    for (src, want) in [
        ("=ISEVEN(4)", true),
        ("=ISEVEN(3)", false),
        ("=ISEVEN(-4)", true),
        ("=ISEVEN(0)", true),
        ("=ISEVEN(2.5)", true),
        ("=ISODD(3)", true),
        ("=ISODD(4)", false),
        ("=ISODD(-3)", true),
    ] {
        match eval_one(src) {
            ResultData::Boolean(b) => assert_eq!(b, want, "for {src}"),
            other => panic!("expected a boolean for {src}, got {other:?}"),
        }
    }
}

#[test]
fn test_prob_checks_only_that_the_probabilities_sum_to_one() {
    // Excel does not reject an individual probability outside [0, 1] --
    // only the total matters. PROB({1,2}, {1.5,-0.5}, 0, 3) is 1 in real
    // Excel, and rejecting the negative turned a pairwise-excluded range
    // that legitimately summed to 1 into #NUM!.
    let mut sheet = create_sheet(&[
        ["1", "0.5", "1.5", "=PROB(A1:A2, B1:B2, 0, 3)"],
        ["2", "0.5", "-0.5", "=PROB(A1:A2, C1:C2, 0, 3)"],
        ["", "", "", "=PROB(A1:A2, A1:A2, 0, 3)"],
    ]);
    sheet.commit(None).unwrap();
    assert_eq!(num_of(&sheet.get_result_data(&CellRef::new(0, 3))), 1.0);
    assert_eq!(num_of(&sheet.get_result_data(&CellRef::new(1, 3))), 1.0);
    // A total that is not 1 is still #NUM!.
    match sheet.get_result_data(&CellRef::new(2, 3)) {
        ResultData::Error(e) => assert_eq!(e, "#NUM!"),
        other => panic!("expected #NUM!, got {other:?}"),
    }
}

fn num_of(r: &ResultData) -> f64 {
    match r {
        ResultData::Float(f) => *f,
        ResultData::Integer(i) => *i as f64,
        other => panic!("expected a number, got {other:?}"),
    }
}

#[test]
fn test_serial_zero_is_excels_phantom_january_zero() {
    // Excel's serial 0 is "January 0, 1900", not 1 January, and it is
    // consistent about it: DAY(0) is 0 while MONTH(0) is 1 and YEAR(0) is
    // 1900. Returning day 1 there put every one of these out by a day.
    assert_eq!(num("=DAY(0)"), 0.0);
    assert_eq!(num("=MONTH(0)"), 1.0);
    assert_eq!(num("=YEAR(0)"), 1900.0);
    // A fraction of a day is still day 0.
    assert_eq!(num("=DAY(0.6299)"), 0.0);
    match eval_one("=TEXT(0.6299, \"yyyy-mm-dd\")") {
        ResultData::String(s) => assert_eq!(s, "1900-01-00"),
        other => panic!("expected 1900-01-00, got {other:?}"),
    }
    // Serial 1 onwards is unchanged, including the 1900 leap-year bug.
    assert_eq!(num("=DAY(1)"), 1.0);
    for (serial, want) in [
        (59.0, "1900-02-28"),
        (60.0, "1900-02-29"),
        (61.0, "1900-03-01"),
    ] {
        match eval_one(&format!("=TEXT({serial}, \"yyyy-mm-dd\")")) {
            ResultData::String(s) => assert_eq!(s, want, "for serial {serial}"),
            other => panic!("expected {want}, got {other:?}"),
        }
    }
}

#[test]
fn test_fuzz_log_first_arg_type_error_wins_over_later_arg_error() {
    // Harvested from fuzz/fuzz_excel.py, seed 946837:
    // LOG(Sheet1[[#Headers],[C]], PEARSON(G5:I5, F1:G3)) -- the header
    // reference is non-numeric text, and PEARSON's mismatched-size ranges
    // (1x3 vs 3x2) are #N/A. Real Excel checks LOG's first argument
    // before ever looking at whether the second is itself an error, so
    // the result is #VALUE! (from the first-argument check), not #N/A
    // (measured via win32com: LOG("C", NA()) is #VALUE!). visi previously
    // had a generic pre-dispatch scan that returned the *first* error
    // found across all arguments regardless of position, so it surfaced
    // the #N/A from argument 2 instead of ever reaching LOG's own
    // first-argument check.
    match eval_one("=LOG(\"C\", NA())") {
        ResultData::Error(e) => assert_eq!(e, "#VALUE!"),
        other => panic!("expected #VALUE!, got {other:?}"),
    }
}

#[test]
fn test_fuzz_atan2_first_arg_type_error_wins_over_later_arg_error() {
    // Harvested from fuzz/fuzz_excel.py, seed 196793:
    // ATAN2(IF(H3 > Sheet1[[#Headers],[A]], H5, I3), MODE.SNGL(Sheet1[B]))
    // -- H3 > "A" is FALSE (a number never exceeds text), so the IF
    // yields I3, a non-numeric text cell; MODE.SNGL has no repeated value
    // and is #N/A. Same first-argument-wins shape as LOG above (measured
    // via win32com: ATAN2("text", NA()) is #VALUE!).
    match eval_one("=ATAN2(\"text\", NA())") {
        ResultData::Error(e) => assert_eq!(e, "#VALUE!"),
        other => panic!("expected #VALUE!, got {other:?}"),
    }
}

#[test]
fn test_fuzz_seriessum_rejects_numeric_looking_text_coefficient() {
    // Harvested from fuzz/fuzz_excel.py, seed 107768: coefficients
    // {<blank>, "2" (forced text), 27, -35}. Real Excel's SERIESSUM(1.49,
    // 1, 2, A1:A4) is #VALUE! -- unlike GCD/LCM/MULTINOMIAL, a
    // numeric-looking string in the coefficients isn't coerced (confirmed
    // directly via win32com with no blank at all either:
    // SERIESSUM(1.49, 1, 2, {"2", 27, -35}) is also #VALUE!). visi
    // previously coerced the string and returned a number.
    let grid = [
        ["", "=SERIESSUM(1.49, 1, 2, A1:A4)"],
        ["\"2\"", ""],
        ["27", ""],
        ["-35", ""],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    match sheet.get_result_data(&CellRef::new(0, 1)) {
        ResultData::Error(e) => assert_eq!(e, "#VALUE!"),
        other => panic!("expected #VALUE!, got {other:?}"),
    }
}
