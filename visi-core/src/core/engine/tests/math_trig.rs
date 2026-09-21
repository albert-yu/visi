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
fn test_fuzz_tanh_three_keeps_correctly_rounded_text_length() {
    let grid = [["=LENB(TANH(3))"]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    let result = sheet.get_result_data(&CellRef::new(0, 0));
    assert!(
        matches!(result, ResultData::Float(16.0)),
        "Expected 16 for correctly rounded 0.99505475368673, got {result:?}"
    );
}

#[test]
fn test_fuzz_tanh_three_matches_high_precision_reference() {
    let grid = [["=TANH(3)"]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    let result = sheet.get_result_data(&CellRef::new(0, 0));
    assert!(
        matches!(result, ResultData::Float(f) if f == 0.9950547536867305),
        "Expected nearest f64 to 0.995054753686730451331880185255488475, got {result:?}"
    );
}

#[test]
fn test_fuzz_coth_stays_below_negative_one_until_true_limit() {
    match eval_one("=ISODD(INT(COTH(-19)))") {
        ResultData::Boolean(v) => assert!(!v, "expected FALSE"),
        other => panic!("expected FALSE, got {other:?}"),
    }
}

#[test]
fn test_fuzz_quotient_zero_does_not_keep_negative_sign_in_atan2() {
    match eval_one("=ATAN2(RADIANS(-45), QUOTIENT(PI(), -37))") {
        ResultData::Float(v) => assert!((v - std::f64::consts::PI).abs() < 1e-12, "got {v}"),
        other => panic!("expected pi, got {other:?}"),
    }
}

#[test]
fn test_fuzz_atan2_negative_zero_y_returns_positive_pi() {
    match eval_one("=ATAN2(-84, PERCENTOF(0, -10))") {
        ResultData::Float(v) => assert!((v - std::f64::consts::PI).abs() < 1e-12, "got {v}"),
        other => panic!("expected pi, got {other:?}"),
    }
}

#[test]
fn test_fuzz_power_type_checks_base_before_exponent_error() {
    match eval_one("=POWER(\"C\", NA())") {
        ResultData::Error(e) => assert_eq!(e, "#VALUE!"),
        other => panic!("expected #VALUE!, got {other:?}"),
    }
}

#[test]
fn test_fuzz_power_negative_base_rejects_huge_exponent() {
    match eval_one("=POWER(-95, SINH(-424.13))") {
        ResultData::Error(e) => assert_eq!(e, "#NUM!"),
        other => panic!("expected #NUM!, got {other:?}"),
    }
}

#[test]
fn test_fuzz_sin_cos_tan_and_reciprocals_num_error_past_2_pow_27() {
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
    assert_eq!(num("=COTH(47692.3)"), 1.0);
    assert!((num("=COTH(1)") - 1.3130352854993312).abs() < 1e-9);
}

#[test]
fn test_fuzz_gcd_first_arg_type_error_wins_over_later_arg_error() {
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
    let expected = 3.14;
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
fn test_gcd_lcm_reject_values_and_results_above_excel_integer_limit() {
    assert!(matches!(eval_one("=GCD(2^54, 2)"), ResultData::Error(ref e) if e == "#NUM!"));
    assert!(matches!(eval_one("=LCM(2^53, 3)"), ResultData::Error(ref e) if e == "#NUM!"));
    assert!(
        matches!(eval_one("=LCM(2^53, 1)"), ResultData::Float(v) if (v - 9_007_199_254_740_992.0).abs() < 1.0)
    );
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

    let r1 = sheet.get_result_data(&CellRef::new(2, 0));
    assert!(matches!(r1, ResultData::Float(v) if (v - 11.0).abs() < 1e-6));

    let r2 = sheet.get_result_data(&CellRef::new(2, 1));
    assert!(matches!(r2, ResultData::Float(v) if (v - 30.0).abs() < 1e-6));

    let r3 = sheet.get_result_data(&CellRef::new(2, 2));
    assert!(matches!(r3, ResultData::Float(v) if (v - 1024.0).abs() < 1e-6));

    let r4 = sheet.get_result_data(&CellRef::new(2, 3));
    assert!(matches!(r4, ResultData::Float(v) if (v - 3.0).abs() < 1e-6));
}

#[test]
fn test_sumsq_result_is_snapped_before_rounddown_observes_it() {
    let mut sheet = create_sheet(&[
        ["D", "=ROUNDDOWN(SUMSQ(A:A), 2)"],
        ["-5", ""],
        ["9", ""],
        ["26", ""],
        ["10", ""],
        ["10", ""],
        ["43", ""],
        ["-188.7", ""],
        ["FALSE", ""],
        ["38", ""],
    ]);
    sheet.commit(None).unwrap();
    match sheet.get_result_data(&CellRef::new(0, 1)) {
        ResultData::Float(f) => assert_eq!(f, 39882.69),
        other => panic!("expected 39882.69, got {other:?}"),
    }
}

#[test]
fn test_ceiling_floor_honor_significance_argument() {
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

    let r6 = sheet.get_result_data(&CellRef::new(0, 5));
    assert!(
        matches!(r6, ResultData::String(ref s) if s == "6.12323399573677E-17+i"),
        "IMSQRT(-1) = {r6:?}"
    );

    let r7 = sheet.get_result_data(&CellRef::new(0, 6));
    assert!(
        matches!(r7, ResultData::String(ref s) if s == "4"),
        "IMPOWER(2,2) = {r7:?}"
    );
}

#[test]
fn test_complex_number_formatting_uses_excel_precision_not_raw_f64() {
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
        assert_eq!(
            got_s.chars().last(),
            expected.chars().last(),
            "col {col}: suffix changed ({got_s} vs {expected})"
        );
    }
}

#[test]
fn test_roman_concise_forms_match_excel() {
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
    assert_eq!(num("=DAYS360(40602, 42543, FALSE)"), 1912.0);
    assert_eq!(num("=DAYS360(40602, 42543)"), 1912.0);
    assert_eq!(num("=DAYS360(40602, 42543, TRUE)"), 1914.0);
}

#[test]
fn test_supplied_blank_argument_is_zero_not_the_default() {
    assert_eq!(num("=LOG(100)"), 2.0);
    assert_eq!(num("=MROUND(10, Z90)"), 0.0);
    for src in ["=LOG(100, Z90)", "=LOG(1, Z90)"] {
        match eval_one(src) {
            ResultData::Error(e) => assert_eq!(e, "#NUM!", "for {src}"),
            other => panic!("expected #NUM! for {src}, got {other:?}"),
        }
    }
    match eval_one("=LOG(1, 1)") {
        ResultData::Error(e) => assert_eq!(e, "#DIV/0!"),
        other => panic!("expected #DIV/0!, got {other:?}"),
    }
}

#[test]
fn test_power_of_zero_to_the_zero_is_num_error() {
    for src in ["=POWER(0, 0)", "=0^0", "=POWER(Y90, Y91)"] {
        match eval_one(src) {
            ResultData::Error(e) => assert_eq!(e, "#NUM!", "for {src}"),
            other => panic!("expected #NUM! for {src}, got {other:?}"),
        }
    }
    assert_eq!(num("=POWER(0, 2)"), 0.0);
    assert_eq!(num("=POWER(2, 0)"), 1.0);
    match eval_one("=POWER(0, -1)") {
        ResultData::Error(e) => assert_eq!(e, "#DIV/0!"),
        other => panic!("expected #DIV/0!, got {other:?}"),
    }
}

#[test]
fn test_sumproduct_treats_non_numeric_entries_as_zero() {
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
    let cases: [(&str, &str, &str, f64, f64); 8] = [
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
        (
            "DATE(2018,2,28)",
            "DATE(2018,3,10)",
            "DATE(2018,4,30)",
            0.0,
            100.35615901837147,
        ),
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
    assert_eq!(num("=MOD(3298534883328, 3)"), 0.0);
    match eval_one("=MOD(6597069766656, 3)") {
        ResultData::Error(e) => assert_eq!(e, "#NUM!"),
        other => panic!("expected #NUM!, got {other:?}"),
    }
    assert_eq!(num("=MOD(10, 3)"), 1.0);
    assert_eq!(num("=MOD(-10, 3)"), 2.0);
    assert_eq!(num("=MOD(10, -3)"), -2.0);
    assert_eq!(num("=MOD(TRUE, 2)"), 1.0);
}

#[test]
fn test_fuzz_mod_stays_exact_at_an_integer_quotient_boundary() {
    assert_eq!(num("=MOD(-47, (47 / -13))"), 0.0);
}

#[test]
fn test_fuzz_mod_fractional_percent_exact_zero() {
    assert_eq!(num("=MOD(52, 1%)"), 0.0);
    assert_eq!(num("=MOD(52, 0.01)"), 0.0);
}

#[test]
fn test_fuzz_mod_tiny_power_against_negative_divisor_is_not_zero() {
    assert!((num("=MOD((-5 ^ -16), (-44 * 85))") + 3740.0).abs() < 1e-9);
}

#[test]
#[allow(clippy::excessive_precision)]
fn test_coupdaysnc_and_acoth_precision() {
    let s = "DATE(2011,8,28)";
    let m = "EDATE(DATE(2011,8,28),18)";
    assert_eq!(num(&format!("=COUPDAYSNC({s}, {m}, 2, 0)")), 2.0);
    assert_eq!(num(&format!("=COUPDAYBS({s}, {m}, 2, 0)")), 178.0);
    assert_eq!(num(&format!("=COUPDAYS({s}, {m}, 2, 0)")), 180.0);
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

    let exact = -0.0060606802661724050957;
    let got = num("=ACOTH(-165)");
    assert!(
        ((got - exact) / exact).abs() < 1e-15,
        "ACOTH(-165) expected {exact}, got {got}"
    );
    let acoth2 = 0.54930614433405484570;
    let got2 = num("=ACOTH(2)");
    assert!(
        ((got2 - acoth2) / acoth2).abs() < 1e-15,
        "ACOTH(2) expected {acoth2}, got {got2}"
    );
}

#[test]
fn test_iseven_and_isodd_past_the_i64_range() {
    assert!(matches!(
        eval_one("=ISEVEN(INT(19^24))"),
        ResultData::Boolean(true)
    ));
    assert!(matches!(
        eval_one("=ISODD(INT(19^24))"),
        ResultData::Boolean(false)
    ));
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
    let mut sheet = create_sheet(&[
        ["1", "0.5", "1.5", "=PROB(A1:A2, B1:B2, 0, 3)"],
        ["2", "0.5", "-0.5", "=PROB(A1:A2, C1:C2, 0, 3)"],
        ["", "", "", "=PROB(A1:A2, A1:A2, 0, 3)"],
    ]);
    sheet.commit(None).unwrap();
    assert_eq!(num_of(&sheet.get_result_data(&CellRef::new(0, 3))), 1.0);
    assert_eq!(num_of(&sheet.get_result_data(&CellRef::new(1, 3))), 1.0);
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
    assert_eq!(num("=DAY(0)"), 0.0);
    assert_eq!(num("=MONTH(0)"), 1.0);
    assert_eq!(num("=YEAR(0)"), 1900.0);
    assert_eq!(num("=DAY(0.6299)"), 0.0);
    match eval_one("=TEXT(0.6299, \"yyyy-mm-dd\")") {
        ResultData::String(s) => assert_eq!(s, "1900-01-00"),
        other => panic!("expected 1900-01-00, got {other:?}"),
    }
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
    match eval_one("=LOG(\"C\", NA())") {
        ResultData::Error(e) => assert_eq!(e, "#VALUE!"),
        other => panic!("expected #VALUE!, got {other:?}"),
    }
}

#[test]
fn test_fuzz_atan2_first_arg_type_error_wins_over_later_arg_error() {
    match eval_one("=ATAN2(\"text\", NA())") {
        ResultData::Error(e) => assert_eq!(e, "#VALUE!"),
        other => panic!("expected #VALUE!, got {other:?}"),
    }
}

#[test]
fn test_fuzz_seriessum_rejects_numeric_looking_text_coefficient() {
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
