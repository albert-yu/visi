use super::*;

#[test]
fn test_upper_rounddown_combo() {
    let sheet_src = [
        // A   B   C     D   E
        ["-60", "", "-6", "", "1"],
        ["cJsjkQ", "183.83", "", "TRUE", "FALSE"],
        ["32", "", "-96", "-25", "-205.145"],
        ["177.49", "-423.909", "", "-27", "n"],
        ["-77", "SusOPQc", "86", "", "-107.2312"],
        [
            "=(MIN(A1, A4) / PRODUCT(C1:E4))",
            "=(C2 * SQRT(-49))",
            "=(B3 * IF((E4 > -49), E1, E1))",
            "=LEFT(\"ABS(-17)\", 2)",
            "=IF((-41 > C3), SUM(B1:B5), C3)",
        ],
        [
            "=IF((UPPER(\"A4\") > C3), ROUNDDOWN(B6, 0), IF((-21 > -30), C2, A6))",
            "=ABS(D4)",
            "=LOWER(\"SUM(D2:E6)\")",
            "-112",
            "=ABS(ROUNDDOWN(B4, 2))",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    // A7
    let cell_ref = CellRef::new(6, 0);
    let target = sheet.get_result_data(&cell_ref);
    match target {
        ResultData::Error(e) => assert_eq!(e, "#NUM!"),
        other => panic!("Expected #NUM!, got {:?}", other),
    }
}

#[test]
fn test_fuzz_right_string_addition_coercion() {
    let sheet_src = [
        ["\"fuKg\"", "74", "-75", "-37", "29"],
        ["6", "FALSE", "FALSE", "", "FALSE"],
        ["-327.3", "-27", "87", "-70", ""],
        ["\"Ec\"", "FALSE", "41", "267.6832", ""],
        ["148.845", "", "-435", "-344.315", "28"],
        [
            "=LOWER(\"30\")",
            "=UPPER(\"(D2 * -2)\")",
            "=(MIN(E3, -6) ^ ROUNDUP(24, 0))",
            "=C2",
            "=(LEN(\"D3\") / -34)",
        ],
        [
            "17",
            "=IF((-12 > SUM(E5, A1)), (D1 ^ A6), E5)",
            "=C1",
            "=ABS(D2)",
            "=42",
        ],
        [
            "=(IF((B2 > D4), C6, E6) ^ ROUNDUP(17, 0))",
            "=MAX(18, (D7 * 39))",
            "=ROUND(LEN(\"C7\"), 1)",
            "=(IF((28 > -30), A1, C6) / D2)",
            "=LEN(\"AND(C6 > 0, A7 < 100)\")",
        ],
        [
            "=C7",
            "-255.8071",
            "=B1",
            "-85",
            "=AND(SUM(D1:D8) > 0, ABS(D4) < 100)",
        ],
        [
            "=IF((25 > OR(43 > 0, -33 < 100)), ABS(-13), LOWER(\"E6\"))",
            "=(LEFT(\"8\", 1) * IF((21 > A6), C2, 47))",
            "=C6",
            "=48",
            "=(RIGHT(\"A1\", 2) + ROUNDDOWN(E1, 1))",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 1));
    match target {
        ResultData::Float(f) => assert_eq!(f, 28.0),
        ResultData::Integer(i) => assert_eq!(i, 28),
        other => panic!("Expected Float(28.0) for B7, got {:?}", other),
    }
}

#[test]
fn test_fuzz_lower_negative_number_string_coercion() {
    let sheet_src = [
        ["", "61", "\"ha\"", "97", "166"],
        ["-8", "TRUE", "-374.6", "424", "328.132"],
        ["21", "-100", "-97", "FALSE", "-10"],
        ["83", "TRUE", "-9", "178.26", "-80"],
        ["\"2H\"", "FALSE", "\"qn\"", "FALSE", "4"],
        ["=4", "2", "=MAX(A5:A5)", "=A2", "=UPPER(\"D4\")"],
        [
            "=AND(-15 > 0, INT(B2) < 100)",
            "=B3",
            "=A2",
            "=LOWER(\"B5\")",
            "=A5",
        ],
        [
            "=ABS(ROUNDUP(34, 1))",
            "=(RIGHT(\"A7\", 2) - (C7 - B4))",
            "=B2",
            "=A1",
            "\"kiK ga\"",
        ],
        [
            "=MIN(MAX(-43, E8), -19)",
            "=(ROUNDUP(B4, 2) / IF((C1 > -33), 7, 19))",
            "=AND(D6 > 0, ROUND(D8, 2) < 100)",
            "=SQRT(D8)",
            "47",
        ],
        [
            "=ABS(INT(A3))",
            "=(C9 + -2)",
            "=MIN(E1:E7)",
            "436.831",
            "=LOWER(\"-30\")",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 1));
    match target {
        ResultData::Float(f) => assert_eq!(f, -2.0),
        other => panic!("Expected Float(-2.0) for B10, got {:?}", other),
    }
}

#[test]
fn test_fuzz_and_left_string_comparison() {
    let sheet_src = [
        ["\"2rLzL\"", "19", "1", "78", "-442.67"],
        ["3", "\" QAap\"", "0", "\"q\"", "TRUE"],
        ["195.715", "-239.5", "FALSE", "52", ""],
        ["", "\"L qPj\"", "-24.5812", "97", "FALSE"],
        ["46", "\"yQpUe\"", "-51", "30", "\"jXFiCL\""],
        [
            "=(D5 ^ D3)",
            "=LEFT(\"(A4 * D4)\", 4)",
            "=ROUNDDOWN((A1 * E3), 2)",
            "2",
            "=(CONCATENATE(\"E1\", \"-30\") / C2)",
        ],
        [
            "=B4",
            "=AND(D4 > 0, D6 < 100)",
            "=ABS(ROUNDDOWN(A2, 0))",
            "=(ROUND(E4, 2) - LEN(\"E4\"))",
            "\"1dDknr\"",
        ],
        [
            "=AVERAGE((B7 + D6), ROUNDDOWN(D4, 0))",
            "=(ROUNDDOWN(D1, 2) * 6)",
            "=PRODUCT(PRODUCT(A1:A5), E2)",
            "=A5",
            "=43",
        ],
        ["=SUM(B1:E6)", "=C2", "FALSE", "=A3", "=SUM(B7:D7)"],
        [
            "=SUM(A7:A8)",
            "=LEFT(\"(C5 * -4)\", 3)",
            "=-9",
            "=IF(((E8 / E6) > -41), C2, IF((D3 > -31), B6, -11))",
            "=AND(LEFT(\"-37\", 3) > 0, MAX(D3:E8) < 100)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 0));
    match target {
        ResultData::Float(f) => assert_eq!(f, 50.0),
        other => panic!("Expected Float(50.0) for A8, got {:?}", other),
    }
}

#[test]
fn test_fuzz_concatenate_if_function_error() {
    let sheet_src = [
        ["TRUE", "FALSE", "TRUE", "329.1", "TRUE"],
        ["\"MwNWS\"", "57", "-46", "-70", "\"ViyztP\""],
        ["58.05", "-84", "-90", "4", "-439.4668"],
        ["-61", "39", "-56", "-41", "305.132"],
        ["-386.937", "\"KXkaE\"", "TRUE", "\"2C\"", ""],
        [
            "=RIGHT(\"E5\", 1)",
            "=(SQRT(C5) - IF((8 > D1), D5, C2))",
            "=47",
            "=OR(B5 > 0, MAX(B5:C5) < 100)",
            "=RIGHT(\"ROUND(-35, 2)\", 1)",
        ],
        ["=A6", "=(AVERAGE(D5:D5) ^ SQRT(A2))", "=D4", "=E5", "=A5"],
        ["=C2", "=-38", "=C5", "TRUE", "=A2"],
        ["", "=PRODUCT(C8, 8)", "=(C6 / MAX(B6:C7))", "=-17", "=D3"],
        [
            "=D6",
            "=ROUNDUP(ROUNDDOWN(D2, 1), 0)",
            "=SUM(E6:E7)",
            "=ROUNDDOWN(-48, 0)",
            "\"POJV\"",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 2));
    assert!(
        matches!(target, ResultData::Error(_)),
        "Expected Error for C9, got {:?}",
        target
    );
}

#[test]
fn test_fuzz_string_coercion_expected_number_error() {
    let sheet_src = [
        ["-292.46", "\"UCmL\"", "\"csEXNMm\"", "\"hpNtY\"", "74"],
        ["FALSE", "-57", "99", "34", "0"],
        ["-45", "25", "445.79", "-19", "\"QYFhRm\""],
        ["48", "410.85", "3", "186.346", "9"],
        ["-89", "24", "58.732", "3", "-74"],
        [
            "10",
            "=E2",
            "=((A3 / 20) + SQRT(B2))",
            "44",
            "=ABS(AND(E3 > 0, B3 < 100))",
        ],
        [
            "=LEFT(\"A2\", 5)",
            "=IF((PRODUCT(E4:E4) > (A4 + C3)), E4, 35)",
            "=INT(-30)",
            "=(CONCATENATE(\"B6\", \"B3\") * SUM(46, A2))",
            "=OR(D3 > 0, MIN(D6:E6) < 100)",
        ],
        [
            "=37",
            "-79",
            "=PRODUCT(A3, E1)",
            "235.4",
            "=AND((2 - -9) > 0, RIGHT(\"43\", 3) < 100)",
        ],
        [
            "=D4",
            "=PRODUCT(E6:E7)",
            "=(MAX(A3:B3) - LOWER(\"-30\"))",
            "=E5",
            "=LEN(\"IF((-50 > A3), -23, E4)\")",
        ],
        [
            "0",
            "\"qq\"",
            "=PRODUCT(A2:C5)",
            "=MAX(AND(C1 > 0, B7 < 100), CONCATENATE(\"D1\", \"12\"))",
            "=ABS(IF((A3 > 46), 7, 41))",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 3));
    assert!(
        matches!(target, ResultData::Error(_)),
        "Expected Error for D10, got {:?}",
        target
    );
}

#[test]
fn test_fuzz_average_right_string_argument() {
    let sheet_src = [
        ["28", "68", "\"E \"", "", "TRUE"],
        ["", "-99", "23", "", "6"],
        ["", "\"bb\"", "TRUE", "-271.85", "-78"],
        ["FALSE", "\"gEqS\"", "-43", "29", "\"A\""],
        ["88", "\"R\"", "-66", "81", "-8"],
        [
            "=(INT(B3) / (-11 ^ D1))",
            "=D5",
            "-80",
            "=AND((7 - D3) > 0, AVERAGE(C3:C5) < 100)",
            "=MAX(-34, PRODUCT(A5:E5))",
        ],
        [
            "=INT(A5)",
            "=A1",
            "=29",
            "=RIGHT(\"IF((D5 > E1), D5, E2)\", 1)",
            "=AVERAGE(A1, RIGHT(\"D6\", 1))",
        ],
        [
            "=16",
            "=PRODUCT(C7, (D4 + C7))",
            "=AVERAGE(C5:D5)",
            "=OR(-23 > 0, -21 < 100)",
            "-56",
        ],
        [
            "=PRODUCT(AND(11 > 0, A1 < 100), (A6 ^ A3))",
            "=C6",
            "=-19",
            "=B5",
            "=A4",
        ],
        ["=E3", "=E4", "=LEFT(\"A1\", 4)", "=B6", "=C1"],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 4));
    match target {
        ResultData::Float(f) => assert_eq!(f, 17.0),
        other => panic!("Expected Float(17.0) for E7, got {:?}", other),
    }
}

#[test]
fn test_fuzz_if_max_right_string_branch() {
    let sheet_src = [
        ["\"YDAf2XmI\"", "FALSE", "", "12.4026", "\"id\""],
        ["\"Gui\"", "-41", "346.342", "\"f\"", "-386"],
        ["12", "-36", "160.585", "-359.5374", "45"],
        ["322.06", "38", "-30", "FALSE", "\"V\""],
        ["", "2", "41", "67", "47.3"],
        [
            "=D5",
            "=INT(LEFT(\"-12\", 1))",
            "=ROUNDDOWN(MIN(E3:E4), 2)",
            "=AND(E5 > 0, 33 < 100)",
            "=D5",
        ],
        [
            "=D3",
            "=E1",
            "=((E6 / 41) + ROUND(8, 2))",
            "=AVERAGE(IF((C1 > C6), B5, B4), ROUND(C6, 2))",
            "=-9",
        ],
        [
            "=D3",
            "=IF((MAX(C7:E7) > -17), RIGHT(\"C1\", 3), AND(16 > 0, C4 < 100))",
            "=SQRT(A2)",
            "=-7",
            "=MAX(UPPER(\"C3\"), MAX(D4:E4))",
        ],
        [
            "=D7",
            "=IF((IF((C6 > B2), D5, D4) > C3), OR(B2 > 0, A8 < 100), (C8 - D5))",
            "=C8",
            "=AND(INT(-7) > 0, OR(E2 > 0, B7 < 100) < 100)",
            "=AVERAGE(ROUNDDOWN(E3, 2), C2)",
        ],
        [
            "-43",
            "=((-46 * C3) / IF((-25 > A3), C7, C3))",
            "=UPPER(\"E3\")",
            "=ROUNDUP(SUM(C3:D5), 0)",
            "=IF((MAX(A2:D9) > B9), ROUNDUP(D5, 2), E4)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target1 = sheet.get_result_data(&CellRef::new(7, 1));
    let target2 = sheet.get_result_data(&CellRef::new(6, 2));
    match target1 {
        ResultData::String(s) => assert_eq!(s, "C1"),
        other => panic!("Expected String(\"C1\"), got {:?}", other),
    }
    match target2 {
        ResultData::Float(f) => assert!((f - 9.6341).abs() < 1e-3),
        other => panic!("Expected Float(~9.6341), got {:?}", other),
    }
}

#[test]
fn test_fuzz_and_upper_power_comparison() {
    let sheet_src = [
        ["\"iPf\"", "FALSE", "327.95", "-9", "-343.8"],
        ["-81", "6", "-59", "\"ixpH\"", "0"],
        ["51", "\"yNu\"", "\"X\"", "-31", "41"],
        ["", "-466.47", "437", "0", ""],
        ["7", "50", "-84", "-90", "-95"],
        [
            "200.2",
            "=OR(INT(D3) > 0, (6 * A3) < 100)",
            "=IF((ABS(E4) > OR(D1 > 0, -31 < 100)), E5, ROUNDDOWN(E5, 0))",
            "=B1",
            "=A5",
        ],
        [
            "=AND(MIN(E3:E6) > 0, UPPER(\"E2\") < 100)",
            "",
            "=LEFT(\"PRODUCT(B2:E6)\", 2)",
            "=D2",
            "=-35",
        ],
        [
            "=C5",
            "=ABS((39 + A7))",
            "=(24 ^ -44)",
            "=IF((C6 > AVERAGE(E2:E5)), 7, D2)",
            "-244.157",
        ],
        [
            "=10",
            "-103.1",
            "=ROUND(E2, 2)",
            "=AVERAGE(PRODUCT(D8, B3), C5)",
            "=ROUND(PRODUCT(20, D4), 2)",
        ],
        [
            "=D8",
            "=SUM(OR(D9 > 0, A1 < 100), ROUND(-9, 0))",
            "=(INT(D7) + CONCATENATE(\"D9\", \"C9\"))",
            "=AND(UPPER(\"E9\") > 0, (E7 ^ A7) < 100)",
            "=E9",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 3));
    match target {
        ResultData::Boolean(b) => assert_eq!(b, true),
        other => panic!("Expected Boolean(true), got {:?}", other),
    }
}

#[test]
fn test_fuzz_subtraction_string_formula_cell() {
    let sheet_src = [
        ["-499.734", "-35", "0", "30", "5"],
        ["-65", "", "\"mARAP\"", "7", ""],
        ["448.29", "-115.1434", "\"TDzexvu\"", "6", ""],
        ["\"uddSCurn\"", "", "98", "49", "\"YQ1G\""],
        ["0", "-5", "\"1lrX\"", "19", "\"C\""],
        [
            "=((E5 / A2) ^ IF((A3 > C5), -48, E4))",
            "=ROUND(AND(-5 > 0, -40 < 100), 2)",
            "=D2",
            "22",
            "=AND(UPPER(\"D4\") > 0, B5 < 100)",
        ],
        [
            "=E5",
            "=PRODUCT(D2:E3)",
            "=(MAX(B5:D6) - SUM(C1:E1))",
            "=D5",
            "=ROUND(B6, 0)",
        ],
        [
            "=SQRT(LOWER(\"A6\"))",
            "=AND((46 + A6) > 0, 28 < 100)",
            "=(MAX(E3:E7) / AND(A6 > 0, E2 < 100))",
            "=SUM(C4:D4)",
            "=A4",
        ],
        ["=C8", "=ABS(B8)", "", "=AND(-18 > 0, INT(8) < 100)", "=-2"],
        [
            "=(B5 - E9)",
            "=(INT(-10) / RIGHT(\"C6\", 3))",
            "=AVERAGE(A6:A6)",
            "=C3",
            "=RIGHT(\"C8\", 2)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    match target {
        ResultData::Float(f) => assert_eq!(f, -3.0),
        other => panic!("Expected Float(-3.0) for A10, got {:?}", other),
    }
    for r in 0..10 {
        for c in 0..5 {
            let res = sheet.get_result_data(&CellRef::new(r, c));
            if matches!(res, ResultData::Error(_)) {
                let col_let = (b'A' + c as u8) as char;
                assert!(col_let >= 'A' && col_let <= 'E');
            }
        }
    }
}

#[test]
fn test_fuzz_multiplication_upper_string_number() {
    let sheet_src = [
        ["-68", "\"KB\"", "\"XE\"", "352.42", "451.693"],
        ["-39.3853", "", "", "29.6", "34.75"],
        ["4", "94", "-6", "-99", ""],
        ["\"W\"", "\"L\"", "11", "94", ""],
        ["49", "", "363.5", "\"zdIPRq\"", "159.09"],
        [
            "=AND(PRODUCT(B1:C2) > 0, 20 < 100)",
            "=D4",
            "14",
            "=C5",
            "-164.7238",
        ],
        [
            "=A4",
            "=ABS(MIN(A2, 24))",
            "=B1",
            "=B2",
            "=SQRT(SUM(-10, C6))",
        ],
        [
            "=(SUM(D3:D3) + ROUNDUP(B7, 0))",
            "=D1",
            "18",
            "=C3",
            "=SUM(A6:C6)",
        ],
        ["=42", "=PRODUCT(7, C6)", "=LEN(\"E2\")", "=8", "=-6"],
        [
            "=-43",
            "=(C9 * UPPER(\"24\"))",
            "=IF((A5 > ROUND(42, 0)), (5 + 43), OR(-2 > 0, E5 < 100))",
            "=12",
            "=D9",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 1));
    match target {
        ResultData::Float(f) => assert_eq!(f, 48.0),
        other => panic!("Expected Float(48.0), got {:?}", other),
    }
}

#[test]
fn test_fuzz_division_by_right_string() {
    let sheet_src = [
        ["-23", "-90.2", "43", "50", "FALSE"],
        ["-24", "74", "-31", "85", "53"],
        ["\"lzFQ1BZ\"", "-98", "44", "6", "-21"],
        ["89", "355.279", "\"Vv\"", "255", "\" hU\""],
        ["79.51000000000001", "-16", "-152.82", "8", "135.7"],
        [
            "-54",
            "=SUM(D3:D5)",
            "=INT(C3)",
            "=(LOWER(\"B4\") - LEN(\"A1\"))",
            "=(PRODUCT(7, D3) / AVERAGE(E3:E4))",
        ],
        ["=B1", "33", "=(LEN(\"-33\") / D4)", "=B4", "TRUE"],
        ["145", "=(D7 / RIGHT(\"B6\", 1))", "59", "223.585", "-86"],
        [
            "=SQRT(RIGHT(\"D8\", 3))",
            "=-17",
            "=ROUND(B5, 1)",
            "=(E6 * -3)",
            "=IF((E4 > (B6 + -22)), (B1 - -29), -47)",
        ],
        [
            "=AVERAGE(CONCATENATE(\"B5\", \"A2\"), 3)",
            "=AVERAGE((E3 * -25), ROUNDDOWN(C1, 2))",
            "=ROUNDDOWN(B1, 0)",
            "=ROUNDDOWN(IF((-45 > E2), C4, 5), 2)",
            "=OR(B8 > 0, MAX(A2:B8) < 100)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 1));
    match target {
        ResultData::Float(f) => assert!(
            (f - 59.213166).abs() < 1e-3,
            "Expected ~59.213166 for B8, got {}",
            f
        ),
        other => panic!("Expected Float(~59.213166) for B8, got {:?}", other),
    }
}

#[test]
fn test_fuzz_left_operand_string_value_error_precedence() {
    let sheet_src = [
        ["-391.0356", "\"1Jc\"", "FALSE", "", "\"rCSFXgC\""],
        ["27", "179", "-314.7", "64", "FALSE"],
        ["-51", "-427.6", "-261.33", "\"e\"", "\"2OtmmtAT\""],
        ["0", "\"CEpp\"", "21", "75.179", "-279.178"],
        ["", "-205", "-227.2985", "414.49", "FALSE"],
        [
            "=14",
            "-85",
            "=(B5 / IF((-14 > D1), E1, E4))",
            "=(D3 * (49 / D1))",
            "=PRODUCT(A1:A4)",
        ],
        [
            "=CONCATENATE(\"E5\", \"SUM(C2:C6)\")",
            "=UPPER(\"ABS(A5)\")",
            "=AND(D2 > 0, (-40 - E3) < 100)",
            "=AVERAGE(B6:E6)",
            "=(D1 + 6)",
        ],
        [
            "=IF(((C4 + A1) > A4), -5, B6)",
            "=LEFT(\"SQRT(A2)\", 5)",
            "-27",
            "=ROUNDUP(B1, 2)",
            "=B5",
        ],
        [
            "=(-15 ^ AND(E3 > 0, A7 < 100))",
            "=SQRT(ROUND(C5, 1))",
            "=E4",
            "=B5",
            "=A7",
        ],
        [
            "=(ROUNDDOWN(C2, 2) * ROUNDUP(E9, 0))",
            "374",
            "=IF((-35 > AND(D9 > 0, C4 < 100)), -40, C5)",
            "=A4",
            "=INT((36 - C7))",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target_d6 = sheet.get_result_data(&CellRef::new(5, 3));
    println!(
        "Seed 316841 evaluated target CellRef(5, 3) D6: {:?}",
        target_d6
    );
    match target_d6 {
        ResultData::Error(ref e) => assert!(
            e.contains("#DIV/0!") || e.contains("#VALUE!"),
            "Expected #DIV/0! or #VALUE! for D6, got {:?}",
            target_d6
        ),
        other => panic!("Expected Error for D6, got {:?}", other),
    }

    let target_d7 = sheet.get_result_data(&CellRef::new(6, 3));
    println!(
        "Seed 316841 evaluated target CellRef(6, 3) D7: {:?}",
        target_d7
    );
    match target_d7 {
        ResultData::Error(ref e) => assert!(
            e.contains("#DIV/0!") || e.contains("#VALUE!"),
            "Expected #DIV/0! or #VALUE! for D7, got {:?}",
            target_d7
        ),
        other => panic!("Expected Error for D7, got {:?}", other),
    }
}

#[test]
fn test_fuzz_sum_cell_string_literal_parsing() {
    let sheet_src = [
        ["", "-215.8", "", "\"dhvueks\"", "32"],
        ["-11", "43", "-228.367", "57", "-37.8"],
        ["28", "61", "-129.451", "-35.61", "FALSE"],
        ["-79.73", "8", "-91", "31.168", "TRUE"],
        ["\"r\"", "", "4", "-3.6295", "\"2\""],
        [
            "=ROUNDDOWN(SUM(E5, A3), 2)",
            "=AVERAGE(C5:E5)",
            "=PRODUCT((C2 - 39), AND(-27 > 0, B1 < 100))",
            "=(AVERAGE(E3:E5) ^ UPPER(\"-28\"))",
            "=(C4 + ROUNDDOWN(-11, 2))",
        ],
        [
            "=13",
            "=A4",
            "=SUM(B4, LEFT(\"A1\", 4))",
            "87",
            "=IF((D6 > D3), 31, IF((D5 > B6), D3, 13))",
        ],
        [
            "=D7",
            "=UPPER(\"43\")",
            "=C7",
            "79",
            "=INT(AND(D3 > 0, E7 < 100))",
        ],
        [
            "=B6",
            "=IF((-45 > SUM(40, B2)), UPPER(\"E6\"), 42)",
            "95",
            "=C7",
            "17",
        ],
        ["-429.6347", "=D7", "=C9", "=20", "=(B1 ^ E4)"],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let e5 = sheet.get_result_data(&CellRef::new(4, 4));
    println!("Seed 623549 E5: {:?}", e5);
    let a6 = sheet.get_result_data(&CellRef::new(5, 0));
    println!("Seed 623549 A6: {:?}", a6);
    let b6 = sheet.get_result_data(&CellRef::new(5, 1));
    println!("Seed 623549 B6: {:?}", b6);
    let d6 = sheet.get_result_data(&CellRef::new(5, 3));
    println!("Seed 623549 D6: {:?}", d6);
    match a6 {
        ResultData::Float(f) => assert_eq!(f, 28.0),
        other => panic!("Expected Float(28.0), got {:?}", other),
    }
    match b6 {
        ResultData::Float(f) => assert!((f - 0.18525).abs() < 1e-3),
        other => panic!("Expected Float(~0.18525), got {:?}", other),
    }
    match d6 {
        ResultData::Error(ref e) => assert!(e.contains("#DIV/0!")),
        other => panic!("Expected Error(#DIV/0!), got {:?}", other),
    }
}

#[test]
fn test_fuzz_min_cell_string_ignore() {
    let sheet_src = [
        ["TRUE", "FALSE", "\"ajTqx2gu\"", "175", "\" yil2xDW\""],
        ["\"XKxZ\"", "TRUE", "-330", "7", "\"dQ\""],
        ["0", "\"tMcVh\"", "", "-75", ""],
        ["89", "\"Lz\"", "", "-87", "-50"],
        ["\"NWXitc\"", "-98", "FALSE", "\"HLFOVkq\"", "\"fc\""],
        ["=E5", "=AVERAGE(B2:B2)", "=B5", "=MAX(B3:E3)", "=0"],
        [
            "=-5",
            "134.5",
            "=UPPER(\"D2\")",
            "=PRODUCT(LEFT(\"D5\", 3), -20)",
            "=LEN(\"(E3 / C1)\")",
        ],
        [
            "=OR(OR(A2 > 0, 4 < 100) > 0, ABS(-47) < 100)",
            "=IF((IF((A1 > A7), 12, C4) > (A5 - D4)), IF((C2 > 13), B4, D1), A1)",
            "=MIN(ABS(E6), B4)",
            "=-32",
            "=-42",
        ],
        [
            "=C7",
            "=IF((ROUNDUP(C2, 0) > (-16 * A1)), C5, -25)",
            "-19",
            "=9",
            "=A6",
        ],
        [
            "=48",
            "=INT(MIN(D6, A9))",
            "=-30",
            "-130.326",
            "=(LOWER(\"D9\") - -24)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 2));
    println!("Seed 923425 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 0.0),
        other => panic!("Expected Float(0.0), got {:?}", other),
    }
}

#[test]
fn test_fuzz_sqrt_len_formula_string() {
    let sheet_src = [
        ["FALSE", "257.4", "54", "90", "-84"],
        ["9", "-24", "", "0", "\"ggbBvUQb\""],
        ["44", "0", "", "3", "FALSE"],
        ["", "TRUE", "-179.503", "99", "\"DwfJTT2\""],
        ["-63", "386.5289", "FALSE", "", "14"],
        [
            "=-8",
            "=IF(((C4 ^ E5) > B2), ABS(-38), ROUNDDOWN(C5, 1))",
            "=LEN(\"-30\")",
            "=C4",
            "=PRODUCT(B4, RIGHT(\"A4\", 3))",
        ],
        [
            "=ROUNDDOWN(AVERAGE(D4:D4), 1)",
            "=E1",
            "=8",
            "=D4",
            "=(AVERAGE(E3:E6) / (47 + E1))",
        ],
        [
            "=E3",
            "31",
            "=UPPER(\"C2\")",
            "=LEN(\"MIN(A4, 12)\")",
            "=PRODUCT((A3 ^ 45), -46)",
        ],
        ["=-39", "=SQRT(D8)", "=44", "=E8", "=34"],
        [
            "=B4",
            "=UPPER(\"C8\")",
            "=OR(D1 > 0, ROUNDUP(C8, 1) < 100)",
            "=SUM(ROUND(-23, 0), E5)",
            "=IF(((B9 * A6) > ABS(D2)), C7, -38)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let b9 = sheet.get_result_data(&CellRef::new(8, 1));
    println!("Seed 433038 B9: {:?}", b9);
    let e10 = sheet.get_result_data(&CellRef::new(9, 4));
    println!("Seed 433038 E10: {:?}", e10);
    match b9 {
        ResultData::Float(f) => assert!((f - 3.3166247903554).abs() < 1e-5),
        other => panic!("Expected Float(~3.3166), got {:?}", other),
    }
    match e10 {
        ResultData::Float(f) => assert_eq!(f, -38.0),
        other => panic!("Expected Float(-38.0), got {:?}", other),
    }
}

#[test]
fn test_fuzz_concatenate_scientific_string() {
    let sheet_src = [
        ["211.623", "-204.42", "91", "113.74", "-39"],
        ["TRUE", "\"uU\"", "\"3eEROE\"", "", "64"],
        ["", "99", "FALSE", "", "166.3114"],
        ["TRUE", "-28", "-3", "\"1QlA\"", ""],
        ["89", "-235.82", "-471.049", "-411.8394", "-15"],
        [
            "=C4",
            "=UPPER(\"(E5 / E3)\")",
            "=ROUNDDOWN(ROUNDDOWN(E4, 1), 1)",
            "=-7",
            "=C4",
        ],
        [
            "=A1",
            "=-30",
            "=AND(OR(E6 > 0, A6 < 100) > 0, AND(D6 > 0, C6 < 100) < 100)",
            "=IF(((A1 ^ B3) > CONCATENATE(\"B6\", \"D1\")), SUM(8, 3), CONCATENATE(\"45\", \"E3\"))",
            "=LEN(\"B4\")",
        ],
        [
            "=D1",
            "=IF((D5 > LOWER(\"A1\")), (25 * -9), CONCATENATE(\"12\", \"B5\"))",
            "=AVERAGE(B3:B3)",
            "\"KHtwf\"",
            "=ROUND(INT(A2), 1)",
        ],
        [
            "=IF((LEN(\"C7\") > ABS(A7)), SQRT(A3), (E6 / D6))",
            "=D3",
            "=PRODUCT(C6:E7)",
            "=(A7 * OR(C4 > 0, 16 < 100))",
            "=ROUNDUP(OR(-29 > 0, D1 < 100), 1)",
        ],
        [
            "=OR(B3 > 0, LOWER(\"E5\") < 100)",
            "=35",
            "=C4",
            "=ABS(ROUND(33, 0))",
            "82",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 3));
    println!("Seed 278797 target: {:?}", target);
    match target {
        ResultData::String(ref s) => assert_eq!(s, "45E3"),
        other => panic!("Expected String(\"45E3\"), got {:?}", other),
    }
}

#[test]
fn test_fuzz_abs_d6_sum_string_cond() {
    let sheet_src = [
        ["100", "\"ggwX\"", "", "84", "\"zs tV\""],
        ["\"UgssE\"", "FALSE", "-39", "", "-78"],
        ["", "12", "\"QxX3nKBd\"", "-95", "-322.18"],
        ["84", "-84", "-17", "", "55"],
        ["94", "-398.86", "69", "", "-52"],
        [
            "=ROUNDDOWN(C2, 0)",
            "=25",
            "87",
            "=IF((SUM(-23, E1) > SUM(C2, B2)), SQRT(D4), B5)",
            "-87",
        ],
        [
            "=D5",
            "\"S3Mfwld\"",
            "=ABS(D6)",
            "=AND(UPPER(\"A4\") > 0, OR(E1 > 0, A6 < 100) < 100)",
            "=IF((OR(45 > 0, C6 < 100) > IF((12 > B4), B2, E2)), RIGHT(\"C1\", 5), AND(30 > 0, D6 < 100))",
        ],
        [
            "=OR(-4 > 0, ROUND(E7, 0) < 100)",
            "=MAX(D4:D4)",
            "=AVERAGE(-14, (18 - E1))",
            "\"I\"",
            "-42",
        ],
        [
            "=CONCATENATE(\"D6\", \"(B1 * D5)\")",
            "=(OR(E4 > 0, E8 < 100) - OR(A3 > 0, -35 < 100))",
            "=C1",
            "=LOWER(\"B3\")",
            "=IF((MIN(B7, D2) > D4), IF((B1 > -46), B1, D4), -40)",
        ],
        [
            "=(C8 * (44 + A8))",
            "=UPPER(\"OR(B9 > 0, A3 < 100)\")",
            "=C9",
            "=ROUND(E4, 2)",
            "=C3",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 2));
    println!("Seed 262160 target: {:?}", target);
}

#[test]
fn test_fuzz_left_string_max_range_if() {
    let sheet_src = [
        ["0", "\"yiCm\"", "72", "\"pa\"", "-245.1091"],
        ["", "-65", "\"3P\"", "4", "431.6"],
        ["", "-96.63", "\"Ux\"", "", "49"],
        ["-14", "TRUE", "20", "-91", "\"l1L\""],
        ["128", "", "72", "-387.3", "-73.70999999999999"],
        [
            "=AND(D2 > 0, A3 < 100)",
            "=ROUNDDOWN(MIN(E1:E1), 2)",
            "=ROUND(CONCATENATE(\"E2\", \"-22\"), 1)",
            "=B2",
            "=0",
        ],
        [
            "=B1",
            "=(LEN(\"C4\") ^ (D2 - 49))",
            "=OR(ROUNDDOWN(D4, 1) > 0, A3 < 100)",
            "=B3",
            "=7",
        ],
        [
            "=SUM(ROUNDDOWN(B2, 1), D7)",
            "=MAX(-39, (-8 ^ C4))",
            "=OR((D1 ^ D7) > 0, PRODUCT(C3, D3) < 100)",
            "=C7",
            "=37",
        ],
        [
            "=IF((LEFT(\"11\", 5) > MAX(D7:D8)), ROUND(D8, 1), 42)",
            "=-40",
            "\"VQqCDql\"",
            "=46",
            "=D2",
        ],
        [
            "=MAX(E7:E9)",
            "=CONCATENATE(\"INT(A7)\", \"E1\")",
            "=LEFT(\"(E3 * D3)\", 4)",
            "-14",
            "-75",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 0));
    println!("Seed 64619 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 1.0),
        ResultData::Integer(i) => assert_eq!(i, 1),
        other => panic!("Expected 1, got {:?}", other),
    }
}

#[test]
fn test_fuzz_sqrt_right_string_multiplication() {
    let sheet_src = [
        ["353.54", "-86", "-353.5", "-30", "8"],
        ["0", "-25", "-75", "-466.059", "\" jSiRFLD\""],
        ["-204", "40", "217.5", "7", ""],
        ["31", "77", "60", "-60", "10"],
        ["5", "-227.38", "42", "", "408.45"],
        ["=LEFT(\"A5\", 4)", "=B4", "FALSE", "=C3", "=ROUND(B1, 0)"],
        [
            "89.6992",
            "=OR((D1 ^ C4) > 0, MIN(B6:B6) < 100)",
            "=MAX(C5:C6)",
            "=42",
            "=INT(A6)",
        ],
        [
            "=-22",
            "=IF((32 > ROUNDUP(E1, 1)), E3, OR(A4 > 0, C7 < 100))",
            "=SQRT(D3)",
            "-331.98",
            "=C7",
        ],
        [
            "=(SQRT(C8) * RIGHT(\"-26\", 4))",
            "=D4",
            "=ABS(B5)",
            "=C6",
            "5",
        ],
        [
            "=SQRT(C4)",
            "=ROUND(ABS(E9), 0)",
            "=SUM(E6:E8)",
            "=ROUNDUP(A7, 2)",
            "=PRODUCT(ROUNDDOWN(C1, 1), SUM(A6:A6))",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 0));
    println!("Seed 419803 target: {:?}", target);
}

#[test]
fn test_fuzz_sqrt_concatenate_date_string() {
    let sheet_src = [
        ["5", "FALSE", "478.85", "83", "-92"],
        ["-86.03", "-220.4", "26", "292.73", "168"],
        ["483.8", "TRUE", "402", "\"ESFAq\"", "0"],
        ["", "-107.8535", "-37.85", "-100", "54"],
        ["-151.8833", "\"1vokLCy1\"", "-79", "", "26"],
        [
            "",
            "=RIGHT(\"(C4 - C2)\", 1)",
            "=ROUNDUP(UPPER(\"A4\"), 2)",
            "=-14",
            "=MIN(C5:C5)",
        ],
        [
            "=E1",
            "=49",
            "=RIGHT(\"MAX(E5:E5)\", 3)",
            "=E5",
            "=MIN(E3:E5)",
        ],
        [
            "=D4",
            "=-3",
            "=IF((D6 > OR(A2 > 0, D2 < 100)), 21, IF((C4 > 24), E6, C4))",
            "=SQRT(CONCATENATE(\"7\", \"-23\"))",
            "=D2",
        ],
        [
            "=LEFT(\"PRODUCT(A1:C5)\", 3)",
            "=AND(ROUND(D7, 0) > 0, (D5 - C2) < 100)",
            "33.9",
            "=MIN(B7:C8)",
            "=C7",
        ],
        [
            "-269.6165",
            "=MAX(50, 30)",
            "=C8",
            "=AVERAGE(A8:E9)",
            "=PRODUCT(PRODUCT(C8, D2), E9)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 3));
    println!("Seed 225586 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert!((f - 215.00232556881798).abs() < 1e-5),
        other => panic!("Expected float 215.00232556881798, got {:?}", other),
    }
}

#[test]
fn test_fuzz_average_lower_e9_value_error() {
    let sheet_src = [
        ["-14", "-2", "FALSE", "", "\"kdj\""],
        ["44", "-32", "", "\"UDocR 31\"", "\"jr\""],
        ["7", "-38", "TRUE", "\"QDTdWDIa\"", ""],
        ["", "-65", "-89", "-355", "15"],
        ["\"nkp\"", "178.6532", "FALSE", "\"21oGJn\"", "\"3\""],
        [
            "=IF((LOWER(\"D5\") > SQRT(C2)), E4, D1)",
            "=ROUND(OR(C2 > 0, C1 < 100), 1)",
            "=UPPER(\"IF((C1 > E3), D4, 1)\")",
            "-34.1584",
            "-26",
        ],
        [
            "=ROUND(12, 0)",
            "=LOWER(\"-29\")",
            "=IF((IF((D6 > A1), E2, C5) > -31), D4, IF((21 > -27), -46, C5))",
            "=OR(A5 > 0, 8 < 100)",
            "=D1",
        ],
        [
            "=ROUNDUP(LOWER(\"-21\"), 0)",
            "=SUM(A6:B6)",
            "=LEN(\"B1\")",
            "=AVERAGE(E3:E3)",
            "=-4",
        ],
        [
            "=E7",
            "-30",
            "=IF((-44 > C1), ROUND(D1, 1), D5)",
            "=A6",
            "=SQRT(A2)",
        ],
        [
            "=((C5 ^ E8) / (D5 * 10))",
            "=AVERAGE(LOWER(\"E9\"), D8)",
            "=LEN(\"(A7 ^ E1)\")",
            "=IF((D8 > C7), SUM(C9:E9), (E4 * D7))",
            "=AVERAGE(A7:C8)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 1));
    println!("Seed 758897 target: {:?}", target);
    println!("Target is {:?}", target);
}

#[test]
fn test_fuzz_if_left_hyphen_string_comparison() {
    let sheet_src = [
        ["99", "TRUE", "321.01", "423", "-13"],
        ["100", "480", "\"y\"", "4", "-98"],
        ["349.764", "-437.2078", "\"uEL\"", "FALSE", "-31"],
        ["-94", "-81", "", "165", "-85"],
        ["", "FALSE", "\"rz\"", "97", "50.4176"],
        [
            "=UPPER(\"(6 ^ -17)\")",
            "0",
            "=ROUND(B5, 0)",
            "28",
            "=(ROUNDUP(-9, 2) * A5)",
        ],
        [
            "=OR(B6 > 0, MIN(B2:B5) < 100)",
            "=(MAX(E5:E5) * D1)",
            "=ABS((6 - C1))",
            "=AND(MAX(D3, -19) > 0, ROUND(17, 0) < 100)",
            "=IF((LEFT(\"-44\", 1) > A6), ROUNDDOWN(D5, 1), CONCATENATE(\"A5\", \"B1\"))",
        ],
        ["=B3", "=A2", "-275.0375", "=E5", "=RIGHT(\"ABS(C1)\", 3)"],
        [
            "=MAX(D5, (E1 ^ A6))",
            "=AND(C2 > 0, A7 < 100)",
            "=MAX(B1:D7)",
            "=CONCATENATE(\"(A3 - B3)\", \"SUM(B7, D3)\")",
            "=AVERAGE(A1:D2)",
        ],
        [
            "=LEN(\"C1\")",
            "=((2 - -4) / (-18 / E4))",
            "=C9",
            "=8",
            "=B7",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 4));
    println!("Seed 153407 target: {:?}", target);
    match target {
        ResultData::String(s) => assert_eq!(s, "A5B1"),
        other => panic!("Expected String A5B1, got {:?}", other),
    }
}

#[test]
fn test_fuzz_if_boolean_gt_int_upper_neg7() {
    let sheet_src = [
        ["-20", "-25.878", "", "-55.5739", "94"],
        ["-137", "", "146.1", "9", "268"],
        ["10", "\"lidXy\"", "\"sDa\"", "-40.7562", "382"],
        ["-77", "-42", "\"BFyBvSvA\"", "FALSE", "\"Qx3pF \""],
        ["341.131", "\"SI eihZF\"", "-39", "68", "356"],
        [
            "TRUE",
            "=ROUNDDOWN(AND(C3 > 0, -22 < 100), 0)",
            "=(LOWER(\"7\") + (48 * E5))",
            "=(IF((B3 > C5), D3, -26) + 11)",
            "=LEN(\"A1\")",
        ],
        [
            "=C6",
            "=IF((AND(E4 > 0, A5 < 100) > INT(C6)), UPPER(\"-7\"), A5)",
            "=C6",
            "=C3",
            "-14",
        ],
        [
            "=SQRT(SUM(A1, A1))",
            "=ABS(MAX(9, C4))",
            "=E4",
            "=IF((ROUNDUP(D3, 1) > (A7 + E1)), D6, -4)",
            "-76",
        ],
        [
            "=IF((ROUNDDOWN(C4, 1) > IF((B8 > B4), D7, E2)), A3, D2)",
            "=D5",
            "=IF((AND(C7 > 0, -18 < 100) > OR(B6 > 0, -30 < 100)), (B8 ^ A4), (A7 - A6))",
            "=17",
            "=D4",
        ],
        [
            "=E7",
            "=((E7 ^ E6) - 46)",
            "=(LEFT(\"E1\", 4) ^ (16 * E4))",
            "=AND(MAX(A2:A3) > 0, LEFT(\"E3\", 1) < 100)",
            "=A1",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 1));
    println!("Seed 547773 target: {:?}", target);
    match target {
        ResultData::String(s) => assert_eq!(s, "-7"),
        ResultData::Float(f) => assert_eq!(f, -7.0),
        ResultData::Integer(i) => assert_eq!(i, -7),
        other => panic!("Expected -7, got {:?}", other),
    }
}

#[test]
fn test_fuzz_if_sqrt_boolean_string_branch() {
    let sheet_src = [
        ["\"pCFGZmpm\"", "-434.8", "TRUE", "-413.79", "169.5"],
        ["2", "-130", "\"CIR\"", "133", "\"SwcK\""],
        ["30", "0", "FALSE", "-13", "69"],
        ["-34", "-99.045", "11", "17", "46"],
        ["-98", "", "\"1UOZV3 f\"", "-90", "-78"],
        [
            "=B5",
            "-83",
            "=PRODUCT((-30 - E5), -11)",
            "=LEFT(\"B4\", 5)",
            "=PRODUCT(A2:C3)",
        ],
        [
            "\"FtNOl\"",
            "=OR(PRODUCT(B1:C2) > 0, (39 / C1) < 100)",
            "93",
            "=-34",
            "=INT(C1)",
        ],
        ["=B7", "=E2", "=-5", "154.153", "=-31"],
        [
            "=17",
            "=MAX(C2:D3)",
            "=OR(17 > 0, PRODUCT(B6:C6) < 100)",
            "\"ccKYKd\"",
            "\"apQN\"",
        ],
        [
            "80",
            "=AND(A5 > 0, A4 < 100)",
            "=E8",
            "=IF((ABS(B2) > SQRT(A8)), D6, AVERAGE(B9:D9))",
            "=3",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 3));
    println!("Seed 815090 target: {:?}", target);
    match target {
        ResultData::String(s) => assert_eq!(s, "B4"),
        other => panic!("Expected B4, got {:?}", other),
    }
}

#[test]
fn test_fuzz_if_number_gt_string_comparison() {
    let sheet_src = [
        ["79", "-61", "10", "\"llTMht\"", ""],
        ["-72", "\"mhZ\"", "-10", "131.678", "-19"],
        ["\"svgONaM\"", "", "-164.077", "-310.3", "-4"],
        ["-444.004", "-58", "-50", "4", "41"],
        ["FALSE", "\"FSyZd\"", "23", "11", "\"I qQ1iAl\""],
        ["=4", "=17", "=-21", "=PRODUCT(LEN(\"9\"), B4)", "=ABS(D3)"],
        [
            "=(-43 / (3 ^ 23))",
            "=A5",
            "=E5",
            "=LEFT(\"ROUND(-29, 1)\", 1)",
            "=AVERAGE(B1:C5)",
        ],
        [
            "=AVERAGE(E6:E7)",
            "=IF((ABS(E7) > LEFT(\"D2\", 2)), UPPER(\"E2\"), A3)",
            "TRUE",
            "",
            "=ROUNDDOWN(LEFT(\"A7\", 3), 1)",
        ],
        [
            "=IF((CONCATENATE(\"C4\", \"E2\") > CONCATENATE(\"18\", \"32\")), AVERAGE(B1, E5), ROUNDUP(A7, 2))",
            "=IF((B2 > UPPER(\"E6\")), CONCATENATE(\"A7\", \"E8\"), (B5 / A6))",
            "=-40",
            "-320",
            "=(IF((E3 > 28), 42, -28) - 4)",
        ],
        [
            "=ROUNDDOWN(AVERAGE(C1, A7), 1)",
            "=A1",
            "-22",
            "=RIGHT(\"A6\", 1)",
            "=UPPER(\"C4\")",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 1));
    println!("Seed 315137 target: {:?}", target);
    match target {
        ResultData::String(s) => assert!(s.contains("svgONaM")),
        other => panic!("Expected svgONaM, got {:?}", other),
    }
}

#[test]
fn test_fuzz_max_concatenate_or_value_error() {
    let sheet_src = [
        ["", "-14", "53", "-187", "82"],
        ["7", "224.428", "251", "\"OirPKSm\"", "-412.847"],
        ["-96", "\"sE\"", "386.3835", "-4", "\"oRQ2oX\""],
        ["38", "-6", "-39", "447", "\"K\""],
        ["-82", "-15", "59", "41", "4"],
        [
            "=E4",
            "=LOWER(\"INT(A1)\")",
            "=IF((A5 > AVERAGE(-32, B2)), ABS(B5), AND(10 > 0, E3 < 100))",
            "=D1",
            "=A2",
        ],
        [
            "=IF((-43 > LOWER(\"A6\")), 50, SQRT(-45))",
            "=A2",
            "39",
            "=A5",
            "=B3",
        ],
        [
            "=ROUND(IF((D4 > E3), C2, C3), 1)",
            "0",
            "12.4211",
            "=ABS(C5)",
            "=C7",
        ],
        [
            "=E2",
            "-367.7296",
            "=SQRT(C8)",
            "=(SQRT(D1) / E5)",
            "=AVERAGE((D4 - A8), C7)",
        ],
        [
            "=LEN(\"A5\")",
            "=MIN(A1:E5)",
            "=MAX(CONCATENATE(\"E6\", \"B7\"), OR(A6 > 0, D9 < 100))",
            "178.7795",
            "=A3",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 2));
    println!("Seed 509664 target: {:?}", target);
    match target {
        ResultData::Error(e) => assert!(e == "#VALUE!" || e == "#NUM!"),
        other => panic!("Expected #VALUE! or #NUM!, got {:?}", other),
    }
}

#[test]
fn test_fuzz_sum_range_string_literal_ignore() {
    let sheet_src = [
        ["3", "7", "2", "-75", "\"2AY\""],
        ["-20", "62", "85", "\"Hu\"", "TRUE"],
        ["", "80", "94", "-20", ""],
        ["FALSE", "2", "\"EgCWnL2J\"", "TRUE", "-27"],
        ["\"3\"", "", "", "TRUE", "\"sCr\""],
        [
            "39",
            "=OR(PRODUCT(C1, B4) > 0, SQRT(E3) < 100)",
            "=SUM(A5:A5)",
            "71",
            "=A4",
        ],
        ["=A2", "=26", "\"mCzRh\"", "", "=UPPER(\"E5\")"],
        [
            "=ABS((E1 ^ -46))",
            "=RIGHT(\"(E4 * D1)\", 3)",
            "=AVERAGE(LEN(\"C3\"), D1)",
            "=AND(ROUNDUP(E6, 1) > 0, PRODUCT(D7:E7) < 100)",
            "=(C4 + CONCATENATE(\"39\", \"7\"))",
        ],
        [
            "=IF((12 > C1), AND(C7 > 0, E1 < 100), (D5 / A6))",
            "=ROUNDUP(A2, 1)",
            "=SUM(A3:B5)",
            "TRUE",
            "=C6",
        ],
        [
            "=D3",
            "=(B1 ^ A3)",
            "=IF((INT(11) > CONCATENATE(\"D5\", \"-26\")), SUM(E7, A6), AND(E2 > 0, C1 < 100))",
            "=B2",
            "482.2",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(5, 2));
    println!("Seed 965939 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 0.0),
        ResultData::Integer(i) => assert_eq!(i, 0),
        other => panic!("Expected 0, got {:?}", other),
    }
}

#[test]
fn test_fuzz_if_upper_string_comparison_val() {
    let sheet_src = [
        ["65", "0", "FALSE", "-97", "55"],
        ["5", "FALSE", "\"JZOdb1\"", "", "FALSE"],
        ["38", "", "-357", "\"hWxw\"", "242.9694"],
        ["-98", "\"RV\"", "-238.682", "-87", "\"G\""],
        ["26.52", "-14", "495.2726", "-12", "\"1QpSKs\""],
        [
            "=(B1 ^ 21)",
            "=ABS(AVERAGE(A4, B5))",
            "=30",
            "=IF((ROUNDUP(D5, 1) > UPPER(\"B4\")), PRODUCT(E4:E4), 35)",
            "=LOWER(\"IF((E1 > 32), C1, E1)\")",
        ],
        [
            "=E1",
            "=IF((PRODUCT(B5:B6) > E4), CONCATENATE(\"D5\", \"C3\"), C2)",
            "",
            "=E6",
            "=C3",
        ],
        [
            "=IF((OR(15 > 0, -15 < 100) > IF((-34 > 35), E7, C1)), D2, AND(A3 > 0, B7 < 100))",
            "-496.535",
            "-22",
            "=C4",
            "=10",
        ],
        [
            "=CONCATENATE(\"D2\", \"12\")",
            "=LEN(\"A3\")",
            "=D6",
            "=IF((-24 > UPPER(\"A4\")), E1, ROUNDDOWN(E7, 2))",
            "=C8",
        ],
        [
            "=A5",
            "=PRODUCT(D4:E4)",
            "=ROUNDDOWN(SQRT(C8), 1)",
            "=IF((-20 > (B8 + E7)), (-11 + A5), SUM(D1, A2))",
            "=LOWER(\"D3\")",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 3));
    println!("Seed 242788 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, -357.0),
        ResultData::Integer(i) => assert_eq!(i, -357),
        other => panic!("Expected -357, got {:?}", other),
    }
}

#[test]
fn test_fuzz_if_sqrt_len_branch_val() {
    let sheet_src = [
        ["-330.75", "-39", "TRUE", "\"Jkspzj\"", "TRUE"],
        ["FALSE", "-53", "\"Mvh\"", "-323.98", "\"UeAkj\""],
        ["FALSE", "57", "1", "FALSE", "414.87"],
        ["-51", "TRUE", "62.04", "52", "TRUE"],
        ["6", "FALSE", "83", "486.6446", "FALSE"],
        [
            "-83",
            "-46",
            "=MIN(AND(D3 > 0, A2 < 100), IF((A4 > -20), 2, D5))",
            "=ROUNDDOWN(E5, 1)",
            "=D4",
        ],
        [
            "-345",
            "=C5",
            "=AND(MIN(D3, D6) > 0, E5 < 100)",
            "=E4",
            "=(LOWER(\"C3\") / E3)",
        ],
        [
            "=AVERAGE(B5, 45)",
            "=LEFT(\"A7\", 3)",
            "=-19",
            "=PRODUCT(A3, A5)",
            "=LEN(\"IF((B1 > E3), -42, E2)\")",
        ],
        [
            "=((B2 + B7) ^ LOWER(\"D2\"))",
            "",
            "=IF((B7 > SQRT(D6)), LEN(\"D7\"), C3)",
            "=ROUNDUP(MIN(D8:D8), 0)",
            "-35",
        ],
        ["2", "=D8", "=C4", "=AVERAGE(A3:E5)", ""],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 2));
    println!("Seed 738552 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 2.0),
        ResultData::Integer(i) => assert_eq!(i, 2),
        other => panic!("Expected 2, got {:?}", other),
    }
}

#[test]
fn test_fuzz_or_multiplication_string_error_order() {
    let sheet_src = [
        ["281.41", "67.274", "-21", "46", "0"],
        ["-424.437", "\"3qUWHp\"", "FALSE", "\"1C3\"", "12.5"],
        ["FALSE", "2", "TRUE", "59", "2"],
        ["-199.85", "", "TRUE", "85", "64"],
        ["\"kuG3fTA\"", "0", "TRUE", "FALSE", "9"],
        [
            "=D5",
            "=-25",
            "TRUE",
            "=OR(-45 > 0, E5 < 100)",
            "=OR((0 * D2) > 0, -20 < 100)",
        ],
        [
            "=E4",
            "=(SUM(D6, B6) - ROUND(D2, 1))",
            "=ROUNDDOWN(A6, 1)",
            "=(D1 * AVERAGE(B2:B6))",
            "=ROUNDUP(AND(E5 > 0, A6 < 100), 0)",
        ],
        [
            "=B3",
            "=(ROUNDDOWN(6, 0) ^ AND(-27 > 0, -38 < 100))",
            "=B5",
            "",
            "=A3",
        ],
        [
            "=C8",
            "=C1",
            "-489.158",
            "=C7",
            "=IF((MAX(C4:D5) > E4), E2, LOWER(\"B2\"))",
        ],
        ["=ROUNDUP((B7 + C4), 0)", "=MIN(D8:D8)", "17", "63", "21"],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(5, 4));
    println!("Seed 574846 target: {:?}", target);
    match target {
        ResultData::Error(e) => assert_eq!(e, "#VALUE!"),
        other => panic!("Expected #VALUE!, got {:?}", other),
    }
}

#[test]
fn test_fuzz_multiplication_string_div_by_zero_precedence() {
    let sheet_src = [
        ["37", "-148.27", "261.06", "68", "179"],
        ["", "\"QUR\"", "TRUE", "-73", "\"vyeKi\""],
        ["", "FALSE", "44", "\"LuoHWK\"", "59.38"],
        ["61", "\"kpGlOd\"", "\"Xtz1\"", "-46.4", "-388.1504"],
        ["17", "42", "-80", "3", "-64"],
        ["=A1", "=INT(-22)", "=-37", "=C2", "=E1"],
        [
            "",
            "-279",
            "=PRODUCT((B2 - C3), A4)",
            "=(D3 * (E3 / A3))",
            "=((D4 - E2) - B3)",
        ],
        ["=ABS(E3)", "=MIN(A5, ABS(D2))", "=29", "=-13", "=D2"],
        [
            "=LEN(\"B3\")",
            "=LEFT(\"AND(22 > 0, 14 < 100)\", 1)",
            "=ROUND(AND(E5 > 0, E6 < 100), 0)",
            "=C6",
            "=INT(LEFT(\"-41\", 2))",
        ],
        [
            "=PRODUCT(LEN(\"D3\"), ABS(-48))",
            "=(E7 * E5)",
            "=INT((D2 + A8))",
            "=-3",
            "=IF((LEN(\"C5\") > E6), 34, LOWER(\"E8\"))",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 3));
    println!("Seed 224583 target: {:?}", target);
    match target {
        ResultData::Error(e) => assert_eq!(e, "#VALUE!"),
        other => panic!("Expected #VALUE!, got {:?}", other),
    }
}

#[test]
fn test_fuzz_or_round_string_error_order() {
    let sheet_src = [
        ["50", "4", "55", "-417", "-58"],
        ["55", "346.94", "\"Y ggk\"", "76", "-422.647"],
        ["7", "\"TtF\"", "46", "269.1", "0"],
        ["\"gG\"", "-41", "361.685", "-198.7", "3"],
        ["", "TRUE", "-71.2086", "221.5164", "FALSE"],
        [
            "=(LEFT(\"B2\", 5) * C3)",
            "=(-22 - ROUND(25, 0))",
            "=B2",
            "=LEFT(\"ROUNDDOWN(13, 1)\", 1)",
            "=SQRT(A5)",
        ],
        ["\"zMyIO\"", "=AVERAGE(A6:D6)", "=3", "=27", "=A4"],
        [
            "=A2",
            "=30",
            "=IF((LOWER(\"A4\") > OR(E6 > 0, -35 < 100)), INT(30), LEN(\"E5\"))",
            "-7",
            "46",
        ],
        [
            "FALSE",
            "=C1",
            "=B2",
            "=(SQRT(19) - E7)",
            "=OR(OR(B8 > 0, D1 < 100) > 0, ROUND(A4, 1) < 100)",
        ],
        [
            "=(-15 - 20)",
            "=RIGHT(\"ROUND(-48, 2)\", 4)",
            "=21",
            "=A3",
            "=(D3 ^ D3)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 4));
    println!("Seed 422835 target: {:?}", target);
    match target {
        ResultData::Error(e) => assert_eq!(e, "#VALUE!"),
        other => panic!("Expected #VALUE!, got {:?}", other),
    }
}

#[test]
fn test_fuzz_sqrt_product_len_val() {
    let sheet_src = [
        ["-92", "\"aoqjZito\"", "\"SaJyt\"", "-419.51", "0"],
        ["-16", "2", "21", "75", "10"],
        ["4", "79", "134.7696", "TRUE", "46"],
        ["72", "24", "", "", "9"],
        ["0", "\"s\"", "57.51", "TRUE", "37"],
        [
            "=22",
            "=D4",
            "=-40",
            "=PRODUCT(LEN(\"C4\"), 30)",
            "=PRODUCT(D5:D5)",
        ],
        [
            "=(E1 ^ D2)",
            "=ROUNDUP(-32, 0)",
            "=E6",
            "=MAX(24, AVERAGE(B1:B1))",
            "=D2",
        ],
        [
            "=SQRT(D6)",
            "=CONCATENATE(\"ROUNDDOWN(C2, 1)\", \"IF((B3 > -13), C7, D1)\")",
            "=((C2 ^ 15) - 29)",
            "=D1",
            "=-8",
        ],
        [
            "=D5",
            "=IF((C7 > AND(50 > 0, D5 < 100)), E8, MAX(E6:E6))",
            "=SQRT(ROUND(A7, 0))",
            "=-14",
            "=IF((A8 > (D2 - B7)), 40, LEFT(\"D1\", 5))",
        ],
        [
            "=(MAX(B5, 25) / IF((D9 > B8), -34, E6))",
            "=(B9 + AND(C2 > 0, E2 < 100))",
            "=D3",
            "",
            "=AVERAGE(B6:E7)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 0));
    println!("Seed 620565 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert!((f - 7.745966692414834).abs() < 1e-4),
        other => panic!("Expected 7.745966692414834, got {:?}", other),
    }
}

#[test]
fn test_fuzz_sqrt_min_cell_ref_string_ignore() {
    let sheet_src = [
        ["8", "7", "-42", "-306.1794", "2"],
        ["-82", "\"FaCgqE\"", "-189.2", "\"xM3\"", "-106"],
        ["80", "\"F\"", "\"n1q\"", "-78", ""],
        ["-137", "-397", "-55", "7", "299.43"],
        ["\"ER\"", "-62", "92", "30.4145", "-492"],
        [
            "=LOWER(\"(A5 / B3)\")",
            "-11",
            "=((D5 / D3) + SUM(D5:E5))",
            "=ROUNDDOWN(-12, 1)",
            "=MIN((E1 ^ E5), B2)",
        ],
        ["=LEN(\"SUM(C4:C6)\")", "=ABS(E3)", "-393.635", "=38", "=E4"],
        [
            "=E5",
            "FALSE",
            "=LEFT(\"D7\", 1)",
            "=E7",
            "=((A7 - C1) + ROUNDUP(B6, 0))",
        ],
        [
            "=B5",
            "=ROUNDDOWN(ABS(E3), 0)",
            "=AVERAGE(C7, PRODUCT(D3:D5))",
            "=OR(E5 > 0, SUM(E7:E7) < 100)",
            "=(C2 ^ PRODUCT(A6:E7))",
        ],
        [
            "=B2",
            "=D2",
            "=PRODUCT(D2, (C7 * B2))",
            "=(SQRT(E6) + 31)",
            "\"m1\"",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 3));
    println!("Seed 874653 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 31.0),
        ResultData::Integer(i) => assert_eq!(i, 31),
        other => panic!("Expected 31, got {:?}", other),
    }
}

#[test]
fn test_fuzz_if_lower_string_constant_branch() {
    let sheet_src = [
        ["-129.6", "\"gHcl\"", "", "73.1585", "67"],
        ["33", "102.9", "\"MfFR\"", "FALSE", "-411.1429"],
        ["-405.31", "8", "36.8849", "198.4695", "FALSE"],
        ["59", "-18", "\"vXut1lNE\"", "-60", ""],
        ["16", "-450.9472", "108.318", "-30", "\"GvWzHSI\""],
        ["=A3", "=B1", "-149.5", "=C1", "=B2"],
        [
            "=IF((SUM(E6, A1) > ROUND(D6, 0)), LOWER(\"14\"), (-8 - -10))",
            "=(IF((-33 > C6), C4, C6) + D1)",
            "=AND(E1 > 0, E1 < 100)",
            "=12",
            "=SUM(A6:D6)",
        ],
        [
            "=RIGHT(\"A3\", 2)",
            "=ROUND(RIGHT(\"-45\", 4), 0)",
            "-29",
            "FALSE",
            "=ABS(ROUND(-17, 0))",
        ],
        ["=C7", "=B7", "=(B6 + -37)", "=(B1 + 15)", "=AVERAGE(C5:E7)"],
        [
            "=-18",
            "=(B2 / UPPER(\"B6\"))",
            "=LEFT(\"AND(34 > 0, B9 < 100)\", 3)",
            "=(1 - D3)",
            "=-20",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 0));
    println!("Seed 419816 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 2.0),
        ResultData::Integer(i) => assert_eq!(i, 2),
        other => panic!("Expected 2, got {:?}", other),
    }
}

#[test]
fn test_fuzz_product_concatenate_left_to_right_value_error() {
    let sheet_src = [
        ["70", "375", "TRUE", "TRUE", "-12"],
        ["60", "\"3QGdh\"", "360.197", "-362", "59"],
        ["-46", "TRUE", "FALSE", "-24", "\"H\""],
        ["-25", "82", "FALSE", "FALSE", ""],
        ["-76", "-73", "\"oCByq\"", "89", "TRUE"],
        [
            "=LOWER(\"ROUND(A4, 1)\")",
            "61",
            "=(SQRT(-1) * 47)",
            "=LEN(\"-30\")",
            "=-36",
        ],
        [
            "=IF((LOWER(\"B1\") > IF((-22 > C3), 3, E2)), 16, A1)",
            "=ABS(A3)",
            "=C5",
            "=ROUNDDOWN(AND(B3 > 0, A3 < 100), 1)",
            "=MIN(B6:E6)",
        ],
        ["=C6", "=C2", "=AVERAGE(-49, ROUNDUP(A6, 2))", "65", "=B7"],
        [
            "=-19",
            "=B1",
            "=PRODUCT(CONCATENATE(\"D4\", \"E5\"), IF((A8 > B8), A5, A3))",
            "=E6",
            "=B3",
        ],
        [
            "=ROUNDDOWN(SQRT(A4), 1)",
            "=10",
            "=IF((E3 > D2), PRODUCT(E5:E9), E2)",
            "=INT(E6)",
            "-175",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 2));
    println!("Seed 158215 target: {:?}", target);
    match target {
        ResultData::Error(e) => assert!(e == "#VALUE!" || e == "#NUM!"),
        other => panic!("Expected #VALUE! or #NUM!, got {:?}", other),
    }
}

#[test]
fn test_fuzz_string_gt_number_comparison_if() {
    let sheet_src = [
        ["\"guyMN s \"", "-89", "12", "86", "-45"],
        ["\"r\"", "93", "FALSE", "", "-28"],
        ["\"mFf\"", "", "372", "-70", "\"aDcy\""],
        ["-202.887", "\"le\"", "45", "-48", "281"],
        ["FALSE", "\"nXAPy\"", "\"bTJiL\"", "-53", "\" QZQ\""],
        [
            "=IF((CONCATENATE(\"C1\", \"B1\") > E5), (13 * D3), -6)",
            "=PRODUCT(A2:B5)",
            "=CONCATENATE(\"A5\", \"E4\")",
            "=(2 / B1)",
            "=IF((E3 > C5), MAX(C2:C3), E4)",
        ],
        [
            "=PRODUCT(E1:E2)",
            "=AVERAGE(A1:D2)",
            "=PRODUCT(E6:E6)",
            "=-25",
            "=(UPPER(\"E4\") * AND(5 > 0, -39 < 100))",
        ],
        [
            "=A7",
            "=OR(OR(B7 > 0, E5 < 100) > 0, (E4 ^ A5) < 100)",
            "=50",
            "=IF((UPPER(\"-22\") > ROUND(E6, 0)), AND(B4 > 0, 24 < 100), RIGHT(\"D6\", 3))",
            "-14.8777",
        ],
        ["=A7", "=(A5 * 43)", "=C6", "-12", "=C4"],
        [
            "=CONCATENATE(\"44\", \"20\")",
            "=OR(SUM(B8:C8) > 0, (11 + A7) < 100)",
            "=ROUNDUP(UPPER(\"B5\"), 1)",
            "=AVERAGE(A2:D6)",
            "=(-3 / E5)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 3));
    println!("Seed 691619 target: {:?}", target);
    match target {
        ResultData::Boolean(b) => assert!(b),
        other => panic!("Expected True, got {:?}", other),
    }
}

#[test]
fn test_fuzz_concatenate_scientific_string_0e5() {
    let sheet_src = [
        ["60", "\"S\"", "", "-488.273", "\"qdY\""],
        ["TRUE", "\"kRZGH\"", "20", "", "6"],
        ["-2", "68", "-3", "-87", "20"],
        ["TRUE", "-96", "\"AbeWNuTT\"", "18", "-50"],
        ["77", "445.53", "-81", "66", "66"],
        [
            "=AVERAGE(C5:E5)",
            "=IF((C2 > E2), B2, (B5 + E4))",
            "=C1",
            "=(IF((D5 > E1), C3, A5) * RIGHT(\"-31\", 5))",
            "=PRODUCT(A1:A5)",
        ],
        [
            "=INT(MAX(A2, B6))",
            "=-23",
            "=D4",
            "=SQRT(ABS(-33))",
            "=-49",
        ],
        [
            "=IF((-19 > 19), ABS(D7), SUM(B3, 25))",
            "=12",
            "=C2",
            "=(E7 * C4)",
            "=SQRT((B1 * E7))",
        ],
        [
            "=47",
            "=-23",
            "=C4",
            "=AVERAGE(E2:E7)",
            "=CONCATENATE(\"0\", \"E5\")",
        ],
        [
            "=CONCATENATE(\"A6\", \"AND(D3 > 0, E9 < 100)\")",
            "=IF((IF((B9 > C4), D3, -12) > A3), AVERAGE(C8:E8), ABS(44))",
            "=ROUND(-13, 1)",
            "=-21",
            "=ABS(E5)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 4));
    println!("Seed 305840 target: {:?}", target);
    match target {
        ResultData::String(s) => assert_eq!(s, "0E5"),
        other => panic!("Expected String 0E5, got {:?}", other),
    }
}

#[test]
fn test_fuzz_if_concatenate_scientific_string() {
    let sheet_src = [
        ["0", "389.7", "", "-58", "47"],
        ["21", "-71", "", "7", "-87"],
        ["-378", "-471", "6", "9", ""],
        ["\"thQpdFj\"", "9", "\"ZZFx\"", "-1", "\"En\""],
        ["-41", "-73", "TRUE", "", "0"],
        [
            "=RIGHT(\"ROUNDUP(E5, 0)\", 4)",
            "=SUM(E3:E3)",
            "=B2",
            "=(D2 ^ UPPER(\"B4\"))",
            "=B3",
        ],
        [
            "\"Z\"",
            "=C2",
            "=IF((IF((38 > C5), 2, E6) > ROUND(B5, 2)), IF((E6 > B4), 1, C1), CONCATENATE(\"-44\", \"E6\"))",
            "=SUM(E6:E6)",
            "\"TJEs\"",
        ],
        [
            "=IF((MIN(A6:D7) > E5), -41, D5)",
            "",
            "=(D3 * B5)",
            "=ROUND(-11, 0)",
            "=-27",
        ],
        [
            "=OR(AVERAGE(C8:E8) > 0, CONCATENATE(\"C7\", \"C2\") < 100)",
            "=ABS(LEFT(\"-32\", 1))",
            "=50",
            "=C6",
            "-64",
        ],
        [
            "=AVERAGE(A3, OR(-47 > 0, E9 < 100))",
            "=D2",
            "=-19",
            "=AND(C8 > 0, B9 < 100)",
            "=(E7 - PRODUCT(B2, C6))",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 2));
    println!("Seed 22551 target: {:?}", target);
    match target {
        ResultData::String(s) => assert_eq!(s, "-44E6"),
        other => panic!("Expected String -44E6, got {:?}", other),
    }
}

#[test]
fn test_fuzz_lower_empty_cell_addition_value_error() {
    let sheet_src = [
        ["", "", "", "", ""],
        ["", "=LOWER(E1) + 42", "", "", ""],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(1, 1));
    match target {
        ResultData::Error(e) => assert_eq!(e, "#VALUE!"),
        other => panic!("Expected #VALUE!, got {:?}", other),
    }
}

#[test]
fn test_fuzz_round_non_numeric_string_value_error() {
    let sheet_src = [
        ["\"TRUE\"", "", "", "", ""],
        ["", "=ROUND(A1, 2)", "", "", ""],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(1, 1));
    match target {
        ResultData::Error(e) => assert_eq!(e, "#VALUE!"),
        other => panic!("Expected #VALUE!, got {:?}", other),
    }
}


