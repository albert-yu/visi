use super::*;

#[test]
fn test_fuzz_cell_reference_zero_coercion() {
    let sheet_src = [
        ["74", "-287.148", "", "\"HrCRG\"", "167.3"],
        ["-66.04900000000001", "-208.0246", "56", "TRUE", ""],
        ["\"RzGODC2\"", "-26", "335.983", "-11", "273.331"],
        ["0", "\"HdRvUPf\"", "-55", "486.727", "209.55"],
        ["\"yHLLKWgU\"", "-34", "FALSE", "\"mPUQspYq\"", "-102.928"],
        [
            "=A3",
            "=CONCATENATE(\"MIN(B2:E4)\", \"A5\")",
            "=-40",
            "=C4",
            "=AVERAGE(B4, B2)",
        ],
        ["-440.6", "=-48", "=(SQRT(40) - ROUND(E5, 1))", "", "=B3"],
        [
            "=OR(IF((B7 > C1), C1, D3) > 0, IF((22 > B4), A5, C2) < 100)",
            "=ROUNDUP(E5, 2)",
            "=(RIGHT(\"A5\", 2) + CONCATENATE(\"C1\", \"C7\"))",
            "=RIGHT(\"19\", 2)",
            "=ABS(ROUNDDOWN(18, 2))",
        ],
        [
            "=(AND(B3 > 0, D5 < 100) * (D3 + C6))",
            "=49",
            "9",
            "=(OR(C3 > 0, A4 < 100) + 29)",
            "=-11",
        ],
        [
            "=D5",
            "=LEN(\"C4\")",
            "=IF((IF((-37 > -30), E3, -39) > SUM(A7:A8)), B6, C7)",
            "=D4",
            "=A5",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 0));
    match target {
        ResultData::Float(f) => assert!(f.abs() < 1e-6),
        other => panic!("Expected Float(0.0) for A9, got {:?}", other),
    }
}

#[test]
fn test_fuzz_constant_literal_cell_evaluation() {
    let sheet_src = [
        ["\"QZaPVt\"", "0", "", "\"wjDauh\"", "275.684"],
        ["-400.3", "-99", "-56", "35", "-33"],
        ["\"rkTuS\"", "-491.12", "65", "118.7921", "20"],
        ["6", "95", "\" uQbVFQI\"", "7", "-54"],
        ["211.4", "-17", "FALSE", "421.516", "-73"],
        [
            "=ABS(E4)",
            "=B1",
            "=(ROUNDUP(A1, 0) / UPPER(\"B4\"))",
            "=C4",
            "=MIN(D5:E5)",
        ],
        [
            "=CONCATENATE(\"MAX(E3:E6)\", \"OR(A4 > 0, E1 < 100)\")",
            "=IF((-29 > B2), E2, ROUND(B6, 2))",
            "TRUE",
            "19",
            "=ROUNDUP(C5, 2)",
        ],
        ["=(E3 ^ B6)", "162.149", "=A5", "=B1", "=B7"],
        [
            "=B3",
            "=(LOWER(\"37\") ^ SQRT(B3))",
            "=LEFT(\"E2\", 5)",
            "=(LOWER(\"D6\") * B4)",
            "=(E8 - 35)",
        ],
        [
            "=B8",
            "=(RIGHT(\"C6\", 1) + SQRT(-21))",
            "=SQRT(PRODUCT(A2:C4))",
            "=D9",
            "=28",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 0));
    match target {
        ResultData::Float(f) => assert_eq!(f, 1.0),
        other => panic!("Expected Float(1.0) for A8, got {:?}", other),
    }
}

#[test]
fn test_fuzz_nested_math_expression_precision() {
    let sheet_src = [
        ["7", "0", "TRUE", "\"rSMs\"", "FALSE"],
        ["\"anel1aK\"", "64", "-61", "-52", ""],
        ["FALSE", "", "19.1863", "4", "37"],
        ["FALSE", "-80", "2", "0", ""],
        ["\"FyvenPp\"", "\"maI\"", "FALSE", "1", "\"N1gFi\""],
        [
            "FALSE",
            "=B2",
            "=-24",
            "=MAX(B2:C3)",
            "=AND(A3 > 0, -50 < 100)",
        ],
        [
            "=16",
            "=(ROUNDUP(6, 0) + LOWER(\"D6\"))",
            "=C5",
            "=AND(ROUNDUP(E6, 2) > 0, OR(E3 > 0, -27 < 100) < 100)",
            "=(B1 + C4)",
        ],
        [
            "=UPPER(\"IF((C3 > E7), -9, B3)\")",
            "=20",
            "",
            "=UPPER(\"(D4 ^ C4)\")",
            "=B4",
        ],
        [
            "=E8",
            "60",
            "=ROUND((E2 / C2), 2)",
            "=IF((43 > C6), AND(E3 > 0, E8 < 100), E1)",
            "-64",
        ],
        ["=D1", "=((B1 / E7) + -43)", "=28", "=B7", "83"],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 1));
    match target {
        ResultData::Float(f) => assert_eq!(f, -43.0),
        other => panic!("Expected Float(-43.0) for B10, got {:?}", other),
    }
}

#[test]
fn test_fuzz_negative_constant_subtraction() {
    let sheet_src = [
        ["1", "36", "20", "TRUE", "6"],
        ["-61", "\"bkh\"", "419.3782", "-80", "100"],
        ["91", "21", "\"oKzthzv\"", "-76", "-97"],
        ["-2", "FALSE", "232.9699", "", "-91.42"],
        ["-90", "-69", "-211.4693", "45", "10"],
        [
            "=AND(IF((D4 > A2), B5, E1) > 0, CONCATENATE(\"A4\", \"D1\") < 100)",
            "97.056",
            "=INT(SUM(C2:D2))",
            "=D2",
            "FALSE",
        ],
        [
            "=IF((D5 > (A2 / C1)), IF((A1 > A3), B3, A6), AND(32 > 0, D3 < 100))",
            "=E3",
            "=CONCATENATE(\"A2\", \"B2\")",
            "=B3",
            "=21",
        ],
        [
            "=MAX(D5:D6)",
            "TRUE",
            "=ROUNDDOWN(B7, 2)",
            "=IF((ROUNDUP(B4, 2) > INT(A6)), (-13 ^ -44), LEN(\"B2\"))",
            "FALSE",
        ],
        ["0", "=ROUNDUP(D6, 0)", "=E6", "=-24", "=D6"],
        [
            "=C1",
            "TRUE",
            "=D9",
            "=MIN(A6:B7)",
            "=CONCATENATE(\"D6\", \"A1\")",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 1));
    match target {
        ResultData::Float(f) => assert_eq!(f, -80.0),
        other => panic!("Expected Float(-80.0) for B9, got {:?}", other),
    }
}

#[test]
fn test_fuzz_zero_result_division_expression() {
    let sheet_src = [
        ["17", "\"cxFkHkuB\"", "TRUE", "TRUE", "-362.7"],
        ["60", "471.19", "", "\"p\"", "0"],
        ["", "-41", "-12.175", "22.876", "-127.099"],
        ["", "435.35", "77", "\"clH\"", "-5"],
        ["166.9", "29", "63", "-31", "3"],
        [
            "=A5",
            "=IF(((-1 + E3) > MIN(A3:C4)), SQRT(E1), C1)",
            "=LOWER(\"A3\")",
            "=RIGHT(\"15\", 2)",
            "44",
        ],
        [
            "500",
            "=AND(E4 > 0, E5 < 100)",
            "=-37",
            "FALSE",
            "=(MIN(C4:E4) - IF((C2 > C2), -3, E4))",
        ],
        [
            "=IF((B6 > ABS(D5)), (-16 * -6), D3)",
            "=E2",
            "=IF((OR(B2 > 0, E6 < 100) > ROUND(D5, 0)), LEN(\"-5\"), A6)",
            "=INT(B7)",
            "=MIN(E2:E2)",
        ],
        ["FALSE", "=-14", "=C5", "=(E7 * LEN(\"D3\"))", "=ABS(B7)"],
        [
            "=LEFT(\"B1\", 2)",
            "=RIGHT(\"C3\", 2)",
            "=33",
            "-133.1505",
            "=E4",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 3));
    match target {
        ResultData::Float(f) => assert_eq!(f, 0.0),
        other => panic!("Expected Float(0.0) for D9, got {:?}", other),
    }
}

#[test]
fn test_fuzz_power_cell_references() {
    let sheet_src = [
        ["-80", "-85", "40", "-95", "30"],
        ["-452.5", "\"LVnJoc\"", "", "30", "-77.1208"],
        ["\"2cB\"", "249.6382", "FALSE", "65", "TRUE"],
        ["-33", "\"UbxkVL \"", "53", "32", "FALSE"],
        ["TRUE", "0", "-38", "\"yOaI\"", "FALSE"],
        ["\"lQNfLl1d\"", "=D3", "4", "=SUM(D3:E5)", "=A5"],
        [
            "",
            "=AVERAGE(D1:E6)",
            "=UPPER(\"SQRT(B6)\")",
            "=(IF((-21 > D1), -14, C4) - SQRT(A4))",
            "=C1",
        ],
        [
            "=MIN(B6:B7)",
            "=(OR(E4 > 0, C3 < 100) * C2)",
            "=OR(LEFT(\"A6\", 1) > 0, E5 < 100)",
            "=32",
            "=B5",
        ],
        [
            "36",
            "=(E4 ^ D7)",
            "=AVERAGE(E8:E8)",
            "=OR(IF((D5 > 20), 41, D4) > 0, (A3 - C3) < 100)",
            "=B6",
        ],
        ["=C7", "=25", "=B7", "=B5", "=MAX(C6:D9)"],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 1));
    assert!(
        matches!(target, ResultData::Error(ref e) if e == "#NUM!"),
        "Expected #NUM! for B9, got {:?}",
        target
    );
}

#[test]
fn test_fuzz_addition_cell_references() {
    let sheet_src = [
        ["-75", "FALSE", "\"meC nxb\"", "FALSE", "FALSE"],
        ["FALSE", "1", "-443", "TRUE", "358.3"],
        ["63", "73", "TRUE", "FALSE", "423.462"],
        ["\"vJvhAgQw\"", "83", "-98", "-182", "-25"],
        ["", "FALSE", "397.034", "\"smWGqaU\"", "TRUE"],
        [
            "=LEFT(\"D3\", 5)",
            "=(A5 + LOWER(\"-20\"))",
            "=OR(C5 > 0, IF((D1 > 38), 28, C2) < 100)",
            "=(C3 * E5)",
            "=A5",
        ],
        [
            "=AVERAGE((B3 + D1), D4)",
            "=AND(B4 > 0, INT(E4) < 100)",
            "=23",
            "=16",
            "=15",
        ],
        [
            "=B5",
            "=B7",
            "=(LOWER(\"A5\") + B2)",
            "=((-46 / C4) ^ B4)",
            "=MAX(SQRT(E1), ABS(D2))",
        ],
        [
            "=(B6 + D8)",
            "=LOWER(\"AVERAGE(C8:E8)\")",
            "=A1",
            "=MIN(D1:E1)",
            "=((8 / D2) + OR(B4 > 0, 18 < 100))",
        ],
        [
            "=26",
            "=A5",
            "=OR(17 > 0, C3 < 100)",
            "=UPPER(\"IF((B2 > C3), C1, -20)\")",
            "=-29",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 0));
    match target {
        ResultData::Float(f) => assert_eq!(f, -20.0),
        other => panic!("Expected Float(-20.0) for A9, got {:?}", other),
    }
}

#[test]
fn test_fuzz_multiplication_cell_references() {
    let sheet_src = [
        ["34", "", "14", "TRUE", "-89"],
        ["", "-34", "\"XMnG\"", "7", "FALSE"],
        ["-75", "FALSE", "-231", "TRUE", "0"],
        ["7", "\"P2z1\"", "", "-76", "17"],
        ["3", "FALSE", "TRUE", "79", "0"],
        [
            "=C4",
            "=PRODUCT(A2:A5)",
            "=UPPER(\"SQRT(A5)\")",
            "=E1",
            "-10.011",
        ],
        [
            "=ROUND(D1, 2)",
            "=INT(RIGHT(\"-3\", 3))",
            "=(D6 + ROUNDUP(C6, 0))",
            "=E4",
            "=INT(SUM(-42, C2))",
        ],
        [
            "=IF((C6 > LEN(\"D4\")), ROUND(A1, 1), B5)",
            "=6",
            "=MIN(A1:C5)",
            "=3",
            "=-7",
        ],
        ["=(ROUNDDOWN(C1, 0) + A4)", "187.4", "=-42", "=INT(A5)", "1"],
        [
            "=(C9 * E1)",
            "-140.95",
            "=UPPER(\"(-46 + A2)\")",
            "=-7",
            "=E1",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    match target {
        ResultData::Float(f) => assert_eq!(f, 3738.0),
        other => panic!("Expected Float(3738.0) for A10, got {:?}", other),
    }
}

#[test]
fn test_fuzz_multiplication_division_expression() {
    let sheet_src = [
        ["\" a1 2oi\"", "8", "49", "10", "\"hUZmik3E\""],
        ["-11.1", "-311.1075", "", "33", "157.41"],
        ["78", "-60", "-49", "-34", "\"M\""],
        ["", "70", "-93", "\"1Yyyi2M\"", "FALSE"],
        ["\"IooXW\"", "0", "-63", "424.7", ""],
        [
            "=IF((-39 > D3), (B5 / B4), OR(A5 > 0, A4 < 100))",
            "=B5",
            "=19",
            "20",
            "=AND(LEN(\"C4\") > 0, MIN(D2, E4) < 100)",
        ],
        [
            "=(30 * (D6 / E6))",
            "322",
            "=(CONCATENATE(\"-27\", \"1\") - OR(D6 > 0, C2 < 100))",
            "\"DIbqt\"",
            "-74",
        ],
        ["=(B6 / A5)", "=ROUND(26, 1)", "=40", "=D6", "=INT(27)"],
        [
            "=ROUNDDOWN(E3, 0)",
            "30",
            "=8",
            "=(IF((-23 > C1), C8, -46) / SQRT(-36))",
            "=INT(A6)",
        ],
        [
            "47",
            "=(LEFT(\"E6\", 4) + E6)",
            "=ROUNDDOWN(4, 1)",
            "=D1",
            "=A3",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 0));
    match target {
        ResultData::Float(f) => assert_eq!(f, 600.0),
        other => panic!("Expected Float(600.0) for A7, got {:?}", other),
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
fn test_fuzz_subtraction_cell_references() {
    let sheet_src = [
        ["FALSE", "-60", "TRUE", "-242", "22"],
        ["TRUE", "", "\"Qo2z\"", "-22", "50"],
        ["\"e \"", "40", "56", "-56", "-74"],
        ["-20", "FALSE", "-220.4508", "", "-33"],
        ["\" To\"", "\"Tk1\"", "FALSE", "\"aNlj\"", "TRUE"],
        [
            "=ABS(C2)",
            "=(IF((A4 > E2), A5, 40) ^ IF((C1 > -23), C1, C5))",
            "=E4",
            "=INT(IF((C4 > C4), B4, E3))",
            "1",
        ],
        [
            "=(D3 - D6)",
            "=(SQRT(-40) ^ MIN(D2:E3))",
            "\"o3rTx\"",
            "=(A6 + D6)",
            "TRUE",
        ],
        [
            "=D5",
            "=IF((-39 > LEN(\"C7\")), A7, A4)",
            "6",
            "-15",
            "=MIN(LEN(\"E6\"), -20)",
        ],
        ["=C3", "=UPPER(\"C3\")", "=(-26 ^ C5)", "=17", "=E3"],
        [
            "=A2",
            "=SQRT(OR(9 > 0, C8 < 100))",
            "=UPPER(\"IF((E7 > D8), 6, -29)\")",
            "=IF((A7 > (E2 + E8)), B5, C9)",
            "=E2",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 0));
    match target {
        ResultData::Float(f) => assert_eq!(f, 18.0),
        other => panic!("Expected Float(18.0) for A7, got {:?}", other),
    }
}

#[test]
fn test_fuzz_positive_base_power_underflow_zero() {
    let sheet_src = [
        ["56", "", "", "\"ovmGQ\"", "2"],
        ["214.841", "\" UcRcJk\"", "-28", "22.02", "TRUE"],
        ["351.7", "", "0", "499.82", "-307.6667"],
        ["\"gVA3\"", "", "0", "311.7", "TRUE"],
        ["44", "399", "2", "-24", "-166.6152"],
        [
            "=(B1 / -26)",
            "\"sv2XVCLY\"",
            "=C2",
            "=SUM(AND(B5 > 0, A5 < 100), A3)",
            "=MAX(B4:C5)",
        ],
        [
            "=CONCATENATE(\"ROUNDUP(E2, 2)\", \"IF((C6 > B2), E5, -37)\")",
            "=(C4 - -1)",
            "=((D6 ^ -39) ^ 48)",
            "=B4",
            "=ROUNDDOWN(LOWER(\"-34\"), 0)",
        ],
        [
            "=MIN(15, C3)",
            "91",
            "=SUM(B2:C3)",
            "=LOWER(\"B7\")",
            "TRUE",
        ],
        [
            "=AVERAGE(D4:D6)",
            "=SQRT(PRODUCT(C7, 15))",
            "=LEN(\"A2\")",
            "=C2",
            "=-17",
        ],
        [
            "=LOWER(\"D2\")",
            "=C4",
            "=-47",
            "=0",
            "=AND(SQRT(E4) > 0, -4 < 100)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 2));
    println!("Seed 522134 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 0.0),
        ResultData::Integer(i) => assert_eq!(i, 0),
        other => panic!("Expected 0.0, got {:?}", other),
    }
}

#[test]
fn test_fuzz_zero_base_positive_exponent_evaluation() {
    let sheet_src = [
        ["0", "5", "10", "15", "20"],
        ["0", "1", "2", "3", "4"],
        ["0", "0", "0", "0", "0"],
        ["1", "2", "3", "4", "5"],
        ["-1", "-2", "-3", "-4", "-5"],
        ["=A1", "=B1", "=C1", "=D1", "=E1"],
        ["=A2", "=B2", "=C2", "=D2", "=E2"],
        ["=A3", "=B3", "=C3", "=D3", "=E3"],
        ["=A4", "=B4", "=C4", "=D4", "=E4"],
        [
            "=A1",
            "=B1",
            "=C1",
            "=D1",
            "=(E3 * (A1 ^ B1))",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 4));
    match target {
        ResultData::Float(f) => assert_eq!(f, 0.0),
        other => panic!("Expected Float(0.0), got {:?}", other),
    }
}

#[test]
fn test_fuzz_int_concatenate_month_two_digit_year_as_date() {
    let sheet_src = [
        ["24.0319", "\"B\"", "106.3654", "376", "38"],
        ["-276", "394", "FALSE", "-277.8755", ""],
        ["FALSE", "TRUE", "5", "1", "9"],
        ["", "-34", "29", "8", "0"],
        ["47", "210.486", "68", "", "-3"],
        ["=C1", "=INT(CONCATENATE(D3, B4))", "=A2", "\"H\"", "=IF((B5 > C1), B1, PRODUCT(B4:E4))"],
        ["=D2", "=AVERAGE(E1:E6)", "=AND(SUM(C1:E4) > 0, IF((A6 > D4), D3, A3) < 100)", "=ROUNDDOWN(C6, 2)", "=IF(((12 - 50) > LEFT(D1, 2)), (B1 - A5), A6)"],
        ["=E4", "=LOWER(B2)", "", "=A4", "=AND(C3 > 0, ROUNDUP(-1, 0) < 100)"],
        ["=-25", "=(C8 - E7)", "=PRODUCT((B5 ^ C4), IF((C3 > B7), 4, A5))", "=LEFT(C8, 3)", "=A7"],
        ["=C2", "=C3", "=B4", "=LOWER((E2 - D3))", "=-40"],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(5, 1));
    println!("Seed 87588 target B6: {:?}", target);
    match target {
        ResultData::Float(f) => assert!(
            (f - 12420.0).abs() < 1e-6,
            "Expected 12420 for B6 (CONCATENATE(D3, B4) = \"1-34\", which Excel parses as the date Jan 1934), got {:?}",
            target
        ),
        ResultData::Integer(i) => assert_eq!(i, 12420),
        other => panic!("Expected 12420 for B6, got {:?}", other),
    }
}
