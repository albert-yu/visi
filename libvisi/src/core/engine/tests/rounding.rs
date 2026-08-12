use super::*;

#[test]
fn test_fuzz_rounddown_scaled_float_precision() {
    let sheet_src = [
        ["\"QA3T3dlz\"", "-304", "38", "-230", "-298.006"],
        ["-159.7126", "", "-224", "", "FALSE"],
        ["43", "1", "", "-65", "0"],
        ["-445.02", "TRUE", "-58", "79", "-44"],
        ["", "63.5", "133.3", "99.2978", "275"],
        [
            "",
            "=E2",
            "=ABS(SUM(E1:E4))",
            "=(A5 - -15)",
            "=OR(PRODUCT(27, B4) > 0, (B4 * A3) < 100)",
        ],
        [
            "=D2",
            "=(INT(B1) * PRODUCT(B6, 46))",
            "=B4",
            "",
            "=UPPER(\"3\")",
        ],
        [
            "=AND(AVERAGE(13, -35) > 0, E3 < 100)",
            "=AVERAGE((E7 - D5), C5)",
            "=B1",
            "=E7",
            "-26",
        ],
        [
            "=A2",
            "=SUM(C6, LOWER(\"E2\"))",
            "=SUM(C3:D4)",
            "",
            "=(C4 ^ 12)",
        ],
        ["=-9", "=B9", "=-28", "=E5", "-399.593"],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 1));
    match target {
        ResultData::Float(f) => assert!(
            (f - 18.5011).abs() < 1e-3,
            "Expected ~18.5011 for B8, got {}",
            f
        ),
        other => panic!("Expected Float(~18.5011) for B8, got {:?}", other),
    }
}

#[test]
fn test_fuzz_roundup_large_exponent_precision() {
    let sheet_src = [
        ["278", "94", "", "\"Wb\"", "-98"],
        ["-19", "FALSE", "", "50", "0"],
        ["3", "400.65", "270.96", "5", "-24"],
        ["469.2", "-81", "-99", "-68", "-23"],
        ["372.65", "\"Gao \"", "62", "", "-54"],
        [
            "=OR(D4 > 0, OR(B3 > 0, A2 < 100) < 100)",
            "88",
            "=LEFT(\"E1\", 4)",
            "=(C3 - C5)",
            "=INT(D2)",
        ],
        [
            "=MAX(B2:D2)",
            "=B6",
            "=(SQRT(E6) + INT(24))",
            "=B3",
            "=ROUNDDOWN(B3, 1)",
        ],
        [
            "FALSE",
            "=INT(B4)",
            "=MAX(ROUNDDOWN(A2, 0), ROUNDDOWN(A5, 2))",
            "=32",
            "=ABS(41)",
        ],
        ["=B7", "=LEN(\"(C3 / E5)\")", "", "=D1", "=17"],
        [
            "=SQRT((-6 / 48))",
            "=B7",
            "=((B9 - D7) / B3)",
            "=IF((LEN(\"E8\") > CONCATENATE(\"E1\", \"C8\")), LOWER(\"E6\"), B8)",
            "=46",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 2));
    match target {
        ResultData::Float(f) => assert!(
            (f - 31.071067).abs() < 1e-3,
            "Expected ~31.071067 for C7, got {}",
            f
        ),
        other => panic!("Expected Float(~31.071067) for C7, got {:?}", other),
    }
}

#[test]
fn test_fuzz_subtraction_division_by_round() {
    let sheet_src = [
        ["-72", "203.37", "\"COp\"", "\"eNZIbOll\"", "\"CscRQlf\""],
        ["-75", "-151", "-79", "296.7224", "\"C\""],
        ["474.8", "TRUE", "TRUE", "-386.7184", "-265.71"],
        ["364.52", "51", "3", "-303.122", "-12"],
        ["-307.44", "46", "-125.6", "-97", ""],
        ["=-3", "-219", "=C3", "\"pHXeHmLw\"", "=PRODUCT(-1, A5)"],
        ["1", "-357.5857", "=ABS(B1)", "1", "=SQRT(AVERAGE(E2:E3))"],
        [
            "=ROUND((C2 + -15), 0)",
            "=AVERAGE(B5:E6)",
            "=E1",
            "=MAX(SUM(-38, C1), LOWER(\"E7\"))",
            "7",
        ],
        [
            "=SUM(INT(A8), UPPER(\"-11\"))",
            "=ABS(C2)",
            "428.9",
            "=ABS(IF((B5 > 18), 5, B6))",
            "=ROUNDUP(C1, 0)",
        ],
        [
            "=RIGHT(\"AND(E9 > 0, 45 < 100)\", 4)",
            "=IF((LEN(\"B8\") > B8), A7, E6)",
            "=D3",
            "=MIN(A5:E9)",
            "=((B6 - A5) / ROUND(-38, 0))",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 0));
    match target {
        ResultData::Float(f) => assert_eq!(f, -105.0),
        other => panic!("Expected Float(-105.0) for A9, got {:?}", other),
    }
}

#[test]
fn test_fuzz_round_nested_precision() {
    let sheet_src = [
        ["0", "-396", "TRUE", "-68", "-24.717"],
        ["\"Zurel\"", "356", "-62", "\"agptcdix\"", "216.82"],
        ["-286.82", "-62", "-306", "53", "24"],
        ["-60", "3", "68", "314", "-74"],
        ["-386.513", "", "-44", "", "\"gWmHnj\""],
        [
            "63",
            "=ROUND(IF((-28 > B3), D2, D2), 2)",
            "",
            "=(OR(-6 > 0, B2 < 100) + RIGHT(\"C1\", 3))",
            "-284.79",
        ],
        ["=A6", "=40", "=ROUNDUP(E6, 2)", "=AVERAGE(C2, B3)", "=E6"],
        [
            "=C2",
            "=D3",
            "=LOWER(\"(B1 / E3)\")",
            "=(E1 + 49)",
            "=CONCATENATE(\"C7\", \"A2\")",
        ],
        [
            "=ROUNDDOWN((A1 - 50), 1)",
            "=IF((C4 > (-42 + A2)), IF((B6 > A4), -7, A8), IF((D6 > 9), 7, 33))",
            "=A4",
            "=UPPER(\"(C3 ^ -48)\")",
            "=ABS(CONCATENATE(\"E7\", \"B2\"))",
        ],
        [
            "=ROUNDDOWN(UPPER(\"C3\"), 1)",
            "0",
            "\"K saoEGH\"",
            "=RIGHT(\"(-24 ^ E2)\", 4)",
            "=D4",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 2));
    match target {
        ResultData::Float(f) => assert!(
            (f - -284.79).abs() < 1e-2,
            "Expected ~-284.79 for C7, got {}",
            f
        ),
        other => panic!("Expected Float(~-284.79) for C7, got {:?}", other),
    }
}

#[test]
fn test_fuzz_power_integer_exponents() {
    let sheet_src = [
        ["47", "TRUE", "253.662", "65", "\"u1xQ E1D\""],
        ["-70", "-60", "\"oXvFcFo\"", "24", "90.5"],
        ["46.9", "69", "\"x\"", "\"HaNh i\"", "TRUE"],
        ["38", "98.191", "", "\"adTyY\"", "-329.96"],
        ["-99", "-216.193", "9", "56", "-14"],
        ["=D5", "=A1", "=C2", "=ROUND(MIN(-34, B5), 0)", "10"],
        [
            "=ROUNDDOWN(18, 0)",
            "=-1",
            "=IF((-30 > SUM(D6:D6)), SUM(D6:E6), MIN(B1, D4))",
            "=47",
            "=RIGHT(\"D2\", 2)",
        ],
        ["-79", "=(E1 - -9)", "-1", "=A3", "78"],
        [
            "=(A8 - MAX(B4, D5))",
            "=MAX(C7, OR(6 > 0, C3 < 100))",
            "=C1",
            "=-42",
            "=E2",
        ],
        [
            "=ABS(IF((C8 > A1), -24, E1))",
            "=ABS(D6)",
            "=A4",
            "=ABS(D6)",
            "\"p\"",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 1));
    match target {
        ResultData::Float(f) => assert_eq!(f, 216.0),
        other => panic!("Expected Float(216.0) for B10, got {:?}", other),
    }
}

#[test]
fn test_fuzz_roundup_power_expression() {
    let sheet_src = [
        ["82", "5", "91", "TRUE", "\"zLW\""],
        ["-52", "5", "\" ey\"", "-231.6", "-70"],
        ["-9", "72", "150", "286.5", "-11"],
        ["\"Nup Bkxr\"", "-12", "FALSE", "-419.2166", "-78"],
        ["\"Jxd2 \"", "12", "", "1", "17"],
        [
            "=UPPER(\"OR(-17 > 0, 10 < 100)\")",
            "=RIGHT(\"ROUNDDOWN(E4, 1)\", 5)",
            "=D4",
            "=(PRODUCT(A4, C4) * C4)",
            "=LEN(\"E5\")",
        ],
        ["=B6", "=ABS(ABS(-33))", "=-20", "\"CkAx1j\"", "=A6"],
        [
            "=C5",
            "=UPPER(\"ROUNDUP(C3, 2)\")",
            "=PRODUCT(IF((28 > -42), -24, D3), -3)",
            "=MIN(A6:E6)",
            "-60.743",
        ],
        [
            "=(-12 ^ D4)",
            "=OR(RIGHT(\"A2\", 2) > 0, C5 < 100)",
            "=(C4 * IF((B8 > D5), C5, 28))",
            "=A5",
            "=AND(-45 > 0, AND(A2 > 0, D3 < 100) < 100)",
        ],
        [
            "-10",
            "=ROUNDUP((A1 ^ C8), 1)",
            "=ROUND(IF((A3 > E7), -6, C4), 0)",
            "=16",
            "=ROUNDUP(D8, 1)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 1));
    match target {
        ResultData::Float(f) => assert!(
            (f - 6.231568e137).abs() / f < 1e-3,
            "Expected ~6.231568e137 for B10, got {}",
            f
        ),
        other => panic!("Expected Float(~6.231568e137) for B10, got {:?}", other),
    }
}

#[test]
fn test_fuzz_abs_cell_reference() {
    let sheet_src = [
        ["", "79.54000000000001", "34.0806", "-410", ""],
        ["-336.1717", "\"Dz\"", "FALSE", "-90", "80"],
        ["-11", "-88.38", "-2", "TRUE", "-291"],
        ["39", "228.77", "\"jo\"", "-41", "8.827999999999999"],
        ["", "309.53", "1", "98", "0"],
        [
            "=-16",
            "=AVERAGE(C5:E5)",
            "=A5",
            "=(ROUNDDOWN(A2, 2) ^ IF((45 > E2), C2, 21))",
            "=UPPER(\"C4\")",
        ],
        [
            "=AND(UPPER(\"B4\") > 0, LEFT(\"C1\", 3) < 100)",
            "=ABS(C6)",
            "=AVERAGE(E1, E5)",
            "=D6",
            "=E2",
        ],
        [
            "=MAX(E3:E3)",
            "-282",
            "=D5",
            "=MAX(IF((B3 > C7), -36, 16), RIGHT(\"C5\", 3))",
            "=PRODUCT(INT(44), E1)",
        ],
        ["FALSE", "=D6", "-56", "=D1", "=-32"],
        [
            "=(SUM(E3:E9) ^ 35)",
            "=ROUNDDOWN((B1 ^ -9), 0)",
            "=B4",
            "=LEN(\"D3\")",
            "=SUM(B2:E3)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 1));
    match target {
        ResultData::Float(f) => assert_eq!(f, 0.0),
        other => panic!("Expected Float(0.0) for B7, got {:?}", other),
    }
    for r in 0..10 {
        for c in 0..5 {
            let res = sheet.get_result_data(&CellRef::new(r, c));
            if matches!(res, ResultData::Error(_)) {
                let col_let = (b'A' + c as u8) as char;
                assert!(('A'..='E').contains(&col_let));
            }
        }
    }
}

#[test]
fn test_fuzz_abs_multiplication_precision() {
    let sheet_src = [
        ["\"ob\"", "", "-8", "2", "322"],
        ["-39", "FALSE", "TRUE", "\"sivgqf\"", ""],
        ["7", "FALSE", "-14", "269", "371.3"],
        ["-202.535", "0", "TRUE", "422.74", "426.4483"],
        ["0", "9", "\"sT3\"", "TRUE", "340.7"],
        [
            "=C3",
            "=PRODUCT((-22 * E1), -48)",
            "=ROUNDUP(43, 1)",
            "=IF((D2 > E1), AVERAGE(B5, A5), A4)",
            "",
        ],
        [
            "=UPPER(\"-26\")",
            "4",
            "=-38",
            "=IF((A5 > ROUND(D4, 1)), LEFT(\"D2\", 1), AVERAGE(D6:D6))",
            "35",
        ],
        ["\"Sxz1Rx Q\"", "=30", "=SUM(C7:D7)", "256.936", "=43"],
        [
            "=ABS((D7 * E3))",
            "=E3",
            "=SQRT(MAX(B8:D8))",
            "=-25",
            "=-47",
        ],
        [
            "=IF((C4 > RIGHT(\"B7\", 3)), E2, C9)",
            "=SQRT(OR(-3 > 0, D1 < 100))",
            "=LEFT(\"C6\", 4)",
            "-71",
            "=D1",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 0));
    match target {
        ResultData::Float(f) => assert!((f - 1670.85).abs() < 1e-3),
        other => panic!("Expected Float(~1670.85), got {:?}", other),
    }
}

#[test]
fn test_fuzz_abs_subtraction_expression() {
    let sheet_src = [
        ["91", "\"d\"", "60", "331.25", "51"],
        ["22.77", "-34", "-15", "-94", "2"],
        ["-22", "155.8", "-320.5996", "\"V\"", "FALSE"],
        ["TRUE", "FALSE", "-460.0053", "-83", "19.159"],
        ["-55", "1", "\"yHs lFZi\"", "\"nZ\"", "FALSE"],
        [
            "=LOWER(\"OR(D2 > 0, A2 < 100)\")",
            "=D3",
            "195.5917",
            "=IF(((C2 * 16) > INT(-23)), ROUND(A1, 2), (B4 - B5))",
            "=13",
        ],
        [
            "=OR(RIGHT(\"B3\", 1) > 0, SQRT(-19) < 100)",
            "=ROUNDUP(ROUNDUP(B5, 0), 2)",
            "=C2",
            "=-7",
            "=-26",
        ],
        ["=(LEN(\"A2\") / A2)", "", "=ABS((D7 + -26))", "=B5", "=D7"],
        [
            "=D1",
            "=46",
            "=(A3 - LEFT(\"37\", 2))",
            "=LEN(\"B8\")",
            "=-11",
        ],
        [
            "85",
            "=C8",
            "=IF(((D5 / 44) > (A6 + C9)), PRODUCT(D8:D8), CONCATENATE(\"E6\", \"E1\"))",
            "\"cmWb\"",
            "=IF((ROUNDUP(D6, 2) > D4), PRODUCT(A9:D9), AND(A4 > 0, -10 < 100))",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 2));
    match target {
        ResultData::Float(f) => assert_eq!(f, 33.0),
        other => panic!("Expected Float(33.0) for C8, got {:?}", other),
    }
    for r in 0..10 {
        for c in 0..5 {
            let res = sheet.get_result_data(&CellRef::new(r, c));
            if matches!(res, ResultData::Error(_)) {
                let col_let = (b"A"[0] + c as u8) as char;
                assert!(('A'..='E').contains(&col_let));
            }
        }
    }
}

#[test]
fn test_fuzz_roundup_division_negative_zero() {
    let sheet_src = [
        ["149.7872", "258.543", "", "12", "FALSE"],
        ["\"3qMv1d\"", "444.9831", "TRUE", "", "\"z\""],
        ["\"LLw2Eeu\"", "\"BQQ\"", "", "\"phJdTzJ\"", "FALSE"],
        ["-94", "FALSE", "TRUE", "59.1905", "\"qswbs\""],
        ["482.7", "7", "60", "\"FHL\"", "80"],
        [
            "=D1",
            "=B5",
            "=-26",
            "=AND(-10 > 0, IF((35 > -23), 36, 49) < 100)",
            "=ROUNDDOWN(-35, 2)",
        ],
        [
            "=C3",
            "=D3",
            "FALSE",
            "=ROUNDUP(SUM(E1, C4), 0)",
            "=IF((D4 > AND(C2 > 0, -33 < 100)), E6, 29)",
        ],
        [
            "=25",
            "=(A7 * (D5 - 37))",
            "-37",
            "=(ROUNDUP(-14, 2) ^ -1)",
            "=SQRT(IF((B1 > A4), 33, C3))",
        ],
        [
            "FALSE",
            "=ROUNDUP((A7 / D8), 2)",
            "=C2",
            "=(E4 ^ AVERAGE(-19, E4))",
            "=UPPER(\"(C2 * B4)\")",
        ],
        [
            "=C3",
            "=(AND(B2 > 0, C8 < 100) + ROUNDUP(A8, 0))",
            "=B4",
            "=MIN(A5:D6)",
            "=C3",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 1));
    match target {
        ResultData::Float(f) => assert_eq!(f.abs(), 0.0),
        other => panic!("Expected Float(~0.0), got {:?}", other),
    }
}

#[test]
fn test_fuzz_round_cell_reference() {
    let sheet_src = [
        [
            "-139.3711",
            "74.28100000000001",
            "6.793",
            "-277",
            "\"QrVXAE\"",
        ],
        ["83", "129.4", "\"WRedVYxM\"", "427.42", "-62"],
        ["FALSE", "TRUE", "91", "83", "2"],
        ["-45", "7", "396", "-52", "0"],
        ["-64", "", "", "", "TRUE"],
        [
            "=SQRT(A1)",
            "=C2",
            "=(B5 + B4)",
            "=(INT(B1) ^ (1 - -8))",
            "\"BHED\"",
        ],
        [
            "=((B5 / D1) ^ -47)",
            "=MIN(A6:C6)",
            "-265.176",
            "=IF((RIGHT(\"B4\", 1) > AND(E2 > 0, C5 < 100)), C4, D1)",
            "=C2",
        ],
        ["34", "=(-42 / (20 - B7))", "=UPPER(\"D3\")", "=E5", "63"],
        [
            "25",
            "=LOWER(\"(B3 - D3)\")",
            "=ABS(UPPER(\"D5\"))",
            "=SUM((A1 ^ C3), -22)",
            "=ROUNDDOWN(IF((E1 > -36), A5, C8), 0)",
        ],
        [
            "=32",
            "=ROUND(C6, 2)",
            "=B7",
            "=(C5 + D8)",
            "=((C7 + B3) / E6)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 1));
    match target {
        ResultData::Float(f) => assert_eq!(f, 7.0),
        other => panic!("Expected Float(7.0), got {:?}", other),
    }
}

#[test]
fn test_fuzz_rounddown_cell_reference() {
    let sheet_src = [
        ["4", "-496.2", "", "5", "-38"],
        ["85", "\"HcOUVBwS\"", "0", "0", "87"],
        ["\"3c\"", "187", "-182.3", "", "-370.66"],
        ["402.72", "-349.5425", "79", "2", "83"],
        ["\"p\"", "56.82", "-4", "60", "-48"],
        [
            "-125.224",
            "=-36",
            "=(IF((A5 > C3), B3, C1) * ABS(C3))",
            "=AND(MIN(E1:E4) > 0, ROUNDDOWN(A4, 0) < 100)",
            "138",
        ],
        [
            "=C5",
            "=CONCATENATE(\"IF((B4 > C3), B6, C6)\", \"INT(33)\")",
            "=E3",
            "=(A4 ^ AVERAGE(B5:B5))",
            "=CONCATENATE(\"D5\", \"B2\")",
        ],
        [
            "=B6",
            "=39",
            "=D6",
            "=PRODUCT(E7:E7)",
            "=CONCATENATE(\"E4\", \"AND(C5 > 0, E5 < 100)\")",
        ],
        [
            "=ROUNDDOWN(B8, 0)",
            "=MAX(B2:C7)",
            "=(B8 + -21)",
            "=SUM(E3:E7)",
            "=ROUND(SUM(B8:C8), 2)",
        ],
        [
            "=AND(C8 > 0, 36 < 100)",
            "=INT(E9)",
            "=OR(PRODUCT(B2, 2) > 0, LOWER(\"-6\") < 100)",
            "=AND(8 > 0, IF((C1 > 31), B1, E9) < 100)",
            "39",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 0));
    match target {
        ResultData::Float(f) => assert_eq!(f, 39.0),
        other => panic!("Expected Float(39.0) for A9, got {:?}", other),
    }
}

#[test]
fn test_fuzz_round_int_cell_reference() {
    let sheet_src = [
        ["-44", "0", "FALSE", "", "64"],
        ["-89", "8", "-39", "TRUE", "\" MHdPCGd\""],
        ["-67", "40", "FALSE", "", "-277.6"],
        ["206.81", "", "\"EPAIFa1j\"", "17", "-48"],
        ["", "-210.35", "-99", "-22", "-322.5"],
        [
            "=E5",
            "=INT(E1)",
            "=SQRT(RIGHT(\"B2\", 4))",
            "=MAX(LEN(\"B2\"), D1)",
            "=D1",
        ],
        [
            "=E5",
            "=AND(C2 > 0, (19 ^ B3) < 100)",
            "=-5",
            "=AND(OR(-1 > 0, E6 < 100) > 0, IF((A6 > C3), E6, C5) < 100)",
            "-29",
        ],
        [
            "=AVERAGE(E1, E2)",
            "=(MIN(A2, 26) + LEN(\"A2\"))",
            "=D5",
            "=D6",
            "=ROUND(PRODUCT(A5:E7), 2)",
        ],
        [
            "=LEFT(\"ROUNDDOWN(-46, 2)\", 5)",
            "=ROUNDUP(E2, 2)",
            "=(D1 * CONCATENATE(\"19\", \"38\"))",
            "=ABS(E4)",
            "=((-21 ^ A3) ^ (D3 ^ -12))",
        ],
        ["=ROUND(B6, 0)", "=-33", "=-1", "=MIN(C7:C9)", "12.395"],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    match target {
        ResultData::Float(f) => assert_eq!(f, 64.0),
        other => panic!("Expected Float(64.0), got {:?}", other),
    }
}

#[test]
fn test_fuzz_round_single_digit_cell() {
    let sheet_src = [
        ["-32", "FALSE", "FALSE", "-39", "87"],
        ["", "-13.52", "TRUE", "89.43000000000001", "94"],
        ["-59", "27", "84", "-482.5904", "2"],
        ["", "\"Q\"", "-30", "484.45", "24"],
        ["TRUE", "7", "TRUE", "5", "11"],
        [
            "=ROUND(INT(C3), 0)",
            "=B3",
            "=((-9 + C1) - LOWER(\"12\"))",
            "=OR(SUM(A3:C4) > 0, OR(A3 > 0, -13 < 100) < 100)",
            "=-12",
        ],
        [
            "=AND(AND(C5 > 0, E5 < 100) > 0, AND(13 > 0, 26 < 100) < 100)",
            "\"iUZS\"",
            "=(IF((42 > -35), E4, E4) ^ ABS(B2))",
            "=RIGHT(\"PRODUCT(E3, E4)\", 2)",
            "=MIN(A4:B4)",
        ],
        ["FALSE", "TRUE", "=(D1 ^ 50)", "88", "=C5"],
        [
            "=(ROUNDUP(C5, 0) ^ E4)",
            "=LEN(\"INT(D1)\")",
            "=ROUNDDOWN(4, 1)",
            "280.31",
            "=B6",
        ],
        [
            "=ROUND(B9, 1)",
            "=UPPER(\"SUM(-11, B8)\")",
            "211.044",
            "\"dl\"",
            "64",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    match target {
        ResultData::Float(f) => assert_eq!(f, 7.0),
        other => panic!("Expected Float(7.0) for A10, got {:?}", other),
    }
}

#[test]
fn test_fuzz_round_e9_precision() {
    let sheet_src = [
        ["TRUE", "61", "1", "-45", "-36"],
        ["FALSE", "-3.5449", "", "2", "-49"],
        ["-11", "-64", "-258.9581", "\"Ah\"", "40"],
        ["\"nkQDAkX\"", "54", "8", "6", "-54"],
        ["18", "\"UJrTlo3\"", "470.7", "128", ""],
        ["=SQRT(ABS(-3))", "=-8", "=7", "224", "=ROUNDUP(ABS(D4), 1)"],
        [
            "=IF((MAX(E5:E5) > AND(B4 > 0, E6 < 100)), AVERAGE(B6:B6), IF((E3 > E4), D6, B4))",
            "=C5",
            "=(AVERAGE(E1, A1) + ABS(-20))",
            "=(IF((25 > C6), A3, 18) * C6)",
            "=UPPER(\"B1\")",
        ],
        [
            "=-5",
            "=(INT(D2) + IF((E4 > E3), A1, E2))",
            "=(27 / (-5 - -10))",
            "-94",
            "=IF((CONCATENATE(\"D4\", \"3\") > LEN(\"E3\")), -16, SUM(A4, B7))",
        ],
        [
            "=C8",
            "=AVERAGE(C5:E8)",
            "=B4",
            "=ROUNDUP(SQRT(4), 1)",
            "=MAX((D6 / C5), D2)",
        ],
        [
            "=ROUND(E9, 0)",
            "=B6",
            "=SUM(E3:E9)",
            "=-39",
            "=SQRT(IF((-23 > E2), B1, A3))",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    println!("Seed 48394 evaluated target CellRef(9, 0): {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 2.0),
        other => panic!("Expected Float(2.0), got {:?}", other),
    }
}

#[test]
fn test_fuzz_rounddown_power_c3() {
    let sheet_src = [
        ["\" NBLpaTv\"", "36.56", "\"DrEszH\"", "", "31"],
        ["7", "0", "", "-435.2", "-41"],
        ["-138.433", "41", "0", "\"u\"", "\"wIxVI\""],
        ["", "\"HAGiDJE\"", "-395", "TRUE", ""],
        ["-58", "\"m\"", "", "0", "\"Uc\""],
        [
            "=LEFT(\"ROUND(D4, 0)\", 1)",
            "=UPPER(\"38\")",
            "=(ROUNDUP(B3, 2) + UPPER(\"D2\"))",
            "=D5",
            "=E2",
        ],
        [
            "=ROUND(MAX(A3, E3), 2)",
            "=(C4 * CONCATENATE(\"E6\", \"39\"))",
            "=D6",
            "=42",
            "=-46",
        ],
        [
            "=IF((B6 > INT(D3)), ROUNDDOWN(B1, 1), B5)",
            "=D3",
            "=LEN(\"C5\")",
            "=IF((16 > B7), SUM(B5:B5), OR(D4 > 0, A3 < 100))",
            "=A1",
        ],
        [
            "=IF((SQRT(A6) > LEFT(\"A5\", 3)), MAX(A6, C6), MIN(E3:E7))",
            "=C1",
            "=A7",
            "=AVERAGE(A7, PRODUCT(E7, C7))",
            "=INT(OR(E4 > 0, B3 < 100))",
        ],
        [
            "=(ROUNDDOWN(E9, 0) ^ C3)",
            "=LEFT(\"(C6 + A8)\", 5)",
            "",
            "=-44",
            "=SQRT(-19)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    println!("Seed 957517 evaluated target CellRef(9, 0): {:?}", target);
}

#[test]
fn test_fuzz_roundup_e6_cell_reference() {
    let sheet_src = [
        ["", "", "", "84", "53"],
        ["FALSE", "\"Q\"", "-40", "73", ""],
        ["43", "22", "", "3", "-68"],
        ["-23", "48", "-361.23", "56", "-140.93"],
        ["469.3", "-14", "-459.021", "-95", "362.62"],
        [
            "=ABS(IF((C3 > -44), A5, 40))",
            "=D3",
            "=B2",
            "8",
            "=ABS(-27)",
        ],
        [
            "=AND(D2 > 0, (A4 + 30) < 100)",
            "35",
            "=ROUNDUP(E6, 1)",
            "=OR(-1 > 0, D2 < 100)",
            "=SUM(D4:D5)",
        ],
        [
            "=E4",
            "\"JUO\"",
            "=B6",
            "=IF((D1 > CONCATENATE(\"E4\", \"1\")), ROUND(B5, 2), E5)",
            "=D6",
        ],
        [
            "=B5",
            "",
            "\"hIVGEy\"",
            "47",
            "=(UPPER(\"E5\") + (29 - -16))",
        ],
        [
            "=OR(5 > 0, -21 < 100)",
            "=-20",
            "=UPPER(\"IF((D6 > C4), C5, C8)\")",
            "=CONCATENATE(\"IF((D5 > C4), E7, -49)\", \"AVERAGE(-5, D5)\")",
            "=ABS(-22)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let c7 = sheet.get_result_data(&CellRef::new(6, 2));
    println!("Seed 931968 C7: {:?}", c7);
}

#[test]
fn test_fuzz_int_sqrt_d6() {
    let sheet_src = [
        ["\"mb2sof1i\"", "78", "444.5", "-55", "342.8"],
        ["-70", "62", "-254", "-183", ""],
        ["TRUE", "0", "\"OrlEQPfz\"", "\"sMRhHT\"", "TRUE"],
        ["", "-16", "\"mhdP\"", "76", "0"],
        ["-9", "\"KFHr2en\"", "-132.9", "36", "-365.88"],
        [
            "=PRODUCT(C5:E5)",
            "TRUE",
            "=PRODUCT(E5:E5)",
            "=SQRT(D4)",
            "=AND(E1 > 0, MAX(B5:C5) < 100)",
        ],
        [
            "\"CkwbTno\"",
            "-32",
            "=-43",
            "=C5",
            "=OR(D3 > 0, -50 < 100)",
        ],
        [
            "=D4",
            "=INT(D6)",
            "\"Ydy\"",
            "=UPPER(\"IF((19 > E7), A7, C4)\")",
            "=PRODUCT(SQRT(C5), A7)",
        ],
        [
            "=A4",
            "=(IF((B6 > -20), A8, 21) + (D4 / D1))",
            "236.302",
            "=OR(23 > 0, LEFT(\"D8\", 2) < 100)",
            "=E8",
        ],
        ["=MIN(D2:E3)", "=A7", "=SUM(E5:E6)", "TRUE", "=E1"],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 1));
    println!("Seed 544657 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 8.0),
        other => panic!("Expected Float(8.0), got {:?}", other),
    }
}

#[test]
fn test_fuzz_round_d6_single_decimal() {
    let sheet_src = [
        ["", "32", "-63", "-67", "41"],
        ["226.3", "184.8", "-87", "", "68"],
        ["", "51", "-53", "4", "249.3257"],
        ["66", "356", "FALSE", "295.24", ""],
        ["124.9", "-11", "\"hIwAgm\"", "TRUE", ""],
        [
            "=IF((INT(47) > SUM(C5:C5)), C1, (C2 / D3))",
            "=B2",
            "=IF((E5 > LOWER(\"C3\")), INT(E4), RIGHT(\"C4\", 2))",
            "=(9 - (E1 + 19))",
            "=E3",
        ],
        [
            "=1",
            "=(IF((B1 > -15), C4, C6) / AND(50 > 0, D2 < 100))",
            "=ROUNDUP(RIGHT(\"-44\", 4), 0)",
            "=-37",
            "0",
        ],
        ["=ROUND(D6, 1)", "=D5", "=D7", "28", "=D7"],
        ["=LEN(\"A1\")", "=E2", "=AVERAGE(C7, 30)", "=5", "=C6"],
        [
            "=E2",
            "=(IF((E9 > D9), 35, A3) * (C8 - C2))",
            "=50",
            "=C4",
            "=IF((PRODUCT(D3:E9) > -36), D1, -25)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 0));
    println!("Seed 780052 target: {:?}", target);
}

#[test]
fn test_fuzz_roundup_negative_tiny_exponent() {
    let sheet_src = [
        ["38", "", "49", "65", "-99"],
        ["33", "2", "\"WQk3\"", "TRUE", ""],
        ["\"d\"", "-96", "-97", "", "-312.17"],
        ["-89", "52", "\"AMQu\"", "50", "TRUE"],
        ["235.2", "FALSE", "\"nSp\"", "94.67", "\"hhoveY\""],
        [
            "=LOWER(\"(C1 / C4)\")",
            "-51",
            "=LEFT(\"-41\", 2)",
            "=21",
            "",
        ],
        [
            "=MIN(B5, SQRT(-21))",
            "=E2",
            "=PRODUCT(B6:B6)",
            "=OR(-28 > 0, D4 < 100)",
            "=PRODUCT(B4:B6)",
        ],
        [
            "=E4",
            "=E3",
            "=IF((B4 > AVERAGE(A5:B6)), C1, ROUND(B5, 1))",
            "=E7",
            "=E3",
        ],
        [
            "=D2",
            "=-25",
            "=LOWER(\"D3\")",
            "=SUM(AVERAGE(C7:C8), E6)",
            "=ROUND(E4, 2)",
        ],
        [
            "=PRODUCT(ROUNDUP(D1, 0), (E1 * A4))",
            "=OR(LEN(\"C1\") > 0, RIGHT(\"C8\", 3) < 100)",
            "=IF((-4 > SUM(-40, 22)), ROUNDUP(-28, 2), A5)",
            "=ROUNDUP((E8 ^ -29), 2)",
            "=A7",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let d10 = sheet.get_result_data(&CellRef::new(9, 3));
    println!("Seed 871434 D10: {:?}", d10);
    match d10 {
        ResultData::Float(f) => assert_eq!(f, -0.01),
        other => panic!("Expected Float(-0.01), got {:?}", other),
    }
}

#[test]
fn test_fuzz_addition_roundup_e7() {
    let sheet_src = [
        ["32", "", "", "2", "288.4"],
        ["-77", "-56", "-473.497", "-44", "219"],
        ["-264.47", "23", "", "", "29"],
        ["\"UI\"", "", "", "442.4109", "20"],
        ["99", "-97", "-373.8", "\"f2z\"", "64"],
        [
            "=E2",
            "=SUM(SUM(3, E5), MAX(C1, C5))",
            "=AND(OR(D3 > 0, C3 < 100) > 0, E3 < 100)",
            "=(-21 ^ OR(E5 > 0, A5 < 100))",
            "=AVERAGE(C5:E5)",
        ],
        [
            "\"v\"",
            "=E4",
            "=1",
            "=B4",
            "=(-47 + IF((D3 > A6), D5, 19))",
        ],
        [
            "=(MIN(A3:C6) + AND(E7 > 0, A6 < 100))",
            "\"a\"",
            "=(A2 - SUM(B4:B4))",
            "=ROUNDUP(UPPER(\"D6\"), 0)",
            "=C5",
        ],
        ["=LEFT(\"-38\", 3)", "=D8", "=44", "0", "=D1"],
        [
            "=(B3 + ROUNDUP(E7, 1))",
            "=A7",
            "=ROUND(UPPER(\"A2\"), 1)",
            "=21",
            "=MIN(D8:D8)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    println!("Seed 457220 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, -5.0),
        ResultData::Integer(i) => assert_eq!(i, -5),
        other => panic!("Expected -5, got {:?}", other),
    }
}

#[test]
fn test_fuzz_round_division_b2_neg13() {
    let sheet_src = [
        ["10", "209.3", "", "94", "-45.488"],
        ["", "283.3379", "224", "", "144.475"],
        ["-51", "-52", "\"SqAEKqrx\"", "-93", "-193.9086"],
        ["27", "21", "16", "TRUE", "\"MSG1hfJ\""],
        ["\"sEq\"", "9", "", "-85.3004", "TRUE"],
        ["=OR(B4 > 0, E5 < 100)", "=E1", "=ABS(B2)", "=A1", "=-40"],
        [
            "=IF((46 > (D2 + D1)), LEN(\"A5\"), C5)",
            "=-32",
            "=E1",
            "=(B2 / -13)",
            "=B3",
        ],
        [
            "=SUM(C1:D2)",
            "-24",
            "=-12",
            "=(B6 * RIGHT(\"A1\", 5))",
            "=AND(PRODUCT(A5, B5) > 0, E4 < 100)",
        ],
        [
            "=MIN(C4:D6)",
            "=MAX(A1:C5)",
            "=ROUND(D7, 2)",
            "=(PRODUCT(-41, B5) - B6)",
            "=-26",
        ],
        [
            "=(B2 - C1)",
            "=ABS(B9)",
            "=(ROUNDUP(-45, 2) + IF((C9 > D2), A7, D4))",
            "\"AeDrkE\"",
            "=(MAX(E4:E9) - ROUNDUP(-14, 1))",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 2));
    println!("Seed 845650 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, -21.8),
        ResultData::Integer(i) => assert_eq!(i, -22),
        other => panic!("Expected -21.8, got {:?}", other),
    }
}

#[test]
fn test_fuzz_round_d7_b10() {
    let sheet_src = [
        ["FALSE", "", "\"MSx\"", "\"acEbw\"", "0"],
        ["-17", "\"JkGE\"", "-83", "177.54", "-76"],
        ["-88", "183.0719", "-33", "-92", "\"hldh\""],
        ["\"sK\"", "-37", "\"tBtsh\"", "7", "60"],
        ["\"GiIU\"", "47", "32", "-216.9266", "\"OXH\""],
        [
            "=RIGHT(\"ROUNDDOWN(-43, 2)\", 2)",
            "=E3",
            "=E1",
            "=(AVERAGE(D4:D5) * E4)",
            "=SUM(E4, 35)",
        ],
        [
            "=ROUND((C2 / -2), 0)",
            "=OR(C6 > 0, UPPER(\"D6\") < 100)",
            "=SQRT(OR(E6 > 0, B1 < 100))",
            "=A3",
            "=D5",
        ],
        ["", "=11", "=A1", "=-26", "\"MSpw2\""],
        [
            "=LEFT(\"23\", 5)",
            "=OR(MIN(A1, B4) > 0, SUM(5, B3) < 100)",
            "97",
            "=ROUNDDOWN(D5, 1)",
            "=AVERAGE(B2:D7)",
        ],
        [
            "=(E3 / A9)",
            "=ROUND(D7, 2)",
            "=AND(AVERAGE(B7:C8) > 0, MIN(C6:D9) < 100)",
            "=A6",
            "=B9",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 1));
    println!("Seed 207640 target: {:?}", target);
}

#[test]
fn test_fuzz_abs_b8_cell_ref() {
    let sheet_src = [
        ["347", "9", "-82", "-37", "-112.89"],
        ["386.74", "TRUE", "TRUE", "100", "48"],
        ["46", "TRUE", "", "\"rASQncw\"", "\"VT\""],
        ["FALSE", "", "TRUE", "-27", "0"],
        ["0", "-381.1", "\"ozPy\"", "46", "41"],
        [
            "=C5",
            "=AND(C1 > 0, MAX(42, A2) < 100)",
            "=IF((A1 > 41), B4, A3)",
            "-26",
            "=SQRT(B3)",
        ],
        [
            "=E3",
            "=LOWER(\"(E6 - A3)\")",
            "=-46",
            "=ROUND(17, 1)",
            "=MIN(ROUNDUP(22, 0), LEFT(\"D3\", 3))",
        ],
        ["=MIN(A7:D7)", "=18", "=(A5 ^ INT(A7))", "=-42", "=10"],
        ["=A6", "=A1", "=E7", "=LEN(\"C3\")", "=LEFT(\"D7\", 1)"],
        [
            "=ABS(B8)",
            "=PRODUCT(IF((-30 > B7), -3, A7), IF((A1 > -37), B6, A6))",
            "311.4486",
            "=LEN(\"C4\")",
            "=E6",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    println!("Seed 142584 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 18.0),
        ResultData::Integer(i) => assert_eq!(i, 18),
        other => panic!("Expected 18, got {:?}", other),
    }
}

#[test]
fn test_fuzz_sqrt_d6_zero_value() {
    let sheet_src = [
        ["\"ZNC3o\"", "0", "479", "0", "-34"],
        ["47", "", "242.42", "36", "149.1177"],
        ["35", "10", "466.1", "-43.34", "9"],
        ["", "34", "", "-59", ""],
        ["", "-42", "343.927", "-36", "-338.9"],
        ["=ABS(A2)", "=-44", "=-30", "=C4", "=ROUNDUP(A2, 1)"],
        [
            "=AND(AND(C3 > 0, 29 < 100) > 0, C3 < 100)",
            "=-9",
            "=A3",
            "49.021",
            "=(B1 / IF((-18 > -26), 6, A3))",
        ],
        ["27", "=C2", "107.8808", "=E6", "-49"],
        [
            "=C1",
            "=SQRT(D6)",
            "=A3",
            "=(-38 + PRODUCT(C5, D6))",
            "=RIGHT(\"AND(D6 > 0, C8 < 100)\", 1)",
        ],
        [
            "=(E1 / RIGHT(\"E6\", 5))",
            "51",
            "=LEFT(\"PRODUCT(C3:E7)\", 5)",
            "-73",
            "=E7",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 1));
    println!("Seed 687697 target: {:?}", target);
}

#[test]
fn test_fuzz_sqrt_e6_power_e2() {
    let sheet_src = [
        ["\"pCQOcOPR\"", "-15", "-75", "36.3", "43"],
        ["", "TRUE", "0", "TRUE", "0"],
        ["7", "\"1\"", "", "TRUE", ""],
        ["-392.04", "382.4681", "\"jtih\"", "TRUE", "92"],
        ["0", "FALSE", "FALSE", "-10", "FALSE"],
        [
            "=-40",
            "-306",
            "=IF((E1 > SUM(B3:E4)), LEFT(\"C5\", 1), (E3 / -41))",
            "349",
            "=OR(OR(E1 > 0, A3 < 100) > 0, OR(B1 > 0, -46 < 100) < 100)",
        ],
        [
            "-58",
            "FALSE",
            "=IF((D5 > (-8 + B2)), (-48 * D6), (A6 / C5))",
            "=-39",
            "=-3",
        ],
        [
            "=B5",
            "=AND(9 > 0, PRODUCT(A1:E2) < 100)",
            "94",
            "=(MAX(-36, D3) * ROUND(A2, 1))",
            "79",
        ],
        [
            "=LEFT(\"(E5 * E5)\", 5)",
            "=(SQRT(E6) ^ E2)",
            "=B2",
            "=ROUNDUP(SUM(B8:D8), 2)",
            "=IF((RIGHT(\"A1\", 4) > ROUNDUP(D2, 1)), AND(9 > 0, 22 < 100), ROUNDUP(-29, 1))",
        ],
        [
            "\"asFud3\"",
            "=IF((D8 > (B4 + A5)), 14, CONCATENATE(\"C3\", \"-4\"))",
            "=-21",
            "406.135",
            "\"KIioktqr\"",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 1));
    println!("Seed 699215 target: {:?}", target);
}

#[test]
fn test_fuzz_power_int_negative_exponent() {
    let sheet_src = [
        ["", "TRUE", "-28", "FALSE", "22"],
        ["-13", "", "289.98", "41", "72"],
        ["-78", "\"b3\"", "-116.6164", "-92", "\"Pzpq\""],
        ["\"GYIJf\"", "6", "69", "-39", "-114.2545"],
        ["84", "-99", "-96", "-30", ""],
        ["=B2", "=D1", "1", "=D5", "=-35"],
        ["=(B1 + D5)", "=A2", "=MIN(C3:E4)", "=C4", "-55"],
        [
            "=UPPER(\"SUM(B5:B5)\")",
            "=ROUNDDOWN(B3, 0)",
            "=-38",
            "=(-25 - ROUNDUP(E7, 1))",
            "=-37",
        ],
        ["=B7", "=D6", "=(17 + C1)", "=(B4 ^ INT(E8))", "TRUE"],
        [
            "=SUM(A5, IF((A5 > B7), 35, C3))",
            "=E6",
            "=AVERAGE(D8:D9)",
            "=37",
            "-68",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 3));
    println!("Seed 32959 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert!((f - 1.615860020532192e-29).abs() < 1e-35),
        ResultData::Integer(i) => assert_eq!(i, 0),
        other => panic!("Expected float, got {:?}", other),
    }
}

#[test]
fn test_fuzz_abs_c8_b10() {
    let sheet_src = [
        ["\"YxtpPvPu\"", "\"oI\"", "TRUE", "-28", "-62"],
        ["-128.04", "", "\"lugMgk\"", "24.45", ""],
        ["TRUE", "39", "13", "", "26"],
        ["14", "4", "FALSE", "100", "-250.1"],
        ["TRUE", "", "\"yWVnZrB\"", "-77", "167.4"],
        [
            "=(UPPER(\"3\") - E2)",
            "=AVERAGE(E3:E5)",
            "=ABS(ROUNDUP(B2, 1))",
            "=LEFT(\"C4\", 5)",
            "=OR(UPPER(\"50\") > 0, C5 < 100)",
        ],
        ["=PRODUCT(A5:D6)", "268.8", "=A4", "=D5", "=38"],
        [
            "=MIN(B4:E4)",
            "=IF((E2 > -45), (C2 / 17), (C4 + 17))",
            "=C7",
            "=26",
            "=IF((C6 > AND(A6 > 0, B7 < 100)), CONCATENATE(\"19\", \"A4\"), PRODUCT(C2, C4))",
        ],
        [
            "=SUM(ROUND(26, 1), (C1 - D2))",
            "=OR(IF((-21 > D8), A7, A6) > 0, ROUNDDOWN(-6, 2) < 100)",
            "=C2",
            "=D1",
            "=AVERAGE(IF((E8 > B5), D5, 8), (B6 / D4))",
        ],
        ["=C8", "=ABS(C8)", "=UPPER(\"B8\")", "=26", "=C3"],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 1));
    println!("Seed 640931 target: {:?}", target);
}

#[test]
fn test_fuzz_rounddown_d7_empty_cell_ref() {
    let sheet_src = [
        ["\"DTopPER\"", "-404.1953", "-22", "13", "6"],
        ["6", "\"VMqL\"", "\"S QCn2gY\"", "458.7529", ""],
        ["55", "-355.3733", "\"b\"", "\"yUpOMSRU\"", "4"],
        ["-35", "-369.5703", "162", "43", "268.344"],
        ["TRUE", "-77", "-63", "27", "-365"],
        ["", "=C4", "=1", "=ROUND(B2, 0)", "=INT(RIGHT(\"C5\", 4))"],
        ["=E5", "=PRODUCT(B5:B5)", "", "=A6", "=D1"],
        [
            "0",
            "=ROUND(PRODUCT(C4:C7), 1)",
            "=ROUNDUP(A2, 2)",
            "=E7",
            "=A4",
        ],
        [
            "=A2",
            "=ROUNDDOWN(D7, 2)",
            "=OR(ROUND(A2, 2) > 0, OR(37 > 0, D2 < 100) < 100)",
            "=SQRT((D2 / B3))",
            "40",
        ],
        [
            "=D6",
            "=B5",
            "=MIN(OR(E3 > 0, B8 < 100), (C9 - B9))",
            "\"zXLgd\"",
            "=IF((INT(D4) > E6), 35, -36)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 1));
    println!("Seed 332459 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 0.0),
        ResultData::Integer(i) => assert_eq!(i, 0),
        other => panic!("Expected 0, got {:?}", other),
    }
}

#[test]
fn test_fuzz_round_e6_empty_cell_ref() {
    let sheet_src = [
        ["492.3757", "-9", "", "\"QAQK\"", "\"bYw3RAT\""],
        ["", "-4", "5", "10", "TRUE"],
        ["\"E\"", "64", "FALSE", "TRUE", "177.1"],
        ["FALSE", "", "\"MkuQBum\"", "-49", "4"],
        ["91", "-26.531", "-94", "62", "-91"],
        [
            "=CONCATENATE(\"OR(0 > 0, D5 < 100)\", \"(0 + A4)\")",
            "=OR(AND(E4 > 0, B1 < 100) > 0, D3 < 100)",
            "=IF((B5 > IF((E2 > 40), B2, E4)), A1, 49)",
            "=AVERAGE((1 + -7), 44)",
            "=ROUNDDOWN(B4, 2)",
        ],
        [
            "=AVERAGE(A4:A5)",
            "-83",
            "=B4",
            "=11",
            "=MAX(MIN(C4:C4), (E2 / A6))",
        ],
        ["=ROUND(E6, 2)", "TRUE", "91", "=SQRT((B5 + 7))", "=D5"],
        [
            "=(IF((E8 > -50), C5, 1) - RIGHT(\"29\", 3))",
            "=IF((-23 > -3), 40, 50)",
            "=((D6 - 12) * A7)",
            "=UPPER(\"AND(-4 > 0, -46 < 100)\")",
            "=AND(8 > 0, B2 < 100)",
        ],
        ["", "=0", "=C2", "=C7", "=A3"],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 0));
    println!("Seed 325598 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 0.0),
        ResultData::Integer(i) => assert_eq!(i, 0),
        other => panic!("Expected 0, got {:?}", other),
    }
}

#[test]
fn test_fuzz_c2_multiplication_rounddown_abs() {
    let sheet_src = [
        ["-2", "\"FoKzo\"", "12", "\"p\"", "-57"],
        ["473.86", "-89", "-4", "-49", "214.8"],
        ["-16", "", "\"QOu2I\"", "214.28", "\"HCM\""],
        ["48", "-84", "-72", "\"o\"", "83"],
        ["9", "10", "103.2", "37", "-76"],
        ["FALSE", "=B4", "=36", "=SQRT(MIN(B5, -43))", "=ABS(2)"],
        ["=C4", "=C5", "=OR(A6 > 0, D1 < 100)", "-236.395", "=E5"],
        [
            "=((B4 - -4) / C3)",
            "=ROUNDUP(UPPER(\"36\"), 2)",
            "=-1",
            "=15",
            "=9",
        ],
        [
            "-90",
            "=(LEFT(\"8\", 1) ^ SQRT(E4))",
            "=(C2 * ROUNDDOWN(E6, 2))",
            "=IF((-24 > ROUNDUP(A7, 0)), (23 ^ E6), OR(A3 > 0, D7 < 100))",
            "=IF((D5 > AVERAGE(-18, C4)), (C5 - D1), A7)",
        ],
        [
            "=D3",
            "=IF((C8 > B4), OR(C9 > 0, -38 < 100), MAX(E8:E8))",
            "=39",
            "=ROUNDUP((D2 * 12), 2)",
            "-448",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 2));
    println!("Seed 667784 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, -8.0),
        ResultData::Integer(i) => assert_eq!(i, -8),
        other => panic!("Expected -8, got {:?}", other),
    }
}

#[test]
fn test_fuzz_int_negative_constant_cell_ref() {
    let sheet_src = [
        ["72", "FALSE", "1", "-31", "52"],
        ["93", "TRUE", "FALSE", "FALSE", "15"],
        ["-33", "3.66", "-42", "75", "41"],
        ["21", "", "38", "\"2bUngO2\"", "131.699"],
        ["56", "TRUE", "6", "", "58"],
        [
            "=AVERAGE((E1 - A1), AND(A3 > 0, -9 < 100))",
            "=5",
            "=(10 + SUM(E4, 9))",
            "=LEFT(\"-28\", 3)",
            "TRUE",
        ],
        [
            "=UPPER(\"IF((E1 > B5), C1, 27)\")",
            "=MAX(OR(C3 > 0, A1 < 100), MAX(D3:E3))",
            "=(LOWER(\"A2\") - C3)",
            "137.3",
            "=C3",
        ],
        [
            "=AND(E5 > 0, -39 < 100)",
            "=-33",
            "314.9",
            "=AND(ABS(-38) > 0, A3 < 100)",
            "-33",
        ],
        [
            "=38",
            "=LEFT(\"OR(-16 > 0, 24 < 100)\", 3)",
            "=D7",
            "=MAX(ROUND(-20, 1), A5)",
            "=A3",
        ],
        [
            "=INT(E9)",
            "=(ROUND(13, 1) + ROUND(E1, 1))",
            "=MIN(C8, ROUND(E4, 1))",
            "=D1",
            "=CONCATENATE(\"(C2 * -12)\", \"IF((B8 > C1), B8, E4)\")",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    println!("Seed 947857 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, -33.0),
        ResultData::Integer(i) => assert_eq!(i, -33),
        other => panic!("Expected -33, got {:?}", other),
    }
}

#[test]
fn test_fuzz_int_positive_constant_cell_ref() {
    let sheet_src = [
        ["60", "472.3", "10", "-43", "-75"],
        ["-97", "437.62", "-11", "-51", "FALSE"],
        ["34", "-287", "\"32Ouh\"", "64", "\"DscZws\""],
        ["", "-480", "\"RDAR\"", "FALSE", "TRUE"],
        ["-50", "-9", "3", "\"S\"", "TRUE"],
        [
            "=E4",
            "=-11",
            "=(LEN(\"B2\") * 35)",
            "=MIN(IF((C4 > D2), E2, B4), (-9 ^ 1))",
            "=IF((LEN(\"A3\") > (D4 / D2)), 4, D3)",
        ],
        [
            "=IF((B5 > C2), 3, 23)",
            "-38.569",
            "=(AVERAGE(B1, 44) ^ A6)",
            "-68",
            "110.243",
        ],
        [
            "=30",
            "-7.66",
            "\"wnYkoMY\"",
            "=IF((B5 > LEFT(\"-50\", 1)), LOWER(\"34\"), E6)",
            "=ROUNDUP(ROUNDUP(C1, 0), 0)",
        ],
        [
            "=(MIN(-17, E4) * B4)",
            "=B8",
            "=(-27 + (D4 - 15))",
            "=INT(E8)",
            "=(D6 * A2)",
        ],
        [
            "\"L2\"",
            "=A7",
            "=INT(E3)",
            "=LEN(\"AND(A9 > 0, A8 < 100)\")",
            "=LEFT(\"D5\", 1)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 3));
    println!("Seed 284690 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 10.0),
        ResultData::Integer(i) => assert_eq!(i, 10),
        other => panic!("Expected 10, got {:?}", other),
    }
}

#[test]
fn test_fuzz_round_e6_empty_ref() {
    let sheet_src = [
        ["\"pXoVoGJ\"", "317.3", "\"3LXfHOa\"", "86", ""],
        ["-349", "\"hvR\"", "23", "\"Dgs\"", "\"Iamx\""],
        ["89", "31", "-338.133", "41", "257.4388"],
        ["36", "TRUE", "\"CX fGNIK\"", "FALSE", "31"],
        ["-440", "-1", "", "58", "93"],
        ["=LEN(\"(0 - A2)\")", "=MIN(A1:C3)", "=C4", "TRUE", "=C5"],
        ["\"Ro\"", "=ROUND(E6, 2)", "=30", "-83", "=SUM(A4:D6)"],
        ["=(C6 * -25)", "=A7", "=2", "=(-5 * -10)", "48"],
        [
            "=SQRT(IF((B5 > -35), -32, B1))",
            "=C7",
            "=IF((RIGHT(\"B6\", 3) > (C5 - E6)), D6, C2)",
            "=4",
            "=INT(PRODUCT(-30, 27))",
        ],
        ["=C6", "-26", "=B5", "=(E9 * (C5 - B2))", "=LEFT(\"A1\", 4)"],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 1));
    println!("Seed 483010 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 0.0),
        ResultData::Integer(i) => assert_eq!(i, 0),
        other => panic!("Expected 0, got {:?}", other),
    }
}

#[test]
fn test_fuzz_roundup_e8_val() {
    let sheet_src = [
        ["89", "6", "-13", "-26", "-5"],
        ["3", "267.908", "\"HrVDbJ\"", "-100", "\"j rEDd\""],
        ["", "-185.2262", "\"XeIdn\"", "\"i\"", "54"],
        ["FALSE", "\"AwTAJ S\"", "55", "-8", "-3"],
        ["\"ADUZGxg\"", "72", "-76", "21", "136"],
        [
            "=28",
            "=D3",
            "=AVERAGE(D1:E4)",
            "=A2",
            "=(UPPER(\"D3\") / LEN(\"-50\"))",
        ],
        [
            "=(E4 * AND(17 > 0, 12 < 100))",
            "=LEFT(\"B3\", 5)",
            "=AVERAGE(D2:E5)",
            "=CONCATENATE(\"ROUNDUP(D3, 0)\", \"(E6 * -21)\")",
            "=C6",
        ],
        [
            "=A7",
            "=ROUNDDOWN(LEN(\"-17\"), 1)",
            "=C5",
            "=LEN(\"IF((C4 > C4), E6, 31)\")",
            "=IF((C1 > (B3 - D2)), LOWER(\"21\"), AVERAGE(A4:B4))",
        ],
        [
            "=B5",
            "=ROUNDUP(E8, 2)",
            "=PRODUCT(44, SUM(E3:E7))",
            "=AND(OR(B3 > 0, E1 < 100) > 0, C8 < 100)",
            "=AND(AND(C3 > 0, C5 < 100) > 0, LEN(\"7\") < 100)",
        ],
        [
            "-46",
            "=D9",
            "=C2",
            "=IF((LEFT(\"-40\", 4) > IF((E3 > D3), A4, C1)), RIGHT(\"-25\", 5), -43)",
            "=B1",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 1));
    println!("Seed 507073 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 21.0),
        ResultData::Integer(i) => assert_eq!(i, 21),
        other => panic!("Expected 21, got {:?}", other),
    }
}

#[test]
fn test_fuzz_rounddown_e6_val() {
    let sheet_src = [
        ["TRUE", "\"n2J\"", "-438.096", "-100", "5"],
        ["0", "22", "8", "", "161.8898"],
        ["\"rfVLxYq\"", "-384.8124", "FALSE", "", "TRUE"],
        ["", "43", "0", "351.73", "-23"],
        ["69", "73", "-260.7", "-77", "-397.07"],
        [
            "=UPPER(\"AVERAGE(19, D2)\")",
            "=26",
            "-52",
            "=(C3 - E2)",
            "=C4",
        ],
        [
            "=C1",
            "=IF(((-39 / -13) > (30 - C6)), (A6 ^ -24), D1)",
            "=(RIGHT(\"A1\", 2) * C3)",
            "=2",
            "=LEN(\"A4\")",
        ],
        [
            "=CONCATENATE(\"AND(A4 > 0, A3 < 100)\", \"PRODUCT(E1, D5)\")",
            "=C2",
            "=ROUNDDOWN(E6, 1)",
            "3",
            "=LEFT(\"IF((-42 > B2), -19, 19)\", 2)",
        ],
        [
            "=D2",
            "=PRODUCT(E8:E8)",
            "=16",
            "=(B2 * RIGHT(\"A3\", 5))",
            "=ROUND(E5, 0)",
        ],
        [
            "=B9",
            "=E7",
            "=LEFT(\"A2\", 3)",
            "=(-30 * C1)",
            "=MIN(E2:E5)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 2));
    println!("Seed 517687 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 0.0),
        ResultData::Integer(i) => assert_eq!(i, 0),
        other => panic!("Expected 0, got {:?}", other),
    }
}

#[test]
fn test_fuzz_abs_c6_val() {
    let sheet_src = [
        ["-21", "\"VEFaOfm\"", "", "-85", "-60"],
        ["0", "TRUE", "-238", "495.7921", "14"],
        ["-22", "63", "FALSE", "\"x\"", "\"MSikQ\""],
        ["19", "TRUE", "\"Wx\"", "8", "TRUE"],
        ["\"bTU\"", "51", "2", "TRUE", "8"],
        [
            "8",
            "=D3",
            "=OR(LOWER(\"E1\") > 0, SQRT(B5) < 100)",
            "=E3",
            "=D2",
        ],
        [
            "=IF((B5 > C6), SQRT(D5), ROUNDUP(-14, 1))",
            "=E3",
            "=AND(RIGHT(\"D3\", 5) > 0, ROUNDUP(B4, 0) < 100)",
            "=AND(UPPER(\"15\") > 0, IF((25 > D6), C2, -43) < 100)",
            "0",
        ],
        [
            "=C6",
            "=ABS(C6)",
            "=C7",
            "=-15",
            "=IF((-33 > ROUNDUP(-8, 2)), (B4 / C2), LOWER(\"C3\"))",
        ],
        [
            "=(OR(A4 > 0, D7 < 100) - CONCATENATE(\"E8\", \"D8\"))",
            "=B6",
            "=SQRT(MIN(E6:E7))",
            "=B7",
            "=LEFT(\"ROUNDDOWN(33, 2)\", 2)",
        ],
        [
            "=(LEN(\"16\") * MAX(-46, A8))",
            "=PRODUCT(B2:B5)",
            "=LEFT(\"(B1 * 31)\", 5)",
            "=AVERAGE(-50, A3)",
            "76",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 1));
    println!("Seed 159433 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 1.0),
        ResultData::Integer(i) => assert_eq!(i, 1),
        other => panic!("Expected 1, got {:?}", other),
    }
}

#[test]
fn test_fuzz_round_addition_val() {
    let sheet_src = [
        ["-464", "", "-69", "-89", "272.45"],
        ["83", "28", "-84", "TRUE", "TRUE"],
        ["-358.99", "\"3L\"", "81", "-45.6497", "-17"],
        ["1", "-64", "-20.5752", "0", "67"],
        ["\"sc \"", "75", "", "7", "\"bgQrkP2\""],
        [
            "=A5",
            "=SUM(A3:D5)",
            "=OR(UPPER(\"C5\") > 0, (D3 - A2) < 100)",
            "=MIN(C3:E5)",
            "=PRODUCT(ROUND(D2, 1), AVERAGE(A2:A5))",
        ],
        [
            "=(ROUNDUP(D4, 0) ^ B6)",
            "=ROUNDUP((18 ^ B2), 1)",
            "=C3",
            "7",
            "FALSE",
        ],
        [
            "=LEFT(\"OR(C2 > 0, -49 < 100)\", 3)",
            "=(-21 - ABS(C3))",
            "-57",
            "=D4",
            "=MAX(D1:D3)",
        ],
        [
            "=A6",
            "201.3376",
            "=ROUND(LEFT(\"-47\", 1), 1)",
            "115.5",
            "=B4",
        ],
        [
            "=(ROUND(C6, 2) + D4)",
            "=ROUNDUP(B9, 2)",
            "FALSE",
            "=ROUNDUP(9, 0)",
            "=AVERAGE(E1:E2)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    println!("Seed 61467 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 1.0),
        ResultData::Integer(i) => assert_eq!(i, 1),
        other => panic!("Expected 1, got {:?}", other),
    }
}

#[test]
fn test_fuzz_round_cell_ref_float_val() {
    let sheet_src = [
        ["FALSE", "-481.188", "459.798", "-31", "\"qXxUInB\""],
        ["\"G\"", "80.01000000000001", "\"LDzgIKy\"", "-11", ""],
        ["FALSE", "-98", "60", "-1", "-12"],
        ["\"BjZ\"", "FALSE", "-23.5", "69", "70"],
        ["\"Vy\"", "\"H\"", "-76", "-56.95", "TRUE"],
        [
            "=IF((D2 > (-16 + C3)), C4, C4)",
            "\"2Lb\"",
            "=B5",
            "=((C3 + 28) + IF((B2 > 40), B1, 17))",
            "=IF((LOWER(\"B4\") > E3), 6, C3)",
        ],
        ["=26", "=D3", "=C2", "=D4", "=C1"],
        [
            "=ROUNDUP(LEN(\"B2\"), 0)",
            "=AND((A1 * E3) > 0, E7 < 100)",
            "=LEN(\"A3\")",
            "=ROUNDUP(D3, 2)",
            "-60",
        ],
        [
            "=CONCATENATE(\"AVERAGE(C2, B7)\", \"AVERAGE(A2, -7)\")",
            "=IF((LEN(\"E8\") > E2), B5, RIGHT(\"C5\", 5))",
            "=(14 - MIN(B4, A8))",
            "=ROUND(E7, 0)",
            "=21",
        ],
        [
            "=(C2 * B9)",
            "=AND(LEFT(\"C8\", 5) > 0, (A2 / C9) < 100)",
            "=(A2 * 10)",
            "=B9",
            "=PRODUCT(ABS(B8), ROUNDDOWN(D2, 0))",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 3));
    println!("Seed 976519 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 460.0),
        ResultData::Integer(i) => assert_eq!(i, 460),
        other => panic!("Expected 460, got {:?}", other),
    }
}

#[test]
fn test_fuzz_power_roundup_rounddown_val() {
    let sheet_src = [
        ["-12.1", "-22", "34", "58", ""],
        ["-17", "98", "\"ssCuwts\"", "-175.8", "-135.4"],
        ["-150.079", "45", "115.15", "-85", "40"],
        ["7", "-67", "-19", "-48", "0"],
        ["49", "\"hB\"", "-34", "-3", "-20"],
        [
            "=PRODUCT(-50, E4)",
            "=((18 * 18) - IF((-1 > 46), E1, B5))",
            "=PRODUCT((A5 * A3), PRODUCT(-43, C5))",
            "=IF((IF((B5 > D1), B1, E2) > B3), MIN(A2:A5), E1)",
            "=SQRT(OR(E3 > 0, D3 < 100))",
        ],
        [
            "331.33",
            "=-3",
            "=SQRT(B3)",
            "=LEN(\"OR(A6 > 0, 8 < 100)\")",
            "=B1",
        ],
        [
            "=((E4 - A1) - D2)",
            "=RIGHT(\"9\", 1)",
            "-60",
            "=-14",
            "=ROUNDDOWN(ROUND(B1, 1), 1)",
        ],
        [
            "=A4",
            "=AVERAGE(IF((B1 > C1), D5, A4), D6)",
            "=14",
            "=LOWER(\"-44\")",
            "=LOWER(\"13\")",
        ],
        [
            "=(ROUNDUP(C9, 1) ^ ROUNDDOWN(17, 0))",
            "=B8",
            "=AND(IF((D7 > 42), E3, D8) > 0, (D8 ^ B6) < 100)",
            "=IF((E8 > ROUNDDOWN(A4, 2)), 43, E3)",
            "=E7",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    println!("Seed 158317 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert!((f - 30491346729331195904.0).abs() < 1e10),
        ResultData::Integer(i) => assert_eq!(i, 30491346729331195904_i128 as i64),
        other => panic!("Expected 30491346729331195904, got {:?}", other),
    }
}

#[test]
fn test_fuzz_rounddown_unevaluated_formula_dependency() {
    let sheet_src = [
        ["474.1487", "-125.33", "", "-1", "4"],
        ["85", "3", "-388.17", "-310.2", "-269"],
        ["-76", "", "-96", "-94", "127.986"],
        ["10", "-20", "86", "30", "-72"],
        ["76", "-37", "TRUE", "88", ""],
        [
            "=B2",
            "75",
            "=C4",
            "=PRODUCT(MIN(E4, D4), IF((E1 > 44), D2, -18))",
            "=IF((A3 > A1), OR(-17 > 0, B5 < 100), IF((4 > C2), C5, D1))",
        ],
        ["83", "=D4", "=E4", "-34", "=C2"],
        [
            "=IF((SQRT(D4) > PRODUCT(C1, A4)), (E1 * E1), C7)",
            "=MAX(MAX(A7:D7), C2)",
            "=ROUNDUP(OR(A6 > 0, D7 < 100), 1)",
            "=MAX(C4:D7)",
            "33",
        ],
        [
            "=A3",
            "=ROUNDDOWN(D6, 0)",
            "=((B2 ^ C2) + (-44 - B1))",
            "=SUM(OR(31 > 0, B3 < 100), ROUNDUP(B3, 2))",
            "=((C8 * -38) ^ -45)",
        ],
        [
            "=C4",
            "=(ROUND(18, 0) * D5)",
            "=E7",
            "21",
            "=((B4 * 32) * (E7 / A3))",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 1));
    match target {
        ResultData::Float(f) => assert_eq!(f, 1296.0),
        ResultData::Integer(i) => assert_eq!(i, 1296),
        other => panic!("Expected Float(1296.0), got {:?}", other),
    }
}

#[test]
fn test_scientific_notation_respects_the_twenty_character_budget() {
    use crate::core::engine::result_data::format_excel_number;

    // Excel's display budget is 20 characters, and in scientific notation
    // the exponent is charged against it -- a three-digit exponent leaves
    // one fewer mantissa digit than a two-digit one. Both of these are
    // exactly 20 characters wide, and both are real Excel's rendering.
    assert_eq!(
        format_excel_number(2.277577478736661e-171),
        "2.2775774787367E-171"
    );
    assert_eq!(format_excel_number(1.0 / 3e20), "3.33333333333333E-21");

    // The mantissa used to be fixed at 15 significant digits regardless of
    // the exponent width, which made the first of these 21 characters.
    assert!(format_excel_number(2.277577478736661e-171).len() <= 20);
    assert!(format_excel_number(-2.277577478736661e-171).len() <= 21); // plus the sign
}

#[test]
fn test_number_to_text_keeps_only_excels_fifteen_significant_digits() {
    use crate::core::engine::result_data::format_excel_number;

    // 43^11 is 929293739471223048 as an f64; Excel shows 15 significant
    // digits and zeroes the rest. The old formatter emitted the f64's own
    // digits here, which leaked precision Excel never displays.
    assert_eq!(format_excel_number(43f64.powi(11)), "929293739471223000");
    assert_eq!(
        format_excel_number(-(43f64.powi(11))),
        "-929293739471223000"
    );

    // The sign is not charged against the 20-character budget: real Excel
    // writes this at a full 15 significant digits despite the minus.
    assert_eq!(
        format_excel_number(-2.05237592634038e-10),
        "-2.05237592634038E-10"
    );

    // Unchanged neighbours.
    assert_eq!(format_excel_number(1e19), "10000000000000000000");
    assert_eq!(format_excel_number(1e20), "1E+20");
    assert_eq!(format_excel_number(976121418126.432), "976121418126.432");
    assert_eq!(format_excel_number(-1.0 / 3.0), "-0.333333333333333");
    assert_eq!(
        format_excel_number(std::f64::consts::PI),
        "3.14159265358979"
    );
}

#[test]
fn test_scientific_notation_rounds_from_the_fifteen_digit_value() {
    use crate::core::engine::result_data::format_excel_number;

    // Excel snaps a result to 15 significant digits and only then formats
    // it, so when a three-digit exponent leaves room for just 14 the two
    // roundings compose. 28^-92 is 7.26877317134744769...e-134: rounding
    // that straight to 14 digits gives ...7474, but snapping to 15 first
    // gives 7.26877317134745 and then 14 gives ...7475 -- which is what
    // Excel prints.
    assert_eq!(format_excel_number(28f64.powi(-92)), "7.2687731713475E-134");
    // The neighbours this could have disturbed, all real Excel values.
    assert_eq!(
        format_excel_number(2.277577478736661e-171),
        "2.2775774787367E-171"
    );
    assert_eq!(
        format_excel_number(-2.05237592634038e-10),
        "-2.05237592634038E-10"
    );
    assert_eq!(format_excel_number(1.0 / 3e20), "3.33333333333333E-21");
    assert_eq!(format_excel_number(1e20), "1E+20");
}
