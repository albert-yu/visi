use super::*;

#[test]
fn test_fuzz_average_range_with_empty_cells() {
    let sheet_src = [
        ["23", "\"qefv\"", "97", "-11", "-33"],
        ["148.2045", "45", "-272.57", "0", "-97"],
        ["\"wiKVP\"", "-58", "FALSE", "\"3HdF\"", "-54"],
        ["\"iltsnHSB\"", "", "\"m\"", "-254.7", ""],
        ["398.5356", "45", "76", "-77", "\"uTL\""],
        [
            "=(CONCATENATE(\"D1\", \"E2\") ^ A3)",
            "=-31",
            "=-7",
            "=D4",
            "=UPPER(\"OR(-38 > 0, 43 < 100)\")",
        ],
        ["149.9", "-4", "\"2ambJCe\"", "=C3", "=MIN(C2:D4)"],
        [
            "=ROUNDDOWN(MAX(C2:C2), 0)",
            "=B6",
            "=CONCATENATE(\"48\", \"C7\")",
            "62",
            "\"Ti\"",
        ],
        [
            "=-46",
            "-30",
            "=22",
            "=E4",
            "=AND(E4 > 0, OR(D5 > 0, E5 < 100) < 100)",
        ],
        [
            "=SUM(ABS(E9), SUM(A2, -9))",
            "=1",
            "=LOWER(\"D8\")",
            "=IF(((C9 / D9) > C4), AND(18 > 0, C3 < 100), SQRT(D7))",
            "=D5",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    match target {
        ResultData::Float(f) => assert!(
            (f - 139.2045).abs() < 1e-3,
            "Expected ~139.2045 for A10, got {}",
            f
        ),
        other => panic!("Expected Float(~139.2045) for A10, got {:?}", other),
    }
}

#[test]
fn test_fuzz_division_by_negative_sum() {
    let sheet_src = [
        ["\"jarupzx\"", "\"WL\"", "-307.77", "TRUE", "39"],
        ["FALSE", "\"eTB\"", "-31", "92", "-54"],
        ["TRUE", "65", "FALSE", "357", "FALSE"],
        ["-65", "76", "-58", "FALSE", "\"oT\""],
        ["275.93", "434.801", "TRUE", "440.643", "FALSE"],
        [
            "=PRODUCT(ROUND(A1, 1), SQRT(C5))",
            "=A1",
            "=D4",
            "=OR(AVERAGE(E2:E4) > 0, D2 < 100)",
            "=E3",
        ],
        [
            "=ROUNDDOWN(OR(48 > 0, B4 < 100), 2)",
            "=-49",
            "=LEN(\"(33 + E3)\")",
            "=-28",
            "=(-17 + D3)",
        ],
        [
            "=ABS(B7)",
            "=ROUNDDOWN(LOWER(\"B2\"), 1)",
            "=CONCATENATE(\"A3\", \"AND(D1 > 0, 32 < 100)\")",
            "=UPPER(\"ROUND(E1, 2)\")",
            "=B7",
        ],
        [
            "=A7",
            "=(A7 + IF((E8 > -34), A5, D4))",
            "=RIGHT(\"E1\", 5)",
            "=CONCATENATE(\"E4\", \"D6\")",
            "=LEN(\"C5\")",
        ],
        [
            "-26",
            "=UPPER(\"-42\")",
            "=LEFT(\"ROUND(B6, 0)\", 5)",
            "\"GCdw2QD\"",
            "=(-41 / SUM(D5, -7))",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 0));
    match target {
        ResultData::Float(f) => assert_eq!(f, 49.0),
        other => panic!("Expected Float(49.0) for A8, got {:?}", other),
    }
}

#[test]
fn test_fuzz_product_nested_float_precision() {
    let sheet_src = [
        ["", "-83", "", "-3", "-290.5375"],
        ["95", "60.005", "\"KQdrYUl\"", "-13", "-70"],
        ["-437.94", "FALSE", "55", "84", "TRUE"],
        ["-181.7", "-68", "\"3z\"", "-332.5255", "-27"],
        ["TRUE", "100", "-14.34", "TRUE", "-96"],
        ["=-43", "=OR(6 > 0, SUM(E3, -4) < 100)", "=D1", "=A1", "=E2"],
        [
            "=SQRT(LEFT(\"C2\", 5))",
            "57",
            "=B6",
            "=ROUNDDOWN(OR(13 > 0, A4 < 100), 2)",
            "=-46",
        ],
        [
            "=A1",
            "=AVERAGE(B5:D6)",
            "=IF((D5 > C3), B2, E3)",
            "303.5083",
            "=C1",
        ],
        [
            "=LEFT(\"AND(D6 > 0, 16 < 100)\", 4)",
            "=E7",
            "=A5",
            "79",
            "=15",
        ],
        [
            "=RIGHT(\"C6\", 1)",
            "=C3",
            "=(ROUNDUP(D9, 1) + LEN(\"42\"))",
            "=(D5 + LEN(\"14\"))",
            "=PRODUCT(E9:E9)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 1));
    match target {
        ResultData::Float(f) => assert!(
            (f - 20.665).abs() < 1e-3,
            "Expected ~20.665 for B8, got {}",
            f
        ),
        other => panic!("Expected Float(~27.5533) for B8, got {:?}", other),
    }
}

#[test]
fn test_fuzz_range_min_max_evaluation() {
    let sheet_src = [
        ["-96", "37", "-118.7", "3", "TRUE"],
        ["491.0698", "-21", "", "-95", "-6"],
        ["-89", "\"OhDtA\"", "-61", "", "-10"],
        ["318.776", "", "369.33", "\"AS\"", "90.20699999999999"],
        ["-75", "TRUE", "", "", "31"],
        [
            "=D4",
            "=PRODUCT(OR(-39 > 0, E1 < 100), A1)",
            "=ROUNDDOWN(AND(D4 > 0, -30 < 100), 0)",
            "=B4",
            "=D4",
        ],
        [
            "\"hPo\"",
            "=SUM(B2:B2)",
            "=-10",
            "=ABS((B5 * 20))",
            "=ROUND(IF((E5 > E1), 25, A1), 1)",
        ],
        [
            "=MAX(SQRT(4), D7)",
            "=LEN(\"-36\")",
            "=-25",
            "=C7",
            "=(RIGHT(\"50\", 3) + (B5 - B6))",
        ],
        [
            "=-40",
            "=MAX(C4, LEN(\"C5\"))",
            "=(ABS(B7) - -47)",
            "=AVERAGE(A7, A7)",
            "=PRODUCT(B5, (C6 * D7))",
        ],
        [
            "=AVERAGE(ROUNDDOWN(E7, 1), AND(E9 > 0, 29 < 100))",
            "=(ROUND(E5, 2) / C1)",
            "=(MAX(B2:B8) / D4)",
            "=D1",
            "=(LEFT(\"C9\", 2) - -47)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 4));
    match target {
        ResultData::Float(f) => assert_eq!(f, 51.0),
        other => panic!("Expected Float(51.0) for E8, got {:?}", other),
    }
}

#[test]
fn test_fuzz_average_nested_max_precision() {
    let sheet_src = [
        ["-45", "FALSE", "-43", "67", "\"bnoLpNG2\""],
        ["48", "82", "-46", "TRUE", "491.9317"],
        ["-19", "13", "\"yxyMfKYY\"", "453.7", "-321.9617"],
        ["", "FALSE", "\"edjWwod\"", "", "348.2"],
        ["8", "-154.7", "7", "", "\"egAUhXV\""],
        [
            "=AND(AVERAGE(B2:B4) > 0, ROUNDDOWN(D5, 0) < 100)",
            "=B4",
            "=ROUNDUP((E2 / 38), 1)",
            "=AND(A3 > 0, 9 < 100)",
            "=38",
        ],
        [
            "=(IF((C1 > -17), A3, E6) ^ D6)",
            "=C1",
            "=D5",
            "=D4",
            "=UPPER(\"14\")",
        ],
        [
            "=A5",
            "=MAX(C2, A6)",
            "=LEN(\"MIN(B2:D5)\")",
            "=MAX(A2:A3)",
            "=(ABS(C4) ^ D3)",
        ],
        [
            "=(SQRT(C4) ^ SUM(C1:E4))",
            "=AVERAGE(C7:D8)",
            "=OR(E7 > 0, 36 < 100)",
            "\"pAs\"",
            "=MAX(D7:D8)",
        ],
        [
            "=UPPER(\"ROUND(C5, 1)\")",
            "=LEN(\"3\")",
            "-51",
            "=E3",
            "=AND(28 > 0, MIN(E3, D3) < 100)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 0));
    match target {
        ResultData::Float(f) => assert_eq!(f, 1.0),
        other => panic!("Expected Float(1.0) for A7, got {:?}", other),
    }
}

#[test]
fn test_fuzz_product_negative_multipliers() {
    let sheet_src = [
        ["79.8169", "-2", "", "FALSE", "TRUE"],
        ["7", "7", "44", "TRUE", ""],
        ["\"Ump\"", "51", "-233", "77", "0"],
        ["", "42", "-9", "-40", "-226.72"],
        ["\"U\"", "10", "-295", "FALSE", "105.5542"],
        ["=24", "-76", "=ABS(-36)", "", "FALSE"],
        [
            "=D4",
            "=(PRODUCT(B3:C5) / 9)",
            "=ROUNDDOWN(E2, 2)",
            "=A1",
            "-20",
        ],
        ["=D4", "=-15", "=-34", "=(SUM(E3:E5) * (C5 - -36))", "=C2"],
        ["=C1", "=((B2 - 22) * AVERAGE(E1:E6))", "=A2", "=E2", "=E5"],
        [
            "=D4",
            "=(ROUNDUP(E9, 0) * MIN(B9:B9))",
            "=IF((AVERAGE(E5:E5) > MIN(E2, C2)), C8, C8)",
            "=(B5 * LOWER(\"E5\"))",
            "=ROUND(LOWER(\"-36\"), 1)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 1));
    match target {
        ResultData::Float(f) => assert!(
            (f - 64217.874).abs() < 1e-3,
            "Expected ~64217.874, got {}",
            f
        ),
        other => panic!("Expected Float for B10, got {:?}", other),
    }
}

#[test]
fn test_fuzz_subtraction_large_range_min() {
    let sheet_src = [
        ["13", "\"XUj\"", "-64", "", "452.8"],
        ["-383.7822", "-131.26", "79", "-38", "91"],
        ["", "0", "-282", "0", "TRUE"],
        ["\"WkkJpo\"", "FALSE", "118.478", "-46", "FALSE"],
        ["-65.7633", "-76", "62", "TRUE", "303.77"],
        ["=A2", "=UPPER(\"A3\")", "=C5", "=MAX(A4:C5)", "=C3"],
        ["\"yJH mhT\"", "=A3", "=(INT(E4) + E6)", "=B2", "=D6"],
        [
            "=(AND(-7 > 0, E4 < 100) * E3)",
            "=IF((UPPER(\"18\") > C4), IF((A6 > 26), E7, C2), ABS(-38))",
            "=AND(E5 > 0, (D1 + D2) < 100)",
            "=(ROUNDUP(D1, 1) / C3)",
            "=-10",
        ],
        [
            "=B1",
            "=(AVERAGE(B3, B2) * IF((13 > B7), -48, 2))",
            "=(AND(D6 > 0, E2 < 100) + (B1 * 24))",
            "=44",
            "=ABS(OR(-10 > 0, A2 < 100))",
        ],
        [
            "=LOWER(\"-6\")",
            "=OR(CONCATENATE(\"33\", \"C1\") > 0, MAX(C8:E8) < 100)",
            "=(PRODUCT(E2:E4) + C3)",
            "=MAX(D2:D9)",
            "3",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 2));
    match target {
        ResultData::Float(f) => assert_eq!(f, -282.0),
        other => panic!("Expected Float(-282.0) for C7, got {:?}", other),
    }
}

#[test]
fn test_fuzz_sum_range_negative_values() {
    let sheet_src = [
        ["-58", "40", "222", "\"3ZE\"", "-174.2544"],
        ["\"yrSbQs\"", "FALSE", "-270.4109", "76", "-71"],
        ["71", "-69", "\"Kej\"", "-450.002", ""],
        ["25", "77", "6", "\"W\"", "\"zAhnlQyo\""],
        ["", "", "\"cm2\"", "\"Hs\"", "5"],
        [
            "=PRODUCT(A2:D5)",
            "=((C4 - A5) - MIN(C4:D4))",
            "=CONCATENATE(\"ABS(B4)\", \"MAX(E1:E1)\")",
            "=A3",
            "=B5",
        ],
        [
            "=-45",
            "=MAX(20, (E3 + C5))",
            "=(ROUND(19, 2) - IF((C1 > 33), D6, A6))",
            "218.1489",
            "=-38",
        ],
        [
            "=B6",
            "=(SQRT(-48) * E2)",
            "=(ROUND(B5, 1) * D2)",
            "=A7",
            "=ROUND(ROUND(B2, 2), 1)",
        ],
        [
            "=B5",
            "=ROUNDUP(IF((E7 > C1), D1, A4), 0)",
            "494",
            "400.726",
            "=OR(-26 > 0, MAX(E8:E8) < 100)",
        ],
        [
            "=MAX(E1:E5)",
            "=OR(E7 > 0, (-39 * C8) < 100)",
            "=CONCATENATE(\"(E1 * D7)\", \"INT(E5)\")",
            "=CONCATENATE(\"MIN(C6:E8)\", \"30\")",
            "211.5",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 2));
    match target {
        ResultData::Float(f) => assert_eq!(f, -52.0),
        other => panic!("Expected Float(-52.0) for C7, got {:?}", other),
    }
}

#[test]
fn test_fuzz_sum_product_cell_references() {
    let sheet_src = [
        ["", "6", "TRUE", "-372.1", "-61"],
        ["64", "-79", "-15", "", "0"],
        ["\"CVZDGbB\"", "", "106.6", "25", ""],
        ["49", "-86", "56.984", "18", "424.2595"],
        ["-99", "-15", "-68", "-50", "5"],
        [
            "=ROUND(39, 1)",
            "=SUM(21, (E2 * A1))",
            "=LOWER(\"C4\")",
            "=D2",
            "=AND(A5 > 0, A3 < 100)",
        ],
        [
            "=D3",
            "=ROUND(C5, 0)",
            "=-36",
            "=ROUNDDOWN((A2 + A4), 2)",
            "=IF((29 > SUM(A5:A6)), B5, SUM(A5:C6))",
        ],
        [
            "=41",
            "=-21",
            "=A4",
            "=IF((C6 > SUM(B1:B7)), AND(D2 > 0, 38 < 100), INT(E5))",
            "=-9",
        ],
        [
            "=IF((MIN(E5:E6) > OR(C1 > 0, E1 < 100)), MAX(D8, D6), AVERAGE(C2:D8))",
            "=-10",
            "=E4",
            "\"WypeLQcC\"",
            "=IF((D8 > IF((E3 > C7), D5, -27)), LEN(\"C3\"), SUM(C5:D7))",
        ],
        [
            "9",
            "=IF((IF((21 > B6), -5, C3) > LEFT(\"-36\", 5)), PRODUCT(A7:B7), ROUNDUP(D7, 2))",
            "=B9",
            "401",
            "31.8",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 1));
    match target {
        ResultData::Float(f) => assert_eq!(f, 113.0),
        other => panic!("Expected Float(113.0) for B10, got {:?}", other),
    }
}

#[test]
fn test_fuzz_negative_integer_range_sum() {
    let sheet_src = [
        ["-58", "\"fij\"", "59", "-42.662", "\"AVg\""],
        ["-354.803", "-58", "", "FALSE", ""],
        ["TRUE", "-62", "\"qh\"", "\"yl\"", "0"],
        ["", "-27", "-29", "100", "104"],
        ["-39", "-220", "-86", "10", "-95"],
        [
            "=A5",
            "=C4",
            "=CONCATENATE(\"INT(-12)\", \"(A1 * A3)\")",
            "=1",
            "=(AND(B4 > 0, B2 < 100) * MAX(A4:B4))",
        ],
        [
            "",
            "1",
            "=((B5 - D3) ^ C6)",
            "=LEN(\"(-48 + D3)\")",
            "=SUM(D1:E5)",
        ],
        [
            "40",
            "=AVERAGE(B7:C7)",
            "=-45",
            "=ROUNDUP(AND(5 > 0, C5 < 100), 0)",
            "=IF((OR(C4 > 0, 30 < 100) > -45), A6, UPPER(\"C3\"))",
        ],
        [
            "=AND(OR(-1 > 0, A1 < 100) > 0, MIN(D6:E7) < 100)",
            "=PRODUCT(ROUND(B8, 1), 38)",
            "=B4",
            "",
            "=PRODUCT(E7:E8)",
        ],
        [
            "-62",
            "=37",
            "=AVERAGE(E1, B5)",
            "=INT(13)",
            "=(AND(-13 > 0, 22 < 100) * IF((36 > D2), A7, 21))",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 2));
    match target {
        ResultData::Float(f) => assert_eq!(f, -220.0),
        other => panic!("Expected Float(-220.0) for C10, got {:?}", other),
    }
}

#[test]
fn test_fuzz_nested_min_max_evaluation() {
    let sheet_src = [
        ["", "88", "", "-219.79", "291"],
        ["\"C\"", "5", "\"sXn\"", "-357.03", "TRUE"],
        ["-51", "-6", "-277.064", "", "\"L\""],
        ["77", "FALSE", "-50", "39", "-330.8664"],
        ["FALSE", "74", "300.7576", "FALSE", "7"],
        [
            "=LOWER(\"ROUNDDOWN(A4, 1)\")",
            "=(MAX(D3, E5) - (-41 * D4))",
            "=(A3 ^ A1)",
            "=A2",
            "=-8",
        ],
        [
            "=(C3 / B3)",
            "=-30",
            "=LEFT(\"AVERAGE(B6:C6)\", 4)",
            "=D3",
            "0",
        ],
        [
            "=(D5 - B1)",
            "=OR(ROUNDDOWN(-37, 0) > 0, MAX(A3:A7) < 100)",
            "=D5",
            "=ROUND((E5 / 15), 2)",
            "=OR((A5 ^ 46) > 0, LEFT(\"D6\", 1) < 100)",
        ],
        [
            "=E4",
            "=SQRT(UPPER(\"A6\"))",
            "=LOWER(\"47\")",
            "=OR(6 > 0, 24 < 100)",
            "=D5",
        ],
        ["=INT(E6)", "=B4", "=((D6 ^ D8) ^ E3)", "-28", "=-3"],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    match target {
        ResultData::Float(f) => assert_eq!(f, -8.0),
        other => panic!("Expected Float(-8.0) for A10, got {:?}", other),
    }
}

#[test]
fn test_fuzz_roundup_sum_expression() {
    let sheet_src = [
        ["0", "", "-41", "", "1"],
        ["", "-84", "55", "-93", "6"],
        ["-1", "436", "3", "-384.63", "4"],
        ["-26", "24", "", "475.29", "0"],
        ["-90.40000000000001", "0", "-15", "\"ArNOnb\"", "100"],
        ["=A4", "=E4", "=OR(2 > 0, A4 < 100)", "=B5", "=C4"],
        [
            "=LEFT(\"ROUNDDOWN(A4, 0)\", 2)",
            "=C3",
            "=ROUNDUP((E1 + D6), 0)",
            "=(E3 * 25)",
            "=AND(CONCATENATE(\"E2\", \"A2\") > 0, (E3 / 14) < 100)",
        ],
        [
            "=AND(-10 > 0, E3 < 100)",
            "=B3",
            "=-50",
            "=(AVERAGE(B5:C6) - -28)",
            "=28",
        ],
        [
            "=LOWER(\"(B3 ^ D1)\")",
            "=IF((UPPER(\"E4\") > RIGHT(\"C4\", 1)), C5, ROUNDDOWN(-10, 2))",
            "=SQRT(LEFT(\"A2\", 3))",
            "=LEFT(\"B6\", 5)",
            "=(19 - (-45 * 49))",
        ],
        [
            "=A3",
            "-344.119",
            "=D2",
            "=(INT(-21) ^ (C8 - D8))",
            "=AVERAGE(MIN(-47, B4), E2)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 2));
    match target {
        ResultData::Float(f) => assert_eq!(f, 1.0),
        other => panic!("Expected Float(1.0) for C7, got {:?}", other),
    }
}

#[test]
fn test_fuzz_max_int_multiplication() {
    let sheet_src = [
        ["-75", "FALSE", "85", "\"nL\"", "FALSE"],
        ["\"oroIgv\"", "\"shG\"", "21", "FALSE", "\"yqt1\""],
        ["", "14.0734", "416.3", "FALSE", "88"],
        ["TRUE", "-73.29000000000001", "3", "-69", "FALSE"],
        ["151.8448", "259.92", "-32", "95", "-350.02"],
        [
            "=MAX(A2:E2)",
            "=ABS(RIGHT(\"30\", 2))",
            "=RIGHT(\"B2\", 2)",
            "=(MAX(D4:D4) - C5)",
            "=6",
        ],
        [
            "=((C3 - C4) - C4)",
            "=INT(OR(C3 > 0, 48 < 100))",
            "=INT(LEN(\"E4\"))",
            "=INT(A2)",
            "=OR(E4 > 0, C6 < 100)",
        ],
        [
            "=D5",
            "=ROUND(SUM(B5:E6), 2)",
            "=D6",
            "=SQRT(IF((-41 > A6), D7, A3))",
            "-195.6794",
        ],
        ["-13", "=A4", "=MAX(INT(D3), (E5 * C8))", "=28", "249"],
        [
            "224.0393",
            "=AVERAGE(D4, UPPER(\"E4\"))",
            "=IF(((B7 + 3) > LEN(\"D2\")), LEN(\"-41\"), (C3 / -45))",
            "=ABS(D8)",
            "=AND(ABS(33) > 0, IF((D5 > 43), A1, -44) < 100)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 2));
    match target {
        ResultData::Float(f) => assert!(
            (f - 12950.74).abs() < 1e-2,
            "Expected ~12950.74 for C9, got {}",
            f
        ),
        other => panic!("Expected Float(~12950.74) for C9, got {:?}", other),
    }
}

#[test]
fn test_fuzz_power_min_expression() {
    let sheet_src = [
        ["88", "-275.4692", "", "", "\"p\""],
        ["-84", "-85", "17", "\"V2ETUs\"", "64"],
        ["100", "-177.9", "", "\"apdKJDV\"", "-359"],
        ["-34", "6", "-33", "", "TRUE"],
        ["-348.65", "", "92", "-96", "6"],
        [
            "=INT(D3)",
            "=(CONCATENATE(\"-19\", \"40\") * OR(C1 > 0, E1 < 100))",
            "\"KBdDN1kV\"",
            "=ROUNDUP(SQRT(-27), 2)",
            "=RIGHT(\"-17\", 4)",
        ],
        [
            "\"r3TN\"",
            "=AND(25 > 0, UPPER(\"-13\") < 100)",
            "=PRODUCT(A6, B5)",
            "89",
            "=AVERAGE(D1, AVERAGE(B3:C5))",
        ],
        [
            "=(MIN(E7, D4) ^ B6)",
            "=-50",
            "=(B4 * ROUNDDOWN(A1, 2))",
            "=UPPER(\"ABS(D1)\")",
            "=-39",
        ],
        [
            "=B8",
            "=MIN(D3:D3)",
            "96",
            "=AVERAGE(ROUND(-25, 1), IF((21 > B5), -46, E3))",
            "=SUM(A8:A8)",
        ],
        [
            "=B5",
            "=(ABS(C6) * INT(C4))",
            "\"q\"",
            "=INT((A9 * -46))",
            "\"NcJ\"",
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
fn test_fuzz_addition_min_negative() {
    let sheet_src = [
        ["0", "4", "\"RfgVD2\"", "\"oTxGNKnS\"", "66"],
        ["-19", "-61", "81", "\"xSB\"", "-426"],
        ["229.424", "25", "FALSE", "0", "-71"],
        ["FALSE", "-360.75", "\"p2s\"", "3", "98"],
        ["", "\"b\"", "", "FALSE", "33"],
        [
            "=AVERAGE(D1:D1)",
            "=D1",
            "=A3",
            "=C4",
            "=MIN(INT(E4), LEN(\"A4\"))",
        ],
        [
            "=IF((LEN(\"C4\") > -8), E1, ABS(B1))",
            "=PRODUCT(C2:D4)",
            "=(LOWER(\"31\") + B3)",
            "=B5",
            "=MAX(E5:E6)",
        ],
        [
            "=ROUNDUP(C2, 0)",
            "=IF((ROUNDDOWN(A1, 0) > E1), (D1 ^ B7), (26 + D4))",
            "=(INT(A2) ^ ABS(-7))",
            "=-25",
            "=MIN(A7, SUM(C6:C7))",
        ],
        ["\"hbKN1Pua\"", "-5", "=E8", "=C6", "23"],
        [
            "=LOWER(\"A9\")",
            "=(D8 + MIN(-32, -17))",
            "=D3",
            "=UPPER(\"E5\")",
            "=ROUND(B6, 0)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 1));
    match target {
        ResultData::Float(f) => assert_eq!(f, -57.0),
        other => panic!("Expected Float(-57.0) for B10, got {:?}", other),
    }
}

#[test]
fn test_fuzz_roundup_sum_precision() {
    let sheet_src = [
        ["\"c2B\"", "-289.203", "395.6", "161", "10"],
        ["164.73", "431", "2", "52", "16"],
        ["-434.5", "23", "-15", "4", "\" vW\""],
        ["-1.81", "-11", "-77", "\"siz\"", "-38.13"],
        ["\"mXBoJ\"", "-380.7094", "\"pTbV\"", "-63", "28"],
        [
            "=SQRT(LOWER(\"B5\"))",
            "=-12",
            "=(IF((E4 > E2), 29, A3) * -1)",
            "=A1",
            "=AVERAGE(C5:E5)",
        ],
        [
            "=B4",
            "=MAX(PRODUCT(C2:C2), OR(B6 > 0, E4 < 100))",
            "-67",
            "-13",
            "=-33",
        ],
        [
            "-340.75",
            "=-35",
            "=MAX(LEFT(\"4\", 2), AVERAGE(A7, -38))",
            "=ABS(UPPER(\"-4\"))",
            "=ROUND(INT(A5), 0)",
        ],
        [
            "=B3",
            "=UPPER(\"C7\")",
            "=C4",
            "=D6",
            "=(A8 - OR(33 > 0, C7 < 100))",
        ],
        [
            "=ROUNDUP(SUM(C1, C4), 2)",
            "=CONCATENATE(\"ROUNDDOWN(B2, 0)\", \"INT(A9)\")",
            "-77",
            "=INT(A4)",
            "5",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    match target {
        ResultData::Float(f) => assert!(
            (f - 318.61).abs() < 1e-3,
            "Expected ~318.61 for A10, got {}",
            f
        ),
        other => panic!("Expected Float(~318.61) for A10, got {:?}", other),
    }
}

#[test]
fn test_fuzz_power_product_addition() {
    let sheet_src = [
        ["7", "74", "57", "FALSE", "43.65"],
        ["37", "42", "-55", "", "86"],
        ["-8", "0", "TRUE", "-483.6633", ""],
        ["89", "230", "4", "\"QCOAtn\"", "FALSE"],
        ["0", "310.526", "-89.5", "34", "-80"],
        [
            "=-10",
            "=A1",
            "=MIN(A2:E3)",
            "=(42 - OR(-45 > 0, D3 < 100))",
            "=IF((ABS(D2) > E4), MAX(E5, -30), B5)",
        ],
        [
            "16",
            "=OR((D2 * B6) > 0, LEFT(\"10\", 3) < 100)",
            "=D2",
            "=D5",
            "462.672",
        ],
        [
            "=PRODUCT(ABS(17), (-17 ^ A4))",
            "=C6",
            "=((C6 ^ D6) + PRODUCT(C3, A3))",
            "7",
            "=SQRT(MIN(E4:E4))",
        ],
        [
            "=(28 * (28 - E5))",
            "=B1",
            "-492",
            "=A2",
            "=(UPPER(\"E1\") / A2)",
        ],
        [
            "=A8",
            "=E5",
            "=A5",
            "=E6",
            "=IF((ABS(C5) > 18), (C2 * 36), (-16 * B2))",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 2));
    match target {
        ResultData::Float(f) => assert!((f - -1.1648460276016541e110).abs() / 1e110 < 1e-3),
        other => panic!("Expected Float(~-1.1648e110), got {:?}", other),
    }
}

#[test]
fn test_fuzz_division_by_sum() {
    let sheet_src = [
        ["14", "-100", "7", "438.807", "-310.23"],
        ["2", "\"JQxaopYx\"", "9", "", "\"K\""],
        ["-5", "93", "-195.001", "0", "46"],
        ["81", "\"EmQomuP\"", "\"b\"", "-198.6326", "FALSE"],
        ["", "\"QPDfo\"", "45", "-11", "0"],
        [
            "=ABS(B5)",
            "=SQRT(B4)",
            "99",
            "=IF((2 > INT(18)), OR(C5 > 0, -34 < 100), A4)",
            "=(INT(D1) / B5)",
        ],
        [
            "=LEN(\"IF((A2 > E1), A2, -35)\")",
            "=MAX(B4:E4)",
            "=MAX(D5, MAX(C3, C1))",
            "492",
            "=ROUND(E6, 2)",
        ],
        [
            "=D6",
            "=(D6 / SUM(C3, C3))",
            "=-28",
            "=ROUNDUP(IF((D7 > E2), A6, B3), 0)",
            "=A6",
        ],
        [
            "=MIN(B5, (D2 * 10))",
            "=(LEN(\"E5\") / -15)",
            "=A5",
            "=UPPER(\"(-10 * D2)\")",
            "=ROUNDDOWN(MIN(C1:D4), 0)",
        ],
        [
            "=MIN(C2:D3)",
            "=50",
            "=E7",
            "=E1",
            "=IF((IF((E7 > 8), E9, -33) > SQRT(D5)), 20, (A5 * 27))",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 1));
    match target {
        ResultData::Float(f) => assert!(
            (f - -0.207691).abs() < 1e-3,
            "Expected ~ -0.207691 for B8, got {}",
            f
        ),
        other => panic!("Expected Float(~ -0.207691) for B8, got {:?}", other),
    }
}

#[test]
fn test_fuzz_rounddown_power_product_small() {
    let sheet_src = [
        ["0", "-279", "444.6148", "FALSE", "25"],
        ["-141.946", "FALSE", "68", "145.3101", "493"],
        ["12.7", "138.2", "8", "-70", "-39.05"],
        ["53.593", "9", "11", "-359", "-32"],
        ["41", "5", "77", "\"EZ\"", "57"],
        [
            "=D3",
            "404.051",
            "=A5",
            "=B4",
            "=(SUM(B2:B4) / ROUND(D4, 1))",
        ],
        [
            "\"I\"",
            "=MAX(D4, E2)",
            "=B6",
            "\"FN\"",
            "=AND(B3 > 0, E3 < 100)",
        ],
        [
            "=46",
            "\"oE\"",
            "=-30",
            "=IF((B4 > -3), E7, ABS(-50))",
            "=(ABS(A2) / ABS(-4))",
        ],
        [
            "=(ROUNDDOWN(B7, 2) ^ PRODUCT(D6:E6))",
            "\"1\"",
            "=E7",
            "\"Cg1U\"",
            "=LEN(\"(C7 - B7)\")",
        ],
        ["50", "=A7", "=AVERAGE(A9:B9)", "=A2", "48"],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 0));
    println!("Seed 251454 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert!((f - 1.1553665264221954e-10).abs() < 1e-18),
        ResultData::Integer(i) => assert_eq!(i, 0),
        other => panic!("Expected tiny float, got {:?}", other),
    }
}

#[test]
fn test_fuzz_sum_power_int_large_negative_exponent() {
    let sheet_src = [
        ["\"dXjDtD\"", "", "-45", "TRUE", "5"],
        ["0", "\"uaEG cUx\"", "-11", "-98", ""],
        ["-72.90000000000001", "61", "\" zr\"", "75", "-96"],
        ["-114.918", "54.135", "FALSE", "-84", "0"],
        ["-63.6609", "\"Rlb\"", "-44.3", "\"SM\"", ""],
        [
            "=MIN(A3, (-18 * -12))",
            "=RIGHT(\"AND(16 > 0, 33 < 100)\", 4)",
            "=(SQRT(10) / ROUND(E4, 2))",
            "=AND(4 > 0, AND(-15 > 0, 50 < 100) < 100)",
            "=A4",
        ],
        [
            "0",
            "=D5",
            "=(MAX(E1:E4) + MIN(D3, E1))",
            "=(SUM(E6:E6) ^ INT(E6))",
            "=IF((SQRT(19) > C5), -23, C4)",
        ],
        [
            "-415.27",
            "=MAX(SUM(C5:D7), (A4 - -27))",
            "=37",
            "=MIN(D2:E4)",
            "-5",
        ],
        [
            "=(SUM(C8:C8) + IF((B7 > C8), B1, B2))",
            "=LEFT(\"(15 + 9)\", 1)",
            "=LEFT(\"ROUNDDOWN(D5, 2)\", 4)",
            "=(AVERAGE(C2:E3) / (A4 + -47))",
            "=ROUNDDOWN(IF((E7 > -13), A4, A1), 1)",
        ],
        [
            "=AND(-43 > 0, -44 < 100)",
            "=IF((B2 > LEFT(\"A7\", 5)), MIN(A1, -22), (D9 - B3))",
            "=SQRT((A6 + -12))",
            "=B9",
            "=19",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 3));
    println!("Seed 475402 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert!((f - -1.1359866023259109e-237).abs() < 1e-240),
        ResultData::Integer(i) => assert_eq!(i, 0),
        other => panic!("Expected float, got {:?}", other),
    }
}

#[test]
fn test_fuzz_rounddown_average_a5() {
    let sheet_src = [
        ["", "53", "41", "9", "486.3"],
        ["235.999", "-219.94", "\"qcBBBm\"", "-42", "\"DaE3PM\""],
        ["\"t\"", "43", "11", "-89", "3"],
        ["\"amn\"", "39", "\"Al2\"", "\"PK\"", "29"],
        ["-27", "\"uiY\"", "449.2", "FALSE", "-85"],
        [
            "",
            "=((D5 / D3) * 0)",
            "=UPPER(\"8\")",
            "=ROUND(AND(E2 > 0, E4 < 100), 1)",
            "=AVERAGE(A5:A5)",
        ],
        ["=ROUNDDOWN((-7 + B4), 2)", "=E6", "=E6", "=B5", "-18"],
        [
            "=(ROUND(-14, 0) / C4)",
            "\"zMxPb\"",
            "=SUM(E5:E5)",
            "=ROUNDDOWN(C7, 0)",
            "=C2",
        ],
        [
            "=C4",
            "-43",
            "=SUM(B2:E5)",
            "-472.5",
            "=(-34 * ROUNDUP(-50, 2))",
        ],
        [
            "21",
            "=MAX(E7:E7)",
            "=IF((-35 > IF((C7 > B6), A9, E7)), (6 - D3), C5)",
            "=MIN(AVERAGE(C9:E9), -3)",
            "=C7",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 3));
    println!("Seed 58883 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, -27.0),
        ResultData::Integer(i) => assert_eq!(i, -27),
        other => panic!("Expected -27, got {:?}", other),
    }
}

#[test]
fn test_fuzz_rounddown_abs_product_precision() {
    let sheet_src = [
        ["\"jgJ\"", "\"lor\"", "FALSE", "", "-423"],
        ["-203", "", "8", "FALSE", "\"jNlVx\""],
        ["FALSE", "-462", "", "-35", "-0.617"],
        ["87", "68", "-51", "TRUE", "-40"],
        ["\"1amXg\"", "54", "76", "", "-34"],
        [
            "=(-29 / E3)",
            "=AND(-1 > 0, LEFT(\"A4\", 3) < 100)",
            "\"zNHR\"",
            "=SUM(E1:E4)",
            "=PRODUCT(D3:E5)",
        ],
        [
            "=(ROUND(D3, 0) ^ CONCATENATE(\"33\", \"-42\"))",
            "=E2",
            "=D1",
            "=37",
            "\"Nr\"",
        ],
        ["=47", "=ROUNDDOWN(ABS(E6), 2)", "=C6", "9", ""],
        ["=50", "=-7", "=LEN(\"D3\")", "=-24", "=A2"],
        [
            "=8",
            "=E6",
            "=-38",
            "=-3",
            "=CONCATENATE(\"A4\", \"AVERAGE(B9:D9)\")",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 1));
    println!("Seed 810486 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 29369.2),
        ResultData::Integer(i) => assert_eq!(i, 29369),
        other => panic!("Expected 29369.2, got {:?}", other),
    }
}

#[test]
fn test_fuzz_average_nested_d2_d8() {
    let sheet_src = [
        ["TRUE", "", "-262.09", "-19", ""],
        ["356.364", "\"dWMrcEg\"", "-402.123", "", "FALSE"],
        ["\"A\"", "-6", "210.2262", "", "-39"],
        ["\"jDjm m\"", "", "-210.36", "-55", "-415.4"],
        ["", "TRUE", "75", "", "267.06"],
        [
            "=LOWER(\"AND(-40 > 0, -42 < 100)\")",
            "=A4",
            "=MIN(B2:D2)",
            "=ABS(B5)",
            "=C2",
        ],
        [
            "TRUE",
            "-176.875",
            "=ROUND((C5 - A4), 1)",
            "=ROUNDUP(IF((E6 > 41), E6, D6), 2)",
            "=ROUNDDOWN(-42, 1)",
        ],
        [
            "=IF((-11 > SUM(E6, A4)), 16, -47)",
            "=AVERAGE(AND(D1 > 0, D7 < 100), (A6 + B3))",
            "=MIN(C6:C7)",
            "=C3",
            "=AND(IF((D4 > D3), A5, -34) > 0, B2 < 100)",
        ],
        [
            "=(SQRT(A3) ^ RIGHT(\"24\", 3))",
            "=C6",
            "=ROUNDDOWN(B5, 2)",
            "=B7",
            "=(B1 / SUM(C4:E6))",
        ],
        [
            "=AND(E4 > 0, MAX(D7, -47) < 100)",
            "=ROUND(A7, 1)",
            "=AVERAGE(AVERAGE(D2:D8), C1)",
            "=-34",
            "=D1",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 2));
    println!("Seed 585372 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert!((f - -111.391725).abs() < 1e-4),
        ResultData::Integer(i) => assert_eq!(i, -111),
        other => panic!("Expected -111.391725, got {:?}", other),
    }
}

#[test]
fn test_fuzz_roundup_min_float_precision_val() {
    let sheet_src = [
        ["", "473.2", "-89", "8", ""],
        ["", "", "6", "\"ozkAb\"", "-85"],
        ["58", "-23", "17", "38", "2"],
        ["400.1", "-92", "5", "", "-33"],
        ["-42", "", "", "-70", "56"],
        [
            "=0",
            "=D1",
            "=A5",
            "=OR(RIGHT(\"E1\", 4) > 0, LEFT(\"E1\", 4) < 100)",
            "=IF((MAX(E2:E5) > E2), PRODUCT(A4, A3), -4)",
        ],
        [
            "=ROUND(46, 2)",
            "=IF((LOWER(\"B4\") > (B1 ^ C3)), D3, (D5 * B3))",
            "=AND(-3 > 0, IF((D6 > 50), A3, A6) < 100)",
            "=ROUND(37, 1)",
            "=(LOWER(\"14\") - IF((B2 > C4), C3, D2))",
        ],
        [
            "=(A6 ^ IF((C3 > 15), B5, C5))",
            "=IF((AVERAGE(15, A7) > (D6 - A3)), D1, ROUNDUP(9, 1))",
            "=MIN(E1:E7)",
            "=IF((B5 > LEN(\"-8\")), -47, (47 * 20))",
            "=OR(46 > 0, (B2 + 8) < 100)",
        ],
        [
            "=(MAX(E2:E4) ^ (42 / E6))",
            "=OR(AND(B4 > 0, B8 < 100) > 0, C2 < 100)",
            "=(E7 + -44)",
            "=A1",
            "=32",
        ],
        [
            "-298.237",
            "=ROUNDUP(MIN(E6, E8), 1)",
            "=D9",
            "=(RIGHT(\"B8\", 5) - (-12 - A2))",
            "=MAX(D9:D9)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 1));
    println!("Seed 854292 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert!((f - 23205.8).abs() < 0.2),
        ResultData::Integer(i) => assert_eq!(i, 23205),
        other => panic!("Expected ~23205.8, got {:?}", other),
    }
}

#[test]
fn test_fuzz_product_direct_text_arg_error_precedence() {
    let sheet_src = [
        ["\"ddzXs\"", "-76", "-37", "FALSE", "-68"],
        ["-3", "-265.23", "", "\"Y2aD2X\"", "FALSE"],
        ["75", "-475.29", "", "FALSE", "-4"],
        ["\"F\"", "\"X\"", "-188", "\"W3W\"", "10"],
        ["", "TRUE", "-91", "\"ESQVPT\"", "396.8"],
        [
            "10",
            "=ROUNDUP(A5, 1)",
            "=AND(D1 > 0, CONCATENATE(-34, -25) < 100)",
            "=C2",
            "=LEFT(E5, 2)",
        ],
        [
            "=-28",
            "=B4",
            "-35",
            "=E3",
            "=PRODUCT(UPPER(A4), (D3 ^ -32))",
        ],
        ["=B5", "=D3", "18", "=LEFT(-17, 3)", "=A7"],
        [
            "=AVERAGE(IF((E5 > B7), C2, C3), OR(E5 > 0, 25 < 100))",
            "=33",
            "=CONCATENATE(SQRT(D1), -22)",
            "=B8",
            "=ROUNDUP(RIGHT(C8, 1), 2)",
        ],
        ["0", "\"Ue\"", "=A4", "=(D9 - 16)", "=C6"],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 4));
    println!("Seed 787979 target E7: {:?}", target);
    match target {
        ResultData::Error(ref e) => assert!(
            e.contains("#VALUE!"),
            "Expected #VALUE! for E7 (Excel evaluates the direct text arg \"F\" before the #DIV/0! from D3^-32), got {:?}",
            target
        ),
        other => panic!("Expected Error(#VALUE!) for E7, got {:?}", other),
    }
}

#[test]
fn test_fuzz_average_range_ignores_text_cell() {
    let sheet_src = [
        ["3", "292.843", "290.419", "-42", "3"],
        ["FALSE", "-76", "TRUE", "-26", "487.552"],
        ["115.5", "1", "-431", "", ""],
        ["16", "74", "-57", "303.83", "\"gIp\""],
        ["-64", "\"PyUBvLT\"", "", "-175.247", ""],
        [
            "",
            "=MIN(CONCATENATE(C3, E2), B3)",
            "=ABS(B4)",
            "=A1",
            "=(MAX(C3:D4) ^ AVERAGE(C3:E3))",
        ],
        ["=E1", "=RIGHT(B4, 3)", "8", "=AVERAGE(A3:C3)", "=D1"],
        [
            "=A3",
            "=AND(IF((D5 > -29), D1, A2) > 0, C6 < 100)",
            "\"3lM3\"",
            "=D5",
            "\"1\"",
        ],
        [
            "=-30",
            "=IF((MAX(E1:E3) > LEFT(-11, 4)), OR(D3 > 0, D7 < 100), MAX(E2, B2))",
            "=UPPER(LEN(A2))",
            "=E3",
            "=(ABS(C2) ^ OR(A6 > 0, D6 < 100))",
        ],
        [
            "=RIGHT(AVERAGE(E6:E9), 4)",
            "=28",
            "=B5",
            "=IF((INT(-47) > PRODUCT(E8:E8)), -1, MIN(A8, 44))",
            "4",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    println!("Seed 29632 target A10: {:?}", target);
    match target {
        ResultData::String(ref s) => assert_eq!(
            s, "6667",
            "Expected \"6667\" for A10 (AVERAGE(E6:E9) must ignore the text cell E8=\"1\", giving -41/3, not treat it as the number 1)"
        ),
        other => panic!("Expected String(\"6667\") for A10, got {:?}", other),
    }
}

#[test]
fn test_product_snaps_once_not_per_factor() {
    // Excel snaps a formula's result to 15 significant digits, and PRODUCT
    // is where that is observable twice over.
    //
    // Applying the snap per factor compounds: over seven factors the
    // partial products drift about 14 ULP, and
    // CONCATENATE(PRODUCT(...), ...) rendered 189124133819.665 where real
    // Excel gives 189124133819.664 -- which is also what the plain
    // a*b*c*... chain produces.
    let mut sheet = create_sheet(&[
        ["-198.552", "=CONCATENATE(\"x\", PRODUCT(A1:A7))"],
        ["139.6292", "=A1*A2*A3*A4*A5*A6*A7"],
        ["10", ""],
        ["-44", ""],
        ["-6", ""],
        ["38", ""],
        ["-68", ""],
    ]);
    sheet.commit(None).unwrap();
    match sheet.get_result_data(&CellRef::new(0, 1)) {
        ResultData::String(s) => assert_eq!(s, "x189124133819.664"),
        other => panic!("expected x189124133819.664, got {other:?}"),
    }

    // ... but dropping the snap entirely is wrong too: it is what makes
    // ROUNDDOWN see 29369.2 rather than 29369.199999999997, which is
    // covered by test_fuzz_rounddown_abs_product_precision.
    let mut sheet = create_sheet(&[[
        "-35",
        "-0.617",
        "-40",
        "-34",
        "=ROUNDDOWN(PRODUCT(A1:D1), 2)",
    ]]);
    sheet.commit(None).unwrap();
    match sheet.get_result_data(&CellRef::new(0, 4)) {
        ResultData::Float(f) => assert!((f - 29369.2).abs() < 1e-9, "got {f}"),
        other => panic!("expected 29369.2, got {other:?}"),
    }
}
