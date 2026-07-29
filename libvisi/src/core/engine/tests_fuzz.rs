use super::*;

fn create_sheet<const ROWS: usize, const COLS: usize>(grid: &[[&str; COLS]; ROWS]) -> Sheet {
    let rows = grid.len();
    let cols = grid[0].len();
    let mut sheet = Sheet::new(SheetInit {
        name: Some("sheet1".to_string()),
        rows,
        cols,
        ..Default::default()
    });
    for (i, row) in grid.iter().enumerate() {
        for (j, val) in row.iter().enumerate() {
            sheet.insert(
                TextCellRef {
                    row: i,
                    col: j,
                    char_offset: 0,
                },
                val,
            )
        }
    }

    sheet
}

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
fn test_fuzz_sqrt_negative_operand_error() {
    let sheet_src = [
        ["TRUE", "-278.17", "-55", "29", "240"],
        ["-35", "\"N\"", "7", "-419", "FALSE"],
        ["16", "10", "0", "FALSE", "92"],
        ["39", "TRUE", "66", "\"zpx3A\"", "356"],
        ["-31", "\"ZLwUBtDn\"", "-133.063", "3", "FALSE"],
        [
            "=IF(((28 * -25) > LOWER(\"-8\")), ROUNDUP(D4, 2), E2)",
            "=(LOWER(\"D4\") - D3)",
            "=ROUNDUP(ROUNDUP(D3, 0), 1)",
            "-43",
            "=UPPER(\"ABS(-46)\")",
        ],
        [
            "=B4",
            "=SUM(OR(-12 > 0, B1 < 100), ABS(D3))",
            "318",
            "=C4",
            "=-1",
        ],
        [
            "=(IF((D7 > D1), -29, B7) / UPPER(\"E4\"))",
            "=46",
            "=IF((B7 > (B5 * D1)), CONCATENATE(\"45\", \"-46\"), OR(E7 > 0, C2 < 100))",
            "=(45 ^ -27)",
            "=ROUND((E5 * D6), 1)",
        ],
        [
            "=(IF((D8 > -19), -8, A4) ^ AVERAGE(20, 5))",
            "=PRODUCT((B4 / B5), ROUND(A2, 0))",
            "=D5",
            "=SQRT(E7)",
            "=AND(IF((A3 > 44), 48, E4) > 0, OR(B7 > 0, -49 < 100) < 100)",
        ],
        [
            "=A3",
            "-497.457",
            "=IF((B5 > (C1 * A6)), E7, A6)",
            "=LEN(\"INT(-5)\")",
            "=-27",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 3));
    assert!(
        matches!(target, ResultData::Error(ref e) if e.contains("#NUM!")),
        "Expected #NUM!, got {:?}",
        target
    );
}

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
fn test_fuzz_and_nested_if_boolean_evaluation() {
    let sheet_src = [
        ["49", "6", "-26", "-98", "\"BnnA\""],
        ["\"lrHaQ\"", "72", "", "\"BZm\"", ""],
        ["\"K jKmvaf\"", "\"qrCkEFAQ\"", "11", "\"yguQaatS\"", "57"],
        ["146.3619", "61", "\"VKdXCqYO\"", "", "-79"],
        ["-37", "", "-362.0854", "63", "FALSE"],
        [
            "=PRODUCT(PRODUCT(D1:D5), OR(-29 > 0, B3 < 100))",
            "=INT(LOWER(\"-7\"))",
            "=ABS((E5 - E5))",
            "=-2",
            "=-44",
        ],
        [
            "=ROUNDUP(AVERAGE(D6:E6), 2)",
            "=D3",
            "=ABS(CONCATENATE(\"24\", \"D3\"))",
            "=ROUNDUP(B5, 1)",
            "-285.5",
        ],
        ["-97", "=-31", "46", "=4", "=MAX(C6:E6)"],
        ["\"b\"", "=A8", "=SUM(E4:E7)", "=MIN(E6:E6)", "=E6"],
        [
            "=IF((A7 > E8), (48 * E9), D7)",
            "=-32",
            "=-33",
            "=C5",
            "=AND(IF((50 > -18), B3, -13) > 0, C3 < 100)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(5, 0));
    match target {
        ResultData::Float(f) => assert!(f.abs() < 1e-6),
        other => panic!("Expected Float(0.0) for A6, got {:?}", other),
    }
}

#[test]
fn test_fuzz_roundup_rounddown_nested_error() {
    let sheet_src = [
        ["57", "0", "66", "-66", "-36"],
        ["236.7614", "104", "78", "87", "64.40000000000001"],
        ["-70", "7", "", "-346.6315", "-25"],
        ["-447", "-490.756", "-94", "\"2zorN\"", "\"PzZo\""],
        ["-63", "-29", "46", "\"xYPgPb\"", ""],
        [
            "=ABS(ROUNDUP(C1, 2))",
            "=-12",
            "\"dxvW\"",
            "=INT(SUM(2, C3))",
            "-90",
        ],
        [
            "=INT(OR(B2 > 0, 20 < 100))",
            "=AND(SUM(C4:C6) > 0, 17 < 100)",
            "=OR(RIGHT(\"B1\", 5) > 0, INT(D4) < 100)",
            "=-37",
            "=AND(A3 > 0, B3 < 100)",
        ],
        ["=PRODUCT(C7:C7)", "-46", "=E6", "=42", "=B3"],
        [
            "=(ROUNDDOWN(B2, 0) + (E2 + D1))",
            "",
            "=(B2 ^ A2)",
            "=OR(SQRT(E5) > 0, (A8 / E3) < 100)",
            "228.3",
        ],
        [
            "=IF((IF((16 > B5), B6, E8) > PRODUCT(B4:E7)), 38, 36)",
            "=(E6 ^ -45)",
            "=E3",
            "=(D6 - LEFT(\"A4\", 4))",
            "=ROUNDUP(ROUNDDOWN(D7, 0), 1)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 2));
    assert!(
        matches!(target, ResultData::Boolean(true) | ResultData::Error(_)),
        "Expected Boolean(true) or Error for C7, got {:?}",
        target
    );
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
fn test_fuzz_min_range_division_by_zero_error() {
    let sheet_src = [
        ["-52", "347.193", "14", "6", "4"],
        ["-466", "-4", "TRUE", "\"y\"", ""],
        ["472", "-36", "-10", "-28.908", "267"],
        ["\"vfclM\"", "-60", "414", "\"TiNHbw\"", "-28"],
        ["FALSE", "FALSE", "", "\"LlUvCUn\"", "FALSE"],
        [
            "=(RIGHT(\"B2\", 3) / (C5 / D3))",
            "=SQRT(LOWER(\"D5\"))",
            "=AND(B4 > 0, -31 < 100)",
            "=D2",
            "=MAX((A3 + 34), B1)",
        ],
        ["=E6", "=((C6 / E2) / E1)", "=-2", "=A5", "=E2"],
        ["=C7", "=MIN(A1:A3)", "=SQRT(ROUNDDOWN(E4, 1))", "76", "8"],
        [
            "=MAX(PRODUCT(D7, B6), IF((E7 > E4), C1, B8))",
            "328",
            "=MIN(D4:D5)",
            "0",
            "=((C8 / E3) / ROUND(-21, 0))",
        ],
        [
            "=-39",
            "=(CONCATENATE(\"-30\", \"A6\") - LEN(\"D5\"))",
            "138.2545",
            "=36",
            "=MIN(C1:C9)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 1));
    assert!(
        matches!(target, ResultData::Error(ref e) if e.contains("#DIV/0!")),
        "Expected #DIV/0! for B7, got {:?}",
        target
    );
}

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
fn test_fuzz_boolean_dependency_cell_evaluation() {
    let sheet_src = [
        ["46", "-124.46", "TRUE", "-56", "75"],
        ["FALSE", "\"PJwI\"", "", "75", "3"],
        ["\"W1bdh\"", "FALSE", "0", "\"hOS\"", "10"],
        ["\"lnxfR\"", "TRUE", "-477.673", "\"QYQf\"", "TRUE"],
        ["\"MB\"", "109.944", "8", "25", "-63"],
        [
            "=IF((ROUND(E3, 1) > (-42 * E1)), (A4 / 0), IF((-23 > D5), C3, 49))",
            "=LOWER(\"AVERAGE(A4:A4)\")",
            "=OR(E4 > 0, LEN(\"C2\") < 100)",
            "\"TADFiW\"",
            "=RIGHT(\"D2\", 4)",
        ],
        [
            "=UPPER(\"(24 - E1)\")",
            "FALSE",
            "=PRODUCT(UPPER(\"D2\"), C1)",
            "1",
            "=PRODUCT(D1:E5)",
        ],
        [
            "=A6",
            "=((48 * D4) ^ C1)",
            "=C2",
            "=IF((ABS(A2) > RIGHT(\"D2\", 2)), B5, 9)",
            "=AND(LEN(\"-5\") > 0, E3 < 100)",
        ],
        [
            "=CONCATENATE(\"ROUND(A1, 1)\", \"AND(D7 > 0, 44 < 100)\")",
            "=E1",
            "=-17",
            "=(IF((20 > -39), -2, A7) / B6)",
            "=-2",
        ],
        [
            "=OR((-19 ^ C8) > 0, 35 < 100)",
            "=-7",
            "=ROUND(D3, 1)",
            "=SUM(E8:E9)",
            "=D8",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    match target {
        ResultData::Boolean(b) => assert_eq!(b, true),
        other => panic!("Expected Boolean(true) for A10, got {:?}", other),
    }
}

#[test]
fn test_fuzz_and_multiplication_num_error() {
    let sheet_src = [
        ["FALSE", "-55", "28", "14", "77"],
        ["52", "483.04", "30", "-324.1", "FALSE"],
        ["-54", "", "483", "TRUE", "-272"],
        ["\"CHue\"", "\"xqit\"", "349", "\"uyDfgqA\"", "-86"],
        ["-72", "TRUE", "87", "", "-13"],
        [
            "-67.511",
            "=26",
            "=C2",
            "=OR(RIGHT(\"-32\", 2) > 0, A4 < 100)",
            "408",
        ],
        ["\"JRW\"", "=E4", "=29", "=SQRT(-8)", "=C4"],
        [
            "=MAX(LOWER(\"-2\"), AND(17 > 0, A7 < 100))",
            "=D1",
            "98.194",
            "-2",
            "=ABS(A6)",
        ],
        ["=E1", "=MIN(B1:E3)", "\"YbeYuyK\"", "=E3", "5"],
        [
            "=(PRODUCT(C3, A2) + (D7 + C9))",
            "=E3",
            "=C2",
            "=IF((ROUNDUP(25, 0) > B6), IF((43 > -8), -42, 28), C8)",
            "=(AND(D1 > 0, D5 < 100) * A5)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    assert!(
        matches!(target, ResultData::Error(ref e) if e.contains("#NUM!")),
        "Expected #NUM! for A10, got {:?}",
        target
    );
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
            "Expected ~27.5533 for B8, got {}",
            f
        ),
        other => panic!("Expected Float(~27.5533) for B8, got {:?}", other),
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
fn test_fuzz_division_by_zero_formula_error() {
    let sheet_src = [
        ["", "77", "1", "\"WdtMTpo\"", "TRUE"],
        ["91", "\"O\"", "50", "332.55", "100"],
        ["-254.2", "-3", "-34", "-12", "\"Fz\""],
        ["\"zE\"", "\"V\"", "-94", "", "-10"],
        ["\"tL\"", "91", "\"pCuZO 2\"", "3", "38"],
        ["=INT(21)", "=B5", "=E1", "=B1", ""],
        [
            "=E1",
            "=AND(LEN(\"A2\") > 0, 17 < 100)",
            "=UPPER(\"INT(A3)\")",
            "TRUE",
            "=IF((OR(E6 > 0, D4 < 100) > 49), C2, B5)",
        ],
        [
            "=AVERAGE(A7, LOWER(\"-15\"))",
            "=-15",
            "-249.4955",
            "=RIGHT(\"15\", 5)",
            "=D7",
        ],
        [
            "=((A7 + D3) - CONCATENATE(\"E3\", \"-38\"))",
            "=24",
            "=SQRT(43)",
            "=INT(AVERAGE(D6:E7))",
            "=SQRT(SUM(E5:E5))",
        ],
        [
            "=-13",
            "=AVERAGE(C9:E9)",
            "=AVERAGE(B7:C9)",
            "=D5",
            "=ROUNDDOWN(D9, 0)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 0));
    match target {
        ResultData::Error(_) => {}
        ResultData::Float(f) => assert_eq!(f, -15.0),
        other => panic!("Expected Float(-15.0) for A8, got {:?}", other),
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
fn test_fuzz_sqrt_log_range_error() {
    let sheet_src = [
        ["\"OgDll3XO\"", "-45", "320.12", "3", "-450.4"],
        ["62", "72", "-339.5", "298.5", "-21"],
        ["\"NxSOHE\"", "", "66", "79", "335.4112"],
        ["53", "\"vsFIR\"", "TRUE", "-307.808", "TRUE"],
        ["-95", "\"ArfF\"", "-45", "", "5"],
        [
            "=LOWER(\"36\")",
            "=UPPER(\"19\")",
            "=IF((ROUNDUP(C3, 0) > C1), PRODUCT(E3:E4), UPPER(\"A2\"))",
            "=C3",
            "=SQRT(PRODUCT(D3, -15))",
        ],
        [
            "=IF((C3 > (E6 ^ B3)), A3, LEN(\"-7\"))",
            "=AVERAGE(INT(-8), INT(C1))",
            "=36",
            "=IF((-31 > C4), UPPER(\"D1\"), C3)",
            "=RIGHT(\"IF((E2 > -13), E4, A2)\", 3)",
        ],
        [
            "=C5",
            "=B7",
            "=PRODUCT(A4:E5)",
            "",
            "=UPPER(\"PRODUCT(C1, 20)\")",
        ],
        [
            "=SUM(B2:C2)",
            "=(E7 - C2)",
            "=AND(D7 > 0, 31 < 100)",
            "=(19 + 9)",
            "0",
        ],
        [
            "=(ROUND(D4, 2) ^ AVERAGE(D7:E7))",
            "=E5",
            "=PRODUCT(A1:C8)",
            "=ROUNDDOWN(AND(C2 > 0, C8 < 100), 1)",
            "TRUE",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 0));
    assert!(
        matches!(target, ResultData::Error(ref e) if e.contains("#NUM!")),
        "Expected #NUM! for A7, got {:?}",
        target
    );
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
fn test_fuzz_and_if_comparison_evaluation() {
    let sheet_src = [
        ["0", "", "46", "", "-69"],
        ["-50.9", "4", "-2.81", "52", "27"],
        ["38", "", "18", "-23", "108.8"],
        ["\"IZKTZKJ\"", "193.2", "-52", "-310.4212", "\"kn\""],
        ["32", "-94.88", "199.402", "TRUE", "-1.82"],
        ["=48", "=LEN(\"D1\")", "", "=(A3 * INT(A3))", "=46"],
        ["=A3", "=26", "=SQRT(42)", "=LOWER(\"B5\")", "=A6"],
        [
            "=AVERAGE(E2, ROUNDDOWN(10, 2))",
            "=C4",
            "=C7",
            "=ROUNDDOWN(MAX(C1, -49), 0)",
            "=-6",
        ],
        [
            "=C4",
            "5",
            "=MAX(E6, 40)",
            "=6",
            "=(OR(A2 > 0, -3 < 100) / ROUND(-38, 0))",
        ],
        [
            "-430.0124",
            "=AND(IF((D1 > D3), 38, D3) > 0, E7 < 100)",
            "=(33 ^ CONCATENATE(\"D8\", \"A8\"))",
            "TRUE",
            "=((C6 / B3) + IF((C1 > E8), D6, C9))",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 1));
    match target {
        ResultData::Boolean(b) => assert!(b, "Expected true for B10"),
        other => panic!("Expected Boolean(true) for B10, got {:?}", other),
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
fn test_fuzz_if_multiplication_comparison() {
    let sheet_src = [
        ["\"mQK1CAkE\"", "\"fUeQvRXH\"", "3", "386.992", "\"1QU\""],
        ["", "7", "\"rTrosz\"", "TRUE", "\"j2q\""],
        ["", "-344.821", "-41", "1", ""],
        ["", "", "", "-91", "-397"],
        ["69", "-14", "463.5482", "FALSE", "\"FWKVlndl\""],
        [
            "=SQRT(ROUNDUP(31, 0))",
            "=A2",
            "=ROUND(ABS(A5), 2)",
            "=37",
            "=B5",
        ],
        ["=E6", "=D1", "=C5", "=C5", "=(-13 + MAX(E6:E6))"],
        [
            "=ROUND(-40, 2)",
            "=(B7 * MAX(D2, C5))",
            "=A1",
            "28",
            "=MAX((B3 + E6), SUM(C3:D5))",
        ],
        [
            "",
            "=IF((ABS(C8) > CONCATENATE(\"31\", \"A6\")), LEN(\"-45\"), B8)",
            "=((B1 ^ E4) / MIN(A6:B8))",
            "=IF((C5 > (C4 * C6)), (B3 / E8), MIN(D1, D5))",
            "=CONCATENATE(\"SUM(-33, A4)\", \"AVERAGE(A5:A8)\")",
        ],
        ["=C3", "FALSE", "=-32", "=INT(E6)", "-287.722"],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 3));
    match target {
        ResultData::Float(f) => assert!(
            (f - -1.036905).abs() < 1e-3,
            "Expected ~ -1.036905 for D9, got {}",
            f
        ),
        other => panic!("Expected Float(~ -1.036905) for D9, got {:?}", other),
    }
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
fn test_fuzz_average_if_branch_evaluation() {
    let sheet_src = [
        ["-23", "", "\"OElW\"", "72", "1"],
        ["413.757", "-2", "60", "-71", "66"],
        ["-86", "8", "33", "TRUE", "FALSE"],
        ["209.28", "79", "\"f\"", "162.008", "\"ZKu\""],
        ["\"2dtEYT\"", "", "\"2qdSZiqu\"", "26", "2"],
        ["=B4", "41", "=AVERAGE(A4:A5)", "=B2", "=E2"],
        [
            "=ROUNDUP(D5, 0)",
            "=(AND(C2 > 0, A6 < 100) + (35 * 0))",
            "=OR(D3 > 0, IF((D6 > E2), A2, E6) < 100)",
            "=SUM((26 - D2), A4)",
            "=IF(((B6 * D5) > ROUNDUP(A6, 2)), (C1 + -42), (A6 - D1))",
        ],
        [
            "=E1",
            "=B6",
            "=ROUND(CONCATENATE(\"-3\", \"E7\"), 1)",
            "21",
            "=MIN(A2:D7)",
        ],
        [
            "=28",
            "TRUE",
            "=RIGHT(\"D3\", 2)",
            "=AVERAGE(IF((E2 > C1), C2, E3), C8)",
            "=(-40 ^ UPPER(\"29\"))",
        ],
        [
            "=E2",
            "=E8",
            "=(1 - A6)",
            "=A8",
            "=(ROUNDUP(21, 1) - SQRT(42))",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 3));
    match target {
        ResultData::Float(f) => assert_eq!(f, -30000000.0),
        other => panic!("Expected Float(-30000000.0) for D9, got {:?}", other),
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
                assert!(col_let >= 'A' && col_let <= 'E');
            }
        }
    }
}

#[test]
fn test_fuzz_sqrt_or_evaluation() {
    let sheet_src = [
        ["FALSE", "\"2DeBb\"", "\"aVisZWJZ\"", "17", "\"O\""],
        ["\"Y3i\"", "\"2jvd\"", "\"Dpnvny\"", "TRUE", "-49"],
        ["33", "0", "98", "\"fljvy\"", "-373.3952"],
        ["-19", "4", "\"gipC1lq\"", "TRUE", "65"],
        ["22", "-62", "\"RrDIE\"", "-51", "59"],
        [
            "=CONCATENATE(\"INT(-28)\", \"-43\")",
            "=28",
            "",
            "=10",
            "=(SUM(C2:D2) * E4)",
        ],
        [
            "=A4",
            "=A6",
            "163.57",
            "=ROUND(CONCATENATE(\"C2\", \"C5\"), 2)",
            "=ROUNDDOWN(C3, 0)",
        ],
        [
            "\"J\"",
            "=34",
            "=ROUND(LOWER(\"-34\"), 1)",
            "-97",
            "=(D5 - RIGHT(\"B5\", 3))",
        ],
        [
            "=SQRT(D6)",
            "-77",
            "=(-46 + (-3 ^ E5))",
            "=OR(A7 > 0, (A4 + E6) < 100)",
            "=C5",
        ],
        [
            "=MIN(A7:C8)",
            "=D2",
            "-283.3129",
            "=ROUNDUP(SUM(D2:D7), 2)",
            "=A5",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target1 = sheet.get_result_data(&CellRef::new(8, 0));
    let target2 = sheet.get_result_data(&CellRef::new(8, 3));
    match target1 {
        ResultData::Float(f) => assert!(
            (f - 3.162277).abs() < 1e-3,
            "Expected ~3.162277 for A9, got {}",
            f
        ),
        other => panic!("Expected Float(~3.162277) for A9, got {:?}", other),
    }
    match target2 {
        ResultData::Boolean(b) => assert!(b, "Expected true for D9"),
        other => panic!("Expected Boolean(true) for D9, got {:?}", other),
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
fn test_fuzz_sum_or_addition_coercion() {
    let sheet_src = [
        ["TRUE", "\"vegfg\"", "-295.32", "TRUE", "25"],
        ["\"JwVJ\"", "-324.78", "7", "-49", "TRUE"],
        ["-202.3", "-353", "FALSE", "-98", "-5"],
        ["-41", "\"Ya\"", "", "19", "4"],
        ["-18", "", "339.594", "16", "39"],
        [
            "=C5",
            "=-6",
            "=IF((C4 > MIN(A1, C5)), (A2 - 7), ROUNDDOWN(31, 2))",
            "=E5",
            "=D3",
        ],
        [
            "=IF((IF((B4 > E1), D6, 26) > AVERAGE(E1, 23)), ROUNDDOWN(15, 1), D2)",
            "=RIGHT(\"(E6 ^ 17)\", 1)",
            "=MAX(B4:D4)",
            "=AND(UPPER(\"B3\") > 0, (B4 / D1) < 100)",
            "=B5",
        ],
        [
            "TRUE",
            "=13",
            "=(IF((48 > -1), A6, -41) - (E3 + B7))",
            "=AND((40 / 36) > 0, LOWER(\"C2\") < 100)",
            "=A7",
        ],
        [
            "=AVERAGE(A8:A8)",
            "=MAX(-42, LOWER(\"C3\"))",
            "=SUM(OR(47 > 0, -49 < 100), (E8 + 35))",
            "=ROUND(SQRT(1), 0)",
            "=B4",
        ],
        [
            "=((D7 - A4) / AND(D5 > 0, E5 < 100))",
            "-15",
            "=B5",
            "=OR(-15 > 0, 10 < 100)",
            "51.9",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 2));
    match target {
        ResultData::Float(f) => assert_eq!(f, 51.0),
        other => panic!("Expected Float(51.0), got {:?}", other),
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
fn test_fuzz_roundup_subtraction_boolean_addition() {
    let sheet_src = [
        ["25", "5", "", "24", "-17"],
        ["-27", "28", "-93.3104", "74", "-71.2"],
        ["9", "-41", "16", "FALSE", "-339"],
        ["7", "", "-90", "-443.5", "285.242"],
        ["16", "24", "FALSE", "-19", "480.2611"],
        [
            "=PRODUCT(CONCATENATE(\"E3\", \"C3\"), PRODUCT(B5:E5))",
            "0",
            "=OR((E4 + 37) > 0, ROUNDDOWN(E4, 2) < 100)",
            "=ABS(MIN(C4:C4))",
            "=SUM(B3:B4)",
        ],
        [
            "=MAX(A4:D5)",
            "=(-3 * (E4 * A5))",
            "=IF((MAX(B2:B4) > (C3 - A3)), D6, CONCATENATE(\"A3\", \"C1\"))",
            "=D4",
            "=A5",
        ],
        [
            "=D5",
            "=C4",
            "=MIN(RIGHT(\"-33\", 3), E4)",
            "=B7",
            "=RIGHT(\"OR(D5 > 0, -28 < 100)\", 5)",
        ],
        ["", "=E8", "=INT(-36)", "TRUE", "=B4"],
        [
            "=LEFT(\"ABS(26)\", 2)",
            "=B4",
            "=(ROUNDUP(C7, 0) - (D1 + B6))",
            "=ROUNDUP(PRODUCT(B6, -6), 0)",
            "=MIN(D5:D9)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 2));
    match target {
        ResultData::Float(f) => assert_eq!(f, 66.0),
        other => panic!("Expected Float(66.0), got {:?}", other),
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
fn test_fuzz_addition_formula_constant() {
    let sheet_src = [
        ["-74", "-14", "-51", "-73.3", ""],
        ["", "FALSE", "\"i2blI\"", "\"IvctY\"", "-76"],
        ["-27", "-55", "21", "42", "83.538"],
        ["-31", "-27", "", "-45.363", "-8"],
        ["0", "74.40000000000001", "-155.7", "", "-53.258"],
        [
            "=SUM(D4, UPPER(\"-35\"))",
            "TRUE",
            "=IF((B1 > OR(-19 > 0, B1 < 100)), PRODUCT(D1:D3), LOWER(\"A1\"))",
            "",
            "",
        ],
        [
            "=OR(CONCATENATE(\"1\", \"13\") > 0, LOWER(\"A6\") < 100)",
            "=(-14 / (E4 * 10))",
            "=MAX(INT(C6), PRODUCT(A4, D5))",
            "=ROUND(C6, 2)",
            "=IF((B5 > LOWER(\"B2\")), IF((B4 > -39), A2, C6), AVERAGE(A1, B5))",
        ],
        [
            "=IF((RIGHT(\"E4\", 4) > B2), -9, C7)",
            "63",
            "=24",
            "=A2",
            "=-31",
        ],
        [
            "=MIN((16 * C3), (A5 / B1))",
            "=IF((C4 > D1), CONCATENATE(\"E8\", \"47\"), 33)",
            "=B3",
            "5",
            "=A1",
        ],
        [
            "=SQRT(ROUNDDOWN(B4, 2))",
            "=(E7 + 30)",
            "=A9",
            "180.58",
            "=(-13 * (C5 / B6))",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 1));
    match target {
        ResultData::Float(f) => assert!((f - 30.2).abs() < 1e-3),
        other => panic!("Expected Float(~30.2), got {:?}", other),
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
                assert!(col_let >= 'A' && col_let <= 'E');
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
fn test_fuzz_product_abs_if_branch() {
    let sheet_src = [
        ["", "92", "-3", "8", "33"],
        ["-318", "0", "-43", "0", "\"hxtO3rV\""],
        ["-477.5", "78.2", "96", "FALSE", "-328.849"],
        ["", "340.91", "", "50", "-418.3604"],
        ["-46", "7", "\"vfH1Q\"", "2", "\"eR\""],
        [
            "=ABS(ROUNDUP(C3, 1))",
            "=RIGHT(\"(B3 * A2)\", 5)",
            "=ROUNDUP(CONCATENATE(\"-26\", \"A3\"), 1)",
            "-62",
            "=IF((D4 > RIGHT(\"B3\", 3)), LEN(\"E3\"), IF((38 > 44), 0, 36))",
        ],
        [
            "=LEN(\"-13\")",
            "68",
            "=AND(A4 > 0, SUM(-36, A6) < 100)",
            "=INT((B6 / E4))",
            "=RIGHT(\"D2\", 5)",
        ],
        ["=SUM(A6:B6)", "\"virpT\"", "=SUM(D1:E6)", "=B6", "=A7"],
        [
            "=E1",
            "=INT((-30 - E2))",
            "=AND(B4 > 0, IF((-25 > A4), 35, B3) < 100)",
            "=SUM(IF((E2 > -48), D5, C5), LOWER(\"A4\"))",
            "=UPPER(\"D8\")",
        ],
        [
            "=18",
            "=PRODUCT(ABS(E8), IF((-4 > -12), -5, A2))",
            "=ROUNDDOWN(SQRT(E6), 0)",
            "=(-4 * B1)",
            "=(E7 ^ LOWER(\"E4\"))",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 1));
    match target {
        ResultData::Float(f) => assert_eq!(f, -15.0),
        other => panic!("Expected Float(-15.0) for B10, got {:?}", other),
    }
}

#[test]
fn test_fuzz_abs_power_negative_base_error() {
    let sheet_src = [
        ["\"mDV2O\"", "0", "\"ww\"", "-37", "-255.2"],
        ["25", "-272.34", "3", "TRUE", "-4"],
        ["84", "\"2Hlb\"", "75", "FALSE", "47"],
        ["220.459", "TRUE", "-8", "-9", "47"],
        ["196.03", "10", "0", "", "8"],
        [
            "=(D3 ^ UPPER(\"A3\"))",
            "=E5",
            "=(E2 * -34)",
            "406.9",
            "=AND(48 > 0, SQRT(-25) < 100)",
        ],
        ["-472.176", "=ABS((E6 ^ A1))", "", "=C5", "=D6"],
        [
            "=ROUND(-4, 1)",
            "=CONCATENATE(\"MIN(D7:E7)\", \"IF((-3 > D6), B1, -16)\")",
            "=(ROUND(D1, 1) - A6)",
            "=(C2 * -30)",
            "=MAX(UPPER(\"-50\"), -8)",
        ],
        [
            "=UPPER(\"ROUND(45, 0)\")",
            "=E2",
            "=IF((ROUNDUP(B6, 2) > MIN(E3:E7)), C5, E2)",
            "=SQRT(INT(A6))",
            "=(B4 ^ LEFT(\"D2\", 2))",
        ],
        [
            "=17",
            "=SUM(LOWER(\"47\"), IF((E3 > C2), B1, D6))",
            "=E3",
            "=LEFT(\"SQRT(C5)\", 3)",
            "-55",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 1));
    assert!(
        matches!(target, ResultData::Error(ref e) if e == "#NUM!"),
        "Expected #NUM! for B7, got {:?}",
        target
    );
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
fn test_fuzz_if_power_comparison_branches() {
    let sheet_src = [
        ["", "", "\"uALF eK\"", "\"XgB\"", ""],
        ["31", "79.81480000000001", "\"3\"", "\"raaPD\"", "4"],
        ["", "-30", "69", "FALSE", "-78"],
        ["-11.21", "72", "-1", "\"jSNe3xpR\"", "65"],
        ["\"nUmwj\"", "10", "88", "99", "TRUE"],
        [
            "=INT(ROUNDUP(C1, 0))",
            "=(D3 ^ IF((D1 > C2), D2, D4))",
            "=A4",
            "=A2",
            "FALSE",
        ],
        [
            "=ROUND(A4, 1)",
            "=RIGHT(\"-3\", 4)",
            "\"uWOqN\"",
            "=PRODUCT(E2:E4)",
            "=D6",
        ],
        [
            "=ROUNDDOWN(B6, 1)",
            "=50",
            "18",
            "=(C6 + ROUND(-35, 0))",
            "=MIN(A3, A1)",
        ],
        [
            "=IF(((E8 ^ A3) > D7), (E6 + 4), (D6 / -5))",
            "=C6",
            "=-43",
            "-95.50839999999999",
            "=IF((A7 > UPPER(\"29\")), E5, INT(A7))",
        ],
        [
            "=AND(ROUND(-5, 1) > 0, A5 < 100)",
            "=D4",
            "\"Y\"",
            "=-28",
            "=B8",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 0));
    assert!(
        matches!(target, ResultData::Error(ref e) if e == "#NUM!"),
        "Expected #NUM! for A9, got {:?}",
        target
    );
}

#[test]
fn test_fuzz_subtraction_and_comparison_division() {
    let sheet_src = [
        ["-85", "0", "85", "FALSE", "-67"],
        ["65", "\"NAU\"", "8", "-93", "1"],
        ["-201.948", "9", "-245", "24.4", "78"],
        ["\"S\"", "102", "-82", "\"A\"", "58"],
        ["-74", "-21", "\"1\"", "399.35", "264.902"],
        [
            "=IF((A3 > ABS(D3)), B3, CONCATENATE(\"B1\", \"B3\"))",
            "=B2",
            "=(-28 + (B4 ^ -48))",
            "0",
            "=MAX(C2:D5)",
        ],
        [
            "=PRODUCT((C5 / D3), D2)",
            "=(IF((B4 > E6), A5, 15) - (C6 + E3))",
            "=C4",
            "=MIN(E4:E6)",
            "=E4",
        ],
        [
            "=AND(IF((D2 > C3), B4, B7) > 0, ROUNDUP(5, 0) < 100)",
            "=D7",
            "39",
            "=(C2 + IF((B2 > A3), 43, B6))",
            "=IF((ROUNDDOWN(E7, 2) > E4), ABS(D1), MAX(E4:E4))",
        ],
        ["=4", "=5", "62", "=MAX(C2:D6)", "\"M\""],
        [
            "=(AND(C3 > 0, C3 < 100) - (E8 / A3))",
            "=D1",
            "=ROUNDUP(ABS(C1), 2)",
            "=LEFT(\"C1\", 3)",
            "=43",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    match target {
        ResultData::Float(f) => assert!(
            (f - 0.287202).abs() < 1e-3,
            "Expected ~0.287202 for A10, got {}",
            f
        ),
        other => panic!("Expected Float(~0.287202) for A10, got {:?}", other),
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
fn test_fuzz_or_zero_power_zero_num_error() {
    let sheet_src = [
        ["", "499.8781", "FALSE", "75", "-325.2142"],
        ["-69", "", "20", "\"F if\"", "-74"],
        ["7", "TRUE", "TRUE", "FALSE", ""],
        ["-42.681", "224", "\"LjESYnsO\"", "0", ""],
        ["\"IfnR\"", "301.17", "72", "88", "-95"],
        [
            "=UPPER(\"AVERAGE(D3:E3)\")",
            "=-24",
            "=OR(B5 > 0, (B2 ^ E3) < 100)",
            "=LEFT(\"OR(16 > 0, C5 < 100)\", 4)",
            "=AND((C5 ^ C1) > 0, A3 < 100)",
        ],
        ["=LEN(\"D5\")", "=49", "=-3", "=SUM(D6:D6)", "\"uRm3\""],
        [
            "=MIN(E2:E3)",
            "=(LOWER(\"C3\") / A5)",
            "=ABS(E5)",
            "=SUM(D3:E3)",
            "=D6",
        ],
        [
            "=MAX(B2:B5)",
            "=8",
            "=C1",
            "TRUE",
            "=(MAX(E7, 40) * ROUND(-34, 0))",
        ],
        [
            "=ROUND(D3, 2)",
            "=ROUNDUP(AND(-23 > 0, D6 < 100), 0)",
            "=B5",
            "=AND((21 / E3) > 0, C5 < 100)",
            "=INT(INT(B3))",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(5, 2));
    match target {
        ResultData::Boolean(b) => assert_eq!(b, true),
        ResultData::Error(ref e) => assert!(e.contains("#NUM!")),
        other => panic!("Expected Boolean(true) or #NUM!, got {:?}", other),
    }
}

#[test]
fn test_fuzz_if_max_int_num_error() {
    let sheet_src = [
        ["0", "TRUE", "-308.27", "1", ""],
        ["\"QhQ\"", "-74", "7", "-20", "-244.9764"],
        ["-2", "101.48", "67", "-85.1962", "-17"],
        ["-446.7525", "-10", "-387.5653", "29", "83.035"],
        ["-27", "-60", "\"RvfkX\"", "-18", ""],
        [
            "=IF((-28 > SUM(E5:E5)), (B2 * 46), IF((37 > D2), C4, A1))",
            "=IF((B1 > B3), OR(A4 > 0, -39 < 100), SQRT(C3))",
            "=(PRODUCT(A1:D4) / (D4 + A1))",
            "=A5",
            "=ROUNDDOWN(D2, 0)",
        ],
        [
            "=OR((B2 ^ B3) > 0, SUM(C5, 19) < 100)",
            "=A1",
            "TRUE",
            "=AND(11 > 0, ABS(E3) < 100)",
            "=AVERAGE(B5:E6)",
        ],
        [
            "=AND(C7 > 0, ABS(C3) < 100)",
            "=OR(-17 > 0, PRODUCT(45, 15) < 100)",
            "=AND(LEFT(\"-40\", 2) > 0, SUM(3, C6) < 100)",
            "=-20",
            "=39",
        ],
        [
            "=IF((MAX(4, E6) > INT(E7)), (E4 - A7), (41 / E6))",
            "=(CONCATENATE(\"C2\", \"A5\") / -32)",
            "=AVERAGE(A6:B6)",
            "=IF((SUM(D7, D7) > A4), ROUNDUP(E6, 0), (E3 + A6))",
            "=ABS(IF((9 > B6), E3, E4))",
        ],
        [
            "=C8",
            "=MAX(D1:D6)",
            "=B3",
            "=((-49 ^ B7) * -42)",
            "=IF((D3 > 4), A3, E4)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 0));
    match target {
        ResultData::Float(f) => assert!((f - 82.035).abs() < 1e-3),
        ResultData::Error(ref e) => assert!(e.contains("#NUM!")),
        other => panic!("Expected Float(82.035) or #NUM!, got {:?}", other),
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
fn test_fuzz_abs_min_range_boolean() {
    let sheet_src = [
        ["-46", "TRUE", "-38", "-340.83", "10"],
        ["", "\"JcSVrdA\"", "-65", "\"Il\"", "-221.6"],
        ["FALSE", "", "-110.62", "-19.28", "62"],
        ["294.2", "", "", "31.55", "-71"],
        ["FALSE", "FALSE", "\"3hW\"", "31", ""],
        [
            "=ROUNDUP(MIN(C2:C2), 1)",
            "=((E4 / B4) - C4)",
            "=OR(ABS(B2) > 0, D4 < 100)",
            "=ABS(D2)",
            "=B2",
        ],
        [
            "=UPPER(\"-44\")",
            "=E5",
            "=IF((IF((B2 > -34), B4, C6) > A1), SQRT(-5), LEFT(\"15\", 4))",
            "=MIN(A4:A6)",
            "=20",
        ],
        [
            "FALSE",
            "=ABS(MIN(C5:C5))",
            "=UPPER(\"ROUNDDOWN(A6, 1)\")",
            "=B5",
            "=MAX(A6:D7)",
        ],
        [
            "-64",
            "145",
            "=(14 - -15)",
            "=PRODUCT(B1:C7)",
            "=AVERAGE(D6:E7)",
        ],
        [
            "=ABS(D7)",
            "=SUM(C6:D8)",
            "=LEFT(\"E1\", 5)",
            "=E5",
            "=MIN(C4:C8)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    println!(
        "Seed 25929 evaluated target CellRef(9, 0) A10: {:?}",
        target
    );
    let d7 = sheet.get_result_data(&CellRef::new(6, 3));
    println!("Seed 25929 D7: {:?}", d7);
    let a6 = sheet.get_result_data(&CellRef::new(5, 0));
    println!("Seed 25929 A6: {:?}", a6);
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
fn test_fuzz_roundup_if_branch_evaluation() {
    let sheet_src = [
        ["0", "-26", "-163", "-156", "\"qy2RBN\""],
        ["", "40", "166.187", "-100", "\"Df1LC\""],
        ["", "-29", "FALSE", "FALSE", "\"zggq\""],
        ["10", "-381.4", "FALSE", "20", "TRUE"],
        ["FALSE", "-53", "177.04", "108.5", "\"gGKd GKH\""],
        [
            "=43",
            "=OR(CONCATENATE(\"A4\", \"-33\") > 0, SUM(B2:E4) < 100)",
            "=47",
            "-48",
            "=4",
        ],
        [
            "=AVERAGE(C6:C6)",
            "=INT(C2)",
            "=AND(ABS(E4) > 0, (-23 - A6) < 100)",
            "=(AVERAGE(A2:D2) + D3)",
            "=SQRT(ROUND(E2, 2))",
        ],
        ["37", "0", "=INT(ABS(B2))", "=D6", "=D5"],
        [
            "=B3",
            "=MIN(LEFT(\"A5\", 2), B7)",
            "=SUM(D7, MAX(D4:E7))",
            "=ROUND(IF((A8 > 23), E6, D4), 2)",
            "=B8",
        ],
        [
            "=MAX((44 ^ E6), AND(E9 > 0, B1 < 100))",
            "=(ABS(E5) + PRODUCT(C7:C8))",
            "=ROUNDUP(ABS(E8), 1)",
            "=A2",
            "=LOWER(\"D2\")",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target_c10 = sheet.get_result_data(&CellRef::new(9, 2));
    println!("Seed 892758 C10: {:?}", target_c10);
    let target_d9 = sheet.get_result_data(&CellRef::new(8, 3));
    println!("Seed 892758 D9: {:?}", target_d9);
}

#[test]
fn test_fuzz_if_power_overflow_num_error() {
    let sheet_src = [
        ["-3", "21.488", "", "", "\"1X1OB\""],
        ["-247.996", "\"XNfFdxbI\"", "4", "\"O\"", "\"VjE\""],
        ["9", "", "-39", "-76", "127.738"],
        ["-105.1", "35", "372.504", "133.7", "40"],
        ["TRUE", "TRUE", "", "FALSE", "402.9274"],
        [
            "=PRODUCT(A4:A5)",
            "=E1",
            "=OR(ROUND(B1, 1) > 0, CONCATENATE(\"-18\", \"47\") < 100)",
            "=-22",
            "=AVERAGE(C1:D4)",
        ],
        [
            "=D1",
            "=(ROUNDDOWN(E5, 2) + (C2 * B3))",
            "215.33",
            "=B2",
            "=IF(((A6 / D1) > AVERAGE(D3, C1)), D5, AVERAGE(C3:C4))",
        ],
        [
            "=ROUNDUP((-22 - A5), 2)",
            "=-3",
            "4",
            "=UPPER(\"OR(48 > 0, B1 < 100)\")",
            "=CONCATENATE(\"E7\", \"E5\")",
        ],
        ["FALSE", "=A1", "=A1", "", "\"ZDYF1\""],
        [
            "=(IF((D1 > A8), -12, C9) ^ ROUNDDOWN(B7, 0))",
            "=AVERAGE(C4, OR(D4 > 0, C8 < 100))",
            "=MIN(4, (-18 + -4))",
            "=OR(E2 > 0, C4 < 100)",
            "=SUM(A3:C4)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target_a10 = sheet.get_result_data(&CellRef::new(9, 0));
    println!("Seed 708047 A10: {:?}", target_a10);
    let d1 = sheet.get_result_data(&CellRef::new(0, 3));
    let a8 = sheet.get_result_data(&CellRef::new(7, 0));
    let b7 = sheet.get_result_data(&CellRef::new(6, 1));
    println!("Seed 708047 D1: {:?}, A8: {:?}, B7: {:?}", d1, a8, b7);
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
fn test_fuzz_if_sqrt_nested_branch() {
    let sheet_src = [
        ["6", "\"mdrrGl\"", "290.85", "FALSE", "-52"],
        ["\"ae\"", "-72", "0", "\"jib\"", "FALSE"],
        ["34", "12", "\"C\"", "FALSE", "29"],
        ["379.061", "\" Vn\"", "-101.701", "FALSE", "2"],
        ["-90", "-81", "\"YJPn\"", "2", "67"],
        [
            "=(IF((D2 > D2), D5, E3) / ABS(-38))",
            "=11",
            "=D3",
            "=OR((B4 - 22) > 0, C3 < 100)",
            "=PRODUCT(D4, D3)",
        ],
        [
            "=(MIN(B4, B1) ^ -1)",
            "=B3",
            "=SQRT(D2)",
            "=ROUNDUP(-10, 0)",
            "=LEFT(\"(D2 + D2)\", 5)",
        ],
        [
            "=UPPER(\"-16\")",
            "=E6",
            "=-38",
            "=MAX(B1:E7)",
            "=UPPER(\"MIN(C4, -37)\")",
        ],
        [
            "=PRODUCT(D7:D7)",
            "=IF((C1 > (-11 * E2)), (E4 + -13), LEFT(\"A2\", 4))",
            "=IF((OR(B7 > 0, D2 < 100) > (E2 + A6)), ABS(C4), LOWER(\"B8\"))",
            "=AND(B7 > 0, UPPER(\"D7\") < 100)",
            "=D5",
        ],
        [
            "=IF((E3 > SQRT(D9)), IF((-19 > -33), 45, 13), LEN(\"E5\"))",
            "=(IF((C7 > E5), -43, C7) + E9)",
            "=A9",
            "88",
            "=A7",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    println!("Seed 369614 evaluated target CellRef(9, 0): {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 45.0),
        other => panic!("Expected Float(45.0), got {:?}", other),
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
fn test_fuzz_int_if_branch_evaluation() {
    let sheet_src = [
        ["\"vMzcZU\"", "", "65", "TRUE", "55"],
        ["432", "41", "92", "FALSE", "-39"],
        ["\"yFkqp3\"", "2", "0", "", "-235.4"],
        ["11", "\"mVClvk\"", "-97", "-331.7", "45.2"],
        ["15", "52", "TRUE", "\"jvnuJ\"", "\"unJzzWMF\""],
        [
            "=E5",
            "=ROUNDDOWN(35, 1)",
            "=(A2 - AVERAGE(E5:E5))",
            "=AVERAGE(E1:E4)",
            "=SUM(A1:D2)",
        ],
        [
            "=E4",
            "=INT(IF((B6 > D5), A6, E6))",
            "=B3",
            "-81",
            "=((E1 - D6) / IF((C3 > A3), A3, E2))",
        ],
        ["=AVERAGE(E4:E6)", "=(MAX(A7, 49) + D7)", "=13", "", "-78"],
        [
            "=D8",
            "9",
            "=AND(A8 > 0, (-45 + E2) < 100)",
            "=1",
            "=SQRT(INT(E8))",
        ],
        [
            "=LEFT(\"A9\", 2)",
            "=((-19 * B1) / B4)",
            "=MIN(C1:D5)",
            "=AVERAGE(B7, E3)",
            "-83.43000000000001",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let b7 = sheet.get_result_data(&CellRef::new(6, 1));
    println!("Seed 672328 B7: {:?}", b7);
    let d10 = sheet.get_result_data(&CellRef::new(9, 3));
    println!("Seed 672328 D10: {:?}", d10);
    match b7 {
        ResultData::Float(f) => assert_eq!(f, 630.0),
        other => panic!("Expected Float(630.0), got {:?}", other),
    }
    match d10 {
        ResultData::Float(f) => assert_eq!(f, 197.3),
        other => panic!("Expected Float(197.3), got {:?}", other),
    }
}

#[test]
fn test_fuzz_sqrt_negative_e8_num_error() {
    let sheet_src = [
        ["-415", "2", "-31.7166", "-72", "-424.5495"],
        ["-169", "-1", "3", "\"zI\"", "370.6"],
        ["252.449", "-298.4798", "", "-22", "-77"],
        ["TRUE", "\"iR\"", "TRUE", "26", "\"11wh\""],
        ["-173.86", "437.575", "12", "TRUE", "\"aG\""],
        ["=LEN(\"(-41 + A5)\")", "63", "-179.314", "-236.94", "=D3"],
        [
            "=C2",
            "=AVERAGE(E6:E6)",
            "=OR(PRODUCT(B4, -33) > 0, AND(42 > 0, -13 < 100) < 100)",
            "=IF((28 > SQRT(-16)), ROUND(D4, 1), ROUNDDOWN(24, 1))",
            "=IF(((B5 * D1) > (A2 * -39)), MAX(33, 46), AVERAGE(-18, E3))",
        ],
        [
            "=AVERAGE(CONCATENATE(\"-27\", \"C4\"), A7)",
            "=IF((UPPER(\"B5\") > 32), A4, OR(D5 > 0, B2 < 100))",
            "=SUM((B3 + -48), -6)",
            "=ROUNDUP(IF((D2 > A1), E6, 13), 2)",
            "=SQRT((E1 - 18))",
        ],
        ["=INT(LEFT(\"-26\", 5))", "\"3PTBp\"", "=E6", "30", "=C2"],
        ["=E6", "=SQRT(E8)", "484", "=-29", "22.24"],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let b10 = sheet.get_result_data(&CellRef::new(9, 1));
    println!("Seed 585783 B10: {:?}", b10);
    let d8 = sheet.get_result_data(&CellRef::new(7, 3));
    println!("Seed 585783 D8: {:?}", d8);
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
fn test_fuzz_and_sqrt_error_precedence() {
    let sheet_src = [
        ["84", "", "31", "340.5032", ""],
        ["-373.2", "-12", "FALSE", "\"Bx\"", "-130.3"],
        ["95", "\"EZ\"", "\"xQSzxQK\"", "326.07", "\"M1Ahbv2p\""],
        ["10", "-86", "76", "57", "TRUE"],
        ["-484.77", "\"Nxkov\"", "-46", "", "401"],
        [
            "=(-5 / -32)",
            "=IF((IF((B2 > E5), 18, 41) > A1), ROUNDUP(E2, 0), 37)",
            "=B1",
            "=-28",
            "=PRODUCT(A4:A5)",
        ],
        [
            "=INT(AND(A6 > 0, A3 < 100))",
            "=(28 * D1)",
            "\"zndnANee\"",
            "\"UWL1l3 \"",
            "113.5114",
        ],
        [
            "=OR(D7 > 0, OR(B2 > 0, B6 < 100) < 100)",
            "\"kH\"",
            "=ROUND(PRODUCT(D6, E5), 2)",
            "=B2",
            "=CONCATENATE(\"15\", \"INT(A7)\")",
        ],
        [
            "0",
            "=(IF((E3 > E8), E1, C4) ^ 44)",
            "=(UPPER(\"37\") ^ PRODUCT(C5:C8))",
            "=(ROUND(20, 2) + UPPER(\"B7\"))",
            "24",
        ],
        [
            "=SQRT(C3)",
            "=IF((A5 > D8), -1, LEN(\"E6\"))",
            "=UPPER(\"MAX(C9, C2)\")",
            "=SUM(OR(-26 > 0, 28 < 100), IF((C5 > C6), B2, -48))",
            "=AND(SQRT(-39) > 0, SQRT(C7) < 100)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let e10 = sheet.get_result_data(&CellRef::new(9, 4));
    println!("Seed 632978 E10: {:?}", e10);
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
fn test_fuzz_if_sqrt_negative_cond_error() {
    let sheet_src = [
        ["104", "", "9", "5", ""],
        ["264", "-287", "", "81", "-38"],
        ["62", "TRUE", "\"jMg\"", "84", "\" uND\""],
        ["TRUE", "", "0", "2", ""],
        ["90", "327", "-32", "", "-399.3008"],
        [
            "=MIN(E5:E5)",
            "=C4",
            "=-12",
            "=IF((IF((B5 > A4), A5, E1) > RIGHT(\"B4\", 3)), B4, A1)",
            "=IF((IF((D4 > 26), -38, -33) > ABS(-7)), D3, (-15 ^ 15))",
        ],
        ["=SUM(A6:B6)", "=C5", "0", "-73", "=(A3 - E1)"],
        [
            "=E4",
            "=E4",
            "=OR((B6 / -30) > 0, 16 < 100)",
            "=(48 - D2)",
            "=E6",
        ],
        [
            "=IF((SQRT(C6) > LEN(\"-46\")), (45 / -28), MIN(D2, E5))",
            "=IF((37 > LEFT(\"-6\", 2)), PRODUCT(B7:B8), SUM(0, E1))",
            "470",
            "=-19",
            "TRUE",
        ],
        ["=PRODUCT(B1:C3)", "=B7", "=SQRT(10)", "0", "=D7"],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 0));
    println!("Seed 558820 target: {:?}", target);
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
fn test_fuzz_if_rounddown_branch() {
    let sheet_src = [
        ["10", "-51", "73", "-50", "-62"],
        ["4", "TRUE", "50", "-340", "60"],
        ["", "40.26", "FALSE", "\"Xm\"", "-98"],
        ["0", "16", "5", "-80", "-47.3"],
        ["\"H1Mbf\"", "-445.9098", "-265.8", "-82", "-86"],
        [
            "=D2",
            "=7",
            "=RIGHT(\"E1\", 4)",
            "=43",
            "=(AND(B3 > 0, -22 < 100) + LEN(\"D3\"))",
        ],
        [
            "=(A6 ^ -30)",
            "=MAX(A5:D6)",
            "=30",
            "=(C6 + (A5 ^ -24))",
            "=AVERAGE(A5:E6)",
        ],
        [
            "=ROUNDDOWN(C5, 0)",
            "=E4",
            "=IF((C3 > ROUNDDOWN(E6, 2)), D5, INT(A6))",
            "=C6",
            "=(16 + B3)",
        ],
        [
            "=(MIN(E7:E8) - LOWER(\"C4\"))",
            "=IF((-50 > -18), -16, SUM(E7:E8))",
            "TRUE",
            "-21",
            "=(MAX(E7, A8) * PRODUCT(A5, C1))",
        ],
        [
            "=((46 + B6) ^ (D2 - -4))",
            "=A1",
            "=MIN(D9:E9)",
            "=AVERAGE(D1:E6)",
            "=(LEFT(\"C4\", 4) / E6)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 2));
    println!("Seed 158982 target: {:?}", target);
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
fn test_fuzz_abs_if_c6_cond() {
    let sheet_src = [
        ["-48", "2", "-292", "4", "76"],
        ["", "TRUE", "-55.17", "-36", "-58"],
        ["\"Id\"", "-4", "", "59", "8"],
        ["FALSE", "59", "FALSE", "386.088", ""],
        ["\"Ow3PoM\"", "81", "64", "FALSE", "FALSE"],
        ["=(-4 + LEFT(\"-21\", 4))", "79", "=-19", "=E3", "51"],
        [
            "=D6",
            "=AVERAGE(29, LEN(\"C2\"))",
            "=MIN(OR(A4 > 0, B1 < 100), SQRT(E2))",
            "\"rpKZse\"",
            "=LEN(\"IF((1 > -37), -11, E3)\")",
        ],
        [
            "=D5",
            "=ABS(IF((C6 > 5), E4, A7))",
            "=MAX(E4:E6)",
            "=CONCATENATE(\"A6\", \"MIN(B4, E3)\")",
            "=D4",
        ],
        [
            "=ROUNDUP(15, 1)",
            "=OR(B6 > 0, ROUNDDOWN(B4, 0) < 100)",
            "=B6",
            "FALSE",
            "=D4",
        ],
        [
            "=AVERAGE(D4:E8)",
            "=D9",
            "=((-25 / A2) * C6)",
            "=ABS(ROUND(15, 0))",
            "TRUE",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 1));
    println!("Seed 767352 target: {:?}", target);
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
fn test_fuzz_round_b6_boolean_ref() {
    let sheet_src = [
        ["\"IYpej\"", "12", "-30", "-53", "357.51"],
        ["\"W\"", "-57", "65", "-398.6184", "0"],
        ["1", "FALSE", "FALSE", "27", "TRUE"],
        ["FALSE", "95", "4", "TRUE", "-75"],
        ["-74", "-34", "203", "TRUE", "82"],
        [
            "=LOWER(\"A5\")",
            "=B3",
            "=AVERAGE(C1, 6)",
            "=B2",
            "=(D3 ^ 39)",
        ],
        [
            "=LEN(\"(E6 + E5)\")",
            "=(IF((18 > B6), A3, -40) + (-14 / A2))",
            "=SQRT(LEN(\"A5\"))",
            "=(-13 + LOWER(\"D5\"))",
            "\"Bf\"",
        ],
        [
            "=ROUND(B6, 1)",
            "=-33",
            "=LEFT(\"MIN(D5:E7)\", 3)",
            "=D5",
            "=PRODUCT(D5:D7)",
        ],
        [
            "0",
            "=IF((UPPER(\"11\") > (E5 ^ C4)), 50, E5)",
            "=(E6 / LOWER(\"C2\"))",
            "=D5",
            "=D2",
        ],
        [
            "88",
            "=IF((C5 > A9), (C5 * E6), ROUNDDOWN(A2, 0))",
            "-21",
            "=C3",
            "=PRODUCT(B2, D7)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 0));
    println!("Seed 734641 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 0.0),
        ResultData::Integer(i) => assert_eq!(i, 0),
        other => panic!("Expected 0, got {:?}", other),
    }
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
fn test_fuzz_if_nested_product_b4_d4() {
    let sheet_src = [
        ["486.89", "\"Eaf\"", "50", "-374.4602", "-124.2236"],
        ["-315", "TRUE", "", "96", ""],
        ["94", "-455.2", "56.703", "", "5"],
        ["\"klN\"", "55", "FALSE", "4", "-69"],
        ["\"adEPPnfK\"", "442.7", "65.8", "-370", ""],
        ["=ABS(ROUNDUP(A2, 2))", "=ROUND(0, 1)", "=-4", "=B3", "=-43"],
        [
            "=ABS(A6)",
            "121",
            "=IF((ROUNDDOWN(D3, 0) > ROUND(A3, 0)), D5, MIN(B3:D3))",
            "=-26",
            "=40",
        ],
        [
            "=IF((IF((B5 > E7), E2, C4) > ROUND(E7, 0)), IF((A7 > -21), B7, 47), PRODUCT(B4:D4))",
            "225.89",
            "TRUE",
            "=E5",
            "=-32",
        ],
        [
            "=CONCATENATE(\"SQRT(E8)\", \"E1\")",
            "=UPPER(\"OR(C4 > 0, D3 < 100)\")",
            "=10",
            "=RIGHT(\"C5\", 2)",
            "=C2",
        ],
        [
            "=E5",
            "=(PRODUCT(A3:E5) / D6)",
            "=SUM(MIN(A3, 19), (-4 / D5))",
            "=IF((D3 > C4), CONCATENATE(\"A5\", \"6\"), -23)",
            "=ABS(B8)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 0));
    println!("Seed 92353 target: {:?}", target);
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
fn test_fuzz_if_subtraction_rounddown_condition() {
    let sheet_src = [
        ["", "-53", "", "-24", "88"],
        ["FALSE", "-290", "17", "8", "-267.6631"],
        ["-125.7", "\"Up23XW\"", "-239.8602", "0", "39"],
        ["94", "74", "TRUE", "81", "-7"],
        ["-129.158", "\"WenDvo\"", "FALSE", "-12", "-107.343"],
        [
            "-88",
            "=AND(-21 > 0, AVERAGE(E3:E4) < 100)",
            "=LEN(\"D1\")",
            "=D5",
            "15",
        ],
        [
            "=D6",
            "210.95",
            "=RIGHT(\"B3\", 1)",
            "=AND(E2 > 0, IF((D1 > E5), -30, E5) < 100)",
            "=((-16 - A5) * D4)",
        ],
        [
            "=ABS(OR(D4 > 0, B7 < 100))",
            "=MAX((D5 + 0), RIGHT(\"A1\", 2))",
            "=C7",
            "=UPPER(\"MIN(E1:E7)\")",
            "=SUM(C1:D7)",
        ],
        [
            "=SUM(E2:E6)",
            "=ABS((E4 * A8))",
            "=IF(((A3 - A7) > ROUNDDOWN(A7, 0)), IF((47 > B1), 48, B5), A4)",
            "=A2",
            "=OR(IF((C4 > A1), D6, D8) > 0, A3 < 100)",
        ],
        [
            "=CONCATENATE(\"B2\", \"(-26 * D6)\")",
            "=10",
            "=D6",
            "=21",
            "=SUM(A1, IF((D6 > E6), B3, -48))",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 2));
    println!("Seed 354063 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 94.0),
        ResultData::Integer(i) => assert_eq!(i, 94),
        other => panic!("Expected 94, got {:?}", other),
    }
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
fn test_fuzz_rounddown_if_empty_cell_ref() {
    let sheet_src = [
        ["-264.7", "94", "-328", "-9", "146.41"],
        ["-25", "157.337", "-67", "-12", "0"],
        ["-366.201", "99", "-49", "9", "5"],
        ["", "", "\"E\"", "5", "26"],
        ["", "83", "TRUE", "3", "174"],
        [
            "=A1",
            "282",
            "=UPPER(\"-31\")",
            "=IF((-38 > PRODUCT(-18, A1)), B2, 36)",
            "=PRODUCT(A4:B5)",
        ],
        ["=42", "=(C2 / IF((3 > -12), D6, -18))", "=B4", "=19", "=E5"],
        [
            "=(MAX(D2, B1) - (D6 + C5))",
            "=E5",
            "=ROUNDUP(IF((E7 > E7), C7, -40), 1)",
            "-86",
            "=CONCATENATE(\"E1\", \"A4\")",
        ],
        [
            "=ROUNDDOWN(IF((B8 > B5), C7, 35), 0)",
            "=40",
            "=UPPER(\"OR(D6 > 0, B5 < 100)\")",
            "=SUM(RIGHT(\"B5\", 5), B8)",
            "=50",
        ],
        [
            "=(IF((-21 > -46), A7, B8) * LEN(\"-28\"))",
            "54",
            "=(A8 + AVERAGE(E9:E9))",
            "=AND((D3 ^ E8) > 0, ROUNDUP(D5, 2) < 100)",
            "=SQRT(ABS(C5))",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 0));
    println!("Seed 101079 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 0.0),
        ResultData::Integer(i) => assert_eq!(i, 0),
        other => panic!("Expected 0, got {:?}", other),
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
fn test_fuzz_if_d4_division_abs_b7() {
    let sheet_src = [
        ["-47", "-6", "77", "\"reWs\"", "7"],
        ["-392.33", "-96", "9", "479.49", "-291.13"],
        ["TRUE", "211.2", "4", "13", "FALSE"],
        ["FALSE", "393", "-47", "-36", "TRUE"],
        ["\"zgPkYl\"", "TRUE", "\"aEL\"", "", "330.8"],
        [
            "=CONCATENATE(\"IF((30 > B4), 13, E1)\", \"C4\")",
            "=SQRT(AND(C4 > 0, A5 < 100))",
            "=SQRT(6)",
            "=MIN(D3:D4)",
            "FALSE",
        ],
        [
            "=CONCATENATE(\"C2\", \"INT(48)\")",
            "=E1",
            "=LEFT(\"ROUNDUP(-17, 0)\", 4)",
            "=(E2 + (-47 / B3))",
            "=CONCATENATE(\"D4\", \"SQRT(36)\")",
        ],
        [
            "=SUM(A7:A7)",
            "=PRODUCT(1, ROUNDUP(25, 2))",
            "=IF((LOWER(\"C1\") > E3), IF((-44 > E1), D7, A2), D6)",
            "=A6",
            "=A5",
        ],
        [
            "=IF((D4 > (-41 / 14)), SUM(D4:E8), ABS(B7))",
            "=(LEFT(\"C4\", 2) * PRODUCT(C8:C8))",
            "-41",
            "=CONCATENATE(\"A7\", \"D3\")",
            "=B4",
        ],
        [
            "-16",
            "=ROUNDDOWN(38, 0)",
            "=(IF((C3 > C1), 50, B6) * 17)",
            "=(IF((E5 > D3), B5, 38) + (-34 - -47))",
            "\"Bj1MXyY\"",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 0));
    println!("Seed 273076 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 7.0),
        ResultData::Integer(i) => assert_eq!(i, 7),
        other => panic!("Expected 7, got {:?}", other),
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
fn test_fuzz_addition_sqrt_negative_num_error() {
    let sheet_src = [
        ["-17", "26", "-100", "\"l\"", "-41"],
        ["", "9", "46", "-175.65", ""],
        ["-27", "", "100", "18.239", "5"],
        ["-141.0941", "\"BnXOprLe\"", "-304.064", "FALSE", "-79"],
        ["-93", "", "\"PI\"", "-66", "-188"],
        [
            "=LOWER(\"AND(40 > 0, -33 < 100)\")",
            "=(SQRT(6) ^ AND(16 > 0, A4 < 100))",
            "=AND(AND(A1 > 0, A3 < 100) > 0, IF((B1 > C4), E2, 9) < 100)",
            "=RIGHT(\"AND(B2 > 0, E1 < 100)\", 1)",
            "=AND(IF((D3 > C2), E5, A4) > 0, (A1 * -24) < 100)",
        ],
        [
            "=C1",
            "=PRODUCT(8, C3)",
            "=-30",
            "=MAX(A3:B4)",
            "=(LEFT(\"E4\", 3) ^ E1)",
        ],
        ["=1", "=C7", "=A6", "=D1", "=C4"],
        [
            "=ROUNDUP(RIGHT(\"C5\", 2), 0)",
            "=E5",
            "-142.96",
            "=IF((MIN(A6:D6) > -24), LEN(\"C8\"), LEN(\"E2\"))",
            "=C4",
        ],
        [
            "=(D2 + SQRT(B8))",
            "=(47 * UPPER(\"E4\"))",
            "=-47",
            "=(IF((43 > 6), -48, 0) * SQRT(-4))",
            "=C4",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    println!("Seed 964755 target: {:?}", target);
    match target {
        ResultData::Error(e) => assert_eq!(e, "#NUM!"),
        other => panic!("Expected #NUM!, got {:?}", other),
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
fn test_fuzz_rounddown_e6_boolean_ref() {
    let sheet_src = [
        ["-281", "10", "", "-177", "\"WOb\""],
        ["", "-153", "58", "\"v\"", "FALSE"],
        ["-353", "\"OJrt\"", "95", "117", "TRUE"],
        ["89", "12", "219.818", "4", "70.5967"],
        ["\"g2OgFkx\"", "0", "9", "98", "475"],
        [
            "=MIN((A1 / D2), B1)",
            "56",
            "=6",
            "=(19 ^ -22)",
            "=AND(LEN(\"2\") > 0, 31 < 100)",
        ],
        ["=C4", "=ROUND(PRODUCT(-5, D2), 0)", "=-36", "3", "=B2"],
        [
            "=A4",
            "=MIN(D6:D6)",
            "=((C1 - D4) - B2)",
            "=IF((LEN(\"C5\") > RIGHT(\"E2\", 2)), ROUND(43, 0), (B2 ^ C4))",
            "=AND(AND(C7 > 0, A4 < 100) > 0, A5 < 100)",
        ],
        [
            "=UPPER(\"MAX(B8, D3)\")",
            "=ROUNDDOWN(E6, 2)",
            "=SUM(INT(E2), ROUND(-39, 2))",
            "=LOWER(\"A2\")",
            "=A3",
        ],
        [
            "=C4",
            "",
            "=CONCATENATE(\"30\", \"OR(E7 > 0, C2 < 100)\")",
            "=12",
            "=ROUNDDOWN(IF((A2 > A7), A2, B3), 0)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 1));
    println!("Seed 442593 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 1.0),
        ResultData::Integer(i) => assert_eq!(i, 1),
        other => panic!("Expected 1, got {:?}", other),
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
fn test_fuzz_roundup_if_val() {
    let sheet_src = [
        ["\"WTcuBFQn\"", "\"v3EwtV\"", "400.724", "-10", "TRUE"],
        ["55.3559", "-23", "61", "\"yflJCHu\"", "-55"],
        ["", "2", "92", "\"L\"", "-84"],
        ["6", "0", "332", "30", ""],
        ["\"HIAZyhx\"", "\"hzAZbX\"", "0", "\"C\"", "\"xnBd\""],
        [
            "=CONCATENATE(\"A4\", \"B2\")",
            "=SQRT(MAX(B4:D5))",
            "=C4",
            "=A5",
            "=UPPER(\"45\")",
        ],
        [
            "=PRODUCT((27 * A1), D6)",
            "=35",
            "=RIGHT(\"ROUND(22, 0)\", 3)",
            "FALSE",
            "=E2",
        ],
        [
            "=D5",
            "=LOWER(\"-23\")",
            "=((28 ^ D5) ^ -5)",
            "=B7",
            "=SUM(INT(-14), D5)",
        ],
        [
            "=RIGHT(\"MIN(E3, A5)\", 2)",
            "=39",
            "=MIN(A3:D6)",
            "=OR(A7 > 0, (-22 - C7) < 100)",
            "=UPPER(\"MAX(14, -26)\")",
        ],
        [
            "=ROUNDUP(IF((E3 > E2), C6, E6), 1)",
            "=48",
            "=MIN(A7:E8)",
            "=(C2 / A9)",
            "-80",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    println!("Seed 719875 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 45.0),
        ResultData::Integer(i) => assert_eq!(i, 45),
        other => panic!("Expected 45, got {:?}", other),
    }
}

#[test]
fn test_fuzz_if_rounddown_cond_val() {
    let sheet_src = [
        ["-104.5", "61", "-48", "-265", "TRUE"],
        ["-8", "433", "-66", "FALSE", "343"],
        ["3", "-3", "415.555", "0", "388.234"],
        ["87", "FALSE", "-41", "-31", "-22"],
        ["0", "99", "-263", "\"S2ddn2p\"", "-168.413"],
        [
            "=B5",
            "=((E3 ^ D5) * (C4 - A2))",
            "=AND(OR(B5 > 0, B4 < 100) > 0, (B2 - B5) < 100)",
            "=B4",
            "=-42",
        ],
        [
            "=IF((ROUNDDOWN(E6, 2) > 50), D1, -5)",
            "=ROUNDUP(B1, 0)",
            "=PRODUCT(E4:E6)",
            "=ROUND(RIGHT(\"C5\", 1), 0)",
            "\"mjY2Pekq\"",
        ],
        [
            "=UPPER(\"ABS(-20)\")",
            "=1",
            "=(-24 + (42 + A5))",
            "=ROUNDUP(INT(C2), 1)",
            "=A1",
        ],
        [
            "=ROUND(SQRT(1), 1)",
            "=ABS(MIN(B6:D6))",
            "=IF((A5 > SUM(E4, D6)), -23, C1)",
            "=(ROUNDDOWN(C5, 1) * OR(E3 > 0, E5 < 100))",
            "=ABS(D7)",
        ],
        [
            "=C5",
            "=IF((RIGHT(\"C2\", 4) > B7), A6, E7)",
            "=IF((ROUNDUP(38, 2) > INT(B8)), C9, UPPER(\"-19\"))",
            "=A2",
            "TRUE",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 0));
    println!("Seed 439684 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, -5.0),
        ResultData::Integer(i) => assert_eq!(i, -5),
        other => panic!("Expected -5, got {:?}", other),
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
fn test_fuzz_sqrt_if_cond_val() {
    let sheet_src = [
        ["82", "486.7", "80", "TRUE", "95"],
        ["393.634", "-306.052", "7", "61", "-293"],
        ["\"EVUOZ\"", "43", "", "\"R\"", "3"],
        ["249", "TRUE", "-445.67", "TRUE", "-42"],
        ["", "50", "73", "-57", "TRUE"],
        [
            "=14",
            "=AND(SUM(-26, 28) > 0, MIN(D4, 27) < 100)",
            "=(D4 + (-1 / B5))",
            "=AVERAGE(D2:D4)",
            "=1",
        ],
        [
            "=OR(ROUNDUP(B4, 2) > 0, C6 < 100)",
            "\"l\"",
            "=(A6 + D4)",
            "=AVERAGE(INT(C3), UPPER(\"B5\"))",
            "=40",
        ],
        [
            "=LOWER(\"B5\")",
            "=SQRT(B6)",
            "=C6",
            "=(ROUNDDOWN(B4, 0) ^ C1)",
            "=UPPER(\"E4\")",
        ],
        [
            "=SUM((E1 + C2), E5)",
            "-92",
            "=48",
            "=LEN(\"15\")",
            "=(C2 / AND(D5 > 0, E5 < 100))",
        ],
        [
            "=LEFT(\"OR(-28 > 0, A7 < 100)\", 1)",
            "=SQRT(IF((43 > 28), D6, C7))",
            "314.56",
            "-89",
            "4",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 1));
    println!("Seed 233700 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert!((f - 7.810249675906654).abs() < 1e-4),
        ResultData::Integer(i) => assert_eq!(i, 7),
        other => panic!("Expected 7.810249675906654, got {:?}", other),
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
fn test_fuzz_if_sqrt_range_cond_val() {
    let sheet_src = [
        ["91", "\"U\"", "11", "57", "TRUE"],
        ["\"DdRE\"", "-27", "0", "-73.229", "-50"],
        ["\"EY 1k\"", "FALSE", "34", "-100", "60"],
        ["\"Kq\"", "27", "228.9", "1", "-21"],
        ["-254.6356", "-42", "TRUE", "-62", "\"pCEeIM\""],
        [
            "=MIN(C1:E4)",
            "=(AND(17 > 0, C5 < 100) * IF((-3 > E1), E2, A3))",
            "=B5",
            "=C4",
            "=OR(-26 > 0, ROUNDDOWN(A2, 1) < 100)",
        ],
        ["=-38", "=C6", "24", "TRUE", "=D5"],
        [
            "=22",
            "=AND(IF((44 > -16), C3, E7) > 0, INT(-36) < 100)",
            "=ABS(E4)",
            "=-29",
            "=C1",
        ],
        ["-85", "=-46", "=C1", "=LEN(\"PRODUCT(A3:A3)\")", "242.9"],
        [
            "=ROUND(UPPER(\"B9\"), 0)",
            "=IF((SQRT(C8) > E3), A2, 17)",
            "=ROUND(C5, 0)",
            "-406",
            "=0",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 1));
    println!("Seed 839355 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 17.0),
        ResultData::Integer(i) => assert_eq!(i, 17),
        other => panic!("Expected 17, got {:?}", other),
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
fn test_fuzz_division_if_branch_val() {
    let sheet_src = [
        ["", "77", "-57", "14", "\"Rj\""],
        ["47", "\"LLu\"", "", "\"ljDQrJ\"", "\"slrI\""],
        ["", "\"gXjEz\"", "", "-415.525", "-52"],
        ["10", "-68", "", "-17", "\"J\""],
        ["53", "16", "289.77", "-82", "24.2276"],
        [
            "=-46",
            "=(B5 - A4)",
            "=-45",
            "=ABS(SUM(D3:E3))",
            "=LEN(\"A2\")",
        ],
        [
            "=(ABS(D6) / IF((B5 > E6), -22, D5))",
            "=-39",
            "=17",
            "=OR(SQRT(27) > 0, MIN(C3:E6) < 100)",
            "=OR((43 ^ 24) > 0, C1 < 100)",
        ],
        [
            "=33",
            "=D1",
            "=SUM(MIN(A6:A6), (13 / E6))",
            "=3",
            "=ABS(ROUND(E3, 0))",
        ],
        [
            "=AND(SQRT(-42) > 0, -25 < 100)",
            "=(-49 + OR(D1 > 0, D5 < 100))",
            "=SUM(C8:E8)",
            "=((-25 * B1) ^ MIN(A4, B5))",
            "=((E3 / A2) / C4)",
        ],
        [
            "=ROUNDUP((E6 - -31), 2)",
            "=MIN(E2:E8)",
            "=OR(E9 > 0, UPPER(\"D1\") < 100)",
            "=17",
            "=D6",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 0));
    println!("Seed 879868 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert!((f - (-21.251136363636363)).abs() < 1e-4),
        other => panic!("Expected -21.251136363636363, got {:?}", other),
    }
}

#[test]
fn test_fuzz_if_rounddown_boolean_subtraction_val() {
    let sheet_src = [
        ["\"dZGZ\"", "-50", "4", "85", "67"],
        ["-41", "TRUE", "", "67", "-75"],
        ["459.99", "6", "\"O1\"", "21", "FALSE"],
        ["-66", "132", "-97", "44", "FALSE"],
        ["-466.6791", "\"N\"", "-41", "118.8", "\"mLFdskj\""],
        [
            "=-43",
            "=ROUNDUP(SQRT(E3), 1)",
            "=MAX(C5:C5)",
            "86",
            "=LEFT(\"E5\", 3)",
        ],
        [
            "=LEFT(\"(D3 / E5)\", 5)",
            "183",
            "=AND((A5 + A3) > 0, IF((C4 > -41), -5, C2) < 100)",
            "-39",
            "=E1",
        ],
        [
            "=PRODUCT(D2:D7)",
            "=AND(B1 > 0, (E1 * E1) < 100)",
            "=C2",
            "=B6",
            "=E2",
        ],
        [
            "=C3",
            "=(E1 + A3)",
            "=10",
            "=AND(UPPER(\"E2\") > 0, D8 < 100)",
            "=(MAX(A8:D8) * D2)",
        ],
        [
            "=-18",
            "-143",
            "=ROUNDDOWN(OR(-26 > 0, 34 < 100), 2)",
            "=IF((A4 > ROUNDDOWN(E8, 1)), (E3 - E8), B2)",
            "=OR(37 > 0, MIN(C4:E6) < 100)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 3));
    println!("Seed 353982 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 75.0),
        ResultData::Integer(i) => assert_eq!(i, 75),
        other => panic!("Expected 75, got {:?}", other),
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
fn test_fuzz_if_sqrt_roundup_empty_cell_ref() {
    let sheet_src = [
        ["-213.0018", "19", "0", "-308.3751", "58"],
        ["", "-97", "", "", ""],
        ["\"oMNfOA\"", "\"BkwUOw\"", "12", "-47", "TRUE"],
        ["-26", "7", "-5", "-485.31", "-335.3205"],
        ["-60", "0", "", "-11", "24"],
        [
            "=5",
            "=ROUNDDOWN((E4 * 3), 2)",
            "=AND(AND(-15 > 0, A4 < 100) > 0, 11 < 100)",
            "FALSE",
            "=C4",
        ],
        ["=E3", "=D3", "=C2", "=B2", "3"],
        [
            "=C7",
            "=10",
            "=RIGHT(\"INT(-24)\", 5)",
            "=A3",
            "=ROUNDUP(C5, 1)",
        ],
        ["=INT((C1 * -39))", "=B4", "=LOWER(\"A6\")", "FALSE", "=E6"],
        [
            "=IF((E7 > 20), (21 * C5), SQRT(E8))",
            "=PRODUCT(LEFT(\"D5\", 5), (C2 - -29))",
            "=-8",
            "=B7",
            "=LEFT(\"C7\", 2)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    println!("Seed 829563 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 0.0),
        ResultData::Integer(i) => assert_eq!(i, 0),
        other => panic!("Expected 0, got {:?}", other),
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
fn test_fuzz_if_average_range_empty_cell_ref() {
    let sheet_src = [
        ["14", "\"qu RkdT\"", "", "\"pZ\"", "-41"],
        ["-64.7", "-94", "\"scJWUwM\"", "TRUE", "\"ii\""],
        ["\"RjekSLLy\"", "37", "223.7869", "-56", "TRUE"],
        ["89", "177.6816", "-51", "", "0"],
        ["", "-399.73", "-392.5864", "", "-20"],
        [
            "=UPPER(\"IF((E2 > D3), E5, 32)\")",
            "-82",
            "=OR(OR(B1 > 0, E2 < 100) > 0, SUM(A4, 38) < 100)",
            "=9",
            "=E4",
        ],
        [
            "=OR(ROUNDUP(C1, 0) > 0, D5 < 100)",
            "=15",
            "=B4",
            "=CONCATENATE(\"IF((1 > A2), B3, E3)\", \"PRODUCT(E6:E6)\")",
            "=A2",
        ],
        [
            "=IF((ABS(E7) > D1), AVERAGE(C6:C6), SUM(A7:A7))",
            "=MAX(B5:B6)",
            "=OR(SUM(12, 2) > 0, LEN(\"C5\") < 100)",
            "=(-13 / B4)",
            "TRUE",
        ],
        [
            "=LEFT(\"SUM(44, D1)\", 3)",
            "\"GdRInUq\"",
            "=IF((ROUNDDOWN(B8, 1) > E6), C8, (14 + A2))",
            "=AVERAGE(E4:E8)",
            "=IF((D5 > C3), SUM(A7:A7), B1)",
        ],
        [
            "=CONCATENATE(\"AND(E3 > 0, C1 < 100)\", \"AVERAGE(C7:D7)\")",
            "=IF((-21 > D7), LEN(\"17\"), OR(46 > 0, D4 < 100))",
            "=-34",
            "=B1",
            "=MIN(A5, UPPER(\"-6\"))",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 0));
    println!("Seed 234176 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 0.0),
        ResultData::Integer(i) => assert_eq!(i, 0),
        other => panic!("Expected 0, got {:?}", other),
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
fn test_fuzz_abs_div_by_zero_subtraction_error() {
    let sheet_src = [
        ["0", "14", "-424", "FALSE", "-22"],
        ["", "319.113", "22", "FALSE", "215.3"],
        ["-71", "-247.8", "-384.719", "", "89"],
        ["6", "-88", "\"ccZu2\"", "227.4153", "-490.9"],
        ["167.2", "-76", "54", "-76", "-73"],
        [
            "=LEN(\"-30\")",
            "=48",
            "=ABS(C1)",
            "=C1",
            "=AND(IF((A2 > A1), B2, E5) > 0, (D2 ^ D2) < 100)",
        ],
        [
            "=IF((LEN(\"B3\") > RIGHT(\"E6\", 1)), SQRT(A6), (-17 * E5))",
            "=AVERAGE(LOWER(\"B5\"), 24)",
            "=MIN(C4:C5)",
            "=LOWER(\"45\")",
            "=-21",
        ],
        ["1", "=B5", "", "\"WInX2\"", "=(E1 / D2)"],
        ["=C1", "=A5", "=(ABS(E8) - A6)", "=C4", "TRUE"],
        [
            "=ABS(B1)",
            "=16",
            "=RIGHT(\"B6\", 2)",
            "=INT(ABS(B3))",
            "=C4",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 2));
    println!("Seed 817752 target: {:?}", target);
    match target {
        ResultData::Error(e) => assert_eq!(e, "#DIV/0!"),
        other => panic!("Expected #DIV/0!, got {:?}", other),
    }
}

#[test]
fn test_fuzz_negative_base_huge_exponent_num_error() {
    let sheet_src = [
        ["-41", "", "199.663", "-421.675", ""],
        ["17", "-36.3925", "316.534", "-30", "TRUE"],
        ["FALSE", "", "TRUE", "6", "154"],
        ["\"2Lpn pQ\"", "-91", "-28", "\"G\"", "-73"],
        ["-84", "64", "", "-37", "105.398"],
        [
            "=(E3 - INT(E5))",
            "10",
            "=C2",
            "-94",
            "=AND(MIN(B1:C1) > 0, B4 < 100)",
        ],
        [
            "=((B1 - A6) + ABS(D3))",
            "=A6",
            "=C6",
            "=LEN(\"E3\")",
            "=PRODUCT(D1:E6)",
        ],
        [
            "=UPPER(\"D2\")",
            "=45",
            "=D5",
            "=ROUNDDOWN(UPPER(\"E7\"), 1)",
            "=SQRT(27)",
        ],
        [
            "=D4",
            "=E3",
            "=AND(IF((C6 > A8), D2, -43) > 0, AND(C5 > 0, A6 < 100) < 100)",
            "=ROUND(B4, 2)",
            "=LEN(\"C7\")",
        ],
        ["=D5", "=(D4 * B5)", "=E4", "=(-26 ^ E7)", "\"ptjbR1\""],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 3));
    println!("Seed 128903 target: {:?}", target);
    match target {
        ResultData::Error(e) => assert_eq!(e, "#NUM!"),
        other => panic!("Expected #NUM!, got {:?}", other),
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
fn test_fuzz_if_sqrt_average_range_num_error() {
    let sheet_src = [
        ["FALSE", "\"pNOEESaj\"", "94", "0", "\"dp3sTYc\""],
        ["-344.363", "", "\"zyNCSI\"", "6", "44"],
        ["58.4", "37", "0", "\"nNgEFl\"", "5"],
        ["-295.431", "-240.2", "19", "\"IxGnv3u\"", "25"],
        ["", "-74", "2", "-87", "-84"],
        [
            "=OR(B4 > 0, ROUNDUP(-19, 2) < 100)",
            "-352.4",
            "94",
            "=AVERAGE(C2:E5)",
            "=ROUNDUP(ABS(D1), 1)",
        ],
        ["-59", "63", "FALSE", "=PRODUCT(B6:B6)", "=B4"],
        [
            "=IF((SQRT(D6) > C2), (A1 / -23), IF((D5 > E5), -6, A7))",
            "=SQRT(-6)",
            "=SQRT(30)",
            "=(ROUND(20, 0) + B3)",
            "=ABS(34)",
        ],
        [
            "312.486",
            "=ROUNDUP((17 + E7), 2)",
            "=AND(CONCATENATE(\"C5\", \"43\") > 0, B2 < 100)",
            "=(A8 * A7)",
            "-2",
        ],
        [
            "=INT(SQRT(B6))",
            "=25",
            "=(IF((E1 > A5), A9, -42) * UPPER(\"D5\"))",
            "=(IF((-17 > B5), E5, B1) * -14)",
            "=ROUND(UPPER(\"D8\"), 2)",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 0));
    println!("Seed 631817 target: {:?}", target);
    match target {
        ResultData::Error(e) => assert_eq!(e, "#NUM!"),
        other => panic!("Expected #NUM!, got {:?}", other),
    }
}

#[test]
fn test_fuzz_sqrt_average_negative_num_error() {
    let sheet_src = [
        ["478.67", "-79", "-63", "-95", "33"],
        ["51", "-99", "\"f\"", "\"yEF\"", "-14"],
        ["TRUE", "\"Vd\"", "", "-150.0543", "-496.7"],
        ["TRUE", "477.525", "105.9347", "TRUE", ""],
        ["FALSE", "TRUE", "-38", "8", "81"],
        [
            "=A3",
            "=MAX(B5:C5)",
            "-394.489",
            "=IF((E5 > IF((B4 > 47), 2, B5)), D3, (D4 - E5))",
            "=AND(D3 > 0, IF((D3 > E1), 28, 29) < 100)",
        ],
        [
            "=(23 * RIGHT(\"E6\", 4))",
            "=D3",
            "=SQRT(D6)",
            "=IF(((A1 + E3) > (-15 + -26)), ROUNDUP(-43, 1), B4)",
            "=C6",
        ],
        ["=C6", "=B2", "=28", "=MAX(LEN(\"3\"), C6)", "=5"],
        [
            "=B5",
            "=AVERAGE(OR(E7 > 0, D5 < 100), IF((C6 > D6), 41, B6))",
            "=E8",
            "=LEN(\"30\")",
            "=SUM(D2:D3)",
        ],
        [
            "=LOWER(\"(C1 / B8)\")",
            "TRUE",
            "=34",
            "=LEFT(\"INT(-5)\", 4)",
            "=C4",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 2));
    println!("Seed 916212 target: {:?}", target);
    match target {
        ResultData::Error(e) => assert_eq!(e, "#NUM!"),
        other => panic!("Expected #NUM!, got {:?}", other),
    }
}

#[test]
fn test_fuzz_sqrt_rounddown_negative_num_error() {
    let sheet_src = [
        ["-113.77", "85", "", "-60", "85"],
        ["-282.7737", "\"S1vRUvgQ\"", "\"FA\"", "-57", "407"],
        ["-8", "65", "390.8202", "-11.7", "0"],
        ["FALSE", "18.0384", "41", "\"ZpGcDI2d\"", "-206"],
        ["\"z\"", "", "-86", "", "-241"],
        ["=46", "=-39", "=A4", "=UPPER(\"SUM(C1, 41)\")", "=33"],
        [
            "=SQRT(LEN(\"E1\"))",
            "=SQRT(A6)",
            "=OR(D4 > 0, -4 < 100)",
            "=AND(A6 > 0, A2 < 100)",
            "",
        ],
        [
            "=IF((MAX(C3:C3) > LEN(\"20\")), IF((32 > E1), E2, E3), 1)",
            "=B1",
            "=E5",
            "=ABS(-18)",
            "=D7",
        ],
        [
            "=IF((LEN(\"-24\") > PRODUCT(22, B7)), SUM(A2, -30), MIN(C6:E7))",
            "=C1",
            "=AVERAGE(LEN(\"A6\"), LOWER(\"C1\"))",
            "=D1",
            "=E3",
        ],
        ["=MAX(D8:E8)", "=B7", "=SQRT(D9)", "=C6", "=(D7 * C6)"],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 2));
    println!("Seed 812174 target: {:?}", target);
    match target {
        ResultData::Error(e) => assert_eq!(e, "#NUM!"),
        other => panic!("Expected #NUM!, got {:?}", other),
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
fn test_fuzz_boolean_subtraction_roundup_val() {
    let sheet_src = [
        ["TRUE", "47.7", "141.122", "74", "46"],
        ["92", "-48", "TRUE", "\"NqhvDZ\"", "69"],
        ["2", "FALSE", "2", "-95", ""],
        ["\"bmWR\"", "26", "481.1", "FALSE", "FALSE"],
        ["-86", "-66", "\"dR\"", "-82.502", "-80"],
        ["=A2", "=A4", "0", "=2", "=ROUNDDOWN(B1, 0)"],
        ["=LOWER(\"INT(C5)\")", "=E5", "=E5", "=SQRT(B2)", "52"],
        [
            "=D3",
            "=IF((LEFT(\"B2\", 1) > C4), D6, C3)",
            "=AVERAGE(D3:E7)",
            "=(ROUND(-50, 0) - 47)",
            "=(INT(48) * B1)",
        ],
        [
            "=41",
            "\"Us\"",
            "=ROUNDUP(D3, 0)",
            "=SQRT(MAX(D7:E7))",
            "=ABS(-3)",
        ],
        [
            "=D2",
            "=(AND(C6 > 0, E5 < 100) - ROUNDUP(E8, 1))",
            "",
            "=C1",
            "\"Gd\"",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 1));
    println!("Seed 355893 target: {:?}", target);
    match target {
        ResultData::Float(f) => assert!((f - -2289.6).abs() < 0.1),
        ResultData::Integer(i) => assert_eq!(i, -2289),
        other => panic!("Expected -2289.6, got {:?}", other),
    }
}

#[test]
fn test_fuzz_nested_logical_and_or_numeric_comparison() {
    let sheet_src = [
        ["15", "42", "-10", "50", "-97"],
        ["92", "-48", "TRUE", "\"NqhvDZ\"", "69"],
        ["2", "FALSE", "2", "-95", ""],
        ["\"bmWR\"", "26", "481.1", "FALSE", "FALSE"],
        ["-86", "-66", "10", "-82.502", "-80"],
        ["=A1", "=B1", "=C1", "=D1", "=E1"],
        ["=A2", "=B2", "=C2", "=D2", "=E2"],
        ["=A3", "=B3", "=C3", "=D3", "=E3"],
        ["=A4", "=B4", "=C4", "=D4", "=E4"],
        [
            "=A1",
            "=B1",
            "=C1",
            "=D1",
            "=AND(OR(15 > 0, D4 < 100) > 0, AND(C5 > 0, B4 < 100))",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 4));
    match target {
        ResultData::Boolean(b) => assert_eq!(b, true),
        other => panic!("Expected Boolean(true), got {:?}", other),
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


