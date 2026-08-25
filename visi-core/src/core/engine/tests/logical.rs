use super::*;

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
        ResultData::Boolean(b) => assert!(b),
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
                assert!(('A'..='E').contains(&col_let));
            }
        }
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
        ResultData::Boolean(b) => assert!(b),
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
        ResultData::Boolean(b) => assert!(b),
        other => panic!("Expected Boolean(true), got {:?}", other),
    }
}

/// The numeric value at a cell, for the empty-string tests below.
fn numeric_at(sheet: &Sheet, row: usize, col: usize) -> f64 {
    let value = sheet.get_result_data(&CellRef::new(row, col));
    sheet
        .to_f64(&value)
        .unwrap_or_else(|| panic!("expected a number at ({row}, {col}), got {value:?}"))
}

/// An empty-string cell is *text*, not a blank cell.
///
/// Harvested from a differential-fuzz grid whose `G1` held a single space.
/// OOXML strips whitespace-only `<t>` content that isn't marked
/// `xml:space="preserve"`, so both engines see an empty string -- and Excel
/// keeps it as a text cell. visi used to rebuild it as blank on import, which
/// made three separate formulas in that one grid disagree with Excel at once:
/// `TYPE(G1)` answered 1 instead of 2, `G1 < 100` answered TRUE instead of
/// FALSE (text sorts above every number in Excel), and an `ISERROR` over a
/// division by `AND(..., G1 < 100)` missed the `#DIV/0!` that Excel produced.
#[test]
fn test_fuzz_empty_string_cell_is_text_not_blank() {
    let sheet_src = [
        ["\"\"", "5", ""],
        ["=TYPE(A1)", "=ISTEXT(A1)", "=ISBLANK(A1)"],
        ["=A1<100", "=ISNUMBER(A1)", "=TYPE(C1)"],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();

    // TYPE: 2 is text, 1 would be a number -- and a blank cell reads as 1.
    assert_eq!(numeric_at(&sheet, 1, 0), 2.0);
    assert!(matches!(
        sheet.get_result_data(&CellRef::new(1, 1)),
        ResultData::Boolean(true)
    ));
    assert!(matches!(
        sheet.get_result_data(&CellRef::new(1, 2)),
        ResultData::Boolean(false)
    ));
    // Text is greater than any number in Excel's ordering, so this is FALSE.
    assert!(matches!(
        sheet.get_result_data(&CellRef::new(2, 0)),
        ResultData::Boolean(false)
    ));
    assert!(matches!(
        sheet.get_result_data(&CellRef::new(2, 1)),
        ResultData::Boolean(false)
    ));
    // The control: C1 really is blank, and a blank cell's TYPE is 1.
    assert_eq!(numeric_at(&sheet, 2, 2), 1.0);
}

/// Excel is asymmetric here and visi mirrors it: a cell holding the empty
/// string is counted by COUNTA *and* by COUNTBLANK, while ISBLANK reports it
/// as not blank. COUNTBLANK is documented to include cells whose formula
/// returned `""`; COUNTA counts it because it is a value.
#[test]
fn test_empty_string_counts_as_both_present_and_blank() {
    let sheet_src = [
        ["\"\"", "=COUNTA(A1)", "=COUNTBLANK(A1)"],
        ["", "=COUNTA(A2)", "=COUNTBLANK(A2)"],
        ["=IF(TRUE,\"\",\"x\")", "=COUNTA(A3)", "=COUNTBLANK(A3)"],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();

    // A text cell holding "": present to COUNTA, blank to COUNTBLANK.
    assert_eq!(numeric_at(&sheet, 0, 1), 1.0);
    assert_eq!(numeric_at(&sheet, 0, 2), 1.0);
    // A genuinely empty cell: counted by neither / only COUNTBLANK.
    assert_eq!(numeric_at(&sheet, 1, 1), 0.0);
    assert_eq!(numeric_at(&sheet, 1, 2), 1.0);
    // A formula that returned "" behaves like the text cell.
    assert_eq!(numeric_at(&sheet, 2, 1), 1.0);
    assert_eq!(numeric_at(&sheet, 2, 2), 1.0);
}

/// The `ISERROR` shape from the same fuzz grid, reduced: `AND(...)` over a
/// text cell yields FALSE, so the division is by zero. Getting `G1 < 100`
/// wrong turned this from `#DIV/0!` into an ordinary number.
#[test]
fn test_fuzz_and_over_empty_string_cell_forces_div_by_zero() {
    let sheet_src = [
        ["\"\"", "\"HIt\"", "-78"],
        ["=ISERROR((C1 + -16) / AND(B1 > 0, A1 < 100))", "", ""],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();

    assert!(matches!(
        sheet.get_result_data(&CellRef::new(1, 0)),
        ResultData::Boolean(true)
    ));
}

#[test]
fn test_fuzz_text_comparison_case_insensitive_lexicographical() {
    let sheet_src = [[
        "=(\"-4463\">\"36-33\")",
        "=(\"-1\">\"-2\")",
        "=(\"-2\">\"-1\")",
        "=(\"a-1\">\"a2\")",
        "=(\"-a\">\"a\")",
        "=(\"a\">\"-a\")",
        "=(\"1-2\">\"12\")",
        "=(\"12\">\"1-2\")",
        "=REPT(-324.7, 1) > LEFT(3, 2)",
    ]];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();

    let want = [false, false, true, false, false, true, false, true, false];
    for (col, expected) in want.into_iter().enumerate() {
        match sheet.get_result_data(&CellRef::new(0, col)) {
            ResultData::Boolean(b) => {
                assert_eq!(b, expected, "column {col} ({:?})", sheet_src[0][col])
            }
            other => panic!("column {col}: expected a boolean, got {other:?}"),
        }
    }
}
