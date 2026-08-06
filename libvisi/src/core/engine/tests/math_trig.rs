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
    assert!(matches!(r6, ResultData::Float(v) if (v - 3.14).abs() < 1e-6));
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
