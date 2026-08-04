use super::*;

#[test]
fn test_date_and_time_functions() {
    let grid = [
        ["=DATE(2024, 8, 3)", "=YEAR(DATE(2024, 8, 3))", "=MONTH(DATE(2024, 8, 3))", "=DAY(DATE(2024, 8, 3))", "=TIME(12, 30, 0)", "=HOUR(0.5)"],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None);

    let r1 = sheet.get_result_data(&CellRef::new(0, 0));
    assert!(matches!(r1, ResultData::Float(v) if (v - 45507.0).abs() < 10.0));

    let r2 = sheet.get_result_data(&CellRef::new(0, 1));
    assert!(matches!(r2, ResultData::Float(v) if (v - 2024.0).abs() < 1e-6));

    let r3 = sheet.get_result_data(&CellRef::new(0, 2));
    assert!(matches!(r3, ResultData::Float(v) if (v - 8.0).abs() < 1e-6));

    let r4 = sheet.get_result_data(&CellRef::new(0, 3));
    assert!(matches!(r4, ResultData::Float(v) if (v - 3.0).abs() < 1e-6));

    let r5 = sheet.get_result_data(&CellRef::new(0, 4));
    assert!(matches!(r5, ResultData::Float(v) if (v - 0.52083333).abs() < 1e-4));

    let r6 = sheet.get_result_data(&CellRef::new(0, 5));
    assert!(matches!(r6, ResultData::Float(v) if (v - 12.0).abs() < 1e-6));
}

#[test]
fn test_engineering_functions() {
    let grid = [
        ["=BIN2DEC(\"1010\")", "=DEC2HEX(255)", "=BITAND(6, 3)", "=DELTA(5, 5)", "=GESTEP(10, 5)", "=CONVERT(1, \"km\", \"m\")"],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None);

    let r1 = sheet.get_result_data(&CellRef::new(0, 0));
    assert!(matches!(r1, ResultData::Float(v) if (v - 10.0).abs() < 1e-6));

    let r2 = sheet.get_result_data(&CellRef::new(0, 1));
    assert!(matches!(r2, ResultData::String(ref s) if s == "FF"));

    let r3 = sheet.get_result_data(&CellRef::new(0, 2));
    assert!(matches!(r3, ResultData::Float(v) if (v - 2.0).abs() < 1e-6));

    let r4 = sheet.get_result_data(&CellRef::new(0, 3));
    assert!(matches!(r4, ResultData::Float(v) if (v - 1.0).abs() < 1e-6));

    let r5 = sheet.get_result_data(&CellRef::new(0, 4));
    assert!(matches!(r5, ResultData::Float(v) if (v - 1.0).abs() < 1e-6));

    let r6 = sheet.get_result_data(&CellRef::new(0, 5));
    assert!(matches!(r6, ResultData::Float(v) if (v - 1000.0).abs() < 1e-6));
}

#[test]
fn test_information_logical_lookup_web_functions() {
    let grid = [
        ["=ISEVEN(4)", "=ISODD(5)", "=TYPE(100)", "=XOR(TRUE, FALSE)", "=ADDRESS(1, 1)", "=ENCODEURL(\"hello world\")"],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None);

    let r1 = sheet.get_result_data(&CellRef::new(0, 0));
    assert!(matches!(r1, ResultData::Boolean(true)));

    let r2 = sheet.get_result_data(&CellRef::new(0, 1));
    assert!(matches!(r2, ResultData::Boolean(true)));

    let r3 = sheet.get_result_data(&CellRef::new(0, 2));
    assert!(matches!(r3, ResultData::Float(v) if (v - 1.0).abs() < 1e-6));

    let r4 = sheet.get_result_data(&CellRef::new(0, 3));
    assert!(matches!(r4, ResultData::Boolean(true)));

    let r5 = sheet.get_result_data(&CellRef::new(0, 4));
    assert!(matches!(r5, ResultData::String(ref s) if s == "$A$1"));

    let r6 = sheet.get_result_data(&CellRef::new(0, 5));
    assert!(matches!(r6, ResultData::String(ref s) if s == "hello%20world"));
}
