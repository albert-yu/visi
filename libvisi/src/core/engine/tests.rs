use super::*;
use std::collections::HashSet;

fn test_eval<T, F>(source: &str, expected: T, val_getter: F) -> Result<(), EngineError>
where
    T: std::cmp::PartialEq + std::fmt::Debug,
    F: Fn(&ResultData) -> Option<T>,
{
    let sheet = Sheet::new(SheetInit::default());
    let (result, _) = sheet.eval(source, None)?;
    let result_value = val_getter(&result).expect("Failed to get value");
    assert_eq!(expected, result_value);
    Ok(())
}

fn get_int_val(r: &ResultData) -> Option<i64> {
    match r {
        ResultData::Integer(i) => Some(*i),
        ResultData::Float(f) => Some(*f as i64),
        _ => None,
    }
}

fn get_float_val(r: &ResultData) -> Option<f64> {
    match r {
        ResultData::Float(f) => Some(*f),
        ResultData::Integer(i) => Some(*i as f64),
        _ => None,
    }
}

fn get_string_val(r: &ResultData) -> Option<String> {
    match r {
        ResultData::String(s) => Some(s.clone()),
        _ => None,
    }
}

fn test_ints(source: &str, expected: i64) -> Result<(), EngineError> {
    test_eval(source, expected, get_int_val)
}

fn test_floats(source: &str, expected: f64) -> Result<(), EngineError> {
    test_eval(source, expected, get_float_val)
}

fn test_strings(source: &str, expected: &str) -> Result<(), EngineError> {
    test_eval(source, expected.to_string(), get_string_val)
}

#[test]
fn test_integer_evaluations() {
    test_ints("-2", -2).unwrap();
    test_ints("=3 + 4", 7).unwrap();
    test_ints("=9 * 7", 63).unwrap();
    test_ints("=12*12", 144).unwrap();
    test_ints("=3**2", 9).unwrap();
    test_floats("=12 / 2", 6.0).unwrap();
    test_floats("=1 / 2", 0.5).unwrap();
}

#[test]
fn test_float_evaluations() {
    test_floats("-2.0", -2.0).unwrap();
    test_floats("=3.0 + 4.0", 7.0).unwrap();
    test_floats("=sum([1, 2, 3.5])", 6.5).unwrap();
}

#[test]
fn test_excel_formula_evaluations() {
    test_floats("=3 + 4", 7.0).unwrap();
    test_floats("=10 - 2 * 3", 4.0).unwrap();
    test_floats("=(10 - 2) * 3", 24.0).unwrap();
    test_floats("=2 ^ 3", 8.0).unwrap();
    test_floats("=10 / 4", 2.5).unwrap();

    test_floats("=SUM(1, 2, 3)", 6.0).unwrap();
    test_floats("=AVERAGE(1, 2, 3)", 2.0).unwrap();
    test_floats("=MIN(5, 3, 9)", 3.0).unwrap();
    test_floats("=MAX(5, 3, 9)", 9.0).unwrap();
    test_floats("=IF(3 > 2, 10, 20)", 10.0).unwrap();
    test_floats("=IF(1 > 2, 10, 20)", 20.0).unwrap();
    test_floats("=COUNT(1, 2, 3, \"foo\")", 3.0).unwrap();

    test_floats("=LN(EXP(1))", 1.0).unwrap();
    test_floats("=LOG10(100)", 2.0).unwrap();
    test_floats("=CEIL(4.2)", 5.0).unwrap();
    test_floats("=FLOOR(4.8)", 4.0).unwrap();
    test_floats("=TAN(0)", 0.0).unwrap();
    test_floats("=ASIN(0)", 0.0).unwrap();
    test_floats("=ACOS(1)", 0.0).unwrap();
    test_floats("=ATAN(0)", 0.0).unwrap();

    let mut sheet = Sheet::new(SheetInit {
        name: Some("table_1".to_string()),
        rows: 5,
        cols: 5,
        ..Default::default()
    });

    sheet.set_cell_src(0, 0, "10".to_string());
    sheet.set_cell_src(0, 1, "20".to_string());
    sheet.set_cell_src(1, 0, "30".to_string());
    sheet.set_cell_src(1, 1, "40".to_string());
    sheet.commit(None).unwrap();

    sheet.set_cell_src(2, 0, "=A1 + B1".to_string());
    sheet.commit(None).unwrap();
    assert_eq!(
        get_float_val(&sheet.get_result_data(&CellRef::new(2, 0))),
        Some(30.0)
    );

    sheet.set_cell_src(2, 1, "=SUM(A1:B2)".to_string());
    sheet.commit(None).unwrap();
    assert_eq!(
        get_float_val(&sheet.get_result_data(&CellRef::new(2, 1))),
        Some(100.0)
    );

    sheet.set_cell_src(2, 2, "=AVERAGE(A1:B2)".to_string());
    sheet.commit(None).unwrap();
    assert_eq!(
        get_float_val(&sheet.get_result_data(&CellRef::new(2, 2))),
        Some(25.0)
    );

    let mut table2 = Sheet::new(SheetInit {
        name: Some("table_2".to_string()),
        rows: 5,
        cols: 5,
        ..Default::default()
    });
    table2.set_cell_src(0, 0, "42".to_string());
    table2.commit(None).unwrap();

    let mut context = Context::default();
    context.sheets.insert("table_2".to_string(), &table2);

    let (res_cross, _) = sheet.eval("=table_2!A1", Some(&context)).unwrap();
    assert_eq!(get_float_val(&res_cross), Some(42.0));
}

fn get_bool_val(r: &ResultData) -> Option<bool> {
    match r {
        ResultData::Boolean(b) => Some(*b),
        _ => None,
    }
}

fn test_booleans(source: &str, expected: bool) -> Result<(), EngineError> {
    test_eval(source, expected, get_bool_val)
}

#[test]
fn test_new_excel_formulas() {
    test_booleans("=AND(1 > 0, 2 > 1)", true).unwrap();
    test_booleans("=AND(1 > 0, 2 < 1)", false).unwrap();
    test_booleans("=OR(1 > 0, 2 < 1)", true).unwrap();
    test_booleans("=OR(1 < 0, 2 < 1)", false).unwrap();
    test_booleans("=NOT(1 > 0)", false).unwrap();
    test_booleans("=NOT(1 < 0)", true).unwrap();

    test_strings("=CONCAT(\"Hello\", \" \", \"World\")", "Hello World").unwrap();
    test_strings("=LEFT(\"hello\", 3)", "hel").unwrap();
    test_strings("=LEFT(\"hello\")", "h").unwrap();
    test_strings("=RIGHT(\"hello\", 3)", "llo").unwrap();
    test_strings("=RIGHT(\"hello\")", "o").unwrap();
    test_strings("=MID(\"hello\", 2, 3)", "ell").unwrap();
    test_floats("=LEN(\"hello\")", 5.0).unwrap();
    test_strings("=TRIM(\"  hello   world  \")", "hello world").unwrap();
    test_strings("=UPPER(\"hello\")", "HELLO").unwrap();
    test_strings("=LOWER(\"HELLO\")", "hello").unwrap();
    test_strings("=PROPER(\"hello world\")", "Hello World").unwrap();

    test_booleans("=ISNUMBER(5)", true).unwrap();
    test_booleans("=ISNUMBER(\"foo\")", false).unwrap();
    test_booleans("=ISTEXT(\"foo\")", true).unwrap();
    test_booleans("=ISTEXT(5)", false).unwrap();
    test_booleans("=ISBLANK(GET(10, 10))", true).unwrap();
    test_booleans("=ISERROR(1 / 0)", true).unwrap();
    test_booleans("=ISERROR(5)", false).unwrap();

    test_floats("=PRODUCT(2, 3, 4)", 24.0).unwrap();
    test_floats("=MOD(10, 3)", 1.0).unwrap();
    test_floats("=MOD(-10, 3)", 2.0).unwrap();
    test_floats("=COUNTA(1, \"foo\", GET(10, 10))", 2.0).unwrap();

    test_floats("=IFERROR(5, 10)", 5.0).unwrap();
    test_floats("=IFERROR(1 / 0, 10)", 10.0).unwrap();

    let (today_res, _) = Sheet::new(SheetInit::default())
        .eval("=TODAY()", None)
        .unwrap();
    if let ResultData::String(s) = today_res {
        assert_eq!(s.len(), 10);
        assert_eq!(&s[4..5], "-");
        assert_eq!(&s[7..8], "-");
    } else {
        panic!("Expected TODAY() to be string");
    }

    let (now_res, _) = Sheet::new(SheetInit::default())
        .eval("=NOW()", None)
        .unwrap();
    if let ResultData::String(s) = now_res {
        assert_eq!(s.len(), 19);
        assert_eq!(&s[10..11], " ");
        assert_eq!(&s[13..14], ":");
    } else {
        panic!("Expected NOW() to be string");
    }

    let mut sheet = Sheet::new(SheetInit {
        name: Some("data_table".to_string()),
        rows: 5,
        cols: 5,
        ..Default::default()
    });

    sheet.set_cell_src(0, 0, "1".to_string());
    sheet.set_cell_src(0, 1, "10".to_string());
    sheet.set_cell_src(0, 2, "Apples".to_string());
    sheet.set_cell_src(1, 0, "2".to_string());
    sheet.set_cell_src(1, 1, "20".to_string());
    sheet.set_cell_src(1, 2, "Oranges".to_string());
    sheet.set_cell_src(2, 0, "3".to_string());
    sheet.set_cell_src(2, 1, "30".to_string());
    sheet.set_cell_src(2, 2, "Apples".to_string());
    sheet.commit(None).unwrap();

    sheet.set_cell_src(3, 0, "=MATCH(2, A1:A3, 0)".to_string());

    sheet.set_cell_src(3, 1, "=INDEX(A1:C3, 2, 2)".to_string());

    sheet.set_cell_src(3, 2, "=VLOOKUP(2, A1:C3, 3, FALSE)".to_string());

    sheet.set_cell_src(3, 3, "=SUMIF(C1:C3, \"Apples\", B1:B3)".to_string());

    sheet.set_cell_src(
        4,
        0,
        "=SUMIFS(B1:B3, C1:C3, \"Apples\", A1:A3, \">1\")".to_string(),
    );

    sheet.set_cell_src(4, 1, "=COUNTIF(C1:C3, \"Apples\")".to_string());

    sheet.set_cell_src(
        4,
        2,
        "=COUNTIFS(C1:C3, \"Apples\", A1:A3, \">1\")".to_string(),
    );

    sheet.set_cell_src(4, 3, "=XLOOKUP(30, B1:B3, C1:C3)".to_string());
    sheet.commit(None).unwrap();

    assert_eq!(
        get_float_val(&sheet.get_result_data(&CellRef::new(3, 0))),
        Some(2.0)
    );
    assert_eq!(
        get_float_val(&sheet.get_result_data(&CellRef::new(3, 1))),
        Some(20.0)
    );
    assert_eq!(
        get_string_val(&sheet.get_result_data(&CellRef::new(3, 2))),
        Some("Oranges".to_string())
    );
    assert_eq!(
        get_float_val(&sheet.get_result_data(&CellRef::new(3, 3))),
        Some(40.0)
    );
    assert_eq!(
        get_float_val(&sheet.get_result_data(&CellRef::new(4, 0))),
        Some(30.0)
    );
    assert_eq!(
        get_float_val(&sheet.get_result_data(&CellRef::new(4, 1))),
        Some(2.0)
    );
    assert_eq!(
        get_float_val(&sheet.get_result_data(&CellRef::new(4, 2))),
        Some(1.0)
    );
    assert_eq!(
        get_string_val(&sheet.get_result_data(&CellRef::new(4, 3))),
        Some("Apples".to_string())
    );

    let (mmult_res, _) = sheet.eval("=MMULT(A1:B2, A1:B2)", None).unwrap();
    if let ResultData::List(list) = mmult_res {
        assert_eq!(list.len(), 4);
        assert_eq!(get_float_val(&list[0]), Some(21.0));
        assert_eq!(get_float_val(&list[1]), Some(210.0));
        assert_eq!(get_float_val(&list[2]), Some(42.0));
        assert_eq!(get_float_val(&list[3]), Some(420.0));
    } else {
        panic!("Expected MMULT to return a list");
    }
}

#[test]
fn test_plot_is_line() {
    let mut sheet = Sheet::new(SheetInit::default());
    sheet.set_cell_src(
        0,
        0,
        "=plot([0, 1, 2], [0, 1, 4], [1.0, 0.0, 0.0, 1.0], 0.01, \"line\")".to_string(),
    );
    sheet.commit(None).unwrap();
    let res = sheet.get_result_data(&CellRef::new(0, 0));
    if let ResultData::Plot {
        is_line, radius, ..
    } = res
    {
        assert!(is_line);
        assert_eq!(radius, 0.01);
    } else {
        panic!("Expected Plot, got {:?}", res);
    }

    let mut table2 = Sheet::new(SheetInit::default());
    table2.set_cell_src(
        0,
        0,
        "=plot([0, 1, 2], [0, 1, 4], [1.0, 0.0, 0.0, 1.0], \"line\")".to_string(),
    );
    table2.commit(None).unwrap();
    let res2 = table2.get_result_data(&CellRef::new(0, 0));
    if let ResultData::Plot {
        is_line, radius, ..
    } = res2
    {
        assert!(is_line);
        assert_eq!(radius, 0.005);
    } else {
        panic!("Expected Plot, got {:?}", res2);
    }

    let mut table3 = Sheet::new(SheetInit::default());
    table3.set_cell_src(0, 0, "=plot([0, 1, 2], [0, 1, 4], \"line\")".to_string());
    table3.commit(None).unwrap();
    let res3 = table3.get_result_data(&CellRef::new(0, 0));
    if let ResultData::Plot { is_line, .. } = res3 {
        assert!(is_line);
    } else {
        panic!("Expected Plot, got {:?}", res3);
    }

    let mut table4 = Sheet::new(SheetInit::default());
    table4.set_cell_src(0, 0, "=plot([0, 1, 2], [0, 1, 4])".to_string());
    table4.commit(None).unwrap();
    let res4 = table4.get_result_data(&CellRef::new(0, 0));
    if let ResultData::Plot { is_line, .. } = res4 {
        assert!(!is_line);
    } else {
        panic!("Expected Plot, got {:?}", res4);
    }

    let mut table5 = Sheet::new(SheetInit::default());
    table5.set_cell_src(
        0,
        0,
        "=plot([0, 1, 2], [0, 1, 4], type=\"line\")".to_string(),
    );
    table5.commit(None).unwrap();
    let res5 = table5.get_result_data(&CellRef::new(0, 0));
    if let ResultData::Plot { is_line, .. } = res5 {
        assert!(is_line);
    } else {
        panic!("Expected Plot, got {:?}", res5);
    }

    let mut table6 = Sheet::new(SheetInit::default());
    table6.set_cell_src(
        0,
        0,
        "=plot([0, 1, 2], [0, 1, 4], radius=0.02, type=\"line\")".to_string(),
    );
    table6.commit(None).unwrap();
    let res6 = table6.get_result_data(&CellRef::new(0, 0));
    if let ResultData::Plot {
        is_line, radius, ..
    } = res6
    {
        assert!(is_line);
        assert_eq!(radius, 0.02);
    } else {
        panic!("Expected Plot, got {:?}", res6);
    }

    let mut table7 = Sheet::new(SheetInit::default());
    table7.set_cell_src(
        0,
        0,
        "=plot([0, 1, 2], [0, 1, 4], color=[0.5, 0.5, 0.5, 1.0], radius=0.03, type=\"scatter\")"
            .to_string(),
    );
    table7.commit(None).unwrap();
    let res7 = table7.get_result_data(&CellRef::new(0, 0));
    if let ResultData::Plot {
        is_line,
        radius,
        color,
        ..
    } = res7
    {
        assert!(!is_line);
        assert_eq!(radius, 0.03);
        assert_eq!(color, [0.5, 0.5, 0.5, 1.0]);
    } else {
        panic!("Expected Plot, got {:?}", res7);
    }

    let mut table8 = Sheet::new(SheetInit::default());
    table8.set_cell_src(
        0,
        0,
        "=plot([0, 1], [0, 1], title=\"My Plot\", xlabel=\"X Axis\", ylabel=\"Y Axis\")"
            .to_string(),
    );
    table8.commit(None).unwrap();
    let res8 = table8.get_result_data(&CellRef::new(0, 0));
    if let ResultData::Plot {
        title,
        xlabel,
        ylabel,
        ..
    } = res8
    {
        assert_eq!(title, Some("My Plot".to_string()));
        assert_eq!(xlabel, Some("X Axis".to_string()));
        assert_eq!(ylabel, Some("Y Axis".to_string()));
    } else {
        panic!("Expected Plot, got {:?}", res8);
    }
}

#[test]
fn test_plot_cell_dimensions() {
    let mut sheet = Sheet::new(SheetInit::default());
    sheet.set_cell_src(0, 0, "=plot([0, 2, 4], [0, 0, 2])".to_string());
    sheet.commit(None).unwrap();
    let res = sheet.get_result_data(&CellRef::new(0, 0));
    assert_eq!(res.plot_cell_dims(), Some((16, 9)));

    let mut table2 = Sheet::new(SheetInit::default());
    table2.set_cell_src(0, 0, "=plot([0, 0, 2], [0, 2, 4])".to_string());
    table2.commit(None).unwrap();
    let res2 = table2.get_result_data(&CellRef::new(0, 0));
    assert_eq!(res2.plot_cell_dims(), Some((9, 16)));

    let empty_plot = ResultData::Plot {
        points: vec![],
        color: [0.0, 0.0, 0.0, 1.0],
        radius: 1.0,
        is_line: false,
        title: None,
        xlabel: None,
        ylabel: None,
    };
    assert_eq!(empty_plot.plot_cell_dims(), Some((16, 16)));

    let single_point_plot = ResultData::Plot {
        points: vec![(1.0, 1.0)],
        color: [0.0, 0.0, 0.0, 1.0],
        radius: 1.0,
        is_line: false,
        title: None,
        xlabel: None,
        ylabel: None,
    };
    assert_eq!(single_point_plot.plot_cell_dims(), Some((16, 16)));
}

#[test]
fn test_precedence() {
    test_ints("=3 + 4 * 5", 23).unwrap();
    test_ints("=3 * 4 + 5", 17).unwrap();
    test_ints("=3 * (4 + 5)", 27).unwrap();
    test_ints("=3 * (4 + 5) * 2", 54).unwrap();
    test_ints("=3 * (4 + 5) * 2 + 1", 55).unwrap();
    test_ints("=3 * (4 + 5) * (2 + 1)", 81).unwrap();
}

#[test]
fn test_load() {
    let mut sheet = Sheet::new(SheetInit::default());

    sheet.insert(
        TextCellRef {
            row: 0,
            col: 0,
            char_offset: 0,
        },
        "1",
    );
    sheet.commit(None).unwrap();

    let (two, _) = sheet.eval("=A1 + 1", None).unwrap();
    match two {
        ResultData::Integer(val) => assert_eq!(val, 2),
        ResultData::Float(val) => assert_eq!(val, 2.0),
        _ => panic!("Expected number result"),
    }

    let (cell_alone, _) = sheet.eval("=A1", None).unwrap();
    match cell_alone {
        ResultData::Integer(val) => assert_eq!(val, 1),
        ResultData::Float(val) => assert_eq!(val, 1.0),
        _ => panic!("Expected number result"),
    }
}

#[test]
fn test_concatenation() {
    test_strings("=\"Hello World\"", "Hello World").unwrap();

    test_strings("=CONCATENATE(\"A\", \"B\")", "AB").unwrap();

    test_strings("=CONCATENATE(\"Hello\", \" World\")", "Hello World").unwrap();
    test_strings("=CONCATENATE(\"ABC\", \"DEF\")", "ABCDEF").unwrap();

    test_strings("=CONCATENATE(str(5), \" items\")", "5 items").unwrap();
    test_strings("=CONCATENATE(\"Value: \", str(42))", "Value: 42").unwrap();
    test_strings("=CONCATENATE(str(3.14), \" is pi\")", "3.14 is pi").unwrap();
    test_strings("=CONCATENATE(\"Result: \", str(True))", "Result: True").unwrap();
}

#[test]
fn test_table_naming() {
    let table1 = Sheet::new(SheetInit::default());
    assert_eq!(table1.name, "table_1");

    let table2 = Sheet::new(SheetInit {
        name: Some("my_table".to_string()),
        ..Default::default()
    });
    assert_eq!(table2.name, "my_table");

    let table3 = Sheet::new(SheetInit {
        name: None,
        ..Default::default()
    });
    assert_eq!(table3.name, "table_1");
}

#[test]
fn test_table_references() {
    let table1 = Sheet::new(SheetInit {
        id: None,
        name: Some("table_1".to_string()),
        rows: 5,
        cols: 5,
    });

    let mut table2 = Sheet::new(SheetInit {
        id: None,
        name: Some("table_2".to_string()),
        rows: 5,
        cols: 5,
    });

    table2.insert(
        TextCellRef {
            row: 0,
            col: 0,
            char_offset: 0,
        },
        "42",
    );
    table2.commit(None).unwrap();

    table2.insert(
        TextCellRef {
            row: 1,
            col: 1,
            char_offset: 0,
        },
        "100",
    );
    table2.commit(None).unwrap();

    let mut context = Context::default();
    context.sheets.insert("table_2".to_string(), &table2);

    let (result, _) = table1.eval("=table_2!A1", Some(&context)).unwrap();
    assert_eq!(get_int_val(&result), Some(42));

    let (result2, _) = table1.eval("=table_2!B2", Some(&context)).unwrap();
    assert_eq!(get_int_val(&result2), Some(100));

    let (result3, _) = table1.eval("=table_2!A1 + 8", Some(&context)).unwrap();
    assert_eq!(get_int_val(&result3), Some(50));
}

#[test]
fn test_context_from_tables() {
    let table1 = Sheet::new(SheetInit {
        id: None,
        name: Some("table_1".to_string()),
        rows: 5,
        cols: 5,
    });

    let mut table2 = Sheet::new(SheetInit {
        id: None,
        name: Some("table_2".to_string()),
        rows: 5,
        cols: 5,
    });

    table2.insert(
        TextCellRef {
            row: 0,
            col: 0,
            char_offset: 0,
        },
        "99",
    );
    table2.commit(None).unwrap();

    let sheets = vec![table1, table2];
    let mut context = Context::new();
    for sheet in &sheets {
        context.add_table(sheet.name.clone(), &sheet);
    }

    assert!(context.sheets.contains_key("table_1"));
    assert!(context.sheets.contains_key("table_2"));

    let (result, _) = sheets[0].eval("=table_2!A1", Some(&context)).unwrap();
    assert_eq!(get_int_val(&result), Some(99));
}

#[test]
fn test_builtin_math_functions() {
    let sheet = Sheet::new(SheetInit::default());

    let (result, _) = sheet.eval("=abs(-5)", None).unwrap();
    assert_eq!(get_int_val(&result), Some(5));

    let (result, _) = sheet.eval("=abs(-3.14)", None).unwrap();
    assert_eq!(get_float_val(&result), Some(3.14));

    let (result, _) = sheet.eval("=sqrt(16)", None).unwrap();
    assert_eq!(get_float_val(&result), Some(4.0));

    let (result, _) = sheet.eval("=sin(0)", None).unwrap();
    assert_eq!(get_float_val(&result), Some(0.0));

    let (result, _) = sheet.eval("=cos(0)", None).unwrap();
    assert_eq!(get_float_val(&result), Some(1.0));

    let (result, _) = sheet.eval("=floor(3.9)", None).unwrap();
    assert_eq!(get_int_val(&result), Some(3));

    let (result, _) = sheet.eval("=ceil(3.1)", None).unwrap();
    assert_eq!(get_int_val(&result), Some(4));

    let (result, _) = sheet.eval("=round(3.7)", None).unwrap();
    assert_eq!(get_int_val(&result), Some(4));

    let (result, _) = sheet.eval("=round(3.2)", None).unwrap();
    assert_eq!(get_int_val(&result), Some(3));

    let (result, _) = sheet.eval("=exp(0)", None).unwrap();
    assert_eq!(get_float_val(&result), Some(1.0));

    let (result, _) = sheet.eval("=ln(1)", None).unwrap();
    assert_eq!(get_float_val(&result), Some(0.0));

    let (result, _) = sheet.eval("=log10(10)", None).unwrap();
    assert_eq!(get_float_val(&result), Some(1.0));

    let (result, _) = sheet.eval("=rand()", None).unwrap();
    if let ResultData::Float(val) = result {
        assert!(val >= 0.0 && val < 1.0);
    } else {
        panic!("Expected RAND to return a Float, got {:?}", result);
    }

    let (result, _) = sheet.eval("=randbetween(5, 10)", None).unwrap();
    if let ResultData::Integer(val) = result {
        assert!(val >= 5 && val <= 10);
    } else {
        panic!(
            "Expected RANDBETWEEN to return an Integer, got {:?}",
            result
        );
    }
}
#[test]
fn test_dependency_propagation() {
    let mut sheet = Sheet::new(SheetInit::default());

    sheet.insert(
        TextCellRef {
            row: 0,
            col: 0,
            char_offset: 0,
        },
        "10",
    );
    sheet.commit(None).unwrap();

    sheet.insert(
        TextCellRef {
            row: 0,
            col: 1,
            char_offset: 0,
        },
        "=A1 + 5",
    );
    sheet.commit(None).unwrap();

    let b1 = sheet.eval("=B1", None).unwrap().0;
    assert_eq!(get_int_val(&b1), Some(15));

    sheet.columns[0].src[0] = "20".to_string();
    sheet.columns[0].dirty_indices.push(0);

    sheet.commit(None).unwrap();

    let b1_new = sheet.eval("=B1", None).unwrap().0;
    assert_eq!(get_int_val(&b1_new), Some(25));

    sheet.insert(
        TextCellRef {
            row: 0,
            col: 2,
            char_offset: 0,
        },
        "=B1 * 2",
    );
    sheet.commit(None).unwrap();

    let c1 = sheet.eval("=C1", None).unwrap().0;
    assert_eq!(get_int_val(&c1), Some(50));

    sheet.insert(
        TextCellRef {
            row: 0,
            col: 0,
            char_offset: 0,
        },
        "1",
    );

    sheet.columns[0].src[0] = String::new();
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 0,
            char_offset: 0,
        },
        "1",
    );
    sheet.commit(None).unwrap();

    let b1_final = sheet.eval("=B1", None).unwrap().0;
    let c1_final = sheet.eval("=C1", None).unwrap().0;

    assert_eq!(get_int_val(&b1_final), Some(6));
    assert_eq!(get_int_val(&c1_final), Some(12));
}
#[test]
fn test_cross_table_dependency_propagation() {
    let mut table1 = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        ..Default::default()
    });

    let mut table2 = Sheet::new(SheetInit {
        name: Some("Sheet2".to_string()),
        ..Default::default()
    });

    table1.insert(
        TextCellRef {
            row: 0,
            col: 0,
            char_offset: 0,
        },
        "10",
    );
    table1.commit(None).unwrap();

    table2.insert(
        TextCellRef {
            row: 0,
            col: 0,
            char_offset: 0,
        },
        "=Sheet1!A1 * 2",
    );

    let mut context = Context::new();
    context.add_table("Sheet1".to_string(), &table1);
    table2.commit(Some(&context)).unwrap();

    let b1 = table2.eval("=A1", None).unwrap().0;
    assert_eq!(get_int_val(&b1), Some(20));

    table1.columns[0].src[0] = "20".to_string();
    table1.columns[0].dirty_indices.push(0);

    let updated_cells_1 = table1.commit(None).unwrap();
    assert!(updated_cells_1.contains(&CellRef::new(0, 0)));

    for cell in updated_cells_1 {
        let dep = Dependency::Remote {
            sheet: "Sheet1".to_string(),
            cell,
        };
        table2.invalidate_dependency(&dep);
    }

    let mut context = Context::new();
    context.add_table("Sheet1".to_string(), &table1);
    table2.commit(Some(&context)).unwrap();

    let b1_new = table2.eval("=A1", None).unwrap().0;
    assert_eq!(get_int_val(&b1_new), Some(40));
}

#[test]
fn test_table_serialization() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("test_table".to_string()),
        ..Default::default()
    });

    let cell1 = CellRef::new(0, 0);
    let cell2 = CellRef::new(1, 1);
    let dep = Dependency::Local(cell1);

    let mut dependents = HashSet::new();
    dependents.insert(cell2);
    sheet.dependencies.insert(dep.clone(), dependents);

    let mut providers = HashSet::new();
    providers.insert(dep);
    sheet.dependencies_rev.insert(cell2, providers);

    let json = serde_json::to_string(&sheet).expect("Failed to serialize sheet");

    let deserialized: Sheet = serde_json::from_str(&json).expect("Failed to deserialize sheet");

    assert_eq!(deserialized.name, sheet.name);

    assert_eq!(deserialized.dependencies.len(), 0);
    assert_eq!(deserialized.dependencies_rev.len(), 0);
}

#[test]
fn test_error_handling() {
    let mut sheet = Sheet::new(SheetInit::default());

    sheet.insert(
        TextCellRef {
            row: 0,
            col: 0,
            char_offset: 0,
        },
        "=1 / 0",
    );
    sheet.commit(None).unwrap();

    let res = sheet.get_result_data(&CellRef::new(0, 0));
    match res {
        ResultData::Error(e) => {
            assert!(
                e.contains("#DIV/0!")
                    || e.contains("ZeroDivisionError")
                    || e.contains("division by zero")
            )
        }
        _ => panic!("Expected error, got {:?}", res),
    }

    sheet.insert(
        TextCellRef {
            row: 1,
            col: 0,
            char_offset: 0,
        },
        "=A1 + 1",
    );
    sheet.commit(None).unwrap();

    let res2 = sheet.get_result_data(&CellRef::new(1, 0));
    match res2 {
        ResultData::Error(_) => {}
        _ => panic!("Expected error to propagate, got {:?}", res2),
    }
}

#[test]
fn test_fuzz_reproducer_seed_843244() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 0,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 1,
            char_offset: 0,
        },
        "-278.17",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 2,
            char_offset: 0,
        },
        "-55",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 3,
            char_offset: 0,
        },
        "29",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 4,
            char_offset: 0,
        },
        "240",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 0,
            char_offset: 0,
        },
        "-35",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 1,
            char_offset: 0,
        },
        "\"N\"",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 2,
            char_offset: 0,
        },
        "7",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 3,
            char_offset: 0,
        },
        "-419",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 4,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 0,
            char_offset: 0,
        },
        "16",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 1,
            char_offset: 0,
        },
        "10",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 2,
            char_offset: 0,
        },
        "0",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 3,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 4,
            char_offset: 0,
        },
        "92",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 0,
            char_offset: 0,
        },
        "39",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 1,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 2,
            char_offset: 0,
        },
        "66",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 3,
            char_offset: 0,
        },
        "\"zpx3A\"",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 4,
            char_offset: 0,
        },
        "356",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 0,
            char_offset: 0,
        },
        "-31",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 1,
            char_offset: 0,
        },
        "\"ZLwUBtDn\"",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 2,
            char_offset: 0,
        },
        "-133.063",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 3,
            char_offset: 0,
        },
        "3",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 4,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 0,
            char_offset: 0,
        },
        "=IF(((28 * -25) > LOWER(\"-8\")), ROUNDUP(D4, 2), E2)",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 1,
            char_offset: 0,
        },
        "=(LOWER(\"D4\") - D3)",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 2,
            char_offset: 0,
        },
        "=ROUNDUP(ROUNDUP(D3, 0), 1)",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 3,
            char_offset: 0,
        },
        "-43",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 4,
            char_offset: 0,
        },
        "=UPPER(\"ABS(-46)\")",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 0,
            char_offset: 0,
        },
        "=B4",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 1,
            char_offset: 0,
        },
        "=SUM(OR(-12 > 0, B1 < 100), ABS(D3))",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 2,
            char_offset: 0,
        },
        "318",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 3,
            char_offset: 0,
        },
        "=C4",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 4,
            char_offset: 0,
        },
        "=-1",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 0,
            char_offset: 0,
        },
        "=(IF((D7 > D1), -29, B7) / UPPER(\"E4\"))",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 1,
            char_offset: 0,
        },
        "=46",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 2,
            char_offset: 0,
        },
        "=IF((B7 > (B5 * D1)), CONCATENATE(\"45\", \"-46\"), OR(E7 > 0, C2 < 100))",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 3,
            char_offset: 0,
        },
        "=(45 ^ -27)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 4,
            char_offset: 0,
        },
        "=ROUND((E5 * D6), 1)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 0,
            char_offset: 0,
        },
        "=(IF((D8 > -19), -8, A4) ^ AVERAGE(20, 5))",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 1,
            char_offset: 0,
        },
        "=PRODUCT((B4 / B5), ROUND(A2, 0))",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 2,
            char_offset: 0,
        },
        "=D5",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 3,
            char_offset: 0,
        },
        "=SQRT(E7)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 4,
            char_offset: 0,
        },
        "=AND(IF((A3 > 44), 48, E4) > 0, OR(B7 > 0, -49 < 100) < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 0,
            char_offset: 0,
        },
        "=A3",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 1,
            char_offset: 0,
        },
        "-497.457",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 2,
            char_offset: 0,
        },
        "=IF((B5 > (C1 * A6)), E7, A6)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 3,
            char_offset: 0,
        },
        "=LEN(\"INT(-5)\")",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 4,
            char_offset: 0,
        },
        "=-27",
    );
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 3));
    println!("test_fuzz_reproducer_seed_843244 evaluated: {:?}", target);
    assert!(
        matches!(target, ResultData::Error(ref e) if e.contains("#NUM!")),
        "Expected #NUM!, got {:?}",
        target
    );
}

#[test]
fn test_fuzz_reproducer_seed_655058() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 0,
            char_offset: 0,
        },
        "\"QA3T3dlz\"",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 1,
            char_offset: 0,
        },
        "-304",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 2,
            char_offset: 0,
        },
        "38",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 3,
            char_offset: 0,
        },
        "-230",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 4,
            char_offset: 0,
        },
        "-298.006",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 0,
            char_offset: 0,
        },
        "-159.7126",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 2,
            char_offset: 0,
        },
        "-224",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 4,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 0,
            char_offset: 0,
        },
        "43",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 1,
            char_offset: 0,
        },
        "1",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 3,
            char_offset: 0,
        },
        "-65",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 4,
            char_offset: 0,
        },
        "0",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 0,
            char_offset: 0,
        },
        "-445.02",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 1,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 2,
            char_offset: 0,
        },
        "-58",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 3,
            char_offset: 0,
        },
        "79",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 4,
            char_offset: 0,
        },
        "-44",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 1,
            char_offset: 0,
        },
        "63.5",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 2,
            char_offset: 0,
        },
        "133.3",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 3,
            char_offset: 0,
        },
        "99.2978",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 4,
            char_offset: 0,
        },
        "275",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 1,
            char_offset: 0,
        },
        "=E2",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 2,
            char_offset: 0,
        },
        "=ABS(SUM(E1:E4))",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 3,
            char_offset: 0,
        },
        "=(A5 - -15)",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 4,
            char_offset: 0,
        },
        "=OR(PRODUCT(27, B4) > 0, (B4 * A3) < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 0,
            char_offset: 0,
        },
        "=D2",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 1,
            char_offset: 0,
        },
        "=(INT(B1) * PRODUCT(B6, 46))",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 2,
            char_offset: 0,
        },
        "=B4",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 4,
            char_offset: 0,
        },
        "=UPPER(\"3\")",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 0,
            char_offset: 0,
        },
        "=AND(AVERAGE(13, -35) > 0, E3 < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 1,
            char_offset: 0,
        },
        "=AVERAGE((E7 - D5), C5)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 2,
            char_offset: 0,
        },
        "=B1",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 3,
            char_offset: 0,
        },
        "=E7",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 4,
            char_offset: 0,
        },
        "-26",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 0,
            char_offset: 0,
        },
        "=A2",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 1,
            char_offset: 0,
        },
        "=SUM(C6, LOWER(\"E2\"))",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 2,
            char_offset: 0,
        },
        "=SUM(C3:D4)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 4,
            char_offset: 0,
        },
        "=(C4 ^ 12)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 0,
            char_offset: 0,
        },
        "=-9",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 1,
            char_offset: 0,
        },
        "=B9",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 2,
            char_offset: 0,
        },
        "=-28",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 3,
            char_offset: 0,
        },
        "=E5",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 4,
            char_offset: 0,
        },
        "-399.593",
    );
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 1));
    println!("test_fuzz_reproducer_seed_655058 evaluated: {:?}", target);
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
fn test_fuzz_reproducer_seed_25814() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 0,
            char_offset: 0,
        },
        "278",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 1,
            char_offset: 0,
        },
        "94",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 3,
            char_offset: 0,
        },
        "\"Wb\"",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 4,
            char_offset: 0,
        },
        "-98",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 0,
            char_offset: 0,
        },
        "-19",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 1,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 3,
            char_offset: 0,
        },
        "50",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 4,
            char_offset: 0,
        },
        "0",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 0,
            char_offset: 0,
        },
        "3",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 1,
            char_offset: 0,
        },
        "400.65",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 2,
            char_offset: 0,
        },
        "270.96",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 3,
            char_offset: 0,
        },
        "5",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 4,
            char_offset: 0,
        },
        "-24",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 0,
            char_offset: 0,
        },
        "469.2",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 1,
            char_offset: 0,
        },
        "-81",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 2,
            char_offset: 0,
        },
        "-99",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 3,
            char_offset: 0,
        },
        "-68",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 4,
            char_offset: 0,
        },
        "-23",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 0,
            char_offset: 0,
        },
        "372.65",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 1,
            char_offset: 0,
        },
        "\"Gao \"",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 2,
            char_offset: 0,
        },
        "62",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 4,
            char_offset: 0,
        },
        "-54",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 0,
            char_offset: 0,
        },
        "=OR(D4 > 0, OR(B3 > 0, A2 < 100) < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 1,
            char_offset: 0,
        },
        "88",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 2,
            char_offset: 0,
        },
        "=LEFT(\"E1\", 4)",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 3,
            char_offset: 0,
        },
        "=(C3 - C5)",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 4,
            char_offset: 0,
        },
        "=INT(D2)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 0,
            char_offset: 0,
        },
        "=MAX(B2:D2)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 1,
            char_offset: 0,
        },
        "=B6",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 2,
            char_offset: 0,
        },
        "=(SQRT(E6) + INT(24))",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 3,
            char_offset: 0,
        },
        "=B3",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 4,
            char_offset: 0,
        },
        "=ROUNDDOWN(B3, 1)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 0,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 1,
            char_offset: 0,
        },
        "=INT(B4)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 2,
            char_offset: 0,
        },
        "=MAX(ROUNDDOWN(A2, 0), ROUNDDOWN(A5, 2))",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 3,
            char_offset: 0,
        },
        "=32",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 4,
            char_offset: 0,
        },
        "=ABS(41)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 0,
            char_offset: 0,
        },
        "=B7",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 1,
            char_offset: 0,
        },
        "=LEN(\"(C3 / E5)\")",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 3,
            char_offset: 0,
        },
        "=D1",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 4,
            char_offset: 0,
        },
        "=17",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 0,
            char_offset: 0,
        },
        "=SQRT((-6 / 48))",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 1,
            char_offset: 0,
        },
        "=B7",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 2,
            char_offset: 0,
        },
        "=((B9 - D7) / B3)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 3,
            char_offset: 0,
        },
        "=IF((LEN(\"E8\") > CONCATENATE(\"E1\", \"C8\")), LOWER(\"E6\"), B8)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 4,
            char_offset: 0,
        },
        "=46",
    );
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 2));
    println!("test_fuzz_reproducer_seed_25814 evaluated: {:?}", target);
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
fn test_fuzz_reproducer_seed_711187() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 0,
            char_offset: 0,
        },
        "23",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 1,
            char_offset: 0,
        },
        "\"qefv\"",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 2,
            char_offset: 0,
        },
        "97",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 3,
            char_offset: 0,
        },
        "-11",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 4,
            char_offset: 0,
        },
        "-33",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 0,
            char_offset: 0,
        },
        "148.2045",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 1,
            char_offset: 0,
        },
        "45",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 2,
            char_offset: 0,
        },
        "-272.57",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 3,
            char_offset: 0,
        },
        "0",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 4,
            char_offset: 0,
        },
        "-97",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 0,
            char_offset: 0,
        },
        "\"wiKVP\"",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 1,
            char_offset: 0,
        },
        "-58",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 2,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 3,
            char_offset: 0,
        },
        "\"3HdF\"",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 4,
            char_offset: 0,
        },
        "-54",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 0,
            char_offset: 0,
        },
        "\"iltsnHSB\"",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 2,
            char_offset: 0,
        },
        "\"m\"",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 3,
            char_offset: 0,
        },
        "-254.7",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 0,
            char_offset: 0,
        },
        "398.5356",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 1,
            char_offset: 0,
        },
        "45",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 2,
            char_offset: 0,
        },
        "76",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 3,
            char_offset: 0,
        },
        "-77",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 4,
            char_offset: 0,
        },
        "\"uTL\"",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 0,
            char_offset: 0,
        },
        "=(CONCATENATE(\"D1\", \"E2\") ^ A3)",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 1,
            char_offset: 0,
        },
        "=-31",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 2,
            char_offset: 0,
        },
        "=-7",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 3,
            char_offset: 0,
        },
        "=D4",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 4,
            char_offset: 0,
        },
        "=UPPER(\"OR(-38 > 0, 43 < 100)\")",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 0,
            char_offset: 0,
        },
        "149.9",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 1,
            char_offset: 0,
        },
        "-4",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 2,
            char_offset: 0,
        },
        "\"2ambJCe\"",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 3,
            char_offset: 0,
        },
        "=C3",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 4,
            char_offset: 0,
        },
        "=MIN(C2:D4)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 0,
            char_offset: 0,
        },
        "=ROUNDDOWN(MAX(C2:C2), 0)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 1,
            char_offset: 0,
        },
        "=B6",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 2,
            char_offset: 0,
        },
        "=CONCATENATE(\"48\", \"C7\")",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 3,
            char_offset: 0,
        },
        "62",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 4,
            char_offset: 0,
        },
        "\"Ti\"",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 0,
            char_offset: 0,
        },
        "=-46",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 1,
            char_offset: 0,
        },
        "-30",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 2,
            char_offset: 0,
        },
        "=22",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 3,
            char_offset: 0,
        },
        "=E4",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 4,
            char_offset: 0,
        },
        "=AND(E4 > 0, OR(D5 > 0, E5 < 100) < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 0,
            char_offset: 0,
        },
        "=SUM(ABS(E9), SUM(A2, -9))",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 1,
            char_offset: 0,
        },
        "=1",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 2,
            char_offset: 0,
        },
        "=LOWER(\"D8\")",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 3,
            char_offset: 0,
        },
        "=IF(((C9 / D9) > C4), AND(18 > 0, C3 < 100), SQRT(D7))",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 4,
            char_offset: 0,
        },
        "=D5",
    );
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    println!("test_fuzz_reproducer_seed_711187 evaluated: {:?}", target);
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
fn test_fuzz_reproducer_seed_870160() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 0,
            char_offset: 0,
        },
        "49",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 1,
            char_offset: 0,
        },
        "6",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 2,
            char_offset: 0,
        },
        "-26",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 3,
            char_offset: 0,
        },
        "-98",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 4,
            char_offset: 0,
        },
        "\"BnnA\"",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 0,
            char_offset: 0,
        },
        "\"lrHaQ\"",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 1,
            char_offset: 0,
        },
        "72",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 3,
            char_offset: 0,
        },
        "\"BZm\"",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 0,
            char_offset: 0,
        },
        "\"K jKmvaf\"",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 1,
            char_offset: 0,
        },
        "\"qrCkEFAQ\"",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 2,
            char_offset: 0,
        },
        "11",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 3,
            char_offset: 0,
        },
        "\"yguQaatS\"",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 4,
            char_offset: 0,
        },
        "57",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 0,
            char_offset: 0,
        },
        "146.3619",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 1,
            char_offset: 0,
        },
        "61",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 2,
            char_offset: 0,
        },
        "\"VKdXCqYO\"",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 4,
            char_offset: 0,
        },
        "-79",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 0,
            char_offset: 0,
        },
        "-37",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 2,
            char_offset: 0,
        },
        "-362.0854",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 3,
            char_offset: 0,
        },
        "63",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 4,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 0,
            char_offset: 0,
        },
        "=PRODUCT(PRODUCT(D1:D5), OR(-29 > 0, B3 < 100))",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 1,
            char_offset: 0,
        },
        "=INT(LOWER(\"-7\"))",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 2,
            char_offset: 0,
        },
        "=ABS((E5 - E5))",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 3,
            char_offset: 0,
        },
        "=-2",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 4,
            char_offset: 0,
        },
        "=-44",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 0,
            char_offset: 0,
        },
        "=ROUNDUP(AVERAGE(D6:E6), 2)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 1,
            char_offset: 0,
        },
        "=D3",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 2,
            char_offset: 0,
        },
        "=ABS(CONCATENATE(\"24\", \"D3\"))",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 3,
            char_offset: 0,
        },
        "=ROUNDUP(B5, 1)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 4,
            char_offset: 0,
        },
        "-285.5",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 0,
            char_offset: 0,
        },
        "-97",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 1,
            char_offset: 0,
        },
        "=-31",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 2,
            char_offset: 0,
        },
        "46",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 3,
            char_offset: 0,
        },
        "=4",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 4,
            char_offset: 0,
        },
        "=MAX(C6:E6)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 0,
            char_offset: 0,
        },
        "\"b\"",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 1,
            char_offset: 0,
        },
        "=A8",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 2,
            char_offset: 0,
        },
        "=SUM(E4:E7)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 3,
            char_offset: 0,
        },
        "=MIN(E6:E6)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 4,
            char_offset: 0,
        },
        "=E6",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 0,
            char_offset: 0,
        },
        "=IF((A7 > E8), (48 * E9), D7)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 1,
            char_offset: 0,
        },
        "=-32",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 2,
            char_offset: 0,
        },
        "=-33",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 3,
            char_offset: 0,
        },
        "=C5",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 4,
            char_offset: 0,
        },
        "=AND(IF((50 > -18), B3, -13) > 0, C3 < 100)",
    );
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(5, 0));
    println!("test_fuzz_reproducer_seed_870160 evaluated: {:?}", target);
    match target {
        ResultData::Float(f) => assert!(f.abs() < 1e-6),
        other => panic!("Expected Float(0.0) for A6, got {:?}", other),
    }
}

#[test]
fn test_fuzz_reproducer_seed_97218() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 0,
            char_offset: 0,
        },
        "57",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 1,
            char_offset: 0,
        },
        "0",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 2,
            char_offset: 0,
        },
        "66",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 3,
            char_offset: 0,
        },
        "-66",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 4,
            char_offset: 0,
        },
        "-36",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 0,
            char_offset: 0,
        },
        "236.7614",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 1,
            char_offset: 0,
        },
        "104",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 2,
            char_offset: 0,
        },
        "78",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 3,
            char_offset: 0,
        },
        "87",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 4,
            char_offset: 0,
        },
        "64.40000000000001",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 0,
            char_offset: 0,
        },
        "-70",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 1,
            char_offset: 0,
        },
        "7",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 3,
            char_offset: 0,
        },
        "-346.6315",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 4,
            char_offset: 0,
        },
        "-25",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 0,
            char_offset: 0,
        },
        "-447",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 1,
            char_offset: 0,
        },
        "-490.756",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 2,
            char_offset: 0,
        },
        "-94",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 3,
            char_offset: 0,
        },
        "\"2zorN\"",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 4,
            char_offset: 0,
        },
        "\"PzZo\"",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 0,
            char_offset: 0,
        },
        "-63",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 1,
            char_offset: 0,
        },
        "-29",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 2,
            char_offset: 0,
        },
        "46",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 3,
            char_offset: 0,
        },
        "\"xYPgPb\"",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 0,
            char_offset: 0,
        },
        "=ABS(ROUNDUP(C1, 2))",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 1,
            char_offset: 0,
        },
        "=-12",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 2,
            char_offset: 0,
        },
        "\"dxvW\"",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 3,
            char_offset: 0,
        },
        "=INT(SUM(2, C3))",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 4,
            char_offset: 0,
        },
        "-90",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 0,
            char_offset: 0,
        },
        "=INT(OR(B2 > 0, 20 < 100))",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 1,
            char_offset: 0,
        },
        "=AND(SUM(C4:C6) > 0, 17 < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 2,
            char_offset: 0,
        },
        "=OR(RIGHT(\"B1\", 5) > 0, INT(D4) < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 3,
            char_offset: 0,
        },
        "=-37",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 4,
            char_offset: 0,
        },
        "=AND(A3 > 0, B3 < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 0,
            char_offset: 0,
        },
        "=PRODUCT(C7:C7)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 1,
            char_offset: 0,
        },
        "-46",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 2,
            char_offset: 0,
        },
        "=E6",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 3,
            char_offset: 0,
        },
        "=42",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 4,
            char_offset: 0,
        },
        "=B3",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 0,
            char_offset: 0,
        },
        "=(ROUNDDOWN(B2, 0) + (E2 + D1))",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 2,
            char_offset: 0,
        },
        "=(B2 ^ A2)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 3,
            char_offset: 0,
        },
        "=OR(SQRT(E5) > 0, (A8 / E3) < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 4,
            char_offset: 0,
        },
        "228.3",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 0,
            char_offset: 0,
        },
        "=IF((IF((16 > B5), B6, E8) > PRODUCT(B4:E7)), 38, 36)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 1,
            char_offset: 0,
        },
        "=(E6 ^ -45)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 2,
            char_offset: 0,
        },
        "=E3",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 3,
            char_offset: 0,
        },
        "=(D6 - LEFT(\"A4\", 4))",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 4,
            char_offset: 0,
        },
        "=ROUNDUP(ROUNDDOWN(D7, 0), 1)",
    );
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 2));
    println!("test_fuzz_reproducer_seed_97218 evaluated: {:?}", target);
    assert!(
        matches!(target, ResultData::Error(_)),
        "Expected Error for C7, got {:?}",
        target
    );
}

#[test]
fn test_fuzz_reproducer_seed_230672() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 0,
            char_offset: 0,
        },
        "\"jarupzx\"",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 1,
            char_offset: 0,
        },
        "\"WL\"",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 2,
            char_offset: 0,
        },
        "-307.77",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 3,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 4,
            char_offset: 0,
        },
        "39",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 0,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 1,
            char_offset: 0,
        },
        "\"eTB\"",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 2,
            char_offset: 0,
        },
        "-31",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 3,
            char_offset: 0,
        },
        "92",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 4,
            char_offset: 0,
        },
        "-54",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 0,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 1,
            char_offset: 0,
        },
        "65",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 2,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 3,
            char_offset: 0,
        },
        "357",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 4,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 0,
            char_offset: 0,
        },
        "-65",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 1,
            char_offset: 0,
        },
        "76",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 2,
            char_offset: 0,
        },
        "-58",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 3,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 4,
            char_offset: 0,
        },
        "\"oT\"",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 0,
            char_offset: 0,
        },
        "275.93",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 1,
            char_offset: 0,
        },
        "434.801",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 2,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 3,
            char_offset: 0,
        },
        "440.643",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 4,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 0,
            char_offset: 0,
        },
        "=PRODUCT(ROUND(A1, 1), SQRT(C5))",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 1,
            char_offset: 0,
        },
        "=A1",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 2,
            char_offset: 0,
        },
        "=D4",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 3,
            char_offset: 0,
        },
        "=OR(AVERAGE(E2:E4) > 0, D2 < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 4,
            char_offset: 0,
        },
        "=E3",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 0,
            char_offset: 0,
        },
        "=ROUNDDOWN(OR(48 > 0, B4 < 100), 2)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 1,
            char_offset: 0,
        },
        "=-49",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 2,
            char_offset: 0,
        },
        "=LEN(\"(33 + E3)\")",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 3,
            char_offset: 0,
        },
        "=-28",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 4,
            char_offset: 0,
        },
        "=(-17 + D3)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 0,
            char_offset: 0,
        },
        "=ABS(B7)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 1,
            char_offset: 0,
        },
        "=ROUNDDOWN(LOWER(\"B2\"), 1)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 2,
            char_offset: 0,
        },
        "=CONCATENATE(\"A3\", \"AND(D1 > 0, 32 < 100)\")",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 3,
            char_offset: 0,
        },
        "=UPPER(\"ROUND(E1, 2)\")",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 4,
            char_offset: 0,
        },
        "=B7",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 0,
            char_offset: 0,
        },
        "=A7",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 1,
            char_offset: 0,
        },
        "=(A7 + IF((E8 > -34), A5, D4))",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 2,
            char_offset: 0,
        },
        "=RIGHT(\"E1\", 5)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 3,
            char_offset: 0,
        },
        "=CONCATENATE(\"E4\", \"D6\")",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 4,
            char_offset: 0,
        },
        "=LEN(\"C5\")",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 0,
            char_offset: 0,
        },
        "-26",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 1,
            char_offset: 0,
        },
        "=UPPER(\"-42\")",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 2,
            char_offset: 0,
        },
        "=LEFT(\"ROUND(B6, 0)\", 5)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 3,
            char_offset: 0,
        },
        "\"GCdw2QD\"",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 4,
            char_offset: 0,
        },
        "=(-41 / SUM(D5, -7))",
    );
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 0));
    println!("test_fuzz_reproducer_seed_230672 evaluated: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 49.0),
        other => panic!("Expected Float(49.0) for A8, got {:?}", other),
    }
}

#[test]
fn test_fuzz_reproducer_seed_140247() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 0,
            char_offset: 0,
        },
        "\"fuKg\"",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 1,
            char_offset: 0,
        },
        "74",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 2,
            char_offset: 0,
        },
        "-75",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 3,
            char_offset: 0,
        },
        "-37",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 4,
            char_offset: 0,
        },
        "29",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 0,
            char_offset: 0,
        },
        "6",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 1,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 2,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 4,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 0,
            char_offset: 0,
        },
        "-327.3",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 1,
            char_offset: 0,
        },
        "-27",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 2,
            char_offset: 0,
        },
        "87",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 3,
            char_offset: 0,
        },
        "-70",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 0,
            char_offset: 0,
        },
        "\"Ec\"",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 1,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 2,
            char_offset: 0,
        },
        "41",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 3,
            char_offset: 0,
        },
        "267.6832",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 0,
            char_offset: 0,
        },
        "148.845",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 2,
            char_offset: 0,
        },
        "-435",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 3,
            char_offset: 0,
        },
        "-344.315",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 4,
            char_offset: 0,
        },
        "28",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 0,
            char_offset: 0,
        },
        "=LOWER(\"30\")",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 1,
            char_offset: 0,
        },
        "=UPPER(\"(D2 * -2)\")",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 2,
            char_offset: 0,
        },
        "=(MIN(E3, -6) ^ ROUNDUP(24, 0))",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 3,
            char_offset: 0,
        },
        "=C2",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 4,
            char_offset: 0,
        },
        "=(LEN(\"D3\") / -34)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 0,
            char_offset: 0,
        },
        "17",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 1,
            char_offset: 0,
        },
        "=IF((-12 > SUM(E5, A1)), (D1 ^ A6), E5)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 2,
            char_offset: 0,
        },
        "=C1",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 3,
            char_offset: 0,
        },
        "=ABS(D2)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 4,
            char_offset: 0,
        },
        "=42",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 0,
            char_offset: 0,
        },
        "=(IF((B2 > D4), C6, E6) ^ ROUNDUP(17, 0))",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 1,
            char_offset: 0,
        },
        "=MAX(18, (D7 * 39))",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 2,
            char_offset: 0,
        },
        "=ROUND(LEN(\"C7\"), 1)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 3,
            char_offset: 0,
        },
        "=(IF((28 > -30), A1, C6) / D2)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 4,
            char_offset: 0,
        },
        "=LEN(\"AND(C6 > 0, A7 < 100)\")",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 0,
            char_offset: 0,
        },
        "=C7",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 1,
            char_offset: 0,
        },
        "-255.8071",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 2,
            char_offset: 0,
        },
        "=B1",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 3,
            char_offset: 0,
        },
        "-85",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 4,
            char_offset: 0,
        },
        "=AND(SUM(D1:D8) > 0, ABS(D4) < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 0,
            char_offset: 0,
        },
        "=IF((25 > OR(43 > 0, -33 < 100)), ABS(-13), LOWER(\"E6\"))",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 1,
            char_offset: 0,
        },
        "=(LEFT(\"8\", 1) * IF((21 > A6), C2, 47))",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 2,
            char_offset: 0,
        },
        "=C6",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 3,
            char_offset: 0,
        },
        "=48",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 4,
            char_offset: 0,
        },
        "=(RIGHT(\"A1\", 2) + ROUNDDOWN(E1, 1))",
    );
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 1));
    println!("test_fuzz_reproducer_seed_140247 evaluated: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 28.0),
        ResultData::Integer(i) => assert_eq!(i, 28),
        other => panic!("Expected Float(28.0) for B7, got {:?}", other),
    }
}

#[test]
fn test_fuzz_reproducer_seed_233445() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 0,
            char_offset: 0,
        },
        "-52",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 1,
            char_offset: 0,
        },
        "347.193",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 2,
            char_offset: 0,
        },
        "14",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 3,
            char_offset: 0,
        },
        "6",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 4,
            char_offset: 0,
        },
        "4",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 0,
            char_offset: 0,
        },
        "-466",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 1,
            char_offset: 0,
        },
        "-4",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 2,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 3,
            char_offset: 0,
        },
        "\"y\"",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 0,
            char_offset: 0,
        },
        "472",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 1,
            char_offset: 0,
        },
        "-36",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 2,
            char_offset: 0,
        },
        "-10",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 3,
            char_offset: 0,
        },
        "-28.908",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 4,
            char_offset: 0,
        },
        "267",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 0,
            char_offset: 0,
        },
        "\"vfclM\"",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 1,
            char_offset: 0,
        },
        "-60",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 2,
            char_offset: 0,
        },
        "414",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 3,
            char_offset: 0,
        },
        "\"TiNHbw\"",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 4,
            char_offset: 0,
        },
        "-28",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 0,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 1,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 3,
            char_offset: 0,
        },
        "\"LlUvCUn\"",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 4,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 0,
            char_offset: 0,
        },
        "=(RIGHT(\"B2\", 3) / (C5 / D3))",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 1,
            char_offset: 0,
        },
        "=SQRT(LOWER(\"D5\"))",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 2,
            char_offset: 0,
        },
        "=AND(B4 > 0, -31 < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 3,
            char_offset: 0,
        },
        "=D2",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 4,
            char_offset: 0,
        },
        "=MAX((A3 + 34), B1)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 0,
            char_offset: 0,
        },
        "=E6",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 1,
            char_offset: 0,
        },
        "=((C6 / E2) / E1)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 2,
            char_offset: 0,
        },
        "=-2",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 3,
            char_offset: 0,
        },
        "=A5",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 4,
            char_offset: 0,
        },
        "=E2",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 0,
            char_offset: 0,
        },
        "=C7",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 1,
            char_offset: 0,
        },
        "=MIN(A1:A3)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 2,
            char_offset: 0,
        },
        "=SQRT(ROUNDDOWN(E4, 1))",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 3,
            char_offset: 0,
        },
        "76",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 4,
            char_offset: 0,
        },
        "8",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 0,
            char_offset: 0,
        },
        "=MAX(PRODUCT(D7, B6), IF((E7 > E4), C1, B8))",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 1,
            char_offset: 0,
        },
        "328",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 2,
            char_offset: 0,
        },
        "=MIN(D4:D5)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 3,
            char_offset: 0,
        },
        "0",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 4,
            char_offset: 0,
        },
        "=((C8 / E3) / ROUND(-21, 0))",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 0,
            char_offset: 0,
        },
        "=-39",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 1,
            char_offset: 0,
        },
        "=(CONCATENATE(\"-30\", \"A6\") - LEN(\"D5\"))",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 2,
            char_offset: 0,
        },
        "138.2545",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 3,
            char_offset: 0,
        },
        "=36",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 4,
            char_offset: 0,
        },
        "=MIN(C1:C9)",
    );
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 1));
    println!("test_fuzz_reproducer_seed_233445 evaluated: {:?}", target);
    assert!(
        matches!(target, ResultData::Error(ref e) if e.contains("#DIV/0!")),
        "Expected #DIV/0! for B7, got {:?}",
        target
    );
}

#[test]
fn test_fuzz_reproducer_seed_58482() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 0,
            char_offset: 0,
        },
        "74",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 1,
            char_offset: 0,
        },
        "-287.148",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 3,
            char_offset: 0,
        },
        "\"HrCRG\"",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 4,
            char_offset: 0,
        },
        "167.3",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 0,
            char_offset: 0,
        },
        "-66.04900000000001",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 1,
            char_offset: 0,
        },
        "-208.0246",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 2,
            char_offset: 0,
        },
        "56",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 3,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 0,
            char_offset: 0,
        },
        "\"RzGODC2\"",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 1,
            char_offset: 0,
        },
        "-26",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 2,
            char_offset: 0,
        },
        "335.983",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 3,
            char_offset: 0,
        },
        "-11",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 4,
            char_offset: 0,
        },
        "273.331",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 0,
            char_offset: 0,
        },
        "0",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 1,
            char_offset: 0,
        },
        "\"HdRvUPf\"",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 2,
            char_offset: 0,
        },
        "-55",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 3,
            char_offset: 0,
        },
        "486.727",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 4,
            char_offset: 0,
        },
        "209.55",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 0,
            char_offset: 0,
        },
        "\"yHLLKWgU\"",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 1,
            char_offset: 0,
        },
        "-34",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 2,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 3,
            char_offset: 0,
        },
        "\"mPUQspYq\"",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 4,
            char_offset: 0,
        },
        "-102.928",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 0,
            char_offset: 0,
        },
        "=A3",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 1,
            char_offset: 0,
        },
        "=CONCATENATE(\"MIN(B2:E4)\", \"A5\")",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 2,
            char_offset: 0,
        },
        "=-40",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 3,
            char_offset: 0,
        },
        "=C4",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 4,
            char_offset: 0,
        },
        "=AVERAGE(B4, B2)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 0,
            char_offset: 0,
        },
        "-440.6",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 1,
            char_offset: 0,
        },
        "=-48",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 2,
            char_offset: 0,
        },
        "=(SQRT(40) - ROUND(E5, 1))",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 4,
            char_offset: 0,
        },
        "=B3",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 0,
            char_offset: 0,
        },
        "=OR(IF((B7 > C1), C1, D3) > 0, IF((22 > B4), A5, C2) < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 1,
            char_offset: 0,
        },
        "=ROUNDUP(E5, 2)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 2,
            char_offset: 0,
        },
        "=(RIGHT(\"A5\", 2) + CONCATENATE(\"C1\", \"C7\"))",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 3,
            char_offset: 0,
        },
        "=RIGHT(\"19\", 2)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 4,
            char_offset: 0,
        },
        "=ABS(ROUNDDOWN(18, 2))",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 0,
            char_offset: 0,
        },
        "=(AND(B3 > 0, D5 < 100) * (D3 + C6))",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 1,
            char_offset: 0,
        },
        "=49",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 2,
            char_offset: 0,
        },
        "9",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 3,
            char_offset: 0,
        },
        "=(OR(C3 > 0, A4 < 100) + 29)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 4,
            char_offset: 0,
        },
        "=-11",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 0,
            char_offset: 0,
        },
        "=D5",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 1,
            char_offset: 0,
        },
        "=LEN(\"C4\")",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 2,
            char_offset: 0,
        },
        "=IF((IF((-37 > -30), E3, -39) > SUM(A7:A8)), B6, C7)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 3,
            char_offset: 0,
        },
        "=D4",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 4,
            char_offset: 0,
        },
        "=A5",
    );
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 0));
    println!("test_fuzz_reproducer_seed_58482 evaluated: {:?}", target);
    match target {
        ResultData::Float(f) => assert!(f.abs() < 1e-6),
        other => panic!("Expected Float(0.0) for A9, got {:?}", other),
    }
}

#[test]
fn test_fuzz_reproducer_seed_867362() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 1,
            char_offset: 0,
        },
        "61",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 2,
            char_offset: 0,
        },
        "\"ha\"",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 3,
            char_offset: 0,
        },
        "97",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 4,
            char_offset: 0,
        },
        "166",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 0,
            char_offset: 0,
        },
        "-8",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 1,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 2,
            char_offset: 0,
        },
        "-374.6",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 3,
            char_offset: 0,
        },
        "424",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 4,
            char_offset: 0,
        },
        "328.132",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 0,
            char_offset: 0,
        },
        "21",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 1,
            char_offset: 0,
        },
        "-100",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 2,
            char_offset: 0,
        },
        "-97",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 3,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 4,
            char_offset: 0,
        },
        "-10",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 0,
            char_offset: 0,
        },
        "83",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 1,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 2,
            char_offset: 0,
        },
        "-9",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 3,
            char_offset: 0,
        },
        "178.26",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 4,
            char_offset: 0,
        },
        "-80",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 0,
            char_offset: 0,
        },
        "\"2H\"",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 1,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 2,
            char_offset: 0,
        },
        "\"qn\"",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 3,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 4,
            char_offset: 0,
        },
        "4",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 0,
            char_offset: 0,
        },
        "=4",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 1,
            char_offset: 0,
        },
        "2",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 2,
            char_offset: 0,
        },
        "=MAX(A5:A5)",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 3,
            char_offset: 0,
        },
        "=A2",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 4,
            char_offset: 0,
        },
        "=UPPER(\"D4\")",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 0,
            char_offset: 0,
        },
        "=AND(-15 > 0, INT(B2) < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 1,
            char_offset: 0,
        },
        "=B3",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 2,
            char_offset: 0,
        },
        "=A2",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 3,
            char_offset: 0,
        },
        "=LOWER(\"B5\")",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 4,
            char_offset: 0,
        },
        "=A5",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 0,
            char_offset: 0,
        },
        "=ABS(ROUNDUP(34, 1))",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 1,
            char_offset: 0,
        },
        "=(RIGHT(\"A7\", 2) - (C7 - B4))",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 2,
            char_offset: 0,
        },
        "=B2",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 3,
            char_offset: 0,
        },
        "=A1",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 4,
            char_offset: 0,
        },
        "\"kiK ga\"",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 0,
            char_offset: 0,
        },
        "=MIN(MAX(-43, E8), -19)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 1,
            char_offset: 0,
        },
        "=(ROUNDUP(B4, 2) / IF((C1 > -33), 7, 19))",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 2,
            char_offset: 0,
        },
        "=AND(D6 > 0, ROUND(D8, 2) < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 3,
            char_offset: 0,
        },
        "=SQRT(D8)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 4,
            char_offset: 0,
        },
        "47",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 0,
            char_offset: 0,
        },
        "=ABS(INT(A3))",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 1,
            char_offset: 0,
        },
        "=(C9 + -2)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 2,
            char_offset: 0,
        },
        "=MIN(E1:E7)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 3,
            char_offset: 0,
        },
        "436.831",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 4,
            char_offset: 0,
        },
        "=LOWER(\"-30\")",
    );
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 1));
    println!("test_fuzz_reproducer_seed_867362 evaluated: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, -2.0),
        other => panic!("Expected Float(-2.0) for B10, got {:?}", other),
    }
}

#[test]
fn test_fuzz_reproducer_seed_497384() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 0,
            char_offset: 0,
        },
        "-72",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 1,
            char_offset: 0,
        },
        "203.37",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 2,
            char_offset: 0,
        },
        "\"COp\"",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 3,
            char_offset: 0,
        },
        "\"eNZIbOll\"",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 4,
            char_offset: 0,
        },
        "\"CscRQlf\"",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 0,
            char_offset: 0,
        },
        "-75",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 1,
            char_offset: 0,
        },
        "-151",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 2,
            char_offset: 0,
        },
        "-79",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 3,
            char_offset: 0,
        },
        "296.7224",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 4,
            char_offset: 0,
        },
        "\"C\"",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 0,
            char_offset: 0,
        },
        "474.8",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 1,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 2,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 3,
            char_offset: 0,
        },
        "-386.7184",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 4,
            char_offset: 0,
        },
        "-265.71",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 0,
            char_offset: 0,
        },
        "364.52",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 1,
            char_offset: 0,
        },
        "51",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 2,
            char_offset: 0,
        },
        "3",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 3,
            char_offset: 0,
        },
        "-303.122",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 4,
            char_offset: 0,
        },
        "-12",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 0,
            char_offset: 0,
        },
        "-307.44",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 1,
            char_offset: 0,
        },
        "46",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 2,
            char_offset: 0,
        },
        "-125.6",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 3,
            char_offset: 0,
        },
        "-97",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 0,
            char_offset: 0,
        },
        "=-3",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 1,
            char_offset: 0,
        },
        "-219",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 2,
            char_offset: 0,
        },
        "=C3",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 3,
            char_offset: 0,
        },
        "\"pHXeHmLw\"",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 4,
            char_offset: 0,
        },
        "=PRODUCT(-1, A5)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 0,
            char_offset: 0,
        },
        "1",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 1,
            char_offset: 0,
        },
        "-357.5857",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 2,
            char_offset: 0,
        },
        "=ABS(B1)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 3,
            char_offset: 0,
        },
        "1",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 4,
            char_offset: 0,
        },
        "=SQRT(AVERAGE(E2:E3))",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 0,
            char_offset: 0,
        },
        "=ROUND((C2 + -15), 0)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 1,
            char_offset: 0,
        },
        "=AVERAGE(B5:E6)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 2,
            char_offset: 0,
        },
        "=E1",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 3,
            char_offset: 0,
        },
        "=MAX(SUM(-38, C1), LOWER(\"E7\"))",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 4,
            char_offset: 0,
        },
        "7",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 0,
            char_offset: 0,
        },
        "=SUM(INT(A8), UPPER(\"-11\"))",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 1,
            char_offset: 0,
        },
        "=ABS(C2)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 2,
            char_offset: 0,
        },
        "428.9",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 3,
            char_offset: 0,
        },
        "=ABS(IF((B5 > 18), 5, B6))",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 4,
            char_offset: 0,
        },
        "=ROUNDUP(C1, 0)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 0,
            char_offset: 0,
        },
        "=RIGHT(\"AND(E9 > 0, 45 < 100)\", 4)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 1,
            char_offset: 0,
        },
        "=IF((LEN(\"B8\") > B8), A7, E6)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 2,
            char_offset: 0,
        },
        "=D3",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 3,
            char_offset: 0,
        },
        "=MIN(A5:E9)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 4,
            char_offset: 0,
        },
        "=((B6 - A5) / ROUND(-38, 0))",
    );
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 0));
    println!("test_fuzz_reproducer_seed_497384 evaluated: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, -105.0),
        other => panic!("Expected Float(-105.0) for A9, got {:?}", other),
    }
}

#[test]
fn test_fuzz_reproducer_seed_302307() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 0,
            char_offset: 0,
        },
        "46",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 1,
            char_offset: 0,
        },
        "-124.46",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 2,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 3,
            char_offset: 0,
        },
        "-56",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 4,
            char_offset: 0,
        },
        "75",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 0,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 1,
            char_offset: 0,
        },
        "\"PJwI\"",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 3,
            char_offset: 0,
        },
        "75",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 4,
            char_offset: 0,
        },
        "3",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 0,
            char_offset: 0,
        },
        "\"W1bdh\"",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 1,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 2,
            char_offset: 0,
        },
        "0",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 3,
            char_offset: 0,
        },
        "\"hOS\"",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 4,
            char_offset: 0,
        },
        "10",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 0,
            char_offset: 0,
        },
        "\"lnxfR\"",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 1,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 2,
            char_offset: 0,
        },
        "-477.673",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 3,
            char_offset: 0,
        },
        "\"QYQf\"",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 4,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 0,
            char_offset: 0,
        },
        "\"MB\"",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 1,
            char_offset: 0,
        },
        "109.944",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 2,
            char_offset: 0,
        },
        "8",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 3,
            char_offset: 0,
        },
        "25",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 4,
            char_offset: 0,
        },
        "-63",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 0,
            char_offset: 0,
        },
        "=IF((ROUND(E3, 1) > (-42 * E1)), (A4 / 0), IF((-23 > D5), C3, 49))",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 1,
            char_offset: 0,
        },
        "=LOWER(\"AVERAGE(A4:A4)\")",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 2,
            char_offset: 0,
        },
        "=OR(E4 > 0, LEN(\"C2\") < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 3,
            char_offset: 0,
        },
        "\"TADFiW\"",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 4,
            char_offset: 0,
        },
        "=RIGHT(\"D2\", 4)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 0,
            char_offset: 0,
        },
        "=UPPER(\"(24 - E1)\")",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 1,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 2,
            char_offset: 0,
        },
        "=PRODUCT(UPPER(\"D2\"), C1)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 3,
            char_offset: 0,
        },
        "1",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 4,
            char_offset: 0,
        },
        "=PRODUCT(D1:E5)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 0,
            char_offset: 0,
        },
        "=A6",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 1,
            char_offset: 0,
        },
        "=((48 * D4) ^ C1)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 2,
            char_offset: 0,
        },
        "=C2",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 3,
            char_offset: 0,
        },
        "=IF((ABS(A2) > RIGHT(\"D2\", 2)), B5, 9)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 4,
            char_offset: 0,
        },
        "=AND(LEN(\"-5\") > 0, E3 < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 0,
            char_offset: 0,
        },
        "=CONCATENATE(\"ROUND(A1, 1)\", \"AND(D7 > 0, 44 < 100)\")",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 1,
            char_offset: 0,
        },
        "=E1",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 2,
            char_offset: 0,
        },
        "=-17",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 3,
            char_offset: 0,
        },
        "=(IF((20 > -39), -2, A7) / B6)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 4,
            char_offset: 0,
        },
        "=-2",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 0,
            char_offset: 0,
        },
        "=OR((-19 ^ C8) > 0, 35 < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 1,
            char_offset: 0,
        },
        "=-7",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 2,
            char_offset: 0,
        },
        "=ROUND(D3, 1)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 3,
            char_offset: 0,
        },
        "=SUM(E8:E9)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 4,
            char_offset: 0,
        },
        "=D8",
    );
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    println!("test_fuzz_reproducer_seed_302307 evaluated: {:?}", target);
    match target {
        ResultData::Boolean(b) => assert_eq!(b, true),
        other => panic!("Expected Boolean(true) for A10, got {:?}", other),
    }
}

#[test]
fn test_fuzz_reproducer_seed_486091() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 0,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 1,
            char_offset: 0,
        },
        "-55",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 2,
            char_offset: 0,
        },
        "28",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 3,
            char_offset: 0,
        },
        "14",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 4,
            char_offset: 0,
        },
        "77",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 0,
            char_offset: 0,
        },
        "52",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 1,
            char_offset: 0,
        },
        "483.04",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 2,
            char_offset: 0,
        },
        "30",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 3,
            char_offset: 0,
        },
        "-324.1",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 4,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 0,
            char_offset: 0,
        },
        "-54",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 2,
            char_offset: 0,
        },
        "483",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 3,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 4,
            char_offset: 0,
        },
        "-272",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 0,
            char_offset: 0,
        },
        "\"CHue\"",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 1,
            char_offset: 0,
        },
        "\"xqit\"",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 2,
            char_offset: 0,
        },
        "349",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 3,
            char_offset: 0,
        },
        "\"uyDfgqA\"",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 4,
            char_offset: 0,
        },
        "-86",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 0,
            char_offset: 0,
        },
        "-72",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 1,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 2,
            char_offset: 0,
        },
        "87",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 4,
            char_offset: 0,
        },
        "-13",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 0,
            char_offset: 0,
        },
        "-67.511",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 1,
            char_offset: 0,
        },
        "=26",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 2,
            char_offset: 0,
        },
        "=C2",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 3,
            char_offset: 0,
        },
        "=OR(RIGHT(\"-32\", 2) > 0, A4 < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 4,
            char_offset: 0,
        },
        "408",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 0,
            char_offset: 0,
        },
        "\"JRW\"",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 1,
            char_offset: 0,
        },
        "=E4",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 2,
            char_offset: 0,
        },
        "=29",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 3,
            char_offset: 0,
        },
        "=SQRT(-8)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 4,
            char_offset: 0,
        },
        "=C4",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 0,
            char_offset: 0,
        },
        "=MAX(LOWER(\"-2\"), AND(17 > 0, A7 < 100))",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 1,
            char_offset: 0,
        },
        "=D1",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 2,
            char_offset: 0,
        },
        "98.194",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 3,
            char_offset: 0,
        },
        "-2",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 4,
            char_offset: 0,
        },
        "=ABS(A6)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 0,
            char_offset: 0,
        },
        "=E1",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 1,
            char_offset: 0,
        },
        "=MIN(B1:E3)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 2,
            char_offset: 0,
        },
        "\"YbeYuyK\"",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 3,
            char_offset: 0,
        },
        "=E3",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 4,
            char_offset: 0,
        },
        "5",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 0,
            char_offset: 0,
        },
        "=(PRODUCT(C3, A2) + (D7 + C9))",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 1,
            char_offset: 0,
        },
        "=E3",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 2,
            char_offset: 0,
        },
        "=C2",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 3,
            char_offset: 0,
        },
        "=IF((ROUNDUP(25, 0) > B6), IF((43 > -8), -42, 28), C8)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 4,
            char_offset: 0,
        },
        "=(AND(D1 > 0, D5 < 100) * A5)",
    );
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    println!("test_fuzz_reproducer_seed_486091 evaluated: {:?}", target);
    assert!(
        matches!(target, ResultData::Error(ref e) if e.contains("#NUM!")),
        "Expected #NUM! for A10, got {:?}",
        target
    );
}

#[test]
fn test_fuzz_reproducer_seed_995940() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 1,
            char_offset: 0,
        },
        "-83",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 3,
            char_offset: 0,
        },
        "-3",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 4,
            char_offset: 0,
        },
        "-290.5375",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 0,
            char_offset: 0,
        },
        "95",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 1,
            char_offset: 0,
        },
        "60.005",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 2,
            char_offset: 0,
        },
        "\"KQdrYUl\"",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 3,
            char_offset: 0,
        },
        "-13",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 4,
            char_offset: 0,
        },
        "-70",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 0,
            char_offset: 0,
        },
        "-437.94",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 1,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 2,
            char_offset: 0,
        },
        "55",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 3,
            char_offset: 0,
        },
        "84",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 4,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 0,
            char_offset: 0,
        },
        "-181.7",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 1,
            char_offset: 0,
        },
        "-68",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 2,
            char_offset: 0,
        },
        "\"3z\"",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 3,
            char_offset: 0,
        },
        "-332.5255",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 4,
            char_offset: 0,
        },
        "-27",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 0,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 1,
            char_offset: 0,
        },
        "100",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 2,
            char_offset: 0,
        },
        "-14.34",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 3,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 4,
            char_offset: 0,
        },
        "-96",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 0,
            char_offset: 0,
        },
        "=-43",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 1,
            char_offset: 0,
        },
        "=OR(6 > 0, SUM(E3, -4) < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 2,
            char_offset: 0,
        },
        "=D1",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 3,
            char_offset: 0,
        },
        "=A1",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 4,
            char_offset: 0,
        },
        "=E2",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 0,
            char_offset: 0,
        },
        "=SQRT(LEFT(\"C2\", 5))",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 1,
            char_offset: 0,
        },
        "57",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 2,
            char_offset: 0,
        },
        "=B6",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 3,
            char_offset: 0,
        },
        "=ROUNDDOWN(OR(13 > 0, A4 < 100), 2)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 4,
            char_offset: 0,
        },
        "=-46",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 0,
            char_offset: 0,
        },
        "=A1",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 1,
            char_offset: 0,
        },
        "=AVERAGE(B5:D6)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 2,
            char_offset: 0,
        },
        "=IF((D5 > C3), B2, E3)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 3,
            char_offset: 0,
        },
        "303.5083",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 4,
            char_offset: 0,
        },
        "=C1",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 0,
            char_offset: 0,
        },
        "=LEFT(\"AND(D6 > 0, 16 < 100)\", 4)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 1,
            char_offset: 0,
        },
        "=E7",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 2,
            char_offset: 0,
        },
        "=A5",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 3,
            char_offset: 0,
        },
        "79",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 4,
            char_offset: 0,
        },
        "=15",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 0,
            char_offset: 0,
        },
        "=RIGHT(\"C6\", 1)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 1,
            char_offset: 0,
        },
        "=C3",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 2,
            char_offset: 0,
        },
        "=(ROUNDUP(D9, 1) + LEN(\"42\"))",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 3,
            char_offset: 0,
        },
        "=(D5 + LEN(\"14\"))",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 4,
            char_offset: 0,
        },
        "=PRODUCT(E9:E9)",
    );
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 1));
    println!("test_fuzz_reproducer_seed_995940 evaluated: {:?}", target);
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
fn test_fuzz_reproducer_seed_538533() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 0,
            char_offset: 0,
        },
        "\"2rLzL\"",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 1,
            char_offset: 0,
        },
        "19",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 2,
            char_offset: 0,
        },
        "1",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 3,
            char_offset: 0,
        },
        "78",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 4,
            char_offset: 0,
        },
        "-442.67",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 0,
            char_offset: 0,
        },
        "3",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 1,
            char_offset: 0,
        },
        "\" QAap\"",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 2,
            char_offset: 0,
        },
        "0",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 3,
            char_offset: 0,
        },
        "\"q\"",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 4,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 0,
            char_offset: 0,
        },
        "195.715",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 1,
            char_offset: 0,
        },
        "-239.5",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 2,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 3,
            char_offset: 0,
        },
        "52",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 1,
            char_offset: 0,
        },
        "\"L qPj\"",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 2,
            char_offset: 0,
        },
        "-24.5812",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 3,
            char_offset: 0,
        },
        "97",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 4,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 0,
            char_offset: 0,
        },
        "46",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 1,
            char_offset: 0,
        },
        "\"yQpUe\"",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 2,
            char_offset: 0,
        },
        "-51",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 3,
            char_offset: 0,
        },
        "30",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 4,
            char_offset: 0,
        },
        "\"jXFiCL\"",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 0,
            char_offset: 0,
        },
        "=(D5 ^ D3)",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 1,
            char_offset: 0,
        },
        "=LEFT(\"(A4 * D4)\", 4)",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 2,
            char_offset: 0,
        },
        "=ROUNDDOWN((A1 * E3), 2)",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 3,
            char_offset: 0,
        },
        "2",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 4,
            char_offset: 0,
        },
        "=(CONCATENATE(\"E1\", \"-30\") / C2)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 0,
            char_offset: 0,
        },
        "=B4",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 1,
            char_offset: 0,
        },
        "=AND(D4 > 0, D6 < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 2,
            char_offset: 0,
        },
        "=ABS(ROUNDDOWN(A2, 0))",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 3,
            char_offset: 0,
        },
        "=(ROUND(E4, 2) - LEN(\"E4\"))",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 4,
            char_offset: 0,
        },
        "\"1dDknr\"",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 0,
            char_offset: 0,
        },
        "=AVERAGE((B7 + D6), ROUNDDOWN(D4, 0))",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 1,
            char_offset: 0,
        },
        "=(ROUNDDOWN(D1, 2) * 6)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 2,
            char_offset: 0,
        },
        "=PRODUCT(PRODUCT(A1:A5), E2)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 3,
            char_offset: 0,
        },
        "=A5",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 4,
            char_offset: 0,
        },
        "=43",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 0,
            char_offset: 0,
        },
        "=SUM(B1:E6)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 1,
            char_offset: 0,
        },
        "=C2",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 2,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 3,
            char_offset: 0,
        },
        "=A3",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 4,
            char_offset: 0,
        },
        "=SUM(B7:D7)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 0,
            char_offset: 0,
        },
        "=SUM(A7:A8)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 1,
            char_offset: 0,
        },
        "=LEFT(\"(C5 * -4)\", 3)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 2,
            char_offset: 0,
        },
        "=-9",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 3,
            char_offset: 0,
        },
        "=IF(((E8 / E6) > -41), C2, IF((D3 > -31), B6, -11))",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 4,
            char_offset: 0,
        },
        "=AND(LEFT(\"-37\", 3) > 0, MAX(D3:E8) < 100)",
    );
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 0));
    println!("test_fuzz_reproducer_seed_538533 evaluated: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 50.0),
        other => panic!("Expected Float(50.0) for A8, got {:?}", other),
    }
}

#[test]
fn test_fuzz_reproducer_seed_971398() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 0,
            char_offset: 0,
        },
        "\"QZaPVt\"",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 1,
            char_offset: 0,
        },
        "0",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 3,
            char_offset: 0,
        },
        "\"wjDauh\"",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 4,
            char_offset: 0,
        },
        "275.684",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 0,
            char_offset: 0,
        },
        "-400.3",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 1,
            char_offset: 0,
        },
        "-99",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 2,
            char_offset: 0,
        },
        "-56",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 3,
            char_offset: 0,
        },
        "35",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 4,
            char_offset: 0,
        },
        "-33",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 0,
            char_offset: 0,
        },
        "\"rkTuS\"",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 1,
            char_offset: 0,
        },
        "-491.12",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 2,
            char_offset: 0,
        },
        "65",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 3,
            char_offset: 0,
        },
        "118.7921",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 4,
            char_offset: 0,
        },
        "20",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 0,
            char_offset: 0,
        },
        "6",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 1,
            char_offset: 0,
        },
        "95",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 2,
            char_offset: 0,
        },
        "\" uQbVFQI\"",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 3,
            char_offset: 0,
        },
        "7",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 4,
            char_offset: 0,
        },
        "-54",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 0,
            char_offset: 0,
        },
        "211.4",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 1,
            char_offset: 0,
        },
        "-17",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 2,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 3,
            char_offset: 0,
        },
        "421.516",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 4,
            char_offset: 0,
        },
        "-73",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 0,
            char_offset: 0,
        },
        "=ABS(E4)",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 1,
            char_offset: 0,
        },
        "=B1",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 2,
            char_offset: 0,
        },
        "=(ROUNDUP(A1, 0) / UPPER(\"B4\"))",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 3,
            char_offset: 0,
        },
        "=C4",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 4,
            char_offset: 0,
        },
        "=MIN(D5:E5)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 0,
            char_offset: 0,
        },
        "=CONCATENATE(\"MAX(E3:E6)\", \"OR(A4 > 0, E1 < 100)\")",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 1,
            char_offset: 0,
        },
        "=IF((-29 > B2), E2, ROUND(B6, 2))",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 2,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 3,
            char_offset: 0,
        },
        "19",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 4,
            char_offset: 0,
        },
        "=ROUNDUP(C5, 2)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 0,
            char_offset: 0,
        },
        "=(E3 ^ B6)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 1,
            char_offset: 0,
        },
        "162.149",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 2,
            char_offset: 0,
        },
        "=A5",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 3,
            char_offset: 0,
        },
        "=B1",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 4,
            char_offset: 0,
        },
        "=B7",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 0,
            char_offset: 0,
        },
        "=B3",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 1,
            char_offset: 0,
        },
        "=(LOWER(\"37\") ^ SQRT(B3))",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 2,
            char_offset: 0,
        },
        "=LEFT(\"E2\", 5)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 3,
            char_offset: 0,
        },
        "=(LOWER(\"D6\") * B4)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 4,
            char_offset: 0,
        },
        "=(E8 - 35)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 0,
            char_offset: 0,
        },
        "=B8",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 1,
            char_offset: 0,
        },
        "=(RIGHT(\"C6\", 1) + SQRT(-21))",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 2,
            char_offset: 0,
        },
        "=SQRT(PRODUCT(A2:C4))",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 3,
            char_offset: 0,
        },
        "=D9",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 4,
            char_offset: 0,
        },
        "=28",
    );
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 0));
    println!("test_fuzz_reproducer_seed_971398 evaluated: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 1.0),
        other => panic!("Expected Float(1.0) for A8, got {:?}", other),
    }
}

#[test]
fn test_fuzz_reproducer_seed_450293() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 0,
            char_offset: 0,
        },
        "7",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 1,
            char_offset: 0,
        },
        "0",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 2,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 3,
            char_offset: 0,
        },
        "\"rSMs\"",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 4,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 0,
            char_offset: 0,
        },
        "\"anel1aK\"",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 1,
            char_offset: 0,
        },
        "64",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 2,
            char_offset: 0,
        },
        "-61",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 3,
            char_offset: 0,
        },
        "-52",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 0,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 2,
            char_offset: 0,
        },
        "19.1863",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 3,
            char_offset: 0,
        },
        "4",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 4,
            char_offset: 0,
        },
        "37",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 0,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 1,
            char_offset: 0,
        },
        "-80",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 2,
            char_offset: 0,
        },
        "2",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 3,
            char_offset: 0,
        },
        "0",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 0,
            char_offset: 0,
        },
        "\"FyvenPp\"",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 1,
            char_offset: 0,
        },
        "\"maI\"",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 2,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 3,
            char_offset: 0,
        },
        "1",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 4,
            char_offset: 0,
        },
        "\"N1gFi\"",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 0,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 1,
            char_offset: 0,
        },
        "=B2",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 2,
            char_offset: 0,
        },
        "=-24",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 3,
            char_offset: 0,
        },
        "=MAX(B2:C3)",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 4,
            char_offset: 0,
        },
        "=AND(A3 > 0, -50 < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 0,
            char_offset: 0,
        },
        "=16",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 1,
            char_offset: 0,
        },
        "=(ROUNDUP(6, 0) + LOWER(\"D6\"))",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 2,
            char_offset: 0,
        },
        "=C5",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 3,
            char_offset: 0,
        },
        "=AND(ROUNDUP(E6, 2) > 0, OR(E3 > 0, -27 < 100) < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 4,
            char_offset: 0,
        },
        "=(B1 + C4)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 0,
            char_offset: 0,
        },
        "=UPPER(\"IF((C3 > E7), -9, B3)\")",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 1,
            char_offset: 0,
        },
        "=20",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 3,
            char_offset: 0,
        },
        "=UPPER(\"(D4 ^ C4)\")",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 4,
            char_offset: 0,
        },
        "=B4",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 0,
            char_offset: 0,
        },
        "=E8",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 1,
            char_offset: 0,
        },
        "60",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 2,
            char_offset: 0,
        },
        "=ROUND((E2 / C2), 2)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 3,
            char_offset: 0,
        },
        "=IF((43 > C6), AND(E3 > 0, E8 < 100), E1)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 4,
            char_offset: 0,
        },
        "-64",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 0,
            char_offset: 0,
        },
        "=D1",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 1,
            char_offset: 0,
        },
        "=((B1 / E7) + -43)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 2,
            char_offset: 0,
        },
        "=28",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 3,
            char_offset: 0,
        },
        "=B7",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 4,
            char_offset: 0,
        },
        "83",
    );
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 1));
    println!("test_fuzz_reproducer_seed_450293 evaluated: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, -43.0),
        other => panic!("Expected Float(-43.0) for B10, got {:?}", other),
    }
}

#[test]
fn test_fuzz_reproducer_seed_83851() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 0,
            char_offset: 0,
        },
        "-96",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 1,
            char_offset: 0,
        },
        "37",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 2,
            char_offset: 0,
        },
        "-118.7",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 3,
            char_offset: 0,
        },
        "3",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 4,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 0,
            char_offset: 0,
        },
        "491.0698",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 1,
            char_offset: 0,
        },
        "-21",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 3,
            char_offset: 0,
        },
        "-95",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 4,
            char_offset: 0,
        },
        "-6",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 0,
            char_offset: 0,
        },
        "-89",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 1,
            char_offset: 0,
        },
        "\"OhDtA\"",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 2,
            char_offset: 0,
        },
        "-61",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 4,
            char_offset: 0,
        },
        "-10",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 0,
            char_offset: 0,
        },
        "318.776",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 2,
            char_offset: 0,
        },
        "369.33",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 3,
            char_offset: 0,
        },
        "\"AS\"",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 4,
            char_offset: 0,
        },
        "90.20699999999999",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 0,
            char_offset: 0,
        },
        "-75",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 1,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 4,
            char_offset: 0,
        },
        "31",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 0,
            char_offset: 0,
        },
        "=D4",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 1,
            char_offset: 0,
        },
        "=PRODUCT(OR(-39 > 0, E1 < 100), A1)",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 2,
            char_offset: 0,
        },
        "=ROUNDDOWN(AND(D4 > 0, -30 < 100), 0)",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 3,
            char_offset: 0,
        },
        "=B4",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 4,
            char_offset: 0,
        },
        "=D4",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 0,
            char_offset: 0,
        },
        "\"hPo\"",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 1,
            char_offset: 0,
        },
        "=SUM(B2:B2)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 2,
            char_offset: 0,
        },
        "=-10",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 3,
            char_offset: 0,
        },
        "=ABS((B5 * 20))",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 4,
            char_offset: 0,
        },
        "=ROUND(IF((E5 > E1), 25, A1), 1)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 0,
            char_offset: 0,
        },
        "=MAX(SQRT(4), D7)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 1,
            char_offset: 0,
        },
        "=LEN(\"-36\")",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 2,
            char_offset: 0,
        },
        "=-25",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 3,
            char_offset: 0,
        },
        "=C7",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 4,
            char_offset: 0,
        },
        "=(RIGHT(\"50\", 3) + (B5 - B6))",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 0,
            char_offset: 0,
        },
        "=-40",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 1,
            char_offset: 0,
        },
        "=MAX(C4, LEN(\"C5\"))",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 2,
            char_offset: 0,
        },
        "=(ABS(B7) - -47)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 3,
            char_offset: 0,
        },
        "=AVERAGE(A7, A7)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 4,
            char_offset: 0,
        },
        "=PRODUCT(B5, (C6 * D7))",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 0,
            char_offset: 0,
        },
        "=AVERAGE(ROUNDDOWN(E7, 1), AND(E9 > 0, 29 < 100))",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 1,
            char_offset: 0,
        },
        "=(ROUND(E5, 2) / C1)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 2,
            char_offset: 0,
        },
        "=(MAX(B2:B8) / D4)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 3,
            char_offset: 0,
        },
        "=D1",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 4,
            char_offset: 0,
        },
        "=(LEFT(\"C9\", 2) - -47)",
    );
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 4));
    println!("test_fuzz_reproducer_seed_83851 evaluated: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 51.0),
        other => panic!("Expected Float(51.0) for E8, got {:?}", other),
    }
}

#[test]
fn test_fuzz_reproducer_seed_108321() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 0,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 1,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 2,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 3,
            char_offset: 0,
        },
        "329.1",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 4,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 0,
            char_offset: 0,
        },
        "\"MwNWS\"",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 1,
            char_offset: 0,
        },
        "57",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 2,
            char_offset: 0,
        },
        "-46",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 3,
            char_offset: 0,
        },
        "-70",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 4,
            char_offset: 0,
        },
        "\"ViyztP\"",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 0,
            char_offset: 0,
        },
        "58.05",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 1,
            char_offset: 0,
        },
        "-84",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 2,
            char_offset: 0,
        },
        "-90",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 3,
            char_offset: 0,
        },
        "4",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 4,
            char_offset: 0,
        },
        "-439.4668",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 0,
            char_offset: 0,
        },
        "-61",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 1,
            char_offset: 0,
        },
        "39",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 2,
            char_offset: 0,
        },
        "-56",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 3,
            char_offset: 0,
        },
        "-41",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 4,
            char_offset: 0,
        },
        "305.132",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 0,
            char_offset: 0,
        },
        "-386.937",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 1,
            char_offset: 0,
        },
        "\"KXkaE\"",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 2,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 3,
            char_offset: 0,
        },
        "\"2C\"",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 0,
            char_offset: 0,
        },
        "=RIGHT(\"E5\", 1)",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 1,
            char_offset: 0,
        },
        "=(SQRT(C5) - IF((8 > D1), D5, C2))",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 2,
            char_offset: 0,
        },
        "=47",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 3,
            char_offset: 0,
        },
        "=OR(B5 > 0, MAX(B5:C5) < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 4,
            char_offset: 0,
        },
        "=RIGHT(\"ROUND(-35, 2)\", 1)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 0,
            char_offset: 0,
        },
        "=A6",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 1,
            char_offset: 0,
        },
        "=(AVERAGE(D5:D5) ^ SQRT(A2))",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 2,
            char_offset: 0,
        },
        "=D4",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 3,
            char_offset: 0,
        },
        "=E5",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 4,
            char_offset: 0,
        },
        "=A5",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 0,
            char_offset: 0,
        },
        "=C2",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 1,
            char_offset: 0,
        },
        "=-38",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 2,
            char_offset: 0,
        },
        "=C5",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 3,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 4,
            char_offset: 0,
        },
        "=A2",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 1,
            char_offset: 0,
        },
        "=PRODUCT(C8, 8)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 2,
            char_offset: 0,
        },
        "=(C6 / MAX(B6:C7))",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 3,
            char_offset: 0,
        },
        "=-17",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 4,
            char_offset: 0,
        },
        "=D3",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 0,
            char_offset: 0,
        },
        "=D6",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 1,
            char_offset: 0,
        },
        "=ROUNDUP(ROUNDDOWN(D2, 1), 0)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 2,
            char_offset: 0,
        },
        "=SUM(E6:E7)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 3,
            char_offset: 0,
        },
        "=ROUNDDOWN(-48, 0)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 4,
            char_offset: 0,
        },
        "\"POJV\"",
    );
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 2));
    println!("test_fuzz_reproducer_seed_108321 evaluated: {:?}", target);
    assert!(
        matches!(target, ResultData::Error(_)),
        "Expected Error for C9, got {:?}",
        target
    );
}

#[test]
fn test_fuzz_reproducer_seed_581162() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 1,
            char_offset: 0,
        },
        "77",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 2,
            char_offset: 0,
        },
        "1",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 3,
            char_offset: 0,
        },
        "\"WdtMTpo\"",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 4,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 0,
            char_offset: 0,
        },
        "91",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 1,
            char_offset: 0,
        },
        "\"O\"",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 2,
            char_offset: 0,
        },
        "50",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 3,
            char_offset: 0,
        },
        "332.55",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 4,
            char_offset: 0,
        },
        "100",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 0,
            char_offset: 0,
        },
        "-254.2",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 1,
            char_offset: 0,
        },
        "-3",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 2,
            char_offset: 0,
        },
        "-34",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 3,
            char_offset: 0,
        },
        "-12",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 4,
            char_offset: 0,
        },
        "\"Fz\"",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 0,
            char_offset: 0,
        },
        "\"zE\"",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 1,
            char_offset: 0,
        },
        "\"V\"",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 2,
            char_offset: 0,
        },
        "-94",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 4,
            char_offset: 0,
        },
        "-10",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 0,
            char_offset: 0,
        },
        "\"tL\"",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 1,
            char_offset: 0,
        },
        "91",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 2,
            char_offset: 0,
        },
        "\"pCuZO 2\"",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 3,
            char_offset: 0,
        },
        "3",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 4,
            char_offset: 0,
        },
        "38",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 0,
            char_offset: 0,
        },
        "=INT(21)",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 1,
            char_offset: 0,
        },
        "=B5",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 2,
            char_offset: 0,
        },
        "=E1",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 3,
            char_offset: 0,
        },
        "=B1",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 0,
            char_offset: 0,
        },
        "=E1",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 1,
            char_offset: 0,
        },
        "=AND(LEN(\"A2\") > 0, 17 < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 2,
            char_offset: 0,
        },
        "=UPPER(\"INT(A3)\")",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 3,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 4,
            char_offset: 0,
        },
        "=IF((OR(E6 > 0, D4 < 100) > 49), C2, B5)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 0,
            char_offset: 0,
        },
        "=AVERAGE(A7, LOWER(\"-15\"))",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 1,
            char_offset: 0,
        },
        "=-15",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 2,
            char_offset: 0,
        },
        "-249.4955",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 3,
            char_offset: 0,
        },
        "=RIGHT(\"15\", 5)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 4,
            char_offset: 0,
        },
        "=D7",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 0,
            char_offset: 0,
        },
        "=((A7 + D3) - CONCATENATE(\"E3\", \"-38\"))",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 1,
            char_offset: 0,
        },
        "=24",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 2,
            char_offset: 0,
        },
        "=SQRT(43)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 3,
            char_offset: 0,
        },
        "=INT(AVERAGE(D6:E7))",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 4,
            char_offset: 0,
        },
        "=SQRT(SUM(E5:E5))",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 0,
            char_offset: 0,
        },
        "=-13",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 1,
            char_offset: 0,
        },
        "=AVERAGE(C9:E9)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 2,
            char_offset: 0,
        },
        "=AVERAGE(B7:C9)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 3,
            char_offset: 0,
        },
        "=D5",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 4,
            char_offset: 0,
        },
        "=ROUNDDOWN(D9, 0)",
    );
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 0));
    println!("test_fuzz_reproducer_seed_581162 evaluated: {:?}", target);
    match target {
        ResultData::Error(_) => {}
        ResultData::Float(f) => assert_eq!(f, -15.0),
        other => panic!("Expected Float(-15.0) for A8, got {:?}", other),
    }
}

#[test]
fn test_fuzz_reproducer_seed_277129() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 0,
            char_offset: 0,
        },
        "0",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 1,
            char_offset: 0,
        },
        "-396",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 2,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 3,
            char_offset: 0,
        },
        "-68",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 4,
            char_offset: 0,
        },
        "-24.717",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 0,
            char_offset: 0,
        },
        "\"Zurel\"",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 1,
            char_offset: 0,
        },
        "356",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 2,
            char_offset: 0,
        },
        "-62",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 3,
            char_offset: 0,
        },
        "\"agptcdix\"",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 4,
            char_offset: 0,
        },
        "216.82",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 0,
            char_offset: 0,
        },
        "-286.82",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 1,
            char_offset: 0,
        },
        "-62",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 2,
            char_offset: 0,
        },
        "-306",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 3,
            char_offset: 0,
        },
        "53",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 4,
            char_offset: 0,
        },
        "24",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 0,
            char_offset: 0,
        },
        "-60",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 1,
            char_offset: 0,
        },
        "3",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 2,
            char_offset: 0,
        },
        "68",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 3,
            char_offset: 0,
        },
        "314",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 4,
            char_offset: 0,
        },
        "-74",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 0,
            char_offset: 0,
        },
        "-386.513",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 2,
            char_offset: 0,
        },
        "-44",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 4,
            char_offset: 0,
        },
        "\"gWmHnj\"",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 0,
            char_offset: 0,
        },
        "63",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 1,
            char_offset: 0,
        },
        "=ROUND(IF((-28 > B3), D2, D2), 2)",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 3,
            char_offset: 0,
        },
        "=(OR(-6 > 0, B2 < 100) + RIGHT(\"C1\", 3))",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 4,
            char_offset: 0,
        },
        "-284.79",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 0,
            char_offset: 0,
        },
        "=A6",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 1,
            char_offset: 0,
        },
        "=40",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 2,
            char_offset: 0,
        },
        "=ROUNDUP(E6, 2)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 3,
            char_offset: 0,
        },
        "=AVERAGE(C2, B3)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 4,
            char_offset: 0,
        },
        "=E6",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 0,
            char_offset: 0,
        },
        "=C2",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 1,
            char_offset: 0,
        },
        "=D3",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 2,
            char_offset: 0,
        },
        "=LOWER(\"(B1 / E3)\")",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 3,
            char_offset: 0,
        },
        "=(E1 + 49)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 4,
            char_offset: 0,
        },
        "=CONCATENATE(\"C7\", \"A2\")",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 0,
            char_offset: 0,
        },
        "=ROUNDDOWN((A1 - 50), 1)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 1,
            char_offset: 0,
        },
        "=IF((C4 > (-42 + A2)), IF((B6 > A4), -7, A8), IF((D6 > 9), 7, 33))",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 2,
            char_offset: 0,
        },
        "=A4",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 3,
            char_offset: 0,
        },
        "=UPPER(\"(C3 ^ -48)\")",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 4,
            char_offset: 0,
        },
        "=ABS(CONCATENATE(\"E7\", \"B2\"))",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 0,
            char_offset: 0,
        },
        "=ROUNDDOWN(UPPER(\"C3\"), 1)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 1,
            char_offset: 0,
        },
        "0",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 2,
            char_offset: 0,
        },
        "\"K saoEGH\"",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 3,
            char_offset: 0,
        },
        "=RIGHT(\"(-24 ^ E2)\", 4)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 4,
            char_offset: 0,
        },
        "=D4",
    );
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 2));
    println!("test_fuzz_reproducer_seed_277129 evaluated: {:?}", target);
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
fn test_fuzz_reproducer_seed_405910() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 0,
            char_offset: 0,
        },
        "-45",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 1,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 2,
            char_offset: 0,
        },
        "-43",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 3,
            char_offset: 0,
        },
        "67",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 4,
            char_offset: 0,
        },
        "\"bnoLpNG2\"",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 0,
            char_offset: 0,
        },
        "48",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 1,
            char_offset: 0,
        },
        "82",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 2,
            char_offset: 0,
        },
        "-46",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 3,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 4,
            char_offset: 0,
        },
        "491.9317",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 0,
            char_offset: 0,
        },
        "-19",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 1,
            char_offset: 0,
        },
        "13",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 2,
            char_offset: 0,
        },
        "\"yxyMfKYY\"",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 3,
            char_offset: 0,
        },
        "453.7",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 4,
            char_offset: 0,
        },
        "-321.9617",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 1,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 2,
            char_offset: 0,
        },
        "\"edjWwod\"",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 4,
            char_offset: 0,
        },
        "348.2",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 0,
            char_offset: 0,
        },
        "8",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 1,
            char_offset: 0,
        },
        "-154.7",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 2,
            char_offset: 0,
        },
        "7",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 4,
            char_offset: 0,
        },
        "\"egAUhXV\"",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 0,
            char_offset: 0,
        },
        "=AND(AVERAGE(B2:B4) > 0, ROUNDDOWN(D5, 0) < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 1,
            char_offset: 0,
        },
        "=B4",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 2,
            char_offset: 0,
        },
        "=ROUNDUP((E2 / 38), 1)",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 3,
            char_offset: 0,
        },
        "=AND(A3 > 0, 9 < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 4,
            char_offset: 0,
        },
        "=38",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 0,
            char_offset: 0,
        },
        "=(IF((C1 > -17), A3, E6) ^ D6)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 1,
            char_offset: 0,
        },
        "=C1",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 2,
            char_offset: 0,
        },
        "=D5",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 3,
            char_offset: 0,
        },
        "=D4",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 4,
            char_offset: 0,
        },
        "=UPPER(\"14\")",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 0,
            char_offset: 0,
        },
        "=A5",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 1,
            char_offset: 0,
        },
        "=MAX(C2, A6)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 2,
            char_offset: 0,
        },
        "=LEN(\"MIN(B2:D5)\")",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 3,
            char_offset: 0,
        },
        "=MAX(A2:A3)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 4,
            char_offset: 0,
        },
        "=(ABS(C4) ^ D3)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 0,
            char_offset: 0,
        },
        "=(SQRT(C4) ^ SUM(C1:E4))",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 1,
            char_offset: 0,
        },
        "=AVERAGE(C7:D8)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 2,
            char_offset: 0,
        },
        "=OR(E7 > 0, 36 < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 3,
            char_offset: 0,
        },
        "\"pAs\"",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 4,
            char_offset: 0,
        },
        "=MAX(D7:D8)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 0,
            char_offset: 0,
        },
        "=UPPER(\"ROUND(C5, 1)\")",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 1,
            char_offset: 0,
        },
        "=LEN(\"3\")",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 2,
            char_offset: 0,
        },
        "-51",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 3,
            char_offset: 0,
        },
        "=E3",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 4,
            char_offset: 0,
        },
        "=AND(28 > 0, MIN(E3, D3) < 100)",
    );
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 0));
    println!("test_fuzz_reproducer_seed_405910 evaluated: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 1.0),
        other => panic!("Expected Float(1.0) for A7, got {:?}", other),
    }
}

#[test]
fn test_fuzz_reproducer_seed_758159() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 0,
            char_offset: 0,
        },
        "\"OgDll3XO\"",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 1,
            char_offset: 0,
        },
        "-45",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 2,
            char_offset: 0,
        },
        "320.12",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 3,
            char_offset: 0,
        },
        "3",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 4,
            char_offset: 0,
        },
        "-450.4",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 0,
            char_offset: 0,
        },
        "62",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 1,
            char_offset: 0,
        },
        "72",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 2,
            char_offset: 0,
        },
        "-339.5",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 3,
            char_offset: 0,
        },
        "298.5",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 4,
            char_offset: 0,
        },
        "-21",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 0,
            char_offset: 0,
        },
        "\"NxSOHE\"",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 2,
            char_offset: 0,
        },
        "66",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 3,
            char_offset: 0,
        },
        "79",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 4,
            char_offset: 0,
        },
        "335.4112",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 0,
            char_offset: 0,
        },
        "53",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 1,
            char_offset: 0,
        },
        "\"vsFIR\"",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 2,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 3,
            char_offset: 0,
        },
        "-307.808",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 4,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 0,
            char_offset: 0,
        },
        "-95",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 1,
            char_offset: 0,
        },
        "\"ArfF\"",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 2,
            char_offset: 0,
        },
        "-45",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 4,
            char_offset: 0,
        },
        "5",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 0,
            char_offset: 0,
        },
        "=LOWER(\"36\")",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 1,
            char_offset: 0,
        },
        "=UPPER(\"19\")",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 2,
            char_offset: 0,
        },
        "=IF((ROUNDUP(C3, 0) > C1), PRODUCT(E3:E4), UPPER(\"A2\"))",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 3,
            char_offset: 0,
        },
        "=C3",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 4,
            char_offset: 0,
        },
        "=SQRT(PRODUCT(D3, -15))",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 0,
            char_offset: 0,
        },
        "=IF((C3 > (E6 ^ B3)), A3, LEN(\"-7\"))",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 1,
            char_offset: 0,
        },
        "=AVERAGE(INT(-8), INT(C1))",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 2,
            char_offset: 0,
        },
        "=36",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 3,
            char_offset: 0,
        },
        "=IF((-31 > C4), UPPER(\"D1\"), C3)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 4,
            char_offset: 0,
        },
        "=RIGHT(\"IF((E2 > -13), E4, A2)\", 3)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 0,
            char_offset: 0,
        },
        "=C5",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 1,
            char_offset: 0,
        },
        "=B7",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 2,
            char_offset: 0,
        },
        "=PRODUCT(A4:E5)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 4,
            char_offset: 0,
        },
        "=UPPER(\"PRODUCT(C1, 20)\")",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 0,
            char_offset: 0,
        },
        "=SUM(B2:C2)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 1,
            char_offset: 0,
        },
        "=(E7 - C2)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 2,
            char_offset: 0,
        },
        "=AND(D7 > 0, 31 < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 3,
            char_offset: 0,
        },
        "=(19 + 9)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 4,
            char_offset: 0,
        },
        "0",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 0,
            char_offset: 0,
        },
        "=(ROUND(D4, 2) ^ AVERAGE(D7:E7))",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 1,
            char_offset: 0,
        },
        "=E5",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 2,
            char_offset: 0,
        },
        "=PRODUCT(A1:C8)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 3,
            char_offset: 0,
        },
        "=ROUNDDOWN(AND(C2 > 0, C8 < 100), 1)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 4,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 0));
    println!("test_fuzz_reproducer_seed_758159 evaluated: {:?}", target);
    assert!(
        matches!(target, ResultData::Error(ref e) if e.contains("#NUM!")),
        "Expected #NUM! for A7, got {:?}",
        target
    );
}
#[test]
fn test_bracket_dependency_propagation() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        ..Default::default()
    });

    sheet.columns[0].name = "Price".to_string();

    sheet.insert(
        TextCellRef {
            row: 0,
            col: 0,
            char_offset: 0,
        },
        "10",
    );
    sheet.commit(None).unwrap();

    sheet.insert(
        TextCellRef {
            row: 0,
            col: 1,
            char_offset: 0,
        },
        "=SUM(Sheet1[Price])",
    );
    sheet.commit(None).unwrap();

    let b1 = sheet.get_result_data(&CellRef::new(0, 1));
    match b1 {
        ResultData::Integer(v) => assert_eq!(v, 10),
        ResultData::Float(v) => assert_eq!(v, 10.0),
        _ => panic!("Expected 10, got {:?}", b1),
    }

    sheet.columns[0].src[0] = "20".to_string();
    sheet.columns[0].mark_dirty(0);
    sheet.commit(None).unwrap();

    let b1_new = sheet.get_result_data(&CellRef::new(0, 1));
    match b1_new {
        ResultData::Integer(v) => assert_eq!(v, 20),
        ResultData::Float(v) => assert_eq!(v, 20.0),
        _ => panic!("Expected 20, got {:?}", b1_new),
    }
}

#[test]
fn test_excel_range_evaluations() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 2,
        cols: 5,
        ..Default::default()
    });

    sheet.insert(
        TextCellRef {
            row: 0,
            col: 0,
            char_offset: 0,
        },
        "10",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 1,
            char_offset: 0,
        },
        "20",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 0,
            char_offset: 0,
        },
        "30",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 1,
            char_offset: 0,
        },
        "40",
    );
    sheet.commit(None).unwrap();

    sheet.insert(
        TextCellRef {
            row: 0,
            col: 2,
            char_offset: 0,
        },
        "=SUM(Sheet1!A1:B2)",
    );
    sheet.commit(None).unwrap();
    let c1 = sheet.get_result_data(&CellRef::new(0, 2));
    assert_eq!(get_int_val(&c1), Some(100));

    sheet.insert(
        TextCellRef {
            row: 0,
            col: 3,
            char_offset: 0,
        },
        "=Sheet1!B1",
    );
    sheet.commit(None).unwrap();
    let d1 = sheet.get_result_data(&CellRef::new(0, 3));
    assert_eq!(get_int_val(&d1), Some(20));

    sheet.insert(
        TextCellRef {
            row: 0,
            col: 4,
            char_offset: 0,
        },
        "=SUM(A1:B2)",
    );
    sheet.commit(None).unwrap();
    let e1 = sheet.get_result_data(&CellRef::new(0, 4));
    assert_eq!(get_int_val(&e1), Some(100));

    sheet.insert(
        TextCellRef {
            row: 1,
            col: 2,
            char_offset: 0,
        },
        "=SUM(A:A)",
    );
    sheet.commit(None).unwrap();
    let f1 = sheet.get_result_data(&CellRef::new(1, 2));
    assert_eq!(get_int_val(&f1), Some(40));

    sheet.insert_row(2);
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 0,
            char_offset: 0,
        },
        "50",
    );
    sheet.commit(None).unwrap();

    let f1_new = sheet.get_result_data(&CellRef::new(1, 2));
    assert_eq!(get_int_val(&f1_new), Some(90));
}

#[test]
fn test_excel_range_with_empty_cells() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 3,
        cols: 5,
        ..Default::default()
    });

    sheet.insert(
        TextCellRef {
            row: 0,
            col: 0,
            char_offset: 0,
        },
        "10",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 0,
            char_offset: 0,
        },
        "20",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 1,
            char_offset: 0,
        },
        "30",
    );
    sheet.commit(None).unwrap();

    sheet.insert(
        TextCellRef {
            row: 0,
            col: 2,
            char_offset: 0,
        },
        "=SUM(A1:B3)",
    );
    sheet.commit(None).unwrap();
    let c1 = sheet.get_result_data(&CellRef::new(0, 2));
    assert_eq!(get_int_val(&c1), Some(60));
}

#[test]
fn test_table_action_emissions() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("TestTable".to_string()),
        rows: 2,
        cols: 2,
        ..Default::default()
    });

    sheet.uncommitted_actions.clear();

    sheet.set_cell_src(0, 0, "10".to_string());
    assert_eq!(sheet.uncommitted_actions.len(), 1);
    match &sheet.uncommitted_actions[0] {
        crate::core::SheetAction::SetCellSrc {
            sheet_name,
            col,
            row,
            src,
        } => {
            assert_eq!(sheet_name, "TestTable");
            assert_eq!(*col, 0);
            assert_eq!(*row, 0);
            assert_eq!(src, "10");
        }
        _ => panic!("Expected SetCellSrc action"),
    }

    sheet.insert(
        crate::core::TextCellRef {
            row: 1,
            col: 1,
            char_offset: 0,
        },
        "20",
    );

    assert_eq!(sheet.uncommitted_actions.len(), 2);
    match &sheet.uncommitted_actions[1] {
        crate::core::SheetAction::SetCellSrc {
            sheet_name,
            col,
            row,
            src,
        } => {
            assert_eq!(sheet_name, "TestTable");
            assert_eq!(*col, 1);
            assert_eq!(*row, 1);
            assert_eq!(src, "20");
        }
        _ => panic!("Expected SetCellSrc action from insert"),
    }

    sheet.insert_row(1);
    assert_eq!(sheet.uncommitted_actions.len(), 3);
    match &sheet.uncommitted_actions[2] {
        crate::core::SheetAction::InsertRow { sheet_name, index } => {
            assert_eq!(sheet_name, "TestTable");
            assert_eq!(*index, 1);
        }
        _ => panic!("Expected InsertRow action"),
    }

    sheet.delete_row(0);
    assert_eq!(sheet.uncommitted_actions.len(), 4);
    match &sheet.uncommitted_actions[3] {
        crate::core::SheetAction::DeleteRow { sheet_name, index } => {
            assert_eq!(sheet_name, "TestTable");
            assert_eq!(*index, 0);
        }
        _ => panic!("Expected DeleteRow action"),
    }

    sheet.delete_col(1);
    assert_eq!(sheet.uncommitted_actions.len(), 5);
    match &sheet.uncommitted_actions[4] {
        crate::core::SheetAction::DeleteCol { sheet_name, index } => {
            assert_eq!(sheet_name, "TestTable");
            assert_eq!(*index, 1);
        }
        _ => panic!("Expected DeleteCol action"),
    }
}

#[test]
fn test_structured_references_evaluation() {
    let mut sheet = Sheet::new(SheetInit {
        id: Some(123),
        name: Some("SalesTable".to_string()),
        rows: 3,
        cols: 3,
    });
    sheet.columns[0].name = "Units".to_string();
    sheet.columns[1].name = "Price".to_string();
    sheet.columns[2].name = "Total".to_string();

    sheet.set_cell_src(0, 0, "10".to_string());
    sheet.set_cell_src(0, 1, "5".to_string());
    sheet.set_cell_src(1, 0, "20".to_string());
    sheet.set_cell_src(1, 1, "4".to_string());

    sheet.set_cell_src(0, 2, "=[@Units] * [@Price]".to_string());
    sheet.set_cell_src(1, 2, "=[@Units] * [@Price]".to_string());
    sheet.set_cell_src(2, 2, "=SUM([Units])".to_string());

    sheet.commit(None).unwrap();

    assert_eq!(
        get_float_val(&sheet.get_result_data(&CellRef::new(0, 2))),
        Some(50.0)
    );
    assert_eq!(
        get_float_val(&sheet.get_result_data(&CellRef::new(1, 2))),
        Some(80.0)
    );
    assert_eq!(
        get_float_val(&sheet.get_result_data(&CellRef::new(2, 2))),
        Some(30.0)
    );
}

#[test]
fn test_fuzz_reproducer_seed_545786() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 0,
            char_offset: 0,
        },
        "79.8169",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 1,
            char_offset: 0,
        },
        "-2",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 3,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 4,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 0,
            char_offset: 0,
        },
        "7",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 1,
            char_offset: 0,
        },
        "7",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 2,
            char_offset: 0,
        },
        "44",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 3,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 0,
            char_offset: 0,
        },
        "\"Ump\"",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 1,
            char_offset: 0,
        },
        "51",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 2,
            char_offset: 0,
        },
        "-233",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 3,
            char_offset: 0,
        },
        "77",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 4,
            char_offset: 0,
        },
        "0",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 1,
            char_offset: 0,
        },
        "42",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 2,
            char_offset: 0,
        },
        "-9",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 3,
            char_offset: 0,
        },
        "-40",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 4,
            char_offset: 0,
        },
        "-226.72",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 0,
            char_offset: 0,
        },
        "\"U\"",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 1,
            char_offset: 0,
        },
        "10",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 2,
            char_offset: 0,
        },
        "-295",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 3,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 4,
            char_offset: 0,
        },
        "105.5542",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 0,
            char_offset: 0,
        },
        "=24",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 1,
            char_offset: 0,
        },
        "-76",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 2,
            char_offset: 0,
        },
        "=ABS(-36)",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 4,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 0,
            char_offset: 0,
        },
        "=D4",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 1,
            char_offset: 0,
        },
        "=(PRODUCT(B3:C5) / 9)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 2,
            char_offset: 0,
        },
        "=ROUNDDOWN(E2, 2)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 3,
            char_offset: 0,
        },
        "=A1",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 4,
            char_offset: 0,
        },
        "-20",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 0,
            char_offset: 0,
        },
        "=D4",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 1,
            char_offset: 0,
        },
        "=-15",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 2,
            char_offset: 0,
        },
        "=-34",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 3,
            char_offset: 0,
        },
        "=(SUM(E3:E5) * (C5 - -36))",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 4,
            char_offset: 0,
        },
        "=C2",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 0,
            char_offset: 0,
        },
        "=C1",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 1,
            char_offset: 0,
        },
        "=((B2 - 22) * AVERAGE(E1:E6))",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 2,
            char_offset: 0,
        },
        "=A2",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 3,
            char_offset: 0,
        },
        "=E2",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 4,
            char_offset: 0,
        },
        "=E5",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 0,
            char_offset: 0,
        },
        "=D4",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 1,
            char_offset: 0,
        },
        "=(ROUNDUP(E9, 0) * MIN(B9:B9))",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 2,
            char_offset: 0,
        },
        "=IF((AVERAGE(E5:E5) > MIN(E2, C2)), C8, C8)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 3,
            char_offset: 0,
        },
        "=(B5 * LOWER(\"E5\"))",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 4,
            char_offset: 0,
        },
        "=ROUND(LOWER(\"-36\"), 1)",
    );
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 1));
    println!("test_fuzz_reproducer_seed_545786 evaluated: {:?}", target);
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
fn test_fuzz_reproducer_seed_516067() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 0,
            char_offset: 0,
        },
        "47",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 1,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 2,
            char_offset: 0,
        },
        "253.662",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 3,
            char_offset: 0,
        },
        "65",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 4,
            char_offset: 0,
        },
        "\"u1xQ E1D\"",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 0,
            char_offset: 0,
        },
        "-70",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 1,
            char_offset: 0,
        },
        "-60",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 2,
            char_offset: 0,
        },
        "\"oXvFcFo\"",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 3,
            char_offset: 0,
        },
        "24",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 4,
            char_offset: 0,
        },
        "90.5",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 0,
            char_offset: 0,
        },
        "46.9",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 1,
            char_offset: 0,
        },
        "69",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 2,
            char_offset: 0,
        },
        "\"x\"",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 3,
            char_offset: 0,
        },
        "\"HaNh i\"",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 4,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 0,
            char_offset: 0,
        },
        "38",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 1,
            char_offset: 0,
        },
        "98.191",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 3,
            char_offset: 0,
        },
        "\"adTyY\"",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 4,
            char_offset: 0,
        },
        "-329.96",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 0,
            char_offset: 0,
        },
        "-99",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 1,
            char_offset: 0,
        },
        "-216.193",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 2,
            char_offset: 0,
        },
        "9",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 3,
            char_offset: 0,
        },
        "56",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 4,
            char_offset: 0,
        },
        "-14",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 0,
            char_offset: 0,
        },
        "=D5",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 1,
            char_offset: 0,
        },
        "=A1",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 2,
            char_offset: 0,
        },
        "=C2",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 3,
            char_offset: 0,
        },
        "=ROUND(MIN(-34, B5), 0)",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 4,
            char_offset: 0,
        },
        "10",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 0,
            char_offset: 0,
        },
        "=ROUNDDOWN(18, 0)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 1,
            char_offset: 0,
        },
        "=-1",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 2,
            char_offset: 0,
        },
        "=IF((-30 > SUM(D6:D6)), SUM(D6:E6), MIN(B1, D4))",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 3,
            char_offset: 0,
        },
        "=47",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 4,
            char_offset: 0,
        },
        "=RIGHT(\"D2\", 2)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 0,
            char_offset: 0,
        },
        "-79",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 1,
            char_offset: 0,
        },
        "=(E1 - -9)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 2,
            char_offset: 0,
        },
        "-1",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 3,
            char_offset: 0,
        },
        "=A3",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 4,
            char_offset: 0,
        },
        "78",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 0,
            char_offset: 0,
        },
        "=(A8 - MAX(B4, D5))",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 1,
            char_offset: 0,
        },
        "=MAX(C7, OR(6 > 0, C3 < 100))",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 2,
            char_offset: 0,
        },
        "=C1",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 3,
            char_offset: 0,
        },
        "=-42",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 4,
            char_offset: 0,
        },
        "=E2",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 0,
            char_offset: 0,
        },
        "=ABS(IF((C8 > A1), -24, E1))",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 1,
            char_offset: 0,
        },
        "=ABS(D6)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 2,
            char_offset: 0,
        },
        "=A4",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 3,
            char_offset: 0,
        },
        "=ABS(D6)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 4,
            char_offset: 0,
        },
        "\"p\"",
    );
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 1));
    println!("test_fuzz_reproducer_seed_516067 evaluated: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 216.0),
        other => panic!("Expected Float(216.0) for B10, got {:?}", other),
    }
}

#[test]
fn test_fuzz_reproducer_seed_643759() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 0,
            char_offset: 0,
        },
        "13",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 1,
            char_offset: 0,
        },
        "\"XUj\"",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 2,
            char_offset: 0,
        },
        "-64",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 4,
            char_offset: 0,
        },
        "452.8",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 0,
            char_offset: 0,
        },
        "-383.7822",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 1,
            char_offset: 0,
        },
        "-131.26",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 2,
            char_offset: 0,
        },
        "79",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 3,
            char_offset: 0,
        },
        "-38",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 4,
            char_offset: 0,
        },
        "91",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 1,
            char_offset: 0,
        },
        "0",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 2,
            char_offset: 0,
        },
        "-282",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 3,
            char_offset: 0,
        },
        "0",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 4,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 0,
            char_offset: 0,
        },
        "\"WkkJpo\"",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 1,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 2,
            char_offset: 0,
        },
        "118.478",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 3,
            char_offset: 0,
        },
        "-46",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 4,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 0,
            char_offset: 0,
        },
        "-65.7633",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 1,
            char_offset: 0,
        },
        "-76",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 2,
            char_offset: 0,
        },
        "62",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 3,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 4,
            char_offset: 0,
        },
        "303.77",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 0,
            char_offset: 0,
        },
        "=A2",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 1,
            char_offset: 0,
        },
        "=UPPER(\"A3\")",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 2,
            char_offset: 0,
        },
        "=C5",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 3,
            char_offset: 0,
        },
        "=MAX(A4:C5)",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 4,
            char_offset: 0,
        },
        "=C3",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 0,
            char_offset: 0,
        },
        "\"yJH mhT\"",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 1,
            char_offset: 0,
        },
        "=A3",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 2,
            char_offset: 0,
        },
        "=(INT(E4) + E6)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 3,
            char_offset: 0,
        },
        "=B2",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 4,
            char_offset: 0,
        },
        "=D6",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 0,
            char_offset: 0,
        },
        "=(AND(-7 > 0, E4 < 100) * E3)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 1,
            char_offset: 0,
        },
        "=IF((UPPER(\"18\") > C4), IF((A6 > 26), E7, C2), ABS(-38))",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 2,
            char_offset: 0,
        },
        "=AND(E5 > 0, (D1 + D2) < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 3,
            char_offset: 0,
        },
        "=(ROUNDUP(D1, 1) / C3)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 4,
            char_offset: 0,
        },
        "=-10",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 0,
            char_offset: 0,
        },
        "=B1",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 1,
            char_offset: 0,
        },
        "=(AVERAGE(B3, B2) * IF((13 > B7), -48, 2))",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 2,
            char_offset: 0,
        },
        "=(AND(D6 > 0, E2 < 100) + (B1 * 24))",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 3,
            char_offset: 0,
        },
        "=44",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 4,
            char_offset: 0,
        },
        "=ABS(OR(-10 > 0, A2 < 100))",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 0,
            char_offset: 0,
        },
        "=LOWER(\"-6\")",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 1,
            char_offset: 0,
        },
        "=OR(CONCATENATE(\"33\", \"C1\") > 0, MAX(C8:E8) < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 2,
            char_offset: 0,
        },
        "=(PRODUCT(E2:E4) + C3)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 3,
            char_offset: 0,
        },
        "=MAX(D2:D9)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 4,
            char_offset: 0,
        },
        "3",
    );
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 2));
    println!("test_fuzz_reproducer_seed_643759 evaluated: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, -282.0),
        other => panic!("Expected Float(-282.0) for C7, got {:?}", other),
    }
}

#[test]
fn test_fuzz_reproducer_seed_8029() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 0,
            char_offset: 0,
        },
        "1",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 1,
            char_offset: 0,
        },
        "36",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 2,
            char_offset: 0,
        },
        "20",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 3,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 4,
            char_offset: 0,
        },
        "6",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 0,
            char_offset: 0,
        },
        "-61",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 1,
            char_offset: 0,
        },
        "\"bkh\"",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 2,
            char_offset: 0,
        },
        "419.3782",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 3,
            char_offset: 0,
        },
        "-80",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 4,
            char_offset: 0,
        },
        "100",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 0,
            char_offset: 0,
        },
        "91",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 1,
            char_offset: 0,
        },
        "21",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 2,
            char_offset: 0,
        },
        "\"oKzthzv\"",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 3,
            char_offset: 0,
        },
        "-76",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 4,
            char_offset: 0,
        },
        "-97",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 0,
            char_offset: 0,
        },
        "-2",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 1,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 2,
            char_offset: 0,
        },
        "232.9699",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 4,
            char_offset: 0,
        },
        "-91.42",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 0,
            char_offset: 0,
        },
        "-90",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 1,
            char_offset: 0,
        },
        "-69",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 2,
            char_offset: 0,
        },
        "-211.4693",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 3,
            char_offset: 0,
        },
        "45",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 4,
            char_offset: 0,
        },
        "10",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 0,
            char_offset: 0,
        },
        "=AND(IF((D4 > A2), B5, E1) > 0, CONCATENATE(\"A4\", \"D1\") < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 1,
            char_offset: 0,
        },
        "97.056",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 2,
            char_offset: 0,
        },
        "=INT(SUM(C2:D2))",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 3,
            char_offset: 0,
        },
        "=D2",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 4,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 0,
            char_offset: 0,
        },
        "=IF((D5 > (A2 / C1)), IF((A1 > A3), B3, A6), AND(32 > 0, D3 < 100))",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 1,
            char_offset: 0,
        },
        "=E3",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 2,
            char_offset: 0,
        },
        "=CONCATENATE(\"A2\", \"B2\")",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 3,
            char_offset: 0,
        },
        "=B3",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 4,
            char_offset: 0,
        },
        "=21",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 0,
            char_offset: 0,
        },
        "=MAX(D5:D6)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 1,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 2,
            char_offset: 0,
        },
        "=ROUNDDOWN(B7, 2)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 3,
            char_offset: 0,
        },
        "=IF((ROUNDUP(B4, 2) > INT(A6)), (-13 ^ -44), LEN(\"B2\"))",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 4,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 0,
            char_offset: 0,
        },
        "0",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 1,
            char_offset: 0,
        },
        "=ROUNDUP(D6, 0)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 2,
            char_offset: 0,
        },
        "=E6",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 3,
            char_offset: 0,
        },
        "=-24",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 4,
            char_offset: 0,
        },
        "=D6",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 0,
            char_offset: 0,
        },
        "=C1",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 1,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 2,
            char_offset: 0,
        },
        "=D9",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 3,
            char_offset: 0,
        },
        "=MIN(A6:B7)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 4,
            char_offset: 0,
        },
        "=CONCATENATE(\"D6\", \"A1\")",
    );
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 1));
    println!("test_fuzz_reproducer_seed_8029 evaluated: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, -80.0),
        other => panic!("Expected Float(-80.0) for B9, got {:?}", other),
    }
}

#[test]
fn test_fuzz_reproducer_seed_278502() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 0,
            char_offset: 0,
        },
        "-58",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 1,
            char_offset: 0,
        },
        "40",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 2,
            char_offset: 0,
        },
        "222",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 3,
            char_offset: 0,
        },
        "\"3ZE\"",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 4,
            char_offset: 0,
        },
        "-174.2544",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 0,
            char_offset: 0,
        },
        "\"yrSbQs\"",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 1,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 2,
            char_offset: 0,
        },
        "-270.4109",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 3,
            char_offset: 0,
        },
        "76",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 4,
            char_offset: 0,
        },
        "-71",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 0,
            char_offset: 0,
        },
        "71",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 1,
            char_offset: 0,
        },
        "-69",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 2,
            char_offset: 0,
        },
        "\"Kej\"",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 3,
            char_offset: 0,
        },
        "-450.002",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 0,
            char_offset: 0,
        },
        "25",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 1,
            char_offset: 0,
        },
        "77",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 2,
            char_offset: 0,
        },
        "6",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 3,
            char_offset: 0,
        },
        "\"W\"",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 4,
            char_offset: 0,
        },
        "\"zAhnlQyo\"",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 2,
            char_offset: 0,
        },
        "\"cm2\"",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 3,
            char_offset: 0,
        },
        "\"Hs\"",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 4,
            char_offset: 0,
        },
        "5",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 0,
            char_offset: 0,
        },
        "=PRODUCT(A2:D5)",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 1,
            char_offset: 0,
        },
        "=((C4 - A5) - MIN(C4:D4))",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 2,
            char_offset: 0,
        },
        "=CONCATENATE(\"ABS(B4)\", \"MAX(E1:E1)\")",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 3,
            char_offset: 0,
        },
        "=A3",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 4,
            char_offset: 0,
        },
        "=B5",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 0,
            char_offset: 0,
        },
        "=-45",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 1,
            char_offset: 0,
        },
        "=MAX(20, (E3 + C5))",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 2,
            char_offset: 0,
        },
        "=(ROUND(19, 2) - IF((C1 > 33), D6, A6))",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 3,
            char_offset: 0,
        },
        "218.1489",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 4,
            char_offset: 0,
        },
        "=-38",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 0,
            char_offset: 0,
        },
        "=B6",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 1,
            char_offset: 0,
        },
        "=(SQRT(-48) * E2)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 2,
            char_offset: 0,
        },
        "=(ROUND(B5, 1) * D2)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 3,
            char_offset: 0,
        },
        "=A7",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 4,
            char_offset: 0,
        },
        "=ROUND(ROUND(B2, 2), 1)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 0,
            char_offset: 0,
        },
        "=B5",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 1,
            char_offset: 0,
        },
        "=ROUNDUP(IF((E7 > C1), D1, A4), 0)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 2,
            char_offset: 0,
        },
        "494",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 3,
            char_offset: 0,
        },
        "400.726",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 4,
            char_offset: 0,
        },
        "=OR(-26 > 0, MAX(E8:E8) < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 0,
            char_offset: 0,
        },
        "=MAX(E1:E5)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 1,
            char_offset: 0,
        },
        "=OR(E7 > 0, (-39 * C8) < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 2,
            char_offset: 0,
        },
        "=CONCATENATE(\"(E1 * D7)\", \"INT(E5)\")",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 3,
            char_offset: 0,
        },
        "=CONCATENATE(\"MIN(C6:E8)\", \"30\")",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 4,
            char_offset: 0,
        },
        "211.5",
    );
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 2));
    println!("test_fuzz_reproducer_seed_278502 evaluated: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, -52.0),
        other => panic!("Expected Float(-52.0) for C7, got {:?}", other),
    }
}

#[test]
fn test_fuzz_reproducer_seed_842487() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 0,
            char_offset: 0,
        },
        "17",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 1,
            char_offset: 0,
        },
        "\"cxFkHkuB\"",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 2,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 3,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 4,
            char_offset: 0,
        },
        "-362.7",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 0,
            char_offset: 0,
        },
        "60",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 1,
            char_offset: 0,
        },
        "471.19",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 3,
            char_offset: 0,
        },
        "\"p\"",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 4,
            char_offset: 0,
        },
        "0",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 1,
            char_offset: 0,
        },
        "-41",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 2,
            char_offset: 0,
        },
        "-12.175",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 3,
            char_offset: 0,
        },
        "22.876",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 4,
            char_offset: 0,
        },
        "-127.099",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 1,
            char_offset: 0,
        },
        "435.35",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 2,
            char_offset: 0,
        },
        "77",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 3,
            char_offset: 0,
        },
        "\"clH\"",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 4,
            char_offset: 0,
        },
        "-5",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 0,
            char_offset: 0,
        },
        "166.9",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 1,
            char_offset: 0,
        },
        "29",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 2,
            char_offset: 0,
        },
        "63",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 3,
            char_offset: 0,
        },
        "-31",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 4,
            char_offset: 0,
        },
        "3",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 0,
            char_offset: 0,
        },
        "=A5",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 1,
            char_offset: 0,
        },
        "=IF(((-1 + E3) > MIN(A3:C4)), SQRT(E1), C1)",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 2,
            char_offset: 0,
        },
        "=LOWER(\"A3\")",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 3,
            char_offset: 0,
        },
        "=RIGHT(\"15\", 2)",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 4,
            char_offset: 0,
        },
        "44",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 0,
            char_offset: 0,
        },
        "500",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 1,
            char_offset: 0,
        },
        "=AND(E4 > 0, E5 < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 2,
            char_offset: 0,
        },
        "=-37",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 3,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 4,
            char_offset: 0,
        },
        "=(MIN(C4:E4) - IF((C2 > C2), -3, E4))",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 0,
            char_offset: 0,
        },
        "=IF((B6 > ABS(D5)), (-16 * -6), D3)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 1,
            char_offset: 0,
        },
        "=E2",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 2,
            char_offset: 0,
        },
        "=IF((OR(B2 > 0, E6 < 100) > ROUND(D5, 0)), LEN(\"-5\"), A6)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 3,
            char_offset: 0,
        },
        "=INT(B7)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 4,
            char_offset: 0,
        },
        "=MIN(E2:E2)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 0,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 1,
            char_offset: 0,
        },
        "=-14",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 2,
            char_offset: 0,
        },
        "=C5",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 3,
            char_offset: 0,
        },
        "=(E7 * LEN(\"D3\"))",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 4,
            char_offset: 0,
        },
        "=ABS(B7)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 0,
            char_offset: 0,
        },
        "=LEFT(\"B1\", 2)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 1,
            char_offset: 0,
        },
        "=RIGHT(\"C3\", 2)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 2,
            char_offset: 0,
        },
        "=33",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 3,
            char_offset: 0,
        },
        "-133.1505",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 4,
            char_offset: 0,
        },
        "=E4",
    );
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 3));
    println!("test_fuzz_reproducer_seed_842487 evaluated: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 0.0),
        other => panic!("Expected Float(0.0) for D9, got {:?}", other),
    }
}

#[test]
fn test_fuzz_reproducer_seed_507065() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 1,
            char_offset: 0,
        },
        "6",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 2,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 3,
            char_offset: 0,
        },
        "-372.1",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 4,
            char_offset: 0,
        },
        "-61",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 0,
            char_offset: 0,
        },
        "64",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 1,
            char_offset: 0,
        },
        "-79",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 2,
            char_offset: 0,
        },
        "-15",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 4,
            char_offset: 0,
        },
        "0",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 0,
            char_offset: 0,
        },
        "\"CVZDGbB\"",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 2,
            char_offset: 0,
        },
        "106.6",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 3,
            char_offset: 0,
        },
        "25",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 0,
            char_offset: 0,
        },
        "49",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 1,
            char_offset: 0,
        },
        "-86",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 2,
            char_offset: 0,
        },
        "56.984",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 3,
            char_offset: 0,
        },
        "18",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 4,
            char_offset: 0,
        },
        "424.2595",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 0,
            char_offset: 0,
        },
        "-99",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 1,
            char_offset: 0,
        },
        "-15",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 2,
            char_offset: 0,
        },
        "-68",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 3,
            char_offset: 0,
        },
        "-50",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 4,
            char_offset: 0,
        },
        "5",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 0,
            char_offset: 0,
        },
        "=ROUND(39, 1)",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 1,
            char_offset: 0,
        },
        "=SUM(21, (E2 * A1))",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 2,
            char_offset: 0,
        },
        "=LOWER(\"C4\")",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 3,
            char_offset: 0,
        },
        "=D2",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 4,
            char_offset: 0,
        },
        "=AND(A5 > 0, A3 < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 0,
            char_offset: 0,
        },
        "=D3",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 1,
            char_offset: 0,
        },
        "=ROUND(C5, 0)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 2,
            char_offset: 0,
        },
        "=-36",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 3,
            char_offset: 0,
        },
        "=ROUNDDOWN((A2 + A4), 2)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 4,
            char_offset: 0,
        },
        "=IF((29 > SUM(A5:A6)), B5, SUM(A5:C6))",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 0,
            char_offset: 0,
        },
        "=41",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 1,
            char_offset: 0,
        },
        "=-21",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 2,
            char_offset: 0,
        },
        "=A4",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 3,
            char_offset: 0,
        },
        "=IF((C6 > SUM(B1:B7)), AND(D2 > 0, 38 < 100), INT(E5))",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 4,
            char_offset: 0,
        },
        "=-9",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 0,
            char_offset: 0,
        },
        "=IF((MIN(E5:E6) > OR(C1 > 0, E1 < 100)), MAX(D8, D6), AVERAGE(C2:D8))",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 1,
            char_offset: 0,
        },
        "=-10",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 2,
            char_offset: 0,
        },
        "=E4",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 3,
            char_offset: 0,
        },
        "\"WypeLQcC\"",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 4,
            char_offset: 0,
        },
        "=IF((D8 > IF((E3 > C7), D5, -27)), LEN(\"C3\"), SUM(C5:D7))",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 0,
            char_offset: 0,
        },
        "9",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 1,
            char_offset: 0,
        },
        "=IF((IF((21 > B6), -5, C3) > LEFT(\"-36\", 5)), PRODUCT(A7:B7), ROUNDUP(D7, 2))",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 2,
            char_offset: 0,
        },
        "=B9",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 3,
            char_offset: 0,
        },
        "401",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 4,
            char_offset: 0,
        },
        "31.8",
    );
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 1));
    println!("test_fuzz_reproducer_seed_507065 evaluated: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 113.0),
        other => panic!("Expected Float(113.0) for B10, got {:?}", other),
    }
}

#[test]
fn test_fuzz_reproducer_seed_368811() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 0,
            char_offset: 0,
        },
        "-58",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 1,
            char_offset: 0,
        },
        "\"fij\"",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 2,
            char_offset: 0,
        },
        "59",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 3,
            char_offset: 0,
        },
        "-42.662",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 4,
            char_offset: 0,
        },
        "\"AVg\"",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 0,
            char_offset: 0,
        },
        "-354.803",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 1,
            char_offset: 0,
        },
        "-58",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 3,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 0,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 1,
            char_offset: 0,
        },
        "-62",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 2,
            char_offset: 0,
        },
        "\"qh\"",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 3,
            char_offset: 0,
        },
        "\"yl\"",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 4,
            char_offset: 0,
        },
        "0",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 1,
            char_offset: 0,
        },
        "-27",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 2,
            char_offset: 0,
        },
        "-29",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 3,
            char_offset: 0,
        },
        "100",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 4,
            char_offset: 0,
        },
        "104",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 0,
            char_offset: 0,
        },
        "-39",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 1,
            char_offset: 0,
        },
        "-220",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 2,
            char_offset: 0,
        },
        "-86",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 3,
            char_offset: 0,
        },
        "10",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 4,
            char_offset: 0,
        },
        "-95",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 0,
            char_offset: 0,
        },
        "=A5",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 1,
            char_offset: 0,
        },
        "=C4",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 2,
            char_offset: 0,
        },
        "=CONCATENATE(\"INT(-12)\", \"(A1 * A3)\")",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 3,
            char_offset: 0,
        },
        "=1",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 4,
            char_offset: 0,
        },
        "=(AND(B4 > 0, B2 < 100) * MAX(A4:B4))",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 1,
            char_offset: 0,
        },
        "1",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 2,
            char_offset: 0,
        },
        "=((B5 - D3) ^ C6)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 3,
            char_offset: 0,
        },
        "=LEN(\"(-48 + D3)\")",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 4,
            char_offset: 0,
        },
        "=SUM(D1:E5)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 0,
            char_offset: 0,
        },
        "40",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 1,
            char_offset: 0,
        },
        "=AVERAGE(B7:C7)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 2,
            char_offset: 0,
        },
        "=-45",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 3,
            char_offset: 0,
        },
        "=ROUNDUP(AND(5 > 0, C5 < 100), 0)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 4,
            char_offset: 0,
        },
        "=IF((OR(C4 > 0, 30 < 100) > -45), A6, UPPER(\"C3\"))",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 0,
            char_offset: 0,
        },
        "=AND(OR(-1 > 0, A1 < 100) > 0, MIN(D6:E7) < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 1,
            char_offset: 0,
        },
        "=PRODUCT(ROUND(B8, 1), 38)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 2,
            char_offset: 0,
        },
        "=B4",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 4,
            char_offset: 0,
        },
        "=PRODUCT(E7:E8)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 0,
            char_offset: 0,
        },
        "-62",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 1,
            char_offset: 0,
        },
        "=37",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 2,
            char_offset: 0,
        },
        "=AVERAGE(E1, B5)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 3,
            char_offset: 0,
        },
        "=INT(13)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 4,
            char_offset: 0,
        },
        "=(AND(-13 > 0, 22 < 100) * IF((36 > D2), A7, 21))",
    );
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 2));
    println!("test_fuzz_reproducer_seed_368811 evaluated: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, -220.0),
        other => panic!("Expected Float(-220.0) for C10, got {:?}", other),
    }
}

#[test]
fn test_fuzz_reproducer_seed_357041() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 1,
            char_offset: 0,
        },
        "88",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 3,
            char_offset: 0,
        },
        "-219.79",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 4,
            char_offset: 0,
        },
        "291",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 0,
            char_offset: 0,
        },
        "\"C\"",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 1,
            char_offset: 0,
        },
        "5",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 2,
            char_offset: 0,
        },
        "\"sXn\"",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 3,
            char_offset: 0,
        },
        "-357.03",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 4,
            char_offset: 0,
        },
        "TRUE",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 0,
            char_offset: 0,
        },
        "-51",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 1,
            char_offset: 0,
        },
        "-6",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 2,
            char_offset: 0,
        },
        "-277.064",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 4,
            char_offset: 0,
        },
        "\"L\"",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 0,
            char_offset: 0,
        },
        "77",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 1,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 2,
            char_offset: 0,
        },
        "-50",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 3,
            char_offset: 0,
        },
        "39",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 4,
            char_offset: 0,
        },
        "-330.8664",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 0,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 1,
            char_offset: 0,
        },
        "74",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 2,
            char_offset: 0,
        },
        "300.7576",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 3,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 4,
            char_offset: 0,
        },
        "7",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 0,
            char_offset: 0,
        },
        "=LOWER(\"ROUNDDOWN(A4, 1)\")",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 1,
            char_offset: 0,
        },
        "=(MAX(D3, E5) - (-41 * D4))",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 2,
            char_offset: 0,
        },
        "=(A3 ^ A1)",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 3,
            char_offset: 0,
        },
        "=A2",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 4,
            char_offset: 0,
        },
        "=-8",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 0,
            char_offset: 0,
        },
        "=(C3 / B3)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 1,
            char_offset: 0,
        },
        "=-30",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 2,
            char_offset: 0,
        },
        "=LEFT(\"AVERAGE(B6:C6)\", 4)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 3,
            char_offset: 0,
        },
        "=D3",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 4,
            char_offset: 0,
        },
        "0",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 0,
            char_offset: 0,
        },
        "=(D5 - B1)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 1,
            char_offset: 0,
        },
        "=OR(ROUNDDOWN(-37, 0) > 0, MAX(A3:A7) < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 2,
            char_offset: 0,
        },
        "=D5",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 3,
            char_offset: 0,
        },
        "=ROUND((E5 / 15), 2)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 4,
            char_offset: 0,
        },
        "=OR((A5 ^ 46) > 0, LEFT(\"D6\", 1) < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 0,
            char_offset: 0,
        },
        "=E4",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 1,
            char_offset: 0,
        },
        "=SQRT(UPPER(\"A6\"))",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 2,
            char_offset: 0,
        },
        "=LOWER(\"47\")",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 3,
            char_offset: 0,
        },
        "=OR(6 > 0, 24 < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 4,
            char_offset: 0,
        },
        "=D5",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 0,
            char_offset: 0,
        },
        "=INT(E6)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 1,
            char_offset: 0,
        },
        "=B4",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 2,
            char_offset: 0,
        },
        "=((D6 ^ D8) ^ E3)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 3,
            char_offset: 0,
        },
        "-28",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 4,
            char_offset: 0,
        },
        "=-3",
    );
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    println!("test_fuzz_reproducer_seed_357041 evaluated: {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, -8.0),
        other => panic!("Expected Float(-8.0) for A10, got {:?}", other),
    }
}

#[test]
fn test_fuzz_reproducer_seed_320979() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 0,
            char_offset: 0,
        },
        "-292.46",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 1,
            char_offset: 0,
        },
        "\"UCmL\"",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 2,
            char_offset: 0,
        },
        "\"csEXNMm\"",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 3,
            char_offset: 0,
        },
        "\"hpNtY\"",
    );
    sheet.insert(
        TextCellRef {
            row: 0,
            col: 4,
            char_offset: 0,
        },
        "74",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 0,
            char_offset: 0,
        },
        "FALSE",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 1,
            char_offset: 0,
        },
        "-57",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 2,
            char_offset: 0,
        },
        "99",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 3,
            char_offset: 0,
        },
        "34",
    );
    sheet.insert(
        TextCellRef {
            row: 1,
            col: 4,
            char_offset: 0,
        },
        "0",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 0,
            char_offset: 0,
        },
        "-45",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 1,
            char_offset: 0,
        },
        "25",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 2,
            char_offset: 0,
        },
        "445.79",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 3,
            char_offset: 0,
        },
        "-19",
    );
    sheet.insert(
        TextCellRef {
            row: 2,
            col: 4,
            char_offset: 0,
        },
        "\"QYFhRm\"",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 0,
            char_offset: 0,
        },
        "48",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 1,
            char_offset: 0,
        },
        "410.85",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 2,
            char_offset: 0,
        },
        "3",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 3,
            char_offset: 0,
        },
        "186.346",
    );
    sheet.insert(
        TextCellRef {
            row: 3,
            col: 4,
            char_offset: 0,
        },
        "9",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 0,
            char_offset: 0,
        },
        "-89",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 1,
            char_offset: 0,
        },
        "24",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 2,
            char_offset: 0,
        },
        "58.732",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 3,
            char_offset: 0,
        },
        "3",
    );
    sheet.insert(
        TextCellRef {
            row: 4,
            col: 4,
            char_offset: 0,
        },
        "-74",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 0,
            char_offset: 0,
        },
        "10",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 1,
            char_offset: 0,
        },
        "=E2",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 2,
            char_offset: 0,
        },
        "=((A3 / 20) + SQRT(B2))",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 3,
            char_offset: 0,
        },
        "44",
    );
    sheet.insert(
        TextCellRef {
            row: 5,
            col: 4,
            char_offset: 0,
        },
        "=ABS(AND(E3 > 0, B3 < 100))",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 0,
            char_offset: 0,
        },
        "=LEFT(\"A2\", 5)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 1,
            char_offset: 0,
        },
        "=IF((PRODUCT(E4:E4) > (A4 + C3)), E4, 35)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 2,
            char_offset: 0,
        },
        "=INT(-30)",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 3,
            char_offset: 0,
        },
        "=(CONCATENATE(\"B6\", \"B3\") * SUM(46, A2))",
    );
    sheet.insert(
        TextCellRef {
            row: 6,
            col: 4,
            char_offset: 0,
        },
        "=OR(D3 > 0, MIN(D6:E6) < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 0,
            char_offset: 0,
        },
        "=37",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 1,
            char_offset: 0,
        },
        "-79",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 2,
            char_offset: 0,
        },
        "=PRODUCT(A3, E1)",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 3,
            char_offset: 0,
        },
        "235.4",
    );
    sheet.insert(
        TextCellRef {
            row: 7,
            col: 4,
            char_offset: 0,
        },
        "=AND((2 - -9) > 0, RIGHT(\"43\", 3) < 100)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 0,
            char_offset: 0,
        },
        "=D4",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 1,
            char_offset: 0,
        },
        "=PRODUCT(E6:E7)",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 2,
            char_offset: 0,
        },
        "=(MAX(A3:B3) - LOWER(\"-30\"))",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 3,
            char_offset: 0,
        },
        "=E5",
    );
    sheet.insert(
        TextCellRef {
            row: 8,
            col: 4,
            char_offset: 0,
        },
        "=LEN(\"IF((-50 > A3), -23, E4)\")",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 0,
            char_offset: 0,
        },
        "0",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 1,
            char_offset: 0,
        },
        "\"qq\"",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 2,
            char_offset: 0,
        },
        "=PRODUCT(A2:C5)",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 3,
            char_offset: 0,
        },
        "=MAX(AND(C1 > 0, B7 < 100), CONCATENATE(\"D1\", \"12\"))",
    );
    sheet.insert(
        TextCellRef {
            row: 9,
            col: 4,
            char_offset: 0,
        },
        "=ABS(IF((A3 > 46), 7, 41))",
    );
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 3));
    println!("test_fuzz_reproducer_seed_320979 evaluated: {:?}", target);
    assert!(
        matches!(target, ResultData::Error(_)),
        "Expected Error for D10, got {:?}",
        target
    );
}

#[test]
fn test_fuzz_reproducer_seed_834997() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "-80");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "-85");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "40");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-95");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "30");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-452.5");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "\"LVnJoc\"");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "30");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "-77.1208");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "\"2cB\"");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "249.6382");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "65");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "-33");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "\"UbxkVL \"");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "53");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "32");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "-38");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "\"yOaI\"");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "\"lQNfLl1d\"");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=D3");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "4");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=SUM(D3:E5)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=A5");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=AVERAGE(D1:E6)");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=UPPER(\"SQRT(B6)\")");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=(IF((-21 > D1), -14, C4) - SQRT(A4))");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=C1");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=MIN(B6:B7)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=(OR(E4 > 0, C3 < 100) * C2)");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=OR(LEFT(\"A6\", 1) > 0, E5 < 100)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=32");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=B5");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "36");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=(E4 ^ D7)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=AVERAGE(E8:E8)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=OR(IF((D5 > 20), 41, D4) > 0, (A3 - C3) < 100)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=B6");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=C7");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=25");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=B7");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=B5");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=MAX(C6:D9)");
    sheet.commit(None).unwrap();
    println!("Running seed 834997 repro...");
    let target = sheet.get_result_data(&CellRef::new(8, 1));
    println!("Seed 834997 evaluated target CellRef(8, 1): {:?}", target);

}

#[test]
fn test_fuzz_reproducer_seed_469392() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "82");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "5");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "91");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "\"zLW\"");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-52");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "5");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "\" ey\"");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "-231.6");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "-70");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "-9");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "72");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "150");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "286.5");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "-11");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "\"Nup Bkxr\"");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "-12");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "-419.2166");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "-78");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "\"Jxd2 \"");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "12");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "1");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "17");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=UPPER(\"OR(-17 > 0, 10 < 100)\")");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=RIGHT(\"ROUNDDOWN(E4, 1)\", 5)");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=D4");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=(PRODUCT(A4, C4) * C4)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=LEN(\"E5\")");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=B6");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=ABS(ABS(-33))");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=-20");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "\"CkAx1j\"");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=A6");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=C5");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=UPPER(\"ROUNDUP(C3, 2)\")");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=PRODUCT(IF((28 > -42), -24, D3), -3)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=MIN(A6:E6)");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "-60.743");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=(-12 ^ D4)");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=OR(RIGHT(\"A2\", 2) > 0, C5 < 100)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=(C4 * IF((B8 > D5), C5, 28))");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=A5");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=AND(-45 > 0, AND(A2 > 0, D3 < 100) < 100)");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "-10");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=ROUNDUP((A1 ^ C8), 1)");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=ROUND(IF((A3 > E7), -6, C4), 0)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=16");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=ROUNDUP(D8, 1)");
    sheet.commit(None).unwrap();
    println!("Running seed 469392 repro...");
    let target = sheet.get_result_data(&CellRef::new(9, 1));
    println!("Seed 469392 evaluated target CellRef(9, 1): {:?}", target);

}

#[test]
fn test_fuzz_reproducer_seed_101881() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "-41");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "1");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "-84");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "55");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "-93");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "6");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "-1");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "436");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "3");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-384.63");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "4");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "-26");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "24");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "475.29");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "-90.40000000000001");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "-15");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "\"ArNOnb\"");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "100");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=A4");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=E4");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=OR(2 > 0, A4 < 100)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=B5");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=C4");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=LEFT(\"ROUNDDOWN(A4, 0)\", 2)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=C3");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=ROUNDUP((E1 + D6), 0)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=(E3 * 25)");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=AND(CONCATENATE(\"E2\", \"A2\") > 0, (E3 / 14) < 100)");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=AND(-10 > 0, E3 < 100)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=B3");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=-50");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=(AVERAGE(B5:C6) - -28)");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=28");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=LOWER(\"(B3 ^ D1)\")");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=IF((UPPER(\"E4\") > RIGHT(\"C4\", 1)), C5, ROUNDDOWN(-10, 2))");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=SQRT(LEFT(\"A2\", 3))");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=LEFT(\"B6\", 5)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=(19 - (-45 * 49))");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=A3");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "-344.119");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=D2");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=(INT(-21) ^ (C8 - D8))");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=AVERAGE(MIN(-47, B4), E2)");
    sheet.commit(None).unwrap();
    println!("Running seed 101881 repro...");
    let target = sheet.get_result_data(&CellRef::new(6, 2));
    println!("Seed 101881 evaluated target CellRef(6, 2): {:?}", target);

}

#[test]
fn test_fuzz_reproducer_seed_833777() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "46");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "-69");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-50.9");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "4");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "-2.81");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "52");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "27");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "38");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "18");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-23");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "108.8");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "\"IZKTZKJ\"");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "193.2");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-52");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "-310.4212");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "\"kn\"");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "32");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "-94.88");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "199.402");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "-1.82");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=48");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=LEN(\"D1\")");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=(A3 * INT(A3))");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=46");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=A3");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=26");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=SQRT(42)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=LOWER(\"B5\")");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=A6");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=AVERAGE(E2, ROUNDDOWN(10, 2))");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=C4");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=C7");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=ROUNDDOWN(MAX(C1, -49), 0)");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=-6");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=C4");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "5");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=MAX(E6, 40)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=6");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=(OR(A2 > 0, -3 < 100) / ROUND(-38, 0))");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "-430.0124");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=AND(IF((D1 > D3), 38, D3) > 0, E7 < 100)");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=(33 ^ CONCATENATE(\"D8\", \"A8\"))");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=((C6 / B3) + IF((C1 > E8), D6, C9))");
    sheet.commit(None).unwrap();
    println!("Running seed 833777 repro...");
    let target = sheet.get_result_data(&CellRef::new(9, 1));
    println!("Seed 833777 evaluated target CellRef(9, 1): {:?}", target);

}

#[test]
fn test_fuzz_reproducer_seed_473592() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "-75");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "85");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "\"nL\"");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "\"oroIgv\"");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "\"shG\"");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "21");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "\"yqt1\"");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "14.0734");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "416.3");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "88");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "-73.29000000000001");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "3");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "-69");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "151.8448");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "259.92");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "-32");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "95");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "-350.02");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=MAX(A2:E2)");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=ABS(RIGHT(\"30\", 2))");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=RIGHT(\"B2\", 2)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=(MAX(D4:D4) - C5)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=6");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=((C3 - C4) - C4)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=INT(OR(C3 > 0, 48 < 100))");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=INT(LEN(\"E4\"))");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=INT(A2)");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=OR(E4 > 0, C6 < 100)");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=D5");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=ROUND(SUM(B5:E6), 2)");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=D6");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=SQRT(IF((-41 > A6), D7, A3))");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "-195.6794");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "-13");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=A4");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=MAX(INT(D3), (E5 * C8))");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=28");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "249");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "224.0393");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=AVERAGE(D4, UPPER(\"E4\"))");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=IF(((B7 + 3) > LEN(\"D2\")), LEN(\"-41\"), (C3 / -45))");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=ABS(D8)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=AND(ABS(33) > 0, IF((D5 > 43), A1, -44) < 100)");
    sheet.commit(None).unwrap();
    println!("Running seed 473592 repro...");
    let target = sheet.get_result_data(&CellRef::new(8, 2));
    println!("Seed 473592 evaluated target CellRef(8, 2): {:?}", target);

}

#[test]
fn test_fuzz_reproducer_seed_717209() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "28");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "68");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "\"E \"");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "-99");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "23");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "6");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "\"bb\"");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-271.85");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "-78");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "\"gEqS\"");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-43");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "29");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "\"A\"");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "88");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "\"R\"");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "-66");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "81");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "-8");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=(INT(B3) / (-11 ^ D1))");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=D5");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "-80");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=AND((7 - D3) > 0, AVERAGE(C3:C5) < 100)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=MAX(-34, PRODUCT(A5:E5))");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=INT(A5)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=A1");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=29");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=RIGHT(\"IF((D5 > E1), D5, E2)\", 1)");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=AVERAGE(A1, RIGHT(\"D6\", 1))");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=16");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=PRODUCT(C7, (D4 + C7))");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=AVERAGE(C5:D5)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=OR(-23 > 0, -21 < 100)");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "-56");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=PRODUCT(AND(11 > 0, A1 < 100), (A6 ^ A3))");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=C6");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=-19");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=B5");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=A4");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=E3");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=E4");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=LEFT(\"A1\", 4)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=B6");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=C1");
    sheet.commit(None).unwrap();
    println!("Running seed 717209 repro...");
    let target = sheet.get_result_data(&CellRef::new(6, 4));
    println!("Seed 717209 evaluated target CellRef(6, 4): {:?}", target);

}

#[test]
fn test_fuzz_reproducer_seed_194393() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "\"mQK1CAkE\"");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "\"fUeQvRXH\"");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "3");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "386.992");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "\"1QU\"");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "7");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "\"rTrosz\"");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "\"j2q\"");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "-344.821");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "-41");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "1");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "-91");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "-397");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "69");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "-14");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "463.5482");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "\"FWKVlndl\"");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=SQRT(ROUNDUP(31, 0))");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=A2");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=ROUND(ABS(A5), 2)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=37");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=B5");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=E6");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=D1");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=C5");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=C5");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=(-13 + MAX(E6:E6))");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=ROUND(-40, 2)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=(B7 * MAX(D2, C5))");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=A1");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "28");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=MAX((B3 + E6), SUM(C3:D5))");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=IF((ABS(C8) > CONCATENATE(\"31\", \"A6\")), LEN(\"-45\"), B8)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=((B1 ^ E4) / MIN(A6:B8))");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=IF((C5 > (C4 * C6)), (B3 / E8), MIN(D1, D5))");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=CONCATENATE(\"SUM(-33, A4)\", \"AVERAGE(A5:A8)\")");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=C3");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=-32");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=INT(E6)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "-287.722");
    sheet.commit(None).unwrap();
    println!("Running seed 194393 repro...");
    let target = sheet.get_result_data(&CellRef::new(8, 3));
    println!("Seed 194393 evaluated target CellRef(8, 3): {:?}", target);

}

#[test]
fn test_fuzz_reproducer_seed_249481() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "-75");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "\"meC nxb\"");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "1");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "-443");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "358.3");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "63");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "73");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "423.462");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "\"vJvhAgQw\"");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "83");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-98");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "-182");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "-25");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "397.034");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "\"smWGqaU\"");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=LEFT(\"D3\", 5)");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=(A5 + LOWER(\"-20\"))");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=OR(C5 > 0, IF((D1 > 38), 28, C2) < 100)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=(C3 * E5)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=A5");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=AVERAGE((B3 + D1), D4)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=AND(B4 > 0, INT(E4) < 100)");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=23");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=16");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=15");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=B5");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=B7");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=(LOWER(\"A5\") + B2)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=((-46 / C4) ^ B4)");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=MAX(SQRT(E1), ABS(D2))");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=(B6 + D8)");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=LOWER(\"AVERAGE(C8:E8)\")");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=A1");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=MIN(D1:E1)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=((8 / D2) + OR(B4 > 0, 18 < 100))");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=26");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=A5");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=OR(17 > 0, C3 < 100)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=UPPER(\"IF((B2 > C3), C1, -20)\")");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=-29");
    sheet.commit(None).unwrap();
    println!("Running seed 249481 repro...");
    let target = sheet.get_result_data(&CellRef::new(8, 0));
    println!("Seed 249481 evaluated target CellRef(8, 0): {:?}", target);

}

#[test]
fn test_fuzz_reproducer_seed_651135() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "88");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "-275.4692");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "\"p\"");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-84");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "-85");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "17");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "\"V2ETUs\"");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "64");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "100");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "-177.9");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "\"apdKJDV\"");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "-359");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "-34");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "6");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-33");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "-348.65");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "92");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "-96");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "6");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=INT(D3)");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=(CONCATENATE(\"-19\", \"40\") * OR(C1 > 0, E1 < 100))");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "\"KBdDN1kV\"");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=ROUNDUP(SQRT(-27), 2)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=RIGHT(\"-17\", 4)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "\"r3TN\"");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=AND(25 > 0, UPPER(\"-13\") < 100)");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=PRODUCT(A6, B5)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "89");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=AVERAGE(D1, AVERAGE(B3:C5))");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=(MIN(E7, D4) ^ B6)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=-50");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=(B4 * ROUNDDOWN(A1, 2))");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=UPPER(\"ABS(D1)\")");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=-39");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=B8");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=MIN(D3:D3)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "96");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=AVERAGE(ROUND(-25, 1), IF((21 > B5), -46, E3))");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=SUM(A8:A8)");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=B5");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=(ABS(C6) * INT(C4))");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "\"q\"");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=INT((A9 * -46))");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "\"NcJ\"");
    sheet.commit(None).unwrap();
    println!("Running seed 651135 repro...");
    let target = sheet.get_result_data(&CellRef::new(7, 0));
    println!("Seed 651135 evaluated target CellRef(7, 0): {:?}", target);

}

#[test]
fn test_fuzz_reproducer_seed_26662() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "4");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "\"RfgVD2\"");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "\"oTxGNKnS\"");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "66");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-19");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "-61");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "81");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "\"xSB\"");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "-426");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "229.424");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "25");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "-71");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "-360.75");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "\"p2s\"");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "3");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "98");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "\"b\"");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "33");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=AVERAGE(D1:D1)");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=D1");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=A3");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=C4");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=MIN(INT(E4), LEN(\"A4\"))");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=IF((LEN(\"C4\") > -8), E1, ABS(B1))");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=PRODUCT(C2:D4)");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=(LOWER(\"31\") + B3)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=B5");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=MAX(E5:E6)");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=ROUNDUP(C2, 0)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=IF((ROUNDDOWN(A1, 0) > E1), (D1 ^ B7), (26 + D4))");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=(INT(A2) ^ ABS(-7))");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=-25");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=MIN(A7, SUM(C6:C7))");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "\"hbKN1Pua\"");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "-5");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=E8");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=C6");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "23");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=LOWER(\"A9\")");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=(D8 + MIN(-32, -17))");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=D3");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=UPPER(\"E5\")");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=ROUND(B6, 0)");
    sheet.commit(None).unwrap();
    println!("Running seed 26662 repro...");
    let target = sheet.get_result_data(&CellRef::new(9, 1));
    println!("Seed 26662 evaluated target CellRef(9, 1): {:?}", target);

}
#[test]
fn test_fuzz_reproducer_seed_938517() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "34");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "14");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "-89");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "-34");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "\"XMnG\"");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "7");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "-75");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "-231");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "7");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "\"P2z1\"");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "-76");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "17");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "3");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "79");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=C4");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=PRODUCT(A2:A5)");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=UPPER(\"SQRT(A5)\")");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=E1");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "-10.011");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=ROUND(D1, 2)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=INT(RIGHT(\"-3\", 3))");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=(D6 + ROUNDUP(C6, 0))");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=E4");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=INT(SUM(-42, C2))");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=IF((C6 > LEN(\"D4\")), ROUND(A1, 1), B5)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=6");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=MIN(A1:C5)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=3");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=-7");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=(ROUNDDOWN(C1, 0) + A4)");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "187.4");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=-42");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=INT(A5)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "1");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=(C9 * E1)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "-140.95");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=UPPER(\"(-46 + A2)\")");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=-7");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=E1");
    sheet.commit(None).unwrap();
    println!("Running seed 938517 repro...");
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    println!("Seed 938517 evaluated target CellRef(9, 0): {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 3738.0),
        other => panic!("Expected Float(3738.0) for A10, got {:?}", other),
    }

}

#[test]
fn test_fuzz_reproducer_seed_673397() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "\"c2B\"");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "-289.203");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "395.6");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "161");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "10");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "164.73");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "431");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "2");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "52");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "16");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "-434.5");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "23");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "-15");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "4");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "\" vW\"");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "-1.81");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "-11");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-77");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "\"siz\"");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "-38.13");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "\"mXBoJ\"");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "-380.7094");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "\"pTbV\"");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "-63");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "28");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=SQRT(LOWER(\"B5\"))");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=-12");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=(IF((E4 > E2), 29, A3) * -1)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=A1");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=AVERAGE(C5:E5)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=B4");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=MAX(PRODUCT(C2:C2), OR(B6 > 0, E4 < 100))");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "-67");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "-13");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=-33");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "-340.75");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=-35");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=MAX(LEFT(\"4\", 2), AVERAGE(A7, -38))");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=ABS(UPPER(\"-4\"))");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=ROUND(INT(A5), 0)");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=B3");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=UPPER(\"C7\")");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=C4");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=D6");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=(A8 - OR(33 > 0, C7 < 100))");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=ROUNDUP(SUM(C1, C4), 2)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=CONCATENATE(\"ROUNDDOWN(B2, 0)\", \"INT(A9)\")");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "-77");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=INT(A4)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "5");
    sheet.commit(None).unwrap();
    println!("Running seed 673397 repro...");
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    println!("Seed 673397 evaluated target CellRef(9, 0): {:?}", target);
    match target {
        ResultData::Float(f) => assert!((f - 318.6).abs() < 1e-3, "Expected ~318.6 for A10, got {}", f),
        other => panic!("Expected Float(~318.6) for A10, got {:?}", other),
    }

}

#[test]
fn test_fuzz_reproducer_seed_41112() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "-23");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "\"OElW\"");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "72");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "1");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "413.757");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "-2");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "60");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "-71");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "66");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "-86");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "8");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "33");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "209.28");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "79");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "\"f\"");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "162.008");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "\"ZKu\"");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "\"2dtEYT\"");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "\"2qdSZiqu\"");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "26");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "2");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=B4");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "41");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=AVERAGE(A4:A5)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=B2");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=E2");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=ROUNDUP(D5, 0)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=(AND(C2 > 0, A6 < 100) + (35 * 0))");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=OR(D3 > 0, IF((D6 > E2), A2, E6) < 100)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=SUM((26 - D2), A4)");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=IF(((B6 * D5) > ROUNDUP(A6, 2)), (C1 + -42), (A6 - D1))");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=E1");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=B6");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=ROUND(CONCATENATE(\"-3\", \"E7\"), 1)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "21");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=MIN(A2:D7)");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=28");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=RIGHT(\"D3\", 2)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=AVERAGE(IF((E2 > C1), C2, E3), C8)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=(-40 ^ UPPER(\"29\"))");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=E2");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=E8");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=(1 - A6)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=A8");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=(ROUNDUP(21, 1) - SQRT(42))");
    sheet.commit(None).unwrap();
    println!("Running seed 41112 repro...");
    let target = sheet.get_result_data(&CellRef::new(8, 3));
    println!("Seed 41112 evaluated target CellRef(8, 3): {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, -30000000.0),
        other => panic!("Expected Float(-30000000.0) for D9, got {:?}", other),
    }

}
#[test]
fn test_fuzz_reproducer_seed_618606() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "79.54000000000001");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "34.0806");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-410");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-336.1717");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "\"Dz\"");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "-90");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "80");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "-11");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "-88.38");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "-2");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "-291");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "39");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "228.77");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "\"jo\"");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "-41");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "8.827999999999999");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "309.53");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "1");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "98");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=-16");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=AVERAGE(C5:E5)");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=A5");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=(ROUNDDOWN(A2, 2) ^ IF((45 > E2), C2, 21))");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=UPPER(\"C4\")");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=AND(UPPER(\"B4\") > 0, LEFT(\"C1\", 3) < 100)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=ABS(C6)");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=AVERAGE(E1, E5)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=D6");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=E2");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=MAX(E3:E3)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "-282");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=D5");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=MAX(IF((B3 > C7), -36, 16), RIGHT(\"C5\", 3))");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=PRODUCT(INT(44), E1)");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=D6");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "-56");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=D1");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=-32");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=(SUM(E3:E9) ^ 35)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=ROUNDDOWN((B1 ^ -9), 0)");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=B4");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=LEN(\"D3\")");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=SUM(B2:E3)");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 1));
    println!("Seed 618606 evaluated target CellRef(6, 1) B7: {:?}", target);
    for r in 0..10 {
        for c in 0..5 {
            let res = sheet.get_result_data(&CellRef::new(r, c));
            if matches!(res, ResultData::Error(_)) {
                let col_let = (b'A' + c as u8) as char;
                println!("  Error cell {}{}: {:?}", col_let, r + 1, res);
            }
        }
    }
}
#[test]
fn test_fuzz_reproducer_seed_935638() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "\"2DeBb\"");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "\"aVisZWJZ\"");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "17");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "\"O\"");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "\"Y3i\"");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "\"2jvd\"");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "\"Dpnvny\"");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "-49");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "33");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "98");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "\"fljvy\"");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "-373.3952");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "-19");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "4");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "\"gipC1lq\"");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "65");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "22");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "-62");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "\"RrDIE\"");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "-51");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "59");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=CONCATENATE(\"INT(-28)\", \"-43\")");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=28");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=10");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=(SUM(C2:D2) * E4)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=A4");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=A6");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "163.57");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=ROUND(CONCATENATE(\"C2\", \"C5\"), 2)");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=ROUNDDOWN(C3, 0)");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "\"J\"");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=34");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=ROUND(LOWER(\"-34\"), 1)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "-97");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=(D5 - RIGHT(\"B5\", 3))");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=SQRT(D6)");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "-77");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=(-46 + (-3 ^ E5))");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=OR(A7 > 0, (A4 + E6) < 100)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=C5");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=MIN(A7:C8)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=D2");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "-283.3129");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=ROUNDUP(SUM(D2:D7), 2)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=A5");
    sheet.commit(None).unwrap();
    let target1 = sheet.get_result_data(&CellRef::new(8, 0));
    println!("Seed 935638 target1 CellRef(8, 0) A9: {:?}", target1);
    let target2 = sheet.get_result_data(&CellRef::new(8, 3));
    println!("Seed 935638 target2 CellRef(8, 3) D9: {:?}", target2);
    for r in 0..10 {
        for c in 0..5 {
            let res = sheet.get_result_data(&CellRef::new(r, c));
            if matches!(res, ResultData::Error(_)) {
                let col_let = (b'A' + c as u8) as char;
                println!("  Error cell {}{}: {:?}", col_let, r + 1, res);
            }
        }
    }
}
#[test]
fn test_fuzz_reproducer_seed_91404() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "\"YDAf2XmI\"");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "12.4026");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "\"id\"");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "\"Gui\"");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "-41");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "346.342");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "\"f\"");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "-386");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "12");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "-36");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "160.585");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-359.5374");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "45");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "322.06");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "38");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-30");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "\"V\"");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "2");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "41");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "67");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "47.3");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=D5");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=INT(LEFT(\"-12\", 1))");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=ROUNDDOWN(MIN(E3:E4), 2)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=AND(E5 > 0, 33 < 100)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=D5");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=D3");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=E1");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=((E6 / 41) + ROUND(8, 2))");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=AVERAGE(IF((C1 > C6), B5, B4), ROUND(C6, 2))");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=-9");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=D3");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=IF((MAX(C7:E7) > -17), RIGHT(\"C1\", 3), AND(16 > 0, C4 < 100))");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=SQRT(A2)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=-7");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=MAX(UPPER(\"C3\"), MAX(D4:E4))");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=D7");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=IF((IF((C6 > B2), D5, D4) > C3), OR(B2 > 0, A8 < 100), (C8 - D5))");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=C8");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=AND(INT(-7) > 0, OR(E2 > 0, B7 < 100) < 100)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=AVERAGE(ROUNDDOWN(E3, 2), C2)");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "-43");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=((-46 * C3) / IF((-25 > A3), C7, C3))");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=UPPER(\"E3\")");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=ROUNDUP(SUM(C3:D5), 0)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=IF((MAX(A2:D9) > B9), ROUNDUP(D5, 2), E4)");
    sheet.commit(None).unwrap();
    let target1 = sheet.get_result_data(&CellRef::new(7, 1));
    println!("Seed 91404 target1 CellRef(7, 1) B8: {:?}", target1);
    let target2 = sheet.get_result_data(&CellRef::new(6, 2));
    println!("Seed 91404 target2 CellRef(6, 2) C7: {:?}", target2);
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
fn test_fuzz_reproducer_seed_770254() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "7");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "74");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "57");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "43.65");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "37");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "42");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "-55");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "86");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "-8");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-483.6633");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "89");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "230");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "4");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "\"QCOAtn\"");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "310.526");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "-89.5");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "34");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "-80");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=-10");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=A1");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=MIN(A2:E3)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=(42 - OR(-45 > 0, D3 < 100))");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=IF((ABS(D2) > E4), MAX(E5, -30), B5)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "16");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=OR((D2 * B6) > 0, LEFT(\"10\", 3) < 100)");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=D2");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=D5");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "462.672");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=PRODUCT(ABS(17), (-17 ^ A4))");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=C6");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=((C6 ^ D6) + PRODUCT(C3, A3))");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "7");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=SQRT(MIN(E4:E4))");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=(28 * (28 - E5))");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=B1");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "-492");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=A2");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=(UPPER(\"E1\") / A2)");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=A8");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=E5");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=A5");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=E6");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=IF((ABS(C5) > 18), (C2 * 36), (-16 * B2))");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 2));
    match target {
        ResultData::Float(f) => assert!((f - -1.1648460276016541e110).abs() / 1e110 < 1e-3),
        other => panic!("Expected Float(~-1.1648e110), got {:?}", other),
    }
}
#[test]
fn test_fuzz_reproducer_seed_542519() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "\"vegfg\"");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "-295.32");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "25");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "\"JwVJ\"");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "-324.78");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "7");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "-49");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "-202.3");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "-353");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-98");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "-5");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "-41");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "\"Ya\"");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "19");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "4");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "-18");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "339.594");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "16");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "39");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=C5");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=-6");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=IF((C4 > MIN(A1, C5)), (A2 - 7), ROUNDDOWN(31, 2))");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=E5");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=D3");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=IF((IF((B4 > E1), D6, 26) > AVERAGE(E1, 23)), ROUNDDOWN(15, 1), D2)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=RIGHT(\"(E6 ^ 17)\", 1)");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=MAX(B4:D4)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=AND(UPPER(\"B3\") > 0, (B4 / D1) < 100)");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=B5");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=13");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=(IF((48 > -1), A6, -41) - (E3 + B7))");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=AND((40 / 36) > 0, LOWER(\"C2\") < 100)");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=A7");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=AVERAGE(A8:A8)");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=MAX(-42, LOWER(\"C3\"))");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=SUM(OR(47 > 0, -49 < 100), (E8 + 35))");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=ROUND(SQRT(1), 0)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=B4");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=((D7 - A4) / AND(D5 > 0, E5 < 100))");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "-15");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=B5");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=OR(-15 > 0, 10 < 100)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "51.9");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 2));
    println!("Seed 542519 evaluated target CellRef(8, 2): {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 51.0),
        other => panic!("Expected Float(51.0), got {:?}", other),
    }
}
#[test]
fn test_fuzz_reproducer_seed_581646() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "\"ob\"");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "-8");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "2");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "322");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-39");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "\"sivgqf\"");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "7");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "-14");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "269");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "371.3");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "-202.535");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "422.74");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "426.4483");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "9");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "\"sT3\"");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "340.7");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=C3");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=PRODUCT((-22 * E1), -48)");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=ROUNDUP(43, 1)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=IF((D2 > E1), AVERAGE(B5, A5), A4)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=UPPER(\"-26\")");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "4");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=-38");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=IF((A5 > ROUND(D4, 1)), LEFT(\"D2\", 1), AVERAGE(D6:D6))");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "35");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "\"Sxz1Rx Q\"");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=30");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=SUM(C7:D7)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "256.936");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=43");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=ABS((D7 * E3))");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=E3");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=SQRT(MAX(B8:D8))");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=-25");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=-47");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=IF((C4 > RIGHT(\"B7\", 3)), E2, C9)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=SQRT(OR(-3 > 0, D1 < 100))");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=LEFT(\"C6\", 4)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "-71");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=D1");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 0));
    println!("Seed 581646 evaluated target CellRef(8, 0): {:?}", target);
    match target {
        ResultData::Float(f) => assert!((f - 1670.85).abs() < 1e-3),
        other => panic!("Expected Float(~1670.85), got {:?}", other),
    }
}
#[test]
fn test_fuzz_reproducer_seed_607909() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "\" a1 2oi\"");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "8");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "49");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "10");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "\"hUZmik3E\"");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-11.1");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "-311.1075");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "33");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "157.41");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "78");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "-60");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "-49");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-34");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "\"M\"");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "70");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-93");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "\"1Yyyi2M\"");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "\"IooXW\"");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "-63");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "424.7");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=IF((-39 > D3), (B5 / B4), OR(A5 > 0, A4 < 100))");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=B5");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=19");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "20");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=AND(LEN(\"C4\") > 0, MIN(D2, E4) < 100)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=(30 * (D6 / E6))");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "322");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=(CONCATENATE(\"-27\", \"1\") - OR(D6 > 0, C2 < 100))");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "\"DIbqt\"");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "-74");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=(B6 / A5)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=ROUND(26, 1)");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=40");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=D6");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=INT(27)");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=ROUNDDOWN(E3, 0)");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "30");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=8");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=(IF((-23 > C1), C8, -46) / SQRT(-36))");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=INT(A6)");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "47");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=(LEFT(\"E6\", 4) + E6)");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=ROUNDDOWN(4, 1)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=D1");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=A3");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 0));
    println!("Seed 607909 evaluated target CellRef(6, 0) A7: {:?}", target);
    for r in 0..10 {
        for c in 0..5 {
            let res = sheet.get_result_data(&CellRef::new(r, c));
            if matches!(res, ResultData::Error(_)) {
                let col_let = (b'A' + c as u8) as char;
                println!("  Error cell {}{}: {:?}", col_let, r + 1, res);
            }
        }
    }
}
#[test]
fn test_fuzz_reproducer_seed_761871() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "25");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "5");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "24");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "-17");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-27");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "28");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "-93.3104");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "74");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "-71.2");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "9");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "-41");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "16");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "-339");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "7");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-90");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "-443.5");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "285.242");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "16");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "24");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "-19");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "480.2611");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=PRODUCT(CONCATENATE(\"E3\", \"C3\"), PRODUCT(B5:E5))");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=OR((E4 + 37) > 0, ROUNDDOWN(E4, 2) < 100)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=ABS(MIN(C4:C4))");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=SUM(B3:B4)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=MAX(A4:D5)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=(-3 * (E4 * A5))");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=IF((MAX(B2:B4) > (C3 - A3)), D6, CONCATENATE(\"A3\", \"C1\"))");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=D4");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=A5");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=D5");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=C4");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=MIN(RIGHT(\"-33\", 3), E4)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=B7");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=RIGHT(\"OR(D5 > 0, -28 < 100)\", 5)");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=E8");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=INT(-36)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=B4");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=LEFT(\"ABS(26)\", 2)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=B4");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=(ROUNDUP(C7, 0) - (D1 + B6))");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=ROUNDUP(PRODUCT(B6, -6), 0)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=MIN(D5:D9)");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 2));
    println!("Seed 761871 evaluated target CellRef(9, 2): {:?}", target);
    match target { ResultData::Float(f) => assert_eq!(f, 66.0), other => panic!("Expected Float(66.0), got {:?}", other) }
}
#[test]
fn test_fuzz_reproducer_seed_119147() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "\"iPf\"");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "327.95");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-9");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "-343.8");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-81");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "6");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "-59");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "\"ixpH\"");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "51");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "\"yNu\"");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "\"X\"");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-31");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "41");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "-466.47");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "437");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "7");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "50");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "-84");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "-90");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "-95");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "200.2");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=OR(INT(D3) > 0, (6 * A3) < 100)");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=IF((ABS(E4) > OR(D1 > 0, -31 < 100)), E5, ROUNDDOWN(E5, 0))");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=B1");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=A5");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=AND(MIN(E3:E6) > 0, UPPER(\"E2\") < 100)");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=LEFT(\"PRODUCT(B2:E6)\", 2)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=D2");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=-35");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=C5");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=ABS((39 + A7))");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=(24 ^ -44)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=IF((C6 > AVERAGE(E2:E5)), 7, D2)");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "-244.157");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=10");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "-103.1");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=ROUND(E2, 2)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=AVERAGE(PRODUCT(D8, B3), C5)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=ROUND(PRODUCT(20, D4), 2)");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=D8");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=SUM(OR(D9 > 0, A1 < 100), ROUND(-9, 0))");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=(INT(D7) + CONCATENATE(\"D9\", \"C9\"))");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=AND(UPPER(\"E9\") > 0, (E7 ^ A7) < 100)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=E9");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 3));
    println!("Seed 119147 evaluated target CellRef(9, 3): {:?}", target);
    match target { ResultData::Boolean(b) => assert_eq!(b, true), other => panic!("Expected Boolean(true), got {:?}", other) }
}
#[test]
fn test_fuzz_reproducer_seed_549258() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "14");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "-100");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "7");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "438.807");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "-310.23");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "2");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "\"JQxaopYx\"");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "9");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "\"K\"");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "-5");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "93");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "-195.001");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "46");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "81");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "\"EmQomuP\"");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "\"b\"");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "-198.6326");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "\"QPDfo\"");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "45");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "-11");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=ABS(B5)");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=SQRT(B4)");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "99");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=IF((2 > INT(18)), OR(C5 > 0, -34 < 100), A4)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=(INT(D1) / B5)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=LEN(\"IF((A2 > E1), A2, -35)\")");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=MAX(B4:E4)");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=MAX(D5, MAX(C3, C1))");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "492");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=ROUND(E6, 2)");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=D6");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=(D6 / SUM(C3, C3))");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=-28");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=ROUNDUP(IF((D7 > E2), A6, B3), 0)");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=A6");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=MIN(B5, (D2 * 10))");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=(LEN(\"E5\") / -15)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=A5");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=UPPER(\"(-10 * D2)\")");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=ROUNDDOWN(MIN(C1:D4), 0)");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=MIN(C2:D3)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=50");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=E7");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=E1");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=IF((IF((E7 > 8), E9, -33) > SQRT(D5)), 20, (A5 * 27))");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 1));
    println!("Seed 549258 evaluated target CellRef(7, 1): {:?}", target);
}
#[test]
fn test_fuzz_reproducer_seed_584686() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "-499.734");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "-35");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "30");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "5");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-65");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "\"mARAP\"");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "7");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "448.29");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "-115.1434");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "\"TDzexvu\"");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "6");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "\"uddSCurn\"");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "98");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "49");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "\"YQ1G\"");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "-5");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "\"1lrX\"");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "19");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "\"C\"");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=((E5 / A2) ^ IF((A3 > C5), -48, E4))");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=ROUND(AND(-5 > 0, -40 < 100), 2)");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=D2");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "22");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=AND(UPPER(\"D4\") > 0, B5 < 100)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=E5");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=PRODUCT(D2:E3)");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=(MAX(B5:D6) - SUM(C1:E1))");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=D5");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=ROUND(B6, 0)");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=SQRT(LOWER(\"A6\"))");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=AND((46 + A6) > 0, 28 < 100)");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=(MAX(E3:E7) / AND(A6 > 0, E2 < 100))");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=SUM(C4:D4)");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=A4");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=C8");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=ABS(B8)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=AND(-18 > 0, INT(8) < 100)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=-2");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=(B5 - E9)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=(INT(-10) / RIGHT(\"C6\", 3))");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=AVERAGE(A6:A6)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=C3");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=RIGHT(\"C8\", 2)");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    println!("Seed 584686 evaluated target CellRef(9, 0) A10: {:?}", target);
    for r in 0..10 {
        for c in 0..5 {
            let res = sheet.get_result_data(&CellRef::new(r, c));
            if matches!(res, ResultData::Error(_)) {
                let col_let = (b'A' + c as u8) as char;
                println!("  Error cell {}{}: {:?}", col_let, r + 1, res);
            }
        }
    }
}
#[test]
fn test_fuzz_reproducer_seed_140361() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "-74");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "-14");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "-51");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-73.3");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "\"i2blI\"");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "\"IvctY\"");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "-76");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "-27");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "-55");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "21");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "42");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "83.538");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "-31");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "-27");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "-45.363");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "-8");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "74.40000000000001");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "-155.7");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "-53.258");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=SUM(D4, UPPER(\"-35\"))");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=IF((B1 > OR(-19 > 0, B1 < 100)), PRODUCT(D1:D3), LOWER(\"A1\"))");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=OR(CONCATENATE(\"1\", \"13\") > 0, LOWER(\"A6\") < 100)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=(-14 / (E4 * 10))");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=MAX(INT(C6), PRODUCT(A4, D5))");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=ROUND(C6, 2)");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=IF((B5 > LOWER(\"B2\")), IF((B4 > -39), A2, C6), AVERAGE(A1, B5))");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=IF((RIGHT(\"E4\", 4) > B2), -9, C7)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "63");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=24");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=A2");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=-31");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=MIN((16 * C3), (A5 / B1))");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=IF((C4 > D1), CONCATENATE(\"E8\", \"47\"), 33)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=B3");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "5");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=A1");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=SQRT(ROUNDDOWN(B4, 2))");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=(E7 + 30)");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=A9");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "180.58");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=(-13 * (C5 / B6))");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 1));
    println!("Seed 140361 evaluated target CellRef(9, 1): {:?}", target);
    match target { ResultData::Float(f) => assert!((f - 30.2).abs() < 1e-3), other => panic!("Expected Float(~30.2), got {:?}", other) }
}
#[test]
fn test_fuzz_reproducer_seed_957558() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "-60");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-242");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "22");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "\"Qo2z\"");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "-22");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "50");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "\"e \"");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "40");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "56");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-56");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "-74");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "-20");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-220.4508");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "-33");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "\" To\"");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "\"Tk1\"");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "\"aNlj\"");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=ABS(C2)");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=(IF((A4 > E2), A5, 40) ^ IF((C1 > -23), C1, C5))");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=E4");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=INT(IF((C4 > C4), B4, E3))");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "1");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=(D3 - D6)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=(SQRT(-40) ^ MIN(D2:E3))");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "\"o3rTx\"");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=(A6 + D6)");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=D5");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=IF((-39 > LEN(\"C7\")), A7, A4)");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "6");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "-15");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=MIN(LEN(\"E6\"), -20)");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=C3");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=UPPER(\"C3\")");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=(-26 ^ C5)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=17");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=E3");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=A2");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=SQRT(OR(9 > 0, C8 < 100))");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=UPPER(\"IF((E7 > D8), 6, -29)\")");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=IF((A7 > (E2 + E8)), B5, C9)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=E2");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 0));
    println!("Seed 957558 evaluated target CellRef(6, 0): {:?}", target);
}
#[test]
fn test_fuzz_reproducer_seed_814112() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "-68");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "\"KB\"");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "\"XE\"");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "352.42");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "451.693");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-39.3853");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "29.6");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "34.75");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "4");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "94");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "-6");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-99");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "\"W\"");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "\"L\"");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "11");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "94");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "49");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "363.5");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "\"zdIPRq\"");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "159.09");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=AND(PRODUCT(B1:C2) > 0, 20 < 100)");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=D4");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "14");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=C5");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "-164.7238");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=A4");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=ABS(MIN(A2, 24))");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=B1");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=B2");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=SQRT(SUM(-10, C6))");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=(SUM(D3:D3) + ROUNDUP(B7, 0))");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=D1");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "18");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=C3");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=SUM(A6:C6)");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=42");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=PRODUCT(7, C6)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=LEN(\"E2\")");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=8");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=-6");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=-43");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=(C9 * UPPER(\"24\"))");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=IF((A5 > ROUND(42, 0)), (5 + 43), OR(-2 > 0, E5 < 100))");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=12");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=D9");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 1));
    println!("Seed 814112 evaluated target CellRef(9, 1): {:?}", target);
    match target {
        ResultData::Float(f) => assert_eq!(f, 48.0),
        other => panic!("Expected Float(48.0), got {:?}", other),
    }
}
#[test]
fn test_fuzz_reproducer_seed_377746() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "91");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "\"d\"");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "60");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "331.25");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "51");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "22.77");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "-34");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "-15");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "-94");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "2");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "-22");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "155.8");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "-320.5996");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "\"V\"");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-460.0053");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "-83");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "19.159");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "-55");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "1");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "\"yHs lFZi\"");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "\"nZ\"");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=LOWER(\"OR(D2 > 0, A2 < 100)\")");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=D3");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "195.5917");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=IF(((C2 * 16) > INT(-23)), ROUND(A1, 2), (B4 - B5))");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=13");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=OR(RIGHT(\"B3\", 1) > 0, SQRT(-19) < 100)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=ROUNDUP(ROUNDUP(B5, 0), 2)");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=C2");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=-7");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=-26");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=(LEN(\"A2\") / A2)");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=ABS((D7 + -26))");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=B5");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=D7");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=D1");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=46");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=(A3 - LEFT(\"37\", 2))");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=LEN(\"B8\")");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=-11");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "85");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=C8");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=IF(((D5 / 44) > (A6 + C9)), PRODUCT(D8:D8), CONCATENATE(\"E6\", \"E1\"))");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "\"cmWb\"");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=IF((ROUNDUP(D6, 2) > D4), PRODUCT(A9:D9), AND(A4 > 0, -10 < 100))");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 2));
    println!("Seed 377746 evaluated target CellRef(7, 2): {:?}", target);
    for r in 0..10 {
        for c in 0..5 {
            let res = sheet.get_result_data(&CellRef::new(r, c));
            if matches!(res, ResultData::Error(_)) {
                let col_let = (b"A"[0] + c as u8) as char;
                println!("  Error cell {}{}: {:?}", col_let, r + 1, res);
            }
        }
    }
}
#[test]
fn test_fuzz_reproducer_seed_691297() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "149.7872");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "258.543");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "12");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "\"3qMv1d\"");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "444.9831");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "\"z\"");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "\"LLw2Eeu\"");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "\"BQQ\"");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "\"phJdTzJ\"");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "-94");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "59.1905");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "\"qswbs\"");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "482.7");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "7");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "60");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "\"FHL\"");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "80");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=D1");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=B5");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=-26");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=AND(-10 > 0, IF((35 > -23), 36, 49) < 100)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=ROUNDDOWN(-35, 2)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=C3");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=D3");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=ROUNDUP(SUM(E1, C4), 0)");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=IF((D4 > AND(C2 > 0, -33 < 100)), E6, 29)");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=25");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=(A7 * (D5 - 37))");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "-37");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=(ROUNDUP(-14, 2) ^ -1)");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=SQRT(IF((B1 > A4), 33, C3))");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=ROUNDUP((A7 / D8), 2)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=C2");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=(E4 ^ AVERAGE(-19, E4))");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=UPPER(\"(C2 * B4)\")");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=C3");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=(AND(B2 > 0, C8 < 100) + ROUNDUP(A8, 0))");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=B4");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=MIN(A5:D6)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=C3");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 1));
    println!("Seed 691297 evaluated target CellRef(8, 1): {:?}", target);
    match target { ResultData::Float(f) => assert_eq!(f.abs(), 0.0), other => panic!("Expected Float(~0.0), got {:?}", other) }
}
#[test]
fn test_fuzz_reproducer_seed_266270() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "-139.3711");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "74.28100000000001");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "6.793");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-277");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "\"QrVXAE\"");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "83");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "129.4");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "\"WRedVYxM\"");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "427.42");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "-62");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "91");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "83");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "2");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "-45");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "7");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "396");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "-52");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "-64");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=SQRT(A1)");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=C2");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=(B5 + B4)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=(INT(B1) ^ (1 - -8))");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "\"BHED\"");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=((B5 / D1) ^ -47)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=MIN(A6:C6)");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "-265.176");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=IF((RIGHT(\"B4\", 1) > AND(E2 > 0, C5 < 100)), C4, D1)");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=C2");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "34");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=(-42 / (20 - B7))");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=UPPER(\"D3\")");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=E5");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "63");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "25");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=LOWER(\"(B3 - D3)\")");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=ABS(UPPER(\"D5\"))");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=SUM((A1 ^ C3), -22)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=ROUNDDOWN(IF((E1 > -36), A5, C8), 0)");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=32");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=ROUND(C6, 2)");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=B7");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=(C5 + D8)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=((C7 + B3) / E6)");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 1));
    println!("Seed 266270 evaluated target CellRef(9, 1): {:?}", target);
    match target { ResultData::Float(f) => assert_eq!(f, 7.0), other => panic!("Expected Float(7.0), got {:?}", other) }
}
#[test]
fn test_fuzz_reproducer_seed_795051() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "92");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "-3");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "8");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "33");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-318");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "-43");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "\"hxtO3rV\"");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "-477.5");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "78.2");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "96");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "-328.849");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "340.91");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "50");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "-418.3604");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "-46");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "7");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "\"vfH1Q\"");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "2");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "\"eR\"");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=ABS(ROUNDUP(C3, 1))");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=RIGHT(\"(B3 * A2)\", 5)");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=ROUNDUP(CONCATENATE(\"-26\", \"A3\"), 1)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "-62");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=IF((D4 > RIGHT(\"B3\", 3)), LEN(\"E3\"), IF((38 > 44), 0, 36))");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=LEN(\"-13\")");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "68");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=AND(A4 > 0, SUM(-36, A6) < 100)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=INT((B6 / E4))");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=RIGHT(\"D2\", 5)");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=SUM(A6:B6)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "\"virpT\"");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=SUM(D1:E6)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=B6");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=A7");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=E1");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=INT((-30 - E2))");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=AND(B4 > 0, IF((-25 > A4), 35, B3) < 100)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=SUM(IF((E2 > -48), D5, C5), LOWER(\"A4\"))");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=UPPER(\"D8\")");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=18");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=PRODUCT(ABS(E8), IF((-4 > -12), -5, A2))");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=ROUNDDOWN(SQRT(E6), 0)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=(-4 * B1)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=(E7 ^ LOWER(\"E4\"))");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 1));
    println!("Seed 795051 evaluated target CellRef(9, 1): {:?}", target);
}
#[test]
fn test_fuzz_reproducer_seed_515564() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "\"mDV2O\"");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "\"ww\"");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-37");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "-255.2");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "25");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "-272.34");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "3");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "-4");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "84");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "\"2Hlb\"");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "75");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "47");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "220.459");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-8");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "-9");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "47");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "196.03");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "10");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "8");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=(D3 ^ UPPER(\"A3\"))");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=E5");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=(E2 * -34)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "406.9");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=AND(48 > 0, SQRT(-25) < 100)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "-472.176");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=ABS((E6 ^ A1))");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=C5");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=D6");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=ROUND(-4, 1)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=CONCATENATE(\"MIN(D7:E7)\", \"IF((-3 > D6), B1, -16)\")");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=(ROUND(D1, 1) - A6)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=(C2 * -30)");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=MAX(UPPER(\"-50\"), -8)");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=UPPER(\"ROUND(45, 0)\")");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=E2");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=IF((ROUNDUP(B6, 2) > MIN(E3:E7)), C5, E2)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=SQRT(INT(A6))");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=(B4 ^ LEFT(\"D2\", 2))");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=17");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=SUM(LOWER(\"47\"), IF((E3 > C2), B1, D6))");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=E3");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=LEFT(\"SQRT(C5)\", 3)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "-55");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 1));
    println!("Seed 515564 evaluated target CellRef(6, 1): {:?}", target);
}
#[test]
fn test_fuzz_reproducer_seed_107138() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "4");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "-496.2");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "5");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "-38");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "85");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "\"HcOUVBwS\"");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "87");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "\"3c\"");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "187");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "-182.3");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "-370.66");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "402.72");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "-349.5425");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "79");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "2");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "83");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "\"p\"");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "56.82");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "-4");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "60");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "-48");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "-125.224");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=-36");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=(IF((A5 > C3), B3, C1) * ABS(C3))");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=AND(MIN(E1:E4) > 0, ROUNDDOWN(A4, 0) < 100)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "138");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=C5");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=CONCATENATE(\"IF((B4 > C3), B6, C6)\", \"INT(33)\")");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=E3");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=(A4 ^ AVERAGE(B5:B5))");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=CONCATENATE(\"D5\", \"B2\")");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=B6");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=39");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=D6");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=PRODUCT(E7:E7)");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=CONCATENATE(\"E4\", \"AND(C5 > 0, E5 < 100)\")");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=ROUNDDOWN(B8, 0)");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=MAX(B2:C7)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=(B8 + -21)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=SUM(E3:E7)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=ROUND(SUM(B8:C8), 2)");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=AND(C8 > 0, 36 < 100)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=INT(E9)");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=OR(PRODUCT(B2, 2) > 0, LOWER(\"-6\") < 100)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=AND(8 > 0, IF((C1 > 31), B1, E9) < 100)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "39");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 0));
    println!("Seed 107138 evaluated target CellRef(8, 0): {:?}", target);
}
#[test]
fn test_fuzz_reproducer_seed_160523() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "\"uALF eK\"");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "\"XgB\"");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "31");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "79.81480000000001");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "\"3\"");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "\"raaPD\"");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "4");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "-30");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "69");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "-78");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "-11.21");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "72");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-1");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "\"jSNe3xpR\"");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "65");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "\"nUmwj\"");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "10");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "88");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "99");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=INT(ROUNDUP(C1, 0))");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=(D3 ^ IF((D1 > C2), D2, D4))");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=A4");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=A2");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=ROUND(A4, 1)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=RIGHT(\"-3\", 4)");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "\"uWOqN\"");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=PRODUCT(E2:E4)");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=D6");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=ROUNDDOWN(B6, 1)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=50");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "18");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=(C6 + ROUND(-35, 0))");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=MIN(A3, A1)");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=IF(((E8 ^ A3) > D7), (E6 + 4), (D6 / -5))");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=C6");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=-43");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "-95.50839999999999");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=IF((A7 > UPPER(\"29\")), E5, INT(A7))");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=AND(ROUND(-5, 1) > 0, A5 < 100)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=D4");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "\"Y\"");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=-28");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=B8");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 0));
    println!("Seed 160523 evaluated target CellRef(8, 0): {:?}", target);
}
#[test]
fn test_fuzz_reproducer_seed_746635() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "-85");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "85");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "-67");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "65");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "\"NAU\"");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "8");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "-93");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "1");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "-201.948");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "9");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "-245");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "24.4");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "78");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "\"S\"");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "102");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-82");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "\"A\"");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "58");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "-74");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "-21");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "\"1\"");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "399.35");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "264.902");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=IF((A3 > ABS(D3)), B3, CONCATENATE(\"B1\", \"B3\"))");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=B2");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=(-28 + (B4 ^ -48))");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=MAX(C2:D5)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=PRODUCT((C5 / D3), D2)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=(IF((B4 > E6), A5, 15) - (C6 + E3))");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=C4");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=MIN(E4:E6)");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=E4");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=AND(IF((D2 > C3), B4, B7) > 0, ROUNDUP(5, 0) < 100)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=D7");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "39");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=(C2 + IF((B2 > A3), 43, B6))");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=IF((ROUNDDOWN(E7, 2) > E4), ABS(D1), MAX(E4:E4))");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=4");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=5");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "62");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=MAX(C2:D6)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "\"M\"");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=(AND(C3 > 0, C3 < 100) - (E8 / A3))");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=D1");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=ROUNDUP(ABS(C1), 2)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=LEFT(\"C1\", 3)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=43");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    println!("Seed 746635 evaluated target CellRef(9, 0): {:?}", target);
}
#[test]
fn test_fuzz_reproducer_seed_448316() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "-23");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "-90.2");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "43");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "50");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-24");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "74");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "-31");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "85");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "53");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "\"lzFQ1BZ\"");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "-98");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "44");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "6");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "-21");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "89");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "355.279");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "\"Vv\"");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "255");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "\" hU\"");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "79.51000000000001");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "-16");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "-152.82");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "8");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "135.7");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "-54");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=SUM(D3:D5)");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=INT(C3)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=(LOWER(\"B4\") - LEN(\"A1\"))");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=(PRODUCT(7, D3) / AVERAGE(E3:E4))");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=B1");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "33");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=(LEN(\"-33\") / D4)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=B4");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "145");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=(D7 / RIGHT(\"B6\", 1))");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "59");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "223.585");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "-86");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=SQRT(RIGHT(\"D8\", 3))");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=-17");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=ROUND(B5, 1)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=(E6 * -3)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=IF((E4 > (B6 + -22)), (B1 - -29), -47)");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=AVERAGE(CONCATENATE(\"B5\", \"A2\"), 3)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=AVERAGE((E3 * -25), ROUNDDOWN(C1, 2))");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=ROUNDDOWN(B1, 0)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=ROUNDDOWN(IF((-45 > E2), C4, 5), 2)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=OR(B8 > 0, MAX(A2:B8) < 100)");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 1));
    println!("Seed 448316 evaluated target CellRef(7, 1): {:?}", target);
}