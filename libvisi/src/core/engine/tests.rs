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
fn test_fuzz_sqrt_negative_operand_error() {
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
    assert!(
        matches!(target, ResultData::Error(ref e) if e.contains("#NUM!")),
        "Expected #NUM!, got {:?}",
        target
    );
}

#[test]
fn test_fuzz_rounddown_scaled_float_precision() {
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
    match target {
        ResultData::Float(f) => assert!(f.abs() < 1e-6),
        other => panic!("Expected Float(0.0) for A6, got {:?}", other),
    }
}

#[test]
fn test_fuzz_roundup_rounddown_nested_error() {
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
    assert!(
        matches!(target, ResultData::Boolean(true) | ResultData::Error(_)),
        "Expected Boolean(true) or Error for C7, got {:?}",
        target
    );
}

#[test]
fn test_fuzz_division_by_negative_sum() {
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
    match target {
        ResultData::Float(f) => assert_eq!(f, 49.0),
        other => panic!("Expected Float(49.0) for A8, got {:?}", other),
    }
}

#[test]
fn test_fuzz_right_string_addition_coercion() {
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
    match target {
        ResultData::Float(f) => assert_eq!(f, 28.0),
        ResultData::Integer(i) => assert_eq!(i, 28),
        other => panic!("Expected Float(28.0) for B7, got {:?}", other),
    }
}

#[test]
fn test_fuzz_min_range_division_by_zero_error() {
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
    assert!(
        matches!(target, ResultData::Error(ref e) if e.contains("#DIV/0!")),
        "Expected #DIV/0! for B7, got {:?}",
        target
    );
}

#[test]
fn test_fuzz_cell_reference_zero_coercion() {
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
    match target {
        ResultData::Float(f) => assert!(f.abs() < 1e-6),
        other => panic!("Expected Float(0.0) for A9, got {:?}", other),
    }
}

#[test]
fn test_fuzz_lower_negative_number_string_coercion() {
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
    match target {
        ResultData::Float(f) => assert_eq!(f, -2.0),
        other => panic!("Expected Float(-2.0) for B10, got {:?}", other),
    }
}

#[test]
fn test_fuzz_subtraction_division_by_round() {
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
    match target {
        ResultData::Float(f) => assert_eq!(f, -105.0),
        other => panic!("Expected Float(-105.0) for A9, got {:?}", other),
    }
}

#[test]
fn test_fuzz_boolean_dependency_cell_evaluation() {
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
    match target {
        ResultData::Boolean(b) => assert_eq!(b, true),
        other => panic!("Expected Boolean(true) for A10, got {:?}", other),
    }
}

#[test]
fn test_fuzz_and_multiplication_num_error() {
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
    assert!(
        matches!(target, ResultData::Error(ref e) if e.contains("#NUM!")),
        "Expected #NUM! for A10, got {:?}",
        target
    );
}

#[test]
fn test_fuzz_product_nested_float_precision() {
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
    match target {
        ResultData::Float(f) => assert_eq!(f, 50.0),
        other => panic!("Expected Float(50.0) for A8, got {:?}", other),
    }
}

#[test]
fn test_fuzz_constant_literal_cell_evaluation() {
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
    match target {
        ResultData::Float(f) => assert_eq!(f, 1.0),
        other => panic!("Expected Float(1.0) for A8, got {:?}", other),
    }
}

#[test]
fn test_fuzz_nested_math_expression_precision() {
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
    match target {
        ResultData::Float(f) => assert_eq!(f, -43.0),
        other => panic!("Expected Float(-43.0) for B10, got {:?}", other),
    }
}

#[test]
fn test_fuzz_range_min_max_evaluation() {
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
    match target {
        ResultData::Float(f) => assert_eq!(f, 51.0),
        other => panic!("Expected Float(51.0) for E8, got {:?}", other),
    }
}

#[test]
fn test_fuzz_concatenate_if_function_error() {
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
    assert!(
        matches!(target, ResultData::Error(_)),
        "Expected Error for C9, got {:?}",
        target
    );
}

#[test]
fn test_fuzz_division_by_zero_formula_error() {
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
    match target {
        ResultData::Error(_) => {}
        ResultData::Float(f) => assert_eq!(f, -15.0),
        other => panic!("Expected Float(-15.0) for A8, got {:?}", other),
    }
}

#[test]
fn test_fuzz_round_nested_precision() {
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
    match target {
        ResultData::Float(f) => assert_eq!(f, 1.0),
        other => panic!("Expected Float(1.0) for A7, got {:?}", other),
    }
}

#[test]
fn test_fuzz_sqrt_log_range_error() {
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
    assert!(
        matches!(target, ResultData::Error(ref e) if e.contains("#NUM!")),
        "Expected #NUM! for A7, got {:?}",
        target
    );
}

#[test]
fn test_fuzz_product_negative_multipliers() {
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
    match target {
        ResultData::Float(f) => assert_eq!(f, 216.0),
        other => panic!("Expected Float(216.0) for B10, got {:?}", other),
    }
}

#[test]
fn test_fuzz_subtraction_large_range_min() {
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
    match target {
        ResultData::Float(f) => assert_eq!(f, -282.0),
        other => panic!("Expected Float(-282.0) for C7, got {:?}", other),
    }
}

#[test]
fn test_fuzz_negative_constant_subtraction() {
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
    match target {
        ResultData::Float(f) => assert_eq!(f, -80.0),
        other => panic!("Expected Float(-80.0) for B9, got {:?}", other),
    }
}

#[test]
fn test_fuzz_sum_range_negative_values() {
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
    match target {
        ResultData::Float(f) => assert_eq!(f, -52.0),
        other => panic!("Expected Float(-52.0) for C7, got {:?}", other),
    }
}

#[test]
fn test_fuzz_zero_result_division_expression() {
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
    match target {
        ResultData::Float(f) => assert_eq!(f, 0.0),
        other => panic!("Expected Float(0.0) for D9, got {:?}", other),
    }
}

#[test]
fn test_fuzz_sum_product_cell_references() {
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
    match target {
        ResultData::Float(f) => assert_eq!(f, 113.0),
        other => panic!("Expected Float(113.0) for B10, got {:?}", other),
    }
}

#[test]
fn test_fuzz_negative_integer_range_sum() {
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
    match target {
        ResultData::Float(f) => assert_eq!(f, -220.0),
        other => panic!("Expected Float(-220.0) for C10, got {:?}", other),
    }
}

#[test]
fn test_fuzz_nested_min_max_evaluation() {
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
    match target {
        ResultData::Float(f) => assert_eq!(f, -8.0),
        other => panic!("Expected Float(-8.0) for A10, got {:?}", other),
    }
}

#[test]
fn test_fuzz_string_coercion_expected_number_error() {
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
    assert!(
        matches!(target, ResultData::Error(_)),
        "Expected Error for D10, got {:?}",
        target
    );
}

#[test]
fn test_fuzz_power_cell_references() {
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
    let target = sheet.get_result_data(&CellRef::new(8, 1));

}

#[test]
fn test_fuzz_roundup_power_expression() {
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
    let target = sheet.get_result_data(&CellRef::new(9, 1));

}

#[test]
fn test_fuzz_roundup_sum_expression() {
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
    let target = sheet.get_result_data(&CellRef::new(6, 2));

}

#[test]
fn test_fuzz_and_if_comparison_evaluation() {
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
    let target = sheet.get_result_data(&CellRef::new(9, 1));

}

#[test]
fn test_fuzz_max_int_multiplication() {
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
    let target = sheet.get_result_data(&CellRef::new(8, 2));

}

#[test]
fn test_fuzz_average_right_string_argument() {
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
    let target = sheet.get_result_data(&CellRef::new(6, 4));

}

#[test]
fn test_fuzz_if_multiplication_comparison() {
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
    let target = sheet.get_result_data(&CellRef::new(8, 3));

}

#[test]
fn test_fuzz_addition_cell_references() {
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
    let target = sheet.get_result_data(&CellRef::new(8, 0));

}

#[test]
fn test_fuzz_power_min_expression() {
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
    let target = sheet.get_result_data(&CellRef::new(7, 0));

}

#[test]
fn test_fuzz_addition_min_negative() {
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
    let target = sheet.get_result_data(&CellRef::new(9, 1));

}

#[test]
fn test_fuzz_multiplication_cell_references() {
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
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    match target {
        ResultData::Float(f) => assert_eq!(f, 3738.0),
        other => panic!("Expected Float(3738.0) for A10, got {:?}", other),
    }

}

#[test]
fn test_fuzz_roundup_sum_precision() {
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
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    match target {
        ResultData::Float(f) => assert!((f - 318.61).abs() < 1e-3, "Expected ~318.61 for A10, got {}", f),
        other => panic!("Expected Float(~318.61) for A10, got {:?}", other),
    }

}

#[test]
fn test_fuzz_average_if_branch_evaluation() {
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
    let target = sheet.get_result_data(&CellRef::new(8, 3));
    match target {
        ResultData::Float(f) => assert_eq!(f, -30000000.0),
        other => panic!("Expected Float(-30000000.0) for D9, got {:?}", other),
    }

}

#[test]
fn test_fuzz_abs_cell_reference() {
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
    for r in 0..10 {
        for c in 0..5 {
            let res = sheet.get_result_data(&CellRef::new(r, c));
            if matches!(res, ResultData::Error(_)) {
                let col_let = (b'A' + c as u8) as char;
            }
        }
    }
}

#[test]
fn test_fuzz_sqrt_or_evaluation() {
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
    let target2 = sheet.get_result_data(&CellRef::new(8, 3));
    for r in 0..10 {
        for c in 0..5 {
            let res = sheet.get_result_data(&CellRef::new(r, c));
            if matches!(res, ResultData::Error(_)) {
                let col_let = (b'A' + c as u8) as char;
            }
        }
    }
}

#[test]
fn test_fuzz_if_max_right_string_branch() {
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
fn test_fuzz_sum_or_addition_coercion() {
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
    match target {
        ResultData::Float(f) => assert_eq!(f, 51.0),
        other => panic!("Expected Float(51.0), got {:?}", other),
    }
}

#[test]
fn test_fuzz_abs_multiplication_precision() {
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
    match target {
        ResultData::Float(f) => assert!((f - 1670.85).abs() < 1e-3),
        other => panic!("Expected Float(~1670.85), got {:?}", other),
    }
}

#[test]
fn test_fuzz_multiplication_division_expression() {
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
    for r in 0..10 {
        for c in 0..5 {
            let res = sheet.get_result_data(&CellRef::new(r, c));
            if matches!(res, ResultData::Error(_)) {
                let col_let = (b'A' + c as u8) as char;
            }
        }
    }
}

#[test]
fn test_fuzz_roundup_subtraction_boolean_addition() {
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
    match target { ResultData::Float(f) => assert_eq!(f, 66.0), other => panic!("Expected Float(66.0), got {:?}", other) }
}

#[test]
fn test_fuzz_and_upper_power_comparison() {
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
    match target { ResultData::Boolean(b) => assert_eq!(b, true), other => panic!("Expected Boolean(true), got {:?}", other) }
}

#[test]
fn test_fuzz_division_by_sum() {
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
}

#[test]
fn test_fuzz_subtraction_string_formula_cell() {
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
    for r in 0..10 {
        for c in 0..5 {
            let res = sheet.get_result_data(&CellRef::new(r, c));
            if matches!(res, ResultData::Error(_)) {
                let col_let = (b'A' + c as u8) as char;
            }
        }
    }
}

#[test]
fn test_fuzz_addition_formula_constant() {
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
    match target { ResultData::Float(f) => assert!((f - 30.2).abs() < 1e-3), other => panic!("Expected Float(~30.2), got {:?}", other) }
}

#[test]
fn test_fuzz_subtraction_cell_references() {
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
}

#[test]
fn test_fuzz_multiplication_upper_string_number() {
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
    match target {
        ResultData::Float(f) => assert_eq!(f, 48.0),
        other => panic!("Expected Float(48.0), got {:?}", other),
    }
}

#[test]
fn test_fuzz_abs_subtraction_expression() {
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
    for r in 0..10 {
        for c in 0..5 {
            let res = sheet.get_result_data(&CellRef::new(r, c));
            if matches!(res, ResultData::Error(_)) {
                let col_let = (b"A"[0] + c as u8) as char;
            }
        }
    }
}

#[test]
fn test_fuzz_roundup_division_negative_zero() {
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
    match target { ResultData::Float(f) => assert_eq!(f.abs(), 0.0), other => panic!("Expected Float(~0.0), got {:?}", other) }
}

#[test]
fn test_fuzz_round_cell_reference() {
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
    match target { ResultData::Float(f) => assert_eq!(f, 7.0), other => panic!("Expected Float(7.0), got {:?}", other) }
}

#[test]
fn test_fuzz_product_abs_if_branch() {
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
}

#[test]
fn test_fuzz_abs_power_negative_base_error() {
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
}

#[test]
fn test_fuzz_rounddown_cell_reference() {
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
}

#[test]
fn test_fuzz_if_power_comparison_branches() {
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
}

#[test]
fn test_fuzz_subtraction_and_comparison_division() {
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
}

#[test]
fn test_fuzz_division_by_right_string() {
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
}

#[test]
fn test_fuzz_round_int_cell_reference() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "-44");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "64");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-89");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "8");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "-39");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "\" MHdPCGd\"");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "-67");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "40");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "-277.6");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "206.81");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "\"EPAIFa1j\"");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "17");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "-48");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "-210.35");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "-99");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "-22");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "-322.5");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=E5");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=INT(E1)");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=SQRT(RIGHT(\"B2\", 4))");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=MAX(LEN(\"B2\"), D1)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=D1");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=E5");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=AND(C2 > 0, (19 ^ B3) < 100)");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=-5");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=AND(OR(-1 > 0, E6 < 100) > 0, IF((A6 > C3), E6, C5) < 100)");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "-29");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=AVERAGE(E1, E2)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=(MIN(A2, 26) + LEN(\"A2\"))");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=D5");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=D6");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=ROUND(PRODUCT(A5:E7), 2)");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=LEFT(\"ROUNDDOWN(-46, 2)\", 5)");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=ROUNDUP(E2, 2)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=(D1 * CONCATENATE(\"19\", \"38\"))");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=ABS(E4)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=((-21 ^ A3) ^ (D3 ^ -12))");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=ROUND(B6, 0)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=-33");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=-1");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=MIN(C7:C9)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "12.395");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    match target {
        ResultData::Float(f) => assert_eq!(f, 64.0),
        other => panic!("Expected Float(64.0), got {:?}", other),
    }
}

#[test]
fn test_fuzz_or_zero_power_zero_num_error() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "499.8781");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "75");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "-325.2142");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-69");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "20");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "\"F if\"");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "-74");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "7");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "-42.681");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "224");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "\"LjESYnsO\"");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "\"IfnR\"");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "301.17");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "72");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "88");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "-95");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=UPPER(\"AVERAGE(D3:E3)\")");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=-24");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=OR(B5 > 0, (B2 ^ E3) < 100)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=LEFT(\"OR(16 > 0, C5 < 100)\", 4)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=AND((C5 ^ C1) > 0, A3 < 100)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=LEN(\"D5\")");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=49");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=-3");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=SUM(D6:D6)");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "\"uRm3\"");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=MIN(E2:E3)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=(LOWER(\"C3\") / A5)");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=ABS(E5)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=SUM(D3:E3)");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=D6");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=MAX(B2:B5)");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=8");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=C1");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=(MAX(E7, 40) * ROUND(-34, 0))");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=ROUND(D3, 2)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=ROUNDUP(AND(-23 > 0, D6 < 100), 0)");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=B5");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=AND((21 / E3) > 0, C5 < 100)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=INT(INT(B3))");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(5, 2));
    match target { ResultData::Boolean(b) => assert_eq!(b, true), ResultData::Error(ref e) => assert!(e.contains("#NUM!")), other => panic!("Expected Boolean(true) or #NUM!, got {:?}", other) }
}

#[test]
fn test_fuzz_if_max_int_num_error() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "-308.27");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "1");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "\"QhQ\"");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "-74");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "7");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "-20");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "-244.9764");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "-2");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "101.48");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "67");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-85.1962");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "-17");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "-446.7525");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "-10");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-387.5653");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "29");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "83.035");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "-27");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "-60");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "\"RvfkX\"");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "-18");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=IF((-28 > SUM(E5:E5)), (B2 * 46), IF((37 > D2), C4, A1))");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=IF((B1 > B3), OR(A4 > 0, -39 < 100), SQRT(C3))");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=(PRODUCT(A1:D4) / (D4 + A1))");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=A5");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=ROUNDDOWN(D2, 0)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=OR((B2 ^ B3) > 0, SUM(C5, 19) < 100)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=A1");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=AND(11 > 0, ABS(E3) < 100)");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=AVERAGE(B5:E6)");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=AND(C7 > 0, ABS(C3) < 100)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=OR(-17 > 0, PRODUCT(45, 15) < 100)");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=AND(LEFT(\"-40\", 2) > 0, SUM(3, C6) < 100)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=-20");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=39");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=IF((MAX(4, E6) > INT(E7)), (E4 - A7), (41 / E6))");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=(CONCATENATE(\"C2\", \"A5\") / -32)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=AVERAGE(A6:B6)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=IF((SUM(D7, D7) > A4), ROUNDUP(E6, 0), (E3 + A6))");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=ABS(IF((9 > B6), E3, E4))");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=C8");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=MAX(D1:D6)");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=B3");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=((-49 ^ B7) * -42)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=IF((D3 > 4), A3, E4)");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 0));
    match target { ResultData::Float(f) => assert!((f - 82.035).abs() < 1e-3), ResultData::Error(ref e) => assert!(e.contains("#NUM!")), other => panic!("Expected Float(82.035) or #NUM!, got {:?}", other) }
}

#[test]
fn test_fuzz_round_single_digit_cell() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "-32");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-39");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "87");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "-13.52");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "89.43000000000001");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "94");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "-59");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "27");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "84");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-482.5904");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "2");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "\"Q\"");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-30");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "484.45");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "24");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "7");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "5");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "11");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=ROUND(INT(C3), 0)");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=B3");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=((-9 + C1) - LOWER(\"12\"))");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=OR(SUM(A3:C4) > 0, OR(A3 > 0, -13 < 100) < 100)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=-12");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=AND(AND(C5 > 0, E5 < 100) > 0, AND(13 > 0, 26 < 100) < 100)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "\"iUZS\"");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=(IF((42 > -35), E4, E4) ^ ABS(B2))");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=RIGHT(\"PRODUCT(E3, E4)\", 2)");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=MIN(A4:B4)");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=(D1 ^ 50)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "88");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=C5");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=(ROUNDUP(C5, 0) ^ E4)");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=LEN(\"INT(D1)\")");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=ROUNDDOWN(4, 1)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "280.31");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=B6");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=ROUND(B9, 1)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=UPPER(\"SUM(-11, B8)\")");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "211.044");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "\"dl\"");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "64");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
}

#[test]
fn test_fuzz_abs_min_range_boolean() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "-46");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "-38");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-340.83");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "10");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "\"JcSVrdA\"");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "-65");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "\"Il\"");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "-221.6");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "-110.62");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-19.28");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "62");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "294.2");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "31.55");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "-71");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "\"3hW\"");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "31");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=ROUNDUP(MIN(C2:C2), 1)");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=((E4 / B4) - C4)");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=OR(ABS(B2) > 0, D4 < 100)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=ABS(D2)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=B2");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=UPPER(\"-44\")");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=E5");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=IF((IF((B2 > -34), B4, C6) > A1), SQRT(-5), LEFT(\"15\", 4))");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=MIN(A4:A6)");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=20");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=ABS(MIN(C5:C5))");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=UPPER(\"ROUNDDOWN(A6, 1)\")");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=B5");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=MAX(A6:D7)");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "-64");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "145");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=(14 - -15)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=PRODUCT(B1:C7)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=AVERAGE(D6:E7)");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=ABS(D7)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=SUM(C6:D8)");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=LEFT(\"E1\", 5)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=E5");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=MIN(C4:C8)");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    println!("Seed 25929 evaluated target CellRef(9, 0) A10: {:?}", target);
    let d7 = sheet.get_result_data(&CellRef::new(6, 3));
    println!("Seed 25929 D7: {:?}", d7);
    let a6 = sheet.get_result_data(&CellRef::new(5, 0));
    println!("Seed 25929 A6: {:?}", a6);
}
#[test]
fn test_fuzz_left_operand_string_value_error_precedence() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "-391.0356");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "\"1Jc\"");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "\"rCSFXgC\"");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "27");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "179");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "-314.7");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "64");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "-51");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "-427.6");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "-261.33");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "\"e\"");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "\"2OtmmtAT\"");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "\"CEpp\"");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "21");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "75.179");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "-279.178");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "-205");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "-227.2985");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "414.49");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=14");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "-85");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=(B5 / IF((-14 > D1), E1, E4))");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=(D3 * (49 / D1))");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=PRODUCT(A1:A4)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=CONCATENATE(\"E5\", \"SUM(C2:C6)\")");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=UPPER(\"ABS(A5)\")");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=AND(D2 > 0, (-40 - E3) < 100)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=AVERAGE(B6:E6)");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=(D1 + 6)");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=IF(((C4 + A1) > A4), -5, B6)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=LEFT(\"SQRT(A2)\", 5)");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "-27");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=ROUNDUP(B1, 2)");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=B5");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=(-15 ^ AND(E3 > 0, A7 < 100))");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=SQRT(ROUND(C5, 1))");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=E4");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=B5");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=A7");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=(ROUNDDOWN(C2, 2) * ROUNDUP(E9, 0))");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "374");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=IF((-35 > AND(D9 > 0, C4 < 100)), -40, C5)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=A4");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=INT((36 - C7))");
    sheet.commit(None).unwrap();
    let target_d6 = sheet.get_result_data(&CellRef::new(5, 3));
    println!("Seed 316841 evaluated target CellRef(5, 3) D6: {:?}", target_d6);
    match target_d6 {
        ResultData::Error(ref e) => assert!(e.contains("#DIV/0!") || e.contains("#VALUE!"), "Expected #DIV/0! or #VALUE! for D6, got {:?}", target_d6),
        other => panic!("Expected Error for D6, got {:?}", other),
    }

    let target_d7 = sheet.get_result_data(&CellRef::new(6, 3));
    println!("Seed 316841 evaluated target CellRef(6, 3) D7: {:?}", target_d7);
    match target_d7 {
        ResultData::Error(ref e) => assert!(e.contains("#DIV/0!") || e.contains("#VALUE!"), "Expected #DIV/0! or #VALUE! for D7, got {:?}", target_d7),
        other => panic!("Expected Error for D7, got {:?}", other),
    }
}
#[test]
fn test_fuzz_roundup_if_branch_evaluation() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "-26");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "-163");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-156");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "\"qy2RBN\"");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "40");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "166.187");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "-100");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "\"Df1LC\"");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "-29");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "\"zggq\"");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "10");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "-381.4");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "20");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "-53");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "177.04");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "108.5");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "\"gGKd GKH\"");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=43");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=OR(CONCATENATE(\"A4\", \"-33\") > 0, SUM(B2:E4) < 100)");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=47");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "-48");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=4");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=AVERAGE(C6:C6)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=INT(C2)");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=AND(ABS(E4) > 0, (-23 - A6) < 100)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=(AVERAGE(A2:D2) + D3)");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=SQRT(ROUND(E2, 2))");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "37");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=INT(ABS(B2))");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=D6");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=D5");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=B3");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=MIN(LEFT(\"A5\", 2), B7)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=SUM(D7, MAX(D4:E7))");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=ROUND(IF((A8 > 23), E6, D4), 2)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=B8");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=MAX((44 ^ E6), AND(E9 > 0, B1 < 100))");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=(ABS(E5) + PRODUCT(C7:C8))");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=ROUNDUP(ABS(E8), 1)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=A2");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=LOWER(\"D2\")");
    sheet.commit(None).unwrap();
    let target_c10 = sheet.get_result_data(&CellRef::new(9, 2));
    println!("Seed 892758 C10: {:?}", target_c10);
    let target_d9 = sheet.get_result_data(&CellRef::new(8, 3));
    println!("Seed 892758 D9: {:?}", target_d9);
}
#[test]
fn test_fuzz_if_power_overflow_num_error() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "-3");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "21.488");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "\"1X1OB\"");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-247.996");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "\"XNfFdxbI\"");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "4");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "\"O\"");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "\"VjE\"");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "9");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "-39");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-76");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "127.738");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "-105.1");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "35");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "372.504");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "133.7");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "40");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "402.9274");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=PRODUCT(A4:A5)");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=E1");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=OR(ROUND(B1, 1) > 0, CONCATENATE(\"-18\", \"47\") < 100)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=-22");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=AVERAGE(C1:D4)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=D1");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=(ROUNDDOWN(E5, 2) + (C2 * B3))");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "215.33");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=B2");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=IF(((A6 / D1) > AVERAGE(D3, C1)), D5, AVERAGE(C3:C4))");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=ROUNDUP((-22 - A5), 2)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=-3");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "4");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=UPPER(\"OR(48 > 0, B1 < 100)\")");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=CONCATENATE(\"E7\", \"E5\")");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=A1");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=A1");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "\"ZDYF1\"");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=(IF((D1 > A8), -12, C9) ^ ROUNDDOWN(B7, 0))");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=AVERAGE(C4, OR(D4 > 0, C8 < 100))");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=MIN(4, (-18 + -4))");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=OR(E2 > 0, C4 < 100)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=SUM(A3:C4)");
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
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "61");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "1");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-45");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "-36");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "-3.5449");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "2");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "-49");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "-11");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "-64");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "-258.9581");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "\"Ah\"");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "40");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "\"nkQDAkX\"");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "54");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "8");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "6");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "-54");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "18");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "\"UJrTlo3\"");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "470.7");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "128");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=SQRT(ABS(-3))");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=-8");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=7");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "224");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=ROUNDUP(ABS(D4), 1)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=IF((MAX(E5:E5) > AND(B4 > 0, E6 < 100)), AVERAGE(B6:B6), IF((E3 > E4), D6, B4))");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=C5");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=(AVERAGE(E1, A1) + ABS(-20))");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=(IF((25 > C6), A3, 18) * C6)");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=UPPER(\"B1\")");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=-5");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=(INT(D2) + IF((E4 > E3), A1, E2))");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=(27 / (-5 - -10))");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "-94");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=IF((CONCATENATE(\"D4\", \"3\") > LEN(\"E3\")), -16, SUM(A4, B7))");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=C8");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=AVERAGE(C5:E8)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=B4");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=ROUNDUP(SQRT(4), 1)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=MAX((D6 / C5), D2)");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=ROUND(E9, 0)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=B6");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=SUM(E3:E9)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=-39");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=SQRT(IF((-23 > E2), B1, A3))");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    println!("Seed 48394 evaluated target CellRef(9, 0): {:?}", target);
    match target { ResultData::Float(f) => assert_eq!(f, 2.0), other => panic!("Expected Float(2.0), got {:?}", other) }
}
#[test]
fn test_fuzz_if_sqrt_nested_branch() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "6");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "\"mdrrGl\"");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "290.85");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "-52");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "\"ae\"");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "-72");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "\"jib\"");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "34");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "12");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "\"C\"");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "29");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "379.061");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "\" Vn\"");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-101.701");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "2");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "-90");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "-81");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "\"YJPn\"");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "2");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "67");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=(IF((D2 > D2), D5, E3) / ABS(-38))");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=11");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=D3");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=OR((B4 - 22) > 0, C3 < 100)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=PRODUCT(D4, D3)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=(MIN(B4, B1) ^ -1)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=B3");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=SQRT(D2)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=ROUNDUP(-10, 0)");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=LEFT(\"(D2 + D2)\", 5)");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=UPPER(\"-16\")");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=E6");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=-38");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=MAX(B1:E7)");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=UPPER(\"MIN(C4, -37)\")");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=PRODUCT(D7:D7)");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=IF((C1 > (-11 * E2)), (E4 + -13), LEFT(\"A2\", 4))");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=IF((OR(B7 > 0, D2 < 100) > (E2 + A6)), ABS(C4), LOWER(\"B8\"))");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=AND(B7 > 0, UPPER(\"D7\") < 100)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=D5");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=IF((E3 > SQRT(D9)), IF((-19 > -33), 45, 13), LEN(\"E5\"))");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=(IF((C7 > E5), -43, C7) + E9)");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=A9");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "88");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=A7");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    println!("Seed 369614 evaluated target CellRef(9, 0): {:?}", target);
    match target { ResultData::Float(f) => assert_eq!(f, 45.0), other => panic!("Expected Float(45.0), got {:?}", other) }
}
#[test]
fn test_fuzz_rounddown_power_c3() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "\" NBLpaTv\"");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "36.56");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "\"DrEszH\"");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "31");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "7");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "-435.2");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "-41");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "-138.433");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "41");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "\"u\"");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "\"wIxVI\"");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "\"HAGiDJE\"");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-395");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "-58");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "\"m\"");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "\"Uc\"");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=LEFT(\"ROUND(D4, 0)\", 1)");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=UPPER(\"38\")");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=(ROUNDUP(B3, 2) + UPPER(\"D2\"))");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=D5");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=E2");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=ROUND(MAX(A3, E3), 2)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=(C4 * CONCATENATE(\"E6\", \"39\"))");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=D6");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=42");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=-46");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=IF((B6 > INT(D3)), ROUNDDOWN(B1, 1), B5)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=D3");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=LEN(\"C5\")");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=IF((16 > B7), SUM(B5:B5), OR(D4 > 0, A3 < 100))");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=A1");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=IF((SQRT(A6) > LEFT(\"A5\", 3)), MAX(A6, C6), MIN(E3:E7))");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=C1");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=A7");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=AVERAGE(A7, PRODUCT(E7, C7))");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=INT(OR(E4 > 0, B3 < 100))");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=(ROUNDDOWN(E9, 0) ^ C3)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=LEFT(\"(C6 + A8)\", 5)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=-44");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=SQRT(-19)");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    println!("Seed 957517 evaluated target CellRef(9, 0): {:?}", target);
}
#[test]
fn test_fuzz_sum_cell_string_literal_parsing() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "-215.8");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "\"dhvueks\"");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "32");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-11");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "43");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "-228.367");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "57");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "-37.8");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "28");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "61");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "-129.451");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-35.61");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "-79.73");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "8");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-91");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "31.168");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "\"r\"");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "4");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "-3.6295");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "\"2\"");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=ROUNDDOWN(SUM(E5, A3), 2)");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=AVERAGE(C5:E5)");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=PRODUCT((C2 - 39), AND(-27 > 0, B1 < 100))");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=(AVERAGE(E3:E5) ^ UPPER(\"-28\"))");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=(C4 + ROUNDDOWN(-11, 2))");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=13");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=A4");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=SUM(B4, LEFT(\"A1\", 4))");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "87");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=IF((D6 > D3), 31, IF((D5 > B6), D3, 13))");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=D7");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=UPPER(\"43\")");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=C7");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "79");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=INT(AND(D3 > 0, E7 < 100))");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=B6");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=IF((-45 > SUM(40, B2)), UPPER(\"E6\"), 42)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "95");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=C7");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "17");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "-429.6347");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=D7");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=C9");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=20");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=(B1 ^ E4)");
    sheet.commit(None).unwrap();
    let e5 = sheet.get_result_data(&CellRef::new(4, 4));
    println!("Seed 623549 E5: {:?}", e5);
    let a6 = sheet.get_result_data(&CellRef::new(5, 0));
    println!("Seed 623549 A6: {:?}", a6);
    let b6 = sheet.get_result_data(&CellRef::new(5, 1));
    println!("Seed 623549 B6: {:?}", b6);
    let d6 = sheet.get_result_data(&CellRef::new(5, 3));
    println!("Seed 623549 D6: {:?}", d6);
    match a6 { ResultData::Float(f) => assert_eq!(f, 28.0), other => panic!("Expected Float(28.0), got {:?}", other) }
    match b6 { ResultData::Float(f) => assert!((f - 0.18525).abs() < 1e-3), other => panic!("Expected Float(~0.18525), got {:?}", other) }
    match d6 { ResultData::Error(ref e) => assert!(e.contains("#DIV/0!")), other => panic!("Expected Error(#DIV/0!), got {:?}", other) }
}
#[test]
fn test_fuzz_int_if_branch_evaluation() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "\"vMzcZU\"");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "65");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "55");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "432");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "41");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "92");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "-39");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "\"yFkqp3\"");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "2");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "-235.4");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "11");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "\"mVClvk\"");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-97");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "-331.7");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "45.2");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "15");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "52");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "\"jvnuJ\"");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "\"unJzzWMF\"");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=E5");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=ROUNDDOWN(35, 1)");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=(A2 - AVERAGE(E5:E5))");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=AVERAGE(E1:E4)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=SUM(A1:D2)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=E4");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=INT(IF((B6 > D5), A6, E6))");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=B3");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "-81");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=((E1 - D6) / IF((C3 > A3), A3, E2))");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=AVERAGE(E4:E6)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=(MAX(A7, 49) + D7)");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=13");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "-78");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=D8");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "9");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=AND(A8 > 0, (-45 + E2) < 100)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=1");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=SQRT(INT(E8))");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=LEFT(\"A9\", 2)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=((-19 * B1) / B4)");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=MIN(C1:D5)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=AVERAGE(B7, E3)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "-83.43000000000001");
    sheet.commit(None).unwrap();
    let b7 = sheet.get_result_data(&CellRef::new(6, 1));
    println!("Seed 672328 B7: {:?}", b7);
    let d10 = sheet.get_result_data(&CellRef::new(9, 3));
    println!("Seed 672328 D10: {:?}", d10);
    match b7 { ResultData::Float(f) => assert_eq!(f, 630.0), other => panic!("Expected Float(630.0), got {:?}", other) }
    match d10 { ResultData::Float(f) => assert_eq!(f, 197.3), other => panic!("Expected Float(197.3), got {:?}", other) }
}
#[test]
fn test_fuzz_sqrt_negative_e8_num_error() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "-415");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "2");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "-31.7166");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-72");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "-424.5495");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-169");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "-1");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "3");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "\"zI\"");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "370.6");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "252.449");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "-298.4798");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-22");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "-77");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "\"iR\"");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "26");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "\"11wh\"");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "-173.86");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "437.575");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "12");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "\"aG\"");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=LEN(\"(-41 + A5)\")");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "63");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "-179.314");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "-236.94");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=D3");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=C2");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=AVERAGE(E6:E6)");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=OR(PRODUCT(B4, -33) > 0, AND(42 > 0, -13 < 100) < 100)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=IF((28 > SQRT(-16)), ROUND(D4, 1), ROUNDDOWN(24, 1))");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=IF(((B5 * D1) > (A2 * -39)), MAX(33, 46), AVERAGE(-18, E3))");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=AVERAGE(CONCATENATE(\"-27\", \"C4\"), A7)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=IF((UPPER(\"B5\") > 32), A4, OR(D5 > 0, B2 < 100))");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=SUM((B3 + -48), -6)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=ROUNDUP(IF((D2 > A1), E6, 13), 2)");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=SQRT((E1 - 18))");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=INT(LEFT(\"-26\", 5))");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "\"3PTBp\"");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=E6");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "30");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=C2");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=E6");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=SQRT(E8)");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "484");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=-29");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "22.24");
    sheet.commit(None).unwrap();
    let b10 = sheet.get_result_data(&CellRef::new(9, 1));
    println!("Seed 585783 B10: {:?}", b10);
    let d8 = sheet.get_result_data(&CellRef::new(7, 3));
    println!("Seed 585783 D8: {:?}", d8);
}
#[test]
fn test_fuzz_roundup_e6_cell_reference() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "84");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "53");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "\"Q\"");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "-40");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "73");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "43");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "22");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "3");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "-68");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "-23");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "48");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-361.23");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "56");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "-140.93");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "469.3");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "-14");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "-459.021");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "-95");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "362.62");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=ABS(IF((C3 > -44), A5, 40))");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=D3");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=B2");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "8");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=ABS(-27)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=AND(D2 > 0, (A4 + 30) < 100)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "35");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=ROUNDUP(E6, 1)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=OR(-1 > 0, D2 < 100)");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=SUM(D4:D5)");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=E4");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "\"JUO\"");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=B6");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=IF((D1 > CONCATENATE(\"E4\", \"1\")), ROUND(B5, 2), E5)");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=D6");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=B5");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "\"hIVGEy\"");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "47");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=(UPPER(\"E5\") + (29 - -16))");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=OR(5 > 0, -21 < 100)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=-20");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=UPPER(\"IF((D6 > C4), C5, C8)\")");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=CONCATENATE(\"IF((D5 > C4), E7, -49)\", \"AVERAGE(-5, D5)\")");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=ABS(-22)");
    sheet.commit(None).unwrap();
    let c7 = sheet.get_result_data(&CellRef::new(6, 2));
    println!("Seed 931968 C7: {:?}", c7);
}
#[test]
fn test_fuzz_min_cell_string_ignore() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "\"ajTqx2gu\"");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "175");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "\" yil2xDW\"");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "\"XKxZ\"");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "-330");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "7");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "\"dQ\"");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "\"tMcVh\"");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-75");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "89");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "\"Lz\"");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "-87");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "-50");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "\"NWXitc\"");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "-98");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "\"HLFOVkq\"");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "\"fc\"");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=E5");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=AVERAGE(B2:B2)");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=B5");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=MAX(B3:E3)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=0");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=-5");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "134.5");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=UPPER(\"D2\")");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=PRODUCT(LEFT(\"D5\", 3), -20)");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=LEN(\"(E3 / C1)\")");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=OR(OR(A2 > 0, 4 < 100) > 0, ABS(-47) < 100)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=IF((IF((A1 > A7), 12, C4) > (A5 - D4)), IF((C2 > 13), B4, D1), A1)");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=MIN(ABS(E6), B4)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=-32");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=-42");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=C7");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=IF((ROUNDUP(C2, 0) > (-16 * A1)), C5, -25)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "-19");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=9");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=A6");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=48");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=INT(MIN(D6, A9))");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=-30");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "-130.326");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=(LOWER(\"D9\") - -24)");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 2));
    println!("Seed 923425 target: {:?}", target);
    match target { ResultData::Float(f) => assert_eq!(f, 0.0), other => panic!("Expected Float(0.0), got {:?}", other) }
}
#[test]
fn test_fuzz_int_sqrt_d6() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "\"mb2sof1i\"");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "78");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "444.5");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-55");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "342.8");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-70");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "62");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "-254");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "-183");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "\"OrlEQPfz\"");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "\"sMRhHT\"");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "-16");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "\"mhdP\"");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "76");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "-9");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "\"KFHr2en\"");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "-132.9");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "36");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "-365.88");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=PRODUCT(C5:E5)");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=PRODUCT(E5:E5)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=SQRT(D4)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=AND(E1 > 0, MAX(B5:C5) < 100)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "\"CkwbTno\"");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "-32");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=-43");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=C5");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=OR(D3 > 0, -50 < 100)");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=D4");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=INT(D6)");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "\"Ydy\"");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=UPPER(\"IF((19 > E7), A7, C4)\")");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=PRODUCT(SQRT(C5), A7)");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=A4");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=(IF((B6 > -20), A8, 21) + (D4 / D1))");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "236.302");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=OR(23 > 0, LEFT(\"D8\", 2) < 100)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=E8");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=MIN(D2:E3)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=A7");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=SUM(E5:E6)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=E1");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 1));
    println!("Seed 544657 target: {:?}", target);
    match target { ResultData::Float(f) => assert_eq!(f, 8.0), other => panic!("Expected Float(8.0), got {:?}", other) }
}
#[test]
fn test_fuzz_round_d6_single_decimal() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "32");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "-63");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-67");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "41");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "226.3");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "184.8");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "-87");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "68");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "51");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "-53");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "4");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "249.3257");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "66");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "356");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "295.24");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "124.9");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "-11");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "\"hIwAgm\"");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=IF((INT(47) > SUM(C5:C5)), C1, (C2 / D3))");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=B2");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=IF((E5 > LOWER(\"C3\")), INT(E4), RIGHT(\"C4\", 2))");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=(9 - (E1 + 19))");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=E3");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=1");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=(IF((B1 > -15), C4, C6) / AND(50 > 0, D2 < 100))");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=ROUNDUP(RIGHT(\"-44\", 4), 0)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=-37");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=ROUND(D6, 1)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=D5");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=D7");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "28");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=D7");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=LEN(\"A1\")");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=E2");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=AVERAGE(C7, 30)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=5");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=C6");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=E2");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=(IF((E9 > D9), 35, A3) * (C8 - C2))");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=50");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=C4");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=IF((PRODUCT(D3:E9) > -36), D1, -25)");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 0));
    println!("Seed 780052 target: {:?}", target);
}
#[test]
fn test_fuzz_and_sqrt_error_precedence() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "84");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "31");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "340.5032");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-373.2");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "-12");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "\"Bx\"");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "-130.3");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "95");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "\"EZ\"");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "\"xQSzxQK\"");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "326.07");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "\"M1Ahbv2p\"");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "10");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "-86");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "76");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "57");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "-484.77");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "\"Nxkov\"");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "-46");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "401");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=(-5 / -32)");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=IF((IF((B2 > E5), 18, 41) > A1), ROUNDUP(E2, 0), 37)");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=B1");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=-28");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=PRODUCT(A4:A5)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=INT(AND(A6 > 0, A3 < 100))");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=(28 * D1)");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "\"zndnANee\"");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "\"UWL1l3 \"");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "113.5114");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=OR(D7 > 0, OR(B2 > 0, B6 < 100) < 100)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "\"kH\"");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=ROUND(PRODUCT(D6, E5), 2)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=B2");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=CONCATENATE(\"15\", \"INT(A7)\")");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=(IF((E3 > E8), E1, C4) ^ 44)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=(UPPER(\"37\") ^ PRODUCT(C5:C8))");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=(ROUND(20, 2) + UPPER(\"B7\"))");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "24");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=SQRT(C3)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=IF((A5 > D8), -1, LEN(\"E6\"))");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=UPPER(\"MAX(C9, C2)\")");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=SUM(OR(-26 > 0, 28 < 100), IF((C5 > C6), B2, -48))");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=AND(SQRT(-39) > 0, SQRT(C7) < 100)");
    sheet.commit(None).unwrap();
    let e10 = sheet.get_result_data(&CellRef::new(9, 4));
    println!("Seed 632978 E10: {:?}", e10);
}
#[test]
fn test_fuzz_roundup_negative_tiny_exponent() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "38");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "49");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "65");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "-99");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "33");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "2");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "\"WQk3\"");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "\"d\"");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "-96");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "-97");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "-312.17");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "-89");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "52");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "\"AMQu\"");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "50");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "235.2");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "\"nSp\"");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "94.67");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "\"hhoveY\"");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=LOWER(\"(C1 / C4)\")");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "-51");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=LEFT(\"-41\", 2)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=21");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=MIN(B5, SQRT(-21))");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=E2");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=PRODUCT(B6:B6)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=OR(-28 > 0, D4 < 100)");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=PRODUCT(B4:B6)");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=E4");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=E3");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=IF((B4 > AVERAGE(A5:B6)), C1, ROUND(B5, 1))");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=E7");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=E3");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=D2");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=-25");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=LOWER(\"D3\")");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=SUM(AVERAGE(C7:C8), E6)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=ROUND(E4, 2)");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=PRODUCT(ROUNDUP(D1, 0), (E1 * A4))");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=OR(LEN(\"C1\") > 0, RIGHT(\"C8\", 3) < 100)");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=IF((-4 > SUM(-40, 22)), ROUNDUP(-28, 2), A5)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=ROUNDUP((E8 ^ -29), 2)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=A7");
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
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "257.4");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "54");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "90");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "-84");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "9");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "-24");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "\"ggbBvUQb\"");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "44");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "3");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-179.503");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "99");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "\"DwfJTT2\"");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "-63");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "386.5289");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "14");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=-8");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=IF(((C4 ^ E5) > B2), ABS(-38), ROUNDDOWN(C5, 1))");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=LEN(\"-30\")");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=C4");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=PRODUCT(B4, RIGHT(\"A4\", 3))");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=ROUNDDOWN(AVERAGE(D4:D4), 1)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=E1");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=8");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=D4");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=(AVERAGE(E3:E6) / (47 + E1))");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=E3");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "31");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=UPPER(\"C2\")");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=LEN(\"MIN(A4, 12)\")");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=PRODUCT((A3 ^ 45), -46)");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=-39");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=SQRT(D8)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=44");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=E8");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=34");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=B4");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=UPPER(\"C8\")");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=OR(D1 > 0, ROUNDUP(C8, 1) < 100)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=SUM(ROUND(-23, 0), E5)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=IF(((B9 * A6) > ABS(D2)), C7, -38)");
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
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "211.623");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "-204.42");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "91");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "113.74");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "-39");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "\"uU\"");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "\"3eEROE\"");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "64");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "99");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "166.3114");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "-28");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-3");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "\"1QlA\"");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "89");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "-235.82");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "-471.049");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "-411.8394");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "-15");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=C4");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=UPPER(\"(E5 / E3)\")");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=ROUNDDOWN(ROUNDDOWN(E4, 1), 1)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=-7");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=C4");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=A1");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=-30");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=AND(OR(E6 > 0, A6 < 100) > 0, AND(D6 > 0, C6 < 100) < 100)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=IF(((A1 ^ B3) > CONCATENATE(\"B6\", \"D1\")), SUM(8, 3), CONCATENATE(\"45\", \"E3\"))");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=LEN(\"B4\")");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=D1");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=IF((D5 > LOWER(\"A1\")), (25 * -9), CONCATENATE(\"12\", \"B5\"))");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=AVERAGE(B3:B3)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "\"KHtwf\"");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=ROUND(INT(A2), 1)");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=IF((LEN(\"C7\") > ABS(A7)), SQRT(A3), (E6 / D6))");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=D3");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=PRODUCT(C6:E7)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=(A7 * OR(C4 > 0, 16 < 100))");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=ROUNDUP(OR(-29 > 0, D1 < 100), 1)");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=OR(B3 > 0, LOWER(\"E5\") < 100)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=35");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=C4");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=ABS(ROUND(33, 0))");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "82");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 3));
    println!("Seed 278797 target: {:?}", target);
    match target { ResultData::String(ref s) => assert_eq!(s, "45E3"), other => panic!("Expected String(\"45E3\"), got {:?}", other) }
}
#[test]
fn test_fuzz_if_sqrt_negative_cond_error() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "104");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "9");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "5");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "264");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "-287");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "81");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "-38");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "62");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "\"jMg\"");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "84");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "\" uND\"");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "2");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "90");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "327");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "-32");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "-399.3008");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=MIN(E5:E5)");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=C4");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=-12");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=IF((IF((B5 > A4), A5, E1) > RIGHT(\"B4\", 3)), B4, A1)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=IF((IF((D4 > 26), -38, -33) > ABS(-7)), D3, (-15 ^ 15))");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=SUM(A6:B6)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=C5");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "-73");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=(A3 - E1)");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=E4");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=E4");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=OR((B6 / -30) > 0, 16 < 100)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=(48 - D2)");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=E6");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=IF((SQRT(C6) > LEN(\"-46\")), (45 / -28), MIN(D2, E5))");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=IF((37 > LEFT(\"-6\", 2)), PRODUCT(B7:B8), SUM(0, E1))");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "470");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=-19");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=PRODUCT(B1:C3)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=B7");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=SQRT(10)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=D7");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 0));
    println!("Seed 558820 target: {:?}", target);
}
#[test]
fn test_fuzz_addition_roundup_e7() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "32");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "2");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "288.4");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-77");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "-56");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "-473.497");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "-44");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "219");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "-264.47");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "23");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "29");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "\"UI\"");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "442.4109");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "20");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "99");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "-97");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "-373.8");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "\"f2z\"");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "64");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=E2");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=SUM(SUM(3, E5), MAX(C1, C5))");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=AND(OR(D3 > 0, C3 < 100) > 0, E3 < 100)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=(-21 ^ OR(E5 > 0, A5 < 100))");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=AVERAGE(C5:E5)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "\"v\"");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=E4");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=1");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=B4");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=(-47 + IF((D3 > A6), D5, 19))");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=(MIN(A3:C6) + AND(E7 > 0, A6 < 100))");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "\"a\"");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=(A2 - SUM(B4:B4))");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=ROUNDUP(UPPER(\"D6\"), 0)");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=C5");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=LEFT(\"-38\", 3)");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=D8");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=44");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=D1");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=(B3 + ROUNDUP(E7, 1))");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=A7");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=ROUND(UPPER(\"A2\"), 1)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=21");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=MIN(D8:D8)");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    println!("Seed 457220 target: {:?}", target);
    match target { ResultData::Float(f) => assert_eq!(f, -5.0), ResultData::Integer(i) => assert_eq!(i, -5), other => panic!("Expected -5, got {:?}", other) }
}
#[test]
fn test_fuzz_if_rounddown_branch() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "10");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "-51");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "73");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-50");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "-62");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "4");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "50");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "-340");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "60");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "40.26");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "\"Xm\"");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "-98");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "16");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "5");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "-80");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "-47.3");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "\"H1Mbf\"");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "-445.9098");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "-265.8");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "-82");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "-86");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=D2");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=7");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=RIGHT(\"E1\", 4)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=43");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=(AND(B3 > 0, -22 < 100) + LEN(\"D3\"))");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=(A6 ^ -30)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=MAX(A5:D6)");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=30");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=(C6 + (A5 ^ -24))");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=AVERAGE(A5:E6)");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=ROUNDDOWN(C5, 0)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=E4");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=IF((C3 > ROUNDDOWN(E6, 2)), D5, INT(A6))");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=C6");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=(16 + B3)");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=(MIN(E7:E8) - LOWER(\"C4\"))");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=IF((-50 > -18), -16, SUM(E7:E8))");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "-21");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=(MAX(E7, A8) * PRODUCT(A5, C1))");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=((46 + B6) ^ (D2 - -4))");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=A1");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=MIN(D9:E9)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=AVERAGE(D1:E6)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=(LEFT(\"C4\", 4) / E6)");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 2));
    println!("Seed 158982 target: {:?}", target);
}
#[test]
fn test_fuzz_round_division_b2_neg13() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "10");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "209.3");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "94");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "-45.488");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "283.3379");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "224");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "144.475");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "-51");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "-52");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "\"SqAEKqrx\"");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-93");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "-193.9086");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "27");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "21");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "16");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "\"MSG1hfJ\"");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "\"sEq\"");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "9");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "-85.3004");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=OR(B4 > 0, E5 < 100)");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=E1");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=ABS(B2)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=A1");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=-40");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=IF((46 > (D2 + D1)), LEN(\"A5\"), C5)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=-32");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=E1");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=(B2 / -13)");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=B3");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=SUM(C1:D2)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "-24");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=-12");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=(B6 * RIGHT(\"A1\", 5))");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=AND(PRODUCT(A5, B5) > 0, E4 < 100)");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=MIN(C4:D6)");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=MAX(A1:C5)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=ROUND(D7, 2)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=(PRODUCT(-41, B5) - B6)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=-26");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=(B2 - C1)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=ABS(B9)");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=(ROUNDUP(-45, 2) + IF((C9 > D2), A7, D4))");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "\"AeDrkE\"");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=(MAX(E4:E9) - ROUNDUP(-14, 1))");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 2));
    println!("Seed 845650 target: {:?}", target);
    match target { ResultData::Float(f) => assert_eq!(f, -21.8), ResultData::Integer(i) => assert_eq!(i, -22), other => panic!("Expected -21.8, got {:?}", other) }
}
#[test]
fn test_fuzz_abs_d6_sum_string_cond() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "100");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "\"ggwX\"");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "84");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "\"zs tV\"");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "\"UgssE\"");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "-39");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "-78");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "12");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "\"QxX3nKBd\"");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-95");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "-322.18");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "84");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "-84");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-17");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "55");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "94");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "-398.86");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "69");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "-52");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=ROUNDDOWN(C2, 0)");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=25");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "87");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=IF((SUM(-23, E1) > SUM(C2, B2)), SQRT(D4), B5)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "-87");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=D5");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "\"S3Mfwld\"");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=ABS(D6)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=AND(UPPER(\"A4\") > 0, OR(E1 > 0, A6 < 100) < 100)");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=IF((OR(45 > 0, C6 < 100) > IF((12 > B4), B2, E2)), RIGHT(\"C1\", 5), AND(30 > 0, D6 < 100))");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=OR(-4 > 0, ROUND(E7, 0) < 100)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=MAX(D4:D4)");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=AVERAGE(-14, (18 - E1))");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "\"I\"");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "-42");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=CONCATENATE(\"D6\", \"(B1 * D5)\")");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=(OR(E4 > 0, E8 < 100) - OR(A3 > 0, -35 < 100))");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=C1");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=LOWER(\"B3\")");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=IF((MIN(B7, D2) > D4), IF((B1 > -46), B1, D4), -40)");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=(C8 * (44 + A8))");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=UPPER(\"OR(B9 > 0, A3 < 100)\")");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=C9");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=ROUND(E4, 2)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=C3");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 2));
    println!("Seed 262160 target: {:?}", target);
}
#[test]
fn test_fuzz_round_d7_b10() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "\"MSx\"");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "\"acEbw\"");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-17");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "\"JkGE\"");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "-83");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "177.54");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "-76");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "-88");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "183.0719");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "-33");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-92");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "\"hldh\"");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "\"sK\"");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "-37");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "\"tBtsh\"");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "7");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "60");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "\"GiIU\"");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "47");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "32");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "-216.9266");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "\"OXH\"");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=RIGHT(\"ROUNDDOWN(-43, 2)\", 2)");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=E3");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=E1");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=(AVERAGE(D4:D5) * E4)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=SUM(E4, 35)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=ROUND((C2 / -2), 0)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=OR(C6 > 0, UPPER(\"D6\") < 100)");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=SQRT(OR(E6 > 0, B1 < 100))");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=A3");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=D5");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=11");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=A1");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=-26");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "\"MSpw2\"");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=LEFT(\"23\", 5)");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=OR(MIN(A1, B4) > 0, SUM(5, B3) < 100)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "97");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=ROUNDDOWN(D5, 1)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=AVERAGE(B2:D7)");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=(E3 / A9)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=ROUND(D7, 2)");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=AND(AVERAGE(B7:C8) > 0, MIN(C6:D9) < 100)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=A6");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=B9");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 1));
    println!("Seed 207640 target: {:?}", target);
}
#[test]
fn test_fuzz_abs_b8_cell_ref() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "347");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "9");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "-82");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-37");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "-112.89");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "386.74");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "100");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "48");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "46");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "\"rASQncw\"");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "\"VT\"");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "-27");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "-381.1");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "\"ozPy\"");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "46");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "41");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=C5");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=AND(C1 > 0, MAX(42, A2) < 100)");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=IF((A1 > 41), B4, A3)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "-26");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=SQRT(B3)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=E3");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=LOWER(\"(E6 - A3)\")");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=-46");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=ROUND(17, 1)");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=MIN(ROUNDUP(22, 0), LEFT(\"D3\", 3))");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=MIN(A7:D7)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=18");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=(A5 ^ INT(A7))");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=-42");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=10");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=A6");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=A1");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=E7");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=LEN(\"C3\")");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=LEFT(\"D7\", 1)");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=ABS(B8)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=PRODUCT(IF((-30 > B7), -3, A7), IF((A1 > -37), B6, A6))");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "311.4486");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=LEN(\"C4\")");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=E6");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    println!("Seed 142584 target: {:?}", target);
    match target { ResultData::Float(f) => assert_eq!(f, 18.0), ResultData::Integer(i) => assert_eq!(i, 18), other => panic!("Expected 18, got {:?}", other) }
}
#[test]
fn test_fuzz_sqrt_d6_zero_value() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "\"ZNC3o\"");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "479");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "-34");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "47");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "242.42");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "36");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "149.1177");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "35");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "10");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "466.1");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-43.34");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "9");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "34");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "-59");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "-42");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "343.927");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "-36");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "-338.9");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=ABS(A2)");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=-44");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=-30");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=C4");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=ROUNDUP(A2, 1)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=AND(AND(C3 > 0, 29 < 100) > 0, C3 < 100)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=-9");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=A3");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "49.021");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=(B1 / IF((-18 > -26), 6, A3))");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "27");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=C2");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "107.8808");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=E6");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "-49");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=C1");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=SQRT(D6)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=A3");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=(-38 + PRODUCT(C5, D6))");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=RIGHT(\"AND(D6 > 0, C8 < 100)\", 1)");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=(E1 / RIGHT(\"E6\", 5))");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "51");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=LEFT(\"PRODUCT(C3:E7)\", 5)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "-73");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=E7");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 1));
    println!("Seed 687697 target: {:?}", target);
}
#[test]
fn test_fuzz_rounddown_power_product_small() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "-279");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "444.6148");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "25");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-141.946");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "68");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "145.3101");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "493");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "12.7");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "138.2");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "8");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-70");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "-39.05");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "53.593");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "9");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "11");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "-359");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "-32");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "41");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "5");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "77");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "\"EZ\"");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "57");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=D3");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "404.051");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=A5");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=B4");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=(SUM(B2:B4) / ROUND(D4, 1))");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "\"I\"");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=MAX(D4, E2)");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=B6");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "\"FN\"");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=AND(B3 > 0, E3 < 100)");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=46");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "\"oE\"");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=-30");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=IF((B4 > -3), E7, ABS(-50))");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=(ABS(A2) / ABS(-4))");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=(ROUNDDOWN(B7, 2) ^ PRODUCT(D6:E6))");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "\"1\"");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=E7");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "\"Cg1U\"");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=LEN(\"(C7 - B7)\")");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "50");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=A7");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=AVERAGE(A9:B9)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=A2");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "48");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 0));
    println!("Seed 251454 target: {:?}", target);
    match target { ResultData::Float(f) => assert!((f - 1.1553665264221954e-10).abs() < 1e-18), ResultData::Integer(i) => assert_eq!(i, 0), other => panic!("Expected tiny float, got {:?}", other) }
}
#[test]
fn test_fuzz_abs_if_c6_cond() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "-48");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "2");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "-292");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "4");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "76");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "-55.17");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "-36");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "-58");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "\"Id\"");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "-4");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "59");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "8");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "59");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "386.088");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "\"Ow3PoM\"");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "81");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "64");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=(-4 + LEFT(\"-21\", 4))");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "79");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=-19");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=E3");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "51");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=D6");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=AVERAGE(29, LEN(\"C2\"))");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=MIN(OR(A4 > 0, B1 < 100), SQRT(E2))");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "\"rpKZse\"");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=LEN(\"IF((1 > -37), -11, E3)\")");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=D5");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=ABS(IF((C6 > 5), E4, A7))");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=MAX(E4:E6)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=CONCATENATE(\"A6\", \"MIN(B4, E3)\")");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=D4");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=ROUNDUP(15, 1)");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=OR(B6 > 0, ROUNDDOWN(B4, 0) < 100)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=B6");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=D4");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=AVERAGE(D4:E8)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=D9");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=((-25 / A2) * C6)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=ABS(ROUND(15, 0))");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "TRUE");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 1));
    println!("Seed 767352 target: {:?}", target);
}
#[test]
fn test_fuzz_left_string_max_range_if() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "\"yiCm\"");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "72");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "\"pa\"");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "-245.1091");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "-65");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "\"3P\"");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "4");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "431.6");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "-96.63");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "\"Ux\"");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "49");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "-14");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "20");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "-91");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "\"l1L\"");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "128");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "72");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "-387.3");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "-73.70999999999999");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=AND(D2 > 0, A3 < 100)");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=ROUNDDOWN(MIN(E1:E1), 2)");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=ROUND(CONCATENATE(\"E2\", \"-22\"), 1)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=B2");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=0");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=B1");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=(LEN(\"C4\") ^ (D2 - 49))");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=OR(ROUNDDOWN(D4, 1) > 0, A3 < 100)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=B3");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=7");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=SUM(ROUNDDOWN(B2, 1), D7)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=MAX(-39, (-8 ^ C4))");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=OR((D1 ^ D7) > 0, PRODUCT(C3, D3) < 100)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=C7");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=37");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=IF((LEFT(\"11\", 5) > MAX(D7:D8)), ROUND(D8, 1), 42)");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=-40");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "\"VQqCDql\"");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=46");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=D2");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=MAX(E7:E9)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=CONCATENATE(\"INT(A7)\", \"E1\")");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=LEFT(\"(E3 * D3)\", 4)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "-14");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "-75");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 0));
    println!("Seed 64619 target: {:?}", target);
    match target { ResultData::Float(f) => assert_eq!(f, 1.0), ResultData::Integer(i) => assert_eq!(i, 1), other => panic!("Expected 1, got {:?}", other) }
}
#[test]
fn test_fuzz_sqrt_e6_power_e2() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "\"pCQOcOPR\"");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "-15");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "-75");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "36.3");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "43");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "7");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "\"1\"");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "-392.04");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "382.4681");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "\"jtih\"");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "92");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "-10");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=-40");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "-306");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=IF((E1 > SUM(B3:E4)), LEFT(\"C5\", 1), (E3 / -41))");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "349");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=OR(OR(E1 > 0, A3 < 100) > 0, OR(B1 > 0, -46 < 100) < 100)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "-58");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=IF((D5 > (-8 + B2)), (-48 * D6), (A6 / C5))");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=-39");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=-3");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=B5");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=AND(9 > 0, PRODUCT(A1:E2) < 100)");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "94");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=(MAX(-36, D3) * ROUND(A2, 1))");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "79");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=LEFT(\"(E5 * E5)\", 5)");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=(SQRT(E6) ^ E2)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=B2");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=ROUNDUP(SUM(B8:D8), 2)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=IF((RIGHT(\"A1\", 4) > ROUNDUP(D2, 1)), AND(9 > 0, 22 < 100), ROUNDUP(-29, 1))");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "\"asFud3\"");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=IF((D8 > (B4 + A5)), 14, CONCATENATE(\"C3\", \"-4\"))");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=-21");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "406.135");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "\"KIioktqr\"");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 1));
    println!("Seed 699215 target: {:?}", target);
}
#[test]
fn test_fuzz_round_b6_boolean_ref() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "\"IYpej\"");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "12");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "-30");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-53");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "357.51");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "\"W\"");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "-57");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "65");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "-398.6184");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "1");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "27");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "95");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "4");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "-75");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "-74");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "-34");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "203");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "82");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=LOWER(\"A5\")");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=B3");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=AVERAGE(C1, 6)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=B2");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=(D3 ^ 39)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=LEN(\"(E6 + E5)\")");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=(IF((18 > B6), A3, -40) + (-14 / A2))");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=SQRT(LEN(\"A5\"))");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=(-13 + LOWER(\"D5\"))");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "\"Bf\"");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=ROUND(B6, 1)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=-33");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=LEFT(\"MIN(D5:E7)\", 3)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=D5");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=PRODUCT(D5:D7)");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=IF((UPPER(\"11\") > (E5 ^ C4)), 50, E5)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=(E6 / LOWER(\"C2\"))");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=D5");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=D2");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "88");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=IF((C5 > A9), (C5 * E6), ROUNDDOWN(A2, 0))");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "-21");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=C3");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=PRODUCT(B2, D7)");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 0));
    println!("Seed 734641 target: {:?}", target);
    match target { ResultData::Float(f) => assert_eq!(f, 0.0), ResultData::Integer(i) => assert_eq!(i, 0), other => panic!("Expected 0, got {:?}", other) }
}
#[test]
fn test_fuzz_power_int_negative_exponent() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "-28");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "22");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-13");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "289.98");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "41");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "72");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "-78");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "\"b3\"");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "-116.6164");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-92");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "\"Pzpq\"");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "\"GYIJf\"");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "6");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "69");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "-39");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "-114.2545");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "84");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "-99");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "-96");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "-30");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=B2");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=D1");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "1");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=D5");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=-35");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=(B1 + D5)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=A2");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=MIN(C3:E4)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=C4");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "-55");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=UPPER(\"SUM(B5:B5)\")");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=ROUNDDOWN(B3, 0)");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=-38");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=(-25 - ROUNDUP(E7, 1))");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=-37");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=B7");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=D6");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=(17 + C1)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=(B4 ^ INT(E8))");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=SUM(A5, IF((A5 > B7), 35, C3))");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=E6");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=AVERAGE(D8:D9)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=37");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "-68");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 3));
    println!("Seed 32959 target: {:?}", target);
    match target { ResultData::Float(f) => assert!((f - 1.615860020532192e-29).abs() < 1e-35), ResultData::Integer(i) => assert_eq!(i, 0), other => panic!("Expected float, got {:?}", other) }
}
#[test]
fn test_fuzz_sum_power_int_large_negative_exponent() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "\"dXjDtD\"");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "-45");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "5");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "\"uaEG cUx\"");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "-11");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "-98");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "-72.90000000000001");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "61");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "\" zr\"");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "75");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "-96");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "-114.918");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "54.135");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "-84");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "-63.6609");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "\"Rlb\"");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "-44.3");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "\"SM\"");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=MIN(A3, (-18 * -12))");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=RIGHT(\"AND(16 > 0, 33 < 100)\", 4)");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=(SQRT(10) / ROUND(E4, 2))");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=AND(4 > 0, AND(-15 > 0, 50 < 100) < 100)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=A4");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=D5");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=(MAX(E1:E4) + MIN(D3, E1))");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=(SUM(E6:E6) ^ INT(E6))");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=IF((SQRT(19) > C5), -23, C4)");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "-415.27");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=MAX(SUM(C5:D7), (A4 - -27))");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=37");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=MIN(D2:E4)");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "-5");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=(SUM(C8:C8) + IF((B7 > C8), B1, B2))");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=LEFT(\"(15 + 9)\", 1)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=LEFT(\"ROUNDDOWN(D5, 2)\", 4)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=(AVERAGE(C2:E3) / (A4 + -47))");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=ROUNDDOWN(IF((E7 > -13), A4, A1), 1)");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=AND(-43 > 0, -44 < 100)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=IF((B2 > LEFT(\"A7\", 5)), MIN(A1, -22), (D9 - B3))");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=SQRT((A6 + -12))");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=B9");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=19");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 3));
    println!("Seed 475402 target: {:?}", target);
    match target { ResultData::Float(f) => assert!((f - -1.1359866023259109e-237).abs() < 1e-240), ResultData::Integer(i) => assert_eq!(i, 0), other => panic!("Expected float, got {:?}", other) }
}
#[test]
fn test_fuzz_sqrt_right_string_multiplication() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "353.54");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "-86");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "-353.5");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-30");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "8");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "-25");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "-75");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "-466.059");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "\" jSiRFLD\"");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "-204");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "40");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "217.5");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "7");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "31");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "77");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "60");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "-60");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "10");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "5");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "-227.38");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "42");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "408.45");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=LEFT(\"A5\", 4)");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=B4");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=C3");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=ROUND(B1, 0)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "89.6992");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=OR((D1 ^ C4) > 0, MIN(B6:B6) < 100)");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=MAX(C5:C6)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=42");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=INT(A6)");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=-22");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=IF((32 > ROUNDUP(E1, 1)), E3, OR(A4 > 0, C7 < 100))");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=SQRT(D3)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "-331.98");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=C7");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=(SQRT(C8) * RIGHT(\"-26\", 4))");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=D4");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=ABS(B5)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=C6");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "5");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=SQRT(C4)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=ROUND(ABS(E9), 0)");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=SUM(E6:E8)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=ROUNDUP(A7, 2)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=PRODUCT(ROUNDDOWN(C1, 1), SUM(A6:A6))");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 0));
    println!("Seed 419803 target: {:?}", target);
}
#[test]
fn test_fuzz_abs_c8_b10() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "\"YxtpPvPu\"");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "\"oI\"");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-28");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "-62");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-128.04");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "\"lugMgk\"");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "24.45");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "39");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "13");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "26");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "14");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "4");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "100");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "-250.1");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "\"yWVnZrB\"");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "-77");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "167.4");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=(UPPER(\"3\") - E2)");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=AVERAGE(E3:E5)");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=ABS(ROUNDUP(B2, 1))");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=LEFT(\"C4\", 5)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=OR(UPPER(\"50\") > 0, C5 < 100)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=PRODUCT(A5:D6)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "268.8");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=A4");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=D5");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=38");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=MIN(B4:E4)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=IF((E2 > -45), (C2 / 17), (C4 + 17))");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=C7");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=26");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=IF((C6 > AND(A6 > 0, B7 < 100)), CONCATENATE(\"19\", \"A4\"), PRODUCT(C2, C4))");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=SUM(ROUND(26, 1), (C1 - D2))");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=OR(IF((-21 > D8), A7, A6) > 0, ROUNDDOWN(-6, 2) < 100)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=C2");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=D1");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=AVERAGE(IF((E8 > B5), D5, 8), (B6 / D4))");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=C8");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=ABS(C8)");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=UPPER(\"B8\")");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=26");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=C3");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 1));
    println!("Seed 640931 target: {:?}", target);
}
#[test]
fn test_fuzz_if_nested_product_b4_d4() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "486.89");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "\"Eaf\"");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "50");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-374.4602");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "-124.2236");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-315");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "96");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "94");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "-455.2");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "56.703");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "5");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "\"klN\"");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "55");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "4");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "-69");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "\"adEPPnfK\"");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "442.7");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "65.8");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "-370");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=ABS(ROUNDUP(A2, 2))");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=ROUND(0, 1)");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=-4");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=B3");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=-43");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=ABS(A6)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "121");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=IF((ROUNDDOWN(D3, 0) > ROUND(A3, 0)), D5, MIN(B3:D3))");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=-26");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=40");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=IF((IF((B5 > E7), E2, C4) > ROUND(E7, 0)), IF((A7 > -21), B7, 47), PRODUCT(B4:D4))");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "225.89");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=E5");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=-32");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=CONCATENATE(\"SQRT(E8)\", \"E1\")");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=UPPER(\"OR(C4 > 0, D3 < 100)\")");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=10");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=RIGHT(\"C5\", 2)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=C2");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=E5");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=(PRODUCT(A3:E5) / D6)");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=SUM(MIN(A3, 19), (-4 / D5))");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=IF((D3 > C4), CONCATENATE(\"A5\", \"6\"), -23)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=ABS(B8)");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 0));
    println!("Seed 92353 target: {:?}", target);
}
#[test]
fn test_fuzz_sqrt_concatenate_date_string() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "5");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "478.85");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "83");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "-92");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-86.03");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "-220.4");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "26");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "292.73");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "168");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "483.8");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "402");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "\"ESFAq\"");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "-107.8535");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-37.85");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "-100");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "54");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "-151.8833");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "\"1vokLCy1\"");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "-79");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "26");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=RIGHT(\"(C4 - C2)\", 1)");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=ROUNDUP(UPPER(\"A4\"), 2)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=-14");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=MIN(C5:C5)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=E1");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=49");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=RIGHT(\"MAX(E5:E5)\", 3)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=E5");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=MIN(E3:E5)");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=D4");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=-3");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=IF((D6 > OR(A2 > 0, D2 < 100)), 21, IF((C4 > 24), E6, C4))");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=SQRT(CONCATENATE(\"7\", \"-23\"))");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=D2");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=LEFT(\"PRODUCT(A1:C5)\", 3)");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=AND(ROUND(D7, 0) > 0, (D5 - C2) < 100)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "33.9");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=MIN(B7:C8)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=C7");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "-269.6165");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=MAX(50, 30)");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=C8");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=AVERAGE(A8:E9)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=PRODUCT(PRODUCT(C8, D2), E9)");
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
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "-14");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "-2");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "\"kdj\"");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "44");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "-32");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "\"UDocR 31\"");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "\"jr\"");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "7");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "-38");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "\"QDTdWDIa\"");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "-65");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-89");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "-355");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "15");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "\"nkp\"");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "178.6532");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "\"21oGJn\"");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "\"3\"");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=IF((LOWER(\"D5\") > SQRT(C2)), E4, D1)");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=ROUND(OR(C2 > 0, C1 < 100), 1)");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=UPPER(\"IF((C1 > E3), D4, 1)\")");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "-34.1584");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "-26");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=ROUND(12, 0)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=LOWER(\"-29\")");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=IF((IF((D6 > A1), E2, C5) > -31), D4, IF((21 > -27), -46, C5))");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=OR(A5 > 0, 8 < 100)");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=D1");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=ROUNDUP(LOWER(\"-21\"), 0)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=SUM(A6:B6)");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=LEN(\"B1\")");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=AVERAGE(E3:E3)");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=-4");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=E7");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "-30");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=IF((-44 > C1), ROUND(D1, 1), D5)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=A6");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=SQRT(A2)");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=((C5 ^ E8) / (D5 * 10))");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=AVERAGE(LOWER(\"E9\"), D8)");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=LEN(\"(A7 ^ E1)\")");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=IF((D8 > C7), SUM(C9:E9), (E4 * D7))");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=AVERAGE(A7:C8)");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 1));
    println!("Seed 758897 target: {:?}", target);
    println!("Target is {:?}", target);
}
#[test]
fn test_fuzz_if_left_hyphen_string_comparison() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "99");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "321.01");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "423");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "-13");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "100");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "480");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "\"y\"");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "4");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "-98");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "349.764");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "-437.2078");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "\"uEL\"");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "-31");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "-94");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "-81");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "165");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "-85");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "\"rz\"");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "97");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "50.4176");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=UPPER(\"(6 ^ -17)\")");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=ROUND(B5, 0)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "28");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=(ROUNDUP(-9, 2) * A5)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=OR(B6 > 0, MIN(B2:B5) < 100)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=(MAX(E5:E5) * D1)");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=ABS((6 - C1))");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=AND(MAX(D3, -19) > 0, ROUND(17, 0) < 100)");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=IF((LEFT(\"-44\", 1) > A6), ROUNDDOWN(D5, 1), CONCATENATE(\"A5\", \"B1\"))");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=B3");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=A2");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "-275.0375");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=E5");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=RIGHT(\"ABS(C1)\", 3)");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=MAX(D5, (E1 ^ A6))");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=AND(C2 > 0, A7 < 100)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=MAX(B1:D7)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=CONCATENATE(\"(A3 - B3)\", \"SUM(B7, D3)\")");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=AVERAGE(A1:D2)");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=LEN(\"C1\")");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=((2 - -4) / (-18 / E4))");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=C9");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=8");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=B7");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 4));
    println!("Seed 153407 target: {:?}", target);
    match target { ResultData::String(s) => assert_eq!(s, "A5B1"), other => panic!("Expected String A5B1, got {:?}", other) }
}
#[test]
fn test_fuzz_rounddown_average_a5() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "53");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "41");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "9");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "486.3");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "235.999");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "-219.94");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "\"qcBBBm\"");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "-42");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "\"DaE3PM\"");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "\"t\"");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "43");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "11");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-89");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "3");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "\"amn\"");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "39");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "\"Al2\"");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "\"PK\"");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "29");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "-27");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "\"uiY\"");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "449.2");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "-85");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=((D5 / D3) * 0)");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=UPPER(\"8\")");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=ROUND(AND(E2 > 0, E4 < 100), 1)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=AVERAGE(A5:A5)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=ROUNDDOWN((-7 + B4), 2)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=E6");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=E6");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=B5");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "-18");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=(ROUND(-14, 0) / C4)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "\"zMxPb\"");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=SUM(E5:E5)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=ROUNDDOWN(C7, 0)");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=C2");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=C4");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "-43");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=SUM(B2:E5)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "-472.5");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=(-34 * ROUNDUP(-50, 2))");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "21");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=MAX(E7:E7)");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=IF((-35 > IF((C7 > B6), A9, E7)), (6 - D3), C5)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=MIN(AVERAGE(C9:E9), -3)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=C7");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 3));
    println!("Seed 58883 target: {:?}", target);
    match target { ResultData::Float(f) => assert_eq!(f, -27.0), ResultData::Integer(i) => assert_eq!(i, -27), other => panic!("Expected -27, got {:?}", other) }
}
#[test]
fn test_fuzz_if_subtraction_rounddown_condition() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "-53");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-24");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "88");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "-290");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "17");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "8");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "-267.6631");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "-125.7");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "\"Up23XW\"");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "-239.8602");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "39");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "94");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "74");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "81");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "-7");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "-129.158");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "\"WenDvo\"");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "-12");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "-107.343");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "-88");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=AND(-21 > 0, AVERAGE(E3:E4) < 100)");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=LEN(\"D1\")");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=D5");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "15");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=D6");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "210.95");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=RIGHT(\"B3\", 1)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=AND(E2 > 0, IF((D1 > E5), -30, E5) < 100)");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=((-16 - A5) * D4)");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=ABS(OR(D4 > 0, B7 < 100))");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=MAX((D5 + 0), RIGHT(\"A1\", 2))");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=C7");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=UPPER(\"MIN(E1:E7)\")");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=SUM(C1:D7)");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=SUM(E2:E6)");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=ABS((E4 * A8))");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=IF(((A3 - A7) > ROUNDDOWN(A7, 0)), IF((47 > B1), 48, B5), A4)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=A2");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=OR(IF((C4 > A1), D6, D8) > 0, A3 < 100)");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=CONCATENATE(\"B2\", \"(-26 * D6)\")");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=10");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=D6");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=21");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=SUM(A1, IF((D6 > E6), B3, -48))");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 2));
    println!("Seed 354063 target: {:?}", target);
    match target { ResultData::Float(f) => assert_eq!(f, 94.0), ResultData::Integer(i) => assert_eq!(i, 94), other => panic!("Expected 94, got {:?}", other) }
}
#[test]
fn test_fuzz_rounddown_d7_empty_cell_ref() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "\"DTopPER\"");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "-404.1953");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "-22");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "13");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "6");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "6");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "\"VMqL\"");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "\"S QCn2gY\"");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "458.7529");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "55");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "-355.3733");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "\"b\"");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "\"yUpOMSRU\"");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "4");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "-35");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "-369.5703");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "162");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "43");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "268.344");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "-77");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "-63");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "27");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "-365");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=C4");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=1");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=ROUND(B2, 0)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=INT(RIGHT(\"C5\", 4))");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=E5");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=PRODUCT(B5:B5)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=A6");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=D1");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=ROUND(PRODUCT(C4:C7), 1)");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=ROUNDUP(A2, 2)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=E7");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=A4");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=A2");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=ROUNDDOWN(D7, 2)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=OR(ROUND(A2, 2) > 0, OR(37 > 0, D2 < 100) < 100)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=SQRT((D2 / B3))");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "40");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=D6");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=B5");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=MIN(OR(E3 > 0, B8 < 100), (C9 - B9))");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "\"zXLgd\"");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=IF((INT(D4) > E6), 35, -36)");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 1));
    println!("Seed 332459 target: {:?}", target);
    match target { ResultData::Float(f) => assert_eq!(f, 0.0), ResultData::Integer(i) => assert_eq!(i, 0), other => panic!("Expected 0, got {:?}", other) }
}
#[test]
fn test_fuzz_round_e6_empty_cell_ref() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "492.3757");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "-9");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "\"QAQK\"");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "\"bYw3RAT\"");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "-4");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "5");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "10");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "\"E\"");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "64");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "177.1");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "\"MkuQBum\"");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "-49");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "4");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "91");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "-26.531");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "-94");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "62");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "-91");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=CONCATENATE(\"OR(0 > 0, D5 < 100)\", \"(0 + A4)\")");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=OR(AND(E4 > 0, B1 < 100) > 0, D3 < 100)");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=IF((B5 > IF((E2 > 40), B2, E4)), A1, 49)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=AVERAGE((1 + -7), 44)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=ROUNDDOWN(B4, 2)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=AVERAGE(A4:A5)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "-83");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=B4");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=11");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=MAX(MIN(C4:C4), (E2 / A6))");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=ROUND(E6, 2)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "91");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=SQRT((B5 + 7))");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=D5");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=(IF((E8 > -50), C5, 1) - RIGHT(\"29\", 3))");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=IF((-23 > -3), 40, 50)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=((D6 - 12) * A7)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=UPPER(\"AND(-4 > 0, -46 < 100)\")");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=AND(8 > 0, B2 < 100)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=0");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=C2");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=C7");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=A3");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 0));
    println!("Seed 325598 target: {:?}", target);
    match target { ResultData::Float(f) => assert_eq!(f, 0.0), ResultData::Integer(i) => assert_eq!(i, 0), other => panic!("Expected 0, got {:?}", other) }
}
#[test]
fn test_fuzz_rounddown_abs_product_precision() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "\"jgJ\"");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "\"lor\"");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "-423");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-203");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "8");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "\"jNlVx\"");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "-462");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-35");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "-0.617");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "87");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "68");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-51");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "-40");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "\"1amXg\"");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "54");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "76");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "-34");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=(-29 / E3)");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=AND(-1 > 0, LEFT(\"A4\", 3) < 100)");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "\"zNHR\"");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=SUM(E1:E4)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=PRODUCT(D3:E5)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=(ROUND(D3, 0) ^ CONCATENATE(\"33\", \"-42\"))");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=E2");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=D1");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=37");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "\"Nr\"");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=47");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=ROUNDDOWN(ABS(E6), 2)");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=C6");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "9");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=50");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=-7");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=LEN(\"D3\")");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=-24");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=A2");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=8");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=E6");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=-38");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=-3");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=CONCATENATE(\"A4\", \"AVERAGE(B9:D9)\")");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 1));
    println!("Seed 810486 target: {:?}", target);
    match target { ResultData::Float(f) => assert_eq!(f, 29369.2), ResultData::Integer(i) => assert_eq!(i, 29369), other => panic!("Expected 29369.2, got {:?}", other) }
}
#[test]
fn test_fuzz_if_boolean_gt_int_upper_neg7() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "-20");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "-25.878");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-55.5739");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "94");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-137");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "146.1");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "9");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "268");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "10");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "\"lidXy\"");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "\"sDa\"");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-40.7562");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "382");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "-77");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "-42");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "\"BFyBvSvA\"");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "\"Qx3pF \"");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "341.131");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "\"SI eihZF\"");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "-39");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "68");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "356");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=ROUNDDOWN(AND(C3 > 0, -22 < 100), 0)");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=(LOWER(\"7\") + (48 * E5))");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=(IF((B3 > C5), D3, -26) + 11)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=LEN(\"A1\")");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=C6");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=IF((AND(E4 > 0, A5 < 100) > INT(C6)), UPPER(\"-7\"), A5)");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=C6");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=C3");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "-14");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=SQRT(SUM(A1, A1))");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=ABS(MAX(9, C4))");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=E4");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=IF((ROUNDUP(D3, 1) > (A7 + E1)), D6, -4)");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "-76");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=IF((ROUNDDOWN(C4, 1) > IF((B8 > B4), D7, E2)), A3, D2)");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=D5");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=IF((AND(C7 > 0, -18 < 100) > OR(B6 > 0, -30 < 100)), (B8 ^ A4), (A7 - A6))");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=17");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=D4");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=E7");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=((E7 ^ E6) - 46)");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=(LEFT(\"E1\", 4) ^ (16 * E4))");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=AND(MAX(A2:A3) > 0, LEFT(\"E3\", 1) < 100)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=A1");
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
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "-264.7");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "94");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "-328");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-9");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "146.41");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-25");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "157.337");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "-67");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "-12");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "-366.201");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "99");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "-49");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "9");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "5");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "\"E\"");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "5");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "26");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "83");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "3");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "174");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=A1");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "282");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=UPPER(\"-31\")");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=IF((-38 > PRODUCT(-18, A1)), B2, 36)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=PRODUCT(A4:B5)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=42");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=(C2 / IF((3 > -12), D6, -18))");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=B4");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=19");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=E5");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=(MAX(D2, B1) - (D6 + C5))");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=E5");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=ROUNDUP(IF((E7 > E7), C7, -40), 1)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "-86");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=CONCATENATE(\"E1\", \"A4\")");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=ROUNDDOWN(IF((B8 > B5), C7, 35), 0)");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=40");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=UPPER(\"OR(D6 > 0, B5 < 100)\")");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=SUM(RIGHT(\"B5\", 5), B8)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=50");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=(IF((-21 > -46), A7, B8) * LEN(\"-28\"))");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "54");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=(A8 + AVERAGE(E9:E9))");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=AND((D3 ^ E8) > 0, ROUNDUP(D5, 2) < 100)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=SQRT(ABS(C5))");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 0));
    println!("Seed 101079 target: {:?}", target);
    match target { ResultData::Float(f) => assert_eq!(f, 0.0), ResultData::Integer(i) => assert_eq!(i, 0), other => panic!("Expected 0, got {:?}", other) }
}
#[test]
fn test_fuzz_if_sqrt_boolean_string_branch() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "\"pCFGZmpm\"");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "-434.8");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-413.79");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "169.5");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "2");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "-130");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "\"CIR\"");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "133");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "\"SwcK\"");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "30");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-13");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "69");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "-34");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "-99.045");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "11");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "17");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "46");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "-98");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "\"1UOZV3 f\"");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "-90");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "-78");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=B5");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "-83");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=PRODUCT((-30 - E5), -11)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=LEFT(\"B4\", 5)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=PRODUCT(A2:C3)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "\"FtNOl\"");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=OR(PRODUCT(B1:C2) > 0, (39 / C1) < 100)");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "93");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=-34");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=INT(C1)");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=B7");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=E2");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=-5");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "154.153");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=-31");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=17");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=MAX(C2:D3)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=OR(17 > 0, PRODUCT(B6:C6) < 100)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "\"ccKYKd\"");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "\"apQN\"");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "80");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=AND(A5 > 0, A4 < 100)");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=E8");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=IF((ABS(B2) > SQRT(A8)), D6, AVERAGE(B9:D9))");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=3");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 3));
    println!("Seed 815090 target: {:?}", target);
    match target { ResultData::String(s) => assert_eq!(s, "B4"), other => panic!("Expected B4, got {:?}", other) }
}
#[test]
fn test_fuzz_if_d4_division_abs_b7() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "-47");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "-6");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "77");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "\"reWs\"");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "7");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-392.33");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "-96");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "9");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "479.49");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "-291.13");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "211.2");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "4");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "13");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "393");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-47");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "-36");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "\"zgPkYl\"");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "\"aEL\"");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "330.8");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=CONCATENATE(\"IF((30 > B4), 13, E1)\", \"C4\")");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=SQRT(AND(C4 > 0, A5 < 100))");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=SQRT(6)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=MIN(D3:D4)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=CONCATENATE(\"C2\", \"INT(48)\")");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=E1");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=LEFT(\"ROUNDUP(-17, 0)\", 4)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=(E2 + (-47 / B3))");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=CONCATENATE(\"D4\", \"SQRT(36)\")");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=SUM(A7:A7)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=PRODUCT(1, ROUNDUP(25, 2))");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=IF((LOWER(\"C1\") > E3), IF((-44 > E1), D7, A2), D6)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=A6");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=A5");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=IF((D4 > (-41 / 14)), SUM(D4:E8), ABS(B7))");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=(LEFT(\"C4\", 2) * PRODUCT(C8:C8))");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "-41");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=CONCATENATE(\"A7\", \"D3\")");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=B4");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "-16");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=ROUNDDOWN(38, 0)");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=(IF((C3 > C1), 50, B6) * 17)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=(IF((E5 > D3), B5, 38) + (-34 - -47))");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "\"Bj1MXyY\"");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 0));
    println!("Seed 273076 target: {:?}", target);
    match target { ResultData::Float(f) => assert_eq!(f, 7.0), ResultData::Integer(i) => assert_eq!(i, 7), other => panic!("Expected 7, got {:?}", other) }
}
#[test]
fn test_fuzz_c2_multiplication_rounddown_abs() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "-2");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "\"FoKzo\"");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "12");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "\"p\"");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "-57");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "473.86");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "-89");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "-4");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "-49");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "214.8");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "-16");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "\"QOu2I\"");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "214.28");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "\"HCM\"");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "48");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "-84");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-72");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "\"o\"");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "83");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "9");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "10");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "103.2");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "37");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "-76");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=B4");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=36");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=SQRT(MIN(B5, -43))");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=ABS(2)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=C4");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=C5");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=OR(A6 > 0, D1 < 100)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "-236.395");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=E5");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=((B4 - -4) / C3)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=ROUNDUP(UPPER(\"36\"), 2)");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=-1");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=15");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=9");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "-90");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=(LEFT(\"8\", 1) ^ SQRT(E4))");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=(C2 * ROUNDDOWN(E6, 2))");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=IF((-24 > ROUNDUP(A7, 0)), (23 ^ E6), OR(A3 > 0, D7 < 100))");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=IF((D5 > AVERAGE(-18, C4)), (C5 - D1), A7)");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=D3");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=IF((C8 > B4), OR(C9 > 0, -38 < 100), MAX(E8:E8))");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=39");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=ROUNDUP((D2 * 12), 2)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "-448");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 2));
    println!("Seed 667784 target: {:?}", target);
    match target { ResultData::Float(f) => assert_eq!(f, -8.0), ResultData::Integer(i) => assert_eq!(i, -8), other => panic!("Expected -8, got {:?}", other) }
}
#[test]
fn test_fuzz_addition_sqrt_negative_num_error() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "-17");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "26");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "-100");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "\"l\"");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "-41");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "9");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "46");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "-175.65");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "-27");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "100");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "18.239");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "5");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "-141.0941");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "\"BnXOprLe\"");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-304.064");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "-79");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "-93");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "\"PI\"");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "-66");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "-188");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=LOWER(\"AND(40 > 0, -33 < 100)\")");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=(SQRT(6) ^ AND(16 > 0, A4 < 100))");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=AND(AND(A1 > 0, A3 < 100) > 0, IF((B1 > C4), E2, 9) < 100)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=RIGHT(\"AND(B2 > 0, E1 < 100)\", 1)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=AND(IF((D3 > C2), E5, A4) > 0, (A1 * -24) < 100)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=C1");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=PRODUCT(8, C3)");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=-30");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=MAX(A3:B4)");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=(LEFT(\"E4\", 3) ^ E1)");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=1");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=C7");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=A6");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=D1");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=C4");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=ROUNDUP(RIGHT(\"C5\", 2), 0)");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=E5");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "-142.96");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=IF((MIN(A6:D6) > -24), LEN(\"C8\"), LEN(\"E2\"))");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=C4");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=(D2 + SQRT(B8))");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=(47 * UPPER(\"E4\"))");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=-47");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=(IF((43 > 6), -48, 0) * SQRT(-4))");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=C4");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    println!("Seed 964755 target: {:?}", target);
    match target { ResultData::Error(e) => assert_eq!(e, "#NUM!"), other => panic!("Expected #NUM!, got {:?}", other) }
}
#[test]
fn test_fuzz_int_negative_constant_cell_ref() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "72");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "1");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-31");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "52");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "93");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "15");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "-33");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "3.66");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "-42");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "75");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "41");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "21");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "38");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "\"2bUngO2\"");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "131.699");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "56");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "6");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "58");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=AVERAGE((E1 - A1), AND(A3 > 0, -9 < 100))");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=5");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=(10 + SUM(E4, 9))");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=LEFT(\"-28\", 3)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=UPPER(\"IF((E1 > B5), C1, 27)\")");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=MAX(OR(C3 > 0, A1 < 100), MAX(D3:E3))");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=(LOWER(\"A2\") - C3)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "137.3");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=C3");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=AND(E5 > 0, -39 < 100)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=-33");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "314.9");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=AND(ABS(-38) > 0, A3 < 100)");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "-33");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=38");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=LEFT(\"OR(-16 > 0, 24 < 100)\", 3)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=D7");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=MAX(ROUND(-20, 1), A5)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=A3");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=INT(E9)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=(ROUND(13, 1) + ROUND(E1, 1))");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=MIN(C8, ROUND(E4, 1))");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=D1");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=CONCATENATE(\"(C2 * -12)\", \"IF((B8 > C1), B8, E4)\")");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    println!("Seed 947857 target: {:?}", target);
    match target { ResultData::Float(f) => assert_eq!(f, -33.0), ResultData::Integer(i) => assert_eq!(i, -33), other => panic!("Expected -33, got {:?}", other) }
}
#[test]
fn test_fuzz_if_number_gt_string_comparison() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "79");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "-61");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "10");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "\"llTMht\"");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-72");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "\"mhZ\"");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "-10");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "131.678");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "-19");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "\"svgONaM\"");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "-164.077");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-310.3");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "-4");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "-444.004");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "-58");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-50");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "4");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "41");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "\"FSyZd\"");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "23");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "11");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "\"I qQ1iAl\"");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=4");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=17");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=-21");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=PRODUCT(LEN(\"9\"), B4)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=ABS(D3)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=(-43 / (3 ^ 23))");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=A5");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=E5");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=LEFT(\"ROUND(-29, 1)\", 1)");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=AVERAGE(B1:C5)");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=AVERAGE(E6:E7)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=IF((ABS(E7) > LEFT(\"D2\", 2)), UPPER(\"E2\"), A3)");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=ROUNDDOWN(LEFT(\"A7\", 3), 1)");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=IF((CONCATENATE(\"C4\", \"E2\") > CONCATENATE(\"18\", \"32\")), AVERAGE(B1, E5), ROUNDUP(A7, 2))");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=IF((B2 > UPPER(\"E6\")), CONCATENATE(\"A7\", \"E8\"), (B5 / A6))");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=-40");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "-320");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=(IF((E3 > 28), 42, -28) - 4)");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=ROUNDDOWN(AVERAGE(C1, A7), 1)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=A1");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "-22");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=RIGHT(\"A6\", 1)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=UPPER(\"C4\")");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 1));
    println!("Seed 315137 target: {:?}", target);
    match target { ResultData::String(s) => assert!(s.contains("svgONaM")), other => panic!("Expected svgONaM, got {:?}", other) }
}
#[test]
fn test_fuzz_int_positive_constant_cell_ref() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "60");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "472.3");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "10");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-43");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "-75");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-97");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "437.62");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "-11");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "-51");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "34");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "-287");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "\"32Ouh\"");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "64");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "\"DscZws\"");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "-480");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "\"RDAR\"");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "-50");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "-9");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "3");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "\"S\"");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=E4");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=-11");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=(LEN(\"B2\") * 35)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=MIN(IF((C4 > D2), E2, B4), (-9 ^ 1))");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=IF((LEN(\"A3\") > (D4 / D2)), 4, D3)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=IF((B5 > C2), 3, 23)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "-38.569");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=(AVERAGE(B1, 44) ^ A6)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "-68");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "110.243");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=30");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "-7.66");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "\"wnYkoMY\"");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=IF((B5 > LEFT(\"-50\", 1)), LOWER(\"34\"), E6)");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=ROUNDUP(ROUNDUP(C1, 0), 0)");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=(MIN(-17, E4) * B4)");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=B8");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=(-27 + (D4 - 15))");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=INT(E8)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=(D6 * A2)");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "\"L2\"");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=A7");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=INT(E3)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=LEN(\"AND(A9 > 0, A8 < 100)\")");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=LEFT(\"D5\", 1)");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 3));
    println!("Seed 284690 target: {:?}", target);
    match target { ResultData::Float(f) => assert_eq!(f, 10.0), ResultData::Integer(i) => assert_eq!(i, 10), other => panic!("Expected 10, got {:?}", other) }
}
#[test]
fn test_fuzz_average_nested_d2_d8() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "-262.09");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-19");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "356.364");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "\"dWMrcEg\"");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "-402.123");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "\"A\"");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "-6");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "210.2262");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "-39");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "\"jDjm m\"");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-210.36");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "-55");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "-415.4");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "75");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "267.06");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=LOWER(\"AND(-40 > 0, -42 < 100)\")");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=A4");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=MIN(B2:D2)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=ABS(B5)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=C2");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "-176.875");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=ROUND((C5 - A4), 1)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=ROUNDUP(IF((E6 > 41), E6, D6), 2)");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=ROUNDDOWN(-42, 1)");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=IF((-11 > SUM(E6, A4)), 16, -47)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=AVERAGE(AND(D1 > 0, D7 < 100), (A6 + B3))");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=MIN(C6:C7)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=C3");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=AND(IF((D4 > D3), A5, -34) > 0, B2 < 100)");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=(SQRT(A3) ^ RIGHT(\"24\", 3))");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=C6");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=ROUNDDOWN(B5, 2)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=B7");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=(B1 / SUM(C4:E6))");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=AND(E4 > 0, MAX(D7, -47) < 100)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=ROUND(A7, 1)");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=AVERAGE(AVERAGE(D2:D8), C1)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=-34");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=D1");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 2));
    println!("Seed 585372 target: {:?}", target);
    match target { ResultData::Float(f) => assert!((f - -111.391725).abs() < 1e-4), ResultData::Integer(i) => assert_eq!(i, -111), other => panic!("Expected -111.391725, got {:?}", other) }
}
#[test]
fn test_fuzz_round_e6_empty_ref() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "\"pXoVoGJ\"");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "317.3");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "\"3LXfHOa\"");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "86");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-349");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "\"hvR\"");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "23");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "\"Dgs\"");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "\"Iamx\"");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "89");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "31");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "-338.133");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "41");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "257.4388");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "36");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "\"CX fGNIK\"");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "31");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "-440");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "-1");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "58");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "93");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=LEN(\"(0 - A2)\")");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=MIN(A1:C3)");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=C4");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=C5");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "\"Ro\"");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=ROUND(E6, 2)");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=30");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "-83");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=SUM(A4:D6)");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=(C6 * -25)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=A7");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=2");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=(-5 * -10)");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "48");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=SQRT(IF((B5 > -35), -32, B1))");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=C7");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=IF((RIGHT(\"B6\", 3) > (C5 - E6)), D6, C2)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=4");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=INT(PRODUCT(-30, 27))");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=C6");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "-26");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=B5");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=(E9 * (C5 - B2))");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=LEFT(\"A1\", 4)");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 1));
    println!("Seed 483010 target: {:?}", target);
    match target { ResultData::Float(f) => assert_eq!(f, 0.0), ResultData::Integer(i) => assert_eq!(i, 0), other => panic!("Expected 0, got {:?}", other) }
}
#[test]
fn test_fuzz_rounddown_e6_boolean_ref() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "-281");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "10");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-177");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "\"WOb\"");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "-153");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "58");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "\"v\"");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "-353");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "\"OJrt\"");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "95");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "117");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "89");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "12");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "219.818");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "4");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "70.5967");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "\"g2OgFkx\"");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "9");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "98");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "475");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=MIN((A1 / D2), B1)");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "56");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=6");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=(19 ^ -22)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=AND(LEN(\"2\") > 0, 31 < 100)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=C4");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=ROUND(PRODUCT(-5, D2), 0)");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=-36");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "3");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=B2");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=A4");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=MIN(D6:D6)");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=((C1 - D4) - B2)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=IF((LEN(\"C5\") > RIGHT(\"E2\", 2)), ROUND(43, 0), (B2 ^ C4))");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=AND(AND(C7 > 0, A4 < 100) > 0, A5 < 100)");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=UPPER(\"MAX(B8, D3)\")");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=ROUNDDOWN(E6, 2)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=SUM(INT(E2), ROUND(-39, 2))");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=LOWER(\"A2\")");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=A3");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=C4");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=CONCATENATE(\"30\", \"OR(E7 > 0, C2 < 100)\")");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=12");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=ROUNDDOWN(IF((A2 > A7), A2, B3), 0)");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 1));
    println!("Seed 442593 target: {:?}", target);
    match target { ResultData::Float(f) => assert_eq!(f, 1.0), ResultData::Integer(i) => assert_eq!(i, 1), other => panic!("Expected 1, got {:?}", other) }
}
#[test]
fn test_fuzz_roundup_e8_val() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "89");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "6");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "-13");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-26");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "-5");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "3");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "267.908");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "\"HrVDbJ\"");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "-100");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "\"j rEDd\"");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "-185.2262");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "\"XeIdn\"");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "\"i\"");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "54");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "\"AwTAJ S\"");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "55");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "-8");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "-3");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "\"ADUZGxg\"");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "72");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "-76");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "21");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "136");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=28");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=D3");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=AVERAGE(D1:E4)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=A2");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=(UPPER(\"D3\") / LEN(\"-50\"))");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=(E4 * AND(17 > 0, 12 < 100))");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=LEFT(\"B3\", 5)");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=AVERAGE(D2:E5)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=CONCATENATE(\"ROUNDUP(D3, 0)\", \"(E6 * -21)\")");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=C6");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=A7");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=ROUNDDOWN(LEN(\"-17\"), 1)");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=C5");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=LEN(\"IF((C4 > C4), E6, 31)\")");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=IF((C1 > (B3 - D2)), LOWER(\"21\"), AVERAGE(A4:B4))");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=B5");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=ROUNDUP(E8, 2)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=PRODUCT(44, SUM(E3:E7))");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=AND(OR(B3 > 0, E1 < 100) > 0, C8 < 100)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=AND(AND(C3 > 0, C5 < 100) > 0, LEN(\"7\") < 100)");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "-46");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=D9");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=C2");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=IF((LEFT(\"-40\", 4) > IF((E3 > D3), A4, C1)), RIGHT(\"-25\", 5), -43)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=B1");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 1));
    println!("Seed 507073 target: {:?}", target);
    match target { ResultData::Float(f) => assert_eq!(f, 21.0), ResultData::Integer(i) => assert_eq!(i, 21), other => panic!("Expected 21, got {:?}", other) }
}
#[test]
fn test_fuzz_rounddown_e6_val() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "\"n2J\"");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "-438.096");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-100");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "5");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "22");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "8");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "161.8898");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "\"rfVLxYq\"");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "-384.8124");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "43");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "351.73");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "-23");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "69");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "73");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "-260.7");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "-77");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "-397.07");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=UPPER(\"AVERAGE(19, D2)\")");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=26");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "-52");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=(C3 - E2)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=C4");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=C1");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=IF(((-39 / -13) > (30 - C6)), (A6 ^ -24), D1)");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=(RIGHT(\"A1\", 2) * C3)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=2");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=LEN(\"A4\")");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=CONCATENATE(\"AND(A4 > 0, A3 < 100)\", \"PRODUCT(E1, D5)\")");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=C2");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=ROUNDDOWN(E6, 1)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "3");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=LEFT(\"IF((-42 > B2), -19, 19)\", 2)");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=D2");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=PRODUCT(E8:E8)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=16");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=(B2 * RIGHT(\"A3\", 5))");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=ROUND(E5, 0)");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=B9");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=E7");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=LEFT(\"A2\", 3)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=(-30 * C1)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=MIN(E2:E5)");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 2));
    println!("Seed 517687 target: {:?}", target);
    match target { ResultData::Float(f) => assert_eq!(f, 0.0), ResultData::Integer(i) => assert_eq!(i, 0), other => panic!("Expected 0, got {:?}", other) }
}
#[test]
fn test_fuzz_abs_c6_val() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "-21");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "\"VEFaOfm\"");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-85");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "-60");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "-238");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "495.7921");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "14");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "-22");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "63");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "\"x\"");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "\"MSikQ\"");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "19");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "\"Wx\"");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "8");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "\"bTU\"");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "51");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "2");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "8");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "8");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=D3");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=OR(LOWER(\"E1\") > 0, SQRT(B5) < 100)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=E3");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=D2");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=IF((B5 > C6), SQRT(D5), ROUNDUP(-14, 1))");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=E3");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=AND(RIGHT(\"D3\", 5) > 0, ROUNDUP(B4, 0) < 100)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=AND(UPPER(\"15\") > 0, IF((25 > D6), C2, -43) < 100)");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=C6");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=ABS(C6)");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=C7");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=-15");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=IF((-33 > ROUNDUP(-8, 2)), (B4 / C2), LOWER(\"C3\"))");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=(OR(A4 > 0, D7 < 100) - CONCATENATE(\"E8\", \"D8\"))");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=B6");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=SQRT(MIN(E6:E7))");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=B7");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=LEFT(\"ROUNDDOWN(33, 2)\", 2)");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=(LEN(\"16\") * MAX(-46, A8))");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=PRODUCT(B2:B5)");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=LEFT(\"(B1 * 31)\", 5)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=AVERAGE(-50, A3)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "76");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 1));
    println!("Seed 159433 target: {:?}", target);
    match target { ResultData::Float(f) => assert_eq!(f, 1.0), ResultData::Integer(i) => assert_eq!(i, 1), other => panic!("Expected 1, got {:?}", other) }
}
#[test]
fn test_fuzz_roundup_if_val() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "\"WTcuBFQn\"");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "\"v3EwtV\"");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "400.724");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-10");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "55.3559");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "-23");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "61");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "\"yflJCHu\"");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "-55");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "2");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "92");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "\"L\"");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "-84");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "6");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "332");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "30");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "\"HIAZyhx\"");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "\"hzAZbX\"");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "\"C\"");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "\"xnBd\"");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=CONCATENATE(\"A4\", \"B2\")");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=SQRT(MAX(B4:D5))");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=C4");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=A5");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=UPPER(\"45\")");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=PRODUCT((27 * A1), D6)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=35");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=RIGHT(\"ROUND(22, 0)\", 3)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=E2");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=D5");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=LOWER(\"-23\")");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=((28 ^ D5) ^ -5)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=B7");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=SUM(INT(-14), D5)");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=RIGHT(\"MIN(E3, A5)\", 2)");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=39");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=MIN(A3:D6)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=OR(A7 > 0, (-22 - C7) < 100)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=UPPER(\"MAX(14, -26)\")");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=ROUNDUP(IF((E3 > E2), C6, E6), 1)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=48");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=MIN(A7:E8)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=(C2 / A9)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "-80");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    println!("Seed 719875 target: {:?}", target);
    match target { ResultData::Float(f) => assert_eq!(f, 45.0), ResultData::Integer(i) => assert_eq!(i, 45), other => panic!("Expected 45, got {:?}", other) }
}
#[test]
fn test_fuzz_if_rounddown_cond_val() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "-104.5");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "61");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "-48");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-265");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-8");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "433");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "-66");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "343");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "3");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "-3");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "415.555");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "388.234");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "87");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-41");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "-31");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "-22");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "99");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "-263");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "\"S2ddn2p\"");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "-168.413");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=B5");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=((E3 ^ D5) * (C4 - A2))");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=AND(OR(B5 > 0, B4 < 100) > 0, (B2 - B5) < 100)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=B4");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=-42");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=IF((ROUNDDOWN(E6, 2) > 50), D1, -5)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=ROUNDUP(B1, 0)");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=PRODUCT(E4:E6)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=ROUND(RIGHT(\"C5\", 1), 0)");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "\"mjY2Pekq\"");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=UPPER(\"ABS(-20)\")");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=1");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=(-24 + (42 + A5))");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=ROUNDUP(INT(C2), 1)");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=A1");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=ROUND(SQRT(1), 1)");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=ABS(MIN(B6:D6))");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=IF((A5 > SUM(E4, D6)), -23, C1)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=(ROUNDDOWN(C5, 1) * OR(E3 > 0, E5 < 100))");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=ABS(D7)");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=C5");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=IF((RIGHT(\"C2\", 4) > B7), A6, E7)");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=IF((ROUNDUP(38, 2) > INT(B8)), C9, UPPER(\"-19\"))");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=A2");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "TRUE");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 0));
    println!("Seed 439684 target: {:?}", target);
    match target { ResultData::Float(f) => assert_eq!(f, -5.0), ResultData::Integer(i) => assert_eq!(i, -5), other => panic!("Expected -5, got {:?}", other) }
}
#[test]
fn test_fuzz_max_concatenate_or_value_error() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "-14");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "53");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-187");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "82");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "7");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "224.428");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "251");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "\"OirPKSm\"");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "-412.847");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "-96");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "\"sE\"");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "386.3835");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-4");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "\"oRQ2oX\"");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "38");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "-6");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-39");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "447");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "\"K\"");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "-82");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "-15");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "59");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "41");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "4");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=E4");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=LOWER(\"INT(A1)\")");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=IF((A5 > AVERAGE(-32, B2)), ABS(B5), AND(10 > 0, E3 < 100))");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=D1");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=A2");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=IF((-43 > LOWER(\"A6\")), 50, SQRT(-45))");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=A2");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "39");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=A5");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=B3");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=ROUND(IF((D4 > E3), C2, C3), 1)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "12.4211");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=ABS(C5)");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=C7");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=E2");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "-367.7296");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=SQRT(C8)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=(SQRT(D1) / E5)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=AVERAGE((D4 - A8), C7)");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=LEN(\"A5\")");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=MIN(A1:E5)");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=MAX(CONCATENATE(\"E6\", \"B7\"), OR(A6 > 0, D9 < 100))");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "178.7795");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=A3");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 2));
    println!("Seed 509664 target: {:?}", target);
    match target { ResultData::Error(e) => assert!(e == "#VALUE!" || e == "#NUM!"), other => panic!("Expected #VALUE! or #NUM!, got {:?}", other) }
}
#[test]
fn test_fuzz_sqrt_if_cond_val() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "82");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "486.7");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "80");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "95");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "393.634");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "-306.052");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "7");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "61");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "-293");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "\"EVUOZ\"");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "43");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "\"R\"");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "3");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "249");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-445.67");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "-42");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "50");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "73");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "-57");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=14");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=AND(SUM(-26, 28) > 0, MIN(D4, 27) < 100)");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=(D4 + (-1 / B5))");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=AVERAGE(D2:D4)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=1");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=OR(ROUNDUP(B4, 2) > 0, C6 < 100)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "\"l\"");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=(A6 + D4)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=AVERAGE(INT(C3), UPPER(\"B5\"))");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=40");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=LOWER(\"B5\")");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=SQRT(B6)");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=C6");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=(ROUNDDOWN(B4, 0) ^ C1)");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=UPPER(\"E4\")");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=SUM((E1 + C2), E5)");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "-92");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=48");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=LEN(\"15\")");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=(C2 / AND(D5 > 0, E5 < 100))");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=LEFT(\"OR(-28 > 0, A7 < 100)\", 1)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=SQRT(IF((43 > 28), D6, C7))");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "314.56");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "-89");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "4");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 1));
    println!("Seed 233700 target: {:?}", target);
    match target { ResultData::Float(f) => assert!((f - 7.810249675906654).abs() < 1e-4), ResultData::Integer(i) => assert_eq!(i, 7), other => panic!("Expected 7.810249675906654, got {:?}", other) }
}
#[test]
fn test_fuzz_round_addition_val() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "-464");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "-69");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-89");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "272.45");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "83");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "28");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "-84");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "-358.99");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "\"3L\"");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "81");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-45.6497");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "-17");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "1");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "-64");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-20.5752");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "67");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "\"sc \"");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "75");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "7");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "\"bgQrkP2\"");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=A5");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=SUM(A3:D5)");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=OR(UPPER(\"C5\") > 0, (D3 - A2) < 100)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=MIN(C3:E5)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=PRODUCT(ROUND(D2, 1), AVERAGE(A2:A5))");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=(ROUNDUP(D4, 0) ^ B6)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=ROUNDUP((18 ^ B2), 1)");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=C3");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "7");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=LEFT(\"OR(C2 > 0, -49 < 100)\", 3)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=(-21 - ABS(C3))");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "-57");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=D4");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=MAX(D1:D3)");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=A6");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "201.3376");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=ROUND(LEFT(\"-47\", 1), 1)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "115.5");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=B4");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=(ROUND(C6, 2) + D4)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=ROUNDUP(B9, 2)");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=ROUNDUP(9, 0)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=AVERAGE(E1:E2)");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    println!("Seed 61467 target: {:?}", target);
    match target { ResultData::Float(f) => assert_eq!(f, 1.0), ResultData::Integer(i) => assert_eq!(i, 1), other => panic!("Expected 1, got {:?}", other) }
}
#[test]
fn test_fuzz_sum_range_string_literal_ignore() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "3");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "7");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "2");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-75");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "\"2AY\"");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-20");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "62");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "85");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "\"Hu\"");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "80");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "94");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-20");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "2");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "\"EgCWnL2J\"");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "-27");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "\"3\"");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "\"sCr\"");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "39");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=OR(PRODUCT(C1, B4) > 0, SQRT(E3) < 100)");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=SUM(A5:A5)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "71");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=A4");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=A2");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=26");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "\"mCzRh\"");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=UPPER(\"E5\")");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=ABS((E1 ^ -46))");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=RIGHT(\"(E4 * D1)\", 3)");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=AVERAGE(LEN(\"C3\"), D1)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=AND(ROUNDUP(E6, 1) > 0, PRODUCT(D7:E7) < 100)");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=(C4 + CONCATENATE(\"39\", \"7\"))");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=IF((12 > C1), AND(C7 > 0, E1 < 100), (D5 / A6))");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=ROUNDUP(A2, 1)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=SUM(A3:B5)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=C6");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=D3");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=(B1 ^ A3)");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=IF((INT(11) > CONCATENATE(\"D5\", \"-26\")), SUM(E7, A6), AND(E2 > 0, C1 < 100))");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=B2");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "482.2");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(5, 2));
    println!("Seed 965939 target: {:?}", target);
    match target { ResultData::Float(f) => assert_eq!(f, 0.0), ResultData::Integer(i) => assert_eq!(i, 0), other => panic!("Expected 0, got {:?}", other) }
}
#[test]
fn test_fuzz_if_sqrt_range_cond_val() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "91");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "\"U\"");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "11");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "57");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "\"DdRE\"");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "-27");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "-73.229");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "-50");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "\"EY 1k\"");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "34");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-100");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "60");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "\"Kq\"");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "27");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "228.9");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "1");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "-21");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "-254.6356");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "-42");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "-62");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "\"pCEeIM\"");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=MIN(C1:E4)");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=(AND(17 > 0, C5 < 100) * IF((-3 > E1), E2, A3))");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=B5");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=C4");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=OR(-26 > 0, ROUNDDOWN(A2, 1) < 100)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=-38");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=C6");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "24");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=D5");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=22");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=AND(IF((44 > -16), C3, E7) > 0, INT(-36) < 100)");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=ABS(E4)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=-29");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=C1");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "-85");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=-46");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=C1");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=LEN(\"PRODUCT(A3:A3)\")");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "242.9");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=ROUND(UPPER(\"B9\"), 0)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=IF((SQRT(C8) > E3), A2, 17)");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=ROUND(C5, 0)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "-406");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=0");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 1));
    println!("Seed 839355 target: {:?}", target);
    match target { ResultData::Float(f) => assert_eq!(f, 17.0), ResultData::Integer(i) => assert_eq!(i, 17), other => panic!("Expected 17, got {:?}", other) }
}
#[test]
fn test_fuzz_if_upper_string_comparison_val() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "65");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-97");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "55");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "5");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "\"JZOdb1\"");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "38");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "-357");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "\"hWxw\"");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "242.9694");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "-98");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "\"RV\"");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-238.682");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "-87");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "\"G\"");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "26.52");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "-14");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "495.2726");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "-12");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "\"1QpSKs\"");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=(B1 ^ 21)");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=ABS(AVERAGE(A4, B5))");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=30");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=IF((ROUNDUP(D5, 1) > UPPER(\"B4\")), PRODUCT(E4:E4), 35)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=LOWER(\"IF((E1 > 32), C1, E1)\")");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=E1");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=IF((PRODUCT(B5:B6) > E4), CONCATENATE(\"D5\", \"C3\"), C2)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=E6");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=C3");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=IF((OR(15 > 0, -15 < 100) > IF((-34 > 35), E7, C1)), D2, AND(A3 > 0, B7 < 100))");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "-496.535");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "-22");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=C4");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=10");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=CONCATENATE(\"D2\", \"12\")");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=LEN(\"A3\")");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=D6");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=IF((-24 > UPPER(\"A4\")), E1, ROUNDDOWN(E7, 2))");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=C8");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=A5");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=PRODUCT(D4:E4)");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=ROUNDDOWN(SQRT(C8), 1)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=IF((-20 > (B8 + E7)), (-11 + A5), SUM(D1, A2))");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=LOWER(\"D3\")");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 3));
    println!("Seed 242788 target: {:?}", target);
    match target { ResultData::Float(f) => assert_eq!(f, -357.0), ResultData::Integer(i) => assert_eq!(i, -357), other => panic!("Expected -357, got {:?}", other) }
}
#[test]
fn test_fuzz_if_sqrt_len_branch_val() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "-330.75");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "-39");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "\"Jkspzj\"");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "-53");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "\"Mvh\"");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "-323.98");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "\"UeAkj\"");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "57");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "1");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "414.87");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "-51");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "62.04");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "52");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "6");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "83");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "486.6446");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "-83");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "-46");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=MIN(AND(D3 > 0, A2 < 100), IF((A4 > -20), 2, D5))");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=ROUNDDOWN(E5, 1)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=D4");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "-345");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=C5");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=AND(MIN(D3, D6) > 0, E5 < 100)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=E4");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=(LOWER(\"C3\") / E3)");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=AVERAGE(B5, 45)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=LEFT(\"A7\", 3)");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=-19");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=PRODUCT(A3, A5)");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=LEN(\"IF((B1 > E3), -42, E2)\")");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=((B2 + B7) ^ LOWER(\"D2\"))");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=IF((B7 > SQRT(D6)), LEN(\"D7\"), C3)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=ROUNDUP(MIN(D8:D8), 0)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "-35");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "2");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=D8");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=C4");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=AVERAGE(A3:E5)");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 2));
    println!("Seed 738552 target: {:?}", target);
    match target { ResultData::Float(f) => assert_eq!(f, 2.0), ResultData::Integer(i) => assert_eq!(i, 2), other => panic!("Expected 2, got {:?}", other) }
}
#[test]
fn test_fuzz_or_multiplication_string_error_order() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "281.41");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "67.274");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "-21");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "46");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-424.437");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "\"3qUWHp\"");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "\"1C3\"");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "12.5");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "2");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "59");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "2");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "-199.85");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "85");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "64");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "\"kuG3fTA\"");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "9");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=D5");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=-25");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=OR(-45 > 0, E5 < 100)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=OR((0 * D2) > 0, -20 < 100)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=E4");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=(SUM(D6, B6) - ROUND(D2, 1))");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=ROUNDDOWN(A6, 1)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=(D1 * AVERAGE(B2:B6))");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=ROUNDUP(AND(E5 > 0, A6 < 100), 0)");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=B3");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=(ROUNDDOWN(6, 0) ^ AND(-27 > 0, -38 < 100))");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=B5");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=A3");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=C8");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=C1");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "-489.158");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=C7");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=IF((MAX(C4:D5) > E4), E2, LOWER(\"B2\"))");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=ROUNDUP((B7 + C4), 0)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=MIN(D8:D8)");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "17");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "63");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "21");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(5, 4));
    println!("Seed 574846 target: {:?}", target);
    match target { ResultData::Error(e) => assert_eq!(e, "#VALUE!"), other => panic!("Expected #VALUE!, got {:?}", other) }
}
#[test]
fn test_fuzz_multiplication_string_div_by_zero_precedence() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "37");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "-148.27");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "261.06");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "68");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "179");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "\"QUR\"");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "-73");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "\"vyeKi\"");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "44");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "\"LuoHWK\"");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "59.38");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "61");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "\"kpGlOd\"");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "\"Xtz1\"");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "-46.4");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "-388.1504");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "17");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "42");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "-80");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "3");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "-64");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=A1");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=INT(-22)");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=-37");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=C2");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=E1");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "-279");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=PRODUCT((B2 - C3), A4)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=(D3 * (E3 / A3))");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=((D4 - E2) - B3)");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=ABS(E3)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=MIN(A5, ABS(D2))");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=29");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=-13");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=D2");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=LEN(\"B3\")");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=LEFT(\"AND(22 > 0, 14 < 100)\", 1)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=ROUND(AND(E5 > 0, E6 < 100), 0)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=C6");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=INT(LEFT(\"-41\", 2))");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=PRODUCT(LEN(\"D3\"), ABS(-48))");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=(E7 * E5)");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=INT((D2 + A8))");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=-3");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=IF((LEN(\"C5\") > E6), 34, LOWER(\"E8\"))");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 3));
    println!("Seed 224583 target: {:?}", target);
    match target { ResultData::Error(e) => assert_eq!(e, "#VALUE!"), other => panic!("Expected #VALUE!, got {:?}", other) }
}
#[test]
fn test_fuzz_or_round_string_error_order() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "50");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "4");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "55");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-417");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "-58");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "55");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "346.94");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "\"Y ggk\"");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "76");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "-422.647");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "7");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "\"TtF\"");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "46");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "269.1");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "\"gG\"");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "-41");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "361.685");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "-198.7");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "3");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "-71.2086");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "221.5164");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=(LEFT(\"B2\", 5) * C3)");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=(-22 - ROUND(25, 0))");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=B2");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=LEFT(\"ROUNDDOWN(13, 1)\", 1)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=SQRT(A5)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "\"zMyIO\"");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=AVERAGE(A6:D6)");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=3");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=27");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=A4");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=A2");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=30");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=IF((LOWER(\"A4\") > OR(E6 > 0, -35 < 100)), INT(30), LEN(\"E5\"))");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "-7");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "46");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=C1");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=B2");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=(SQRT(19) - E7)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=OR(OR(B8 > 0, D1 < 100) > 0, ROUND(A4, 1) < 100)");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=(-15 - 20)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=RIGHT(\"ROUND(-48, 2)\", 4)");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=21");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=A3");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=(D3 ^ D3)");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 4));
    println!("Seed 422835 target: {:?}", target);
    match target { ResultData::Error(e) => assert_eq!(e, "#VALUE!"), other => panic!("Expected #VALUE!, got {:?}", other) }
}
#[test]
fn test_fuzz_division_if_branch_val() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "77");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "-57");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "14");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "\"Rj\"");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "47");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "\"LLu\"");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "\"ljDQrJ\"");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "\"slrI\"");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "\"gXjEz\"");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-415.525");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "-52");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "10");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "-68");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "-17");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "\"J\"");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "53");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "16");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "289.77");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "-82");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "24.2276");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=-46");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=(B5 - A4)");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=-45");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=ABS(SUM(D3:E3))");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=LEN(\"A2\")");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=(ABS(D6) / IF((B5 > E6), -22, D5))");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=-39");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=17");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=OR(SQRT(27) > 0, MIN(C3:E6) < 100)");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=OR((43 ^ 24) > 0, C1 < 100)");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=33");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=D1");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=SUM(MIN(A6:A6), (13 / E6))");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=3");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=ABS(ROUND(E3, 0))");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=AND(SQRT(-42) > 0, -25 < 100)");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=(-49 + OR(D1 > 0, D5 < 100))");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=SUM(C8:E8)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=((-25 * B1) ^ MIN(A4, B5))");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=((E3 / A2) / C4)");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=ROUNDUP((E6 - -31), 2)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=MIN(E2:E8)");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=OR(E9 > 0, UPPER(\"D1\") < 100)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=17");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=D6");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 0));
    println!("Seed 879868 target: {:?}", target);
    match target { ResultData::Float(f) => assert!((f - (-21.251136363636363)).abs() < 1e-4), other => panic!("Expected -21.251136363636363, got {:?}", other) }
}
#[test]
fn test_fuzz_if_rounddown_boolean_subtraction_val() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "\"dZGZ\"");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "-50");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "4");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "85");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "67");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-41");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "67");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "-75");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "459.99");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "6");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "\"O1\"");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "21");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "-66");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "132");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-97");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "44");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "-466.6791");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "\"N\"");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "-41");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "118.8");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "\"mLFdskj\"");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=-43");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=ROUNDUP(SQRT(E3), 1)");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=MAX(C5:C5)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "86");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=LEFT(\"E5\", 3)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=LEFT(\"(D3 / E5)\", 5)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "183");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=AND((A5 + A3) > 0, IF((C4 > -41), -5, C2) < 100)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "-39");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=E1");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=PRODUCT(D2:D7)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=AND(B1 > 0, (E1 * E1) < 100)");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=C2");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=B6");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=E2");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=C3");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=(E1 + A3)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=10");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=AND(UPPER(\"E2\") > 0, D8 < 100)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=(MAX(A8:D8) * D2)");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=-18");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "-143");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=ROUNDDOWN(OR(-26 > 0, 34 < 100), 2)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=IF((A4 > ROUNDDOWN(E8, 1)), (E3 - E8), B2)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=OR(37 > 0, MIN(C4:E6) < 100)");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 3));
    println!("Seed 353982 target: {:?}", target);
    match target { ResultData::Float(f) => assert_eq!(f, 75.0), ResultData::Integer(i) => assert_eq!(i, 75), other => panic!("Expected 75, got {:?}", other) }
}
#[test]
fn test_fuzz_sqrt_product_len_val() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "-92");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "\"aoqjZito\"");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "\"SaJyt\"");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-419.51");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-16");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "2");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "21");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "75");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "10");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "4");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "79");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "134.7696");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "46");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "72");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "24");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "9");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "\"s\"");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "57.51");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "37");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=22");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=D4");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=-40");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=PRODUCT(LEN(\"C4\"), 30)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=PRODUCT(D5:D5)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=(E1 ^ D2)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=ROUNDUP(-32, 0)");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=E6");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=MAX(24, AVERAGE(B1:B1))");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=D2");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=SQRT(D6)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=CONCATENATE(\"ROUNDDOWN(C2, 1)\", \"IF((B3 > -13), C7, D1)\")");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=((C2 ^ 15) - 29)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=D1");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=-8");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=D5");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=IF((C7 > AND(50 > 0, D5 < 100)), E8, MAX(E6:E6))");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=SQRT(ROUND(A7, 0))");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=-14");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=IF((A8 > (D2 - B7)), 40, LEFT(\"D1\", 5))");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=(MAX(B5, 25) / IF((D9 > B8), -34, E6))");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=(B9 + AND(C2 > 0, E2 < 100))");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=D3");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=AVERAGE(B6:E7)");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 0));
    println!("Seed 620565 target: {:?}", target);
    match target { ResultData::Float(f) => assert!((f - 7.745966692414834).abs() < 1e-4), other => panic!("Expected 7.745966692414834, got {:?}", other) }
}
#[test]
fn test_fuzz_sqrt_min_cell_ref_string_ignore() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "8");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "7");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "-42");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-306.1794");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "2");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-82");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "\"FaCgqE\"");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "-189.2");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "\"xM3\"");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "-106");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "80");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "\"F\"");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "\"n1q\"");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-78");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "-137");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "-397");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-55");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "7");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "299.43");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "\"ER\"");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "-62");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "92");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "30.4145");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "-492");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=LOWER(\"(A5 / B3)\")");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "-11");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=((D5 / D3) + SUM(D5:E5))");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=ROUNDDOWN(-12, 1)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=MIN((E1 ^ E5), B2)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=LEN(\"SUM(C4:C6)\")");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=ABS(E3)");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "-393.635");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=38");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=E4");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=E5");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=LEFT(\"D7\", 1)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=E7");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=((A7 - C1) + ROUNDUP(B6, 0))");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=B5");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=ROUNDDOWN(ABS(E3), 0)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=AVERAGE(C7, PRODUCT(D3:D5))");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=OR(E5 > 0, SUM(E7:E7) < 100)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=(C2 ^ PRODUCT(A6:E7))");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=B2");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=D2");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=PRODUCT(D2, (C7 * B2))");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=(SQRT(E6) + 31)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "\"m1\"");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 3));
    println!("Seed 874653 target: {:?}", target);
    match target { ResultData::Float(f) => assert_eq!(f, 31.0), ResultData::Integer(i) => assert_eq!(i, 31), other => panic!("Expected 31, got {:?}", other) }
}
#[test]
fn test_fuzz_if_sqrt_roundup_empty_cell_ref() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "-213.0018");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "19");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-308.3751");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "58");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "-97");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "\"oMNfOA\"");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "\"BkwUOw\"");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "12");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-47");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "-26");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "7");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-5");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "-485.31");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "-335.3205");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "-60");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "-11");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "24");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=5");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=ROUNDDOWN((E4 * 3), 2)");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=AND(AND(-15 > 0, A4 < 100) > 0, 11 < 100)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=C4");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=E3");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=D3");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=C2");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=B2");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "3");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=C7");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=10");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=RIGHT(\"INT(-24)\", 5)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=A3");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=ROUNDUP(C5, 1)");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=INT((C1 * -39))");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=B4");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=LOWER(\"A6\")");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=E6");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=IF((E7 > 20), (21 * C5), SQRT(E8))");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=PRODUCT(LEFT(\"D5\", 5), (C2 - -29))");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=-8");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=B7");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=LEFT(\"C7\", 2)");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    println!("Seed 829563 target: {:?}", target);
    match target { ResultData::Float(f) => assert_eq!(f, 0.0), ResultData::Integer(i) => assert_eq!(i, 0), other => panic!("Expected 0, got {:?}", other) }
}
#[test]
fn test_fuzz_if_lower_string_constant_branch() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "-129.6");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "\"gHcl\"");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "73.1585");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "67");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "33");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "102.9");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "\"MfFR\"");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "-411.1429");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "-405.31");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "8");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "36.8849");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "198.4695");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "59");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "-18");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "\"vXut1lNE\"");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "-60");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "16");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "-450.9472");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "108.318");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "-30");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "\"GvWzHSI\"");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=A3");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=B1");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "-149.5");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=C1");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=B2");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=IF((SUM(E6, A1) > ROUND(D6, 0)), LOWER(\"14\"), (-8 - -10))");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=(IF((-33 > C6), C4, C6) + D1)");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=AND(E1 > 0, E1 < 100)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=12");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=SUM(A6:D6)");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=RIGHT(\"A3\", 2)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=ROUND(RIGHT(\"-45\", 4), 0)");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "-29");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=ABS(ROUND(-17, 0))");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=C7");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=B7");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=(B6 + -37)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=(B1 + 15)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=AVERAGE(C5:E7)");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=-18");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=(B2 / UPPER(\"B6\"))");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=LEFT(\"AND(34 > 0, B9 < 100)\", 3)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=(1 - D3)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=-20");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 0));
    println!("Seed 419816 target: {:?}", target);
    match target { ResultData::Float(f) => assert_eq!(f, 2.0), ResultData::Integer(i) => assert_eq!(i, 2), other => panic!("Expected 2, got {:?}", other) }
}
#[test]
fn test_fuzz_if_average_range_empty_cell_ref() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "14");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "\"qu RkdT\"");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "\"pZ\"");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "-41");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-64.7");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "-94");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "\"scJWUwM\"");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "\"ii\"");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "\"RjekSLLy\"");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "37");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "223.7869");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-56");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "89");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "177.6816");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-51");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "-399.73");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "-392.5864");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "-20");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=UPPER(\"IF((E2 > D3), E5, 32)\")");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "-82");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=OR(OR(B1 > 0, E2 < 100) > 0, SUM(A4, 38) < 100)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=9");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=E4");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=OR(ROUNDUP(C1, 0) > 0, D5 < 100)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=15");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=B4");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=CONCATENATE(\"IF((1 > A2), B3, E3)\", \"PRODUCT(E6:E6)\")");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=A2");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=IF((ABS(E7) > D1), AVERAGE(C6:C6), SUM(A7:A7))");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=MAX(B5:B6)");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=OR(SUM(12, 2) > 0, LEN(\"C5\") < 100)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=(-13 / B4)");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=LEFT(\"SUM(44, D1)\", 3)");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "\"GdRInUq\"");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=IF((ROUNDDOWN(B8, 1) > E6), C8, (14 + A2))");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=AVERAGE(E4:E8)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=IF((D5 > C3), SUM(A7:A7), B1)");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=CONCATENATE(\"AND(E3 > 0, C1 < 100)\", \"AVERAGE(C7:D7)\")");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=IF((-21 > D7), LEN(\"17\"), OR(46 > 0, D4 < 100))");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=-34");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=B1");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=MIN(A5, UPPER(\"-6\"))");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 0));
    println!("Seed 234176 target: {:?}", target);
    match target { ResultData::Float(f) => assert_eq!(f, 0.0), ResultData::Integer(i) => assert_eq!(i, 0), other => panic!("Expected 0, got {:?}", other) }
}
#[test]
fn test_fuzz_round_cell_ref_float_val() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "-481.188");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "459.798");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-31");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "\"qXxUInB\"");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "\"G\"");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "80.01000000000001");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "\"LDzgIKy\"");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "-11");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "-98");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "60");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-1");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "-12");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "\"BjZ\"");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-23.5");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "69");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "70");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "\"Vy\"");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "\"H\"");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "-76");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "-56.95");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=IF((D2 > (-16 + C3)), C4, C4)");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "\"2Lb\"");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=B5");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=((C3 + 28) + IF((B2 > 40), B1, 17))");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=IF((LOWER(\"B4\") > E3), 6, C3)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=26");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=D3");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=C2");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=D4");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=C1");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=ROUNDUP(LEN(\"B2\"), 0)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=AND((A1 * E3) > 0, E7 < 100)");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=LEN(\"A3\")");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=ROUNDUP(D3, 2)");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "-60");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=CONCATENATE(\"AVERAGE(C2, B7)\", \"AVERAGE(A2, -7)\")");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=IF((LEN(\"E8\") > E2), B5, RIGHT(\"C5\", 5))");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=(14 - MIN(B4, A8))");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=ROUND(E7, 0)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=21");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=(C2 * B9)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=AND(LEFT(\"C8\", 5) > 0, (A2 / C9) < 100)");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=(A2 * 10)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=B9");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=PRODUCT(ABS(B8), ROUNDDOWN(D2, 0))");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 3));
    println!("Seed 976519 target: {:?}", target);
    match target { ResultData::Float(f) => assert_eq!(f, 460.0), ResultData::Integer(i) => assert_eq!(i, 460), other => panic!("Expected 460, got {:?}", other) }
}
#[test]
fn test_fuzz_product_concatenate_left_to_right_value_error() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "70");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "375");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "-12");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "60");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "\"3QGdh\"");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "360.197");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "-362");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "59");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "-46");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-24");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "\"H\"");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "-25");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "82");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "-76");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "-73");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "\"oCByq\"");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "89");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=LOWER(\"ROUND(A4, 1)\")");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "61");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=(SQRT(-1) * 47)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=LEN(\"-30\")");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=-36");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=IF((LOWER(\"B1\") > IF((-22 > C3), 3, E2)), 16, A1)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=ABS(A3)");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=C5");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=ROUNDDOWN(AND(B3 > 0, A3 < 100), 1)");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=MIN(B6:E6)");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=C6");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=C2");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=AVERAGE(-49, ROUNDUP(A6, 2))");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "65");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=B7");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=-19");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=B1");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=PRODUCT(CONCATENATE(\"D4\", \"E5\"), IF((A8 > B8), A5, A3))");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=E6");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=B3");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=ROUNDDOWN(SQRT(A4), 1)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=10");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=IF((E3 > D2), PRODUCT(E5:E9), E2)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=INT(E6)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "-175");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 2));
    println!("Seed 158215 target: {:?}", target);
    match target { ResultData::Error(e) => assert!(e == "#VALUE!" || e == "#NUM!"), other => panic!("Expected #VALUE! or #NUM!, got {:?}", other) }
}
#[test]
fn test_fuzz_abs_div_by_zero_subtraction_error() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "14");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "-424");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "-22");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "319.113");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "22");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "215.3");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "-71");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "-247.8");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "-384.719");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "89");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "6");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "-88");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "\"ccZu2\"");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "227.4153");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "-490.9");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "167.2");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "-76");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "54");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "-76");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "-73");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=LEN(\"-30\")");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=48");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=ABS(C1)");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=C1");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=AND(IF((A2 > A1), B2, E5) > 0, (D2 ^ D2) < 100)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=IF((LEN(\"B3\") > RIGHT(\"E6\", 1)), SQRT(A6), (-17 * E5))");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=AVERAGE(LOWER(\"B5\"), 24)");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=MIN(C4:C5)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=LOWER(\"45\")");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=-21");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "1");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=B5");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "\"WInX2\"");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=(E1 / D2)");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=C1");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=A5");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=(ABS(E8) - A6)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=C4");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=ABS(B1)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=16");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=RIGHT(\"B6\", 2)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=INT(ABS(B3))");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=C4");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 2));
    println!("Seed 817752 target: {:?}", target);
    match target { ResultData::Error(e) => assert_eq!(e, "#DIV/0!"), other => panic!("Expected #DIV/0!, got {:?}", other) }
}
#[test]
fn test_fuzz_negative_base_huge_exponent_num_error() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "-41");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "199.663");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-421.675");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "17");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "-36.3925");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "316.534");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "-30");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "6");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "154");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "\"2Lpn pQ\"");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "-91");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-28");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "\"G\"");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "-73");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "-84");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "64");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "-37");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "105.398");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=(E3 - INT(E5))");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "10");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=C2");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "-94");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=AND(MIN(B1:C1) > 0, B4 < 100)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=((B1 - A6) + ABS(D3))");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=A6");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=C6");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=LEN(\"E3\")");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=PRODUCT(D1:E6)");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=UPPER(\"D2\")");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=45");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=D5");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=ROUNDDOWN(UPPER(\"E7\"), 1)");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=SQRT(27)");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=D4");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=E3");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=AND(IF((C6 > A8), D2, -43) > 0, AND(C5 > 0, A6 < 100) < 100)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=ROUND(B4, 2)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=LEN(\"C7\")");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=D5");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=(D4 * B5)");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=E4");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=(-26 ^ E7)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "\"ptjbR1\"");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 3));
    println!("Seed 128903 target: {:?}", target);
    match target { ResultData::Error(e) => assert_eq!(e, "#NUM!"), other => panic!("Expected #NUM!, got {:?}", other) }
}
#[test]
fn test_fuzz_power_roundup_rounddown_val() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "-12.1");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "-22");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "34");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "58");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-17");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "98");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "\"ssCuwts\"");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "-175.8");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "-135.4");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "-150.079");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "45");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "115.15");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-85");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "40");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "7");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "-67");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "-19");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "-48");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "49");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "\"hB\"");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "-34");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "-3");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "-20");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=PRODUCT(-50, E4)");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=((18 * 18) - IF((-1 > 46), E1, B5))");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=PRODUCT((A5 * A3), PRODUCT(-43, C5))");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=IF((IF((B5 > D1), B1, E2) > B3), MIN(A2:A5), E1)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=SQRT(OR(E3 > 0, D3 < 100))");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "331.33");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=-3");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=SQRT(B3)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=LEN(\"OR(A6 > 0, 8 < 100)\")");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=B1");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=((E4 - A1) - D2)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=RIGHT(\"9\", 1)");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "-60");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=-14");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=ROUNDDOWN(ROUND(B1, 1), 1)");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=A4");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=AVERAGE(IF((B1 > C1), D5, A4), D6)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=14");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=LOWER(\"-44\")");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=LOWER(\"13\")");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=(ROUNDUP(C9, 1) ^ ROUNDDOWN(17, 0))");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=B8");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=AND(IF((D7 > 42), E3, D8) > 0, (D8 ^ B6) < 100)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=IF((E8 > ROUNDDOWN(A4, 2)), 43, E3)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=E7");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 0));
    println!("Seed 158317 target: {:?}", target);
    match target { ResultData::Float(f) => assert!((f - 30491346729331195904.0).abs() < 1e10), ResultData::Integer(i) => assert_eq!(i, 30491346729331195904_i128 as i64), other => panic!("Expected 30491346729331195904, got {:?}", other) }
}
#[test]
fn test_fuzz_if_sqrt_average_range_num_error() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "\"pNOEESaj\"");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "94");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "\"dp3sTYc\"");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-344.363");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "\"zyNCSI\"");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "6");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "44");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "58.4");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "37");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "\"nNgEFl\"");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "5");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "-295.431");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "-240.2");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "19");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "\"IxGnv3u\"");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "25");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "-74");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "2");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "-87");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "-84");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=OR(B4 > 0, ROUNDUP(-19, 2) < 100)");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "-352.4");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "94");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=AVERAGE(C2:E5)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=ROUNDUP(ABS(D1), 1)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "-59");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "63");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=PRODUCT(B6:B6)");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=B4");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=IF((SQRT(D6) > C2), (A1 / -23), IF((D5 > E5), -6, A7))");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=SQRT(-6)");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=SQRT(30)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=(ROUND(20, 0) + B3)");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=ABS(34)");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "312.486");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=ROUNDUP((17 + E7), 2)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=AND(CONCATENATE(\"C5\", \"43\") > 0, B2 < 100)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=(A8 * A7)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "-2");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=INT(SQRT(B6))");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=25");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=(IF((E1 > A5), A9, -42) * UPPER(\"D5\"))");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=(IF((-17 > B5), E5, B1) * -14)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=ROUND(UPPER(\"D8\"), 2)");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 0));
    println!("Seed 631817 target: {:?}", target);
    match target { ResultData::Error(e) => assert_eq!(e, "#NUM!"), other => panic!("Expected #NUM!, got {:?}", other) }
}
#[test]
fn test_fuzz_sqrt_average_negative_num_error() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "478.67");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "-79");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "-63");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-95");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "33");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "51");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "-99");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "\"f\"");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "\"yEF\"");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "-14");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "\"Vd\"");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-150.0543");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "-496.7");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "477.525");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "105.9347");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "-38");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "8");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "81");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=A3");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=MAX(B5:C5)");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "-394.489");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=IF((E5 > IF((B4 > 47), 2, B5)), D3, (D4 - E5))");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=AND(D3 > 0, IF((D3 > E1), 28, 29) < 100)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=(23 * RIGHT(\"E6\", 4))");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=D3");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=SQRT(D6)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=IF(((A1 + E3) > (-15 + -26)), ROUNDUP(-43, 1), B4)");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=C6");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=C6");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=B2");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=28");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=MAX(LEN(\"3\"), C6)");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=5");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=B5");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=AVERAGE(OR(E7 > 0, D5 < 100), IF((C6 > D6), 41, B6))");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=E8");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=LEN(\"30\")");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=SUM(D2:D3)");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=LOWER(\"(C1 / B8)\")");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=34");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=LEFT(\"INT(-5)\", 4)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=C4");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 2));
    println!("Seed 916212 target: {:?}", target);
    match target { ResultData::Error(e) => assert_eq!(e, "#NUM!"), other => panic!("Expected #NUM!, got {:?}", other) }
}
#[test]
fn test_fuzz_sqrt_rounddown_negative_num_error() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "-113.77");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "85");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-60");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "85");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "-282.7737");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "\"S1vRUvgQ\"");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "\"FA\"");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "-57");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "407");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "-8");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "65");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "390.8202");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-11.7");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "18.0384");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "41");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "\"ZpGcDI2d\"");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "-206");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "\"z\"");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "-86");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "-241");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=46");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=-39");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=A4");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=UPPER(\"SUM(C1, 41)\")");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=33");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=SQRT(LEN(\"E1\"))");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=SQRT(A6)");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=OR(D4 > 0, -4 < 100)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=AND(A6 > 0, A2 < 100)");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=IF((MAX(C3:C3) > LEN(\"20\")), IF((32 > E1), E2, E3), 1)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=B1");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=E5");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=ABS(-18)");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=D7");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=IF((LEN(\"-24\") > PRODUCT(22, B7)), SUM(A2, -30), MIN(C6:E7))");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=C1");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=AVERAGE(LEN(\"A6\"), LOWER(\"C1\"))");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=D1");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=E3");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=MAX(D8:E8)");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=B7");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=SQRT(D9)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=C6");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=(D7 * C6)");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 2));
    println!("Seed 812174 target: {:?}", target);
    match target { ResultData::Error(e) => assert_eq!(e, "#NUM!"), other => panic!("Expected #NUM!, got {:?}", other) }
}
#[test]
fn test_fuzz_positive_base_power_underflow_zero() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "56");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "\"ovmGQ\"");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "2");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "214.841");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "\" UcRcJk\"");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "-28");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "22.02");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "351.7");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "499.82");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "-307.6667");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "\"gVA3\"");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "0");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "311.7");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "44");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "399");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "2");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "-24");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "-166.6152");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=(B1 / -26)");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "\"sv2XVCLY\"");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=C2");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=SUM(AND(B5 > 0, A5 < 100), A3)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=MAX(B4:C5)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=CONCATENATE(\"ROUNDUP(E2, 2)\", \"IF((C6 > B2), E5, -37)\")");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=(C4 - -1)");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=((D6 ^ -39) ^ 48)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=B4");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=ROUNDDOWN(LOWER(\"-34\"), 0)");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=MIN(15, C3)");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "91");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=SUM(B2:C3)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=LOWER(\"B7\")");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=AVERAGE(D4:D6)");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=SQRT(PRODUCT(C7, 15))");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=LEN(\"A2\")");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=C2");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=-17");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=LOWER(\"D2\")");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=C4");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=-47");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=0");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=AND(SQRT(E4) > 0, -4 < 100)");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(6, 2));
    println!("Seed 522134 target: {:?}", target);
    match target { ResultData::Float(f) => assert_eq!(f, 0.0), ResultData::Integer(i) => assert_eq!(i, 0), other => panic!("Expected 0.0, got {:?}", other) }
}
#[test]
fn test_fuzz_string_gt_number_comparison_if() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "\"guyMN s \"");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "-89");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "12");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "86");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "-45");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "\"r\"");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "93");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "-28");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "\"mFf\"");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "372");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-70");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "\"aDcy\"");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "-202.887");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "\"le\"");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "45");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "-48");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "281");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "FALSE");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "\"nXAPy\"");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "\"bTJiL\"");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "-53");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "\" QZQ\"");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=IF((CONCATENATE(\"C1\", \"B1\") > E5), (13 * D3), -6)");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=PRODUCT(A2:B5)");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=CONCATENATE(\"A5\", \"E4\")");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=(2 / B1)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=IF((E3 > C5), MAX(C2:C3), E4)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=PRODUCT(E1:E2)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=AVERAGE(A1:D2)");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=PRODUCT(E6:E6)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=-25");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=(UPPER(\"E4\") * AND(5 > 0, -39 < 100))");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=A7");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=OR(OR(B7 > 0, E5 < 100) > 0, (E4 ^ A5) < 100)");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=50");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=IF((UPPER(\"-22\") > ROUND(E6, 0)), AND(B4 > 0, 24 < 100), RIGHT(\"D6\", 3))");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "-14.8777");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=A7");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=(A5 * 43)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=C6");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "-12");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=C4");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=CONCATENATE(\"44\", \"20\")");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=OR(SUM(B8:C8) > 0, (11 + A7) < 100)");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=ROUNDUP(UPPER(\"B5\"), 1)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=AVERAGE(A2:D6)");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=(-3 / E5)");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(7, 3));
    println!("Seed 691619 target: {:?}", target);
    match target { ResultData::Boolean(b) => assert!(b), other => panic!("Expected True, got {:?}", other) }
}
#[test]
fn test_fuzz_concatenate_scientific_string_0e5() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 0, char_offset: 0 }, "60");
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "\"S\"");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "-488.273");
    sheet.insert(TextCellRef { row: 0, col: 4, char_offset: 0 }, "\"qdY\"");
    sheet.insert(TextCellRef { row: 1, col: 0, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 1, col: 1, char_offset: 0 }, "\"kRZGH\"");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "20");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "6");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "-2");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "68");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "-3");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "-87");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "20");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "TRUE");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "-96");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "\"AbeWNuTT\"");
    sheet.insert(TextCellRef { row: 3, col: 3, char_offset: 0 }, "18");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "-50");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "77");
    sheet.insert(TextCellRef { row: 4, col: 1, char_offset: 0 }, "445.53");
    sheet.insert(TextCellRef { row: 4, col: 2, char_offset: 0 }, "-81");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "66");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "66");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=AVERAGE(C5:E5)");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=IF((C2 > E2), B2, (B5 + E4))");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=C1");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=(IF((D5 > E1), C3, A5) * RIGHT(\"-31\", 5))");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=PRODUCT(A1:A5)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=INT(MAX(A2, B6))");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=-23");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=D4");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=SQRT(ABS(-33))");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=-49");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=IF((-19 > 19), ABS(D7), SUM(B3, 25))");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=12");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=C2");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=(E7 * C4)");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=SQRT((B1 * E7))");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=47");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=-23");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=C4");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=AVERAGE(E2:E7)");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=CONCATENATE(\"0\", \"E5\")");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "=CONCATENATE(\"A6\", \"AND(D3 > 0, E9 < 100)\")");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=IF((IF((B9 > C4), D3, -12) > A3), AVERAGE(C8:E8), ABS(44))");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=ROUND(-13, 1)");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=-21");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=ABS(E5)");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(8, 4));
    println!("Seed 305840 target: {:?}", target);
    match target { ResultData::String(s) => assert_eq!(s, "0E5"), other => panic!("Expected String 0E5, got {:?}", other) }
}
#[test]
fn test_fuzz_roundup_min_float_precision_val() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 10,
        cols: 5,
        ..Default::default()
    });
    sheet.insert(TextCellRef { row: 0, col: 1, char_offset: 0 }, "473.2");
    sheet.insert(TextCellRef { row: 0, col: 2, char_offset: 0 }, "-89");
    sheet.insert(TextCellRef { row: 0, col: 3, char_offset: 0 }, "8");
    sheet.insert(TextCellRef { row: 1, col: 2, char_offset: 0 }, "6");
    sheet.insert(TextCellRef { row: 1, col: 3, char_offset: 0 }, "\"ozkAb\"");
    sheet.insert(TextCellRef { row: 1, col: 4, char_offset: 0 }, "-85");
    sheet.insert(TextCellRef { row: 2, col: 0, char_offset: 0 }, "58");
    sheet.insert(TextCellRef { row: 2, col: 1, char_offset: 0 }, "-23");
    sheet.insert(TextCellRef { row: 2, col: 2, char_offset: 0 }, "17");
    sheet.insert(TextCellRef { row: 2, col: 3, char_offset: 0 }, "38");
    sheet.insert(TextCellRef { row: 2, col: 4, char_offset: 0 }, "2");
    sheet.insert(TextCellRef { row: 3, col: 0, char_offset: 0 }, "400.1");
    sheet.insert(TextCellRef { row: 3, col: 1, char_offset: 0 }, "-92");
    sheet.insert(TextCellRef { row: 3, col: 2, char_offset: 0 }, "5");
    sheet.insert(TextCellRef { row: 3, col: 4, char_offset: 0 }, "-33");
    sheet.insert(TextCellRef { row: 4, col: 0, char_offset: 0 }, "-42");
    sheet.insert(TextCellRef { row: 4, col: 3, char_offset: 0 }, "-70");
    sheet.insert(TextCellRef { row: 4, col: 4, char_offset: 0 }, "56");
    sheet.insert(TextCellRef { row: 5, col: 0, char_offset: 0 }, "=0");
    sheet.insert(TextCellRef { row: 5, col: 1, char_offset: 0 }, "=D1");
    sheet.insert(TextCellRef { row: 5, col: 2, char_offset: 0 }, "=A5");
    sheet.insert(TextCellRef { row: 5, col: 3, char_offset: 0 }, "=OR(RIGHT(\"E1\", 4) > 0, LEFT(\"E1\", 4) < 100)");
    sheet.insert(TextCellRef { row: 5, col: 4, char_offset: 0 }, "=IF((MAX(E2:E5) > E2), PRODUCT(A4, A3), -4)");
    sheet.insert(TextCellRef { row: 6, col: 0, char_offset: 0 }, "=ROUND(46, 2)");
    sheet.insert(TextCellRef { row: 6, col: 1, char_offset: 0 }, "=IF((LOWER(\"B4\") > (B1 ^ C3)), D3, (D5 * B3))");
    sheet.insert(TextCellRef { row: 6, col: 2, char_offset: 0 }, "=AND(-3 > 0, IF((D6 > 50), A3, A6) < 100)");
    sheet.insert(TextCellRef { row: 6, col: 3, char_offset: 0 }, "=ROUND(37, 1)");
    sheet.insert(TextCellRef { row: 6, col: 4, char_offset: 0 }, "=(LOWER(\"14\") - IF((B2 > C4), C3, D2))");
    sheet.insert(TextCellRef { row: 7, col: 0, char_offset: 0 }, "=(A6 ^ IF((C3 > 15), B5, C5))");
    sheet.insert(TextCellRef { row: 7, col: 1, char_offset: 0 }, "=IF((AVERAGE(15, A7) > (D6 - A3)), D1, ROUNDUP(9, 1))");
    sheet.insert(TextCellRef { row: 7, col: 2, char_offset: 0 }, "=MIN(E1:E7)");
    sheet.insert(TextCellRef { row: 7, col: 3, char_offset: 0 }, "=IF((B5 > LEN(\"-8\")), -47, (47 * 20))");
    sheet.insert(TextCellRef { row: 7, col: 4, char_offset: 0 }, "=OR(46 > 0, (B2 + 8) < 100)");
    sheet.insert(TextCellRef { row: 8, col: 0, char_offset: 0 }, "=(MAX(E2:E4) ^ (42 / E6))");
    sheet.insert(TextCellRef { row: 8, col: 1, char_offset: 0 }, "=OR(AND(B4 > 0, B8 < 100) > 0, C2 < 100)");
    sheet.insert(TextCellRef { row: 8, col: 2, char_offset: 0 }, "=(E7 + -44)");
    sheet.insert(TextCellRef { row: 8, col: 3, char_offset: 0 }, "=A1");
    sheet.insert(TextCellRef { row: 8, col: 4, char_offset: 0 }, "=32");
    sheet.insert(TextCellRef { row: 9, col: 0, char_offset: 0 }, "-298.237");
    sheet.insert(TextCellRef { row: 9, col: 1, char_offset: 0 }, "=ROUNDUP(MIN(E6, E8), 1)");
    sheet.insert(TextCellRef { row: 9, col: 2, char_offset: 0 }, "=D9");
    sheet.insert(TextCellRef { row: 9, col: 3, char_offset: 0 }, "=(RIGHT(\"B8\", 5) - (-12 - A2))");
    sheet.insert(TextCellRef { row: 9, col: 4, char_offset: 0 }, "=MAX(D9:D9)");
    sheet.commit(None).unwrap();
    let target = sheet.get_result_data(&CellRef::new(9, 1));
    println!("Seed 854292 target: {:?}", target);
    match target { ResultData::Float(f) => assert!((f - 23205.8).abs() < 0.2), ResultData::Integer(i) => assert_eq!(i, 23205), other => panic!("Expected ~23205.8, got {:?}", other) }
}