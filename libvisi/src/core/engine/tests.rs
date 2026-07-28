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
            assert!(e.contains("ZeroDivisionError") || e.contains("division by zero"))
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
    if let Ok(bytes) = std::fs::read("fuzz_results/failures/fail_iter_5_seed_545786/source.xlsx") {
        if let Ok((sheets, _)) = crate::core::xlsx::import_xlsx_data(&bytes, &[], |_, _, _| {}) {
            let mut sheet = sheets[0].sheet.clone();
            sheet.commit(None).unwrap();
            let b10 = sheet.get_result_data(&CellRef::new(9, 1));
            println!("B10 evaluated: {:?}", b10);
            match b10 {
                ResultData::Float(f) => assert!((f - 64217.874).abs() < 1e-3, "Expected ~64217.874, got {}", f),
                other => panic!("Expected Float for B10, got {:?}", other),
            }
        }
    }
}

#[test]
fn test_fuzz_reproducer_seed_516067() {
    if let Ok(bytes) = std::fs::read("fuzz_results/failures/fail_iter_5_seed_516067/source.xlsx") {
        if let Ok((sheets, _)) = crate::core::xlsx::import_xlsx_data(&bytes, &[], |_, _, _| {}) {
            let mut sheet = sheets[0].sheet.clone();
            sheet.commit(None).unwrap();
            let b10 = sheet.get_result_data(&CellRef::new(9, 1));
            println!("B10 evaluated: {:?}", b10);
            match b10 {
                ResultData::Float(f) => assert_eq!(f, 216.0),
                other => panic!("Expected Float(216.0) for B10, got {:?}", other),
            }
        }
    }
}
