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
    test_floats("=CEILING(4.2)", 5.0).unwrap();
    test_floats("=FLOOR(4.8)", 4.0).unwrap();
    test_floats("=TAN(0)", 0.0).unwrap();
    test_floats("=ASIN(0)", 0.0).unwrap();
    test_floats("=ACOS(1)", 0.0).unwrap();
    test_floats("=ATAN(0)", 0.0).unwrap();

    test_booleans("=TRUE()", true).unwrap();
    test_booleans("=FALSE()", false).unwrap();
    test_booleans("=AND(TRUE(), TRUE)", true).unwrap();

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

fn approx_float(source: &str, expected: f64) {
    let sheet = Sheet::new(SheetInit::default());
    let (result, _) = sheet.eval(source, None).unwrap();
    let f = get_float_val(&result).unwrap_or_else(|| panic!("{source} did not return a number"));
    assert!((f - expected).abs() < 5e-3, "{source}: {f} != {expected}");
}

#[test]
fn test_financial_functions() {
    // These mirror the pure-math tests in `core::finance` but go through
    // the parser/dispatch path end-to-end.
    approx_float("=PMT(0.08/12, 10, 10000)", -1037.03);
    approx_float("=FV(0.06/12, 10, -200, -500, 1)", 2581.40);
    approx_float("=PV(0.08/12, 20*12, 500)", -59777.15);
    approx_float("=NPER(0.12/12, -100, -1000, 10000, 1)", 59.6739);
    approx_float(
        "=IPMT(0.10/12, 1, 36, 8000000) + PPMT(0.10/12, 1, 36, 8000000)",
        -258137.4976,
    );
    approx_float("=ISPMT(0.10/12, 1, 3*12, 8000000)", -64814.81481481482);
    approx_float("=NPV(0.10, -10000, 3000, 4200, 6800)", 1188.44);
    approx_float("=SLN(30000, 7500, 10)", 2250.0);
    approx_float("=SYD(30000, 7500, 10, 1)", 4090.91);
    approx_float("=DDB(2400, 300, 10, 1)", 480.0);
    approx_float("=EFFECT(0.0525, 4)", 0.0535427302);
    approx_float("=NOMINAL(0.0535427302, 4)", 0.0525);
    approx_float("=DOLLARDE(1.02, 16)", 1.125);
    approx_float("=DOLLARFR(1.125, 16)", 1.02);
    approx_float("=RRI(10, 1000, 2000)", 0.0717735);
    approx_float("=PDURATION(0.025, 2000, 2200)", 3.86045);
    approx_float("=CUMIPMT(0.09/12, 30*12, 125000, 13, 24, 0)", -11135.23);
    approx_float("=CUMPRINC(0.09/12, 30*12, 125000, 13, 24, 0)", -934.11);

    // No `{...}` array-literal syntax in the parser, so exercise the
    // range-argument financial functions (IRR/FVSCHEDULE/XNPV) against
    // real cells instead of inline arrays.
    let mut fin_sheet = Sheet::new(SheetInit {
        name: Some("fin".to_string()),
        rows: 6,
        cols: 4,
        ..Default::default()
    });
    for (i, v) in [-70000, 12000, 15000, 18000, 21000, 26000]
        .iter()
        .enumerate()
    {
        fin_sheet.set_cell_src(i, 0, v.to_string());
    }
    for (i, v) in [0.09, 0.11, 0.1].iter().enumerate() {
        fin_sheet.set_cell_src(i, 1, v.to_string());
    }
    for (i, v) in [-10000, 2750, 4250, 3250, 2750].iter().enumerate() {
        fin_sheet.set_cell_src(i, 2, v.to_string());
    }
    // Excel serials for 2008-01-01, 2008-03-01, 2008-10-30, 2009-02-15, 2009-04-01
    for (i, v) in [39448, 39508, 39751, 39859, 39904].iter().enumerate() {
        fin_sheet.set_cell_src(i, 3, v.to_string());
    }
    fin_sheet.commit(None).unwrap();

    let (irr_res, _) = fin_sheet.eval("=IRR(A1:A6)", None).unwrap();
    assert!((get_float_val(&irr_res).unwrap() - 0.0866).abs() < 5e-3);

    let (fvs_res, _) = fin_sheet.eval("=FVSCHEDULE(1, B1:B3)", None).unwrap();
    assert!((get_float_val(&fvs_res).unwrap() - 1.33089).abs() < 5e-3);

    let (xnpv_res, _) = fin_sheet.eval("=XNPV(0.09, C1:C5, D1:D5)", None).unwrap();
    assert!((get_float_val(&xnpv_res).unwrap() - 2086.65).abs() < 5e-2);
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
    test_strings("=CONCATENATE(\"Result: \", str(True))", "Result: TRUE").unwrap();
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
        context.add_table(sheet.name.clone(), sheet);
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
    #[allow(clippy::approx_constant)]
    let expected = 3.14; // ABS(-3.14), not an approximation of PI
    assert_eq!(get_float_val(&result), Some(expected));

    let (result, _) = sheet.eval("=sqrt(16)", None).unwrap();
    assert_eq!(get_float_val(&result), Some(4.0));

    let (result, _) = sheet.eval("=sin(0)", None).unwrap();
    assert_eq!(get_float_val(&result), Some(0.0));

    let (result, _) = sheet.eval("=cos(0)", None).unwrap();
    assert_eq!(get_float_val(&result), Some(1.0));

    let (result, _) = sheet.eval("=floor(3.9)", None).unwrap();
    assert_eq!(get_int_val(&result), Some(3));

    let (result, _) = sheet.eval("=ceiling(3.1)", None).unwrap();
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
        assert!((0.0..1.0).contains(&val));
    } else {
        panic!("Expected RAND to return a Float, got {:?}", result);
    }

    let (result, _) = sheet.eval("=randbetween(5, 10)", None).unwrap();
    if let ResultData::Integer(val) = result {
        assert!((5..=10).contains(&val));
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
fn test_structured_reference_basic_column() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("table_1".to_string()),
        rows: 3,
        cols: 3,
        ..Default::default()
    });
    sheet.columns[0].name = "Sales".to_string();
    sheet.columns[1].name = "Cost".to_string();

    sheet.set_cell_src(0, 0, "10".to_string());
    sheet.set_cell_src(1, 0, "20".to_string());
    sheet.set_cell_src(2, 0, "30".to_string());
    // Unqualified structured reference (same sheet).
    sheet.set_cell_src(0, 2, "=SUM([Sales])".to_string());
    // Qualified with the table (sheet) name.
    sheet.set_cell_src(1, 2, "=SUM(table_1[Sales])".to_string());
    sheet.commit(None).unwrap();

    assert_eq!(
        get_float_val(&sheet.get_result_data(&CellRef::new(0, 2))),
        Some(60.0)
    );
    assert_eq!(
        get_float_val(&sheet.get_result_data(&CellRef::new(1, 2))),
        Some(60.0)
    );
}

#[test]
fn test_structured_reference_this_row() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("table_1".to_string()),
        rows: 2,
        cols: 3,
        ..Default::default()
    });
    sheet.columns[0].name = "Sales".to_string();
    sheet.columns[1].name = "Cost".to_string();
    sheet.columns[2].name = "Profit".to_string();

    sheet.set_cell_src(0, 0, "100".to_string());
    sheet.set_cell_src(0, 1, "40".to_string());
    sheet.set_cell_src(1, 0, "200".to_string());
    sheet.set_cell_src(1, 1, "50".to_string());

    sheet.set_cell_src(0, 2, "=[@Sales] - [@Cost]".to_string());
    sheet.set_cell_src(1, 2, "=[@Sales] - [@Cost]".to_string());
    sheet.commit(None).unwrap();

    assert_eq!(
        get_float_val(&sheet.get_result_data(&CellRef::new(0, 2))),
        Some(60.0)
    );
    assert_eq!(
        get_float_val(&sheet.get_result_data(&CellRef::new(1, 2))),
        Some(150.0)
    );
}

#[test]
fn test_structured_reference_headers_and_totals_sections() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("table_1".to_string()),
        rows: 2,
        cols: 2,
        ..Default::default()
    });
    sheet.columns[0].name = "Sales".to_string();
    sheet.columns[1].name = "Cost".to_string();

    let (res, _) = sheet.eval("=table_1[[#Headers],[Sales]]", None).unwrap();
    assert_eq!(get_string_val(&res), Some("Sales".to_string()));

    let (res_bare, _) = sheet.eval("=[[#Headers],[Cost]]", None).unwrap();
    assert_eq!(get_string_val(&res_bare), Some("Cost".to_string()));

    // No totals row concept exists on the underlying table, so a totals
    // reference resolves to an empty/None result rather than erroring.
    let (res_totals, _) = sheet.eval("=[[#Totals],[Sales]]", None).unwrap();
    assert!(matches!(res_totals, ResultData::None));
}

#[test]
fn test_structured_reference_cross_sheet() {
    let mut sheet2 = Sheet::new(SheetInit {
        name: Some("Sheet2".to_string()),
        rows: 2,
        cols: 1,
        ..Default::default()
    });
    sheet2.columns[0].name = "Revenue".to_string();
    sheet2.set_cell_src(0, 0, "111".to_string());
    sheet2.set_cell_src(1, 0, "222".to_string());
    sheet2.commit(None).unwrap();

    let mut sheet1 = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 1,
        cols: 1,
        ..Default::default()
    });
    sheet1.set_cell_src(0, 0, "=SUM(Sheet2[Revenue])".to_string());

    let mut context = Context::default();
    context.sheets.insert("Sheet2".to_string(), &sheet2);
    sheet1.commit(Some(&context)).unwrap();

    assert_eq!(
        get_float_val(&sheet1.get_result_data(&CellRef::new(0, 0))),
        Some(333.0)
    );
}

#[test]
fn test_structured_reference_aggregates_ignore_text_like_a_range() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("table_1".to_string()),
        rows: 3,
        cols: 2,
        ..Default::default()
    });
    sheet.columns[0].name = "Sales".to_string();

    sheet.set_cell_src(0, 0, "10".to_string());
    sheet.set_cell_src(1, 0, "\"n/a\"".to_string());
    sheet.set_cell_src(2, 0, "20".to_string());
    sheet.set_cell_src(0, 1, "=SUM([Sales])".to_string());
    sheet.set_cell_src(1, 1, "=AVERAGE([Sales])".to_string());
    sheet.commit(None).unwrap();

    // A structured reference behaves like a range reference: non-numeric
    // text cells are ignored rather than raising #VALUE!.
    assert_eq!(
        get_float_val(&sheet.get_result_data(&CellRef::new(0, 1))),
        Some(30.0)
    );
    assert_eq!(
        get_float_val(&sheet.get_result_data(&CellRef::new(1, 1))),
        Some(15.0)
    );
}

#[test]
fn test_structured_reference_this_row_used_directly_in_sum_ignores_text() {
    // `[@Column]` evaluates to a single scalar cell value (like a plain
    // `CellRef`), not a `List`. When passed directly as a function argument
    // (e.g. `SUM([@Sales])`), a non-numeric text cell must be ignored just
    // like `SUM(A1)` would ignore a text cell in A1 -- not raise #VALUE!.
    let mut sheet = Sheet::new(SheetInit {
        name: Some("table_1".to_string()),
        rows: 2,
        cols: 2,
        ..Default::default()
    });
    sheet.columns[0].name = "Sales".to_string();

    sheet.set_cell_src(0, 0, "10".to_string());
    sheet.set_cell_src(1, 0, "\"n/a\"".to_string());
    sheet.set_cell_src(0, 1, "=SUM([@Sales])".to_string());
    sheet.set_cell_src(1, 1, "=SUM([@Sales])".to_string());
    sheet.commit(None).unwrap();

    assert_eq!(
        get_float_val(&sheet.get_result_data(&CellRef::new(0, 1))),
        Some(10.0)
    );
    assert_eq!(
        get_float_val(&sheet.get_result_data(&CellRef::new(1, 1))),
        Some(0.0)
    );
}

#[test]
fn test_structured_reference_whole_table_data_section() {
    let mut sheet2 = Sheet::new(SheetInit {
        name: Some("Sheet2".to_string()),
        rows: 2,
        cols: 2,
        ..Default::default()
    });
    sheet2.columns[0].name = "A".to_string();
    sheet2.columns[1].name = "B".to_string();
    sheet2.set_cell_src(0, 0, "1".to_string());
    sheet2.set_cell_src(1, 0, "2".to_string());
    sheet2.set_cell_src(0, 1, "3".to_string());
    sheet2.set_cell_src(1, 1, "4".to_string());
    sheet2.commit(None).unwrap();

    let mut sheet1 = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 1,
        cols: 1,
        ..Default::default()
    });
    // `[#Data]` with no column name spans every column in the table.
    sheet1.set_cell_src(0, 0, "=SUM(Sheet2[#Data])".to_string());

    let mut context = Context::default();
    context.sheets.insert("Sheet2".to_string(), &sheet2);
    sheet1.commit(Some(&context)).unwrap();

    assert_eq!(
        get_float_val(&sheet1.get_result_data(&CellRef::new(0, 0))),
        Some(10.0)
    );
}

#[test]
fn test_structured_reference_whole_row_no_column() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("table_1".to_string()),
        rows: 2,
        cols: 3,
        ..Default::default()
    });
    sheet.columns[0].name = "Sales".to_string();
    sheet.columns[1].name = "Cost".to_string();
    sheet.columns[2].name = "Extra".to_string();

    sheet.set_cell_src(0, 0, "10".to_string());
    sheet.set_cell_src(0, 1, "3".to_string());
    sheet.set_cell_src(0, 2, "2".to_string());
    sheet.commit(None).unwrap();

    // `[@]` (this row, no column) spans every column of the current row.
    let (res, _) = sheet.eval_with_row("=SUM([@])", None, Some(0)).unwrap();
    assert_eq!(get_float_val(&res), Some(15.0));
}

#[test]
fn test_structured_reference_missing_column_errors() {
    let sheet = Sheet::new(SheetInit {
        name: Some("table_1".to_string()),
        rows: 2,
        cols: 2,
        ..Default::default()
    });
    assert!(sheet.eval("=[NoSuchColumn]", None).is_err());
}

#[test]
fn test_structured_reference_missing_table_errors() {
    let sheet = Sheet::new(SheetInit {
        name: Some("table_1".to_string()),
        rows: 1,
        cols: 1,
        ..Default::default()
    });
    assert!(sheet.eval("=NoSuchTable[Col]", None).is_err());
}

#[test]
fn test_structured_reference_recomputes_on_column_change() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("table_1".to_string()),
        rows: 2,
        cols: 2,
        ..Default::default()
    });
    sheet.columns[0].name = "Sales".to_string();

    sheet.set_cell_src(0, 0, "10".to_string());
    sheet.set_cell_src(1, 0, "20".to_string());
    sheet.set_cell_src(0, 1, "=SUM([Sales])".to_string());
    sheet.commit(None).unwrap();
    assert_eq!(
        get_float_val(&sheet.get_result_data(&CellRef::new(0, 1))),
        Some(30.0)
    );

    // Changing a cell elsewhere in the referenced column should invalidate
    // and recompute the dependent structured-reference formula.
    sheet.set_cell_src(1, 0, "50".to_string());
    sheet.commit(None).unwrap();
    assert_eq!(
        get_float_val(&sheet.get_result_data(&CellRef::new(0, 1))),
        Some(60.0)
    );
}

#[test]
fn test_excel_table_structured_reference_respects_table_row_bounds() {
    // Column 0 has data both above and below the defined table's row range;
    // a structured reference into the table must only see the table's own
    // rows, unlike the legacy whole-sheet fallback which scans every row.
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 5,
        cols: 2,
        ..Default::default()
    });
    sheet.set_cell_src(0, 0, "999".to_string()); // above the table
    sheet.set_cell_src(1, 0, "Amount".to_string()); // header row
    sheet.set_cell_src(2, 0, "10".to_string());
    sheet.set_cell_src(3, 0, "20".to_string());
    sheet.set_cell_src(4, 0, "888".to_string()); // below the table
    sheet.commit(None).unwrap();

    sheet
        .add_table("Sales".to_string(), 1, 0, 3, 0, true, false)
        .unwrap();

    let (res, _) = sheet.eval("=SUM(Sales[Amount])", None).unwrap();
    assert_eq!(get_float_val(&res), Some(30.0));
}

#[test]
fn test_excel_table_totals_row_returns_real_value() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 4,
        cols: 1,
        ..Default::default()
    });
    sheet.set_cell_src(0, 0, "Amount".to_string());
    sheet.set_cell_src(1, 0, "10".to_string());
    sheet.set_cell_src(2, 0, "20".to_string());
    sheet.set_cell_src(3, 0, "=SUM(A2:A3)".to_string());
    sheet.commit(None).unwrap();

    sheet
        .add_table("Sales".to_string(), 0, 0, 3, 0, true, true)
        .unwrap();

    let (res, _) = sheet.eval("=Sales[[#Totals],[Amount]]", None).unwrap();
    assert_eq!(get_float_val(&res), Some(30.0));
}

#[test]
fn test_excel_table_totals_section_without_totals_row_is_none() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 3,
        cols: 1,
        ..Default::default()
    });
    sheet.set_cell_src(0, 0, "Amount".to_string());
    sheet.set_cell_src(1, 0, "10".to_string());
    sheet.set_cell_src(2, 0, "20".to_string());
    sheet.commit(None).unwrap();

    // has_totals_row = false: no totals row is reserved at all.
    sheet
        .add_table("Sales".to_string(), 0, 0, 2, 0, true, false)
        .unwrap();

    let (res, _) = sheet.eval("=Sales[[#Totals],[Amount]]", None).unwrap();
    assert!(matches!(res, ResultData::None));
}

#[test]
fn test_excel_table_headers_use_stored_table_column_name() {
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 3,
        cols: 1,
        ..Default::default()
    });
    sheet.set_cell_src(0, 0, "Amount".to_string());
    sheet.set_cell_src(1, 0, "10".to_string());
    sheet.set_cell_src(2, 0, "20".to_string());
    sheet.commit(None).unwrap();

    sheet
        .add_table("Sales".to_string(), 0, 0, 2, 0, true, false)
        .unwrap();

    let (res, _) = sheet.eval("=Sales[[#Headers],[Amount]]", None).unwrap();
    assert_eq!(get_string_val(&res), Some("Amount".to_string()));
}

#[test]
fn test_excel_table_cross_sheet_reference() {
    let mut sheet2 = Sheet::new(SheetInit {
        name: Some("Sheet2".to_string()),
        rows: 3,
        cols: 1,
        ..Default::default()
    });
    sheet2.set_cell_src(0, 0, "Amount".to_string());
    sheet2.set_cell_src(1, 0, "100".to_string());
    sheet2.set_cell_src(2, 0, "200".to_string());
    sheet2.commit(None).unwrap();
    sheet2
        .add_table("Revenue".to_string(), 0, 0, 2, 0, true, false)
        .unwrap();

    let sheet1 = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 1,
        cols: 1,
        ..Default::default()
    });

    let mut context = Context::default();
    context.sheets.insert("Sheet2".to_string(), &sheet2);

    let (res, _) = sheet1
        .eval("=SUM(Revenue[Amount])", Some(&context))
        .unwrap();
    assert_eq!(get_float_val(&res), Some(300.0));
}

#[test]
fn test_excel_table_structured_reference_survives_commit() {
    // Regression test: `commit()` re-derives each formula's evaluated
    // source text via `compile_formula`/`serialize_formula` on every run
    // (see Sheet::commit), not just the first time. That recompilation
    // step must recognize a real ExcelTable's columns -- if it instead
    // tried to resolve them the legacy way (as plain DataColumn names, all
    // of which are blank for a table's columns), it would rewrite the
    // formula's column name to a bogus placeholder and break it, even
    // though a direct `sheet.eval()` call (bypassing compile_formula) would
    // have worked fine.
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 4,
        cols: 3,
        ..Default::default()
    });
    sheet.set_cell_src(0, 0, "Name".to_string());
    sheet.set_cell_src(0, 1, "Amount".to_string());
    sheet.set_cell_src(1, 0, "Widget".to_string());
    sheet.set_cell_src(1, 1, "10".to_string());
    sheet.set_cell_src(2, 0, "Gadget".to_string());
    sheet.set_cell_src(2, 1, "20".to_string());
    sheet.commit(None).unwrap();

    sheet
        .add_table("Sales".to_string(), 0, 0, 2, 1, true, false)
        .unwrap();

    sheet.set_cell_src(0, 2, "=SUM(Sales[Amount])".to_string());
    sheet.set_cell_src(1, 2, "=Sales[[#Headers],[Amount]]".to_string());
    sheet.commit(None).unwrap();

    assert_eq!(
        get_float_val(&sheet.get_result_data(&CellRef::new(0, 2))),
        Some(30.0)
    );
    assert_eq!(
        get_string_val(&sheet.get_result_data(&CellRef::new(1, 2))),
        Some("Amount".to_string())
    );

    // Committing again (as e.g. re-evaluating an already-saved workbook
    // would) must keep working -- this is what actually exercises the
    // repeated compile_formula/serialize_formula round-trip.
    sheet.mark_all_dirty();
    sheet.commit(None).unwrap();
    assert_eq!(
        get_float_val(&sheet.get_result_data(&CellRef::new(0, 2))),
        Some(30.0)
    );
}

#[test]
fn test_excel_table_column_reference_dependency_is_row_scoped_not_whole_column() {
    // A structured column reference must depend on only the table's own
    // data rows (like a bounded range reference, e.g. A1:A100), not the
    // whole sheet column. Verified against real Excel: placing a summary
    // formula like `=SUM(Inventory[Price])` in the same column as the
    // table but outside its rows is NOT circular there, and changing it
    // doesn't need a whole-column dependency to invalidate correctly --
    // only per-row dependencies on the table's own rows do.
    let mut sheet = Sheet::new(SheetInit {
        name: Some("Sheet1".to_string()),
        rows: 6,
        cols: 2,
        ..Default::default()
    });
    sheet.set_cell_src(0, 0, "Item".to_string());
    sheet.set_cell_src(0, 1, "Price".to_string());
    sheet.set_cell_src(1, 0, "Widget".to_string());
    sheet.set_cell_src(1, 1, "9.99".to_string());
    sheet.set_cell_src(2, 0, "Gadget".to_string());
    sheet.set_cell_src(2, 1, "19.99".to_string());
    sheet.set_cell_src(3, 0, "Gizmo".to_string());
    sheet.set_cell_src(3, 1, "4.5".to_string());
    sheet.commit(None).unwrap();

    sheet
        .add_table("Inventory".to_string(), 0, 0, 3, 1, true, false)
        .unwrap();

    // Row 5 (0-based): same "Price" column (1) as the table, but well
    // outside the table's own rows (0..=3).
    sheet.set_cell_src(5, 1, "=SUM(Inventory[Price])".to_string());
    sheet.commit(None).unwrap();

    assert_eq!(
        get_float_val(&sheet.get_result_data(&CellRef::new(5, 1))),
        Some(34.48)
    );

    // No whole-column dependency should exist for column 1 -- only
    // per-row dependencies on the table's own data rows (1..=3; row 0 is
    // the header, excluded). If this ever regresses to a whole-column
    // dependency, the summary formula would depend on its own cell (since
    // it also lives in column 1) and false-positive as circular, which
    // real Excel does not do.
    assert!(
        !sheet.dependencies.contains_key(&Dependency::LocalColumn(1)),
        "structured table reference must not register a whole-column dependency"
    );
    for r in 1..=3 {
        assert!(
            sheet
                .dependencies
                .contains_key(&Dependency::Local(CellRef::new(r, 1))),
            "expected a per-row dependency on row {r}"
        );
    }

    // Changing an in-table cell must still correctly invalidate and
    // recompute the summary formula.
    sheet.set_cell_src(1, 1, "100".to_string());
    sheet.commit(None).unwrap();
    assert_eq!(
        get_float_val(&sheet.get_result_data(&CellRef::new(5, 1))),
        Some(124.49)
    );
}
