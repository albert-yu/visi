use super::*;

#[test]
fn test_text_manipulation_functions() {
    let grid = [[
        "=CHAR(65)",
        "=CODE(\"A\")",
        "=EXACT(\"abc\", \"abc\")",
        "=REPT(\"a\", 3)",
        "=SUBSTITUTE(\"banana\", \"a\", \"o\")",
        "=UNICHAR(65)",
    ]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    let r1 = sheet.get_result_data(&CellRef::new(0, 0));
    assert!(matches!(r1, ResultData::String(ref s) if s == "A"));

    let r2 = sheet.get_result_data(&CellRef::new(0, 1));
    assert!(matches!(r2, ResultData::Float(v) if (v - 65.0).abs() < 1e-6));

    let r3 = sheet.get_result_data(&CellRef::new(0, 2));
    assert!(matches!(r3, ResultData::Boolean(true)));

    let r4 = sheet.get_result_data(&CellRef::new(0, 3));
    assert!(matches!(r4, ResultData::String(ref s) if s == "aaa"));

    let r5 = sheet.get_result_data(&CellRef::new(0, 4));
    assert!(matches!(r5, ResultData::String(ref s) if s == "bonono"));

    let r6 = sheet.get_result_data(&CellRef::new(0, 5));
    assert!(matches!(r6, ResultData::String(ref s) if s == "A"));
}

#[test]
fn test_text_search_and_split_functions() {
    let grid = [[
        "=FIND(\"bar\", \"foobar\")",
        "=TEXTBEFORE(\"hello-world\", \"-\")",
        "=TEXTAFTER(\"hello-world\", \"-\")",
        "=TEXTJOIN(\", \", TRUE, \"a\", \"b\", \"c\")",
        "=VALUE(\"123.45\")",
    ]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    let r1 = sheet.get_result_data(&CellRef::new(0, 0));
    assert!(matches!(r1, ResultData::Float(v) if (v - 4.0).abs() < 1e-6));

    let r2 = sheet.get_result_data(&CellRef::new(0, 1));
    assert!(matches!(r2, ResultData::String(ref s) if s == "hello"));

    let r3 = sheet.get_result_data(&CellRef::new(0, 2));
    assert!(matches!(r3, ResultData::String(ref s) if s == "world"));

    let r4 = sheet.get_result_data(&CellRef::new(0, 3));
    assert!(matches!(r4, ResultData::String(ref s) if s == "a, b, c"));

    let r5 = sheet.get_result_data(&CellRef::new(0, 4));
    assert!(matches!(r5, ResultData::Float(v) if (v - 123.45).abs() < 1e-6));
}

#[test]
fn test_currency_and_formatting_functions() {
    let grid = [[
        "=DOLLAR(1234.56)",
        "=FIXED(1234.56, 1)",
        "=TEXT(0.25, \"0.0%\")",
        "=NUMBERVALUE(\"1,234.56\", \".\", \",\")",
    ]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    let r1 = sheet.get_result_data(&CellRef::new(0, 0));
    assert!(matches!(r1, ResultData::String(ref s) if s == "$1,234.56"));

    let r2 = sheet.get_result_data(&CellRef::new(0, 1));
    assert!(matches!(r2, ResultData::String(ref s) if s == "1,234.6"));

    let r3 = sheet.get_result_data(&CellRef::new(0, 2));
    assert!(matches!(r3, ResultData::String(ref s) if s == "25.0%"));

    let r4 = sheet.get_result_data(&CellRef::new(0, 3));
    assert!(matches!(r4, ResultData::Float(v) if (v - 1234.56).abs() < 1e-6));
}

#[test]
fn test_jis_is_inverse_of_asc() {
    let grid = [["=ASC(JIS(\"AB 1\"))", "=JIS(\"AB\")"]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    let r1 = sheet.get_result_data(&CellRef::new(0, 0));
    assert!(matches!(r1, ResultData::String(ref s) if s == "AB 1"));

    let r2 = sheet.get_result_data(&CellRef::new(0, 1));
    assert!(matches!(r2, ResultData::String(ref s) if s == "\u{FF21}\u{FF22}"));
}

#[test]
fn test_filterxml_basic_and_errors() {
    let grid = [[
        "=FILTERXML(\"<a><b id='2'>hi</b><b id='9'>lo</b></a>\", \"/a/b[@id='9']\")",
        "=FILTERXML(\"<a><b>x</b></a>\", \"/a/b/text()\")",
        "=FILTERXML(\"<a><b>x</b></a>\", \"/a/c\")",
        "=FILTERXML(\"<a><b>x</b>\", \"/a/b\")",
    ]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    let r1 = sheet.get_result_data(&CellRef::new(0, 0));
    assert!(matches!(r1, ResultData::String(ref s) if s == "lo"));

    let r2 = sheet.get_result_data(&CellRef::new(0, 1));
    assert!(matches!(r2, ResultData::String(ref s) if s == "x"));

    let r3 = sheet.get_result_data(&CellRef::new(0, 2));
    assert!(matches!(r3, ResultData::Error(ref e) if e == "#N/A"));

    let r4 = sheet.get_result_data(&CellRef::new(0, 3));
    assert!(matches!(r4, ResultData::Error(ref e) if e == "#VALUE!"));
}
