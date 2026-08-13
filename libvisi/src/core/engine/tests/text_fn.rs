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
fn test_regex_functions_use_real_regex_not_literal_substring() {
    let grid = [[
        "=REGEXEXTRACT(\"order-12345\", \"[0-9]+\")",
        "=REGEXREPLACE(\"order-12345\", \"[0-9]+\", \"X\")",
        "=REGEXTEST(\"order-12345\", \"[0-9]+\")",
        "=REGEXTEST(\"order-abcde\", \"[0-9]+\")",
    ]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    let r1 = sheet.get_result_data(&CellRef::new(0, 0));
    assert!(matches!(r1, ResultData::String(ref s) if s == "12345"));

    let r2 = sheet.get_result_data(&CellRef::new(0, 1));
    assert!(matches!(r2, ResultData::String(ref s) if s == "order-X"));

    let r3 = sheet.get_result_data(&CellRef::new(0, 2));
    assert!(matches!(r3, ResultData::Boolean(true)));

    let r4 = sheet.get_result_data(&CellRef::new(0, 3));
    assert!(matches!(r4, ResultData::Boolean(false)));
}

#[test]
fn test_text_number_format_codes() {
    // TEXT()'s number-format handling used to be a crude stub: `%` always
    // hardcoded exactly 1 decimal place regardless of the format string,
    // and `,` (thousands grouping), `$` (currency), and date-token
    // formats ("yyyy-mm-dd") weren't implemented at all -- the raw
    // number was returned unformatted. Found via differential fuzzing
    // (every TEXT() call with one of these formats mismatched real
    // Excel).
    let grid = [[
        "=TEXT(-7679.0669, \"$#,##0.00\")",
        "=TEXT(3021.1929, \"#,##0\")",
        "=TEXT(6436.3899, \"0%\")",
        "=TEXT(DATE(1910,8,29), \"yyyy-mm-dd\")",
    ]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    let r1 = sheet.get_result_data(&CellRef::new(0, 0));
    assert!(
        matches!(r1, ResultData::String(ref s) if s == "-$7,679.07"),
        "{r1:?}"
    );

    let r2 = sheet.get_result_data(&CellRef::new(0, 1));
    assert!(
        matches!(r2, ResultData::String(ref s) if s == "3,021"),
        "{r2:?}"
    );

    // 6436.3899 * 100 = 643639.9%, "0" has no decimal places -> rounds
    // to a whole percent, not the old stub's hardcoded ".0".
    let r3 = sheet.get_result_data(&CellRef::new(0, 2));
    assert!(
        matches!(r3, ResultData::String(ref s) if s == "643639%"),
        "{r3:?}"
    );

    let r4 = sheet.get_result_data(&CellRef::new(0, 3));
    assert!(
        matches!(r4, ResultData::String(ref s) if s == "1910-08-29"),
        "{r4:?}"
    );
}

#[test]
fn test_proper_capitalizes_letter_after_digits() {
    // PROPER used `is_alphanumeric()` to decide whether a character
    // could consume the "capitalize the next letter" flag, so a run of
    // digits incorrectly ate it the same way a letter would --
    // PROPER("123abc") returned "123abc" unchanged instead of "123Abc".
    // Per Microsoft's own definition, PROPER capitalizes a letter
    // preceded by "any character that is not a letter", which includes
    // digits, not just punctuation/spacing.
    let grid = [["=PROPER(\"123abc\")"]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    let r = sheet.get_result_data(&CellRef::new(0, 0));
    assert!(
        matches!(r, ResultData::String(ref s) if s == "123Abc"),
        "{r:?}"
    );
}

#[test]
fn test_arraytotext_joins_every_element_including_text_and_blanks() {
    // ARRAYTOTEXT used flatten_stat_numbers, which silently drops any
    // non-numeric cell (its lenient mode is built for SUM/AVERAGE-style
    // aggregates) -- so a range with any text or blank cells produced a
    // result missing those elements entirely, instead of joining every
    // element's own text the way real Excel does.
    let grid = [
        ["1", "\"hello\"", "TRUE", ""],
        ["=ARRAYTOTEXT(A1:D1)", "", "", ""],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    let r = sheet.get_result_data(&CellRef::new(1, 0));
    assert!(
        matches!(r, ResultData::String(ref s) if s == "1, hello, TRUE, "),
        "{r:?}"
    );
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

#[test]
fn test_number_to_text_matches_excel_decimal_range() {
    // Excel keeps plain decimal notation until the decimal rendering would
    // exceed 20 characters. The old magnitude cutoffs (1e-5 .. 1e11) were
    // far narrower and turned numbers Excel writes out in full into
    // scientific notation -- CONCATENATE over SINH(...) produced
    // "9.76121418126432E+11" where Excel gives "976121418126.432".
    // Every expectation is verbatim real-Excel output.
    use crate::core::engine::result_data::format_excel_number;
    for (value, expected) in [
        (976121418126.432_f64, "976121418126.432"),
        (1e15, "1000000000000000"),
        (1e18, "1000000000000000000"),
        (1e19, "10000000000000000000"),
        (1e20, "1E+20"),
        (0.000001, "0.000001"),
        (0.000000001, "0.000000001"),
        (0.000001207666770903, "0.000001207666770903"),
        (0.00000120766677090395, "1.20766677090395E-06"),
        (0.5, "0.5"),
        (0.0, "0"),
    ] {
        let got = format_excel_number(value);
        assert_eq!(got, expected, "format_excel_number({value})");
    }
}

#[test]
fn test_text_rounds_half_away_from_zero() {
    // TEXT used Rust's `{:.N}` formatting, which rounds the *binary* value
    // to nearest-even; Excel rounds half away from zero on the decimal it
    // displays. TEXT(-3873.705, "0.00") is -3873.71 in Excel but came out
    // as -3873.70 here.
    for (value, fmt, expected) in [
        (-3873.705_f64, "0.00", "-3873.71"),
        (2.675, "0.00", "2.68"),
        (1.005, "0.00", "1.01"),
    ] {
        let got = crate::core::text::text_fn(value, fmt);
        assert_eq!(got, Ok(expected.to_string()), "TEXT({value}, {fmt:?})");
    }
}
