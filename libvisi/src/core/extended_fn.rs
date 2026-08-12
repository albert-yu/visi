// High-precision Information, Database, Lookup, Logical, Web, and Cube functions for libvisi

use crate::core::engine::ResultData;

// ============================================================================
// 1. Information Functions
// ============================================================================

pub fn error_type(err: &str) -> Result<f64, String> {
    match err {
        "#NULL!" => Ok(1.0),
        "#DIV/0!" => Ok(2.0),
        "#VALUE!" => Ok(3.0),
        "#REF!" => Ok(4.0),
        "#NAME?" => Ok(5.0),
        "#NUM!" => Ok(6.0),
        "#N/A" => Ok(7.0),
        "#GETTING_DATA" => Ok(8.0),
        _ => Err("#N/A".to_string()),
    }
}

pub fn iserr(val: &ResultData) -> bool {
    if let ResultData::Error(e) = val {
        e != "#N/A"
    } else {
        false
    }
}

pub fn iseven(number: f64) -> bool {
    (number.floor() as i64) % 2 == 0
}
pub fn isodd(number: f64) -> bool {
    (number.floor() as i64) % 2 != 0
}
pub fn islogical(val: &ResultData) -> bool {
    matches!(val, ResultData::Boolean(_))
}
pub fn isnontext(val: &ResultData) -> bool {
    !matches!(val, ResultData::String(_))
}
pub fn n_fn(val: &ResultData) -> f64 {
    match val {
        ResultData::Float(f) => *f,
        ResultData::Integer(i) => *i as f64,
        ResultData::Boolean(b) if *b => 1.0,
        _ => 0.0,
    }
}
pub fn na_fn() -> ResultData {
    ResultData::Error("#N/A".to_string())
}
pub fn type_fn(val: &ResultData) -> f64 {
    match val {
        ResultData::Float(_) | ResultData::Integer(_) => 1.0,
        ResultData::String(_) => 2.0,
        ResultData::Boolean(_) => 4.0,
        ResultData::Error(_) => 16.0,
        ResultData::List(_) => 64.0,
        _ => 1.0,
    }
}

// ============================================================================
// 2. Logical & Array Functions
// ============================================================================

pub fn xor_fn(bools: &[bool]) -> bool {
    bools.iter().filter(|&&b| b).count() % 2 != 0
}

pub fn ifna(val: ResultData, alt: ResultData) -> ResultData {
    if let ResultData::Error(ref e) = val
        && e == "#N/A"
    {
        return alt;
    }
    val
}

// ============================================================================
// 3. Lookup & Reference Helpers
// ============================================================================

pub fn address_fn(
    row_num: f64,
    col_num: f64,
    abs_num: Option<f64>,
    _a1: Option<bool>,
    _sheet_name: Option<&str>,
) -> Result<String, String> {
    let r = row_num.floor() as usize;
    let c = col_num.floor() as usize;
    let abs_type = abs_num.unwrap_or(1.0).floor() as i32;

    if r == 0 || c == 0 {
        return Err("#VALUE!".to_string());
    }

    let mut col_str = String::new();
    let mut curr_c = c - 1;
    loop {
        let rem = curr_c % 26;
        col_str.insert(0, (b'A' + rem as u8) as char);
        if curr_c < 26 {
            break;
        }
        curr_c = curr_c / 26 - 1;
    }
    match abs_type {
        1 => Ok(format!("${}${}", col_str, r)),
        2 => Ok(format!("${}{}", col_str, r)),
        3 => Ok(format!("{}${}", col_str, r)),
        4 => Ok(format!("{}{}", col_str, r)),
        _ => Ok(format!("${}${}", col_str, r)),
    }
}

// ============================================================================
// 4. Web & Stub Functions
// ============================================================================

pub fn encodeurl(text: &str) -> Result<String, String> {
    let mut res = String::new();
    for b in text.bytes() {
        if b.is_ascii_alphanumeric() || b == b'-' || b == b'_' || b == b'.' || b == b'~' {
            res.push(b as char);
        } else {
            res.push_str(&format!("%{:02X}", b));
        }
    }
    Ok(res)
}
