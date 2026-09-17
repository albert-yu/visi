use serde::{Deserialize, Serialize};

/// What a cell holds once it has been evaluated.
///
/// We intentionally omit a date type. In Excel, a date is a plain
/// numeric serial and the notation is saved separately, as
/// `CellStyle::num_format`. This way, we can treat dates like
/// any other numeric type (e.g. `SUM`).
///
/// An Excel error is a *value*, not a Rust error, e.g. `=1/0` evaluates
/// successfully to `Error("#DIV/0!")`. See [`EngineError`] for Rust errors.
///
/// [`EngineError`]: crate::core::EngineError
#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum ResultData {
    /// A blank cell. Coerces to 0 or `""` depending on what reads it.
    None,
    /// `TRUE` or `FALSE`.
    Boolean(bool),
    /// A whole number.
    Integer(i64),
    /// A number that is not a whole number, or one too large for an `i64`.
    /// A date is a `Float` holding its Excel serial.
    Float(f64),
    /// Text.
    String(String),
    /// An ordered sequence, for the engine-specific functions that return one.
    /// Not an Excel array.
    List(Vec<ResultData>),
    /// Key/value pairs, for the engine-specific functions that return them.
    Dict(Vec<(ResultData, ResultData)>),
    /// An Excel error value, held as its code: `#DIV/0!`, `#VALUE!`, `#N/A`.
    Error(String),
}

impl std::fmt::Display for ResultData {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            ResultData::None => write!(f, ""),
            ResultData::Boolean(b) => write!(f, "{}", if *b { "TRUE" } else { "FALSE" }),
            ResultData::Integer(i) => write!(f, "{}", i),
            ResultData::Float(fl) => write!(f, "{}", format_excel_number(*fl)),
            ResultData::String(s) => write!(f, "{}", s),
            ResultData::List(l) => {
                let items: Vec<String> = l.iter().map(|i| i.to_string()).collect();
                write!(f, "[{}]", items.join(", "))
            }
            ResultData::Dict(d) => {
                let items: Vec<String> = d.iter().map(|(k, v)| format!("{}: {}", k, v)).collect();
                write!(f, "{{ {} }}", items.join(", "))
            }
            ResultData::Error(e) => write!(f, "Error: {}", e),
        }
    }
}

fn round_digits_half_up(digits: &str, exp: i32, keep: usize) -> (String, i32) {
    if digits.len() <= keep {
        return (digits.to_string(), exp);
    }
    let mut kept: Vec<u8> = digits.as_bytes()[..keep].to_vec();
    let round_up = digits.as_bytes()[keep] >= b'5';
    let mut exp = exp;
    if round_up {
        let mut i = keep;
        loop {
            if i == 0 {
                kept.insert(0, b'1');
                kept.pop();
                exp += 1;
                break;
            }
            i -= 1;
            if kept[i] == b'9' {
                kept[i] = b'0';
            } else {
                kept[i] += 1;
                break;
            }
        }
    }
    let mut out = String::from_utf8(kept).expect("ascii digits");
    while out.len() > 1 && out.ends_with('0') {
        out.pop();
    }
    (out, exp)
}

pub(crate) const EXCEL_ERROR_CODES: &[&str] = &[
    "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#CALC!", "#SPILL!",
];

pub(crate) fn is_excel_error_code(src: &str) -> bool {
    EXCEL_ERROR_CODES
        .iter()
        .any(|e| src.eq_ignore_ascii_case(e))
}

pub(crate) fn format_excel_number(f: f64) -> String {
    if f == 0.0 {
        return "0".to_string();
    }
    if f.is_nan() || f.is_infinite() {
        return "#NUM!".to_string();
    }

    let sci = format!("{:.14e}", f);
    let (mantissa, exp_str) = sci.split_once('e').expect("{:e} always emits an exponent");
    let exp: i32 = exp_str.parse().expect("{:e} emits an integer exponent");
    let sign = if f < 0.0 { "-" } else { "" };
    let digits: String = mantissa.chars().filter(|c| c.is_ascii_digit()).collect();
    let digits = digits.trim_end_matches('0');
    let digits = if digits.is_empty() { "0" } else { digits };

    let decimal_len = if exp >= 0 {
        let int_digits = (exp + 1) as usize;
        let frac_digits = digits.len().saturating_sub(int_digits);
        int_digits + usize::from(frac_digits > 0) + frac_digits
    } else {
        2 + (-exp - 1) as usize + digits.len()
    };

    if decimal_len <= 20 {
        if exp >= 0 {
            let int_digits = (exp + 1) as usize;
            let mut out = String::from(sign);
            if digits.len() <= int_digits {
                out.push_str(digits);
                out.push_str(&"0".repeat(int_digits - digits.len()));
            } else {
                out.push_str(&digits[..int_digits]);
                out.push('.');
                out.push_str(&digits[int_digits..]);
            }
            out
        } else {
            format!("{}0.{}{}", sign, "0".repeat((-exp - 1) as usize), digits)
        }
    } else {
        let suffix_budget_len = if exp.abs() >= 99 {
            5
        } else {
            format!("E{:+03}", exp).len()
        };
        let frac_digits = 18usize.saturating_sub(suffix_budget_len).min(14);

        let (rounded_digits, exp) = round_digits_half_up(digits, exp, frac_digits + 1);
        let mut mantissa = String::from(sign);
        mantissa.push_str(&rounded_digits[..1]);
        if rounded_digits.len() > 1 {
            mantissa.push('.');
            mantissa.push_str(&rounded_digits[1..]);
        }
        format!("{}E{:+03}", mantissa, exp)
    }
}
