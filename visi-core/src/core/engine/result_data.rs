//! The value type a cell evaluates to.

use serde::{Deserialize, Serialize};

/// What a cell holds once it has been evaluated.
///
/// There is deliberately **no date variant**. As in Excel, a date is a plain
/// numeric serial and the notation it was typed in lives on the cell, as
/// `CellStyle::num_format` -- so `ISNUMBER` is true for a date, `SUM` counts
/// it, and every numeric path works on it untouched. Only rendering consults
/// the format, through `Sheet::get_display_string`.
///
/// An Excel error is a *value*, not a Rust error: `=1/0` evaluates
/// successfully to `Error("#DIV/0!")`. See [`EngineError`] for the failures
/// that are not values.
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

/// Rounds a significant-digit string to `keep` digits, half away from zero,
/// trimming the trailing zeros Excel does not display. Returns the digits
/// and the (possibly incremented) exponent -- rounding 999 up to 100 shifts
/// the decimal point.
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
                // Every digit carried: 999... becomes 1000..., one decimal
                // place further left.
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

/// The Excel error values, spelled exactly as a cell shows them.
///
/// A closed set: these are the only strings a cell can hold that are an error
/// rather than text, which is what makes recognising one on entry safe.
pub(crate) const EXCEL_ERROR_CODES: &[&str] = &[
    "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#CALC!", "#SPILL!",
];

/// Whether a literal cell entry is one of Excel's error values.
///
/// Typing `#NUM!` into Excel produces the error, not the text -- measured,
/// along with the same thing happening when VBA assigns the string through
/// `Range.Value`. So `Sheet::commit` recognises one, and `xlsx::text_cell_src`
/// quotes it on import for the same reason it quotes `TRUE` and `6/22/26`:
/// a cell Excel told us is *text* has to survive the round trip as text.
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

    // Excel displays 15 significant digits and no more. Everything below
    // works from the *rounded* scientific form rather than from the f64
    // directly, so digits past that precision are dropped instead of
    // leaking out: (-43)^11 is 21611482313284248 in f64, but Excel writes
    // 21611482313284200.
    let sci = format!("{:.14e}", f);
    let (mantissa, exp_str) = sci.split_once('e').expect("{:e} always emits an exponent");
    let exp: i32 = exp_str.parse().expect("{:e} emits an integer exponent");
    let sign = if f < 0.0 { "-" } else { "" };
    let digits: String = mantissa.chars().filter(|c| c.is_ascii_digit()).collect();
    let digits = digits.trim_end_matches('0');
    let digits = if digits.is_empty() { "0" } else { digits };

    // Excel keeps plain decimal notation for as long as the decimal
    // rendering stays within 20 characters, and only then falls back to
    // scientific. The minus sign is *not* charged against that budget --
    // real Excel writes -2.05237592634038E-10, which is 21 characters.
    // Verified against real Excel: 1e18 and 1e19 render in full (19 and 20
    // characters) while 1e20 (21) goes scientific, and 0.000001207666770903
    // renders in full (20) while 0.00000120766677090395 (22) goes
    // scientific.
    let decimal_len = if exp >= 0 {
        let int_digits = (exp + 1) as usize;
        let frac_digits = digits.len().saturating_sub(int_digits);
        int_digits + usize::from(frac_digits > 0) + frac_digits
    } else {
        // "0." + leading zeros + significant digits
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
        // The 20-character budget applies to the scientific rendering too,
        // and the exponent is charged against it: a three-digit exponent
        // leaves one fewer mantissa digit than a two-digit one. Real Excel
        // writes PHI(28) as "2.2775774787367E-171" (13 fractional digits)
        // and CSCH(-23) as "-2.05237592634038E-10" (14), both 20 characters
        // once the sign is set aside.
        // Excel starts reserving the three-exponent-digit budget at |99|,
        // even though the textual exponent still has only two digits there:
        // 6.37409849780041E-98 keeps 14 fractional digits, while
        // 6.3740984978004E-99 keeps 13. The actual `E+99`/`E-99` suffix is
        // still rendered with two digits; only the mantissa budget shrinks.
        let suffix_budget_len = if exp.abs() >= 99 {
            5
        } else {
            format!("E{:+03}", exp).len()
        };
        let frac_digits = 18usize.saturating_sub(suffix_budget_len).min(14);

        // Rounded from the *15-significant-digit* value, not from the raw
        // f64. Excel snaps a result to 15 significant digits and only then
        // formats it, so when a three-digit exponent leaves room for just
        // 14 the two roundings compose. 28^-92 is
        // 7.26877317134744769...e-134: rounding that straight to 14 digits
        // gives ...7474, but snapping to 15 first gives 7.26877317134745
        // and then 14 gives ...7475, which is what Excel prints.
        //
        // Working from the digit string rather than re-rounding the f64
        // keeps the two steps exact, and rounds half away from zero, as
        // Excel does elsewhere (DOLLAR/FIXED/TEXT).
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
