use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum ResultData {
    None,
    Boolean(bool),
    Integer(i64),
    Float(f64),
    String(String),
    List(Vec<ResultData>),
    Dict(Vec<(ResultData, ResultData)>),
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

pub fn format_excel_number(f: f64) -> String {
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
    //
    // That is a much wider decimal range than the magnitude cutoffs this
    // used to apply (1e-5 .. 1e11), which turned e.g. 976121418126.432 --
    // which real Excel writes out in full -- into "9.76121418126432E+11".
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
        let suffix_len = format!("E{:+03}", exp).len();
        let frac_digits = 18usize.saturating_sub(suffix_len).min(14);

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
