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
    Plot {
        points: Vec<(f32, f32)>,
        color: [f32; 4],
        radius: f32,
        is_line: bool,
        title: Option<String>,
        xlabel: Option<String>,
        ylabel: Option<String>,
    },
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
            ResultData::Plot { .. } => write!(f, "[Plot]"),
            ResultData::Error(e) => write!(f, "Error: {}", e),
        }
    }
}

impl ResultData {
    pub fn plot_cell_dims(&self) -> Option<(usize, usize)> {
        if let ResultData::Plot { points, .. } = self {
            if points.is_empty() {
                return Some((16, 16));
            }
            let mut x_min = f32::INFINITY;
            let mut x_max = f32::NEG_INFINITY;
            let mut y_min = f32::INFINITY;
            let mut y_max = f32::NEG_INFINITY;
            for &(x, y) in points {
                if x < x_min {
                    x_min = x;
                }
                if x > x_max {
                    x_max = x;
                }
                if y < y_min {
                    y_min = y;
                }
                if y > y_max {
                    y_max = y;
                }
            }
            if !x_min.is_finite() || !x_max.is_finite() || !y_min.is_finite() || !y_max.is_finite()
            {
                return Some((16, 16));
            }
            let dx = x_max - x_min;
            let dy = y_max - y_min;
            if dx == 0.0 && dy == 0.0 {
                return Some((16, 16));
            }

            // Proportional scaling: the larger of the two dimensions should be 15
            let w;
            let h;
            if dx >= dy {
                w = 15;
                if dx > 0.0 {
                    h = ((15.0 * dy / dx).round() as usize).max(1);
                } else {
                    h = 1;
                }
            } else {
                h = 15;
                if dy > 0.0 {
                    w = ((15.0 * dx / dy).round() as usize).max(1);
                } else {
                    w = 1;
                }
            }
            // Add a column to the left and a row to the bottom for axis labels
            Some((w + 1, h + 1))
        } else {
            None
        }
    }
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

        // Re-rounding can carry into the exponent (9.99e5 -> 1.0e6), so
        // take the exponent from this rendering rather than the earlier one.
        let sci = format!("{:.*e}", frac_digits, f);
        let (mantissa, exp_str) = sci.split_once('e').expect("{:e} always emits an exponent");
        let exp: i32 = exp_str.parse().expect("{:e} emits an integer exponent");
        let mut mantissa = mantissa.to_string();
        if mantissa.contains('.') {
            while mantissa.ends_with('0') {
                mantissa.pop();
            }
            if mantissa.ends_with('.') {
                mantissa.pop();
            }
        }
        format!("{}E{:+03}", mantissa, exp)
    }
}
