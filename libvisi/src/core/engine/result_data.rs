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

    let abs_f = f.abs();
    let exp = abs_f.log10().floor() as i32;

    // Excel keeps plain decimal notation for as long as the decimal
    // rendering (at 15 significant digits, trailing zeros trimmed) stays
    // within 20 characters, and only then falls back to scientific.
    //
    // That is a much wider decimal range than the magnitude cutoffs this
    // used to apply (1e-5 .. 1e11), which turned e.g. 976121418126.432 --
    // which real Excel writes out in full -- into "9.76121418126432E+11".
    // Verified against real Excel: 1e18 and 1e19 render in full (19 and 20
    // characters) while 1e20 (21) goes scientific, and 0.000001207666770903
    // renders in full (20) while 0.00000120766677090395 (22) goes
    // scientific.
    let significant = {
        // Digits actually needed, once 15-significant-digit rounding has
        // trimmed whatever trailing zeros it produces.
        let mantissa = format!("{:.*e}", 14, abs_f);
        let digits = mantissa
            .split('e')
            .next()
            .unwrap_or("")
            .replace('.', "")
            .trim_end_matches('0')
            .len();
        digits.max(1)
    };
    let decimal_len = if exp >= 0 {
        let int_digits = (exp + 1) as usize;
        let frac_digits = significant.saturating_sub(int_digits);
        int_digits + usize::from(frac_digits > 0) + frac_digits
    } else {
        // "0." + leading zeros + significant digits
        2 + (-exp - 1) as usize + significant
    };

    if decimal_len > 20 {
        let mantissa = f / 10.0f64.powi(exp);
        let factor = 10.0f64.powi(14);
        let rounded_mantissa = (mantissa * factor).round() / factor;
        let mut s = format!("{:.14}", rounded_mantissa);
        if s.contains('.') {
            while s.ends_with('0') {
                s.pop();
            }
            if s.ends_with('.') {
                s.pop();
            }
        }
        format!("{}E{:+03}", s, exp)
    } else {
        let decimals = (14 - exp).clamp(0, 19) as usize;
        let formatted = format!("{:.1$}", f, decimals);
        let mut trimmed = formatted;
        if trimmed.contains('.') {
            while trimmed.ends_with('0') {
                trimmed.pop();
            }
            if trimmed.ends_with('.') {
                trimmed.pop();
            }
        }
        trimmed
    }
}
