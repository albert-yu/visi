// High-precision Text functions for libvisi
// Implements Excel-compatible string manipulation, unicode, formatting, splitting, joining, and search routines.

pub fn arraytotext(items: &[String], format: Option<f64>) -> Result<String, String> {
    let fmt = format.unwrap_or(0.0).round() as i32;
    if fmt == 1 {
        let quoted: Vec<String> = items.iter().map(|s| format!("\"{}\"", s)).collect();
        Ok(format!("{{{}}}", quoted.join("; ")))
    } else {
        Ok(items.join(", "))
    }
}

pub fn asc(text: &str) -> Result<String, String> {
    let mut res = String::new();
    for c in text.chars() {
        let code = c as u32;
        if (0xFF01..=0xFF5E).contains(&code) {
            if let Some(ch) = char::from_u32(code - 0xfee0) {
                res.push(ch);
            } else {
                res.push(c);
            }
        } else if code == 0x3000 {
            res.push(' ');
        } else {
            res.push(c);
        }
    }
    Ok(res)
}

pub fn jis(text: &str) -> Result<String, String> {
    let mut res = String::new();
    for c in text.chars() {
        let code = c as u32;
        if (0x0021..=0x007E).contains(&code) {
            if let Some(ch) = char::from_u32(code + 0xfee0) {
                res.push(ch);
            } else {
                res.push(c);
            }
        } else if c == ' ' {
            res.push('\u{3000}');
        } else {
            res.push(c);
        }
    }
    Ok(res)
}

pub fn bahttext(number: f64) -> Result<String, String> {
    if number.is_nan() || number.is_infinite() {
        return Err("#VALUE!".to_string());
    }
    let is_neg = number < 0.0;
    let abs_num = number.abs();
    let baht = abs_num.floor() as u64;
    let satang = ((abs_num - baht as f64) * 100.0).round() as u64;

    let digits = [
        "ศูนย์",
        "หนึ่ง",
        "สอง",
        "สาม",
        "สี่",
        "ห้า",
        "หก",
        "เจ็ด",
        "แปด",
        "เก้า",
    ];
    let positions = ["", "สิบ", "ร้อย", "พัน", "หมื่น", "แสน", "ล้าน"];

    fn convert_group(n: u64, digits: &[&str; 10], positions: &[&str; 7]) -> String {
        if n == 0 {
            return String::new();
        }
        let s = n.to_string();
        let len = s.len();
        let mut res = String::new();
        for (i, ch) in s.chars().enumerate() {
            let d = ch.to_digit(10).unwrap() as usize;
            let pos = len - 1 - i;
            if d != 0 {
                if pos == 0 && d == 1 && len > 1 {
                    res.push_str("เอ็ด");
                } else if pos == 1 && d == 2 {
                    res.push_str("ยี่สิบ");
                } else if pos == 1 && d == 1 {
                    res.push_str("สิบ");
                } else {
                    res.push_str(digits[d]);
                    res.push_str(positions[pos % 6]);
                }
            }
        }
        res
    }

    let mut result = String::new();
    if is_neg {
        result.push_str("ลบ");
    }

    if baht == 0 && satang == 0 {
        return Ok("ศูนย์บาทถ้วน".to_string());
    }

    if baht > 0 {
        result.push_str(&convert_group(baht, &digits, &positions));
        result.push_str("บาท");
    }

    if satang == 0 {
        result.push_str("ถ้วน");
    } else {
        result.push_str(&convert_group(satang, &digits, &positions));
        result.push_str("สตางค์");
    }

    Ok(result)
}

pub fn char_fn(number: f64) -> Result<String, String> {
    let n = number.floor() as u32;
    if !(1..=255).contains(&n) {
        Err("#VALUE!".to_string())
    } else if let Some(c) = char::from_u32(n) {
        Ok(c.to_string())
    } else {
        Err("#VALUE!".to_string())
    }
}

pub fn clean(text: &str) -> Result<String, String> {
    Ok(text.chars().filter(|&c| (c as u32) >= 32).collect())
}

pub fn code(text: &str) -> Result<f64, String> {
    if text.is_empty() {
        Err("#VALUE!".to_string())
    } else {
        let first = text.chars().next().unwrap();
        Ok(first as u32 as f64)
    }
}

pub fn dbcs(text: &str) -> Result<String, String> {
    let mut res = String::new();
    for c in text.chars() {
        let code = c as u32;
        if (0x0021..=0x007E).contains(&code) {
            if let Some(ch) = char::from_u32(code + 0xfee0) {
                res.push(ch);
            } else {
                res.push(c);
            }
        } else if code == 0x0020 {
            res.push('\u{3000}');
        } else {
            res.push(c);
        }
    }
    Ok(res)
}

pub fn detectlanguage(_text: &str) -> Result<String, String> {
    // Real language detection isn't implemented; always report English.
    Ok("en".to_string())
}

/// Rounds half away from zero at `decimals` places, the rule Excel's
/// DOLLAR/FIXED display uses. Rust's `format!("{:.N}", x)` instead rounds
/// the *binary* value to nearest-even, so a decimal literal like 3395.85
/// (whose nearest f64 is a hair below 3395.85) formats down to "3395.8"
/// where Excel shows "3395.9". Confirmed against real Excel:
/// DOLLAR(3395.85, 1) = "$3,395.9", DOLLAR(2.5, 0) = "$3",
/// DOLLAR(-2.5, 0) = "($3)".
fn round_half_away_from_zero(value: f64, decimals: usize) -> f64 {
    let factor = 10f64.powi(decimals as i32);
    let scaled = value * factor;
    // Nudge by one ulp-ish epsilon so a value that is only *just* below
    // the .5 boundary because of binary representation still rounds up.
    let eps = scaled.abs() * f64::EPSILON * 4.0;
    let adjusted = if scaled >= 0.0 {
        scaled + eps
    } else {
        scaled - eps
    };
    (adjusted.abs().round().copysign(adjusted)) / factor
}

pub fn dollar(number: f64, decimals: Option<f64>) -> Result<String, String> {
    let dec = decimals.unwrap_or(2.0).round() as usize;
    let is_neg = number < 0.0;
    let abs_num = round_half_away_from_zero(number, dec).abs();
    let formatted = format!("{:.*}", dec, abs_num);
    let parts: Vec<&str> = formatted.split('.').collect();
    let int_part = parts[0];

    let mut with_commas = String::new();
    let len = int_part.len();
    for (i, c) in int_part.chars().enumerate() {
        if i > 0 && (len - i).is_multiple_of(3) {
            with_commas.push(',');
        }
        with_commas.push(c);
    }

    let num_str = if parts.len() > 1 {
        format!("{}.{}", with_commas, parts[1])
    } else {
        with_commas
    };

    if is_neg {
        Ok(format!("(${})", num_str))
    } else {
        Ok(format!("${}", num_str))
    }
}

pub fn exact(text1: &str, text2: &str) -> Result<bool, String> {
    Ok(text1 == text2)
}

pub fn find(find_text: &str, within_text: &str, start_num: Option<f64>) -> Result<f64, String> {
    let start = start_num.unwrap_or(1.0).floor() as usize;
    if start < 1 {
        return Err("#VALUE!".to_string());
    }
    let chars: Vec<char> = within_text.chars().collect();
    if start > chars.len() + 1 {
        return Err("#VALUE!".to_string());
    }
    let search_slice: String = chars[start - 1..].iter().collect();

    if let Some(pos) = search_slice.find(find_text) {
        let prefix = &search_slice[..pos];
        Ok((start + prefix.chars().count()) as f64)
    } else {
        Err("#VALUE!".to_string())
    }
}

pub fn fixed(
    number: f64,
    decimals: Option<f64>,
    no_commas: Option<bool>,
) -> Result<String, String> {
    let dec = decimals.unwrap_or(2.0).round() as usize;
    let skip_commas = no_commas.unwrap_or(false);
    let is_neg = number < 0.0;
    let abs_num = round_half_away_from_zero(number, dec).abs();
    let formatted = format!("{:.*}", dec, abs_num);

    if skip_commas {
        if is_neg {
            Ok(format!("-{}", formatted))
        } else {
            Ok(formatted)
        }
    } else {
        let parts: Vec<&str> = formatted.split('.').collect();
        let int_part = parts[0];
        let mut with_commas = String::new();
        let len = int_part.len();
        for (i, c) in int_part.chars().enumerate() {
            if i > 0 && (len - i).is_multiple_of(3) {
                with_commas.push(',');
            }
            with_commas.push(c);
        }
        let res = if parts.len() > 1 {
            format!("{}.{}", with_commas, parts[1])
        } else {
            with_commas
        };
        if is_neg {
            Ok(format!("-{}", res))
        } else {
            Ok(res)
        }
    }
}

pub fn numbervalue(
    text: &str,
    decimal_sep: Option<&str>,
    group_sep: Option<&str>,
) -> Result<f64, String> {
    let dec = decimal_sep.unwrap_or(".");
    let group = group_sep.unwrap_or(",");
    let mut cleaned = text.trim().to_string();
    cleaned = cleaned.replace(group, "");
    cleaned = cleaned.replace(dec, ".");
    match cleaned.parse::<f64>() {
        Ok(v) => Ok(v),
        Err(_) => Err("#VALUE!".to_string()),
    }
}

pub fn phonetic(reference: &str) -> Result<String, String> {
    Ok(reference.to_string())
}

pub fn regexextract(text: &str, pattern: &str) -> Result<String, String> {
    let re = regex::Regex::new(pattern).map_err(|_| "#VALUE!".to_string())?;
    match re.find(text) {
        Some(m) => Ok(m.as_str().to_string()),
        None => Err("#N/A".to_string()),
    }
}

pub fn regexreplace(text: &str, pattern: &str, replacement: &str) -> Result<String, String> {
    let re = regex::Regex::new(pattern).map_err(|_| "#VALUE!".to_string())?;
    Ok(re.replace_all(text, replacement).into_owned())
}

pub fn regextest(text: &str, pattern: &str) -> Result<bool, String> {
    let re = regex::Regex::new(pattern).map_err(|_| "#VALUE!".to_string())?;
    Ok(re.is_match(text))
}

pub fn replace_fn(
    old_text: &str,
    start_num: f64,
    num_chars: f64,
    new_text: &str,
) -> Result<String, String> {
    let start = start_num.floor() as usize;
    let n = num_chars.floor() as usize;
    if start < 1 {
        return Err("#VALUE!".to_string());
    }
    let chars: Vec<char> = old_text.chars().collect();
    let start_idx = (start - 1).min(chars.len());
    let end_idx = (start_idx + n).min(chars.len());

    let mut res = String::new();
    res.extend(&chars[..start_idx]);
    res.push_str(new_text);
    res.extend(&chars[end_idx..]);
    Ok(res)
}

pub fn rept(text: &str, count: f64) -> Result<String, String> {
    let cnt = count.floor() as usize;
    if count < 0.0 {
        Err("#VALUE!".to_string())
    } else {
        Ok(text.repeat(cnt))
    }
}

pub fn search(find_text: &str, within_text: &str, start_num: Option<f64>) -> Result<f64, String> {
    let start = start_num.unwrap_or(1.0).floor() as usize;
    if start < 1 {
        return Err("#VALUE!".to_string());
    }
    let lower_find = find_text.to_lowercase();
    let lower_within = within_text.to_lowercase();
    let chars: Vec<char> = lower_within.chars().collect();
    if start > chars.len() + 1 {
        return Err("#VALUE!".to_string());
    }
    let search_slice: String = chars[start - 1..].iter().collect();

    // Simple wildcard handling for ? and *
    let clean_find = lower_find.replace(['?', '*'], "");
    if clean_find.is_empty() {
        return Ok(start as f64);
    }
    if let Some(pos) = search_slice.find(&clean_find) {
        let prefix = &search_slice[..pos];
        Ok((start + prefix.chars().count()) as f64)
    } else {
        Err("#VALUE!".to_string())
    }
}

pub fn substitute(
    text: &str,
    old_text: &str,
    new_text: &str,
    instance: Option<f64>,
) -> Result<String, String> {
    if old_text.is_empty() {
        return Ok(text.to_string());
    }
    if let Some(inst_val) = instance {
        let inst = inst_val.floor() as usize;
        if inst < 1 {
            return Err("#VALUE!".to_string());
        }
        let mut curr_inst = 0;
        let mut res = String::new();
        let mut last_idx = 0;
        for (idx, _) in text.match_indices(old_text) {
            curr_inst += 1;
            if curr_inst == inst {
                res.push_str(&text[last_idx..idx]);
                res.push_str(new_text);
                res.push_str(&text[idx + old_text.len()..]);
                return Ok(res);
            }
            last_idx = idx;
        }
        Ok(text.to_string())
    } else {
        Ok(text.replace(old_text, new_text))
    }
}

pub fn t_fn(val: &str, is_string: bool) -> String {
    if is_string {
        val.to_string()
    } else {
        String::new()
    }
}

fn add_thousands_separators(digits: &str) -> String {
    let bytes = digits.as_bytes();
    let mut out = String::with_capacity(digits.len() + digits.len() / 3);
    for (i, b) in bytes.iter().enumerate() {
        if i > 0 && (bytes.len() - i).is_multiple_of(3) {
            out.push(',');
        }
        out.push(*b as char);
    }
    out
}

fn format_date_text(val: f64, format_text: &str) -> Result<String, String> {
    if val < 0.0 {
        return Err("#VALUE!".to_string());
    }
    let (y, m, d) = crate::core::date_fn::serial_to_ymd(val);
    // Longest tokens first so "yyyy" isn't partially eaten by a "yy" pass.
    let replaced = format_text
        .replace("yyyy", &format!("{:04}", y))
        .replace("yy", &format!("{:02}", y % 100))
        .replace("mm", &format!("{:02}", m))
        .replace("dd", &format!("{:02}", d));
    Ok(replaced)
}

/// A pragmatic subset of Excel's TEXT() number-format mini-language: `$`
/// currency prefix, `%` percentage (value *100, suffixed), `,` thousands
/// grouping, and `.0...`/`.#...` decimal-place count, plus common date
/// tokens (`yyyy`/`yy`/`mm`/`dd`). Not a full format-code parser (no
/// custom positive/negative/zero sections, no scientific notation, no
/// fractions, ...) -- covers what this engine's own formula generation
/// and fuzzing actually exercise.
pub fn text_fn(val: f64, format_text: &str) -> Result<String, String> {
    let fmt = format_text.trim();

    let has_date_tokens =
        (fmt.contains('y') || fmt.contains('d')) && !fmt.contains('0') && !fmt.contains('#');
    if has_date_tokens {
        return format_date_text(val, fmt);
    }

    let has_currency = fmt.contains('$');
    let has_percent = fmt.contains('%');
    let has_comma = fmt.contains(',');
    let dec_count = match fmt.find('.') {
        Some(idx) => fmt[idx + 1..]
            .chars()
            .take_while(|c| *c == '0' || *c == '#')
            .count(),
        None => 0,
    };

    let scaled = if has_percent { val * 100.0 } else { val };
    let is_negative = scaled < 0.0;
    // Excel rounds half away from zero on the decimal it shows, whereas
    // Rust's `{:.N}` rounds the underlying *binary* value to nearest-even:
    // TEXT(-3873.705, "0.00") is -3873.71 in Excel but formats as -3873.70
    // here without this. Same rule DOLLAR/FIXED needed.
    let formatted = format!(
        "{:.*}",
        dec_count,
        round_half_away_from_zero(scaled, dec_count).abs()
    );
    let (int_part, dec_part) = match formatted.split_once('.') {
        Some((i, d)) => (i.to_string(), Some(d.to_string())),
        None => (formatted, None),
    };
    let int_part = if has_comma {
        add_thousands_separators(&int_part)
    } else {
        int_part
    };

    let mut result = int_part;
    if let Some(d) = dec_part {
        result.push('.');
        result.push_str(&d);
    }
    if has_currency {
        result.insert(0, '$');
    }
    if is_negative {
        result.insert(0, '-');
    }
    if has_percent {
        result.push('%');
    }
    Ok(result)
}

pub fn textafter(text: &str, delimiter: &str, instance: Option<f64>) -> Result<String, String> {
    let inst = instance.unwrap_or(1.0).floor() as usize;
    if inst < 1 || delimiter.is_empty() {
        return Err("#VALUE!".to_string());
    }
    let mut count = 0;
    for (idx, _) in text.match_indices(delimiter) {
        count += 1;
        if count == inst {
            return Ok(text[idx + delimiter.len()..].to_string());
        }
    }
    Err("#N/A".to_string())
}

pub fn textbefore(text: &str, delimiter: &str, instance: Option<f64>) -> Result<String, String> {
    let inst = instance.unwrap_or(1.0).floor() as usize;
    if inst < 1 || delimiter.is_empty() {
        return Err("#VALUE!".to_string());
    }
    let mut count = 0;
    for (idx, _) in text.match_indices(delimiter) {
        count += 1;
        if count == inst {
            return Ok(text[..idx].to_string());
        }
    }
    Err("#N/A".to_string())
}

pub fn textjoin(delimiter: &str, ignore_empty: bool, texts: &[String]) -> Result<String, String> {
    let filtered: Vec<&str> = texts
        .iter()
        .map(|s| s.as_str())
        .filter(|s| !ignore_empty || !s.is_empty())
        .collect();
    Ok(filtered.join(delimiter))
}

pub fn textsplit(text: &str, col_delim: &str) -> Result<Vec<String>, String> {
    if col_delim.is_empty() {
        return Ok(vec![text.to_string()]);
    }
    Ok(text.split(col_delim).map(|s| s.to_string()).collect())
}

pub fn translate(text: &str, _from: &str, _to: &str) -> Result<String, String> {
    Ok(text.to_string())
}

pub fn unichar(number: f64) -> Result<String, String> {
    let n = number.floor() as u32;
    if let Some(c) = char::from_u32(n) {
        Ok(c.to_string())
    } else {
        Err("#VALUE!".to_string())
    }
}

pub fn unicode(text: &str) -> Result<f64, String> {
    if text.is_empty() {
        Err("#VALUE!".to_string())
    } else {
        let first = text.chars().next().unwrap();
        Ok(first as u32 as f64)
    }
}

pub fn value(text: &str) -> Result<f64, String> {
    let s = text.trim();
    if let Ok(v) = s.parse::<f64>() {
        Ok(v)
    } else {
        Err("#VALUE!".to_string())
    }
}

pub fn valuetotext(val: &str, format: Option<f64>) -> Result<String, String> {
    let fmt = format.unwrap_or(0.0).round() as i32;
    if fmt == 1 {
        Ok(format!("\"{}\"", val))
    } else {
        Ok(val.to_string())
    }
}
