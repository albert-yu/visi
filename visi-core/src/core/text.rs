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
    Ok("en".to_string())
}

fn round_half_away_from_zero(value: f64, decimals: usize) -> f64 {
    let factor = 10f64.powi(decimals as i32);
    let scaled = value * factor;
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
    Ok(crate::core::date::render_date_code(
        crate::core::date::excel_serial_to_date(val),
        format_text,
        crate::core::date::StringCase::Title,
    ))
}

fn split_format_sections(format_text: &str) -> Vec<String> {
    let mut sections = Vec::new();
    let mut current = String::new();
    let mut in_quote = false;
    let mut escaped = false;
    for c in format_text.chars() {
        if escaped {
            current.push(c);
            escaped = false;
            continue;
        }
        if c == '\\' {
            current.push(c);
            escaped = true;
            continue;
        }
        if c == '"' {
            in_quote = !in_quote;
            current.push(c);
            continue;
        }
        if c == ';' && !in_quote {
            sections.push(current.trim().to_string());
            current.clear();
        } else {
            current.push(c);
        }
    }
    sections.push(current.trim().to_string());
    sections
}

fn section_for_number(sections: &[String], val: f64) -> (&str, bool) {
    if val < 0.0 {
        if sections.len() >= 2 && !sections[1].is_empty() {
            return (&sections[1], false);
        }
        return (
            sections.first().map(String::as_str).unwrap_or("General"),
            true,
        );
    }
    if val == 0.0 && sections.len() >= 3 && !sections[2].is_empty() {
        return (&sections[2], false);
    }
    (
        sections.first().map(String::as_str).unwrap_or("General"),
        false,
    )
}

fn clean_format_literal(text: &[char]) -> String {
    let mut out = String::new();
    let mut i = 0;
    while i < text.len() {
        match text[i] {
            '"' => {
                i += 1;
                while i < text.len() && text[i] != '"' {
                    out.push(text[i]);
                    i += 1;
                }
                if i < text.len() {
                    i += 1;
                }
            }
            '\\' => {
                if let Some(next) = text.get(i + 1) {
                    out.push(*next);
                    i += 2;
                } else {
                    i += 1;
                }
            }
            '_' | '*' => i += 2,
            '[' => {
                i += 1;
                while i < text.len() && text[i] != ']' {
                    i += 1;
                }
                if i < text.len() {
                    i += 1;
                }
            }
            c if c == '@' => {
                out.push(c);
                i += 1;
            }
            '0' | '#' | '?' | '.' | ',' => i += 1,
            c => {
                out.push(c);
                i += 1;
            }
        }
    }
    out
}

fn placeholder_bounds(chars: &[char]) -> Option<(usize, usize)> {
    let first = chars.iter().position(|c| matches!(c, '0' | '#' | '?'))?;
    let mut last = first;
    let mut i = first;
    while i < chars.len() {
        match chars[i] {
            '0' | '#' | '?' | '.' | ',' | '+' | '-' => {
                last = i;
                i += 1;
            }
            'E' | 'e' => {
                last = i;
                i += 1;
                if i < chars.len() && matches!(chars[i], '+' | '-') {
                    last = i;
                    i += 1;
                }
            }
            _ => i += 1,
        }
    }
    Some((first, last))
}

fn trim_optional_decimal(mut decimals: String, tokens: &[char]) -> String {
    while decimals.len() > tokens.iter().filter(|c| **c == '0').count()
        && decimals.ends_with('0')
        && tokens
            .get(decimals.len().saturating_sub(1))
            .is_some_and(|c| *c != '0')
    {
        decimals.pop();
    }
    decimals
}

fn format_fixed_number(abs_val: f64, pattern: &str) -> String {
    let pattern_chars: Vec<char> = pattern.chars().collect();
    let decimal_idx = pattern_chars.iter().position(|c| *c == '.');
    let int_pattern: String = pattern_chars[..decimal_idx.unwrap_or(pattern_chars.len())]
        .iter()
        .collect();
    let dec_pattern: String = decimal_idx
        .map(|idx| pattern_chars[idx + 1..].iter().collect())
        .unwrap_or_default();
    let dec_tokens: Vec<char> = dec_pattern
        .chars()
        .take_while(|c| matches!(c, '0' | '#' | '?'))
        .collect();
    let max_dec = dec_tokens.len();
    let min_dec = dec_tokens.iter().filter(|c| **c == '0').count();
    let min_int = int_pattern.chars().filter(|c| *c == '0').count();
    let grouping = int_pattern.contains(',');
    let rounded = round_half_away_from_zero(abs_val, max_dec).abs();
    let formatted = format!("{:.*}", max_dec, rounded);
    let (mut int_part, dec_part) = match formatted.split_once('.') {
        Some((i, d)) => (i.to_string(), d.to_string()),
        None => (formatted, String::new()),
    };
    if int_part.len() < min_int {
        int_part = format!("{}{}", "0".repeat(min_int - int_part.len()), int_part);
    }
    if grouping {
        int_part = add_thousands_separators(&int_part);
    }
    let mut out = int_part;
    let decimals = if max_dec == 0 {
        String::new()
    } else {
        trim_optional_decimal(dec_part, &dec_tokens)
    };
    if !decimals.is_empty() || min_dec > 0 || (decimal_idx.is_some() && max_dec > 0) {
        out.push('.');
        out.push_str(&decimals);
        if decimals.len() < min_dec {
            out.push_str(&"0".repeat(min_dec - decimals.len()));
        }
    }
    out
}

fn format_scientific_number(abs_val: f64, pattern: &str) -> Option<String> {
    let upper = pattern.to_ascii_uppercase();
    let (mantissa_pattern, exponent_pattern) = upper.split_once('E')?;
    let dec_count = mantissa_pattern
        .split_once('.')
        .map(|(_, d)| d.chars().filter(|c| matches!(c, '0' | '#' | '?')).count())
        .unwrap_or(0);
    let exp_digits = exponent_pattern
        .chars()
        .filter(|c| matches!(c, '0' | '#' | '?'))
        .count()
        .max(1);
    let raw = format!("{:.*E}", dec_count, abs_val);
    let (mantissa, exp) = raw.split_once('E')?;
    let exp_num: i32 = exp.parse().ok()?;
    let sign = if exponent_pattern.contains('-') && exp_num < 0 {
        "-"
    } else if exponent_pattern.contains('+') {
        if exp_num < 0 { "-" } else { "+" }
    } else {
        ""
    };
    Some(format!(
        "{mantissa}E{sign}{:0width$}",
        exp_num.abs(),
        width = exp_digits
    ))
}

fn is_time_format_code(code: &str) -> bool {
    let mut in_quote = false;
    let mut in_bracket = false;
    let mut escaped = false;
    for c in code.chars() {
        if escaped {
            escaped = false;
            continue;
        }
        match c {
            '\\' => escaped = true,
            '"' if !in_bracket => in_quote = !in_quote,
            '[' if !in_quote => in_bracket = true,
            ']' if in_bracket => in_bracket = false,
            _ if in_quote || in_bracket => {}
            _ if matches!(c.to_ascii_lowercase(), 'h' | 's') => return true,
            _ => {}
        }
    }
    false
}

pub(crate) fn format_number_format(val: f64, format_text: &str) -> Result<String, String> {
    let fmt = format_text.trim();
    if fmt.is_empty() || fmt.eq_ignore_ascii_case("General") || is_time_format_code(fmt) {
        return Ok(crate::core::engine::result_data::format_excel_number(val));
    }
    if crate::core::date::is_date_code(fmt) {
        return format_date_text(val, fmt);
    }

    let sections = split_format_sections(fmt);
    let (section, use_default_minus) = section_for_number(&sections, val);
    let chars: Vec<char> = section.chars().collect();
    let Some((first, last)) = placeholder_bounds(&chars) else {
        return Ok(clean_format_literal(&chars));
    };
    let prefix = clean_format_literal(&chars[..first]);
    let suffix = clean_format_literal(&chars[last + 1..]);
    let number_pattern: String = chars[first..=last]
        .iter()
        .filter(|c| matches!(c, '0' | '#' | '?' | '.' | ',' | 'E' | 'e' | '+' | '-'))
        .collect();
    let percent_count = section.chars().filter(|c| *c == '%').count() as i32;
    let scaled = val.abs() * 100f64.powi(percent_count);
    let number = if number_pattern.to_ascii_uppercase().contains('E') {
        format_scientific_number(scaled, &number_pattern)
            .unwrap_or_else(|| format_fixed_number(scaled, &number_pattern))
    } else {
        format_fixed_number(scaled, &number_pattern)
    };
    let mut out = String::new();
    if use_default_minus {
        out.push('-');
    }
    out.push_str(&prefix);
    out.push_str(&number);
    out.push_str(&suffix);
    Ok(out)
}

pub(crate) fn format_text_format(val: &str, format_text: &str) -> Option<String> {
    let sections = split_format_sections(format_text);
    let section = if sections.len() >= 4 {
        sections[3].as_str()
    } else {
        sections.first().map(String::as_str).unwrap_or("@")
    };
    if !section.contains('@') {
        return None;
    }
    Some(clean_format_literal(&section.chars().collect::<Vec<_>>()).replace('@', val))
}

pub fn text_fn(val: f64, format_text: &str) -> Result<String, String> {
    format_number_format(val, format_text)
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

#[allow(dead_code)]
pub fn value(text: &str) -> Result<f64, String> {
    value_with_locale(text, &crate::core::locale::Locale::en_us())
}

pub fn value_with_locale(text: &str, locale: &crate::core::locale::Locale) -> Result<f64, String> {
    let s = text.trim();
    if let Ok(v) = s.parse::<f64>() {
        Ok(v)
    } else if let Some((date, _)) = crate::core::date::parse_date_with_locale(s, locale) {
        Ok(crate::core::date::date_to_excel_serial(date))
    } else if let Some(f) = crate::core::date_fn::parse_time_fraction(s) {
        Ok(f)
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
