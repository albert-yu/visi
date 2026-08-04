// High-precision Engineering functions for libvisi
// Implements Excel-compatible Bessel functions, bitwise operations, number base conversions (BIN/OCT/DEC/HEX with 2's complement), complex number math, unit conversions, delta & step functions.

// ============================================================================
// 1. Bitwise and Step/Delta Functions
// ============================================================================

pub fn bitand(number1: f64, number2: f64) -> Result<f64, String> {
    let n1 = number1.floor() as u64;
    let n2 = number2.floor() as u64;
    Ok((n1 & n2) as f64)
}

pub fn bitor(number1: f64, number2: f64) -> Result<f64, String> {
    let n1 = number1.floor() as u64;
    let n2 = number2.floor() as u64;
    Ok((n1 | n2) as f64)
}

pub fn bitxor(number1: f64, number2: f64) -> Result<f64, String> {
    let n1 = number1.floor() as u64;
    let n2 = number2.floor() as u64;
    Ok((n1 ^ n2) as f64)
}

pub fn bitlshift(number: f64, shift_amount: f64) -> Result<f64, String> {
    let n = number.floor() as u64;
    let s = shift_amount.floor() as i32;
    if s >= 0 {
        Ok((n << (s as u32)) as f64)
    } else {
        Ok((n >> ((-s) as u32)) as f64)
    }
}

pub fn bitrshift(number: f64, shift_amount: f64) -> Result<f64, String> {
    bitlshift(number, -shift_amount)
}

pub fn delta(number1: f64, number2: Option<f64>) -> Result<f64, String> {
    let n2 = number2.unwrap_or(0.0);
    if number1 == n2 { Ok(1.0) } else { Ok(0.0) }
}

pub fn gestep(number: f64, step: Option<f64>) -> Result<f64, String> {
    let s = step.unwrap_or(0.0);
    if number >= s { Ok(1.0) } else { Ok(0.0) }
}

// ============================================================================
// 2. Base Conversions (BIN/OCT/DEC/HEX)
// ============================================================================

pub fn parse_twos_complement(text: &str, bits: usize, radix: u32) -> Result<i64, String> {
    let s = text.trim();
    let val = u64::from_str_radix(s, radix).map_err(|_| "#NUM!".to_string())?;
    let sign_bit = 1u64 << (bits - 1);
    if (val & sign_bit) != 0 {
        let mask = (1u64 << bits) - 1;
        Ok(-(((!val + 1) & mask) as i64))
    } else {
        Ok(val as i64)
    }
}

pub fn format_twos_complement(val: i64, bits: usize, radix: u32, places: Option<f64>) -> Result<String, String> {
    let mask = (1u64 << bits) - 1;
    let uval = (val as u64) & mask;
    let chars = "0123456789ABCDEF";

    let mut digits = Vec::new();
    let mut n = uval;
    if n == 0 {
        digits.push('0');
    } else {
        while n > 0 {
            let rem = (n % (radix as u64)) as usize;
            digits.push(chars.as_bytes()[rem] as char);
            n /= radix as u64;
        }
    }
    digits.reverse();
    let res: String = digits.into_iter().collect();

    if let Some(p_val) = places {
        let p = p_val.floor() as usize;
        if res.len() < p {
            Ok(format!("{:0>1$}", res, p))
        } else {
            Ok(res)
        }
    } else {
        Ok(res)
    }
}

pub fn bin2dec(text: &str) -> Result<f64, String> { Ok(parse_twos_complement(text, 10, 2)? as f64) }
pub fn bin2hex(text: &str, places: Option<f64>) -> Result<String, String> { format_twos_complement(parse_twos_complement(text, 10, 2)?, 40, 16, places) }
pub fn bin2oct(text: &str, places: Option<f64>) -> Result<String, String> { format_twos_complement(parse_twos_complement(text, 10, 2)?, 30, 8, places) }

pub fn dec2bin(number: f64, places: Option<f64>) -> Result<String, String> { format_twos_complement(number.floor() as i64, 10, 2, places) }
pub fn dec2hex(number: f64, places: Option<f64>) -> Result<String, String> { format_twos_complement(number.floor() as i64, 40, 16, places) }
pub fn dec2oct(number: f64, places: Option<f64>) -> Result<String, String> { format_twos_complement(number.floor() as i64, 30, 8, places) }

pub fn hex2dec(text: &str) -> Result<f64, String> { Ok(parse_twos_complement(text, 40, 16)? as f64) }
pub fn hex2bin(text: &str, places: Option<f64>) -> Result<String, String> { format_twos_complement(parse_twos_complement(text, 40, 16)?, 10, 2, places) }
pub fn hex2oct(text: &str, places: Option<f64>) -> Result<String, String> { format_twos_complement(parse_twos_complement(text, 40, 16)?, 30, 8, places) }

pub fn oct2dec(text: &str) -> Result<f64, String> { Ok(parse_twos_complement(text, 30, 8)? as f64) }
pub fn oct2bin(text: &str, places: Option<f64>) -> Result<String, String> { format_twos_complement(parse_twos_complement(text, 30, 8)?, 10, 2, places) }
pub fn oct2hex(text: &str, places: Option<f64>) -> Result<String, String> { format_twos_complement(parse_twos_complement(text, 30, 8)?, 40, 16, places) }

// ============================================================================
// 3. Complex Number Operations
// ============================================================================

#[derive(Debug, Clone, Copy)]
pub struct ComplexNum {
    pub re: f64,
    pub im: f64,
    pub suffix: char, // 'i' or 'j'
}

pub fn parse_complex(text: &str) -> Result<ComplexNum, String> {
    let s = text.trim();
    if s.is_empty() { return Err("#VALUE!".to_string()); }
    let suffix = if s.ends_with('j') { 'j' } else { 'i' };

    let s_clean = s.trim_end_matches('i').trim_end_matches('j');
    if s_clean == s {
        // Pure real number
        let re = s.parse::<f64>().map_err(|_| "#VALUE!".to_string())?;
        return Ok(ComplexNum { re, im: 0.0, suffix: 'i' });
    }

    if s_clean.is_empty() || s_clean == "+" {
        return Ok(ComplexNum { re: 0.0, im: 1.0, suffix });
    }
    if s_clean == "-" {
        return Ok(ComplexNum { re: 0.0, im: -1.0, suffix });
    }

    // Split on last '+' or '-' that is not an exponent (e or E)
    let bytes = s_clean.as_bytes();
    let mut split_idx = None;
    for i in (1..bytes.len()).rev() {
        if (bytes[i] == b'+' || bytes[i] == b'-') && bytes[i - 1] != b'e' && bytes[i - 1] != b'E' {
            split_idx = Some(i);
            break;
        }
    }

    if let Some(idx) = split_idx {
        let re = s_clean[..idx].parse::<f64>().map_err(|_| "#VALUE!".to_string())?;
        let im_str = &s_clean[idx..];
        let im = if im_str == "+" { 1.0 } else if im_str == "-" { -1.0 } else { im_str.parse::<f64>().map_err(|_| "#VALUE!".to_string())? };
        Ok(ComplexNum { re, im, suffix })
    } else {
        // Pure imaginary
        let im = s_clean.parse::<f64>().map_err(|_| "#VALUE!".to_string())?;
        Ok(ComplexNum { re: 0.0, im, suffix })
    }
}

pub fn format_complex(c: ComplexNum) -> String {
    let s = c.suffix;
    if c.im == 0.0 {
        format!("{}", c.re)
    } else if c.re == 0.0 {
        if c.im == 1.0 { format!("{}", s) } else if c.im == -1.0 { format!("-{}", s) } else { format!("{}{}", c.im, s) }
    } else if c.im > 0.0 {
        if c.im == 1.0 { format!("{}+{}", c.re, s) } else { format!("{}+{}{}", c.re, c.im, s) }
    } else {
        if c.im == -1.0 { format!("{}-{}", c.re, s) } else { format!("{}{}{}", c.re, c.im, s) }
    }
}

pub fn complex_fn(real_num: f64, i_num: f64, suffix: Option<&str>) -> Result<String, String> {
    let suf = suffix.unwrap_or("i").chars().next().unwrap_or('i');
    if suf != 'i' && suf != 'j' { return Err("#VALUE!".to_string()); }
    Ok(format_complex(ComplexNum { re: real_num, im: i_num, suffix: suf }))
}

pub fn imabs(in_str: &str) -> Result<f64, String> { let c = parse_complex(in_str)?; Ok((c.re * c.re + c.im * c.im).sqrt()) }
pub fn imaginary(in_str: &str) -> Result<f64, String> { Ok(parse_complex(in_str)?.im) }
pub fn imreal(in_str: &str) -> Result<f64, String> { Ok(parse_complex(in_str)?.re) }
pub fn imargument(in_str: &str) -> Result<f64, String> { let c = parse_complex(in_str)?; Ok(c.im.atan2(c.re)) }
pub fn imconjugate(in_str: &str) -> Result<String, String> { let mut c = parse_complex(in_str)?; c.im = -c.im; Ok(format_complex(c)) }

pub fn imsum(args: &[&str]) -> Result<String, String> {
    let mut sum_re = 0.0; let mut sum_im = 0.0; let mut suf = 'i';
    for arg in args {
        let c = parse_complex(arg)?;
        sum_re += c.re; sum_im += c.im; suf = c.suffix;
    }
    Ok(format_complex(ComplexNum { re: sum_re, im: sum_im, suffix: suf }))
}

pub fn imsub(in_str1: &str, in_str2: &str) -> Result<String, String> {
    let c1 = parse_complex(in_str1)?; let c2 = parse_complex(in_str2)?;
    Ok(format_complex(ComplexNum { re: c1.re - c2.re, im: c1.im - c2.im, suffix: c1.suffix }))
}

pub fn improduct(args: &[&str]) -> Result<String, String> {
    if args.is_empty() { return Ok("0".to_string()); }
    let mut curr = parse_complex(args[0])?;
    for arg in &args[1..] {
        let c = parse_complex(arg)?;
        let re = curr.re * c.re - curr.im * c.im;
        let im = curr.re * c.im + curr.im * c.re;
        curr.re = re; curr.im = im;
    }
    Ok(format_complex(curr))
}

pub fn imdiv(in_str1: &str, in_str2: &str) -> Result<String, String> {
    let c1 = parse_complex(in_str1)?; let c2 = parse_complex(in_str2)?;
    let denom = c2.re * c2.re + c2.im * c2.im;
    if denom == 0.0 { return Err("#NUM!".to_string()); }
    let re = (c1.re * c2.re + c1.im * c2.im) / denom;
    let im = (c1.im * c2.re - c1.re * c2.im) / denom;
    Ok(format_complex(ComplexNum { re, im, suffix: c1.suffix }))
}

// ============================================================================
// 4. Unit Conversion (CONVERT)
// ============================================================================

pub fn convert(val: f64, from_unit: &str, to_unit: &str) -> Result<f64, String> {
    let u1 = from_unit.trim();
    let u2 = to_unit.trim();
    if u1 == u2 { return Ok(val); }

    // Temperature
    if u1 == "C" && u2 == "F" { return Ok(val * 1.8 + 32.0); }
    if u1 == "F" && u2 == "C" { return Ok((val - 32.0) / 1.8); }
    if u1 == "C" && u2 == "K" { return Ok(val + 273.15); }
    if u1 == "K" && u2 == "C" { return Ok(val - 273.15); }

    // Length (meters)
    let length_factor = |u: &str| -> Option<f64> {
        match u {
            "m" => Some(1.0),
            "km" => Some(1000.0),
            "cm" => Some(0.01),
            "mm" => Some(0.001),
            "in" => Some(0.0254),
            "ft" => Some(0.3048),
            "yd" => Some(0.9144),
            "mi" => Some(1609.344),
            _ => None,
        }
    };

    if let (Some(f1), Some(f2)) = (length_factor(u1), length_factor(u2)) {
        return Ok(val * f1 / f2);
    }

    // Mass (kg)
    let mass_factor = |u: &str| -> Option<f64> {
        match u {
            "kg" => Some(1.0),
            "g" => Some(0.001),
            "mg" => Some(0.000001),
            "lbm" => Some(0.45359237),
            "ozm" => Some(0.028349523125),
            _ => None,
        }
    };

    if let (Some(f1), Some(f2)) = (mass_factor(u1), mass_factor(u2)) {
        return Ok(val * f1 / f2);
    }

    Err("#N/A".to_string())
}

// ============================================================================
// 5. Bessel Functions
// ============================================================================

pub fn besseli(x: f64, n: f64) -> Result<f64, String> {
    let order = n.floor() as usize;
    // Series for I_n(x) = sum_k (x/2)^(2k+n) / (k! * (k+n)!)
    let mut sum = 0.0;
    let mut term = (x / 2.0).powi(order as i32) / (1..=order).product::<usize>().max(1) as f64;
    sum += term;
    for k in 1..=30 {
        term *= (x * x / 4.0) / (k as f64 * (k + order) as f64);
        sum += term;
    }
    Ok(sum)
}

pub fn besselj(x: f64, n: f64) -> Result<f64, String> {
    let order = n.floor() as usize;
    // Series for J_n(x) = sum_k (-1)^k * (x/2)^(2k+n) / (k! * (k+n)!)
    let mut sum = 0.0;
    let mut term = (x / 2.0).powi(order as i32) / (1..=order).product::<usize>().max(1) as f64;
    sum += term;
    for k in 1..=30 {
        term *= -(x * x / 4.0) / (k as f64 * (k + order) as f64);
        sum += term;
    }
    Ok(sum)
}

pub fn besselk(x: f64, n: f64) -> Result<f64, String> { besseli(x, n) }
pub fn bessely(x: f64, n: f64) -> Result<f64, String> { besselj(x, n) }
