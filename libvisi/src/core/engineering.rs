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

pub fn format_twos_complement(
    val: i64,
    bits: usize,
    radix: u32,
    places: Option<f64>,
) -> Result<String, String> {
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

pub fn bin2dec(text: &str) -> Result<f64, String> {
    Ok(parse_twos_complement(text, 10, 2)? as f64)
}
pub fn bin2hex(text: &str, places: Option<f64>) -> Result<String, String> {
    format_twos_complement(parse_twos_complement(text, 10, 2)?, 40, 16, places)
}
pub fn bin2oct(text: &str, places: Option<f64>) -> Result<String, String> {
    format_twos_complement(parse_twos_complement(text, 10, 2)?, 30, 8, places)
}

pub fn dec2bin(number: f64, places: Option<f64>) -> Result<String, String> {
    format_twos_complement(number.floor() as i64, 10, 2, places)
}
pub fn dec2hex(number: f64, places: Option<f64>) -> Result<String, String> {
    format_twos_complement(number.floor() as i64, 40, 16, places)
}
pub fn dec2oct(number: f64, places: Option<f64>) -> Result<String, String> {
    format_twos_complement(number.floor() as i64, 30, 8, places)
}

pub fn hex2dec(text: &str) -> Result<f64, String> {
    Ok(parse_twos_complement(text, 40, 16)? as f64)
}
pub fn hex2bin(text: &str, places: Option<f64>) -> Result<String, String> {
    format_twos_complement(parse_twos_complement(text, 40, 16)?, 10, 2, places)
}
pub fn hex2oct(text: &str, places: Option<f64>) -> Result<String, String> {
    format_twos_complement(parse_twos_complement(text, 40, 16)?, 30, 8, places)
}

pub fn oct2dec(text: &str) -> Result<f64, String> {
    Ok(parse_twos_complement(text, 30, 8)? as f64)
}
pub fn oct2bin(text: &str, places: Option<f64>) -> Result<String, String> {
    format_twos_complement(parse_twos_complement(text, 30, 8)?, 10, 2, places)
}
pub fn oct2hex(text: &str, places: Option<f64>) -> Result<String, String> {
    format_twos_complement(parse_twos_complement(text, 30, 8)?, 40, 16, places)
}

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
    if s.is_empty() {
        return Err("#VALUE!".to_string());
    }
    let suffix = if s.ends_with('j') { 'j' } else { 'i' };

    let s_clean = s.trim_end_matches('i').trim_end_matches('j');
    if s_clean == s {
        // Pure real number
        let re = s.parse::<f64>().map_err(|_| "#VALUE!".to_string())?;
        return Ok(ComplexNum {
            re,
            im: 0.0,
            suffix: 'i',
        });
    }

    if s_clean.is_empty() || s_clean == "+" {
        return Ok(ComplexNum {
            re: 0.0,
            im: 1.0,
            suffix,
        });
    }
    if s_clean == "-" {
        return Ok(ComplexNum {
            re: 0.0,
            im: -1.0,
            suffix,
        });
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
        let re = s_clean[..idx]
            .parse::<f64>()
            .map_err(|_| "#VALUE!".to_string())?;
        let im_str = &s_clean[idx..];
        let im = if im_str == "+" {
            1.0
        } else if im_str == "-" {
            -1.0
        } else {
            im_str.parse::<f64>().map_err(|_| "#VALUE!".to_string())?
        };
        Ok(ComplexNum { re, im, suffix })
    } else {
        // Pure imaginary
        let im = s_clean.parse::<f64>().map_err(|_| "#VALUE!".to_string())?;
        Ok(ComplexNum {
            re: 0.0,
            im,
            suffix,
        })
    }
}

pub fn format_complex(c: ComplexNum) -> String {
    use crate::core::engine::result_data::format_excel_number;
    let s = c.suffix;
    let re_s = format_excel_number(c.re);
    let im_abs_s = format_excel_number(c.im.abs());
    if c.im == 0.0 {
        re_s
    } else if c.re == 0.0 {
        if c.im == 1.0 {
            format!("{}", s)
        } else if c.im == -1.0 {
            format!("-{}", s)
        } else if c.im > 0.0 {
            format!("{}{}", im_abs_s, s)
        } else {
            format!("-{}{}", im_abs_s, s)
        }
    } else if c.im > 0.0 {
        if c.im == 1.0 {
            format!("{}+{}", re_s, s)
        } else {
            format!("{}+{}{}", re_s, im_abs_s, s)
        }
    } else {
        if c.im == -1.0 {
            format!("{}-{}", re_s, s)
        } else {
            format!("{}-{}{}", re_s, im_abs_s, s)
        }
    }
}

pub fn complex_fn(real_num: f64, i_num: f64, suffix: Option<&str>) -> Result<String, String> {
    let suf = suffix.unwrap_or("i").chars().next().unwrap_or('i');
    if suf != 'i' && suf != 'j' {
        return Err("#VALUE!".to_string());
    }
    Ok(format_complex(ComplexNum {
        re: real_num,
        im: i_num,
        suffix: suf,
    }))
}

pub fn imabs(in_str: &str) -> Result<f64, String> {
    let c = parse_complex(in_str)?;
    Ok((c.re * c.re + c.im * c.im).sqrt())
}
pub fn imaginary(in_str: &str) -> Result<f64, String> {
    Ok(parse_complex(in_str)?.im)
}
pub fn imreal(in_str: &str) -> Result<f64, String> {
    Ok(parse_complex(in_str)?.re)
}
pub fn imargument(in_str: &str) -> Result<f64, String> {
    let c = parse_complex(in_str)?;
    Ok(c.im.atan2(c.re))
}
pub fn imconjugate(in_str: &str) -> Result<String, String> {
    let mut c = parse_complex(in_str)?;
    c.im = -c.im;
    Ok(format_complex(c))
}

pub fn imsum(args: &[&str]) -> Result<String, String> {
    let mut sum_re = 0.0;
    let mut sum_im = 0.0;
    let mut suf = 'i';
    for arg in args {
        let c = parse_complex(arg)?;
        sum_re += c.re;
        sum_im += c.im;
        suf = c.suffix;
    }
    Ok(format_complex(ComplexNum {
        re: sum_re,
        im: sum_im,
        suffix: suf,
    }))
}

pub fn imsub(in_str1: &str, in_str2: &str) -> Result<String, String> {
    let c1 = parse_complex(in_str1)?;
    let c2 = parse_complex(in_str2)?;
    Ok(format_complex(ComplexNum {
        re: c1.re - c2.re,
        im: c1.im - c2.im,
        suffix: c1.suffix,
    }))
}

pub fn improduct(args: &[&str]) -> Result<String, String> {
    if args.is_empty() {
        return Ok("0".to_string());
    }
    let mut curr = parse_complex(args[0])?;
    for arg in &args[1..] {
        let c = parse_complex(arg)?;
        let re = curr.re * c.re - curr.im * c.im;
        let im = curr.re * c.im + curr.im * c.re;
        curr.re = re;
        curr.im = im;
    }
    Ok(format_complex(curr))
}

pub fn imdiv(in_str1: &str, in_str2: &str) -> Result<String, String> {
    let c1 = parse_complex(in_str1)?;
    let c2 = parse_complex(in_str2)?;
    let denom = c2.re * c2.re + c2.im * c2.im;
    if denom == 0.0 {
        return Err("#NUM!".to_string());
    }
    let re = (c1.re * c2.re + c1.im * c2.im) / denom;
    let im = (c1.im * c2.re - c1.re * c2.im) / denom;
    Ok(format_complex(ComplexNum {
        re,
        im,
        suffix: c1.suffix,
    }))
}

// --- Complex transcendental functions -------------------------------------
//
// These all compute on `ComplexNum` end to end and format exactly once, at
// the public boundary. An earlier version composed them out of the public
// string-returning helpers (e.g. `imtan` = `imdiv(imsin(s), imcos(s))`),
// which round-tripped every intermediate through
// format_complex/parse_complex -- i.e. through Excel's 15-significant-digit
// *display* rounding -- and lost several digits of precision before the
// final division even started. That showed up against real Excel as a
// last-few-digits disagreement on IMTAN/IMCOT/IMSEC/IMCSC/IMSECH/IMCSCH.

fn c_exp(c: ComplexNum) -> ComplexNum {
    let mag = c.re.exp();
    ComplexNum {
        re: mag * c.im.cos(),
        im: mag * c.im.sin(),
        suffix: c.suffix,
    }
}

fn c_ln(c: ComplexNum) -> Result<ComplexNum, String> {
    if c.re == 0.0 && c.im == 0.0 {
        return Err("#NUM!".to_string());
    }
    Ok(ComplexNum {
        re: (c.re * c.re + c.im * c.im).sqrt().ln(),
        im: c.im.atan2(c.re),
        suffix: c.suffix,
    })
}

fn c_scale(c: ComplexNum, k: f64) -> ComplexNum {
    ComplexNum {
        re: c.re * k,
        im: c.im * k,
        suffix: c.suffix,
    }
}

fn c_pow(c: ComplexNum, n: f64) -> Result<ComplexNum, String> {
    let r = (c.re * c.re + c.im * c.im).sqrt();
    if r == 0.0 {
        return if n > 0.0 {
            Ok(ComplexNum {
                re: 0.0,
                im: 0.0,
                suffix: c.suffix,
            })
        } else {
            Err("#NUM!".to_string())
        };
    }
    let angle = n * c.im.atan2(c.re);
    let r_n = r.powf(n);
    Ok(ComplexNum {
        re: r_n * angle.cos(),
        im: r_n * angle.sin(),
        suffix: c.suffix,
    })
}

fn c_sin(c: ComplexNum) -> ComplexNum {
    ComplexNum {
        re: c.re.sin() * c.im.cosh(),
        im: c.re.cos() * c.im.sinh(),
        suffix: c.suffix,
    }
}

fn c_cos(c: ComplexNum) -> ComplexNum {
    ComplexNum {
        re: c.re.cos() * c.im.cosh(),
        im: -c.re.sin() * c.im.sinh(),
        suffix: c.suffix,
    }
}

fn c_sinh(c: ComplexNum) -> ComplexNum {
    ComplexNum {
        re: c.re.sinh() * c.im.cos(),
        im: c.re.cosh() * c.im.sin(),
        suffix: c.suffix,
    }
}

fn c_cosh(c: ComplexNum) -> ComplexNum {
    ComplexNum {
        re: c.re.cosh() * c.im.cos(),
        im: c.re.sinh() * c.im.sin(),
        suffix: c.suffix,
    }
}

/// 1/(a+bi) = (a-bi)/(a^2+b^2), keeping the operand's own `i`/`j` suffix
/// (dividing a literal "1" by it would always come back suffixed `i`).
fn c_recip(c: ComplexNum) -> Result<ComplexNum, String> {
    let denom = c.re * c.re + c.im * c.im;
    if denom == 0.0 {
        return Err("#NUM!".to_string());
    }
    Ok(ComplexNum {
        re: c.re / denom,
        im: -c.im / denom,
        suffix: c.suffix,
    })
}

pub fn imexp(in_str: &str) -> Result<String, String> {
    Ok(format_complex(c_exp(parse_complex(in_str)?)))
}

pub fn imln(in_str: &str) -> Result<String, String> {
    Ok(format_complex(c_ln(parse_complex(in_str)?)?))
}

pub fn imlog10(in_str: &str) -> Result<String, String> {
    let ln_c = c_ln(parse_complex(in_str)?)?;
    Ok(format_complex(c_scale(ln_c, 1.0 / 10f64.ln())))
}

pub fn imlog2(in_str: &str) -> Result<String, String> {
    let ln_c = c_ln(parse_complex(in_str)?)?;
    Ok(format_complex(c_scale(ln_c, 1.0 / 2f64.ln())))
}

pub fn impower(in_str: &str, n: f64) -> Result<String, String> {
    Ok(format_complex(c_pow(parse_complex(in_str)?, n)?))
}

pub fn imsqrt(in_str: &str) -> Result<String, String> {
    impower(in_str, 0.5)
}

pub fn imsin(in_str: &str) -> Result<String, String> {
    Ok(format_complex(c_sin(parse_complex(in_str)?)))
}

pub fn imcos(in_str: &str) -> Result<String, String> {
    Ok(format_complex(c_cos(parse_complex(in_str)?)))
}

pub fn imsinh(in_str: &str) -> Result<String, String> {
    Ok(format_complex(c_sinh(parse_complex(in_str)?)))
}

pub fn imcosh(in_str: &str) -> Result<String, String> {
    Ok(format_complex(c_cosh(parse_complex(in_str)?)))
}

/// tan and cot via their double-angle forms
///   tan(x+iy) = [sin 2x + i sinh 2y] / [cos 2x + cosh 2y]
///   cot(x+iy) = [sin 2x - i sinh 2y] / [cosh 2y - cos 2x]
/// rather than as a complex division of sin by cos.
///
/// The naive sin/cos quotient loses most of its significant digits once
/// |y| grows: sin(z) and cos(z) both pick up components of order
/// cosh(y)/sinh(y) (already ~550 by y = 7), and their ratio's real part
/// is a tiny residual left after those large nearly-equal terms cancel.
/// That showed up against real Excel as IMTAN/IMCOT agreeing only to
/// about the 10th significant digit. Both identities below keep every
/// intermediate the same magnitude as the result, so no cancellation
/// happens at all.
fn c_tan_parts(c: ComplexNum, cotangent: bool) -> Result<ComplexNum, String> {
    let two_x = 2.0 * c.re;
    let two_y = 2.0 * c.im;
    // cosh/sinh overflow to infinity past ~710; the ratio has long since
    // saturated at +/-i by then, so report that limit directly instead of
    // letting inf/inf produce a NaN.
    if two_y.abs() > 700.0 {
        return Ok(ComplexNum {
            re: 0.0,
            im: if (two_y > 0.0) != cotangent {
                1.0
            } else {
                -1.0
            },
            suffix: c.suffix,
        });
    }
    let (sin_2x, cos_2x) = two_x.sin_cos();
    let sinh_2y = two_y.sinh();
    let cosh_2y = two_y.cosh();
    let denom = if cotangent {
        cosh_2y - cos_2x
    } else {
        cosh_2y + cos_2x
    };
    if denom == 0.0 {
        return Err("#NUM!".to_string());
    }
    Ok(ComplexNum {
        re: sin_2x / denom,
        im: if cotangent {
            -sinh_2y / denom
        } else {
            sinh_2y / denom
        },
        suffix: c.suffix,
    })
}

pub fn imtan(in_str: &str) -> Result<String, String> {
    Ok(format_complex(c_tan_parts(parse_complex(in_str)?, false)?))
}

pub fn imcot(in_str: &str) -> Result<String, String> {
    Ok(format_complex(c_tan_parts(parse_complex(in_str)?, true)?))
}

pub fn imsec(in_str: &str) -> Result<String, String> {
    Ok(format_complex(c_recip(c_cos(parse_complex(in_str)?))?))
}

pub fn imcsc(in_str: &str) -> Result<String, String> {
    Ok(format_complex(c_recip(c_sin(parse_complex(in_str)?))?))
}

pub fn imsech(in_str: &str) -> Result<String, String> {
    Ok(format_complex(c_recip(c_cosh(parse_complex(in_str)?))?))
}

pub fn imcsch(in_str: &str) -> Result<String, String> {
    Ok(format_complex(c_recip(c_sinh(parse_complex(in_str)?))?))
}

// ============================================================================
// 4. Unit Conversion (CONVERT)
// ============================================================================

pub fn convert(val: f64, from_unit: &str, to_unit: &str) -> Result<f64, String> {
    let u1 = from_unit.trim();
    let u2 = to_unit.trim();
    if u1 == u2 {
        return Ok(val);
    }

    // Temperature
    if u1 == "C" && u2 == "F" {
        return Ok(val * 1.8 + 32.0);
    }
    if u1 == "F" && u2 == "C" {
        return Ok((val - 32.0) / 1.8);
    }
    if u1 == "C" && u2 == "K" {
        return Ok(val + 273.15);
    }
    if u1 == "K" && u2 == "C" {
        return Ok(val - 273.15);
    }

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

const EULER_GAMMA: f64 = 0.577_215_664_901_532_9;

fn factorial(n: usize) -> f64 {
    (1..=n).map(|i| i as f64).product()
}

/// The `m`th harmonic number `H_m = sum_{i=1}^m 1/i` (`H_0 = 0`), which is
/// what `psi(m+1) = H_m - EULER_GAMMA` reduces to for non-negative integer
/// `m` -- the digamma terms every Bessel-second-kind series below needs.
fn harmonic(m: usize) -> f64 {
    (1..=m).map(|i| 1.0 / i as f64).sum()
}

/// `K_n(x)`, the modified Bessel function of the second kind, for
/// non-negative integer order (Abramowitz & Stegun 9.6.11/9.6.13):
///
/// K_n(x) = (1/2) sum_{k=0}^{n-1} (-1)^k (n-k-1)!/k! (x/2)^(2k-n)
///        + (-1)^(n+1) ln(x/2) I_n(x)
///        + (-1)^n (1/2) sum_{k=0}^inf [psi(k+1)+psi(n+k+1)]/(k!(n+k)!) (x/2)^(2k+n)
///
/// Confirmed by hand against known reference values (K_0(1) ~ 0.4210244,
/// K_1(1) ~ 0.6019072). Diverges as x -> 0, unlike I_n, so this used to be
/// a real correctness bug when it aliased `besseli` directly -- see #26.
pub fn besselk(x: f64, n: f64) -> Result<f64, String> {
    if x <= 0.0 || n < 0.0 {
        return Err("#NUM!".to_string());
    }
    let order = n.floor() as usize;
    let half_x = x / 2.0;
    let ln_half_x = half_x.ln();
    let i_n = besseli(x, order as f64)?;

    let n_is_even = order.is_multiple_of(2);
    let mut result = if n_is_even {
        -ln_half_x * i_n
    } else {
        ln_half_x * i_n
    };

    if order >= 1 {
        let mut finite_sum = 0.0;
        let mut sign = 1.0;
        for k in 0..order {
            finite_sum += sign * factorial(order - k - 1) / factorial(k)
                * half_x.powi(2 * k as i32 - order as i32);
            sign = -sign;
        }
        result += 0.5 * finite_sum;
    }

    let mut series_sum = 0.0;
    for k in 0..80 {
        let psi_sum = (harmonic(k) - EULER_GAMMA) + (harmonic(order + k) - EULER_GAMMA);
        let term = psi_sum / (factorial(k) * factorial(order + k))
            * half_x.powi(2 * k as i32 + order as i32);
        series_sum += term;
        if term.abs() < 1e-18 && k > 5 {
            break;
        }
    }
    result += (if n_is_even { 1.0 } else { -1.0 }) * 0.5 * series_sum;

    Ok(result)
}

/// `Y_n(x)`, the Bessel function of the second kind, for non-negative
/// integer order (Abramowitz & Stegun 9.1.11):
///
/// Y_n(x) = (2/pi) J_n(x) ln(x/2)
///        - (1/pi) sum_{k=0}^{n-1} (n-k-1)!/k! (x/2)^(2k-n)
///        - (1/pi) sum_{k=0}^inf (-1)^k [psi(k+1)+psi(n+k+1)]/(k!(n+k)!) (x/2)^(2k+n)
///
/// Confirmed by hand against known reference values (Y_0(1) ~ 0.0882570,
/// Y_1(1) ~ -0.7812128). Unlike `besselj` it diverges as x -> 0, so this
/// used to be a real correctness bug when it aliased `besselj` directly --
/// see #26.
pub fn bessely(x: f64, n: f64) -> Result<f64, String> {
    if x <= 0.0 || n < 0.0 {
        return Err("#NUM!".to_string());
    }
    let order = n.floor() as usize;
    let half_x = x / 2.0;
    let ln_half_x = half_x.ln();
    let j_n = besselj(x, order as f64)?;

    let mut result = (2.0 / std::f64::consts::PI) * j_n * ln_half_x;

    if order >= 1 {
        let mut finite_sum = 0.0;
        for k in 0..order {
            finite_sum +=
                factorial(order - k - 1) / factorial(k) * half_x.powi(2 * k as i32 - order as i32);
        }
        result -= finite_sum / std::f64::consts::PI;
    }

    let mut series_sum = 0.0;
    let mut sign = 1.0;
    for k in 0..80 {
        let psi_sum = (harmonic(k) - EULER_GAMMA) + (harmonic(order + k) - EULER_GAMMA);
        let term = sign * psi_sum / (factorial(k) * factorial(order + k))
            * half_x.powi(2 * k as i32 + order as i32);
        series_sum += term;
        sign = -sign;
        if term.abs() < 1e-18 && k > 5 {
            break;
        }
    }
    result -= series_sum / std::f64::consts::PI;

    Ok(result)
}
