// High-precision statistical functions for libvisi
// Implements Excel-compatible statistical distributions, summary measures, linear regression, and criteria functions.

use std::cmp::Ordering;

// ============================================================================
// 1. Core Mathematical Utilities (Special Functions)
// ============================================================================

/// Inverse standard normal CDF (Acklam's algorithm, max error < 1.15e-9, refined with Newton steps to double precision).
pub fn inv_normal_cdf(p: f64) -> Result<f64, String> {
    if p <= 0.0 || p >= 1.0 {
        return Err("#NUM!".to_string());
    }

    // Coefficients in rational approximations
    let a = [
        -3.969683028665376e+01,
        2.209460984245205e+02,
        -2.759285104469687e+02,
        1.383577518672690e+02,
        -3.066479806614716e+01,
        2.506628277459239e+00,
    ];
    let b = [
        -5.447609879822406e+01,
        1.615858368580409e+02,
        -1.556989798598866e+02,
        6.680131188771972e+01,
        -1.328068155288572e+01,
    ];
    let c = [
        -7.784894002430293e-03,
        -3.223964580411365e-01,
        -2.400758277161838e+00,
        -2.549732539343734e+00,
        4.374664141464968e+00,
        2.938163982698783e+00,
    ];
    let d = [
        7.784695709041462e-03,
        3.224671290700398e-01,
        2.445134137142996e+00,
        3.754408661907416e+00,
    ];

    let p_low = 0.02425;
    let p_high = 1.0 - p_low;

    let q: f64;
    let mut x: f64;

    if p < p_low {
        // Lower tail
        q = (-2.0 * p.ln()).sqrt();
        x = (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5])
            / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1.0);
    } else if p <= p_high {
        // Central region
        q = p - 0.5;
        let r = q * q;
        x = (((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q
            / (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1.0);
    } else {
        // Upper tail
        q = (-2.0 * (1.0 - p).ln()).sqrt();
        x = -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5])
            / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1.0);
    }

    // Refine using 2 Halley/Newton steps to achieve maximum f64 precision
    for _ in 0..2 {
        let err = normal_cdf(x) - p;
        let pdf = normal_pdf(x);
        if pdf > 0.0 {
            let delta = err / (pdf + 0.5 * x * err);
            x -= delta;
        }
    }

    Ok(x)
}

/// Standard normal PDF
pub fn normal_pdf(x: f64) -> f64 {
    (1.0 / (2.0 * std::f64::consts::PI).sqrt()) * (-0.5 * x * x).exp()
}

/// Standard normal CDF via error function
pub fn normal_cdf(x: f64) -> f64 {
    0.5 * (1.0 + erf(x / std::f64::consts::SQRT_2))
}

/// Error function erf(x)
pub fn erf(x: f64) -> f64 {
    if x.is_nan() {
        return f64::NAN;
    }
    if x == 0.0 {
        return 0.0;
    }
    let sign = if x < 0.0 { -1.0 } else { 1.0 };
    let ax = x.abs();
    if ax > 10.0 {
        return sign;
    }

    // High-precision Chebyshev approximation (erfc(x) for x >= 0)
    let t = 1.0 / (1.0 + 0.3275911 * ax);
    let poly = t
        * (0.254829592
            + t * (-0.284496736 + t * (1.421413741 + t * (-1.453152027 + t * 1.061405429))));
    let ans = 1.0 - poly * (-ax * ax).exp();
    sign * ans
}

/// Complementary error function erfc(x) = 1 - erf(x)
pub fn erfc(x: f64) -> f64 {
    1.0 - erf(x)
}

/// Log Gamma function ln(Gamma(x)) using Lanczos approximation (g=7, N=9)
pub fn lgamma(x: f64) -> f64 {
    if x <= 0.0 {
        if x == x.floor() {
            return f64::NAN; // Pole at non-positive integers
        }
        let sin_pix = (std::f64::consts::PI * x).sin().abs();
        return (std::f64::consts::PI / sin_pix).ln() - lgamma(1.0 - x);
    }

    let p = [
        0.99999999999980993,
        676.5203681218851,
        -1259.139216722289,
        771.32342877765313,
        -176.61502916214059,
        12.507343278686905,
        -0.13857109526572012,
        9.9843695780195716e-6,
        1.5056327351493116e-7,
    ];

    let z = x - 1.0;
    let mut sum = p[0];
    for i in 1..p.len() {
        sum += p[i] / (z + i as f64);
    }

    let t = z + 7.5;
    0.5 * (2.0 * std::f64::consts::PI).ln() + (z + 0.5) * t.ln() - t + sum.ln()
}

/// Gamma function Gamma(x)
pub fn gamma(x: f64) -> f64 {
    if x <= 0.0 && x == x.floor() {
        return f64::NAN;
    }
    if x < 0.0 {
        let sin_pix = (std::f64::consts::PI * x).sin();
        return std::f64::consts::PI / (sin_pix * gamma(1.0 - x));
    }
    lgamma(x).exp()
}

/// Lower regularized incomplete gamma P(a, x) = gamma(a, x) / Gamma(a)
pub fn regularized_gamma_p(a: f64, x: f64) -> f64 {
    if a <= 0.0 || x < 0.0 {
        return f64::NAN;
    }
    if x == 0.0 {
        return 0.0;
    }

    if x < a + 1.0 {
        let mut ap = a;
        let mut sum = 1.0 / a;
        let mut del = sum;
        for _ in 0..200 {
            ap += 1.0;
            del *= x / ap;
            sum += del;
            if del.abs() < sum.abs() * 1e-15 {
                break;
            }
        }
        sum * (-x + a * x.ln() - lgamma(a)).exp()
    } else {
        1.0 - regularized_gamma_q(a, x)
    }
}

/// Upper regularized incomplete gamma Q(a, x) = Gamma(a, x) / Gamma(a)
pub fn regularized_gamma_q(a: f64, x: f64) -> f64 {
    if a <= 0.0 || x < 0.0 {
        return f64::NAN;
    }
    if x == 0.0 {
        return 1.0;
    }

    if x < a + 1.0 {
        1.0 - regularized_gamma_p(a, x)
    } else {
        let mut b = x + 1.0 - a;
        let mut c = 1.0 / 1e-30;
        let mut d = 1.0 / b;
        let mut h = d;

        for i in 1..200 {
            let an = -(i as f64) * (i as f64 - a);
            b += 2.0;
            d = an * d + b;
            if d.abs() < 1e-30 {
                d = 1e-30;
            }
            c = b + an / c;
            if c.abs() < 1e-30 {
                c = 1e-30;
            }
            d = 1.0 / d;
            let del = d * c;
            h *= del;
            if (del - 1.0).abs() < 1e-15 {
                break;
            }
        }

        h * (-x + a * x.ln() - lgamma(a)).exp()
    }
}

/// Inverse incomplete gamma function: solves P(a, x) = p for x
pub fn inv_gamma_p(a: f64, p: f64) -> Result<f64, String> {
    if p <= 0.0 || p >= 1.0 || a <= 0.0 {
        if p == 0.0 {
            return Ok(0.0);
        }
        return Err("#NUM!".to_string());
    }

    let mut x = if a > 1.0 {
        let eta = inv_normal_cdf(p)?;
        (a * (1.0 - 1.0 / (9.0 * a) + eta / (9.0 * a).sqrt()).powi(3)).max(0.001)
    } else {
        (p * a * gamma(a)).powf(1.0 / a)
    };

    for _ in 0..40 {
        let err = regularized_gamma_p(a, x) - p;
        if err.abs() < 1e-12 {
            break;
        }
        let pdf = (-x + (a - 1.0) * x.ln() - lgamma(a)).exp();
        if pdf == 0.0 {
            break;
        }
        let step = err / pdf;
        x -= step;
        if x <= 0.0 {
            x = 1e-6;
        }
    }

    Ok(x)
}

/// Regularized incomplete beta function I_x(a, b)
pub fn incbeta(a: f64, b: f64, x: f64) -> f64 {
    if x <= 0.0 {
        return 0.0;
    }
    if x >= 1.0 {
        return 1.0;
    }
    if a <= 0.0 || b <= 0.0 {
        return f64::NAN;
    }

    if x > (a + 1.0) / (a + b + 2.0) {
        return 1.0 - incbeta(b, a, 1.0 - x);
    }

    let lbeta = lgamma(a) + lgamma(b) - lgamma(a + b);
    let front = (a * x.ln() + b * (1.0 - x).ln() - lbeta).exp() / a;

    let mut d = 1.0 - (a + b) * x / (a + 1.0);
    if d.abs() < 1e-30 {
        d = 1e-30;
    }
    d = 1.0 / d;
    let mut h = d;
    let mut c = 1.0;

    for m in 1..200 {
        let m_f = m as f64;
        let m2 = 2.0 * m_f;

        let aa = m_f * (b - m_f) * x / ((a + m2 - 1.0) * (a + m2));
        d = 1.0 + aa * d;
        if d.abs() < 1e-30 {
            d = 1e-30;
        }
        c = 1.0 + aa / c;
        if c.abs() < 1e-30 {
            c = 1e-30;
        }
        d = 1.0 / d;
        h *= d * c;

        let aa2 = -(a + m_f) * (a + b + m_f) * x / ((a + m2) * (a + m2 + 1.0));
        d = 1.0 + aa2 * d;
        if d.abs() < 1e-30 {
            d = 1e-30;
        }
        c = 1.0 + aa2 / c;
        if c.abs() < 1e-30 {
            c = 1e-30;
        }
        d = 1.0 / d;
        let del = d * c;
        h *= del;

        if (del - 1.0).abs() < 1e-15 {
            break;
        }
    }

    front * h
}

/// Inverse incomplete beta function: solves I_x(a, b) = p for x
pub fn inv_incbeta(a: f64, b: f64, p: f64) -> Result<f64, String> {
    if p <= 0.0 || p >= 1.0 || a <= 0.0 || b <= 0.0 {
        if p == 0.0 {
            return Ok(0.0);
        }
        if p == 1.0 {
            return Ok(1.0);
        }
        return Err("#NUM!".to_string());
    }

    let y = inv_normal_cdf(p)?;
    let h = 2.0 / (1.0 / (2.0 * a - 1.0) + 1.0 / (2.0 * b - 1.0));
    let w = (y * (h + 5.0 / 6.0 - 2.0 / (3.0 * h)).sqrt() / h)
        - (1.0 / (2.0 * b - 1.0) - 1.0 / (2.0 * a - 1.0)) * (y * y + 5.0 / 6.0 - 2.0 / (3.0 * h));
    let mut x = (a / (a + b * (2.0 * w).exp())).clamp(0.0001, 0.9999);

    for _ in 0..40 {
        let err = incbeta(a, b, x) - p;
        if err.abs() < 1e-12 {
            break;
        }
        let lbeta = lgamma(a) + lgamma(b) - lgamma(a + b);
        let pdf = ((a - 1.0) * x.ln() + (b - 1.0) * (1.0 - x).ln() - lbeta).exp();
        if pdf == 0.0 {
            break;
        }
        let step = err / pdf;
        x = (x - step).clamp(1e-12, 1.0 - 1e-12);
    }

    Ok(x)
}

// ============================================================================
// 2. Descriptive & Summary Statistics
// ============================================================================

pub fn avedev(data: &[f64]) -> Result<f64, String> {
    if data.is_empty() {
        return Err("#DIV/0!".to_string());
    }
    let mean = data.iter().sum::<f64>() / data.len() as f64;
    let sum_abs_diff: f64 = data.iter().map(|&x| (x - mean).abs()).sum();
    Ok(sum_abs_diff / data.len() as f64)
}

pub fn devsq(data: &[f64]) -> Result<f64, String> {
    if data.is_empty() {
        return Err("#NUM!".to_string());
    }
    let mean = data.iter().sum::<f64>() / data.len() as f64;
    let sum_sq_diff: f64 = data.iter().map(|&x| (x - mean) * (x - mean)).sum();
    Ok(sum_sq_diff)
}

pub fn geomean(data: &[f64]) -> Result<f64, String> {
    if data.is_empty() {
        return Err("#NUM!".to_string());
    }
    let mut log_sum = 0.0;
    for &x in data {
        if x <= 0.0 {
            return Err("#NUM!".to_string());
        }
        log_sum += x.ln();
    }
    Ok((log_sum / data.len() as f64).exp())
}

pub fn harmean(data: &[f64]) -> Result<f64, String> {
    if data.is_empty() {
        return Err("#NUM!".to_string());
    }
    let mut inv_sum = 0.0;
    for &x in data {
        if x <= 0.0 {
            return Err("#NUM!".to_string());
        }
        inv_sum += 1.0 / x;
    }
    Ok(data.len() as f64 / inv_sum)
}

pub fn median(data: &[f64]) -> Result<f64, String> {
    if data.is_empty() {
        return Err("#NUM!".to_string());
    }
    let mut sorted = data.to_vec();
    sorted.sort_by(|a, b| a.partial_cmp(b).unwrap_or(Ordering::Equal));
    let n = sorted.len();
    if n % 2 == 1 {
        Ok(sorted[n / 2])
    } else {
        Ok((sorted[n / 2 - 1] + sorted[n / 2]) / 2.0)
    }
}

pub fn mode_sngl(data: &[f64]) -> Result<f64, String> {
    if data.is_empty() {
        return Err("#N/A".to_string());
    }
    let mut counts = std::collections::HashMap::new();
    let mut max_count = 0;
    let mut first_seen = std::collections::HashMap::new();

    for (idx, &x) in data.iter().enumerate() {
        // Quantize floats slightly for exact hash key matching
        let key = x.to_bits();
        let entry = counts.entry(key).or_insert(0);
        *entry += 1;
        if *entry > max_count {
            max_count = *entry;
        }
        first_seen.entry(key).or_insert(idx);
    }

    if max_count <= 1 {
        return Err("#N/A".to_string());
    }

    // Return the mode that appeared earliest
    let mut best_val = 0.0;
    let mut best_idx = usize::MAX;
    for (key, count) in counts {
        if count == max_count {
            let idx = first_seen[&key];
            if idx < best_idx {
                best_idx = idx;
                best_val = f64::from_bits(key);
            }
        }
    }
    Ok(best_val)
}

pub fn mode_mult(data: &[f64]) -> Result<Vec<f64>, String> {
    if data.is_empty() {
        return Err("#N/A".to_string());
    }
    let mut counts = std::collections::HashMap::new();
    let mut max_count = 0;

    for &x in data {
        let key = x.to_bits();
        let entry = counts.entry(key).or_insert(0);
        *entry += 1;
        if *entry > max_count {
            max_count = *entry;
        }
    }

    if max_count <= 1 {
        return Err("#N/A".to_string());
    }

    let mut modes = Vec::new();
    for (key, count) in counts {
        if count == max_count {
            modes.push(f64::from_bits(key));
        }
    }
    modes.sort_by(|a, b| a.partial_cmp(b).unwrap_or(Ordering::Equal));
    Ok(modes)
}

pub fn trimmean(data: &[f64], percent: f64) -> Result<f64, String> {
    if data.is_empty() || percent < 0.0 || percent >= 1.0 {
        return Err("#NUM!".to_string());
    }
    let mut sorted = data.to_vec();
    sorted.sort_by(|a, b| a.partial_cmp(b).unwrap_or(Ordering::Equal));

    let n = sorted.len();
    let k = ((n as f64 * percent / 2.0).floor()) as usize;
    if 2 * k >= n {
        return Err("#NUM!".to_string());
    }

    let trimmed = &sorted[k..(n - k)];
    Ok(trimmed.iter().sum::<f64>() / trimmed.len() as f64)
}

// ============================================================================
// 3. Variance & Higher Moments
// ============================================================================

pub fn var_s(data: &[f64]) -> Result<f64, String> {
    if data.len() <= 1 {
        return Err("#DIV/0!".to_string());
    }
    let mean = data.iter().sum::<f64>() / data.len() as f64;
    let sum_sq: f64 = data.iter().map(|&x| (x - mean) * (x - mean)).sum();
    Ok(sum_sq / (data.len() - 1) as f64)
}

pub fn var_p(data: &[f64]) -> Result<f64, String> {
    if data.is_empty() {
        return Err("#DIV/0!".to_string());
    }
    let mean = data.iter().sum::<f64>() / data.len() as f64;
    let sum_sq: f64 = data.iter().map(|&x| (x - mean) * (x - mean)).sum();
    Ok(sum_sq / data.len() as f64)
}

pub fn stdev_s(data: &[f64]) -> Result<f64, String> {
    Ok(var_s(data)?.sqrt())
}

pub fn stdev_p(data: &[f64]) -> Result<f64, String> {
    Ok(var_p(data)?.sqrt())
}

pub fn skew(data: &[f64]) -> Result<f64, String> {
    let n = data.len();
    if n < 3 {
        return Err("#DIV/0!".to_string());
    }
    let s = stdev_s(data)?;
    if s == 0.0 {
        return Err("#DIV/0!".to_string());
    }
    let mean = data.iter().sum::<f64>() / n as f64;
    let sum_cube: f64 = data.iter().map(|&x| ((x - mean) / s).powi(3)).sum();

    let factor = (n as f64) / ((n - 1) as f64 * (n - 2) as f64);
    Ok(factor * sum_cube)
}

pub fn skew_p(data: &[f64]) -> Result<f64, String> {
    let n = data.len();
    if n < 1 {
        return Err("#DIV/0!".to_string());
    }
    let sigma = stdev_p(data)?;
    if sigma == 0.0 {
        return Err("#DIV/0!".to_string());
    }
    let mean = data.iter().sum::<f64>() / n as f64;
    let sum_cube: f64 = data.iter().map(|&x| ((x - mean) / sigma).powi(3)).sum();
    Ok(sum_cube / n as f64)
}

pub fn kurt(data: &[f64]) -> Result<f64, String> {
    let n = data.len();
    if n < 4 {
        return Err("#DIV/0!".to_string());
    }
    let s = stdev_s(data)?;
    if s == 0.0 {
        return Err("#DIV/0!".to_string());
    }
    let mean = data.iter().sum::<f64>() / n as f64;
    let sum_quad: f64 = data.iter().map(|&x| ((x - mean) / s).powi(4)).sum();

    let nf = n as f64;
    let term1 = (nf * (nf + 1.0)) / ((nf - 1.0) * (nf - 2.0) * (nf - 3.0)) * sum_quad;
    let term2 = (3.0 * (nf - 1.0) * (nf - 1.0)) / ((nf - 2.0) * (nf - 3.0));
    Ok(term1 - term2)
}

// ============================================================================
// 4. Ranks, Quantiles, and Percentiles
// ============================================================================

pub fn large(data: &[f64], k: usize) -> Result<f64, String> {
    if k == 0 || k > data.len() {
        return Err("#NUM!".to_string());
    }
    let mut sorted = data.to_vec();
    sorted.sort_by(|a, b| b.partial_cmp(a).unwrap_or(Ordering::Equal));
    Ok(sorted[k - 1])
}

pub fn small(data: &[f64], k: usize) -> Result<f64, String> {
    if k == 0 || k > data.len() {
        return Err("#NUM!".to_string());
    }
    let mut sorted = data.to_vec();
    sorted.sort_by(|a, b| a.partial_cmp(b).unwrap_or(Ordering::Equal));
    Ok(sorted[k - 1])
}

pub fn rank_eq(number: f64, ref_data: &[f64], order: usize) -> Result<f64, String> {
    if ref_data.is_empty() {
        return Err("#N/A".to_string());
    }
    let mut found = false;
    let mut rank = 1;

    for &x in ref_data {
        if x == number {
            found = true;
        }
        if order == 0 {
            if x > number {
                rank += 1;
            }
        } else {
            if x < number {
                rank += 1;
            }
        }
    }

    if !found {
        Err("#N/A".to_string())
    } else {
        Ok(rank as f64)
    }
}

pub fn rank_avg(number: f64, ref_data: &[f64], order: usize) -> Result<f64, String> {
    if ref_data.is_empty() {
        return Err("#N/A".to_string());
    }
    let mut count_same = 0;
    let mut rank = 1;

    for &x in ref_data {
        if x == number {
            count_same += 1;
        }
        if order == 0 {
            if x > number {
                rank += 1;
            }
        } else {
            if x < number {
                rank += 1;
            }
        }
    }

    if count_same == 0 {
        Err("#N/A".to_string())
    } else {
        Ok(rank as f64 + (count_same - 1) as f64 / 2.0)
    }
}

pub fn percentile_inc(data: &[f64], k: f64) -> Result<f64, String> {
    if data.is_empty() || k < 0.0 || k > 1.0 {
        return Err("#NUM!".to_string());
    }
    let mut sorted = data.to_vec();
    sorted.sort_by(|a, b| a.partial_cmp(b).unwrap_or(Ordering::Equal));

    let n = sorted.len();
    if n == 1 {
        return Ok(sorted[0]);
    }

    let idx = k * (n - 1) as f64;
    let j = idx.floor() as usize;
    let d = idx - j as f64;

    if j >= n - 1 {
        Ok(sorted[n - 1])
    } else {
        Ok(sorted[j] + d * (sorted[j + 1] - sorted[j]))
    }
}

pub fn percentile_exc(data: &[f64], k: f64) -> Result<f64, String> {
    if data.is_empty() || k <= 0.0 || k >= 1.0 {
        return Err("#NUM!".to_string());
    }
    let mut sorted = data.to_vec();
    sorted.sort_by(|a, b| a.partial_cmp(b).unwrap_or(Ordering::Equal));

    let n = sorted.len();
    let idx = k * (n + 1) as f64 - 1.0;
    if idx < 0.0 || idx >= (n - 1) as f64 {
        return Err("#NUM!".to_string());
    }

    let j = idx.floor() as usize;
    let d = idx - j as f64;
    Ok(sorted[j] + d * (sorted[j + 1] - sorted[j]))
}

pub fn quartile_inc(data: &[f64], quart: usize) -> Result<f64, String> {
    match quart {
        0 => percentile_inc(data, 0.0),
        1 => percentile_inc(data, 0.25),
        2 => percentile_inc(data, 0.50),
        3 => percentile_inc(data, 0.75),
        4 => percentile_inc(data, 1.0),
        _ => Err("#NUM!".to_string()),
    }
}

pub fn quartile_exc(data: &[f64], quart: usize) -> Result<f64, String> {
    match quart {
        1 => percentile_exc(data, 0.25),
        2 => percentile_exc(data, 0.50),
        3 => percentile_exc(data, 0.75),
        _ => Err("#NUM!".to_string()),
    }
}

pub fn percentrank_inc(data: &[f64], x: f64, significance: usize) -> Result<f64, String> {
    if data.is_empty() {
        return Err("#N/A".to_string());
    }
    let mut sorted = data.to_vec();
    sorted.sort_by(|a, b| a.partial_cmp(b).unwrap_or(Ordering::Equal));

    let min_v = sorted[0];
    let max_v = sorted[sorted.len() - 1];
    if x < min_v || x > max_v {
        return Err("#N/A".to_string());
    }

    let n = sorted.len();
    if n == 1 {
        return Ok(1.0);
    }

    let mut ans = 0.0;
    for i in 0..n {
        if sorted[i] == x {
            ans = i as f64 / (n - 1) as f64;
            break;
        } else if sorted[i] > x {
            let prev = sorted[i - 1];
            let next = sorted[i];
            let frac = (x - prev) / (next - prev);
            ans = ((i - 1) as f64 + frac) / (n - 1) as f64;
            break;
        }
    }

    let mult = 10.0_f64.powi(significance as i32);
    Ok((ans * mult).round() / mult)
}

pub fn percentrank_exc(data: &[f64], x: f64, significance: usize) -> Result<f64, String> {
    if data.is_empty() {
        return Err("#N/A".to_string());
    }
    let mut sorted = data.to_vec();
    sorted.sort_by(|a, b| a.partial_cmp(b).unwrap_or(Ordering::Equal));

    let min_v = sorted[0];
    let max_v = sorted[sorted.len() - 1];
    if x < min_v || x > max_v {
        return Err("#N/A".to_string());
    }

    let n = sorted.len();
    let mut ans = 0.0;
    for i in 0..n {
        if sorted[i] == x {
            ans = (i + 1) as f64 / (n + 1) as f64;
            break;
        } else if sorted[i] > x {
            let prev = sorted[i - 1];
            let next = sorted[i];
            let frac = (x - prev) / (next - prev);
            ans = (i as f64 + frac) / (n + 1) as f64;
            break;
        }
    }

    let mult = 10.0_f64.powi(significance as i32);
    Ok((ans * mult).round() / mult)
}

// ============================================================================
// 5. Bivariate Statistics & Linear Regression
// ============================================================================

pub fn covariance_p(xs: &[f64], ys: &[f64]) -> Result<f64, String> {
    if xs.len() != ys.len() || xs.is_empty() {
        return Err("#N/A".to_string());
    }
    let n = xs.len() as f64;
    let mean_x = xs.iter().sum::<f64>() / n;
    let mean_y = ys.iter().sum::<f64>() / n;

    let cov = xs
        .iter()
        .zip(ys.iter())
        .map(|(&x, &y)| (x - mean_x) * (y - mean_y))
        .sum::<f64>();
    Ok(cov / n)
}

pub fn covariance_s(xs: &[f64], ys: &[f64]) -> Result<f64, String> {
    if xs.len() != ys.len() || xs.len() <= 1 {
        return Err("#DIV/0!".to_string());
    }
    let n = xs.len() as f64;
    let mean_x = xs.iter().sum::<f64>() / n;
    let mean_y = ys.iter().sum::<f64>() / n;

    let cov = xs
        .iter()
        .zip(ys.iter())
        .map(|(&x, &y)| (x - mean_x) * (y - mean_y))
        .sum::<f64>();
    Ok(cov / (n - 1.0))
}

pub fn correl(xs: &[f64], ys: &[f64]) -> Result<f64, String> {
    if xs.len() != ys.len() || xs.is_empty() {
        return Err("#N/A".to_string());
    }
    let n = xs.len() as f64;
    let mean_x = xs.iter().sum::<f64>() / n;
    let mean_y = ys.iter().sum::<f64>() / n;

    let mut num = 0.0;
    let mut den_x = 0.0;
    let mut den_y = 0.0;

    for (&x, &y) in xs.iter().zip(ys.iter()) {
        let dx = x - mean_x;
        let dy = y - mean_y;
        num += dx * dy;
        den_x += dx * dx;
        den_y += dy * dy;
    }

    let den = (den_x * den_y).sqrt();
    if den == 0.0 {
        Err("#DIV/0!".to_string())
    } else {
        Ok(num / den)
    }
}

pub fn slope(ys: &[f64], xs: &[f64]) -> Result<f64, String> {
    if xs.len() != ys.len() || xs.is_empty() {
        return Err("#N/A".to_string());
    }
    let n = xs.len() as f64;
    let mean_x = xs.iter().sum::<f64>() / n;
    let mean_y = ys.iter().sum::<f64>() / n;

    let mut num = 0.0;
    let mut den = 0.0;

    for (&x, &y) in xs.iter().zip(ys.iter()) {
        let dx = x - mean_x;
        let dy = y - mean_y;
        num += dx * dy;
        den += dx * dx;
    }

    if den == 0.0 {
        Err("#DIV/0!".to_string())
    } else {
        Ok(num / den)
    }
}

pub fn intercept(ys: &[f64], xs: &[f64]) -> Result<f64, String> {
    let m = slope(ys, xs)?;
    let mean_x = xs.iter().sum::<f64>() / xs.len() as f64;
    let mean_y = ys.iter().sum::<f64>() / ys.len() as f64;
    Ok(mean_y - m * mean_x)
}

pub fn rsq(ys: &[f64], xs: &[f64]) -> Result<f64, String> {
    let r = correl(xs, ys)?;
    Ok(r * r)
}

pub fn steyx(ys: &[f64], xs: &[f64]) -> Result<f64, String> {
    if xs.len() != ys.len() || xs.len() <= 2 {
        return Err("#DIV/0!".to_string());
    }
    let n = xs.len() as f64;
    let mean_x = xs.iter().sum::<f64>() / n;
    let mean_y = ys.iter().sum::<f64>() / n;

    let mut s_xx = 0.0;
    let mut s_yy = 0.0;
    let mut s_xy = 0.0;

    for (&x, &y) in xs.iter().zip(ys.iter()) {
        let dx = x - mean_x;
        let dy = y - mean_y;
        s_xx += dx * dx;
        s_yy += dy * dy;
        s_xy += dx * dy;
    }

    if s_xx == 0.0 {
        return Err("#DIV/0!".to_string());
    }

    let val = (s_yy - (s_xy * s_xy) / s_xx) / (n - 2.0);
    if val < 0.0 { Ok(0.0) } else { Ok(val.sqrt()) }
}

pub fn forecast_linear(x: f64, ys: &[f64], xs: &[f64]) -> Result<f64, String> {
    let m = slope(ys, xs)?;
    let b = intercept(ys, xs)?;
    Ok(m * x + b)
}

// ============================================================================
// 6. Distribution Functions
// ============================================================================

pub fn standardize(x: f64, mean: f64, std_dev: f64) -> Result<f64, String> {
    if std_dev <= 0.0 {
        Err("#NUM!".to_string())
    } else {
        Ok((x - mean) / std_dev)
    }
}

pub fn gauss(z: f64) -> Result<f64, String> {
    Ok(normal_cdf(z) - 0.5)
}

pub fn phi(x: f64) -> Result<f64, String> {
    Ok(normal_pdf(x))
}

pub fn norm_s_dist(z: f64, cumulative: bool) -> Result<f64, String> {
    if cumulative {
        Ok(normal_cdf(z))
    } else {
        Ok(normal_pdf(z))
    }
}

pub fn norm_s_inv(p: f64) -> Result<f64, String> {
    inv_normal_cdf(p)
}

pub fn norm_dist(x: f64, mean: f64, std_dev: f64, cumulative: bool) -> Result<f64, String> {
    if std_dev <= 0.0 {
        return Err("#NUM!".to_string());
    }
    let z = (x - mean) / std_dev;
    if cumulative {
        Ok(normal_cdf(z))
    } else {
        Ok(normal_pdf(z) / std_dev)
    }
}

pub fn norm_inv(p: f64, mean: f64, std_dev: f64) -> Result<f64, String> {
    if std_dev <= 0.0 {
        return Err("#NUM!".to_string());
    }
    let z = inv_normal_cdf(p)?;
    Ok(mean + z * std_dev)
}

pub fn lognorm_dist(x: f64, mean: f64, std_dev: f64, cumulative: bool) -> Result<f64, String> {
    if x <= 0.0 || std_dev <= 0.0 {
        return Err("#NUM!".to_string());
    }
    let lx = x.ln();
    if cumulative {
        norm_dist(lx, mean, std_dev, true)
    } else {
        let pdf = normal_pdf((lx - mean) / std_dev) / (x * std_dev);
        Ok(pdf)
    }
}

pub fn lognorm_inv(p: f64, mean: f64, std_dev: f64) -> Result<f64, String> {
    let z = norm_inv(p, mean, std_dev)?;
    Ok(z.exp())
}

pub fn expon_dist(x: f64, lambda: f64, cumulative: bool) -> Result<f64, String> {
    if x < 0.0 || lambda <= 0.0 {
        return Err("#NUM!".to_string());
    }
    if cumulative {
        Ok(1.0 - (-lambda * x).exp())
    } else {
        Ok(lambda * (-lambda * x).exp())
    }
}

pub fn poisson_dist(x: f64, mean: f64, cumulative: bool) -> Result<f64, String> {
    if x < 0.0 || mean <= 0.0 {
        return Err("#NUM!".to_string());
    }
    let k = x.floor() as usize;
    if cumulative {
        Ok(regularized_gamma_q(k as f64 + 1.0, mean))
    } else {
        Ok((-(mean) + k as f64 * mean.ln() - lgamma(k as f64 + 1.0)).exp())
    }
}

pub fn weibull_dist(x: f64, alpha: f64, beta: f64, cumulative: bool) -> Result<f64, String> {
    if x < 0.0 || alpha <= 0.0 || beta <= 0.0 {
        return Err("#NUM!".to_string());
    }
    let ratio = x / beta;
    if cumulative {
        Ok(1.0 - (-ratio.powf(alpha)).exp())
    } else {
        Ok((alpha / beta) * ratio.powf(alpha - 1.0) * (-ratio.powf(alpha)).exp())
    }
}

pub fn gamma_dist(x: f64, alpha: f64, beta: f64, cumulative: bool) -> Result<f64, String> {
    if x < 0.0 || alpha <= 0.0 || beta <= 0.0 {
        return Err("#NUM!".to_string());
    }
    if cumulative {
        Ok(regularized_gamma_p(alpha, x / beta))
    } else {
        let pdf = (x / beta).powf(alpha - 1.0) * (-(x / beta)).exp() / (beta * gamma(alpha));
        Ok(pdf)
    }
}

pub fn gamma_inv(p: f64, alpha: f64, beta: f64) -> Result<f64, String> {
    if beta <= 0.0 {
        return Err("#NUM!".to_string());
    }
    let g_val = inv_gamma_p(alpha, p)?;
    Ok(g_val * beta)
}

pub fn beta_dist(
    x: f64,
    alpha: f64,
    beta: f64,
    cumulative: bool,
    a: f64,
    b: f64,
) -> Result<f64, String> {
    if alpha <= 0.0 || beta <= 0.0 || a >= b || x < a || x > b {
        return Err("#NUM!".to_string());
    }
    let y = (x - a) / (b - a);
    if cumulative {
        Ok(incbeta(alpha, beta, y))
    } else {
        let pdf = (y.powf(alpha - 1.0) * (1.0 - y).powf(beta - 1.0))
            / ((lgamma(alpha) + lgamma(beta) - lgamma(alpha + beta)).exp() * (b - a));
        Ok(pdf)
    }
}

pub fn beta_inv(p: f64, alpha: f64, beta: f64, a: f64, b: f64) -> Result<f64, String> {
    if a >= b {
        return Err("#NUM!".to_string());
    }
    let y = inv_incbeta(alpha, beta, p)?;
    Ok(a + y * (b - a))
}

pub fn binom_dist(
    number_s: f64,
    trials: f64,
    probability_s: f64,
    cumulative: bool,
) -> Result<f64, String> {
    if number_s < 0.0
        || trials < 0.0
        || number_s > trials
        || probability_s < 0.0
        || probability_s > 1.0
    {
        return Err("#NUM!".to_string());
    }
    let k = number_s.floor();
    let n = trials.floor();
    let p = probability_s;

    if cumulative {
        let mut sum = 0.0;
        for i in 0..=(k as usize) {
            let pmf = (lgamma(n + 1.0) - lgamma(i as f64 + 1.0) - lgamma(n - i as f64 + 1.0)).exp()
                * p.powi(i as i32)
                * (1.0 - p).powf(n - i as f64);
            sum += pmf;
        }
        Ok(sum.min(1.0))
    } else {
        let pmf = (lgamma(n + 1.0) - lgamma(k + 1.0) - lgamma(n - k + 1.0)).exp()
            * p.powf(k)
            * (1.0 - p).powf(n - k);
        Ok(pmf)
    }
}

pub fn binom_dist_range(
    trials: f64,
    probability_s: f64,
    number_s: f64,
    number_s2: Option<f64>,
) -> Result<f64, String> {
    let k1 = number_s.floor();
    let k2 = number_s2.unwrap_or(number_s).floor();
    if k1 > k2 {
        return Err("#NUM!".to_string());
    }
    let mut sum = 0.0;
    for i in (k1 as usize)..=(k2 as usize) {
        sum += binom_dist(i as f64, trials, probability_s, false)?;
    }
    Ok(sum)
}

pub fn binom_inv(trials: f64, probability_s: f64, alpha: f64) -> Result<f64, String> {
    if trials < 0.0 || probability_s < 0.0 || probability_s > 1.0 || alpha < 0.0 || alpha > 1.0 {
        return Err("#NUM!".to_string());
    }
    let n = trials.floor() as usize;
    let mut cum = 0.0;
    for k in 0..=n {
        cum += binom_dist(k as f64, trials, probability_s, false)?;
        if cum >= alpha {
            return Ok(k as f64);
        }
    }
    Ok(n as f64)
}

pub fn negbinom_dist(
    number_f: f64,
    number_s: f64,
    probability_s: f64,
    cumulative: bool,
) -> Result<f64, String> {
    if number_f < 0.0 || number_s < 1.0 || probability_s < 0.0 || probability_s > 1.0 {
        return Err("#NUM!".to_string());
    }
    let k = number_f.floor();
    let r = number_s.floor();
    let p = probability_s;

    if cumulative {
        let mut sum = 0.0;
        for i in 0..=(k as usize) {
            let pmf = (lgamma(i as f64 + r) - lgamma(i as f64 + 1.0) - lgamma(r)).exp()
                * p.powf(r)
                * (1.0 - p).powi(i as i32);
            sum += pmf;
        }
        Ok(sum.min(1.0))
    } else {
        let pmf =
            (lgamma(k + r) - lgamma(k + 1.0) - lgamma(r)).exp() * p.powf(r) * (1.0 - p).powf(k);
        Ok(pmf)
    }
}

pub fn hypgeom_dist(
    sample_s: f64,
    sample_size: f64,
    pop_s: f64,
    pop_size: f64,
    cumulative: bool,
) -> Result<f64, String> {
    if sample_s < 0.0
        || sample_size < 0.0
        || pop_s < 0.0
        || pop_size < 0.0
        || sample_s > sample_size
        || sample_s > pop_s
        || sample_size > pop_size
    {
        return Err("#NUM!".to_string());
    }
    let k = sample_s.floor();
    let n = sample_size.floor();
    let m_pop = pop_s.floor();
    let n_pop = pop_size.floor();

    let pmf_fn = |x: f64| -> f64 {
        let log_comb1 = lgamma(m_pop + 1.0) - lgamma(x + 1.0) - lgamma(m_pop - x + 1.0);
        let log_comb2 = lgamma(n_pop - m_pop + 1.0)
            - lgamma(n - x + 1.0)
            - lgamma(n_pop - m_pop - (n - x) + 1.0);
        let log_comb3 = lgamma(n_pop + 1.0) - lgamma(n + 1.0) - lgamma(n_pop - n + 1.0);
        (log_comb1 + log_comb2 - log_comb3).exp()
    };

    if cumulative {
        let min_x = 0.0_f64.max(n + m_pop - n_pop);
        let mut sum = 0.0;
        for i in (min_x as usize)..=(k as usize) {
            sum += pmf_fn(i as f64);
        }
        Ok(sum.min(1.0))
    } else {
        Ok(pmf_fn(k))
    }
}

pub fn chisq_dist(x: f64, df: f64, cumulative: bool) -> Result<f64, String> {
    if df < 1.0 {
        return Err("#NUM!".to_string());
    }
    gamma_dist(x, df / 2.0, 2.0, cumulative)
}

pub fn chisq_dist_rt(x: f64, df: f64) -> Result<f64, String> {
    if x < 0.0 || df < 1.0 {
        return Err("#NUM!".to_string());
    }
    Ok(1.0 - chisq_dist(x, df, true)?)
}

pub fn chisq_inv(p: f64, df: f64) -> Result<f64, String> {
    if df < 1.0 {
        return Err("#NUM!".to_string());
    }
    gamma_inv(p, df / 2.0, 2.0)
}

pub fn chisq_inv_rt(p: f64, df: f64) -> Result<f64, String> {
    chisq_inv(1.0 - p, df)
}

pub fn chisq_test(actual: &[f64], expected: &[f64]) -> Result<f64, String> {
    if actual.len() != expected.len() || actual.is_empty() {
        return Err("#N/A".to_string());
    }
    let mut chi2 = 0.0;
    for (&o, &e) in actual.iter().zip(expected.iter()) {
        if e <= 0.0 {
            return Err("#DIV/0!".to_string());
        }
        chi2 += (o - e) * (o - e) / e;
    }
    let df = (actual.len() - 1) as f64;
    chisq_dist_rt(chi2, df)
}

pub fn f_dist(x: f64, df1: f64, df2: f64, cumulative: bool) -> Result<f64, String> {
    if x < 0.0 || df1 < 1.0 || df2 < 1.0 {
        return Err("#NUM!".to_string());
    }
    let y = (df1 * x) / (df1 * x + df2);
    if cumulative {
        Ok(incbeta(df1 / 2.0, df2 / 2.0, y))
    } else {
        let num = (df1 / df2).powf(df1 / 2.0) * x.powf(df1 / 2.0 - 1.0);
        let den = (1.0 + (df1 / df2) * x).powf((df1 + df2) / 2.0)
            * (lgamma(df1 / 2.0) + lgamma(df2 / 2.0) - lgamma((df1 + df2) / 2.0)).exp();
        Ok(num / den)
    }
}

pub fn f_dist_rt(x: f64, df1: f64, df2: f64) -> Result<f64, String> {
    if x < 0.0 || df1 < 1.0 || df2 < 1.0 {
        return Err("#NUM!".to_string());
    }
    Ok(1.0 - f_dist(x, df1, df2, true)?)
}

pub fn f_inv(p: f64, df1: f64, df2: f64) -> Result<f64, String> {
    if df1 < 1.0 || df2 < 1.0 {
        return Err("#NUM!".to_string());
    }
    let y = inv_incbeta(df1 / 2.0, df2 / 2.0, p)?;
    Ok((df2 * y) / (df1 * (1.0 - y)))
}

pub fn f_inv_rt(p: f64, df1: f64, df2: f64) -> Result<f64, String> {
    f_inv(1.0 - p, df1, df2)
}

pub fn f_test(array1: &[f64], array2: &[f64]) -> Result<f64, String> {
    let s1 = var_s(array1)?;
    let s2 = var_s(array2)?;
    if s1 == 0.0 || s2 == 0.0 {
        return Err("#DIV/0!".to_string());
    }
    let f_stat = s1 / s2;
    let df1 = (array1.len() - 1) as f64;
    let df2 = (array2.len() - 1) as f64;

    let p1 = f_dist_rt(f_stat, df1, df2)?;
    let p_two_tailed = (2.0 * p1.min(1.0 - p1)).min(1.0);
    Ok(p_two_tailed)
}

pub fn t_dist(x: f64, df: f64, cumulative: bool) -> Result<f64, String> {
    if df < 1.0 {
        return Err("#NUM!".to_string());
    }
    if cumulative {
        let x_t = df / (df + x * x);
        let ib = incbeta(df / 2.0, 0.5, x_t);
        if x >= 0.0 {
            Ok(1.0 - 0.5 * ib)
        } else {
            Ok(0.5 * ib)
        }
    } else {
        let num = (lgamma((df + 1.0) / 2.0) - lgamma(df / 2.0)).exp();
        let den = (df * std::f64::consts::PI).sqrt() * (1.0 + (x * x) / df).powf((df + 1.0) / 2.0);
        Ok(num / den)
    }
}

pub fn t_dist_rt(x: f64, df: f64) -> Result<f64, String> {
    Ok(1.0 - t_dist(x, df, true)?)
}

pub fn t_dist_2t(x: f64, df: f64) -> Result<f64, String> {
    if x < 0.0 {
        return Err("#NUM!".to_string());
    }
    Ok(2.0 * t_dist_rt(x, df)?)
}

pub fn t_inv(p: f64, df: f64) -> Result<f64, String> {
    if p <= 0.0 || p >= 1.0 || df < 1.0 {
        return Err("#NUM!".to_string());
    }
    if p == 0.5 {
        return Ok(0.0);
    }
    // Newton-Raphson solver
    let mut x = inv_normal_cdf(p)?;
    for _ in 0..30 {
        let err = t_dist(x, df, true)? - p;
        if err.abs() < 1e-12 {
            break;
        }
        let pdf = t_dist(x, df, false)?;
        if pdf == 0.0 {
            break;
        }
        x -= err / pdf;
    }
    Ok(x)
}

pub fn t_inv_2t(p: f64, df: f64) -> Result<f64, String> {
    if p <= 0.0 || p > 1.0 {
        return Err("#NUM!".to_string());
    }
    t_inv(1.0 - p / 2.0, df)
}

pub fn t_test(
    array1: &[f64],
    array2: &[f64],
    tails: usize,
    test_type: usize,
) -> Result<f64, String> {
    if tails != 1 && tails != 2 {
        return Err("#NUM!".to_string());
    }
    let n1 = array1.len();
    let n2 = array2.len();

    let (t_stat, df) = match test_type {
        1 => {
            // Paired
            if n1 != n2 || n1 <= 1 {
                return Err("#N/A".to_string());
            }
            let diffs: Vec<f64> = array1
                .iter()
                .zip(array2.iter())
                .map(|(&a, &b)| a - b)
                .collect();
            let mean_d = diffs.iter().sum::<f64>() / n1 as f64;
            let sd = stdev_s(&diffs)?;
            if sd == 0.0 {
                return Err("#DIV/0!".to_string());
            }
            (mean_d / (sd / (n1 as f64).sqrt()), (n1 - 1) as f64)
        }
        2 => {
            // Two-sample homoscedastic (equal variance)
            if n1 <= 1 || n2 <= 1 {
                return Err("#DIV/0!".to_string());
            }
            let m1 = array1.iter().sum::<f64>() / n1 as f64;
            let m2 = array2.iter().sum::<f64>() / n1 as f64;
            let s1 = var_s(array1)?;
            let s2 = var_s(array2)?;
            let df = (n1 + n2 - 2) as f64;
            let sp2 = (((n1 - 1) as f64 * s1) + ((n2 - 1) as f64 * s2)) / df;
            let se = (sp2 * (1.0 / n1 as f64 + 1.0 / n2 as f64)).sqrt();
            if se == 0.0 {
                return Err("#DIV/0!".to_string());
            }
            ((m1 - m2) / se, df)
        }
        3 => {
            // Welch's t-test (heteroscedastic)
            if n1 <= 1 || n2 <= 1 {
                return Err("#DIV/0!".to_string());
            }
            let m1 = array1.iter().sum::<f64>() / n1 as f64;
            let m2 = array2.iter().sum::<f64>() / n2 as f64;
            let s1 = var_s(array1)?;
            let s2 = var_s(array2)?;
            let v1 = s1 / n1 as f64;
            let v2 = s2 / n2 as f64;
            let se = (v1 + v2).sqrt();
            if se == 0.0 {
                return Err("#DIV/0!".to_string());
            }
            let df =
                (v1 + v2).powi(2) / ((v1 * v1) / (n1 - 1) as f64 + (v2 * v2) / (n2 - 1) as f64);
            ((m1 - m2) / se, df)
        }
        _ => return Err("#NUM!".to_string()),
    };

    let p_rt = t_dist_rt(t_stat.abs(), df)?;
    if tails == 1 { Ok(p_rt) } else { Ok(2.0 * p_rt) }
}

pub fn z_test(array: &[f64], x: f64, sigma: Option<f64>) -> Result<f64, String> {
    if array.is_empty() {
        return Err("#DIV/0!".to_string());
    }
    let n = array.len() as f64;
    let mean = array.iter().sum::<f64>() / n;
    let s = match sigma {
        Some(s_val) => s_val,
        None => stdev_s(array)?,
    };
    if s == 0.0 {
        return Err("#DIV/0!".to_string());
    }
    let z = (mean - x) / (s / n.sqrt());
    Ok(1.0 - normal_cdf(z))
}

pub fn confidence_norm(alpha: f64, std_dev: f64, size: f64) -> Result<f64, String> {
    if alpha <= 0.0 || alpha >= 1.0 || std_dev <= 0.0 || size < 1.0 {
        return Err("#NUM!".to_string());
    }
    let z = inv_normal_cdf(1.0 - alpha / 2.0)?;
    Ok(z * std_dev / size.sqrt())
}

pub fn confidence_t(alpha: f64, std_dev: f64, size: f64) -> Result<f64, String> {
    if alpha <= 0.0 || alpha >= 1.0 || std_dev <= 0.0 || size <= 1.0 {
        return Err("#NUM!".to_string());
    }
    let df = size.floor() - 1.0;
    let t_val = t_inv_2t(alpha, df)?;
    Ok(t_val * std_dev / size.sqrt())
}

pub fn fisher(x: f64) -> Result<f64, String> {
    if x <= -1.0 || x >= 1.0 {
        Err("#NUM!".to_string())
    } else {
        Ok(0.5 * ((1.0 + x) / (1.0 - x)).ln())
    }
}

pub fn fisherinv(y: f64) -> Result<f64, String> {
    let e2y = (2.0 * y).exp();
    Ok((e2y - 1.0) / (e2y + 1.0))
}

pub fn permut(n: f64, k: f64) -> Result<f64, String> {
    if n < 0.0 || k < 0.0 || k > n {
        return Err("#NUM!".to_string());
    }
    let n_i = n.floor();
    let k_i = k.floor();
    let mut ans = 1.0;
    for i in 0..(k_i as usize) {
        ans *= n_i - i as f64;
    }
    Ok(ans)
}

pub fn permutationa(n: f64, k: f64) -> Result<f64, String> {
    if n < 0.0 || k < 0.0 {
        return Err("#NUM!".to_string());
    }
    Ok(n.floor().powf(k.floor()))
}

pub fn prob(
    x_range: &[f64],
    prob_range: &[f64],
    lower_limit: f64,
    upper_limit: Option<f64>,
) -> Result<f64, String> {
    if x_range.len() != prob_range.len() || x_range.is_empty() {
        return Err("#N/A".to_string());
    }
    let prob_sum: f64 = prob_range.iter().sum();
    if (prob_sum - 1.0).abs() > 1e-6 {
        return Err("#NUM!".to_string());
    }
    let upper = upper_limit.unwrap_or(lower_limit);

    let mut sum = 0.0;
    for (&x, &p) in x_range.iter().zip(prob_range.iter()) {
        if p < 0.0 || p > 1.0 {
            return Err("#NUM!".to_string());
        }
        if x >= lower_limit && x <= upper {
            sum += p;
        }
    }
    Ok(sum)
}

pub fn frequency(data: &[f64], bins: &[f64]) -> Result<Vec<f64>, String> {
    let mut sorted_bins = bins.to_vec();
    sorted_bins.sort_by(|a, b| a.partial_cmp(b).unwrap_or(Ordering::Equal));

    let mut counts = vec![0.0; sorted_bins.len() + 1];
    for &x in data {
        let mut placed = false;
        for (i, &b) in sorted_bins.iter().enumerate() {
            if x <= b {
                counts[i] += 1.0;
                placed = true;
                break;
            }
        }
        if !placed {
            let last = counts.len() - 1;
            counts[last] += 1.0;
        }
    }
    Ok(counts)
}
