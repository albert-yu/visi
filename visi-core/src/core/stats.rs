// High-precision statistical functions for visi-core
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
        1.383_577_518_672_69e2,
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
    // Via erfc, not `0.5 * (1 + erf(...))`. In the left tail erf approaches
    // -1, so that form cancels catastrophically and eventually rounds to
    // exactly 0: NORM.S.DIST(-11, TRUE) came out as 0 instead of
    // 1.9106595744986622e-28, which real Excel reports in full. erfc keeps
    // the tail accurate all the way down (Excel still resolves
    // NORM.S.DIST(-30, TRUE) as 4.9067139271479094e-198).
    0.5 * erfc(-x / std::f64::consts::SQRT_2)
}

/// Error function erf(x). Delegates to `libm` (a pure-Rust fdlibm port,
/// full double precision).
pub fn erf(x: f64) -> f64 {
    libm::erf(x)
}

/// Complementary error function erfc(x) = 1 - erf(x). Uses libm's own
/// erfc directly (not `1.0 - erf(x)`) since that subtraction loses
/// precision for large x, where erf(x) is very close to 1.
pub fn erfc(x: f64) -> f64 {
    libm::erfc(x)
}

/// log|Gamma(x)|. Delegates to `libm` (a pure-Rust fdlibm port).
///
/// libm returns +inf at the non-positive-integer poles;
/// normalized to #NUM! at the dispatch boundary.
pub fn lgamma(x: f64) -> f64 {
    libm::lgamma(x)
}

/// Gamma function Gamma(x). Uses `libm::tgamma` rather than
/// `lgamma(x).exp()`: going through the logarithm and back costs several
/// significant digits, which shows up directly against Excel at integer
/// arguments where the answer is a factorial. GAMMA(34) is exactly 33! =
/// 8683317618811886495518194401280000000, which Excel displays as
/// 8.68331761881189E+36; the exp(lgamma) form gave
/// 8.68331761881199E+36, wrong from the 14th digit.
pub fn gamma(x: f64) -> f64 {
    if x <= 0.0 && x == x.floor() {
        return f64::NAN;
    }
    libm::tgamma(x)
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
/// Gamma for the integer and half-integer arguments the F, t, chi-square
/// and beta families produce (every one of them is some `df / 2`), built
/// by recurrence from `sqrt(pi)` rather than taken from `libm::tgamma`.
///
/// The recurrence multiplies small exact half-integers, so it stays near
/// half an ULP where this crate's `tgamma` drifts: `tgamma(1.5)` is 2.6
/// ULP high, and that error lands directly in the beta prefactor of
/// `incbeta`. Returns `None` for anything else, leaving the caller on
/// `tgamma`.
fn gamma_half_integer(a: f64) -> Option<f64> {
    let two_a = a * 2.0;
    if two_a <= 0.0 || two_a.fract() != 0.0 || two_a > 400.0 {
        return None;
    }
    let two_a = two_a as u32;
    if two_a.is_multiple_of(2) {
        // a is a positive integer: Gamma(a) = (a - 1)!
        let n = two_a / 2;
        let mut r = 1.0f64;
        for k in 2..n {
            r *= f64::from(k);
        }
        Some(r)
    } else {
        // a = n + 1/2: Gamma(a) = (n - 1/2)(n - 3/2)...(1/2) * sqrt(pi)
        let n = (two_a - 1) / 2;
        let mut r = std::f64::consts::PI.sqrt();
        for k in 0..n {
            r *= f64::from(k) + 0.5;
        }
        Some(r)
    }
}

/// `Gamma(a)` for the incomplete-beta prefactor, preferring the exact
/// recurrence where it applies.
fn beta_gamma(a: f64) -> f64 {
    gamma_half_integer(a).unwrap_or_else(|| libm::tgamma(a))
}

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

    // The prefactor x^a (1-x)^b / (a * B(a,b)), computed from tgamma
    // directly rather than as exp(a*ln x + b*ln(1-x) - lbeta).
    //
    // The log form routes everything through a single exponential, so the
    // *absolute* error of its argument becomes the *relative* error of the
    // result -- and lgamma(a+b) alone contributes ~1 ULP of a number that
    // can be 5 or more, which is ~5e-16 straight into the exponent. Over a
    // spread of (a, b, x) drawn from F-distribution degrees of freedom,
    // that cost a median of 12.7 ULP and a p90 of 51.6; going through
    // tgamma gives 1.9 and 4.6.
    //
    // Falls back to the log form whenever tgamma overflows (large a + b)
    // or the powers underflow, which is exactly where the logarithm earns
    // its keep -- hence a finiteness check rather than a fixed cutoff.
    let beta = beta_gamma(a) * beta_gamma(b) / beta_gamma(a + b);
    // `1 - x` rounds, and raising it to the power `b` multiplies that
    // rounding by `b` -- for b = 50 a half-ULP slip in `1 - x` became 15
    // ULP in the result. Recover the exact residual (`om + om_err` is
    // `1 - x` exactly) and apply the first-order correction, which brings
    // that same case back to 0.3 ULP.
    //
    // `(1.0 - om) - x` is exact either way: for x >= 0.5 the original
    // subtraction was already exact by Sterbenz and the residual is 0,
    // and for x < 0.5 `om` lands in (0.5, 1] where `1.0 - om` is exact.
    let om = 1.0 - x;
    let om_err = (1.0 - om) - x;
    let pow_om = om.powf(b) * (1.0 + b * om_err / om);
    let direct = x.powf(a) * pow_om / (a * beta);
    let front = if beta.is_finite() && beta > 0.0 && direct.is_finite() && direct > 0.0 {
        direct
    } else {
        let lbeta = lgamma(a) + lgamma(b) - lgamma(a + b);
        (a * x.ln() + b * (1.0 - x).ln() - lbeta).exp() / a
    };

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

        // Machine epsilon rather than the textbook 1e-15: that threshold
        // leaves about 1e-15 relative error, which is exactly the size of
        // the disagreements this was producing in the 15th significant
        // digit (FTEST over one fuzzed pair came out 0.941716332833876
        // where the true value is 0.94171633283387507 and Excel prints
        // 0.941716332833875).
        if (del - 1.0).abs() < f64::EPSILON {
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

    // Safeguarded Newton: a Newton step when it stays inside the current
    // bracket, otherwise a bisection step. incbeta is monotonically
    // increasing in x on [0, 1], so [0, 1] is always a valid starting
    // bracket and bisection alone would already converge -- Newton is
    // only an accelerator here, never something that can run away.
    //
    // The previous version was unguarded Newton that simply *clamped* an
    // overshooting step to [1e-12, 1-1e-12]. Those clamps are absorbing:
    // once a step overshot, x stuck to the boundary and the loop returned
    // it as the answer. That surfaced against real Excel as BETAINV
    // answering a flat 1e-12 or 0.999999999999, and (since
    // F.INV/F.INV.RT/FINV map y back through `df2*y / (df1*(1-y))`, which
    // blows up as y approaches 1) as F.INV returning ~1e12 instead of a
    // small number.
    let y = inv_normal_cdf(p)?;
    let h = 2.0 / (1.0 / (2.0 * a - 1.0) + 1.0 / (2.0 * b - 1.0));
    let w = (y * (h + 5.0 / 6.0 - 2.0 / (3.0 * h)).sqrt() / h)
        - (1.0 / (2.0 * b - 1.0) - 1.0 / (2.0 * a - 1.0)) * (y * y + 5.0 / 6.0 - 2.0 / (3.0 * h));
    let initial = a / (a + b * (2.0 * w).exp());
    let mut x = if initial.is_finite() && initial > 0.0 && initial < 1.0 {
        initial
    } else {
        0.5
    };

    let lbeta = lgamma(a) + lgamma(b) - lgamma(a + b);
    let mut lo = 0.0_f64;
    let mut hi = 1.0_f64;

    for _ in 0..200 {
        let err = incbeta(a, b, x) - p;
        if err.abs() < 1e-14 {
            return Ok(x);
        }
        // incbeta is increasing, so err < 0 means x is still too small.
        if err < 0.0 {
            lo = x;
        } else {
            hi = x;
        }

        let pdf = ((a - 1.0) * x.ln() + (b - 1.0) * (1.0 - x).ln() - lbeta).exp();
        let newton = if pdf > 0.0 && pdf.is_finite() {
            x - err / pdf
        } else {
            f64::NAN
        };
        let next = if newton.is_finite() && newton > lo && newton < hi {
            newton
        } else {
            0.5 * (lo + hi)
        };

        if (next - x).abs() <= 1e-16 * x.abs().max(f64::MIN_POSITIVE) {
            return Ok(next);
        }
        x = next;
        if hi - lo <= f64::EPSILON {
            break;
        }
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
    for &x in data {
        if x <= 0.0 {
            return Err("#NUM!".to_string());
        }
    }
    if data.iter().all(|&x| x == data[0]) {
        return Ok(data[0]);
    }
    let mut log_sum = 0.0;
    for &x in data {
        log_sum += x.ln();
    }
    Ok((log_sum / data.len() as f64).exp())
}

pub fn harmean(data: &[f64]) -> Result<f64, String> {
    if data.is_empty() {
        return Err("#N/A".to_string());
    }
    for &x in data {
        if x <= 0.0 {
            return Err("#NUM!".to_string());
        }
    }
    if data.iter().all(|&x| x == data[0]) {
        return Ok(data[0]);
    }
    let mut inv_sum = 0.0;
    for &x in data {
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
    let mut first_seen = std::collections::HashMap::new();

    for (idx, &x) in data.iter().enumerate() {
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

    // Real Excel returns tied modes in the order they first appear in the
    // data (same rule MODE.SNGL above uses to break a tie down to one
    // value), not sorted by value -- verified via
    // fuzz/fuzz_excel.py (INDEX(MODE.MULT(...), 1) picked the
    // first-encountered mode, not the smallest, when three values tied).
    let mut modes: Vec<(usize, f64)> = counts
        .into_iter()
        .filter(|&(_, count)| count == max_count)
        .map(|(key, _)| (first_seen[&key], f64::from_bits(key)))
        .collect();
    modes.sort_by_key(|&(idx, _)| idx);
    Ok(modes.into_iter().map(|(_, v)| v).collect())
}

pub fn trimmean(data: &[f64], percent: f64) -> Result<f64, String> {
    if data.is_empty() || !(0.0..1.0).contains(&percent) {
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
    // Skewness needs at least three observations; Excel reports #DIV/0!
    // below that for both SKEW and SKEW.P (confirmed directly -- two
    // values give #DIV/0!, three compute).
    if n < 3 {
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
    if data.is_empty() || !(0.0..=1.0).contains(&k) {
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
    if idx < 0.0 || idx > (n - 1) as f64 {
        return Err("#NUM!".to_string());
    }

    let j = idx.floor() as usize;
    let d = idx - j as f64;
    if j >= n - 1 {
        Ok(sorted[n - 1])
    } else {
        Ok(sorted[j] + d * (sorted[j + 1] - sorted[j]))
    }
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

    // Excel's PERCENTRANK truncates to `significance` digits rather than
    // rounding (a raw value of e.g. 0.4545 gives 0.454, not 0.455). The
    // nudge matters: a rank that is mathematically exactly 0.4 can land a
    // hair below it in f64 (0.39999999999999997), and truncating *that*
    // yields 0.399 where Excel reports 0.4.
    let mult = 10.0_f64.powi(significance as i32);
    let scaled = ans * mult;
    Ok((scaled + scaled.abs().max(1.0) * f64::EPSILON * 4.0).floor() / mult)
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

    // Excel's PERCENTRANK truncates to `significance` digits rather than
    // rounding (a raw value of e.g. 0.4545 gives 0.454, not 0.455). The
    // nudge matters: a rank that is mathematically exactly 0.4 can land a
    // hair below it in f64 (0.39999999999999997), and truncating *that*
    // yields 0.399 where Excel reports 0.4.
    let mult = 10.0_f64.powi(significance as i32);
    let scaled = ans * mult;
    Ok((scaled + scaled.abs().max(1.0) * f64::EPSILON * 4.0).floor() / mult)
}

// ============================================================================
// 5. Bivariate Statistics & Linear Regression
// ============================================================================

pub fn covariance_p(xs: &[f64], ys: &[f64]) -> Result<f64, String> {
    // Length mismatch is #N/A, but zero usable pairs is #DIV/0! -- both
    // confirmed against real Excel (e.g. CORREL over two ranges whose
    // every pair contains a text cell gives #DIV/0!, not #N/A).
    if xs.len() != ys.len() {
        return Err("#N/A".to_string());
    }
    if xs.is_empty() {
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
    // Length mismatch is #N/A, but zero usable pairs is #DIV/0! -- both
    // confirmed against real Excel (e.g. CORREL over two ranges whose
    // every pair contains a text cell gives #DIV/0!, not #N/A).
    if xs.len() != ys.len() {
        return Err("#N/A".to_string());
    }
    if xs.is_empty() {
        return Err("#DIV/0!".to_string());
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
    // Length mismatch is #N/A, but zero usable pairs is #DIV/0! -- both
    // confirmed against real Excel (e.g. CORREL over two ranges whose
    // every pair contains a text cell gives #DIV/0!, not #N/A).
    if xs.len() != ys.len() {
        return Err("#N/A".to_string());
    }
    if xs.is_empty() {
        return Err("#DIV/0!".to_string());
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
    if number_s < 0.0 || trials < 0.0 || number_s > trials || !(0.0..=1.0).contains(&probability_s)
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
    if trials < 0.0 || !(0.0..=1.0).contains(&probability_s) || !(0.0..=1.0).contains(&alpha) {
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
    if number_f < 0.0 || number_s < 1.0 || !(0.0..=1.0).contains(&probability_s) {
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
        // C(m_pop, x) is architecturally 0 once x is outside [0, m_pop],
        // and likewise C(n_pop - m_pop, n - x) once (n - x) is outside
        // [0, n_pop - m_pop] -- the lgamma-based log-combination formula
        // below assumes valid choose() arguments and produces a pole
        // (lgamma of a non-positive integer -> NaN) rather than 0 outside
        // that range, so this has to be checked before calling it.
        if x < 0.0 || x > m_pop || (n - x) < 0.0 || (n - x) > n_pop - m_pop {
            return 0.0;
        }
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
    // The upper incomplete gamma directly, not `1 - CDF`. For a large
    // statistic the CDF is within an ULP of 1 and the subtraction
    // underflows to exactly 0 -- CHITEST over a series with one large
    // term returned 0 where real Excel resolves 6.4e-103.
    Ok(regularized_gamma_q(df / 2.0, x / 2.0))
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

/// `categories` is the number of cells the two ranges originally held,
/// which is not the same as `actual.len()`: the caller has already dropped
/// pairs where either side was non-numeric, but Excel takes the degrees of
/// freedom from the *original* dimensions. With one text cell in a 2-cell
/// pair, one pair survives and Excel still evaluates against df = 1 rather
/// than the df = 0 the survivor count would give.
pub fn chisq_test(actual: &[f64], expected: &[f64], categories: usize) -> Result<f64, String> {
    if actual.len() != expected.len() {
        return Err("#N/A".to_string());
    }
    // No surviving pair is not an error: the statistic is simply 0, and
    // with the degrees of freedom coming from `categories` the answer is
    // 1. Real Excel returns 1 for two 3-cell ranges whose every pair holds
    // something non-numeric. (A raw size mismatch, and a `categories` of
    // fewer than 2, are both rejected by the caller before this point.)
    let mut chi2 = 0.0;
    for (&o, &e) in actual.iter().zip(expected.iter()) {
        // An expected frequency of exactly zero is the division itself
        // failing, so that is #DIV/0! and is checked here.
        //
        // A *negative* expected frequency is not rejected per element,
        // which is the non-obvious part. Excel just divides by it, letting
        // that term push the statistic down, and only reports #NUM! if the
        // total comes out negative -- so whether a negative expected value
        // is an error depends on the other terms:
        //
        //   CHITEST({1,2,3}, {5,-4,3})              = #NUM!   (chi2 = -5.8)
        //   CHITEST({-478.8,352.51,8.5}, {38,8.5,-75}) = 0    (chi2 ~ 20859)
        //
        // Rejecting `e < 0` up front got the second case wrong.
        if e == 0.0 {
            return Err("#DIV/0!".to_string());
        }
        chi2 += (o - e) * (o - e) / e;
    }
    if chi2 < 0.0 {
        return Err("#NUM!".to_string());
    }
    let df = (categories.max(1) - 1) as f64;
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
    // Via the symmetry I_y(a, b) = 1 - I_(1-y)(b, a), rather than
    // subtracting the left tail from 1. For a large F statistic that left
    // tail sits within an ULP or two of 1, so `1.0 - cdf` throws away most
    // of the answer's significant digits -- F.TEST agreed with Excel only
    // to about 12 of them. Forming 1-y directly as df2 / (df1*x + df2)
    // sidesteps the cancellation.
    let y_complement = df2 / (df1 * x + df2);
    Ok(incbeta(df2 / 2.0, df1 / 2.0, y_complement))
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
            // Paired. A genuine length mismatch is #N/A (though the
            // caller already checks raw sizes before pairwise-excluding),
            // but *too few usable pairs* is #DIV/0! -- there's no
            // denominator to divide by. Confirmed against real Excel:
            // T.TEST over two 5-cell ranges whose pairwise-valid overlap
            // is a single pair reports #DIV/0!, as does a pair of
            // identical (zero-variance) samples.
            if n1 != n2 {
                return Err("#N/A".to_string());
            }
            if n1 <= 1 {
                return Err("#DIV/0!".to_string());
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
            let m2 = array2.iter().sum::<f64>() / n2 as f64;
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
    // (e^2y - 1) / (e^2y + 1) is tanh(y), but computing it that way
    // overflows to inf/inf for y beyond ~355 and came back as #NUM! where
    // Excel simply reports 1. tanh saturates instead, which is also what
    // the identity is worth in f64 long before that point.
    Ok(y.tanh())
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
    if x_range.len() != prob_range.len() {
        return Err("#N/A".to_string());
    }
    // No usable probabilities at all is a probability sum of 0, which
    // fails the "must sum to 1" rule below -- Excel reports #NUM! for it,
    // not #N/A (confirmed with a probability range that is entirely text).
    let prob_sum: f64 = prob_range.iter().sum();
    if (prob_sum - 1.0).abs() > 1e-6 {
        return Err("#NUM!".to_string());
    }
    let upper = upper_limit.unwrap_or(lower_limit);

    let mut sum = 0.0;
    // Excel checks only that the probabilities sum to 1; it does *not*
    // reject an individual one outside [0, 1]. PROB({1,2}, {1.5,-0.5}, 0,
    // 3) is 1 in real Excel, and rejecting the negative there turned a
    // pairwise-excluded range that legitimately summed to 1 into #NUM!.
    for (&x, &p) in x_range.iter().zip(prob_range.iter()) {
        if x >= lower_limit && x <= upper {
            sum += p;
        }
    }
    Ok(sum)
}

pub fn frequency(data: &[f64], bins: &[f64]) -> Result<Vec<f64>, String> {
    // Excel sorts the bins internally to work out the interval each value
    // falls in, but reports each interval's count back at that bin's
    // *original* position in bins_array, with the overflow count last.
    // Returning the counts in sorted order instead
    // silently permutes the result whenever bins_array isn't already
    // ascending. Verified against real Excel with bins [25, -10, 8] over
    // data [5, -20, 30, 1, 12]: Excel gives [1, 1, 2, 1], i.e. the sorted
    // counts [1, 2, 1] mapped back through each bin's rank, then overflow.
    let mut order: Vec<usize> = (0..bins.len()).collect();
    order.sort_by(|&a, &b| bins[a].partial_cmp(&bins[b]).unwrap_or(Ordering::Equal));
    let sorted_bins: Vec<f64> = order.iter().map(|&i| bins[i]).collect();

    let mut sorted_counts = vec![0.0; sorted_bins.len() + 1];
    for &x in data {
        let mut placed = false;
        for (i, &b) in sorted_bins.iter().enumerate() {
            if x <= b {
                sorted_counts[i] += 1.0;
                placed = true;
                break;
            }
        }
        if !placed {
            let last = sorted_counts.len() - 1;
            sorted_counts[last] += 1.0;
        }
    }

    let mut counts = vec![0.0; bins.len() + 1];
    for (rank, &orig_idx) in order.iter().enumerate() {
        counts[orig_idx] = sorted_counts[rank];
    }
    counts[bins.len()] = sorted_counts[bins.len()];
    Ok(counts)
}
