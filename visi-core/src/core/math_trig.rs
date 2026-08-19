// High-precision Math and Trigonometry functions for visi-core
// Implements Excel-compatible trigonometric, hyperbolic, matrix, matrix inversion, series, combinatorics, and matrix/array math.

// ============================================================================
// 1. Trigonometric and Hyperbolic Functions
// ============================================================================

/// Excel's SIN/COS/TAN (and so anything built on them: COT, CSC, SEC)
/// refuse an argument at or beyond `2^27` radians with `#NUM!`, rather
/// than returning whatever a library sin/cos happens to reduce it to --
/// past that magnitude a double's ~15-16 significant digits can no
/// longer resolve which multiple of 2*pi the value is near, so any
/// answer would be numerically meaningless. Measured directly (win32com,
/// real Windows Excel): `SIN(134217727)` (2^27 - 1) is a real number,
/// `SIN(134217728)` (2^27) is `#NUM!`, and COS/TAN share the identical
/// boundary. fuzz/fuzz_excel.py seed 676008 hit this via
/// `CSC(F4^47)` where `F4^47` is on the order of 1e101.
pub(crate) const TRIG_ARG_LIMIT: f64 = 134_217_728.0; // 2^27

pub(crate) fn check_trig_domain(x: f64) -> Result<(), String> {
    if x.abs() >= TRIG_ARG_LIMIT {
        Err("#NUM!".to_string())
    } else {
        Ok(())
    }
}

pub fn acosh(x: f64) -> Result<f64, String> {
    if x < 1.0 {
        Err("#NUM!".to_string())
    } else {
        Ok(x.acosh())
    }
}

pub fn acot(x: f64) -> Result<f64, String> {
    // Returns acot(x) in range (0, PI)
    if x == 0.0 {
        Ok(std::f64::consts::FRAC_PI_2)
    } else {
        let val = (1.0 / x).atan();
        if val < 0.0 {
            Ok(val + std::f64::consts::PI)
        } else {
            Ok(val)
        }
    }
}

pub fn acoth(x: f64) -> Result<f64, String> {
    if x.abs() <= 1.0 {
        Err("#NUM!".to_string())
    } else {
        // atanh(1/x), not 0.5 * ln((x+1)/(x-1)): for large |x| that ratio
        // approaches 1 and the logarithm loses most of its significant
        // digits. ACOTH(-165) is -0.006060680266172405..., and the
        // logarithm form gave -0.006060680266172425 -- wrong from the
        // 15th digit, which is exactly where Excel's display lands.
        Ok((1.0 / x).atanh())
    }
}

pub fn asinh(x: f64) -> Result<f64, String> {
    Ok(x.asinh())
}

pub fn atan2(x: f64, y: f64) -> Result<f64, String> {
    // Excel's ATAN2 syntax is ATAN2(x_num, y_num) where x is 1st arg, y is 2nd arg
    if x == 0.0 && y == 0.0 {
        Err("#DIV/0!".to_string())
    } else {
        Ok(y.atan2(x))
    }
}

pub fn atanh(x: f64) -> Result<f64, String> {
    if x.abs() >= 1.0 {
        Err("#NUM!".to_string())
    } else {
        Ok(x.atanh())
    }
}

pub fn cosh(x: f64) -> Result<f64, String> {
    Ok(x.cosh())
}

pub fn cot(x: f64) -> Result<f64, String> {
    check_trig_domain(x)?;
    let tan_val = x.tan();
    if tan_val == 0.0 {
        Err("#DIV/0!".to_string())
    } else {
        Ok(1.0 / tan_val)
    }
}

pub fn coth(x: f64) -> Result<f64, String> {
    if x == 0.0 {
        Err("#DIV/0!".to_string())
    } else {
        // Not `x.cosh() / x.sinh()`: both overflow to `f64::INFINITY` well
        // before `|x|` gets anywhere near where `coth` itself is
        // ill-behaved (~710), leaving `inf / inf = NaN` -- which this
        // engine's NaN guard then reports as `#NUM!` for an `x` where
        // Excel returns a perfectly good answer near +-1 (measured:
        // COTH(47692.3) is `1` in real Excel; fuzz/fuzz_excel.py seed
        // 711993). `tanh` saturates to +-1 directly with no such overflow,
        // so go through it the way `cot` already goes through `tan`.
        Ok(1.0 / x.tanh())
    }
}

pub fn csc(x: f64) -> Result<f64, String> {
    check_trig_domain(x)?;
    let sin_val = x.sin();
    if sin_val == 0.0 {
        Err("#DIV/0!".to_string())
    } else {
        Ok(1.0 / sin_val)
    }
}

pub fn csch(x: f64) -> Result<f64, String> {
    let sinh_val = x.sinh();
    if sinh_val == 0.0 {
        Err("#DIV/0!".to_string())
    } else {
        Ok(1.0 / sinh_val)
    }
}

pub fn degrees(radians: f64) -> Result<f64, String> {
    Ok(radians * (180.0 / std::f64::consts::PI))
}

pub fn radians(degrees: f64) -> Result<f64, String> {
    Ok(degrees * (std::f64::consts::PI / 180.0))
}

pub fn sec(x: f64) -> Result<f64, String> {
    check_trig_domain(x)?;
    let cos_val = x.cos();
    if cos_val == 0.0 {
        Err("#DIV/0!".to_string())
    } else {
        Ok(1.0 / cos_val)
    }
}

pub fn sech(x: f64) -> Result<f64, String> {
    Ok(1.0 / x.cosh())
}

pub fn sinh(x: f64) -> Result<f64, String> {
    Ok(x.sinh())
}

pub fn sqrtpi(x: f64) -> Result<f64, String> {
    if x < 0.0 {
        Err("#NUM!".to_string())
    } else {
        Ok((x * std::f64::consts::PI).sqrt())
    }
}

pub fn tanh(x: f64) -> Result<f64, String> {
    Ok(x.tanh())
}

// ============================================================================
// 2. Rounding and Integer Arithmetic
// ============================================================================

pub fn ceiling_math(x: f64, significance: Option<f64>, mode: Option<f64>) -> Result<f64, String> {
    let sig = significance.unwrap_or(1.0);
    if sig == 0.0 {
        return Ok(0.0);
    }
    let m = mode.unwrap_or(0.0);
    let sig_abs = sig.abs();

    if x >= 0.0 {
        Ok((x / sig_abs).ceil() * sig_abs)
    } else {
        if m != 0.0 {
            // Round away from zero (towards negative infinity)
            Ok((x / sig_abs).floor() * sig_abs)
        } else {
            // Round towards zero
            Ok((x / sig_abs).ceil() * sig_abs)
        }
    }
}

pub fn floor_math(x: f64, significance: Option<f64>, mode: Option<f64>) -> Result<f64, String> {
    let sig = significance.unwrap_or(1.0);
    if sig == 0.0 {
        return Ok(0.0);
    }
    let m = mode.unwrap_or(0.0);
    let sig_abs = sig.abs();

    if x >= 0.0 {
        Ok((x / sig_abs).floor() * sig_abs)
    } else {
        if m != 0.0 {
            // Round towards zero
            Ok((x / sig_abs).ceil() * sig_abs)
        } else {
            // Round away from zero (towards negative infinity)
            Ok((x / sig_abs).floor() * sig_abs)
        }
    }
}

pub fn even(x: f64) -> Result<f64, String> {
    if x == 0.0 {
        return Ok(0.0);
    }
    let sign = x.signum();
    let ax = x.abs();
    let mut ceiled = ax.ceil();
    if (ceiled as i64) % 2 != 0 {
        ceiled += 1.0;
    }
    Ok(sign * ceiled)
}

pub fn odd(x: f64) -> Result<f64, String> {
    if x == 0.0 {
        return Ok(1.0);
    }
    let sign = x.signum();
    let ax = x.abs();
    let mut ceiled = ax.ceil();
    if (ceiled as i64) % 2 == 0 {
        ceiled += 1.0;
    }
    Ok(sign * ceiled)
}

pub fn mround(x: f64, multiple: f64) -> Result<f64, String> {
    if multiple == 0.0 {
        return Ok(0.0);
    }
    if (x > 0.0 && multiple < 0.0) || (x < 0.0 && multiple > 0.0) {
        return Err("#NUM!".to_string());
    }
    Ok((x / multiple).round() * multiple)
}

pub fn quotient(numerator: f64, denominator: f64) -> Result<f64, String> {
    if denominator == 0.0 {
        Err("#DIV/0!".to_string())
    } else {
        let q = (numerator / denominator).trunc();
        Ok(if q == 0.0 { 0.0 } else { q })
    }
}

pub fn sign(x: f64) -> Result<f64, String> {
    if x > 0.0 {
        Ok(1.0)
    } else if x < 0.0 {
        Ok(-1.0)
    } else {
        Ok(0.0)
    }
}

pub fn trunc(x: f64, digits: Option<f64>) -> Result<f64, String> {
    let d = digits.unwrap_or(0.0).round() as i32;
    let factor = 10.0_f64.powi(d);
    Ok((x * factor).trunc() / factor)
}

// ============================================================================
// 3. Number Base Conversions & Roman Numerals
// ============================================================================

pub fn base(number: f64, radix: f64, min_length: Option<f64>) -> Result<String, String> {
    let num = number.floor() as i64;
    let r = radix.floor() as u32;
    if num < 0 || !(2..=36).contains(&r) {
        return Err("#NUM!".to_string());
    }
    let min_len = min_length.unwrap_or(0.0).floor() as usize;

    let chars = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    if num == 0 {
        let s = "0".to_string();
        if s.len() < min_len {
            return Ok(format!("{:0>1$}", s, min_len));
        }
        return Ok(s);
    }

    let mut n = num as u64;
    let mut digits = Vec::new();
    while n > 0 {
        let rem = (n % (r as u64)) as usize;
        digits.push(chars.as_bytes()[rem] as char);
        n /= r as u64;
    }
    digits.reverse();
    let res: String = digits.into_iter().collect();
    if res.len() < min_len {
        Ok(format!("{:0>1$}", res, min_len))
    } else {
        Ok(res)
    }
}

pub fn decimal(text: &str, radix: f64) -> Result<f64, String> {
    let r = radix.floor() as u32;
    if !(2..=36).contains(&r) {
        return Err("#NUM!".to_string());
    }
    let s = text.trim();
    if s.is_empty() {
        return Err("#VALUE!".to_string());
    }
    match u64::from_str_radix(s, r) {
        Ok(val) => Ok(val as f64),
        Err(_) => Err("#NUM!".to_string()),
    }
}

pub fn arabic(text: &str) -> Result<f64, String> {
    let s = text.trim().to_uppercase();
    if s.is_empty() {
        return Ok(0.0);
    }
    let is_neg = s.starts_with('-');
    let roman = if is_neg { &s[1..] } else { &s[..] };

    let val_of = |c: char| -> Result<i64, String> {
        match c {
            'I' => Ok(1),
            'V' => Ok(5),
            'X' => Ok(10),
            'L' => Ok(50),
            'C' => Ok(100),
            'D' => Ok(500),
            'M' => Ok(1000),
            _ => Err("#VALUE!".to_string()),
        }
    };

    let mut total = 0i64;
    let mut prev = 0i64;

    for c in roman.chars().rev() {
        let curr = val_of(c)?;
        if curr < prev {
            total -= curr;
        } else {
            total += curr;
            prev = curr;
        }
    }

    if is_neg {
        Ok(-total as f64)
    } else {
        Ok(total as f64)
    }
}

/// Excel's ROMAN, including its four progressively "concise" forms.
///
/// Form 0 is classic notation, which only subtracts a power of ten one or
/// two places below the numeral it precedes (CM, CD, XC, XL, IX, IV).
/// Forms 1-4 unlock progressively longer "reaches" and additionally allow
/// the half-power numerals (V, L, D) to be subtracted, which is how Excel
/// gets its shorter non-classical spellings. `form` used to be ignored
/// outright, so every form rendered as form 0.
///
/// Writing the numerals in descending order M D C L X V I (indices 0..6),
/// let `reach` be the index distance from the minuend to the subtrahend.
/// The concision level a pair needs is then:
///   * power-of-ten subtrahend (I, X, C):  `2 * ((reach - 1) / 2)`
///   * half-power subtrahend  (V, L, D):   `1 + 2 * ((reach - 2) / 2)`
///
/// That reproduces every documented pair at exactly the form Excel first
/// uses it: CD/XL/IV and CM/XC/IX at form 0; LD/LM/VL at 1; XD/XM at 2;
/// VD/VM at 3; ID/IM at 4. Verified against real Excel across all five
/// forms of 45, 499, 990, 1481 and 1999.
pub fn roman(number: f64, form: Option<f64>) -> Result<String, String> {
    let n = number.floor() as i64;
    if !(1..=3999).contains(&n) {
        return Err("#VALUE!".to_string());
    }
    let level = form.unwrap_or(0.0).floor().clamp(0.0, 4.0) as usize;

    const NUMERALS: [(i64, &str); 7] = [
        (1000, "M"),
        (500, "D"),
        (100, "C"),
        (50, "L"),
        (10, "X"),
        (5, "V"),
        (1, "I"),
    ];

    let mut candidates: Vec<(i64, String)> = NUMERALS
        .iter()
        .map(|(v, sym)| (*v, (*sym).to_string()))
        .collect();
    for (i, (big, big_sym)) in NUMERALS.iter().enumerate() {
        for (j, (small, small_sym)) in NUMERALS.iter().enumerate().skip(i + 1) {
            let reach = j - i;
            // Even indices are the powers of ten (M, C, X, I); odd ones
            // the half-powers (D, L, V).
            let needed = if j % 2 == 0 {
                2 * ((reach - 1) / 2)
            } else if reach < 2 {
                // e.g. "DM"/"LC"/"VX" -- worth the same as the plain
                // numeral, never a real spelling.
                continue;
            } else {
                1 + 2 * ((reach - 2) / 2)
            };
            if needed <= level {
                candidates.push((big - small, format!("{}{}", small_sym, big_sym)));
            }
        }
    }
    // Greedy needs richest-first; ties go to the shorter spelling.
    candidates.sort_by(|a, b| b.0.cmp(&a.0).then(a.1.len().cmp(&b.1.len())));

    let mut rem = n;
    let mut res = String::new();
    for (val, sym) in &candidates {
        while rem >= *val {
            res.push_str(sym);
            rem -= *val;
        }
        if rem == 0 {
            break;
        }
    }
    Ok(res)
}

// ============================================================================
// 4. Combinatorics, Factorials, Power, Series, GCD/LCM
// ============================================================================

pub fn combin(n: f64, k: f64) -> Result<f64, String> {
    let ni = n.floor() as i64;
    let ki = k.floor() as i64;
    if ni < 0 || ki < 0 || ki > ni {
        return Err("#NUM!".to_string());
    }
    if ki == 0 || ki == ni {
        return Ok(1.0);
    }
    let k_min = ki.min(ni - ki);
    let mut ans = 1.0;
    for i in 1..=k_min {
        ans = ans * (ni - i + 1) as f64 / i as f64;
    }
    Ok(ans.round())
}

pub fn combina(n: f64, k: f64) -> Result<f64, String> {
    let ni = n.floor() as i64;
    let ki = k.floor() as i64;
    if ni < 0 || ki < 0 {
        return Err("#NUM!".to_string());
    }
    if ni == 0 && ki == 0 {
        return Ok(1.0);
    }
    if ni == 0 {
        return Ok(0.0);
    }
    combin((ni + ki - 1) as f64, ki as f64)
}

pub fn fact(n: f64) -> Result<f64, String> {
    let ni = n.floor() as i64;
    if !(0..=170).contains(&ni) {
        return Err("#NUM!".to_string());
    }
    let mut ans = 1.0;
    for i in 1..=ni {
        ans *= i as f64;
    }
    Ok(ans)
}

pub fn factdouble(n: f64) -> Result<f64, String> {
    let ni = n.floor() as i64;
    if !(-1..=300).contains(&ni) {
        return Err("#NUM!".to_string());
    }
    if ni <= 0 {
        return Ok(1.0);
    }
    let mut ans = 1.0;
    let mut i = ni;
    while i > 0 {
        ans *= i as f64;
        i -= 2;
    }
    Ok(ans)
}

pub fn gcd(nums: &[f64]) -> Result<f64, String> {
    if nums.is_empty() {
        return Err("#VALUE!".to_string());
    }
    // Excel's GCD/LCM are defined only for non-negative arguments and
    // report #NUM! for a negative one, rather than quietly working on its
    // magnitude the way `.abs()` below otherwise would.
    if nums.iter().any(|n| *n < 0.0) {
        return Err("#NUM!".to_string());
    }
    let mut result = nums[0].floor().abs() as u64;
    fn gcd_two(mut a: u64, mut b: u64) -> u64 {
        while b != 0 {
            let t = b;
            b = a % b;
            a = t;
        }
        a
    }
    for &num in &nums[1..] {
        let n = num.floor().abs() as u64;
        result = gcd_two(result, n);
    }
    Ok(result as f64)
}

pub fn lcm(nums: &[f64]) -> Result<f64, String> {
    if nums.is_empty() {
        return Err("#VALUE!".to_string());
    }
    if nums.iter().any(|n| *n < 0.0) {
        return Err("#NUM!".to_string());
    }
    fn gcd_two(mut a: u64, mut b: u64) -> u64 {
        while b != 0 {
            let t = b;
            b = a % b;
            a = t;
        }
        a
    }
    let mut result = nums[0].floor().abs() as u64;
    if result == 0 {
        return Ok(0.0);
    }
    for &num in &nums[1..] {
        let n = num.floor().abs() as u64;
        if n == 0 {
            return Ok(0.0);
        }
        let g = gcd_two(result, n);
        result = (result / g) * n;
    }
    Ok(result as f64)
}

pub fn multinomial(nums: &[f64]) -> Result<f64, String> {
    let mut sum_n = 0i64;
    for &num in nums {
        let ni = num.floor() as i64;
        if ni < 0 {
            return Err("#NUM!".to_string());
        }
        sum_n += ni;
    }
    let top = fact(sum_n as f64)?;
    let mut bot = 1.0;
    for &num in nums {
        bot *= fact(num)?;
    }
    if bot == 0.0 {
        Err("#DIV/0!".to_string())
    } else {
        Ok((top / bot).round())
    }
}

pub fn power(number: f64, p: f64) -> Result<f64, String> {
    if number == 0.0 && p == 0.0 {
        // Excel declines to pick a value for 0^0: both POWER(0, 0) and
        // 0^0 are #NUM!. The `^` operator already did this; POWER was
        // returning 1, so POWER(<blank>, <blank>) silently became 1.
        Err("#NUM!".to_string())
    } else if number == 0.0 && p < 0.0 {
        Err("#DIV/0!".to_string())
    } else if number < 0.0 && (p.floor() != p || p.abs() > 1e6) {
        Err("#NUM!".to_string())
    } else {
        Ok(number.powf(p))
    }
}

pub fn seriessum(x: f64, n: f64, m: f64, coefficients: &[f64]) -> Result<f64, String> {
    let mut sum = 0.0;
    for (i, &a) in coefficients.iter().enumerate() {
        let p = n + (i as f64) * m;
        sum += a * x.powf(p);
    }
    Ok(sum)
}

// ============================================================================
// 5. Matrix and Array Operations
// ============================================================================

pub fn mdeterm(matrix: &[Vec<f64>]) -> Result<f64, String> {
    let n = matrix.len();
    if n == 0 || matrix.iter().any(|row| row.len() != n) {
        return Err("#VALUE!".to_string());
    }

    let mut mat = matrix.to_vec();
    let mut det = 1.0;

    for i in 0..n {
        let mut pivot = i;
        for j in (i + 1)..n {
            if mat[j][i].abs() > mat[pivot][i].abs() {
                pivot = j;
            }
        }
        if mat[pivot][i] == 0.0 {
            return Ok(0.0);
        }
        if i != pivot {
            mat.swap(i, pivot);
            det = -det;
        }
        det *= mat[i][i];

        let pivot_val = mat[i][i];
        for j in (i + 1)..n {
            let factor = mat[j][i] / pivot_val;
            let row_i = mat[i][i + 1..n].to_vec();
            for (target, src) in mat[j][i + 1..n].iter_mut().zip(row_i.iter()) {
                *target -= factor * src;
            }
        }
    }

    Ok(det)
}

pub fn minverse(matrix: &[Vec<f64>]) -> Result<Vec<Vec<f64>>, String> {
    let n = matrix.len();
    if n == 0 || matrix.iter().any(|row| row.len() != n) {
        return Err("#VALUE!".to_string());
    }

    let mut aug = vec![vec![0.0; 2 * n]; n];
    for i in 0..n {
        for j in 0..n {
            aug[i][j] = matrix[i][j];
        }
        aug[i][n + i] = 1.0;
    }

    for i in 0..n {
        let mut pivot = i;
        for j in (i + 1)..n {
            if aug[j][i].abs() > aug[pivot][i].abs() {
                pivot = j;
            }
        }
        if aug[pivot][i].abs() < 1e-12 {
            return Err("#NUM!".to_string());
        }
        if i != pivot {
            aug.swap(i, pivot);
        }

        let div = aug[i][i];
        for val in aug[i].iter_mut() {
            *val /= div;
        }

        for j in 0..n {
            if j != i {
                let factor = aug[j][i];
                let row_i = aug[i].clone();
                for (target, src) in aug[j].iter_mut().zip(row_i.iter()) {
                    *target -= factor * src;
                }
            }
        }
    }

    let mut inv = vec![vec![0.0; n]; n];
    for i in 0..n {
        for j in 0..n {
            inv[i][j] = aug[i][n + j];
        }
    }
    Ok(inv)
}

pub fn munit(dimension: f64) -> Result<Vec<Vec<f64>>, String> {
    let n = dimension.floor() as usize;
    if n == 0 {
        return Err("#VALUE!".to_string());
    }
    let mut mat = vec![vec![0.0; n]; n];
    for (i, row) in mat.iter_mut().enumerate() {
        row[i] = 1.0;
    }
    Ok(mat)
}

pub fn percentof(data_value: f64, target_value: f64) -> Result<f64, String> {
    if target_value == 0.0 {
        Err("#DIV/0!".to_string())
    } else {
        Ok(data_value / target_value)
    }
}

pub fn sumproduct(arrays: &[Vec<f64>]) -> Result<f64, String> {
    if arrays.is_empty() {
        return Ok(0.0);
    }
    let len = arrays[0].len();
    if arrays.iter().any(|arr| arr.len() != len) {
        return Err("#VALUE!".to_string());
    }

    let mut sum = 0.0;
    for i in 0..len {
        let mut prod = 1.0;
        for arr in arrays {
            prod *= arr[i];
        }
        sum += prod;
    }
    Ok(sum)
}

pub fn sumsq(nums: &[f64]) -> Result<f64, String> {
    Ok(nums.iter().map(|&x| x * x).sum())
}

pub fn sumx2my2(xs: &[f64], ys: &[f64]) -> Result<f64, String> {
    if xs.len() != ys.len() {
        return Err("#N/A".to_string());
    }
    Ok(xs.iter().zip(ys.iter()).map(|(&x, &y)| x * x - y * y).sum())
}

pub fn sumx2py2(xs: &[f64], ys: &[f64]) -> Result<f64, String> {
    if xs.len() != ys.len() {
        return Err("#N/A".to_string());
    }
    Ok(xs.iter().zip(ys.iter()).map(|(&x, &y)| x * x + y * y).sum())
}

pub fn sumxmy2(xs: &[f64], ys: &[f64]) -> Result<f64, String> {
    if xs.len() != ys.len() {
        return Err("#N/A".to_string());
    }
    Ok(xs
        .iter()
        .zip(ys.iter())
        .map(|(&x, &y)| (x - y) * (x - y))
        .sum())
}

pub fn sequence(
    rows: f64,
    cols: Option<f64>,
    start: Option<f64>,
    step: Option<f64>,
) -> Result<Vec<Vec<f64>>, String> {
    let r = rows.floor() as usize;
    let c = cols.unwrap_or(1.0).floor() as usize;
    if r == 0 || c == 0 {
        return Err("#VALUE!".to_string());
    }
    let mut curr = start.unwrap_or(1.0);
    let st = step.unwrap_or(1.0);

    let mut grid = vec![vec![0.0; c]; r];
    for row in grid.iter_mut() {
        for val in row.iter_mut() {
            *val = curr;
            curr += st;
        }
    }
    Ok(grid)
}

pub fn randarray(
    rows: Option<f64>,
    cols: Option<f64>,
    min: Option<f64>,
    max: Option<f64>,
    whole_number: Option<bool>,
) -> Result<Vec<Vec<f64>>, String> {
    use rand::Rng;
    let r = rows.unwrap_or(1.0).floor() as usize;
    let c = cols.unwrap_or(1.0).floor() as usize;
    let min_val = min.unwrap_or(0.0);
    let max_val = max.unwrap_or(1.0);
    let is_whole = whole_number.unwrap_or(false);

    if r == 0 || c == 0 || min_val > max_val {
        return Err("#VALUE!".to_string());
    }

    let mut rng = rand::thread_rng();
    let mut grid = vec![vec![0.0; c]; r];

    for row in grid.iter_mut() {
        for cell in row.iter_mut() {
            let val = rng.gen_range(min_val..=max_val);
            *cell = if is_whole { val.round() } else { val };
        }
    }
    Ok(grid)
}
