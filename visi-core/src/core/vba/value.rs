use std::fmt;
use std::rc::Rc;

use super::host::ObjRef;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct VbaError {
    pub number: i32,

    pub description: String,
}

impl VbaError {
    pub fn invalid_call() -> Self {
        Self::new(5, "Invalid procedure call or argument")
    }
    pub fn overflow() -> Self {
        Self::new(6, "Overflow")
    }
    pub fn subscript() -> Self {
        Self::new(9, "Subscript out of range")
    }
    pub fn div_by_zero() -> Self {
        Self::new(11, "Division by zero")
    }
    pub fn type_mismatch() -> Self {
        Self::new(13, "Type mismatch")
    }
    pub fn invalid_null() -> Self {
        Self::new(94, "Invalid use of Null")
    }

    pub fn new(number: i32, description: impl Into<String>) -> Self {
        Self {
            number,
            description: description.into(),
        }
    }
}

impl fmt::Display for VbaError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "run-time error {}: {}", self.number, self.description)
    }
}

impl std::error::Error for VbaError {}

pub type VResult<T> = Result<T, VbaError>;

#[derive(Debug, Clone, PartialEq)]
pub enum Variant {
    Empty,
    Null,
    Boolean(bool),
    Integer(i16),
    Long(i32),
    Single(f32),
    Double(f64),
    Currency(i64),
    Date(f64),
    Str(String),
    ErrValue(i32),
    Object(ObjRef),
    Array(Rc<VarArray>),
}

#[derive(Debug, Clone, PartialEq)]
pub struct VarArray {
    pub rows: usize,
    pub cols: usize,
    pub values: Vec<Variant>,
}

impl VarArray {
    pub fn get(&self, row: usize, col: usize) -> VResult<Variant> {
        if row < 1 || col < 1 || row > self.rows || col > self.cols {
            return Err(VbaError::subscript());
        }
        Ok(self.values[(row - 1) * self.cols + (col - 1)].clone())
    }

    pub fn ubound(&self, dim: usize) -> VResult<usize> {
        match dim {
            1 => Ok(self.rows),
            2 => Ok(self.cols),
            _ => Err(VbaError::subscript()),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ArithMode {
    Constant,

    Promote,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
enum NumClass {
    Integer,
    Long,
    Currency,
    Single,
    Double,
}

impl Variant {
    pub fn type_name(&self) -> &'static str {
        match self {
            Variant::Empty => "Empty",
            Variant::Null => "Null",
            Variant::Boolean(_) => "Boolean",
            Variant::Integer(_) => "Integer",
            Variant::Long(_) => "Long",
            Variant::Single(_) => "Single",
            Variant::Double(_) => "Double",
            Variant::Currency(_) => "Currency",
            Variant::Date(_) => "Date",
            Variant::Str(_) => "String",
            Variant::ErrValue(_) => "Error",
            Variant::Object(o) => o.type_name(),
            Variant::Array(_) => "Variant()",
        }
    }

    pub fn error_number(&self) -> Option<i32> {
        match self {
            Variant::ErrValue(n) => Some(*n),
            _ => None,
        }
    }

    pub fn as_object(&self) -> Option<&ObjRef> {
        match self {
            Variant::Object(o) => Some(o),
            _ => None,
        }
    }

    pub fn is_null(&self) -> bool {
        matches!(self, Variant::Null)
    }

    pub fn is_empty(&self) -> bool {
        matches!(self, Variant::Empty)
    }

    fn num_class(&self) -> Option<NumClass> {
        Some(match self {
            Variant::Empty | Variant::Boolean(_) | Variant::Integer(_) => NumClass::Integer,
            Variant::Long(_) => NumClass::Long,
            Variant::Currency(_) => NumClass::Currency,
            Variant::Single(_) => NumClass::Single,
            Variant::Double(_) | Variant::Date(_) | Variant::Str(_) => NumClass::Double,
            Variant::Null => return None,
            Variant::ErrValue(_) | Variant::Object(_) | Variant::Array(_) => NumClass::Double,
        })
    }

    pub fn to_f64(&self) -> VResult<f64> {
        Ok(match self {
            Variant::Empty => 0.0,
            Variant::Boolean(b) => {
                if *b {
                    -1.0
                } else {
                    0.0
                }
            }
            Variant::Integer(v) => *v as f64,
            Variant::Long(v) => *v as f64,
            Variant::Single(v) => *v as f64,
            Variant::Double(v) | Variant::Date(v) => *v,
            Variant::Currency(v) => *v as f64 / 10_000.0,
            Variant::Str(s) => parse_vba_number(s)?,
            Variant::Null => return Err(VbaError::invalid_null()),
            Variant::ErrValue(_) | Variant::Object(_) | Variant::Array(_) => {
                return Err(VbaError::type_mismatch());
            }
        })
    }

    pub fn to_vba_string(&self) -> VResult<String> {
        Ok(match self {
            Variant::Empty => String::new(),
            Variant::Null => return Err(VbaError::invalid_null()),
            Variant::Boolean(b) => if *b { "True" } else { "False" }.to_string(),
            Variant::Integer(v) => v.to_string(),
            Variant::Long(v) => v.to_string(),
            Variant::Single(v) => format_number(*v as f64),
            Variant::Double(v) => format_number(*v),
            Variant::Currency(v) => format_currency(*v),
            Variant::Date(v) => format_vba_date(*v),
            Variant::Str(s) => s.clone(),
            Variant::ErrValue(n) => format!("Error {n}"),
            Variant::Object(_) | Variant::Array(_) => return Err(VbaError::type_mismatch()),
        })
    }

    pub fn to_bool(&self) -> VResult<bool> {
        match self {
            Variant::Boolean(b) => Ok(*b),
            Variant::Null => Err(VbaError::invalid_null()),
            Variant::Str(s) if bool_word(s).is_some() => Ok(bool_word(s).unwrap_or(false)),
            other => Ok(other.to_f64()? != 0.0),
        }
    }

    pub fn to_bool_condition(&self) -> VResult<bool> {
        match self {
            Variant::Null => Ok(false),
            other => other.to_bool(),
        }
    }

    pub fn from_literal(value: f64, has_fraction_or_exponent: bool) -> Variant {
        if has_fraction_or_exponent {
            return Variant::Double(value);
        }
        if value >= i16::MIN as f64 && value <= i16::MAX as f64 {
            Variant::Integer(value as i16)
        } else if value >= i32::MIN as f64 && value <= i32::MAX as f64 {
            Variant::Long(value as i32)
        } else {
            Variant::Double(value)
        }
    }

    fn pack_mode(value: f64, class: NumClass, mode: ArithMode) -> VResult<Variant> {
        if mode == ArithMode::Constant {
            return Self::pack(value, class);
        }
        let mut class = class;
        loop {
            match Self::pack(value, class) {
                Ok(v) => return Ok(v),
                Err(e) => {
                    class = match class {
                        NumClass::Integer => NumClass::Long,
                        NumClass::Long | NumClass::Currency | NumClass::Single => NumClass::Double,
                        NumClass::Double => return Err(e),
                    };
                }
            }
        }
    }

    fn pack(value: f64, class: NumClass) -> VResult<Variant> {
        Ok(match class {
            NumClass::Integer => {
                if value < i16::MIN as f64 || value > i16::MAX as f64 {
                    return Err(VbaError::overflow());
                }
                Variant::Integer(value as i16)
            }
            NumClass::Long => {
                if value < i32::MIN as f64 || value > i32::MAX as f64 {
                    return Err(VbaError::overflow());
                }
                Variant::Long(value as i32)
            }
            NumClass::Currency => {
                let scaled = bankers_round(value * 10_000.0);
                if scaled < i64::MIN as f64 || scaled > i64::MAX as f64 {
                    return Err(VbaError::overflow());
                }
                Variant::Currency(scaled as i64)
            }
            NumClass::Single => {
                let as_single = value as f32;
                if !as_single.is_finite() && value.is_finite() {
                    return Err(VbaError::overflow());
                }
                Variant::Single(as_single)
            }
            NumClass::Double => Variant::Double(value),
        })
    }

    fn arith_type(lhs: &Variant, rhs: &Variant) -> VResult<NumClass> {
        let l = lhs.num_class().ok_or_else(VbaError::invalid_null)?;
        let r = rhs.num_class().ok_or_else(VbaError::invalid_null)?;
        let pair = (l.min(r), l.max(r));
        if pair == (NumClass::Long, NumClass::Single) {
            return Ok(NumClass::Double);
        }
        Ok(l.max(r))
    }
}

pub fn bool_word(s: &str) -> Option<bool> {
    let t = s.trim();
    if t.eq_ignore_ascii_case("true") {
        Some(true)
    } else if t.eq_ignore_ascii_case("false") {
        Some(false)
    } else {
        None
    }
}

fn logical_operand(v: &Variant) -> Variant {
    match v {
        Variant::Str(s) => match bool_word(s) {
            Some(b) => Variant::Boolean(b),
            None => v.clone(),
        },
        _ => v.clone(),
    }
}

fn logical_pair(lhs: &Variant, rhs: &Variant, kinds: (Operand, Operand)) -> (Variant, Variant) {
    let both_static = kinds.0 != Operand::Runtime && kinds.1 != Operand::Runtime;
    let l_bool = matches!(lhs, Variant::Boolean(_));
    let r_bool = matches!(rhs, Variant::Boolean(_));
    (
        if r_bool && both_static {
            lhs.clone()
        } else {
            logical_operand(lhs)
        },
        if l_bool && both_static {
            rhs.clone()
        } else {
            logical_operand(rhs)
        },
    )
}

pub fn bankers_round(v: f64) -> f64 {
    let floor = v.floor();
    let diff = v - floor;
    if diff > 0.5 {
        floor + 1.0
    } else if diff < 0.5 {
        floor
    } else {
        if (floor / 2.0).fract() == 0.0 {
            floor
        } else {
            floor + 1.0
        }
    }
}

pub fn parse_vba_number(s: &str) -> VResult<f64> {
    let t = s.trim();
    if t.is_empty() {
        return Err(VbaError::type_mismatch());
    }
    if let Some(hex) = t.strip_prefix("&H").or_else(|| t.strip_prefix("&h")) {
        return i64::from_str_radix(hex, 16)
            .map(|v| v as f64)
            .map_err(|_| VbaError::type_mismatch());
    }
    if let Some(oct) = t.strip_prefix("&O").or_else(|| t.strip_prefix("&o")) {
        return i64::from_str_radix(oct, 8)
            .map(|v| v as f64)
            .map_err(|_| VbaError::type_mismatch());
    }
    let normalised = t.replace(['d', 'D'], "e");
    let negated_by_suffix = normalised.ends_with('-');
    let body = match normalised.strip_suffix(['-', '+']) {
        Some(rest) => {
            let rest = rest.trim_end();
            if rest.starts_with(['-', '+']) {
                return Err(VbaError::type_mismatch());
            }
            rest
        }
        None => normalised.as_str(),
    };
    let value = body
        .parse::<f64>()
        .map(|v| if negated_by_suffix { -v } else { v })
        .map_err(|_| VbaError::type_mismatch())?;
    if !value.is_finite() {
        return Err(VbaError::overflow());
    }
    Ok(value)
}

pub fn numeric_prefix(s: &str) -> Option<f64> {
    let t: String = s.chars().filter(|c| !c.is_whitespace()).collect();
    let mut best = None;
    for (i, _) in t.char_indices() {
        if let Ok(v) = t[..=i].parse::<f64>()
            && v.is_finite()
        {
            best = Some(v);
        }
    }
    best
}

pub fn format_number(v: f64) -> String {
    if v == 0.0 {
        return if v.is_sign_negative() { "-0" } else { "0" }.to_string();
    }
    if v.is_nan() {
        return "NaN".to_string();
    }
    if v.is_infinite() {
        return if v > 0.0 { "INF" } else { "-INF" }.to_string();
    }

    let sci = format!("{v:.14e}");
    let (mantissa, exp_text) = sci.split_once('e').unwrap_or((sci.as_str(), "0"));
    let exp: i32 = exp_text.parse().unwrap_or(0);

    if (-4..16).contains(&exp) {
        let decimals = (14 - exp).max(0) as usize;
        let mut r = format!("{v:.decimals$}");
        if r.contains('.') {
            r = r.trim_end_matches('0').trim_end_matches('.').to_string();
        }
        let sig_digits = r
            .chars()
            .filter(|c| c.is_ascii_digit())
            .skip_while(|c| *c == '0')
            .count();
        if exp == -2 && sig_digits >= 15 || exp < -2 && sig_digits > 15 {
        } else {
            return if r == "-0" { "-0".to_string() } else { r };
        }
    }

    let mantissa = if mantissa.contains('.') {
        mantissa.trim_end_matches('0').trim_end_matches('.')
    } else {
        mantissa
    };
    let sign = if exp < 0 { "-" } else { "+" };
    format!("{mantissa}E{sign}{:02}", exp.abs())
}

fn format_currency(scaled: i64) -> String {
    let whole = scaled / 10_000;
    let frac = (scaled % 10_000).abs();
    if frac == 0 {
        whole.to_string()
    } else {
        let frac_str = format!("{frac:04}");
        let frac_str = frac_str.trim_end_matches('0');
        let sign = if scaled < 0 && whole == 0 { "-" } else { "" };
        format!("{sign}{whole}.{frac_str}")
    }
}

pub fn add(lhs: &Variant, rhs: &Variant, mode: ArithMode) -> VResult<Variant> {
    if lhs.is_null() || rhs.is_null() {
        return Ok(Variant::Null);
    }
    if let (Variant::Str(a), Variant::Str(b)) = (lhs, rhs) {
        return Ok(Variant::Str(format!("{a}{b}")));
    }
    if let (Variant::Empty, Variant::Str(s)) | (Variant::Str(s), Variant::Empty) = (lhs, rhs) {
        return Ok(Variant::Str(s.clone()));
    }
    keep_date(lhs, rhs, arith(lhs, rhs, mode, |a, b| a + b)?)
}

pub fn sub(lhs: &Variant, rhs: &Variant, mode: ArithMode) -> VResult<Variant> {
    keep_date(lhs, rhs, arith(lhs, rhs, mode, |a, b| a - b)?)
}

fn keep_date(lhs: &Variant, rhs: &Variant, result: Variant) -> VResult<Variant> {
    let one_date = matches!(lhs, Variant::Date(_)) != matches!(rhs, Variant::Date(_));
    if !one_date {
        return Ok(result);
    }
    Ok(match result {
        Variant::Integer(_)
        | Variant::Long(_)
        | Variant::Single(_)
        | Variant::Double(_)
        | Variant::Currency(_) => Variant::Date(result.to_f64()?),
        other => other,
    })
}

pub fn format_vba_date(serial: f64) -> String {
    let days = serial.floor();
    let frac = serial - days;
    let date_part = if days == 0.0 {
        None
    } else {
        let d = crate::core::date::excel_serial_to_date(days);
        Some(format!(
            "{}/{}/{:02}",
            d.month,
            d.day,
            d.year.rem_euclid(100)
        ))
    };
    let total_seconds = (frac * 86_400.0).round() as i64;
    let time_part = if total_seconds == 0 && date_part.is_some() {
        None
    } else {
        let (h24, m, s) = (
            total_seconds / 3600,
            (total_seconds / 60) % 60,
            total_seconds % 60,
        );
        let (h12, ampm) = match h24 {
            0 => (12, "AM"),
            1..=11 => (h24, "AM"),
            12 => (12, "PM"),
            _ => (h24 - 12, "PM"),
        };
        Some(format!("{h12}:{m:02}:{s:02} {ampm}"))
    };
    match (date_part, time_part) {
        (Some(d), Some(t)) => format!("{d} {t}"),
        (Some(d), None) => d,
        (None, Some(t)) => t,
        (None, None) => "12:00:00 AM".to_string(),
    }
}

pub fn mul(lhs: &Variant, rhs: &Variant, mode: ArithMode) -> VResult<Variant> {
    arith(lhs, rhs, mode, |a, b| a * b)
}

pub fn div(lhs: &Variant, rhs: &Variant) -> VResult<Variant> {
    if lhs.is_null() || rhs.is_null() {
        if !lhs.is_null() {
            lhs.to_f64()?;
        }
        if !rhs.is_null() {
            rhs.to_f64()?;
        }
        return Ok(Variant::Null);
    }
    let a = lhs.to_f64()?;
    let b = rhs.to_f64()?;
    if b == 0.0 {
        return Err(if a == 0.0 {
            VbaError::overflow()
        } else {
            VbaError::div_by_zero()
        });
    }
    let r = a / b;
    if !r.is_finite() || !a.is_finite() || !b.is_finite() {
        return Err(VbaError::overflow());
    }
    Ok(Variant::Double(r))
}

pub fn int_div(lhs: &Variant, rhs: &Variant) -> VResult<Variant> {
    let (a, b, class) = int_operands(lhs, rhs)?;
    let Some((a, b, class)) = zip3(a, b, class) else {
        return Ok(Variant::Null);
    };
    if b == 0 {
        return Err(VbaError::div_by_zero());
    }
    Variant::pack((a / b) as f64, class)
}

pub fn modulo(lhs: &Variant, rhs: &Variant) -> VResult<Variant> {
    let (a, b, class) = int_operands(lhs, rhs)?;
    let Some((a, b, class)) = zip3(a, b, class) else {
        return Ok(Variant::Null);
    };
    if b == 0 {
        return Err(VbaError::div_by_zero());
    }
    Variant::pack((a % b) as f64, class)
}

fn zip3(a: Option<i64>, b: Option<i64>, class: NumClass) -> Option<(i64, i64, NumClass)> {
    Some((a?, b?, class))
}

fn int_operands(lhs: &Variant, rhs: &Variant) -> VResult<(Option<i64>, Option<i64>, NumClass)> {
    fn one(v: &Variant) -> VResult<Option<i64>> {
        if v.is_null() {
            return Ok(None);
        }
        let r = bankers_round(v.to_f64()?);
        if !r.is_finite() || r < i32::MIN as f64 || r > i32::MAX as f64 {
            return Err(VbaError::overflow());
        }
        Ok(Some(r as i64))
    }

    let (l, r) = logical_pair(lhs, rhs, (Operand::Runtime, Operand::Runtime));

    let a = one(&l)?;
    let b = one(&r)?;
    if a.is_none() || b.is_none() {
        return Ok((None, None, NumClass::Long));
    }

    let class = match Variant::arith_type(&l, &r)? {
        NumClass::Integer => NumClass::Integer,
        _ => NumClass::Long,
    };
    Ok((a, b, class))
}

pub fn pow(lhs: &Variant, rhs: &Variant, mode: ArithMode) -> VResult<Variant> {
    if lhs.is_null() || rhs.is_null() {
        if !lhs.is_null() {
            lhs.to_f64()?;
        }
        if !rhs.is_null() {
            rhs.to_f64()?;
        }
        return Ok(Variant::Null);
    }
    let base = lhs.to_f64()?;
    let exp = rhs.to_f64()?;
    if base < 0.0 && exp.fract() != 0.0 {
        return Err(VbaError::invalid_call());
    }
    if base == 0.0 && exp < 0.0 {
        return Err(VbaError::invalid_call());
    }
    let r = base.powf(exp);
    let _ = mode;
    if !r.is_finite() && base.is_finite() && exp.is_finite() {
        return Err(VbaError::overflow());
    }
    Ok(Variant::Double(r))
}

pub fn concat(lhs: &Variant, rhs: &Variant) -> VResult<Variant> {
    if lhs.is_null() && rhs.is_null() {
        return Ok(Variant::Null);
    }
    if matches!(lhs, Variant::ErrValue(_)) || matches!(rhs, Variant::ErrValue(_)) {
        return Err(VbaError::type_mismatch());
    }
    let a = if lhs.is_null() {
        String::new()
    } else {
        lhs.to_vba_string()?
    };
    let b = if rhs.is_null() {
        String::new()
    } else {
        rhs.to_vba_string()?
    };
    Ok(Variant::Str(format!("{a}{b}")))
}

fn arith(
    lhs: &Variant,
    rhs: &Variant,
    mode: ArithMode,
    f: impl Fn(f64, f64) -> f64,
) -> VResult<Variant> {
    if lhs.is_null() || rhs.is_null() {
        if !lhs.is_null() {
            lhs.to_f64()?;
        }
        if !rhs.is_null() {
            rhs.to_f64()?;
        }
        return Ok(Variant::Null);
    }
    let class = Variant::arith_type(lhs, rhs)?;
    let (a, b) = (lhs.to_f64()?, rhs.to_f64()?);
    let r = f(a, b);
    if !r.is_finite() || !a.is_finite() || !b.is_finite() {
        return Err(VbaError::overflow());
    }
    Variant::pack_mode(r, class, mode)
}

pub fn neg(v: &Variant, mode: ArithMode) -> VResult<Variant> {
    if v.is_null() {
        return Ok(Variant::Null);
    }
    if mode == ArithMode::Constant
        && let Variant::Long(n) = v
        && *n == i32::MIN
    {
        return Ok(Variant::Long(i32::MIN));
    }
    let class = v.num_class().ok_or_else(VbaError::invalid_null)?;
    Variant::pack_mode(-v.to_f64()?, class, mode)
}

pub fn pos(v: &Variant, mode: ArithMode) -> VResult<Variant> {
    if v.is_null() {
        return Ok(Variant::Null);
    }
    let class = v.num_class().ok_or_else(VbaError::invalid_null)?;
    Variant::pack_mode(v.to_f64()?, class, mode)
}

pub fn not(v: &Variant) -> VResult<Variant> {
    let v = &logical_operand(v);
    match v {
        Variant::Null => Ok(Variant::Null),
        Variant::Boolean(b) => Ok(Variant::Boolean(!b)),
        other => {
            let n = bankers_round(other.to_f64()?);
            let class = match other.num_class().ok_or_else(VbaError::invalid_null)? {
                NumClass::Integer => NumClass::Integer,
                _ => NumClass::Long,
            };
            Variant::pack(!(n as i64) as f64, class)
        }
    }
}

pub fn logical(
    lhs: &Variant,
    rhs: &Variant,
    kinds: (Operand, Operand),
    f: impl Fn(i64, i64) -> i64,
) -> VResult<Variant> {
    let (l, r) = logical_pair(lhs, rhs, kinds);
    let (lhs, rhs) = (&l, &r);
    if let (Variant::Boolean(a), Variant::Boolean(b)) = (lhs, rhs) {
        let r = f(if *a { -1 } else { 0 }, if *b { -1 } else { 0 });
        return Ok(Variant::Boolean(r != 0));
    }
    let one = |v: &Variant| -> VResult<Option<f64>> {
        if v.is_null() {
            return Ok(None);
        }
        let r = bankers_round(v.to_f64()?);
        if !r.is_finite() || r < i32::MIN as f64 || r > i32::MAX as f64 {
            return Err(VbaError::overflow());
        }
        Ok(Some(r))
    };
    let (a, b) = (one(lhs)?, one(rhs)?);
    let (Some(a), Some(b)) = (a, b) else {
        return Ok(Variant::Null);
    };
    let class = match Variant::arith_type(lhs, rhs)? {
        NumClass::Integer => NumClass::Integer,
        _ => NumClass::Long,
    };
    Variant::pack(f(a as i64, b as i64) as f64, class)
}

pub fn and(lhs: &Variant, rhs: &Variant, kinds: (Operand, Operand)) -> VResult<Variant> {
    if let Some(v) = three_valued(lhs, rhs, kinds, false)? {
        return Ok(v);
    }
    logical(lhs, rhs, kinds, |x, y| x & y)
}

pub fn or(lhs: &Variant, rhs: &Variant, kinds: (Operand, Operand)) -> VResult<Variant> {
    if let Some(v) = three_valued(lhs, rhs, kinds, true)? {
        return Ok(v);
    }
    logical(lhs, rhs, kinds, |x, y| x | y)
}

pub fn imp(lhs: &Variant, rhs: &Variant, kinds: (Operand, Operand)) -> VResult<Variant> {
    or(&not(lhs)?, rhs, kinds)
}

fn three_valued(
    lhs: &Variant,
    rhs: &Variant,
    kinds: (Operand, Operand),
    deciding: bool,
) -> VResult<Option<Variant>> {
    let (known, kind) = match (lhs.is_null(), rhs.is_null()) {
        (true, true) => return Ok(Some(Variant::Null)),
        (true, false) => (rhs, kinds.1),
        (false, true) => (lhs, kinds.0),
        (false, false) => return Ok(None),
    };
    let folded;
    let known = if matches!(known, Variant::Str(_)) && kind == Operand::Runtime {
        folded = logical_operand(known);
        &folded
    } else {
        known
    };
    if matches!(known, Variant::Boolean(_)) {
        if known.to_bool()? != deciding {
            return Ok(Some(Variant::Null));
        }
        return Ok(Some(known.clone()));
    }
    let class = match known.num_class().ok_or_else(VbaError::invalid_null)? {
        NumClass::Integer => NumClass::Integer,
        _ => NumClass::Long,
    };
    let rounded = bankers_round(known.to_f64()?);
    if (rounded != 0.0) != deciding {
        return Ok(Some(Variant::Null));
    }
    Variant::pack(rounded, class).map(Some)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Operand {
    Literal,

    ConstExpr,

    Static,

    Runtime,
}

impl Operand {
    fn is_const(self) -> bool {
        matches!(self, Operand::Literal | Operand::ConstExpr)
    }
}

pub fn compare_ctx(
    lhs: &Variant,
    rhs: &Variant,
    lhs_kind: Operand,
    rhs_kind: Operand,
) -> VResult<Option<std::cmp::Ordering>> {
    use std::cmp::Ordering;
    if lhs.is_null() || rhs.is_null() {
        return Ok(None);
    }

    let both_stringy = matches!(
        (lhs, rhs),
        (Variant::Str(_), Variant::Str(_))
            | (Variant::Str(_), Variant::Empty)
            | (Variant::Empty, Variant::Str(_))
    );
    if both_stringy {
        return Ok(Some(lhs.to_vba_string()?.cmp(&rhs.to_vba_string()?)));
    }

    let numeric = |v: &Variant| -> VResult<f64> { v.to_f64() };

    match (lhs, rhs) {
        (Variant::Str(text), other) | (other, Variant::Str(text)) => {
            let str_on_left = matches!(lhs, Variant::Str(_));
            let (str_kind, num_kind) = if str_on_left {
                (lhs_kind, rhs_kind)
            } else {
                (rhs_kind, lhs_kind)
            };
            if matches!(other, Variant::Boolean(_)) {
                if num_kind != Operand::Runtime {
                    let as_bool =
                        bool_word(text).or_else(|| parse_vba_number(text).ok().map(|n| n != 0.0));
                    if let Some(a) = as_bool {
                        let ord = cmp_f64(bool_as_number(a), bool_as_number(other.to_bool()?));
                        return Ok(Some(if str_on_left { ord } else { ord.reverse() }));
                    }
                    if str_kind != Operand::Runtime {
                        return Err(VbaError::type_mismatch());
                    }
                    let ord = Ordering::Greater;
                    return Ok(Some(if str_on_left { ord } else { ord.reverse() }));
                }
                if str_kind != Operand::Runtime {
                    let ord = text.as_str().cmp(other.to_vba_string()?.as_str());
                    return Ok(Some(if str_on_left { ord } else { ord.reverse() }));
                }
            }

            let str_typed = str_kind.is_const() || str_kind == Operand::Static;

            let ord = if num_kind == Operand::Static {
                match parse_vba_number(text) {
                    Ok(a) => cmp_f64(a, numeric(other)?),
                    Err(e) if str_typed || e.number != 13 => return Err(e),
                    Err(_) => Ordering::Greater,
                }
            } else if num_kind.is_const() && str_typed {
                match numeric_prefix(text) {
                    Some(a) => cmp_f64(a, numeric(other)?),
                    None => return Err(VbaError::type_mismatch()),
                }
            } else if num_kind.is_const() {
                match parse_vba_number(text) {
                    Ok(a) => cmp_f64(a, numeric(other)?),
                    Err(e) if e.number != 13 => return Err(e),
                    Err(_) => Ordering::Greater,
                }
            } else if str_typed {
                text.as_str().cmp(other.to_vba_string()?.as_str())
            } else {
                Ordering::Greater
            };
            Ok(Some(if str_on_left { ord } else { ord.reverse() }))
        }
        _ => Ok(Some(cmp_f64(numeric(lhs)?, numeric(rhs)?))),
    }
}

fn bool_as_number(b: bool) -> f64 {
    if b { -1.0 } else { 0.0 }
}

fn cmp_f64(a: f64, b: f64) -> std::cmp::Ordering {
    a.partial_cmp(&b).unwrap_or(std::cmp::Ordering::Equal)
}

pub fn compare(lhs: &Variant, rhs: &Variant) -> VResult<Option<std::cmp::Ordering>> {
    compare_ctx(lhs, rhs, Operand::Runtime, Operand::Runtime)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn shows(v: &Variant) -> String {
        format!(
            "{}|{}",
            v.type_name(),
            v.to_vba_string().unwrap_or_default()
        )
    }

    #[test]
    fn literal_typing_matches_excel() {
        assert_eq!(shows(&Variant::from_literal(1.0, false)), "Integer|1");
        assert_eq!(
            shows(&Variant::from_literal(32767.0, false)),
            "Integer|32767"
        );
        assert_eq!(shows(&Variant::from_literal(32768.0, false)), "Long|32768");
        assert_eq!(
            shows(&Variant::from_literal(2147483647.0, false)),
            "Long|2147483647"
        );
        assert_eq!(
            shows(&Variant::from_literal(2147483648.0, false)),
            "Double|2147483648"
        );
        assert_eq!(shows(&Variant::from_literal(1.5, true)), "Double|1.5");
        assert_eq!(
            shows(&Variant::from_literal(100000.0, true)),
            "Double|100000"
        );
        assert_eq!(
            shows(&Variant::from_literal(100000.0, false)),
            "Long|100000"
        );
    }

    #[test]
    fn constant_arithmetic_overflows_but_runtime_arithmetic_promotes() {
        assert_eq!(
            shows(
                &add(
                    &Variant::Integer(32767),
                    &Variant::Integer(1),
                    ArithMode::Promote
                )
                .unwrap()
            ),
            "Long|32768"
        );
        assert_eq!(
            shows(
                &add(
                    &Variant::Long(2147483647),
                    &Variant::Integer(1),
                    ArithMode::Promote
                )
                .unwrap()
            ),
            "Double|2147483648"
        );
        assert_eq!(
            shows(
                &mul(
                    &Variant::Long(100000),
                    &Variant::Long(100000),
                    ArithMode::Promote
                )
                .unwrap()
            ),
            "Double|10000000000"
        );
        assert_eq!(
            shows(
                &add(
                    &Variant::Integer(1),
                    &Variant::Integer(1),
                    ArithMode::Constant
                )
                .unwrap()
            ),
            "Integer|2"
        );
        assert_eq!(
            add(
                &Variant::Integer(32767),
                &Variant::Integer(1),
                ArithMode::Constant
            )
            .unwrap_err()
            .number,
            6
        );
        assert_eq!(
            add(
                &Variant::Long(2147483647),
                &Variant::Integer(1),
                ArithMode::Constant
            )
            .unwrap_err()
            .number,
            6
        );
        assert_eq!(
            mul(
                &Variant::Long(100000),
                &Variant::Long(100000),
                ArithMode::Constant
            )
            .unwrap_err()
            .number,
            6
        );
    }

    #[test]
    fn plus_concatenates_only_when_both_sides_are_strings() {
        assert_eq!(
            shows(
                &add(
                    &Variant::Str("1".into()),
                    &Variant::Integer(1),
                    ArithMode::Constant
                )
                .unwrap()
            ),
            "Double|2"
        );
        assert_eq!(
            shows(
                &add(
                    &Variant::Str("1".into()),
                    &Variant::Str("2".into()),
                    ArithMode::Constant
                )
                .unwrap()
            ),
            "String|12"
        );
        assert_eq!(
            add(
                &Variant::Str("abc".into()),
                &Variant::Integer(1),
                ArithMode::Constant
            )
            .unwrap_err()
            .number,
            13
        );
        assert_eq!(
            shows(
                &add(
                    &Variant::Str("  3  ".into()),
                    &Variant::Integer(1),
                    ArithMode::Constant
                )
                .unwrap()
            ),
            "Double|4"
        );
    }

    #[test]
    fn division_is_always_double() {
        assert_eq!(
            shows(&div(&Variant::Integer(1), &Variant::Integer(2)).unwrap()),
            "Double|0.5"
        );
        assert_eq!(
            shows(&div(&Variant::Integer(4), &Variant::Integer(2)).unwrap()),
            "Double|2"
        );
        assert_eq!(
            div(&Variant::Integer(1), &Variant::Integer(0))
                .unwrap_err()
                .number,
            11
        );
    }

    #[test]
    fn int_div_and_mod_round_their_operands_first() {
        assert_eq!(
            shows(&int_div(&Variant::Integer(7), &Variant::Integer(2)).unwrap()),
            "Integer|3"
        );
        assert_eq!(
            shows(&int_div(&Variant::Integer(-7), &Variant::Integer(2)).unwrap()),
            "Integer|-3"
        );
        assert_eq!(
            shows(&int_div(&Variant::Double(7.6), &Variant::Integer(2)).unwrap()),
            "Long|4"
        );
        assert_eq!(
            shows(&modulo(&Variant::Integer(7), &Variant::Integer(2)).unwrap()),
            "Integer|1"
        );
        assert_eq!(
            shows(&modulo(&Variant::Integer(-7), &Variant::Integer(2)).unwrap()),
            "Integer|-1"
        );
        assert_eq!(
            shows(&modulo(&Variant::Double(7.6), &Variant::Integer(2)).unwrap()),
            "Long|0"
        );
    }

    #[test]
    fn pow_is_always_double() {
        assert_eq!(
            shows(
                &pow(
                    &Variant::Integer(2),
                    &Variant::Integer(2),
                    ArithMode::Constant
                )
                .unwrap()
            ),
            "Double|4"
        );
    }

    #[test]
    fn runtime_pow_overflow_raises() {
        assert_eq!(
            pow(
                &Variant::Double(3.0),
                &Variant::Long(32767),
                ArithMode::Promote
            )
            .unwrap_err()
            .number,
            6
        );
    }

    #[test]
    fn booleans_are_minus_one() {
        assert_eq!(
            shows(
                &add(
                    &Variant::Boolean(true),
                    &Variant::Integer(1),
                    ArithMode::Constant
                )
                .unwrap()
            ),
            "Integer|0"
        );
        assert_eq!(
            shows(
                &add(
                    &Variant::Boolean(true),
                    &Variant::Boolean(true),
                    ArithMode::Constant
                )
                .unwrap()
            ),
            "Integer|-2"
        );
    }

    #[test]
    fn logical_operators_are_bitwise_unless_both_sides_are_boolean() {
        assert_eq!(
            shows(
                &logical(
                    &Variant::Boolean(true),
                    &Variant::Boolean(false),
                    (Operand::Literal, Operand::Literal),
                    |a, b| a & b
                )
                .unwrap()
            ),
            "Boolean|False"
        );
        assert_eq!(
            shows(
                &logical(
                    &Variant::Integer(5),
                    &Variant::Integer(3),
                    (Operand::Literal, Operand::Literal),
                    |a, b| a & b
                )
                .unwrap()
            ),
            "Integer|1"
        );
        assert_eq!(shows(&not(&Variant::Integer(5)).unwrap()), "Integer|-6");
        assert_eq!(
            shows(&not(&Variant::Boolean(true)).unwrap()),
            "Boolean|False"
        );
    }

    #[test]
    fn null_and_uses_the_bitwise_rounded_truthiness_of_the_known_operand() {
        assert_eq!(
            shows(
                &and(
                    &Variant::Null,
                    &Variant::Double(0.0001),
                    (Operand::Runtime, Operand::Runtime)
                )
                .unwrap()
            ),
            "Long|0"
        );
    }

    #[test]
    fn null_propagates_through_arithmetic_but_not_concatenation() {
        assert!(
            add(&Variant::Null, &Variant::Integer(1), ArithMode::Constant)
                .unwrap()
                .is_null()
        );
        assert_eq!(
            shows(&concat(&Variant::Null, &Variant::Str("a".into())).unwrap()),
            "String|a"
        );
    }

    #[test]
    fn empty_is_both_zero_and_the_empty_string() {
        assert_eq!(
            shows(&add(&Variant::Empty, &Variant::Integer(1), ArithMode::Constant).unwrap()),
            "Integer|1"
        );
        assert_eq!(
            shows(&concat(&Variant::Empty, &Variant::Str("a".into())).unwrap()),
            "String|a"
        );
        assert_eq!(
            compare(&Variant::Empty, &Variant::Integer(0)).unwrap(),
            Some(std::cmp::Ordering::Equal)
        );
        assert_eq!(
            compare(&Variant::Empty, &Variant::Str(String::new())).unwrap(),
            Some(std::cmp::Ordering::Equal)
        );
        for (l, r) in [
            (Variant::Empty, Variant::Str("a".into())),
            (Variant::Str("a".into()), Variant::Empty),
        ] {
            assert_eq!(
                shows(&add(&l, &r, ArithMode::Constant).unwrap()),
                "String|a"
            );
        }
        assert_eq!(
            shows(
                &add(
                    &Variant::Empty,
                    &Variant::Str("1".into()),
                    ArithMode::Constant
                )
                .unwrap()
            ),
            "String|1"
        );
        assert_eq!(
            shows(
                &add(
                    &Variant::Empty,
                    &Variant::Str(String::new()),
                    ArithMode::Constant
                )
                .unwrap()
            ),
            "String|"
        );
        assert_eq!(
            shows(&add(&Variant::Empty, &Variant::Empty, ArithMode::Constant).unwrap()),
            "Integer|0"
        );
        assert!(
            sub(
                &Variant::Empty,
                &Variant::Str("a".into()),
                ArithMode::Constant
            )
            .is_err()
        );
        assert!(
            mul(
                &Variant::Empty,
                &Variant::Str("a".into()),
                ArithMode::Constant
            )
            .is_err()
        );
    }

    #[test]
    fn rounding_is_bankers_not_half_away_from_zero() {
        assert_eq!(bankers_round(0.5), 0.0);
        assert_eq!(bankers_round(1.5), 2.0);
        assert_eq!(bankers_round(2.5), 2.0);
        assert_eq!(bankers_round(-0.5), -0.0);
        assert_eq!(bankers_round(-1.5), -2.0);
        assert_eq!(bankers_round(3.5), 4.0);
    }

    #[test]
    fn concat_stringifies_numbers() {
        assert_eq!(
            shows(&concat(&Variant::Integer(1), &Variant::Integer(2)).unwrap()),
            "String|12"
        );
    }

    #[test]
    fn negative_zero_keeps_its_sign() {
        assert_eq!(format_number(-0.0), "-0");
        assert_eq!(format_number(0.0), "0");
        assert_eq!(format_number(-0.4 + 0.4), "0");
        assert_eq!(
            shows(&neg(&Variant::Double(0.0), ArithMode::Constant).unwrap()),
            "Double|-0"
        );
    }

    #[test]
    fn cstr_shows_exactly_fifteen_significant_digits() {
        assert_eq!(format_number(9.156321295314252e-5), "9.15632129531425E-05");
        assert_eq!(
            format_number(-0.000985221674876847),
            "-0.000985221674876847"
        );
        assert_eq!(format_number(0.000457247370827618), "0.000457247370827618");
        assert_eq!(format_number(0.0001), "0.0001");
        assert_eq!(format_number(0.00001), "1E-05");
        assert_eq!(format_number(1e15), "1000000000000000");
        assert_eq!(format_number(1e16), "1E+16");
    }

    #[test]
    fn number_formatting_matches_cstr() {
        assert_eq!(format_number(1.5), "1.5");
        assert_eq!(format_number(2.0), "2");
        assert_eq!(format_number(0.0), "0");
        assert_eq!(format_number(-1.25), "-1.25");
        assert_eq!(format_number(1000.0), "1000");
        assert_eq!(format_number(0.1), "0.1");
        assert_eq!(format_number(1.0 / 15.0), "6.66666666666667E-02");
    }

    #[test]
    fn string_to_number_accepts_what_vba_accepts() {
        assert_eq!(parse_vba_number("1e3").unwrap(), 1000.0);
        assert_eq!(parse_vba_number("  3  ").unwrap(), 3.0);
        assert_eq!(parse_vba_number("1.5").unwrap(), 1.5);
        assert_eq!(parse_vba_number("&HFF").unwrap(), 255.0);
        assert_eq!(parse_vba_number("").unwrap_err().number, 13);
        assert_eq!(parse_vba_number("abc").unwrap_err().number, 13);
        assert_eq!(parse_vba_number("1-").unwrap(), -1.0);
        assert_eq!(parse_vba_number("1+").unwrap(), 1.0);
        assert_eq!(parse_vba_number("2.5-").unwrap(), -2.5);
        assert_eq!(parse_vba_number("1 -").unwrap(), -1.0);
        assert_eq!(parse_vba_number("1E2-").unwrap(), -100.0);
        assert_eq!(parse_vba_number("-1-").unwrap_err().number, 13);
        assert_eq!(parse_vba_number("1--").unwrap_err().number, 13);
        assert_eq!(parse_vba_number("-").unwrap_err().number, 13);
    }
}
