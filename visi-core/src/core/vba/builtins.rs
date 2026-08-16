//! VBA's intrinsic functions, for the Phase 1 interpreter.
//!
//! Only the host-free ones: conversion, maths, strings, and type inspection.
//! Anything that reaches a workbook (`Application.WorksheetFunction`,
//! `Range`) is Phase 2, and anything non-deterministic (`Now`, `Rnd`,
//! `Timer`) is deliberately absent — a differential fuzzer cannot compare a
//! value that differs between the two runs by construction, so offering them
//! would only produce noise.
//!
//! [`call`] returns `Ok(None)` for a name it does not know, so the caller can
//! go on to report "Sub or Function not defined" itself rather than having
//! two places that decide what exists.

use super::value::{self, VResult, Variant, VbaError};

/// Intrinsics that inspect a `Null` rather than propagating or rejecting it.
const HANDLES_NULL: &[&str] = &[
    "isnull",
    "isempty",
    "isnumeric",
    "isdate",
    "isobject",
    "iserror",
    "isarray",
    "typename",
    "vartype",
    "iif",
];

/// Intrinsics that raise error 94 on a `Null` argument.
///
/// Measured, not derived, because no principle is visible behind the split:
/// `Hex` and `Oct` propagate a `Null` while `Chr` and `Asc` reject it;
/// `String` propagates while `Space` rejects; `Trim` propagates while
/// `StrReverse` rejects. Everything not listed here (and not in
/// [`HANDLES_NULL`]) propagates -- including `CVar`, the one `C*` conversion
/// that does.
const REJECTS_NULL: &[&str] = &[
    "cstr",
    "cint",
    "clng",
    "cdbl",
    "csng",
    "cbool",
    "ccur",
    "val",
    "sgn",
    "sqr",
    "exp",
    "log",
    "sin",
    "cos",
    "tan",
    "atn",
    "space",
    "strreverse",
    "chr",
    "chrw",
    "asc",
    "ascw",
    "replace",
];

/// One argument, or `Empty` when it was omitted.
fn arg(args: &[Variant], i: usize) -> Variant {
    args.get(i).cloned().unwrap_or(Variant::Empty)
}

fn need(args: &[Variant], i: usize) -> VResult<Variant> {
    args.get(i).cloned().ok_or_else(VbaError::invalid_call)
}

/// `Null` in, `Null` out — the rule almost every intrinsic follows.
fn any_null(args: &[Variant]) -> bool {
    args.iter().any(|a| a.is_null())
}

/// Calls an intrinsic by name, or returns `None` if there is no such name.
pub fn call(name: &str, args: &[Variant]) -> VResult<Option<Variant>> {
    let lower = name.to_ascii_lowercase();

    // How an intrinsic treats a `Null` argument, from a sweep of all 46 of
    // them against real Excel (see docs/vba-error-ordering.md).
    if any_null(args) && !HANDLES_NULL.contains(&lower.as_str()) {
        if REJECTS_NULL.contains(&lower.as_str()) {
            return Err(VbaError::invalid_null());
        }
        return Ok(Some(Variant::Null));
    }

    let v = match lower.as_str() {
        // ---- type inspection --------------------------------------------
        "typename" => Variant::Str(arg(args, 0).type_name().to_string()),
        "vartype" => Variant::Integer(vartype(&arg(args, 0))),
        "isnull" => Variant::Boolean(arg(args, 0).is_null()),
        "isempty" => Variant::Boolean(arg(args, 0).is_empty()),
        "isnumeric" => Variant::Boolean(is_numeric(&arg(args, 0))),
        "isdate" => Variant::Boolean(matches!(arg(args, 0), Variant::Date(_))),
        "isobject" => Variant::Boolean(matches!(arg(args, 0), Variant::Object(_))),
        // Measured: `IsError(CVErr(2042))` and `IsError` of a cell holding
        // `=1/0` are both True. This is the whole point of the error Variant
        // -- `Application.VLookup` returning one is how a macro tests for a
        // failed lookup without trapping a run-time error.
        "iserror" => Variant::Boolean(matches!(arg(args, 0), Variant::ErrValue(_))),
        "isarray" => Variant::Boolean(matches!(arg(args, 0), Variant::Array(_))),
        "cverr" => Variant::ErrValue(need(args, 0)?.to_f64()? as i32),

        // ---- arrays ------------------------------------------------------
        // Only a `Range.Value` read produces one of these, so the bounds are
        // always 1-based and two-dimensional.
        "ubound" | "lbound" => {
            let Variant::Array(a) = need(args, 0)? else {
                return Err(VbaError::type_mismatch());
            };
            let dim = match args.get(1) {
                Some(d) => d.to_f64()? as usize,
                None => 1,
            };
            if lower == "lbound" {
                a.ubound(dim)?;
                Variant::Long(1)
            } else {
                Variant::Long(a.ubound(dim)? as i32)
            }
        }

        // ---- conversion --------------------------------------------------
        "cstr" => Variant::Str(need(args, 0)?.to_vba_string()?),
        "cint" => pack_int(numeric_arg(args, 0)?)?,
        "clng" => pack_long(numeric_arg(args, 0)?)?,
        "cdbl" => Variant::Double(numeric_arg(args, 0)?),
        "csng" => {
            let f = need(args, 0)?.to_f64()? as f32;
            if !f.is_finite() {
                return Err(VbaError::overflow());
            }
            Variant::Single(f)
        }
        "cbool" => Variant::Boolean(need(args, 0)?.to_bool()?),
        "ccur" => {
            let scaled = value::bankers_round(need(args, 0)?.to_f64()? * 10_000.0);
            if !scaled.is_finite() || scaled.abs() > i64::MAX as f64 {
                return Err(VbaError::overflow());
            }
            Variant::Currency(scaled as i64)
        }
        "cvar" => need(args, 0)?,
        // Always a Double, whatever the argument: `Val(1%)`, `Val(255)` and
        // `Val("100000")` are all Doubles. An earlier version typed the
        // result like a literal, inferred from a fuzz case where `Val` may
        // never have run -- measuring it directly settled it.
        "val" => Variant::Double(val_of(&arg(args, 0))),

        // ---- maths -------------------------------------------------------
        // Abs and Sgn keep the argument's own numeric width, which is
        // observable through TypeName.
        "abs" => same_width(&need(args, 0)?, f64::abs)?,
        "sgn" => Variant::Integer(sgn(need(args, 0)?.to_f64()?)),
        // Int floors, Fix truncates: Int(-1.5) is -2 and Fix(-1.5) is -1.
        "int" => same_width(&need(args, 0)?, f64::floor)?,
        "fix" => same_width(&need(args, 0)?, f64::trunc)?,
        "sqr" => {
            let x = need(args, 0)?.to_f64()?;
            if x < 0.0 {
                return Err(VbaError::invalid_call());
            }
            Variant::Double(x.sqrt())
        }
        "exp" => Variant::Double(need(args, 0)?.to_f64()?.exp()),
        "log" => {
            let x = need(args, 0)?.to_f64()?;
            if x <= 0.0 {
                return Err(VbaError::invalid_call());
            }
            Variant::Double(x.ln())
        }
        "sin" => Variant::Double(need(args, 0)?.to_f64()?.sin()),
        "cos" => Variant::Double(need(args, 0)?.to_f64()?.cos()),
        "tan" => Variant::Double(need(args, 0)?.to_f64()?.tan()),
        "atn" => Variant::Double(need(args, 0)?.to_f64()?.atan()),
        // VBA's Round is banker's, matching every other conversion.
        "round" => {
            let x = need(args, 0)?.to_f64()?;
            let places = match args.get(1) {
                Some(v) => v.to_f64()? as i32,
                None => 0,
            };
            let factor = 10f64.powi(places);
            Variant::Double(value::bankers_round(x * factor) / factor)
        }

        // ---- strings -----------------------------------------------------
        "len" => Variant::Long(chars(&need(args, 0)?)?.len() as i32),
        "left" => {
            let s = chars(&need(args, 0)?)?;
            let n = count_arg(args, 1)?;
            Variant::Str(s.iter().take(n).collect())
        }
        "right" => {
            let s = chars(&need(args, 0)?)?;
            let n = count_arg(args, 1)?;
            let skip = s.len().saturating_sub(n);
            Variant::Str(s.iter().skip(skip).collect())
        }
        "mid" => {
            let s = chars(&need(args, 0)?)?;
            // Mid is 1-based, and start < 1 is an error rather than a clamp.
            let start = need(args, 1)?.to_f64()?;
            if start < 1.0 {
                return Err(VbaError::invalid_call());
            }
            let start = start as usize - 1;
            let taken: Vec<char> = s.iter().skip(start).copied().collect();
            match args.get(2) {
                Some(n) => {
                    let n = to_count(n.to_f64()?)?;
                    Variant::Str(taken.into_iter().take(n).collect())
                }
                None => Variant::Str(taken.into_iter().collect()),
            }
        }
        "instr" => instr(args)?,
        "ucase" => Variant::Str(need(args, 0)?.to_vba_string()?.to_uppercase()),
        "lcase" => Variant::Str(need(args, 0)?.to_vba_string()?.to_lowercase()),
        "trim" => Variant::Str(need(args, 0)?.to_vba_string()?.trim().to_string()),
        "ltrim" => Variant::Str(need(args, 0)?.to_vba_string()?.trim_start().to_string()),
        "rtrim" => Variant::Str(need(args, 0)?.to_vba_string()?.trim_end().to_string()),
        "space" => Variant::Str(" ".repeat(count_arg(args, 0)?)),
        "string" => {
            let n = count_arg(args, 0)?;
            let c = match need(args, 1)? {
                Variant::Str(s) => s.chars().next().unwrap_or(' '),
                other => char_from_code(other.to_f64()?)?,
            };
            Variant::Str(std::iter::repeat_n(c, n).collect())
        }
        "chr" | "chrw" => Variant::Str(char_from_code(need(args, 0)?.to_f64()?)?.to_string()),
        "asc" | "ascw" => {
            let s = need(args, 0)?.to_vba_string()?;
            match s.chars().next() {
                Some(c) => Variant::Integer(c as u32 as i16),
                None => return Err(VbaError::invalid_call()),
            }
        }
        "strreverse" => Variant::Str(need(args, 0)?.to_vba_string()?.chars().rev().collect()),
        "replace" => {
            let s = need(args, 0)?.to_vba_string()?;
            let find = need(args, 1)?.to_vba_string()?;
            let with = need(args, 2)?.to_vba_string()?;
            if find.is_empty() {
                Variant::Str(s)
            } else {
                Variant::Str(s.replace(&find, &with))
            }
        }
        "strcomp" => {
            let a = need(args, 0)?.to_vba_string()?;
            let b = need(args, 1)?.to_vba_string()?;
            Variant::Integer(match a.cmp(&b) {
                std::cmp::Ordering::Less => -1,
                std::cmp::Ordering::Equal => 0,
                std::cmp::Ordering::Greater => 1,
            })
        }
        // `RGB` composes the BGR `Long` that `Interior.Color` and
        // `Font.Color` take -- the low byte is red, which is why `&HFF0000`
        // is blue. See `vba::color`. Measured: a component above 255 clamps
        // rather than carrying into the next byte (`RGB(300, 0, 0)` is 255),
        // and a negative one is error 5 rather than clamping to zero.
        "rgb" => {
            let part = |i: usize| -> VResult<i64> {
                let v = to_i64(need(args, i)?.to_f64()?)?;
                if v < 0 {
                    return Err(VbaError::new(5, "Invalid procedure call or argument"));
                }
                Ok(v)
            };
            let (r, g, b) = (part(0)?, part(1)?, part(2)?);
            Variant::Long(crate::core::vba::color::rgb(r, g, b))
        }
        "hex" => Variant::Str(format!("{:X}", to_i64(need(args, 0)?.to_f64()?)?)),
        "oct" => Variant::Str(format!("{:o}", to_i64(need(args, 0)?.to_f64()?)?)),

        // ---- logic -------------------------------------------------------
        // IIf evaluates both arms, as VBA does -- it is a function call, not
        // a conditional expression, which is a classic VBA footgun.
        "iif" => {
            let cond = need(args, 0)?;
            if !cond.is_null() && cond.to_bool()? {
                arg(args, 1)
            } else {
                arg(args, 2)
            }
        }

        _ => return Ok(None),
    };
    Ok(Some(v))
}

fn vartype(v: &Variant) -> i16 {
    match v {
        Variant::Empty => 0,
        Variant::Null => 1,
        Variant::Integer(_) => 2,
        Variant::Long(_) => 3,
        Variant::Single(_) => 4,
        Variant::Double(_) => 5,
        Variant::Currency(_) => 6,
        Variant::Date(_) => 7,
        Variant::Str(_) => 8,
        Variant::Boolean(_) => 11,
        // The documented VarType constants: vbObject, vbError, and an array
        // is vbArray (8192) added to its element type, which for a range read
        // is always vbVariant (12).
        Variant::Object(_) => 9,
        Variant::ErrValue(_) => 10,
        Variant::Array(_) => 8192 + 12,
    }
}

/// A numeric argument for the explicit conversions, which unlike arithmetic
/// accept an error value and give its `CVErr` number back.
///
/// Measured: `v = Application.VLookup(...)` failing makes `CLng(v)` `2042`,
/// while `v + 1` is error 13. Only this path may look through an error value.
fn numeric_arg(args: &[Variant], i: usize) -> VResult<f64> {
    let v = need(args, i)?;
    match v.error_number() {
        Some(n) => Ok(n as f64),
        None => v.to_f64(),
    }
}

fn is_numeric(v: &Variant) -> bool {
    match v {
        Variant::Str(s) => value::parse_vba_number(s).is_ok() && !s.trim().is_empty(),
        Variant::Empty | Variant::Null => false,
        Variant::ErrValue(_) | Variant::Object(_) | Variant::Array(_) => false,
        _ => true,
    }
}

/// `Val` stops at the first character that cannot continue a number, and
/// returns 0 rather than erroring — unlike implicit coercion.
fn val_of(v: &Variant) -> f64 {
    let Ok(s) = v.to_vba_string() else {
        return 0.0;
    };
    let t: String = s.chars().filter(|c| !c.is_whitespace()).collect();
    let mut end = 0;
    for (i, _) in t.char_indices() {
        if t[..=i].parse::<f64>().is_ok() {
            end = i + 1;
        }
    }
    t[..end].parse::<f64>().unwrap_or(0.0)
}

fn sgn(v: f64) -> i16 {
    if v > 0.0 {
        1
    } else if v < 0.0 {
        -1
    } else {
        0
    }
}

/// Applies a numeric function while keeping the argument's width, which is
/// what makes `Int(-1.5)` a `Double` and `Abs(-1%)` an `Integer`.
fn same_width(v: &Variant, f: impl Fn(f64) -> f64) -> VResult<Variant> {
    let r = f(v.to_f64()?);
    Ok(match v {
        Variant::Integer(_) | Variant::Boolean(_) | Variant::Empty => {
            if r < i16::MIN as f64 || r > i16::MAX as f64 {
                return Err(VbaError::overflow());
            }
            Variant::Integer(r as i16)
        }
        Variant::Long(_) => {
            if r < i32::MIN as f64 || r > i32::MAX as f64 {
                return Err(VbaError::overflow());
            }
            Variant::Long(r as i32)
        }
        Variant::Single(_) => Variant::Single(r as f32),
        Variant::Currency(_) => Variant::Currency((r * 10_000.0).round() as i64),
        _ => Variant::Double(r),
    })
}

fn pack_int(v: f64) -> VResult<Variant> {
    let r = value::bankers_round(v);
    if r < i16::MIN as f64 || r > i16::MAX as f64 {
        return Err(VbaError::overflow());
    }
    Ok(Variant::Integer(r as i16))
}

fn pack_long(v: f64) -> VResult<Variant> {
    let r = value::bankers_round(v);
    if r < i32::MIN as f64 || r > i32::MAX as f64 {
        return Err(VbaError::overflow());
    }
    Ok(Variant::Long(r as i32))
}

fn to_i64(v: f64) -> VResult<i64> {
    let r = value::bankers_round(v);
    if !r.is_finite() || r.abs() > i64::MAX as f64 {
        return Err(VbaError::overflow());
    }
    Ok(r as i64)
}

/// A count argument, which VBA rounds rather than truncates.
///
/// Measured: `Space(2.6)` is three spaces, not two.
fn to_count(v: f64) -> VResult<usize> {
    let r = value::bankers_round(v);
    if r < 0.0 {
        return Err(VbaError::invalid_call());
    }
    Ok(r as usize)
}

fn count_arg(args: &[Variant], i: usize) -> VResult<usize> {
    to_count(need(args, i)?.to_f64()?)
}

fn chars(v: &Variant) -> VResult<Vec<char>> {
    Ok(v.to_vba_string()?.chars().collect())
}

fn char_from_code(code: f64) -> VResult<char> {
    let c = value::bankers_round(code);
    if !(0.0..=65535.0).contains(&c) {
        return Err(VbaError::invalid_call());
    }
    char::from_u32(c as u32).ok_or_else(VbaError::invalid_call)
}

/// `InStr` in both its arities: `InStr(haystack, needle)` and
/// `InStr(start, haystack, needle)`.
fn instr(args: &[Variant]) -> VResult<Variant> {
    let (start, hay, needle) = if args.len() >= 3 {
        let s = need(args, 0)?.to_f64()?;
        if s < 1.0 {
            return Err(VbaError::invalid_call());
        }
        (
            s as usize - 1,
            need(args, 1)?.to_vba_string()?,
            need(args, 2)?.to_vba_string()?,
        )
    } else {
        (
            0,
            need(args, 0)?.to_vba_string()?,
            need(args, 1)?.to_vba_string()?,
        )
    };
    let hay_chars: Vec<char> = hay.chars().collect();
    // A zero-length string to search in is 0, whatever the needle -- so
    // `InStr("", "")` is 0 while `InStr("a", "")` is 1. Measured; this used
    // to report 1 for the empty/empty pair, on the reasoning that an empty
    // needle matches at the start, which is true only when there is a string
    // to match in. `Empty` reaches here as `""` and behaves the same way.
    if hay_chars.is_empty() {
        return Ok(Variant::Long(0));
    }
    if start >= hay_chars.len() {
        return Ok(Variant::Long(0));
    }
    // Empty needle matches at the start position, as VBA has it.
    if needle.is_empty() {
        return Ok(Variant::Long(start as i32 + 1));
    }
    let tail: String = hay_chars[start.min(hay_chars.len())..].iter().collect();
    Ok(match tail.find(&needle) {
        Some(byte_idx) => {
            let char_idx = tail[..byte_idx].chars().count();
            Variant::Long((start + char_idx + 1) as i32)
        }
        None => Variant::Long(0),
    })
}
