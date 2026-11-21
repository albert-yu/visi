use std::rc::Rc;

use crate::core::{date, date_fn, finance};

use super::value::{self, VResult, VarArray, Variant, VbaError};

const HANDLES_NULL: &[&str] = &[
    "isnull",
    "isempty",
    "isnumeric",
    "isdate",
    "isobject",
    "iserror",
    "ismissing",
    "isarray",
    "typename",
    "vartype",
    "iif",
    "array",
    "choose",
    "switch",
];

const REJECTS_NULL: &[&str] = &[
    "cstr",
    "cbyte",
    "cint",
    "clng",
    "clnglng",
    "clngptr",
    "cdbl",
    "csng",
    "cbool",
    "ccur",
    "cdec",
    "cdate",
    "cvdate",
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
    "split",
    "join",
    "filter",
    "dateserial",
    "timeserial",
    "datevalue",
    "timevalue",
];

fn arg(args: &[Variant], i: usize) -> Variant {
    args.get(i).cloned().unwrap_or(Variant::Empty)
}

fn need(args: &[Variant], i: usize) -> VResult<Variant> {
    args.get(i).cloned().ok_or_else(VbaError::invalid_call)
}

fn any_null(args: &[Variant]) -> bool {
    args.iter().any(|a| a.is_null())
}

#[cfg(test)]
#[rustfmt::skip]
pub(super) const IMPLEMENTED_NAMES: &[&str] = &[
    "abs", "array", "asc", "ascw", "atn", "beep", "cbool", "cbyte", "ccur", "cdate", "cdbl",
    "cdec", "choose", "chr", "chrw", "cint", "clng", "clnglng", "clngptr", "cos", "csng", "cstr",
    "cvar", "cvdate", "cverr", "date", "dateadd", "datediff", "datepart", "dateserial",
    "datevalue", "day", "ddb", "doevents", "exp", "filter", "fix", "format", "formatcurrency",
    "formatdatetime", "formatnumber", "formatpercent", "fv", "hex", "hour", "iif", "instr",
    "instrrev", "int", "ipmt", "irr", "isarray", "isdate", "isempty", "iserror", "ismissing",
    "isnull", "isnumeric", "isobject", "join", "lbound", "lcase", "left", "len", "log", "ltrim",
    "mid", "minute", "mirr", "month", "monthname", "now", "nper", "npv", "oct", "pmt", "ppmt",
    "pv", "rate", "replace", "rgb", "right", "round", "rtrim", "second", "sgn", "sin", "sln",
    "space", "split", "sqr", "str", "strcomp", "strconv", "string", "strreverse", "switch", "syd",
    "tan", "time", "timer", "timeserial", "timevalue", "trim", "typename", "ubound", "ucase",
    "val", "vartype", "weekday", "weekdayname", "year",
];

#[cfg(test)]
pub(super) fn implemented_names() -> impl Iterator<Item = &'static str> {
    IMPLEMENTED_NAMES.iter().copied()
}

pub fn call(name: &str, args: &[Variant]) -> VResult<Option<Variant>> {
    call_with_compare(name, args, false)
}

pub fn call_with_compare(
    name: &str,
    args: &[Variant],
    default_text_compare: bool,
) -> VResult<Option<Variant>> {
    let lower = name.to_ascii_lowercase();

    if any_null(args) && !HANDLES_NULL.contains(&lower.as_str()) {
        if REJECTS_NULL.contains(&lower.as_str()) {
            return Err(VbaError::invalid_null());
        }
        return Ok(Some(Variant::Null));
    }

    let v = match lower.as_str() {
        "typename" => Variant::Str(arg(args, 0).type_name().to_string()),
        "vartype" => Variant::Integer(vartype(&arg(args, 0))),
        "isnull" => Variant::Boolean(arg(args, 0).is_null()),
        "isempty" => Variant::Boolean(arg(args, 0).is_empty()),
        "isnumeric" => Variant::Boolean(is_numeric(&arg(args, 0))),
        "isdate" => Variant::Boolean(is_date(&arg(args, 0))),
        "isobject" => Variant::Boolean(matches!(arg(args, 0), Variant::Object(_))),
        "iserror" => Variant::Boolean(matches!(
            arg(args, 0),
            Variant::ErrValue(_) | Variant::Missing
        )),
        "ismissing" => Variant::Boolean(matches!(arg(args, 0), Variant::Missing)),
        "isarray" => Variant::Boolean(matches!(arg(args, 0), Variant::Array(_))),
        "cverr" => Variant::ErrValue(need(args, 0)?.to_f64()? as i32),

        "ubound" | "lbound" => {
            let Variant::Array(a) = need(args, 0)? else {
                return Err(VbaError::type_mismatch());
            };
            let dim = match args.get(1) {
                Some(d) => {
                    let v = value::bankers_round(d.to_f64()?);
                    if v < 1.0 {
                        return Err(VbaError::subscript());
                    }
                    v as usize
                }
                None => 1,
            };
            if lower == "lbound" {
                Variant::Long(a.lbound(dim)?)
            } else {
                Variant::Long(a.ubound(dim)?)
            }
        }
        "array" => Variant::Array(std::rc::Rc::new(value::VarArray::new_1d(0, args.to_vec()))),

        "cstr" => Variant::Str(need(args, 0)?.to_vba_string()?),
        "cbyte" => {
            let r = value::bankers_round(numeric_arg(args, 0)?);
            if !(0.0..=255.0).contains(&r) {
                return Err(VbaError::overflow());
            }
            Variant::Byte(r as u8)
        }
        "cint" => pack_int(numeric_arg(args, 0)?)?,
        "clng" => pack_long(numeric_arg(args, 0)?)?,
        "clnglng" | "clngptr" => {
            let r = value::bankers_round(numeric_arg(args, 0)?);
            if !r.is_finite() || r < i64::MIN as f64 || r > i64::MAX as f64 {
                return Err(VbaError::overflow());
            }
            Variant::LongLong(r as i64)
        }
        "cdec" => value::to_decimal(&need(args, 0)?)?,
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
        "val" => Variant::Double(val_of(&arg(args, 0))),

        "abs" => same_width(&need(args, 0)?, f64::abs)?,
        "sgn" => Variant::Integer(sgn(need(args, 0)?.to_f64()?)),
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
        "round" => {
            let x = need(args, 0)?.to_f64()?;
            let places = match args.get(1) {
                Some(v) => v.to_f64()? as i32,
                None => 0,
            };
            let factor = 10f64.powi(places);
            Variant::Double(value::bankers_round(x * factor) / factor)
        }

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
        "instr" => instr(args, default_text_compare)?,
        "instrrev" => instrrev(args, default_text_compare)?,
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
        "str" => {
            let n = need(args, 0)?.to_f64()?;
            let s = value::format_number(n);
            if n >= 0.0 && !s.starts_with('-') {
                Variant::Str(format!(" {s}"))
            } else {
                Variant::Str(s)
            }
        }
        "strconv" => {
            let s = need(args, 0)?.to_vba_string()?;
            let conv = value::bankers_round(need(args, 1)?.to_f64()?) as i32;
            Variant::Str(str_conv(&s, conv))
        }
        "replace" => replace_fn(args, default_text_compare)?,
        "strcomp" => {
            let a = need(args, 0)?.to_vba_string()?;
            let b = need(args, 1)?.to_vba_string()?;
            let text_cmp = resolve_compare(args.get(2), default_text_compare)?;
            let ord = if text_cmp {
                a.to_lowercase().cmp(&b.to_lowercase())
            } else {
                a.cmp(&b)
            };
            Variant::Integer(match ord {
                std::cmp::Ordering::Less => -1,
                std::cmp::Ordering::Equal => 0,
                std::cmp::Ordering::Greater => 1,
            })
        }
        "split" => split_fn(args, default_text_compare)?,
        "join" => join_fn(args)?,
        "filter" => filter_fn(args, default_text_compare)?,
        "choose" => choose_fn(args)?,
        "switch" => switch_fn(args)?,
        "format" => format_fn(args)?,
        "formatnumber" => format_num_kind(args, NumFmtKind::Number)?,
        "formatcurrency" => format_num_kind(args, NumFmtKind::Currency)?,
        "formatpercent" => format_num_kind(args, NumFmtKind::Percent)?,
        "formatdatetime" => format_datetime_fn(args)?,
        "cdate" | "cvdate" => Variant::Date(to_date_serial(&need(args, 0)?)?),
        "dateserial" => {
            let y = value::bankers_round(need(args, 0)?.to_f64()?) as i32;
            let y = if (0..=29).contains(&y) {
                2000 + y
            } else if (30..=99).contains(&y) {
                1900 + y
            } else {
                y
            };
            let m = value::bankers_round(need(args, 1)?.to_f64()?) as i32;
            let d = value::bankers_round(need(args, 2)?.to_f64()?) as i32;
            Variant::Date(date_fn::ymd_to_serial(y, m, d))
        }
        "timeserial" => {
            let h = need(args, 0)?.to_f64()?;
            let m = need(args, 1)?.to_f64()?;
            let s = need(args, 2)?.to_f64()?;
            Variant::Date((h * 3600.0 + m * 60.0 + s) / 86_400.0)
        }
        "datevalue" => {
            let serial = to_date_serial(&need(args, 0)?)?;
            Variant::Date(serial.trunc())
        }
        "timevalue" => {
            let serial = to_time_serial(&need(args, 0)?)?;
            Variant::Date(serial)
        }
        "year" => {
            let serial = to_date_serial(&need(args, 0)?)?;
            let (y, _, _) = date_fn::serial_to_ymd(serial.trunc());
            Variant::Integer(y as i16)
        }
        "month" => {
            let serial = to_date_serial(&need(args, 0)?)?;
            let (_, m, _) = date_fn::serial_to_ymd(serial.trunc());
            Variant::Integer(m as i16)
        }
        "day" => {
            let serial = to_date_serial(&need(args, 0)?)?;
            let (_, _, d) = date_fn::serial_to_ymd(serial.trunc());
            Variant::Integer(d as i16)
        }
        "hour" => {
            let serial = to_date_serial(&need(args, 0)?)?;
            let secs = (serial.fract().abs() * 86_400.0).round() as i64 % 86_400;
            Variant::Integer((secs / 3600) as i16)
        }
        "minute" => {
            let serial = to_date_serial(&need(args, 0)?)?;
            let secs = (serial.fract().abs() * 86_400.0).round() as i64 % 86_400;
            Variant::Integer(((secs / 60) % 60) as i16)
        }
        "second" => {
            let serial = to_date_serial(&need(args, 0)?)?;
            let secs = (serial.fract().abs() * 86_400.0).round() as i64 % 86_400;
            Variant::Integer((secs % 60) as i16)
        }
        "weekday" => {
            let serial = to_date_serial(&need(args, 0)?)?;
            let first_dow = match args.get(1) {
                Some(v) if !v.is_empty() => value::bankers_round(v.to_f64()?) as i32,
                _ => 1,
            };
            let sun_based = ((serial.trunc() as i64 - 1).rem_euclid(7) + 1) as i32;
            let shift = if first_dow <= 1 { 1 } else { first_dow };
            let res = ((sun_based - shift).rem_euclid(7)) + 1;
            Variant::Integer(res as i16)
        }
        "weekdayname" => {
            let wd = value::bankers_round(need(args, 0)?.to_f64()?) as i32;
            let abbr = args
                .get(1)
                .map(|v| v.to_bool())
                .transpose()?
                .unwrap_or(false);
            let first = args
                .get(2)
                .map(|v| v.to_f64())
                .transpose()?
                .map(|n| value::bankers_round(n) as i32)
                .unwrap_or(1);
            if !(1..=7).contains(&wd) {
                return Err(VbaError::invalid_call());
            }
            let shift = if first <= 1 { 1 } else { first };
            let idx = ((wd - 1 + shift - 1).rem_euclid(7)) as usize;
            const FULL: [&str; 7] = [
                "Sunday",
                "Monday",
                "Tuesday",
                "Wednesday",
                "Thursday",
                "Friday",
                "Saturday",
            ];
            let name = FULL[idx];
            Variant::Str(if abbr { &name[..3] } else { name }.to_string())
        }
        "monthname" => {
            let m = value::bankers_round(need(args, 0)?.to_f64()?) as i32;
            let abbr = args
                .get(1)
                .map(|v| v.to_bool())
                .transpose()?
                .unwrap_or(false);
            if !(1..=12).contains(&m) {
                return Err(VbaError::invalid_call());
            }
            const MONTHS: [&str; 12] = [
                "January",
                "February",
                "March",
                "April",
                "May",
                "June",
                "July",
                "August",
                "September",
                "October",
                "November",
                "December",
            ];
            let name = MONTHS[m as usize - 1];
            Variant::Str(if abbr { &name[..3] } else { name }.to_string())
        }
        "dateadd" => date_add_fn(args)?,
        "datediff" => date_diff_fn(args)?,
        "datepart" => date_part_fn(args)?,
        "now" => Variant::Date(current_ole_serial()),
        "date" => Variant::Date(current_ole_serial().trunc()),
        "time" => Variant::Date(current_ole_serial().fract().abs()),
        "timer" => Variant::Single((current_ole_serial().fract().abs() * 86_400.0) as f32),
        "ddb" => {
            let cost = need(args, 0)?.to_f64()?;
            let salvage = need(args, 1)?.to_f64()?;
            let life = need(args, 2)?.to_f64()?;
            let period = need(args, 3)?.to_f64()?;
            let factor = args
                .get(4)
                .filter(|v| !v.is_empty())
                .map(|v| v.to_f64())
                .transpose()?
                .unwrap_or(2.0);
            Variant::Double(finance::ddb(cost, salvage, life, period, factor))
        }
        "fv" => {
            let rate = need(args, 0)?.to_f64()?;
            let nper = need(args, 1)?.to_f64()?;
            let pmt = need(args, 2)?.to_f64()?;
            let pv = opt_f64(args, 3, 0.0)?;
            let due = opt_f64(args, 4, 0.0)?;
            Variant::Double(finance::fv(rate, nper, pmt, pv, due))
        }
        "ipmt" => {
            let rate = need(args, 0)?.to_f64()?;
            let per = need(args, 1)?.to_f64()?;
            let nper = need(args, 2)?.to_f64()?;
            let pv = need(args, 3)?.to_f64()?;
            let fv = opt_f64(args, 4, 0.0)?;
            let due = opt_f64(args, 5, 0.0)?;
            Variant::Double(finance::ipmt(rate, per, nper, pv, fv, due))
        }
        "irr" => {
            let vals = array_to_f64s(&need(args, 0)?)?;
            let guess = opt_f64(args, 1, 0.1)?;
            let r = finance::irr(&vals, guess).ok_or_else(VbaError::invalid_call)?;
            Variant::Double(r)
        }
        "mirr" => {
            let vals = array_to_f64s(&need(args, 0)?)?;
            let frate = need(args, 1)?.to_f64()?;
            let rrate = need(args, 2)?.to_f64()?;
            let r = finance::mirr(&vals, frate, rrate).ok_or_else(VbaError::invalid_call)?;
            Variant::Double(r)
        }
        "nper" => {
            let rate = need(args, 0)?.to_f64()?;
            let pmt = need(args, 1)?.to_f64()?;
            let pv = need(args, 2)?.to_f64()?;
            let fv = opt_f64(args, 3, 0.0)?;
            let due = opt_f64(args, 4, 0.0)?;
            let r = finance::nper(rate, pmt, pv, fv, due).ok_or_else(VbaError::invalid_call)?;
            Variant::Double(r)
        }
        "npv" => {
            let rate = need(args, 0)?.to_f64()?;
            let vals = array_to_f64s(&need(args, 1)?)?;
            Variant::Double(finance::npv(rate, &vals))
        }
        "pmt" => {
            let rate = need(args, 0)?.to_f64()?;
            let nper = need(args, 1)?.to_f64()?;
            let pv = need(args, 2)?.to_f64()?;
            let fv = opt_f64(args, 3, 0.0)?;
            let due = opt_f64(args, 4, 0.0)?;
            Variant::Double(finance::pmt(rate, nper, pv, fv, due))
        }
        "ppmt" => {
            let rate = need(args, 0)?.to_f64()?;
            let per = need(args, 1)?.to_f64()?;
            let nper = need(args, 2)?.to_f64()?;
            let pv = need(args, 3)?.to_f64()?;
            let fv = opt_f64(args, 4, 0.0)?;
            let due = opt_f64(args, 5, 0.0)?;
            Variant::Double(finance::ppmt(rate, per, nper, pv, fv, due))
        }
        "pv" => {
            let rate = need(args, 0)?.to_f64()?;
            let nper = need(args, 1)?.to_f64()?;
            let pmt = need(args, 2)?.to_f64()?;
            let fv = opt_f64(args, 3, 0.0)?;
            let due = opt_f64(args, 4, 0.0)?;
            Variant::Double(finance::pv(rate, nper, pmt, fv, due))
        }
        "rate" => {
            let nper = need(args, 0)?.to_f64()?;
            let pmt = need(args, 1)?.to_f64()?;
            let pv = need(args, 2)?.to_f64()?;
            let fv = opt_f64(args, 3, 0.0)?;
            let due = opt_f64(args, 4, 0.0)?;
            let guess = opt_f64(args, 5, 0.1)?;
            let r =
                finance::rate(nper, pmt, pv, fv, due, guess).ok_or_else(VbaError::invalid_call)?;
            Variant::Double(r)
        }
        "sln" => {
            let cost = need(args, 0)?.to_f64()?;
            let salvage = need(args, 1)?.to_f64()?;
            let life = need(args, 2)?.to_f64()?;
            if life == 0.0 {
                return Err(VbaError::div_by_zero());
            }
            Variant::Double(finance::sln(cost, salvage, life))
        }
        "syd" => {
            let cost = need(args, 0)?.to_f64()?;
            let salvage = need(args, 1)?.to_f64()?;
            let life = need(args, 2)?.to_f64()?;
            let per = need(args, 3)?.to_f64()?;
            Variant::Double(finance::syd(cost, salvage, life, per))
        }
        "doevents" => Variant::Integer(0),
        "beep" => Variant::Empty,
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
        Variant::Object(_) => 9,
        Variant::ErrValue(_) | Variant::Missing => 10,
        Variant::Boolean(_) => 11,
        Variant::Decimal(_) => 14,
        Variant::Byte(_) => 17,
        Variant::LongLong(_) => 20,
        Variant::Udt(_) => 36,
        Variant::Array(_) => 8192 + 12,
    }
}

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
        Variant::Empty => true,
        Variant::Null => false,
        Variant::Missing
        | Variant::ErrValue(_)
        | Variant::Object(_)
        | Variant::Array(_)
        | Variant::Udt(_) => false,
        _ => true,
    }
}

fn val_of(v: &Variant) -> f64 {
    let Ok(s) = v.to_vba_string() else {
        return 0.0;
    };
    let t: String = s.chars().filter(|c| !c.is_whitespace()).collect();
    if let Some(hex) = t.strip_prefix("&H").or_else(|| t.strip_prefix("&h")) {
        let digits: String = hex.chars().take_while(|c| c.is_ascii_hexdigit()).collect();
        return i64::from_str_radix(&digits, 16).unwrap_or(0) as f64;
    }
    if let Some(oct) = t.strip_prefix("&O").or_else(|| t.strip_prefix("&o")) {
        let digits: String = oct.chars().take_while(|c| matches!(c, '0'..='7')).collect();
        return i64::from_str_radix(&digits, 8).unwrap_or(0) as f64;
    }
    let t = t.replace(['d', 'D'], "e");
    let mut end = 0;
    for (i, _) in t.char_indices() {
        if t[..=i].parse::<f64>().is_ok() {
            end = i + 1;
        }
    }
    t[..end].parse::<f64>().unwrap_or(0.0)
}

fn is_date(v: &Variant) -> bool {
    match v {
        Variant::Date(_) => true,
        Variant::Str(s) => to_date_serial_str(s).is_ok(),
        _ => false,
    }
}

fn to_date_serial(v: &Variant) -> VResult<f64> {
    match v {
        Variant::Date(d) => Ok(*d),
        Variant::Str(s) => to_date_serial_str(s),
        other => other.to_f64(),
    }
}

fn to_date_serial_str(s: &str) -> VResult<f64> {
    let t = s.trim();
    if let Some((d, _)) = date::parse_date(t) {
        return Ok(date::date_to_excel_serial(d));
    }
    if let Some(tf) = date_fn::parse_time_fraction(t) {
        return Ok(tf);
    }
    if let Some((dp, tp)) = t.split_once(' ')
        && let Some((d, _)) = date::parse_date(dp.trim())
        && let Some(tf) = date_fn::parse_time_fraction(tp.trim())
    {
        let day = date::date_to_excel_serial(d);
        return Ok(if day < 0.0 { day - tf } else { day + tf });
    }
    Err(VbaError::type_mismatch())
}

fn to_time_serial(v: &Variant) -> VResult<f64> {
    match v {
        Variant::Date(d) => Ok(d.fract().abs()),
        Variant::Str(s) => {
            let t = s.trim();
            if let Some(tf) = date_fn::parse_time_fraction(t) {
                return Ok(tf);
            }
            if let Some((_, tp)) = t.split_once(' ')
                && let Some(tf) = date_fn::parse_time_fraction(tp.trim())
            {
                return Ok(tf);
            }
            Err(VbaError::type_mismatch())
        }
        other => Ok(other.to_f64()?.fract().abs()),
    }
}

fn current_ole_serial() -> f64 {
    let secs = web_time::SystemTime::now()
        .duration_since(web_time::UNIX_EPOCH)
        .map(|d| d.as_secs_f64())
        .unwrap_or(0.0);
    25569.0 + secs / 86_400.0
}

fn opt_f64(args: &[Variant], idx: usize, default: f64) -> VResult<f64> {
    match args.get(idx) {
        Some(v) if !v.is_empty() && !matches!(v, Variant::Missing) => v.to_f64(),
        _ => Ok(default),
    }
}

fn array_to_f64s(v: &Variant) -> VResult<Vec<f64>> {
    let Variant::Array(arr) = v else {
        return Err(VbaError::type_mismatch());
    };
    arr.values.iter().map(|x| x.to_f64()).collect()
}

fn resolve_compare(arg: Option<&Variant>, default_text: bool) -> VResult<bool> {
    match arg {
        Some(v) if !v.is_empty() && !matches!(v, Variant::Missing) => {
            let n = value::bankers_round(v.to_f64()?) as i32;
            match n {
                -1 => Ok(default_text),
                0 => Ok(false),
                1 => Ok(true),
                _ => Err(VbaError::invalid_call()),
            }
        }
        _ => Ok(default_text),
    }
}

fn str_conv(s: &str, conv: i32) -> String {
    match conv {
        1 => s.to_uppercase(),
        2 => s.to_lowercase(),
        3 => {
            let mut out = String::with_capacity(s.len());
            let mut cap_next = true;
            for ch in s.chars() {
                if ch.is_alphabetic() {
                    if cap_next {
                        out.extend(ch.to_uppercase());
                        cap_next = false;
                    } else {
                        out.extend(ch.to_lowercase());
                    }
                } else {
                    out.push(ch);
                    cap_next = ch.is_whitespace() || ch == '-' || ch == '/';
                }
            }
            out
        }
        _ => s.to_string(),
    }
}

fn replace_fn(args: &[Variant], default_text: bool) -> VResult<Variant> {
    let s = need(args, 0)?.to_vba_string()?;
    let find = need(args, 1)?.to_vba_string()?;
    let with = need(args, 2)?.to_vba_string()?;
    let start = match args.get(3) {
        Some(v) if !v.is_empty() && !matches!(v, Variant::Missing) => {
            let st = value::bankers_round(v.to_f64()?) as i64;
            if st < 1 {
                return Err(VbaError::invalid_call());
            }
            st as usize - 1
        }
        _ => 0,
    };
    let count = match args.get(4) {
        Some(v) if !v.is_empty() && !matches!(v, Variant::Missing) => {
            let c = value::bankers_round(v.to_f64()?) as i64;
            if c < -1 {
                return Err(VbaError::invalid_call());
            }
            c
        }
        _ => -1,
    };
    let text_cmp = resolve_compare(args.get(5), default_text)?;
    let chars: Vec<char> = s.chars().collect();
    if start >= chars.len() {
        return Ok(Variant::Str(String::new()));
    }
    let slice: String = chars[start..].iter().collect();
    if find.is_empty() || count == 0 {
        return Ok(Variant::Str(slice));
    }
    if !text_cmp {
        let res = if count < 0 {
            slice.replace(&find, &with)
        } else {
            slice.replacen(&find, &with, count as usize)
        };
        return Ok(Variant::Str(res));
    }
    let hay_chars: Vec<char> = slice.chars().collect();
    let needle_lower: Vec<char> = find.to_lowercase().chars().collect();
    let mut out = String::new();
    let mut i = 0;
    let mut replaced = 0i64;
    while i < hay_chars.len() {
        let can_match = (count < 0 || replaced < count)
            && i + needle_lower.len() <= hay_chars.len()
            && hay_chars[i..i + needle_lower.len()]
                .iter()
                .zip(needle_lower.iter())
                .all(|(a, b)| a.to_lowercase().eq(std::iter::once(*b)));
        if can_match {
            out.push_str(&with);
            i += needle_lower.len();
            replaced += 1;
        } else {
            out.push(hay_chars[i]);
            i += 1;
        }
    }
    Ok(Variant::Str(out))
}

fn split_fn(args: &[Variant], default_text: bool) -> VResult<Variant> {
    let expr = need(args, 0)?.to_vba_string()?;
    let delim = match args.get(1) {
        Some(v) if !matches!(v, Variant::Missing) => v.to_vba_string()?,
        _ => " ".to_string(),
    };
    let limit = match args.get(2) {
        Some(v) if !v.is_empty() && !matches!(v, Variant::Missing) => {
            value::bankers_round(v.to_f64()?) as i64
        }
        _ => -1,
    };
    let text_cmp = resolve_compare(args.get(3), default_text)?;
    if expr.is_empty() || limit == 0 {
        return Ok(Variant::Array(Rc::new(VarArray::new_1d(0, Vec::new()))));
    }
    if delim.is_empty() || limit == 1 {
        return Ok(Variant::Array(Rc::new(VarArray::new_1d(
            0,
            vec![Variant::Str(expr)],
        ))));
    }
    let parts: Vec<Variant> = if !text_cmp {
        if limit < 0 {
            expr.split(&delim)
                .map(|p| Variant::Str(p.to_string()))
                .collect()
        } else {
            expr.splitn(limit as usize, &delim)
                .map(|p| Variant::Str(p.to_string()))
                .collect()
        }
    } else {
        let hay: Vec<char> = expr.chars().collect();
        let ndl: Vec<char> = delim.to_lowercase().chars().collect();
        let mut res = Vec::new();
        let mut start = 0usize;
        let mut i = 0usize;
        while i + ndl.len() <= hay.len() {
            if limit > 0 && res.len() + 1 >= limit as usize {
                break;
            }
            let matched = hay[i..i + ndl.len()]
                .iter()
                .zip(ndl.iter())
                .all(|(a, b)| a.to_lowercase().eq(std::iter::once(*b)));
            if matched {
                res.push(Variant::Str(hay[start..i].iter().collect()));
                i += ndl.len();
                start = i;
            } else {
                i += 1;
            }
        }
        res.push(Variant::Str(hay[start..].iter().collect()));
        res
    };
    Ok(Variant::Array(Rc::new(VarArray::new_1d(0, parts))))
}

fn join_fn(args: &[Variant]) -> VResult<Variant> {
    let Variant::Array(arr) = need(args, 0)? else {
        return Err(VbaError::type_mismatch());
    };
    let delim = match args.get(1) {
        Some(v) if !matches!(v, Variant::Missing) => v.to_vba_string()?,
        _ => " ".to_string(),
    };
    let mut parts = Vec::with_capacity(arr.values.len());
    for v in &arr.values {
        parts.push(v.to_vba_string()?);
    }
    Ok(Variant::Str(parts.join(&delim)))
}

fn filter_fn(args: &[Variant], default_text: bool) -> VResult<Variant> {
    let Variant::Array(arr) = need(args, 0)? else {
        return Err(VbaError::type_mismatch());
    };
    let needle = need(args, 1)?.to_vba_string()?;
    let include = match args.get(2) {
        Some(v) if !v.is_empty() && !matches!(v, Variant::Missing) => v.to_bool()?,
        _ => true,
    };
    let text_cmp = resolve_compare(args.get(3), default_text)?;
    let needle_cmp = if text_cmp {
        needle.to_lowercase()
    } else {
        needle.clone()
    };
    let mut out = Vec::new();
    for v in &arr.values {
        let s = v.to_vba_string()?;
        let found = if text_cmp {
            s.to_lowercase().contains(&needle_cmp)
        } else {
            s.contains(&needle_cmp)
        };
        if found == include {
            out.push(Variant::Str(s));
        }
    }
    Ok(Variant::Array(Rc::new(VarArray::new_1d(0, out))))
}

fn choose_fn(args: &[Variant]) -> VResult<Variant> {
    let first = need(args, 0)?;
    if first.is_null() {
        return Ok(Variant::Null);
    }
    let idx = value::bankers_round(first.to_f64()?) as i64;
    if idx < 1 || (idx as usize) >= args.len() {
        return Ok(Variant::Null);
    }
    Ok(args[idx as usize].clone())
}

fn switch_fn(args: &[Variant]) -> VResult<Variant> {
    if !args.len().is_multiple_of(2) {
        return Err(VbaError::invalid_call());
    }
    let mut i = 0;
    while i < args.len() {
        if !args[i].is_null() && args[i].to_bool()? {
            return Ok(args[i + 1].clone());
        }
        i += 2;
    }
    Ok(Variant::Null)
}

#[derive(Clone, Copy)]
enum NumFmtKind {
    Number,
    Currency,
    Percent,
}

fn format_num_kind(args: &[Variant], kind: NumFmtKind) -> VResult<Variant> {
    let val = need(args, 0)?.to_f64()?;
    let decimals = match args.get(1) {
        Some(v) if !v.is_empty() && !matches!(v, Variant::Missing) => {
            let d = value::bankers_round(v.to_f64()?) as i32;
            if d < 0 { 2usize } else { d as usize }
        }
        _ => 2usize,
    };
    let scaled = match kind {
        NumFmtKind::Percent => val * 100.0,
        _ => val,
    };
    let factor = 10f64.powi(decimals as i32);
    let rounded = value::bankers_round(scaled * factor) / factor;
    let formatted = format_grouped(rounded.abs(), decimals, true);
    let neg = rounded < 0.0;
    let out = match kind {
        NumFmtKind::Number => {
            if neg {
                format!("-{formatted}")
            } else {
                formatted
            }
        }
        NumFmtKind::Currency => {
            if neg {
                format!("(${formatted})")
            } else {
                format!("${formatted}")
            }
        }
        NumFmtKind::Percent => {
            if neg {
                format!("-{formatted}%")
            } else {
                format!("{formatted}%")
            }
        }
    };
    Ok(Variant::Str(out))
}

fn format_grouped(abs_val: f64, decimals: usize, group: bool) -> String {
    let raw = format!("{abs_val:.decimals$}");
    if !group {
        return raw;
    }
    let (int_part, frac_part) = match raw.split_once('.') {
        Some((i, f)) => (i, Some(f)),
        None => (raw.as_str(), None),
    };
    let mut grouped = String::new();
    for (idx, ch) in int_part.chars().rev().enumerate() {
        if idx > 0 && idx % 3 == 0 {
            grouped.push(',');
        }
        grouped.push(ch);
    }
    let int_grouped: String = grouped.chars().rev().collect();
    match frac_part {
        Some(f) => format!("{int_grouped}.{f}"),
        None => int_grouped,
    }
}

fn format_datetime_fn(args: &[Variant]) -> VResult<Variant> {
    let serial = to_date_serial(&need(args, 0)?)?;
    let named = match args.get(1) {
        Some(v) if !v.is_empty() && !matches!(v, Variant::Missing) => {
            value::bankers_round(v.to_f64()?) as i32
        }
        _ => 0,
    };
    let (y, m, d) = date_fn::serial_to_ymd(serial.trunc());
    let secs = (serial.fract().abs() * 86_400.0).round() as i64 % 86_400;
    let (h24, min, sec) = (secs / 3600, (secs / 60) % 60, secs % 60);
    let (h12, ampm) = match h24 {
        0 => (12, "AM"),
        1..=11 => (h24, "AM"),
        12 => (12, "PM"),
        _ => (h24 - 12, "PM"),
    };
    let out = match named {
        0 => value::format_vba_date(serial),
        1 => format!("{}, {} {}, {}", weekday_full(serial), month_full(m), d, y),
        2 => format!("{m}/{d}/{y}"),
        3 => format!("{h12}:{min:02}:{sec:02} {ampm}"),
        4 => format!("{h24:02}:{min:02}"),
        _ => return Err(VbaError::invalid_call()),
    };
    Ok(Variant::Str(out))
}

fn weekday_full(serial: f64) -> &'static str {
    const FULL: [&str; 7] = [
        "Sunday",
        "Monday",
        "Tuesday",
        "Wednesday",
        "Thursday",
        "Friday",
        "Saturday",
    ];
    let idx = (serial.trunc() as i64 - 1).rem_euclid(7) as usize;
    FULL[idx]
}

fn month_full(m: i32) -> &'static str {
    const MONTHS: [&str; 12] = [
        "January",
        "February",
        "March",
        "April",
        "May",
        "June",
        "July",
        "August",
        "September",
        "October",
        "November",
        "December",
    ];
    MONTHS[(m.clamp(1, 12) - 1) as usize]
}

fn format_fn(args: &[Variant]) -> VResult<Variant> {
    let val = need(args, 0)?;
    let Some(fmt_arg) = args.get(1) else {
        return Ok(Variant::Str(val.to_vba_string()?));
    };
    if fmt_arg.is_empty() || matches!(fmt_arg, Variant::Missing) {
        return Ok(Variant::Str(val.to_vba_string()?));
    }
    let fmt = fmt_arg.to_vba_string()?;
    match fmt.to_ascii_lowercase().as_str() {
        "general number" => Ok(Variant::Str(value::format_number(val.to_f64()?))),
        "currency" => format_num_kind(&[val], NumFmtKind::Currency),
        "fixed" => Ok(Variant::Str(format!("{:.2}", val.to_f64()?))),
        "standard" => format_num_kind(&[val], NumFmtKind::Number),
        "percent" => format_num_kind(&[val], NumFmtKind::Percent),
        "scientific" => Ok(Variant::Str(format!("{:.2E}", val.to_f64()?))),
        "yes/no" => Ok(Variant::Str(
            if val.to_bool()? { "Yes" } else { "No" }.to_string(),
        )),
        "true/false" => Ok(Variant::Str(
            if val.to_bool()? { "True" } else { "False" }.to_string(),
        )),
        "on/off" => Ok(Variant::Str(
            if val.to_bool()? { "On" } else { "Off" }.to_string(),
        )),
        "general date" => format_datetime_fn(&[val, Variant::Integer(0)]),
        "long date" => format_datetime_fn(&[val, Variant::Integer(1)]),
        "short date" => format_datetime_fn(&[val, Variant::Integer(2)]),
        "long time" => format_datetime_fn(&[val, Variant::Integer(3)]),
        "short time" => format_datetime_fn(&[val, Variant::Integer(4)]),
        ">" => Ok(Variant::Str(val.to_vba_string()?.to_uppercase())),
        "<" => Ok(Variant::Str(val.to_vba_string()?.to_lowercase())),
        _ => {
            if fmt
                .chars()
                .any(|c| matches!(c, 'y' | 'Y' | 'd' | 'D' | 'h' | 'H' | 'n' | 'N' | 's' | 'S'))
                && !fmt.contains('0')
                && !fmt.contains('#')
            {
                let serial = to_date_serial(&val)?;
                return Ok(Variant::Str(format_custom_date(serial, &fmt)));
            }
            if fmt.contains('0') || fmt.contains('#') {
                let n = val.to_f64()?;
                return Ok(Variant::Str(format_custom_number(n, &fmt)));
            }
            Ok(Variant::Str(val.to_vba_string()?))
        }
    }
}

fn format_custom_number(val: f64, fmt: &str) -> String {
    let is_pct = fmt.contains('%');
    let has_group = fmt.contains(',');
    let scaled = if is_pct { val * 100.0 } else { val };
    let decimals = match fmt.split_once('.') {
        Some((_, frac)) => frac.chars().filter(|c| *c == '0' || *c == '#').count(),
        None => 0,
    };
    let factor = 10f64.powi(decimals as i32);
    let rounded = value::bankers_round(scaled * factor) / factor;
    let body = format_grouped(rounded.abs(), decimals, has_group);
    let prefix: String = fmt
        .chars()
        .take_while(|c| !matches!(c, '0' | '#' | '.' | ',' | '%'))
        .collect();
    let suffix = if is_pct { "%" } else { "" };
    if rounded < 0.0 {
        format!("-{prefix}{body}{suffix}")
    } else {
        format!("{prefix}{body}{suffix}")
    }
}

fn format_custom_date(serial: f64, fmt: &str) -> String {
    let (y, m, d) = date_fn::serial_to_ymd(serial.trunc());
    let secs = (serial.fract().abs() * 86_400.0).round() as i64 % 86_400;
    let (h24, min, sec) = (secs / 3600, (secs / 60) % 60, secs % 60);
    let has_ampm = fmt.to_ascii_lowercase().contains("am/pm");
    let h_disp = if has_ampm {
        match h24 {
            0 => 12,
            1..=12 => h24,
            _ => h24 - 12,
        }
    } else {
        h24
    };
    let ampm = if h24 < 12 { "AM" } else { "PM" };
    let chars: Vec<char> = fmt.chars().collect();
    let mut out = String::new();
    let mut i = 0;
    while i < chars.len() {
        let rest: String = chars[i..].iter().collect();
        let lower_rest = rest.to_ascii_lowercase();
        if lower_rest.starts_with("am/pm") {
            out.push_str(ampm);
            i += 5;
        } else if lower_rest.starts_with("yyyy") {
            out.push_str(&format!("{y:04}"));
            i += 4;
        } else if lower_rest.starts_with("yy") {
            out.push_str(&format!("{:02}", y.rem_euclid(100)));
            i += 2;
        } else if lower_rest.starts_with("mmmm") {
            out.push_str(month_full(m));
            i += 4;
        } else if lower_rest.starts_with("mmm") {
            out.push_str(&month_full(m)[..3]);
            i += 3;
        } else if lower_rest.starts_with("mm") {
            out.push_str(&format!("{m:02}"));
            i += 2;
        } else if lower_rest.starts_with('m') {
            out.push_str(&m.to_string());
            i += 1;
        } else if lower_rest.starts_with("dddd") {
            out.push_str(weekday_full(serial));
            i += 4;
        } else if lower_rest.starts_with("ddd") {
            out.push_str(&weekday_full(serial)[..3]);
            i += 3;
        } else if lower_rest.starts_with("dd") {
            out.push_str(&format!("{d:02}"));
            i += 2;
        } else if lower_rest.starts_with('d') {
            out.push_str(&d.to_string());
            i += 1;
        } else if lower_rest.starts_with("hh") {
            out.push_str(&format!("{h_disp:02}"));
            i += 2;
        } else if lower_rest.starts_with('h') {
            out.push_str(&h_disp.to_string());
            i += 1;
        } else if lower_rest.starts_with("nn") {
            out.push_str(&format!("{min:02}"));
            i += 2;
        } else if lower_rest.starts_with('n') {
            out.push_str(&min.to_string());
            i += 1;
        } else if lower_rest.starts_with("ss") {
            out.push_str(&format!("{sec:02}"));
            i += 2;
        } else if lower_rest.starts_with('s') {
            out.push_str(&sec.to_string());
            i += 1;
        } else {
            out.push(chars[i]);
            i += 1;
        }
    }
    out
}

fn date_add_fn(args: &[Variant]) -> VResult<Variant> {
    let interval = need(args, 0)?.to_vba_string()?.to_ascii_lowercase();
    let number = value::bankers_round(need(args, 1)?.to_f64()?) as i32;
    let serial = to_date_serial(&need(args, 2)?)?;
    let (y, m, d) = date_fn::serial_to_ymd(serial.trunc());
    let frac = serial.fract();
    let res = match interval.as_str() {
        "yyyy" => {
            let ny = y + number;
            let nd = d.min(date_fn::days_in_month(ny, m));
            date_fn::ymd_to_serial(ny, m, nd) + frac
        }
        "q" => {
            let total_m = (m - 1) + number * 3;
            let ny = y + total_m.div_euclid(12);
            let nm = total_m.rem_euclid(12) + 1;
            let nd = d.min(date_fn::days_in_month(ny, nm));
            date_fn::ymd_to_serial(ny, nm, nd) + frac
        }
        "m" => {
            let total_m = (m - 1) + number;
            let ny = y + total_m.div_euclid(12);
            let nm = total_m.rem_euclid(12) + 1;
            let nd = d.min(date_fn::days_in_month(ny, nm));
            date_fn::ymd_to_serial(ny, nm, nd) + frac
        }
        "y" | "d" | "w" => serial + f64::from(number),
        "ww" => serial + f64::from(number) * 7.0,
        "h" => serial + f64::from(number) / 24.0,
        "n" => serial + f64::from(number) / 1440.0,
        "s" => serial + f64::from(number) / 86_400.0,
        _ => return Err(VbaError::invalid_call()),
    };
    Ok(Variant::Date(res))
}

fn date_diff_fn(args: &[Variant]) -> VResult<Variant> {
    let interval = need(args, 0)?.to_vba_string()?.to_ascii_lowercase();
    let d1 = to_date_serial(&need(args, 1)?)?;
    let d2 = to_date_serial(&need(args, 2)?)?;
    let (y1, m1, _) = date_fn::serial_to_ymd(d1.trunc());
    let (y2, m2, _) = date_fn::serial_to_ymd(d2.trunc());
    let diff: i64 = match interval.as_str() {
        "yyyy" => i64::from(y2 - y1),
        "q" => i64::from((y2 - y1) * 4 + ((m2 - 1) / 3 - (m1 - 1) / 3)),
        "m" => i64::from((y2 - y1) * 12 + (m2 - m1)),
        "y" | "d" => (d2.trunc() - d1.trunc()) as i64,
        "w" | "ww" => ((d2.trunc() - d1.trunc()) / 7.0).trunc() as i64,
        "h" => ((d2 - d1) * 24.0).trunc() as i64,
        "n" => ((d2 - d1) * 1440.0).trunc() as i64,
        "s" => ((d2 - d1) * 86_400.0).round() as i64,
        _ => return Err(VbaError::invalid_call()),
    };
    Ok(Variant::Long(diff as i32))
}

fn date_part_fn(args: &[Variant]) -> VResult<Variant> {
    let interval = need(args, 0)?.to_vba_string()?.to_ascii_lowercase();
    let serial = to_date_serial(&need(args, 1)?)?;
    let (y, m, d) = date_fn::serial_to_ymd(serial.trunc());
    let secs = (serial.fract().abs() * 86_400.0).round() as i64 % 86_400;
    let res: i32 = match interval.as_str() {
        "yyyy" => y,
        "q" => (m - 1) / 3 + 1,
        "m" => m,
        "y" => (serial.trunc() - date_fn::ymd_to_serial(y, 1, 1)) as i32 + 1,
        "d" => d,
        "w" => ((serial.trunc() as i64 - 1).rem_euclid(7) + 1) as i32,
        "ww" => {
            let doy = (serial.trunc() - date_fn::ymd_to_serial(y, 1, 1)) as i32;
            doy / 7 + 1
        }
        "h" => (secs / 3600) as i32,
        "n" => ((secs / 60) % 60) as i32,
        "s" => (secs % 60) as i32,
        _ => return Err(VbaError::invalid_call()),
    };
    Ok(Variant::Integer(res as i16))
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

fn instr(args: &[Variant], default_text: bool) -> VResult<Variant> {
    let (start, hay, needle, text_cmp) = if args.len() >= 3 {
        let s = need(args, 0)?.to_f64()?;
        if s < 1.0 {
            return Err(VbaError::invalid_call());
        }
        (
            s as usize - 1,
            need(args, 1)?.to_vba_string()?,
            need(args, 2)?.to_vba_string()?,
            resolve_compare(args.get(3), default_text)?,
        )
    } else {
        (
            0,
            need(args, 0)?.to_vba_string()?,
            need(args, 1)?.to_vba_string()?,
            default_text,
        )
    };
    let hay_chars: Vec<char> = hay.chars().collect();
    if hay_chars.is_empty() {
        return Ok(Variant::Long(0));
    }
    if start >= hay_chars.len() {
        return Ok(Variant::Long(0));
    }
    if needle.is_empty() {
        return Ok(Variant::Long(start as i32 + 1));
    }
    let tail: String = hay_chars[start.min(hay_chars.len())..].iter().collect();
    let found = if text_cmp {
        let tl = tail.to_lowercase();
        let nl = needle.to_lowercase();
        tl.find(&nl).map(|b| tl[..b].chars().count())
    } else {
        tail.find(&needle).map(|b| tail[..b].chars().count())
    };
    Ok(match found {
        Some(char_idx) => Variant::Long((start + char_idx + 1) as i32),
        None => Variant::Long(0),
    })
}

fn instrrev(args: &[Variant], default_text: bool) -> VResult<Variant> {
    let hay = need(args, 0)?.to_vba_string()?;
    let needle = need(args, 1)?.to_vba_string()?;
    let hay_chars: Vec<char> = hay.chars().collect();
    let start = match args.get(2) {
        Some(v) if !v.is_empty() && !matches!(v, Variant::Missing) => {
            let s = value::bankers_round(v.to_f64()?) as i64;
            if s == -1 {
                hay_chars.len()
            } else if s < 1 {
                return Err(VbaError::invalid_call());
            } else {
                (s as usize).min(hay_chars.len())
            }
        }
        _ => hay_chars.len(),
    };
    let text_cmp = resolve_compare(args.get(3), default_text)?;
    if hay_chars.is_empty() {
        return Ok(Variant::Long(0));
    }
    if needle.is_empty() {
        return Ok(Variant::Long(start as i32));
    }
    let prefix: String = hay_chars[..start].iter().collect();
    let found = if text_cmp {
        let pl = prefix.to_lowercase();
        let nl = needle.to_lowercase();
        pl.rfind(&nl).map(|b| pl[..b].chars().count())
    } else {
        prefix.rfind(&needle).map(|b| prefix[..b].chars().count())
    };
    Ok(match found {
        Some(char_idx) => Variant::Long((char_idx + 1) as i32),
        None => Variant::Long(0),
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn every_listed_name_is_dispatched() {
        for name in IMPLEMENTED_NAMES {
            assert!(
                !matches!(call(name, &[]), Ok(None)),
                "`{name}` is in IMPLEMENTED_NAMES but `call` does not handle it"
            );
        }
    }

    #[test]
    fn an_unknown_name_is_reported_as_unknown() {
        assert!(matches!(call("NoSuchIntrinsic", &[]), Ok(None)));
    }
}
