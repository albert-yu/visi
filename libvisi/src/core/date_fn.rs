// High-precision Date and Time functions for libvisi
// Implements Excel-compatible serial date calculations, 1900 leap year bug support, day/month/year extractions, and workday/networkdays routines.

pub fn ymd_to_serial(year: i32, month: i32, day: i32) -> f64 {
    let mut y = year;
    let mut m = month;
    if m > 12 {
        y += (m - 1) / 12;
        m = (m - 1) % 12 + 1;
    } else if m < 1 {
        let adj = (12 - m) / 12;
        y -= adj;
        m += adj * 12;
    }

    // Days before year y (using Gregorian rules)
    let y1 = y - 1;
    let mut days = y1 * 365 + y1 / 4 - y1 / 100 + y1 / 400;

    let leap = (y % 4 == 0 && y % 100 != 0) || (y % 400 == 0);
    let month_days = [
        0,
        31,
        if leap { 29 } else { 28 },
        31,
        30,
        31,
        30,
        31,
        31,
        30,
        31,
        30,
        31,
    ];

    for i in 1..m {
        days += month_days[i as usize];
    }
    days += day;

    // Excel 1900 epoch offset (1900-01-01 is serial 1).
    // Dec 31 1BC is day 0 in this count.
    // Excel includes non-existent leap day Feb 29, 1900 (serial 60).
    let mut serial = (days - 693595) as f64;
    if serial >= 60.0 {
        serial += 1.0;
    }
    serial
}

pub fn serial_to_ymd(serial: f64) -> (i32, i32, i32) {
    let mut s = serial.floor() as i64;
    if s < 1 {
        return (1900, 1, 1);
    }
    if s == 60 {
        return (1900, 2, 29);
    }
    if s > 60 {
        s -= 1;
    }
    // Shift to Dec 31 1BC offset
    let days = s + 693595;

    let mut y = (days as f64 / 365.2425) as i64 + 2;
    let mut y1 = y - 1;
    let mut d_count: i64 = y1 * 365 + y1 / 4 - y1 / 100 + y1 / 400;

    while d_count >= days {
        y -= 1;
        y1 = y - 1;
        d_count = y1 * 365 + y1 / 4 - y1 / 100 + y1 / 400;
    }

    let mut rem_days = days - d_count;
    let leap = (y % 4 == 0 && y % 100 != 0) || (y % 400 == 0);
    let month_days = [
        0,
        31,
        if leap { 29 } else { 28 },
        31,
        30,
        31,
        30,
        31,
        31,
        30,
        31,
        30,
        31,
    ];

    let mut m = 1;
    while m <= 12 && rem_days > month_days[m as usize] {
        rem_days -= month_days[m as usize];
        m += 1;
    }

    (y as i32, m, rem_days as i32)
}

pub fn date_fn(year: f64, month: f64, day: f64) -> Result<f64, String> {
    let y = year.floor() as i32;
    let m = month.floor() as i32;
    let d = day.floor() as i32;
    if !(0..=9999).contains(&y) {
        Err("#NUM!".to_string())
    } else {
        let full_y = if y < 1900 { 1900 + y } else { y };
        Ok(ymd_to_serial(full_y, m, d))
    }
}

pub fn day_fn(serial: f64) -> Result<f64, String> {
    if serial < 0.0 {
        Err("#NUM!".to_string())
    } else {
        let (_, _, d) = serial_to_ymd(serial);
        Ok(d as f64)
    }
}

pub fn month_fn(serial: f64) -> Result<f64, String> {
    if serial < 0.0 {
        Err("#NUM!".to_string())
    } else {
        let (_, m, _) = serial_to_ymd(serial);
        Ok(m as f64)
    }
}

pub fn year_fn(serial: f64) -> Result<f64, String> {
    if serial < 0.0 {
        Err("#NUM!".to_string())
    } else {
        let (y, _, _) = serial_to_ymd(serial);
        Ok(y as f64)
    }
}

pub fn days(end_date: f64, start_date: f64) -> Result<f64, String> {
    Ok(end_date.floor() - start_date.floor())
}

pub fn days360(start_date: f64, end_date: f64, method: Option<bool>) -> Result<f64, String> {
    let (y1, m1, mut d1) = serial_to_ymd(start_date);
    let (y2, m2, mut d2) = serial_to_ymd(end_date);
    let is_euro = method.unwrap_or(false);

    if is_euro {
        if d1 == 31 { d1 = 30; }
        if d2 == 31 { d2 = 30; }
    } else {
        if d1 == 31 { d1 = 30; }
        if d2 == 31 && d1 == 30 { d2 = 30; }
    }

    Ok(((y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1)) as f64)
}

pub fn edate(start_date: f64, months: f64) -> Result<f64, String> {
    let (y, m, d) = serial_to_ymd(start_date);
    let total_m = m + months.floor() as i32;
    let target_y = y + (total_m - 1).div_euclid(12);
    let target_m = (total_m - 1).rem_euclid(12) + 1;

    let leap = (target_y % 4 == 0 && target_y % 100 != 0) || (target_y % 400 == 0);
    let max_days = [0, 31, if leap { 29 } else { 28 }, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][target_m as usize];
    let target_d = d.min(max_days);

    Ok(ymd_to_serial(target_y, target_m, target_d) + start_date.fract())
}

pub fn eomonth(start_date: f64, months: f64) -> Result<f64, String> {
    let (y, m, _) = serial_to_ymd(start_date);
    let total_m = m + months.floor() as i32;
    let target_y = y + (total_m - 1).div_euclid(12);
    let target_m = (total_m - 1).rem_euclid(12) + 1;

    let leap = (target_y % 4 == 0 && target_y % 100 != 0) || (target_y % 400 == 0);
    let max_days = [0, 31, if leap { 29 } else { 28 }, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][target_m as usize];

    Ok(ymd_to_serial(target_y, target_m, max_days))
}

pub fn time_fn(hour: f64, minute: f64, second: f64) -> Result<f64, String> {
    let h = hour.floor();
    let m = minute.floor();
    let s = second.floor();
    let total_seconds = h * 3600.0 + m * 60.0 + s;
    let fraction = (total_seconds / 86400.0).rem_euclid(1.0);
    Ok(fraction)
}

pub fn hour_fn(serial: f64) -> Result<f64, String> {
    let frac = serial.fract().abs();
    let secs = (frac * 86400.0).round() as u64;
    Ok(((secs / 3600) % 24) as f64)
}

pub fn minute_fn(serial: f64) -> Result<f64, String> {
    let frac = serial.fract().abs();
    let secs = (frac * 86400.0).round() as u64;
    Ok(((secs / 60) % 60) as f64)
}

pub fn second_fn(serial: f64) -> Result<f64, String> {
    let frac = serial.fract().abs();
    let secs = (frac * 86400.0).round() as u64;
    Ok((secs % 60) as f64)
}

pub fn weekday(serial: f64, return_type: Option<f64>) -> Result<f64, String> {
    let s = serial.floor() as i64;
    if s < 1 {
        return Err("#NUM!".to_string());
    }
    let r_type = return_type.unwrap_or(1.0).floor() as i32;
    // 1900-01-01 (serial 1) was Sunday
    let base_day = (s + 6) % 7; // 0=Sunday, 1=Monday, ..., 6=Saturday

    match r_type {
        1 => Ok((base_day + 1) as f64), // 1=Sunday..7=Saturday
        2 => Ok((if base_day == 0 { 7 } else { base_day }) as f64), // 1=Monday..7=Sunday
        3 => Ok((if base_day == 0 { 6 } else { base_day - 1 }) as f64), // 0=Monday..6=Sunday
        11 => Ok((if base_day == 0 { 7 } else { base_day }) as f64),
        12 => Ok(((base_day + 5) % 7 + 1) as f64),
        _ => Ok((base_day + 1) as f64),
    }
}

pub fn weeknum(serial: f64, return_type: Option<f64>) -> Result<f64, String> {
    let (y, _, _) = serial_to_ymd(serial);
    let jan1 = ymd_to_serial(y, 1, 1);
    let r_type = return_type.unwrap_or(1.0).floor() as i32;

    if r_type == 21 {
        return isoweeknum(serial);
    }
    let days_diff = serial.floor() - jan1;
    let jan1_weekday = weekday(jan1, Some(r_type as f64))?;
    Ok(((days_diff + jan1_weekday - 1.0) / 7.0).floor() + 1.0)
}

pub fn isoweeknum(serial: f64) -> Result<f64, String> {
    let (y, _m, _d) = serial_to_ymd(serial);
    let day_of_year = serial.floor() - ymd_to_serial(y, 1, 1) + 1.0;
    let wday = weekday(serial, Some(2.0))?; // 1=Mon..7=Sun
    let iso_week = ((day_of_year - wday + 10.0) / 7.0).floor();

    if iso_week < 1.0 {
        return isoweeknum(ymd_to_serial(y - 1, 12, 31));
    }
    if iso_week > 52.0 {
        let next_jan1_wday = weekday(ymd_to_serial(y + 1, 1, 1), Some(2.0))?;
        if next_jan1_wday <= 4.0 && (day_of_year >= 363.0 - (if y % 4 == 0 { 1.0 } else { 0.0 })) {
            return Ok(1.0);
        }
    }
    Ok(iso_week)
}

pub fn networkdays(start_date: f64, end_date: f64, holidays: &[f64]) -> Result<f64, String> {
    let s = start_date.floor() as i64;
    let e = end_date.floor() as i64;
    if s > e {
        return Ok(-networkdays(end_date, start_date, holidays)?);
    }
    let hol_set: Vec<i64> = holidays.iter().map(|h| h.floor() as i64).collect();

    let mut count = 0;
    for day in s..=e {
        let w = weekday(day as f64, Some(2.0))? as i64; // 1=Mon..7=Sun
        if w <= 5 && !hol_set.contains(&day) {
            count += 1;
        }
    }
    Ok(count as f64)
}

pub fn workday(start_date: f64, days: f64, holidays: &[f64]) -> Result<f64, String> {
    let mut curr = start_date.floor() as i64;
    let mut remaining = days.floor() as i64;
    let step = if remaining >= 0 { 1 } else { -1 };
    let hol_set: Vec<i64> = holidays.iter().map(|h| h.floor() as i64).collect();

    while remaining != 0 {
        curr += step;
        let w = weekday(curr as f64, Some(2.0))? as i64;
        if w <= 5 && !hol_set.contains(&curr) {
            remaining -= step;
        }
    }
    Ok(curr as f64)
}

/// Actual/actual year length for basis-1 day-count conventions. Confirmed
/// against real Excel via the differential fuzzer to have two regimes:
/// for a span of at most 366 days (including one that crosses a calendar
/// year boundary, e.g. Dec into Jan), it's simply whether the *later*
/// date's own calendar year is a leap year -- not a days-weighted blend
/// of the two years. Only once the span genuinely covers multiple full
/// calendar years does it become the average of 365/366 across every
/// year from `start`'s year through `end`'s year inclusive (e.g. a span
/// covering 4 calendar years with a single leap year among them averages
/// to (365*3+366)/4 = 365.25). Shared by `YEARFRAC` and the bond/discount
/// functions in `finance.rs` that use this same basis-1 convention.
pub fn actual_actual_year_days(start: f64, end: f64) -> f64 {
    let d1 = start.min(end);
    let d2 = start.max(end);
    if d2 - d1 <= 366.0 {
        let (y, _, _) = serial_to_ymd(d2);
        let leap = (y % 4 == 0 && y % 100 != 0) || (y % 400 == 0);
        return if leap { 366.0 } else { 365.0 };
    }
    let (y1, _, _) = serial_to_ymd(d1);
    let (y2, _, _) = serial_to_ymd(d2);
    let n = y2 - y1 + 1;
    let total: i32 = (y1..=y2)
        .map(|y| {
            let leap = (y % 4 == 0 && y % 100 != 0) || (y % 400 == 0);
            if leap { 366 } else { 365 }
        })
        .sum();
    total as f64 / n as f64
}

pub fn yearfrac(start_date: f64, end_date: f64, basis: Option<f64>) -> Result<f64, String> {
    let b = basis.unwrap_or(0.0).floor() as i32;
    let d1 = start_date.min(end_date);
    let d2 = start_date.max(end_date);
    let diff = d2 - d1;

    match b {
        0 => Ok(days360(d1, d2, Some(false))? / 360.0),
        1 => Ok(diff / actual_actual_year_days(d1, d2)),
        2 => Ok(diff / 360.0),
        3 => Ok(diff / 365.0),
        4 => Ok(days360(d1, d2, Some(true))? / 360.0),
        _ => Err("#NUM!".to_string()),
    }
}
