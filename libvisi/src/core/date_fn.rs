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
        // Serial 0 is Excel's phantom "January 0, 1900", not 1 January:
        // DAY(0) is 0, MONTH(0) is 1, YEAR(0) is 1900, and
        // TEXT(0.6299, "yyyy-mm-dd") is "1900-01-00". Returning day 1 here
        // made every one of those off by a day.
        return (1900, 1, 0);
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

/// Parses a date-only text into an (year, month, day) triple, without
/// resolving it to a serial number yet (datevalue and the date portion of
/// value() in text.rs both need this). Supports the formats Excel's own
/// DATEVALUE recognizes without relying on the current locale: ISO
/// `YYYY-MM-DD` / `YYYY/MM/DD`, and US-style `M/D/YYYY` (2- or 4-digit
/// year, 2-digit year assumed 20xx for 00-29 and 19xx for 30-99, matching
/// Excel's own pivot point).
pub fn parse_date_parts(text: &str) -> Option<(i32, i32, i32)> {
    let s = text.trim();
    if s.is_empty() {
        return None;
    }
    let parts: Vec<&str> = if s.contains('-') {
        s.split('-').collect()
    } else if s.contains('/') {
        s.split('/').collect()
    } else {
        return None;
    };
    if parts.len() != 3 {
        return None;
    }
    let nums: Vec<i32> = parts
        .iter()
        .filter_map(|p| p.trim().parse::<i32>().ok())
        .collect();
    if nums.len() != 3 {
        return None;
    }

    // ISO order (year first) if the first component looks like a year.
    let (y, m, d) = if parts[0].trim().len() == 4 {
        (nums[0], nums[1], nums[2])
    } else {
        // US order: month/day/year, with a 2- or 4-digit year.
        let mut y = nums[2];
        if parts[2].trim().len() <= 2 {
            y = if y <= 29 { 2000 + y } else { 1900 + y };
        }
        (y, nums[0], nums[1])
    };

    if !(1..=12).contains(&m) || !(1..=31).contains(&d) {
        return None;
    }
    Some((y, m, d))
}

/// Days in a given month, honouring leap years.
pub fn days_in_month(year: i32, month: i32) -> i32 {
    match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 => {
            if (year % 4 == 0 && year % 100 != 0) || year % 400 == 0 {
                29
            } else {
                28
            }
        }
        _ => 30,
    }
}

/// DATEDIF's calendar difference, decomposed into whole years, whole
/// months and leftover days -- the shared basis for all six unit codes.
///
/// The point of the borrowing below is that DATEDIF counts *completed*
/// intervals. If the end day-of-month hasn't reached the start's yet,
/// that month isn't complete: the day difference borrows the length of
/// the month preceding the end date, and the month count drops by one
/// (which can in turn borrow a year). Computing the parts independently
/// instead -- plain `m2 - m1`, plain `d2 - d1` -- overcounts by a month
/// whenever the end day is earlier, and can even go negative ("MD" of a
/// pair whose end day precedes its start day used to report -9).
fn datedif_parts(start: f64, end: f64) -> (i32, i32, i32) {
    let (y1, m1, d1) = serial_to_ymd(start);
    let (y2, m2, d2) = serial_to_ymd(end);

    let mut years = y2 - y1;
    let mut months = m2 - m1;
    let mut days = d2 - d1;

    if days < 0 {
        months -= 1;
        let (prev_y, prev_m) = if m2 == 1 { (y2 - 1, 12) } else { (y2, m2 - 1) };
        days += days_in_month(prev_y, prev_m);
    }
    if months < 0 {
        years -= 1;
        months += 12;
    }
    (years, months, days)
}

pub fn datedif(start: f64, end: f64, unit: &str) -> Result<f64, String> {
    if end < start {
        return Err("#NUM!".to_string());
    }
    let (years, months, days) = datedif_parts(start, end);
    let (y1, m1, d1) = serial_to_ymd(start);
    match unit.to_uppercase().as_str() {
        "Y" => Ok(years as f64),
        "M" => Ok((years * 12 + months) as f64),
        "D" => Ok(end.floor() - start.floor()),
        "MD" => Ok(days as f64),
        "YM" => Ok(months as f64),
        // Days since the most recent anniversary of the start date, i.e.
        // the day count with whole years removed.
        "YD" => {
            let anniversary = ymd_to_serial(y1 + years, m1, d1);
            Ok(end.floor() - anniversary)
        }
        _ => Err("#NUM!".to_string()),
    }
}

pub fn datevalue(text: &str) -> Result<f64, String> {
    match parse_date_parts(text) {
        Some((y, m, d)) => Ok(ymd_to_serial(y, m, d)),
        None => Err("#VALUE!".to_string()),
    }
}

/// Parses a time-only text into a day fraction (0.0..1.0). Supports
/// `H:MM`, `H:MM:SS`, and either with a trailing `AM`/`PM` marker
/// (case-insensitive, with or without a separating space).
pub fn parse_time_fraction(text: &str) -> Option<f64> {
    let s = text.trim();
    if s.is_empty() {
        return None;
    }
    let lower = s.to_lowercase();
    let (body, meridiem) = if let Some(stripped) = lower.strip_suffix("am") {
        (stripped.trim(), Some(true))
    } else if let Some(stripped) = lower.strip_suffix("pm") {
        (stripped.trim(), Some(false))
    } else {
        (lower.as_str(), None)
    };

    let parts: Vec<&str> = body.split(':').collect();
    if parts.len() < 2 || parts.len() > 3 {
        return None;
    }
    let mut h: f64 = parts[0].trim().parse().ok()?;
    let m: f64 = parts[1].trim().parse().ok()?;
    let s_val: f64 = if parts.len() == 3 {
        parts[2].trim().parse().ok()?
    } else {
        0.0
    };
    if !(0.0..60.0).contains(&m) || !(0.0..60.0).contains(&s_val) {
        return None;
    }

    if let Some(is_am) = meridiem {
        if !(1.0..=12.0).contains(&h) {
            return None;
        }
        h = if is_am {
            if h == 12.0 { 0.0 } else { h }
        } else if h == 12.0 {
            12.0
        } else {
            h + 12.0
        };
    } else if !(0.0..24.0).contains(&h) {
        return None;
    }

    let total_seconds = h * 3600.0 + m * 60.0 + s_val;
    Some((total_seconds / 86400.0).rem_euclid(1.0))
}

pub fn timevalue(text: &str) -> Result<f64, String> {
    parse_time_fraction(text).ok_or_else(|| "#VALUE!".to_string())
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
        if d1 == 31 {
            d1 = 30;
        }
        if d2 == 31 {
            d2 = 30;
        }
    } else {
        // The DAYS360 *function's* US method, which is not quite the NASD
        // 30/360 convention Excel's own YEARFRAC and bond functions use --
        // see `days_30_360_nasd`, which documents the two differences and
        // the cases that separate them. The end-of-February rule below is
        // what distinguishes this from the European method; without it the
        // two agree on every pair that avoids a February month-end, which
        // is why omitting it went unnoticed for a long time.
        if m1 == 2 && d1 == days_in_month(y1, m1) {
            d1 = 30;
        }
        if d1 == 31 {
            d1 = 30;
        }
        // Note this tests the *adjusted* d1, so a February month-end start
        // does pull a 31st end date back to the 30th here.
        if d2 == 31 && d1 == 30 {
            d2 = 30;
        }
    }

    Ok(((y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1)) as f64)
}

/// NASD 30/360, but with a *month-end end date* first pulled back to the
/// 30th -- including February's month end, which is what separates it from
/// the plain European rule.
///
/// Excel uses this for the two ODDLPRICE/ODDLYIELD quantities whose end
/// date is a **coupon date** (the quasi-coupon period length, and the
/// last-interest-to-maturity span), while the spans that end at the
/// *settlement* date use plain `days_30_360_nasd`. So the same pair of
/// dates can count differently depending on which role it plays:
///
/// ```text
/// last_interest 2017-12-27, maturity 2018-02-28, settlement 2018-01-04
///   last_interest -> maturity    63   (28 Feb is a month end -> 30)
///   settlement    -> maturity    54   (plain NASD, 28 Feb stays 28)
/// ```
///
/// Fitted against 20 real-Excel ODDLPRICE values covering month-end and
/// non-month-end maturities, leap and non-leap Februaries, and month-end
/// last-interest dates.
pub fn days_30_360_coupon_end(start_date: f64, end_date: f64) -> f64 {
    let (y1, m1, mut d1) = serial_to_ymd(start_date);
    let (y2, m2, mut d2) = serial_to_ymd(end_date);

    if m1 == 2 && d1 == days_in_month(y1, m1) {
        d1 = 30;
    }
    if d1 == 31 {
        d1 = 30;
    }
    if d2 == days_in_month(y2, m2) {
        d2 = 30;
    }

    ((y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1)) as f64
}

/// The NASD 30/360 day count that Excel's YEARFRAC (basis 0) and the bond
/// functions use. This is *not* what the DAYS360 function computes, which
/// is why it lives here separately rather than sharing `days360`'s US
/// branch. Two rules differ, and each shows up on its own:
///
/// - When both ends are February month-ends, this pulls the end date to
///   the 30th as well. DAYS360 does not:
///   `YEARFRAC(2003-02-28, 2005-02-28, 0) * 360` is 720 while
///   `DAYS360(2003-02-28, 2005-02-28, FALSE)` is 718.
/// - The "end date on the 31st comes back to the 30th" rule tests the
///   start day *before* it was adjusted, so a February month-end start
///   does not trigger it: `YEARFRAC(2003-02-28, 2005-03-31, 0) * 360` is
///   751 where `DAYS360` gives 750.
///
/// Verified against real Excel over twelve date pairs chosen to separate
/// the two rule sets.
pub fn days_30_360_nasd(start_date: f64, end_date: f64) -> f64 {
    let (y1, m1, mut d1) = serial_to_ymd(start_date);
    let (y2, m2, mut d2) = serial_to_ymd(end_date);

    let d1_is_feb_eom = m1 == 2 && d1 == days_in_month(y1, m1);
    let d2_is_feb_eom = m2 == 2 && d2 == days_in_month(y2, m2);
    // Tested before d1 is adjusted below.
    let d1_was_month_end = d1 == 30 || d1 == 31;

    if d1_is_feb_eom && d2_is_feb_eom {
        d2 = 30;
    }
    if d1_is_feb_eom {
        d1 = 30;
    }
    if d2 == 31 && d1_was_month_end {
        d2 = 30;
    }
    if d1 == 31 {
        d1 = 30;
    }

    ((y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1)) as f64
}

pub fn edate(start_date: f64, months: f64) -> Result<f64, String> {
    let (y, m, d) = serial_to_ymd(start_date);
    let total_m = m + months.floor() as i32;
    let target_y = y + (total_m - 1).div_euclid(12);
    let target_m = (total_m - 1).rem_euclid(12) + 1;

    let leap = (target_y % 4 == 0 && target_y % 100 != 0) || (target_y % 400 == 0);
    let max_days = [
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
    ][target_m as usize];
    let target_d = d.min(max_days);

    Ok(ymd_to_serial(target_y, target_m, target_d) + start_date.fract())
}

pub fn eomonth(start_date: f64, months: f64) -> Result<f64, String> {
    let (y, m, _) = serial_to_ymd(start_date);
    let total_m = m + months.floor() as i32;
    let target_y = y + (total_m - 1).div_euclid(12);
    let target_m = (total_m - 1).rem_euclid(12) + 1;

    let leap = (target_y % 4 == 0 && target_y % 100 != 0) || (target_y % 400 == 0);
    let max_days = [
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
    ][target_m as usize];

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
        // Within a single year the denominator is 366 when either the
        // period actually contains a 29 February, or the whole period
        // lies inside one leap year; otherwise 365. Taking the *end*
        // year's leap-ness alone (what this used to do) is wrong for a
        // short period that ends in a leap year before the leap day.
        //
        // Four real-Excel data points pin all three branches down:
        //   2016-06-01 -> 2016-09-01  366  (wholly inside leap 2016)
        //   2027-12-26 -> 2028-03-26  366  (spans 29 Feb 2028)
        //   2023-08-05 -> 2024-01-05  365  (ends in a leap year, before
        //                                   the leap day)
        //   2015-02-15 -> 2016-02-13  365  (same, and YEARFRAC there is
        //                                   363/365 = 0.994520547945205
        //                                   to the last digit)
        let (y1, _, _) = serial_to_ymd(d1);
        let (y2, _, _) = serial_to_ymd(d2);
        let is_leap = |y: i32| (y % 4 == 0 && y % 100 != 0) || (y % 400 == 0);
        let wholly_inside_leap_year = y1 == y2 && is_leap(y1);
        let contains_leap_day = (y1..=y2).any(|y| {
            is_leap(y) && {
                let feb29 = ymd_to_serial(y, 2, 29);
                feb29 >= d1 && feb29 <= d2
            }
        });
        return if wholly_inside_leap_year || contains_leap_day {
            366.0
        } else {
            365.0
        };
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
        // The NASD convention, not the DAYS360 function's US method --
        // Excel's own YEARFRAC and DAYS360 genuinely disagree on February
        // month-ends. See `days_30_360_nasd`.
        0 => Ok(days_30_360_nasd(d1, d2) / 360.0),
        1 => Ok(diff / actual_actual_year_days(d1, d2)),
        2 => Ok(diff / 360.0),
        3 => Ok(diff / 365.0),
        4 => Ok(days360(d1, d2, Some(true))? / 360.0),
        _ => Err("#NUM!".to_string()),
    }
}
