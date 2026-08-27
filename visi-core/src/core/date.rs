//! Recognizing dates written as text, and converting them to Excel serials.
//!
//! [`parse_date`] infers both the date and the [`DateFormat`] it was written
//! in; [`date_to_excel_serial`] converts to Excel's day count, reproducing the
//! 1900 leap-year bug.
//!
//! The `DateFormat` half is what lets a date cell echo back in the notation it
//! was typed in, the way Excel does: `6/22/26` stays `6/22/26` rather than
//! normalizing to ISO. [`DateFormat::to_format_code`] lowers it to an Excel
//! number-format code and [`render_date_code`] renders that code, so this
//! module and `text::text_fn`'s `TEXT()` share one date formatter instead of
//! keeping two.
//!
//! The value itself stays a plain numeric serial, as it is in Excel -- the
//! notation lives on the cell, as `CellStyle::num_format`. `engine::sheet`
//! records it when it recognizes a literal and renders through it in
//! `get_display_string`; `xlsx` maps it to and from a worksheet `numFmt`.
//! Month-name casing is the one detail a format code cannot carry, so it
//! survives [`format_date`] but not a round trip through a worksheet --
//! which is Excel's behavior too.
//!
//! `DateFormat` records the separator, field order, year width and month-name
//! spelling, but not whether a numeric month or day was zero-padded --
//! `06/22/2026` and `6/22/2026` are the same format. Rendering is unpadded
//! there, which is what Excel also does with `m/d/yyyy`.

use crate::core::locale::{DateOrder, Locale};

const MONTHS_FULL: [&str; 12] = [
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
const MONTHS_SHORT: [&str; 12] = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
];

/// How a month name was capitalized in the text a date was typed as.
///
/// A format code cannot carry casing, so this rides alongside
/// [`DateFormat::to_format_code`] and is lost on a round trip through a
/// worksheet -- as it is in Excel.
#[derive(Clone, Copy, Debug, PartialEq, Eq, serde::Serialize, serde::Deserialize)]
pub enum StringCase {
    /// All lowercase, as in `22-jun-2026`.
    Lower,
    /// All uppercase, as in `22-JUN-2026`.
    Upper,
    /// Leading capital, rest lowercase: `22-Jun-2026`. The default.
    Title,
    /// Mixed in some other way; rendered as the canonical title case.
    Original,
}

pub fn detect_case(s: &str) -> StringCase {
    if s.chars().all(|c| c.is_uppercase()) {
        StringCase::Upper
    } else if s.chars().all(|c| c.is_lowercase()) {
        StringCase::Lower
    } else {
        let mut chars = s.chars();
        if let Some(first) = chars.next()
            && first.is_uppercase()
            && chars.all(|c| c.is_lowercase())
        {
            return StringCase::Title;
        }
        StringCase::Original
    }
}

/// A calendar date, with no time-of-day and no timezone.
///
/// Only an intermediate: cells hold an Excel serial, not a `SimpleDate`. This
/// is what [`parse_date`] produces and what [`date_to_excel_serial`] consumes,
/// so the calendar arithmetic happens in one place.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SimpleDate {
    /// Full year, four digits -- a two-digit year is widened by [`parse_date`].
    pub year: i32,
    /// Month, 1-12.
    pub month: u32,
    /// Day of month, 1-31.
    pub day: u32,
}

pub fn is_leap_year(year: i32) -> bool {
    year % 4 == 0 && (year % 100 != 0 || year % 400 == 0)
}

pub fn days_in_month(year: i32, month: u32) -> u32 {
    match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 => {
            if is_leap_year(year) {
                29
            } else {
                28
            }
        }
        _ => 0,
    }
}

/// The notation a date was written in: field order, separator, year width and
/// month-name spelling.
///
/// This is *detection* output, not the storage form. A cell stores an Excel
/// serial plus the format code this lowers to (`CellStyle::num_format`), which
/// is why a `DateFormat` can express a little more than survives a save --
/// month-name casing has no format-code equivalent, and zero-padding of a
/// numeric month or day is not recorded at all, so `06/22/2026` and
/// `6/22/2026` are the same variant and both render unpadded.
///
/// The two-part variants fill in the missing field: a month/day pair takes
/// `parse_date`'s default year, a month/year pair takes day 1.
#[derive(Debug, Clone, Copy, PartialEq, Eq, serde::Serialize, serde::Deserialize)]
pub enum DateFormat {
    /// Year-month-day, all numeric: `2026-06-22`.
    Ymd {
        /// Character separating the fields, `-` or `/`.
        sep: char,
    },
    /// Month-day-year, all numeric: `06/22/2026`, `6/22/26`.
    Mdy {
        /// Character separating the fields, `-` or `/`.
        sep: char,
        /// Digits the year was written with: 2 or 4.
        year_len: usize,
    },
    /// Day-month-year, all numeric: `22-06-2026`.
    Dmy {
        /// Character separating the fields, `-` or `/`.
        sep: char,
        /// Digits the year was written with: 2 or 4.
        year_len: usize,
    },
    /// Day, month name, year: `22-Jun-2026`, `22-June-26`.
    DMmmY {
        /// Character separating the fields, `-` or `/`.
        sep: char,
        /// Digits the year was written with: 2 or 4.
        year_len: usize,
        /// Casing the month name was typed in.
        month_case: StringCase,
        /// `true` for a full name (`June`), `false` for an abbreviation (`Jun`).
        month_full: bool,
    },
    /// Month name, day, year: `Jun-22-2026`, `June-22-26`.
    MmmDY {
        /// Character separating the fields, `-` or `/`.
        sep: char,
        /// Digits the year was written with: 2 or 4.
        year_len: usize,
        /// Casing the month name was typed in.
        month_case: StringCase,
        /// `true` for a full name (`June`), `false` for an abbreviation (`Jun`).
        month_full: bool,
    },
    /// Year, month name, day: `2026-Jun-22`.
    YMmmD {
        /// Character separating the fields, `-` or `/`.
        sep: char,
        /// Digits the year was written with: 2 or 4.
        year_len: usize,
        /// Casing the month name was typed in.
        month_case: StringCase,
        /// `true` for a full name (`June`), `false` for an abbreviation (`Jun`).
        month_full: bool,
    },

    /// Numeric month and day, year assumed: `6/22`.
    Md {
        /// Character separating the fields, `-` or `/`.
        sep: char,
    },
    /// Numeric day and month, year assumed: `22/6`.
    Dm {
        /// Character separating the fields, `-`, `/`, or `.`.
        sep: char,
    },
    /// Numeric month and year, day assumed to be the 1st: `6/2026`.
    My {
        /// Character separating the fields, `-` or `/`.
        sep: char,
        /// Digits the year was written with: 2 or 4.
        year_len: usize,
    },
    /// Day then month name, year assumed: `22-Jun`.
    DMmm {
        /// Character separating the fields, `-` or `/`.
        sep: char,
        /// Casing the month name was typed in.
        month_case: StringCase,
        /// `true` for a full name (`June`), `false` for an abbreviation (`Jun`).
        month_full: bool,
    },
    /// Month name then day, year assumed: `Jun-22`.
    MmmD {
        /// Character separating the fields, `-` or `/`.
        sep: char,
        /// Casing the month name was typed in.
        month_case: StringCase,
        /// `true` for a full name (`June`), `false` for an abbreviation (`Jun`).
        month_full: bool,
    },
    /// Month name then year, day assumed to be the 1st: `Jun-2026`.
    MmmY {
        /// Character separating the fields, `-` or `/`.
        sep: char,
        /// Digits the year was written with: 2 or 4.
        year_len: usize,
        /// Casing the month name was typed in.
        month_case: StringCase,
        /// `true` for a full name (`June`), `false` for an abbreviation (`Jun`).
        month_full: bool,
    },
    /// Year then month name, day assumed to be the 1st: `2026-Jun`.
    YMmm {
        /// Character separating the fields, `-` or `/`.
        sep: char,
        /// Digits the year was written with: 2 or 4.
        year_len: usize,
        /// Casing the month name was typed in.
        month_case: StringCase,
        /// `true` for a full name (`June`), `false` for an abbreviation (`Jun`).
        month_full: bool,
    },
}

impl DateFormat {
    /// Lowers to an Excel number-format code (`m/d/yy`, `d-mmm-yyyy`, ...).
    ///
    /// This is the interchange form: it is what gets written to the worksheet
    /// as a `numFmt` and what [`render_date_code`] consumes. Month-name casing
    /// has no representation in a format code, so it rides alongside as
    /// [`DateFormat::month_case`].
    pub fn to_format_code(&self) -> String {
        // A month name is "mmm"/"mmmm"; a numeric month is bare "m" because
        // `DateFormat` does not record zero-padding.
        fn month_word(full: bool) -> &'static str {
            if full { "mmmm" } else { "mmm" }
        }
        fn year(len: usize) -> &'static str {
            if len == 2 { "yy" } else { "yyyy" }
        }

        match *self {
            DateFormat::Ymd { sep } => format!("yyyy{sep}mm{sep}dd"),
            DateFormat::Mdy { sep, year_len } => format!("m{sep}d{sep}{}", year(year_len)),
            DateFormat::Dmy { sep, year_len } => format!("d{sep}m{sep}{}", year(year_len)),
            DateFormat::DMmmY {
                sep,
                year_len,
                month_full,
                ..
            } => format!("d{sep}{}{sep}{}", month_word(month_full), year(year_len)),
            DateFormat::MmmDY {
                sep,
                year_len,
                month_full,
                ..
            } => format!("{}{sep}d{sep}{}", month_word(month_full), year(year_len)),
            DateFormat::YMmmD {
                sep,
                year_len,
                month_full,
                ..
            } => format!("{}{sep}{}{sep}d", year(year_len), month_word(month_full)),
            DateFormat::Md { sep } => format!("m{sep}d"),
            DateFormat::Dm { sep } => format!("d{sep}m"),
            DateFormat::My { sep, year_len } => format!("m{sep}{}", year(year_len)),
            DateFormat::DMmm {
                sep, month_full, ..
            } => format!("d{sep}{}", month_word(month_full)),
            DateFormat::MmmD {
                sep, month_full, ..
            } => format!("{}{sep}d", month_word(month_full)),
            DateFormat::MmmY {
                sep,
                year_len,
                month_full,
                ..
            } => format!("{}{sep}{}", month_word(month_full), year(year_len)),
            DateFormat::YMmm {
                sep,
                year_len,
                month_full,
                ..
            } => format!("{}{sep}{}", year(year_len), month_word(month_full)),
        }
    }

    /// The casing the month name was typed in, for the formats that have one.
    pub fn month_case(&self) -> StringCase {
        match *self {
            DateFormat::DMmmY { month_case, .. }
            | DateFormat::MmmDY { month_case, .. }
            | DateFormat::YMmmD { month_case, .. }
            | DateFormat::DMmm { month_case, .. }
            | DateFormat::MmmD { month_case, .. }
            | DateFormat::MmmY { month_case, .. }
            | DateFormat::YMmm { month_case, .. } => month_case,
            _ => StringCase::Title,
        }
    }
}

fn apply_case(s: &str, case: StringCase) -> String {
    match case {
        StringCase::Upper => s.to_uppercase(),
        StringCase::Lower => s.to_lowercase(),
        // Month names are stored title-cased already.
        StringCase::Title | StringCase::Original => s.to_string(),
    }
}

/// Renders a date through an Excel number-format code.
///
/// Handles the date tokens visi recognizes: runs of `y` (1-2 -> 2-digit year,
/// 3+ -> 4-digit), `m` (1 -> bare month, 2 -> zero-padded, 3 -> `Jun`, 4+ ->
/// `June`) and `d` (1 -> bare day, 2+ -> zero-padded). Anything else is copied
/// through verbatim, so separators and literal text survive.
///
/// Tokens are matched as runs in a single pass rather than by successive
/// string replacement, which is what keeps a substituted month name from being
/// re-scanned -- `December` contains an `m` and `May` a `y`.
pub fn render_date_code(date: SimpleDate, code: &str, month_case: StringCase) -> String {
    let chars: Vec<char> = code.chars().collect();
    let mut out = String::with_capacity(code.len() + 8);
    let mut i = 0;

    while i < chars.len() {
        let c = chars[i];
        if c == '\\' {
            if let Some(next) = chars.get(i + 1) {
                out.push(*next);
                i += 2;
            } else {
                i += 1;
            }
            continue;
        }
        let lower = c.to_ascii_lowercase();
        if !matches!(lower, 'y' | 'm' | 'd') {
            out.push(c);
            i += 1;
            continue;
        }

        let mut run = 0;
        while i + run < chars.len() && chars[i + run].to_ascii_lowercase() == lower {
            run += 1;
        }
        i += run;

        match lower {
            'y' => {
                if run <= 2 {
                    out.push_str(&format!("{:02}", date.year.rem_euclid(100)));
                } else {
                    out.push_str(&format!("{:04}", date.year));
                }
            }
            'm' => {
                let idx = (date.month as usize).saturating_sub(1);
                match run {
                    1 => out.push_str(&date.month.to_string()),
                    2 => out.push_str(&format!("{:02}", date.month)),
                    3 => out.push_str(&apply_case(
                        MONTHS_SHORT.get(idx).copied().unwrap_or(""),
                        month_case,
                    )),
                    _ => out.push_str(&apply_case(
                        MONTHS_FULL.get(idx).copied().unwrap_or(""),
                        month_case,
                    )),
                }
            }
            _ => {
                if run == 1 {
                    out.push_str(&date.day.to_string());
                } else {
                    out.push_str(&format!("{:02}", date.day));
                }
            }
        }
    }
    out
}

/// Renders a date back in the notation [`parse_date`] recognized it in.
pub fn format_date(date: SimpleDate, format: &DateFormat) -> String {
    render_date_code(date, &format.to_format_code(), format.month_case())
}

/// Whether a number-format code renders a date, as opposed to a numeric
/// format like `0.00` or `#,##0`.
///
/// Deliberately narrow: it wants a `y`/`m`/`d` token and no digit placeholder,
/// so an unrecognized or numeric code falls back to plain number rendering
/// rather than being mangled into a date.
pub fn is_date_code(code: &str) -> bool {
    let has_date_token = code
        .chars()
        .any(|c| matches!(c.to_ascii_lowercase(), 'y' | 'm' | 'd'));
    let has_number_placeholder = code.contains('0') || code.contains('#');
    has_date_token && !has_number_placeholder
}

/// The inverse of [`date_to_excel_serial`], for rendering a computed serial.
pub fn excel_serial_to_date(serial: f64) -> SimpleDate {
    let (year, month, day) = crate::core::date_fn::serial_to_ymd(serial);
    SimpleDate {
        year,
        month: month.max(0) as u32,
        day: day.max(0) as u32,
    }
}

fn parse_digits(part: &str) -> Option<i32> {
    let trimmed = part.trim().trim_end_matches('.');
    if !trimmed.is_empty() && trimmed.chars().all(|c| c.is_ascii_digit()) {
        trimmed.parse::<i32>().ok()
    } else {
        None
    }
}

/// Recognizes a date written as text using the default US locale.
pub fn parse_date(src: &str) -> Option<(SimpleDate, DateFormat)> {
    parse_date_with_locale(src, &Locale::en_us())
}

/// Recognizes a date written as text according to a specific [`Locale`].
///
/// Returns `None` for anything that is not a date, which is how
/// `Sheet::commit` decides whether a literal becomes a plain number or a
/// number carrying a date format. Text that merely *looks* like a date is
/// therefore quoted on import (`xlsx::text_cell_src`) to keep it text.
pub fn parse_date_with_locale(src: &str, locale: &Locale) -> Option<(SimpleDate, DateFormat)> {
    let default_year = locale.default_year();
    let src_trim = src.trim();
    if src_trim.is_empty() {
        return None;
    }

    // Try standard delimiters: '-', '/', '.'
    for &sep in &['-', '/', '.'] {
        let parts: Vec<&str> = src_trim.split(sep).collect();

        // --- 3 PARTS ---
        if parts.len() == 3 {
            // Check if there is a word month in the parts
            let mut month_word_info = None;
            for (i, part) in parts.iter().enumerate() {
                if let Some((m, is_full)) = locale.match_month_word(part) {
                    month_word_info = Some((i, m, is_full));
                    break;
                }
            }

            if let Some((month_idx, month, is_full)) = month_word_info {
                // If one part is a word month, the other two must be digits
                let mut digit_parts = Vec::new();
                for (i, part) in parts.iter().enumerate() {
                    if i != month_idx
                        && let Some(val) = parse_digits(part)
                    {
                        digit_parts.push((i, val, part.trim().trim_end_matches('.').len()));
                    }
                }

                if digit_parts.len() == 2 {
                    let case = detect_case(parts[month_idx].trim().trim_end_matches('.'));

                    // Case A: Day-Month-Year (e.g., 22-Jun-2026, 22-Jun-26)
                    // month_idx is 1. digit_parts[0] is index 0 (day), digit_parts[1] is index 2 (year).
                    if month_idx == 1 && digit_parts[0].0 == 0 && digit_parts[1].0 == 2 {
                        let day = digit_parts[0].1 as u32;
                        let year_raw = digit_parts[1].1;
                        let year_len = digit_parts[1].2;
                        let year = if year_len == 2 {
                            locale.expand_two_digit_year(year_raw)
                        } else {
                            year_raw
                        };
                        if day >= 1 && day <= days_in_month(year, month) {
                            return Some((
                                SimpleDate { year, month, day },
                                DateFormat::DMmmY {
                                    sep,
                                    year_len,
                                    month_case: case,
                                    month_full: is_full,
                                },
                            ));
                        }
                    }

                    // Case B: Month-Day-Year (e.g., Jun-22-2026)
                    // month_idx is 0. digit_parts[0] is index 1 (day), digit_parts[1] is index 2 (year).
                    if month_idx == 0 && digit_parts[0].0 == 1 && digit_parts[1].0 == 2 {
                        let day = digit_parts[0].1 as u32;
                        let year_raw = digit_parts[1].1;
                        let year_len = digit_parts[1].2;
                        let year = if year_len == 2 {
                            locale.expand_two_digit_year(year_raw)
                        } else {
                            year_raw
                        };
                        if day >= 1 && day <= days_in_month(year, month) {
                            return Some((
                                SimpleDate { year, month, day },
                                DateFormat::MmmDY {
                                    sep,
                                    year_len,
                                    month_case: case,
                                    month_full: is_full,
                                },
                            ));
                        }
                    }

                    // Case C: Year-Month-Day (e.g. 2026-Jun-22)
                    // month_idx is 1. digit_parts[0] is index 0 (year), digit_parts[1] is index 2 (day).
                    if month_idx == 1 && digit_parts[0].0 == 0 && digit_parts[1].0 == 2 {
                        let year_raw = digit_parts[0].1;
                        let year_len = digit_parts[0].2;
                        let year = if year_len == 2 {
                            locale.expand_two_digit_year(year_raw)
                        } else {
                            year_raw
                        };
                        let day = digit_parts[1].1 as u32;
                        if day >= 1 && day <= days_in_month(year, month) {
                            return Some((
                                SimpleDate { year, month, day },
                                DateFormat::YMmmD {
                                    sep,
                                    year_len,
                                    month_case: case,
                                    month_full: is_full,
                                },
                            ));
                        }
                    }
                }
            } else {
                // All 3 parts are digits (e.g. 2026-06-22, 06-22-2026, 22-06-2026)
                if let (Some(val0), Some(val1), Some(val2)) = (
                    parse_digits(parts[0]),
                    parse_digits(parts[1]),
                    parse_digits(parts[2]),
                ) {
                    let len0 = parts[0].trim().len();
                    let len2 = parts[2].trim().len();

                    // Option A: YMD (Year first) - len0 == 4
                    if len0 == 4 {
                        let year = val0;
                        let month = val1 as u32;
                        let day = val2 as u32;
                        if (1..=12).contains(&month)
                            && day >= 1
                            && day <= days_in_month(year, month)
                        {
                            return Some((
                                SimpleDate { year, month, day },
                                DateFormat::Ymd { sep },
                            ));
                        }
                    }

                    // Option B: MDY or DMY (Year last) - len2 == 4 or 2
                    if len2 == 4 || len2 == 2 {
                        let year_raw = val2;
                        let year = if len2 == 2 {
                            locale.expand_two_digit_year(year_raw)
                        } else {
                            year_raw
                        };

                        match locale.date_order {
                            DateOrder::Dmy => {
                                let day = val0 as u32;
                                let month = val1 as u32;
                                if (1..=12).contains(&month)
                                    && day >= 1
                                    && day <= days_in_month(year, month)
                                {
                                    return Some((
                                        SimpleDate { year, month, day },
                                        DateFormat::Dmy {
                                            sep,
                                            year_len: len2,
                                        },
                                    ));
                                }
                            }
                            DateOrder::Mdy => {
                                if val0 > 12 && val1 <= 12 {
                                    // Unambiguous DMY fallback
                                    let day = val0 as u32;
                                    let month = val1 as u32;
                                    if (1..=12).contains(&month)
                                        && day >= 1
                                        && day <= days_in_month(year, month)
                                    {
                                        return Some((
                                            SimpleDate { year, month, day },
                                            DateFormat::Dmy {
                                                sep,
                                                year_len: len2,
                                            },
                                        ));
                                    }
                                } else {
                                    let month = val0 as u32;
                                    let day = val1 as u32;
                                    if (1..=12).contains(&month)
                                        && day >= 1
                                        && day <= days_in_month(year, month)
                                    {
                                        return Some((
                                            SimpleDate { year, month, day },
                                            DateFormat::Mdy {
                                                sep,
                                                year_len: len2,
                                            },
                                        ));
                                    }
                                }
                            }
                            DateOrder::Ymd => {
                                if len0 == 2 && len2 <= 2 {
                                    let y = locale.expand_two_digit_year(val0);
                                    let m = val1 as u32;
                                    let d = val2 as u32;
                                    if (1..=12).contains(&m) && d >= 1 && d <= days_in_month(y, m) {
                                        return Some((
                                            SimpleDate {
                                                year: y,
                                                month: m,
                                                day: d,
                                            },
                                            DateFormat::Ymd { sep },
                                        ));
                                    }
                                }
                                let month = val0 as u32;
                                let day = val1 as u32;
                                if (1..=12).contains(&month)
                                    && day >= 1
                                    && day <= days_in_month(year, month)
                                {
                                    return Some((
                                        SimpleDate { year, month, day },
                                        DateFormat::Mdy {
                                            sep,
                                            year_len: len2,
                                        },
                                    ));
                                }
                            }
                        }
                    }
                }
            }
        }

        // --- 2 PARTS ---
        if parts.len() == 2 {
            // Check if there is a word month in the parts
            let mut month_word_info = None;
            for (i, part) in parts.iter().enumerate() {
                if let Some((m, is_full)) = locale.match_month_word(part) {
                    month_word_info = Some((i, m, is_full));
                    break;
                }
            }

            if let Some((month_idx, month, is_full)) = month_word_info {
                let digit_idx = if month_idx == 0 { 1 } else { 0 };
                if let Some(digit_val) = parse_digits(parts[digit_idx]) {
                    let digit_len = parts[digit_idx].trim().trim_end_matches('.').len();
                    let case = detect_case(parts[month_idx].trim().trim_end_matches('.'));

                    // Case A: Month-Year (e.g. Jun-2026 or Jun-26 or 2026-Jun)
                    if digit_len == 4 || (digit_len == 2 && digit_val == default_year % 100) {
                        let year = if digit_len == 2 {
                            locale.expand_two_digit_year(digit_val)
                        } else {
                            digit_val
                        };
                        if (1..=12).contains(&month) {
                            if month_idx == 0 {
                                return Some((
                                    SimpleDate {
                                        year,
                                        month,
                                        day: 1,
                                    },
                                    DateFormat::MmmY {
                                        sep,
                                        year_len: digit_len,
                                        month_case: case,
                                        month_full: is_full,
                                    },
                                ));
                            } else {
                                return Some((
                                    SimpleDate {
                                        year,
                                        month,
                                        day: 1,
                                    },
                                    DateFormat::YMmm {
                                        sep,
                                        year_len: digit_len,
                                        month_case: case,
                                        month_full: is_full,
                                    },
                                ));
                            }
                        }
                    } else {
                        // Case B: Day-Month or Month-Day (assumes default_year)
                        let day = digit_val as u32;
                        if day >= 1 && day <= days_in_month(default_year, month) {
                            if month_idx == 1 {
                                // digit_idx is 0 (Day) -> e.g. 22-Jun
                                return Some((
                                    SimpleDate {
                                        year: default_year,
                                        month,
                                        day,
                                    },
                                    DateFormat::DMmm {
                                        sep,
                                        month_case: case,
                                        month_full: is_full,
                                    },
                                ));
                            } else {
                                // digit_idx is 1 (Day) -> e.g. Jun-22
                                return Some((
                                    SimpleDate {
                                        year: default_year,
                                        month,
                                        day,
                                    },
                                    DateFormat::MmmD {
                                        sep,
                                        month_case: case,
                                        month_full: is_full,
                                    },
                                ));
                            }
                        }
                    }
                }
            } else {
                // All 2 parts are digits (e.g. 6/22, 6/2026)
                if let (Some(val0), Some(val1)) = (parse_digits(parts[0]), parse_digits(parts[1])) {
                    let len0 = parts[0].trim().len();
                    let len1 = parts[1].trim().len();

                    // Option A: Month-Year (e.g. 6/2026)
                    if len1 == 4 {
                        let month = val0 as u32;
                        let year = val1;
                        if (1..=12).contains(&month) {
                            return Some((
                                SimpleDate {
                                    year,
                                    month,
                                    day: 1,
                                },
                                DateFormat::My { sep, year_len: 4 },
                            ));
                        }
                    } else if len0 == 4 {
                        let year = val0;
                        let month = val1 as u32;
                        if (1..=12).contains(&month) {
                            return Some((
                                SimpleDate {
                                    year,
                                    month,
                                    day: 1,
                                },
                                DateFormat::My { sep, year_len: 4 },
                            ));
                        }
                    } else if locale.date_order == DateOrder::Dmy {
                        let day = val0 as u32;
                        let month = val1 as u32;
                        if (1..=12).contains(&month)
                            && day >= 1
                            && day <= days_in_month(default_year, month)
                        {
                            return Some((
                                SimpleDate {
                                    year: default_year,
                                    month,
                                    day,
                                },
                                DateFormat::Dm { sep },
                            ));
                        }
                        if (1..=12).contains(&(val0 as u32)) && len1 == 2 && val1 > 31 {
                            let year = locale.expand_two_digit_year(val1);
                            return Some((
                                SimpleDate {
                                    year,
                                    month: val0 as u32,
                                    day: 1,
                                },
                                DateFormat::My { sep, year_len: 2 },
                            ));
                        }
                    } else {
                        // Option B: Month-Day (assumes default_year)
                        let month = val0 as u32;
                        let day = val1 as u32;
                        if (1..=12).contains(&month)
                            && day >= 1
                            && day <= days_in_month(default_year, month)
                        {
                            return Some((
                                SimpleDate {
                                    year: default_year,
                                    month,
                                    day,
                                },
                                DateFormat::Md { sep },
                            ));
                        }

                        // Option C: Month-Year with a 2-digit year that isn't a
                        // valid day (e.g. "1-34" -> Jan 1934), matching Excel's
                        // fallback when the second part can't be a day.
                        if (1..=12).contains(&month) && len1 == 2 && val1 > 31 {
                            let year = locale.expand_two_digit_year(val1);
                            return Some((
                                SimpleDate {
                                    year,
                                    month,
                                    day: 1,
                                },
                                DateFormat::My { sep, year_len: 2 },
                            ));
                        }
                    }
                }
            }
        }
    }

    // Also check space/comma-separated strings (e.g. "June 22, 2026", "22 June 2026", "22. Juni 2026")
    if src_trim.contains(' ') || src_trim.contains(',') {
        let space_parts: Vec<&str> = src_trim
            .split([' ', ','])
            .filter(|p| !p.trim().is_empty())
            .map(|p| p.trim())
            .collect();

        if space_parts.len() == 3 {
            let mut month_word_info = None;
            for (i, part) in space_parts.iter().enumerate() {
                if let Some((m, is_full)) = locale.match_month_word(part) {
                    month_word_info = Some((i, m, is_full));
                    break;
                }
            }

            if let Some((month_idx, month, is_full)) = month_word_info {
                let mut digit_parts = Vec::new();
                for (i, part) in space_parts.iter().enumerate() {
                    if i != month_idx
                        && let Some(val) = parse_digits(part)
                    {
                        digit_parts.push((i, val, part.trim_end_matches('.').len()));
                    }
                }

                if digit_parts.len() == 2 {
                    let case = detect_case(space_parts[month_idx].trim_end_matches('.'));
                    // e.g. "22 June 2026" or "22. Juni 2026" (month_idx == 1, digit_parts[0] is index 0, digit_parts[1] is index 2)
                    if month_idx == 1 && digit_parts[0].0 == 0 && digit_parts[1].0 == 2 {
                        let day = digit_parts[0].1 as u32;
                        let year_raw = digit_parts[1].1;
                        let year_len = digit_parts[1].2;
                        let year = if year_len == 2 {
                            locale.expand_two_digit_year(year_raw)
                        } else {
                            year_raw
                        };
                        if day >= 1 && day <= days_in_month(year, month) {
                            return Some((
                                SimpleDate { year, month, day },
                                DateFormat::DMmmY {
                                    sep: '-',
                                    year_len,
                                    month_case: case,
                                    month_full: is_full,
                                },
                            ));
                        }
                    }

                    // e.g. "June 22, 2026" (month_idx == 0, digit_parts[0] is index 1, digit_parts[1] is index 2)
                    if month_idx == 0 && digit_parts[0].0 == 1 && digit_parts[1].0 == 2 {
                        let day = digit_parts[0].1 as u32;
                        let year_raw = digit_parts[1].1;
                        let year_len = digit_parts[1].2;
                        let year = if year_len == 2 {
                            locale.expand_two_digit_year(year_raw)
                        } else {
                            year_raw
                        };
                        if day >= 1 && day <= days_in_month(year, month) {
                            return Some((
                                SimpleDate { year, month, day },
                                DateFormat::MmmDY {
                                    sep: '-',
                                    year_len,
                                    month_case: case,
                                    month_full: is_full,
                                },
                            ));
                        }
                    }
                }
            }
        }
    }

    None
}

/// Converts a date to Excel's day count, where 1 is 1900-01-01.
///
/// Reproduces Excel's 1900 leap-year bug -- serial 60 is the nonexistent
/// 1900-02-29 -- by adding a day for every date after 1900-02-28, which is
/// what makes serials agree with Excel's for every date a workbook is likely
/// to contain. Dates before 1900 have no serial and return `0.0`.
pub fn date_to_excel_serial(date: SimpleDate) -> f64 {
    if date.year < 1900 {
        return 0.0;
    }
    let mut days = 0;
    for y in 1900..date.year {
        days += if is_leap_year(y) { 366 } else { 365 };
    }
    for m in 1..date.month {
        days += days_in_month(date.year, m) as i32;
    }
    days += date.day as i32;
    if date.year > 1900 || (date.year == 1900 && date.month > 2) {
        days += 1;
    }
    days as f64
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Every case pins both the parsed date and the [`DateFormat`] that
    /// `parse_date` inferred.
    #[test]
    fn test_date_parsing_and_format_detection() {
        let cases: &[(&str, SimpleDate, DateFormat)] = &[
            (
                "2026-06-22",
                SimpleDate {
                    year: 2026,
                    month: 6,
                    day: 22,
                },
                DateFormat::Ymd { sep: '-' },
            ),
            (
                "2026/06/22",
                SimpleDate {
                    year: 2026,
                    month: 6,
                    day: 22,
                },
                DateFormat::Ymd { sep: '/' },
            ),
            (
                "06-22-2026",
                SimpleDate {
                    year: 2026,
                    month: 6,
                    day: 22,
                },
                DateFormat::Mdy {
                    sep: '-',
                    year_len: 4,
                },
            ),
            (
                "22-06-2026",
                SimpleDate {
                    year: 2026,
                    month: 6,
                    day: 22,
                },
                DateFormat::Dmy {
                    sep: '-',
                    year_len: 4,
                },
            ),
            (
                "06/22/26",
                SimpleDate {
                    year: 2026,
                    month: 6,
                    day: 22,
                },
                DateFormat::Mdy {
                    sep: '/',
                    year_len: 2,
                },
            ),
            // 2-digit years below the pivot roll back into the 1900s.
            (
                "06/22/99",
                SimpleDate {
                    year: 1999,
                    month: 6,
                    day: 22,
                },
                DateFormat::Mdy {
                    sep: '/',
                    year_len: 2,
                },
            ),
            (
                "22-Jun-2026",
                SimpleDate {
                    year: 2026,
                    month: 6,
                    day: 22,
                },
                DateFormat::DMmmY {
                    sep: '-',
                    year_len: 4,
                    month_case: StringCase::Title,
                    month_full: false,
                },
            ),
            (
                "22-June-2026",
                SimpleDate {
                    year: 2026,
                    month: 6,
                    day: 22,
                },
                DateFormat::DMmmY {
                    sep: '-',
                    year_len: 4,
                    month_case: StringCase::Title,
                    month_full: true,
                },
            ),
            (
                "Jun-22-2026",
                SimpleDate {
                    year: 2026,
                    month: 6,
                    day: 22,
                },
                DateFormat::MmmDY {
                    sep: '-',
                    year_len: 4,
                    month_case: StringCase::Title,
                    month_full: false,
                },
            ),
            // 2-part forms infer the missing component.
            (
                "6/22",
                SimpleDate {
                    year: 2026,
                    month: 6,
                    day: 22,
                },
                DateFormat::Md { sep: '/' },
            ),
            (
                "22-Jun",
                SimpleDate {
                    year: 2026,
                    month: 6,
                    day: 22,
                },
                DateFormat::DMmm {
                    sep: '-',
                    month_case: StringCase::Title,
                    month_full: false,
                },
            ),
            (
                "Jun-22",
                SimpleDate {
                    year: 2026,
                    month: 6,
                    day: 22,
                },
                DateFormat::MmmD {
                    sep: '-',
                    month_case: StringCase::Title,
                    month_full: false,
                },
            ),
            (
                "6/2026",
                SimpleDate {
                    year: 2026,
                    month: 6,
                    day: 1,
                },
                DateFormat::My {
                    sep: '/',
                    year_len: 4,
                },
            ),
            (
                "Jun-26",
                SimpleDate {
                    year: 2026,
                    month: 6,
                    day: 1,
                },
                DateFormat::MmmY {
                    sep: '-',
                    year_len: 2,
                    month_case: StringCase::Title,
                    month_full: false,
                },
            ),
            (
                "2026-Jun",
                SimpleDate {
                    year: 2026,
                    month: 6,
                    day: 1,
                },
                DateFormat::YMmm {
                    sep: '-',
                    year_len: 4,
                    month_case: StringCase::Title,
                    month_full: false,
                },
            ),
        ];

        for (src, want_date, want_format) in cases {
            let (date, format) = parse_date(src).unwrap_or_else(|| panic!("{src} did not parse"));
            assert_eq!(date, *want_date, "date mismatch for {src}");
            assert_eq!(format, *want_format, "format mismatch for {src}");
        }
    }

    #[test]
    fn test_locale_aware_date_parsing() {
        let de = Locale::de_de();
        let gb = Locale::en_gb();
        let us = Locale::en_us();

        // In German locale: 22.06.2026 is Day=22, Month=6
        let (date, fmt) = parse_date_with_locale("22.06.2026", &de).unwrap();
        assert_eq!(
            date,
            SimpleDate {
                year: 2026,
                month: 6,
                day: 22
            }
        );
        assert_eq!(
            fmt,
            DateFormat::Dmy {
                sep: '.',
                year_len: 4
            }
        );

        // German month word: 22. Juni 2026
        let (date, _) = parse_date_with_locale("22. Juni 2026", &de).unwrap();
        assert_eq!(
            date,
            SimpleDate {
                year: 2026,
                month: 6,
                day: 22
            }
        );

        // In UK locale: 05/06/2026 is 5th June 2026
        let (date, _) = parse_date_with_locale("05/06/2026", &gb).unwrap();
        assert_eq!(
            date,
            SimpleDate {
                year: 2026,
                month: 6,
                day: 5
            }
        );

        // In US locale: 05/06/2026 is May 6th 2026
        let (date, _) = parse_date_with_locale("05/06/2026", &us).unwrap();
        assert_eq!(
            date,
            SimpleDate {
                year: 2026,
                month: 5,
                day: 6
            }
        );

        // In UK/DE locale: 2-part "22/6" is Day 22, Month 6
        let (date, fmt) = parse_date_with_locale("22/6", &gb).unwrap();
        assert_eq!(
            date,
            SimpleDate {
                year: 2026,
                month: 6,
                day: 22
            }
        );
        assert_eq!(fmt, DateFormat::Dm { sep: '/' });

        // French month name: "14 juillet 2026"
        let fr = Locale::fr_fr();
        let (date, _) = parse_date_with_locale("14 juillet 2026", &fr).unwrap();
        assert_eq!(
            date,
            SimpleDate {
                year: 2026,
                month: 7,
                day: 14
            }
        );

        // Spanish month name: "12 de octubre de 2026" or "12 octubre 2026"
        let es = Locale::es_es();
        let (date, _) = parse_date_with_locale("12 octubre 2026", &es).unwrap();
        assert_eq!(
            date,
            SimpleDate {
                year: 2026,
                month: 10,
                day: 12
            }
        );
    }

    /// The point of detecting a format at all: a date echoes back in the
    /// notation it was typed in.
    #[test]
    fn test_format_date_round_trips_the_typed_notation() {
        let sources = [
            "2026-06-22",
            "2026/06/22",
            "6/22/26",
            "22-Jun-2026",
            "22-June-2026",
            "Jun-22-2026",
            "22-Jun",
            "Jun-22",
            "6/2026",
            "Jun-26",
            "2026-Jun",
        ];
        for src in sources {
            let (date, format) = parse_date(src).unwrap_or_else(|| panic!("{src} did not parse"));
            assert_eq!(
                format_date(date, &format),
                src,
                "round trip failed for {src}"
            );
        }
    }

    #[test]
    fn test_format_date_normalizes_zero_padding() {
        let (date, format) = parse_date("22-06-2026").unwrap();
        assert_eq!(format_date(date, &format), "22-6-2026");

        let (date, format) = parse_date("06/22/2026").unwrap();
        assert_eq!(format_date(date, &format), "6/22/2026");

        // Year-first keeps its padding: the code really is yyyy-mm-dd.
        let (date, format) = parse_date("2026-06-22").unwrap();
        assert_eq!(format_date(date, &format), "2026-06-22");
    }

    #[test]
    fn test_format_date_preserves_month_name_case() {
        for src in ["22-JUN-2026", "22-jun-2026"] {
            let (date, format) = parse_date(src).unwrap();
            assert_eq!(format_date(date, &format), src);
        }
        let (date, format) = parse_date("22-JUN-2026").unwrap();
        assert_eq!(format.to_format_code(), "d-mmm-yyyy");
        assert_eq!(
            render_date_code(date, &format.to_format_code(), StringCase::Title),
            "22-Jun-2026"
        );
    }

    #[test]
    fn test_render_date_code_does_not_rescan_substituted_month_names() {
        let dec = SimpleDate {
            year: 2026,
            month: 12,
            day: 5,
        };
        assert_eq!(
            render_date_code(dec, "mmmm d, yyyy", StringCase::Title),
            "December 5, 2026"
        );
        let may = SimpleDate {
            year: 2026,
            month: 5,
            day: 5,
        };
        assert_eq!(render_date_code(may, "mmm-yy", StringCase::Title), "May-26");
    }

    #[test]
    fn test_render_date_code_token_widths() {
        let d = SimpleDate {
            year: 2026,
            month: 6,
            day: 7,
        };
        assert_eq!(
            render_date_code(d, "yyyy-mm-dd", StringCase::Title),
            "2026-06-07"
        );
        assert_eq!(render_date_code(d, "m/d/yy", StringCase::Title), "6/7/26");
        assert_eq!(render_date_code(d, "mmmm", StringCase::Title), "June");
        // Non-token characters pass through untouched.
        assert_eq!(
            render_date_code(d, "[yyyy] week of d", StringCase::Title),
            "[2026] week of 7"
        );
    }

    #[test]
    fn test_render_date_code_honors_excel_escape_prefix() {
        let d = SimpleDate {
            year: 2026,
            month: 6,
            day: 7,
        };
        assert_eq!(
            render_date_code(d, "yyyy\\-mm\\-dd", StringCase::Title),
            "2026-06-07"
        );
        assert_eq!(
            render_date_code(d, "d\\-mmm\\-yyyy", StringCase::Title),
            "7-Jun-2026"
        );
    }

    #[test]
    fn test_is_date_code_rejects_numeric_formats() {
        assert!(is_date_code("m/d/yy"));
        assert!(is_date_code("yyyy-mm-dd"));
        assert!(!is_date_code("0.00"));
        assert!(!is_date_code("#,##0"));
        assert!(!is_date_code(""));
    }

    #[test]
    fn test_invalid_dates_do_not_parse() {
        assert!(parse_date("2026-02-30").is_none());
        assert!(parse_date("2025-02-29").is_none()); // non-leap year
        assert!(parse_date("13/22/2026").is_none()); // invalid month
        assert!(parse_date("06-32-2026").is_none()); // invalid day
    }

    #[test]
    fn test_date_to_excel_serial() {
        // Excel's epoch: 1900-01-01 is serial 1.
        assert_eq!(
            date_to_excel_serial(SimpleDate {
                year: 1900,
                month: 1,
                day: 1
            }),
            1.0
        );
        // Excel's deliberate 1900 leap-year bug means 1900-03-01 is 61, not 60.
        assert_eq!(
            date_to_excel_serial(SimpleDate {
                year: 1900,
                month: 3,
                day: 1
            }),
            61.0
        );
        assert_eq!(
            date_to_excel_serial(SimpleDate {
                year: 2026,
                month: 6,
                day: 22
            }),
            46195.0
        );
    }
}
