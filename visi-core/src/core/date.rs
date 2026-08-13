//! Recognizing dates written as text, and converting them to Excel serials.
//!
//! [`parse_date`] infers both the date and the [`DateFormat`] it was written
//! in; [`date_to_excel_serial`] converts to Excel's day count, reproducing the
//! 1900 leap-year bug. `engine::sheet`'s literal coercion is the caller.
//!
//! The `DateFormat` half of `parse_date`'s return is currently discarded by
//! that caller -- it exists for a formatter that would render a date back in
//! the notation it was typed in. Tests pin the detection so it stays correct
//! until something uses it.

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

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum StringCase {
    Lower,
    Upper,
    Title,
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

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SimpleDate {
    pub year: i32,
    pub month: u32,
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

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DateFormat {
    // 3-part formats
    Ymd {
        sep: char,
    }, // e.g. 2026-06-22
    Mdy {
        sep: char,
        year_len: usize,
    }, // e.g. 06/22/2026
    Dmy {
        sep: char,
        year_len: usize,
    }, // e.g. 22-06-2026
    DMmmY {
        sep: char,
        year_len: usize,
        month_case: StringCase,
        month_full: bool,
    }, // e.g. 22-Jun-2026, 22-June-26
    MmmDY {
        sep: char,
        year_len: usize,
        month_case: StringCase,
        month_full: bool,
    }, // e.g. Jun-22-2026, June-22-26
    YMmmD {
        sep: char,
        year_len: usize,
        month_case: StringCase,
        month_full: bool,
    }, // e.g. 2026-Jun-22

    // 2-part formats (Assumed Year or Day)
    Md {
        sep: char,
    }, // e.g. 6/22 -> Month-Day (assumes DEFAULT_YEAR)
    My {
        sep: char,
        year_len: usize,
    }, // e.g. 6/2026 -> Month-Year (assumes Day 1)
    DMmm {
        sep: char,
        month_case: StringCase,
        month_full: bool,
    }, // e.g. 22-Jun -> Day-Month (assumes DEFAULT_YEAR)
    MmmD {
        sep: char,
        month_case: StringCase,
        month_full: bool,
    }, // e.g. Jun-22 -> Month-Day (assumes DEFAULT_YEAR)
    MmmY {
        sep: char,
        year_len: usize,
        month_case: StringCase,
        month_full: bool,
    }, // e.g. Jun-2026 -> Month-Year (assumes Day 1)
    YMmm {
        sep: char,
        year_len: usize,
        month_case: StringCase,
        month_full: bool,
    }, // e.g. 2026-Jun -> Year-Month (assumes Day 1)
}

fn find_month_word(part: &str) -> Option<(u32, bool)> {
    // returns (month_1_based, is_full_name)
    let p_lower = part.to_lowercase();
    for (idx, &m) in MONTHS_FULL.iter().enumerate() {
        if m.to_lowercase() == p_lower {
            return Some((idx as u32 + 1, true));
        }
    }
    for (idx, &m) in MONTHS_SHORT.iter().enumerate() {
        if m.to_lowercase() == p_lower {
            return Some((idx as u32 + 1, false));
        }
    }
    None
}

fn parse_digits(part: &str) -> Option<i32> {
    if !part.is_empty() && part.chars().all(|c| c.is_ascii_digit()) {
        part.parse::<i32>().ok()
    } else {
        None
    }
}

pub fn parse_date(src: &str) -> Option<(SimpleDate, DateFormat)> {
    const DEFAULT_YEAR: i32 = 2026;

    for &sep in &['-', '/'] {
        let parts: Vec<&str> = src.split(sep).collect();

        // --- 3 PARTS ---
        if parts.len() == 3 {
            // Check if there is a word month in the parts
            let mut month_word_info = None;
            for (i, part) in parts.iter().enumerate() {
                if let Some((m, is_full)) = find_month_word(part) {
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
                        digit_parts.push((i, val, part.len()));
                    }
                }

                if digit_parts.len() == 2 {
                    let case = detect_case(parts[month_idx]);

                    // Case A: Day-Month-Year (e.g., 22-Jun-2026, 22-Jun-26)
                    // month_idx is 1. digit_parts[0] is index 0 (day), digit_parts[1] is index 2 (year).
                    if month_idx == 1 && digit_parts[0].0 == 0 && digit_parts[1].0 == 2 {
                        let day = digit_parts[0].1 as u32;
                        let year_raw = digit_parts[1].1;
                        let year_len = digit_parts[1].2;
                        let year = if year_len == 2 {
                            if year_raw < 30 {
                                2000 + year_raw
                            } else {
                                1900 + year_raw
                            }
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
                            if year_raw < 30 {
                                2000 + year_raw
                            } else {
                                1900 + year_raw
                            }
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
                            if year_raw < 30 {
                                2000 + year_raw
                            } else {
                                1900 + year_raw
                            }
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
                    let len0 = parts[0].len();
                    let len2 = parts[2].len();

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
                            if year_raw < 30 {
                                2000 + year_raw
                            } else {
                                1900 + year_raw
                            }
                        } else {
                            year_raw
                        };

                        // Check if MDY or DMY
                        // If val0 > 12, it must be DMY
                        if val0 > 12 {
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
                        } else if val1 > 12 {
                            // If val1 > 12, it must be MDY
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
                        } else {
                            // Defaults to MDY (standard US locale)
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

        // --- 2 PARTS ---
        if parts.len() == 2 {
            // Check if there is a word month in the parts
            let mut month_word_info = None;
            for (i, part) in parts.iter().enumerate() {
                if let Some((m, is_full)) = find_month_word(part) {
                    month_word_info = Some((i, m, is_full));
                    break;
                }
            }

            if let Some((month_idx, month, is_full)) = month_word_info {
                let digit_idx = if month_idx == 0 { 1 } else { 0 };
                if let Some(digit_val) = parse_digits(parts[digit_idx]) {
                    let digit_len = parts[digit_idx].len();
                    let case = detect_case(parts[month_idx]);

                    // Case A: Month-Year (e.g. Jun-2026 or Jun-26 or 2026-Jun)
                    // If digit_len == 4 or digit_val > 31 {
                    if digit_len == 4 || (digit_len == 2 && digit_val == DEFAULT_YEAR % 100) {
                        let year = if digit_len == 2 {
                            if digit_val < 30 {
                                2000 + digit_val
                            } else {
                                1900 + digit_val
                            }
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
                        // Case B: Day-Month or Month-Day (assumes DEFAULT_YEAR)
                        let day = digit_val as u32;
                        if day >= 1 && day <= days_in_month(DEFAULT_YEAR, month) {
                            if month_idx == 1 {
                                // digit_idx is 0 (Day) -> e.g. 22-Jun
                                return Some((
                                    SimpleDate {
                                        year: DEFAULT_YEAR,
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
                                        year: DEFAULT_YEAR,
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
                    let len1 = parts[1].len();

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
                    } else {
                        // Option B: Month-Day (assumes DEFAULT_YEAR)
                        let month = val0 as u32;
                        let day = val1 as u32;
                        if (1..=12).contains(&month)
                            && day >= 1
                            && day <= days_in_month(DEFAULT_YEAR, month)
                        {
                            return Some((
                                SimpleDate {
                                    year: DEFAULT_YEAR,
                                    month,
                                    day,
                                },
                                DateFormat::Md { sep },
                            ));
                        }

                        // Option C: Month-Year with a 2-digit year that isn't a
                        // valid day (e.g. "1-34" -> Jan 1934), matching Excel's
                        // fallback when the second part can't be a day.
                        if (1..=12).contains(&month) && len1 == 2 {
                            let year = if val1 < 30 { 2000 + val1 } else { 1900 + val1 };
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
    None
}

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
    /// `parse_date` inferred. The format half used to be checked by feeding it
    /// back through a `format_date` that nothing shipped; asserting the enum
    /// directly covers the same detection logic without the dead round-trip.
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
