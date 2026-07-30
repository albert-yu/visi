const WEEKDAYS_FULL: [&str; 7] = [
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
    "Sunday",
];
const WEEKDAYS_SHORT: [&str; 7] = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];

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
        if let Some(first) = chars.next() {
            if first.is_uppercase() && chars.all(|c| c.is_lowercase()) {
                return StringCase::Title;
            }
        }
        StringCase::Original
    }
}

pub fn apply_case(s: &str, case: StringCase) -> String {
    match case {
        StringCase::Lower => s.to_lowercase(),
        StringCase::Upper => s.to_uppercase(),
        StringCase::Title => {
            let mut chars = s.chars();
            match chars.next() {
                None => String::new(),
                Some(f) => f.to_uppercase().collect::<String>() + &chars.as_str().to_lowercase(),
            }
        }
        StringCase::Original => s.to_string(),
    }
}

pub fn try_fill_weekday(src: &str, offset: i32) -> Option<String> {
    let src_lower = src.to_lowercase();
    for (idx, &day) in WEEKDAYS_FULL.iter().enumerate() {
        if day.to_lowercase() == src_lower {
            let new_idx = ((idx as i32 + offset) % 7 + 7) % 7;
            let case = detect_case(src);
            return Some(apply_case(WEEKDAYS_FULL[new_idx as usize], case));
        }
    }
    for (idx, &day) in WEEKDAYS_SHORT.iter().enumerate() {
        if day.to_lowercase() == src_lower {
            let new_idx = ((idx as i32 + offset) % 7 + 7) % 7;
            let case = detect_case(src);
            return Some(apply_case(WEEKDAYS_SHORT[new_idx as usize], case));
        }
    }
    None
}

pub fn try_fill_month(src: &str, offset: i32) -> Option<String> {
    let src_lower = src.to_lowercase();
    for (idx, &month) in MONTHS_FULL.iter().enumerate() {
        if month.to_lowercase() == src_lower {
            let new_idx = ((idx as i32 + offset) % 12 + 12) % 12;
            let case = detect_case(src);
            return Some(apply_case(MONTHS_FULL[new_idx as usize], case));
        }
    }
    for (idx, &month) in MONTHS_SHORT.iter().enumerate() {
        if month.to_lowercase() == src_lower {
            let new_idx = ((idx as i32 + offset) % 12 + 12) % 12;
            let case = detect_case(src);
            return Some(apply_case(MONTHS_SHORT[new_idx as usize], case));
        }
    }
    None
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

pub fn add_days(mut date: SimpleDate, mut days: i32) -> SimpleDate {
    if days == 0 {
        return date;
    }
    if days > 0 {
        while days > 0 {
            let dim = days_in_month(date.year, date.month);
            let days_left_in_month = dim - date.day + 1;
            if days < days_left_in_month as i32 {
                date.day += days as u32;
                break;
            } else {
                days -= days_left_in_month as i32;
                date.day = 1;
                date.month += 1;
                if date.month > 12 {
                    date.month = 1;
                    date.year += 1;
                }
            }
        }
    } else {
        let mut days_to_sub = -days;
        while days_to_sub > 0 {
            let days_available_in_month = date.day;
            if days_to_sub < days_available_in_month as i32 {
                date.day -= days_to_sub as u32;
                break;
            } else {
                days_to_sub -= days_available_in_month as i32;
                date.month -= 1;
                if date.month == 0 {
                    date.month = 12;
                    date.year -= 1;
                }
                date.day = days_in_month(date.year, date.month);
            }
        }
    }
    date
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

pub fn add_months(date: SimpleDate, months: i32) -> SimpleDate {
    let total_months = date.year * 12 + (date.month - 1) as i32 + months;
    let mut new_year = total_months / 12;
    let mut new_month = total_months % 12 + 1;
    if new_month <= 0 {
        new_month += 12;
        new_year -= 1;
    }
    SimpleDate {
        year: new_year,
        month: new_month as u32,
        day: 1, // Month-based formats always assume Day 1
    }
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
                    if i != month_idx {
                        if let Some(val) = parse_digits(part) {
                            digit_parts.push((i, val, part.len()));
                        }
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
                        if month >= 1
                            && month <= 12
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
                            if month >= 1
                                && month <= 12
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
                            if month >= 1
                                && month <= 12
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
                            if month >= 1
                                && month <= 12
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
                        if month >= 1 && month <= 12 {
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
                        if month >= 1 && month <= 12 {
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
                        if month >= 1
                            && month <= 12
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
                        if month >= 1 && month <= 12 && len1 == 2 {
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

pub fn format_date(date: SimpleDate, format: &DateFormat) -> String {
    match format {
        DateFormat::Ymd { sep } => {
            format!(
                "{:04}{}{:02}{}{:02}",
                date.year, sep, date.month, sep, date.day
            )
        }
        DateFormat::Mdy { sep, year_len } => {
            let formatted_year = if *year_len == 2 {
                format!("{:02}", date.year % 100)
            } else {
                format!("{:04}", date.year)
            };
            format!(
                "{:02}{}{:02}{}{}",
                date.month, sep, date.day, sep, formatted_year
            )
        }
        DateFormat::Dmy { sep, year_len } => {
            let formatted_year = if *year_len == 2 {
                format!("{:02}", date.year % 100)
            } else {
                format!("{:04}", date.year)
            };
            format!(
                "{:02}{}{:02}{}{}",
                date.day, sep, date.month, sep, formatted_year
            )
        }
        DateFormat::DMmmY {
            sep,
            year_len,
            month_case,
            month_full,
        } => {
            let formatted_year = if *year_len == 2 {
                format!("{:02}", date.year % 100)
            } else {
                format!("{:04}", date.year)
            };
            let month_list = if *month_full {
                &MONTHS_FULL
            } else {
                &MONTHS_SHORT
            };
            let month_name = month_list[(date.month - 1) as usize];
            let formatted_month = apply_case(month_name, *month_case);
            format!(
                "{:02}{}{}{}{}",
                date.day, sep, formatted_month, sep, formatted_year
            )
        }
        DateFormat::MmmDY {
            sep,
            year_len,
            month_case,
            month_full,
        } => {
            let formatted_year = if *year_len == 2 {
                format!("{:02}", date.year % 100)
            } else {
                format!("{:04}", date.year)
            };
            let month_list = if *month_full {
                &MONTHS_FULL
            } else {
                &MONTHS_SHORT
            };
            let month_name = month_list[(date.month - 1) as usize];
            let formatted_month = apply_case(month_name, *month_case);
            format!(
                "{}{}{:02}{}{}",
                formatted_month, sep, date.day, sep, formatted_year
            )
        }
        DateFormat::YMmmD {
            sep,
            year_len,
            month_case,
            month_full,
        } => {
            let formatted_year = if *year_len == 2 {
                format!("{:02}", date.year % 100)
            } else {
                format!("{:04}", date.year)
            };
            let month_list = if *month_full {
                &MONTHS_FULL
            } else {
                &MONTHS_SHORT
            };
            let month_name = month_list[(date.month - 1) as usize];
            let formatted_month = apply_case(month_name, *month_case);
            format!(
                "{}{}{}{}{:02}",
                formatted_year, sep, formatted_month, sep, date.day
            )
        }
        DateFormat::Md { sep } => {
            format!("{:02}{}{:02}", date.month, sep, date.day)
        }
        DateFormat::My { sep, year_len } => {
            let formatted_year = if *year_len == 2 {
                format!("{:02}", date.year % 100)
            } else {
                format!("{:04}", date.year)
            };
            format!("{:02}{}{}", date.month, sep, formatted_year)
        }
        DateFormat::DMmm {
            sep,
            month_case,
            month_full,
        } => {
            let month_list = if *month_full {
                &MONTHS_FULL
            } else {
                &MONTHS_SHORT
            };
            let month_name = month_list[(date.month - 1) as usize];
            let formatted_month = apply_case(month_name, *month_case);
            format!("{:02}{}{}", date.day, sep, formatted_month)
        }
        DateFormat::MmmD {
            sep,
            month_case,
            month_full,
        } => {
            let month_list = if *month_full {
                &MONTHS_FULL
            } else {
                &MONTHS_SHORT
            };
            let month_name = month_list[(date.month - 1) as usize];
            let formatted_month = apply_case(month_name, *month_case);
            format!("{}{}{:02}", formatted_month, sep, date.day)
        }
        DateFormat::MmmY {
            sep,
            year_len,
            month_case,
            month_full,
        } => {
            let formatted_year = if *year_len == 2 {
                format!("{:02}", date.year % 100)
            } else {
                format!("{:04}", date.year)
            };
            let month_list = if *month_full {
                &MONTHS_FULL
            } else {
                &MONTHS_SHORT
            };
            let month_name = month_list[(date.month - 1) as usize];
            let formatted_month = apply_case(month_name, *month_case);
            format!("{}{}{}", formatted_month, sep, formatted_year)
        }
        DateFormat::YMmm {
            sep,
            year_len,
            month_case,
            month_full,
        } => {
            let formatted_year = if *year_len == 2 {
                format!("{:02}", date.year % 100)
            } else {
                format!("{:04}", date.year)
            };
            let month_list = if *month_full {
                &MONTHS_FULL
            } else {
                &MONTHS_SHORT
            };
            let month_name = month_list[(date.month - 1) as usize];
            let formatted_month = apply_case(month_name, *month_case);
            format!("{}{}{}", formatted_year, sep, formatted_month)
        }
    }
}

pub fn try_fill_date(src: &str, offset: i32) -> Option<String> {
    let (date, format) = parse_date(src)?;

    let is_month_based = match &format {
        DateFormat::My { .. } | DateFormat::MmmY { .. } | DateFormat::YMmm { .. } => true,
        _ => false,
    };

    let new_date = if is_month_based {
        add_months(date, offset)
    } else {
        add_days(date, offset)
    };

    Some(format_date(new_date, &format))
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

pub fn date_to_days(date: SimpleDate) -> i32 {
    let mut days = 0;
    // Years
    for y in 1..date.year {
        days += if is_leap_year(y) { 366 } else { 365 };
    }
    // Months
    for m in 1..date.month {
        days += days_in_month(date.year, m) as i32;
    }
    // Days
    days += date.day as i32;
    days
}

pub fn days_to_date(mut days: i32) -> SimpleDate {
    let mut year = 1;
    loop {
        let y_days = if is_leap_year(year) { 366 } else { 365 };
        if days <= y_days {
            break;
        }
        days -= y_days;
        year += 1;
    }

    let mut month = 1;
    loop {
        let m_days = days_in_month(year, month) as i32;
        if days <= m_days {
            break;
        }
        days -= m_days;
        month += 1;
    }

    SimpleDate {
        year,
        month,
        day: days as u32,
    }
}

pub fn parse_weekday(s: &str) -> Option<(usize, bool)> {
    let trimmed = s.trim().to_lowercase();
    for (i, w) in WEEKDAYS_FULL.iter().enumerate() {
        if w.to_lowercase() == trimmed {
            return Some((i, false));
        }
    }
    for (i, w) in WEEKDAYS_SHORT.iter().enumerate() {
        if w.to_lowercase() == trimmed {
            return Some((i, true));
        }
    }
    None
}

pub fn detect_weekday_pattern(
    vals: &[String],
    target_idx: usize,
    case_type: StringCase,
) -> Option<String> {
    if vals.len() < 2 {
        return None;
    }
    let mut idxs = Vec::new();
    let mut is_short = true;
    for s in vals {
        let (idx, short) = parse_weekday(s)?;
        idxs.push(idx as isize);
        is_short = short;
    }

    let mut diffs = Vec::new();
    for w in idxs.windows(2) {
        diffs.push((w[1] - w[0]).rem_euclid(7));
    }
    if diffs.is_empty() || !diffs.windows(2).all(|w| w[0] == w[1]) {
        return None;
    }
    let step = diffs[0];
    let next_idx = (idxs[0] + target_idx as isize * step).rem_euclid(7) as usize;

    let weekday_list = if is_short {
        &WEEKDAYS_SHORT
    } else {
        &WEEKDAYS_FULL
    };
    let mut result = weekday_list[next_idx].to_string();

    match case_type {
        StringCase::Upper => {
            result = result.to_uppercase();
        }
        StringCase::Lower => {
            result = result.to_lowercase();
        }
        _ => {}
    }
    Some(result)
}

pub fn parse_month(s: &str) -> Option<(usize, bool)> {
    let trimmed = s.trim().to_lowercase();
    for (i, m) in MONTHS_FULL.iter().enumerate() {
        if m.to_lowercase() == trimmed {
            return Some((i, false));
        }
    }
    for (i, m) in MONTHS_SHORT.iter().enumerate() {
        if m.to_lowercase() == trimmed {
            return Some((i, true));
        }
    }
    None
}

pub fn detect_month_pattern(
    vals: &[String],
    target_idx: usize,
    case_type: StringCase,
) -> Option<String> {
    if vals.len() < 2 {
        return None;
    }
    let mut idxs = Vec::new();
    let mut is_short = true;
    for s in vals {
        let (idx, short) = parse_month(s)?;
        idxs.push(idx as isize);
        is_short = short;
    }

    let mut diffs = Vec::new();
    for w in idxs.windows(2) {
        diffs.push((w[1] - w[0]).rem_euclid(12));
    }
    if diffs.is_empty() || !diffs.windows(2).all(|w| w[0] == w[1]) {
        return None;
    }
    let step = diffs[0];
    let next_idx = (idxs[0] + target_idx as isize * step).rem_euclid(12) as usize;

    let month_list = if is_short {
        &MONTHS_SHORT
    } else {
        &MONTHS_FULL
    };
    let mut result = month_list[next_idx].to_string();

    match case_type {
        StringCase::Upper => {
            result = result.to_uppercase();
        }
        StringCase::Lower => {
            result = result.to_lowercase();
        }
        _ => {}
    }
    Some(result)
}

pub fn date_to_months(date: SimpleDate) -> i32 {
    (date.year - 1) * 12 + (date.month - 1) as i32
}

pub fn months_to_date(months: i32, day: u32) -> SimpleDate {
    let year = months / 12 + 1;
    let month = (months % 12) + 1;
    SimpleDate {
        year,
        month: month as u32,
        day: day.min(days_in_month(year, month as u32)),
    }
}

pub fn detect_date_pattern(vals: &[String], target_idx: usize) -> Option<String> {
    if vals.len() < 2 {
        return None;
    }
    let mut dates = Vec::new();
    let mut formats = Vec::new();
    for s in vals {
        let (date, format) = parse_date(s)?;
        dates.push(date);
        formats.push(format);
    }

    let is_month_based = match &formats[0] {
        DateFormat::My { .. } | DateFormat::MmmY { .. } | DateFormat::YMmm { .. } => true,
        _ => false,
    };

    let next_date = if is_month_based {
        let months: Vec<i32> = dates.iter().map(|d| date_to_months(*d)).collect();
        let step_months = (months[months.len() - 1] - months[0]) as f64 / (months.len() - 1) as f64;
        let next_months = (months[0] as f64 + target_idx as f64 * step_months).round() as i32;
        months_to_date(next_months, dates[0].day)
    } else {
        let days: Vec<i32> = dates.iter().map(|d| date_to_days(*d)).collect();
        let step_days = (days[days.len() - 1] - days[0]) as f64 / (days.len() - 1) as f64;
        let next_days = (days[0] as f64 + target_idx as f64 * step_days).round() as i32;
        days_to_date(next_days)
    };

    Some(format_date(next_date, &formats[0]))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_date_parsing_and_formatting() {
        // Test YMD
        let parsed = parse_date("2026-06-22");
        assert!(parsed.is_some());
        let (date, format) = parsed.unwrap();
        assert_eq!(
            date,
            SimpleDate {
                year: 2026,
                month: 6,
                day: 22
            }
        );
        assert_eq!(format_date(date, &format), "2026-06-22");

        let parsed_slash = parse_date("2026/06/22");
        assert!(parsed_slash.is_some());
        let (date, format) = parsed_slash.unwrap();
        assert_eq!(
            date,
            SimpleDate {
                year: 2026,
                month: 6,
                day: 22
            }
        );
        assert_eq!(format_date(date, &format), "2026/06/22");

        // Test MDY
        let parsed_mdy = parse_date("06-22-2026");
        assert!(parsed_mdy.is_some());
        let (date, format) = parsed_mdy.unwrap();
        assert_eq!(
            date,
            SimpleDate {
                year: 2026,
                month: 6,
                day: 22
            }
        );
        assert_eq!(format_date(date, &format), "06-22-2026");

        // Test DMY (unambiguous day > 12)
        let parsed_dmy = parse_date("22-06-2026");
        assert!(parsed_dmy.is_some());
        let (date, format) = parsed_dmy.unwrap();
        assert_eq!(
            date,
            SimpleDate {
                year: 2026,
                month: 6,
                day: 22
            }
        );
        assert_eq!(format_date(date, &format), "22-06-2026");

        // Test 2-digit years MDY
        let parsed_2digit = parse_date("06/22/26");
        assert!(parsed_2digit.is_some());
        let (date, format) = parsed_2digit.unwrap();
        assert_eq!(
            date,
            SimpleDate {
                year: 2026,
                month: 6,
                day: 22
            }
        );
        assert_eq!(format_date(date, &format), "06/22/26");

        // Test 2-digit years older century
        let parsed_2digit_old = parse_date("06/22/99");
        assert!(parsed_2digit_old.is_some());
        let (date, format) = parsed_2digit_old.unwrap();
        assert_eq!(
            date,
            SimpleDate {
                year: 1999,
                month: 6,
                day: 22
            }
        );
        assert_eq!(format_date(date, &format), "06/22/99");

        // Test DMmmY (word month)
        let parsed_word_month = parse_date("22-Jun-2026");
        assert!(parsed_word_month.is_some());
        let (date, format) = parsed_word_month.unwrap();
        assert_eq!(
            date,
            SimpleDate {
                year: 2026,
                month: 6,
                day: 22
            }
        );
        assert_eq!(format_date(date, &format), "22-Jun-2026");

        let parsed_full_month = parse_date("22-June-2026");
        assert!(parsed_full_month.is_some());
        let (date, format) = parsed_full_month.unwrap();
        assert_eq!(
            date,
            SimpleDate {
                year: 2026,
                month: 6,
                day: 22
            }
        );
        assert_eq!(format_date(date, &format), "22-June-2026");

        // Test MmmDY (month first word format)
        let parsed_month_first = parse_date("Jun-22-2026");
        assert!(parsed_month_first.is_some());
        let (date, format) = parsed_month_first.unwrap();
        assert_eq!(
            date,
            SimpleDate {
                year: 2026,
                month: 6,
                day: 22
            }
        );
        assert_eq!(format_date(date, &format), "Jun-22-2026");

        // Test 2-part Month-Day (assumed year 2026)
        let parsed_2part_md = parse_date("6/22");
        assert!(parsed_2part_md.is_some());
        let (date, format) = parsed_2part_md.unwrap();
        assert_eq!(
            date,
            SimpleDate {
                year: 2026,
                month: 6,
                day: 22
            }
        );
        assert_eq!(format_date(date, &format), "06/22");

        let parsed_2part_dmmm = parse_date("22-Jun");
        assert!(parsed_2part_dmmm.is_some());
        let (date, format) = parsed_2part_dmmm.unwrap();
        assert_eq!(
            date,
            SimpleDate {
                year: 2026,
                month: 6,
                day: 22
            }
        );
        assert_eq!(format_date(date, &format), "22-Jun");

        let parsed_2part_mmmd = parse_date("Jun-22");
        assert!(parsed_2part_mmmd.is_some());
        let (date, format) = parsed_2part_mmmd.unwrap();
        assert_eq!(
            date,
            SimpleDate {
                year: 2026,
                month: 6,
                day: 22
            }
        );
        assert_eq!(format_date(date, &format), "Jun-22");

        // Test 2-part Month-Year (assumed day 1)
        let parsed_2part_my = parse_date("6/2026");
        assert!(parsed_2part_my.is_some());
        let (date, format) = parsed_2part_my.unwrap();
        assert_eq!(
            date,
            SimpleDate {
                year: 2026,
                month: 6,
                day: 1
            }
        );
        assert_eq!(format_date(date, &format), "06/2026");

        let parsed_2part_mmmy = parse_date("Jun-26");
        assert!(parsed_2part_mmmy.is_some());
        let (date, format) = parsed_2part_mmmy.unwrap();
        assert_eq!(
            date,
            SimpleDate {
                year: 2026,
                month: 6,
                day: 1
            }
        );
        assert_eq!(format_date(date, &format), "Jun-26");

        let parsed_2part_ymmm = parse_date("2026-Jun");
        assert!(parsed_2part_ymmm.is_some());
        let (date, format) = parsed_2part_ymmm.unwrap();
        assert_eq!(
            date,
            SimpleDate {
                year: 2026,
                month: 6,
                day: 1
            }
        );
        assert_eq!(format_date(date, &format), "2026-Jun");

        // Test invalid dates (should fail validation)
        assert!(parse_date("2026-02-30").is_none());
        assert!(parse_date("2025-02-29").is_none()); // non-leap year
        assert!(parse_date("13/22/2026").is_none()); // invalid month
        assert!(parse_date("06-32-2026").is_none()); // invalid day
    }

    #[test]
    fn test_date_addition() {
        let date = SimpleDate {
            year: 2026,
            month: 6,
            day: 22,
        };

        // Test positive day offset crossing month boundaries
        assert_eq!(
            add_days(date, 10),
            SimpleDate {
                year: 2026,
                month: 7,
                day: 2
            }
        );
        // Test negative day offset crossing month boundaries
        assert_eq!(
            add_days(date, -23),
            SimpleDate {
                year: 2026,
                month: 5,
                day: 30
            }
        );
        // Test leap year day addition
        let leap_date = SimpleDate {
            year: 2024,
            month: 2,
            day: 28,
        };
        assert_eq!(
            add_days(leap_date, 1),
            SimpleDate {
                year: 2024,
                month: 2,
                day: 29
            }
        );
        assert_eq!(
            add_days(leap_date, 2),
            SimpleDate {
                year: 2024,
                month: 3,
                day: 1
            }
        );

        // Test month addition
        assert_eq!(
            add_months(date, 2),
            SimpleDate {
                year: 2026,
                month: 8,
                day: 1
            }
        );
        assert_eq!(
            add_months(date, -7),
            SimpleDate {
                year: 2025,
                month: 11,
                day: 1
            }
        );
    }
}
