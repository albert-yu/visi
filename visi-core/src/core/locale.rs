//! Locale awareness for spreadsheet regional settings and date parsing.
//!
//! Provides [`Locale`] and [`DateOrder`] configuration matching Excel's
//! regional behaviors: date field order (MDY vs DMY vs YMD), primary
//! separators, localized month names in major languages, and 2-digit year
//! pivot settings.

use serde::{Deserialize, Serialize};

/// Date field ordering of a locale.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub enum DateOrder {
    /// Month-Day-Year (e.g. `en-US`: `06/22/2026`).
    Mdy,
    /// Day-Month-Year (e.g. `en-GB`, `de-DE`, `fr-FR`: `22/06/2026`).
    Dmy,
    /// Year-Month-Day (e.g. `zh-CN`, `ja-JP`, ISO: `2026-06-22`).
    Ymd,
}

impl std::fmt::Display for DateOrder {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            DateOrder::Mdy => write!(f, "MDY"),
            DateOrder::Dmy => write!(f, "DMY"),
            DateOrder::Ymd => write!(f, "YMD"),
        }
    }
}

/// A spreadsheet regional locale configuring date and number parsing rules.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct Locale {
    /// Language / country tag (e.g. `"en-US"`, `"en-GB"`, `"de-DE"`, `"fr-FR"`).
    pub code: String,
    /// Primary date field ordering.
    pub date_order: DateOrder,
    /// Primary date separator (`'/'`, `'.' `, or `'-'`).
    pub primary_date_sep: char,
    /// Windows LCID (e.g., `0x0409` for en-US, `0x0809` for en-GB, `0x0407` for de-DE).
    pub lcid: u16,
    /// 2-digit year pivot threshold (default 29 -> 00..29 is 20xx, 30..99 is 19xx).
    pub two_digit_year_pivot: u32,
    /// Default calendar year for 2-part dates (if `None`, defaults to 2026).
    pub default_year: Option<i32>,
}

impl Default for Locale {
    fn default() -> Self {
        Self::en_us()
    }
}

impl Locale {
    /// English (United States) locale: `MDY`, separator `/`, LCID `0x0409`.
    pub fn en_us() -> Self {
        Self {
            code: "en-US".to_string(),
            date_order: DateOrder::Mdy,
            primary_date_sep: '/',
            lcid: 0x0409,
            two_digit_year_pivot: 29,
            default_year: None,
        }
    }

    /// English (United Kingdom) locale: `DMY`, separator `/`, LCID `0x0809`.
    pub fn en_gb() -> Self {
        Self {
            code: "en-GB".to_string(),
            date_order: DateOrder::Dmy,
            primary_date_sep: '/',
            lcid: 0x0809,
            two_digit_year_pivot: 29,
            default_year: None,
        }
    }

    /// German (Germany) locale: `DMY`, separator `.`, LCID `0x0407`.
    pub fn de_de() -> Self {
        Self {
            code: "de-DE".to_string(),
            date_order: DateOrder::Dmy,
            primary_date_sep: '.',
            lcid: 0x0407,
            two_digit_year_pivot: 29,
            default_year: None,
        }
    }

    /// French (France) locale: `DMY`, separator `/`, LCID `0x040C`.
    pub fn fr_fr() -> Self {
        Self {
            code: "fr-FR".to_string(),
            date_order: DateOrder::Dmy,
            primary_date_sep: '/',
            lcid: 0x040C,
            two_digit_year_pivot: 29,
            default_year: None,
        }
    }

    /// Spanish (Spain) locale: `DMY`, separator `/`, LCID `0x0C0A`.
    pub fn es_es() -> Self {
        Self {
            code: "es-ES".to_string(),
            date_order: DateOrder::Dmy,
            primary_date_sep: '/',
            lcid: 0x0C0A,
            two_digit_year_pivot: 29,
            default_year: None,
        }
    }

    /// Italian (Italy) locale: `DMY`, separator `/`, LCID `0x0410`.
    pub fn it_it() -> Self {
        Self {
            code: "it-IT".to_string(),
            date_order: DateOrder::Dmy,
            primary_date_sep: '/',
            lcid: 0x0410,
            two_digit_year_pivot: 29,
            default_year: None,
        }
    }

    /// Portuguese (Brazil) locale: `DMY`, separator `/`, LCID `0x0416`.
    pub fn pt_br() -> Self {
        Self {
            code: "pt-BR".to_string(),
            date_order: DateOrder::Dmy,
            primary_date_sep: '/',
            lcid: 0x0416,
            two_digit_year_pivot: 29,
            default_year: None,
        }
    }

    /// Dutch (Netherlands) locale: `DMY`, separator `-`, LCID `0x0413`.
    pub fn nl_nl() -> Self {
        Self {
            code: "nl-NL".to_string(),
            date_order: DateOrder::Dmy,
            primary_date_sep: '-',
            lcid: 0x0413,
            two_digit_year_pivot: 29,
            default_year: None,
        }
    }

    /// Russian (Russia) locale: `DMY`, separator `.`, LCID `0x0419`.
    pub fn ru_ru() -> Self {
        Self {
            code: "ru-RU".to_string(),
            date_order: DateOrder::Dmy,
            primary_date_sep: '.',
            lcid: 0x0419,
            two_digit_year_pivot: 29,
            default_year: None,
        }
    }

    /// Chinese (Simplified, China) locale: `YMD`, separator `/`, LCID `0x0804`.
    pub fn zh_cn() -> Self {
        Self {
            code: "zh-CN".to_string(),
            date_order: DateOrder::Ymd,
            primary_date_sep: '/',
            lcid: 0x0804,
            two_digit_year_pivot: 29,
            default_year: None,
        }
    }

    /// Japanese (Japan) locale: `YMD`, separator `/`, LCID `0x0411`.
    pub fn ja_jp() -> Self {
        Self {
            code: "ja-JP".to_string(),
            date_order: DateOrder::Ymd,
            primary_date_sep: '/',
            lcid: 0x0411,
            two_digit_year_pivot: 29,
            default_year: None,
        }
    }

    /// Look up or construct a locale from a BCP 47 or POSIX language tag.
    ///
    /// Recognizes tags like `"en-US"`, `"en_US"`, `"en-GB"`, `"de"`, `"de-DE"`,
    /// `"fr"`, `"es"`, `"it"`, `"pt"`, `"nl"`, `"ru"`, `"zh"`, `"ja"`, etc.
    pub fn from_code(code: &str) -> Option<Self> {
        let norm = code.trim().replace('_', "-").to_lowercase();
        if norm.starts_with("en-gb")
            || norm.starts_with("en-au")
            || norm.starts_with("en-nz")
            || norm.starts_with("en-ie")
        {
            Some(Self::en_gb())
        } else if norm.starts_with("en") {
            Some(Self::en_us())
        } else if norm.starts_with("de") {
            Some(Self::de_de())
        } else if norm.starts_with("fr") {
            Some(Self::fr_fr())
        } else if norm.starts_with("es") {
            Some(Self::es_es())
        } else if norm.starts_with("it") {
            Some(Self::it_it())
        } else if norm.starts_with("pt") {
            Some(Self::pt_br())
        } else if norm.starts_with("nl") {
            Some(Self::nl_nl())
        } else if norm.starts_with("ru") {
            Some(Self::ru_ru())
        } else if norm.starts_with("zh") {
            Some(Self::zh_cn())
        } else if norm.starts_with("ja") {
            Some(Self::ja_jp())
        } else {
            None
        }
    }

    /// The default year for 2-part dates (e.g. `6/22` or `22-Jun`).
    pub fn default_year(&self) -> i32 {
        self.default_year.unwrap_or(2026)
    }

    /// Convert a 2-digit year to a 4-digit year using this locale's pivot.
    pub fn expand_two_digit_year(&self, year: i32) -> i32 {
        if (0..=self.two_digit_year_pivot as i32).contains(&year) {
            2000 + year
        } else if (0..100).contains(&year) {
            1900 + year
        } else {
            year
        }
    }

    /// Attempts to match a month name/word token against this locale's month
    /// dictionaries (or English as fallback). Returns `(month_1_based, is_full_name)`.
    pub fn match_month_word(&self, word: &str) -> Option<(u32, bool)> {
        let cleaned = word.trim().trim_end_matches('.');
        if cleaned.is_empty() {
            return None;
        }

        let lang = self.code.split('-').next().unwrap_or("en").to_lowercase();
        if let Some(res) = match_month_for_lang(&lang, cleaned) {
            return Some(res);
        }

        if lang != "en"
            && let Some(res) = match_month_for_lang("en", cleaned)
        {
            return Some(res);
        }

        None
    }
}

struct MonthEntry {
    name: &'static str,
    month: u32,
    full: bool,
}

fn normalize_str(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for c in s.chars() {
        match c.to_ascii_lowercase() {
            'é' | 'è' | 'ê' | 'ë' => out.push('e'),
            'á' | 'à' | 'â' | 'ä' | 'ã' => out.push('a'),
            'í' | 'ì' | 'î' | 'ï' => out.push('i'),
            'ó' | 'ò' | 'ô' | 'ö' | 'õ' => out.push('o'),
            'ú' | 'ù' | 'û' | 'ü' => out.push('u'),
            'ñ' => out.push('n'),
            'ç' => out.push('c'),
            other => out.push(other),
        }
    }
    out
}

fn match_month_for_lang(lang: &str, word: &str) -> Option<(u32, bool)> {
    let norm = normalize_str(word);
    let entries = get_month_entries_for_lang(lang);
    for entry in entries {
        if normalize_str(entry.name) == norm {
            return Some((entry.month, entry.full));
        }
    }
    None
}

fn get_month_entries_for_lang(lang: &str) -> &'static [MonthEntry] {
    match lang {
        "de" => &GERMAN_MONTHS,
        "fr" => &FRENCH_MONTHS,
        "es" => &SPANISH_MONTHS,
        "it" => &ITALIAN_MONTHS,
        "pt" => &PORTUGUESE_MONTHS,
        "nl" => &DUTCH_MONTHS,
        "ru" => &RUSSIAN_MONTHS,
        "zh" | "ja" => &CJK_MONTHS,
        _ => &ENGLISH_MONTHS,
    }
}

const ENGLISH_MONTHS: [MonthEntry; 24] = [
    MonthEntry {
        name: "january",
        month: 1,
        full: true,
    },
    MonthEntry {
        name: "february",
        month: 2,
        full: true,
    },
    MonthEntry {
        name: "march",
        month: 3,
        full: true,
    },
    MonthEntry {
        name: "april",
        month: 4,
        full: true,
    },
    MonthEntry {
        name: "may",
        month: 5,
        full: true,
    },
    MonthEntry {
        name: "june",
        month: 6,
        full: true,
    },
    MonthEntry {
        name: "july",
        month: 7,
        full: true,
    },
    MonthEntry {
        name: "august",
        month: 8,
        full: true,
    },
    MonthEntry {
        name: "september",
        month: 9,
        full: true,
    },
    MonthEntry {
        name: "october",
        month: 10,
        full: true,
    },
    MonthEntry {
        name: "november",
        month: 11,
        full: true,
    },
    MonthEntry {
        name: "december",
        month: 12,
        full: true,
    },
    MonthEntry {
        name: "jan",
        month: 1,
        full: false,
    },
    MonthEntry {
        name: "feb",
        month: 2,
        full: false,
    },
    MonthEntry {
        name: "mar",
        month: 3,
        full: false,
    },
    MonthEntry {
        name: "apr",
        month: 4,
        full: false,
    },
    MonthEntry {
        name: "may",
        month: 5,
        full: false,
    },
    MonthEntry {
        name: "jun",
        month: 6,
        full: false,
    },
    MonthEntry {
        name: "jul",
        month: 7,
        full: false,
    },
    MonthEntry {
        name: "aug",
        month: 8,
        full: false,
    },
    MonthEntry {
        name: "sep",
        month: 9,
        full: false,
    },
    MonthEntry {
        name: "oct",
        month: 10,
        full: false,
    },
    MonthEntry {
        name: "nov",
        month: 11,
        full: false,
    },
    MonthEntry {
        name: "dec",
        month: 12,
        full: false,
    },
];

const GERMAN_MONTHS: [MonthEntry; 27] = [
    MonthEntry {
        name: "januar",
        month: 1,
        full: true,
    },
    MonthEntry {
        name: "februar",
        month: 2,
        full: true,
    },
    MonthEntry {
        name: "marz",
        month: 3,
        full: true,
    },
    MonthEntry {
        name: "maerz",
        month: 3,
        full: true,
    },
    MonthEntry {
        name: "april",
        month: 4,
        full: true,
    },
    MonthEntry {
        name: "mai",
        month: 5,
        full: true,
    },
    MonthEntry {
        name: "juni",
        month: 6,
        full: true,
    },
    MonthEntry {
        name: "juli",
        month: 7,
        full: true,
    },
    MonthEntry {
        name: "august",
        month: 8,
        full: true,
    },
    MonthEntry {
        name: "september",
        month: 9,
        full: true,
    },
    MonthEntry {
        name: "oktober",
        month: 10,
        full: true,
    },
    MonthEntry {
        name: "november",
        month: 11,
        full: true,
    },
    MonthEntry {
        name: "dezember",
        month: 12,
        full: true,
    },
    MonthEntry {
        name: "jan",
        month: 1,
        full: false,
    },
    MonthEntry {
        name: "feb",
        month: 2,
        full: false,
    },
    MonthEntry {
        name: "mar",
        month: 3,
        full: false,
    },
    MonthEntry {
        name: "mrz",
        month: 3,
        full: false,
    },
    MonthEntry {
        name: "apr",
        month: 4,
        full: false,
    },
    MonthEntry {
        name: "mai",
        month: 5,
        full: false,
    },
    MonthEntry {
        name: "jun",
        month: 6,
        full: false,
    },
    MonthEntry {
        name: "jul",
        month: 7,
        full: false,
    },
    MonthEntry {
        name: "aug",
        month: 8,
        full: false,
    },
    MonthEntry {
        name: "sep",
        month: 9,
        full: false,
    },
    MonthEntry {
        name: "okt",
        month: 10,
        full: false,
    },
    MonthEntry {
        name: "nov",
        month: 11,
        full: false,
    },
    MonthEntry {
        name: "dez",
        month: 12,
        full: false,
    },
    MonthEntry {
        name: "sept",
        month: 9,
        full: false,
    },
];

const FRENCH_MONTHS: [MonthEntry; 28] = [
    MonthEntry {
        name: "janvier",
        month: 1,
        full: true,
    },
    MonthEntry {
        name: "fevrier",
        month: 2,
        full: true,
    },
    MonthEntry {
        name: "mars",
        month: 3,
        full: true,
    },
    MonthEntry {
        name: "avril",
        month: 4,
        full: true,
    },
    MonthEntry {
        name: "mai",
        month: 5,
        full: true,
    },
    MonthEntry {
        name: "juin",
        month: 6,
        full: true,
    },
    MonthEntry {
        name: "juillet",
        month: 7,
        full: true,
    },
    MonthEntry {
        name: "aout",
        month: 8,
        full: true,
    },
    MonthEntry {
        name: "septembre",
        month: 9,
        full: true,
    },
    MonthEntry {
        name: "octobre",
        month: 10,
        full: true,
    },
    MonthEntry {
        name: "novembre",
        month: 11,
        full: true,
    },
    MonthEntry {
        name: "decembre",
        month: 12,
        full: true,
    },
    MonthEntry {
        name: "janv",
        month: 1,
        full: false,
    },
    MonthEntry {
        name: "jan",
        month: 1,
        full: false,
    },
    MonthEntry {
        name: "fevr",
        month: 2,
        full: false,
    },
    MonthEntry {
        name: "fev",
        month: 2,
        full: false,
    },
    MonthEntry {
        name: "mar",
        month: 3,
        full: false,
    },
    MonthEntry {
        name: "avr",
        month: 4,
        full: false,
    },
    MonthEntry {
        name: "mai",
        month: 5,
        full: false,
    },
    MonthEntry {
        name: "juin",
        month: 6,
        full: false,
    },
    MonthEntry {
        name: "juil",
        month: 7,
        full: false,
    },
    MonthEntry {
        name: "aout",
        month: 8,
        full: false,
    },
    MonthEntry {
        name: "aou",
        month: 8,
        full: false,
    },
    MonthEntry {
        name: "sept",
        month: 9,
        full: false,
    },
    MonthEntry {
        name: "sep",
        month: 9,
        full: false,
    },
    MonthEntry {
        name: "oct",
        month: 10,
        full: false,
    },
    MonthEntry {
        name: "nov",
        month: 11,
        full: false,
    },
    MonthEntry {
        name: "dec",
        month: 12,
        full: false,
    },
];

const SPANISH_MONTHS: [MonthEntry; 26] = [
    MonthEntry {
        name: "enero",
        month: 1,
        full: true,
    },
    MonthEntry {
        name: "febrero",
        month: 2,
        full: true,
    },
    MonthEntry {
        name: "marzo",
        month: 3,
        full: true,
    },
    MonthEntry {
        name: "abril",
        month: 4,
        full: true,
    },
    MonthEntry {
        name: "mayo",
        month: 5,
        full: true,
    },
    MonthEntry {
        name: "junio",
        month: 6,
        full: true,
    },
    MonthEntry {
        name: "julio",
        month: 7,
        full: true,
    },
    MonthEntry {
        name: "agosto",
        month: 8,
        full: true,
    },
    MonthEntry {
        name: "septiembre",
        month: 9,
        full: true,
    },
    MonthEntry {
        name: "setiembre",
        month: 9,
        full: true,
    },
    MonthEntry {
        name: "octubre",
        month: 10,
        full: true,
    },
    MonthEntry {
        name: "noviembre",
        month: 11,
        full: true,
    },
    MonthEntry {
        name: "diciembre",
        month: 12,
        full: true,
    },
    MonthEntry {
        name: "ene",
        month: 1,
        full: false,
    },
    MonthEntry {
        name: "feb",
        month: 2,
        full: false,
    },
    MonthEntry {
        name: "mar",
        month: 3,
        full: false,
    },
    MonthEntry {
        name: "abr",
        month: 4,
        full: false,
    },
    MonthEntry {
        name: "may",
        month: 5,
        full: false,
    },
    MonthEntry {
        name: "jun",
        month: 6,
        full: false,
    },
    MonthEntry {
        name: "jul",
        month: 7,
        full: false,
    },
    MonthEntry {
        name: "ago",
        month: 8,
        full: false,
    },
    MonthEntry {
        name: "sep",
        month: 9,
        full: false,
    },
    MonthEntry {
        name: "set",
        month: 9,
        full: false,
    },
    MonthEntry {
        name: "oct",
        month: 10,
        full: false,
    },
    MonthEntry {
        name: "nov",
        month: 11,
        full: false,
    },
    MonthEntry {
        name: "dic",
        month: 12,
        full: false,
    },
];

const ITALIAN_MONTHS: [MonthEntry; 24] = [
    MonthEntry {
        name: "gennaio",
        month: 1,
        full: true,
    },
    MonthEntry {
        name: "febbraio",
        month: 2,
        full: true,
    },
    MonthEntry {
        name: "marzo",
        month: 3,
        full: true,
    },
    MonthEntry {
        name: "aprile",
        month: 4,
        full: true,
    },
    MonthEntry {
        name: "maggio",
        month: 5,
        full: true,
    },
    MonthEntry {
        name: "giugno",
        month: 6,
        full: true,
    },
    MonthEntry {
        name: "luglio",
        month: 7,
        full: true,
    },
    MonthEntry {
        name: "agosto",
        month: 8,
        full: true,
    },
    MonthEntry {
        name: "settembre",
        month: 9,
        full: true,
    },
    MonthEntry {
        name: "ottobre",
        month: 10,
        full: true,
    },
    MonthEntry {
        name: "novembre",
        month: 11,
        full: true,
    },
    MonthEntry {
        name: "dicembre",
        month: 12,
        full: true,
    },
    MonthEntry {
        name: "gen",
        month: 1,
        full: false,
    },
    MonthEntry {
        name: "feb",
        month: 2,
        full: false,
    },
    MonthEntry {
        name: "mar",
        month: 3,
        full: false,
    },
    MonthEntry {
        name: "apr",
        month: 4,
        full: false,
    },
    MonthEntry {
        name: "mag",
        month: 5,
        full: false,
    },
    MonthEntry {
        name: "giu",
        month: 6,
        full: false,
    },
    MonthEntry {
        name: "lug",
        month: 7,
        full: false,
    },
    MonthEntry {
        name: "ago",
        month: 8,
        full: false,
    },
    MonthEntry {
        name: "set",
        month: 9,
        full: false,
    },
    MonthEntry {
        name: "ott",
        month: 10,
        full: false,
    },
    MonthEntry {
        name: "nov",
        month: 11,
        full: false,
    },
    MonthEntry {
        name: "dic",
        month: 12,
        full: false,
    },
];

const PORTUGUESE_MONTHS: [MonthEntry; 24] = [
    MonthEntry {
        name: "janeiro",
        month: 1,
        full: true,
    },
    MonthEntry {
        name: "fevereiro",
        month: 2,
        full: true,
    },
    MonthEntry {
        name: "marco",
        month: 3,
        full: true,
    },
    MonthEntry {
        name: "abril",
        month: 4,
        full: true,
    },
    MonthEntry {
        name: "maio",
        month: 5,
        full: true,
    },
    MonthEntry {
        name: "junho",
        month: 6,
        full: true,
    },
    MonthEntry {
        name: "julho",
        month: 7,
        full: true,
    },
    MonthEntry {
        name: "agosto",
        month: 8,
        full: true,
    },
    MonthEntry {
        name: "setembro",
        month: 9,
        full: true,
    },
    MonthEntry {
        name: "outubro",
        month: 10,
        full: true,
    },
    MonthEntry {
        name: "novembro",
        month: 11,
        full: true,
    },
    MonthEntry {
        name: "dezembro",
        month: 12,
        full: true,
    },
    MonthEntry {
        name: "jan",
        month: 1,
        full: false,
    },
    MonthEntry {
        name: "fev",
        month: 2,
        full: false,
    },
    MonthEntry {
        name: "mar",
        month: 3,
        full: false,
    },
    MonthEntry {
        name: "abr",
        month: 4,
        full: false,
    },
    MonthEntry {
        name: "mai",
        month: 5,
        full: false,
    },
    MonthEntry {
        name: "jun",
        month: 6,
        full: false,
    },
    MonthEntry {
        name: "jul",
        month: 7,
        full: false,
    },
    MonthEntry {
        name: "ago",
        month: 8,
        full: false,
    },
    MonthEntry {
        name: "set",
        month: 9,
        full: false,
    },
    MonthEntry {
        name: "out",
        month: 10,
        full: false,
    },
    MonthEntry {
        name: "nov",
        month: 11,
        full: false,
    },
    MonthEntry {
        name: "dez",
        month: 12,
        full: false,
    },
];

const DUTCH_MONTHS: [MonthEntry; 25] = [
    MonthEntry {
        name: "januari",
        month: 1,
        full: true,
    },
    MonthEntry {
        name: "februari",
        month: 2,
        full: true,
    },
    MonthEntry {
        name: "maart",
        month: 3,
        full: true,
    },
    MonthEntry {
        name: "april",
        month: 4,
        full: true,
    },
    MonthEntry {
        name: "mei",
        month: 5,
        full: true,
    },
    MonthEntry {
        name: "juni",
        month: 6,
        full: true,
    },
    MonthEntry {
        name: "juli",
        month: 7,
        full: true,
    },
    MonthEntry {
        name: "augustus",
        month: 8,
        full: true,
    },
    MonthEntry {
        name: "september",
        month: 9,
        full: true,
    },
    MonthEntry {
        name: "oktober",
        month: 10,
        full: true,
    },
    MonthEntry {
        name: "november",
        month: 11,
        full: true,
    },
    MonthEntry {
        name: "december",
        month: 12,
        full: true,
    },
    MonthEntry {
        name: "jan",
        month: 1,
        full: false,
    },
    MonthEntry {
        name: "feb",
        month: 2,
        full: false,
    },
    MonthEntry {
        name: "mrt",
        month: 3,
        full: false,
    },
    MonthEntry {
        name: "mar",
        month: 3,
        full: false,
    },
    MonthEntry {
        name: "apr",
        month: 4,
        full: false,
    },
    MonthEntry {
        name: "mei",
        month: 5,
        full: false,
    },
    MonthEntry {
        name: "jun",
        month: 6,
        full: false,
    },
    MonthEntry {
        name: "jul",
        month: 7,
        full: false,
    },
    MonthEntry {
        name: "aug",
        month: 8,
        full: false,
    },
    MonthEntry {
        name: "sep",
        month: 9,
        full: false,
    },
    MonthEntry {
        name: "okt",
        month: 10,
        full: false,
    },
    MonthEntry {
        name: "nov",
        month: 11,
        full: false,
    },
    MonthEntry {
        name: "dec",
        month: 12,
        full: false,
    },
];

const RUSSIAN_MONTHS: [MonthEntry; 24] = [
    MonthEntry {
        name: "январь",
        month: 1,
        full: true,
    },
    MonthEntry {
        name: "февраль",
        month: 2,
        full: true,
    },
    MonthEntry {
        name: "март",
        month: 3,
        full: true,
    },
    MonthEntry {
        name: "апрель",
        month: 4,
        full: true,
    },
    MonthEntry {
        name: "май",
        month: 5,
        full: true,
    },
    MonthEntry {
        name: "июнь",
        month: 6,
        full: true,
    },
    MonthEntry {
        name: "июль",
        month: 7,
        full: true,
    },
    MonthEntry {
        name: "август",
        month: 8,
        full: true,
    },
    MonthEntry {
        name: "сентябрь",
        month: 9,
        full: true,
    },
    MonthEntry {
        name: "октябрь",
        month: 10,
        full: true,
    },
    MonthEntry {
        name: "ноябрь",
        month: 11,
        full: true,
    },
    MonthEntry {
        name: "декабрь",
        month: 12,
        full: true,
    },
    MonthEntry {
        name: "янв",
        month: 1,
        full: false,
    },
    MonthEntry {
        name: "фев",
        month: 2,
        full: false,
    },
    MonthEntry {
        name: "мар",
        month: 3,
        full: false,
    },
    MonthEntry {
        name: "апр",
        month: 4,
        full: false,
    },
    MonthEntry {
        name: "май",
        month: 5,
        full: false,
    },
    MonthEntry {
        name: "июн",
        month: 6,
        full: false,
    },
    MonthEntry {
        name: "июл",
        month: 7,
        full: false,
    },
    MonthEntry {
        name: "авг",
        month: 8,
        full: false,
    },
    MonthEntry {
        name: "сен",
        month: 9,
        full: false,
    },
    MonthEntry {
        name: "окт",
        month: 10,
        full: false,
    },
    MonthEntry {
        name: "ноя",
        month: 11,
        full: false,
    },
    MonthEntry {
        name: "дек",
        month: 12,
        full: false,
    },
];

const CJK_MONTHS: [MonthEntry; 12] = [
    MonthEntry {
        name: "1月",
        month: 1,
        full: true,
    },
    MonthEntry {
        name: "2月",
        month: 2,
        full: true,
    },
    MonthEntry {
        name: "3月",
        month: 3,
        full: true,
    },
    MonthEntry {
        name: "4月",
        month: 4,
        full: true,
    },
    MonthEntry {
        name: "5月",
        month: 5,
        full: true,
    },
    MonthEntry {
        name: "6月",
        month: 6,
        full: true,
    },
    MonthEntry {
        name: "7月",
        month: 7,
        full: true,
    },
    MonthEntry {
        name: "8月",
        month: 8,
        full: true,
    },
    MonthEntry {
        name: "9月",
        month: 9,
        full: true,
    },
    MonthEntry {
        name: "10月",
        month: 10,
        full: true,
    },
    MonthEntry {
        name: "11月",
        month: 11,
        full: true,
    },
    MonthEntry {
        name: "12月",
        month: 12,
        full: true,
    },
];
