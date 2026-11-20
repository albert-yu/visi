pub(crate) const NS_SPREADSHEETML_TRANSITIONAL: &str =
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
pub(crate) const NS_SPREADSHEETML_STRICT: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";

pub(crate) const NS_OFFICE_RELS_TRANSITIONAL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub(crate) const NS_OFFICE_RELS_STRICT: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships";

pub(crate) const NS_DRAWINGML_MAIN_TRANSITIONAL: &str =
    "http://schemas.openxmlformats.org/drawingml/2006/main";
pub(crate) const NS_DRAWINGML_MAIN_STRICT: &str = "http://purl.oclc.org/ooxml/drawingml/main";

pub(crate) const NS_DRAWINGML_CHART_TRANSITIONAL: &str =
    "http://schemas.openxmlformats.org/drawingml/2006/chart";
pub(crate) const NS_DRAWINGML_CHART_STRICT: &str = "http://purl.oclc.org/ooxml/drawingml/chart";

pub(crate) const NS_DRAWINGML_SPREADSHEET_TRANSITIONAL: &str =
    "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
pub(crate) const NS_DRAWINGML_SPREADSHEET_STRICT: &str =
    "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing";

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub(crate) enum NamespaceFlavor {
    #[default]
    Transitional,
    Strict,
}

impl NamespaceFlavor {
    pub(crate) fn from_xmlns(xmlns: &str) -> Self {
        if xmlns.starts_with("http://purl.oclc.org/ooxml/") {
            Self::Strict
        } else {
            Self::Transitional
        }
    }

    pub(crate) fn spreadsheetml(self) -> &'static str {
        match self {
            Self::Transitional => NS_SPREADSHEETML_TRANSITIONAL,
            Self::Strict => NS_SPREADSHEETML_STRICT,
        }
    }

    pub(crate) fn office_rels(self) -> &'static str {
        match self {
            Self::Transitional => NS_OFFICE_RELS_TRANSITIONAL,
            Self::Strict => NS_OFFICE_RELS_STRICT,
        }
    }

    pub(crate) fn drawingml_main(self) -> &'static str {
        match self {
            Self::Transitional => NS_DRAWINGML_MAIN_TRANSITIONAL,
            Self::Strict => NS_DRAWINGML_MAIN_STRICT,
        }
    }

    pub(crate) fn drawingml_chart(self) -> &'static str {
        match self {
            Self::Transitional => NS_DRAWINGML_CHART_TRANSITIONAL,
            Self::Strict => NS_DRAWINGML_CHART_STRICT,
        }
    }

    pub(crate) fn drawingml_spreadsheet(self) -> &'static str {
        match self {
            Self::Transitional => NS_DRAWINGML_SPREADSHEET_TRANSITIONAL,
            Self::Strict => NS_DRAWINGML_SPREADSHEET_STRICT,
        }
    }

    pub(crate) fn rel_type(self, suffix: &str) -> String {
        match self {
            Self::Transitional => {
                format!(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/{suffix}"
                )
            }
            Self::Strict => {
                format!("http://purl.oclc.org/ooxml/officeDocument/relationships/{suffix}")
            }
        }
    }
}

pub(crate) fn needs_xml_space_preserve(s: &str) -> bool {
    if s.is_empty() {
        return false;
    }
    let first = s.chars().next().unwrap();
    let last = s.chars().last().unwrap();
    first.is_whitespace()
        || last.is_whitespace()
        || s.contains('\n')
        || s.contains('\r')
        || s.contains('\t')
}

pub(crate) fn decode_xstring(s: &str) -> String {
    if !s.contains("_x") {
        return s.to_string();
    }
    let bytes = s.as_bytes();
    let mut out = String::with_capacity(s.len());
    let mut i = 0usize;
    while i < bytes.len() {
        if i + 6 < bytes.len()
            && bytes[i] == b'_'
            && bytes[i + 1] == b'x'
            && bytes[i + 6] == b'_'
            && let Ok(hex) = std::str::from_utf8(&bytes[i + 2..i + 6])
            && let Ok(code) = u32::from_str_radix(hex, 16)
        {
            if (0xD800..=0xDBFF).contains(&code)
                && i + 13 < bytes.len()
                && bytes[i + 7] == b'_'
                && bytes[i + 8] == b'x'
                && bytes[i + 13] == b'_'
                && let Ok(low_hex) = std::str::from_utf8(&bytes[i + 9..i + 13])
                && let Ok(low) = u32::from_str_radix(low_hex, 16)
                && (0xDC00..=0xDFFF).contains(&low)
            {
                let combined = 0x10000 + ((code - 0xD800) << 10) + (low - 0xDC00);
                if let Some(ch) = char::from_u32(combined) {
                    out.push(ch);
                    i += 14;
                    continue;
                }
            }
            if let Some(ch) = char::from_u32(code) {
                out.push(ch);
                i += 7;
                continue;
            }
        }
        let ch = s[i..].chars().next().unwrap();
        out.push(ch);
        i += ch.len_utf8();
    }
    out
}

pub(crate) fn encode_xstring(s: &str) -> String {
    let bytes = s.as_bytes();
    let mut out = String::with_capacity(s.len());
    let mut i = 0usize;
    while i < bytes.len() {
        if i + 6 < bytes.len()
            && bytes[i] == b'_'
            && bytes[i + 1] == b'x'
            && bytes[i + 6] == b'_'
            && bytes[i + 2..i + 6].iter().all(u8::is_ascii_hexdigit)
        {
            out.push_str("_x005F_");
            i += 1;
            continue;
        }
        let ch = s[i..].chars().next().unwrap();
        let cp = ch as u32;
        if cp < 0x20 && cp != 0x09 && cp != 0x0A && cp != 0x0D {
            out.push_str(&format!("_x{cp:04X}_"));
        } else {
            out.push(ch);
        }
        i += ch.len_utf8();
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn xstring_round_trips_control_chars_and_literal_escapes() {
        assert_eq!(decode_xstring("a_x000A_b"), "a\nb");
        assert_eq!(decode_xstring("_x005F_x000A_"), "_x000A_");
        assert_eq!(decode_xstring("plain text"), "plain text");
        assert_eq!(encode_xstring("a\x01b"), "a_x0001_b");
        assert_eq!(encode_xstring("_x000A_"), "_x005F_x000A_");
        assert_eq!(decode_xstring(&encode_xstring("_x000A_")), "_x000A_");
        assert_eq!(decode_xstring(&encode_xstring("a\x07b")), "a\x07b");
    }

    #[test]
    fn xml_space_preserve_detection() {
        assert!(!needs_xml_space_preserve(""));
        assert!(!needs_xml_space_preserve("hello world"));
        assert!(needs_xml_space_preserve(" leading"));
        assert!(needs_xml_space_preserve("trailing "));
        assert!(needs_xml_space_preserve("line1\nline2"));
    }
}
