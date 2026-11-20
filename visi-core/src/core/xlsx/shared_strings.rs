use std::collections::HashMap;

use quick_xml::events::Event;

use crate::core::xlsx::xml_util::{
    NamespaceFlavor, decode_xstring, encode_xstring, needs_xml_space_preserve,
};

pub(crate) fn parse_shared_strings(xml: &str) -> Result<Vec<String>, String> {
    let mut strings = Vec::new();
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buf = Vec::new();

    let mut in_si = false;
    let mut in_t = false;
    let mut phonetic_depth = 0usize;
    let mut current = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"si" => {
                    in_si = true;
                    current.clear();
                }
                b"rPh" if in_si => {
                    phonetic_depth += 1;
                }
                b"t" if in_si && phonetic_depth == 0 => {
                    in_t = true;
                }
                _ => {}
            },
            Ok(Event::Empty(ref e)) => {
                if e.local_name().as_ref() == b"si" {
                    strings.push(String::new());
                }
            }
            Ok(Event::Text(e)) if in_t && phonetic_depth == 0 => {
                let decoded = e
                    .decode()
                    .map_err(|err| format!("sharedStrings text decode error: {err}"))?;
                let unescaped = quick_xml::escape::unescape(&decoded)
                    .map_err(|err| format!("sharedStrings text unescape error: {err}"))?;
                current.push_str(&unescaped);
            }
            Ok(Event::GeneralRef(e)) if in_t && phonetic_depth == 0 => {
                let raw = e
                    .decode()
                    .map_err(|err| format!("sharedStrings entity decode error: {err}"))?;
                let entity = format!("&{raw};");
                let resolved = quick_xml::escape::unescape(&entity)
                    .map(|v| v.into_owned())
                    .unwrap_or(entity);
                current.push_str(&resolved);
            }
            Ok(Event::CData(e)) if in_t && phonetic_depth == 0 => {
                current.push_str(&String::from_utf8_lossy(e.as_ref()));
            }
            Ok(Event::End(ref e)) => match e.local_name().as_ref() {
                b"si" => {
                    in_si = false;
                    in_t = false;
                    phonetic_depth = 0;
                    strings.push(decode_xstring(&current));
                    current.clear();
                }
                b"rPh" if phonetic_depth > 0 => {
                    phonetic_depth -= 1;
                }
                b"t" => {
                    in_t = false;
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(err) => return Err(format!("Failed to parse sharedStrings.xml: {err}")),
            _ => {}
        }
        buf.clear();
    }

    Ok(strings)
}

#[derive(Debug, Default)]
pub(crate) struct SharedStringTable {
    strings: Vec<String>,
    index: HashMap<String, usize>,
    total_count: usize,
}

impl SharedStringTable {
    pub(crate) fn intern(&mut self, value: &str) -> usize {
        self.total_count += 1;
        if let Some(&idx) = self.index.get(value) {
            return idx;
        }
        let idx = self.strings.len();
        self.strings.push(value.to_string());
        self.index.insert(value.to_string(), idx);
        idx
    }

    pub(crate) fn is_empty(&self) -> bool {
        self.strings.is_empty()
    }

    pub(crate) fn to_xml(&self, flavor: NamespaceFlavor) -> String {
        let ns = flavor.spreadsheetml();
        let mut xml = format!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><sst xmlns=\"{ns}\" count=\"{}\" uniqueCount=\"{}\">",
            self.total_count,
            self.strings.len()
        );
        for s in &self.strings {
            let encoded = encode_xstring(s);
            if needs_xml_space_preserve(&encoded) {
                xml.push_str("<si><t xml:space=\"preserve\">");
            } else {
                xml.push_str("<si><t>");
            }
            xml.push_str(&super::escape_xml_text(&encoded));
            xml.push_str("</t></si>");
        }
        xml.push_str("</sst>");
        xml
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_plain_rich_text_empty_and_phonetic_shared_strings() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4">
            <si><t>hello &amp; world</t></si>
            <si><r><t xml:space="preserve">bold </t></r><r><t>text</t></r><rPh sb="0" eb="1"><t>ignored</t></rPh></si>
            <si><t/></si>
            <si><t>a_x000A_b</t></si>
        </sst>"#;
        let parsed = parse_shared_strings(xml).unwrap();
        assert_eq!(parsed, vec!["hello & world", "bold text", "", "a\nb"]);
    }

    #[test]
    fn shared_string_table_deduplicates_and_round_trips() {
        let mut sst = SharedStringTable::default();
        assert_eq!(sst.intern("alpha"), 0);
        assert_eq!(sst.intern(" padded "), 1);
        assert_eq!(sst.intern("alpha"), 0);
        let xml = sst.to_xml(NamespaceFlavor::Transitional);
        let parsed = parse_shared_strings(&xml).unwrap();
        assert_eq!(parsed, vec!["alpha", " padded "]);
    }
}
