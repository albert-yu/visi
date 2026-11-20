use std::collections::HashSet;

use quick_xml::events::Event;

use crate::core::xlsx::xml_util::NamespaceFlavor;

#[derive(Debug, Default)]
pub(crate) struct ParsedMetadata {
    pub(crate) dynamic_array_cm_indices: HashSet<u32>,
}

impl ParsedMetadata {
    pub(crate) fn is_dynamic_array_cm(&self, cm: u32) -> bool {
        self.dynamic_array_cm_indices.contains(&cm)
    }
}

pub(crate) fn parse_metadata_xml(xml: &str) -> ParsedMetadata {
    let mut out = ParsedMetadata::default();
    let mut metadata_types: Vec<String> = Vec::new();
    let mut in_metadata_types = false;
    let mut in_cell_metadata = false;
    let mut cm_index = 0u32;
    let mut current_bk_is_xldapr = false;

    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"metadataTypes" => in_metadata_types = true,
                b"metadataType" if in_metadata_types => {
                    metadata_types.push(super::get_attr(e, b"name").unwrap_or_default());
                }
                b"cellMetadata" => {
                    in_cell_metadata = true;
                    cm_index = 0;
                }
                b"bk" if in_cell_metadata => {
                    cm_index += 1;
                    current_bk_is_xldapr = false;
                }
                b"rc" if in_cell_metadata => {
                    if let Some(t_idx) =
                        super::get_attr(e, b"t").and_then(|s| s.parse::<usize>().ok())
                        && t_idx >= 1
                        && metadata_types
                            .get(t_idx - 1)
                            .is_some_and(|n| n.eq_ignore_ascii_case("XLDAPR"))
                    {
                        current_bk_is_xldapr = true;
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(ref e)) => match e.local_name().as_ref() {
                b"metadataType" if in_metadata_types => {
                    metadata_types.push(super::get_attr(e, b"name").unwrap_or_default());
                }
                b"rc" if in_cell_metadata => {
                    if let Some(t_idx) =
                        super::get_attr(e, b"t").and_then(|s| s.parse::<usize>().ok())
                        && t_idx >= 1
                        && metadata_types
                            .get(t_idx - 1)
                            .is_some_and(|n| n.eq_ignore_ascii_case("XLDAPR"))
                    {
                        current_bk_is_xldapr = true;
                    }
                }
                _ => {}
            },
            Ok(Event::End(ref e)) => match e.local_name().as_ref() {
                b"metadataTypes" => in_metadata_types = false,
                b"cellMetadata" => in_cell_metadata = false,
                b"bk" if in_cell_metadata && current_bk_is_xldapr => {
                    out.dynamic_array_cm_indices.insert(cm_index);
                }
                _ => {}
            },
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    out
}

pub(crate) fn write_dynamic_array_metadata_xml(flavor: NamespaceFlavor) -> String {
    let ns = flavor.spreadsheetml();
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><metadata xmlns=\"{ns}\" xmlns:xda=\"http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray\"><metadataTypes count=\"1\"><metadataType name=\"XLDAPR\" minSupportedVersion=\"120000\" copy=\"1\" pasteAll=\"1\" pasteValues=\"1\" merge=\"1\" splitFirst=\"1\" rowColShift=\"1\" clearFormats=\"1\" clearComments=\"1\" assign=\"1\" coerce=\"1\"/></metadataTypes><futureMetadata name=\"XLDAPR\" count=\"1\"><bk><extLst><ext uri=\"{{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}}\"><xda:dynamicArrayProperties fDynamic=\"1\" fCollapsed=\"0\"/></ext></extLst></bk></futureMetadata><cellMetadata count=\"1\"><bk><rc t=\"1\" v=\"0\"/></bk></cellMetadata></metadata>"
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn dynamic_array_metadata_round_trips() {
        let xml = write_dynamic_array_metadata_xml(NamespaceFlavor::Transitional);
        let parsed = parse_metadata_xml(&xml);
        assert!(parsed.is_dynamic_array_cm(1));
        assert!(!parsed.is_dynamic_array_cm(2));
    }
}
