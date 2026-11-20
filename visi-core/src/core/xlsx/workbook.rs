use quick_xml::events::Event;

use crate::core::opc::{OpcPackage, rel_type_matches};
use crate::core::xlsx::xml_util::{NamespaceFlavor, decode_xstring};

#[derive(Debug, Clone)]
pub(crate) struct WorkbookSheetMeta {
    pub(crate) name: String,
    pub(crate) sheet_id: u32,
    pub(crate) r_id: String,
    pub(crate) state: Option<String>,
    pub(crate) part_path: String,
}

#[derive(Debug, Clone)]
pub(crate) struct ParsedWorkbook {
    pub(crate) workbook_part: String,
    pub(crate) flavor: NamespaceFlavor,
    pub(crate) date1904: bool,
    pub(crate) sheets: Vec<WorkbookSheetMeta>,
    pub(crate) shared_strings_part: Option<String>,
    pub(crate) styles_part: Option<String>,
}

pub(crate) fn read_workbook(pkg: &OpcPackage) -> Result<ParsedWorkbook, String> {
    let workbook_part = pkg
        .find_workbook_part()
        .ok_or_else(|| "Missing workbook part in OPC package".to_string())?;
    let xml = pkg
        .get_str(&workbook_part)
        .ok_or_else(|| format!("Missing workbook XML at '{workbook_part}'"))?;

    let rels = pkg.read_relationships(&workbook_part);
    let mut rel_targets: std::collections::HashMap<String, (String, String)> =
        std::collections::HashMap::new();
    let mut shared_strings_part = None;
    let mut styles_part = None;

    for rel in &rels {
        if rel.is_external() {
            continue;
        }
        let target = rel.resolved_target(&workbook_part);
        if rel_type_matches(&rel.rel_type, "sharedStrings") && shared_strings_part.is_none() {
            shared_strings_part = Some(target.clone());
        } else if rel_type_matches(&rel.rel_type, "styles") && styles_part.is_none() {
            styles_part = Some(target.clone());
        }
        rel_targets.insert(rel.id.clone(), (rel.rel_type.clone(), target));
    }

    if shared_strings_part.is_none() && pkg.has_part("xl/sharedStrings.xml") {
        shared_strings_part = Some("xl/sharedStrings.xml".to_string());
    }
    if styles_part.is_none() && pkg.has_part("xl/styles.xml") {
        styles_part = Some("xl/styles.xml".to_string());
    }

    let mut flavor = NamespaceFlavor::Transitional;
    let mut date1904 = false;
    let mut sheets = Vec::new();

    let mut reader = quick_xml::Reader::from_str(&xml);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => match e.local_name().as_ref() {
                b"workbook" => {
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"xmlns"
                            && let Ok(val) = std::str::from_utf8(attr.value.as_ref())
                        {
                            flavor = NamespaceFlavor::from_xmlns(val);
                        }
                    }
                }
                b"workbookPr" => {
                    if let Some(val) = super::get_attr(e, b"date1904") {
                        date1904 = val == "1" || val.eq_ignore_ascii_case("true");
                    }
                }
                b"sheet" => {
                    if let (Some(name), Some(r_id)) =
                        (super::get_attr(e, b"name"), super::get_attr(e, b"id"))
                    {
                        let sheet_id = super::get_attr(e, b"sheetId")
                            .and_then(|s| s.parse::<u32>().ok())
                            .unwrap_or((sheets.len() + 1) as u32);
                        let state = super::get_attr(e, b"state");
                        if let Some((rel_type, part_path)) = rel_targets.get(&r_id)
                            && rel_type_matches(rel_type, "worksheet")
                        {
                            sheets.push(WorkbookSheetMeta {
                                name: decode_xstring(&name),
                                sheet_id,
                                r_id,
                                state,
                                part_path: part_path.clone(),
                            });
                        }
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(err) => return Err(format!("Failed to parse workbook.xml: {err}")),
            _ => {}
        }
        buf.clear();
    }

    Ok(ParsedWorkbook {
        workbook_part,
        flavor,
        date1904,
        sheets,
        shared_strings_part,
        styles_part,
    })
}

pub(crate) fn write_workbook_xml(sheets: &[(String, String)], flavor: NamespaceFlavor) -> String {
    let ns = flavor.spreadsheetml();
    let ns_r = flavor.office_rels();
    let mut xml = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><workbook xmlns=\"{ns}\" xmlns:r=\"{ns_r}\"><bookViews><workbookView xWindow=\"0\" yWindow=\"0\" windowWidth=\"24000\" windowHeight=\"15000\"/></bookViews><sheets>"
    );
    for (idx, (name, r_id)) in sheets.iter().enumerate() {
        xml.push_str(&format!(
            "<sheet name=\"{}\" sheetId=\"{}\" r:id=\"{}\"/>",
            super::escape_xml(name),
            idx + 1,
            super::escape_xml(r_id)
        ));
    }
    xml.push_str("</sheets><calcPr calcId=\"191029\"/></workbook>");
    xml
}
