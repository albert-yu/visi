use quick_xml::events::Event;

use crate::core::opc::{OpcPackage, rel_type_matches};
use crate::core::workbook::DefinedName;
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
    pub(crate) defined_names: Vec<DefinedName>,
    pub(crate) shared_strings_part: Option<String>,
    pub(crate) styles_part: Option<String>,
    pub(crate) theme_part: Option<String>,
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
    let mut theme_part = None;

    for rel in &rels {
        if rel.is_external() {
            continue;
        }
        let target = rel.resolved_target(&workbook_part);
        if rel_type_matches(&rel.rel_type, "sharedStrings") && shared_strings_part.is_none() {
            shared_strings_part = Some(target.clone());
        } else if rel_type_matches(&rel.rel_type, "styles") && styles_part.is_none() {
            styles_part = Some(target.clone());
        } else if rel_type_matches(&rel.rel_type, "theme") && theme_part.is_none() {
            theme_part = Some(target.clone());
        }
        rel_targets.insert(rel.id.clone(), (rel.rel_type.clone(), target));
    }

    if shared_strings_part.is_none() && pkg.has_part("xl/sharedStrings.xml") {
        shared_strings_part = Some("xl/sharedStrings.xml".to_string());
    }
    if styles_part.is_none() && pkg.has_part("xl/styles.xml") {
        styles_part = Some("xl/styles.xml".to_string());
    }
    if theme_part.is_none() && pkg.has_part("xl/theme/theme1.xml") {
        theme_part = Some("xl/theme/theme1.xml".to_string());
    }

    let mut flavor = NamespaceFlavor::Transitional;
    let mut date1904 = false;
    let mut sheets = Vec::new();
    let mut defined_names: Vec<DefinedName> = Vec::new();
    let mut active_dn: Option<DefinedName> = None;

    let mut reader = quick_xml::Reader::from_str(&xml);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
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
                b"definedName" => {
                    if let Some(name) = super::get_attr(e, b"name") {
                        let local_sheet_id = super::get_attr(e, b"localSheetId")
                            .and_then(|s| s.parse::<usize>().ok());
                        let hidden = super::get_attr(e, b"hidden")
                            .is_some_and(|v| v == "1" || v.eq_ignore_ascii_case("true"));
                        let comment = super::get_attr(e, b"comment");
                        active_dn = Some(DefinedName {
                            name: decode_xstring(&name),
                            formula: String::new(),
                            local_sheet_id,
                            hidden,
                            comment,
                        });
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(ref e)) => match e.local_name().as_ref() {
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
            Ok(Event::Text(e)) if active_dn.is_some() => {
                if let Ok(decoded) = e.decode()
                    && let Ok(unescaped) = quick_xml::escape::unescape(&decoded)
                    && let Some(ref mut dn) = active_dn
                {
                    dn.formula.push_str(&unescaped);
                }
            }
            Ok(Event::GeneralRef(e)) if active_dn.is_some() => {
                if let Ok(raw) = e.decode()
                    && let Some(ref mut dn) = active_dn
                {
                    let entity = format!("&{raw};");
                    let resolved = quick_xml::escape::unescape(&entity)
                        .map(|v| v.into_owned())
                        .unwrap_or(entity);
                    dn.formula.push_str(&resolved);
                }
            }
            Ok(Event::End(ref e)) => {
                if e.local_name().as_ref() == b"definedName"
                    && let Some(mut dn) = active_dn.take()
                {
                    dn.formula = decode_xstring(dn.formula.trim());
                    defined_names.push(dn);
                }
            }
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
        defined_names,
        shared_strings_part,
        styles_part,
        theme_part,
    })
}

pub(crate) fn parse_sheet_visibility(state: Option<&str>) -> crate::core::SheetVisibility {
    match state {
        Some(s) if s.eq_ignore_ascii_case("hidden") => crate::core::SheetVisibility::Hidden,
        Some(s) if s.eq_ignore_ascii_case("veryHidden") => crate::core::SheetVisibility::VeryHidden,
        _ => crate::core::SheetVisibility::Visible,
    }
}

pub(crate) fn sheet_visibility_attr(
    visibility: crate::core::SheetVisibility,
) -> Option<&'static str> {
    match visibility {
        crate::core::SheetVisibility::Visible => None,
        crate::core::SheetVisibility::Hidden => Some("hidden"),
        crate::core::SheetVisibility::VeryHidden => Some("veryHidden"),
    }
}

pub(crate) fn write_workbook_xml(
    sheets: &[(String, String, crate::core::SheetVisibility)],
    defined_names: &[DefinedName],
    date1904: bool,
    flavor: NamespaceFlavor,
) -> String {
    let ns = flavor.spreadsheetml();
    let ns_r = flavor.office_rels();
    let mut xml = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><workbook xmlns=\"{ns}\" xmlns:r=\"{ns_r}\">"
    );
    if date1904 {
        xml.push_str("<workbookPr date1904=\"1\"/>");
    }
    xml.push_str("<bookViews><workbookView xWindow=\"0\" yWindow=\"0\" windowWidth=\"24000\" windowHeight=\"15000\"/></bookViews><sheets>");
    for (idx, (name, r_id, visibility)) in sheets.iter().enumerate() {
        xml.push_str(&format!(
            "<sheet name=\"{}\" sheetId=\"{}\" r:id=\"{}\"",
            super::escape_xml(name),
            idx + 1,
            super::escape_xml(r_id)
        ));
        if let Some(state) = sheet_visibility_attr(*visibility) {
            xml.push_str(&format!(" state=\"{state}\""));
        }
        xml.push_str("/>");
    }
    xml.push_str("</sheets>");
    if !defined_names.is_empty() {
        xml.push_str("<definedNames>");
        for dn in defined_names {
            xml.push_str(&format!(
                "<definedName name=\"{}\"",
                super::escape_xml(&dn.name)
            ));
            if let Some(local_id) = dn.local_sheet_id {
                xml.push_str(&format!(" localSheetId=\"{local_id}\""));
            }
            if dn.hidden {
                xml.push_str(" hidden=\"1\"");
            }
            if let Some(ref comment) = dn.comment {
                xml.push_str(&format!(" comment=\"{}\"", super::escape_xml(comment)));
            }
            xml.push('>');
            xml.push_str(&super::escape_xml_text(&dn.formula));
            xml.push_str("</definedName>");
        }
        xml.push_str("</definedNames>");
    }
    xml.push_str("<calcPr calcId=\"191029\"/></workbook>");
    xml
}
