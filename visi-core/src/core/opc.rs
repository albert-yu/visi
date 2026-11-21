use std::collections::HashMap;
use std::io::{Cursor, Read, Write};

use quick_xml::events::Event;
use zip::write::SimpleFileOptions;

pub(crate) const NS_CONTENT_TYPES: &str =
    "http://schemas.openxmlformats.org/package/2006/content-types";
pub(crate) const NS_RELATIONSHIPS: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships";

pub(crate) const REL_OFFICE_DOCUMENT: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

pub(crate) const REL_WORKSHEET: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
pub(crate) const REL_SHARED_STRINGS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
pub(crate) const REL_STYLES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
pub(crate) const REL_TABLE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
pub(crate) const REL_DRAWING: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
pub(crate) const REL_CHART: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
pub(crate) const REL_COMMENTS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
pub(crate) const REL_HYPERLINK: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
pub(crate) const REL_METADATA: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata";
pub(crate) const REL_CORE_PROPERTIES: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
pub(crate) const REL_EXTENDED_PROPERTIES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
pub(crate) const REL_PIVOT_CACHE_DEF: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";
pub(crate) const REL_PIVOT_CACHE_RECORDS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords";
pub(crate) const REL_PIVOT_TABLE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";
pub(crate) const REL_VBA_PROJECT: &str =
    "http://schemas.microsoft.com/office/2006/relationships/vbaProject";

pub(crate) const CT_RELS: &str = "application/vnd.openxmlformats-package.relationships+xml";
pub(crate) const CT_XML: &str = "application/xml";
pub(crate) const CT_WORKBOOK: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
pub(crate) const CT_WORKBOOK_MACRO_ENABLED: &str =
    "application/vnd.ms-excel.sheet.macroEnabled.main+xml";
pub(crate) const CT_WORKSHEET: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
pub(crate) const CT_SHARED_STRINGS: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
pub(crate) const CT_STYLES: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
pub(crate) const CT_TABLE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
pub(crate) const CT_DRAWING: &str = "application/vnd.openxmlformats-officedocument.drawing+xml";
pub(crate) const CT_CHART: &str =
    "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
pub(crate) const CT_METADATA: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml";
pub(crate) const CT_COMMENTS: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
pub(crate) const CT_CORE_PROPERTIES: &str =
    "application/vnd.openxmlformats-package.core-properties+xml";
pub(crate) const CT_EXTENDED_PROPERTIES: &str =
    "application/vnd.openxmlformats-officedocument.extended-properties+xml";
pub(crate) const CT_PIVOT_CACHE_DEF: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml";
pub(crate) const CT_PIVOT_CACHE_RECORDS: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml";
pub(crate) const CT_PIVOT_TABLE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml";
pub(crate) const CT_VBA_PROJECT: &str = "application/vnd.ms-office.vbaProject";

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Relationship {
    pub(crate) id: String,
    pub(crate) rel_type: String,
    pub(crate) target: String,
    pub(crate) target_mode: Option<String>,
}

impl Relationship {
    pub(crate) fn is_external(&self) -> bool {
        self.target_mode
            .as_deref()
            .is_some_and(|m| m.eq_ignore_ascii_case("External"))
    }

    pub(crate) fn resolved_target(&self, source_part: &str) -> String {
        if self.is_external() {
            self.target.clone()
        } else {
            resolve_target(source_part, &self.target)
        }
    }
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub(crate) struct ContentTypes {
    pub(crate) defaults: Vec<(String, String)>,
    pub(crate) overrides: Vec<(String, String)>,
}

impl ContentTypes {
    pub(crate) fn parse(xml: &str) -> Self {
        let mut out = Self::default();
        let mut reader = quick_xml::Reader::from_str(xml);
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                    match e.local_name().as_ref() {
                        b"Default" => {
                            if let (Some(ext), Some(ct)) = (
                                get_attr_local(e, b"Extension"),
                                get_attr_local(e, b"ContentType"),
                            ) {
                                out.defaults.push((ext, ct));
                            }
                        }
                        b"Override" => {
                            if let (Some(part), Some(ct)) = (
                                get_attr_local(e, b"PartName"),
                                get_attr_local(e, b"ContentType"),
                            ) {
                                out.overrides.push((part, ct));
                            }
                        }
                        _ => {}
                    }
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
            buf.clear();
        }
        out
    }

    pub(crate) fn ensure_default(&mut self, extension: &str, content_type: &str) {
        if !self
            .defaults
            .iter()
            .any(|(ext, _)| ext.eq_ignore_ascii_case(extension))
        {
            self.defaults
                .push((extension.to_string(), content_type.to_string()));
        }
    }

    pub(crate) fn set_override(&mut self, part_name: &str, content_type: &str) {
        let normalized = format!("/{}", normalize_part_name(part_name));
        if let Some((_, existing)) = self
            .overrides
            .iter_mut()
            .find(|(p, _)| p.eq_ignore_ascii_case(&normalized))
        {
            *existing = content_type.to_string();
        } else {
            self.overrides.push((normalized, content_type.to_string()));
        }
    }

    pub(crate) fn remove_override(&mut self, part_name: &str) -> bool {
        let normalized = format!("/{}", normalize_part_name(part_name));
        let before = self.overrides.len();
        self.overrides
            .retain(|(p, _)| !p.eq_ignore_ascii_case(&normalized));
        self.overrides.len() != before
    }

    #[cfg(test)]
    pub(crate) fn content_type_for(&self, part_name: &str) -> Option<&str> {
        let normalized = format!("/{}", normalize_part_name(part_name));
        if let Some((_, ct)) = self
            .overrides
            .iter()
            .find(|(p, _)| p.eq_ignore_ascii_case(&normalized))
        {
            return Some(ct.as_str());
        }
        let ext = part_name.rsplit('.').next()?;
        self.defaults
            .iter()
            .find(|(e, _)| e.eq_ignore_ascii_case(ext))
            .map(|(_, ct)| ct.as_str())
    }

    pub(crate) fn to_xml(&self) -> String {
        let mut xml = format!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"{NS_CONTENT_TYPES}\">"
        );
        for (ext, ct) in &self.defaults {
            xml.push_str("<Default Extension=\"");
            xml.push_str(&escape_xml_attr(ext));
            xml.push_str("\" ContentType=\"");
            xml.push_str(&escape_xml_attr(ct));
            xml.push_str("\"/>");
        }
        for (part, ct) in &self.overrides {
            xml.push_str("<Override PartName=\"");
            xml.push_str(&escape_xml_attr(part));
            xml.push_str("\" ContentType=\"");
            xml.push_str(&escape_xml_attr(ct));
            xml.push_str("\"/>");
        }
        xml.push_str("</Types>");
        xml
    }
}

#[derive(Debug, Clone, Default)]
pub(crate) struct OpcPackage {
    parts: Vec<(String, Vec<u8>)>,
    index_by_lower: HashMap<String, usize>,
}

impl OpcPackage {
    pub(crate) fn new() -> Self {
        Self::default()
    }

    pub(crate) fn from_bytes(buffer: &[u8]) -> Result<Self, String> {
        let mut archive = zip::ZipArchive::new(Cursor::new(buffer))
            .map_err(|e| format!("Failed to open OPC ZIP package: {e}"))?;
        let mut pkg = Self::new();
        for idx in 0..archive.len() {
            let mut entry = archive
                .by_index(idx)
                .map_err(|e| format!("Failed to read ZIP entry {idx}: {e}"))?;
            if entry.is_dir() {
                continue;
            }
            let name = normalize_part_name(entry.name());
            let mut bytes = Vec::with_capacity(entry.size().min(16 * 1024 * 1024) as usize);
            entry
                .read_to_end(&mut bytes)
                .map_err(|e| format!("Failed to decompress ZIP entry '{name}': {e}"))?;
            pkg.set_part(name, bytes);
        }
        Ok(pkg)
    }

    pub(crate) fn to_bytes(&self) -> Result<Vec<u8>, String> {
        let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let options = SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated)
            .compression_level(Some(6));
        for (name, bytes) in &self.parts {
            writer
                .start_file(name, options)
                .map_err(|e| format!("Failed to start ZIP entry '{name}': {e}"))?;
            writer
                .write_all(bytes)
                .map_err(|e| format!("Failed to write ZIP entry '{name}': {e}"))?;
        }
        writer
            .finish()
            .map_err(|e| format!("Failed to finalize OPC ZIP package: {e}"))
            .map(Cursor::into_inner)
    }

    pub(crate) fn set_part(&mut self, part_name: impl Into<String>, bytes: Vec<u8>) {
        let normalized = normalize_part_name(&part_name.into());
        let key = normalized.to_ascii_lowercase();
        if let Some(&idx) = self.index_by_lower.get(&key) {
            self.parts[idx].1 = bytes;
        } else {
            let idx = self.parts.len();
            self.parts.push((normalized, bytes));
            self.index_by_lower.insert(key, idx);
        }
    }

    pub(crate) fn has_part(&self, part_name: &str) -> bool {
        let key = normalize_part_name(part_name).to_ascii_lowercase();
        self.index_by_lower.contains_key(&key)
    }

    pub(crate) fn remove_part(&mut self, part_name: &str) -> bool {
        let key = normalize_part_name(part_name).to_ascii_lowercase();
        let Some(idx) = self.index_by_lower.remove(&key) else {
            return false;
        };
        self.parts.remove(idx);
        self.index_by_lower.clear();
        for (i, (name, _)) in self.parts.iter().enumerate() {
            self.index_by_lower.insert(name.to_ascii_lowercase(), i);
        }
        true
    }

    pub(crate) fn get_bytes(&self, part_name: &str) -> Option<&[u8]> {
        let key = normalize_part_name(part_name).to_ascii_lowercase();
        let &idx = self.index_by_lower.get(&key)?;
        Some(&self.parts[idx].1)
    }

    pub(crate) fn get_str(&self, part_name: &str) -> Option<String> {
        let bytes = self.get_bytes(part_name)?;
        let text = String::from_utf8(bytes.to_vec()).ok()?;
        if let Some(stripped) = text.strip_prefix('\u{FEFF}') {
            Some(stripped.to_string())
        } else {
            Some(text)
        }
    }

    pub(crate) fn read_relationships(&self, source_part: &str) -> Vec<Relationship> {
        let rels_path = rels_path_for(source_part);
        let Some(xml) = self.get_str(&rels_path) else {
            return Vec::new();
        };
        parse_relationships(&xml)
    }

    pub(crate) fn write_relationships(&mut self, source_part: &str, rels: &[Relationship]) {
        let rels_path = rels_path_for(source_part);
        self.set_part(rels_path, serialize_relationships(rels).into_bytes());
    }

    pub(crate) fn read_content_types(&self) -> ContentTypes {
        self.get_str("[Content_Types].xml")
            .map(|xml| ContentTypes::parse(&xml))
            .unwrap_or_default()
    }

    pub(crate) fn write_content_types(&mut self, content_types: &ContentTypes) {
        self.set_part("[Content_Types].xml", content_types.to_xml().into_bytes());
    }

    pub(crate) fn find_workbook_part(&self) -> Option<String> {
        let root_rels = self.read_relationships("");
        for rel in &root_rels {
            if !rel.is_external() && rel_type_matches(&rel.rel_type, "officeDocument") {
                return Some(rel.resolved_target(""));
            }
        }
        if self.has_part("xl/workbook.xml") {
            return Some("xl/workbook.xml".to_string());
        }
        None
    }

    #[cfg(test)]
    pub(crate) fn validate_opc(&self) -> Result<(), Vec<String>> {
        let mut errors = Vec::new();
        if !self.has_part("[Content_Types].xml") {
            errors.push("Missing required [Content_Types].xml".to_string());
        }
        if !self.has_part("_rels/.rels") {
            errors.push("Missing required _rels/.rels".to_string());
        }
        let ct = self.read_content_types();
        for (part_name, _) in &self.parts {
            if part_name.eq_ignore_ascii_case("[Content_Types].xml") {
                continue;
            }
            if ct.content_type_for(part_name).is_none() {
                errors.push(format!(
                    "Part '{part_name}' has no matching content type in [Content_Types].xml"
                ));
            }
            if part_name.to_ascii_lowercase().ends_with(".rels") {
                let source_part = source_part_for_rels(part_name);
                if let Some(xml) = self.get_str(part_name) {
                    for rel in parse_relationships(&xml) {
                        if !rel.is_external() {
                            let target = rel.resolved_target(&source_part);
                            if !self.has_part(&target) {
                                errors.push(format!(
                                    "Dangling relationship '{}' in '{}' -> '{}'",
                                    rel.id, part_name, target
                                ));
                            }
                        }
                    }
                }
            }
        }
        if errors.is_empty() {
            Ok(())
        } else {
            Err(errors)
        }
    }
}

pub(crate) fn normalize_part_name(path: &str) -> String {
    let replaced = path.replace('\\', "/");
    let trimmed = replaced.trim_start_matches('/');
    let mut segments: Vec<&str> = Vec::new();
    for seg in trimmed.split('/') {
        match seg {
            "" | "." => {}
            ".." => {
                segments.pop();
            }
            other => segments.push(other),
        }
    }
    segments.join("/")
}

pub(crate) fn resolve_target(source_part: &str, target: &str) -> String {
    let target = target.replace('\\', "/");
    if let Some(stripped) = target.strip_prefix('/') {
        return normalize_part_name(stripped);
    }
    let normalized_source = normalize_part_name(source_part);
    let base_dir = match normalized_source.rfind('/') {
        Some(pos) => &normalized_source[..pos],
        None => "",
    };
    if base_dir.is_empty() {
        normalize_part_name(&target)
    } else {
        normalize_part_name(&format!("{base_dir}/{target}"))
    }
}

pub(crate) fn rels_path_for(part_name: &str) -> String {
    let normalized = normalize_part_name(part_name);
    if normalized.is_empty() {
        return "_rels/.rels".to_string();
    }
    match normalized.rfind('/') {
        Some(pos) => format!(
            "{}/_rels/{}.rels",
            &normalized[..pos],
            &normalized[pos + 1..]
        ),
        None => format!("_rels/{normalized}.rels"),
    }
}

#[cfg(test)]
fn source_part_for_rels(rels_part: &str) -> String {
    let normalized = normalize_part_name(rels_part);
    if normalized.eq_ignore_ascii_case("_rels/.rels") {
        return String::new();
    }
    let without_suffix = normalized
        .strip_suffix(".rels")
        .or_else(|| normalized.strip_suffix(".RELS"))
        .unwrap_or(&normalized);
    if let Some(pos) = without_suffix.rfind("/_rels/") {
        format!(
            "{}/{}",
            &without_suffix[..pos],
            &without_suffix[pos + "/_rels/".len()..]
        )
    } else if let Some(rest) = without_suffix.strip_prefix("_rels/") {
        rest.to_string()
    } else {
        without_suffix.to_string()
    }
}

pub(crate) fn rel_type_matches(actual: &str, suffix: &str) -> bool {
    actual
        .rsplit('/')
        .next()
        .is_some_and(|last| last.eq_ignore_ascii_case(suffix))
        && actual.contains("/relationships/")
}

pub(crate) fn parse_relationships(xml: &str) -> Vec<Relationship> {
    let mut rels = Vec::new();
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                if e.local_name().as_ref() == b"Relationship"
                    && let (Some(id), Some(rel_type), Some(target)) = (
                        get_attr_local(e, b"Id"),
                        get_attr_local(e, b"Type"),
                        get_attr_local(e, b"Target"),
                    )
                {
                    let target_mode = get_attr_local(e, b"TargetMode");
                    rels.push(Relationship {
                        id,
                        rel_type,
                        target,
                        target_mode,
                    });
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    rels
}

pub(crate) fn serialize_relationships(rels: &[Relationship]) -> String {
    let mut xml = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"{NS_RELATIONSHIPS}\">"
    );
    for rel in rels {
        xml.push_str("<Relationship Id=\"");
        xml.push_str(&escape_xml_attr(&rel.id));
        xml.push_str("\" Type=\"");
        xml.push_str(&escape_xml_attr(&rel.rel_type));
        xml.push_str("\" Target=\"");
        xml.push_str(&escape_xml_attr(&rel.target));
        xml.push('"');
        if let Some(ref mode) = rel.target_mode {
            xml.push_str(" TargetMode=\"");
            xml.push_str(&escape_xml_attr(mode));
            xml.push('"');
        }
        xml.push_str("/>");
    }
    xml.push_str("</Relationships>");
    xml
}

pub(crate) fn next_rid(rels: &[Relationship]) -> String {
    let max_id = rels
        .iter()
        .filter_map(|r| r.id.strip_prefix("rId")?.parse::<usize>().ok())
        .max()
        .unwrap_or(0);
    format!("rId{}", max_id + 1)
}

fn get_attr_local(e: &quick_xml::events::BytesStart, name: &[u8]) -> Option<String> {
    for attr in e.attributes().flatten() {
        let key = attr.key.as_ref();
        let local = match key.iter().position(|&b| b == b':') {
            Some(pos) => &key[pos + 1..],
            None => key,
        };
        if local == name {
            let raw = String::from_utf8(attr.value.into_owned()).ok()?;
            return Some(
                quick_xml::escape::unescape(&raw)
                    .map(|v| v.into_owned())
                    .unwrap_or(raw),
            );
        }
    }
    None
}

fn escape_xml_attr(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn resolves_relative_and_absolute_targets() {
        assert_eq!(
            resolve_target("xl/worksheets/sheet1.xml", "../tables/table1.xml"),
            "xl/tables/table1.xml"
        );
        assert_eq!(
            resolve_target("xl/worksheets/sheet1.xml", "/xl/tables/table1.xml"),
            "xl/tables/table1.xml"
        );
        assert_eq!(
            resolve_target("xl/workbook.xml", "worksheets/sheet1.xml"),
            "xl/worksheets/sheet1.xml"
        );
        assert_eq!(resolve_target("", "xl/workbook.xml"), "xl/workbook.xml");
    }

    #[test]
    fn computes_rels_paths_and_inverse() {
        assert_eq!(rels_path_for(""), "_rels/.rels");
        assert_eq!(
            rels_path_for("xl/workbook.xml"),
            "xl/_rels/workbook.xml.rels"
        );
        assert_eq!(
            rels_path_for("xl/worksheets/sheet1.xml"),
            "xl/worksheets/_rels/sheet1.xml.rels"
        );
        assert_eq!(source_part_for_rels("_rels/.rels"), "");
        assert_eq!(
            source_part_for_rels("xl/_rels/workbook.xml.rels"),
            "xl/workbook.xml"
        );
        assert_eq!(
            source_part_for_rels("xl/worksheets/_rels/sheet1.xml.rels"),
            "xl/worksheets/sheet1.xml"
        );
    }

    #[test]
    fn matches_both_transitional_and_strict_rel_types() {
        assert!(rel_type_matches(REL_OFFICE_DOCUMENT, "officeDocument"));
        assert!(rel_type_matches(
            "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument",
            "officeDocument"
        ));
        assert!(rel_type_matches(
            "http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet",
            "worksheet"
        ));
        assert!(!rel_type_matches(REL_STYLES, "worksheet"));
    }

    #[test]
    fn opc_package_round_trip_and_validation() {
        let mut pkg = OpcPackage::new();
        let mut ct = ContentTypes::default();
        ct.ensure_default("rels", CT_RELS);
        ct.ensure_default("xml", CT_XML);
        ct.set_override("xl/workbook.xml", CT_WORKBOOK);
        pkg.write_content_types(&ct);
        pkg.write_relationships(
            "",
            &[Relationship {
                id: "rId1".to_string(),
                rel_type: REL_OFFICE_DOCUMENT.to_string(),
                target: "xl/workbook.xml".to_string(),
                target_mode: None,
            }],
        );
        pkg.set_part("xl/workbook.xml", b"<workbook/>".to_vec());

        assert!(pkg.validate_opc().is_ok());
        let bytes = pkg.to_bytes().unwrap();
        let reopened = OpcPackage::from_bytes(&bytes).unwrap();
        assert!(reopened.validate_opc().is_ok());
        assert_eq!(
            reopened.find_workbook_part().as_deref(),
            Some("xl/workbook.xml")
        );
        assert_eq!(
            reopened.get_str("XL/WORKBOOK.XML").as_deref(),
            Some("<workbook/>")
        );
    }

    #[test]
    fn opc_validation_catches_dangling_relationship() {
        let mut pkg = OpcPackage::new();
        let mut ct = ContentTypes::default();
        ct.ensure_default("rels", CT_RELS);
        pkg.write_content_types(&ct);
        pkg.write_relationships(
            "",
            &[Relationship {
                id: "rId1".to_string(),
                rel_type: REL_OFFICE_DOCUMENT.to_string(),
                target: "xl/missing.xml".to_string(),
                target_mode: None,
            }],
        );
        let errs = pkg.validate_opc().unwrap_err();
        assert!(errs.iter().any(|e| e.contains("Dangling relationship")));
    }
}
