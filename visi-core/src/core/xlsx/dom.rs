use quick_xml::events::Event;

pub(crate) const WORKBOOK_SCHEMA_ORDER: &[&str] = &[
    "fileVersion",
    "fileSharing",
    "workbookPr",
    "workbookProtection",
    "bookViews",
    "sheets",
    "functionGroups",
    "externalReferences",
    "definedNames",
    "calcPr",
    "oleSize",
    "customWorkbookViews",
    "pivotCaches",
    "smartTagPr",
    "smartTagTypes",
    "webPublishing",
    "fileRecoveryPr",
    "webPublishObjects",
    "extLst",
];

pub(crate) const WORKSHEET_SCHEMA_ORDER: &[&str] = &[
    "sheetPr",
    "dimension",
    "sheetViews",
    "sheetFormatPr",
    "cols",
    "sheetData",
    "sheetCalcPr",
    "sheetProtection",
    "protectedRanges",
    "scenarios",
    "autoFilter",
    "sortState",
    "dataConsolidate",
    "customSheetViews",
    "mergeCells",
    "phoneticPr",
    "conditionalFormatting",
    "dataValidations",
    "hyperlinks",
    "printOptions",
    "pageMargins",
    "pageSetup",
    "headerFooter",
    "rowBreaks",
    "colBreaks",
    "customProperties",
    "cellWatches",
    "ignoredErrors",
    "smartTags",
    "drawing",
    "legacyDrawing",
    "legacyDrawingHF",
    "drawingHF",
    "picture",
    "oleObjects",
    "controls",
    "webPublishItems",
    "tableParts",
    "extLst",
];

pub(crate) const STYLESHEET_SCHEMA_ORDER: &[&str] = &[
    "numFmts",
    "fonts",
    "fills",
    "borders",
    "cellStyleXfs",
    "cellXfs",
    "cellStyles",
    "dxfs",
    "tableStyles",
    "colors",
    "extLst",
];

pub(crate) const TABLE_SCHEMA_ORDER: &[&str] = &[
    "autoFilter",
    "sortState",
    "tableColumns",
    "tableStyleInfo",
    "extLst",
];

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) enum XmlNode {
    Element(XmlElement),
    Text(String),
    CData(String),
    Comment(String),
    PI(String),
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct XmlElement {
    pub(crate) name: String,
    pub(crate) attrs: Vec<(String, String)>,
    pub(crate) children: Vec<XmlNode>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct XmlDocument {
    pub(crate) xml_decl: Option<String>,
    pub(crate) root: XmlElement,
}

impl XmlElement {
    pub(crate) fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            attrs: Vec::new(),
            children: Vec::new(),
        }
    }

    pub(crate) fn local_name(&self) -> &str {
        match self.name.find(':') {
            Some(pos) => &self.name[pos + 1..],
            None => &self.name,
        }
    }

    pub(crate) fn prefix(&self) -> Option<&str> {
        self.name.find(':').map(|pos| &self.name[..pos])
    }

    pub(crate) fn qualified_child_name(&self, local: &str) -> String {
        match self.prefix() {
            Some(prefix) => format!("{prefix}:{local}"),
            None => local.to_string(),
        }
    }

    pub(crate) fn attr(&self, name: &str) -> Option<&str> {
        self.attrs
            .iter()
            .find(|(k, _)| k == name)
            .map(|(_, v)| v.as_str())
    }

    pub(crate) fn attr_local(&self, local: &str) -> Option<&str> {
        self.attrs
            .iter()
            .find(|(k, _)| match k.find(':') {
                Some(pos) => &k[pos + 1..] == local,
                None => k == local,
            })
            .map(|(_, v)| v.as_str())
    }

    pub(crate) fn set_attr(&mut self, name: impl Into<String>, value: impl Into<String>) {
        let name = name.into();
        let value = value.into();
        if let Some((_, existing)) = self.attrs.iter_mut().find(|(k, _)| k == &name) {
            *existing = value;
        } else {
            self.attrs.push((name, value));
        }
    }

    pub(crate) fn remove_attr(&mut self, name: &str) -> bool {
        let len = self.attrs.len();
        self.attrs.retain(|(k, _)| k != name);
        self.attrs.len() != len
    }

    pub(crate) fn child(&self, local: &str) -> Option<&XmlElement> {
        self.children.iter().find_map(|node| match node {
            XmlNode::Element(el) if el.local_name() == local => Some(el),
            _ => None,
        })
    }

    pub(crate) fn child_mut(&mut self, local: &str) -> Option<&mut XmlElement> {
        self.children.iter_mut().find_map(|node| match node {
            XmlNode::Element(el) if el.local_name() == local => Some(el),
            _ => None,
        })
    }

    pub(crate) fn children_by_local<'a>(
        &'a self,
        local: &'a str,
    ) -> impl Iterator<Item = &'a XmlElement> {
        self.children.iter().filter_map(move |node| match node {
            XmlNode::Element(el) if el.local_name() == local => Some(el),
            _ => None,
        })
    }

    pub(crate) fn child_elements(&self) -> impl Iterator<Item = &XmlElement> {
        self.children.iter().filter_map(|node| match node {
            XmlNode::Element(el) => Some(el),
            _ => None,
        })
    }

    pub(crate) fn remove_children_by_local(&mut self, local: &str) {
        self.children
            .retain(|node| !matches!(node, XmlNode::Element(el) if el.local_name() == local));
    }

    pub(crate) fn text_content(&self) -> String {
        let mut out = String::new();
        for child in &self.children {
            match child {
                XmlNode::Text(t) | XmlNode::CData(t) => out.push_str(t),
                XmlNode::Element(el) => out.push_str(&el.text_content()),
                _ => {}
            }
        }
        out
    }

    pub(crate) fn insert_child_in_schema_order(&mut self, elem: XmlElement, schema_order: &[&str]) {
        let target_rank = schema_order
            .iter()
            .position(|&name| name == elem.local_name());
        let insert_idx = match target_rank {
            None => self.children.len(),
            Some(rank) => {
                let mut pos = self.children.len();
                for (idx, child) in self.children.iter().enumerate() {
                    if let XmlNode::Element(el) = child
                        && let Some(child_rank) = schema_order
                            .iter()
                            .position(|&name| name == el.local_name())
                        && child_rank > rank
                    {
                        pos = idx;
                        break;
                    }
                }
                pos
            }
        };
        self.children.insert(insert_idx, XmlNode::Element(elem));
    }

    pub(crate) fn ensure_child_in_schema_order(
        &mut self,
        local: &str,
        schema_order: &[&str],
    ) -> &mut XmlElement {
        if let Some(idx) = self
            .children
            .iter()
            .position(|node| matches!(node, XmlNode::Element(el) if el.local_name() == local))
        {
            if let XmlNode::Element(ref mut el) = self.children[idx] {
                return el;
            }
            unreachable!();
        }
        let qname = self.qualified_child_name(local);
        let elem = XmlElement::new(qname);
        self.insert_child_in_schema_order(elem, schema_order);
        self.child_mut(local).unwrap()
    }

    pub(crate) fn write_xml(&self, out: &mut String) {
        out.push('<');
        out.push_str(&self.name);
        for (k, v) in &self.attrs {
            out.push(' ');
            out.push_str(k);
            out.push_str("=\"");
            write_escaped_attr(v, out);
            out.push('"');
        }
        if self.children.is_empty() {
            out.push_str("/>");
            return;
        }
        out.push('>');
        for child in &self.children {
            match child {
                XmlNode::Element(el) => el.write_xml(out),
                XmlNode::Text(t) => write_escaped_text(t, out),
                XmlNode::CData(c) => {
                    out.push_str("<![CDATA[");
                    out.push_str(c);
                    out.push_str("]]>");
                }
                XmlNode::Comment(c) => {
                    out.push_str("<!--");
                    out.push_str(c);
                    out.push_str("-->");
                }
                XmlNode::PI(p) => {
                    out.push_str("<?");
                    out.push_str(p);
                    out.push_str("?>");
                }
            }
        }
        out.push_str("</");
        out.push_str(&self.name);
        out.push('>');
    }
}

impl XmlDocument {
    pub(crate) fn new(root: XmlElement) -> Self {
        Self {
            xml_decl: Some(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>".to_string(),
            ),
            root,
        }
    }

    pub(crate) fn parse(xml: &str) -> Result<Self, String> {
        let mut reader = quick_xml::Reader::from_str(xml);
        let mut buf = Vec::new();
        let mut xml_decl = None;
        let mut stack: Vec<XmlElement> = Vec::new();
        let mut root: Option<XmlElement> = None;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Eof) => break,
                Ok(Event::Decl(e)) => {
                    let raw = String::from_utf8_lossy(e.as_ref());
                    xml_decl = Some(format!("<?{raw}?>"));
                }
                Ok(Event::Start(e)) => {
                    stack.push(parse_start_element(&e)?);
                }
                Ok(Event::Empty(e)) => {
                    let elem = parse_start_element(&e)?;
                    if let Some(parent) = stack.last_mut() {
                        parent.children.push(XmlNode::Element(elem));
                    } else {
                        root = Some(elem);
                    }
                }
                Ok(Event::End(_)) => {
                    if let Some(elem) = stack.pop() {
                        if let Some(parent) = stack.last_mut() {
                            parent.children.push(XmlNode::Element(elem));
                        } else {
                            root = Some(elem);
                        }
                    }
                }
                Ok(Event::Text(e)) => {
                    let decoded = e
                        .decode()
                        .map_err(|err| format!("XML text decode error: {err}"))?;
                    let text = quick_xml::escape::unescape(&decoded)
                        .map_err(|err| format!("XML text unescape error: {err}"))?
                        .into_owned();
                    if let Some(parent) = stack.last_mut() {
                        if let Some(XmlNode::Text(existing)) = parent.children.last_mut() {
                            existing.push_str(&text);
                        } else {
                            parent.children.push(XmlNode::Text(text));
                        }
                    }
                }
                Ok(Event::GeneralRef(e)) => {
                    let raw = e
                        .decode()
                        .map_err(|err| format!("XML entity decode error: {err}"))?;
                    let entity = format!("&{raw};");
                    let resolved = quick_xml::escape::unescape(&entity)
                        .map(|v| v.into_owned())
                        .unwrap_or(entity);
                    if let Some(parent) = stack.last_mut() {
                        if let Some(XmlNode::Text(existing)) = parent.children.last_mut() {
                            existing.push_str(&resolved);
                        } else {
                            parent.children.push(XmlNode::Text(resolved));
                        }
                    }
                }
                Ok(Event::CData(e)) => {
                    let text = String::from_utf8_lossy(e.as_ref()).into_owned();
                    if let Some(parent) = stack.last_mut() {
                        parent.children.push(XmlNode::CData(text));
                    }
                }
                Ok(Event::Comment(e)) => {
                    let text = String::from_utf8_lossy(e.as_ref()).into_owned();
                    if let Some(parent) = stack.last_mut() {
                        parent.children.push(XmlNode::Comment(text));
                    }
                }
                Ok(Event::PI(e)) => {
                    let text = String::from_utf8_lossy(e.as_ref()).into_owned();
                    if let Some(parent) = stack.last_mut() {
                        parent.children.push(XmlNode::PI(text));
                    }
                }
                Ok(Event::DocType(_)) => {}
                Err(err) => return Err(format!("XML parse error: {err}")),
            }
            buf.clear();
        }

        let root = root.ok_or_else(|| "XML document has no root element".to_string())?;
        Ok(Self { xml_decl, root })
    }

    pub(crate) fn to_xml_string(&self) -> String {
        let mut out = String::new();
        if let Some(ref decl) = self.xml_decl {
            out.push_str(decl);
        }
        self.root.write_xml(&mut out);
        out
    }

    pub(crate) fn validate_child_sequence(&self, schema_order: &[&str]) -> Result<(), Vec<String>> {
        let mut errors = Vec::new();
        let mut last_rank: Option<(usize, &str)> = None;
        for el in self.root.child_elements() {
            if let Some(rank) = schema_order.iter().position(|&n| n == el.local_name()) {
                if let Some((prev_rank, prev_name)) = last_rank
                    && rank < prev_rank
                {
                    errors.push(format!(
                        "Element <{}> appears after <{prev_name}>, violating XSD sequence order",
                        el.local_name()
                    ));
                }
                last_rank = Some((rank, el.local_name()));
            }
        }
        if errors.is_empty() {
            Ok(())
        } else {
            Err(errors)
        }
    }
}

fn parse_start_element(e: &quick_xml::events::BytesStart) -> Result<XmlElement, String> {
    let name = String::from_utf8(e.name().as_ref().to_vec())
        .map_err(|err| format!("Invalid UTF-8 element name: {err}"))?;
    let mut attrs = Vec::new();
    for attr in e.attributes().flatten() {
        let key = String::from_utf8(attr.key.as_ref().to_vec())
            .map_err(|err| format!("Invalid UTF-8 attribute key: {err}"))?;
        let raw_val = String::from_utf8(attr.value.into_owned())
            .map_err(|err| format!("Invalid UTF-8 attribute value: {err}"))?;
        let val = quick_xml::escape::unescape(&raw_val)
            .map(|v| v.into_owned())
            .unwrap_or(raw_val);
        attrs.push((key, val));
    }
    Ok(XmlElement {
        name,
        attrs,
        children: Vec::new(),
    })
}

fn write_escaped_attr(s: &str, out: &mut String) {
    for ch in s.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\n' => out.push_str("&#xA;"),
            '\r' => out.push_str("&#xD;"),
            '\t' => out.push_str("&#x9;"),
            _ => out.push(ch),
        }
    }
}

fn write_escaped_text(s: &str, out: &mut String) {
    for ch in s.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            _ => out.push(ch),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn dom_preserves_unknown_elements_attributes_and_namespaces() {
        let xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><x:worksheet xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><x:sheetPr codeName=\"Sheet1\"/><x:sheetData><x:row r=\"1\"><x:c r=\"A1\" t=\"inlineStr\"><x:is><x:t xml:space=\"preserve\"> a &amp; b </x:t></x:is></x:c></x:row></x:sheetData><x:extLst><x:ext uri=\"{1234}\"><custom foo=\"bar\"/></x:ext></x:extLst></x:worksheet>";
        let doc = XmlDocument::parse(xml).unwrap();
        assert_eq!(doc.to_xml_string(), xml);
        assert!(doc.validate_child_sequence(WORKSHEET_SCHEMA_ORDER).is_ok());
    }

    #[test]
    fn dom_inserts_children_in_xsd_sequence_order() {
        let xml = "<worksheet><sheetPr/><sheetData/><extLst/></worksheet>";
        let mut doc = XmlDocument::parse(xml).unwrap();
        doc.root
            .insert_child_in_schema_order(XmlElement::new("cols"), WORKSHEET_SCHEMA_ORDER);
        doc.root
            .insert_child_in_schema_order(XmlElement::new("tableParts"), WORKSHEET_SCHEMA_ORDER);
        doc.root
            .insert_child_in_schema_order(XmlElement::new("dimension"), WORKSHEET_SCHEMA_ORDER);

        let names: Vec<&str> = doc
            .root
            .child_elements()
            .map(XmlElement::local_name)
            .collect();
        assert_eq!(
            names,
            vec![
                "sheetPr",
                "dimension",
                "cols",
                "sheetData",
                "tableParts",
                "extLst"
            ]
        );
        assert!(doc.validate_child_sequence(WORKSHEET_SCHEMA_ORDER).is_ok());
    }

    #[test]
    fn dom_sequence_validator_detects_out_of_order_children() {
        let xml = "<worksheet><sheetData/><cols/></worksheet>";
        let doc = XmlDocument::parse(xml).unwrap();
        let errs = doc
            .validate_child_sequence(WORKSHEET_SCHEMA_ORDER)
            .unwrap_err();
        assert_eq!(errs.len(), 1);
        assert!(errs[0].contains("cols"));
    }
}
