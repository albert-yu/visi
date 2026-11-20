use std::collections::{HashMap, HashSet};
use std::io::{Cursor, Read};

use crate::core::xlsx::dom::{
    STYLESHEET_SCHEMA_ORDER, TABLE_SCHEMA_ORDER, WORKBOOK_SCHEMA_ORDER, WORKSHEET_SCHEMA_ORDER,
    XmlDocument, XmlElement, XmlNode,
};
use zip::ZipArchive;

const VALID_ERRORS: &[&str] = &[
    "#NULL!",
    "#DIV/0!",
    "#VALUE!",
    "#REF!",
    "#NAME?",
    "#NUM!",
    "#N/A",
    "#GETTING_DATA",
];

const COMMENTS_SCHEMA_ORDER: &[&str] = &["authors", "commentList", "extLst"];
const METADATA_SCHEMA_ORDER: &[&str] = &[
    "metadataTypes",
    "metadataStrings",
    "mdxMetadata",
    "futureMetadata",
    "cellMetadata",
    "valueMetadata",
    "extLst",
];

fn local_name(name: &str) -> &str {
    match name.rsplit_once(':') {
        Some((_, local)) => local,
        None => name,
    }
}

fn normalize_part(name: &str) -> String {
    let cleaned = name.replace('\\', "/");
    let trimmed = cleaned.trim_start_matches('/');
    let mut parts: Vec<&str> = Vec::new();
    for seg in trimmed.split('/') {
        match seg {
            "" | "." => {}
            ".." => {
                parts.pop();
            }
            s => parts.push(s),
        }
    }
    parts.join("/")
}

fn source_part_for_rels(rels_part: &str) -> String {
    let norm = normalize_part(rels_part);
    if norm.eq_ignore_ascii_case("_rels/.rels") {
        return String::new();
    }
    let without_ext = norm
        .strip_suffix(".rels")
        .or_else(|| norm.strip_suffix(".RELS"))
        .unwrap_or(&norm);
    if let Some((prefix, rest)) = without_ext.rsplit_once("/_rels/") {
        format!("{prefix}/{rest}")
    } else if let Some(rest) = without_ext.strip_prefix("_rels/") {
        rest.to_string()
    } else {
        without_ext.to_string()
    }
}

fn resolve_target(source_part: &str, target: &str) -> String {
    let cleaned = target.replace('\\', "/");
    if cleaned.starts_with('/') {
        return normalize_part(&cleaned);
    }
    if source_part.is_empty() {
        return normalize_part(&cleaned);
    }
    match source_part.rsplit_once('/') {
        Some((dir, _)) => normalize_part(&format!("{dir}/{cleaned}")),
        None => normalize_part(&cleaned),
    }
}

fn parse_cell_ref(s: &str) -> Option<(u32, u32)> {
    let bytes = s.as_bytes();
    let mut i = 0;
    let mut col: u32 = 0;
    while i < bytes.len() && bytes[i].is_ascii_alphabetic() {
        let ch = bytes[i].to_ascii_uppercase();
        col = col.checked_mul(26)?.checked_add((ch - b'A' + 1) as u32)?;
        i += 1;
    }
    if i == 0 || i == bytes.len() {
        return None;
    }
    let row: u32 = s[i..].parse().ok()?;
    if row == 0 {
        return None;
    }
    Some((row, col))
}

fn parse_range_ref(s: &str) -> Option<(u32, u32, u32, u32)> {
    match s.split_once(':') {
        Some((a, b)) => {
            let (r1, c1) = parse_cell_ref(a)?;
            let (r2, c2) = parse_cell_ref(b)?;
            Some((r1, c1, r2, c2))
        }
        None => {
            let (r, c) = parse_cell_ref(s)?;
            Some((r, c, r, c))
        }
    }
}

fn attr_val<'a>(el: &'a XmlElement, key: &str) -> Option<&'a str> {
    el.attrs
        .iter()
        .find(|(k, _)| k == key)
        .map(|(_, v)| v.as_str())
}

fn attr_val_local<'a>(el: &'a XmlElement, local: &str) -> Option<&'a str> {
    el.attrs
        .iter()
        .find(|(k, _)| local_name(k) == local)
        .map(|(_, v)| v.as_str())
}

fn child_elements(el: &XmlElement) -> impl Iterator<Item = &XmlElement> {
    el.children.iter().filter_map(|n| match n {
        XmlNode::Element(e) => Some(e),
        _ => None,
    })
}

fn text_content(el: &XmlElement) -> String {
    let mut out = String::new();
    for node in &el.children {
        match node {
            XmlNode::Text(t) | XmlNode::CData(t) => out.push_str(t),
            _ => {}
        }
    }
    out
}

fn check_sequence(root: &XmlElement, part_name: &str, order: &[&str], errors: &mut Vec<String>) {
    let mut last_rank: Option<usize> = None;
    let mut last_name = "";
    for ch in child_elements(root) {
        let lname = local_name(&ch.name);
        if let Some(rank) = order.iter().position(|&n| n == lname) {
            if let Some(prev) = last_rank
                && rank < prev
            {
                errors.push(format!(
                    "{part_name}: <{lname}> appears after <{last_name}> (violates XSD sequence)"
                ));
            }
            last_rank = Some(rank);
            last_name = lname;
        }
    }
}

fn check_count_attr(el: &XmlElement, part_name: &str, errors: &mut Vec<String>) {
    if let Some(count_str) = attr_val(el, "count") {
        match count_str.parse::<usize>() {
            Ok(declared) => {
                let actual = child_elements(el).count();
                if declared != actual {
                    errors.push(format!(
                        "{part_name}: <{}> count={declared} != child count {actual}",
                        local_name(&el.name)
                    ));
                }
            }
            Err(_) => errors.push(format!(
                "{part_name}: <{}> non-integer count={count_str:?}",
                local_name(&el.name)
            )),
        }
    }
}

fn has_uri_scheme(target: &str) -> bool {
    let Some((scheme, _)) = target.split_once(':') else {
        return false;
    };
    !scheme.is_empty()
        && scheme.as_bytes()[0].is_ascii_alphabetic()
        && scheme
            .bytes()
            .all(|b| b.is_ascii_alphanumeric() || matches!(b, b'+' | b'-' | b'.'))
}

pub(crate) fn validate_ooxml_package(bytes: &[u8]) -> Result<(), Vec<String>> {
    let mut errors = Vec::new();
    let mut archive = match ZipArchive::new(Cursor::new(bytes)) {
        Ok(a) => a,
        Err(e) => return Err(vec![format!("Invalid ZIP container: {e}")]),
    };

    let mut part_bytes: HashMap<String, (String, Vec<u8>)> = HashMap::new();
    for i in 0..archive.len() {
        let mut entry = match archive.by_index(i) {
            Ok(e) => e,
            Err(e) => {
                errors.push(format!("ZIP entry {i} error: {e}"));
                continue;
            }
        };
        if entry.is_dir() {
            continue;
        }
        let raw_name = entry.name().to_string();
        if raw_name.starts_with('/') {
            errors.push(format!("{raw_name}: OPC part name must not start with '/'"));
        }
        if raw_name.replace('\\', "/").split('/').any(|s| s == "..") {
            errors.push(format!("{raw_name}: OPC part name must not contain '..'"));
        }
        let norm_lower = normalize_part(&raw_name).to_ascii_lowercase();
        let mut buf = Vec::new();
        if let Err(e) = entry.read_to_end(&mut buf) {
            errors.push(format!("{raw_name}: read error: {e}"));
            continue;
        }
        if let Some((existing, _)) = part_bytes.get(&norm_lower) {
            errors.push(format!(
                "{raw_name}: duplicate OPC part name colliding with {existing}"
            ));
        } else {
            part_bytes.insert(norm_lower, (raw_name, buf));
        }
    }

    if !part_bytes.contains_key("[content_types].xml") {
        errors.push("Missing [Content_Types].xml".to_string());
    }
    if !part_bytes.contains_key("_rels/.rels") {
        errors.push("Missing _rels/.rels".to_string());
    }

    let mut default_exts: HashSet<String> = HashSet::new();
    let mut override_parts: HashSet<String> = HashSet::new();

    if let Some((_, ct_raw)) = part_bytes.get("[content_types].xml") {
        match std::str::from_utf8(ct_raw)
            .map_err(|e| e.to_string())
            .and_then(XmlDocument::parse)
        {
            Ok(doc) => {
                for ch in child_elements(&doc.root) {
                    match local_name(&ch.name) {
                        "Default" => {
                            if let Some(ext) = attr_val(ch, "Extension") {
                                let ext_lower = ext.to_ascii_lowercase();
                                if !default_exts.insert(ext_lower.clone()) {
                                    errors.push(format!(
                                        "[Content_Types].xml: duplicate Default Extension={ext_lower:?}"
                                    ));
                                }
                            }
                        }
                        "Override" => {
                            if let Some(pname) = attr_val(ch, "PartName") {
                                let norm = normalize_part(pname).to_ascii_lowercase();
                                if !override_parts.insert(norm) {
                                    errors.push(format!(
                                        "[Content_Types].xml: duplicate Override PartName={pname:?}"
                                    ));
                                }
                            }
                        }
                        _ => {}
                    }
                }
            }
            Err(e) => errors.push(format!("[Content_Types].xml: XML parse error: {e}")),
        }
    }

    let mut parsed_docs: HashMap<String, XmlDocument> = HashMap::new();
    let mut rels_by_source: HashMap<String, HashSet<String>> = HashMap::new();

    for (norm_lower, (orig_name, raw)) in &part_bytes {
        if norm_lower != "[content_types].xml" {
            let ext = norm_lower.rsplit_once('.').map(|(_, e)| e).unwrap_or("");
            if !override_parts.contains(norm_lower) && !default_exts.contains(ext) {
                errors.push(format!(
                    "{orig_name}: no matching Default or Override in [Content_Types].xml"
                ));
            }
        }
        if !norm_lower.ends_with(".xml") && !norm_lower.ends_with(".rels") {
            continue;
        }
        let doc = match std::str::from_utf8(raw)
            .map_err(|e| e.to_string())
            .and_then(XmlDocument::parse)
        {
            Ok(d) => d,
            Err(e) => {
                errors.push(format!("{orig_name}: XML parse error: {e}"));
                continue;
            }
        };

        match local_name(&doc.root.name) {
            "workbook" => check_sequence(&doc.root, orig_name, WORKBOOK_SCHEMA_ORDER, &mut errors),
            "worksheet" => {
                check_sequence(&doc.root, orig_name, WORKSHEET_SCHEMA_ORDER, &mut errors)
            }
            "styleSheet" => {
                check_sequence(&doc.root, orig_name, STYLESHEET_SCHEMA_ORDER, &mut errors)
            }
            "table" => check_sequence(&doc.root, orig_name, TABLE_SCHEMA_ORDER, &mut errors),
            "comments" => check_sequence(&doc.root, orig_name, COMMENTS_SCHEMA_ORDER, &mut errors),
            "metadata" => check_sequence(&doc.root, orig_name, METADATA_SCHEMA_ORDER, &mut errors),
            _ => {}
        }

        if norm_lower.ends_with(".rels") {
            let src = source_part_for_rels(orig_name).to_ascii_lowercase();
            let mut seen_ids = HashSet::new();
            for rel in child_elements(&doc.root) {
                if local_name(&rel.name) != "Relationship" {
                    continue;
                }
                let rid = attr_val(rel, "Id").unwrap_or("");
                let rtype = attr_val(rel, "Type").unwrap_or("");
                let target = attr_val(rel, "Target").unwrap_or("");
                let mode = attr_val(rel, "TargetMode").unwrap_or("");
                if rid.is_empty() {
                    errors.push(format!("{orig_name}: Relationship missing Id"));
                } else if !seen_ids.insert(rid.to_string()) {
                    errors.push(format!("{orig_name}: duplicate Relationship Id={rid:?}"));
                }
                if rtype.is_empty() {
                    errors.push(format!("{orig_name}: Relationship {rid:?} missing Type"));
                }
                if mode.eq_ignore_ascii_case("external") {
                    continue;
                }
                if has_uri_scheme(target) {
                    errors.push(format!(
                        "{orig_name}: Relationship {rid:?} has URI target {target:?} without TargetMode='External'"
                    ));
                    continue;
                }
                let resolved = resolve_target(&src, target).to_ascii_lowercase();
                if !part_bytes.contains_key(&resolved) {
                    errors.push(format!(
                        "{orig_name}: dangling relationship {rid} -> {target} ({resolved})"
                    ));
                }
            }
            rels_by_source.insert(src, seen_ids);
        }

        parsed_docs.insert(norm_lower.clone(), doc);
    }

    let mut sst_count: Option<usize> = None;
    let mut cell_xfs_count: Option<usize> = None;

    for (norm_lower, doc) in &parsed_docs {
        let orig_name = &part_bytes[norm_lower].0;
        match local_name(&doc.root.name) {
            "sst" => {
                sst_count = Some(
                    child_elements(&doc.root)
                        .filter(|e| local_name(&e.name) == "si")
                        .count(),
                );
            }
            "styleSheet" => {
                for ch in child_elements(&doc.root) {
                    let lname = local_name(&ch.name);
                    if matches!(
                        lname,
                        "numFmts"
                            | "fonts"
                            | "fills"
                            | "borders"
                            | "cellStyleXfs"
                            | "cellXfs"
                            | "cellStyles"
                            | "dxfs"
                    ) {
                        check_count_attr(ch, orig_name, &mut errors);
                    }
                    if lname == "fills" {
                        let fill_count = child_elements(ch).count();
                        if fill_count < 2 {
                            errors.push(format!(
                                "{orig_name}: <fills> must have >= 2 entries, found {fill_count}"
                            ));
                        }
                    }
                    if lname == "cellXfs" {
                        let count = child_elements(ch).count();
                        cell_xfs_count = Some(count);
                        if count < 1 {
                            errors.push(format!("{orig_name}: <cellXfs> must have >= 1 entry"));
                        }
                    }
                }
            }
            _ => {}
        }
    }

    for (norm_lower, doc) in &parsed_docs {
        let orig_name = &part_bytes[norm_lower].0;
        match local_name(&doc.root.name) {
            "workbook" => {
                let wb_rels = rels_by_source.get(norm_lower);
                let mut seen_names = HashSet::new();
                let mut seen_ids = HashSet::new();
                for ch in child_elements(&doc.root) {
                    if local_name(&ch.name) != "sheets" {
                        continue;
                    }
                    for sheet_el in child_elements(ch) {
                        if local_name(&sheet_el.name) != "sheet" {
                            continue;
                        }
                        let sname = attr_val(sheet_el, "name").unwrap_or("");
                        let sid = attr_val(sheet_el, "sheetId").unwrap_or("");
                        let rid = attr_val_local(sheet_el, "id").unwrap_or("");
                        if sname.is_empty()
                            || sname.chars().count() > 31
                            || sname
                                .chars()
                                .any(|c| matches!(c, '\\' | '/' | '?' | '*' | '[' | ']' | ':'))
                        {
                            errors.push(format!("{orig_name}: invalid sheet name {sname:?}"));
                        } else if !seen_names.insert(sname.to_ascii_lowercase()) {
                            errors.push(format!("{orig_name}: duplicate sheet name {sname:?}"));
                        }
                        if sid.parse::<u32>().ok().filter(|&n| n >= 1).is_none() {
                            errors.push(format!("{orig_name}: invalid sheetId {sid:?}"));
                        } else if !seen_ids.insert(sid.to_string()) {
                            errors.push(format!("{orig_name}: duplicate sheetId {sid:?}"));
                        }
                        if !rid.is_empty() && !wb_rels.is_some_and(|r| r.contains(rid)) {
                            errors.push(format!(
                                "{orig_name}: sheet {sname:?} references missing relationship {rid:?}"
                            ));
                        }
                    }
                }
            }
            "worksheet" => {
                for ch in child_elements(&doc.root) {
                    match local_name(&ch.name) {
                        "mergeCells" => {
                            check_count_attr(ch, orig_name, &mut errors);
                            let mut rects: Vec<(u32, u32, u32, u32, String)> = Vec::new();
                            for mc in child_elements(ch) {
                                if local_name(&mc.name) != "mergeCell" {
                                    continue;
                                }
                                let mref = attr_val(mc, "ref").unwrap_or("");
                                match parse_range_ref(mref) {
                                    Some((r1, c1, r2, c2)) => {
                                        for (pr1, pc1, pr2, pc2, pref) in &rects {
                                            if !(r2 < *pr1 || r1 > *pr2 || c2 < *pc1 || c1 > *pc2) {
                                                errors.push(format!(
                                                    "{orig_name}: overlapping mergeCells {pref:?} and {mref:?}"
                                                ));
                                            }
                                        }
                                        rects.push((r1, c1, r2, c2, mref.to_string()));
                                    }
                                    None => errors.push(format!(
                                        "{orig_name}: invalid mergeCell ref {mref:?}"
                                    )),
                                }
                            }
                        }
                        "sheetData" => {
                            let mut prev_row: u32 = 0;
                            for row_el in child_elements(ch) {
                                if local_name(&row_el.name) != "row" {
                                    continue;
                                }
                                let cur_row = match attr_val(row_el, "r") {
                                    Some(r_str) => match r_str.parse::<u32>() {
                                        Ok(r) if r >= 1 => {
                                            if r <= prev_row {
                                                errors.push(format!(
                                                    "{orig_name}: row/@r={r} not strictly ascending after {prev_row}"
                                                ));
                                            }
                                            r
                                        }
                                        _ => {
                                            errors.push(format!(
                                                "{orig_name}: invalid row/@r={r_str:?}"
                                            ));
                                            continue;
                                        }
                                    },
                                    None => prev_row + 1,
                                };
                                prev_row = cur_row;
                                let mut prev_col: u32 = 0;
                                for c_el in child_elements(row_el) {
                                    if local_name(&c_el.name) != "c" {
                                        continue;
                                    }
                                    let cref = attr_val(c_el, "r");
                                    if let Some(r_str) = cref {
                                        match parse_cell_ref(r_str) {
                                            Some((crow, ccol)) => {
                                                if crow != cur_row {
                                                    errors.push(format!(
                                                        "{orig_name}: cell {r_str:?} row {crow} != parent row {cur_row}"
                                                    ));
                                                }
                                                if ccol <= prev_col {
                                                    errors.push(format!(
                                                        "{orig_name}: cell {r_str:?} not strictly ascending in row {cur_row}"
                                                    ));
                                                }
                                                prev_col = ccol;
                                            }
                                            None => errors.push(format!(
                                                "{orig_name}: invalid cell ref {r_str:?}"
                                            )),
                                        }
                                    }
                                    if let (Some(s_str), Some(max_xfs)) =
                                        (attr_val(c_el, "s"), cell_xfs_count)
                                        && s_str
                                            .parse::<usize>()
                                            .ok()
                                            .is_none_or(|idx| idx >= max_xfs)
                                    {
                                        errors.push(format!(
                                            "{orig_name}: cell {:?} s={s_str:?} out of bounds (cellXfs={max_xfs})",
                                            cref.unwrap_or("?")
                                        ));
                                    }
                                    let t_attr = attr_val(c_el, "t").unwrap_or("n");
                                    let v_text = child_elements(c_el)
                                        .find(|sub| local_name(&sub.name) == "v")
                                        .map(text_content);
                                    if let Some(v) = v_text {
                                        match t_attr {
                                            "s" => {
                                                if let Some(max_sst) = sst_count
                                                    && v.parse::<usize>()
                                                        .ok()
                                                        .is_none_or(|idx| idx >= max_sst)
                                                {
                                                    errors.push(format!(
                                                        "{orig_name}: cell {:?} sst index {v:?} out of bounds (sst={max_sst})",
                                                        cref.unwrap_or("?")
                                                    ));
                                                }
                                            }
                                            "b" => {
                                                if v != "0" && v != "1" {
                                                    errors.push(format!(
                                                        "{orig_name}: cell {:?} boolean {v:?} must be '0' or '1'",
                                                        cref.unwrap_or("?")
                                                    ));
                                                }
                                            }
                                            "e" if !VALID_ERRORS.contains(&v.as_str()) => {
                                                errors.push(format!(
                                                    "{orig_name}: cell {:?} invalid error literal {v:?}",
                                                    cref.unwrap_or("?")
                                                ));
                                            }
                                            _ => {}
                                        }
                                    }
                                }
                            }
                        }
                        _ => {}
                    }
                }
            }
            "comments" => {
                let mut seen_refs = HashSet::new();
                for ch in child_elements(&doc.root) {
                    if local_name(&ch.name) != "commentList" {
                        continue;
                    }
                    for cm in child_elements(ch) {
                        if local_name(&cm.name) != "comment" {
                            continue;
                        }
                        let cref = attr_val(cm, "ref").unwrap_or("");
                        if parse_cell_ref(cref).is_none() {
                            errors.push(format!("{orig_name}: invalid comment ref {cref:?}"));
                        } else if !seen_refs.insert(cref.to_ascii_uppercase()) {
                            errors.push(format!("{orig_name}: duplicate comment ref {cref:?}"));
                        }
                    }
                }
            }
            _ => {}
        }
    }

    if errors.is_empty() {
        Ok(())
    } else {
        Err(errors)
    }
}

pub(crate) fn assert_valid_ooxml(bytes: &[u8]) {
    if let Err(errors) = validate_ooxml_package(bytes) {
        panic!("OOXML spec validation failed:\n{}", errors.join("\n"));
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::core::{
        CellStyle, DefinedName, FreezePane, MergedCellRange, SheetVisibility, WorkbookManager,
    };

    #[test]
    fn test_fresh_and_patched_workbook_passes_ooxml_spec_validation() {
        let mut mgr = WorkbookManager::new_empty().unwrap();
        mgr.add_sheet("Summary").unwrap();
        mgr.set_cell(0, 0, 0, "10".to_string());
        mgr.set_cell(0, 0, 1, "20".to_string());
        mgr.set_cell(0, 0, 2, "=SUM(A1:B1)".to_string());
        mgr.set_cell(0, 1, 0, "  padded  ".to_string());
        mgr.set_cell(0, 1, 1, "TRUE".to_string());
        mgr.set_cell(0, 1, 2, "=1/0".to_string());
        mgr.set_cell(0, 2, 0, "=XLOOKUP(10, A1:A1, B1:B1)".to_string());
        mgr.sheets[0].merged_cells.push(MergedCellRange {
            start_row: 4,
            start_col: 0,
            end_row: 5,
            end_col: 1,
        });
        mgr.sheets[0].freeze_pane = Some(FreezePane { rows: 1, cols: 1 });
        mgr.sheets[1].visibility = SheetVisibility::Hidden;
        mgr.defined_names.push(DefinedName {
            name: "TaxRate".to_string(),
            formula: "0.08".to_string(),
            local_sheet_id: None,
            hidden: false,
            comment: None,
        });
        let style = CellStyle {
            bold: Some(true),
            italic: Some(true),
            ..Default::default()
        };
        mgr.sheets[0].set_cell_style(0, 0, style);
        mgr.evaluate().unwrap();

        let fresh_bytes = mgr.save_bytes().unwrap();
        assert_valid_ooxml(&fresh_bytes);

        let mut loaded = WorkbookManager::load_bytes(&fresh_bytes).unwrap();
        loaded.set_cell(0, 0, 0, "50".to_string());
        loaded.set_cell(0, 3, 0, "=TaxRate*A1".to_string());
        loaded.evaluate().unwrap();

        let patched_bytes = loaded.save_bytes().unwrap();
        assert_valid_ooxml(&patched_bytes);
    }
}
