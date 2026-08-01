//! `vbaProject.bin` (CFB/OLE binary) import and export.
//!
//! Mirrors `pivot_xlsx.rs`'s two-entry-point shape (`import_*`/`inject_*`
//! called from `xlsx.rs`'s `import_xlsx_data`/`export_xlsx_data`) and reuses
//! its zip/OOXML plumbing (`get_zip_file_content`, `parse_workbook_sheets`,
//! `get_attr`), but the payload itself is a binary CFB container, not XML.
//!
//! Design (validated against real Excel via a scratchpad proof-of-concept,
//! not just the MS-OVBA spec): on export, the `dir` stream's
//! PROJECTINFORMATION+PROJECTREFERENCES prefix and the `_VBA_PROJECT`
//! performance-cache stream are copied verbatim from the imported donor
//! file (`VbaProject::raw_donor`) -- never regenerated -- since a real
//! existing project may have references (MSForms, Office, etc.) this
//! codebase cannot yet synthesize. Everything else (the `dir` stream's
//! PROJECTMODULES section, `PROJECT`, `PROJECTwm`, and every module's
//! stream) is always rebuilt fresh from the current `VbaProject` state,
//! since those need to reflect add/remove/rename/edit operations. See the
//! module-level docs on `crate::core::vba` for the full picture.

use crate::core::engine::Sheet;
use crate::core::ovba;
use crate::core::vba::{VbaModule, VbaModuleKind, VbaProject};
use crate::core::xlsx::{get_attr, get_zip_file_content, parse_workbook_rels};
use std::collections::{HashMap, HashSet};
use std::io::{Read, Write};

const REL_VBA_PROJECT: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vbaProject";

// ---------------------------------------------------------------------------
// Import
// ---------------------------------------------------------------------------

/// Reads `xl/vbaProject.bin` out of an xlsx zip, if present, and reconstructs
/// a `VbaProject`. Returns `Ok(None)` if the workbook has no VBA project --
/// never an error, matching how a plain `.xlsx` is the overwhelmingly common
/// case. `sheet_id_by_name` maps each imported sheet's display name to its
/// engine id, used to resolve document modules' `codeName` binding back to
/// a sheet id.
pub fn import_vba_project(
    buffer: &[u8],
    sheet_id_by_name: &HashMap<String, u64>,
) -> Result<Option<VbaProject>, String> {
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(buffer))
        .map_err(|e| format!("Failed to open xlsx zip for VBA import: {}", e))?;

    let vba_bin = match get_zip_file_bytes(&mut archive, "xl/vbaProject.bin") {
        Some(b) => b,
        None => return Ok(None),
    };

    let workbook_xml = get_zip_file_content(&mut archive, "xl/workbook.xml").unwrap_or_default();
    let code_name_by_sheet_name = parse_sheet_code_names(&workbook_xml);
    let mut sheet_id_by_code_name: HashMap<String, u64> = HashMap::new();
    for (sheet_name, code_name) in &code_name_by_sheet_name {
        if let Some(&id) = sheet_id_by_name.get(sheet_name) {
            sheet_id_by_code_name.insert(code_name.clone(), id);
        }
    }

    parse_vba_project_from_cfb_bytes(vba_bin, &sheet_id_by_code_name).map(Some)
}

/// Parses a raw (not zip-wrapped) `vbaProject.bin` byte buffer into a
/// `VbaProject`. Shared by the xlsx-zip import path above and by seeding a
/// brand-new project from the bundled template asset (which is stored as
/// raw CFB bytes, not wrapped in a zip).
pub fn parse_vba_project_from_cfb_bytes(
    vba_bin: Vec<u8>,
    sheet_id_by_code_name: &HashMap<String, u64>,
) -> Result<VbaProject, String> {
    let mut cfb_file = cfb::CompoundFile::open(std::io::Cursor::new(vba_bin.clone()))
        .map_err(|e| format!("Failed to open vbaProject.bin as a CFB container: {}", e))?;

    let project_text = read_stream_string(&mut cfb_file, "/PROJECT")?;
    let project_id = parse_project_id(&project_text)
        .unwrap_or_else(|| "{00000000-0000-0000-0000-000000000000}".to_string());
    let document_module_names = parse_document_module_names(&project_text);

    let dir_raw = read_stream_bytes(&mut cfb_file, "/VBA/dir")?;
    let dir = ovba::decompress(&dir_raw);
    let module_specs = parse_module_specs(&dir)?;

    let mut modules = Vec::with_capacity(module_specs.len());
    for spec in module_specs {
        let raw = read_stream_bytes(&mut cfb_file, &format!("/VBA/{}", spec.name))?;
        let text_offset = spec.text_offset.min(raw.len());
        let prefix_bytes = raw[..text_offset].to_vec();
        let source_bytes = ovba::decompress(&raw[text_offset..]);
        let source = String::from_utf8_lossy(&source_bytes).into_owned();

        let kind = if spec.name == "ThisWorkbook" || document_module_names.contains(&spec.name) {
            VbaModuleKind::Document
        } else if spec.is_document_shaped {
            // MODULETYPE 0x0022 without a matching PROJECT-stream
            // "Document=" line: per spec this bit also covers class
            // modules. Class modules were never validated end-to-end
            // against real Excel by this feature's proof-of-concept --
            // best-effort import so an existing file isn't corrupted by a
            // no-op round-trip, not a fully-supported module kind.
            VbaModuleKind::Class
        } else {
            VbaModuleKind::Standard
        };
        let bound_sheet_id = if kind == VbaModuleKind::Document && spec.name != "ThisWorkbook" {
            sheet_id_by_code_name.get(&spec.name).copied()
        } else {
            None
        };

        modules.push(VbaModule {
            name: spec.name,
            kind,
            source,
            bound_sheet_id,
            prefix_bytes,
        });
    }

    Ok(VbaProject {
        project_id,
        modules,
        raw_donor: vba_bin,
        seed_prefix_bytes: Vec::new(),
    })
}

fn get_zip_file_bytes<R: Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
    path: &str,
) -> Option<Vec<u8>> {
    let mut file = archive.by_name(path).ok()?;
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).ok()?;
    Some(buf)
}

fn read_stream_bytes<F: Read + std::io::Seek>(
    cfb_file: &mut cfb::CompoundFile<F>,
    path: &str,
) -> Result<Vec<u8>, String> {
    let mut stream = cfb_file
        .open_stream(path)
        .map_err(|e| format!("Failed to open '{}' in vbaProject.bin: {}", path, e))?;
    let mut buf = Vec::new();
    stream
        .read_to_end(&mut buf)
        .map_err(|e| format!("Failed to read '{}' in vbaProject.bin: {}", path, e))?;
    Ok(buf)
}

fn read_stream_string<F: Read + std::io::Seek>(
    cfb_file: &mut cfb::CompoundFile<F>,
    path: &str,
) -> Result<String, String> {
    let bytes = read_stream_bytes(cfb_file, path)?;
    Ok(String::from_utf8_lossy(&bytes).into_owned())
}

/// The `PROJECT` stream's `ID="{...}"` line.
fn parse_project_id(project_text: &str) -> Option<String> {
    for line in project_text.lines() {
        if let Some(rest) = line.strip_prefix("ID=") {
            return Some(rest.trim().trim_matches('"').to_string());
        }
    }
    None
}

/// Every module name named in a `Document=Name/&H...` line.
fn parse_document_module_names(project_text: &str) -> HashSet<String> {
    project_text
        .lines()
        .filter_map(|line| line.strip_prefix("Document="))
        .filter_map(|rest| rest.split('/').next())
        .map(|s| s.to_string())
        .collect()
}

/// Maps each `<sheet>` element's `name` to its `codeName` attribute, for
/// sheets that have one.
fn parse_sheet_code_names(workbook_xml: &str) -> HashMap<String, String> {
    let mut map = HashMap::new();
    let mut reader = quick_xml::reader::Reader::from_str(workbook_xml);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Eof) => break,
            Ok(quick_xml::events::Event::Empty(e)) | Ok(quick_xml::events::Event::Start(e)) => {
                let local = e.name().local_name().into_inner();
                if local == b"sheet" {
                    if let (Some(name), Some(code_name)) =
                        (get_attr(&e, b"name"), get_attr(&e, b"codeName"))
                    {
                        map.insert(name, code_name);
                    }
                }
            }
            _ => {}
        }
        buf.clear();
    }
    map
}

struct ModuleSpec {
    name: String,
    text_offset: usize,
    /// True if MODULETYPE was 0x0022 (document/class per spec), false if
    /// 0x0021 (procedural/standard).
    is_document_shaped: bool,
}

/// Generic Id(u16)+Size(u32)+Data record reader, with the one documented
/// fixed-layout exception (PROJECTVERSION, id=0x0009, 12 bytes total).
/// Returns `(id, data, next_pos)`.
fn read_dir_record(dir: &[u8], pos: usize) -> Result<(u16, &[u8], usize), String> {
    if pos + 6 > dir.len() {
        return Err("dir stream truncated while reading a record header".to_string());
    }
    let id = u16::from_le_bytes([dir[pos], dir[pos + 1]]);
    if id == 0x0009 {
        if pos + 12 > dir.len() {
            return Err("dir stream truncated inside PROJECTVERSION".to_string());
        }
        return Ok((id, &dir[pos + 6..pos + 6], pos + 12));
    }
    let size = u32::from_le_bytes([dir[pos + 2], dir[pos + 3], dir[pos + 4], dir[pos + 5]]) as usize;
    let data_start = pos + 6;
    let data_end = data_start + size;
    if data_end > dir.len() {
        return Err(format!(
            "dir stream truncated: record 0x{:04X} at {} claims {} bytes but only {} remain",
            id,
            pos,
            size,
            dir.len() - data_start
        ));
    }
    Ok((id, &dir[data_start..data_end], data_end))
}

/// Finds the byte offset of the PROJECTMODULES record (0x000F), i.e. the
/// end of the PROJECTINFORMATION+PROJECTREFERENCES prefix.
fn find_projectmodules_start(dir: &[u8]) -> Result<usize, String> {
    let mut pos = 0;
    while pos + 6 <= dir.len() {
        let id = u16::from_le_bytes([dir[pos], dir[pos + 1]]);
        if id == 0x000F {
            return Ok(pos);
        }
        let (_, _, next) = read_dir_record(dir, pos)?;
        pos = next;
    }
    Err("PROJECTMODULES record (0x000F) not found in dir stream".to_string())
}

fn parse_module_specs(dir: &[u8]) -> Result<Vec<ModuleSpec>, String> {
    let modules_start = find_projectmodules_start(dir)?;
    let (id, data, mut pos) = read_dir_record(dir, modules_start)?;
    if id != 0x000F {
        return Err("expected PROJECTMODULES record".to_string());
    }
    if data.len() < 2 {
        return Err("PROJECTMODULES record too short".to_string());
    }
    let count = u16::from_le_bytes([data[0], data[1]]) as usize;

    // PROJECTCOOKIE (0x0013) follows unconditionally.
    let (id, _, next) = read_dir_record(dir, pos)?;
    if id != 0x0013 {
        return Err("expected PROJECTCOOKIE record after PROJECTMODULES".to_string());
    }
    pos = next;

    let mut specs = Vec::with_capacity(count);
    let mut cur_name: Option<String> = None;
    let mut cur_offset: Option<usize> = None;
    let mut cur_document_shaped: Option<bool> = None;
    while pos + 6 <= dir.len() {
        let (id, data, next) = read_dir_record(dir, pos)?;
        match id {
            0x0010 => break, // PROJECTTERMINATOR
            0x0019 => cur_name = Some(String::from_utf8_lossy(data).into_owned()),
            0x0031 => {
                if data.len() >= 4 {
                    cur_offset = Some(
                        u32::from_le_bytes([data[0], data[1], data[2], data[3]]) as usize,
                    );
                }
            }
            0x0021 => cur_document_shaped = Some(false),
            0x0022 => cur_document_shaped = Some(true),
            0x002B => {
                // MODULETERMINATOR: flush the accumulated module.
                let name = cur_name
                    .take()
                    .ok_or("MODULETERMINATOR reached with no MODULENAME seen")?;
                specs.push(ModuleSpec {
                    name,
                    text_offset: cur_offset.take().unwrap_or(0),
                    is_document_shaped: cur_document_shaped.take().unwrap_or(false),
                });
            }
            _ => {}
        }
        pos = next;
    }
    Ok(specs)
}

// ---------------------------------------------------------------------------
// Export
// ---------------------------------------------------------------------------

/// Splices a rebuilt `vbaProject.bin` into an already-produced xlsx zip
/// (patching `[Content_Types].xml`/`workbook.xml.rels`/`workbook.xml` as
/// needed), or returns `xlsx_bytes` unchanged if `vba` is `None`.
pub fn export_vba_project(
    xlsx_bytes: Vec<u8>,
    sheets: &[Sheet],
    vba: Option<&VbaProject>,
) -> Result<Vec<u8>, String> {
    let project = match vba {
        Some(p) if !p.modules.is_empty() => p,
        _ => return Ok(xlsx_bytes),
    };

    let new_vba_bin = build_vba_project_bin(project)?;

    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(&xlsx_bytes[..]))
        .map_err(|e| format!("Failed to open generated xlsx zip for VBA export: {}", e))?;
    let already_had_vba = archive.by_name("xl/vbaProject.bin").is_ok();

    let workbook_xml =
        get_zip_file_content(&mut archive, "xl/workbook.xml").ok_or("Missing xl/workbook.xml")?;
    let workbook_rels_xml = get_zip_file_content(&mut archive, "xl/_rels/workbook.xml.rels")
        .ok_or("Missing xl/_rels/workbook.xml.rels")?;
    let content_types = get_zip_file_content(&mut archive, "[Content_Types].xml")
        .ok_or("Missing [Content_Types].xml")?;
    drop(archive);

    let mut new_content_types = content_types.replacen(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
        "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
        1,
    );
    if !already_had_vba {
        new_content_types = new_content_types.replacen(
            "</Types>",
            "<Default Extension=\"bin\" ContentType=\"application/vnd.ms-office.vbaProject\"/></Types>",
            1,
        );
    }

    let mut new_workbook_rels_xml = workbook_rels_xml;
    if !already_had_vba {
        let rid_to_target = parse_workbook_rels(&new_workbook_rels_xml);
        let mut max_rid = rid_to_target
            .keys()
            .filter_map(|rid| rid.strip_prefix("rId").and_then(|n| n.parse::<usize>().ok()))
            .max()
            .unwrap_or(0);
        max_rid += 1;
        new_workbook_rels_xml = new_workbook_rels_xml.replacen(
            "</Relationships>",
            &format!(
                "<Relationship Id=\"rId{}\" Type=\"{}\" Target=\"vbaProject.bin\"/></Relationships>",
                max_rid, REL_VBA_PROJECT
            ),
            1,
        );
    }

    let new_workbook_xml = if already_had_vba {
        // A real macro-enabled donor's own workbook.xml already carries the
        // codeName wiring (it came from a real Excel-authored file); don't
        // touch it.
        workbook_xml
    } else {
        patch_workbook_code_names(&workbook_xml, project, sheets)
    };

    rewrite_zip_with_vba_part(
        &xlsx_bytes,
        new_content_types,
        new_workbook_xml,
        new_workbook_rels_xml,
        new_vba_bin,
    )
}

/// Adds `codeName="ThisWorkbook"` to `<workbookPr>` and a matching
/// `codeName="..."` to each `<sheet>` element bound to a Document module.
fn patch_workbook_code_names(workbook_xml: &str, project: &VbaProject, sheets: &[Sheet]) -> String {
    let mut xml = if workbook_xml.contains("<workbookPr/>") {
        workbook_xml.replacen("<workbookPr/>", "<workbookPr codeName=\"ThisWorkbook\"/>", 1)
    } else if let Some(pos) = workbook_xml.find("<workbookPr ") {
        let insert_at = pos + "<workbookPr ".len();
        format!(
            "{}codeName=\"ThisWorkbook\" {}",
            &workbook_xml[..insert_at],
            &workbook_xml[insert_at..]
        )
    } else {
        workbook_xml.to_string()
    };

    for module in &project.modules {
        if module.kind != VbaModuleKind::Document || module.name == "ThisWorkbook" {
            continue;
        }
        let Some(sheet_id) = module.bound_sheet_id else {
            continue;
        };
        let Some(sheet) = sheets.iter().find(|s| s.id == sheet_id) else {
            continue;
        };
        let needle = format!("name=\"{}\"", escape_xml(&sheet.name));
        if let Some(pos) = xml.find(&needle) {
            let insert_at = pos + needle.len();
            xml = format!(
                "{} codeName=\"{}\"{}",
                &xml[..insert_at],
                escape_xml(&module.name),
                &xml[insert_at..]
            );
        }
    }
    xml
}

fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

fn rewrite_zip_with_vba_part(
    original: &[u8],
    content_types: String,
    workbook_xml: String,
    workbook_rels_xml: String,
    vba_bin: Vec<u8>,
) -> Result<Vec<u8>, String> {
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(original))
        .map_err(|e| format!("Failed to re-open generated xlsx zip: {}", e))?;
    let mut writer = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
    let options = zip::write::SimpleFileOptions::default();
    let mut wrote_vba = false;

    for i in 0..archive.len() {
        let mut file = archive.by_index(i).map_err(|e| e.to_string())?;
        let name = file.name().to_string();
        let mut buf = Vec::new();
        file.read_to_end(&mut buf).map_err(|e| e.to_string())?;
        drop(file);

        writer.start_file(&name, options).map_err(|e| e.to_string())?;
        if name == "[Content_Types].xml" {
            writer.write_all(content_types.as_bytes()).map_err(|e| e.to_string())?;
        } else if name == "xl/workbook.xml" {
            writer.write_all(workbook_xml.as_bytes()).map_err(|e| e.to_string())?;
        } else if name == "xl/_rels/workbook.xml.rels" {
            writer.write_all(workbook_rels_xml.as_bytes()).map_err(|e| e.to_string())?;
        } else if name == "xl/vbaProject.bin" {
            writer.write_all(&vba_bin).map_err(|e| e.to_string())?;
            wrote_vba = true;
        } else {
            writer.write_all(&buf).map_err(|e| e.to_string())?;
        }
    }

    if !wrote_vba {
        writer
            .start_file("xl/vbaProject.bin", options)
            .map_err(|e| e.to_string())?;
        writer.write_all(&vba_bin).map_err(|e| e.to_string())?;
    }

    let cursor = writer.finish().map_err(|e| e.to_string())?;
    Ok(cursor.into_inner())
}

/// Builds a complete `vbaProject.bin` byte buffer: the donor's dir-stream
/// PROJECTINFORMATION+PROJECTREFERENCES prefix and `_VBA_PROJECT` cache are
/// copied verbatim; the dir stream's PROJECTMODULES section, `PROJECT`,
/// `PROJECTwm`, and every module stream are rebuilt fresh from `project`'s
/// current state.
pub fn build_vba_project_bin(project: &VbaProject) -> Result<Vec<u8>, String> {
    let mut donor = cfb::CompoundFile::open(std::io::Cursor::new(project.raw_donor.clone()))
        .map_err(|e| format!("Failed to open donor vbaProject.bin: {}", e))?;

    let donor_dir_raw = read_stream_bytes(&mut donor, "/VBA/dir")?;
    let donor_dir = ovba::decompress(&donor_dir_raw);
    let modules_start = find_projectmodules_start(&donor_dir)?;
    let dir_prefix = &donor_dir[..modules_start];

    let vba_project_cache = read_stream_bytes(&mut donor, "/VBA/_VBA_PROJECT")?;

    let mut new_dir = Vec::new();
    new_dir.extend_from_slice(dir_prefix);
    write_record(&mut new_dir, 0x000F, &(project.modules.len() as u16).to_le_bytes());
    write_record(&mut new_dir, 0x0013, &0xFFFFu16.to_le_bytes()); // PROJECTCOOKIE (value doesn't matter -- verified against real Excel)
    for module in &project.modules {
        write_record(&mut new_dir, 0x0019, module.name.as_bytes());
        write_record(&mut new_dir, 0x0047, &utf16le(&module.name));
        write_record(&mut new_dir, 0x001A, module.name.as_bytes());
        write_record(&mut new_dir, 0x0032, &utf16le(&module.name));
        write_record(&mut new_dir, 0x001C, &[]);
        write_record(&mut new_dir, 0x0048, &[]);
        write_record(
            &mut new_dir,
            0x0031,
            &(module.prefix_bytes.len() as u32).to_le_bytes(),
        );
        write_record(&mut new_dir, 0x001E, &0u32.to_le_bytes());
        write_record(&mut new_dir, 0x002C, &0xFFFFu16.to_le_bytes()); // MODULECOOKIE
        let module_type_id = match module.kind {
            VbaModuleKind::Standard => 0x0021,
            VbaModuleKind::Document | VbaModuleKind::Class => 0x0022,
        };
        write_record(&mut new_dir, module_type_id, &[]);
        write_record(&mut new_dir, 0x002B, &[]);
    }
    write_record(&mut new_dir, 0x0010, &[]); // PROJECTTERMINATOR
    let new_dir_compressed = ovba::compress(&new_dir);

    let new_project_text = build_project_stream(project);
    let new_wm = build_projectwm_stream(project);

    let mut cf = cfb::CompoundFile::create_with_version(
        cfb::Version::V3,
        std::io::Cursor::new(Vec::new()),
    )
    .map_err(|e| format!("Failed to create CFB container: {}", e))?;
    cf.create_storage("VBA")
        .map_err(|e| format!("Failed to create VBA storage: {}", e))?;
    cf.create_stream("VBA/dir")
        .and_then(|mut s| s.write_all(&new_dir_compressed))
        .map_err(|e| format!("Failed to write dir stream: {}", e))?;
    for module in &project.modules {
        let compressed_source = ovba::compress(module.source.as_bytes());
        cf.create_stream(format!("VBA/{}", module.name))
            .and_then(|mut s| {
                s.write_all(&module.prefix_bytes)?;
                s.write_all(&compressed_source)
            })
            .map_err(|e| format!("Failed to write module '{}' stream: {}", module.name, e))?;
    }
    cf.create_stream("VBA/_VBA_PROJECT")
        .and_then(|mut s| s.write_all(&vba_project_cache))
        .map_err(|e| format!("Failed to write _VBA_PROJECT stream: {}", e))?;
    cf.create_stream("PROJECT")
        .and_then(|mut s| s.write_all(new_project_text.as_bytes()))
        .map_err(|e| format!("Failed to write PROJECT stream: {}", e))?;
    cf.create_stream("PROJECTwm")
        .and_then(|mut s| s.write_all(&new_wm))
        .map_err(|e| format!("Failed to write PROJECTwm stream: {}", e))?;

    Ok(cf.into_inner().into_inner())
}

fn write_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(data.len() as u32).to_le_bytes());
    out.extend_from_slice(data);
}

fn utf16le(s: &str) -> Vec<u8> {
    s.encode_utf16().flat_map(|u| u.to_le_bytes()).collect()
}

/// `CMG`/`DPB`/`GC` protection-state lines are deliberately omitted:
/// verified against real Excel that pairing a donor's real values with a
/// different (or regenerated) project ID makes the whole project report as
/// "unviewable". Since this codebase doesn't support setting a VBA project
/// password, an unprotected project (no such lines) is always correct here.
fn build_project_stream(project: &VbaProject) -> String {
    let mut s = String::new();
    s.push_str(&format!("ID=\"{}\"\r\n", project.project_id));
    for module in &project.modules {
        if module.kind == VbaModuleKind::Document {
            s.push_str(&format!("Document={}/&H00000000\r\n", module.name));
        }
    }
    for module in &project.modules {
        if module.kind != VbaModuleKind::Document {
            s.push_str(&format!("Module={}\r\n", module.name));
        }
    }
    s.push_str("Name=\"VBAProject\"\r\n");
    s.push_str("HelpContextID=\"0\"\r\n");
    s.push_str("VersionCompatible32=\"393222000\"\r\n");
    s.push_str("\r\n[Host Extender Info]\r\n");
    s.push_str("&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000\r\n");
    s.push_str("\r\n[Workspace]\r\n");
    for module in &project.modules {
        s.push_str(&format!("{}=0, 0, 0, 0, C\r\n", module.name));
    }
    s
}

fn build_projectwm_stream(project: &VbaProject) -> Vec<u8> {
    let mut out = Vec::new();
    for module in &project.modules {
        out.extend_from_slice(module.name.as_bytes());
        out.push(0x00);
        out.extend_from_slice(&utf16le(&module.name));
        out.extend_from_slice(&[0x00, 0x00]);
    }
    out.extend_from_slice(&[0x00, 0x00]);
    out
}
