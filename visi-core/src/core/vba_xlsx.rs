use crate::core::ovba;
use crate::core::vba::{VbaModule, VbaModuleKind, VbaProject};
use crate::core::xlsx::{escape_xml, get_attr, get_zip_file_content, parse_workbook_rels};
use std::collections::{HashMap, HashSet};
use std::io::{Read, Write};

const REL_VBA_PROJECT: &str = "http://schemas.microsoft.com/office/2006/relationships/vbaProject";

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
    let protection_lines = parse_protection_lines(&project_text);

    let dir_raw = read_stream_bytes(&mut cfb_file, "/VBA/dir")?;
    let dir = ovba::decompress(&dir_raw)?;
    let codepage = parse_codepage(&dir).unwrap_or(1252);
    let module_specs = parse_module_specs(&dir, codepage)?;

    let mut modules = Vec::with_capacity(module_specs.len());
    for spec in module_specs {
        let stream_name = spec.stream_name.as_deref().unwrap_or(&spec.name);
        let raw = read_stream_bytes(&mut cfb_file, &format!("/VBA/{stream_name}"))?;
        let text_offset = spec.text_offset.min(raw.len());
        let prefix_bytes = raw[..text_offset].to_vec();
        let cached_compressed_source = raw[text_offset..].to_vec();
        let source_bytes = ovba::decompress(&cached_compressed_source)?;
        let source = decode_mbcs(&source_bytes, codepage);

        let kind = if spec.name == "ThisWorkbook" || document_module_names.contains(&spec.name) {
            VbaModuleKind::Document
        } else if spec.is_document_shaped {
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
            module_cookie: spec.module_cookie,
            cached_compressed_source: Some(cached_compressed_source),
        });
    }

    Ok(VbaProject {
        project_id,
        modules,
        raw_donor: vba_bin,
        seed_prefix_bytes: Vec::new(),
        seed_module_cookie: 0xFFFF,
        protection_lines,
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

fn parse_project_id(project_text: &str) -> Option<String> {
    for line in project_text.lines() {
        if let Some(rest) = line.strip_prefix("ID=") {
            return Some(rest.trim().trim_matches('"').to_string());
        }
    }
    None
}

fn parse_document_module_names(project_text: &str) -> HashSet<String> {
    project_text
        .lines()
        .filter_map(|line| line.strip_prefix("Document="))
        .filter_map(|rest| rest.split('/').next())
        .map(|s| s.to_string())
        .collect()
}

fn parse_protection_lines(project_text: &str) -> Option<String> {
    let lines: Vec<&str> = project_text
        .lines()
        .filter(|line| {
            line.starts_with("CMG=") || line.starts_with("DPB=") || line.starts_with("GC=")
        })
        .collect();
    if lines.is_empty() {
        None
    } else {
        Some(lines.join("\r\n"))
    }
}

fn parse_sheet_code_names(workbook_xml: &str) -> HashMap<String, String> {
    let mut map = HashMap::new();
    let mut reader = quick_xml::reader::Reader::from_str(workbook_xml);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Eof) => break,
            Ok(quick_xml::events::Event::Empty(e)) | Ok(quick_xml::events::Event::Start(e)) => {
                let local = e.name().local_name().into_inner();
                if local == b"sheet"
                    && let (Some(name), Some(code_name)) =
                        (get_attr(&e, b"name"), get_attr(&e, b"codeName"))
                {
                    map.insert(name, code_name);
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
    stream_name: Option<String>,
    doc_string: Vec<u8>,
    doc_string_unicode: Vec<u8>,
    help_context: u32,
    read_only: bool,
    private: bool,
    text_offset: usize,
    is_document_shaped: bool,
    module_cookie: u16,
}

const CP1252_HIGH: [char; 32] = [
    '\u{20AC}', '\u{0081}', '\u{201A}', '\u{0192}', '\u{201E}', '\u{2026}', '\u{2020}', '\u{2021}',
    '\u{02C6}', '\u{2030}', '\u{0160}', '\u{2039}', '\u{0152}', '\u{008D}', '\u{017D}', '\u{008F}',
    '\u{0090}', '\u{2018}', '\u{2019}', '\u{201C}', '\u{201D}', '\u{2022}', '\u{2013}', '\u{2014}',
    '\u{02DC}', '\u{2122}', '\u{0161}', '\u{203A}', '\u{0153}', '\u{009D}', '\u{017E}', '\u{0178}',
];

const MAC_ROMAN_HIGH: [char; 128] = [
    '\u{00C4}', '\u{00C5}', '\u{00C7}', '\u{00C9}', '\u{00D1}', '\u{00D6}', '\u{00DC}', '\u{00E1}',
    '\u{00E0}', '\u{00E2}', '\u{00E4}', '\u{00E3}', '\u{00E5}', '\u{00E7}', '\u{00E9}', '\u{00E8}',
    '\u{00EA}', '\u{00EB}', '\u{00ED}', '\u{00EC}', '\u{00EE}', '\u{00EF}', '\u{00F1}', '\u{00F3}',
    '\u{00F2}', '\u{00F4}', '\u{00F6}', '\u{00F5}', '\u{00FA}', '\u{00F9}', '\u{00FB}', '\u{00FC}',
    '\u{2020}', '\u{00B0}', '\u{00A2}', '\u{00A3}', '\u{00A7}', '\u{2022}', '\u{00B6}', '\u{00DF}',
    '\u{00AE}', '\u{00A9}', '\u{2122}', '\u{00B4}', '\u{00A8}', '\u{2260}', '\u{00C6}', '\u{00D8}',
    '\u{221E}', '\u{00B1}', '\u{2264}', '\u{2265}', '\u{00A5}', '\u{00B5}', '\u{2202}', '\u{2211}',
    '\u{220F}', '\u{03C0}', '\u{222B}', '\u{00AA}', '\u{00BA}', '\u{03A9}', '\u{00E6}', '\u{00F8}',
    '\u{00BF}', '\u{00A1}', '\u{00AC}', '\u{221A}', '\u{0192}', '\u{2248}', '\u{2206}', '\u{00AB}',
    '\u{00BB}', '\u{2026}', '\u{00A0}', '\u{00C0}', '\u{00C3}', '\u{00D5}', '\u{0152}', '\u{0153}',
    '\u{2013}', '\u{2014}', '\u{201C}', '\u{201D}', '\u{2018}', '\u{2019}', '\u{00F7}', '\u{25CA}',
    '\u{00FF}', '\u{0178}', '\u{2044}', '\u{20AC}', '\u{2039}', '\u{203A}', '\u{FB01}', '\u{FB02}',
    '\u{2021}', '\u{00B7}', '\u{201A}', '\u{201E}', '\u{2030}', '\u{00C2}', '\u{00CA}', '\u{00C1}',
    '\u{00CB}', '\u{00C8}', '\u{00CD}', '\u{00CE}', '\u{00CF}', '\u{00CC}', '\u{00D3}', '\u{00D4}',
    '\u{F8FF}', '\u{00D2}', '\u{00DA}', '\u{00DB}', '\u{00D9}', '\u{0131}', '\u{02C6}', '\u{02DC}',
    '\u{00AF}', '\u{02D8}', '\u{02D9}', '\u{02DA}', '\u{00B8}', '\u{02DD}', '\u{02DB}', '\u{02C7}',
];

pub(crate) fn decode_mbcs(bytes: &[u8], codepage: u16) -> String {
    match codepage {
        65001 => String::from_utf8_lossy(bytes).into_owned(),
        10000 => bytes
            .iter()
            .map(|&b| {
                if b < 0x80 {
                    b as char
                } else {
                    MAC_ROMAN_HIGH[(b - 0x80) as usize]
                }
            })
            .collect(),
        _ => bytes
            .iter()
            .map(|&b| {
                if (0x80..=0x9F).contains(&b) {
                    CP1252_HIGH[(b - 0x80) as usize]
                } else {
                    b as char
                }
            })
            .collect(),
    }
}

pub(crate) fn encode_mbcs(s: &str, codepage: u16) -> Vec<u8> {
    match codepage {
        65001 => s.as_bytes().to_vec(),
        10000 => s
            .chars()
            .map(|c| {
                if (c as u32) < 0x80 {
                    c as u8
                } else if let Some(pos) = MAC_ROMAN_HIGH.iter().position(|&m| m == c) {
                    0x80 + pos as u8
                } else {
                    b'?'
                }
            })
            .collect(),
        _ => s
            .chars()
            .map(|c| {
                let cp = c as u32;
                if cp < 0x80 || (0xA0..=0xFF).contains(&cp) {
                    cp as u8
                } else if let Some(pos) = CP1252_HIGH.iter().position(|&m| m == c) {
                    0x80 + pos as u8
                } else {
                    b'?'
                }
            })
            .collect(),
    }
}

fn parse_codepage(dir: &[u8]) -> Option<u16> {
    let mut pos = 0;
    while pos + 6 <= dir.len() {
        let Ok((id, data, next)) = read_dir_record(dir, pos) else {
            break;
        };
        if id == 0x0003 && data.len() >= 2 {
            return Some(u16::from_le_bytes([data[0], data[1]]));
        }
        if id == 0x000F {
            break;
        }
        pos = next;
    }
    None
}

fn decode_utf16le(data: &[u8]) -> String {
    let units: Vec<u16> = data
        .chunks_exact(2)
        .map(|c| u16::from_le_bytes([c[0], c[1]]))
        .collect();
    String::from_utf16_lossy(&units)
}

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
    let size =
        u32::from_le_bytes([dir[pos + 2], dir[pos + 3], dir[pos + 4], dir[pos + 5]]) as usize;
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

fn parse_module_specs(dir: &[u8], codepage: u16) -> Result<Vec<ModuleSpec>, String> {
    let modules_start = find_projectmodules_start(dir)?;
    let (id, data, mut pos) = read_dir_record(dir, modules_start)?;
    if id != 0x000F {
        return Err("expected PROJECTMODULES record".to_string());
    }
    if data.len() < 2 {
        return Err("PROJECTMODULES record too short".to_string());
    }
    let count = u16::from_le_bytes([data[0], data[1]]) as usize;

    let (id, _, next) = read_dir_record(dir, pos)?;
    if id != 0x0013 {
        return Err("expected PROJECTCOOKIE record after PROJECTMODULES".to_string());
    }
    pos = next;

    let mut specs = Vec::with_capacity(count);
    let mut cur_name: Option<String> = None;
    let mut cur_stream_name: Option<String> = None;
    let mut cur_doc: Vec<u8> = Vec::new();
    let mut cur_doc_unicode: Vec<u8> = Vec::new();
    let mut cur_help: u32 = 0;
    let mut cur_read_only = false;
    let mut cur_private = false;
    let mut cur_offset: Option<usize> = None;
    let mut cur_document_shaped: Option<bool> = None;
    let mut cur_cookie: Option<u16> = None;
    while pos + 6 <= dir.len() {
        let (id, data, next) = read_dir_record(dir, pos)?;
        match id {
            0x0010 => break,
            0x0019 => cur_name = Some(decode_mbcs(data, codepage)),
            0x0047 => cur_name = Some(decode_utf16le(data)),
            0x001A => {
                if cur_stream_name.is_none() {
                    cur_stream_name = Some(decode_mbcs(data, codepage));
                }
            }
            0x0032 => cur_stream_name = Some(decode_utf16le(data)),
            0x001C => cur_doc = data.to_vec(),
            0x0048 => cur_doc_unicode = data.to_vec(),
            0x001E => {
                if data.len() >= 4 {
                    cur_help = u32::from_le_bytes([data[0], data[1], data[2], data[3]]);
                }
            }
            0x0025 => cur_read_only = true,
            0x0028 => cur_private = true,
            0x0031 => {
                if data.len() >= 4 {
                    cur_offset =
                        Some(u32::from_le_bytes([data[0], data[1], data[2], data[3]]) as usize);
                }
            }
            0x0021 => cur_document_shaped = Some(false),
            0x0022 => cur_document_shaped = Some(true),
            0x002C => {
                if data.len() >= 2 {
                    cur_cookie = Some(u16::from_le_bytes([data[0], data[1]]));
                }
            }
            0x002B => {
                let name = cur_name
                    .take()
                    .ok_or("MODULETERMINATOR reached with no MODULENAME seen")?;
                specs.push(ModuleSpec {
                    name,
                    stream_name: cur_stream_name.take(),
                    doc_string: std::mem::take(&mut cur_doc),
                    doc_string_unicode: std::mem::take(&mut cur_doc_unicode),
                    help_context: std::mem::take(&mut cur_help),
                    read_only: std::mem::take(&mut cur_read_only),
                    private: std::mem::take(&mut cur_private),
                    text_offset: cur_offset.take().unwrap_or(0),
                    is_document_shaped: cur_document_shaped.take().unwrap_or(false),
                    module_cookie: cur_cookie.take().unwrap_or(0xFFFF),
                });
            }
            _ => {}
        }
        pos = next;
    }
    Ok(specs)
}

pub fn export_vba_project(
    xlsx_bytes: Vec<u8>,
    vba: Option<&VbaProject>,
    sheet_id_to_worksheet_name: &HashMap<u64, String>,
) -> Result<Vec<u8>, String> {
    let project = match vba {
        Some(p) if !p.modules.is_empty() => p,
        _ => return Ok(xlsx_bytes),
    };

    let new_vba_bin = build_vba_project_bin(project)?;

    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(&xlsx_bytes[..]))
        .map_err(|e| format!("Failed to open generated xlsx zip for VBA export: {}", e))?;
    let workbook_xml =
        get_zip_file_content(&mut archive, "xl/workbook.xml").ok_or("Missing xl/workbook.xml")?;
    let workbook_rels_xml = get_zip_file_content(&mut archive, "xl/_rels/workbook.xml.rels")
        .ok_or("Missing xl/_rels/workbook.xml.rels")?;
    let content_types = get_zip_file_content(&mut archive, "[Content_Types].xml")
        .ok_or("Missing [Content_Types].xml")?;
    drop(archive);

    let new_content_types = patch_content_types(&content_types);
    let new_workbook_rels_xml = patch_workbook_rels(&workbook_rels_xml);
    let new_workbook_xml =
        patch_workbook_code_names(&workbook_xml, project, sheet_id_to_worksheet_name);

    rewrite_zip_with_vba_part(
        &xlsx_bytes,
        new_content_types,
        new_workbook_xml,
        new_workbook_rels_xml,
        new_vba_bin,
    )
}

fn patch_content_types(content_types: &str) -> String {
    let mut xml = content_types.replacen(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
        "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
        1,
    );
    if !xml.contains("Extension=\"bin\"") {
        xml = xml.replacen(
            "</Types>",
            "<Default Extension=\"bin\" ContentType=\"application/vnd.ms-office.vbaProject\"/></Types>",
            1,
        );
    }
    xml
}

fn patch_workbook_rels(workbook_rels_xml: &str) -> String {
    if workbook_rels_xml.contains(REL_VBA_PROJECT) {
        return workbook_rels_xml.to_string();
    }
    let rid_to_target = parse_workbook_rels(workbook_rels_xml);
    let max_rid = rid_to_target
        .keys()
        .filter_map(|rid| {
            rid.strip_prefix("rId")
                .and_then(|n| n.parse::<usize>().ok())
        })
        .max()
        .unwrap_or(0)
        + 1;
    workbook_rels_xml.replacen(
        "</Relationships>",
        &format!(
            "<Relationship Id=\"rId{}\" Type=\"{}\" Target=\"vbaProject.bin\"/></Relationships>",
            max_rid, REL_VBA_PROJECT
        ),
        1,
    )
}

fn patch_workbook_code_names(
    workbook_xml: &str,
    project: &VbaProject,
    sheet_id_to_worksheet_name: &HashMap<u64, String>,
) -> String {
    let mut xml = if workbook_xml.contains("<workbookPr/>") {
        workbook_xml.replacen(
            "<workbookPr/>",
            "<workbookPr codeName=\"ThisWorkbook\"/>",
            1,
        )
    } else if let Some(pos) = workbook_xml.find("<workbookPr ") {
        let tag_end = workbook_xml[pos..]
            .find('>')
            .map(|o| pos + o)
            .unwrap_or(workbook_xml.len());
        if workbook_xml[pos..tag_end].contains("codeName=") {
            workbook_xml.to_string()
        } else {
            let insert_at = pos + "<workbookPr ".len();
            format!(
                "{}codeName=\"ThisWorkbook\" {}",
                &workbook_xml[..insert_at],
                &workbook_xml[insert_at..]
            )
        }
    } else if let Some(pos) = workbook_xml.find("<workbook ")
        && let Some(end_rel) = workbook_xml[pos..].find('>')
    {
        let insert_at = pos + end_rel + 1;
        format!(
            "{}<workbookPr codeName=\"ThisWorkbook\"/>{}",
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
        let Some(worksheet_name) = sheet_id_to_worksheet_name.get(&sheet_id) else {
            continue;
        };
        let needle = format!("name=\"{}\"", escape_xml(worksheet_name));
        if let Some(pos) = xml.find(&needle) {
            let tag_start = xml[..pos].rfind('<').unwrap_or(pos);
            let tag_end = xml[pos..].find('>').map(|o| pos + o).unwrap_or(xml.len());
            if !xml[tag_start..tag_end].contains("codeName=") {
                let insert_at = pos + needle.len();
                xml = format!(
                    "{} codeName=\"{}\"{}",
                    &xml[..insert_at],
                    escape_xml(&module.name),
                    &xml[insert_at..]
                );
            }
        }
    }
    xml
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

    for i in 0..archive.len() {
        let mut file = archive.by_index(i).map_err(|e| e.to_string())?;
        let name = file.name().to_string();
        if name == "xl/vbaProject.bin"
            || name == "xl/vbaProjectSignature.bin"
            || name == "xl/_rels/vbaProject.bin.rels"
        {
            continue;
        }
        let mut buf = Vec::new();
        file.read_to_end(&mut buf).map_err(|e| e.to_string())?;
        drop(file);

        writer
            .start_file(&name, options)
            .map_err(|e| e.to_string())?;
        if name == "[Content_Types].xml" {
            writer
                .write_all(content_types.as_bytes())
                .map_err(|e| e.to_string())?;
        } else if name == "xl/workbook.xml" {
            writer
                .write_all(workbook_xml.as_bytes())
                .map_err(|e| e.to_string())?;
        } else if name == "xl/_rels/workbook.xml.rels" {
            writer
                .write_all(workbook_rels_xml.as_bytes())
                .map_err(|e| e.to_string())?;
        } else {
            writer.write_all(&buf).map_err(|e| e.to_string())?;
        }
    }

    writer
        .start_file("xl/vbaProject.bin", options)
        .map_err(|e| e.to_string())?;
    writer.write_all(&vba_bin).map_err(|e| e.to_string())?;

    let cursor = writer.finish().map_err(|e| e.to_string())?;
    Ok(cursor.into_inner())
}

pub fn build_vba_project_bin(project: &VbaProject) -> Result<Vec<u8>, String> {
    let mut donor = cfb::CompoundFile::open(std::io::Cursor::new(project.raw_donor.clone()))
        .map_err(|e| format!("Failed to open donor vbaProject.bin: {}", e))?;

    let donor_dir_raw = read_stream_bytes(&mut donor, "/VBA/dir")?;
    let donor_dir = ovba::decompress(&donor_dir_raw)?;
    let codepage = parse_codepage(&donor_dir).unwrap_or(1252);
    let donor_specs: HashMap<String, ModuleSpec> = parse_module_specs(&donor_dir, codepage)
        .unwrap_or_default()
        .into_iter()
        .map(|s| (s.name.to_ascii_lowercase(), s))
        .collect();
    let donor_base_classes: Vec<String> = read_stream_string(&mut donor, "/PROJECT")
        .unwrap_or_default()
        .lines()
        .filter(|l| l.starts_with("BaseClass="))
        .map(|l| l.to_string())
        .collect();
    let modules_start = find_projectmodules_start(&donor_dir)?;
    let dir_prefix = &donor_dir[..modules_start];

    let vba_project_cache = read_stream_bytes(&mut donor, "/VBA/_VBA_PROJECT")?;

    let mut new_dir = Vec::new();
    new_dir.extend_from_slice(dir_prefix);
    write_record(
        &mut new_dir,
        0x000F,
        &(project.modules.len() as u16).to_le_bytes(),
    );
    write_record(&mut new_dir, 0x0013, &0xFFFFu16.to_le_bytes());
    for module in &project.modules {
        let donor_spec = donor_specs.get(&module.name.to_ascii_lowercase());
        let stream_name = donor_spec
            .and_then(|s| s.stream_name.as_deref())
            .unwrap_or(&module.name);
        let name_mbcs = encode_mbcs(&module.name, codepage);
        let stream_mbcs = encode_mbcs(stream_name, codepage);
        let doc_mbcs = donor_spec.map(|s| s.doc_string.as_slice()).unwrap_or(&[]);
        let doc_uni = donor_spec
            .map(|s| s.doc_string_unicode.as_slice())
            .unwrap_or(&[]);
        let help_ctx = donor_spec.map(|s| s.help_context).unwrap_or(0);

        write_record(&mut new_dir, 0x0019, &name_mbcs);
        write_record(&mut new_dir, 0x0047, &utf16le(&module.name));
        write_record(&mut new_dir, 0x001A, &stream_mbcs);
        write_record(&mut new_dir, 0x0032, &utf16le(stream_name));
        write_record(&mut new_dir, 0x001C, doc_mbcs);
        write_record(&mut new_dir, 0x0048, doc_uni);
        write_record(
            &mut new_dir,
            0x0031,
            &(module.prefix_bytes.len() as u32).to_le_bytes(),
        );
        write_record(&mut new_dir, 0x001E, &help_ctx.to_le_bytes());
        write_record(&mut new_dir, 0x002C, &module.module_cookie.to_le_bytes());
        let module_type_id = match module.kind {
            VbaModuleKind::Standard => 0x0021,
            VbaModuleKind::Document | VbaModuleKind::Class => 0x0022,
        };
        write_record(&mut new_dir, module_type_id, &[]);
        if donor_spec.is_some_and(|s| s.read_only) {
            write_record(&mut new_dir, 0x0025, &[]);
        }
        if donor_spec.is_some_and(|s| s.private) {
            write_record(&mut new_dir, 0x0028, &[]);
        }
        write_record(&mut new_dir, 0x002B, &[]);
    }
    write_record(&mut new_dir, 0x0010, &[]);
    let new_dir_compressed = ovba::compress(&new_dir)?;

    let new_project_text = build_project_stream(project, &donor_base_classes);
    let new_wm = build_projectwm_stream(project, codepage);

    let mut cf =
        cfb::CompoundFile::create_with_version(cfb::Version::V3, std::io::Cursor::new(Vec::new()))
            .map_err(|e| format!("Failed to create CFB container: {}", e))?;
    cf.create_storage("VBA")
        .map_err(|e| format!("Failed to create VBA storage: {}", e))?;
    cf.create_stream("VBA/dir")
        .and_then(|mut s| s.write_all(&new_dir_compressed))
        .map_err(|e| format!("Failed to write dir stream: {}", e))?;
    for module in &project.modules {
        let stream_name = donor_specs
            .get(&module.name.to_ascii_lowercase())
            .and_then(|s| s.stream_name.as_deref())
            .unwrap_or(&module.name);
        let compressed_source = match &module.cached_compressed_source {
            Some(cached) => cached.clone(),
            None => ovba::compress(&encode_mbcs(&module.source, codepage))?,
        };
        cf.create_stream(format!("VBA/{stream_name}"))
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
        .and_then(|mut s| s.write_all(&encode_mbcs(&new_project_text, codepage)))
        .map_err(|e| format!("Failed to write PROJECT stream: {}", e))?;
    cf.create_stream("PROJECTwm")
        .and_then(|mut s| s.write_all(&new_wm))
        .map_err(|e| format!("Failed to write PROJECTwm stream: {}", e))?;

    copy_extra_donor_entries(&mut donor, &mut cf)?;

    Ok(cf.into_inner().into_inner())
}

fn copy_extra_donor_entries<R: Read + std::io::Seek, W: Read + Write + std::io::Seek>(
    donor: &mut cfb::CompoundFile<R>,
    dst: &mut cfb::CompoundFile<W>,
) -> Result<(), String> {
    let entries: Vec<(std::path::PathBuf, bool)> = donor
        .walk()
        .filter_map(|e| {
            let p = e.path().to_path_buf();
            let s = p.to_string_lossy();
            if s == "/"
                || s == "/VBA"
                || s.starts_with("/VBA/")
                || s == "/PROJECT"
                || s == "/PROJECTwm"
            {
                return None;
            }
            Some((p, e.is_storage()))
        })
        .collect();
    for (path, is_storage) in &entries {
        if *is_storage {
            dst.create_storage_all(path)
                .map_err(|e| format!("Failed to copy donor storage '{}': {}", path.display(), e))?;
        }
    }
    for (path, is_storage) in &entries {
        if !*is_storage {
            let mut data = Vec::new();
            donor
                .open_stream(path)
                .and_then(|mut s| s.read_to_end(&mut data))
                .map_err(|e| format!("Failed to read donor stream '{}': {}", path.display(), e))?;
            if let Some(parent) = path.parent()
                && parent.as_os_str() != "/"
            {
                let _ = dst.create_storage_all(parent);
            }
            dst.create_stream(path)
                .and_then(|mut s| s.write_all(&data))
                .map_err(|e| format!("Failed to write donor stream '{}': {}", path.display(), e))?;
        }
    }
    Ok(())
}

pub(crate) fn write_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(data.len() as u32).to_le_bytes());
    out.extend_from_slice(data);
}

fn utf16le(s: &str) -> Vec<u8> {
    s.encode_utf16().flat_map(|u| u.to_le_bytes()).collect()
}

fn build_project_stream(project: &VbaProject, base_classes: &[String]) -> String {
    let mut s = String::new();
    s.push_str(&format!("ID=\"{}\"\r\n", project.project_id));
    for module in &project.modules {
        if module.kind == VbaModuleKind::Document {
            s.push_str(&format!("Document={}/&H00000000\r\n", module.name));
        }
    }
    for module in &project.modules {
        if module.kind == VbaModuleKind::Standard {
            s.push_str(&format!("Module={}\r\n", module.name));
        }
    }
    for module in &project.modules {
        if module.kind == VbaModuleKind::Class {
            s.push_str(&format!("Class={}\r\n", module.name));
        }
    }
    for line in base_classes {
        s.push_str(line);
        s.push_str("\r\n");
    }
    s.push_str("Name=\"VBAProject\"\r\n");
    s.push_str("HelpContextID=\"0\"\r\n");
    s.push_str("VersionCompatible32=\"393222000\"\r\n");
    if let Some(protection_lines) = &project.protection_lines {
        s.push_str(protection_lines);
        s.push_str("\r\n");
    }
    s.push_str("\r\n[Host Extender Info]\r\n");
    s.push_str("&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000\r\n");
    s.push_str("\r\n[Workspace]\r\n");
    for module in &project.modules {
        s.push_str(&format!("{}=0, 0, 0, 0, C\r\n", module.name));
    }
    s
}

fn build_projectwm_stream(project: &VbaProject, codepage: u16) -> Vec<u8> {
    let mut out = Vec::new();
    for module in &project.modules {
        out.extend_from_slice(&encode_mbcs(&module.name, codepage));
        out.push(0x00);
        out.extend_from_slice(&utf16le(&module.name));
        out.extend_from_slice(&[0x00, 0x00]);
    }
    out.extend_from_slice(&[0x00, 0x00]);
    out
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::core::vba::{VbaModuleKind, VbaProject};

    #[test]
    fn build_vba_project_bin_reuses_cached_compressed_source() {
        let mut project = VbaProject::new_empty();
        let cached =
            ovba::compress(b"Attribute VB_Name = \"M\"\r\nSub Cached()\r\nEnd Sub\r\n").unwrap();
        project.modules.push(VbaModule {
            name: "M".to_string(),
            kind: VbaModuleKind::Standard,
            source: "Attribute VB_Name = \"M\"\r\nSub Fresh()\r\nEnd Sub\r\n".to_string(),
            bound_sheet_id: None,
            prefix_bytes: project.seed_prefix_bytes.clone(),
            module_cookie: project.seed_module_cookie,
            cached_compressed_source: Some(cached),
        });

        let bin = build_vba_project_bin(&project).unwrap();
        let mut cfb_file = cfb::CompoundFile::open(std::io::Cursor::new(bin)).unwrap();
        let mut raw = Vec::new();
        cfb_file
            .open_stream("/VBA/M")
            .unwrap()
            .read_to_end(&mut raw)
            .unwrap();
        let text_offset = project.modules[0].prefix_bytes.len();
        let source_bytes = ovba::decompress(&raw[text_offset..]).unwrap();
        let source = String::from_utf8(source_bytes).unwrap();
        assert!(
            source.contains("Cached"),
            "expected the cached compressed bytes to be reused verbatim instead of \
             recompressing `source` fresh, got: {source}"
        );
    }

    #[test]
    fn build_vba_project_bin_compresses_fresh_when_no_cache() {
        let mut project = VbaProject::new_empty();
        project.modules.push(VbaModule {
            name: "M".to_string(),
            kind: VbaModuleKind::Standard,
            source: "Attribute VB_Name = \"M\"\r\nSub Fresh()\r\nEnd Sub\r\n".to_string(),
            bound_sheet_id: None,
            prefix_bytes: project.seed_prefix_bytes.clone(),
            module_cookie: project.seed_module_cookie,
            cached_compressed_source: None,
        });

        let bin = build_vba_project_bin(&project).unwrap();
        let mut cfb_file = cfb::CompoundFile::open(std::io::Cursor::new(bin)).unwrap();
        let mut raw = Vec::new();
        cfb_file
            .open_stream("/VBA/M")
            .unwrap()
            .read_to_end(&mut raw)
            .unwrap();
        let text_offset = project.modules[0].prefix_bytes.len();
        let source_bytes = ovba::decompress(&raw[text_offset..]).unwrap();
        let source = String::from_utf8(source_bytes).unwrap();
        assert!(source.contains("Fresh"));
    }
}
