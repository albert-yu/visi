use std::collections::HashSet;
use std::io::Read;

use crate::core::ovba;
use crate::core::vba_xlsx::decode_mbcs;

#[allow(dead_code)]
pub fn validate_vba_project_bin(bytes: &[u8]) -> Result<(), Vec<String>> {
    let mut errors = Vec::new();
    let mut cf = match cfb::CompoundFile::open(std::io::Cursor::new(bytes)) {
        Ok(cf) => cf,
        Err(e) => return Err(vec![format!("invalid CFB container: {e}")]),
    };

    for required in ["/PROJECT", "/PROJECTwm", "/VBA/dir", "/VBA/_VBA_PROJECT"] {
        if cf.open_stream(required).is_err() {
            errors.push(format!("missing required MS-OVBA stream {required}"));
        }
    }
    if !errors.is_empty() {
        return Err(errors);
    }

    let dir_raw = read_stream(&mut cf, "/VBA/dir").unwrap_or_default();
    let dir = match ovba::decompress(&dir_raw) {
        Ok(d) => d,
        Err(e) => {
            errors.push(format!("failed to decompress /VBA/dir: {e}"));
            return Err(errors);
        }
    };

    let mut pos = 0usize;
    let mut seen_syskind = false;
    let mut seen_lcid = false;
    let mut seen_lcid_invoke = false;
    let mut codepage: Option<u16> = None;
    let mut seen_name = false;
    let mut seen_version = false;
    let mut declared_module_count: Option<usize> = None;
    let mut in_modules = false;
    let mut seen_dir_terminator = false;
    let mut modules: Vec<(String, String, usize, u16)> = Vec::new();
    let mut cur_name: Option<String> = None;
    let mut cur_stream: Option<String> = None;
    let mut cur_offset: Option<usize> = None;
    let mut cur_type: Option<u16> = None;

    while pos + 6 <= dir.len() {
        let id = u16::from_le_bytes([dir[pos], dir[pos + 1]]);
        let (data, next) = if id == 0x0009 {
            if pos + 12 > dir.len() {
                errors.push("truncated PROJECTVERSION record".to_string());
                break;
            }
            (&dir[pos + 6..pos + 12], pos + 12)
        } else {
            let size = u32::from_le_bytes([dir[pos + 2], dir[pos + 3], dir[pos + 4], dir[pos + 5]])
                as usize;
            if pos + 6 + size > dir.len() {
                errors.push(format!(
                    "truncated dir record 0x{id:04X} claiming {size} bytes"
                ));
                break;
            }
            (&dir[pos + 6..pos + 6 + size], pos + 6 + size)
        };

        if !in_modules {
            match id {
                0x0001 => {
                    seen_syskind = true;
                    if data.len() != 4 {
                        errors.push("PROJECTSYSKIND must be 4 bytes".to_string());
                    }
                }
                0x0002 => seen_lcid = true,
                0x0014 => seen_lcid_invoke = true,
                0x0003 if data.len() == 2 => {
                    codepage = Some(u16::from_le_bytes([data[0], data[1]]));
                }
                0x0004 => seen_name = true,
                0x0009 => seen_version = true,
                0x000F if data.len() == 2 => {
                    declared_module_count = Some(u16::from_le_bytes([data[0], data[1]]) as usize);
                    in_modules = true;
                }
                _ => {}
            }
        } else {
            let cp = codepage.unwrap_or(1252);
            match id {
                0x0010 => {
                    seen_dir_terminator = true;
                    break;
                }
                0x0019 => cur_name = Some(decode_mbcs(data, cp)),
                0x001A => {
                    if cur_stream.is_none() {
                        cur_stream = Some(decode_mbcs(data, cp));
                    }
                }
                0x0031 if data.len() == 4 => {
                    cur_offset =
                        Some(u32::from_le_bytes([data[0], data[1], data[2], data[3]]) as usize);
                }
                0x0021 | 0x0022 => cur_type = Some(id),
                0x002B => {
                    let name = cur_name.take().unwrap_or_default();
                    let stream = cur_stream.take().unwrap_or_else(|| name.clone());
                    let offset = cur_offset.take().unwrap_or(0);
                    let mtype = cur_type.take().unwrap_or(0);
                    if name.is_empty() {
                        errors.push("module missing MODULENAME (0x0019)".to_string());
                    }
                    if mtype == 0 {
                        errors.push(format!("module '{name}' missing MODULETYPE"));
                    }
                    modules.push((name, stream, offset, mtype));
                }
                _ => {}
            }
        }
        pos = next;
    }

    if !seen_syskind {
        errors.push("missing PROJECTSYSKIND (0x0001)".to_string());
    }
    if !seen_lcid {
        errors.push("missing PROJECTLCID (0x0002)".to_string());
    }
    if !seen_lcid_invoke {
        errors.push("missing PROJECTLCIDINVOKE (0x0014)".to_string());
    }
    if codepage.is_none() {
        errors.push("missing PROJECTCODEPAGE (0x0003)".to_string());
    }
    if !seen_name {
        errors.push("missing PROJECTNAME (0x0004)".to_string());
    }
    if !seen_version {
        errors.push("missing PROJECTVERSION (0x0009)".to_string());
    }
    if !seen_dir_terminator {
        errors.push("missing dir stream terminator (0x0010)".to_string());
    }
    if let Some(expected) = declared_module_count
        && expected != modules.len()
    {
        errors.push(format!(
            "PROJECTMODULES count {expected} does not match parsed module count {}",
            modules.len()
        ));
    }

    let cp = codepage.unwrap_or(1252);
    let project_bytes = read_stream(&mut cf, "/PROJECT").unwrap_or_default();
    let project_text = decode_mbcs(&project_bytes, cp);
    if !project_text.lines().any(|l| l.starts_with("ID=")) {
        errors.push("/PROJECT missing ID= property".to_string());
    }
    if !project_text.lines().any(|l| l.starts_with("Name=")) {
        errors.push("/PROJECT missing Name= property".to_string());
    }
    let project_entries: HashSet<String> = project_text
        .lines()
        .filter_map(|l| {
            l.strip_prefix("Document=")
                .and_then(|r| r.split('/').next())
                .or_else(|| l.strip_prefix("Module="))
                .or_else(|| l.strip_prefix("Class="))
                .or_else(|| l.strip_prefix("BaseClass="))
        })
        .map(|s| s.to_ascii_lowercase())
        .collect();

    for (name, stream, offset, _) in &modules {
        if !project_entries.contains(&name.to_ascii_lowercase()) {
            errors.push(format!(
                "module '{name}' present in /VBA/dir but missing from /PROJECT"
            ));
        }
        match read_stream(&mut cf, &format!("/VBA/{stream}")) {
            Ok(raw) => {
                if *offset > raw.len() {
                    errors.push(format!(
                        "module '{name}' TextOffset {offset} exceeds stream length {}",
                        raw.len()
                    ));
                } else if let Err(e) = ovba::decompress(&raw[*offset..]) {
                    errors.push(format!("module '{name}' source failed to decompress: {e}"));
                }
            }
            Err(e) => errors.push(format!("missing module stream '/VBA/{stream}': {e}")),
        }
    }

    let wm = read_stream(&mut cf, "/PROJECTwm").unwrap_or_default();
    if wm.len() < 2 || wm[wm.len() - 2..] != [0x00, 0x00] {
        errors.push("/PROJECTwm must end with 0x0000 terminator".to_string());
    }

    if errors.is_empty() {
        Ok(())
    } else {
        Err(errors)
    }
}

#[allow(dead_code)]
fn read_stream<R: Read + std::io::Seek>(
    cf: &mut cfb::CompoundFile<R>,
    path: &str,
) -> Result<Vec<u8>, String> {
    let mut s = cf.open_stream(path).map_err(|e| e.to_string())?;
    let mut out = Vec::new();
    s.read_to_end(&mut out).map_err(|e| e.to_string())?;
    Ok(out)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::core::vba::{VbaModule, VbaModuleKind, VbaProject};
    use crate::core::vba_xlsx::{
        build_vba_project_bin, export_vba_project, parse_vba_project_from_cfb_bytes,
    };
    use std::collections::HashMap;
    use std::io::Write;

    #[test]
    fn ms_ovba_validator_accepts_class_standard_document_and_cp1252_roundtrip() {
        let mut project = VbaProject::new_empty();
        project.modules.push(VbaModule {
            name: "ThisWorkbook".to_string(),
            kind: VbaModuleKind::Document,
            source: "Attribute VB_Name = \"ThisWorkbook\"\r\n".to_string(),
            bound_sheet_id: None,
            prefix_bytes: project.seed_prefix_bytes.clone(),
            module_cookie: project.seed_module_cookie,
            cached_compressed_source: None,
        });
        project.modules.push(VbaModule {
            name: "Greeter".to_string(),
            kind: VbaModuleKind::Class,
            source: "Attribute VB_Name = \"Greeter\"\r\nPublic Function Euro() As String\r\nEuro = \"€ café ñ\"\r\nEnd Function\r\n".to_string(),
            bound_sheet_id: None,
            prefix_bytes: project.seed_prefix_bytes.clone(),
            module_cookie: project.seed_module_cookie,
            cached_compressed_source: None,
        });
        project.modules.push(VbaModule {
            name: "ModMain".to_string(),
            kind: VbaModuleKind::Standard,
            source: "Attribute VB_Name = \"ModMain\"\r\nSub Run()\r\nEnd Sub\r\n".to_string(),
            bound_sheet_id: None,
            prefix_bytes: project.seed_prefix_bytes.clone(),
            module_cookie: project.seed_module_cookie,
            cached_compressed_source: None,
        });

        let bin = build_vba_project_bin(&project).unwrap();
        validate_vba_project_bin(&bin).unwrap();

        let mut cf = cfb::CompoundFile::open(std::io::Cursor::new(&bin)).unwrap();
        let proj_text = decode_mbcs(&read_stream(&mut cf, "/PROJECT").unwrap(), 1252);
        assert!(proj_text.contains("Document=ThisWorkbook/&H00000000"));
        assert!(proj_text.contains("Class=Greeter"));
        assert!(proj_text.contains("Module=ModMain"));

        let parsed = parse_vba_project_from_cfb_bytes(bin, &HashMap::new()).unwrap();
        let greeter = parsed.find_module("Greeter").unwrap();
        assert_eq!(greeter.kind, VbaModuleKind::Class);
        assert!(greeter.source.contains("€ café ñ"));
    }

    #[test]
    fn export_vba_project_is_idempotent_and_strips_stale_signature() {
        let mut zip_buf = Vec::new();
        {
            let mut zw = zip::ZipWriter::new(std::io::Cursor::new(&mut zip_buf));
            let opt = zip::write::SimpleFileOptions::default();
            zw.start_file("[Content_Types].xml", opt).unwrap();
            zw.write_all(b"<Types><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/></Types>").unwrap();
            zw.start_file("xl/workbook.xml", opt).unwrap();
            zw.write_all(b"<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>").unwrap();
            zw.start_file("xl/_rels/workbook.xml.rels", opt).unwrap();
            zw.write_all(b"<Relationships><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>").unwrap();
            zw.start_file("xl/vbaProjectSignature.bin", opt).unwrap();
            zw.write_all(b"stale-sig").unwrap();
            zw.finish().unwrap();
        }

        let mut project = VbaProject::new_empty();
        project.modules.push(VbaModule {
            name: "Module1".to_string(),
            kind: VbaModuleKind::Standard,
            source: "Attribute VB_Name = \"Module1\"\r\nSub Hi()\r\nEnd Sub\r\n".to_string(),
            bound_sheet_id: None,
            prefix_bytes: project.seed_prefix_bytes.clone(),
            module_cookie: project.seed_module_cookie,
            cached_compressed_source: None,
        });

        let once = export_vba_project(zip_buf, Some(&project), &HashMap::new()).unwrap();
        let twice = export_vba_project(once, Some(&project), &HashMap::new()).unwrap();
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(twice)).unwrap();
        assert!(archive.by_name("xl/vbaProjectSignature.bin").is_err());

        let mut ct = String::new();
        archive
            .by_name("[Content_Types].xml")
            .unwrap()
            .read_to_string(&mut ct)
            .unwrap();
        assert_eq!(ct.matches("Extension=\"bin\"").count(), 1);

        let mut rels = String::new();
        archive
            .by_name("xl/_rels/workbook.xml.rels")
            .unwrap()
            .read_to_string(&mut rels)
            .unwrap();
        assert_eq!(rels.matches("vbaProject.bin").count(), 1);

        let mut wb = String::new();
        archive
            .by_name("xl/workbook.xml")
            .unwrap()
            .read_to_string(&mut wb)
            .unwrap();
        assert_eq!(wb.matches("codeName=\"ThisWorkbook\"").count(), 1);
    }

    #[test]
    fn validator_rejects_corrupt_cfb() {
        assert!(validate_vba_project_bin(b"not a cfb").is_err());
    }
}
