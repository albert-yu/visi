use std::io::{Cursor, Read, Write};

use zip::write::SimpleFileOptions;

use crate::core::Sheet;

use super::shared_strings::SharedStrings;

#[derive(Default)]
pub(crate) struct WorksheetLayout {
    pub(crate) column_widths: Vec<(usize, usize, f64)>,
    pub(crate) row_heights: Vec<(usize, f64)>,
}

const ROOT_RELS: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.bin\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/></Relationships>";
const CORE_XML: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\"><dc:creator>Visi</dc:creator></cp:coreProperties>";
const APP_XML: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"><Application>Visi</Application></Properties>";

fn content_types(sheet_count: usize) -> String {
    let mut xml = String::from(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.bin\" ContentType=\"application/vnd.ms-excel.sheet.binary.macroEnabled.main\"/>",
    );
    for index in 1..=sheet_count {
        xml.push_str(&format!("<Override PartName=\"/xl/worksheets/sheet{index}.bin\" ContentType=\"application/vnd.ms-excel.worksheet\"/>"));
    }
    xml.push_str("<Override PartName=\"/xl/styles.bin\" ContentType=\"application/vnd.ms-excel.styles\"/><Override PartName=\"/xl/sharedStrings.bin\" ContentType=\"application/vnd.ms-excel.sharedStrings\"/><Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/><Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/></Types>");
    xml
}

fn workbook_relationships(sheet_count: usize) -> String {
    let mut xml = String::from(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">",
    );
    for index in 1..=sheet_count {
        xml.push_str(&format!("<Relationship Id=\"rId{index}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{index}.bin\"/>"));
    }
    xml.push_str(&format!("<Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.bin\"/><Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.bin\"/></Relationships>", sheet_count + 1, sheet_count + 2));
    xml
}

fn add_file(
    writer: &mut zip::ZipWriter<Cursor<Vec<u8>>>,
    name: &str,
    bytes: &[u8],
) -> Result<(), String> {
    let options = SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated)
        .compression_level(Some(6));
    writer
        .start_file(name, options)
        .map_err(|error| error.to_string())?;
    writer.write_all(bytes).map_err(|error| error.to_string())
}

pub(crate) fn write(sheets: &[Sheet]) -> Result<Vec<u8>, String> {
    if sheets.is_empty() {
        return Err("cannot write an XLSB workbook without worksheets".to_string());
    }
    let mut strings = SharedStrings::default();
    let mut worksheet_parts = Vec::with_capacity(sheets.len());
    for sheet in sheets {
        worksheet_parts.push(super::worksheet::build(sheet, &mut strings)?);
    }
    let names = sheets
        .iter()
        .map(|sheet| sheet.name.clone())
        .collect::<Vec<_>>();
    let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
    add_file(
        &mut writer,
        "[Content_Types].xml",
        content_types(sheets.len()).as_bytes(),
    )?;
    add_file(&mut writer, "_rels/.rels", ROOT_RELS.as_bytes())?;
    add_file(&mut writer, "docProps/core.xml", CORE_XML.as_bytes())?;
    add_file(&mut writer, "docProps/app.xml", APP_XML.as_bytes())?;
    add_file(
        &mut writer,
        "xl/workbook.bin",
        &super::workbook::build(&names),
    )?;
    add_file(
        &mut writer,
        "xl/_rels/workbook.bin.rels",
        workbook_relationships(sheets.len()).as_bytes(),
    )?;
    add_file(&mut writer, "xl/styles.bin", &super::styles::build())?;
    add_file(&mut writer, "xl/sharedStrings.bin", &strings.encode())?;
    for (index, part) in worksheet_parts.iter().enumerate() {
        add_file(
            &mut writer,
            &format!("xl/worksheets/sheet{}.bin", index + 1),
            part,
        )?;
    }
    writer
        .finish()
        .map_err(|error| error.to_string())
        .map(Cursor::into_inner)
}

fn read_varint(data: &[u8], offset: &mut usize) -> Option<u32> {
    let mut value = 0u32;
    let mut shift = 0;
    loop {
        let byte = *data.get(*offset)?;
        *offset += 1;
        value |= u32::from(byte & 0x7f) << shift;
        if byte & 0x80 == 0 {
            return Some(value);
        }
        shift += 7;
        if shift > 28 {
            return None;
        }
    }
}

fn normalize_workbook_records(data: &[u8]) -> Result<(Vec<u8>, bool), String> {
    let mut offset = 0usize;
    let mut out = Vec::with_capacity(data.len());
    let mut changed = false;
    while offset < data.len() {
        let id = read_varint(data, &mut offset)
            .ok_or_else(|| "truncated XLSB workbook record id".to_string())?;
        let len = read_varint(data, &mut offset)
            .ok_or_else(|| "truncated XLSB workbook record length".to_string())?
            as usize;
        let end = offset
            .checked_add(len)
            .filter(|end| *end <= data.len())
            .ok_or_else(|| "XLSB workbook record extends past the part boundary".to_string())?;
        let payload = &data[offset..end];
        offset = end;
        if id == 0x009c && payload.len() >= 20 {
            let calamine_len = u32::from_le_bytes(payload[8..12].try_into().unwrap()) as usize;
            let standard_len = u32::from_le_bytes(payload[12..16].try_into().unwrap()) as usize;
            let calamine_end = 12usize.saturating_add(calamine_len.saturating_mul(2));
            let standard_end = 16usize.saturating_add(standard_len.saturating_mul(2));
            let calamine_layout = calamine_len <= 255
                && calamine_end <= payload.len()
                && String::from_utf16(
                    &payload[12..calamine_end]
                        .chunks_exact(2)
                        .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
                        .collect::<Vec<_>>(),
                )
                .is_ok_and(|value| value.starts_with("rId"));
            let standard_layout = standard_len <= 255
                && standard_end <= payload.len()
                && String::from_utf16(
                    &payload[16..standard_end]
                        .chunks_exact(2)
                        .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
                        .collect::<Vec<_>>(),
                )
                .is_ok_and(|value| value.starts_with("rId"));
            if !calamine_layout && standard_layout {
                let mut normalized = Vec::with_capacity(payload.len() - 4);
                normalized.extend_from_slice(&payload[..4]);
                normalized.extend_from_slice(&payload[8..]);
                super::record::write_record(id, &normalized, &mut out);
                changed = true;
                continue;
            }
        }
        super::record::write_record(id, payload, &mut out);
    }
    Ok((out, changed))
}

pub(crate) fn normalize_for_calamine(buffer: &[u8]) -> Result<Vec<u8>, String> {
    let mut archive = zip::ZipArchive::new(Cursor::new(buffer))
        .map_err(|error| format!("failed to open XLSB package: {error}"))?;
    let mut entries = Vec::with_capacity(archive.len());
    let mut changed = false;
    for index in 0..archive.len() {
        let mut file = archive.by_index(index).map_err(|error| error.to_string())?;
        let name = file.name().to_string();
        let is_dir = file.is_dir();
        let mut data = Vec::new();
        file.read_to_end(&mut data)
            .map_err(|error| error.to_string())?;
        if name.eq_ignore_ascii_case("xl/workbook.bin") {
            let (normalized, workbook_changed) = normalize_workbook_records(&data)?;
            data = normalized;
            changed |= workbook_changed;
        }
        entries.push((name, is_dir, data));
    }
    if !changed {
        return Ok(buffer.to_vec());
    }
    let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated)
        .compression_level(Some(6));
    for (name, is_dir, data) in entries {
        if is_dir {
            writer
                .add_directory(name, options)
                .map_err(|error| error.to_string())?;
        } else {
            writer
                .start_file(name, options)
                .map_err(|error| error.to_string())?;
            writer.write_all(&data).map_err(|error| error.to_string())?;
        }
    }
    writer
        .finish()
        .map_err(|error| error.to_string())
        .map(Cursor::into_inner)
}

fn formula_tokens_supported(mut tokens: &[u8]) -> Result<(), String> {
    while let Some((&token, rest)) = tokens.split_first() {
        tokens = rest;
        let skip = match token {
            0x01 => return Err("shared formula tokens are not supported in XLSB input".to_string()),
            0x03..=0x16 => 0,
            0x17 => {
                let length = tokens
                    .get(..2)
                    .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]) as usize)
                    .ok_or_else(|| "truncated XLSB string formula token".to_string())?;
                2 + length * 2
            }
            0x18 => match tokens.first() {
                Some(0x19) => 13,
                Some(0x1d) => 5,
                _ => return Err("unsupported XLSB extended formula token".to_string()),
            },
            0x19 => match tokens.first() {
                Some(0x01 | 0x02 | 0x08 | 0x10 | 0x20 | 0x21 | 0x40 | 0x41 | 0x80) => 3,
                Some(0x04) => 11,
                _ => return Err("unsupported XLSB formula attribute".to_string()),
            },
            0x1c | 0x1d => 1,
            0x1e => 2,
            0x1f => 8,
            0x20 | 0x40 | 0x60 => {
                return Err("array constants are not supported in XLSB input".to_string());
            }
            0x21 | 0x41 | 0x61 => 2,
            0x22 | 0x42 | 0x62 => 3,
            0x23 | 0x43 | 0x63 => 4,
            0x24 | 0x44 | 0x64 | 0x2a | 0x4a | 0x6a => 6,
            0x25 | 0x45 | 0x65 | 0x2b | 0x4b | 0x6b => 12,
            0x29 | 0x49 | 0x69 => {
                let length = tokens
                    .get(..2)
                    .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]) as usize)
                    .ok_or_else(|| "truncated XLSB memory formula token".to_string())?;
                let nested = tokens
                    .get(2..2 + length)
                    .ok_or_else(|| "truncated XLSB memory formula payload".to_string())?;
                formula_tokens_supported(nested)?;
                2 + length
            }
            0x39 | 0x59 | 0x79 => {
                return Err(
                    "external workbook formula names are not supported in XLSB input".to_string(),
                );
            }
            0x3a | 0x5a | 0x7a | 0x3c | 0x5c | 0x7c => 8,
            0x3b | 0x5b | 0x7b | 0x3d | 0x5d | 0x7d => 14,
            _ => return Err(format!("unsupported XLSB formula token 0x{token:02x}")),
        };
        tokens = tokens
            .get(skip..)
            .ok_or_else(|| "truncated XLSB formula token payload".to_string())?;
    }
    Ok(())
}

fn validate_formula_record(id: u32, payload: &[u8]) -> Result<(), String> {
    let formula_offset = match id {
        0x0008 => {
            let length = payload
                .get(8..12)
                .map(|bytes| u32::from_le_bytes(bytes.try_into().unwrap()) as usize)
                .ok_or_else(|| "truncated XLSB string formula".to_string())?;
            14 + length * 2
        }
        0x0009 => 18,
        0x000a | 0x000b => 11,
        _ => return Ok(()),
    };
    let formula = payload
        .get(formula_offset..)
        .ok_or_else(|| "truncated XLSB formula record".to_string())?;
    let length = formula
        .get(..4)
        .map(|bytes| u32::from_le_bytes(bytes.try_into().unwrap()) as usize)
        .ok_or_else(|| "truncated XLSB formula length".to_string())?;
    let tokens = formula
        .get(4..4 + length)
        .ok_or_else(|| "truncated XLSB formula payload".to_string())?;
    formula_tokens_supported(tokens)
}

fn read_wide_string(data: &[u8], offset: &mut usize) -> Option<String> {
    let length = u32::from_le_bytes(data.get(*offset..*offset + 4)?.try_into().ok()?) as usize;
    *offset += 4;
    let end = offset.checked_add(length.checked_mul(2)?)?;
    let encoded = data
        .get(*offset..end)?
        .chunks_exact(2)
        .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
        .collect::<Vec<_>>();
    *offset = end;
    String::from_utf16(&encoded).ok()
}

fn workbook_sheets(data: &[u8]) -> Result<Vec<(String, String)>, String> {
    let mut offset = 0usize;
    let mut sheets = Vec::new();
    while offset < data.len() {
        let id = read_varint(data, &mut offset)
            .ok_or_else(|| "truncated XLSB workbook record id".to_string())?;
        let length = read_varint(data, &mut offset)
            .ok_or_else(|| "truncated XLSB workbook record length".to_string())?
            as usize;
        let end = offset
            .checked_add(length)
            .filter(|end| *end <= data.len())
            .ok_or_else(|| "XLSB workbook record extends past the part boundary".to_string())?;
        let payload = &data[offset..end];
        offset = end;
        if id != 0x009c || payload.len() < 16 {
            continue;
        }
        let current_length = u32::from_le_bytes(payload[8..12].try_into().unwrap()) as usize;
        let current_end = 12usize.saturating_add(current_length.saturating_mul(2));
        let mut string_offset = if current_length <= 255 && current_end <= payload.len() {
            8
        } else {
            12
        };
        let relationship = read_wide_string(payload, &mut string_offset)
            .ok_or_else(|| "invalid XLSB worksheet relationship id".to_string())?;
        let name = read_wide_string(payload, &mut string_offset)
            .ok_or_else(|| "invalid XLSB worksheet name".to_string())?;
        sheets.push((name, relationship));
    }
    Ok(sheets)
}

fn parse_workbook_relationships(
    xml: &str,
) -> Result<std::collections::HashMap<String, String>, String> {
    let mut relationships = std::collections::HashMap::new();
    let mut reader = quick_xml::reader::Reader::from_str(xml);
    loop {
        match reader.read_event() {
            Ok(quick_xml::events::Event::Eof) => break,
            Ok(quick_xml::events::Event::Empty(event))
            | Ok(quick_xml::events::Event::Start(event))
                if event.local_name().as_ref() == b"Relationship" =>
            {
                let mut id = None;
                let mut target = None;
                for attribute in event.attributes().flatten() {
                    let value = attribute
                        .decoded_and_normalized_value(
                            quick_xml::XmlVersion::Implicit1_0,
                            reader.decoder(),
                        )
                        .map_err(|error| error.to_string())?
                        .into_owned();
                    match attribute.key.local_name().as_ref() {
                        b"Id" => id = Some(value),
                        b"Target" => target = Some(value),
                        _ => {}
                    }
                }
                if let (Some(id), Some(target)) = (id, target) {
                    relationships.insert(id, target);
                }
            }
            Ok(_) => {}
            Err(error) => return Err(error.to_string()),
        }
    }
    Ok(relationships)
}

fn worksheet_layout(data: &[u8]) -> Result<WorksheetLayout, String> {
    let mut layout = WorksheetLayout::default();
    let mut offset = 0usize;
    while offset < data.len() {
        let id = read_varint(data, &mut offset)
            .ok_or_else(|| "truncated XLSB worksheet record id".to_string())?;
        let length = read_varint(data, &mut offset)
            .ok_or_else(|| "truncated XLSB worksheet record length".to_string())?
            as usize;
        let end = offset
            .checked_add(length)
            .filter(|end| *end <= data.len())
            .ok_or_else(|| "XLSB worksheet record extends past the part boundary".to_string())?;
        let payload = &data[offset..end];
        offset = end;
        if id == 0x003c {
            let fields = payload
                .get(..18)
                .ok_or_else(|| "truncated XLSB column information".to_string())?;
            let first = u32::from_le_bytes(fields[0..4].try_into().unwrap()) as usize;
            let last = u32::from_le_bytes(fields[4..8].try_into().unwrap()) as usize;
            let width = u32::from_le_bytes(fields[8..12].try_into().unwrap()) as f64 / 256.0;
            layout.column_widths.push((first, last, width));
        } else if id == 0x0000 {
            let fields = payload
                .get(..12)
                .ok_or_else(|| "truncated XLSB row header".to_string())?;
            if fields[11] & 0x20 != 0 {
                let row = u32::from_le_bytes(fields[0..4].try_into().unwrap()) as usize;
                let height = u16::from_le_bytes(fields[8..10].try_into().unwrap()) as f64 / 20.0;
                layout.row_heights.push((row, height));
            }
        }
    }
    Ok(layout)
}

pub(crate) fn read_layouts(
    buffer: &[u8],
) -> Result<std::collections::HashMap<String, WorksheetLayout>, String> {
    let mut archive = zip::ZipArchive::new(Cursor::new(buffer))
        .map_err(|error| format!("failed to open XLSB package: {error}"))?;
    let mut workbook_data = Vec::new();
    archive
        .by_name("xl/workbook.bin")
        .map_err(|error| error.to_string())?
        .read_to_end(&mut workbook_data)
        .map_err(|error| error.to_string())?;
    let mut relationships_xml = String::new();
    archive
        .by_name("xl/_rels/workbook.bin.rels")
        .map_err(|error| error.to_string())?
        .read_to_string(&mut relationships_xml)
        .map_err(|error| error.to_string())?;
    let relationships = parse_workbook_relationships(&relationships_xml)?;
    let mut layouts = std::collections::HashMap::new();
    for (name, relationship) in workbook_sheets(&workbook_data)? {
        let target = relationships
            .get(&relationship)
            .ok_or_else(|| format!("missing XLSB worksheet relationship '{relationship}'"))?;
        let path = if target.starts_with('/') {
            target.trim_start_matches('/').to_string()
        } else {
            format!("xl/{}", target.trim_start_matches("../"))
        };
        let mut data = Vec::new();
        archive
            .by_name(&path)
            .map_err(|error| error.to_string())?
            .read_to_end(&mut data)
            .map_err(|error| error.to_string())?;
        layouts.insert(name, worksheet_layout(&data)?);
    }
    Ok(layouts)
}

fn validate_styles(data: &[u8]) -> Result<(), String> {
    let mut offset = 0usize;
    let mut in_cell_formats = false;
    while offset < data.len() {
        let id = read_varint(data, &mut offset)
            .ok_or_else(|| "truncated XLSB style record id".to_string())?;
        let length = read_varint(data, &mut offset)
            .ok_or_else(|| "truncated XLSB style record length".to_string())?
            as usize;
        let end = offset
            .checked_add(length)
            .filter(|end| *end <= data.len())
            .ok_or_else(|| "XLSB style record extends past the part boundary".to_string())?;
        let payload = &data[offset..end];
        offset = end;
        match id {
            0x0269 => in_cell_formats = true,
            0x026a => in_cell_formats = false,
            0x002c => {
                return Err("custom number formats are not supported in XLSB input".to_string());
            }
            0x002f if in_cell_formats => {
                let fields = payload
                    .get(..10)
                    .ok_or_else(|| "truncated XLSB cell format".to_string())?;
                let number_format = u16::from_le_bytes([fields[2], fields[3]]);
                let font = u16::from_le_bytes([fields[4], fields[5]]);
                let fill = u16::from_le_bytes([fields[6], fields[7]]);
                let border = u16::from_le_bytes([fields[8], fields[9]]);
                let supported_date_format =
                    (14..=22).contains(&number_format) || (45..=47).contains(&number_format);
                if number_format != 0 && !supported_date_format {
                    return Err(format!(
                        "number format {number_format} is not supported in XLSB input"
                    ));
                }
                if font != 0 || fill != 0 || border != 0 {
                    return Err(
                        "font, fill, and border formatting are not supported in XLSB input"
                            .to_string(),
                    );
                }
            }
            _ => {}
        }
    }
    Ok(())
}

fn validate_worksheet_records(data: &[u8]) -> Result<(), String> {
    let mut offset = 0usize;
    while offset < data.len() {
        let id = read_varint(data, &mut offset)
            .ok_or_else(|| "truncated XLSB worksheet record id".to_string())?;
        let length = read_varint(data, &mut offset)
            .ok_or_else(|| "truncated XLSB worksheet record length".to_string())?
            as usize;
        let end = offset
            .checked_add(length)
            .filter(|end| *end <= data.len())
            .ok_or_else(|| "XLSB worksheet record extends past the part boundary".to_string())?;
        let payload = &data[offset..end];
        offset = end;
        if id == 0x0000 {
            let flags = *payload
                .get(11)
                .ok_or_else(|| "truncated XLSB row header".to_string())?;
            if flags & !0x20 != 0 {
                return Err("hidden and outlined rows are not supported in XLSB input".to_string());
            }
        } else if id == 0x003c {
            let fields = payload
                .get(..18)
                .ok_or_else(|| "truncated XLSB column information".to_string())?;
            let style = u32::from_le_bytes(fields[12..16].try_into().unwrap());
            let flags = u16::from_le_bytes(fields[16..18].try_into().unwrap());
            if style != 0 || flags & !0x06 != 0 {
                return Err(
                    "hidden, outlined, or formatted columns are not supported in XLSB input"
                        .to_string(),
                );
            }
        }
        let feature = match id {
            0x0040 | 0x023d | 0x02a9 | 0x041d | 0x041e => Some("data validation"),
            0x00a1 => Some("auto filters"),
            0x00b0..=0x00b2 => Some("merged cells"),
            0x01aa => Some("array formulas"),
            0x01ab => Some("shared formulas"),
            0x01cd | 0x01ce | 0x0416 | 0x0417 | 0x046f | 0x0470 | 0x047a => {
                Some("conditional formatting")
            }
            0x01ee => Some("hyperlinks"),
            0x1000 | 0x1001 => Some("dynamic array properties"),
            _ => None,
        };
        if let Some(feature) = feature {
            return Err(format!(
                "unsupported XLSB worksheet feature '{feature}' in record 0x{id:04x}"
            ));
        }
        validate_formula_record(id, payload)?;
    }
    Ok(())
}

pub(crate) fn validate_import(buffer: &[u8]) -> Result<(), String> {
    let mut archive = zip::ZipArchive::new(Cursor::new(buffer))
        .map_err(|error| format!("failed to open XLSB package: {error}"))?;
    let rejected_paths = super::table::PATH_MARKERS
        .iter()
        .chain(super::chart::PATH_MARKERS)
        .chain(super::pivot::PATH_MARKERS)
        .chain(super::vba::PATH_MARKERS)
        .copied()
        .chain([
            "/externallinks/",
            "/connections",
            "/comments",
            "/threadedcomments/",
            "/printersettings/",
            "/media/",
            "/embeddings/",
            "/ctrlprops/",
            "/metadata.bin",
        ])
        .collect::<Vec<_>>();
    for index in 0..archive.len() {
        let mut file = archive.by_index(index).map_err(|error| error.to_string())?;
        let name = file.name().to_ascii_lowercase();
        if let Some(feature) = rejected_paths.iter().find(|path| name.contains(**path)) {
            return Err(format!(
                "unsupported XLSB feature in package part '{name}' ({feature})"
            ));
        }
        if name.starts_with("xl/worksheets/sheet") && name.ends_with(".bin") {
            let mut data = Vec::new();
            file.read_to_end(&mut data)
                .map_err(|error| error.to_string())?;
            validate_worksheet_records(&data)?;
        } else if name == "xl/styles.bin" {
            let mut data = Vec::new();
            file.read_to_end(&mut data)
                .map_err(|error| error.to_string())?;
            validate_styles(&data)?;
        }
    }
    for index in 0..archive.len() {
        let mut file = archive.by_index(index).map_err(|error| error.to_string())?;
        if !file.name().ends_with(".rels") {
            continue;
        }
        let mut xml = String::new();
        file.read_to_string(&mut xml)
            .map_err(|error| error.to_string())?;
        let lower = xml.to_ascii_lowercase();
        let rejected_relationships = super::table::RELATIONSHIP_MARKERS
            .iter()
            .chain(super::chart::RELATIONSHIP_MARKERS)
            .chain(super::pivot::RELATIONSHIP_MARKERS)
            .chain(super::vba::RELATIONSHIP_MARKERS)
            .copied()
            .chain([
                "/externallink",
                "/connection",
                "/comments",
                "/threadedcomment",
                "/hyperlink",
                "/printersettings",
                "/image",
                "/oleobject",
                "/control",
            ]);
        for relationship in rejected_relationships {
            if lower.contains(relationship) {
                return Err(format!(
                    "unsupported XLSB relationship type containing '{relationship}'"
                ));
            }
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn rejects_non_default_cell_formatting() {
        let mut data = Vec::new();
        super::super::record::write_empty_record(0x0269, &mut data);
        let mut format = [0u8; 16];
        format[4..6].copy_from_slice(&1u16.to_le_bytes());
        super::super::record::write_record(0x002f, &format, &mut data);
        super::super::record::write_empty_record(0x026a, &mut data);
        assert_eq!(
            validate_styles(&data).unwrap_err(),
            "font, fill, and border formatting are not supported in XLSB input"
        );
    }

    #[test]
    fn rejects_shared_formula_records() {
        let mut data = Vec::new();
        super::super::record::write_record(0x01ab, &[], &mut data);
        assert_eq!(
            validate_worksheet_records(&data).unwrap_err(),
            "unsupported XLSB worksheet feature 'shared formulas' in record 0x01ab"
        );
    }

    #[test]
    fn rejects_array_formula_tokens() {
        let mut payload = vec![0; 18];
        payload.extend_from_slice(&15u32.to_le_bytes());
        payload.push(0x20);
        payload.extend_from_slice(&[0; 14]);
        let mut data = Vec::new();
        super::super::record::write_record(0x0009, &payload, &mut data);
        assert_eq!(
            validate_worksheet_records(&data).unwrap_err(),
            "array constants are not supported in XLSB input"
        );
    }

    #[test]
    fn normalizes_standard_bundle_sheet_layout_for_calamine() {
        let mut payload = Vec::new();
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&1u32.to_le_bytes());
        super::super::record::write_wide_string("rId1", &mut payload);
        super::super::record::write_wide_string("Sheet1", &mut payload);
        let mut input = Vec::new();
        super::super::record::write_record(0x009c, &payload, &mut input);
        let (normalized, changed) = normalize_workbook_records(&input).unwrap();
        assert!(changed);
        let mut offset = 0;
        assert_eq!(read_varint(&normalized, &mut offset), Some(0x009c));
        let len = read_varint(&normalized, &mut offset).unwrap() as usize;
        let normalized_payload = &normalized[offset..offset + len];
        assert_eq!(
            u32::from_le_bytes(normalized_payload[8..12].try_into().unwrap()),
            4
        );
        assert_eq!(
            String::from_utf16(
                &normalized_payload[12..20]
                    .chunks_exact(2)
                    .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
                    .collect::<Vec<_>>()
            )
            .unwrap(),
            "rId1"
        );
    }
}
