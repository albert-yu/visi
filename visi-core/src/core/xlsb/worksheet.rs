use crate::core::engine::{CellType, ResultData, Sheet};

use super::formula::write_formula_cell;
use super::record::{cell_header, encode_rk, write_empty_record, write_record, write_row_header};
use super::shared_strings::SharedStrings;
use super::styles::style_index;

fn error_code(value: &str) -> u8 {
    match value.trim().to_ascii_uppercase().as_str() {
        "#NULL!" => 0x00,
        "#DIV/0!" => 0x07,
        "#REF!" => 0x17,
        "#NAME?" => 0x1d,
        "#NUM!" => 0x24,
        "#N/A" => 0x2a,
        "#GETTING_DATA" => 0x2b,
        _ => 0x0f,
    }
}

fn write_sheet_header(
    first_row: usize,
    last_row: usize,
    first_col: usize,
    last_col: usize,
    out: &mut Vec<u8>,
) {
    write_empty_record(0x0081, out);
    let mut dimensions = [0u8; 16];
    dimensions[..4].copy_from_slice(&(first_row as u32).to_le_bytes());
    dimensions[4..8].copy_from_slice(&(last_row as u32).to_le_bytes());
    dimensions[8..12].copy_from_slice(&(first_col as u32).to_le_bytes());
    dimensions[12..16].copy_from_slice(&(last_col as u32).to_le_bytes());
    write_record(0x0094, &dimensions, out);
}

fn write_column_widths(sheet: &Sheet, out: &mut Vec<u8>) -> Result<(), String> {
    if !sheet.columns.iter().any(|column| column.width.is_some()) {
        return Ok(());
    }
    write_empty_record(0x0186, out);
    for (index, column) in sheet.columns.iter().enumerate() {
        let Some(width) = column.width else {
            continue;
        };
        if !width.is_finite() || width <= 0.0 || width > 255.0 {
            return Err("column widths must be finite and between 0 and 255".to_string());
        }
        let mut payload = [0u8; 18];
        payload[..4].copy_from_slice(&(index as u32).to_le_bytes());
        payload[4..8].copy_from_slice(&(index as u32).to_le_bytes());
        payload[8..12].copy_from_slice(&((width * 256.0).round() as u32).to_le_bytes());
        payload[16..18].copy_from_slice(&2u16.to_le_bytes());
        write_record(0x003c, &payload, out);
    }
    write_empty_record(0x0187, out);
    Ok(())
}

fn write_blank(col: usize, style: u32, out: &mut Vec<u8>) {
    let mut header = cell_header(col, style);
    header[6..8].copy_from_slice(&0x0100u16.to_le_bytes());
    write_record(0x0001, &header, out);
}

fn write_number(col: usize, style: u32, value: f64, out: &mut Vec<u8>) {
    if let Some(rk) = encode_rk(value) {
        let mut payload = Vec::from(cell_header(col, style));
        payload.extend_from_slice(&rk.to_le_bytes());
        write_record(0x0002, &payload, out);
    } else {
        let mut payload = Vec::from(cell_header(col, style));
        payload.extend_from_slice(&value.to_le_bytes());
        write_record(0x0005, &payload, out);
    }
}

fn write_literal(
    col: usize,
    style: u32,
    source: &str,
    cell_type: CellType,
    value: &ResultData,
    strings: &mut SharedStrings,
    out: &mut Vec<u8>,
) -> Result<(), String> {
    let source_value;
    let value = if matches!(value, ResultData::None) && !source.is_empty() {
        source_value = match cell_type {
            CellType::Int => source
                .parse::<i64>()
                .map(ResultData::Integer)
                .unwrap_or_else(|_| ResultData::String(source.to_string())),
            CellType::Float | CellType::DateTime => source
                .parse::<f64>()
                .map(ResultData::Float)
                .unwrap_or_else(|_| ResultData::String(source.to_string())),
            CellType::Bool => ResultData::Boolean(source.eq_ignore_ascii_case("true")),
            CellType::Error => ResultData::Error(source.to_string()),
            _ => ResultData::String(source.to_string()),
        };
        &source_value
    } else {
        value
    };
    match value {
        ResultData::None => write_blank(col, style, out),
        ResultData::Boolean(value) => {
            let mut payload = Vec::from(cell_header(col, style));
            payload.push(u8::from(*value));
            write_record(0x0004, &payload, out);
        }
        ResultData::Integer(value) => write_number(col, style, *value as f64, out),
        ResultData::Float(value) if value.is_finite() => write_number(col, style, *value, out),
        ResultData::Float(_) => {
            let mut payload = Vec::from(cell_header(col, style));
            payload.push(0x24);
            write_record(0x0003, &payload, out);
        }
        ResultData::String(value) => {
            let index = strings.intern(value);
            let mut payload = Vec::from(cell_header(col, style));
            payload.extend_from_slice(&index.to_le_bytes());
            write_record(0x0007, &payload, out);
        }
        ResultData::Error(value) => {
            let mut payload = Vec::from(cell_header(col, style));
            payload.push(error_code(value));
            write_record(0x0003, &payload, out);
        }
        ResultData::List(_) | ResultData::Dict(_) => {
            return Err("list and dictionary values are not supported in XLSB output".to_string());
        }
    }
    Ok(())
}

fn occupied_bounds(sheet: &Sheet) -> Option<(usize, usize, usize, usize)> {
    let mut bounds: Option<(usize, usize, usize, usize)> = None;
    for (col, column) in sheet.columns.iter().enumerate() {
        if column.width.is_some() {
            bounds = Some(match bounds {
                None => (0, 0, col, col),
                Some((first_row, last_row, first_col, last_col)) => {
                    (first_row, last_row, first_col.min(col), last_col.max(col))
                }
            });
        }
    }
    for (row, height) in sheet.row_heights.iter().enumerate() {
        if height.is_some() {
            bounds = Some(match bounds {
                None => (row, row, 0, 0),
                Some((first_row, last_row, first_col, last_col)) => {
                    (first_row.min(row), last_row.max(row), first_col, last_col)
                }
            });
        }
    }
    for (col_index, column) in sheet.columns.iter().enumerate() {
        for row in 0..column.len() {
            let occupied = column.src(row).is_some_and(|value| !value.is_empty())
                || !matches!(column.value(row), None | Some(ResultData::None))
                || column.styles.get(row).is_some_and(Option::is_some);
            if occupied {
                bounds = Some(match bounds {
                    None => (row, row, col_index, col_index),
                    Some((first_row, last_row, first_col, last_col)) => (
                        first_row.min(row),
                        last_row.max(row),
                        first_col.min(col_index),
                        last_col.max(col_index),
                    ),
                });
            }
        }
    }
    bounds
}

pub(crate) fn build(sheet: &Sheet, strings: &mut SharedStrings) -> Result<Vec<u8>, String> {
    let bounds = occupied_bounds(sheet).unwrap_or((0, 0, 0, 0));
    let mut out = Vec::new();
    write_sheet_header(bounds.0, bounds.1, bounds.2, bounds.3, &mut out);
    write_column_widths(sheet, &mut out)?;
    write_empty_record(0x0091, &mut out);
    for row in bounds.0..=bounds.1 {
        let occupied_cols = (bounds.2..=bounds.3)
            .filter(|&col| {
                sheet.columns.get(col).is_some_and(|column| {
                    column.src(row).is_some_and(|value| !value.is_empty())
                        || !matches!(column.value(row), None | Some(ResultData::None))
                        || column.styles.get(row).is_some_and(Option::is_some)
                })
            })
            .collect::<Vec<_>>();
        let height = sheet.row_heights.get(row).copied().flatten();
        let Some(first_col) = occupied_cols.first().copied().or(height.map(|_| 0)) else {
            continue;
        };
        let last_col = occupied_cols.last().copied().unwrap_or(first_col);
        write_row_header(row, first_col, last_col, height, &mut out)?;
        for col in occupied_cols {
            let column = &sheet.columns[col];
            let source = column.src(row).unwrap_or_default();
            let value = column.value(row).unwrap_or(ResultData::None);
            let cell_type = column.cell_type(row).unwrap_or(CellType::Empty);
            let style = style_index(column.styles.get(row).and_then(Option::as_ref), cell_type)?;
            if source.starts_with('=') {
                write_formula_cell(col, style, source, &value, &sheet.name, &mut out)?;
            } else {
                write_literal(col, style, source, cell_type, &value, strings, &mut out)?;
            }
        }
    }
    write_empty_record(0x0092, &mut out);
    write_empty_record(0x0082, &mut out);
    Ok(out)
}
