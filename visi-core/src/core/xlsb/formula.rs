use crate::core::engine::ResultData;
use crate::core::parser::{Expr, Op, parse_excel_formula};

use super::function_registry::function_id;
use super::record::{cell_header, write_record, write_wide_string};

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

fn encode_string(value: &str, out: &mut Vec<u8>) -> Result<(), String> {
    let encoded = value.encode_utf16().collect::<Vec<_>>();
    let len = u16::try_from(encoded.len()).map_err(|_| "formula string is too long".to_string())?;
    out.push(0x17);
    out.extend_from_slice(&len.to_le_bytes());
    for unit in encoded {
        out.extend_from_slice(&unit.to_le_bytes());
    }
    Ok(())
}

fn encode_ref(row: usize, col: usize, out: &mut Vec<u8>) -> Result<(), String> {
    let row = u32::try_from(row).map_err(|_| "formula row is out of range".to_string())?;
    let col = u16::try_from(col).map_err(|_| "formula column is out of range".to_string())?;
    if row > 1_048_575 || col > 16_383 {
        return Err("formula reference is outside Excel's grid".to_string());
    }
    out.push(0x44);
    out.extend_from_slice(&row.to_le_bytes());
    out.extend_from_slice(&col.to_le_bytes());
    Ok(())
}

fn encode_area(
    start_row: usize,
    start_col: usize,
    end_row: usize,
    end_col: usize,
    out: &mut Vec<u8>,
) -> Result<(), String> {
    let start_row =
        u32::try_from(start_row).map_err(|_| "formula row is out of range".to_string())?;
    let end_row = u32::try_from(end_row).map_err(|_| "formula row is out of range".to_string())?;
    let start_col =
        u16::try_from(start_col).map_err(|_| "formula column is out of range".to_string())?;
    let end_col =
        u16::try_from(end_col).map_err(|_| "formula column is out of range".to_string())?;
    if end_row > 1_048_575 || end_col > 16_383 {
        return Err("formula range is outside Excel's grid".to_string());
    }
    out.push(0x25);
    out.extend_from_slice(&start_row.to_le_bytes());
    out.extend_from_slice(&end_row.to_le_bytes());
    out.extend_from_slice(&start_col.to_le_bytes());
    out.extend_from_slice(&end_col.to_le_bytes());
    Ok(())
}

fn ensure_local_sheet(sheet: &Option<String>, current_sheet: &str) -> Result<(), String> {
    if let Some(sheet) = sheet
        && !sheet.eq_ignore_ascii_case(current_sheet)
    {
        return Err(format!(
            "cross-sheet formula references are not supported in XLSB output: {sheet}"
        ));
    }
    Ok(())
}

fn encode_expr(expr: &Expr, current_sheet: &str, out: &mut Vec<u8>) -> Result<(), String> {
    match expr {
        Expr::Number(value) => {
            if value.fract() == 0.0 && (0.0..=u16::MAX as f64).contains(value) {
                out.push(0x1e);
                out.extend_from_slice(&(*value as u16).to_le_bytes());
            } else {
                out.push(0x1f);
                out.extend_from_slice(&value.to_le_bytes());
            }
        }
        Expr::String(value) => encode_string(value, out)?,
        Expr::Boolean(value) => {
            out.push(0x1d);
            out.push(u8::from(*value));
        }
        Expr::Error(value) => {
            out.push(0x1c);
            out.push(error_code(value));
        }
        Expr::CellRef {
            sheet, row, col, ..
        } => {
            ensure_local_sheet(sheet, current_sheet)?;
            encode_ref(*row, *col, out)?;
        }
        Expr::RangeRef {
            sheet,
            start_row,
            start_col,
            end_row,
            end_col,
            ..
        } => {
            ensure_local_sheet(sheet, current_sheet)?;
            encode_area(*start_row, *start_col, *end_row, *end_col, out)?;
        }
        Expr::FunctionCall { name, args } => {
            for arg in args {
                encode_expr(arg, current_sheet, out)?;
            }
            let upper = name.to_ascii_uppercase();
            if upper == "SUM" && args.len() == 1 {
                out.extend_from_slice(&[0x19, 0x10, 0x00, 0x00]);
            } else {
                let id = function_id(&upper)
                    .ok_or_else(|| format!("unsupported XLSB formula function: {name}"))?;
                let argc = u8::try_from(args.len())
                    .map_err(|_| format!("too many arguments for XLSB formula function: {name}"))?;
                out.push(0x42);
                out.push(argc);
                out.extend_from_slice(&id.to_le_bytes());
            }
        }
        Expr::BinaryOp { op, left, right } => {
            encode_expr(left, current_sheet, out)?;
            encode_expr(right, current_sheet, out)?;
            out.push(match op {
                Op::Add => 0x03,
                Op::Sub => 0x04,
                Op::Mul => 0x05,
                Op::Div => 0x06,
                Op::Exp => 0x07,
                Op::Concat => 0x08,
                Op::Lt => 0x09,
                Op::Le => 0x0a,
                Op::Eq => 0x0b,
                Op::Gt => 0x0c,
                Op::Ge => 0x0d,
                Op::Ne => 0x0e,
                Op::Intersect => 0x0f,
                Op::Union => 0x10,
                _ => return Err(format!("unsupported XLSB binary formula operator: {op:?}")),
            });
        }
        Expr::UnaryOp { op, expr } => {
            encode_expr(expr, current_sheet, out)?;
            out.push(match op {
                Op::Add => 0x12,
                Op::Sub => 0x13,
                Op::Percent => 0x14,
                _ => return Err(format!("unsupported XLSB unary formula operator: {op:?}")),
            });
        }
        Expr::List(_) | Expr::Slice { .. } | Expr::Identifier(_) | Expr::StructuredRef { .. } => {
            return Err(format!("unsupported XLSB formula expression: {expr:?}"));
        }
    }
    Ok(())
}

pub(crate) fn encode_tokens(source: &str, current_sheet: &str) -> Result<Vec<u8>, String> {
    let formula = source.strip_prefix('=').unwrap_or(source);
    let expr = parse_excel_formula(formula)
        .map_err(|error| format!("failed to parse XLSB formula '{source}': {error}"))?;
    let mut out = Vec::new();
    encode_expr(&expr, current_sheet, &mut out)?;
    Ok(out)
}

fn append_formula(tokens: &[u8], payload: &mut Vec<u8>) {
    payload.extend_from_slice(&(tokens.len() as u32).to_le_bytes());
    payload.extend_from_slice(tokens);
    payload.extend_from_slice(&0u32.to_le_bytes());
}

pub(crate) fn write_formula_cell(
    col: usize,
    style: u32,
    source: &str,
    cached: &ResultData,
    current_sheet: &str,
    out: &mut Vec<u8>,
) -> Result<(), String> {
    let tokens = encode_tokens(source, current_sheet)?;
    let header = cell_header(col, style);
    let mut payload = Vec::new();
    payload.extend_from_slice(&header);
    match cached {
        ResultData::String(value) => {
            write_wide_string(value, &mut payload);
            payload.extend_from_slice(&0u16.to_le_bytes());
            append_formula(&tokens, &mut payload);
            write_record(0x0008, &payload, out);
        }
        ResultData::Boolean(value) => {
            payload.push(u8::from(*value));
            payload.extend_from_slice(&0u16.to_le_bytes());
            append_formula(&tokens, &mut payload);
            write_record(0x000a, &payload, out);
        }
        ResultData::Error(value) => {
            payload.push(error_code(value));
            payload.extend_from_slice(&0u16.to_le_bytes());
            append_formula(&tokens, &mut payload);
            write_record(0x000b, &payload, out);
        }
        ResultData::Integer(value) => {
            payload.extend_from_slice(&(*value as f64).to_le_bytes());
            payload.extend_from_slice(&0u16.to_le_bytes());
            append_formula(&tokens, &mut payload);
            write_record(0x0009, &payload, out);
        }
        ResultData::Float(value) => {
            payload.extend_from_slice(&value.to_le_bytes());
            payload.extend_from_slice(&0u16.to_le_bytes());
            append_formula(&tokens, &mut payload);
            write_record(0x0009, &payload, out);
        }
        ResultData::None => {
            payload.extend_from_slice(&0f64.to_le_bytes());
            payload.extend_from_slice(&0u16.to_le_bytes());
            append_formula(&tokens, &mut payload);
            write_record(0x0009, &payload, out);
        }
        ResultData::List(_) | ResultData::Dict(_) => {
            return Err(
                "array and dictionary formula caches are not supported in XLSB output".to_string(),
            );
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn encodes_postfix_formula_tokens() {
        assert_eq!(
            encode_tokens("=A1+2", "Sheet1").unwrap(),
            [0x44, 0, 0, 0, 0, 0, 0, 0x1e, 2, 0, 0x03]
        );
    }

    #[test]
    fn sum_uses_a_reference_class_range_and_attr_sum() {
        assert_eq!(
            encode_tokens("=SUM(A1:A2)", "Sheet1").unwrap(),
            [0x25, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0x19, 0x10, 0, 0,]
        );
    }

    #[test]
    fn rejects_cross_sheet_references_until_3d_tokens_are_supported() {
        assert!(encode_tokens("=Sheet2!A1", "Sheet1").is_err());
    }
}
