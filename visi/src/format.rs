use crate::cli::OutputFormat;
use crate::utils::col_idx_to_letters;
use serde_json::{Value, json};
use visi_core::core::engine::{CellRef, Sheet};

/// Get string representation of cell value (raw formula string vs calculated result)
pub fn get_cell_display_val(sheet: &Sheet, row: usize, col: usize, raw: bool) -> String {
    let cell_ref = CellRef::new(row, col);
    if raw {
        sheet.get_src(&cell_ref).cloned().unwrap_or_default()
    } else {
        let val_str = sheet.get_display_string(&cell_ref);
        if val_str.is_empty() {
            sheet.get_src(&cell_ref).cloned().unwrap_or_default()
        } else {
            val_str
        }
    }
}

/// Render a cell range grid according to specified OutputFormat
#[allow(clippy::too_many_arguments)]
pub fn render_grid(
    sheet: &Sheet,
    min_row: usize,
    min_col: usize,
    max_row: usize,
    max_col: usize,
    raw: bool,
    use_headers: bool,
    format: OutputFormat,
) -> String {
    let row_count = if max_row >= min_row {
        max_row - min_row + 1
    } else {
        0
    };
    let col_count = if max_col >= min_col {
        max_col - min_col + 1
    } else {
        0
    };

    if row_count == 0 || col_count == 0 {
        return "(empty grid)".to_string();
    }

    let mut matrix: Vec<Vec<String>> = Vec::with_capacity(row_count);
    for r in min_row..=max_row {
        let mut row_vals = Vec::with_capacity(col_count);
        for c in min_col..=max_col {
            row_vals.push(get_cell_display_val(sheet, r, c, raw));
        }
        matrix.push(row_vals);
    }

    match format {
        OutputFormat::Table => render_ascii_table(sheet, min_row, min_col, &matrix, use_headers),
        OutputFormat::Csv => render_delimited(&matrix, ",", use_headers),
        OutputFormat::Tsv => render_delimited(&matrix, "\t", use_headers),
        OutputFormat::Json => render_json(&matrix, min_col, use_headers),
    }
}

/// Render pretty ASCII table
fn render_ascii_table(
    _sheet: &Sheet,
    min_row: usize,
    min_col: usize,
    matrix: &[Vec<String>],
    use_headers: bool,
) -> String {
    if matrix.is_empty() || matrix[0].is_empty() {
        return String::new();
    }

    let num_rows = matrix.len();
    let num_cols = matrix[0].len();

    let mut header_row_labels = Vec::with_capacity(num_cols);
    if use_headers {
        header_row_labels = matrix[0].clone();
    } else {
        for c in 0..num_cols {
            header_row_labels.push(col_idx_to_letters(min_col + c));
        }
    }

    let start_data_row = if use_headers { 1 } else { 0 };

    let mut col_widths = Vec::with_capacity(num_cols + 1);
    let max_row_label = (min_row + num_rows).to_string();
    let mut label_col_width = max_row_label.len().max(2);
    if use_headers {
        label_col_width = label_col_width.max(1);
    }
    col_widths.push(label_col_width);

    for (c, header_label) in header_row_labels.iter().enumerate().take(num_cols) {
        let mut max_w = header_label.len();
        for row in matrix.iter().skip(start_data_row) {
            if c < row.len() {
                max_w = max_w.max(row[c].len());
            }
        }
        col_widths.push(max_w.max(1));
    }

    let mut out = String::new();

    let make_separator = || {
        let mut line = String::from("+");
        for w in &col_widths {
            line.push_str(&"-".repeat(*w + 2));
            line.push('+');
        }
        line.push('\n');
        line
    };

    out.push_str(&make_separator());

    out.push('|');
    out.push_str(&format!(" {:^width$} |", "#", width = col_widths[0]));
    for (c, header_label) in header_row_labels.iter().enumerate().take(num_cols) {
        out.push_str(&format!(
            " {:^width$} |",
            header_label,
            width = col_widths[c + 1]
        ));
    }
    out.push('\n');
    out.push_str(&make_separator());

    for (r_idx, row) in matrix.iter().enumerate().skip(start_data_row) {
        let row_num = min_row + r_idx + 1;
        out.push('|');
        out.push_str(&format!(" {:>width$} |", row_num, width = col_widths[0]));
        for (c, cell_val) in row.iter().enumerate().take(num_cols) {
            out.push_str(&format!(
                " {:<width$} |",
                cell_val,
                width = col_widths[c + 1]
            ));
        }
        out.push('\n');
    }

    out.push_str(&make_separator());
    out
}

/// Render CSV / TSV delimited text
fn render_delimited(matrix: &[Vec<String>], delimiter: &str, _use_headers: bool) -> String {
    let mut out = String::new();

    for row in matrix.iter() {
        let escaped_cells: Vec<String> = row
            .iter()
            .map(|cell| {
                if cell.contains(delimiter) || cell.contains('"') || cell.contains('\n') {
                    format!("\"{}\"", cell.replace('"', "\"\""))
                } else {
                    cell.clone()
                }
            })
            .collect();
        out.push_str(&escaped_cells.join(delimiter));
        out.push('\n');
    }

    out
}

/// Render JSON array of rows or JSON array of objects if headers is true
fn render_json(matrix: &[Vec<String>], min_col: usize, use_headers: bool) -> String {
    if matrix.is_empty() {
        return "[]".to_string();
    }

    if use_headers && matrix.len() > 1 {
        let headers = &matrix[0];
        let mut rows_json = Vec::new();

        for row in matrix.iter().skip(1) {
            let mut obj = serde_json::Map::new();
            for (c, header) in headers.iter().enumerate() {
                let key = if header.trim().is_empty() {
                    col_idx_to_letters(min_col + c)
                } else {
                    header.clone()
                };
                let val_str = row.get(c).cloned().unwrap_or_default();
                let json_val = parse_json_value(&val_str);
                obj.insert(key, json_val);
            }
            rows_json.push(Value::Object(obj));
        }

        serde_json::to_string_pretty(&rows_json).unwrap_or_else(|_| "[]".to_string())
    } else {
        let json_matrix: Vec<Vec<Value>> = matrix
            .iter()
            .map(|row| row.iter().map(|s| parse_json_value(s)).collect())
            .collect();

        serde_json::to_string_pretty(&json_matrix).unwrap_or_else(|_| "[]".to_string())
    }
}

/// Convert string into appropriate JSON primitive type (number, bool, or string)
fn parse_json_value(s: &str) -> Value {
    let trimmed = s.trim();
    if trimmed.is_empty() {
        Value::Null
    } else if let Ok(i) = trimmed.parse::<i64>() {
        json!(i)
    } else if let Ok(f) = trimmed.parse::<f64>() {
        json!(f)
    } else if trimmed.eq_ignore_ascii_case("true") {
        json!(true)
    } else if trimmed.eq_ignore_ascii_case("false") {
        json!(false)
    } else {
        json!(s)
    }
}
