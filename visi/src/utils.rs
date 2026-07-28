use std::process;

pub const EXIT_SUCCESS: i32 = 0;
pub const EXIT_USAGE_ERROR: i32 = 1;
pub const EXIT_IO_ERROR: i32 = 2;
pub const EXIT_ENGINE_ERROR: i32 = 3;

/// Exit the program with a specific exit code and print error message to stderr.
pub fn exit_with_error(msg: impl AsRef<str>, code: i32) -> ! {
    eprintln!("Error: {}", msg.as_ref());
    process::exit(code);
}

/// Convert column index (0-based) to Excel letter notation (e.g. 0 -> "A", 25 -> "Z", 26 -> "AA")
pub fn col_idx_to_letters(mut col: usize) -> String {
    let mut letters = String::new();
    loop {
        let remainder = col % 26;
        letters.insert(0, (b'A' + remainder as u8) as char);
        if col < 26 {
            break;
        }
        col = col / 26 - 1;
    }
    letters
}

/// Convert Excel letter notation (e.g. "A", "Z", "AA") to 0-based column index
pub fn col_letters_to_idx(letters: &str) -> Result<usize, String> {
    let s = letters.trim().to_uppercase();
    if s.is_empty() {
        return Err("Column letter string cannot be empty".to_string());
    }
    let mut col = 0usize;
    for c in s.chars() {
        if !c.is_ascii_alphabetic() {
            return Err(format!("Invalid character '{}' in column letters", c));
        }
        col = col * 26 + (c as usize - 'A' as usize + 1);
    }
    if col == 0 {
        Err("Invalid column specification".to_string())
    } else {
        Ok(col - 1)
    }
}

/// Parse column index from string, accepting either letters ("A", "BC") or 1-based number ("1", "5")
pub fn parse_col_spec(spec: &str) -> Result<usize, String> {
    let trimmed = spec.trim();
    if trimmed.is_empty() {
        return Err("Column specification cannot be empty".to_string());
    }
    if trimmed.chars().all(|c| c.is_ascii_digit()) {
        let num: usize = trimmed
            .parse()
            .map_err(|_| format!("Invalid column number '{}'", trimmed))?;
        if num == 0 {
            return Err("Column index must be 1-based (minimum 1)".to_string());
        }
        Ok(num - 1)
    } else {
        col_letters_to_idx(trimmed)
    }
}

/// Parse row index from string, expecting 1-based row number ("1", "100") -> 0-based index
pub fn parse_row_spec(spec: &str) -> Result<usize, String> {
    let trimmed = spec.trim();
    let num: usize = trimmed
        .parse()
        .map_err(|_| format!("Invalid row number '{}'", trimmed))?;
    if num == 0 {
        return Err("Row number must be 1-based (minimum 1)".to_string());
    }
    Ok(num - 1)
}

/// Parse a cell reference string like "A1", "C10", or "Sheet1!B5"
/// Returns (optional_sheet_name, row_idx, col_idx)
pub fn parse_cell_ref(cell_str: &str) -> Result<(Option<String>, usize, usize), String> {
    let trimmed = cell_str.trim();
    if trimmed.is_empty() {
        return Err("Cell reference cannot be empty".to_string());
    }

    let (sheet_part, cell_part) = if let Some(pos) = trimmed.rfind('!') {
        let sheet = &trimmed[..pos];
        let cell = &trimmed[pos + 1..];
        (Some(sheet.trim_matches('\'').to_string()), cell)
    } else {
        (None, trimmed)
    };

    let chars: Vec<char> = cell_part.chars().collect();
    let mut letters = String::new();
    let mut digits = String::new();

    for &c in &chars {
        if c == '$' {
            continue; // Ignore absolute reference symbols like $A$1
        }
        if c.is_ascii_alphabetic() {
            if !digits.is_empty() {
                return Err(format!("Invalid cell reference format: '{}'", cell_str));
            }
            letters.push(c);
        } else if c.is_ascii_digit() {
            digits.push(c);
        } else {
            return Err(format!("Invalid character '{}' in cell reference", c));
        }
    }

    if letters.is_empty() || digits.is_empty() {
        return Err(format!(
            "Invalid cell reference '{}', must combine column letter(s) and row number (e.g. A1)",
            cell_str
        ));
    }

    let col_idx = col_letters_to_idx(&letters)?;
    let row_idx = parse_row_spec(&digits)?;

    Ok((sheet_part, row_idx, col_idx))
}

/// Parse a range reference string like "A1:C10", "Sheet1!A1:B5", or "A1"
/// Returns (optional_sheet_name, start_row, start_col, end_row, end_col)
pub fn parse_range_ref(
    range_str: &str,
) -> Result<(Option<String>, usize, usize, usize, usize), String> {
    let trimmed = range_str.trim();
    if trimmed.is_empty() {
        return Err("Range reference cannot be empty".to_string());
    }

    let (sheet_part, range_part) = if let Some(pos) = trimmed.rfind('!') {
        let sheet = &trimmed[..pos];
        let range = &trimmed[pos + 1..];
        (Some(sheet.trim_matches('\'').to_string()), range)
    } else {
        (None, trimmed)
    };

    if let Some(colon_pos) = range_part.find(':') {
        let start_str = &range_part[..colon_pos];
        let end_str = &range_part[colon_pos + 1..];

        let (_, start_row, start_col) = parse_cell_ref(start_str)?;
        let (_, end_row, end_col) = parse_cell_ref(end_str)?;

        let min_row = start_row.min(end_row);
        let max_row = start_row.max(end_row);
        let min_col = start_col.min(end_col);
        let max_col = start_col.max(end_col);

        Ok((sheet_part, min_row, min_col, max_row, max_col))
    } else {
        let (_, row_idx, col_idx) = parse_cell_ref(range_part)?;
        Ok((sheet_part, row_idx, col_idx, row_idx, col_idx))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_col_conversions() {
        assert_eq!(col_idx_to_letters(0), "A");
        assert_eq!(col_idx_to_letters(25), "Z");
        assert_eq!(col_idx_to_letters(26), "AA");
        assert_eq!(col_idx_to_letters(27), "AB");

        assert_eq!(col_letters_to_idx("A").unwrap(), 0);
        assert_eq!(col_letters_to_idx("z").unwrap(), 25);
        assert_eq!(col_letters_to_idx("AA").unwrap(), 26);
        assert_eq!(col_letters_to_idx("AB").unwrap(), 27);
    }

    #[test]
    fn test_parse_cell_ref() {
        let (sheet, row, col) = parse_cell_ref("A1").unwrap();
        assert_eq!(sheet, None);
        assert_eq!(row, 0);
        assert_eq!(col, 0);

        let (sheet, row, col) = parse_cell_ref("Sheet1!C5").unwrap();
        assert_eq!(sheet, Some("Sheet1".to_string()));
        assert_eq!(row, 4);
        assert_eq!(col, 2);
    }

    #[test]
    fn test_parse_range_ref() {
        let (sheet, s_row, s_col, e_row, e_col) = parse_range_ref("A1:C10").unwrap();
        assert_eq!(sheet, None);
        assert_eq!((s_row, s_col), (0, 0));
        assert_eq!((e_row, e_col), (9, 2));

        let (sheet, s_row, s_col, e_row, e_col) = parse_range_ref("'Data Sheet'!B2:D4").unwrap();
        assert_eq!(sheet, Some("Data Sheet".to_string()));
        assert_eq!((s_row, s_col), (1, 1));
        assert_eq!((e_row, e_col), (3, 3));
    }
}
