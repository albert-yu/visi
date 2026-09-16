use std::process;

pub const EXIT_SUCCESS: i32 = 0;
pub const EXIT_USAGE_ERROR: i32 = 1;
pub const EXIT_IO_ERROR: i32 = 2;
pub const EXIT_ENGINE_ERROR: i32 = 3;

pub fn exit_with_error(msg: impl std::fmt::Display, code: i32) -> ! {
    eprintln!("Error: {}", msg);
    process::exit(code);
}

pub fn col_idx_to_letters(col: usize) -> String {
    visi_core::core::col_idx_to_letters(col)
}

/// Convert Excel letter notation (e.g. "A", "Z", "AA")
/// to 0-based column index
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

/// Parse row label from string,
/// expecting 1-based row number ("1", "100") -> 0-based index
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

pub fn parse_cell_ref(cell_str: &str) -> Result<(Option<String>, usize, usize), String> {
    visi_core::core::parse_cell_ref(cell_str)
}

pub fn parse_range_ref(
    range_str: &str,
) -> Result<(Option<String>, usize, usize, usize, usize), String> {
    visi_core::core::parse_range_ref(range_str)
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
