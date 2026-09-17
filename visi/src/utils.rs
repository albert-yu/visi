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

pub fn col_letters_to_idx(letters: &str) -> Result<usize, String> {
    let s = letters.trim().to_uppercase();
    if s.is_empty() {
        return Err("Column letter string cannot be empty".to_string());
    }
    let col = visi_core::core::col_letters_to_idx(&s);
    Ok(col)
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
