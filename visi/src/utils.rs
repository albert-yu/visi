use std::process;

pub use visi_core::core::{
    col_idx_to_letters, col_letters_to_idx, parse_cell_ref, parse_range_ref, parse_row_spec,
};

pub const EXIT_SUCCESS: i32 = 0;
pub const EXIT_USAGE_ERROR: i32 = 1;
pub const EXIT_IO_ERROR: i32 = 2;
pub const EXIT_ENGINE_ERROR: i32 = 3;

pub fn exit_with_error(msg: impl std::fmt::Display, code: i32) -> ! {
    eprintln!("Error: {}", msg);
    process::exit(code);
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
