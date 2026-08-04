use super::*;

fn create_sheet<const ROWS: usize, const COLS: usize>(grid: &[[&str; COLS]; ROWS]) -> Sheet {
    let rows = grid.len();
    let cols = grid[0].len();
    let mut sheet = Sheet::new(SheetInit {
        name: Some("sheet1".to_string()),
        rows,
        cols,
        ..Default::default()
    });
    for (i, row) in grid.iter().enumerate() {
        for (j, val) in row.iter().enumerate() {
            sheet.insert(
                TextCellRef {
                    row: i,
                    col: j,
                    char_offset: 0,
                },
                val,
            )
        }
    }

    sheet
}

mod aggregate;
mod logical;
mod math;
mod math_trig;
mod rounding;
mod stats;
mod text;
mod unit;
