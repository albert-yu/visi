use super::*;

pub(crate) fn eval1(source: &str) -> ResultData {
    let sheet = Sheet::new(SheetInit::default());
    sheet.eval(source, None).unwrap().0
}

pub(crate) fn assert_float_close(result: &ResultData, expected: f64, tol: f64) {
    match result {
        ResultData::Float(f) => assert!((f - expected).abs() < tol, "expected {expected}, got {f}"),
        ResultData::Integer(i) => assert!(
            (*i as f64 - expected).abs() < tol,
            "expected {expected}, got {i}"
        ),
        other => panic!("expected numeric result close to {expected}, got {other:?}"),
    }
}

pub(crate) fn create_sheet<const ROWS: usize, const COLS: usize>(
    grid: &[[&str; COLS]; ROWS],
) -> Sheet {
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
mod database;
mod dynamic_array;
mod extended;
mod financial;
mod lambda;
mod locale;
mod logical;
mod lookup_ref;
mod math;
mod math_trig;
mod rounding;
mod stats;
mod text;
mod text_fn;
mod unit;
