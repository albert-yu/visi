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

#[test]
fn test_upper_rounddown_combo() {
    let sheet_src = [
        // A   B   C     D   E
        ["-60", "", "-6", "", "1"],
        ["cJsjkQ", "183.83", "", "TRUE", "FALSE"],
        ["32", "", "-96", "-25", "-205.145"],
        ["177.49", "-423.909", "", "-27", "n"],
        ["-77", "SusOPQc", "86", "", "-107.2312"],
        [
            "=(MIN(A1, A4) / PRODUCT(C1:E4))",
            "=(C2 * SQRT(-49))",
            "=(B3 * IF((E4 > -49), E1, E1))",
            "=LEFT(\"ABS(-17)\", 2)",
            "=IF((-41 > C3), SUM(B1:B5), C3)",
        ],
        [
            "=IF((UPPER(\"A4\") > C3), ROUNDDOWN(B6, 0), IF((-21 > -30), C2, A6))",
            "=ABS(D4)",
            "=LOWER(\"SUM(D2:E6)\")",
            "-112",
            "=ABS(ROUNDDOWN(B4, 2))",
        ],
    ];
    let mut sheet = create_sheet(&sheet_src);
    sheet.commit(None).unwrap();
    // A7
    let cell_ref = CellRef::new(6, 0);
    let target = sheet.get_result_data(&cell_ref);
    match target {
        ResultData::Error(e) => assert_eq!(e, "#NUM!"),
        other => panic!("Expected #NUM!, got {:?}", other),
    }
}
