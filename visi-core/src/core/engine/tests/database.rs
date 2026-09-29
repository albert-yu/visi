use super::*;

// --- Database (D*) functions --------------------------------------------
//
// Uses Microsoft's own classic "Tree/Height/Age/Yield/Profit" documented
// example dataset. Every expected value below was confirmed against real
// Microsoft Excel via the differential fuzzer.
//
//     A       B       C    D      E        G     H     J       K
//  1  Tree    Height  Age  Yield  Profit   Tree  Height Tree    Tree
//  2  Apple   18      20   14     105      Apple >12    Pear    Cherry
//  3  Pear    12      12   10     96                    Cherry  Tree
//  4  Cherry  13      14   9      105
//  5  Apple   14      15   10     75       Tree
//  6  Pear    9       8    8      76.8     Pear
//  7  Apple   8       9    6      45       Cherry
//
// G1:H2 = Tree="Apple" AND Height>12 (matches rows 2 and 5: Profit 105, 75)
// J1:J3 = Tree="Pear" OR Tree="Cherry" (matches rows 3, 4, 6: Profit 96, 105, 76.8)
// J1:J2 alone = Tree="Pear" (matches rows 3 and 6 -> ambiguous for DGET)
// K1:K2 = Tree="Cherry" (matches row 4 only -> unique for DGET)
// K3:K4 = Tree header with a blank criteria row -> matches everything
fn database_test_sheet() -> Sheet {
    let grid: [[&str; 12]; 7] = [
        [
            "Tree",
            "Height",
            "Age",
            "Yield",
            "Profit",
            "",
            "Tree",
            "Height",
            "",
            "Tree",
            "Tree",
            "=DSUM(A1:E7, \"Profit\", G1:H2)",
        ],
        [
            "Apple",
            "18",
            "20",
            "14",
            "105",
            "",
            "Apple",
            ">12",
            "",
            "Pear",
            "Cherry",
            "=DAVERAGE(A1:E7, \"Yield\", G1:H2)",
        ],
        [
            "Pear",
            "12",
            "12",
            "10",
            "96",
            "",
            "",
            "",
            "",
            "Cherry",
            "Tree",
            "=DCOUNT(A1:E7, \"Age\", G1:H2)",
        ],
        [
            "Cherry",
            "13",
            "14",
            "9",
            "105",
            "",
            "",
            "",
            "",
            "",
            "",
            "=DCOUNTA(A1:E7, \"Tree\", G1:H2)",
        ],
        [
            "Apple",
            "14",
            "15",
            "10",
            "75",
            "",
            "Tree",
            "",
            "",
            "",
            "",
            "=DMAX(A1:E7, \"Profit\", G1:H2)",
        ],
        [
            "Pear",
            "9",
            "8",
            "8",
            "76.8",
            "",
            "Pear",
            "",
            "",
            "",
            "",
            "=DMIN(A1:E7, \"Profit\", G1:H2)",
        ],
        [
            "Apple",
            "8",
            "9",
            "6",
            "45",
            "",
            "Cherry",
            "",
            "",
            "",
            "",
            "=DPRODUCT(A1:E7, \"Yield\", G1:H2)",
        ],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    sheet
}

#[test]
fn test_database_functions_and_criteria_match_real_excel() {
    let sheet = database_test_sheet();
    assert_float_close(&sheet.get_result_data(&CellRef::new(0, 11)), 180.0, 1e-9); // DSUM
    assert_float_close(&sheet.get_result_data(&CellRef::new(1, 11)), 12.0, 1e-9); // DAVERAGE
    assert_float_close(&sheet.get_result_data(&CellRef::new(2, 11)), 2.0, 1e-9); // DCOUNT
    assert_float_close(&sheet.get_result_data(&CellRef::new(3, 11)), 2.0, 1e-9); // DCOUNTA
    assert_float_close(&sheet.get_result_data(&CellRef::new(4, 11)), 105.0, 1e-9); // DMAX
    assert_float_close(&sheet.get_result_data(&CellRef::new(5, 11)), 75.0, 1e-9); // DMIN
    assert_float_close(&sheet.get_result_data(&CellRef::new(6, 11)), 140.0, 1e-9); // DPRODUCT
}

#[test]
fn test_dget_unique_match_or_error() {
    let mut sheet = database_test_sheet();
    sheet.set_cell_src(0, 11, "=DGET(A1:E7, \"Profit\", K1:K2)".to_string()); // unique Cherry match
    sheet.set_cell_src(1, 11, "=DGET(A1:E7, \"Profit\", J1:J2)".to_string()); // 2 Pear matches -> ambiguous
    sheet.set_cell_src(2, 10, "Tree".to_string());
    sheet.set_cell_src(3, 10, "Mango".to_string());
    sheet.set_cell_src(2, 11, "=DGET(A1:E7, \"Profit\", K3:K4)".to_string()); // no matches
    sheet.commit(None).unwrap();
    assert_float_close(&sheet.get_result_data(&CellRef::new(0, 11)), 105.0, 1e-9);
    assert!(matches!(
        sheet.get_result_data(&CellRef::new(1, 11)),
        ResultData::Error(ref e) if e.contains("#NUM!")
    ));
    assert!(matches!(
        sheet.get_result_data(&CellRef::new(2, 11)),
        ResultData::Error(ref e) if e.contains("#VALUE!")
    ));
}

#[test]
fn test_database_or_across_criteria_rows_and_field_by_index() {
    let mut sheet = database_test_sheet();
    // Pear OR Cherry, field selected by 1-based index (5 = Profit).
    sheet.set_cell_src(0, 11, "=DSUM(A1:E7, 5, J1:J3)".to_string());
    sheet.commit(None).unwrap();
    assert_float_close(&sheet.get_result_data(&CellRef::new(0, 11)), 277.8, 1e-6);
}

#[test]
fn test_database_blank_criteria_row_matches_every_record() {
    let mut sheet = database_test_sheet();
    sheet.set_cell_src(0, 11, "=DSTDEV(A1:E7, \"Profit\", K3:K4)".to_string());
    sheet.set_cell_src(1, 11, "=DVARP(A1:E7, \"Profit\", K3:K4)".to_string());
    sheet.commit(None).unwrap();
    assert_float_close(
        &sheet.get_result_data(&CellRef::new(0, 11)),
        23.149946004256645,
        1e-6,
    );
    assert_float_close(
        &sheet.get_result_data(&CellRef::new(1, 11)),
        446.59999999999934,
        1e-6,
    );
}

#[test]
fn test_database_aggregation_ignores_blank_and_boolean_range_values() {
    // Aggregating range arguments in DSUM/DCOUNT/DPRODUCT/DAVERAGE ignores
    // blanks and TRUE/FALSE range cells rather than coercing them to numeric 0/1.
    let grid: [[&str; 4]; 4] = [
        ["Key", "Val", "=DSUM(A1:B4, \"Val\", D1:D2)", "Key"],
        ["x", "10", "", "x"],
        ["x", "", "=DCOUNT(A1:B4, \"Val\", D1:D2)", ""],
        ["x", "TRUE", "=DPRODUCT(A1:B4, \"Val\", D1:D2)", ""],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    // Only the genuine number (10) should count/sum/multiply; the blank
    // and the boolean must be excluded, not treated as 0/1.
    assert_float_close(&sheet.get_result_data(&CellRef::new(0, 2)), 10.0, 1e-9);
    assert_float_close(&sheet.get_result_data(&CellRef::new(2, 2)), 1.0, 1e-9);
    assert_float_close(&sheet.get_result_data(&CellRef::new(3, 2)), 10.0, 1e-9);
}

#[test]
fn test_database_numeric_criteria_excludes_non_numeric_cells() {
    // Numeric criteria (e.g. ">"/"<") match genuine numbers only;
    // blank/text/boolean database cells do not match.
    let grid: [[&str; 4]; 6] = [
        ["Key", "Val", "=DCOUNT(A1:B6, \"Val\", D1:D2)", "Val"],
        ["x", "1", "", "<1000"],
        ["x", "", "", ""],     // blank -- must not match "<1000" as if it were 0
        ["x", "text", "", ""], // text -- must not match either
        ["x", "TRUE", "", ""], // boolean -- must not match either
        ["x", "-5", "", ""],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    // Only the two genuine numbers (1 and -5) should match and count.
    assert_float_close(&sheet.get_result_data(&CellRef::new(0, 2)), 2.0, 1e-9);
}
