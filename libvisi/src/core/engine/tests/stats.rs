use super::*;

#[test]
fn test_statistical_summary_functions() {
    let grid = [
        ["10", "20", "30", "40", "50"],
        [
            "=AVEDEV(A1:E1)",
            "=AVERAGEA(A1:E1)",
            "=MEDIAN(A1:E1)",
            "=GEOMEAN(A1:E1)",
            "=HARMEAN(A1:E1)",
        ],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    let r1 = sheet.get_result_data(&CellRef::new(1, 0));
    assert!(matches!(r1, ResultData::Float(v) if (v - 12.0).abs() < 1e-6));

    let r2 = sheet.get_result_data(&CellRef::new(1, 1));
    assert!(matches!(r2, ResultData::Float(v) if (v - 30.0).abs() < 1e-6));

    let r3 = sheet.get_result_data(&CellRef::new(1, 2));
    assert!(matches!(r3, ResultData::Float(v) if (v - 30.0).abs() < 1e-6));

    let r4 = sheet.get_result_data(&CellRef::new(1, 3));
    assert!(matches!(r4, ResultData::Float(v) if (v - 26.0517108).abs() < 1e-4));

    let r5 = sheet.get_result_data(&CellRef::new(1, 4));
    assert!(matches!(r5, ResultData::Float(v) if (v - 21.89781).abs() < 1e-4));
}

#[test]
fn test_variance_and_stdev() {
    let grid = [
        ["10", "20", "30", "40", "50"],
        [
            "=VAR.S(A1:E1)",
            "=VAR.P(A1:E1)",
            "=STDEV.S(A1:E1)",
            "=STDEV.P(A1:E1)",
            "=DEVSQ(A1:E1)",
        ],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    let r1 = sheet.get_result_data(&CellRef::new(1, 0));
    assert!(matches!(r1, ResultData::Float(v) if (v - 250.0).abs() < 1e-6));

    let r2 = sheet.get_result_data(&CellRef::new(1, 1));
    assert!(matches!(r2, ResultData::Float(v) if (v - 200.0).abs() < 1e-6));

    let r3 = sheet.get_result_data(&CellRef::new(1, 2));
    assert!(matches!(r3, ResultData::Float(v) if (v - 250.0_f64.sqrt()).abs() < 1e-6));

    let r4 = sheet.get_result_data(&CellRef::new(1, 3));
    assert!(matches!(r4, ResultData::Float(v) if (v - 200.0_f64.sqrt()).abs() < 1e-6));

    let r5 = sheet.get_result_data(&CellRef::new(1, 4));
    assert!(matches!(r5, ResultData::Float(v) if (v - 1000.0).abs() < 1e-6));
}

#[test]
fn test_criteria_and_quantiles() {
    let grid = [
        ["5", "10", "15", "20", "25"],
        [
            "=AVERAGEIF(A1:E1, \">10\")",
            "=LARGE(A1:E1, 2)",
            "=SMALL(A1:E1, 1)",
            "=PERCENTILE.INC(A1:E1, 0.5)",
            "=QUARTILE.INC(A1:E1, 3)",
        ],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    let r1 = sheet.get_result_data(&CellRef::new(1, 0));
    assert!(matches!(r1, ResultData::Float(v) if (v - 20.0).abs() < 1e-6));

    let r2 = sheet.get_result_data(&CellRef::new(1, 1));
    assert!(matches!(r2, ResultData::Float(v) if (v - 20.0).abs() < 1e-6));

    let r3 = sheet.get_result_data(&CellRef::new(1, 2));
    assert!(matches!(r3, ResultData::Float(v) if (v - 5.0).abs() < 1e-6));

    let r4 = sheet.get_result_data(&CellRef::new(1, 3));
    assert!(matches!(r4, ResultData::Float(v) if (v - 15.0).abs() < 1e-6));

    let r5 = sheet.get_result_data(&CellRef::new(1, 4));
    assert!(matches!(r5, ResultData::Float(v) if (v - 20.0).abs() < 1e-6));
}

#[test]
fn test_distributions_and_special_math() {
    let grid = [[
        "=NORM.S.DIST(0, TRUE)",
        "=NORM.S.INV(0.5)",
        "=EXPON.DIST(1, 1, TRUE)",
        "=POISSON.DIST(2, 2, FALSE)",
        "=STANDARDIZE(15, 10, 2.5)",
    ]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    let r1 = sheet.get_result_data(&CellRef::new(0, 0));
    assert!(matches!(r1, ResultData::Float(v) if (v - 0.5).abs() < 1e-6));

    let r2 = sheet.get_result_data(&CellRef::new(0, 1));
    assert!(matches!(r2, ResultData::Float(v) if v.abs() < 1e-6));

    let r3 = sheet.get_result_data(&CellRef::new(0, 2));
    assert!(matches!(r3, ResultData::Float(v) if (v - (1.0 - (-1.0_f64).exp())).abs() < 1e-6));

    let r4 = sheet.get_result_data(&CellRef::new(0, 3));
    assert!(matches!(r4, ResultData::Float(v) if (v - 0.270670566).abs() < 1e-5));

    let r5 = sheet.get_result_data(&CellRef::new(0, 4));
    assert!(matches!(r5, ResultData::Float(v) if (v - 2.0).abs() < 1e-6));
}

#[test]
fn test_regression_and_correlation() {
    let grid = [
        ["1", "2", "3", "4", "5"],
        ["2", "4", "6", "8", "10"],
        [
            "=CORREL(A1:E1, A2:E2)",
            "=SLOPE(A2:E2, A1:E1)",
            "=INTERCEPT(A2:E2, A1:E1)",
            "=FORECAST(6, A2:E2, A1:E1)",
            "=RSQ(A2:E2, A1:E1)",
        ],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    let r1 = sheet.get_result_data(&CellRef::new(2, 0));
    assert!(matches!(r1, ResultData::Float(v) if (v - 1.0).abs() < 1e-6));

    let r2 = sheet.get_result_data(&CellRef::new(2, 1));
    assert!(matches!(r2, ResultData::Float(v) if (v - 2.0).abs() < 1e-6));

    let r3 = sheet.get_result_data(&CellRef::new(2, 2));
    assert!(matches!(r3, ResultData::Float(v) if v.abs() < 1e-6));

    let r4 = sheet.get_result_data(&CellRef::new(2, 3));
    assert!(matches!(r4, ResultData::Float(v) if (v - 12.0).abs() < 1e-6));

    let r5 = sheet.get_result_data(&CellRef::new(2, 4));
    assert!(matches!(r5, ResultData::Float(v) if (v - 1.0).abs() < 1e-6));
}
