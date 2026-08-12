use super::*;

/// Evaluates a standalone formula (no cell grid needed) and returns its result.
fn eval1(source: &str) -> ResultData {
    let sheet = Sheet::new(SheetInit::default());
    sheet.eval(source, None).unwrap().0
}

fn assert_float_close(result: &ResultData, expected: f64, tol: f64) {
    match result {
        ResultData::Float(f) => assert!((f - expected).abs() < tol, "expected {expected}, got {f}"),
        ResultData::Integer(i) => assert!(
            (*i as f64 - expected).abs() < tol,
            "expected {expected}, got {i}"
        ),
        other => panic!("expected numeric result close to {expected}, got {other:?}"),
    }
}

#[test]
fn test_date_serial_matches_real_excel_reference_values() {
    // ymd_to_serial's Excel-epoch offset was off by one for every date
    // after the fictitious Feb 29, 1900 (serial 60) -- found via
    // differential fuzzing against real Excel while validating the new
    // bond-pricing functions below, which all depend on serial-date
    // arithmetic. These are well-known, independently verifiable
    // Excel serial numbers (e.g. 25569 = Unix epoch, a widely cited
    // Excel/Unix-timestamp conversion constant), confirmed again here
    // directly against real Microsoft Excel via the differential fuzzer.
    assert_float_close(&eval1("=DATE(1900,1,1)"), 1.0, 1e-9);
    assert_float_close(&eval1("=DATE(1970,1,1)"), 25569.0, 1e-9);
    assert_float_close(&eval1("=DATE(2000,1,1)"), 36526.0, 1e-9);
    assert_float_close(&eval1("=DATE(2007,11,21)"), 39407.0, 1e-9);
    assert_float_close(&eval1("=DATE(2021,1,1)"), 44197.0, 1e-9);
}

#[test]
fn test_choose_basic() {
    assert_float_close(&eval1("=CHOOSE(1, 10, 20, 30)"), 10.0, 1e-9);
    assert_float_close(&eval1("=CHOOSE(3, 10, 20, 30)"), 30.0, 1e-9);
    assert_eq!(
        eval1("=CHOOSE(2, \"a\", \"b\", \"c\")").to_string(),
        "b".to_string()
    );
}

#[test]
fn test_choose_out_of_range_errors() {
    assert!(
        matches!(eval1("=CHOOSE(0, 10, 20)"), ResultData::Error(ref e) if e.contains("#VALUE!"))
    );
    assert!(
        matches!(eval1("=CHOOSE(5, 10, 20)"), ResultData::Error(ref e) if e.contains("#VALUE!"))
    );
}

#[test]
fn test_choose_lazy_evaluation_skips_unselected_branch_errors() {
    // Real Excel does not evaluate unselected CHOOSE branches; NA() in the
    // unselected branch must not surface.
    let result = eval1("=CHOOSE(1, 42, NA())");
    assert_float_close(&result, 42.0, 1e-9);
}

#[test]
fn test_yearfrac_basis1_uses_actual_actual_year_average() {
    // Confirmed against real Excel via the differential fuzzer: basis 1
    // averages 365/366 across every calendar year the span touches, not
    // the average Julian year (365.2425) the previous implementation used.
    assert_float_close(
        &eval1("=YEARFRAC(DATE(1998,8,8),DATE(1998,9,8),1)"),
        0.08493150684931507,
        1e-9,
    );
    assert_float_close(
        &eval1("=YEARFRAC(DATE(2016,1,1),DATE(2016,6,1),1)"),
        0.41530054644808745,
        1e-9,
    );
    assert_float_close(
        &eval1("=YEARFRAC(DATE(2017,6,1),DATE(2020,9,1),1)"),
        3.252566735112936,
        1e-6,
    );
}

// --- Day-count / bond-pricing financial functions -----------------------
//
// Every expected value below was confirmed directly against real
// Microsoft Excel via the differential fuzzer (fuzz/fuzz_excel.py),
// either from a Microsoft-documented example or a fuzzer-found input.

#[test]
fn test_coupon_date_functions_match_microsoft_docs_example() {
    // Microsoft's own COUPDAYS/COUPDAYBS/COUPNUM/COUPPCD documentation
    // example: settlement 1/25/2011, maturity 11/15/2011, semiannual,
    // actual/actual.
    assert_float_close(
        &eval1("=COUPDAYBS(DATE(2011,1,25), DATE(2011,11,15), 2, 1)"),
        71.0,
        1e-9,
    );
    assert_float_close(
        &eval1("=COUPDAYS(DATE(2011,1,25), DATE(2011,11,15), 2, 1)"),
        181.0,
        1e-9,
    );
    assert_float_close(
        &eval1("=COUPNUM(DATE(2011,1,25), DATE(2011,11,15), 2)"),
        2.0,
        1e-9,
    );
    assert_float_close(
        &eval1("=COUPPCD(DATE(2011,1,25), DATE(2011,11,15), 2)"),
        40497.0,
        1e-9,
    );
}

#[test]
fn test_coupncd_handles_day_of_month_clamping_correctly() {
    // Regression for a real bug found via the differential fuzzer: walking
    // a coupon schedule by chaining EDATE calls let day-31 clamping in a
    // short month (e.g. April) permanently overwrite the day-of-month for
    // every later step. COUPNCD/COUPPCD/COUPNUM now re-derive each quasi-
    // coupon date from the maturity anchor instead of chaining.
    assert_float_close(
        &eval1("=COUPNCD(DATE(2007,11,21), EDATE(DATE(2007,11,21),30), 4)"),
        39499.0,
        1e-9,
    );
}

#[test]
fn test_coupdaysnc_uses_real_calendar_days_not_coupdays_minus_coupdaybs() {
    // COUPDAYSNC applies the same real day-count convention as COUPDAYBS
    // directly to (settlement, next-coupon) -- it is not simply
    // COUPDAYS - COUPDAYBS, since COUPDAYS is an idealized period length
    // on bases 0/2/3/4 that generally doesn't equal the period's actual
    // calendar length.
    assert_float_close(
        &eval1("=COUPDAYSNC(DATE(2030,7,25), EDATE(DATE(2030,7,25),48), 2, 3)"),
        184.0,
        1e-9,
    );
}

#[test]
fn test_price_and_yield_are_inverses_and_match_real_excel() {
    assert_float_close(
        &eval1("=PRICE(DATE(1997,5,13), EDATE(DATE(1997,5,13),42), 0.0606, 0.0885, 100, 2, 3)"),
        91.75707270015924,
        1e-6,
    );
    assert_float_close(
        &eval1("=PRICE(DATE(2015,3,15), EDATE(DATE(2015,3,15),72), 0.0248, 0.0816, 100, 1, 2)"),
        73.86901031390656,
        1e-6,
    );
    assert_float_close(
        &eval1("=YIELD(DATE(2032,10,16), EDATE(DATE(2032,10,16),51), 0.0846, 103.71, 105, 4, 3)"),
        0.0840390355112947,
        1e-9,
    );
}

#[test]
fn test_duration_matches_real_excel() {
    assert_float_close(
        &eval1("=DURATION(DATE(2008,1,1), DATE(2016,1,1), 0.08, 0.09, 2, 1)"),
        5.993774955545186,
        1e-6,
    );
}

#[test]
fn test_disc_pricedisc_yielddisc_match_real_excel() {
    assert_float_close(
        &eval1("=DISC(DATE(1998,4,12), EDATE(DATE(1998,4,12),4), 93.41, 100, 1)"),
        0.197159836065574,
        1e-9,
    );
    assert_float_close(
        &eval1("=PRICEDISC(DATE(1996,2,2), EDATE(DATE(1996,2,2),14), 0.044, 100, 1)"),
        94.88372093023256,
        1e-6,
    );
}

#[test]
fn test_disc_basis1_year_length_has_two_regimes() {
    // Confirmed against real Excel via the differential fuzzer:
    // basis-1 Y has two regimes. For a span of at most 366 days --
    // including one that crosses a calendar-year boundary, like
    // Dec -> Mar -- Y is simply whether the *later* date's own calendar
    // year is leap (not a blend of both years' lengths). Only once the
    // span genuinely covers multiple full calendar years does it become
    // the average of 365/366 across every year touched.
    assert_float_close(
        &eval1("=DISC(DATE(2016,6,1), DATE(2016,9,1), 93, 100, 1)"),
        0.278478260869565,
        1e-9,
    );
    assert_float_close(
        &eval1("=PRICEDISC(DATE(2027,12,26), EDATE(DATE(2027,12,26),3), 0.0899, 100, 1)"),
        97.76478142076503,
        1e-6,
    );
    assert_float_close(
        &eval1("=DISC(DATE(2017,6,1), DATE(2020,9,1), 70, 100, 1)"),
        0.0922348484848485,
        1e-9,
    );
}

#[test]
fn test_pricemat_yieldmat_basis1_uses_issue_to_settlement_span() {
    // Unlike DISC's settlement-to-maturity span, PRICEMAT/YIELDMAT's
    // basis-1 year length is based on the (issue, settlement) span, not
    // the full (often multi-year) issue-to-maturity DIM span -- confirmed
    // against real Excel across two cases whose issue and settlement
    // years' leap status disagree with each other.
    assert_float_close(
        &eval1(
            "=YIELDMAT(EDATE(DATE(2012,9,3),2), EDATE(DATE(2012,9,3),25), DATE(2012,9,3), 0.079, 119.55, 1)",
        ),
        -0.019331059183910253,
        1e-9,
    );
    assert_float_close(
        &eval1(
            "=YIELDMAT(EDATE(DATE(1995,9,4),6), EDATE(DATE(1995,9,4),11), DATE(1995,9,4), 0.0226, 95.27, 1)",
        ),
        0.14082750571984862,
        1e-9,
    );
}

#[test]
fn test_received_and_intrate_match_real_excel() {
    assert_float_close(
        &eval1("=RECEIVED(DATE(2008,2,15), DATE(2008,5,15), 1000000, 0.0575, 2)"),
        1014584.6544071021,
        1e-3,
    );
    assert_float_close(
        &eval1("=INTRATE(DATE(2020,9,4), EDATE(DATE(2020,9,4),13), 24974.48, 5441.6, 1)"),
        -0.7237025672261941,
        1e-9,
    );
}

#[test]
fn test_tbill_functions_match_real_excel() {
    assert_float_close(
        &eval1("=TBILLPRICE(DATE(2008,3,31), DATE(2008,6,1), 0.09)"),
        98.45,
        1e-9,
    );
    assert_float_close(
        &eval1("=TBILLYIELD(DATE(2008,3,31), DATE(2008,6,1), 98.45)"),
        0.09141696292534264,
        1e-9,
    );
    assert_float_close(
        &eval1("=TBILLEQ(DATE(2008,3,31), DATE(2008,6,1), 0.0914)"),
        0.09415149356594302,
        1e-9,
    );
}

#[test]
fn test_accrint_totals_from_issue_regardless_of_calc_method() {
    // Confirmed against real Excel via the differential fuzzer across
    // regular, odd-first-period, and multi-period cases: calc_method
    // (TRUE vs FALSE) never changes ACCRINT's result in practice, so both
    // must total the same accrued-since-issue amount.
    assert_float_close(
        &eval1(
            "=ACCRINT(DATE(2017,6,7), DATE(2018,6,7), DATE(2018,7,7), 0.0547, 14735.46, 1, 0, TRUE)",
        ),
        873.1988004999998,
        1e-6,
    );
    assert_float_close(
        &eval1(
            "=ACCRINT(DATE(2017,6,7), DATE(2018,6,7), DATE(2018,7,7), 0.0547, 14735.46, 1, 0, FALSE)",
        ),
        873.1988004999998,
        1e-6,
    );
}

#[test]
fn test_accrintm_matches_real_excel() {
    assert_float_close(
        &eval1("=ACCRINTM(DATE(1998,8,8), EDATE(DATE(1998,8,8),1), 0.096, 37328.54, 1)"),
        304.3554384657534,
        1e-6,
    );
}

#[test]
fn test_amorlinc_amordegrc_reject_basis_2() {
    // Confirmed against real Excel: unlike every other function in
    // finance.rs, AMORLINC/AMORDEGRC reject basis 2 (actual/360).
    assert!(matches!(
        eval1("=AMORLINC(17737.01, DATE(2026,5,16), EDATE(DATE(2026,5,16),9), 5082.98, 1, 0.5, 2)"),
        ResultData::Error(ref e) if e.contains("#NUM!")
    ));
    assert!(matches!(
        eval1("=AMORDEGRC(8000.89, DATE(1995,11,26), EDATE(DATE(1995,11,26),9), 1484.45, 12, 0.05, 2)"),
        ResultData::Error(ref e) if e.contains("#NUM!")
    ));
}

#[test]
fn test_amordegrc_rejects_life_at_or_below_two_years() {
    // Confirmed against real Excel: the threshold is exactly life > 2
    // (rate < 0.5), not life >= 3 where the next coefficient bracket
    // starts -- life == 2 (rate == 0.5) is already rejected.
    assert!(matches!(
        eval1("=AMORDEGRC(9832.03, DATE(2024,9,7), EDATE(DATE(2024,9,7),5), 2414.1, 0, 0.5, 3)"),
        ResultData::Error(ref e) if e.contains("#NUM!")
    ));
    assert_float_close(
        &eval1("=AMORDEGRC(9832.03, DATE(2024,9,7), EDATE(DATE(2024,9,7),5), 2414.1, 0, 0.499, 3)"),
        3085.0,
        1e-6,
    );
}

#[test]
fn test_amordegrc_period_sequence_matches_real_excel() {
    // A full period-by-period depreciation schedule confirmed against real
    // Excel, including the first-period prorate and final-period taper.
    let expected = [
        5699.0, 8515.0, 6742.0, 5338.0, 4226.0, 3346.0, 2649.0, 2098.0, 1661.0, 1315.0,
    ];
    for (period, exp) in expected.iter().enumerate() {
        let f = format!(
            "=AMORDEGRC(46587.76, DATE(2028,7,27), EDATE(DATE(2028,7,27),7), 3292.26, {period}, 0.0833, 1)"
        );
        assert_float_close(&eval1(&f), *exp, 1e-6);
    }
}

#[test]
fn test_oddfprice_oddfyield_match_real_excel() {
    assert_float_close(
        &eval1(
            "=ODDFPRICE((DATE(2012,7,19)+26), EDATE((DATE(2012,7,19)+60),24), DATE(2012,7,19), (DATE(2012,7,19)+60), 0.082, 0.0923, 100, 4, 3)",
        ),
        98.06021551292406,
        1e-6,
    );
    // Basis 2: ODDFPRICE keeps COUPDAYS's idealized 360/365-per-freq value
    // for E on every basis except 1 (unlike ODDLPRICE/ODDLYIELD below).
    assert_float_close(
        &eval1(
            "=ODDFPRICE(DATE(2012,7,19)+26, EDATE(DATE(2012,7,19)+60,24), DATE(2012,7,19), DATE(2012,7,19)+60, 0.082, 0.0923, 100, 4, 2)",
        ),
        98.05901871412289,
        1e-6,
    );
    assert_float_close(
        &eval1(
            "=ODDFYIELD((DATE(2008,2,16)+20), EDATE((DATE(2008,2,16)+25),12), DATE(2008,2,16), (DATE(2008,2,16)+25), 0.0923, 117.26, 100, 4, 1)",
        ),
        -0.07045544328557858,
        1e-6,
    );
}

#[test]
fn test_oddlprice_oddlyield_match_real_excel() {
    assert_float_close(
        &eval1(
            "=ODDLYIELD((DATE(2029,6,23)+28), (DATE(2029,6,23)+44), DATE(2029,6,23), 0.0404, 115.8, 100, 4, 2)",
        ),
        -3.0950656625987194,
        1e-6,
    );
    assert_float_close(
        &eval1(
            "=ODDLYIELD((DATE(2002,7,10)+45), (DATE(2002,7,10)+114), DATE(2002,7,10), 0.0775, 96.4, 100, 1, 2)",
        ),
        0.2752128427867764,
        1e-6,
    );
    // Basis 3: unlike ODDFPRICE/ODDFYIELD, ODDLPRICE/ODDLYIELD use the
    // *actual* adjacent-period length for E on every basis, including 3.
    assert_float_close(
        &eval1(
            "=ODDLYIELD((DATE(2004,7,25)+51), (DATE(2004,7,25)+57), DATE(2004,7,25), 0.0682, 80.65, 100, 2, 3)",
        ),
        14.62856320740453,
        1e-6,
    );
    // Basis 1 and a freq=2/basis=2 case: both regression-test E being
    // anchored at the regular period *following* last_interest, not the
    // one *preceding* maturity -- an earlier version used the maturity
    // anchor, which only coincidentally matched when the two periods
    // happened to have the same calendar length.
    assert_float_close(
        &eval1(
            "=ODDLYIELD((DATE(2001,3,3)+10), (DATE(2001,3,3)+123), DATE(2001,3,3), 0.0546, 93.74, 100, 4, 1)",
        ),
        0.27529020678767524,
        1e-6,
    );
    assert_float_close(
        &eval1(
            "=ODDLPRICE(DATE(2019,6,21)+65, DATE(2019,6,21)+90, DATE(2019,6,21), 0.0478, 0.0376, 100, 2, 2)",
        ),
        100.06731898218347,
        1e-6,
    );
}

// EUROCONVERT is the one function in this batch NOT verified against real
// Excel: it requires the "Euro Currency Tools" add-in, which returns
// #NAME? in this environment's Excel regardless of arguments (confirmed
// by direct check). These expected values are computed by hand from
// Microsoft's published fixed euro-conversion rates and rounding rule.

#[test]
fn test_euroconvert_direct_and_reverse() {
    assert_float_close(&eval1("=EUROCONVERT(100, \"EUR\", \"DEM\")"), 195.58, 1e-9);
    assert_float_close(&eval1("=EUROCONVERT(100, \"DEM\", \"EUR\")"), 51.13, 1e-9);
}

#[test]
fn test_euroconvert_triangulates_through_eur() {
    assert_float_close(&eval1("=EUROCONVERT(1, \"DEM\", \"FRF\")"), 3.35, 1e-9);
    assert_float_close(
        &eval1("=EUROCONVERT(1, \"DEM\", \"FRF\", TRUE)"),
        3.353854885138279,
        1e-9,
    );
}

#[test]
fn test_euroconvert_same_currency_is_a_no_op() {
    assert_float_close(&eval1("=EUROCONVERT(42, \"EUR\", \"EUR\")"), 42.0, 1e-9);
}

#[test]
fn test_euroconvert_rounds_zero_decimal_currencies_to_whole_units() {
    // ITL/ESP/BEF/LUF had no meaningful subunit in everyday use.
    assert_float_close(
        &eval1("=EUROCONVERT(100, \"EUR\", \"ITL\")"),
        193627.0,
        1e-9,
    );
}

#[test]
fn test_euroconvert_rejects_unknown_currency_code() {
    assert!(matches!(
        eval1("=EUROCONVERT(100, \"EUR\", \"USD\")"),
        ResultData::Error(ref e) if e.contains("#VALUE!")
    ));
}

#[test]
fn test_euroconvert_rejects_triangulation_precision_below_3() {
    assert!(matches!(
        eval1("=EUROCONVERT(1, \"DEM\", \"FRF\", FALSE, 2)"),
        ResultData::Error(ref e) if e.contains("#NUM!")
    ));
}

#[test]
fn test_odd_period_functions_reject_settlement_at_or_before_anchor() {
    // Confirmed against real Excel via the differential fuzzer: settlement
    // must be strictly after issue (ODDFPRICE/ODDFYIELD) or last_interest
    // (ODDLPRICE/ODDLYIELD) -- settlement == issue/last_interest is #NUM!,
    // not a zero-length odd period.
    assert!(matches!(
        eval1("=ODDFPRICE(DATE(2030,4,13)+0, EDATE(DATE(2030,4,13)+44,84), DATE(2030,4,13), DATE(2030,4,13)+44, 0.073, 0.0368, 100, 1, 3)"),
        ResultData::Error(ref e) if e.contains("#NUM!")
    ));
    assert!(matches!(
        eval1("=ODDFYIELD(DATE(2023,11,11)+0, EDATE(DATE(2023,11,11)+20,12), DATE(2023,11,11), DATE(2023,11,11)+20, 0.0562, 89.35, 105, 4, 3)"),
        ResultData::Error(ref e) if e.contains("#NUM!")
    ));
    assert!(matches!(
        eval1("=ODDLPRICE(DATE(2005,6,28)+0, DATE(2005,6,28)+58, DATE(2005,6,28), 0.0353, 0.0938, 105, 4, 4)"),
        ResultData::Error(ref e) if e.contains("#NUM!")
    ));
    assert!(matches!(
        eval1("=ODDLYIELD(DATE(2010,7,11)+0, DATE(2010,7,11)+15, DATE(2010,7,11), 0.028, 110.37, 100, 2, 1)"),
        ResultData::Error(ref e) if e.contains("#NUM!")
    ));
}

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
    // Regression for a real bug found via the differential fuzzer:
    // aggregating with `to_f64` (which coerces blank -> 0 and booleans ->
    // 1/0 for scalar arithmetic) let a single blank matched row zero out
    // DPRODUCT entirely, and skewed DCOUNT/DSUM/DAVERAGE by counting
    // blanks and TRUE/FALSE as numeric 0/1 -- Excel ignores both within a
    // range argument, same as SUM/COUNT/AVERAGE do.
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
    // Regression for a real bug found via the differential fuzzer: a
    // ">"/"<" criterion was matching blank, text, and boolean database
    // cells by silently comparing them as if they were 0 (`to_f64`'s
    // scalar-arithmetic default). Excel only ever matches genuine numbers
    // against a numeric criterion -- blank/text/boolean cells must fail
    // it outright, the same way they're excluded from range aggregation.
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

// --- LAMBDA family (LAMBDA, MAP, REDUCE, SCAN, BYROW, BYCOL, MAKEARRAY,
// ISOMITTED) --------------------------------------------------------------
//
// MAP and REDUCE results were confirmed against real Microsoft Excel.
// BYROW/BYCOL/MAKEARRAY, a bare uninvoked LAMBDA, and even SCAN could not
// be reliably confirmed that way: any dynamic-array-spilling formula --
// even a bare `=SEQUENCE(3)` with no LAMBDA involved at all -- breaks this
// environment's Excel AppleScript automation bridge intermittently
// (confirmed directly, and not specific to any one function; the same
// formula sometimes succeeds and sometimes doesn't across repeat runs).
// Their expected values below are independently verified instead, either
// by hand-calculated arithmetic or against Microsoft's own documented
// SCAN example (SCAN(0, {1,2,3}, LAMBDA(a,v,a+v)) => {1,3,6}).

#[test]
fn test_lambda_bare_is_uncallable() {
    // The parser has no `(expr)(args)` immediate-invocation syntax (that
    // would require calling an arbitrary sub-expression, not just a bare
    // identifier), so an uninvoked, unnamed LAMBDA can't produce a value,
    // matching Excel's own #CALC! for this case.
    assert!(matches!(
        eval1("=LAMBDA(x, x*2)"),
        ResultData::Error(ref e) if e.contains("#CALC!")
    ));
}

#[test]
fn test_isomitted_best_effort() {
    // Best-effort implementation (see the doc comment in evaluate_function
    // above `ISOMITTED`'s dispatch): every lambda invocation path here
    // always supplies exactly as many values as declared parameters, so a
    // declared, in-scope parameter is never actually omitted -- this only
    // exercises the "identifier not found in scope at all" case.
    assert!(matches!(
        eval1("=INDEX(MAP(1, LAMBDA(x, ISOMITTED(x))), 1)"),
        ResultData::Boolean(false)
    ));
    assert!(matches!(
        eval1("=ISOMITTED(some_undeclared_name)"),
        ResultData::Boolean(true)
    ));
}

#[test]
fn test_map_single_and_multiple_arrays_match_real_excel() {
    let grid: [[&str; 3]; 3] = [
        ["10", "1", "=INDEX(MAP(A1:A3, LAMBDA(x, x*2)), 2)"],
        ["20", "2", "=INDEX(MAP(A1:A3, B1:B3, LAMBDA(x,y, x+y)), 3)"],
        ["30", "3", ""],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    assert_float_close(&sheet.get_result_data(&CellRef::new(0, 2)), 40.0, 1e-9); // 20*2
    assert_float_close(&sheet.get_result_data(&CellRef::new(1, 2)), 33.0, 1e-9); // 30+3
}

#[test]
fn test_reduce_and_scan_match_real_excel() {
    let grid: [[&str; 2]; 3] = [
        ["1", "=REDUCE(0, A1:A3, LAMBDA(acc,v, acc+v))"],
        ["2", "=INDEX(SCAN(0, A1:A3, LAMBDA(acc,v, acc+v)), 3)"],
        ["3", ""],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    assert_float_close(&sheet.get_result_data(&CellRef::new(0, 1)), 6.0, 1e-9);
    assert_float_close(&sheet.get_result_data(&CellRef::new(1, 1)), 6.0, 1e-9);
}

#[test]
fn test_reduce_two_arg_form_seeds_from_first_element() {
    // Real Excel's REDUCE always takes 3 arguments, but initial_value is
    // documented as optional; since this parser has no syntax to express
    // "omitted argument" (no leading/trailing empty comma support), a
    // plain 2-argument REDUCE(array, lambda) is accepted as that
    // omitted-initial-value form: the array's own first element seeds the
    // accumulator, and the rest are folded in.
    let grid: [[&str; 2]; 3] = [
        ["1", "=REDUCE(A1:A3, LAMBDA(acc,v, acc+v))"],
        ["2", ""],
        ["3", ""],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    assert_float_close(&sheet.get_result_data(&CellRef::new(0, 1)), 6.0, 1e-9);
}

#[test]
fn test_byrow_and_bycol_sum_hand_verified() {
    // Data: [[1,2,3],[4,5,6],[7,8,9]]. Row sums: 6, 15, 24.
    // Column sums: 12, 15, 18. Not verifiable against real Excel here (see
    // module doc comment above) -- hand-verified arithmetic instead.
    let grid: [[&str; 5]; 3] = [
        ["1", "2", "3", "", ""],
        ["4", "5", "6", "", ""],
        ["7", "8", "9", "", ""],
    ];
    let mut sheet = create_sheet(&grid);
    for i in 0..3 {
        sheet.set_cell_src(
            i,
            3,
            format!("=INDEX(BYROW(A1:C3, LAMBDA(zz, SUM(zz))), {})", i + 1),
        );
        sheet.set_cell_src(
            i,
            4,
            format!("=INDEX(BYCOL(A1:C3, LAMBDA(zz, SUM(zz))), {})", i + 1),
        );
    }
    sheet.commit(None).unwrap();
    for (i, expected_row) in [6.0, 15.0, 24.0].iter().enumerate() {
        assert_float_close(
            &sheet.get_result_data(&CellRef::new(i, 3)),
            *expected_row,
            1e-9,
        );
    }
    for (i, expected_col) in [12.0, 15.0, 18.0].iter().enumerate() {
        assert_float_close(
            &sheet.get_result_data(&CellRef::new(i, 4)),
            *expected_col,
            1e-9,
        );
    }
}

#[test]
fn test_makearray_builds_row_major_flat_array_hand_verified() {
    // MAKEARRAY(2, 3, LAMBDA(r,c, r*10+c)) should build
    // [[11,12,13],[21,22,23]] flattened row-major: [11,12,13,21,22,23].
    let grid: [[&str; 1]; 6] = [
        ["=INDEX(MAKEARRAY(2, 3, LAMBDA(rr,cc, rr*10+cc)), 1)"],
        ["=INDEX(MAKEARRAY(2, 3, LAMBDA(rr,cc, rr*10+cc)), 2)"],
        ["=INDEX(MAKEARRAY(2, 3, LAMBDA(rr,cc, rr*10+cc)), 3)"],
        ["=INDEX(MAKEARRAY(2, 3, LAMBDA(rr,cc, rr*10+cc)), 4)"],
        ["=INDEX(MAKEARRAY(2, 3, LAMBDA(rr,cc, rr*10+cc)), 5)"],
        ["=INDEX(MAKEARRAY(2, 3, LAMBDA(rr,cc, rr*10+cc)), 6)"],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    for (i, expected) in [11.0, 12.0, 13.0, 21.0, 22.0, 23.0].iter().enumerate() {
        assert_float_close(&sheet.get_result_data(&CellRef::new(i, 0)), *expected, 1e-9);
    }
}

#[test]
fn test_index_two_arg_form_is_one_based() {
    // Regression for a real bug found while testing MAP/BYROW: the 2-arg
    // INDEX(array, n) form returned the element one *past* the requested
    // 1-based position (INDEX(list,1) returned list's 2nd element), while
    // the 3-arg row/col form was already correctly 1-based. The
    // standalone INDEX fuzz generator only ever used the 3-arg form, so
    // this never surfaced until a real 2-arg call (from MAP's own
    // implementation) exercised it.
    let grid: [[&str; 2]; 3] = [
        ["10", "=INDEX(A1:A3, 1)"],
        ["20", "=INDEX(A1:A3, 2)"],
        ["30", "=INDEX(A1:A3, 3)"],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    assert_float_close(&sheet.get_result_data(&CellRef::new(0, 1)), 10.0, 1e-9);
    assert_float_close(&sheet.get_result_data(&CellRef::new(1, 1)), 20.0, 1e-9);
    assert_float_close(&sheet.get_result_data(&CellRef::new(2, 1)), 30.0, 1e-9);
}

// --- Range/workbook metadata introspection --------------------------------
//
// Every expected value below was confirmed against real Microsoft Excel
// (FORMULATEXT/ISFORMULA/SHEETS needed an `_xlfn.` prefix to be recognized
// at all when written by openpyxl -- confirmed as the actual cause of a
// real #NAME? mismatch, not a bug in this implementation).

#[test]
fn test_row_and_column_return_array_for_multi_row_or_col_range() {
    // Regression for a real bug found via the differential fuzzer:
    // ROW/COLUMN against a multi-row/multi-column reference must return
    // an array (one entry per row/column spanned), not just the first
    // position -- `=SUM(ROW(A1:A5))` is 1+2+3+4+5=15, not 1.
    let grid: [[&str; 2]; 5] = [
        ["10", "=SUM(ROW(A1:A5))"],
        ["20", "=INDEX(ROW(A1:A5), 3)"],
        ["30", "=COLUMN(A1)"],
        ["40", ""],
        ["50", ""],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    assert_float_close(&sheet.get_result_data(&CellRef::new(0, 1)), 15.0, 1e-9);
    assert_float_close(&sheet.get_result_data(&CellRef::new(1, 1)), 3.0, 1e-9);
    assert_float_close(&sheet.get_result_data(&CellRef::new(2, 1)), 1.0, 1e-9);
}

#[test]
fn test_rows_columns_areas_match_real_excel() {
    let grid: [[&str; 5]; 3] = [
        ["1", "2", "3", "=ROWS(A1:C3)", "=COLUMNS(A1:C3)"],
        ["4", "5", "6", "=AREAS(A1:C3)", ""],
        ["7", "8", "9", "", ""],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    assert_float_close(&sheet.get_result_data(&CellRef::new(0, 3)), 3.0, 1e-9);
    assert_float_close(&sheet.get_result_data(&CellRef::new(0, 4)), 3.0, 1e-9);
    assert_float_close(&sheet.get_result_data(&CellRef::new(1, 3)), 1.0, 1e-9);
}

#[test]
fn test_isref_distinguishes_references_from_values() {
    assert!(matches!(eval1("=ISREF(A1)"), ResultData::Boolean(true)));
    assert!(matches!(eval1("=ISREF(5)"), ResultData::Boolean(false)));
}

#[test]
fn test_formulatext_and_isformula() {
    let grid: [[&str; 3]; 1] = [["10", "=A1*2", "=FORMULATEXT(B1)"]];
    let mut sheet = create_sheet(&grid);
    sheet.set_cell_src(0, 2, "=ISFORMULA(B1)".to_string());
    sheet.commit(None).unwrap();
    assert!(matches!(
        sheet.get_result_data(&CellRef::new(0, 2)),
        ResultData::Boolean(true)
    ));
    let grid2: [[&str; 2]; 1] = [["10", "=FORMULATEXT(A1)"]];
    let mut sheet2 = create_sheet(&grid2);
    sheet2.commit(None).unwrap();
    assert!(matches!(
        sheet2.get_result_data(&CellRef::new(0, 1)),
        ResultData::Error(ref e) if e.contains("#N/A")
    ));
}

#[test]
fn test_hyperlink_returns_friendly_name_or_link() {
    assert_eq!(
        eval1("=HYPERLINK(\"https://example.com\")").to_string(),
        "https://example.com"
    );
    assert_eq!(
        eval1("=HYPERLINK(\"https://example.com\", \"Click\")").to_string(),
        "Click"
    );
}

#[test]
fn test_sheets_counts_context_sheets() {
    assert!(matches!(eval1("=SHEETS()"), ResultData::Float(f) if f == 1.0));
}

#[test]
fn test_sheet_reports_real_workbook_ordinal() {
    // Regression for #26: SHEET() always returned 1 regardless of true
    // position, since `Context.sheets` is an unordered `HashMap` and
    // nothing threaded real order down from `WorkbookManager::sheets` (a
    // `Vec`, where order is genuine) into `Sheet::evaluate_function`.
    let table1 = Sheet::new(SheetInit {
        id: None,
        name: Some("table_1".to_string()),
        rows: 2,
        cols: 2,
    });
    let table2 = Sheet::new(SheetInit {
        id: None,
        name: Some("table_2".to_string()),
        rows: 2,
        cols: 2,
    });
    let table3 = Sheet::new(SheetInit {
        id: None,
        name: Some("table_3".to_string()),
        rows: 2,
        cols: 2,
    });
    let sheets = [table1, table2, table3];

    let mut context = Context::new();
    for sheet in &sheets {
        context.add_table(sheet.name.clone(), sheet);
    }
    context.sheet_order = sheets.iter().map(|s| s.name.clone()).collect();

    let (r1, _) = sheets[0].eval("=SHEET()", Some(&context)).unwrap();
    assert_float_close(&r1, 1.0, 1e-9);
    let (r2, _) = sheets[1].eval("=SHEET()", Some(&context)).unwrap();
    assert_float_close(&r2, 2.0, 1e-9);
    let (r3, _) = sheets[2].eval("=SHEET()", Some(&context)).unwrap();
    assert_float_close(&r3, 3.0, 1e-9);

    // A reference into another sheet reports *that* sheet's ordinal, not
    // the formula's own.
    let (r4, _) = sheets[0]
        .eval("=SHEET(table_3!A1)", Some(&context))
        .unwrap();
    assert_float_close(&r4, 3.0, 1e-9);

    // A plain text sheet name is also accepted, same as real Excel.
    let (r5, _) = sheets[0]
        .eval("=SHEET(\"table_2\")", Some(&context))
        .unwrap();
    assert_float_close(&r5, 2.0, 1e-9);

    // No context at all (standalone eval outside a WorkbookManager pass)
    // keeps the old documented fallback of 1.
    assert!(matches!(eval1("=SHEET()"), ResultData::Float(f) if f == 1.0));
}

#[test]
fn test_indirect_resolves_cell_and_range_text() {
    let grid: [[&str; 2]; 3] = [
        ["10", "=INDIRECT(\"A1\")"],
        ["20", "=SUM(INDIRECT(\"A1:A3\"))"],
        ["30", ""],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    assert_float_close(&sheet.get_result_data(&CellRef::new(0, 1)), 10.0, 1e-9);
    assert_float_close(&sheet.get_result_data(&CellRef::new(1, 1)), 60.0, 1e-9);
}

#[test]
fn test_offset_shifts_and_resizes_reference() {
    let grid: [[&str; 2]; 3] = [
        ["10", "=OFFSET(A1, 1, 0)"],
        ["20", "=SUM(OFFSET(A1, 0, 0, 3, 1))"],
        ["30", ""],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    assert_float_close(&sheet.get_result_data(&CellRef::new(0, 1)), 20.0, 1e-9);
    assert_float_close(&sheet.get_result_data(&CellRef::new(1, 1)), 60.0, 1e-9);
}

#[test]
fn test_cell_info_subset() {
    let grid: [[&str; 2]; 5] = [
        ["10", "=CELL(\"row\", A3)"],
        ["20", "=CELL(\"col\", A3)"],
        ["30", "=CELL(\"address\", A3)"],
        ["40", "=CELL(\"contents\", A1)"],
        ["50", ""],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    assert_float_close(&sheet.get_result_data(&CellRef::new(0, 1)), 3.0, 1e-9);
    assert_float_close(&sheet.get_result_data(&CellRef::new(1, 1)), 1.0, 1e-9);
    assert_eq!(
        sheet.get_result_data(&CellRef::new(2, 1)).to_string(),
        "$A$3"
    );
    assert_float_close(&sheet.get_result_data(&CellRef::new(3, 1)), 10.0, 1e-9);
}

// --- Dynamic array reshaping/lookup batch ---
//
// Each test below builds its 2D input with `SEQUENCE(rows, cols)` so it
// doesn't need a cell grid, then pulls values back out with `INDEX`/`SUM`.
// HSTACK, VSTACK, UNIQUE, SORT, XMATCH, and FILTER (with a genuine boolean
// helper range rather than a broadcast comparison -- this engine's
// comparison operators don't broadcast across ranges, a separate,
// pre-existing, out-of-scope limitation) were confirmed byte-for-byte
// against real Microsoft Excel via the differential fuzzer. WRAPROWS,
// WRAPCOLS, CHOOSEROWS, CHOOSECOLS, DROP, TAKE, EXPAND, TOCOL, TOROW,
// SORTBY, TRIMRANGE, and LOOKUP (on genuinely sorted input) were likewise
// confirmed against real Excel. TRANSPOSE could not be verified against
// real Excel: every authoring variant tried (bare, `_xlfn.`,
// `_xlfn._xlws.`, standalone, and nested inside SUM/INDEX) gave `#VALUE!`
// in real Excel when the formula was written by openpyxl rather than
// Excel itself, even though TRANSPOSE predates dynamic arrays entirely --
// this points at an openpyxl authoring limitation (TRANSPOSE has always
// required the legacy CSE `t="array"` formula flag, which openpyxl's plain
// string assignment doesn't produce) rather than a bug in visi's TRANSPOSE
// logic, which is hand-verified correct below.

#[test]
fn test_transpose_swaps_rows_and_cols() {
    // SEQUENCE(2,3) = [[1,2,3],[4,5,6]]; transposed = [[1,4],[2,5],[3,6]].
    assert_float_close(&eval1("=INDEX(TRANSPOSE(SEQUENCE(2,3)),3,1)"), 3.0, 1e-9);
    assert_float_close(&eval1("=INDEX(TRANSPOSE(SEQUENCE(2,3)),1,2)"), 4.0, 1e-9);
    assert_float_close(&eval1("=SUM(TRANSPOSE(SEQUENCE(2,3)))"), 21.0, 1e-9);
}

#[test]
fn test_index_recovers_shape_of_nested_reshape_function() {
    // Regression test: INDEX's 3-arg row/col form used to only recover a
    // real 2D shape when its first argument was a bare `RangeRef`,
    // defaulting to `num_cols = 1` (i.e. "one long row") for anything
    // else -- including the output of another array-reshaping function.
    // Found via `INDEX(EXPAND(A1:B2,3,3,0),3,3)` against real Excel:
    // Excel returned the pad value 0, visi returned 5 (silently reading
    // the wrong flat offset). Fixed by teaching `array_shape` to recurse
    // into known array-producing function calls (`function_call_cols`).
    let grid: [[&str; 3]; 2] = [
        ["1", "2", "=INDEX(EXPAND(A1:B2,3,3,0),3,3)"],
        ["4", "5", ""],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    assert_float_close(&sheet.get_result_data(&CellRef::new(0, 2)), 0.0, 1e-9);
}

#[test]
fn test_hstack_vstack_combine_arrays() {
    // HSTACK(SEQUENCE(2,1), SEQUENCE(2,1)) side-by-side: [[1,1],[2,2]].
    assert_float_close(
        &eval1("=INDEX(HSTACK(SEQUENCE(2,1),SEQUENCE(2,1)),2,1)"),
        2.0,
        1e-9,
    );
    assert_float_close(
        &eval1("=INDEX(HSTACK(SEQUENCE(2,1),SEQUENCE(2,1)),2,2)"),
        2.0,
        1e-9,
    );
    // VSTACK(SEQUENCE(1,2), SEQUENCE(1,2)) stacked: [[1,2],[1,2]].
    assert_float_close(
        &eval1("=INDEX(VSTACK(SEQUENCE(1,2),SEQUENCE(1,2)),2,2)"),
        2.0,
        1e-9,
    );
    assert_float_close(
        &eval1("=SUM(HSTACK(SEQUENCE(2,1),SEQUENCE(2,1)))"),
        6.0,
        1e-9,
    );
}

#[test]
fn test_chooserows_chosecols_select_by_index() {
    // SEQUENCE(3,3) = [[1,2,3],[4,5,6],[7,8,9]].
    assert_float_close(&eval1("=INDEX(CHOOSEROWS(SEQUENCE(3,3),2),1,1)"), 4.0, 1e-9);
    assert_float_close(
        &eval1("=INDEX(CHOOSEROWS(SEQUENCE(3,3),-1),1,1)"),
        7.0,
        1e-9,
    );
    assert_float_close(&eval1("=INDEX(CHOOSECOLS(SEQUENCE(3,3),2),1,1)"), 2.0, 1e-9);
}

#[test]
fn test_drop_take_slice_from_either_end() {
    assert_float_close(&eval1("=SUM(DROP(SEQUENCE(3,3),1))"), 39.0, 1e-9);
    assert_float_close(&eval1("=SUM(DROP(SEQUENCE(3,3),-1))"), 21.0, 1e-9);
    assert_float_close(&eval1("=SUM(TAKE(SEQUENCE(3,3),2))"), 21.0, 1e-9);
    assert_float_close(&eval1("=SUM(TAKE(SEQUENCE(3,3),-1))"), 24.0, 1e-9);
}

#[test]
fn test_expand_pads_with_given_value() {
    assert_float_close(&eval1("=INDEX(EXPAND(SEQUENCE(2,2),3,3,0),3,3)"), 0.0, 1e-9);
    assert_float_close(&eval1("=INDEX(EXPAND(SEQUENCE(2,2),3,3,0),1,1)"), 1.0, 1e-9);
}

#[test]
fn test_tocol_torow_flatten() {
    assert_float_close(&eval1("=SUM(TOCOL(SEQUENCE(3,3)))"), 45.0, 1e-9);
    assert_float_close(&eval1("=SUM(TOROW(SEQUENCE(3,3)))"), 45.0, 1e-9);
    assert_float_close(&eval1("=INDEX(TOCOL(SEQUENCE(2,2)),3)"), 3.0, 1e-9);
}

#[test]
fn test_wraprows_wrapcols_reshape_flat_sequence() {
    // Confirmed against real Excel: WRAPROWS(SEQUENCE(7),3,0) wraps
    // [1..7] into rows of 3, padding the last row with 0; WRAPCOLS does
    // the same column-major.
    assert_float_close(&eval1("=INDEX(WRAPROWS(SEQUENCE(7),3,0),3,1)"), 7.0, 1e-9);
    assert_float_close(&eval1("=SUM(WRAPROWS(SEQUENCE(7),3,0))"), 28.0, 1e-9);
    assert_float_close(&eval1("=INDEX(WRAPCOLS(SEQUENCE(7),3,0),1,3)"), 7.0, 1e-9);
    assert_float_close(&eval1("=SUM(WRAPCOLS(SEQUENCE(7),3,0))"), 28.0, 1e-9);
}

#[test]
fn test_unique_sort_sortby_filter_trimrange() {
    // UNIQUE(A1:A5) -> {10,20,5,30}; SUM should drop the duplicate 20.
    let grid_u: [[&str; 3]; 5] = [
        ["10", "", "=SUM(UNIQUE(A1:A5))"],
        ["20", "", ""],
        ["5", "", ""],
        ["20", "", ""],
        ["30", "", ""],
    ];
    let mut sheet_u = create_sheet(&grid_u);
    sheet_u.commit(None).unwrap();
    assert_float_close(&sheet_u.get_result_data(&CellRef::new(0, 2)), 65.0, 1e-9);

    // SORT(A1:A5, 1, -1) descending -> first element should be the max, 30.
    let grid_s: [[&str; 3]; 5] = [
        ["10", "", "=INDEX(SORT(A1:A5,1,-1),1)"],
        ["20", "", ""],
        ["5", "", ""],
        ["20", "", ""],
        ["30", "", ""],
    ];
    let mut sheet_s = create_sheet(&grid_s);
    sheet_s.commit(None).unwrap();
    assert_float_close(&sheet_s.get_result_data(&CellRef::new(0, 2)), 30.0, 1e-9);

    // SORTBY(A1:A5, B1:B5, -1): sort A by B descending; B's max (50) is row1 (A=10).
    let grid_sb: [[&str; 3]; 5] = [
        ["10", "50", "=INDEX(SORTBY(A1:A5,B1:B5,-1),1)"],
        ["20", "0", ""],
        ["5", "1", ""],
        ["20", "1", ""],
        ["30", "1", ""],
    ];
    let mut sheet_sb = create_sheet(&grid_sb);
    sheet_sb.commit(None).unwrap();
    assert_float_close(&sheet_sb.get_result_data(&CellRef::new(0, 2)), 10.0, 1e-9);

    // FILTER(A1:A5, B1:B5) with a genuine boolean helper range (not a
    // broadcast comparison -- this engine's comparison operators don't
    // broadcast across a range, a separate pre-existing limitation).
    let grid_f: [[&str; 3]; 5] = [
        ["10", "0", "=SUM(FILTER(A1:A5,B1:B5))"],
        ["20", "1", ""],
        ["5", "0", ""],
        ["20", "1", ""],
        ["30", "1", ""],
    ];
    let mut sheet_f = create_sheet(&grid_f);
    sheet_f.commit(None).unwrap();
    assert_float_close(&sheet_f.get_result_data(&CellRef::new(0, 2)), 70.0, 1e-9);

    // TRIMRANGE(A1:A5) with no blank/error padding is a pass-through.
    assert_float_close(&eval1("=SUM(TRIMRANGE(SEQUENCE(3,3)))"), 45.0, 1e-9);
}

#[test]
fn test_sort_and_sortby_always_place_blanks_last() {
    // Regression test: SORT/SORTBY used the same blank-coerces-to-0/""/
    // false comparator as ordinary `<`/`>` operators, so descending order
    // reversed that coercion and put a blank cell *first* (0 being the
    // largest of a set of negative numbers) instead of last. Real Excel
    // documents that both SORT and SORTBY always place blanks last,
    // regardless of sort direction -- found via the differential fuzzer:
    // `SORT({-215.8,,-100,-240.97,-88},1,-1)` gave 0 instead of -88 for
    // its first element.
    let grid: [[&str; 4]; 5] = [
        ["-215.8", "10", "", "=INDEX(SORT(A1:A5,1,-1),1)"],
        ["", "20", "4", "=INDEX(SORTBY(B1:B5,C1:C5,-1),1)"],
        ["-100", "30", "3", ""],
        ["-240.97", "40", "2", ""],
        ["-88", "50", "1", ""],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    // SORT(A1:A5,1,-1): A2 is blank, sorted last regardless of direction,
    // so the largest real number (-88) is first.
    assert_float_close(&sheet.get_result_data(&CellRef::new(0, 3)), -88.0, 1e-9);
    // SORTBY(B1:B5,C1:C5,-1): C1 is blank, sorted last regardless of
    // direction, so the row with C's largest real value (row2, C=4, B=20)
    // comes first.
    assert_float_close(&sheet.get_result_data(&CellRef::new(1, 3)), 20.0, 1e-9);
}

#[test]
fn test_lookup_vector_form_on_sorted_data() {
    // LOOKUP requires an ascending-sorted lookup vector; Excel's own docs
    // call behavior on unsorted input unpredictable (confirmed
    // divergent-but-Excel-undefined via the differential fuzzer on
    // unsorted input, not treated as a visi bug). On sorted input, visi
    // matches real Excel exactly for exact, mid-range, below-range, and
    // above-range lookups.
    let grid: [[&str; 2]; 5] = [
        ["5", "50"],
        ["10", "100"],
        ["20", "200"],
        ["20", "201"],
        ["30", "300"],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    let cases: [(&str, f64); 3] = [
        ("=LOOKUP(20,A1:A5,B1:B5)", 201.0),
        ("=LOOKUP(15,A1:A5,B1:B5)", 100.0),
        ("=LOOKUP(100,A1:A5,B1:B5)", 300.0),
    ];
    for (formula, expected) in cases {
        let (result, _) = sheet.eval(formula, None).unwrap();
        assert_float_close(&result, expected, 1e-9);
    }
    assert!(matches!(
        sheet.eval("=LOOKUP(1,A1:A5,B1:B5)", None).unwrap().0,
        ResultData::Error(ref e) if e.contains("#N/A")
    ));
}

#[test]
fn test_xmatch_supports_next_smaller_and_larger_modes() {
    let grid: [[&str; 1]; 5] = [["10"], ["20"], ["5"], ["20"], ["30"]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    let (exact, _) = sheet.eval("=XMATCH(20,A1:A5)", None).unwrap();
    assert_float_close(&exact, 2.0, 1e-9);
    let (next_smaller, _) = sheet.eval("=XMATCH(25,A1:A5,-1)", None).unwrap();
    assert_float_close(&next_smaller, 2.0, 1e-9);
}

#[test]
fn test_bare_row_and_column_report_current_cell_position() {
    // No-arg ROW()/COLUMN() report the position of the cell the formula
    // itself lives in -- distinct from the reference-argument form, which
    // reports the referenced range instead (already covered elsewhere).
    let grid: [[&str; 3]; 2] = [["=ROW()", "=COLUMN()", "x"], ["x", "x", "=ROW()+COLUMN()"]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    assert_float_close(&sheet.get_result_data(&CellRef::new(0, 0)), 1.0, 1e-9);
    assert_float_close(&sheet.get_result_data(&CellRef::new(0, 1)), 2.0, 1e-9);
    assert_float_close(&sheet.get_result_data(&CellRef::new(1, 2)), 5.0, 1e-9);
}

#[test]
fn test_let_binds_names_in_sequence_and_rejects_duplicate_names() {
    assert_float_close(&eval1("=LET(x, 5, x * 2)"), 10.0, 1e-9);
    // Later pairs can reference earlier ones in the same LET.
    assert_float_close(&eval1("=LET(x, 5, y, x + 1, x + y)"), 11.0, 1e-9);
    assert!(matches!(
        eval1("=LET(x, 1, x, 2, x)"),
        ResultData::Error(ref e) if e == "#VALUE!"
    ));
}

#[test]
fn test_randarray_respects_shape_bounds_and_whole_number_flag() {
    match eval1("=RANDARRAY(2,3,10,20,TRUE)") {
        ResultData::List(rows) => {
            assert_eq!(rows.len(), 2);
            for row in &rows {
                match row {
                    ResultData::List(cells) => {
                        assert_eq!(cells.len(), 3);
                        for cell in cells {
                            let v = match cell {
                                ResultData::Float(f) => *f,
                                ResultData::Integer(i) => *i as f64,
                                other => panic!("expected numeric cell, got {other:?}"),
                            };
                            assert!((10.0..=20.0).contains(&v), "{v} out of [10,20]");
                            assert_eq!(v.fract(), 0.0, "expected a whole number, got {v}");
                        }
                    }
                    other => panic!("expected a row List, got {other:?}"),
                }
            }
        }
        other => panic!("expected a List of rows, got {other:?}"),
    }
}

#[test]
fn test_hlookup_exact_and_approximate_match_on_text_header_row() {
    // extract_matrix-based HLOOKUP used to coerce every cell through
    // to_f64 and silently drop non-numeric ones, so a text header row (the
    // common HLOOKUP case) never matched -- see #26.
    let grid: [[&str; 3]; 2] = [["Jan", "Feb", "Mar"], ["10", "20", "30"]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    let (exact, _) = sheet.eval("=HLOOKUP(\"Feb\",A1:C2,2,FALSE)", None).unwrap();
    assert_float_close(&exact, 20.0, 1e-9);

    assert!(matches!(
        sheet.eval("=HLOOKUP(\"Nope\",A1:C2,2,FALSE)", None).unwrap().0,
        ResultData::Error(ref e) if e == "#N/A"
    ));

    let numeric_grid: [[&str; 3]; 2] = [["10", "20", "30"], ["a", "b", "c"]];
    let mut numeric_sheet = create_sheet(&numeric_grid);
    numeric_sheet.commit(None).unwrap();
    // Approximate match: largest header <= 25 is 20, in column 2.
    let (approx, _) = numeric_sheet.eval("=HLOOKUP(25,A1:C2,2)", None).unwrap();
    assert_eq!(approx.to_string(), "b");
}

#[test]
fn test_date_functions_match_documented_excel_examples() {
    // Jan 1, 2024 (serial 45292) is a Monday. WORKDAY skips both weekend
    // days landing on Mon 1/8/2024 (serial 45299) five working days later
    // (Tue-Fri, then Mon).
    assert_float_close(&eval1("=WORKDAY(DATE(2024,1,1),5)"), 45299.0, 1e-9);
    // Inclusive of both endpoints, excluding the Sat/Sun in between:
    // Jan 1, 2, 3, 4, 5, 8 = 6 working days.
    assert_float_close(
        &eval1("=NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,8))"),
        6.0,
        1e-9,
    );
    // EOMONTH(0) is the same month's last day; EOMONTH(1) rolls into
    // Feb 2024, a leap year (serial 45351 = Feb 29), so the last day is
    // the 29th.
    assert_float_close(&eval1("=DAY(EOMONTH(DATE(2024,1,15),0))"), 31.0, 1e-9);
    assert_float_close(&eval1("=EOMONTH(DATE(2024,1,15),1)"), 45351.0, 1e-9);
    // WEEKNUM with the default return type (week starts Sunday): Jan 1,
    // 2024 (a Monday) is always week 1; the first Sunday (Jan 7) starts
    // week 2.
    assert_float_close(&eval1("=WEEKNUM(DATE(2024,1,1))"), 1.0, 1e-9);
    assert_float_close(&eval1("=WEEKNUM(DATE(2024,1,7))"), 2.0, 1e-9);
    // Microsoft's own DAYS360 documentation example (US/NASD method).
    assert_float_close(
        &eval1("=DAYS360(DATE(2011,1,30),DATE(2011,2,1))"),
        1.0,
        1e-9,
    );
}

#[test]
fn test_besselk_bessely_match_known_reference_values() {
    // Regression for #26: BESSELK/BESSELY used to alias BESSELI/BESSELJ
    // directly, which is a distinct-function bug, not an imprecision --
    // K_n/Y_n diverge as x -> 0 while I_n/J_n stay finite there. Expected
    // values are well-known constants (Abramowitz & Stegun tables).
    assert_float_close(&eval1("=BESSELK(1,0)"), 0.4210244382, 1e-8);
    assert_float_close(&eval1("=BESSELK(1,1)"), 0.6019072301, 1e-8);
    assert_float_close(&eval1("=BESSELY(1,0)"), 0.0882569642, 1e-8);
    assert_float_close(&eval1("=BESSELY(1,1)"), -0.7812128213, 1e-8);
    // Sanity check the still-correct BESSELI/BESSELJ weren't disturbed.
    assert_float_close(&eval1("=BESSELI(1,0)"), 1.2660658778, 1e-8);
    assert_float_close(&eval1("=BESSELJ(1,0)"), 0.7651976866, 1e-8);
}

#[test]
fn test_complex_number_functions_round_trip() {
    assert_eq!(eval1("=COMPLEX(3,4)").to_string(), "3+4i");
    assert_float_close(&eval1("=IMABS(\"3+4i\")"), 5.0, 1e-9);
    assert_float_close(&eval1("=IMREAL(\"3+4i\")"), 3.0, 1e-9);
    assert_float_close(&eval1("=IMAGINARY(\"3+4i\")"), 4.0, 1e-9);
    assert_eq!(eval1("=IMCONJUGATE(\"3+4i\")").to_string(), "3-4i");
    assert_eq!(eval1("=IMSUM(\"3+4i\",\"1-2i\")").to_string(), "4+2i");
    assert_eq!(eval1("=IMSUB(\"3+4i\",\"1-2i\")").to_string(), "2+6i");
    // (3+4i)(1-2i) = 3-6i+4i-8i^2 = 3-2i+8 = 11-2i
    assert_eq!(eval1("=IMPRODUCT(\"3+4i\",\"1-2i\")").to_string(), "11-2i");
    // (3+4i)/(1-2i) = (3+4i)(1+2i)/5 = (3+6i+4i-8)/5 = (-5+10i)/5 = -1+2i
    assert_eq!(eval1("=IMDIV(\"3+4i\",\"1-2i\")").to_string(), "-1+2i");
}

#[test]
fn test_cube_webservice_image_report_unavailable_connections_not_echo_stub_args() {
    // Regression for #26: these used to echo their last argument back as
    // a placeholder result -- a plausible-looking wrong value is worse
    // than a visible error, since it can silently corrupt a downstream
    // calculation with no signal anything is wrong. None has a local data
    // source this engine can serve (a live OLAP cube connection, actual
    // network access, real image decoding); the error codes match what
    // real Excel shows once its equivalent live connection/resource is
    // unavailable (#N/A for the CUBE* family, mirroring RTD/STOCKHISTORY
    // just below; #VALUE! for WEBSERVICE/IMAGE, per Microsoft's own docs).
    for formula in [
        "=CUBEKPIMEMBER(\"conn\",\"kpi\",1)",
        "=CUBEMEMBER(\"conn\",\"member\")",
        "=CUBEMEMBERPROPERTY(\"conn\",\"member\",\"prop\")",
        "=CUBERANKEDMEMBER(\"conn\",\"set\",1)",
        "=CUBESET(\"conn\",\"set\")",
        "=CUBESETCOUNT(\"set\")",
        "=CUBEVALUE(\"conn\",\"member\")",
    ] {
        assert!(
            matches!(eval1(formula), ResultData::Error(ref e) if e == "#N/A"),
            "expected #N/A for {formula}"
        );
    }
    assert!(matches!(
        eval1("=WEBSERVICE(\"https://example.com\")"),
        ResultData::Error(ref e) if e == "#VALUE!"
    ));
    assert!(matches!(
        eval1("=IMAGE(\"https://example.com/pic.png\")"),
        ResultData::Error(ref e) if e == "#VALUE!"
    ));
}

#[test]
fn test_stockhistory_and_rtd_report_unavailable_data_source() {
    // Neither has a local data source this engine can serve (a live
    // Microsoft stock-data cloud connection, a registered Windows COM RTD
    // server) -- #N/A matches real Excel's own display once its
    // equivalent live connection is unavailable.
    assert!(matches!(
        eval1("=STOCKHISTORY(\"MSFT\",\"2024-01-01\",\"2024-01-31\")"),
        ResultData::Error(ref e) if e == "#N/A"
    ));
    assert!(matches!(
        eval1("=RTD(\"prog.id\",\"server\",\"topic\")"),
        ResultData::Error(ref e) if e == "#N/A"
    ));
}

#[test]
fn test_coupon_schedule_is_end_of_month_when_maturity_is() {
    // A maturity on the last day of its month puts the whole coupon
    // schedule on month-ends, so a step that lands in a leap year takes the
    // 29th rather than the anchor's 28th. Stepping by a fixed day-of-month
    // instead made COUPPCD report the settlement date itself.
    // Values are verbatim real Excel.
    let settlement = "DATE(2024,2,28)";
    let maturity = "EDATE(DATE(2024,2,28),180)"; // 2039-02-28, a month end
    // 2023-02-28
    assert_float_close(
        &eval1(&format!("=COUPPCD({settlement}, {maturity}, 1)")),
        44985.0,
        1e-9,
    );
    // 2024-02-29 -- the 29th, not the 28th.
    assert_float_close(
        &eval1(&format!("=COUPNCD({settlement}, {maturity}, 1)")),
        45351.0,
        1e-9,
    );
    assert_float_close(
        &eval1(&format!("=COUPNUM({settlement}, {maturity}, 1)")),
        16.0,
        1e-9,
    );
    // Semi-annual on the same bond: 2023-08-31, again a month end.
    assert_float_close(
        &eval1(&format!("=COUPPCD({settlement}, {maturity}, 2)")),
        45169.0,
        1e-9,
    );

    // A maturity that is *not* a month end keeps its day-of-month.
    // 2024-05-15 settling against 2039-06-15, semi-annual.
    assert_float_close(
        &eval1("=COUPPCD(DATE(2024,5,15), EDATE(DATE(2024,5,15),181), 2)"),
        45275.0, // 2023-12-15
        1e-9,
    );
    assert_float_close(
        &eval1("=COUPNCD(DATE(2024,5,15), EDATE(DATE(2024,5,15),181), 2)"),
        45458.0, // 2024-06-15
        1e-9,
    );
}

#[test]
fn test_amordegrc_keeps_full_precision_in_the_running_balance() {
    // The declining balance carries full precision; only the returned
    // figure is rounded. Rounding each period and subtracting the rounded
    // amount compounds the error and shifts a later period by a whole unit
    // -- period 2 below is 4624.4757 carried exactly (Excel: 4624), but
    // 4624.508 -> 4625 once the two preceding periods have been rounded.
    // All four values are verbatim real Excel.
    let call = |period: i32| {
        format!(
            "=AMORDEGRC(27370.88, DATE(1998,10,17), EDATE(DATE(1998,10,17),2), 6352.4, {period}, 0.0909, 0)"
        )
    };
    for (period, want) in [(0, 1037.0), (1, 5984.0), (2, 4624.0), (3, 3574.0)] {
        assert_float_close(&eval1(&call(period)), want, 1e-9);
    }
}
