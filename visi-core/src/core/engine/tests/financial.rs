use super::*;

#[test]
fn test_yearfrac_basis1_uses_actual_actual_year_average() {
    // Confirmed against real Excel via the differential fuzzer: basis 1
    // averages 365/366 across every calendar year the span touches, not
    // the average Julian year (365.2425) the previous implementation used.
    assert_float_close(
        &eval_formula("=YEARFRAC(DATE(1998,8,8),DATE(1998,9,8),1)"),
        0.08493150684931507,
        1e-9,
    );
    assert_float_close(
        &eval_formula("=YEARFRAC(DATE(2016,1,1),DATE(2016,6,1),1)"),
        0.41530054644808745,
        1e-9,
    );
    assert_float_close(
        &eval_formula("=YEARFRAC(DATE(2017,6,1),DATE(2020,9,1),1)"),
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
        &eval_formula("=COUPDAYBS(DATE(2011,1,25), DATE(2011,11,15), 2, 1)"),
        71.0,
        1e-9,
    );
    assert_float_close(
        &eval_formula("=COUPDAYS(DATE(2011,1,25), DATE(2011,11,15), 2, 1)"),
        181.0,
        1e-9,
    );
    assert_float_close(
        &eval_formula("=COUPNUM(DATE(2011,1,25), DATE(2011,11,15), 2)"),
        2.0,
        1e-9,
    );
    assert_float_close(
        &eval_formula("=COUPPCD(DATE(2011,1,25), DATE(2011,11,15), 2)"),
        40497.0,
        1e-9,
    );
}

#[test]
fn test_coupncd_handles_day_of_month_clamping_correctly() {
    // Walking a coupon schedule by anchoring each quasi-coupon date from the
    // maturity anchor handles day-of-month clamping in short months (e.g. April).
    assert_float_close(
        &eval_formula("=COUPNCD(DATE(2007,11,21), EDATE(DATE(2007,11,21),30), 4)"),
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
        &eval_formula("=COUPDAYSNC(DATE(2030,7,25), EDATE(DATE(2030,7,25),48), 2, 3)"),
        184.0,
        1e-9,
    );
}

#[test]
fn test_coupdays_basis1_quarterly_november_to_february_schedule() {
    assert_float_close(
        &eval_formula("=COUPDAYS(DATE(2000,11,28), EDATE(DATE(2000,11,28),54), 4, 1)"),
        91.0,
        1e-9,
    );
    assert_float_close(
        &eval_formula("=COUPDAYS(DATE(2006,11,8), EDATE(DATE(2006,11,8),15), 4, 1)"),
        92.0,
        1e-9,
    );
    assert_float_close(
        &eval_formula("=COUPDAYS(DATE(2005,11,16), EDATE(DATE(2005,11,16),42), 4, 1)"),
        92.0,
        1e-9,
    );
}

#[test]
fn test_price_and_yield_are_inverses_and_match_real_excel() {
    assert_float_close(
        &eval_formula(
            "=PRICE(DATE(1997,5,13), EDATE(DATE(1997,5,13),42), 0.0606, 0.0885, 100, 2, 3)",
        ),
        91.75707270015924,
        1e-6,
    );
    assert_float_close(
        &eval_formula(
            "=PRICE(DATE(2015,3,15), EDATE(DATE(2015,3,15),72), 0.0248, 0.0816, 100, 1, 2)",
        ),
        73.86901031390656,
        1e-6,
    );
    assert_float_close(
        &eval_formula(
            "=YIELD(DATE(2032,10,16), EDATE(DATE(2032,10,16),51), 0.0846, 103.71, 105, 4, 3)",
        ),
        0.0840390355112947,
        1e-9,
    );
}

#[test]
fn test_duration_matches_real_excel() {
    assert_float_close(
        &eval_formula("=DURATION(DATE(2008,1,1), DATE(2016,1,1), 0.08, 0.09, 2, 1)"),
        5.993774955545186,
        1e-6,
    );
}

#[test]
fn test_disc_pricedisc_yielddisc_match_real_excel() {
    assert_float_close(
        &eval_formula("=DISC(DATE(1998,4,12), EDATE(DATE(1998,4,12),4), 93.41, 100, 1)"),
        0.197159836065574,
        1e-9,
    );
    assert_float_close(
        &eval_formula("=PRICEDISC(DATE(1996,2,2), EDATE(DATE(1996,2,2),14), 0.044, 100, 1)"),
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
        &eval_formula("=DISC(DATE(2016,6,1), DATE(2016,9,1), 93, 100, 1)"),
        0.278478260869565,
        1e-9,
    );
    assert_float_close(
        &eval_formula("=PRICEDISC(DATE(2027,12,26), EDATE(DATE(2027,12,26),3), 0.0899, 100, 1)"),
        97.76478142076503,
        1e-6,
    );
    assert_float_close(
        &eval_formula("=DISC(DATE(2017,6,1), DATE(2020,9,1), 70, 100, 1)"),
        0.0922348484848485,
        1e-9,
    );

    // The two 366 cases above are "wholly inside a leap year" and "spans
    // 29 February". A short period that merely *ends* in a leap year,
    // before the leap day, takes 365 -- taking the end year's leap-ness
    // alone got these wrong.
    assert_float_close(
        &eval_formula("=PRICEDISC(DATE(2023,8,5), EDATE(DATE(2023,8,5),5), 0.0226, 100, 1)"),
        99.05265753424658,
        1e-9,
    );
    assert_float_close(
        &eval_formula("=YEARFRAC(DATE(2015,2,15), DATE(2016,2,13), 1)"),
        0.994520547945205,
        1e-12,
    );
}

#[test]
fn test_pricemat_yieldmat_basis0_does_not_promote_maturity_february_month_end() {
    assert_float_close(
        &eval_formula(
            "=PRICEMAT(EDATE(DATE(2002,3,28),6), EDATE(DATE(2002,3,28),11), DATE(2002,3,28), 0.0224, 0.04, 0)",
        ),
        99.26032786885246,
        1e-9,
    );
    assert_float_close(
        &eval_formula(
            "=YIELDMAT(EDATE(DATE(2002,3,28),6), EDATE(DATE(2002,3,28),11), DATE(2002,3,28), 0.0224, 99.26032786885246, 0)",
        ),
        0.04,
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
        &eval_formula(
            "=YIELDMAT(EDATE(DATE(2012,9,3),2), EDATE(DATE(2012,9,3),25), DATE(2012,9,3), 0.079, 119.55, 1)",
        ),
        -0.019331059183910253,
        1e-9,
    );
    assert_float_close(
        &eval_formula(
            "=YIELDMAT(EDATE(DATE(1995,9,4),6), EDATE(DATE(1995,9,4),11), DATE(1995,9,4), 0.0226, 95.27, 1)",
        ),
        0.14082750571984862,
        1e-9,
    );
}

#[test]
fn test_received_and_intrate_match_real_excel() {
    assert_float_close(
        &eval_formula("=RECEIVED(DATE(2008,2,15), DATE(2008,5,15), 1000000, 0.0575, 2)"),
        1014584.6544071021,
        1e-3,
    );
    assert_float_close(
        &eval_formula("=INTRATE(DATE(2020,9,4), EDATE(DATE(2020,9,4),13), 24974.48, 5441.6, 1)"),
        -0.7237025672261941,
        1e-9,
    );
}

#[test]
fn test_tbill_functions_match_real_excel() {
    assert_float_close(
        &eval_formula("=TBILLPRICE(DATE(2008,3,31), DATE(2008,6,1), 0.09)"),
        98.45,
        1e-9,
    );
    assert_float_close(
        &eval_formula("=TBILLYIELD(DATE(2008,3,31), DATE(2008,6,1), 98.45)"),
        0.09141696292534264,
        1e-9,
    );
    assert_float_close(
        &eval_formula("=TBILLEQ(DATE(2008,3,31), DATE(2008,6,1), 0.0914)"),
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
        &eval_formula(
            "=ACCRINT(DATE(2017,6,7), DATE(2018,6,7), DATE(2018,7,7), 0.0547, 14735.46, 1, 0, TRUE)",
        ),
        873.1988004999998,
        1e-6,
    );
    assert_float_close(
        &eval_formula(
            "=ACCRINT(DATE(2017,6,7), DATE(2018,6,7), DATE(2018,7,7), 0.0547, 14735.46, 1, 0, FALSE)",
        ),
        873.1988004999998,
        1e-6,
    );
}

#[test]
fn test_accrint_basis0_february_month_end_issue_uses_whole_span() {
    assert_float_close(
        &eval_formula(
            "=ACCRINT(DATE(2003,2,28), EDATE(DATE(2003,2,28),6), EDATE(DATE(2003,2,28),24), 0.0171, 34973.86, 2, 0, FALSE)",
        ),
        1196.106012,
        1e-6,
    );
    assert_float_close(
        &eval_formula(
            "=ACCRINT(DATE(2004,2,29), EDATE(DATE(2004,2,29),6), EDATE(DATE(2004,2,29),18), 0.05, 10000, 2, 0, FALSE)",
        ),
        748.6111111111111,
        1e-9,
    );
}

#[test]
fn test_accrintm_matches_real_excel() {
    assert_float_close(
        &eval_formula("=ACCRINTM(DATE(1998,8,8), EDATE(DATE(1998,8,8),1), 0.096, 37328.54, 1)"),
        304.3554384657534,
        1e-6,
    );
}

#[test]
fn test_amorlinc_amordegrc_reject_basis_2() {
    // Confirmed against real Excel: unlike every other function in
    // finance.rs, AMORLINC/AMORDEGRC reject basis 2 (actual/360).
    assert!(matches!(
        eval_formula("=AMORLINC(17737.01, DATE(2026,5,16), EDATE(DATE(2026,5,16),9), 5082.98, 1, 0.5, 2)"),
        ResultData::Error(ref e) if e.contains("#NUM!")
    ));
    assert!(matches!(
        eval_formula("=AMORDEGRC(8000.89, DATE(1995,11,26), EDATE(DATE(1995,11,26),9), 1484.45, 12, 0.05, 2)"),
        ResultData::Error(ref e) if e.contains("#NUM!")
    ));
}

#[test]
fn test_amordegrc_rejects_life_at_or_below_two_years() {
    // Confirmed against real Excel: the threshold is exactly life > 2
    // (rate < 0.5), not life >= 3 where the next coefficient bracket
    // starts -- life == 2 (rate == 0.5) is already rejected.
    assert!(matches!(
        eval_formula("=AMORDEGRC(9832.03, DATE(2024,9,7), EDATE(DATE(2024,9,7),5), 2414.1, 0, 0.5, 3)"),
        ResultData::Error(ref e) if e.contains("#NUM!")
    ));
    assert_float_close(
        &eval_formula(
            "=AMORDEGRC(9832.03, DATE(2024,9,7), EDATE(DATE(2024,9,7),5), 2414.1, 0, 0.499, 3)",
        ),
        3085.0,
        1e-6,
    );
}

#[test]
fn test_amordegrc_returns_declining_amount_before_later_zero_periods() {
    assert_float_close(
        &eval_formula(
            "=AMORDEGRC(48665.34, DATE(1998,12,22), EDATE(DATE(1998,12,22),11), 7901.84, 13, 0.05, 0)",
        ),
        1085.0,
        1e-9,
    );
}

#[test]
fn test_fuzz_amordegrc_uses_rounded_first_period_for_remaining_balance() {
    assert_float_close(
        &eval_formula(
            "=AMORDEGRC(38511.1, DATE(2027,8,12), EDATE(DATE(2027,8,12),10), 4806.98, 2, 0.1111, 3)",
        ),
        5933.0,
        1e-9,
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
        assert_float_close(&eval_formula(&f), *exp, 1e-6);
    }
}

#[test]
fn test_oddfprice_oddfyield_match_real_excel() {
    assert_float_close(
        &eval_formula(
            "=ODDFPRICE((DATE(2012,7,19)+26), EDATE((DATE(2012,7,19)+60),24), DATE(2012,7,19), (DATE(2012,7,19)+60), 0.082, 0.0923, 100, 4, 3)",
        ),
        98.06021551292406,
        1e-6,
    );
    // Basis 2: ODDFPRICE keeps COUPDAYS's idealized 360/365-per-freq value
    // for E on every basis except 1 (unlike ODDLPRICE/ODDLYIELD below).
    assert_float_close(
        &eval_formula(
            "=ODDFPRICE(DATE(2012,7,19)+26, EDATE(DATE(2012,7,19)+60,24), DATE(2012,7,19), DATE(2012,7,19)+60, 0.082, 0.0923, 100, 4, 2)",
        ),
        98.05901871412289,
        1e-6,
    );
    assert_float_close(
        &eval_formula(
            "=ODDFYIELD((DATE(2008,2,16)+20), EDATE((DATE(2008,2,16)+25),12), DATE(2008,2,16), (DATE(2008,2,16)+25), 0.0923, 117.26, 100, 4, 1)",
        ),
        -0.07045544328557858,
        1e-6,
    );
}

#[test]
fn test_oddlprice_oddlyield_match_real_excel() {
    assert_float_close(
        &eval_formula(
            "=ODDLYIELD((DATE(2029,6,23)+28), (DATE(2029,6,23)+44), DATE(2029,6,23), 0.0404, 115.8, 100, 4, 2)",
        ),
        -3.0950656625987194,
        1e-6,
    );
    assert_float_close(
        &eval_formula(
            "=ODDLYIELD((DATE(2002,7,10)+45), (DATE(2002,7,10)+114), DATE(2002,7,10), 0.0775, 96.4, 100, 1, 2)",
        ),
        0.2752128427867764,
        1e-6,
    );
    // Basis 3: unlike ODDFPRICE/ODDFYIELD, ODDLPRICE/ODDLYIELD use the
    // *actual* adjacent-period length for E on every basis, including 3.
    assert_float_close(
        &eval_formula(
            "=ODDLYIELD((DATE(2004,7,25)+51), (DATE(2004,7,25)+57), DATE(2004,7,25), 0.0682, 80.65, 100, 2, 3)",
        ),
        14.62856320740453,
        1e-6,
    );
    // Basis 1 and a freq=2/basis=2 case: tests E being anchored at the
    // regular period *following* last_interest, not the one *preceding* maturity.
    assert_float_close(
        &eval_formula(
            "=ODDLYIELD((DATE(2001,3,3)+10), (DATE(2001,3,3)+123), DATE(2001,3,3), 0.0546, 93.74, 100, 4, 1)",
        ),
        0.27529020678767524,
        1e-6,
    );
    assert_float_close(
        &eval_formula(
            "=ODDLPRICE(DATE(2019,6,21)+65, DATE(2019,6,21)+90, DATE(2019,6,21), 0.0478, 0.0376, 100, 2, 2)",
        ),
        100.06731898218347,
        1e-6,
    );
}

#[test]
fn test_fuzz_oddlyield_zero_dsc() {
    assert_float_close(
        &eval_formula(
            "=ODDLYIELD((DATE(2004, 8, 19) + 11), (DATE(2004, 8, 19) + 12), DATE(2004, 8, 19), 0.0395, 96.66, 105, 4, 0)",
        ),
        0.0,
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
    assert_float_close(
        &eval_formula("=EUROCONVERT(100, \"EUR\", \"DEM\")"),
        195.58,
        1e-9,
    );
    assert_float_close(
        &eval_formula("=EUROCONVERT(100, \"DEM\", \"EUR\")"),
        51.13,
        1e-9,
    );
}

#[test]
fn test_euroconvert_triangulates_through_eur() {
    assert_float_close(
        &eval_formula("=EUROCONVERT(1, \"DEM\", \"FRF\")"),
        3.35,
        1e-9,
    );
    assert_float_close(
        &eval_formula("=EUROCONVERT(1, \"DEM\", \"FRF\", TRUE)"),
        3.353854885138279,
        1e-9,
    );
}

#[test]
fn test_euroconvert_same_currency_is_a_no_op() {
    assert_float_close(
        &eval_formula("=EUROCONVERT(42, \"EUR\", \"EUR\")"),
        42.0,
        1e-9,
    );
}

#[test]
fn test_euroconvert_rounds_zero_decimal_currencies_to_whole_units() {
    // ITL/ESP/BEF/LUF had no meaningful subunit in everyday use.
    assert_float_close(
        &eval_formula("=EUROCONVERT(100, \"EUR\", \"ITL\")"),
        193627.0,
        1e-9,
    );
}

#[test]
fn test_euroconvert_rejects_unknown_currency_code() {
    assert!(matches!(
        eval_formula("=EUROCONVERT(100, \"EUR\", \"USD\")"),
        ResultData::Error(ref e) if e.contains("#VALUE!")
    ));
}

#[test]
fn test_euroconvert_rejects_triangulation_precision_below_3() {
    assert!(matches!(
        eval_formula("=EUROCONVERT(1, \"DEM\", \"FRF\", FALSE, 2)"),
        ResultData::Error(ref e) if e.contains("#NUM!")
    ));
}

#[test]
fn test_oddf_functions_match_excel_num_for_rejected_coupon_schedules() {
    assert!(matches!(
        eval_formula("=ODDFYIELD(DATE(2035,8,1)+150, EDATE(DATE(2035,8,1)+211,108), DATE(2035,8,1), DATE(2035,8,1)+211, 0.0422, 108.76, 105, 1, 2)"),
        ResultData::Error(ref e) if e.contains("#NUM!")
    ));
    assert!(matches!(
        eval_formula("=ODDFPRICE(DATE(2024,9,10)+25, EDATE(DATE(2024,9,10)+50,42), DATE(2024,9,10), DATE(2024,9,10)+50, 0.0548, 0.0988, 105, 2, 0)"),
        ResultData::Error(ref e) if e.contains("#NUM!")
    ));
}

#[test]
fn test_oddf_functions_reject_settlement_at_or_after_first_coupon() {
    assert!(matches!(
        eval_formula("=ODDFPRICE(DATE(1995,6,6)+57, EDATE(DATE(1995,6,6)+57,30), DATE(1995,6,6), DATE(1995,6,6)+57, 0.05, 0.05, 100, 2, 0)"),
        ResultData::Error(ref e) if e.contains("#NUM!")
    ));
    assert!(matches!(
        eval_formula("=ODDFYIELD(DATE(1995,6,6)+58, EDATE(DATE(1995,6,6)+57,30), DATE(1995,6,6), DATE(1995,6,6)+57, 0.05, 100, 100, 2, 0)"),
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
        eval_formula("=ODDFPRICE(DATE(2030,4,13)+0, EDATE(DATE(2030,4,13)+44,84), DATE(2030,4,13), DATE(2030,4,13)+44, 0.073, 0.0368, 100, 1, 3)"),
        ResultData::Error(ref e) if e.contains("#NUM!")
    ));
    assert!(matches!(
        eval_formula("=ODDFYIELD(DATE(2023,11,11)+0, EDATE(DATE(2023,11,11)+20,12), DATE(2023,11,11), DATE(2023,11,11)+20, 0.0562, 89.35, 105, 4, 3)"),
        ResultData::Error(ref e) if e.contains("#NUM!")
    ));
    assert!(matches!(
        eval_formula("=ODDLPRICE(DATE(2005,6,28)+0, DATE(2005,6,28)+58, DATE(2005,6,28), 0.0353, 0.0938, 105, 4, 4)"),
        ResultData::Error(ref e) if e.contains("#NUM!")
    ));
    assert!(matches!(
        eval_formula("=ODDLYIELD(DATE(2010,7,11)+0, DATE(2010,7,11)+15, DATE(2010,7,11), 0.028, 110.37, 100, 2, 1)"),
        ResultData::Error(ref e) if e.contains("#NUM!")
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
        &eval_formula(&format!("=COUPPCD({settlement}, {maturity}, 1)")),
        44985.0,
        1e-9,
    );
    // 2024-02-29 -- the 29th, not the 28th.
    assert_float_close(
        &eval_formula(&format!("=COUPNCD({settlement}, {maturity}, 1)")),
        45351.0,
        1e-9,
    );
    assert_float_close(
        &eval_formula(&format!("=COUPNUM({settlement}, {maturity}, 1)")),
        16.0,
        1e-9,
    );
    // Semi-annual on the same bond: 2023-08-31, again a month end.
    assert_float_close(
        &eval_formula(&format!("=COUPPCD({settlement}, {maturity}, 2)")),
        45169.0,
        1e-9,
    );

    // A maturity that is *not* a month end keeps its day-of-month.
    // 2024-05-15 settling against 2039-06-15, semi-annual.
    assert_float_close(
        &eval_formula("=COUPPCD(DATE(2024,5,15), EDATE(DATE(2024,5,15),181), 2)"),
        45275.0, // 2023-12-15
        1e-9,
    );
    assert_float_close(
        &eval_formula("=COUPNCD(DATE(2024,5,15), EDATE(DATE(2024,5,15),181), 2)"),
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
        assert_float_close(&eval_formula(&call(period)), want, 1e-9);
    }
}
