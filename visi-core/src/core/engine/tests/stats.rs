use super::*;

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
fn test_probability_distributions_match_known_reference_values() {
    // Expected values are closed-form (binomial/Poisson/exponential/
    // Weibull-with-shape-1/hypergeometric/negative-binomial coefficients
    // computed by hand) or, for CHISQ.DIST, the classic df=1 chi-square
    // critical value (3.841458821 <-> p=0.05). T.DIST/F.DIST are checked
    // against independent numeric integration of their PDFs (Simpson's
    // rule over ~4e5 points in a scratch Python script), not against this
    // engine's own incomplete-beta implementation, to avoid confirming a
    // shared bug. None of these had any regression test before -- #26
    // calls out the whole NORM.*/BETA.*/GAMMA*/CHISQ.* family as
    // fuzzer-covered but locally untested.
    // NORM.DIST/GAUSS go through `erf`'s Abramowitz & Stegun 7.1.26
    // approximation, whose documented max absolute error is ~1.5e-7 --
    // looser tolerance here reflects that implementation limit, not
    // uncertainty in the reference value itself.
    assert_float_close(&eval1("=NORM.DIST(1,0,1,TRUE)"), 0.8413447461, 1e-6);
    assert_float_close(&eval1("=PHI(0)"), 0.3989422804, 1e-8);
    assert_float_close(&eval1("=GAUSS(1.96)"), 0.4750021049, 1e-6);
    assert_float_close(&eval1("=GAMMA(5)"), 24.0, 1e-9);
    assert_float_close(&eval1("=GAMMALN(5)"), 3.1780538303, 1e-8);
    assert_float_close(&eval1("=GAMMA.DIST(2,3,1,FALSE)"), 0.2706705665, 1e-8);
    assert_float_close(&eval1("=GAMMA.DIST(2,3,1,TRUE)"), 0.3233235838, 1e-8);
    assert_float_close(&eval1("=BETA.DIST(0.5,2,2,FALSE,0,1)"), 1.5, 1e-8);
    assert_float_close(&eval1("=BETA.DIST(0.5,2,2,TRUE,0,1)"), 0.5, 1e-8);
    assert_float_close(&eval1("=BINOM.DIST(3,10,0.5,FALSE)"), 0.1171875, 1e-9);
    assert_float_close(&eval1("=BINOM.DIST(3,10,0.5,TRUE)"), 0.171875, 1e-9);
    assert_float_close(&eval1("=BINOM.DIST.RANGE(10,0.5,0,3)"), 0.171875, 1e-9);
    assert_float_close(&eval1("=BINOM.INV(10,0.5,0.5)"), 5.0, 1e-9);
    assert_float_close(&eval1("=POISSON.DIST(3,2,FALSE)"), 0.1804470443, 1e-8);
    assert_float_close(&eval1("=POISSON.DIST(3,2,TRUE)"), 0.8571234605, 1e-8);
    assert_float_close(&eval1("=WEIBULL.DIST(1,1,1,TRUE)"), 0.6321205588, 1e-8);
    assert_float_close(&eval1("=HYPGEOM.DIST(1,2,4,10,FALSE)"), 0.5333333333, 1e-8);
    assert_float_close(&eval1("=HYPGEOM.DIST(1,2,4,10,TRUE)"), 0.8666666667, 1e-8);
    assert_float_close(&eval1("=NEGBINOM.DIST(2,3,0.5,FALSE)"), 0.1875, 1e-9);
    assert_float_close(&eval1("=LOGNORM.DIST(1,0,1,TRUE)"), 0.5, 1e-9);
    assert_float_close(&eval1("=CHISQ.DIST(3.841458821,1,TRUE)"), 0.95, 1e-6);
    assert_float_close(&eval1("=CHISQ.DIST.RT(3.841458821,1)"), 0.05, 1e-6);
    // Independent numeric-integration reference values (see test doc comment).
    assert_float_close(&eval1("=T.DIST(1.5,10,TRUE)"), 0.9177463367, 1e-6);
    assert_float_close(&eval1("=F.DIST(2,5,10,TRUE)"), 0.8358050491, 1e-6);
}

#[test]
fn test_chisq_test_matches_independent_numeric_reference() {
    let grid = [["10", "20", "30"], ["15", "15", "30"]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    let (result, _) = sheet.eval("=CHISQ.TEST(A1:C1,A2:C2)", None).unwrap();
    assert_float_close(&result, 0.1888756028, 1e-6);
}

#[test]
fn test_rank_percentile_and_bivariate_functions_match_hand_computed_values() {
    let grid = [["10", "20", "20", "30", "40"]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    let (rank_eq_desc, _) = sheet.eval("=RANK.EQ(20,A1:E1,0)", None).unwrap();
    assert_float_close(&rank_eq_desc, 3.0, 1e-9);
    let (rank_eq_asc, _) = sheet.eval("=RANK.EQ(20,A1:E1,1)", None).unwrap();
    assert_float_close(&rank_eq_asc, 2.0, 1e-9);
    let (rank_avg, _) = sheet.eval("=RANK.AVG(20,A1:E1,0)", None).unwrap();
    assert_float_close(&rank_avg, 3.5, 1e-9);
    let (pct_exc, _) = sheet.eval("=PERCENTILE.EXC(A1:E1,0.5)", None).unwrap();
    assert_float_close(&pct_exc, 20.0, 1e-6);
    let (pct_exc_max, _) = sheet.eval("=PERCENTILE.EXC(A1:E1,5/6)", None).unwrap();
    assert_float_close(&pct_exc_max, 40.0, 1e-6);
    let (quart_exc, _) = sheet.eval("=QUARTILE.EXC(A1:E1,1)", None).unwrap();
    assert_float_close(&quart_exc, 15.0, 1e-6);
    let (prank_inc, _) = sheet.eval("=PERCENTRANK.INC(A1:E1,25)", None).unwrap();
    assert_float_close(&prank_inc, 0.625, 1e-6);
    let (prank_exc, _) = sheet.eval("=PERCENTRANK.EXC(A1:E1,25)", None).unwrap();
    assert_float_close(&prank_exc, 0.583, 1e-6);

    let bivar_grid = [["1", "2", "3", "4", "5"], ["2", "4", "5", "4", "5"]];
    let mut bivar_sheet = create_sheet(&bivar_grid);
    bivar_sheet.commit(None).unwrap();
    let (cov_p, _) = bivar_sheet
        .eval("=COVARIANCE.P(A1:E1,A2:E2)", None)
        .unwrap();
    assert_float_close(&cov_p, 1.2, 1e-9);
    let (cov_s, _) = bivar_sheet
        .eval("=COVARIANCE.S(A1:E1,A2:E2)", None)
        .unwrap();
    assert_float_close(&cov_s, 1.5, 1e-9);
    let (steyx, _) = bivar_sheet.eval("=STEYX(A2:E2,A1:E1)", None).unwrap();
    assert_float_close(&steyx, 0.8944271910, 1e-8);

    let trim_grid = [["1", "2", "3", "4", "5", "6", "7", "8", "9", "10"]];
    let mut trim_sheet = create_sheet(&trim_grid);
    trim_sheet.commit(None).unwrap();
    let (trimmean, _) = trim_sheet.eval("=TRIMMEAN(A1:J1,0.2)", None).unwrap();
    assert_float_close(&trimmean, 5.5, 1e-9);

    let moment_grid = [["2", "4", "4", "4", "5", "5", "7", "9"]];
    let mut moment_sheet = create_sheet(&moment_grid);
    moment_sheet.commit(None).unwrap();
    let (skew, _) = moment_sheet.eval("=SKEW(A1:H1)", None).unwrap();
    assert_float_close(&skew, 0.8184875534, 1e-8);
    let (skew_p, _) = moment_sheet.eval("=SKEW.P(A1:H1)", None).unwrap();
    assert_float_close(&skew_p, 0.65625, 1e-8);
    let (kurt, _) = moment_sheet.eval("=KURT(A1:H1)", None).unwrap();
    assert_float_close(&kurt, 0.940625, 1e-6);
}

#[test]
fn test_mode_sngl_and_mode_mult() {
    let grid = [["1", "2", "2", "3", "3", "4"]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    // Both 2 and 3 appear twice; MODE.SNGL keeps whichever appeared first.
    let (mode_sngl, _) = sheet.eval("=MODE.SNGL(A1:F1)", None).unwrap();
    assert_float_close(&mode_sngl, 2.0, 1e-9);
    let (mode_mult, _) = sheet.eval("=MODE.MULT(A1:F1)", None).unwrap();
    match mode_mult {
        ResultData::List(items) => {
            let vals: Vec<f64> = items
                .iter()
                .map(|v| match v {
                    ResultData::Float(f) => *f,
                    ResultData::Integer(i) => *i as f64,
                    other => panic!("expected numeric mode, got {other:?}"),
                })
                .collect();
            assert_eq!(vals, vec![2.0, 3.0]);
        }
        other => panic!("expected a List of modes, got {other:?}"),
    }
}

#[test]
fn test_fuzz_mode_mult_orders_ties_by_first_appearance_not_value() {
    // Harvested from fuzz/fuzz_excel.py, seed 101977: F1:G5 = {34, 0, 34,
    // "b", 10, 0, 479.283, -26, -50, 10}, a three-way tie between 34, 0
    // and 10 (each appears twice). Real Excel's INDEX(MODE.MULT(...), 1)
    // is 34 -- the first value to reach the tied count while scanning the
    // range -- not 0, the smallest. visi previously sorted tied modes by
    // value ascending.
    let grid = [
        ["34", "0"],
        ["0", "479.283"],
        ["34", "-26"],
        ["\"b\"", "-50"],
        ["10", "10"],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    let (first_mode, _) = sheet.eval("=INDEX(MODE.MULT(A1:B5), 1)", None).unwrap();
    assert_float_close(&first_mode, 34.0, 1e-9);
}

#[test]
fn test_t_test_uses_each_samples_own_length_for_its_mean() {
    // Regression for #26: t_test's test_type==2 (equal-variance two-sample)
    // branch divided array2's sum by array1's length (n1) instead of its
    // own (n2) to get its mean -- silently correct only when the two
    // samples happen to be the same length. Reference p-value computed by
    // an independent Python script (closed-form pooled-variance t-stat,
    // then numeric integration of the t-PDF for the p-value -- not this
    // engine's own incomplete-beta routine).
    let grid = [
        ["10", "12", "14", "16", "18", ""],
        ["20", "22", "24", "", "", ""],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    let (result, _) = sheet.eval("=T.TEST(A1:E1,A2:C2,2,2)", None).unwrap();
    assert_float_close(&result, 0.008237263171, 1e-6);
}

#[test]
fn test_f_test_and_confidence_intervals_match_independent_reference() {
    let grid = [
        ["10", "20", "30", "40", "50"],
        ["5", "10", "15", "20", "25"],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    // F.DIST.RT computed by independent numeric integration (see the
    // F.DIST test above); F.TEST's two-tailed p-value is
    // 2*min(p, 1-p) of that.
    let (f_test, _) = sheet.eval("=F.TEST(A1:E1,A2:E2)", None).unwrap();
    assert_float_close(&f_test, 0.208, 1e-4);

    // z_(0.975) = 1.959963985 and the df=24 t-critical value
    // t_(0.975,24) = 2.063898568 are well-known constants (also confirmed
    // here via independent bisection against a from-scratch t-PDF
    // integration), so CONFIDENCE.NORM/CONFIDENCE.T reduce to simple
    // arithmetic on them. CONFIDENCE.NORM's tolerance is loosened to match
    // `erf`'s ~1.5e-7 max error (inv_normal_cdf's Newton refinement step
    // uses `normal_cdf`, which is erf-based, so it inherits that bound).
    assert_float_close(&eval1("=CONFIDENCE.NORM(0.05,10,25)"), 3.919927969, 1e-5);
    assert_float_close(&eval1("=CONFIDENCE.T(0.05,10,25)"), 4.127797137, 1e-6);
}

#[test]
fn test_inverse_distributions_round_trip_through_their_forward_dist() {
    // The forward DIST directions above are checked against independent
    // reference values; these INV functions are verified as genuine
    // inverses of those already-verified forwards, which still catches a
    // broken/no-op INV without needing a second independent ground truth
    // for each one.
    assert_float_close(&eval1("=NORM.S.DIST(NORM.S.INV(0.9),TRUE)"), 0.9, 1e-6);
    assert_float_close(&eval1("=NORM.DIST(NORM.INV(0.3,5,2),5,2,TRUE)"), 0.3, 1e-6);
    assert_float_close(
        &eval1("=GAMMA.DIST(GAMMA.INV(0.4,3,2),3,2,TRUE)"),
        0.4,
        1e-5,
    );
    assert_float_close(
        &eval1("=BETA.DIST(BETA.INV(0.4,2,3,0,1),2,3,TRUE,0,1)"),
        0.4,
        1e-5,
    );
    assert_float_close(&eval1("=CHISQ.DIST(CHISQ.INV(0.9,5),5,TRUE)"), 0.9, 1e-5);
    assert_float_close(&eval1("=CHISQ.DIST.RT(CHISQ.INV.RT(0.1,5),5)"), 0.1, 1e-5);
    assert_float_close(&eval1("=T.DIST(T.INV(0.8,10),10,TRUE)"), 0.8, 1e-5);
    assert_float_close(
        &eval1("=LOGNORM.DIST(LOGNORM.INV(0.4,0,1),0,1,TRUE)"),
        0.4,
        1e-5,
    );
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

#[test]
fn test_inv_normal_cdf_matches_real_excel_to_near_double_precision() {
    // erf() used to be the classic Abramowitz & Stegun 7.1.26 rational
    // approximation, with a documented max error of ~1.5e-7 -- that
    // bounded the precision of everything built on it (normal_cdf,
    // inv_normal_cdf, and therefore CONFIDENCE/CONFIDENCE.NORM/
    // NORM.S.INV/NORMSINV/NORMINV/LOGINV/LOGNORM.INV), which is why all
    // of them mismatched real Excel in the ~7th significant digit on
    // every differential fuzzing run. Now delegates to libm's erf/erfc
    // (a pure-Rust fdlibm port, full double precision).
    // 1.959963984540054 is the well-known two-sided 95% confidence
    // z-value.
    assert_float_close(&eval1("=NORM.S.INV(0.975)"), 1.959963984540054, 1e-9);
}

#[test]
fn test_tdist_honors_tails_argument() {
    // Legacy TDIST(x, df, tails) was wired to the same handler as
    // T.DIST.2T (always two-tailed), completely ignoring the `tails`
    // argument -- TDIST(x, df, 1) (one-tailed) returned exactly double
    // the correct value, since the two-tailed probability is 2x the
    // one-tailed probability for a symmetric distribution. Found via
    // differential fuzzing (an exact 2x discrepancy against real Excel).
    let one_tailed = eval1("=TDIST(2, 10, 1)");
    let rt = eval1("=T.DIST.RT(2, 10)");
    assert_float_close(
        &one_tailed,
        match rt {
            ResultData::Float(v) => v,
            other => panic!("expected float, got {other:?}"),
        },
        1e-9,
    );

    let two_tailed = eval1("=TDIST(2, 10, 2)");
    let one_val = match one_tailed {
        ResultData::Float(v) => v,
        other => panic!("expected float, got {other:?}"),
    };
    assert_float_close(&two_tailed, one_val * 2.0, 1e-9);
}

#[test]
fn test_percentrank_truncates_to_significance_not_rounds() {
    // PERCENTRANK.INC/.EXC used `.round()` when limiting the result to
    // `significance` digits, but real Excel truncates instead -- e.g. a
    // raw value of 0.055555... at significance 3 displays as 0.055, not
    // the rounded 0.056. Found via differential fuzzing (a handful of
    // PERCENTRANK calls were off by exactly 0.001 against real Excel).
    let grid = [
        ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10"],
        [
            "=PERCENTRANK(A1:J1, 1.5)",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
        ],
    ];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    // rank = (0 + 0.5) / 9 = 0.0555... -> truncated to 0.055, not
    // rounded up to 0.056.
    let r = sheet.get_result_data(&CellRef::new(1, 0));
    assert!(
        matches!(r, ResultData::Float(v) if (v - 0.055).abs() < 1e-9),
        "{r:?}"
    );
}

#[test]
fn test_hypgeomdist_legacy_is_pmf_only_and_out_of_support_is_zero() {
    // Two independent HYPGEOM.DIST/HYPGEOMDIST bugs found via
    // differential fuzzing:
    //  - Legacy HYPGEOMDIST takes no `cumulative` argument at all -- it's
    //    always the point probability mass -- but the dispatcher defaulted
    //    a missing 5th argument to `true` (cumulative), so it silently
    //    summed the PMF from 0 up through the given count instead of
    //    just returning that one point's probability.
    //  - The PMF's log-combination formula assumed valid choose()
    //    arguments and produced a pole (NaN) instead of 0 once a count
    //    fell outside the distribution's actual support, propagating as
    //    #NUM! instead of the mathematically correct 0.
    let grid = [[
        "=HYPGEOMDIST(1, 4, 19, 45)",
        "=HYPGEOM.DIST(1, 4, 19, 45, FALSE)",
        "=HYPGEOM.DIST(4, 25, 13, 25, FALSE)",
    ]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();

    let r1 = sheet.get_result_data(&CellRef::new(0, 0));
    let r2 = sheet.get_result_data(&CellRef::new(0, 1));
    match (&r1, &r2) {
        (ResultData::Float(a), ResultData::Float(b)) => {
            assert!(
                (a - b).abs() < 1e-9,
                "HYPGEOMDIST(1,4,19,45)={a} should equal the non-cumulative HYPGEOM.DIST={b}"
            );
        }
        _ => panic!("expected floats, got {r1:?} / {r2:?}"),
    }

    let r3 = sheet.get_result_data(&CellRef::new(0, 2));
    assert!(
        matches!(r3, ResultData::Float(v) if v.abs() < 1e-9),
        "{r3:?}"
    );
}

#[test]
fn test_inverse_beta_and_f_distributions_converge_to_excel_values() {
    // inv_incbeta (which BETA.INV/BETAINV/F.INV/F.INV.RT/FINV all go
    // through) used unguarded Newton iteration that *clamped* an
    // overshooting step to [1e-12, 1-1e-12]. Those clamps are absorbing,
    // so once a step overshot, x stuck to the boundary and got returned
    // as the answer -- BETAINV would report a flat 1e-12 or
    // 0.999999999999, and F.INV (which maps the result back through
    // df2*y/(df1*(1-y)), a pole as y approaches 1) would report ~1e12.
    // Now safeguarded: Newton only when the step stays inside the current
    // bracket, bisection otherwise. Every expected value below was read
    // straight out of real Excel.
    for (f, expected) in [
        ("=BETAINV(0.945, 9.128, 5.585)", 0.8079143872863086),
        ("=_xlfn.BETA.INV(0.077, 4.347, 1.607)", 0.45886530331058883),
        ("=_xlfn.F.INV(0.119, 8, 1)", 0.3281233164680227),
        ("=_xlfn.F.INV(0.883, 1, 8)", 3.0866529196587305),
        ("=_xlfn.F.INV.RT(0.942, 18, 3)", 0.33370421499396513),
        ("=_xlfn.F.INV.RT(0.876, 17, 6)", 0.5030517141697566),
        ("=FINV(0.709, 1, 11)", 0.14670778600563464),
        ("=FINV(0.868, 10, 9)", 0.4767239715231606),
        ("=_xlfn.F.INV.RT(0.38, 17, 1)", 3.9202240523326743),
    ] {
        let got = eval1(f);
        match got {
            ResultData::Float(v) => {
                let rel = (v - expected).abs() / expected.abs().max(1e-300);
                assert!(rel < 1e-9, "{f}: got {v}, want {expected} (rel {rel:e})");
            }
            other => panic!("{f}: got {other:?}, want {expected}"),
        }
    }
}

#[test]
fn test_chitest_single_category_is_not_available() {
    // One category means zero degrees of freedom, so there is no
    // chi-square distribution to evaluate against and Excel reports #N/A.
    // The check lives at the call site because it is judged on the ranges'
    // *raw* size: applying it to the pairwise-filtered values instead
    // would turn a two-cell pair that merely holds one text cell into
    // #N/A, where Excel still reports the underlying #DIV/0!.
    // Numeric single cells, matching the case probed against real Excel
    // (a *blank* single-cell operand is a different rule -- see
    // paired_args -- and would report #VALUE!).
    let grid = [["5", "7", "=CHITEST(A1:A1, B1:B1)"]];
    let mut sheet = create_sheet(&grid);
    sheet.commit(None).unwrap();
    let got = sheet.get_result_data(&CellRef::new(0, 2));
    assert!(
        matches!(got, ResultData::Error(ref e) if e == "#N/A"),
        "{got:?}"
    );
    // Two categories still compute.
    assert!(crate::core::stats::chisq_test(&[10.0, 20.0], &[10.0, 20.0], 2).is_ok());
}

// ---------------------------------------------------------------------
// Direct-argument text coercion in the statistical family.
//
// All values below are from real Excel 16.111.3. The rule splits on how
// the text arrived: supplied directly as an argument it is coerced (and
// is an error if it will not coerce), while text reached through a
// reference is skipped. Getting this wrong is quiet rather than loud --
// DEVSQ("abc", 3, 4, 5) used to answer 2, the spread of the remaining
// three numbers, instead of #VALUE!.
// ---------------------------------------------------------------------

fn assert_err(source: &str, expected: &str) {
    match eval1(source) {
        ResultData::Error(e) => assert_eq!(e, expected, "for {source}"),
        other => panic!("expected {expected} for {source}, got {other:?}"),
    }
}

#[test]
fn test_direct_numeric_text_is_coerced_by_stat_family() {
    assert_float_close(&eval1("=SUM(\"12\", 3, 4, 5)"), 24.0, 1e-9);
    assert_float_close(&eval1("=AVERAGE(\"12\", 3, 4, 5)"), 6.0, 1e-9);
    assert_float_close(&eval1("=DEVSQ(\"12\", 3, 4, 5)"), 50.0, 1e-9);
    assert_float_close(&eval1("=STDEV(\"12\", 3, 4, 5)"), 4.08248290463863, 1e-12);
    assert_float_close(&eval1("=VAR(\"12\", 3, 4, 5)"), 16.666666666666668, 1e-12);
    assert_float_close(&eval1("=MEDIAN(\"12\", 3, 4, 5)"), 4.5, 1e-9);
    assert_float_close(&eval1("=SUMSQ(\"12\", 3, 4, 5)"), 194.0, 1e-9);
    assert_float_close(
        &eval1("=GEOMEAN(\"12\", 3, 4, 5)"),
        5.180040128222703,
        1e-12,
    );
    assert_float_close(&eval1("=AVEDEV(\"12\", 3, 4, 5)"), 3.0, 1e-9);
    assert_float_close(&eval1("=SKEW(\"12\", 3, 4, 5)"), 1.7636326148038874, 1e-12);
    assert_float_close(&eval1("=KURT(\"12\", 3, 4, 5)"), 3.2279999999999944, 1e-12);
}

#[test]
fn test_harmean_no_numeric_values_is_not_available() {
    // Real Excel reports #N/A, not #NUM!, when HARMEAN's referenced inputs
    // contain no numbers after blanks/text are skipped. IFNA therefore catches
    // it; this surfaced in fuzz/fuzz_excel.py seed 879563.
    let mut sheet = create_sheet(&[
        ["", "=HARMEAN(A1:A2)", "=IFNA(HARMEAN(A1:A2), 5)"],
        ["\"paren(test)\"", "", ""],
    ]);
    sheet.commit(None).unwrap();
    match sheet.get_result_data(&CellRef::new(0, 1)) {
        ResultData::Error(e) => assert_eq!(e, "#N/A"),
        other => panic!("expected #N/A, got {other:?}"),
    }
    assert_eq!(sheet.get_display_string(&CellRef::new(0, 2)), "5");
}

#[test]
fn test_direct_uncoercible_text_is_value_error_in_stat_family() {
    for f in [
        "SUM", "AVERAGE", "DEVSQ", "STDEV", "VAR", "MEDIAN", "MAX", "MIN", "PRODUCT", "SUMSQ",
        "GEOMEAN", "AVEDEV", "SKEW", "KURT",
    ] {
        assert_err(&format!("={f}(\"abc\", 3, 4, 5)"), "#VALUE!");
    }
}

#[test]
fn test_count_never_errors_on_text() {
    // COUNT is the deliberate exception: numeric text typed directly
    // counts, uncoercible text is simply not counted, and neither is an
    // error. Both values are real Excel's.
    assert_float_close(&eval1("=COUNT(\"12\", 3, 4, 5)"), 4.0, 1e-9);
    assert_float_close(&eval1("=COUNT(\"abc\", 3, 4, 5)"), 3.0, 1e-9);
}

#[test]
fn test_averagea_family_direct_vs_referenced_text() {
    // Real Excel, with A1 holding the *text* "12":
    //   AVERAGEA("12", 3) = 7.5   direct text is coerced
    //   AVERAGEA(A1, 3)   = 1.5   text in a reference counts as 0
    //   AVERAGEA("abc", 3)= #VALUE!
    //   AVERAGEA(TRUE, 3) = 2
    assert_float_close(&eval1("=AVERAGEA(\"12\", 3)"), 7.5, 1e-9);
    assert_float_close(&eval1("=AVERAGEA(TRUE, 3)"), 2.0, 1e-9);
    assert_float_close(&eval1("=MAXA(\"12\", 3)"), 12.0, 1e-9);
    assert_err("=AVERAGEA(\"abc\", 3)", "#VALUE!");

    let mut sheet = create_sheet(&[["=\"12\"", "=AVERAGEA(A1, 3)"]]);
    sheet.commit(None).unwrap();
    assert_float_close(&sheet.get_result_data(&CellRef::new(0, 1)), 1.5, 1e-9);
}

#[test]
fn test_erf_family_coerces_numeric_text_but_rejects_booleans() {
    // Real Excel: ERF("1") and ERF(" 1 ") are both 0.8427007929497149,
    // ERF("-39") is -1, while ERF(TRUE), ERF(FALSE) and ERF("abc") are
    // all #VALUE!. A blank argument is 0, so ERF(<blank>) is 0.
    assert_float_close(&eval1("=ERF(\"1\")"), 0.8427007929497149, 1e-15);
    assert_float_close(&eval1("=ERF(\" 1 \")"), 0.8427007929497149, 1e-15);
    assert_float_close(&eval1("=ERF(\"-39\")"), -1.0, 1e-15);
    assert_float_close(&eval1("=ERFC(\"1\")"), 0.15729920705028513, 1e-15);
    assert_err("=ERF(TRUE)", "#VALUE!");
    assert_err("=ERF(FALSE)", "#VALUE!");
    assert_err("=ERFC(TRUE)", "#VALUE!");
    assert_err("=ERF(\"abc\")", "#VALUE!");
}

#[test]
fn test_chitest_rejects_only_a_negative_total_not_negative_expected_values() {
    // A negative *expected* frequency is not an error on its own. Excel
    // divides by it and lets that term pull the statistic down, reporting
    // #NUM! only if the total comes out negative -- so an identical
    // negative expected value is fine in one series and fatal in another.
    // All three values are real Excel's.
    //
    //   A1:C1 = 1, 2, 3        A2:C2 = 5, -4, 3   -> chi2 = -5.8  -> #NUM!
    //   A3:C3 = -478.8, 352.51, 8.5
    //   A4:C4 = 38, 8.5, -75                      -> chi2 ~ 20859 -> 0
    let mut sheet = create_sheet(&[
        ["1", "2", "3", "=CHITEST(A1:C1, A2:C2)"],
        ["5", "-4", "3", "=CHITEST(A3:C3, A4:C4)"],
        ["-478.8", "352.51", "8.5", "=CHITEST(A1:C1, A5:C5)"],
        ["38", "8.5", "-75", ""],
        ["5", "0", "3", ""],
    ]);
    sheet.commit(None).unwrap();

    match sheet.get_result_data(&CellRef::new(0, 3)) {
        ResultData::Error(e) => assert_eq!(e, "#NUM!", "negative statistic is #NUM!"),
        other => panic!("expected #NUM!, got {other:?}"),
    }
    assert_float_close(&sheet.get_result_data(&CellRef::new(1, 3)), 0.0, 1e-12);
    // An expected frequency of exactly zero is the division failing.
    match sheet.get_result_data(&CellRef::new(2, 3)) {
        ResultData::Error(e) => assert_eq!(e, "#DIV/0!", "zero expected is #DIV/0!"),
        other => panic!("expected #DIV/0!, got {other:?}"),
    }
}

#[test]
fn test_normal_cdf_keeps_its_left_tail() {
    // Computed via erfc rather than 0.5 * (1 + erf(x/sqrt(2))), which
    // cancels catastrophically once erf approaches -1 and eventually
    // rounds to exactly 0 -- NORM.S.DIST(-11, TRUE) used to return 0, so
    // even SIGN() of it disagreed with Excel. Both reference values are
    // real Excel's, and it resolves the tail well past -30.
    assert_float_close(
        &eval1("=NORM.S.DIST(-11, TRUE)"),
        1.9106595744986622e-28,
        1e-40,
    );
    assert_float_close(
        &eval1("=NORM.S.DIST(-30, TRUE)"),
        4.9067139271479094e-198,
        1e-210,
    );
    assert_float_close(&eval1("=SIGN(NORM.S.DIST(-11, TRUE))"), 1.0, 1e-12);
    // The body of the distribution is unchanged.
    assert_float_close(&eval1("=NORM.S.DIST(0, TRUE)"), 0.5, 1e-15);
    assert_float_close(
        &eval1("=NORM.S.DIST(1.96, TRUE)"),
        0.9750021048517795,
        1e-15,
    );
}

#[test]
fn test_paired_sums_error_only_when_a_range_holds_no_numbers() {
    // The rule is *not* "no pair survived exclusion" -- that is simply 0.
    // Real Excel reports #DIV/0! when one of the ranges contains no
    // numeric value at all, and otherwise computes over whatever pairs
    // survive. The two are easy to confuse because they usually coincide:
    //
    //   [53, TRUE] vs [TRUE, -10]   every pair dropped, yet the answer is
    //                               0 -- each range does hold a number
    //   [1, 2]     vs ["a", "b"]    #DIV/0! -- the second holds none
    //   [-116.9395, 53] vs [TRUE, -10]  = 2909, i.e. 53^2 + (-10)^2
    //
    // All values below are real Excel's.
    let mut sheet = create_sheet(&[
        // A       B         C      D        E       F
        ["-116.9395", "53", "=TRUE", "=\"I3w\"", "=TRUE", "-10"],
        ["1", "2", "=\"a\"", "=\"b\"", "=TRUE", "=TRUE"],
        [
            "=SUMX2PY2(A1:C1, D1:F1)",
            "=SUMX2PY2(B1:C1, E1:F1)",
            "=SUMX2PY2(A1:B1, E1:F1)",
            "=SUMX2PY2(A2:B2, C2:D2)",
            "=SUMX2PY2(A2:B2, E2:F2)",
            "=SUMXMY2(A1:C1, D1:F1)",
        ],
    ]);
    sheet.commit(None).unwrap();

    // Every pair dropped, but both ranges hold a number: 0, not an error.
    assert_float_close(&sheet.get_result_data(&CellRef::new(2, 0)), 0.0, 1e-12);
    assert_float_close(&sheet.get_result_data(&CellRef::new(2, 1)), 0.0, 1e-12);
    assert_float_close(&sheet.get_result_data(&CellRef::new(2, 5)), 0.0, 1e-12);
    // One boolean pair dropped, one numeric pair kept.
    assert_float_close(&sheet.get_result_data(&CellRef::new(2, 2)), 2909.0, 1e-9);
    // A range with nothing numeric in it at all.
    for col in [3, 4] {
        match sheet.get_result_data(&CellRef::new(2, col)) {
            ResultData::Error(e) => assert_eq!(e, "#DIV/0!", "column {col}"),
            other => panic!("expected #DIV/0! in column {col}, got {other:?}"),
        }
    }
    // CORREL over a series with no variance is #DIV/0! independently.
    assert_err("=CORREL(A1:C1, D1:F1)", "#DIV/0!");
}

#[test]
// The mpmath reference values below are kept at their full width rather than
// truncated to the shortest round-tripping literal: the digits past f64 are
// what the sub-ULP tolerances here are justified against.
#[allow(clippy::excessive_precision)]
fn test_f_right_tail_avoids_cancellation_and_fisherinv_saturates() {
    // F.DIST.RT used to be computed as 1 - CDF. For a large F statistic
    // the CDF is within an ULP or two of 1, so that subtraction discarded
    // most of the answer's digits; it now goes through the incomplete
    // beta's symmetry instead. All reference values are real Excel's, and
    // everything below agrees with it to better than 2e-13 relative.
    // Reference values here are 40-digit mpmath evaluations of the
    // regularized incomplete beta, not Excel's. visi is closer to the
    // truth than Excel on this first one: 3.5e-16 relative against
    // Excel's 1.3e-14.
    assert_float_close(
        &eval1("=F.DIST.RT(120.02429320013077, 2, 4)"),
        2.6863796553017013481e-4,
        // visi is off by 1.1e-19 here; Excel's own answer is off by
        // 3.5e-18, so this still asserts a comfortable margin over it.
        1e-18,
    );
    assert_float_close(
        &eval1("=F.DIST.RT(1000000, 2, 4)"),
        3.9999840000480035e-12,
        1e-24,
    );
    assert_float_close(&eval1("=F.DIST.RT(2, 3, 7)"), 0.20269364248665092207, 1e-15);
    assert_float_close(&eval1("=F.DIST.RT(0.5, 10, 20)"), 0.8701603741696, 1e-12);
    assert_float_close(&eval1("=F.DIST.RT(1, 5, 5)"), 0.4999999999999999, 1e-13);
    assert_float_close(
        &eval1("=FDIST(4.28, 3, 10)"),
        0.034670525913903016847,
        1e-16,
    );
    // The left tail is unaffected.
    assert_float_close(&eval1("=F.DIST(2, 3, 7, TRUE)"), 0.7973063575133491, 1e-13);

    // FISHERINV is tanh; the (e^2y - 1)/(e^2y + 1) spelling overflowed to
    // inf/inf past y ~ 355 and reported #NUM! where Excel reports 1.
    assert_float_close(&eval1("=FISHERINV(1000)"), 1.0, 1e-15);
    assert_float_close(&eval1("=FISHERINV(-1000)"), -1.0, 1e-15);
    assert_float_close(&eval1("=FISHERINV(0.5)"), 0.46211715726000974, 1e-15);
}

#[test]
fn test_chitest_takes_degrees_of_freedom_from_the_raw_range_size() {
    // CHITEST drops pairs where either side is non-numeric, but takes the
    // degrees of freedom from the ranges' *original* size. With one text
    // cell in a two-cell pair, one pair survives and Excel still evaluates
    // against df = 1 -- using the survivor count would give df = 0 and
    // #NUM!, which is what visi used to return.
    //
    // Reference values are real Excel's. The second also exercises the
    // right tail: computing it as 1 - CDF underflowed to exactly 0.
    let mut sheet = create_sheet(&[
        ["-70", "=\"zz\"", "8.6291", "309.431", "3", "4"],
        [
            "=CHITEST(A1:B1, C1:D1)",
            "=CHITEST(A1:C1, D1:F1)",
            "=CHITEST(A1:A1, C1:C1)",
            "",
            "",
            "",
        ],
    ]);
    sheet.commit(None).unwrap();
    assert_float_close(
        &sheet.get_result_data(&CellRef::new(1, 0)),
        7.81883827261815e-158,
        1e-170,
    );
    assert_float_close(
        &sheet.get_result_data(&CellRef::new(1, 1)),
        6.38808797549415e-103,
        1e-115,
    );
    // A single category leaves no degrees of freedom at all.
    match sheet.get_result_data(&CellRef::new(1, 2)) {
        ResultData::Error(e) => assert_eq!(e, "#N/A"),
        other => panic!("expected #N/A, got {other:?}"),
    }
}

#[test]
fn test_a_lone_blank_cell_is_a_missing_operand() {
    // Excel distinguishes one blank cell from an array of blanks:
    //   SUMPRODUCT(<one blank cell>)   = #VALUE!   (missing operand)
    //   SUMPRODUCT(<two blank cells>)  = 0
    //   SUMPRODUCT(-50, <blank>)       = #VALUE!
    //   SUMPRODUCT(<one text cell>)    = 0         (text is not blank)
    // MULTINOMIAL draws the line in a different place: a blank operand
    // is only *missing* when there is nothing else, so MULTINOMIAL(3,
    // <blank>) is 1 with the blank counting as 0, while SUMPRODUCT rejects
    // a lone blank even beside a number. Z50/Z51 are empty in a fresh
    // sheet.
    for src in [
        "=SUMPRODUCT(Z50:Z50)",
        "=SUMPRODUCT(-50, Z50)",
        "=MULTINOMIAL(Z50)",
        "=MULTINOMIAL(Z50, Z51)",
    ] {
        match eval1(src) {
            ResultData::Error(e) => assert_eq!(e, "#VALUE!", "for {src}"),
            other => panic!("expected #VALUE! for {src}, got {other:?}"),
        }
    }
    assert_float_close(&eval1("=SUMPRODUCT(Z50:Z51)"), 0.0, 1e-12);
    assert_float_close(&eval1("=MULTINOMIAL(3, Z50)"), 1.0, 1e-12);
    assert_float_close(&eval1("=MULTINOMIAL(Z50, 3)"), 1.0, 1e-12);

    let mut sheet = create_sheet(&[["=\"abc\"", "=SUMPRODUCT(A1:A1)"]]);
    sheet.commit(None).unwrap();
    assert_float_close(&sheet.get_result_data(&CellRef::new(0, 1)), 0.0, 1e-12);
}

#[test]
fn test_gamma_keeps_full_precision_at_integer_arguments() {
    // GAMMA(34) is exactly 33! = 8683317618811886495518194401280000000,
    // which Excel displays as 8.68331761881189E+36. Computing it as
    // exp(lgamma(x)) costs several significant digits and gave
    // 8.68331761881199E+36 -- wrong from the 14th.
    assert_float_close(&eval1("=GAMMA(34)"), 8.68331761881189e36, 1e24);
    assert_float_close(&eval1("=GAMMA(5)"), 24.0, 1e-12);
    assert_float_close(&eval1("=GAMMA(11)"), 3628800.0, 1e-6);
    // Non-integer and negative arguments are unchanged.
    assert_float_close(&eval1("=GAMMA(0.5)"), 1.7724538509055159, 1e-15);
    assert_float_close(&eval1("=GAMMA(-1.5)"), 2.3632718012073544, 1e-14);
}

#[test]
fn test_chitest_with_no_surviving_pair_is_one_not_not_available() {
    // Every pair holds something non-numeric, so the statistic is 0 and
    // -- with the degrees of freedom taken from the raw range size -- the
    // p-value is 1. Real Excel returns 1 here; visi reported #N/A, which
    // then propagated (ERF(CHITEST(...)) should be erf(1)).
    let mut sheet = create_sheet(&[
        ["=\"rBN\"", "-323.7702", "=CHITEST(A1:A3, B1:B3)"],
        ["", "=\"6-323.7702\"", "=ERF(CHITEST(A1:A3, B1:B3))"],
        ["27", "=\"B\"", ""],
    ]);
    sheet.commit(None).unwrap();
    assert_float_close(&sheet.get_result_data(&CellRef::new(0, 2)), 1.0, 1e-12);
    assert_float_close(
        &sheet.get_result_data(&CellRef::new(1, 2)),
        0.8427007929497149,
        1e-15,
    );

    // A range holding no number at all is a different case, and is
    // #DIV/0! -- the same rule the paired sums use. Above, each range
    // still had one number in it (27 and -323.7702).
    let mut sheet = create_sheet(&[
        ["=\"cUVCpj\"", "=TRUE", "-60.63", "=\"XDWK\""],
        ["=\"mpSHAC\"", "", "-331.95", "=TRUE"],
        ["=CHITEST(A1:A2, B1:B2)", "=CHITEST(C1:C2, D1:D2)", "", ""],
    ]);
    sheet.commit(None).unwrap();
    for col in 0..2 {
        match sheet.get_result_data(&CellRef::new(2, col)) {
            ResultData::Error(e) => assert_eq!(e, "#DIV/0!", "column {col}"),
            other => panic!("expected #DIV/0! in column {col}, got {other:?}"),
        }
    }
}

#[test]
fn test_mode_family_rejects_a_lone_blank_operand() {
    // MODE is stricter than its neighbours about a blank operand:
    // MODE(x, <blank>) is #VALUE! in real Excel while MEDIAN(x, <blank>)
    // is just x. All three spellings behave the same way.
    for src in [
        "=MODE(241.965, Z90)",
        "=MODE.SNGL(241.965, Z90)",
        "=MODE.MULT(241.965, Z90)",
        "=MODE(241.965, 241.965, Z90)",
    ] {
        match eval1(src) {
            ResultData::Error(e) => assert_eq!(e, "#VALUE!", "for {src}"),
            other => panic!("expected #VALUE! for {src}, got {other:?}"),
        }
    }
    // Unchanged neighbours, and MODE's ordinary behaviour.
    assert_float_close(&eval1("=MEDIAN(241.965, Z90)"), 241.965, 1e-12);
    assert_float_close(&eval1("=MODE(241.965, 241.965)"), 241.965, 1e-12);
    assert_float_close(&eval1("=MODE(1, 1, 2)"), 1.0, 1e-12);
    // No repeated value is still #N/A, not #VALUE!.
    assert_err("=MODE(1, 2)", "#N/A");
}

#[test]
fn test_fuzz_ifna_sees_left_hand_shape_error_before_right_hand_value_error() {
    // Harvested from fuzz/fuzz_excel.py seed 984916. The arithmetic operator
    // surfaces the left #N/A from SUMXMY2's shape mismatch before evaluating a
    // later POWER argument that would be #VALUE!, so IFNA catches it.
    assert_float_close(
        &eval1("=IFNA((SUMXMY2(A1:B3,C1:C2)+POWER(\"x\",-97)),0)"),
        0.0,
        1e-12,
    );

    // Harvested from fuzz/fuzz_excel.py seed 965229. SLOPE's shape mismatch
    // supplies LOG's optional base argument as #N/A; LOG must propagate that
    // optional-argument error instead of turning it into #VALUE!, so IFNA
    // catches it and returns the fallback.
    let mut sheet = create_sheet(&[
        [
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "-121.667",
            "=IFNA(LOG(ATAN2(PI(), J2), SLOPE(J1:J$4, $H2:J2)), F3)",
        ],
        ["", "", "", "", "", "", "", "FALSE", "c", "24", ""],
        ["", "", "", "", "", "pi", "", "", "", "18", ""],
        ["", "", "", "", "", "", "", "", "", "", ""],
    ]);
    sheet.commit(None).unwrap();
    assert_eq!(sheet.get_display_string(&CellRef::new(0, 10)), "pi");
}

#[test]
fn test_shape_mismatch_outranks_the_no_numbers_rule() {
    // A shape mismatch is #N/A and wins over everything else, including a
    // range that also holds no numeric value at all. Getting the order
    // wrong matters beyond the error class: IFNA/ISNA are watching for
    // exactly #N/A, so a spurious #DIV/0! escapes them.
    let mut sheet = create_sheet(&[
        ["1", "=\"x\"", "5", "=SUMXMY2(A1:A3, B1:B2)"],
        ["2", "=\"y\"", "6", "=SUMXMY2(A1:A3, C1:C2)"],
        ["3", "", "", "=CHITEST(A1:A1, A1:B5)"],
    ]);
    sheet.commit(None).unwrap();
    for row in 0..3 {
        match sheet.get_result_data(&CellRef::new(row, 3)) {
            ResultData::Error(e) => assert_eq!(e, "#N/A", "row {row}"),
            other => panic!("expected #N/A in row {row}, got {other:?}"),
        }
    }
}

#[test]
fn test_fuzz_chitest_mismatched_range_with_no_numbers_is_value() {
    // Harvested from fuzz/fuzz_excel.py seeds 409614 and 800412. A pure
    // numeric shape mismatch remains #N/A, but if one of CHITEST's mismatched
    // ranges has no numeric value at all, real Excel reports #VALUE! instead.
    let mut sheet = create_sheet(&[
        ["", "", "0", "=CHITEST(A1:A1, B1:C1)"],
        ["90", "90.1", "", "=CHITEST(A2:B2, C2:C2)"],
        ["90", "90.1", "-52", "=CHITEST(A3:B3, C3:C3)"],
        ["90", "90.1", "=TRUE", "=CHITEST(A4:B4, C4:C4)"],
    ]);
    sheet.commit(None).unwrap();

    for row in 0..2 {
        match sheet.get_result_data(&CellRef::new(row, 3)) {
            ResultData::Error(e) => assert_eq!(e, "#VALUE!", "row {row}"),
            other => panic!("expected #VALUE! in row {row}, got {other:?}"),
        }
    }
    for row in 2..=3 {
        match sheet.get_result_data(&CellRef::new(row, 3)) {
            ResultData::Error(e) => assert_eq!(e, "#N/A", "row {row}"),
            other => panic!("expected #N/A in row {row}, got {other:?}"),
        }
    }
}

#[test]
fn test_gcd_family_coerces_numeric_text_but_not_booleans() {
    // GCD, LCM and MULTINOMIAL coerce text that looks numeric and reject
    // everything else -- so this is narrower than "text is #VALUE!", which
    // is what it used to do. All values are real Excel's.
    assert_float_close(&eval1("=GCD(\"12\", 8)"), 4.0, 1e-12);
    assert_float_close(&eval1("=LCM(\"4\", 6)"), 12.0, 1e-12);
    assert_float_close(&eval1("=MULTINOMIAL(\"3\", 2)"), 10.0, 1e-9);
    assert_float_close(&eval1("=MULTINOMIAL(RIGHT(\"a5\", 1), 2)"), 21.0, 1e-9);
    // Excel computes these slightly below the exact integer, so
    // INT(MULTINOMIAL(0.1,40)) becomes 0 there. The combinatorial result is
    // exactly 1, and visi keeps it.
    assert_float_close(&eval1("=INT(MULTINOMIAL(0.1, 40))"), 1.0, 1e-12);
    for src in [
        "=GCD(\"x\", 8)",
        "=GCD(TRUE, 8)",
        "=LCM(\"x\", 6)",
        "=MULTINOMIAL(\"x\", 2)",
    ] {
        assert_err(src, "#VALUE!");
    }
}

#[test]
// As above: the 60-digit mpmath references stay at full width so the ULP
// claims in the comments can be checked against them.
#[allow(clippy::excessive_precision)]
fn test_incomplete_beta_prefactor_accuracy() {
    // The beta prefactor is computed from tgamma rather than as
    // exp(a*ln x + b*ln(1-x) - lbeta), which put the absolute error of a
    // logarithm straight into the relative error of the result. All
    // expected values are 50-digit mpmath evaluations of the regularized
    // incomplete beta, not Excel's.
    //
    // The FTEST case is the one this was chased down for: the true value
    // is 0.94171633283387507291, which renders at 15 digits as
    // 0.941716332833875. visi used to be ~10 ULP high and print ...876.
    let mut sheet = create_sheet(&[
        ["127.95", "127.95"],
        ["5", "5"],
        ["28", "28"],
        ["-24.3108", "92"],
        ["0", "=FTEST(A1:A11, B1:B4)"],
        ["-40", ""],
        ["-45", ""],
        ["43", ""],
        ["96", ""],
        ["-66", ""],
        ["1", ""],
    ]);
    sheet.commit(None).unwrap();
    let got = match sheet.get_result_data(&CellRef::new(4, 1)) {
        ResultData::Float(f) => f,
        other => panic!("expected a number, got {other:?}"),
    };
    assert!(
        (got - 0.94171633283387507291).abs() < 3e-16,
        "FTEST expected 0.94171633283387507291, got {got}"
    );
    assert_eq!(format!("{got:.15}"), "0.941716332833875");

    // The (1-x)^b correction: without it, a half-ULP rounding of `1 - x`
    // is multiplied by b. This case (a=5, b=50) was 15 ULP out and is now
    // within 1. Reference from 60-digit mpmath.
    assert_float_close(
        &eval1("=BETA.DIST(0.0378, 5, 50, TRUE)"),
        0.052899172535742447319,
        3e-17,
    );

    // Spot checks across the parameter space, all within ~2 ULP.
    assert_float_close(&eval1("=F.DIST.RT(0.5, 10, 20)"), 0.8701603741696, 1e-15);
    assert_float_close(&eval1("=BETA.DIST(0.5, 2, 3, TRUE)"), 0.6875, 1e-15);
    assert_float_close(
        &eval1("=T.DIST(1.5, 10, TRUE)"),
        0.91774633677727990958,
        1e-15,
    );
}
