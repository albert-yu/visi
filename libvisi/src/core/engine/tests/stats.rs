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
