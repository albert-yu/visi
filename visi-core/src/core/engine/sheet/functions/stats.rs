//! Statistical function dispatch.
//!
//! Split out of the parent module's `evaluate_function`, which tries each
//! family in turn.

use super::{FnCall, res_to_rd};
use crate::core::engine::cell::{Dependency, EngineError};
use crate::core::engine::result_data::ResultData;
use crate::core::engine::sheet::Sheet;

impl Sheet {
    /// Evaluates `call` if this family owns its name, else `None`.
    pub(super) fn eval_stats_fn(
        &self,
        call: FnCall<'_>,
        deps: &mut Vec<Dependency>,
    ) -> Option<Result<ResultData, EngineError>> {
        // The body returns `Result` so its arms can keep using `?`; whether
        // the name belongs to this family is signalled alongside.
        let mut owned = true;
        let r = self.eval_stats_dispatch(call, deps, &mut owned);
        owned.then_some(r)
    }

    fn eval_stats_dispatch(
        &self,
        call: FnCall<'_>,
        _deps: &mut Vec<Dependency>,
        owned: &mut bool,
    ) -> Result<ResultData, EngineError> {
        let FnCall {
            upper_name,
            evaluated_args,
            arg_is_direct,
            ..
        } = call;
        match call.upper_name {
            // --- STATISTICAL FUNCTIONS ---
            "AVEDEV" => {
                let nums: Vec<f64> =
                    match self.flatten_args_stat_numbers(evaluated_args, arg_is_direct) {
                        Ok(v) => v,
                        Err(e) => return Ok(ResultData::Error(e)),
                    };
                res_to_rd(crate::core::stats::avedev(&nums))
            }
            "AVERAGEA" => {
                let nums: Vec<f64> =
                    match self.flatten_args_stat_numbers_a(evaluated_args, arg_is_direct) {
                        Ok(v) => v,
                        Err(e) => return Ok(ResultData::Error(e)),
                    };
                if nums.is_empty() {
                    Ok(ResultData::Error("#DIV/0!".to_string()))
                } else {
                    Ok(ResultData::Float(
                        nums.iter().sum::<f64>() / nums.len() as f64,
                    ))
                }
            }
            "AVERAGEIF" => {
                if evaluated_args.len() < 2 {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                }
                let range_list = match &evaluated_args[0] {
                    ResultData::List(l) => l,
                    _ => return Ok(ResultData::Error("#DIV/0!".to_string())),
                };
                let criteria = &evaluated_args[1];
                let avg_range = if evaluated_args.len() >= 3 {
                    match &evaluated_args[2] {
                        ResultData::List(l) => l,
                        _ => range_list,
                    }
                } else {
                    range_list
                };
                let mut sum = 0.0;
                let mut count = 0;
                for (i, val) in range_list.iter().enumerate() {
                    if self.match_criteria(val, criteria)
                        && let Some(target_val) = avg_range.get(i)
                        && let Some(f) = Self::aggregate_range_number(target_val)
                    {
                        sum += f;
                        count += 1;
                    }
                }
                if count == 0 {
                    Ok(ResultData::Error("#DIV/0!".to_string()))
                } else {
                    Ok(ResultData::Float(sum / count as f64))
                }
            }
            "AVERAGEIFS" => {
                if evaluated_args.len() < 3 || (evaluated_args.len() - 1) % 2 != 0 {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                }
                let avg_range = match &evaluated_args[0] {
                    ResultData::List(l) => l,
                    _ => return Ok(ResultData::Error("#DIV/0!".to_string())),
                };
                let mut criteria_pairs = Vec::new();
                let mut i = 1;
                while i < evaluated_args.len() {
                    let crit_range = match &evaluated_args[i] {
                        ResultData::List(l) => l,
                        _ => return Ok(ResultData::Error("#DIV/0!".to_string())),
                    };
                    let crit_val = &evaluated_args[i + 1];
                    criteria_pairs.push((crit_range, crit_val));
                    i += 2;
                }
                let mut sum = 0.0;
                let mut count = 0;
                for (idx, target_val) in avg_range.iter().enumerate() {
                    let mut all_match = true;
                    for (crit_range, crit_val) in &criteria_pairs {
                        if idx >= crit_range.len()
                            || !self.match_criteria(&crit_range[idx], crit_val)
                        {
                            all_match = false;
                            break;
                        }
                    }
                    if all_match && let Some(f) = Self::aggregate_range_number(target_val) {
                        sum += f;
                        count += 1;
                    }
                }
                if count == 0 {
                    Ok(ResultData::Error("#DIV/0!".to_string()))
                } else {
                    Ok(ResultData::Float(sum / count as f64))
                }
            }
            "BETA.DIST" | "BETADIST" => {
                let x = self.to_f64_arg(evaluated_args.first(), "BETA.DIST")?;
                let alpha = self.to_f64_arg(evaluated_args.get(1), "BETA.DIST")?;
                let beta = self.to_f64_arg(evaluated_args.get(2), "BETA.DIST")?;
                let cumulative = evaluated_args
                    .get(3)
                    .map(|v| self.to_bool(v))
                    .unwrap_or(true);
                let a = evaluated_args
                    .get(4)
                    .and_then(|v| self.to_f64(v))
                    .unwrap_or(0.0);
                let b = evaluated_args
                    .get(5)
                    .and_then(|v| self.to_f64(v))
                    .unwrap_or(1.0);
                res_to_rd(crate::core::stats::beta_dist(
                    x, alpha, beta, cumulative, a, b,
                ))
            }
            "BETA.INV" | "BETAINV" => {
                let p = self.to_f64_arg(evaluated_args.first(), "BETA.INV")?;
                let alpha = self.to_f64_arg(evaluated_args.get(1), "BETA.INV")?;
                let beta = self.to_f64_arg(evaluated_args.get(2), "BETA.INV")?;
                let a = evaluated_args
                    .get(3)
                    .and_then(|v| self.to_f64(v))
                    .unwrap_or(0.0);
                let b = evaluated_args
                    .get(4)
                    .and_then(|v| self.to_f64(v))
                    .unwrap_or(1.0);
                res_to_rd(crate::core::stats::beta_inv(p, alpha, beta, a, b))
            }
            "BINOM.DIST" | "BINOMDIST" => {
                let k = self.to_f64_arg(evaluated_args.first(), "BINOM.DIST")?;
                let n = self.to_f64_arg(evaluated_args.get(1), "BINOM.DIST")?;
                let p = self.to_f64_arg(evaluated_args.get(2), "BINOM.DIST")?;
                let cumulative = evaluated_args
                    .get(3)
                    .map(|v| self.to_bool(v))
                    .unwrap_or(false);
                res_to_rd(crate::core::stats::binom_dist(k, n, p, cumulative))
            }
            "BINOM.DIST.RANGE" => {
                let n = self.to_f64_arg(evaluated_args.first(), "BINOM.DIST.RANGE")?;
                let p = self.to_f64_arg(evaluated_args.get(1), "BINOM.DIST.RANGE")?;
                let k1 = self.to_f64_arg(evaluated_args.get(2), "BINOM.DIST.RANGE")?;
                let k2 = evaluated_args.get(3).and_then(|v| self.to_f64(v));
                res_to_rd(crate::core::stats::binom_dist_range(n, p, k1, k2))
            }
            "BINOM.INV" | "CRITBINOM" => {
                let n = self.to_f64_arg(evaluated_args.first(), "BINOM.INV")?;
                let p = self.to_f64_arg(evaluated_args.get(1), "BINOM.INV")?;
                let alpha = self.to_f64_arg(evaluated_args.get(2), "BINOM.INV")?;
                res_to_rd(crate::core::stats::binom_inv(n, p, alpha))
            }
            "CHISQ.DIST" => {
                let x = self.to_f64_arg(evaluated_args.first(), "CHISQ.DIST")?;
                let df = self.to_f64_arg(evaluated_args.get(1), "CHISQ.DIST")?;
                let cumulative = evaluated_args
                    .get(2)
                    .map(|v| self.to_bool(v))
                    .unwrap_or(true);
                res_to_rd(crate::core::stats::chisq_dist(x, df, cumulative))
            }
            "CHISQ.DIST.RT" | "CHIDIST" => {
                let x = self.to_f64_arg(evaluated_args.first(), "CHISQ.DIST.RT")?;
                let df = self.to_f64_arg(evaluated_args.get(1), "CHISQ.DIST.RT")?;
                res_to_rd(crate::core::stats::chisq_dist_rt(x, df))
            }
            "CHISQ.INV" => {
                let p = self.to_f64_arg(evaluated_args.first(), "CHISQ.INV")?;
                let df = self.to_f64_arg(evaluated_args.get(1), "CHISQ.INV")?;
                res_to_rd(crate::core::stats::chisq_inv(p, df))
            }
            "CHISQ.INV.RT" | "CHIINV" => {
                let p = self.to_f64_arg(evaluated_args.first(), "CHISQ.INV.RT")?;
                let df = self.to_f64_arg(evaluated_args.get(1), "CHISQ.INV.RT")?;
                res_to_rd(crate::core::stats::chisq_inv_rt(p, df))
            }
            "CHISQ.TEST" | "CHITEST" => {
                // Like the paired statistical functions, CHITEST
                // compares its two ranges' *raw* cell counts first --
                // a mismatch is #N/A even when a range also holds an
                // error value. It does not, however, pairwise-exclude
                // the way CORREL and friends do (Excel keeps the
                // original dimensions when working out the degrees of
                // freedom), so the values themselves still come from
                // the lenient flatten.
                let mut first_err = None;
                let a_raw = self.positional_numbers(evaluated_args.first(), &mut first_err);
                let e_raw = self.positional_numbers(evaluated_args.get(1), &mut first_err);
                if a_raw.len() != e_raw.len() {
                    // A pure shape mismatch is #N/A, but a one-cell blank
                    // operand is missing and reports #VALUE! even when the
                    // other range has a different shape. One-cell text/boolean
                    // operands are still shape mismatches (#N/A).
                    if Self::is_empty_scalar_operand(
                        evaluated_args.first().unwrap_or(&ResultData::None),
                    ) || Self::is_empty_scalar_operand(
                        evaluated_args.get(1).unwrap_or(&ResultData::None),
                    ) {
                        return Ok(ResultData::Error("#VALUE!".to_string()));
                    }
                    return Ok(ResultData::Error("#N/A".to_string()));
                }
                // A single category leaves zero degrees of freedom, so
                // there is no chi-square distribution to evaluate
                // against and Excel reports #N/A. Judged on the *raw*
                // range size: applying it after pairwise filtering
                // would turn a two-cell pair that merely holds one text
                // cell into #N/A, where Excel still reports the
                // underlying #DIV/0!.
                if a_raw.len() < 2 {
                    return Ok(ResultData::Error("#N/A".to_string()));
                }
                if let Some(e) = first_err {
                    return Ok(ResultData::Error(e));
                }
                // Same shape as the paired sums, and checked only after
                // the #N/A cases above: a range holding no numeric
                // value at all is #DIV/0!, while a range that merely
                // loses every *pair* to exclusion still computes -- the
                // statistic is 0, so the answer is 1.
                if self.paired_sum_has_no_numbers(evaluated_args.first())
                    || self.paired_sum_has_no_numbers(evaluated_args.get(1))
                {
                    return Ok(ResultData::Error("#DIV/0!".to_string()));
                }
                // Values are taken pairwise so a non-numeric cell in
                // one range can't leave the two sides different lengths
                // and turn a computable call into a spurious #N/A --
                // Excel still returns a value there (CHITEST over a
                // 2-cell pair whose expected range holds one text cell
                // computes rather than failing).
                let (actual, expected) =
                    match self.paired_args(evaluated_args.first(), evaluated_args.get(1)) {
                        Ok(v) => v,
                        Err(e) => return Ok(ResultData::Error(e)),
                    };
                // Degrees of freedom come from the raw range size, not
                // from how many pairs survived the filtering above.
                res_to_rd(crate::core::stats::chisq_test(
                    &actual,
                    &expected,
                    a_raw.len(),
                ))
            }
            "CONFIDENCE.NORM" | "CONFIDENCE" => {
                let alpha = self.to_f64_arg(evaluated_args.first(), "CONFIDENCE.NORM")?;
                let std_dev = self.to_f64_arg(evaluated_args.get(1), "CONFIDENCE.NORM")?;
                let size = self.to_f64_arg(evaluated_args.get(2), "CONFIDENCE.NORM")?;
                res_to_rd(crate::core::stats::confidence_norm(alpha, std_dev, size))
            }
            "CONFIDENCE.T" => {
                let alpha = self.to_f64_arg(evaluated_args.first(), "CONFIDENCE.T")?;
                let std_dev = self.to_f64_arg(evaluated_args.get(1), "CONFIDENCE.T")?;
                let size = self.to_f64_arg(evaluated_args.get(2), "CONFIDENCE.T")?;
                res_to_rd(crate::core::stats::confidence_t(alpha, std_dev, size))
            }
            "CORREL" | "PEARSON" => {
                let (xs, ys) = match self.paired_args(evaluated_args.first(), evaluated_args.get(1))
                {
                    Ok(v) => v,
                    Err(e) => return Ok(ResultData::Error(e)),
                };
                res_to_rd(crate::core::stats::correl(&xs, &ys))
            }
            "COUNTBLANK" => {
                let mut count = 0;
                fn count_blank_rec(arg: &ResultData) -> usize {
                    match arg {
                        ResultData::None => 1,
                        ResultData::String(s) if s.is_empty() => 1,
                        ResultData::List(list) => list.iter().map(count_blank_rec).sum(),
                        _ => 0,
                    }
                }
                for arg in evaluated_args {
                    count += count_blank_rec(arg);
                }
                Ok(ResultData::Float(count as f64))
            }
            "COVARIANCE.P" | "COVAR" => {
                let (xs, ys) = match self.paired_args(evaluated_args.first(), evaluated_args.get(1))
                {
                    Ok(v) => v,
                    Err(e) => return Ok(ResultData::Error(e)),
                };
                res_to_rd(crate::core::stats::covariance_p(&xs, &ys))
            }
            "COVARIANCE.S" => {
                let (xs, ys) = match self.paired_args(evaluated_args.first(), evaluated_args.get(1))
                {
                    Ok(v) => v,
                    Err(e) => return Ok(ResultData::Error(e)),
                };
                res_to_rd(crate::core::stats::covariance_s(&xs, &ys))
            }
            "DEVSQ" => {
                let nums: Vec<f64> =
                    match self.flatten_args_stat_numbers(evaluated_args, arg_is_direct) {
                        Ok(v) => v,
                        Err(e) => return Ok(ResultData::Error(e)),
                    };
                res_to_rd(crate::core::stats::devsq(&nums))
            }
            "EXPON.DIST" | "EXPONDIST" => {
                let x = self.to_f64_arg(evaluated_args.first(), "EXPON.DIST")?;
                let lambda = self.to_f64_arg(evaluated_args.get(1), "EXPON.DIST")?;
                let cumulative = evaluated_args
                    .get(2)
                    .map(|v| self.to_bool(v))
                    .unwrap_or(true);
                res_to_rd(crate::core::stats::expon_dist(x, lambda, cumulative))
            }
            "F.DIST" => {
                let x = self.to_f64_arg(evaluated_args.first(), "F.DIST")?;
                let df1 = self.to_f64_arg(evaluated_args.get(1), "F.DIST")?;
                let df2 = self.to_f64_arg(evaluated_args.get(2), "F.DIST")?;
                let cumulative = evaluated_args
                    .get(3)
                    .map(|v| self.to_bool(v))
                    .unwrap_or(true);
                res_to_rd(crate::core::stats::f_dist(x, df1, df2, cumulative))
            }
            "F.DIST.RT" | "FDIST" => {
                let x = self.to_f64_arg(evaluated_args.first(), "F.DIST.RT")?;
                let df1 = self.to_f64_arg(evaluated_args.get(1), "F.DIST.RT")?;
                let df2 = self.to_f64_arg(evaluated_args.get(2), "F.DIST.RT")?;
                res_to_rd(crate::core::stats::f_dist_rt(x, df1, df2))
            }
            "F.INV" => {
                let p = self.to_f64_arg(evaluated_args.first(), "F.INV")?;
                let df1 = self.to_f64_arg(evaluated_args.get(1), "F.INV")?;
                let df2 = self.to_f64_arg(evaluated_args.get(2), "F.INV")?;
                res_to_rd(crate::core::stats::f_inv(p, df1, df2))
            }
            "F.INV.RT" | "FINV" => {
                let p = self.to_f64_arg(evaluated_args.first(), "F.INV.RT")?;
                let df1 = self.to_f64_arg(evaluated_args.get(1), "F.INV.RT")?;
                let df2 = self.to_f64_arg(evaluated_args.get(2), "F.INV.RT")?;
                res_to_rd(crate::core::stats::f_inv_rt(p, df1, df2))
            }
            "F.TEST" | "FTEST" => {
                let array1: Vec<f64> = evaluated_args
                    .first()
                    .map(|arg| self.flatten_stat_numbers(arg, false))
                    .unwrap_or_default();
                let array2: Vec<f64> = evaluated_args
                    .get(1)
                    .map(|arg| self.flatten_stat_numbers(arg, false))
                    .unwrap_or_default();
                res_to_rd(crate::core::stats::f_test(&array1, &array2))
            }
            "FISHER" => {
                let x = self.to_f64_arg(evaluated_args.first(), "FISHER")?;
                res_to_rd(crate::core::stats::fisher(x))
            }
            "FISHERINV" => {
                let y = self.to_f64_arg(evaluated_args.first(), "FISHERINV")?;
                res_to_rd(crate::core::stats::fisherinv(y))
            }
            "FORECAST" | "FORECAST.LINEAR" => {
                let x = self.to_f64_arg(evaluated_args.first(), "FORECAST")?;
                let (ys, xs) = match self.paired_args(evaluated_args.get(1), evaluated_args.get(2))
                {
                    Ok(v) => v,
                    Err(e) => return Ok(ResultData::Error(e)),
                };
                res_to_rd(crate::core::stats::forecast_linear(x, &ys, &xs))
            }
            "FORECAST.ETS" | "FORECAST.ETS.CONFINT" => {
                // FORECAST.ETS(target, values, timeline,
                //              [seasonality], [data_completion], [aggregation])
                // FORECAST.ETS.CONFINT(target, values, timeline,
                //              [confidence], [seasonality], [data_completion], [aggregation])
                let is_confint = upper_name == "FORECAST.ETS.CONFINT";
                let target = self.to_f64_arg(evaluated_args.first(), "FORECAST.ETS")?;
                let (values, timeline) =
                    match self.paired_args(evaluated_args.get(1), evaluated_args.get(2)) {
                        Ok(v) => v,
                        Err(e) => return Ok(ResultData::Error(e)),
                    };
                let (confidence, seasonality_idx) = if is_confint {
                    (self.opt_f64_arg(evaluated_args, 3, 0.95)?, 4)
                } else {
                    (0.95, 3)
                };
                let seasonality = self.opt_f64_arg(evaluated_args, seasonality_idx, 1.0)?;
                let completion = self.opt_f64_arg(evaluated_args, seasonality_idx + 1, 1.0)? != 0.0;

                let series = match crate::core::ets::build_series(&values, &timeline, completion) {
                    Ok(s) => s,
                    Err(e) => return Ok(ResultData::Error(e)),
                };
                let h = match crate::core::ets::horizon(
                    series.start,
                    series.step,
                    series.values.len(),
                    target,
                ) {
                    Ok(h) => h,
                    Err(e) => return Ok(ResultData::Error(e)),
                };
                let model =
                    match crate::core::ets::prepare(&values, &timeline, seasonality, completion) {
                        Ok(m) => m,
                        Err(e) => return Ok(ResultData::Error(e)),
                    };
                if is_confint {
                    res_to_rd(model.confint(h, confidence))
                } else {
                    Ok(ResultData::Float(model.forecast(h)))
                }
            }
            "FORECAST.ETS.SEASONALITY" => {
                // FORECAST.ETS.SEASONALITY(values, timeline,
                //                          [data_completion], [aggregation])
                // -- note there is no leading target-date argument.
                let (values, timeline) =
                    match self.paired_args(evaluated_args.first(), evaluated_args.get(1)) {
                        Ok(v) => v,
                        Err(e) => return Ok(ResultData::Error(e)),
                    };
                let completion = self.opt_f64_arg(evaluated_args, 2, 1.0)? != 0.0;
                match crate::core::ets::build_series(&values, &timeline, completion) {
                    Ok(series) => Ok(ResultData::Float(crate::core::ets::detect_period(
                        &series.values,
                    ) as f64)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "FORECAST.ETS.STAT" => {
                // FORECAST.ETS.STAT(values, timeline, statistic_type,
                //                   [seasonality], [data_completion], [aggregation])
                let (values, timeline) =
                    match self.paired_args(evaluated_args.first(), evaluated_args.get(1)) {
                        Ok(v) => v,
                        Err(e) => return Ok(ResultData::Error(e)),
                    };
                let which = self.to_f64_arg(evaluated_args.get(2), "FORECAST.ETS.STAT")?;
                let seasonality = self.opt_f64_arg(evaluated_args, 3, 1.0)?;
                let completion = self.opt_f64_arg(evaluated_args, 4, 1.0)? != 0.0;
                match crate::core::ets::prepare(&values, &timeline, seasonality, completion) {
                    Ok(model) => res_to_rd(model.stat(which.round() as usize)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "FREQUENCY" => {
                let data: Vec<f64> = evaluated_args
                    .first()
                    .map(|arg| self.flatten_stat_numbers(arg, false))
                    .unwrap_or_default();
                let bins: Vec<f64> = evaluated_args
                    .get(1)
                    .map(|arg| self.flatten_stat_numbers(arg, false))
                    .unwrap_or_default();
                match crate::core::stats::frequency(&data, &bins) {
                    Ok(counts) => Ok(ResultData::List(
                        counts.into_iter().map(ResultData::Float).collect(),
                    )),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "GAMMA" => {
                let x = self.to_f64_arg(evaluated_args.first(), "GAMMA")?;
                let val = crate::core::stats::gamma(x);
                if val.is_nan() {
                    Ok(ResultData::Error("#NUM!".to_string()))
                } else {
                    Ok(ResultData::Float(val))
                }
            }
            "GAMMA.DIST" | "GAMMADIST" => {
                let x = self.to_f64_arg(evaluated_args.first(), "GAMMA.DIST")?;
                let alpha = self.to_f64_arg(evaluated_args.get(1), "GAMMA.DIST")?;
                let beta = self.to_f64_arg(evaluated_args.get(2), "GAMMA.DIST")?;
                let cumulative = evaluated_args
                    .get(3)
                    .map(|v| self.to_bool(v))
                    .unwrap_or(true);
                res_to_rd(crate::core::stats::gamma_dist(x, alpha, beta, cumulative))
            }
            "GAMMA.INV" | "GAMMAINV" => {
                let p = self.to_f64_arg(evaluated_args.first(), "GAMMA.INV")?;
                let alpha = self.to_f64_arg(evaluated_args.get(1), "GAMMA.INV")?;
                let beta = self.to_f64_arg(evaluated_args.get(2), "GAMMA.INV")?;
                res_to_rd(crate::core::stats::gamma_inv(p, alpha, beta))
            }
            "GAMMALN" | "GAMMALN.PRECISE" => {
                // ln(Gamma(x)) is only defined for x > 0 in Excel --
                // GAMMALN(-5), GAMMALN(0) and GAMMALN of a large
                // negative are all #NUM!. The underlying lgamma here
                // uses the reflection formula and happily returns a
                // value for negative non-integers, so the domain has to
                // be enforced at the boundary.
                let x = self.to_f64_arg(evaluated_args.first(), "GAMMALN")?;
                if x <= 0.0 {
                    return Ok(ResultData::Error("#NUM!".to_string()));
                }
                let val = crate::core::stats::lgamma(x);
                if val.is_nan() {
                    Ok(ResultData::Error("#NUM!".to_string()))
                } else {
                    Ok(ResultData::Float(val))
                }
            }
            "GAUSS" => {
                let z = self.to_f64_arg(evaluated_args.first(), "GAUSS")?;
                res_to_rd(crate::core::stats::gauss(z))
            }
            "GEOMEAN" => {
                let nums: Vec<f64> =
                    match self.flatten_args_stat_numbers(evaluated_args, arg_is_direct) {
                        Ok(v) => v,
                        Err(e) => return Ok(ResultData::Error(e)),
                    };
                res_to_rd(crate::core::stats::geomean(&nums))
            }
            "GROWTH" | "LOGEST" => {
                // LINEST/TREND/GROWTH/LOGEST are the *array* form of
                // the regression family and, unlike scalar FORECAST
                // (which drops a non-numeric pair and carries on),
                // real Excel rejects any non-numeric cell outright
                // with #VALUE! -- confirmed by probing all five
                // against the same text-containing range.
                let ys = match self.flatten_numbers_only_arg(evaluated_args.first()) {
                    Ok(v) => v,
                    Err(e) => return Ok(ResultData::Error(e)),
                };
                let xs = match evaluated_args.get(1) {
                    Some(arg) => match self.flatten_numbers_only(arg) {
                        Ok(v) => v,
                        Err(e) => return Ok(ResultData::Error(e)),
                    },
                    None => (1..=ys.len()).map(|i| i as f64).collect(),
                };
                let ln_ys: Vec<f64> = ys.iter().map(|y| y.ln()).collect();
                let m = match crate::core::stats::slope(&ln_ys, &xs) {
                    Ok(v) => v,
                    Err(e) => return Ok(ResultData::Error(e)),
                };
                let b = match crate::core::stats::intercept(&ln_ys, &xs) {
                    Ok(v) => v,
                    Err(e) => return Ok(ResultData::Error(e)),
                };
                if upper_name == "LOGEST" {
                    Ok(ResultData::List(vec![
                        ResultData::Float(m.exp()),
                        ResultData::Float(b.exp()),
                    ]))
                } else {
                    let new_x = evaluated_args
                        .get(2)
                        .and_then(|v| self.to_f64(v))
                        .unwrap_or_else(|| xs.first().copied().unwrap_or(1.0));
                    Ok(ResultData::Float((b + m * new_x).exp()))
                }
            }
            "HARMEAN" => {
                let nums: Vec<f64> =
                    match self.flatten_args_stat_numbers(evaluated_args, arg_is_direct) {
                        Ok(v) => v,
                        Err(e) => return Ok(ResultData::Error(e)),
                    };
                res_to_rd(crate::core::stats::harmean(&nums))
            }
            "HYPGEOM.DIST" | "HYPGEOMDIST" => {
                let sample_s = self.to_f64_arg(evaluated_args.first(), "HYPGEOM.DIST")?;
                let sample_size = self.to_f64_arg(evaluated_args.get(1), "HYPGEOM.DIST")?;
                let pop_s = self.to_f64_arg(evaluated_args.get(2), "HYPGEOM.DIST")?;
                let pop_size = self.to_f64_arg(evaluated_args.get(3), "HYPGEOM.DIST")?;
                // Legacy HYPGEOMDIST takes no cumulative flag at all --
                // it's always the point probability mass, never the
                // cumulative sum (unlike HYPGEOM.DIST, whose 5th
                // argument is required and selects between the two).
                let cumulative = if upper_name == "HYPGEOMDIST" {
                    false
                } else {
                    evaluated_args
                        .get(4)
                        .map(|v| self.to_bool(v))
                        .unwrap_or(true)
                };
                res_to_rd(crate::core::stats::hypgeom_dist(
                    sample_s,
                    sample_size,
                    pop_s,
                    pop_size,
                    cumulative,
                ))
            }
            "INTERCEPT" => {
                let (ys, xs) = match self.paired_args(evaluated_args.first(), evaluated_args.get(1))
                {
                    Ok(v) => v,
                    Err(e) => return Ok(ResultData::Error(e)),
                };
                res_to_rd(crate::core::stats::intercept(&ys, &xs))
            }
            "KURT" => {
                let nums: Vec<f64> =
                    match self.flatten_args_stat_numbers(evaluated_args, arg_is_direct) {
                        Ok(v) => v,
                        Err(e) => return Ok(ResultData::Error(e)),
                    };
                res_to_rd(crate::core::stats::kurt(&nums))
            }
            "LARGE" => {
                let nums: Vec<f64> = evaluated_args
                    .first()
                    .map(|arg| self.flatten_stat_numbers(arg, false))
                    .unwrap_or_default();
                let k = self.to_f64_arg(evaluated_args.get(1), "LARGE")?.round() as usize;
                res_to_rd(crate::core::stats::large(&nums, k))
            }
            "LINEST" | "TREND" => {
                // LINEST/TREND/GROWTH/LOGEST are the *array* form of
                // the regression family and, unlike scalar FORECAST
                // (which drops a non-numeric pair and carries on),
                // real Excel rejects any non-numeric cell outright
                // with #VALUE! -- confirmed by probing all five
                // against the same text-containing range.
                let ys = match self.flatten_numbers_only_arg(evaluated_args.first()) {
                    Ok(v) => v,
                    Err(e) => return Ok(ResultData::Error(e)),
                };
                let xs = match evaluated_args.get(1) {
                    Some(arg) => match self.flatten_numbers_only(arg) {
                        Ok(v) => v,
                        Err(e) => return Ok(ResultData::Error(e)),
                    },
                    None => (1..=ys.len()).map(|i| i as f64).collect(),
                };
                let m = match crate::core::stats::slope(&ys, &xs) {
                    Ok(v) => v,
                    Err(e) => return Ok(ResultData::Error(e)),
                };
                let b = match crate::core::stats::intercept(&ys, &xs) {
                    Ok(v) => v,
                    Err(e) => return Ok(ResultData::Error(e)),
                };
                if upper_name == "LINEST" {
                    Ok(ResultData::List(vec![
                        ResultData::Float(m),
                        ResultData::Float(b),
                    ]))
                } else {
                    let new_x = evaluated_args
                        .get(2)
                        .and_then(|v| self.to_f64(v))
                        .unwrap_or_else(|| xs.first().copied().unwrap_or(1.0));
                    Ok(ResultData::Float(m * new_x + b))
                }
            }
            "LOGNORM.DIST" | "LOGNORMDIST" => {
                let x = self.to_f64_arg(evaluated_args.first(), "LOGNORM.DIST")?;
                let mean = self.to_f64_arg(evaluated_args.get(1), "LOGNORM.DIST")?;
                let std_dev = self.to_f64_arg(evaluated_args.get(2), "LOGNORM.DIST")?;
                let cumulative = evaluated_args
                    .get(3)
                    .map(|v| self.to_bool(v))
                    .unwrap_or(true);
                res_to_rd(crate::core::stats::lognorm_dist(
                    x, mean, std_dev, cumulative,
                ))
            }
            "LOGNORM.INV" | "LOGINV" => {
                let p = self.to_f64_arg(evaluated_args.first(), "LOGNORM.INV")?;
                let mean = self.to_f64_arg(evaluated_args.get(1), "LOGNORM.INV")?;
                let std_dev = self.to_f64_arg(evaluated_args.get(2), "LOGNORM.INV")?;
                res_to_rd(crate::core::stats::lognorm_inv(p, mean, std_dev))
            }
            "MAXA" => {
                let nums: Vec<f64> =
                    match self.flatten_args_stat_numbers_a(evaluated_args, arg_is_direct) {
                        Ok(v) => v,
                        Err(e) => return Ok(ResultData::Error(e)),
                    };
                if nums.is_empty() {
                    Ok(ResultData::Float(0.0))
                } else {
                    Ok(ResultData::Float(
                        nums.iter().cloned().fold(f64::NEG_INFINITY, f64::max),
                    ))
                }
            }
            "MAXIFS" => {
                if evaluated_args.len() < 3 || (evaluated_args.len() - 1) % 2 != 0 {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                }
                let max_range = match &evaluated_args[0] {
                    ResultData::List(l) => l,
                    _ => return Ok(ResultData::Float(0.0)),
                };
                let mut criteria_pairs = Vec::new();
                let mut i = 1;
                while i < evaluated_args.len() {
                    let crit_range = match &evaluated_args[i] {
                        ResultData::List(l) => l,
                        _ => return Ok(ResultData::Float(0.0)),
                    };
                    let crit_val = &evaluated_args[i + 1];
                    criteria_pairs.push((crit_range, crit_val));
                    i += 2;
                }
                let mut max_val = f64::NEG_INFINITY;
                let mut found = false;
                for (idx, target_val) in max_range.iter().enumerate() {
                    let mut all_match = true;
                    for (crit_range, crit_val) in &criteria_pairs {
                        if idx >= crit_range.len()
                            || !self.match_criteria(&crit_range[idx], crit_val)
                        {
                            all_match = false;
                            break;
                        }
                    }
                    if all_match && let Some(f) = Self::aggregate_range_number(target_val) {
                        max_val = max_val.max(f);
                        found = true;
                    }
                }
                if !found {
                    Ok(ResultData::Float(0.0))
                } else {
                    Ok(ResultData::Float(max_val))
                }
            }
            "MEDIAN" => {
                let nums: Vec<f64> =
                    match self.flatten_args_stat_numbers(evaluated_args, arg_is_direct) {
                        Ok(v) => v,
                        Err(e) => return Ok(ResultData::Error(e)),
                    };
                res_to_rd(crate::core::stats::median(&nums))
            }
            "MINA" => {
                let nums: Vec<f64> =
                    match self.flatten_args_stat_numbers_a(evaluated_args, arg_is_direct) {
                        Ok(v) => v,
                        Err(e) => return Ok(ResultData::Error(e)),
                    };
                if nums.is_empty() {
                    Ok(ResultData::Float(0.0))
                } else {
                    Ok(ResultData::Float(
                        nums.iter().cloned().fold(f64::INFINITY, f64::min),
                    ))
                }
            }
            "MINIFS" => {
                if evaluated_args.len() < 3 || (evaluated_args.len() - 1) % 2 != 0 {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                }
                let min_range = match &evaluated_args[0] {
                    ResultData::List(l) => l,
                    _ => return Ok(ResultData::Float(0.0)),
                };
                let mut criteria_pairs = Vec::new();
                let mut i = 1;
                while i < evaluated_args.len() {
                    let crit_range = match &evaluated_args[i] {
                        ResultData::List(l) => l,
                        _ => return Ok(ResultData::Float(0.0)),
                    };
                    let crit_val = &evaluated_args[i + 1];
                    criteria_pairs.push((crit_range, crit_val));
                    i += 2;
                }
                let mut min_val = f64::INFINITY;
                let mut found = false;
                for (idx, target_val) in min_range.iter().enumerate() {
                    let mut all_match = true;
                    for (crit_range, crit_val) in &criteria_pairs {
                        if idx >= crit_range.len()
                            || !self.match_criteria(&crit_range[idx], crit_val)
                        {
                            all_match = false;
                            break;
                        }
                    }
                    if all_match && let Some(f) = Self::aggregate_range_number(target_val) {
                        min_val = min_val.min(f);
                        found = true;
                    }
                }
                if !found {
                    Ok(ResultData::Float(0.0))
                } else {
                    Ok(ResultData::Float(min_val))
                }
            }
            "MODE.MULT" => {
                // The MODE family rejects a lone blank operand where
                // its neighbours tolerate it -- MODE(x, <blank>) is
                // #VALUE! while MEDIAN(x, <blank>) is x. Applies to
                // all three spellings. See is_empty_scalar_operand.
                if evaluated_args.iter().any(Self::is_empty_scalar_operand) {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                }
                let nums: Vec<f64> =
                    match self.flatten_args_stat_numbers(evaluated_args, arg_is_direct) {
                        Ok(v) => v,
                        Err(e) => return Ok(ResultData::Error(e)),
                    };
                match crate::core::stats::mode_mult(&nums) {
                    Ok(modes) => Ok(ResultData::List(
                        modes.into_iter().map(ResultData::Float).collect(),
                    )),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "MODE.SNGL" | "MODE" => {
                // The MODE family rejects a lone blank operand where
                // its neighbours tolerate it -- MODE(x, <blank>) is
                // #VALUE! while MEDIAN(x, <blank>) is x. Applies to
                // all three spellings. See is_empty_scalar_operand.
                if evaluated_args.iter().any(Self::is_empty_scalar_operand) {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                }
                let nums: Vec<f64> =
                    match self.flatten_args_stat_numbers(evaluated_args, arg_is_direct) {
                        Ok(v) => v,
                        Err(e) => return Ok(ResultData::Error(e)),
                    };
                res_to_rd(crate::core::stats::mode_sngl(&nums))
            }
            "NEGBINOM.DIST" | "NEGBINOMDIST" => {
                let k = self.to_f64_arg(evaluated_args.first(), "NEGBINOM.DIST")?;
                let r = self.to_f64_arg(evaluated_args.get(1), "NEGBINOM.DIST")?;
                let p = self.to_f64_arg(evaluated_args.get(2), "NEGBINOM.DIST")?;
                let cumulative = evaluated_args
                    .get(3)
                    .map(|v| self.to_bool(v))
                    .unwrap_or(false);
                res_to_rd(crate::core::stats::negbinom_dist(k, r, p, cumulative))
            }
            "NORM.DIST" | "NORMDIST" => {
                let x = self.to_f64_arg(evaluated_args.first(), "NORM.DIST")?;
                let mean = self.to_f64_arg(evaluated_args.get(1), "NORM.DIST")?;
                let std_dev = self.to_f64_arg(evaluated_args.get(2), "NORM.DIST")?;
                let cumulative = evaluated_args
                    .get(3)
                    .map(|v| self.to_bool(v))
                    .unwrap_or(true);
                res_to_rd(crate::core::stats::norm_dist(x, mean, std_dev, cumulative))
            }
            "NORM.INV" | "NORMINV" => {
                let p = self.to_f64_arg(evaluated_args.first(), "NORM.INV")?;
                let mean = self.to_f64_arg(evaluated_args.get(1), "NORM.INV")?;
                let std_dev = self.to_f64_arg(evaluated_args.get(2), "NORM.INV")?;
                res_to_rd(crate::core::stats::norm_inv(p, mean, std_dev))
            }
            "NORM.S.DIST" | "NORMSDIST" => {
                let z = self.to_f64_arg(evaluated_args.first(), "NORM.S.DIST")?;
                let cumulative = evaluated_args
                    .get(1)
                    .map(|v| self.to_bool(v))
                    .unwrap_or(true);
                res_to_rd(crate::core::stats::norm_s_dist(z, cumulative))
            }
            "NORM.S.INV" | "NORMSINV" => {
                let p = self.to_f64_arg(evaluated_args.first(), "NORM.S.INV")?;
                res_to_rd(crate::core::stats::norm_s_inv(p))
            }
            "PERCENTILE.EXC" => {
                let nums: Vec<f64> = evaluated_args
                    .first()
                    .map(|arg| self.flatten_stat_numbers(arg, false))
                    .unwrap_or_default();
                let k = self.to_f64_arg(evaluated_args.get(1), "PERCENTILE.EXC")?;
                res_to_rd(crate::core::stats::percentile_exc(&nums, k))
            }
            "PERCENTILE.INC" | "PERCENTILE" => {
                let nums: Vec<f64> = evaluated_args
                    .first()
                    .map(|arg| self.flatten_stat_numbers(arg, false))
                    .unwrap_or_default();
                let k = self.to_f64_arg(evaluated_args.get(1), "PERCENTILE.INC")?;
                res_to_rd(crate::core::stats::percentile_inc(&nums, k))
            }
            "PERCENTRANK.EXC" => {
                let nums: Vec<f64> = evaluated_args
                    .first()
                    .map(|arg| self.flatten_stat_numbers(arg, false))
                    .unwrap_or_default();
                let x = self.to_f64_arg(evaluated_args.get(1), "PERCENTRANK.EXC")?;
                let sig = evaluated_args
                    .get(2)
                    .and_then(|v| self.to_f64(v))
                    .unwrap_or(3.0) as usize;
                res_to_rd(crate::core::stats::percentrank_exc(&nums, x, sig))
            }
            "PERCENTRANK.INC" | "PERCENTRANK" => {
                let nums: Vec<f64> = evaluated_args
                    .first()
                    .map(|arg| self.flatten_stat_numbers(arg, false))
                    .unwrap_or_default();
                let x = self.to_f64_arg(evaluated_args.get(1), "PERCENTRANK.INC")?;
                let sig = evaluated_args
                    .get(2)
                    .and_then(|v| self.to_f64(v))
                    .unwrap_or(3.0) as usize;
                res_to_rd(crate::core::stats::percentrank_inc(&nums, x, sig))
            }
            "PERMUT" => {
                let n = self.to_f64_arg(evaluated_args.first(), "PERMUT")?;
                let k = self.to_f64_arg(evaluated_args.get(1), "PERMUT")?;
                res_to_rd(crate::core::stats::permut(n, k))
            }
            "PERMUTATIONA" => {
                let n = self.to_f64_arg(evaluated_args.first(), "PERMUTATIONA")?;
                let k = self.to_f64_arg(evaluated_args.get(1), "PERMUTATIONA")?;
                res_to_rd(crate::core::stats::permutationa(n, k))
            }
            "PHI" => {
                let x = self.to_f64_arg(evaluated_args.first(), "PHI")?;
                res_to_rd(crate::core::stats::phi(x))
            }
            "POISSON.DIST" | "POISSON" => {
                let x = self.to_f64_arg(evaluated_args.first(), "POISSON.DIST")?;
                let mean = self.to_f64_arg(evaluated_args.get(1), "POISSON.DIST")?;
                let cumulative = evaluated_args
                    .get(2)
                    .map(|v| self.to_bool(v))
                    .unwrap_or(true);
                res_to_rd(crate::core::stats::poisson_dist(x, mean, cumulative))
            }
            "PROB" => {
                let (x_range, prob_range) =
                    match self.paired_args(evaluated_args.first(), evaluated_args.get(1)) {
                        Ok(v) => v,
                        Err(e) => return Ok(ResultData::Error(e)),
                    };
                let lower = self.to_f64_arg(evaluated_args.get(2), "PROB")?;
                let upper = evaluated_args.get(3).and_then(|v| self.to_f64(v));
                res_to_rd(crate::core::stats::prob(
                    &x_range,
                    &prob_range,
                    lower,
                    upper,
                ))
            }
            "QUARTILE.EXC" => {
                let nums: Vec<f64> = evaluated_args
                    .first()
                    .map(|arg| self.flatten_stat_numbers(arg, false))
                    .unwrap_or_default();
                let q = self
                    .to_f64_arg(evaluated_args.get(1), "QUARTILE.EXC")?
                    .round() as usize;
                res_to_rd(crate::core::stats::quartile_exc(&nums, q))
            }
            "QUARTILE.INC" | "QUARTILE" => {
                let nums: Vec<f64> = evaluated_args
                    .first()
                    .map(|arg| self.flatten_stat_numbers(arg, false))
                    .unwrap_or_default();
                let q = self
                    .to_f64_arg(evaluated_args.get(1), "QUARTILE.INC")?
                    .round() as usize;
                res_to_rd(crate::core::stats::quartile_inc(&nums, q))
            }
            "RANK.AVG" => {
                let number = self.to_f64_arg(evaluated_args.first(), "RANK.AVG")?;
                let ref_data: Vec<f64> = evaluated_args
                    .get(1)
                    .map(|arg| self.flatten_stat_numbers(arg, false))
                    .unwrap_or_default();
                let order = evaluated_args
                    .get(2)
                    .and_then(|v| self.to_f64(v))
                    .unwrap_or(0.0) as usize;
                res_to_rd(crate::core::stats::rank_avg(number, &ref_data, order))
            }
            "RANK.EQ" | "RANK" => {
                let number = self.to_f64_arg(evaluated_args.first(), "RANK.EQ")?;
                let ref_data: Vec<f64> = evaluated_args
                    .get(1)
                    .map(|arg| self.flatten_stat_numbers(arg, false))
                    .unwrap_or_default();
                let order = evaluated_args
                    .get(2)
                    .and_then(|v| self.to_f64(v))
                    .unwrap_or(0.0) as usize;
                res_to_rd(crate::core::stats::rank_eq(number, &ref_data, order))
            }
            "RSQ" => {
                let (ys, xs) = match self.paired_args(evaluated_args.first(), evaluated_args.get(1))
                {
                    Ok(v) => v,
                    Err(e) => return Ok(ResultData::Error(e)),
                };
                res_to_rd(crate::core::stats::rsq(&ys, &xs))
            }
            "SKEW" => {
                let nums: Vec<f64> =
                    match self.flatten_args_stat_numbers(evaluated_args, arg_is_direct) {
                        Ok(v) => v,
                        Err(e) => return Ok(ResultData::Error(e)),
                    };
                res_to_rd(crate::core::stats::skew(&nums))
            }
            "SKEW.P" => {
                let nums: Vec<f64> =
                    match self.flatten_args_stat_numbers(evaluated_args, arg_is_direct) {
                        Ok(v) => v,
                        Err(e) => return Ok(ResultData::Error(e)),
                    };
                res_to_rd(crate::core::stats::skew_p(&nums))
            }
            "SLOPE" => {
                let (ys, xs) = match self.paired_args(evaluated_args.first(), evaluated_args.get(1))
                {
                    Ok(v) => v,
                    Err(e) => return Ok(ResultData::Error(e)),
                };
                res_to_rd(crate::core::stats::slope(&ys, &xs))
            }
            "SMALL" => {
                let nums: Vec<f64> = evaluated_args
                    .first()
                    .map(|arg| self.flatten_stat_numbers(arg, false))
                    .unwrap_or_default();
                let k = self.to_f64_arg(evaluated_args.get(1), "SMALL")?.round() as usize;
                res_to_rd(crate::core::stats::small(&nums, k))
            }
            "STANDARDIZE" => {
                let x = self.to_f64_arg(evaluated_args.first(), "STANDARDIZE")?;
                let mean = self.to_f64_arg(evaluated_args.get(1), "STANDARDIZE")?;
                let std_dev = self.to_f64_arg(evaluated_args.get(2), "STANDARDIZE")?;
                res_to_rd(crate::core::stats::standardize(x, mean, std_dev))
            }
            "STDEV.P" | "STDEVP" => {
                let nums: Vec<f64> =
                    match self.flatten_args_stat_numbers(evaluated_args, arg_is_direct) {
                        Ok(v) => v,
                        Err(e) => return Ok(ResultData::Error(e)),
                    };
                res_to_rd(crate::core::stats::stdev_p(&nums))
            }
            "STDEV.S" | "STDEV" => {
                let nums: Vec<f64> =
                    match self.flatten_args_stat_numbers(evaluated_args, arg_is_direct) {
                        Ok(v) => v,
                        Err(e) => return Ok(ResultData::Error(e)),
                    };
                res_to_rd(crate::core::stats::stdev_s(&nums))
            }
            "STDEVA" => {
                let nums: Vec<f64> =
                    match self.flatten_args_stat_numbers_a(evaluated_args, arg_is_direct) {
                        Ok(v) => v,
                        Err(e) => return Ok(ResultData::Error(e)),
                    };
                res_to_rd(crate::core::stats::stdev_s(&nums))
            }
            "STDEVPA" => {
                let nums: Vec<f64> =
                    match self.flatten_args_stat_numbers_a(evaluated_args, arg_is_direct) {
                        Ok(v) => v,
                        Err(e) => return Ok(ResultData::Error(e)),
                    };
                res_to_rd(crate::core::stats::stdev_p(&nums))
            }
            "STEYX" => {
                let (ys, xs) = match self.paired_args(evaluated_args.first(), evaluated_args.get(1))
                {
                    Ok(v) => v,
                    Err(e) => return Ok(ResultData::Error(e)),
                };
                res_to_rd(crate::core::stats::steyx(&ys, &xs))
            }
            "T.DIST" => {
                let x = self.to_f64_arg(evaluated_args.first(), "T.DIST")?;
                let df = self.to_f64_arg(evaluated_args.get(1), "T.DIST")?;
                let cumulative = evaluated_args
                    .get(2)
                    .map(|v| self.to_bool(v))
                    .unwrap_or(true);
                res_to_rd(crate::core::stats::t_dist(x, df, cumulative))
            }
            "T.DIST.2T" => {
                let x = self.to_f64_arg(evaluated_args.first(), "T.DIST.2T")?;
                let df = self.to_f64_arg(evaluated_args.get(1), "T.DIST.2T")?;
                res_to_rd(crate::core::stats::t_dist_2t(x, df))
            }
            "TDIST" => {
                let x = self.to_f64_arg(evaluated_args.first(), "TDIST")?;
                let df = self.to_f64_arg(evaluated_args.get(1), "TDIST")?;
                let tails = self.to_f64_arg(evaluated_args.get(2), "TDIST")?;
                if tails == 1.0 {
                    res_to_rd(crate::core::stats::t_dist_rt(x, df))
                } else {
                    res_to_rd(crate::core::stats::t_dist_2t(x, df))
                }
            }
            "T.DIST.RT" => {
                let x = self.to_f64_arg(evaluated_args.first(), "T.DIST.RT")?;
                let df = self.to_f64_arg(evaluated_args.get(1), "T.DIST.RT")?;
                res_to_rd(crate::core::stats::t_dist_rt(x, df))
            }
            "T.INV" => {
                let p = self.to_f64_arg(evaluated_args.first(), "T.INV")?;
                let df = self.to_f64_arg(evaluated_args.get(1), "T.INV")?;
                res_to_rd(crate::core::stats::t_inv(p, df))
            }
            "T.INV.2T" | "TINV" => {
                let p = self.to_f64_arg(evaluated_args.first(), "T.INV.2T")?;
                let df = self.to_f64_arg(evaluated_args.get(1), "T.INV.2T")?;
                res_to_rd(crate::core::stats::t_inv_2t(p, df))
            }
            "T.TEST" | "TTEST" => {
                let tails = evaluated_args
                    .get(2)
                    .and_then(|v| self.to_f64(v))
                    .unwrap_or(2.0) as usize;
                let test_type = evaluated_args
                    .get(3)
                    .and_then(|v| self.to_f64(v))
                    .unwrap_or(1.0) as usize;
                // Only test_type 1 is the *paired* test, where the two
                // arrays must be the same size (#N/A otherwise) and a
                // non-numeric cell drops its whole (x, y) pair. Types 2
                // and 3 are two-*sample* tests that compare two
                // independent groups, so they accept different sizes
                // and each array drops its own non-numerics
                // independently. Both confirmed against real Excel:
                // `TTEST(4-cell-with-text, 4-cell, 1, 2)` equals
                // `TTEST(full-4-cell, 3-cell-survivor, 1, 2)`, while
                // the same call with type 1 instead equals the
                // 3-vs-3 pairwise-survivor form.
                let (array1, array2) = if test_type == 1 {
                    match self.paired_args(evaluated_args.first(), evaluated_args.get(1)) {
                        Ok(v) => v,
                        Err(e) => return Ok(ResultData::Error(e)),
                    }
                } else {
                    (
                        evaluated_args
                            .first()
                            .map(|arg| self.flatten_stat_numbers(arg, false))
                            .unwrap_or_default(),
                        evaluated_args
                            .get(1)
                            .map(|arg| self.flatten_stat_numbers(arg, false))
                            .unwrap_or_default(),
                    )
                };
                res_to_rd(crate::core::stats::t_test(
                    &array1, &array2, tails, test_type,
                ))
            }
            "TRIMMEAN" => {
                let nums: Vec<f64> = evaluated_args
                    .first()
                    .map(|arg| self.flatten_stat_numbers(arg, false))
                    .unwrap_or_default();
                let percent = self.to_f64_arg(evaluated_args.get(1), "TRIMMEAN")?;
                res_to_rd(crate::core::stats::trimmean(&nums, percent))
            }
            "VAR.P" | "VARP" => {
                let nums: Vec<f64> =
                    match self.flatten_args_stat_numbers(evaluated_args, arg_is_direct) {
                        Ok(v) => v,
                        Err(e) => return Ok(ResultData::Error(e)),
                    };
                res_to_rd(crate::core::stats::var_p(&nums))
            }
            "VAR.S" | "VAR" => {
                let nums: Vec<f64> =
                    match self.flatten_args_stat_numbers(evaluated_args, arg_is_direct) {
                        Ok(v) => v,
                        Err(e) => return Ok(ResultData::Error(e)),
                    };
                res_to_rd(crate::core::stats::var_s(&nums))
            }
            "VARA" => {
                let nums: Vec<f64> =
                    match self.flatten_args_stat_numbers_a(evaluated_args, arg_is_direct) {
                        Ok(v) => v,
                        Err(e) => return Ok(ResultData::Error(e)),
                    };
                res_to_rd(crate::core::stats::var_s(&nums))
            }
            "VARPA" => {
                let nums: Vec<f64> =
                    match self.flatten_args_stat_numbers_a(evaluated_args, arg_is_direct) {
                        Ok(v) => v,
                        Err(e) => return Ok(ResultData::Error(e)),
                    };
                res_to_rd(crate::core::stats::var_p(&nums))
            }
            "WEIBULL.DIST" | "WEIBULL" => {
                let x = self.to_f64_arg(evaluated_args.first(), "WEIBULL.DIST")?;
                let alpha = self.to_f64_arg(evaluated_args.get(1), "WEIBULL.DIST")?;
                let beta = self.to_f64_arg(evaluated_args.get(2), "WEIBULL.DIST")?;
                let cumulative = evaluated_args
                    .get(3)
                    .map(|v| self.to_bool(v))
                    .unwrap_or(true);
                res_to_rd(crate::core::stats::weibull_dist(x, alpha, beta, cumulative))
            }
            "Z.TEST" | "ZTEST" => {
                let array: Vec<f64> = evaluated_args
                    .first()
                    .map(|arg| self.flatten_stat_numbers(arg, false))
                    .unwrap_or_default();
                let x = self.to_f64_arg(evaluated_args.get(1), "Z.TEST")?;
                let sigma = evaluated_args.get(2).and_then(|v| self.to_f64(v));
                res_to_rd(crate::core::stats::z_test(&array, x, sigma))
            }
            _ => {
                *owned = false;
                Ok(ResultData::None)
            }
        }
    }
}
