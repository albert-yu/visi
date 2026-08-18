//! Math and trigonometry function dispatch.
//!
//! Split out of the parent module's `evaluate_function`, which tries each
//! family in turn.

use super::{FnCall, res_to_rd};
use crate::core::engine::cell::{Dependency, EngineError};
use crate::core::engine::result_data::ResultData;
use crate::core::engine::sheet::Sheet;

impl Sheet {
    /// Evaluates `call` if this family owns its name, else `None`.
    pub(super) fn eval_math_trig_fn(
        &self,
        call: FnCall<'_>,
        deps: &mut Vec<Dependency>,
    ) -> Option<Result<ResultData, EngineError>> {
        // The body returns `Result` so its arms can keep using `?`; whether
        // the name belongs to this family is signalled alongside.
        let mut owned = true;
        let r = self.eval_math_trig_dispatch(call, deps, &mut owned);
        owned.then_some(r)
    }

    fn eval_math_trig_dispatch(
        &self,
        call: FnCall<'_>,
        _deps: &mut Vec<Dependency>,
        owned: &mut bool,
    ) -> Result<ResultData, EngineError> {
        let FnCall {
            upper_name,
            args,
            evaluated_args,
            arg_is_direct,
            ..
        } = call;
        match call.upper_name {
            // --- MATH AND TRIGONOMETRY FUNCTIONS ---
            "ACOSH" => {
                let x = self.to_f64_arg(evaluated_args.first(), "ACOSH")?;
                res_to_rd(crate::core::math_trig::acosh(x))
            }
            "ACOT" => {
                let x = self.to_f64_arg(evaluated_args.first(), "ACOT")?;
                res_to_rd(crate::core::math_trig::acot(x))
            }
            "ACOTH" => {
                let x = self.to_f64_arg(evaluated_args.first(), "ACOTH")?;
                res_to_rd(crate::core::math_trig::acoth(x))
            }
            "AGGREGATE" => {
                // AGGREGATE(function_num, options, ref1, ...) -- unlike
                // SUBTOTAL(function_num, ref1, ...), its *second*
                // argument is the options flag, not data. Sharing
                // SUBTOTAL's handler (which skips only the first
                // argument) folded that options value straight into
                // the aggregated numbers, so e.g. AGGREGATE(4, 6, ...)
                // computed MAX over the data *plus a literal 6*.
                let fn_num = self
                    .to_f64_arg(evaluated_args.first(), "AGGREGATE")?
                    .round() as usize;
                let options = evaluated_args
                    .get(1)
                    .and_then(|v| self.to_f64(v))
                    .unwrap_or(0.0)
                    .round() as usize;
                // Function numbers 14-19 (LARGE/SMALL/PERCENTILE.INC/
                // QUARTILE.INC/PERCENTILE.EXC/QUARTILE.EXC) take a
                // trailing k argument after the array.
                let takes_k = (14..=19).contains(&fn_num);
                let data_end = if takes_k {
                    evaluated_args.len().saturating_sub(1)
                } else {
                    evaluated_args.len()
                };
                let k = if takes_k {
                    evaluated_args
                        .last()
                        .and_then(|v| self.to_f64(v))
                        .unwrap_or(1.0)
                } else {
                    1.0
                };
                // Options 2/3/6/7 mean "ignore error values"; every
                // option this engine can express other than that still
                // propagates an error in the data, matching Excel.
                let ignores_errors = matches!(options, 2 | 3 | 6 | 7);
                let data_args = &evaluated_args[2.min(evaluated_args.len())..data_end];
                if !ignores_errors && let Some(err) = Self::find_error_in_args(data_args) {
                    return Ok(err);
                }
                let nums: Vec<f64> = data_args
                    .iter()
                    .flat_map(|arg| self.flatten_stat_numbers(arg, false))
                    .collect();
                match fn_num {
                    1 => res_to_rd(if nums.is_empty() {
                        Err("#DIV/0!".to_string())
                    } else {
                        Ok(nums.iter().sum::<f64>() / nums.len() as f64)
                    }),
                    2 | 3 => Ok(ResultData::Float(nums.len() as f64)),
                    // MAX/MIN over nothing is 0, not an infinity --
                    // which the dispatch-level NaN/infinity guard would
                    // otherwise turn into #NUM!.
                    4 => Ok(ResultData::Float(if nums.is_empty() {
                        0.0
                    } else {
                        nums.iter().cloned().fold(f64::NEG_INFINITY, f64::max)
                    })),
                    5 => Ok(ResultData::Float(if nums.is_empty() {
                        0.0
                    } else {
                        nums.iter().cloned().fold(f64::INFINITY, f64::min)
                    })),
                    6 => Ok(ResultData::Float(nums.iter().product())),
                    7 => res_to_rd(crate::core::stats::stdev_s(&nums)),
                    8 => res_to_rd(crate::core::stats::stdev_p(&nums)),
                    9 => Ok(ResultData::Float(nums.iter().sum())),
                    10 => res_to_rd(crate::core::stats::var_s(&nums)),
                    11 => res_to_rd(crate::core::stats::var_p(&nums)),
                    12 => res_to_rd(crate::core::stats::median(&nums)),
                    13 => res_to_rd(crate::core::stats::mode_sngl(&nums)),
                    14 => res_to_rd(crate::core::stats::large(&nums, k.round() as usize)),
                    15 => res_to_rd(crate::core::stats::small(&nums, k.round() as usize)),
                    16 => res_to_rd(crate::core::stats::percentile_inc(&nums, k)),
                    17 => res_to_rd(crate::core::stats::quartile_inc(&nums, k.round() as usize)),
                    18 => res_to_rd(crate::core::stats::percentile_exc(&nums, k)),
                    19 => res_to_rd(crate::core::stats::quartile_exc(&nums, k.round() as usize)),
                    _ => Ok(ResultData::Error("#VALUE!".to_string())),
                }
            }
            "SUBTOTAL" => {
                let fn_num = self.to_f64_arg(evaluated_args.first(), "SUBTOTAL")?.round() as usize;
                let nums: Vec<f64> = evaluated_args
                    .iter()
                    .skip(1)
                    .flat_map(|arg| self.flatten_stat_numbers(arg, false))
                    .collect();
                match fn_num % 100 {
                    1 => res_to_rd(if nums.is_empty() {
                        Err("#DIV/0!".to_string())
                    } else {
                        Ok(nums.iter().sum::<f64>() / nums.len() as f64)
                    }),
                    2 | 3 => Ok(ResultData::Float(nums.len() as f64)),
                    // MAX/MIN over nothing is 0, matching plain
                    // MAX/MIN (and not an infinity, which the
                    // dispatch-level guard would turn into #NUM!).
                    4 => Ok(ResultData::Float(if nums.is_empty() {
                        0.0
                    } else {
                        nums.iter().cloned().fold(f64::NEG_INFINITY, f64::max)
                    })),
                    5 => Ok(ResultData::Float(if nums.is_empty() {
                        0.0
                    } else {
                        nums.iter().cloned().fold(f64::INFINITY, f64::min)
                    })),
                    6 => Ok(ResultData::Float(nums.iter().product())),
                    7 => res_to_rd(crate::core::stats::stdev_s(&nums)),
                    8 => res_to_rd(crate::core::stats::stdev_p(&nums)),
                    9 => Ok(ResultData::Float(nums.iter().sum())),
                    10 => res_to_rd(crate::core::stats::var_s(&nums)),
                    11 => res_to_rd(crate::core::stats::var_p(&nums)),
                    12 => res_to_rd(crate::core::stats::median(&nums)),
                    _ => Ok(ResultData::Float(nums.iter().sum())),
                }
            }
            "ARABIC" => {
                let text = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                res_to_rd(crate::core::math_trig::arabic(&text))
            }
            "ASINH" => {
                let x = self.to_f64_arg(evaluated_args.first(), "ASINH")?;
                res_to_rd(crate::core::math_trig::asinh(x))
            }
            "ATAN2" => {
                let x = self.to_f64_arg(evaluated_args.first(), "ATAN2")?;
                let y = self.to_f64_arg(evaluated_args.get(1), "ATAN2")?;
                res_to_rd(crate::core::math_trig::atan2(x, y))
            }
            "ATANH" => {
                let x = self.to_f64_arg(evaluated_args.first(), "ATANH")?;
                res_to_rd(crate::core::math_trig::atanh(x))
            }
            "BASE" => {
                let num = self.to_f64_arg(evaluated_args.first(), "BASE")?;
                let radix = self.to_f64_arg(evaluated_args.get(1), "BASE")?;
                let min_len = evaluated_args.get(2).and_then(|v| self.to_f64(v));
                match crate::core::math_trig::base(num, radix, min_len) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "CEILING.MATH" | "CEILING.PRECISE" | "ISO.CEILING" => {
                let x = self.to_f64_arg(evaluated_args.first(), "CEILING.MATH")?;
                let sig = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                let mode = evaluated_args.get(2).and_then(|v| self.to_f64(v));
                res_to_rd(crate::core::math_trig::ceiling_math(x, sig, mode))
            }
            "COMBIN" => {
                let n = self.to_f64_arg(evaluated_args.first(), "COMBIN")?;
                let k = self.to_f64_arg(evaluated_args.get(1), "COMBIN")?;
                res_to_rd(crate::core::math_trig::combin(n, k))
            }
            "COMBINA" => {
                let n = self.to_f64_arg(evaluated_args.first(), "COMBINA")?;
                let k = self.to_f64_arg(evaluated_args.get(1), "COMBINA")?;
                res_to_rd(crate::core::math_trig::combina(n, k))
            }
            "COSH" => {
                let x = self.to_f64_arg(evaluated_args.first(), "COSH")?;
                res_to_rd(crate::core::math_trig::cosh(x))
            }
            "COT" => {
                let x = self.to_f64_arg(evaluated_args.first(), "COT")?;
                res_to_rd(crate::core::math_trig::cot(x))
            }
            "COTH" => {
                let x = self.to_f64_arg(evaluated_args.first(), "COTH")?;
                res_to_rd(crate::core::math_trig::coth(x))
            }
            "CSC" => {
                let x = self.to_f64_arg(evaluated_args.first(), "CSC")?;
                res_to_rd(crate::core::math_trig::csc(x))
            }
            "CSCH" => {
                let x = self.to_f64_arg(evaluated_args.first(), "CSCH")?;
                res_to_rd(crate::core::math_trig::csch(x))
            }
            "DECIMAL" => {
                let text = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let radix = self.to_f64_arg(evaluated_args.get(1), "DECIMAL")?;
                res_to_rd(crate::core::math_trig::decimal(&text, radix))
            }
            "DEGREES" => {
                let rad = self.to_f64_arg(evaluated_args.first(), "DEGREES")?;
                res_to_rd(crate::core::math_trig::degrees(rad))
            }
            "EVEN" => {
                let x = self.to_f64_arg(evaluated_args.first(), "EVEN")?;
                res_to_rd(crate::core::math_trig::even(x))
            }
            "FACT" => {
                let n = self.to_f64_arg(evaluated_args.first(), "FACT")?;
                res_to_rd(crate::core::math_trig::fact(n))
            }
            "FACTDOUBLE" => {
                // FACTDOUBLE(TRUE) is #VALUE! even though
                // FACTDOUBLE(1) is 1. See first_arg_is_boolean.
                if Self::first_arg_is_boolean(evaluated_args) {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                }
                let n = self.to_f64_arg(evaluated_args.first(), "FACTDOUBLE")?;
                res_to_rd(crate::core::math_trig::factdouble(n))
            }
            "FLOOR.MATH" | "FLOOR.PRECISE" => {
                let x = self.to_f64_arg(evaluated_args.first(), "FLOOR.MATH")?;
                let sig = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                let mode = evaluated_args.get(2).and_then(|v| self.to_f64(v));
                res_to_rd(crate::core::math_trig::floor_math(x, sig, mode))
            }
            "GCD" | "LCM" => {
                // Unlike MULTINOMIAL, a blank operand here is *omitted*, not
                // zero: real Excel gives LCM(1, <blank>) = 1 (as if LCM(1)),
                // not LCM(1, 0) = 0. Measured with fuzz/fuzz_excel.py seed
                // 308076.
                let mut nums = Vec::new();
                for arg in evaluated_args {
                    match self.flatten_skipping_blanks(Some(arg)) {
                        Ok(v) => nums.extend(v),
                        Err(e) => return Ok(ResultData::Error(e)),
                    }
                }
                if upper_name == "GCD" {
                    res_to_rd(crate::core::math_trig::gcd(&nums))
                } else {
                    res_to_rd(crate::core::math_trig::lcm(&nums))
                }
            }
            "LOG" => {
                let num = self.to_f64_arg(evaluated_args.first(), "LOG")?;
                let base = self.opt_f64_arg(evaluated_args, 1, 10.0)?;
                // Base 1 is #DIV/0!, not #NUM!: log(n)/log(1) divides
                // by zero. Everything else out of domain stays #NUM!
                // (both confirmed against real Excel).
                if base == 1.0 {
                    Ok(ResultData::Error("#DIV/0!".to_string()))
                } else if num <= 0.0 || base <= 0.0 {
                    Ok(ResultData::Error("#NUM!".to_string()))
                } else {
                    Ok(ResultData::Float(num.log(base)))
                }
            }
            "MDETERM" => {
                let matrix = match (args.first(), evaluated_args.first()) {
                    (Some(e), Some(v)) => self.matrix_from_arg(e, v),
                    _ => Vec::new(),
                };
                res_to_rd(crate::core::math_trig::mdeterm(&matrix))
            }
            "MINVERSE" => {
                let matrix = match (args.first(), evaluated_args.first()) {
                    (Some(e), Some(v)) => self.matrix_from_arg(e, v),
                    _ => Vec::new(),
                };
                match crate::core::math_trig::minverse(&matrix) {
                    Ok(inv) => Ok(ResultData::List(
                        inv.into_iter()
                            .map(|row| {
                                ResultData::List(row.into_iter().map(ResultData::Float).collect())
                            })
                            .collect(),
                    )),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "MROUND" => {
                let x = self.to_f64_arg(evaluated_args.first(), "MROUND")?;
                let mult = self.to_f64_arg(evaluated_args.get(1), "MROUND")?;
                res_to_rd(crate::core::math_trig::mround(x, mult))
            }
            "MULTINOMIAL" => {
                // Like GCD/LCM, MULTINOMIAL rejects a non-numeric cell
                // outright (#VALUE!) instead of skipping it the way
                // SUM does -- a blank inside a range still counts as 0.
                // ... and a blank operand is only a *missing* operand
                // when there is nothing else: MULTINOMIAL(<blank>) and
                // MULTINOMIAL(<blank>, <blank>) are #VALUE! while
                // MULTINOMIAL(3, <blank>) is 1, the blank counting as
                // 0. That is narrower than SUMPRODUCT, where any lone
                // blank operand is #VALUE! even beside a number.
                if !evaluated_args.is_empty()
                    && evaluated_args.iter().all(Self::is_empty_scalar_operand)
                {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                }
                let mut nums = Vec::new();
                for arg in evaluated_args {
                    match self.flatten_strict_numbers(arg) {
                        Ok(v) => nums.extend(v),
                        Err(e) => return Ok(ResultData::Error(e)),
                    }
                }
                res_to_rd(crate::core::math_trig::multinomial(&nums))
            }
            "MUNIT" => {
                let dim = self.to_f64_arg(evaluated_args.first(), "MUNIT")?;
                match crate::core::math_trig::munit(dim) {
                    Ok(mat) => Ok(ResultData::List(
                        mat.into_iter()
                            .map(|row| {
                                ResultData::List(row.into_iter().map(ResultData::Float).collect())
                            })
                            .collect(),
                    )),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "ODD" => {
                let x = self.to_f64_arg(evaluated_args.first(), "ODD")?;
                res_to_rd(crate::core::math_trig::odd(x))
            }
            "PERCENTOF" => {
                // PERCENTOF(subset, all) is SUM(subset)/SUM(all), and
                // it inherits SUM's leniency rather than erroring on a
                // non-numeric argument: real Excel gives 0 for
                // PERCENTOF(<text>, 10) (the numerator sums to 0) and
                // #DIV/0! for PERCENTOF(10, <text>) or PERCENTOF(10, 0)
                // (the denominator does). Routing both arguments
                // through to_f64_arg instead made any text #VALUE!.
                // Text *inside a referenced range* sums as 0 (so
                // PERCENTOF(<text cell>, 10) is 0 and
                // PERCENTOF(10, <text cell>) is #DIV/0!), but a
                // directly-supplied non-numeric value -- a literal, or
                // the result of a nested call like LOWER(...) -- is
                // #VALUE!. That's the same direct-vs-reference split
                // the SUM/AVERAGE helpers already make.
                let mut sums = [0.0f64; 2];
                for (i, slot) in sums.iter_mut().enumerate() {
                    let Some(v) = evaluated_args.get(i) else {
                        continue;
                    };
                    if arg_is_direct.get(i).copied().unwrap_or(false) {
                        match v {
                            ResultData::None => {}
                            other => match self.to_f64(other) {
                                Some(f) => *slot = f,
                                None => {
                                    return Ok(ResultData::Error("#VALUE!".to_string()));
                                }
                            },
                        }
                    } else {
                        *slot = self.flatten_stat_numbers(v, false).iter().sum();
                    }
                }
                let [data_val, target_val] = sums;
                if target_val == 0.0 {
                    Ok(ResultData::Error("#DIV/0!".to_string()))
                } else {
                    res_to_rd(crate::core::math_trig::percentof(data_val, target_val))
                }
            }
            "PI" => Ok(ResultData::Float(std::f64::consts::PI)),
            "POWER" => {
                let num = self.to_f64_arg(evaluated_args.first(), "POWER")?;
                let p = self.to_f64_arg(evaluated_args.get(1), "POWER")?;
                res_to_rd(crate::core::math_trig::power(num, p))
            }
            "QUOTIENT" => {
                // QUOTIENT rejects *booleans* but still coerces
                // numeric text: QUOTIENT(12, TRUE) is #VALUE! while
                // QUOTIENT("12", 5) is 2 and QUOTIENT(12, "ab") is
                // #VALUE!. (MOD differs again -- MOD(TRUE, 2) is 1.)
                // So this is to_f64's coercion with booleans excluded,
                // not a numbers-only rule -- rejecting numeric strings
                // too made QUOTIENT over a CONCATENATE/RIGHT result
                // #VALUE! where Excel computes.
                let coerce = |v: Option<&ResultData>| -> Option<f64> {
                    match v {
                        Some(ResultData::Boolean(_)) => None,
                        Some(other) => self.to_f64(other),
                        None => None,
                    }
                };
                let (Some(num), Some(den)) = (
                    coerce(evaluated_args.first()),
                    coerce(evaluated_args.get(1)),
                ) else {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                };
                res_to_rd(crate::core::math_trig::quotient(num, den))
            }
            "RADIANS" => {
                let deg = self.to_f64_arg(evaluated_args.first(), "RADIANS")?;
                res_to_rd(crate::core::math_trig::radians(deg))
            }
            "RANDARRAY" => {
                let rows = evaluated_args.first().and_then(|v| self.to_f64(v));
                let cols = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                let min = evaluated_args.get(2).and_then(|v| self.to_f64(v));
                let max = evaluated_args.get(3).and_then(|v| self.to_f64(v));
                let whole = evaluated_args.get(4).map(|v| self.to_bool(v));
                match crate::core::math_trig::randarray(rows, cols, min, max, whole) {
                    Ok(grid) => Ok(ResultData::List(
                        grid.into_iter()
                            .map(|row| {
                                ResultData::List(row.into_iter().map(ResultData::Float).collect())
                            })
                            .collect(),
                    )),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "ROMAN" => {
                let num = self.to_f64_arg(evaluated_args.first(), "ROMAN")?;
                let form = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                match crate::core::math_trig::roman(num, form) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "SEC" => {
                let x = self.to_f64_arg(evaluated_args.first(), "SEC")?;
                res_to_rd(crate::core::math_trig::sec(x))
            }
            "SECH" => {
                let x = self.to_f64_arg(evaluated_args.first(), "SECH")?;
                res_to_rd(crate::core::math_trig::sech(x))
            }
            "SEQUENCE" => {
                let rows = self.to_f64_arg(evaluated_args.first(), "SEQUENCE")?;
                let cols = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                let start = evaluated_args.get(2).and_then(|v| self.to_f64(v));
                let step = evaluated_args.get(3).and_then(|v| self.to_f64(v));
                match crate::core::math_trig::sequence(rows, cols, start, step) {
                    Ok(grid) => Ok(ResultData::List(
                        grid.into_iter()
                            .map(|row| {
                                ResultData::List(row.into_iter().map(ResultData::Float).collect())
                            })
                            .collect(),
                    )),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "SERIESSUM" => {
                let x = self.to_f64_arg(evaluated_args.first(), "SERIESSUM")?;
                let n = self.to_f64_arg(evaluated_args.get(1), "SERIESSUM")?;
                let m = self.to_f64_arg(evaluated_args.get(2), "SERIESSUM")?;
                let coeffs = match self.flatten_skipping_blanks(evaluated_args.get(3)) {
                    Ok(v) => v,
                    Err(e) => return Ok(ResultData::Error(e)),
                };
                res_to_rd(crate::core::math_trig::seriessum(x, n, m, &coeffs))
            }
            "SIGN" => {
                let x = self.to_f64_arg(evaluated_args.first(), "SIGN")?;
                res_to_rd(crate::core::math_trig::sign(x))
            }
            "SINH" => {
                let x = self.to_f64_arg(evaluated_args.first(), "SINH")?;
                res_to_rd(crate::core::math_trig::sinh(x))
            }
            "SQRTPI" => {
                if Self::first_arg_is_boolean(evaluated_args) {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                }
                let x = self.to_f64_arg(evaluated_args.first(), "SQRTPI")?;
                res_to_rd(crate::core::math_trig::sqrtpi(x))
            }
            "SUMPRODUCT" => {
                // SUMPRODUCT treats non-numeric entries as zeros rather
                // than skipping or rejecting them, which matters twice
                // over: the term contributes 0, and -- because the
                // entry still occupies its slot -- the arrays stay the
                // same length so the remaining terms keep lining up.
                // Dropping them instead made SUMPRODUCT(2, "abc")
                // #VALUE! (length 1 against length 0) where real Excel
                // answers 0, and SUMPRODUCT({1,2}, {3,"x"}) is 3.
                let mut arrays: Vec<Vec<f64>> = Vec::new();
                let mut first_err = None;
                for arg in evaluated_args {
                    // A single blank cell is a missing operand, not an
                    // empty array: SUMPRODUCT over one blank cell is
                    // #VALUE! where over two it is 0.
                    if Self::is_empty_scalar_operand(arg) {
                        return Ok(ResultData::Error("#VALUE!".to_string()));
                    }
                    let mut slots = Vec::new();
                    self.flatten_positional(arg, &mut slots, &mut first_err);
                    arrays.push(slots.into_iter().map(|v| v.unwrap_or(0.0)).collect());
                }
                if let Some(e) = first_err {
                    return Ok(ResultData::Error(e));
                }
                res_to_rd(crate::core::math_trig::sumproduct(&arrays))
            }
            "SUMSQ" => {
                let nums: Vec<f64> =
                    match self.flatten_args_stat_numbers(evaluated_args, arg_is_direct) {
                        Ok(v) => v,
                        Err(e) => return Ok(ResultData::Error(e)),
                    };
                res_to_rd(crate::core::math_trig::sumsq(&nums))
            }
            "SUMX2MY2" => {
                // paired_args first: a shape mismatch is #N/A and takes
                // precedence over everything below it, even when a
                // range also holds no numbers at all.
                let (xs, ys) = match self.paired_args(evaluated_args.first(), evaluated_args.get(1))
                {
                    Ok(v) => v,
                    Err(e) => return Ok(ResultData::Error(e)),
                };
                if self.paired_sum_has_no_numbers(evaluated_args.first())
                    || self.paired_sum_has_no_numbers(evaluated_args.get(1))
                {
                    return Ok(ResultData::Error("#DIV/0!".to_string()));
                }
                res_to_rd(crate::core::math_trig::sumx2my2(&xs, &ys))
            }
            "SUMX2PY2" => {
                // paired_args first: a shape mismatch is #N/A and takes
                // precedence over everything below it, even when a
                // range also holds no numbers at all.
                let (xs, ys) = match self.paired_args(evaluated_args.first(), evaluated_args.get(1))
                {
                    Ok(v) => v,
                    Err(e) => return Ok(ResultData::Error(e)),
                };
                if self.paired_sum_has_no_numbers(evaluated_args.first())
                    || self.paired_sum_has_no_numbers(evaluated_args.get(1))
                {
                    return Ok(ResultData::Error("#DIV/0!".to_string()));
                }
                res_to_rd(crate::core::math_trig::sumx2py2(&xs, &ys))
            }
            "SUMXMY2" => {
                // paired_args first: a shape mismatch is #N/A and takes
                // precedence over everything below it, even when a
                // range also holds no numbers at all.
                let (xs, ys) = match self.paired_args(evaluated_args.first(), evaluated_args.get(1))
                {
                    Ok(v) => v,
                    Err(e) => return Ok(ResultData::Error(e)),
                };
                if self.paired_sum_has_no_numbers(evaluated_args.first())
                    || self.paired_sum_has_no_numbers(evaluated_args.get(1))
                {
                    return Ok(ResultData::Error("#DIV/0!".to_string()));
                }
                res_to_rd(crate::core::math_trig::sumxmy2(&xs, &ys))
            }
            "TANH" => {
                let x = self.to_f64_arg(evaluated_args.first(), "TANH")?;
                res_to_rd(crate::core::math_trig::tanh(x))
            }
            "TRUNC" => {
                let x = self.to_f64_arg(evaluated_args.first(), "TRUNC")?;
                let digits = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                res_to_rd(crate::core::math_trig::trunc(x, digits))
            }
            _ => {
                *owned = false;
                Ok(ResultData::None)
            }
        }
    }
}
