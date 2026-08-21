//! Date and time function dispatch.
//!
//! Split out of the parent module's `evaluate_function`, which tries each
//! family in turn.

use super::{FnCall, res_to_rd};
use crate::core::engine::cell::{Dependency, EngineError};
use crate::core::engine::result_data::ResultData;
use crate::core::engine::sheet::Sheet;

impl Sheet {
    /// Evaluates `call` if this family owns its name, else `None`.
    pub(super) fn eval_date_time_fn(
        &self,
        call: FnCall<'_>,
        deps: &mut Vec<Dependency>,
    ) -> Option<Result<ResultData, EngineError>> {
        // The body returns `Result` so its arms can keep using `?`; whether
        // the name belongs to this family is signalled alongside.
        let mut owned = true;
        let r = self.eval_date_time_dispatch(call, deps, &mut owned);
        owned.then_some(r)
    }

    fn eval_date_time_dispatch(
        &self,
        call: FnCall<'_>,
        _deps: &mut Vec<Dependency>,
        owned: &mut bool,
    ) -> Result<ResultData, EngineError> {
        let FnCall { evaluated_args, .. } = call;
        match call.upper_name {
            // --- DATE AND TIME FUNCTIONS ---
            "DATE" => {
                let y = self.to_f64_arg(evaluated_args.first(), "DATE")?;
                let m = self.to_f64_arg(evaluated_args.get(1), "DATE")?;
                let d = self.to_f64_arg(evaluated_args.get(2), "DATE")?;
                res_to_rd(crate::core::date_fn::date_fn(y, m, d))
            }
            "DATEDIF" => {
                let start = self.to_f64_arg(evaluated_args.first(), "DATEDIF")?;
                let end = self.to_f64_arg(evaluated_args.get(1), "DATEDIF")?;
                let unit = evaluated_args
                    .get(2)
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                res_to_rd(crate::core::date_fn::datedif(start, end, &unit))
            }
            "DATEVALUE" => match evaluated_args.first() {
                Some(ResultData::String(text)) => res_to_rd(crate::core::date_fn::datevalue(text)),
                None => res_to_rd(crate::core::date_fn::datevalue("")),
                _ => Ok(ResultData::Error("#VALUE!".to_string())),
            },
            "DAY" => {
                let s = self.to_f64_arg(evaluated_args.first(), "DAY")?;
                res_to_rd(crate::core::date_fn::day_fn(s))
            }
            "DAYS" => {
                let e = self.to_f64_arg(evaluated_args.first(), "DAYS")?;
                let s = self.to_f64_arg(evaluated_args.get(1), "DAYS")?;
                res_to_rd(crate::core::date_fn::days(e, s))
            }
            "DAYS360" => {
                let s = self.to_f64_arg(evaluated_args.first(), "DAYS360")?;
                let e = self.to_f64_arg(evaluated_args.get(1), "DAYS360")?;
                let method = evaluated_args.get(2).map(|v| self.to_bool(v));
                res_to_rd(crate::core::date_fn::days360(s, e, method))
            }
            "EDATE" => {
                let s = self.to_f64_arg(evaluated_args.first(), "EDATE")?;
                let m = self.to_f64_arg(evaluated_args.get(1), "EDATE")?;
                res_to_rd(crate::core::date_fn::edate(s, m))
            }
            "EOMONTH" => {
                let s = self.to_f64_arg(evaluated_args.first(), "EOMONTH")?;
                let m = self.to_f64_arg(evaluated_args.get(1), "EOMONTH")?;
                res_to_rd(crate::core::date_fn::eomonth(s, m))
            }
            "HOUR" => {
                let s = self.to_f64_arg(evaluated_args.first(), "HOUR")?;
                res_to_rd(crate::core::date_fn::hour_fn(s))
            }
            "ISOWEEKNUM" => {
                let s = self.to_f64_arg(evaluated_args.first(), "ISOWEEKNUM")?;
                res_to_rd(crate::core::date_fn::isoweeknum(s))
            }
            "MINUTE" => {
                let s = self.to_f64_arg(evaluated_args.first(), "MINUTE")?;
                res_to_rd(crate::core::date_fn::minute_fn(s))
            }
            "MONTH" => {
                let s = self.to_f64_arg(evaluated_args.first(), "MONTH")?;
                res_to_rd(crate::core::date_fn::month_fn(s))
            }
            "NETWORKDAYS" | "NETWORKDAYS.INTL" => {
                let s = self.to_f64_arg(evaluated_args.first(), "NETWORKDAYS")?;
                let e = self.to_f64_arg(evaluated_args.get(1), "NETWORKDAYS")?;
                let holidays: Vec<f64> = evaluated_args
                    .get(2)
                    .map(|arg| self.flatten_stat_numbers(arg, false))
                    .unwrap_or_default();
                res_to_rd(crate::core::date_fn::networkdays(s, e, &holidays))
            }
            "SECOND" => {
                let s = self.to_f64_arg(evaluated_args.first(), "SECOND")?;
                res_to_rd(crate::core::date_fn::second_fn(s))
            }
            "TIME" => {
                let h = self.to_f64_arg(evaluated_args.first(), "TIME")?;
                let m = self.to_f64_arg(evaluated_args.get(1), "TIME")?;
                let s = self.to_f64_arg(evaluated_args.get(2), "TIME")?;
                res_to_rd(crate::core::date_fn::time_fn(h, m, s))
            }
            "TIMEVALUE" => match evaluated_args.first() {
                Some(ResultData::String(text)) => res_to_rd(crate::core::date_fn::timevalue(text)),
                None => res_to_rd(crate::core::date_fn::timevalue("")),
                _ => Ok(ResultData::Error("#VALUE!".to_string())),
            },
            "WEEKDAY" => {
                let s = self.to_f64_arg(evaluated_args.first(), "WEEKDAY")?;
                let r_type = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                res_to_rd(crate::core::date_fn::weekday(s, r_type))
            }
            "WEEKNUM" => {
                let s = self.to_f64_arg(evaluated_args.first(), "WEEKNUM")?;
                let r_type = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                res_to_rd(crate::core::date_fn::weeknum(s, r_type))
            }
            "WORKDAY" | "WORKDAY.INTL" => {
                let s = self.to_f64_arg(evaluated_args.first(), "WORKDAY")?;
                let days = self.to_f64_arg(evaluated_args.get(1), "WORKDAY")?;
                let holidays: Vec<f64> = evaluated_args
                    .get(2)
                    .map(|arg| self.flatten_stat_numbers(arg, false))
                    .unwrap_or_default();
                res_to_rd(crate::core::date_fn::workday(s, days, &holidays))
            }
            "YEAR" => {
                let s = self.to_f64_arg(evaluated_args.first(), "YEAR")?;
                res_to_rd(crate::core::date_fn::year_fn(s))
            }
            "YEARFRAC" => {
                let s = self.to_f64_arg(evaluated_args.first(), "YEARFRAC")?;
                let e = self.to_f64_arg(evaluated_args.get(1), "YEARFRAC")?;
                let basis = evaluated_args.get(2).and_then(|v| self.to_f64(v));
                res_to_rd(crate::core::date_fn::yearfrac(s, e, basis))
            }
            _ => {
                *owned = false;
                Ok(ResultData::None)
            }
        }
    }
}
