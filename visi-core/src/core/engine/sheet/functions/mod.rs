//! Excel function dispatch.
//!
//! `evaluate_function` does what has to happen before a function's arguments
//! can be evaluated -- prefix stripping, the lazy/short-circuit functions, and
//! the per-family error-propagation policy -- then offers the evaluated call
//! to each family module in turn. A family returns `None` for a name it does
//! not own, and the chain falls through to "Unknown function".
//!
//! Splitting on family keeps any one file reviewable. The cost is that a call
//! may be tested against several families' matches rather than one; each is a
//! compiler-optimized string match, so this is a small constant, but it is on
//! the hot path and worth a benchmark before adding more families.

mod date_time;
mod engineering;
mod info_lookup;
mod math_trig;
mod stats;
mod text;

use super::{Context, LetScope, Sheet};
use crate::core::engine::cell::{Dependency, EngineError, EvalError};
use crate::core::engine::result_data::ResultData;
use crate::core::parser::Expr;

/// One evaluated function call, as handed to a family module.
///
/// `Copy`, so a family can destructure it and still read `call.upper_name`.
#[derive(Clone, Copy)]
pub(super) struct FnCall<'a> {
    /// Uppercased and stripped of `_xlfn.`/`_xlws.`; what families match on.
    pub upper_name: &'a str,
    /// The unevaluated argument expressions, for the functions that need the
    /// AST rather than the value.
    pub args: &'a [Expr],
    /// The evaluated arguments.
    pub evaluated_args: &'a [ResultData],
    /// Per argument: whether it came from a direct cell reference rather than
    /// a computed expression.
    pub arg_is_direct: &'a [bool],
    /// The other sheets, for cross-sheet references.
    pub context: Option<&'a Context<'a>>,
    /// The row the call is being evaluated for, if any.
    pub row: Option<usize>,
    /// The column the call is being evaluated for, if any.
    pub col: Option<usize>,
    /// Enclosing `LET` bindings.
    pub scope: &'a LetScope<'a>,
}

/// Adapts a numeric function's `Result<f64, String>` to a cell value, where
/// the error string is an Excel error code rather than a Rust failure.
pub(super) fn res_to_rd(res: Result<f64, String>) -> Result<ResultData, EngineError> {
    match res {
        Ok(v) => Ok(ResultData::Float(v)),
        Err(e) => Ok(ResultData::Error(e)),
    }
}

/// A NaN can only come from a math function evaluated outside its domain
/// (ASIN/ACOS of |x|>1, SQRT/LN/LOG10 of a negative, ...), and an infinity
/// only from one that overflowed (POWER(42, 600), EXP(1000)). Excel has
/// neither -- it reports #NUM! for both -- so rather than bolting a
/// domain/overflow guard onto each of those call sites, normalize here at the
/// single point every function result flows through.
fn post_process(r: Result<ResultData, EngineError>) -> Result<ResultData, EngineError> {
    match r {
        Ok(ResultData::Float(f)) if !f.is_finite() => Ok(ResultData::Error("#NUM!".to_string())),
        other => other,
    }
}

impl Sheet {
    /// Evaluates a function call by name.
    ///
    /// Handles the lazy/short-circuit functions itself, since their arguments
    /// must not be evaluated up front, then evaluates the remaining arguments
    /// and offers the call to each family module in turn.
    #[allow(clippy::too_many_arguments)]
    pub(super) fn evaluate_function(
        &self,
        name: &str,
        args: &[crate::core::parser::Expr],
        context: Option<&Context>,
        row: Option<usize>,
        col: Option<usize>,
        deps: &mut Vec<Dependency>,
        scope: &LetScope<'_>,
    ) -> Result<ResultData, EngineError> {
        use crate::core::parser::Expr;
        let mut upper_name = name.to_uppercase();
        if upper_name.starts_with("_XLFN.") {
            upper_name = upper_name["_XLFN.".len()..].to_string();
        }
        // Real Excel's OOXML writer additionally nests some dynamic-array
        // worksheet functions (UNIQUE, SORT, FILTER, ...) under a second
        // `_xlws.` prefix inside `_xlfn.` -- e.g. `_xlfn._xlws.SORT`, not
        // just `_xlfn.SORT`. Without stripping it too, the un-stripped
        // name never matches any dispatch arm here -- confirmed as a real
        // mismatch by checking real Excel's own OOXML export for these
        // functions directly.
        if upper_name.starts_with("_XLWS.") {
            upper_name = upper_name["_XLWS.".len()..].to_string();
        }

        if upper_name == "LET" {
            return self.evaluate_let(args, context, row, col, deps, scope);
        }

        if upper_name == "IF" {
            if args.len() < 3 {
                return Err(EngineError::EvalError(EvalError::UnknownFunction(
                    "IF requires 3 arguments".to_string(),
                )));
            }
            let cond_val = self.evaluate_ast(&args[0], context, row, col, deps, scope)?;
            if let ResultData::Error(_) = cond_val {
                return Ok(cond_val);
            }
            let condition = match self.to_bool_opt(&cond_val) {
                Some(b) => b,
                None => return Ok(ResultData::Error("#VALUE!".to_string())),
            };
            if condition {
                return self.evaluate_ast(&args[1], context, row, col, deps, scope);
            } else {
                return self.evaluate_ast(&args[2], context, row, col, deps, scope);
            }
        }

        if upper_name == "IFERROR" {
            if args.len() < 2 {
                return Err(EngineError::EvalError(EvalError::UnknownFunction(
                    "IFERROR requires 2 arguments".to_string(),
                )));
            }
            let first_res = self.evaluate_ast(&args[0], context, row, col, deps, scope);
            match first_res {
                Ok(ResultData::Error(_)) | Err(_) => {
                    return self.evaluate_ast(&args[1], context, row, col, deps, scope);
                }
                Ok(val) => return Ok(val),
            }
        }

        if upper_name == "IFNA" {
            if args.len() < 2 {
                return Err(EngineError::EvalError(EvalError::UnknownFunction(
                    "IFNA requires 2 arguments".to_string(),
                )));
            }
            let first_val = self.evaluate_ast(&args[0], context, row, col, deps, scope)?;
            if let ResultData::Error(ref e) = first_val
                && e == "#N/A"
            {
                return self.evaluate_ast(&args[1], context, row, col, deps, scope);
            }
            return Ok(first_val);
        }

        if upper_name == "IFS" {
            // Lazily evaluated: only the arms up to and including the
            // first TRUE condition are ever computed, so an error
            // sitting in a later (unselected) value never propagates.
            // Confirmed against real Excel: `IFS(TRUE, 42, TRUE, 1/0)`
            // is 42, while `IFS(FALSE, 42, TRUE, 1/0)` is #DIV/0!.
            let mut i = 0;
            while i + 1 < args.len() {
                let cond = self.evaluate_ast(&args[i], context, row, col, deps, scope)?;
                if let ResultData::Error(_) = cond {
                    return Ok(cond);
                }
                if self.to_bool(&cond) {
                    return self.evaluate_ast(&args[i + 1], context, row, col, deps, scope);
                }
                i += 2;
            }
            return Ok(ResultData::Error("#N/A".to_string()));
        }

        if upper_name == "SWITCH" {
            // Lazily evaluated for the same reason as IFS: an error in
            // a value arm that isn't selected must not propagate
            // (`SWITCH(2, 1, 1/0, 2, 99, -1)` is 99 in real Excel).
            if args.len() < 3 {
                return Ok(ResultData::Error("#VALUE!".to_string()));
            }
            let target = self.evaluate_ast(&args[0], context, row, col, deps, scope)?;
            if let ResultData::Error(_) = target {
                return Ok(target);
            }
            let mut i = 1;
            while i + 1 < args.len() {
                let case = self.evaluate_ast(&args[i], context, row, col, deps, scope)?;
                if let ResultData::Error(_) = case {
                    return Ok(case);
                }
                if target.to_string() == case.to_string() {
                    return self.evaluate_ast(&args[i + 1], context, row, col, deps, scope);
                }
                i += 2;
            }
            // A trailing odd argument is the default.
            if i < args.len() {
                return self.evaluate_ast(&args[i], context, row, col, deps, scope);
            }
            return Ok(ResultData::Error("#N/A".to_string()));
        }

        if upper_name == "CHOOSE" {
            if args.len() < 2 {
                return Err(EngineError::EvalError(EvalError::UnknownFunction(
                    "CHOOSE requires at least 2 arguments".to_string(),
                )));
            }
            let idx_val = self.evaluate_ast(&args[0], context, row, col, deps, scope)?;
            if let ResultData::Error(_) = idx_val {
                return Ok(idx_val);
            }
            let idx = match self.to_f64(&idx_val) {
                Some(f) => f.round() as isize,
                None => return Ok(ResultData::Error("#VALUE!".to_string())),
            };
            let choices = &args[1..];
            if idx >= 1 && (idx as usize) <= choices.len() {
                return self.evaluate_ast(
                    &choices[(idx - 1) as usize],
                    context,
                    row,
                    col,
                    deps,
                    scope,
                );
            } else {
                return Ok(ResultData::Error("#VALUE!".to_string()));
            }
        }

        if upper_name == "LAMBDA" {
            // A bare, uninvoked LAMBDA (not nested as another
            // function's argument, e.g. `=LAMBDA(x, x*2)` alone in a
            // cell) has nothing to apply it to -- the parser doesn't
            // support the `LAMBDA(...)(args)` immediate-invocation
            // syntax (that would need the grammar to allow calling an
            // arbitrary sub-expression, not just a bare identifier),
            // so this mirrors Excel's #CALC! for an unusable lambda.
            return Ok(ResultData::Error("#CALC!".to_string()));
        }

        if matches!(
            upper_name.as_str(),
            "MAP" | "BYROW" | "BYCOL" | "REDUCE" | "SCAN" | "MAKEARRAY"
        ) {
            return self.evaluate_lambda_function(
                upper_name.as_str(),
                args,
                context,
                row,
                col,
                deps,
                scope,
            );
        }

        if upper_name == "ISOMITTED" {
            // Best-effort: every lambda invocation path here
            // (MAP/BYROW/BYCOL/REDUCE/SCAN/MAKEARRAY) always supplies
            // exactly as many argument values as the lambda declares
            // parameters, so a declared parameter is never actually
            // left unbound -- this can only ever observe "not found
            // in scope at all", which is the honest limitation to
            // report rather than silently guessing.
            let is_omitted = match args.first() {
                Some(Expr::Identifier(name)) => scope.get(name).is_none(),
                _ => false,
            };
            return Ok(ResultData::Boolean(is_omitted));
        }

        if matches!(
            upper_name.as_str(),
            "ROW"
                | "ROWS"
                | "COLUMN"
                | "COLUMNS"
                | "AREAS"
                | "ISREF"
                | "FORMULATEXT"
                | "ISFORMULA"
                | "INDIRECT"
                | "OFFSET"
                | "SHEET"
                | "SHEETS"
                | "CELL"
                | "INFO"
        ) {
            return self.evaluate_range_info_function(
                upper_name.as_str(),
                args,
                context,
                row,
                col,
                deps,
                scope,
            );
        }

        if matches!(
            upper_name.as_str(),
            "TRANSPOSE"
                | "HSTACK"
                | "VSTACK"
                | "CHOOSEROWS"
                | "CHOOSECOLS"
                | "DROP"
                | "EXPAND"
                | "TAKE"
                | "TOCOL"
                | "TOROW"
                | "WRAPROWS"
                | "WRAPCOLS"
                | "UNIQUE"
                | "SORT"
                | "SORTBY"
                | "FILTER"
                | "TRIMRANGE"
        ) {
            return self.evaluate_array_reshape_function(
                upper_name.as_str(),
                args,
                context,
                row,
                col,
                deps,
                scope,
            );
        }

        if upper_name == "GETPIVOTDATA" {
            return self.evaluate_getpivotdata(args, context, row, col, deps, scope);
        }

        if upper_name == "ISERROR" {
            if args.is_empty() {
                return Ok(ResultData::Boolean(false));
            }
            let res = self.evaluate_ast(&args[0], context, row, col, deps, scope);
            return match res {
                Ok(ResultData::Error(_)) | Err(_) => Ok(ResultData::Boolean(true)),
                _ => Ok(ResultData::Boolean(false)),
            };
        }

        if upper_name == "ISNA" {
            if args.is_empty() {
                return Ok(ResultData::Boolean(false));
            }
            let res = self.evaluate_ast(&args[0], context, row, col, deps, scope);
            return match res {
                Ok(ResultData::Error(e)) => Ok(ResultData::Boolean(e.contains("#N/A"))),
                _ => Ok(ResultData::Boolean(false)),
            };
        }

        let mut evaluated_args = Vec::new();
        let mut arg_is_direct = Vec::new();
        for arg in args {
            let is_direct_arg = match arg {
                Expr::CellRef { .. } | Expr::RangeRef { .. } | Expr::StructuredRef { .. } => false,
                Expr::FunctionCall { name, .. } => {
                    let n = name.to_uppercase();
                    n != "IF" && n != "IFERROR" && n != "CHOOSE"
                }
                _ => true,
            };
            arg_is_direct.push(is_direct_arg);
            let eval_res = match self.evaluate_ast(arg, context, row, col, deps, scope) {
                Ok(r) => r,
                Err(EngineError::EvalError(EvalError::UnknownFunction(err_str)))
                    if err_str.starts_with('#') =>
                {
                    ResultData::Error(err_str)
                }
                Err(e) => return Err(e),
            };
            evaluated_args.push(eval_res);
        }

        let uses_ordered_arg_error_check = matches!(
            upper_name.as_str(),
            "SUM" | "AVERAGE" | "MIN" | "MAX" | "PRODUCT"
        );
        // The type-introspection functions must see an error value
        // rather than have it propagate past them: real Excel answers
        // TYPE(1/0) = 16, ISNONTEXT(1/0) = TRUE, and
        // ISTEXT/ISNUMBER/ISLOGICAL/ISBLANK(1/0) = FALSE. (Math
        // functions like ISODD do still propagate -- ISODD(1/0) is
        // #DIV/0! -- so they stay out of this list.)
        let inspects_errors = matches!(
            upper_name.as_str(),
            "IFERROR"
                | "ISERROR"
                | "ISNA"
                | "ISERR"
                | "ERROR.TYPE"
                | "TYPE"
                | "ISTEXT"
                | "ISNONTEXT"
                | "ISNUMBER"
                | "ISLOGICAL"
                | "ISBLANK"
        );
        if !inspects_errors
                // COUNTA counts an error argument as one more non-blank
                // value, and COUNT skips it, rather than either
                // propagating it (both match real Excel).
                && upper_name != "COUNTA"
                && upper_name != "COUNT"
                // COUNTBLANK just asks which cells are empty; an error in
                // the range is a non-blank cell, not a reason to fail.
                && upper_name != "COUNTBLANK"
                // AGGREGATE decides for itself whether to propagate or
                // ignore an error in its data, based on its `options`
                // argument, so it must see the raw arguments.
                && upper_name != "AGGREGATE"
                // The paired statistical functions check their two ranges'
                // shapes before anything else -- a size mismatch is #N/A
                // even when a range also holds an error value -- so they
                // re-raise errors themselves (see paired_args).
                && !matches!(
                    upper_name.as_str(),
                    "CORREL"
                        | "PEARSON"
                        | "COVAR"
                        | "COVARIANCE.P"
                        | "COVARIANCE.S"
                        | "SLOPE"
                        | "INTERCEPT"
                        | "RSQ"
                        | "STEYX"
                        | "FORECAST"
                        | "FORECAST.LINEAR"
                        | "SUMX2MY2"
                        | "SUMX2PY2"
                        | "SUMXMY2"
                        | "CHISQ.TEST"
                        | "CHITEST"
                )
                && !uses_ordered_arg_error_check
                && let Some(err) = Self::find_error_in_args(&evaluated_args)
        {
            return Ok(err);
        }

        // A NaN can only come from a math function evaluated outside
        // its domain (ASIN/ACOS of |x|>1, SQRT/LN/LOG10 of a negative,
        // ...), and an infinity only from one that overflowed
        // (POWER(42, 600), EXP(1000)). Excel has neither -- it reports
        // #NUM! for both -- so rather than bolting a domain/overflow
        // guard onto each of those call sites individually, normalize
        // here at the single point every function result flows
        // through.
        let call = FnCall {
            upper_name: upper_name.as_str(),
            args,
            evaluated_args: &evaluated_args,
            arg_is_direct: &arg_is_direct,
            context,
            row,
            col,
            scope,
        };
        if let Some(r) = self.eval_stats_fn(call, deps) {
            return post_process(r);
        }
        if let Some(r) = self.eval_math_trig_fn(call, deps) {
            return post_process(r);
        }
        if let Some(r) = self.eval_text_fn(call, deps) {
            return post_process(r);
        }
        if let Some(r) = self.eval_date_time_fn(call, deps) {
            return post_process(r);
        }
        if let Some(r) = self.eval_engineering_fn(call, deps) {
            return post_process(r);
        }
        if let Some(r) = self.eval_info_lookup_fn(call, deps) {
            return post_process(r);
        }
        post_process(Err(EngineError::EvalError(EvalError::UnknownFunction(
            format!("Unknown function: {}", name),
        ))))
    }
}
