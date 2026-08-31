//! Information, logical, database, lookup, web and cube function dispatch.
//!
//! Split out of the parent module's `evaluate_function`, which tries each
//! family in turn.

use super::{FnCall, res_to_rd};
use crate::core::engine::cell::{CellRef, Dependency, EngineError, EvalError};
use crate::core::engine::result_data::ResultData;
use crate::core::engine::sheet::Sheet;
use crate::core::finance;
use crate::core::parser::Expr;

impl Sheet {
    /// Evaluates `call` if this family owns its name, else `None`.
    pub(super) fn eval_info_lookup_fn(
        &self,
        call: FnCall<'_>,
        deps: &mut Vec<Dependency>,
    ) -> Option<Result<ResultData, EngineError>> {
        // The body returns `Result` so its arms can keep using `?`; whether
        // the name belongs to this family is signalled alongside.
        let mut owned = true;
        let r = self.eval_info_lookup_dispatch(call, deps, &mut owned);
        owned.then_some(r)
    }

    fn eval_info_lookup_dispatch(
        &self,
        call: FnCall<'_>,
        deps: &mut Vec<Dependency>,
        owned: &mut bool,
    ) -> Result<ResultData, EngineError> {
        let FnCall {
            upper_name,
            args,
            evaluated_args,
            arg_is_direct,
            context,
            row,
            col,
            scope,
            ..
        } = call;
        match call.upper_name {
            // --- INFORMATION & LOGICAL & DATABASE & LOOKUP & WEB & CUBE FUNCTIONS ---
            "ERROR.TYPE" => {
                let t = match evaluated_args.first() {
                    Some(ResultData::Error(e)) => e.clone(),
                    _ => String::new(),
                };
                res_to_rd(crate::core::extended_fn::error_type(&t))
            }
            "ISERR" => {
                let val = evaluated_args.first().cloned().unwrap_or(ResultData::None);
                Ok(ResultData::Boolean(crate::core::extended_fn::iserr(&val)))
            }
            "ISEVEN" => {
                let n = self.to_f64_arg(evaluated_args.first(), "ISEVEN")?;
                Ok(ResultData::Boolean(crate::core::extended_fn::iseven(n)))
            }
            "ISLOGICAL" => {
                let val = evaluated_args.first().cloned().unwrap_or(ResultData::None);
                Ok(ResultData::Boolean(crate::core::extended_fn::islogical(
                    &val,
                )))
            }
            "ISNONTEXT" => {
                let val = evaluated_args.first().cloned().unwrap_or(ResultData::None);
                Ok(ResultData::Boolean(crate::core::extended_fn::isnontext(
                    &val,
                )))
            }
            "ISODD" => {
                let n = self.to_f64_arg(evaluated_args.first(), "ISODD")?;
                Ok(ResultData::Boolean(crate::core::extended_fn::isodd(n)))
            }
            "N" => {
                let val = evaluated_args.first().cloned().unwrap_or(ResultData::None);
                Ok(ResultData::Float(crate::core::extended_fn::n_fn(&val)))
            }
            "NA" => Ok(crate::core::extended_fn::na_fn()),
            "TYPE" => {
                let val = evaluated_args.first().cloned().unwrap_or(ResultData::None);
                Ok(ResultData::Float(crate::core::extended_fn::type_fn(&val)))
            }
            "XOR" => {
                let bools: Vec<bool> = evaluated_args.iter().map(|v| self.to_bool(v)).collect();
                Ok(ResultData::Boolean(crate::core::extended_fn::xor_fn(
                    &bools,
                )))
            }
            "ADDRESS" => {
                let r = self.to_f64_arg(evaluated_args.first(), "ADDRESS")?;
                let c = self.to_f64_arg(evaluated_args.get(1), "ADDRESS")?;
                let abs_n = evaluated_args.get(2).and_then(|v| self.to_f64(v));
                let a1 = evaluated_args.get(3).map(|v| self.to_bool(v));
                let s_name = evaluated_args.get(4).map(|v| v.to_string());
                match crate::core::extended_fn::address_fn(r, c, abs_n, a1, s_name.as_deref()) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "HLOOKUP" => {
                if evaluated_args.len() < 3 {
                    return Err(EngineError::EvalError(EvalError::UnknownFunction(
                        "HLOOKUP requires at least 3 arguments".to_string(),
                    )));
                }
                let lookup_val = &evaluated_args[0];
                let row_idx = self.to_f64(&evaluated_args[2]).unwrap_or(1.0) as usize;
                let range_lookup = if evaluated_args.len() >= 4 {
                    self.to_bool(&evaluated_args[3])
                } else {
                    true
                };

                // The mirror image of VLOOKUP just below: the range's
                // flat, row-major `List` is reshaped using the
                // *unevaluated* range's column span, the first *row*
                // (not column) is searched, and the match is read back
                // out of the target row.
                if let ResultData::List(list) = &evaluated_args[1] {
                    let num_cols = match &args[1] {
                        Expr::RangeRef {
                            start_col, end_col, ..
                        } => end_col - start_col + 1,
                        _ => list.len(),
                    };

                    let num_rows = list.len().checked_div(num_cols).unwrap_or(0);
                    if num_rows == 0 || row_idx == 0 || row_idx > num_rows {
                        return Ok(ResultData::Error("#N/A".to_string()));
                    }

                    let first_row = &list[..num_cols];
                    let mut found_col_idx: Option<usize> = None;
                    if !range_lookup {
                        for (c, item) in first_row.iter().enumerate() {
                            if Self::exact_lookup_matches(lookup_val, item) {
                                found_col_idx = Some(c);
                                break;
                            }
                        }
                    } else {
                        let lookup_f = self.to_f64(lookup_val).unwrap_or(0.0);
                        for (c, item) in first_row.iter().enumerate() {
                            let val_f = self.to_f64(item).unwrap_or(0.0);
                            if val_f <= lookup_f {
                                found_col_idx = Some(c);
                            } else {
                                break;
                            }
                        }
                    }

                    match found_col_idx {
                        Some(c) => Ok(list[(row_idx - 1) * num_cols + c].clone()),
                        None => Ok(ResultData::Error("#N/A".to_string())),
                    }
                } else {
                    Ok(ResultData::Error("#N/A".to_string()))
                }
            }
            "ENCODEURL" => {
                let text = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                match crate::core::extended_fn::encodeurl(&text) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            // Both need a live external data source this engine has no
            // access to (Microsoft's undocumented stock-data cloud
            // service for STOCKHISTORY; a registered Windows COM
            // IRtdServer for RTD) -- #N/A matches what real Excel shows
            // once that connection is unavailable, rather than the
            // misleading echo-the-last-argument placeholder these used
            // to fall through to.
            "STOCKHISTORY" | "RTD" => Ok(ResultData::Error("#N/A".to_string())),
            "DAVERAGE" | "DCOUNT" | "DCOUNTA" | "DGET" | "DMAX" | "DMIN" | "DPRODUCT"
            | "DSTDEV" | "DSTDEVP" | "DSUM" | "DVAR" | "DVARP" => {
                self.evaluate_database_function(upper_name, args, evaluated_args, context)
            }
            "HYPERLINK" => {
                // No clickable-hyperlink concept in this engine --
                // returns the display value a formula-based consumer
                // would see: friendly_name if given, else the raw
                // link_location text.
                match evaluated_args.get(1) {
                    Some(friendly) => Ok(friendly.clone()),
                    None => Ok(evaluated_args
                        .first()
                        .cloned()
                        .unwrap_or(ResultData::Error("#VALUE!".to_string()))),
                }
            }
            // No OLAP cube connection concept exists in this engine --
            // #N/A matches what real Excel shows once a cube function's
            // underlying connection is unavailable, the same reasoning
            // already applied to RTD/STOCKHISTORY above -- a plausible-looking
            // wrong value is worse than a visible error, since it can silently
            // corrupt a downstream calculation with no signal anything is wrong.
            "CUBEKPIMEMBER" | "CUBEMEMBER" | "CUBEMEMBERPROPERTY" | "CUBERANKEDMEMBER"
            | "CUBESET" | "CUBESETCOUNT" | "CUBEVALUE" => Ok(ResultData::Error("#N/A".to_string())),
            // WEBSERVICE needs actual network access to an arbitrary
            // URL; #VALUE! matches Microsoft's own documented error
            // for a request that can't be completed.
            "WEBSERVICE" => Ok(ResultData::Error("#VALUE!".to_string())),
            // IMAGE needs to fetch/decode real image data, which this
            // engine has no concept of; #VALUE! matches real Excel's
            // error for a source it can't resolve to a usable image.
            "IMAGE" => Ok(ResultData::Error("#VALUE!".to_string())),
            // GROUPBY/PIVOTBY are genuine, deterministic array
            // functions (not connection-dependent like the above).
            // Properly implementing Excel's full row/column-field grouping and
            // dynamic-array spill semantics is real, separately-scoped work
            // (this engine already has the pivot-table grouping machinery in
            // pivot.rs that a real implementation would build on).
            "GROUPBY" | "PIVOTBY" => Ok(evaluated_args.last().cloned().unwrap_or(ResultData::None)),
            "FILTERXML" => {
                let xml = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let xpath = evaluated_args
                    .get(1)
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                match crate::core::xml::filterxml(&xml, &xpath) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }

            "SUM" => {
                if let Some(err) = self.check_arg_errors(evaluated_args, arg_is_direct) {
                    return Ok(err);
                }
                let mut sum = 0.0;
                for (i, arg) in evaluated_args.iter().enumerate() {
                    sum += self.sum_helper(arg, arg_is_direct[i]);
                }
                Ok(ResultData::Float(sum))
            }
            "AVERAGE" => {
                if let Some(err) = self.check_arg_errors(evaluated_args, arg_is_direct) {
                    return Ok(err);
                }
                let mut sum = 0.0;
                let mut count = 0;
                for (i, arg) in evaluated_args.iter().enumerate() {
                    let (s, c) = self.average_helper(arg, arg_is_direct[i]);
                    sum += s;
                    count += c;
                }
                if count == 0 {
                    Ok(ResultData::Error("#DIV/0!".to_string()))
                } else {
                    Ok(ResultData::Float(sum / count as f64))
                }
            }
            "COUNT" => {
                // A boolean counts when it is typed directly as an
                // argument, but not when it merely sits inside a
                // referenced range -- Excel's documented split, and the
                // same is_direct distinction the SUM/AVERAGE helpers
                // already make.
                let mut count = 0;
                for (i, arg) in evaluated_args.iter().enumerate() {
                    let direct = arg_is_direct.get(i).copied().unwrap_or(false);
                    if direct && matches!(arg, ResultData::Boolean(_)) {
                        count += 1;
                    } else if direct
                        && matches!(arg, ResultData::String(_))
                        && self.to_f64(arg).is_some()
                    {
                        // Numeric text typed directly counts too --
                        // COUNT("12", 3, 4, 5) is 4. Text that will not
                        // coerce is simply not counted; unlike the rest
                        // of the family COUNT never reports #VALUE!.
                        count += 1;
                    } else {
                        count += self.count_helper(arg);
                    }
                }
                Ok(ResultData::Float(count as f64))
            }
            "MIN" => {
                if let Some(err) = self.check_arg_errors(evaluated_args, arg_is_direct) {
                    return Ok(err);
                }
                let mut min_val = f64::INFINITY;
                for (i, arg) in evaluated_args.iter().enumerate() {
                    min_val = min_val.min(self.min_helper(arg, arg_is_direct[i]));
                }
                if min_val.is_infinite() {
                    Ok(ResultData::Float(0.0))
                } else {
                    Ok(ResultData::Float(min_val))
                }
            }
            "MAX" => {
                if let Some(err) = self.check_arg_errors(evaluated_args, arg_is_direct) {
                    return Ok(err);
                }
                let mut max_val = f64::NEG_INFINITY;
                for (i, arg) in evaluated_args.iter().enumerate() {
                    max_val = max_val.max(self.max_helper(arg, arg_is_direct[i]));
                }
                if max_val.is_infinite() {
                    Ok(ResultData::Float(0.0))
                } else {
                    Ok(ResultData::Float(max_val))
                }
            }

            "STR" => {
                if evaluated_args.is_empty() {
                    Ok(ResultData::String(String::new()))
                } else {
                    Ok(ResultData::String(evaluated_args[0].to_string()))
                }
            }
            "SQRT" => {
                let val = self.to_f64_arg(evaluated_args.first(), "SQRT")?;
                if val < 0.0 {
                    Ok(ResultData::Error("#NUM!".to_string()))
                } else {
                    Ok(ResultData::Float(val.sqrt()))
                }
            }
            "RAND" => {
                use rand::Rng;
                let mut rng = rand::thread_rng();
                Ok(ResultData::Float(rng.r#gen::<f64>()))
            }
            "RANDBETWEEN" => {
                if evaluated_args.len() < 2 {
                    return Err(EngineError::EvalError(EvalError::UnknownFunction(
                        "RANDBETWEEN requires 2 arguments".to_string(),
                    )));
                }
                let bottom = self
                    .to_f64_arg(evaluated_args.first(), "RANDBETWEEN")?
                    .round() as i64;
                let top = self
                    .to_f64_arg(evaluated_args.get(1), "RANDBETWEEN")?
                    .round() as i64;
                use rand::Rng;
                let mut rng = rand::thread_rng();
                let val = if bottom <= top {
                    rng.gen_range(bottom..=top)
                } else {
                    rng.gen_range(top..=bottom)
                };
                Ok(ResultData::Integer(val))
            }
            "SIN" => {
                let val = self.to_f64_arg(evaluated_args.first(), "SIN")?;
                res_to_rd(crate::core::math_trig::check_trig_domain(val).map(|()| val.sin()))
            }
            "COS" => {
                let val = self.to_f64_arg(evaluated_args.first(), "COS")?;
                res_to_rd(crate::core::math_trig::check_trig_domain(val).map(|()| val.cos()))
            }
            "TAN" => {
                let val = self.to_f64_arg(evaluated_args.first(), "TAN")?;
                res_to_rd(crate::core::math_trig::check_trig_domain(val).map(|()| val.tan()))
            }
            "ACOS" => {
                let val = self.to_f64_arg(evaluated_args.first(), "ACOS")?;
                Ok(ResultData::Float(val.acos()))
            }
            "ASIN" => {
                let val = self.to_f64_arg(evaluated_args.first(), "ASIN")?;
                Ok(ResultData::Float(val.asin()))
            }
            "ATAN" => {
                let val = self.to_f64_arg(evaluated_args.first(), "ATAN")?;
                Ok(ResultData::Float(val.atan()))
            }
            "FLOOR" => {
                let val = self.to_f64_arg(evaluated_args.first(), "FLOOR")?;
                let sig = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                res_to_rd(crate::core::math_trig::floor_math(val, sig, None))
            }
            "CEILING" => {
                let val = self.to_f64_arg(evaluated_args.first(), "CEILING")?;
                let sig = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                res_to_rd(crate::core::math_trig::ceiling_math(val, sig, None))
            }
            "LOG10" => {
                let val = self.to_f64_arg(evaluated_args.first(), "LOG10")?;
                Ok(ResultData::Float(val.log10()))
            }
            "LN" => {
                let val = self.to_f64_arg(evaluated_args.first(), "LN")?;
                Ok(ResultData::Float(val.ln()))
            }
            "EXP" => {
                let val = self.to_f64_arg(evaluated_args.first(), "EXP")?;
                Ok(ResultData::Float(val.exp()))
            }
            "GET" => {
                if evaluated_args.len() == 2 {
                    let row = self.to_f64(&evaluated_args[0]).unwrap_or(0.0) as usize;
                    let col = self.to_f64(&evaluated_args[1]).unwrap_or(0.0) as usize;
                    let cell_ref = CellRef::new(row, col);
                    deps.push(Dependency::Local(cell_ref));
                    Ok(self.get_result_data(&cell_ref))
                } else if evaluated_args.len() == 3 {
                    let sheet_name = evaluated_args[0].to_string();
                    let row = self.to_f64(&evaluated_args[1]).unwrap_or(0.0) as usize;
                    let col = self.to_f64(&evaluated_args[2]).unwrap_or(0.0) as usize;
                    let cell_ref = CellRef::new(row, col);

                    if sheet_name == self.name {
                        deps.push(Dependency::Local(cell_ref));
                        Ok(self.get_result_data(&cell_ref))
                    } else {
                        deps.push(Dependency::Remote {
                            sheet: sheet_name.clone(),
                            cell: cell_ref,
                        });
                        if let Some(ctx) = context {
                            if let Some(t) = ctx.sheets.get(&sheet_name) {
                                Ok(t.get_result_data(&cell_ref))
                            } else {
                                Err(EngineError::EvalError(EvalError::UnknownFunction(format!(
                                    "Sheet not found: {}",
                                    sheet_name
                                ))))
                            }
                        } else {
                            Err(EngineError::EvalError(EvalError::UnknownFunction(
                                "No context".to_string(),
                            )))
                        }
                    }
                } else {
                    Err(EngineError::EvalError(EvalError::UnknownFunction(
                        "get() takes 2 or 3 arguments".to_string(),
                    )))
                }
            }
            "GET_COL" => {
                if evaluated_args.len() == 2 {
                    let sheet_name = evaluated_args[0].to_string();
                    let col_name = evaluated_args[1].to_string();
                    let is_self = sheet_name == self.name;

                    if let Some(ctx) = context {
                        if let Some(sheet) = ctx.sheets.get(&sheet_name) {
                            if let Some(col_idx) =
                                sheet.columns.iter().position(|c| c.name == col_name)
                            {
                                if is_self {
                                    deps.push(Dependency::LocalColumn(col_idx));
                                } else {
                                    deps.push(Dependency::RemoteColumn {
                                        sheet: sheet_name.clone(),
                                        col: col_idx,
                                    });
                                }
                                let mut results = Vec::new();
                                for row in 0..sheet.row_count() {
                                    let cell_ref = CellRef::new(row, col_idx);
                                    results.push(sheet.get_result_data(&cell_ref));
                                }
                                Ok(ResultData::List(results))
                            } else {
                                Err(EngineError::EvalError(EvalError::UnknownFunction(format!(
                                    "Column not found: {}",
                                    col_name
                                ))))
                            }
                        } else {
                            Err(EngineError::EvalError(EvalError::UnknownFunction(format!(
                                "Sheet not found: {}",
                                sheet_name
                            ))))
                        }
                    } else if is_self {
                        if let Some(col_idx) = self.columns.iter().position(|c| c.name == col_name)
                        {
                            deps.push(Dependency::LocalColumn(col_idx));
                            let mut results = Vec::new();
                            for row in 0..self.row_count() {
                                let cell_ref = CellRef::new(row, col_idx);
                                results.push(self.get_result_data(&cell_ref));
                            }
                            Ok(ResultData::List(results))
                        } else {
                            Err(EngineError::EvalError(EvalError::UnknownFunction(format!(
                                "Column not found: {}",
                                col_name
                            ))))
                        }
                    } else {
                        Err(EngineError::EvalError(EvalError::UnknownFunction(
                            "No context to resolve sheet reference".to_string(),
                        )))
                    }
                } else {
                    Err(EngineError::EvalError(EvalError::UnknownFunction(
                        "get_col() takes 2 arguments".to_string(),
                    )))
                }
            }
            "ABS" => {
                let val = self.to_f64_arg(evaluated_args.first(), "ABS")?;
                Ok(ResultData::Float(val.abs()))
            }
            "INT" => {
                let val = self.to_f64_arg(evaluated_args.first(), "INT")?;
                Ok(ResultData::Float(val.floor()))
            }
            "ROUND" => {
                let val = self.to_f64_arg(evaluated_args.first(), "ROUND")?;
                let digits = if evaluated_args.len() >= 2 {
                    self.to_f64_arg(evaluated_args.get(1), "ROUND")? as i32
                } else {
                    0
                };
                let factor = 10.0f64.powi(digits);
                let mut scaled = val * factor;
                if scaled.abs() >= 1e-12 {
                    scaled = (scaled * 1e12).round() / 1e12;
                }
                Ok(ResultData::Float(scaled.round() / factor))
            }
            "ROUNDUP" => {
                let val = self.to_f64_arg(evaluated_args.first(), "ROUNDUP")?;
                let digits = if evaluated_args.len() >= 2 {
                    self.to_f64_arg(evaluated_args.get(1), "ROUNDUP")? as i32
                } else {
                    0
                };
                let factor = 10.0f64.powi(digits);
                let mut scaled = val * factor;
                if scaled.abs() >= 1e-12 {
                    scaled = (scaled * 1e12).round() / 1e12;
                }
                let rounded = if val >= 0.0 {
                    scaled.ceil() / factor
                } else {
                    scaled.floor() / factor
                };
                Ok(ResultData::Float(rounded))
            }
            "ROUNDDOWN" => {
                let val = self.to_f64_arg(evaluated_args.first(), "ROUNDDOWN")?;
                let digits = if evaluated_args.len() >= 2 {
                    self.to_f64_arg(evaluated_args.get(1), "ROUNDDOWN")? as i32
                } else {
                    0
                };
                let factor = 10.0f64.powi(digits);
                let mut scaled = val * factor;
                if scaled.abs() >= 1e-12 {
                    scaled = (scaled * 1e12).round() / 1e12;
                }
                let rounded = if val >= 0.0 {
                    scaled.floor() / factor
                } else {
                    scaled.ceil() / factor
                };
                Ok(ResultData::Float(rounded))
            }
            "SLICE" => {
                if evaluated_args.len() == 3 {
                    if let ResultData::List(list) = &evaluated_args[0] {
                        let start = self.to_f64(&evaluated_args[1]).unwrap_or(0.0) as isize;
                        let mut end = self.to_f64(&evaluated_args[2]).unwrap_or(-1.0) as isize;

                        let len = list.len() as isize;
                        let start_idx = if start < 0 {
                            (len + start).max(0)
                        } else {
                            start.min(len)
                        } as usize;

                        if end == -1 {
                            end = len;
                        }
                        let end_idx = if end < 0 {
                            (len + end).max(0)
                        } else {
                            end.min(len)
                        } as usize;

                        let sliced = if start_idx < end_idx && start_idx < list.len() {
                            list[start_idx..end_idx.min(list.len())].to_vec()
                        } else {
                            Vec::new()
                        };
                        Ok(ResultData::List(sliced))
                    } else {
                        Ok(ResultData::None)
                    }
                } else {
                    Err(EngineError::EvalError(EvalError::UnknownFunction(
                        "SLICE requires 3 arguments".to_string(),
                    )))
                }
            }
            "INDEX" => {
                if evaluated_args.len() == 2 {
                    if let ResultData::List(raw) = &evaluated_args[0] {
                        let (list, _) = Self::flatten_row_major(raw.clone());
                        let idx = self.to_f64(&evaluated_args[1]).unwrap_or(0.0) as isize;
                        let len = list.len() as isize;
                        // 1-based like every other INDEX form -- found
                        // via the differential fuzzer (LAMBDA/MAP/
                        // BYROW testing was the first thing to ever
                        // exercise this 2-arg path; the standalone
                        // INDEX fuzz generator only ever used the
                        // 3-arg row/col form) that this returned the
                        // element one past the requested position.
                        let real_idx = if idx < 0 { len + idx } else { idx - 1 };
                        if real_idx >= 0 && real_idx < len {
                            Ok(list[real_idx as usize].clone())
                        } else {
                            Ok(ResultData::None)
                        }
                    } else {
                        Ok(ResultData::None)
                    }
                } else if evaluated_args.len() == 3 {
                    let row_num = self.to_f64(&evaluated_args[1]).unwrap_or(1.0) as isize;
                    let col_num = self.to_f64(&evaluated_args[2]).unwrap_or(1.0) as isize;

                    if let ResultData::List(raw) = &evaluated_args[0] {
                        let (list, nested_cols) = Self::flatten_row_major(raw.clone());
                        let num_cols = match nested_cols {
                            Some(c) => c as isize,
                            None => match &args[0] {
                                Expr::RangeRef {
                                    start_col, end_col, ..
                                } => (end_col - start_col + 1) as isize,
                                Expr::FunctionCall { name, args: fargs } => self
                                    .function_call_cols(name, fargs, context, row, col, deps, scope)
                                    .unwrap_or(1)
                                    as isize,
                                _ => 1,
                            },
                        };
                        let r_idx = row_num - 1;
                        let c_idx = col_num - 1;
                        let flat_idx = r_idx * num_cols + c_idx;
                        if flat_idx >= 0 && flat_idx < list.len() as isize {
                            Ok(list[flat_idx as usize].clone())
                        } else {
                            Ok(ResultData::None)
                        }
                    } else {
                        Ok(ResultData::None)
                    }
                } else {
                    Err(EngineError::EvalError(EvalError::UnknownFunction(
                        "INDEX requires 2 or 3 arguments".to_string(),
                    )))
                }
            }
            "GET_COL_IDX" => {
                if evaluated_args.len() == 2 {
                    let sheet_name = evaluated_args[0].to_string();
                    let col_idx = self.to_f64(&evaluated_args[1]).unwrap_or(0.0) as usize;
                    let is_self = sheet_name == self.name;

                    if let Some(ctx) = context {
                        if let Some(sheet) = ctx.sheets.get(&sheet_name) {
                            if is_self {
                                deps.push(Dependency::LocalColumn(col_idx));
                            } else {
                                deps.push(Dependency::RemoteColumn {
                                    sheet: sheet_name.clone(),
                                    col: col_idx,
                                });
                            }
                            let mut results = Vec::new();
                            for row in 0..sheet.row_count() {
                                let cell_ref = CellRef::new(row, col_idx);
                                results.push(sheet.get_result_data(&cell_ref));
                            }
                            Ok(ResultData::List(results))
                        } else {
                            Err(EngineError::EvalError(EvalError::UnknownFunction(format!(
                                "Sheet not found: {}",
                                sheet_name
                            ))))
                        }
                    } else if is_self {
                        deps.push(Dependency::LocalColumn(col_idx));
                        let mut results = Vec::new();
                        for row in 0..self.row_count() {
                            let cell_ref = CellRef::new(row, col_idx);
                            results.push(self.get_result_data(&cell_ref));
                        }
                        Ok(ResultData::List(results))
                    } else {
                        Err(EngineError::EvalError(EvalError::UnknownFunction(
                            "No context to resolve sheet reference".to_string(),
                        )))
                    }
                } else if evaluated_args.len() == 1 {
                    let col_idx = self.to_f64(&evaluated_args[0]).unwrap_or(0.0) as usize;
                    deps.push(Dependency::LocalColumn(col_idx));
                    let mut results = Vec::new();
                    for row in 0..self.row_count() {
                        let cell_ref = CellRef::new(row, col_idx);
                        results.push(self.get_result_data(&cell_ref));
                    }
                    Ok(ResultData::List(results))
                } else {
                    Err(EngineError::EvalError(EvalError::UnknownFunction(
                        "GET_COL_IDX requires 1 or 2 arguments".to_string(),
                    )))
                }
            }
            "COUNTA" => {
                let mut count = 0;
                for arg in evaluated_args {
                    count += self.counta_helper(arg);
                }
                Ok(ResultData::Float(count as f64))
            }
            "CONCAT" | "CONCATENATE" => {
                let mut out = String::new();
                for arg in evaluated_args {
                    self.concat_helper(arg, &mut out);
                }
                Ok(ResultData::String(out))
            }

            "AND" => {
                if evaluated_args.is_empty() {
                    return Ok(ResultData::Boolean(true));
                }
                let mut res = true;
                let mut first_err = None;
                for arg in evaluated_args {
                    match arg {
                        ResultData::Error(e) => {
                            if first_err.is_none() {
                                first_err = Some(ResultData::Error(e.clone()));
                            }
                        }
                        ResultData::List(list) => {
                            for item in list {
                                if let ResultData::Error(e) = item {
                                    if first_err.is_none() {
                                        first_err = Some(ResultData::Error(e.clone()));
                                    }
                                } else if !self.to_bool(item) {
                                    res = false;
                                    if first_err.is_none() {
                                        return Ok(ResultData::Boolean(false));
                                    }
                                    break;
                                }
                            }
                        }
                        other => {
                            if !self.to_bool(other) {
                                res = false;
                                if first_err.is_none() {
                                    return Ok(ResultData::Boolean(false));
                                }
                            }
                        }
                    }
                }
                if let Some(err) = first_err {
                    Ok(err)
                } else {
                    Ok(ResultData::Boolean(res))
                }
            }
            "OR" => {
                if evaluated_args.is_empty() {
                    return Ok(ResultData::Boolean(false));
                }
                let mut res = false;
                let mut first_err = None;
                for arg in evaluated_args {
                    match arg {
                        ResultData::Error(e) => {
                            if first_err.is_none() {
                                first_err = Some(ResultData::Error(e.clone()));
                            }
                        }
                        ResultData::List(list) => {
                            for item in list {
                                if let ResultData::Error(e) = item {
                                    if first_err.is_none() {
                                        first_err = Some(ResultData::Error(e.clone()));
                                    }
                                } else if self.to_bool(item) {
                                    res = true;
                                    if first_err.is_none() {
                                        return Ok(ResultData::Boolean(true));
                                    }
                                    break;
                                }
                            }
                        }
                        other => {
                            if self.to_bool(other) {
                                res = true;
                                if first_err.is_none() {
                                    return Ok(ResultData::Boolean(true));
                                }
                            }
                        }
                    }
                }
                if let Some(err) = first_err {
                    Ok(err)
                } else {
                    Ok(ResultData::Boolean(res))
                }
            }
            "TRUE" => Ok(ResultData::Boolean(true)),
            "FALSE" => Ok(ResultData::Boolean(false)),
            "NOT" => {
                if let Some(err) = Self::find_error_in_args(evaluated_args) {
                    return Ok(err);
                }
                let val = evaluated_args.first().ok_or_else(|| {
                    EngineError::EvalError(EvalError::UnknownFunction(
                        "NOT requires 1 argument".to_string(),
                    ))
                })?;
                Ok(ResultData::Boolean(!self.to_bool(val)))
            }
            "LEFT" => {
                if evaluated_args.is_empty() {
                    return Ok(ResultData::String(String::new()));
                }
                let s = evaluated_args[0].to_string();
                let num_chars = if evaluated_args.len() >= 2 {
                    self.to_f64(&evaluated_args[1]).unwrap_or(1.0) as usize
                } else {
                    1
                };
                let prefix: String = s.chars().take(num_chars).collect();
                Ok(ResultData::String(prefix))
            }
            "RIGHT" => {
                if evaluated_args.is_empty() {
                    return Ok(ResultData::String(String::new()));
                }
                let s = evaluated_args[0].to_string();
                let num_chars = if evaluated_args.len() >= 2 {
                    self.to_f64(&evaluated_args[1]).unwrap_or(1.0) as usize
                } else {
                    1
                };
                let total_chars = s.chars().count();
                let skip_chars = total_chars.saturating_sub(num_chars);
                let suffix: String = s.chars().skip(skip_chars).collect();
                Ok(ResultData::String(suffix))
            }
            "MID" => {
                if evaluated_args.len() < 3 {
                    return Err(EngineError::EvalError(EvalError::UnknownFunction(
                        "MID requires 3 arguments".to_string(),
                    )));
                }
                let s = evaluated_args[0].to_string();
                let start_num = self.to_f64(&evaluated_args[1]).unwrap_or(1.0) as usize;
                let num_chars = self.to_f64(&evaluated_args[2]).unwrap_or(0.0) as usize;

                let start_idx = start_num.saturating_sub(1);
                let mid_str: String = s.chars().skip(start_idx).take(num_chars).collect();
                Ok(ResultData::String(mid_str))
            }
            "LEN" => {
                if evaluated_args.is_empty() {
                    return Ok(ResultData::Float(0.0));
                }
                let s = evaluated_args[0].to_string();
                Ok(ResultData::Float(s.chars().count() as f64))
            }
            "TRIM" => {
                if evaluated_args.is_empty() {
                    return Ok(ResultData::String(String::new()));
                }
                let s = evaluated_args[0].to_string();
                let trimmed = s.trim();
                let mut result = String::new();
                let mut last_was_space = false;
                for c in trimmed.chars() {
                    if c.is_whitespace() {
                        if !last_was_space {
                            result.push(' ');
                            last_was_space = true;
                        }
                    } else {
                        result.push(c);
                        last_was_space = false;
                    }
                }
                Ok(ResultData::String(result))
            }
            "UPPER" => {
                if evaluated_args.is_empty() {
                    return Ok(ResultData::None);
                }
                match &evaluated_args[0] {
                    ResultData::None => Ok(ResultData::String(String::new())),
                    ResultData::Error(e) => Ok(ResultData::Error(e.clone())),
                    v => Ok(ResultData::String(v.to_string().to_uppercase())),
                }
            }
            "LOWER" => {
                if evaluated_args.is_empty() {
                    return Ok(ResultData::None);
                }
                match &evaluated_args[0] {
                    ResultData::None => Ok(ResultData::String(String::new())),
                    ResultData::Error(e) => Ok(ResultData::Error(e.clone())),
                    v => Ok(ResultData::String(v.to_string().to_lowercase())),
                }
            }
            "PROPER" => {
                if evaluated_args.is_empty() {
                    return Ok(ResultData::None);
                }
                match &evaluated_args[0] {
                    ResultData::None => Ok(ResultData::String(String::new())),
                    ResultData::Error(e) => Ok(ResultData::Error(e.clone())),
                    v => Ok(ResultData::String(self.proper(&v.to_string()))),
                }
            }
            "ISNUMBER" => {
                if evaluated_args.is_empty() {
                    return Ok(ResultData::Boolean(false));
                }
                match &evaluated_args[0] {
                    ResultData::Float(_) | ResultData::Integer(_) => Ok(ResultData::Boolean(true)),
                    _ => Ok(ResultData::Boolean(false)),
                }
            }
            "ISTEXT" => {
                if evaluated_args.is_empty() {
                    return Ok(ResultData::Boolean(false));
                }
                match &evaluated_args[0] {
                    ResultData::String(_) => Ok(ResultData::Boolean(true)),
                    _ => Ok(ResultData::Boolean(false)),
                }
            }
            "ISBLANK" => {
                if evaluated_args.is_empty() {
                    return Ok(ResultData::Boolean(true));
                }
                match &evaluated_args[0] {
                    ResultData::None => Ok(ResultData::Boolean(true)),
                    // Only an *absent* value is blank. A cell holding the
                    // empty string is text, and so is a formula that returned
                    // "" -- Excel reports ISBLANK as FALSE for both.
                    _ => Ok(ResultData::Boolean(false)),
                }
            }
            "ISERROR" => {
                if evaluated_args.is_empty() {
                    return Ok(ResultData::Boolean(false));
                }
                match &evaluated_args[0] {
                    ResultData::Error(_) => Ok(ResultData::Boolean(true)),
                    _ => Ok(ResultData::Boolean(false)),
                }
            }
            "PRODUCT" => {
                if let Some(err) = self.check_arg_errors(evaluated_args, arg_is_direct) {
                    return Ok(err);
                }
                let mut prod = 1.0;
                let mut has_nums = false;
                for (i, arg) in evaluated_args.iter().enumerate() {
                    let is_dir = arg_is_direct.get(i).copied().unwrap_or(false);
                    let (p, h) = self.product_helper(arg, is_dir);
                    if h {
                        prod *= p;
                        has_nums = true;
                    }
                }
                if has_nums {
                    // Excel snaps a formula's result to 15 significant
                    // digits, and that is observable beyond display:
                    // PRODUCT(-35, -0.617, -40, -34) is
                    // 29369.199999999997 in raw f64, and
                    // ROUNDDOWN(.., 2) of it gives 29369.19, but Excel
                    // answers 29369.2 because the snap happens first.
                    //
                    // Crucially it is applied *once*, to the finished
                    // product. Doing it per factor compounds: over
                    // seven factors PRODUCT drifted ~14 ULP and
                    // rendered 189124133819.665 where Excel gives
                    // 189124133819.664.
                    Ok(ResultData::Float(Self::clean_float(prod)))
                } else {
                    Ok(ResultData::Float(0.0))
                }
            }
            "MOD" => {
                if evaluated_args.len() < 2 {
                    return Err(EngineError::EvalError(EvalError::UnknownFunction(
                        "MOD requires 2 arguments".to_string(),
                    )));
                }
                let n = self.to_f64_arg(evaluated_args.first(), "MOD")?;
                let d = self.to_f64_arg(evaluated_args.get(1), "MOD")?;
                if d == 0.0 {
                    return Ok(ResultData::Error("#DIV/0!".to_string()));
                }
                // Excel gives up once the quotient gets large enough
                // that `n - d * INT(n / d)` stops being meaningful, and
                // reports #NUM! rather than a number built out of noise
                // -- MOD(28^31, 3) is #NUM! there.
                //
                // The cutoff is on the quotient, not on either operand
                // (MOD(1E15, 1E7) is fine, MOD(1E13, 3) is not), and is
                // identical for different divisors. Bisected against
                // real Excel to between 1.024 and 1.026 times 2^40; the
                // exact constant isn't a round number and isn't worth
                // more probes, so this uses 2^40. That is very slightly
                // conservative -- inside that 0.2%-wide band visi
                // reports #NUM! a little before Excel does -- but it is
                // right everywhere else, which is where the quotients
                // that actually turn up land.
                const MOD_QUOTIENT_LIMIT: f64 = 1_099_511_627_776.0; // 2^40
                let quotient = n / d;
                if !quotient.is_finite() || quotient.abs() > MOD_QUOTIENT_LIMIT {
                    return Ok(ResultData::Error("#NUM!".to_string()));
                }
                let val = n - d * quotient.floor();
                Ok(ResultData::Float(val))
            }
            "TODAY" => {
                let ((y, m, d), _) = self.get_ymd_hms();
                Ok(ResultData::String(format!("{:04}-{:02}-{:02}", y, m, d)))
            }
            "NOW" => {
                let ((y, m, d), (hr, min, sec)) = self.get_ymd_hms();
                Ok(ResultData::String(format!(
                    "{:04}-{:02}-{:02} {:02}:{:02}:{:02}",
                    y, m, d, hr, min, sec
                )))
            }
            "MATCH" => {
                if evaluated_args.len() < 2 {
                    return Err(EngineError::EvalError(EvalError::UnknownFunction(
                        "MATCH requires at least 2 arguments".to_string(),
                    )));
                }
                let lookup_val = &evaluated_args[0];
                let match_type = if evaluated_args.len() >= 3 {
                    self.to_f64(&evaluated_args[2]).unwrap_or(1.0) as isize
                } else {
                    1
                };

                if let ResultData::List(list) = &evaluated_args[1] {
                    let mut match_idx: Option<usize> = None;
                    if match_type == 0 {
                        for (idx, item) in list.iter().enumerate() {
                            if Self::exact_lookup_matches(lookup_val, item) {
                                match_idx = Some(idx);
                                break;
                            }
                        }
                    } else if match_type == 1 {
                        let lookup_f = self.to_f64(lookup_val).unwrap_or(0.0);
                        for (idx, item) in list.iter().enumerate() {
                            let item_f = self.to_f64(item).unwrap_or(0.0);
                            if item_f <= lookup_f {
                                match_idx = Some(idx);
                            } else {
                                break;
                            }
                        }
                    } else if match_type == -1 {
                        let lookup_f = self.to_f64(lookup_val).unwrap_or(0.0);
                        for (idx, item) in list.iter().enumerate() {
                            let item_f = self.to_f64(item).unwrap_or(0.0);
                            if item_f >= lookup_f {
                                match_idx = Some(idx);
                            } else {
                                break;
                            }
                        }
                    }

                    if let Some(idx) = match_idx {
                        Ok(ResultData::Integer((idx + 1) as i64))
                    } else {
                        Ok(ResultData::Error("#N/A".to_string()))
                    }
                } else {
                    Ok(ResultData::Error("#N/A".to_string()))
                }
            }
            "LOOKUP" => {
                if evaluated_args.len() < 2 {
                    return Err(EngineError::EvalError(EvalError::UnknownFunction(
                        "LOOKUP requires at least 2 arguments".to_string(),
                    )));
                }
                let lookup_val = &evaluated_args[0];
                let lookup_vec: Vec<ResultData> = match &evaluated_args[1] {
                    ResultData::List(l) => l.clone(),
                    other => vec![other.clone()],
                };
                let result_vec: Vec<ResultData> = match evaluated_args.get(2) {
                    Some(ResultData::List(l)) => l.clone(),
                    Some(other) => vec![other.clone()],
                    None => lookup_vec.clone(),
                };
                let lookup_f = self.to_f64(lookup_val);
                let mut match_idx: Option<usize> = None;
                for (idx, item) in lookup_vec.iter().enumerate() {
                    let is_match = match lookup_f {
                        Some(lf) => self.to_f64(item).map(|f| f <= lf).unwrap_or(false),
                        None => item.to_string() <= lookup_val.to_string(),
                    };
                    if is_match {
                        match_idx = Some(idx);
                    } else {
                        break;
                    }
                }
                match match_idx {
                    Some(idx) => Ok(result_vec
                        .get(idx)
                        .cloned()
                        .unwrap_or(ResultData::Error("#N/A".to_string()))),
                    None => Ok(ResultData::Error("#N/A".to_string())),
                }
            }
            "XMATCH" => {
                if evaluated_args.len() < 2 {
                    return Err(EngineError::EvalError(EvalError::UnknownFunction(
                        "XMATCH requires at least 2 arguments".to_string(),
                    )));
                }
                let lookup_val = &evaluated_args[0];
                let arr: Vec<ResultData> = match &evaluated_args[1] {
                    ResultData::List(l) => l.clone(),
                    other => vec![other.clone()],
                };
                // search_mode (a 4th argument) isn't supported beyond
                // the default forward linear search, nor is wildcard
                // match_mode (2).
                let match_mode = evaluated_args
                    .get(2)
                    .and_then(|v| self.to_f64(v))
                    .unwrap_or(0.0) as isize;
                let found = match match_mode {
                    -1 | 1 => {
                        let lf = self.to_f64(lookup_val).unwrap_or(0.0);
                        let mut best: Option<(usize, f64)> = None;
                        for (i, item) in arr.iter().enumerate() {
                            if let Some(f) = self.to_f64(item) {
                                let candidate = if match_mode == -1 { f <= lf } else { f >= lf };
                                let better = match best {
                                    Some((_, bf)) => {
                                        if match_mode == -1 {
                                            f > bf
                                        } else {
                                            f < bf
                                        }
                                    }
                                    None => true,
                                };
                                if candidate && better {
                                    best = Some((i, f));
                                }
                            }
                        }
                        best.map(|(i, _)| i)
                    }
                    // XMATCH deliberately does NOT use
                    // exact_lookup_matches: unlike MATCH/VLOOKUP,
                    // real Excel's XMATCH *does* match a blank lookup
                    // value against a blank cell (XMATCH over a blank
                    // A1 in A1:A4 returns 1 where MATCH returns #N/A).
                    _ => arr
                        .iter()
                        .position(|item| item.to_string() == lookup_val.to_string()),
                };
                match found {
                    Some(i) => Ok(ResultData::Float((i + 1) as f64)),
                    None => Ok(ResultData::Error("#N/A".to_string())),
                }
            }
            "VLOOKUP" => {
                if evaluated_args.len() < 3 {
                    return Err(EngineError::EvalError(EvalError::UnknownFunction(
                        "VLOOKUP requires at least 3 arguments".to_string(),
                    )));
                }
                let lookup_val = &evaluated_args[0];
                let col_idx = self.to_f64(&evaluated_args[2]).unwrap_or(1.0) as usize - 1;
                let range_lookup = if evaluated_args.len() >= 4 {
                    self.to_bool(&evaluated_args[3])
                } else {
                    true
                };

                if let ResultData::List(list) = &evaluated_args[1] {
                    let num_cols = match &args[1] {
                        Expr::RangeRef {
                            start_col, end_col, ..
                        } => end_col - start_col + 1,
                        _ => 1,
                    };

                    let num_rows = list.len() / num_cols;
                    if num_rows == 0 || num_cols == 0 {
                        return Ok(ResultData::Error("#N/A".to_string()));
                    }

                    let mut found_row_idx: Option<usize> = None;
                    if !range_lookup {
                        for r in 0..num_rows {
                            let first_col_val = &list[r * num_cols];
                            if Self::exact_lookup_matches(lookup_val, first_col_val) {
                                found_row_idx = Some(r);
                                break;
                            }
                        }
                    } else {
                        let lookup_f = self.to_f64(lookup_val).unwrap_or(0.0);
                        for r in 0..num_rows {
                            let first_col_val = &list[r * num_cols];
                            let val_f = self.to_f64(first_col_val).unwrap_or(0.0);
                            if val_f <= lookup_f {
                                found_row_idx = Some(r);
                            } else {
                                break;
                            }
                        }
                    }

                    if let Some(r) = found_row_idx {
                        if col_idx < num_cols {
                            Ok(list[r * num_cols + col_idx].clone())
                        } else {
                            Ok(ResultData::Error("#REF!".to_string()))
                        }
                    } else {
                        Ok(ResultData::Error("#N/A".to_string()))
                    }
                } else {
                    Ok(ResultData::Error("#N/A".to_string()))
                }
            }
            "XLOOKUP" => {
                if evaluated_args.len() < 3 {
                    return Err(EngineError::EvalError(EvalError::UnknownFunction(
                        "XLOOKUP requires at least 3 arguments".to_string(),
                    )));
                }
                let lookup_val = &evaluated_args[0];
                let if_not_found = if evaluated_args.len() >= 4 {
                    evaluated_args[3].clone()
                } else {
                    ResultData::Error("#N/A".to_string())
                };
                let search_mode = if evaluated_args.len() >= 6 {
                    self.to_f64(&evaluated_args[5]).unwrap_or(1.0) as isize
                } else {
                    1
                };

                if let (ResultData::List(lookup_list), ResultData::List(return_list)) =
                    (&evaluated_args[1], &evaluated_args[2])
                {
                    let mut found_idx: Option<usize> = None;
                    let len = lookup_list.len();

                    let iter_indices: Vec<usize> = if search_mode == -1 {
                        (0..len).rev().collect()
                    } else {
                        (0..len).collect()
                    };

                    for idx in iter_indices {
                        // Like XMATCH (and unlike VLOOKUP/MATCH),
                        // XLOOKUP matches a blank lookup value against
                        // a blank cell rather than reporting #N/A.
                        if lookup_list[idx].to_string() == lookup_val.to_string() {
                            found_idx = Some(idx);
                            break;
                        }
                    }

                    if let Some(idx) = found_idx {
                        if idx < return_list.len() {
                            Ok(return_list[idx].clone())
                        } else {
                            Ok(ResultData::Error("#REF!".to_string()))
                        }
                    } else {
                        Ok(if_not_found)
                    }
                } else {
                    Ok(if_not_found)
                }
            }
            "SUMIF" => {
                if evaluated_args.len() < 2 {
                    return Err(EngineError::EvalError(EvalError::UnknownFunction(
                        "SUMIF requires at least 2 arguments".to_string(),
                    )));
                }
                let range_list = match &evaluated_args[0] {
                    ResultData::List(l) => l,
                    _ => return Ok(ResultData::Float(0.0)),
                };
                let criteria = &evaluated_args[1];
                let sum_list = if evaluated_args.len() >= 3 {
                    match &evaluated_args[2] {
                        ResultData::List(l) => l,
                        _ => range_list,
                    }
                } else {
                    range_list
                };

                let mut sum = 0.0;
                for idx in 0..range_list.len() {
                    if idx < sum_list.len() && self.match_criteria(&range_list[idx], criteria) {
                        sum += Self::aggregate_range_number(&sum_list[idx]).unwrap_or(0.0);
                    }
                }
                Ok(ResultData::Float(sum))
            }
            "SUMIFS" => {
                if evaluated_args.len() < 3 || evaluated_args.len() % 2 == 0 {
                    return Err(EngineError::EvalError(EvalError::UnknownFunction(
                        "SUMIFS requires sum_range and at least one criteria_range/criteria pair"
                            .to_string(),
                    )));
                }
                let sum_list = match &evaluated_args[0] {
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

                let mut sum = 0.0;
                for idx in 0..sum_list.len() {
                    let mut all_match = true;
                    for (crit_range, crit_val) in &criteria_pairs {
                        if idx >= crit_range.len()
                            || !self.match_criteria(&crit_range[idx], crit_val)
                        {
                            all_match = false;
                            break;
                        }
                    }
                    if all_match {
                        sum += Self::aggregate_range_number(&sum_list[idx]).unwrap_or(0.0);
                    }
                }
                Ok(ResultData::Float(sum))
            }
            "COUNTIF" => {
                if evaluated_args.len() < 2 {
                    return Err(EngineError::EvalError(EvalError::UnknownFunction(
                        "COUNTIF requires 2 arguments".to_string(),
                    )));
                }
                let range_list = match &evaluated_args[0] {
                    ResultData::List(l) => l,
                    _ => return Ok(ResultData::Float(0.0)),
                };
                let criteria = &evaluated_args[1];
                let mut count = 0;
                for val in range_list {
                    if self.match_criteria(val, criteria) {
                        count += 1;
                    }
                }
                Ok(ResultData::Float(count as f64))
            }
            "COUNTIFS" => {
                if evaluated_args.len() < 2 || evaluated_args.len() % 2 != 0 {
                    return Err(EngineError::EvalError(EvalError::UnknownFunction(
                        "COUNTIFS requires at least one criteria_range/criteria pair".to_string(),
                    )));
                }
                let mut criteria_pairs = Vec::new();
                let mut i = 0;
                while i < evaluated_args.len() {
                    let crit_range = match &evaluated_args[i] {
                        ResultData::List(l) => l,
                        _ => return Ok(ResultData::Float(0.0)),
                    };
                    let crit_val = &evaluated_args[i + 1];
                    criteria_pairs.push((crit_range, crit_val));
                    i += 2;
                }

                let mut count = 0;
                let first_len = criteria_pairs[0].0.len();
                for idx in 0..first_len {
                    let mut all_match = true;
                    for (crit_range, crit_val) in &criteria_pairs {
                        if idx >= crit_range.len()
                            || !self.match_criteria(&crit_range[idx], crit_val)
                        {
                            all_match = false;
                            break;
                        }
                    }
                    if all_match {
                        count += 1;
                    }
                }
                Ok(ResultData::Float(count as f64))
            }
            "MMULT" => {
                if evaluated_args.len() < 2 {
                    return Err(EngineError::EvalError(EvalError::UnknownFunction(
                        "MMULT requires 2 arguments".to_string(),
                    )));
                }

                if let (ResultData::List(list1), ResultData::List(list2)) =
                    (&evaluated_args[0], &evaluated_args[1])
                {
                    let (rows1, cols1) = match &args[0] {
                        Expr::RangeRef {
                            sheet,
                            start_row,
                            end_row,
                            start_col,
                            end_col,
                            ..
                        } => {
                            let is_self = match sheet {
                                Some(name) => name == &self.name,
                                None => true,
                            };
                            let actual_end = if *end_row == usize::MAX {
                                if is_self {
                                    self.row_count().saturating_sub(1)
                                } else if let Some(ctx) = context {
                                    ctx.sheets
                                        .get(sheet.as_ref().unwrap())
                                        .map(|t| t.row_count().saturating_sub(1))
                                        .unwrap_or(0)
                                } else {
                                    0
                                }
                            } else {
                                *end_row
                            };
                            (actual_end - start_row + 1, end_col - start_col + 1)
                        }
                        _ => (1, list1.len()),
                    };

                    let (rows2, cols2) = match &args[1] {
                        Expr::RangeRef {
                            sheet,
                            start_row,
                            end_row,
                            start_col,
                            end_col,
                            ..
                        } => {
                            let is_self = match sheet {
                                Some(name) => name == &self.name,
                                None => true,
                            };
                            let actual_end = if *end_row == usize::MAX {
                                if is_self {
                                    self.row_count().saturating_sub(1)
                                } else if let Some(ctx) = context {
                                    ctx.sheets
                                        .get(sheet.as_ref().unwrap())
                                        .map(|t| t.row_count().saturating_sub(1))
                                        .unwrap_or(0)
                                } else {
                                    0
                                }
                            } else {
                                *end_row
                            };
                            (actual_end - start_row + 1, end_col - start_col + 1)
                        }
                        _ => {
                            if list2.len() == cols1 {
                                (cols1, 1)
                            } else {
                                (1, list2.len())
                            }
                        }
                    };

                    if cols1 != rows2 {
                        return Ok(ResultData::Error("#VALUE!".to_string()));
                    }

                    // A non-numeric cell anywhere in either operand
                    // makes the whole call #VALUE! in real Excel, not
                    // a silent 0 -- MMULT doesn't ignore text the way
                    // SUM/AVERAGE-style aggregates do.
                    fn as_plain_number(v: &ResultData) -> Option<f64> {
                        match v {
                            ResultData::Float(f) => Some(*f),
                            ResultData::Integer(i) => Some(*i as f64),
                            _ => None,
                        }
                    }
                    let mut result_list = Vec::with_capacity(rows1 * cols2);
                    for r in 0..rows1 {
                        for c in 0..cols2 {
                            let mut val = 0.0;
                            for k in 0..cols1 {
                                // Only a real number is acceptable --
                                // MMULT rejects text, booleans and
                                // blanks alike (all confirmed #VALUE!
                                // against real Excel), so this can't
                                // use to_f64's lenient coercion.
                                let (Some(v1), Some(v2)) = (
                                    as_plain_number(&list1[r * cols1 + k]),
                                    as_plain_number(&list2[k * cols2 + c]),
                                ) else {
                                    return Ok(ResultData::Error("#VALUE!".to_string()));
                                };
                                val += v1 * v2;
                            }
                            result_list.push(ResultData::Float(val));
                        }
                    }
                    Ok(ResultData::List(result_list))
                } else {
                    Ok(ResultData::Error("#VALUE!".to_string()))
                }
            }
            "PV" => {
                let rate = self.to_f64_arg(evaluated_args.first(), "PV")?;
                let nper = self.to_f64_arg(evaluated_args.get(1), "PV")?;
                let pmt = self.to_f64_arg(evaluated_args.get(2), "PV")?;
                let fv = self.opt_f64(evaluated_args, 3, 0.0);
                let pmt_type = self.opt_f64(evaluated_args, 4, 0.0);
                Ok(ResultData::Float(finance::pv(
                    rate, nper, pmt, fv, pmt_type,
                )))
            }
            "FV" => {
                let rate = self.to_f64_arg(evaluated_args.first(), "FV")?;
                let nper = self.to_f64_arg(evaluated_args.get(1), "FV")?;
                let pmt = self.to_f64_arg(evaluated_args.get(2), "FV")?;
                let pv = self.opt_f64(evaluated_args, 3, 0.0);
                let pmt_type = self.opt_f64(evaluated_args, 4, 0.0);
                Ok(ResultData::Float(finance::fv(
                    rate, nper, pmt, pv, pmt_type,
                )))
            }
            "PMT" => {
                let rate = self.to_f64_arg(evaluated_args.first(), "PMT")?;
                let nper = self.to_f64_arg(evaluated_args.get(1), "PMT")?;
                let pv = self.to_f64_arg(evaluated_args.get(2), "PMT")?;
                let fv = self.opt_f64(evaluated_args, 3, 0.0);
                let pmt_type = self.opt_f64(evaluated_args, 4, 0.0);
                Ok(ResultData::Float(finance::pmt(
                    rate, nper, pv, fv, pmt_type,
                )))
            }
            "NPER" => {
                let rate = self.to_f64_arg(evaluated_args.first(), "NPER")?;
                let pmt = self.to_f64_arg(evaluated_args.get(1), "NPER")?;
                let pv = self.to_f64_arg(evaluated_args.get(2), "NPER")?;
                let fv = self.opt_f64(evaluated_args, 3, 0.0);
                let pmt_type = self.opt_f64(evaluated_args, 4, 0.0);
                match finance::nper(rate, pmt, pv, fv, pmt_type) {
                    Some(v) => Ok(ResultData::Float(v)),
                    None => Ok(ResultData::Error("#NUM!".to_string())),
                }
            }
            "RATE" => {
                let nper = self.to_f64_arg(evaluated_args.first(), "RATE")?;
                let pmt = self.to_f64_arg(evaluated_args.get(1), "RATE")?;
                let pv = self.to_f64_arg(evaluated_args.get(2), "RATE")?;
                let fv = self.opt_f64(evaluated_args, 3, 0.0);
                let pmt_type = self.opt_f64(evaluated_args, 4, 0.0);
                let guess = self.opt_f64(evaluated_args, 5, 0.1);
                match finance::rate(nper, pmt, pv, fv, pmt_type, guess) {
                    Some(v) => Ok(ResultData::Float(v)),
                    None => Ok(ResultData::Error("#NUM!".to_string())),
                }
            }
            "IPMT" => {
                let rate = self.to_f64_arg(evaluated_args.first(), "IPMT")?;
                let per = self.to_f64_arg(evaluated_args.get(1), "IPMT")?;
                let nper = self.to_f64_arg(evaluated_args.get(2), "IPMT")?;
                let pv = self.to_f64_arg(evaluated_args.get(3), "IPMT")?;
                let fv = self.opt_f64(evaluated_args, 4, 0.0);
                let pmt_type = self.opt_f64(evaluated_args, 5, 0.0);
                Ok(ResultData::Float(finance::ipmt(
                    rate, per, nper, pv, fv, pmt_type,
                )))
            }
            "PPMT" => {
                let rate = self.to_f64_arg(evaluated_args.first(), "PPMT")?;
                let per = self.to_f64_arg(evaluated_args.get(1), "PPMT")?;
                let nper = self.to_f64_arg(evaluated_args.get(2), "PPMT")?;
                let pv = self.to_f64_arg(evaluated_args.get(3), "PPMT")?;
                let fv = self.opt_f64(evaluated_args, 4, 0.0);
                let pmt_type = self.opt_f64(evaluated_args, 5, 0.0);
                Ok(ResultData::Float(finance::ppmt(
                    rate, per, nper, pv, fv, pmt_type,
                )))
            }
            "CUMIPMT" => {
                let rate = self.to_f64_arg(evaluated_args.first(), "CUMIPMT")?;
                let nper = self.to_f64_arg(evaluated_args.get(1), "CUMIPMT")?;
                let pv = self.to_f64_arg(evaluated_args.get(2), "CUMIPMT")?;
                let start = self.to_f64_arg(evaluated_args.get(3), "CUMIPMT")?;
                let end = self.to_f64_arg(evaluated_args.get(4), "CUMIPMT")?;
                let pmt_type = self.to_f64_arg(evaluated_args.get(5), "CUMIPMT")?;
                Ok(ResultData::Float(finance::cumipmt(
                    rate, nper, pv, start, end, pmt_type,
                )))
            }
            "CUMPRINC" => {
                let rate = self.to_f64_arg(evaluated_args.first(), "CUMPRINC")?;
                let nper = self.to_f64_arg(evaluated_args.get(1), "CUMPRINC")?;
                let pv = self.to_f64_arg(evaluated_args.get(2), "CUMPRINC")?;
                let start = self.to_f64_arg(evaluated_args.get(3), "CUMPRINC")?;
                let end = self.to_f64_arg(evaluated_args.get(4), "CUMPRINC")?;
                let pmt_type = self.to_f64_arg(evaluated_args.get(5), "CUMPRINC")?;
                Ok(ResultData::Float(finance::cumprinc(
                    rate, nper, pv, start, end, pmt_type,
                )))
            }
            "NPV" => {
                let rate = self.to_f64_arg(evaluated_args.first(), "NPV")?;
                let mut values = Vec::new();
                for (i, arg) in evaluated_args.iter().enumerate().skip(1) {
                    values.extend(self.flatten_finance_numbers(arg, arg_is_direct[i]));
                }
                Ok(ResultData::Float(finance::npv(rate, &values)))
            }
            "IRR" => {
                let values = evaluated_args
                    .first()
                    .map(|v| self.flatten_finance_numbers(v, arg_is_direct[0]))
                    .unwrap_or_default();
                let guess = self.opt_f64(evaluated_args, 1, 0.1);
                match finance::irr(&values, guess) {
                    Some(v) => Ok(ResultData::Float(v)),
                    None => Ok(ResultData::Error("#NUM!".to_string())),
                }
            }
            "MIRR" => {
                let values = evaluated_args
                    .first()
                    .map(|v| self.flatten_finance_numbers(v, arg_is_direct[0]))
                    .unwrap_or_default();
                let finance_rate = self.to_f64_arg(evaluated_args.get(1), "MIRR")?;
                let reinvest_rate = self.to_f64_arg(evaluated_args.get(2), "MIRR")?;
                match finance::mirr(&values, finance_rate, reinvest_rate) {
                    Some(v) => Ok(ResultData::Float(v)),
                    None => Ok(ResultData::Error("#NUM!".to_string())),
                }
            }
            "XNPV" => {
                let rate = self.to_f64_arg(evaluated_args.first(), "XNPV")?;
                let values = evaluated_args
                    .get(1)
                    .map(|v| self.flatten_finance_numbers(v, arg_is_direct[1]))
                    .unwrap_or_default();
                let dates = evaluated_args
                    .get(2)
                    .map(|v| self.flatten_finance_numbers(v, arg_is_direct[2]))
                    .unwrap_or_default();
                if values.is_empty() || values.len() != dates.len() {
                    return Ok(ResultData::Error("#NUM!".to_string()));
                }
                Ok(ResultData::Float(finance::xnpv(rate, &values, &dates)))
            }
            "XIRR" => {
                let values = evaluated_args
                    .first()
                    .map(|v| self.flatten_finance_numbers(v, arg_is_direct[0]))
                    .unwrap_or_default();
                let dates = evaluated_args
                    .get(1)
                    .map(|v| self.flatten_finance_numbers(v, arg_is_direct[1]))
                    .unwrap_or_default();
                let guess = self.opt_f64(evaluated_args, 2, 0.1);
                if values.is_empty() || values.len() != dates.len() {
                    return Ok(ResultData::Error("#NUM!".to_string()));
                }
                match finance::xirr(&values, &dates, guess) {
                    Some(v) => Ok(ResultData::Float(v)),
                    None => Ok(ResultData::Error("#NUM!".to_string())),
                }
            }
            "SLN" => {
                let cost = self.to_f64_arg(evaluated_args.first(), "SLN")?;
                let salvage = self.to_f64_arg(evaluated_args.get(1), "SLN")?;
                let life = self.to_f64_arg(evaluated_args.get(2), "SLN")?;
                Ok(ResultData::Float(finance::sln(cost, salvage, life)))
            }
            "SYD" => {
                let cost = self.to_f64_arg(evaluated_args.first(), "SYD")?;
                let salvage = self.to_f64_arg(evaluated_args.get(1), "SYD")?;
                let life = self.to_f64_arg(evaluated_args.get(2), "SYD")?;
                let per = self.to_f64_arg(evaluated_args.get(3), "SYD")?;
                Ok(ResultData::Float(finance::syd(cost, salvage, life, per)))
            }
            "DB" => {
                let cost = self.to_f64_arg(evaluated_args.first(), "DB")?;
                let salvage = self.to_f64_arg(evaluated_args.get(1), "DB")?;
                let life = self.to_f64_arg(evaluated_args.get(2), "DB")?;
                let period = self.to_f64_arg(evaluated_args.get(3), "DB")?;
                let month = self.opt_f64(evaluated_args, 4, 12.0);
                Ok(ResultData::Float(finance::db(
                    cost, salvage, life, period, month,
                )))
            }
            "DDB" => {
                let cost = self.to_f64_arg(evaluated_args.first(), "DDB")?;
                let salvage = self.to_f64_arg(evaluated_args.get(1), "DDB")?;
                let life = self.to_f64_arg(evaluated_args.get(2), "DDB")?;
                let period = self.to_f64_arg(evaluated_args.get(3), "DDB")?;
                let factor = self.opt_f64(evaluated_args, 4, 2.0);
                Ok(ResultData::Float(finance::ddb(
                    cost, salvage, life, period, factor,
                )))
            }
            "VDB" => {
                let cost = self.to_f64_arg(evaluated_args.first(), "VDB")?;
                let salvage = self.to_f64_arg(evaluated_args.get(1), "VDB")?;
                let life = self.to_f64_arg(evaluated_args.get(2), "VDB")?;
                let start = self.to_f64_arg(evaluated_args.get(3), "VDB")?;
                let end = self.to_f64_arg(evaluated_args.get(4), "VDB")?;
                let factor = self.opt_f64(evaluated_args, 5, 2.0);
                let no_switch = evaluated_args
                    .get(6)
                    .map(|v| self.to_bool(v))
                    .unwrap_or(false);
                match finance::vdb(cost, salvage, life, start, end, factor, no_switch) {
                    Some(v) => Ok(ResultData::Float(v)),
                    None => Ok(ResultData::Error("#NUM!".to_string())),
                }
            }
            "EFFECT" => {
                let nominal_rate = self.to_f64_arg(evaluated_args.first(), "EFFECT")?;
                let npery = self.to_f64_arg(evaluated_args.get(1), "EFFECT")?;
                Ok(ResultData::Float(finance::effect(nominal_rate, npery)))
            }
            "NOMINAL" => {
                let effect_rate = self.to_f64_arg(evaluated_args.first(), "NOMINAL")?;
                let npery = self.to_f64_arg(evaluated_args.get(1), "NOMINAL")?;
                Ok(ResultData::Float(finance::nominal(effect_rate, npery)))
            }
            "DOLLARDE" => {
                let fractional_dollar = self.to_f64_arg(evaluated_args.first(), "DOLLARDE")?;
                let fraction = self.to_f64_arg(evaluated_args.get(1), "DOLLARDE")?;
                match finance::dollarde(fractional_dollar, fraction) {
                    Some(v) => Ok(ResultData::Float(v)),
                    None => Ok(ResultData::Error("#NUM!".to_string())),
                }
            }
            "DOLLARFR" => {
                let decimal_dollar = self.to_f64_arg(evaluated_args.first(), "DOLLARFR")?;
                let fraction = self.to_f64_arg(evaluated_args.get(1), "DOLLARFR")?;
                match finance::dollarfr(decimal_dollar, fraction) {
                    Some(v) => Ok(ResultData::Float(v)),
                    None => Ok(ResultData::Error("#NUM!".to_string())),
                }
            }
            "FVSCHEDULE" => {
                let principal = self.to_f64_arg(evaluated_args.first(), "FVSCHEDULE")?;
                let schedule = evaluated_args
                    .get(1)
                    .map(|v| self.flatten_finance_numbers(v, arg_is_direct[1]))
                    .unwrap_or_default();
                Ok(ResultData::Float(finance::fvschedule(principal, &schedule)))
            }
            "RRI" => {
                let nper = self.to_f64_arg(evaluated_args.first(), "RRI")?;
                let pv = self.to_f64_arg(evaluated_args.get(1), "RRI")?;
                let fv = self.to_f64_arg(evaluated_args.get(2), "RRI")?;
                match finance::rri(nper, pv, fv) {
                    Some(v) => Ok(ResultData::Float(v)),
                    None => Ok(ResultData::Error("#NUM!".to_string())),
                }
            }
            "PDURATION" => {
                let rate = self.to_f64_arg(evaluated_args.first(), "PDURATION")?;
                let pv = self.to_f64_arg(evaluated_args.get(1), "PDURATION")?;
                let fv = self.to_f64_arg(evaluated_args.get(2), "PDURATION")?;
                match finance::pduration(rate, pv, fv) {
                    Some(v) => Ok(ResultData::Float(v)),
                    None => Ok(ResultData::Error("#NUM!".to_string())),
                }
            }
            "ISPMT" => {
                let rate = self.to_f64_arg(evaluated_args.first(), "ISPMT")?;
                let per = self.to_f64_arg(evaluated_args.get(1), "ISPMT")?;
                let nper = self.to_f64_arg(evaluated_args.get(2), "ISPMT")?;
                let pv = self.to_f64_arg(evaluated_args.get(3), "ISPMT")?;
                Ok(ResultData::Float(finance::ispmt(rate, per, nper, pv)))
            }
            "COUPDAYBS" => {
                let settlement = self.to_f64_arg(evaluated_args.first(), "COUPDAYBS")?;
                let maturity = self.to_f64_arg(evaluated_args.get(1), "COUPDAYBS")?;
                let frequency = self.to_f64_arg(evaluated_args.get(2), "COUPDAYBS")?;
                let basis = self.opt_f64(evaluated_args, 3, 0.0);
                Ok(ResultData::Float(finance::coupdaybs(
                    settlement, maturity, frequency, basis,
                )))
            }
            "COUPDAYS" => {
                let settlement = self.to_f64_arg(evaluated_args.first(), "COUPDAYS")?;
                let maturity = self.to_f64_arg(evaluated_args.get(1), "COUPDAYS")?;
                let frequency = self.to_f64_arg(evaluated_args.get(2), "COUPDAYS")?;
                let basis = self.opt_f64(evaluated_args, 3, 0.0);
                Ok(ResultData::Float(finance::coupdays(
                    settlement, maturity, frequency, basis,
                )))
            }
            "COUPDAYSNC" => {
                let settlement = self.to_f64_arg(evaluated_args.first(), "COUPDAYSNC")?;
                let maturity = self.to_f64_arg(evaluated_args.get(1), "COUPDAYSNC")?;
                let frequency = self.to_f64_arg(evaluated_args.get(2), "COUPDAYSNC")?;
                let basis = self.opt_f64(evaluated_args, 3, 0.0);
                Ok(ResultData::Float(finance::coupdaysnc(
                    settlement, maturity, frequency, basis,
                )))
            }
            "COUPNCD" => {
                let settlement = self.to_f64_arg(evaluated_args.first(), "COUPNCD")?;
                let maturity = self.to_f64_arg(evaluated_args.get(1), "COUPNCD")?;
                let frequency = self.to_f64_arg(evaluated_args.get(2), "COUPNCD")?;
                Ok(ResultData::Float(finance::coupncd(
                    settlement, maturity, frequency,
                )))
            }
            "COUPNUM" => {
                let settlement = self.to_f64_arg(evaluated_args.first(), "COUPNUM")?;
                let maturity = self.to_f64_arg(evaluated_args.get(1), "COUPNUM")?;
                let frequency = self.to_f64_arg(evaluated_args.get(2), "COUPNUM")?;
                Ok(ResultData::Float(finance::coupnum(
                    settlement, maturity, frequency,
                )))
            }
            "COUPPCD" => {
                let settlement = self.to_f64_arg(evaluated_args.first(), "COUPPCD")?;
                let maturity = self.to_f64_arg(evaluated_args.get(1), "COUPPCD")?;
                let frequency = self.to_f64_arg(evaluated_args.get(2), "COUPPCD")?;
                Ok(ResultData::Float(finance::couppcd(
                    settlement, maturity, frequency,
                )))
            }
            "PRICE" => {
                let settlement = self.to_f64_arg(evaluated_args.first(), "PRICE")?;
                let maturity = self.to_f64_arg(evaluated_args.get(1), "PRICE")?;
                let rate = self.to_f64_arg(evaluated_args.get(2), "PRICE")?;
                let yld = self.to_f64_arg(evaluated_args.get(3), "PRICE")?;
                let redemption = self.to_f64_arg(evaluated_args.get(4), "PRICE")?;
                let frequency = self.to_f64_arg(evaluated_args.get(5), "PRICE")?;
                let basis = self.opt_f64(evaluated_args, 6, 0.0);
                Ok(ResultData::Float(finance::price(
                    settlement, maturity, rate, yld, redemption, frequency, basis,
                )))
            }
            "YIELD" => {
                let settlement = self.to_f64_arg(evaluated_args.first(), "YIELD")?;
                let maturity = self.to_f64_arg(evaluated_args.get(1), "YIELD")?;
                let rate = self.to_f64_arg(evaluated_args.get(2), "YIELD")?;
                let pr = self.to_f64_arg(evaluated_args.get(3), "YIELD")?;
                let redemption = self.to_f64_arg(evaluated_args.get(4), "YIELD")?;
                let frequency = self.to_f64_arg(evaluated_args.get(5), "YIELD")?;
                let basis = self.opt_f64(evaluated_args, 6, 0.0);
                match finance::yield_(settlement, maturity, rate, pr, redemption, frequency, basis)
                {
                    Some(v) => Ok(ResultData::Float(v)),
                    None => Ok(ResultData::Error("#NUM!".to_string())),
                }
            }
            "DURATION" => {
                let settlement = self.to_f64_arg(evaluated_args.first(), "DURATION")?;
                let maturity = self.to_f64_arg(evaluated_args.get(1), "DURATION")?;
                let coupon = self.to_f64_arg(evaluated_args.get(2), "DURATION")?;
                let yld = self.to_f64_arg(evaluated_args.get(3), "DURATION")?;
                let frequency = self.to_f64_arg(evaluated_args.get(4), "DURATION")?;
                let basis = self.opt_f64(evaluated_args, 5, 0.0);
                Ok(ResultData::Float(finance::duration(
                    settlement, maturity, coupon, yld, frequency, basis,
                )))
            }
            "MDURATION" => {
                let settlement = self.to_f64_arg(evaluated_args.first(), "MDURATION")?;
                let maturity = self.to_f64_arg(evaluated_args.get(1), "MDURATION")?;
                let coupon = self.to_f64_arg(evaluated_args.get(2), "MDURATION")?;
                let yld = self.to_f64_arg(evaluated_args.get(3), "MDURATION")?;
                let frequency = self.to_f64_arg(evaluated_args.get(4), "MDURATION")?;
                let basis = self.opt_f64(evaluated_args, 5, 0.0);
                Ok(ResultData::Float(finance::mduration(
                    settlement, maturity, coupon, yld, frequency, basis,
                )))
            }
            "DISC" => {
                let settlement = self.to_f64_arg(evaluated_args.first(), "DISC")?;
                let maturity = self.to_f64_arg(evaluated_args.get(1), "DISC")?;
                let pr = self.to_f64_arg(evaluated_args.get(2), "DISC")?;
                let redemption = self.to_f64_arg(evaluated_args.get(3), "DISC")?;
                let basis = self.opt_f64(evaluated_args, 4, 0.0);
                Ok(ResultData::Float(finance::disc(
                    settlement, maturity, pr, redemption, basis,
                )))
            }
            "PRICEDISC" => {
                let settlement = self.to_f64_arg(evaluated_args.first(), "PRICEDISC")?;
                let maturity = self.to_f64_arg(evaluated_args.get(1), "PRICEDISC")?;
                let discount = self.to_f64_arg(evaluated_args.get(2), "PRICEDISC")?;
                let redemption = self.to_f64_arg(evaluated_args.get(3), "PRICEDISC")?;
                let basis = self.opt_f64(evaluated_args, 4, 0.0);
                Ok(ResultData::Float(finance::pricedisc(
                    settlement, maturity, discount, redemption, basis,
                )))
            }
            "YIELDDISC" => {
                let settlement = self.to_f64_arg(evaluated_args.first(), "YIELDDISC")?;
                let maturity = self.to_f64_arg(evaluated_args.get(1), "YIELDDISC")?;
                let pr = self.to_f64_arg(evaluated_args.get(2), "YIELDDISC")?;
                let redemption = self.to_f64_arg(evaluated_args.get(3), "YIELDDISC")?;
                let basis = self.opt_f64(evaluated_args, 4, 0.0);
                Ok(ResultData::Float(finance::yielddisc(
                    settlement, maturity, pr, redemption, basis,
                )))
            }
            "PRICEMAT" => {
                let settlement = self.to_f64_arg(evaluated_args.first(), "PRICEMAT")?;
                let maturity = self.to_f64_arg(evaluated_args.get(1), "PRICEMAT")?;
                let issue = self.to_f64_arg(evaluated_args.get(2), "PRICEMAT")?;
                let rate = self.to_f64_arg(evaluated_args.get(3), "PRICEMAT")?;
                let yld = self.to_f64_arg(evaluated_args.get(4), "PRICEMAT")?;
                let basis = self.opt_f64(evaluated_args, 5, 0.0);
                Ok(ResultData::Float(finance::pricemat(
                    settlement, maturity, issue, rate, yld, basis,
                )))
            }
            "YIELDMAT" => {
                let settlement = self.to_f64_arg(evaluated_args.first(), "YIELDMAT")?;
                let maturity = self.to_f64_arg(evaluated_args.get(1), "YIELDMAT")?;
                let issue = self.to_f64_arg(evaluated_args.get(2), "YIELDMAT")?;
                let rate = self.to_f64_arg(evaluated_args.get(3), "YIELDMAT")?;
                let pr = self.to_f64_arg(evaluated_args.get(4), "YIELDMAT")?;
                let basis = self.opt_f64(evaluated_args, 5, 0.0);
                Ok(ResultData::Float(finance::yieldmat(
                    settlement, maturity, issue, rate, pr, basis,
                )))
            }
            "RECEIVED" => {
                let settlement = self.to_f64_arg(evaluated_args.first(), "RECEIVED")?;
                let maturity = self.to_f64_arg(evaluated_args.get(1), "RECEIVED")?;
                let investment = self.to_f64_arg(evaluated_args.get(2), "RECEIVED")?;
                let discount = self.to_f64_arg(evaluated_args.get(3), "RECEIVED")?;
                let basis = self.opt_f64(evaluated_args, 4, 0.0);
                Ok(ResultData::Float(finance::received(
                    settlement, maturity, investment, discount, basis,
                )))
            }
            "INTRATE" => {
                let settlement = self.to_f64_arg(evaluated_args.first(), "INTRATE")?;
                let maturity = self.to_f64_arg(evaluated_args.get(1), "INTRATE")?;
                let investment = self.to_f64_arg(evaluated_args.get(2), "INTRATE")?;
                let redemption = self.to_f64_arg(evaluated_args.get(3), "INTRATE")?;
                let basis = self.opt_f64(evaluated_args, 4, 0.0);
                Ok(ResultData::Float(finance::intrate(
                    settlement, maturity, investment, redemption, basis,
                )))
            }
            "TBILLPRICE" => {
                let settlement = self.to_f64_arg(evaluated_args.first(), "TBILLPRICE")?;
                let maturity = self.to_f64_arg(evaluated_args.get(1), "TBILLPRICE")?;
                let discount = self.to_f64_arg(evaluated_args.get(2), "TBILLPRICE")?;
                Ok(ResultData::Float(finance::tbillprice(
                    settlement, maturity, discount,
                )))
            }
            "TBILLYIELD" => {
                let settlement = self.to_f64_arg(evaluated_args.first(), "TBILLYIELD")?;
                let maturity = self.to_f64_arg(evaluated_args.get(1), "TBILLYIELD")?;
                let pr = self.to_f64_arg(evaluated_args.get(2), "TBILLYIELD")?;
                Ok(ResultData::Float(finance::tbillyield(
                    settlement, maturity, pr,
                )))
            }
            "TBILLEQ" => {
                let settlement = self.to_f64_arg(evaluated_args.first(), "TBILLEQ")?;
                let maturity = self.to_f64_arg(evaluated_args.get(1), "TBILLEQ")?;
                let discount = self.to_f64_arg(evaluated_args.get(2), "TBILLEQ")?;
                match finance::tbilleq(settlement, maturity, discount) {
                    Some(v) => Ok(ResultData::Float(v)),
                    None => Ok(ResultData::Error("#NUM!".to_string())),
                }
            }
            "ACCRINTM" => {
                let issue = self.to_f64_arg(evaluated_args.first(), "ACCRINTM")?;
                let settlement = self.to_f64_arg(evaluated_args.get(1), "ACCRINTM")?;
                let rate = self.to_f64_arg(evaluated_args.get(2), "ACCRINTM")?;
                let par = self.to_f64_arg(evaluated_args.get(3), "ACCRINTM")?;
                let basis = self.opt_f64(evaluated_args, 4, 0.0);
                res_to_rd(finance::accrintm(issue, settlement, rate, par, basis))
            }
            "ACCRINT" => {
                let issue = self.to_f64_arg(evaluated_args.first(), "ACCRINT")?;
                let first_interest = self.to_f64_arg(evaluated_args.get(1), "ACCRINT")?;
                let settlement = self.to_f64_arg(evaluated_args.get(2), "ACCRINT")?;
                let rate = self.to_f64_arg(evaluated_args.get(3), "ACCRINT")?;
                let par = self.to_f64_arg(evaluated_args.get(4), "ACCRINT")?;
                let frequency = self.to_f64_arg(evaluated_args.get(5), "ACCRINT")?;
                let basis = self.opt_f64(evaluated_args, 6, 0.0);
                let calc_method = evaluated_args
                    .get(7)
                    .map(|v| self.to_bool(v))
                    .unwrap_or(true);
                Ok(ResultData::Float(finance::accrint(
                    issue,
                    first_interest,
                    settlement,
                    rate,
                    par,
                    frequency,
                    basis,
                    calc_method,
                )))
            }
            "AMORLINC" => {
                let cost = self.to_f64_arg(evaluated_args.first(), "AMORLINC")?;
                let date_purchased = self.to_f64_arg(evaluated_args.get(1), "AMORLINC")?;
                let first_period = self.to_f64_arg(evaluated_args.get(2), "AMORLINC")?;
                let salvage = self.to_f64_arg(evaluated_args.get(3), "AMORLINC")?;
                let period = self.to_f64_arg(evaluated_args.get(4), "AMORLINC")?;
                let rate = self.to_f64_arg(evaluated_args.get(5), "AMORLINC")?;
                let basis = self.opt_f64(evaluated_args, 6, 0.0);
                res_to_rd(finance::amorlinc(
                    cost,
                    date_purchased,
                    first_period,
                    salvage,
                    period,
                    rate,
                    basis,
                ))
            }
            "AMORDEGRC" => {
                let cost = self.to_f64_arg(evaluated_args.first(), "AMORDEGRC")?;
                let date_purchased = self.to_f64_arg(evaluated_args.get(1), "AMORDEGRC")?;
                let first_period = self.to_f64_arg(evaluated_args.get(2), "AMORDEGRC")?;
                let salvage = self.to_f64_arg(evaluated_args.get(3), "AMORDEGRC")?;
                let period = self.to_f64_arg(evaluated_args.get(4), "AMORDEGRC")?;
                let rate = self.to_f64_arg(evaluated_args.get(5), "AMORDEGRC")?;
                let basis = self.opt_f64(evaluated_args, 6, 0.0);
                res_to_rd(finance::amordegrc(
                    cost,
                    date_purchased,
                    first_period,
                    salvage,
                    period,
                    rate,
                    basis,
                ))
            }
            "ODDFPRICE" => {
                let settlement = self.to_f64_arg(evaluated_args.first(), "ODDFPRICE")?;
                let maturity = self.to_f64_arg(evaluated_args.get(1), "ODDFPRICE")?;
                let issue = self.to_f64_arg(evaluated_args.get(2), "ODDFPRICE")?;
                let first_coupon = self.to_f64_arg(evaluated_args.get(3), "ODDFPRICE")?;
                let rate = self.to_f64_arg(evaluated_args.get(4), "ODDFPRICE")?;
                let yld = self.to_f64_arg(evaluated_args.get(5), "ODDFPRICE")?;
                let redemption = self.to_f64_arg(evaluated_args.get(6), "ODDFPRICE")?;
                let frequency = self.to_f64_arg(evaluated_args.get(7), "ODDFPRICE")?;
                let basis = self.opt_f64(evaluated_args, 8, 0.0);
                if settlement <= issue {
                    return Ok(ResultData::Error("#NUM!".to_string()));
                }
                Ok(ResultData::Float(finance::oddfprice(
                    settlement,
                    maturity,
                    issue,
                    first_coupon,
                    rate,
                    yld,
                    redemption,
                    frequency,
                    basis,
                )))
            }
            "ODDFYIELD" => {
                let settlement = self.to_f64_arg(evaluated_args.first(), "ODDFYIELD")?;
                let maturity = self.to_f64_arg(evaluated_args.get(1), "ODDFYIELD")?;
                let issue = self.to_f64_arg(evaluated_args.get(2), "ODDFYIELD")?;
                let first_coupon = self.to_f64_arg(evaluated_args.get(3), "ODDFYIELD")?;
                let rate = self.to_f64_arg(evaluated_args.get(4), "ODDFYIELD")?;
                let pr = self.to_f64_arg(evaluated_args.get(5), "ODDFYIELD")?;
                let redemption = self.to_f64_arg(evaluated_args.get(6), "ODDFYIELD")?;
                let frequency = self.to_f64_arg(evaluated_args.get(7), "ODDFYIELD")?;
                let basis = self.opt_f64(evaluated_args, 8, 0.0);
                if settlement <= issue {
                    return Ok(ResultData::Error("#NUM!".to_string()));
                }
                match finance::oddfyield(
                    settlement,
                    maturity,
                    issue,
                    first_coupon,
                    rate,
                    pr,
                    redemption,
                    frequency,
                    basis,
                ) {
                    Some(v) => Ok(ResultData::Float(v)),
                    None => Ok(ResultData::Error("#NUM!".to_string())),
                }
            }
            "ODDLPRICE" => {
                let settlement = self.to_f64_arg(evaluated_args.first(), "ODDLPRICE")?;
                let maturity = self.to_f64_arg(evaluated_args.get(1), "ODDLPRICE")?;
                let last_interest = self.to_f64_arg(evaluated_args.get(2), "ODDLPRICE")?;
                let rate = self.to_f64_arg(evaluated_args.get(3), "ODDLPRICE")?;
                let yld = self.to_f64_arg(evaluated_args.get(4), "ODDLPRICE")?;
                let redemption = self.to_f64_arg(evaluated_args.get(5), "ODDLPRICE")?;
                let frequency = self.to_f64_arg(evaluated_args.get(6), "ODDLPRICE")?;
                let basis = self.opt_f64(evaluated_args, 7, 0.0);
                if settlement <= last_interest {
                    return Ok(ResultData::Error("#NUM!".to_string()));
                }
                Ok(ResultData::Float(finance::oddlprice(
                    settlement,
                    maturity,
                    last_interest,
                    rate,
                    yld,
                    redemption,
                    frequency,
                    basis,
                )))
            }
            "ODDLYIELD" => {
                let settlement = self.to_f64_arg(evaluated_args.first(), "ODDLYIELD")?;
                let maturity = self.to_f64_arg(evaluated_args.get(1), "ODDLYIELD")?;
                let last_interest = self.to_f64_arg(evaluated_args.get(2), "ODDLYIELD")?;
                let rate = self.to_f64_arg(evaluated_args.get(3), "ODDLYIELD")?;
                let pr = self.to_f64_arg(evaluated_args.get(4), "ODDLYIELD")?;
                let redemption = self.to_f64_arg(evaluated_args.get(5), "ODDLYIELD")?;
                let frequency = self.to_f64_arg(evaluated_args.get(6), "ODDLYIELD")?;
                let basis = self.opt_f64(evaluated_args, 7, 0.0);
                if settlement <= last_interest {
                    return Ok(ResultData::Error("#NUM!".to_string()));
                }
                Ok(ResultData::Float(finance::oddlyield(
                    settlement,
                    maturity,
                    last_interest,
                    rate,
                    pr,
                    redemption,
                    frequency,
                    basis,
                )))
            }
            "EUROCONVERT" => {
                let number = self.to_f64_arg(evaluated_args.first(), "EUROCONVERT")?;
                let source = evaluated_args
                    .get(1)
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let target = evaluated_args
                    .get(2)
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let full_precision = evaluated_args
                    .get(3)
                    .map(|v| self.to_bool(v))
                    .unwrap_or(false);
                let triangulation_precision = evaluated_args.get(4).and_then(|v| self.to_f64(v));
                res_to_rd(finance::euroconvert(
                    number,
                    &source,
                    &target,
                    full_precision,
                    triangulation_precision,
                ))
            }
            _ => {
                *owned = false;
                Ok(ResultData::None)
            }
        }
    }
}
