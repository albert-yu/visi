//! Engineering function dispatch.
//!
//! Split out of the parent module's `evaluate_function`, which tries each
//! family in turn.

use super::{FnCall, res_to_rd};
use crate::core::engine::cell::{Dependency, EngineError};
use crate::core::engine::result_data::ResultData;
use crate::core::engine::sheet::Sheet;

impl Sheet {
    /// Evaluates `call` if this family owns its name, else `None`.
    pub(super) fn eval_engineering_fn(
        &self,
        call: FnCall<'_>,
        deps: &mut Vec<Dependency>,
    ) -> Option<Result<ResultData, EngineError>> {
        // The body returns `Result` so its arms can keep using `?`; whether
        // the name belongs to this family is signalled alongside.
        let mut owned = true;
        let r = self.eval_engineering_dispatch(call, deps, &mut owned);
        owned.then_some(r)
    }

    fn eval_engineering_dispatch(
        &self,
        call: FnCall<'_>,
        _deps: &mut Vec<Dependency>,
        owned: &mut bool,
    ) -> Result<ResultData, EngineError> {
        let FnCall {
            upper_name,
            evaluated_args,
            ..
        } = call;
        match call.upper_name {
            // --- ENGINEERING FUNCTIONS ---
            "BESSELI" => {
                let x = self.to_f64_arg(evaluated_args.first(), "BESSELI")?;
                let n = self.to_f64_arg(evaluated_args.get(1), "BESSELI")?;
                res_to_rd(crate::core::engineering::besseli(x, n))
            }
            "BESSELJ" => {
                let x = self.to_f64_arg(evaluated_args.first(), "BESSELJ")?;
                let n = self.to_f64_arg(evaluated_args.get(1), "BESSELJ")?;
                res_to_rd(crate::core::engineering::besselj(x, n))
            }
            "BESSELK" => {
                let x = self.to_f64_arg(evaluated_args.first(), "BESSELK")?;
                let n = self.to_f64_arg(evaluated_args.get(1), "BESSELK")?;
                res_to_rd(crate::core::engineering::besselk(x, n))
            }
            "BESSELY" => {
                let x = self.to_f64_arg(evaluated_args.first(), "BESSELY")?;
                let n = self.to_f64_arg(evaluated_args.get(1), "BESSELY")?;
                res_to_rd(crate::core::engineering::bessely(x, n))
            }
            "BIN2DEC" => {
                if Self::first_arg_is_boolean(evaluated_args) {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                }
                let t = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                res_to_rd(crate::core::engineering::bin2dec(&t))
            }
            "BIN2HEX" => {
                let t = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let p = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                match crate::core::engineering::bin2hex(&t, p) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "BIN2OCT" => {
                let t = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let p = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                match crate::core::engineering::bin2oct(&t, p) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "BITAND" => {
                let n1 = self.to_f64_arg(evaluated_args.first(), "BITAND")?;
                let n2 = self.to_f64_arg(evaluated_args.get(1), "BITAND")?;
                res_to_rd(crate::core::engineering::bitand(n1, n2))
            }
            "BITLSHIFT" => {
                let n = self.to_f64_arg(evaluated_args.first(), "BITLSHIFT")?;
                let s = self.to_f64_arg(evaluated_args.get(1), "BITLSHIFT")?;
                res_to_rd(crate::core::engineering::bitlshift(n, s))
            }
            "BITOR" => {
                let n1 = self.to_f64_arg(evaluated_args.first(), "BITOR")?;
                let n2 = self.to_f64_arg(evaluated_args.get(1), "BITOR")?;
                res_to_rd(crate::core::engineering::bitor(n1, n2))
            }
            "BITRSHIFT" => {
                let n = self.to_f64_arg(evaluated_args.first(), "BITRSHIFT")?;
                let s = self.to_f64_arg(evaluated_args.get(1), "BITRSHIFT")?;
                res_to_rd(crate::core::engineering::bitrshift(n, s))
            }
            "BITXOR" => {
                let n1 = self.to_f64_arg(evaluated_args.first(), "BITXOR")?;
                let n2 = self.to_f64_arg(evaluated_args.get(1), "BITXOR")?;
                res_to_rd(crate::core::engineering::bitxor(n1, n2))
            }
            "COMPLEX" => {
                let r = self.to_f64_arg(evaluated_args.first(), "COMPLEX")?;
                let i = self.to_f64_arg(evaluated_args.get(1), "COMPLEX")?;
                let s = evaluated_args.get(2).map(|v| v.to_string());
                match crate::core::engineering::complex_fn(r, i, s.as_deref()) {
                    Ok(res) => Ok(ResultData::String(res)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "CONVERT" => {
                let val = self.to_f64_arg(evaluated_args.first(), "CONVERT")?;
                let u1 = evaluated_args
                    .get(1)
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let u2 = evaluated_args
                    .get(2)
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                res_to_rd(crate::core::engineering::convert(val, &u1, &u2))
            }
            "DEC2BIN" => {
                let n = self.to_f64_arg(evaluated_args.first(), "DEC2BIN")?;
                let p = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                match crate::core::engineering::dec2bin(n, p) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "DEC2HEX" => {
                let n = self.to_f64_arg(evaluated_args.first(), "DEC2HEX")?;
                let p = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                match crate::core::engineering::dec2hex(n, p) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "DEC2OCT" => {
                let n = self.to_f64_arg(evaluated_args.first(), "DEC2OCT")?;
                let p = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                match crate::core::engineering::dec2oct(n, p) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "DELTA" => {
                let n1 = self.to_f64_arg(evaluated_args.first(), "DELTA")?;
                let n2 = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                res_to_rd(crate::core::engineering::delta(n1, n2))
            }
            "ERF" | "ERFC" | "ERF.PRECISE" | "ERFC.PRECISE" => {
                // Unlike SQRT/ABS/INT/MOD (which all accept a boolean
                // as 1/0), the error functions reject booleans: real
                // Excel answers #VALUE! for ERF(TRUE) and ERF(FALSE).
                // Numeric *text* is coerced though, from a literal or
                // from a text cell, and surrounding whitespace is
                // tolerated -- ERF("1") and ERF(" 1 ") both give
                // 0.8427007929497149. Non-numeric text is #VALUE!.
                // A blank argument coerces to 0 (ERF(<blank>) is 0 and
                // ERFC(<blank>) is 1). Same rule as QUOTIENT.
                let x = match evaluated_args.first() {
                    None | Some(ResultData::None) => 0.0,
                    Some(v) => {
                        // A one-cell range arrives as a one-element List.
                        let scalar = match v {
                            ResultData::List(items) if items.len() == 1 => &items[0],
                            other => other,
                        };
                        if matches!(scalar, ResultData::Boolean(_)) {
                            // See first_arg_is_boolean.
                            return Ok(ResultData::Error("#VALUE!".to_string()));
                        }
                        match self.to_f64(scalar) {
                            Some(f) => f,
                            None => return Ok(ResultData::Error("#VALUE!".to_string())),
                        }
                    }
                };
                let v = if upper_name.starts_with("ERFC") {
                    crate::core::stats::erfc(x)
                } else {
                    crate::core::stats::erf(x)
                };
                res_to_rd(Ok(v))
            }
            "GESTEP" => {
                let n = self.to_f64_arg(evaluated_args.first(), "GESTEP")?;
                let step = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                res_to_rd(crate::core::engineering::gestep(n, step))
            }
            "HEX2BIN" => {
                let t = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let p = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                match crate::core::engineering::hex2bin(&t, p) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "HEX2DEC" => {
                let t = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                res_to_rd(crate::core::engineering::hex2dec(&t))
            }
            "HEX2OCT" => {
                let t = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let p = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                match crate::core::engineering::hex2oct(&t, p) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "IMABS" => {
                let t = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                res_to_rd(crate::core::engineering::imabs(&t))
            }
            "IMAGINARY" => {
                let t = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                res_to_rd(crate::core::engineering::imaginary(&t))
            }
            "IMARGUMENT" => {
                let t = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                res_to_rd(crate::core::engineering::imargument(&t))
            }
            "IMCONJUGATE" => {
                let t = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                match crate::core::engineering::imconjugate(&t) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "IMDIV" => {
                let t1 = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let t2 = evaluated_args
                    .get(1)
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                match crate::core::engineering::imdiv(&t1, &t2) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "IMPRODUCT" => {
                let strs: Vec<String> = evaluated_args.iter().map(|v| v.to_string()).collect();
                let refs: Vec<&str> = strs.iter().map(|s| s.as_str()).collect();
                match crate::core::engineering::improduct(&refs) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "IMREAL" => {
                let t = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                res_to_rd(crate::core::engineering::imreal(&t))
            }
            "IMSUB" => {
                let t1 = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let t2 = evaluated_args
                    .get(1)
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                match crate::core::engineering::imsub(&t1, &t2) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "IMSUM" => {
                let strs: Vec<String> = evaluated_args.iter().map(|v| v.to_string()).collect();
                let refs: Vec<&str> = strs.iter().map(|s| s.as_str()).collect();
                match crate::core::engineering::imsum(&refs) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "OCT2BIN" => {
                let t = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let p = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                match crate::core::engineering::oct2bin(&t, p) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "OCT2DEC" => {
                let t = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                res_to_rd(crate::core::engineering::oct2dec(&t))
            }
            "OCT2HEX" => {
                let t = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let p = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                match crate::core::engineering::oct2hex(&t, p) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "IMCOS" | "IMCOSH" | "IMCOT" | "IMCSC" | "IMCSCH" | "IMEXP" | "IMLN" | "IMLOG10"
            | "IMLOG2" | "IMSEC" | "IMSECH" | "IMSIN" | "IMSINH" | "IMSQRT" | "IMTAN" => {
                let t = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let result = match upper_name {
                    "IMCOS" => crate::core::engineering::imcos(&t),
                    "IMCOSH" => crate::core::engineering::imcosh(&t),
                    "IMCOT" => crate::core::engineering::imcot(&t),
                    "IMCSC" => crate::core::engineering::imcsc(&t),
                    "IMCSCH" => crate::core::engineering::imcsch(&t),
                    "IMEXP" => crate::core::engineering::imexp(&t),
                    "IMLN" => crate::core::engineering::imln(&t),
                    "IMLOG10" => crate::core::engineering::imlog10(&t),
                    "IMLOG2" => crate::core::engineering::imlog2(&t),
                    "IMSEC" => crate::core::engineering::imsec(&t),
                    "IMSECH" => crate::core::engineering::imsech(&t),
                    "IMSIN" => crate::core::engineering::imsin(&t),
                    "IMSINH" => crate::core::engineering::imsinh(&t),
                    "IMSQRT" => crate::core::engineering::imsqrt(&t),
                    "IMTAN" => crate::core::engineering::imtan(&t),
                    _ => unreachable!(),
                };
                match result {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "IMPOWER" => {
                let t = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let n = self.to_f64_arg(evaluated_args.get(1), "IMPOWER")?;
                match crate::core::engineering::impower(&t, n) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            _ => {
                *owned = false;
                Ok(ResultData::None)
            }
        }
    }
}
