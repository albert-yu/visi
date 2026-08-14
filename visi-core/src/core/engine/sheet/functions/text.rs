//! Text function dispatch.
//!
//! Split out of the parent module's `evaluate_function`, which tries each
//! family in turn.

use super::{FnCall, res_to_rd};
use crate::core::engine::cell::{Dependency, EngineError};
use crate::core::engine::result_data::ResultData;
use crate::core::engine::sheet::Sheet;

impl Sheet {
    /// Evaluates `call` if this family owns its name, else `None`.
    pub(super) fn eval_text_fn(
        &self,
        call: FnCall<'_>,
        deps: &mut Vec<Dependency>,
    ) -> Option<Result<ResultData, EngineError>> {
        // The body returns `Result` so its arms can keep using `?`; whether
        // the name belongs to this family is signalled alongside.
        let mut owned = true;
        let r = self.eval_text_dispatch(call, deps, &mut owned);
        owned.then_some(r)
    }

    fn eval_text_dispatch(
        &self,
        call: FnCall<'_>,
        _deps: &mut Vec<Dependency>,
        owned: &mut bool,
    ) -> Result<ResultData, EngineError> {
        let FnCall { evaluated_args, .. } = call;
        match call.upper_name {
            // --- TEXT FUNCTIONS ---
            "ARRAYTOTEXT" => {
                // Every element's own text (numbers via
                // format_excel_number, TRUE/FALSE, raw strings, ...)
                // via ResultData's Display -- not flatten_stat_numbers,
                // which silently drops non-numeric cells and so only
                // ever produced a text/bool-free (and often empty)
                // result for a mixed range.
                fn flatten_text(val: &ResultData, out: &mut Vec<String>) {
                    match val {
                        ResultData::List(items) => {
                            for item in items {
                                flatten_text(item, out);
                            }
                        }
                        other => out.push(other.to_string()),
                    }
                }
                let mut items = Vec::new();
                if let Some(arg) = evaluated_args.first() {
                    flatten_text(arg, &mut items);
                }
                // A *single* empty cell has no text to render at all
                // and is #VALUE!. A multi-cell range of blanks is not:
                // ARRAYTOTEXT over two empty cells is "," in real
                // Excel, i.e. the separators still show.
                if items.len() == 1 && items[0].is_empty() {
                    return Ok(ResultData::Error("#VALUE!".to_string()));
                }
                let fmt = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                match crate::core::text::arraytotext(&items, fmt) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "ASC" => {
                let text = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                match crate::core::text::asc(&text) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "JIS" => {
                let text = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                match crate::core::text::jis(&text) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "BAHTTEXT" => {
                let num = self.to_f64_arg(evaluated_args.first(), "BAHTTEXT")?;
                match crate::core::text::bahttext(num) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "CHAR" => {
                let num = self.to_f64_arg(evaluated_args.first(), "CHAR")?;
                match crate::core::text::char_fn(num) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "CLEAN" => {
                let text = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                match crate::core::text::clean(&text) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "CODE" => {
                let text = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                res_to_rd(crate::core::text::code(&text))
            }
            "DBCS" => {
                let text = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                match crate::core::text::dbcs(&text) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "DETECTLANGUAGE" => {
                let text = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                match crate::core::text::detectlanguage(&text) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "DOLLAR" => {
                let num = self.to_f64_arg(evaluated_args.first(), "DOLLAR")?;
                let dec = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                match crate::core::text::dollar(num, dec) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "EXACT" => {
                let t1 = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let t2 = evaluated_args
                    .get(1)
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                match crate::core::text::exact(&t1, &t2) {
                    Ok(b) => Ok(ResultData::Boolean(b)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "FIND" | "FINDB" => {
                let find_text = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let within_text = evaluated_args
                    .get(1)
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let start_num = evaluated_args.get(2).and_then(|v| self.to_f64(v));
                res_to_rd(crate::core::text::find(&find_text, &within_text, start_num))
            }
            "FIXED" => {
                let num = self.to_f64_arg(evaluated_args.first(), "FIXED")?;
                let dec = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                let no_commas = evaluated_args.get(2).map(|v| self.to_bool(v));
                match crate::core::text::fixed(num, dec, no_commas) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "NUMBERVALUE" => {
                let text = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let dec = evaluated_args.get(1).map(|v| v.to_string());
                let grp = evaluated_args.get(2).map(|v| v.to_string());
                res_to_rd(crate::core::text::numbervalue(
                    &text,
                    dec.as_deref(),
                    grp.as_deref(),
                ))
            }
            "PHONETIC" => {
                let text = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                match crate::core::text::phonetic(&text) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "REGEXEXTRACT" => {
                let text = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let pat = evaluated_args
                    .get(1)
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                match crate::core::text::regexextract(&text, &pat) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "REGEXREPLACE" => {
                let text = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let pat = evaluated_args
                    .get(1)
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let rep = evaluated_args
                    .get(2)
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                match crate::core::text::regexreplace(&text, &pat, &rep) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "REGEXTEST" => {
                let text = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let pat = evaluated_args
                    .get(1)
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                match crate::core::text::regextest(&text, &pat) {
                    Ok(b) => Ok(ResultData::Boolean(b)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "REPLACE" | "REPLACEB" => {
                let old_text = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let start_num = self.to_f64_arg(evaluated_args.get(1), "REPLACE")?;
                let num_chars = self.to_f64_arg(evaluated_args.get(2), "REPLACE")?;
                let new_text = evaluated_args
                    .get(3)
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                match crate::core::text::replace_fn(&old_text, start_num, num_chars, &new_text) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "REPT" => {
                let text = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let cnt = self.to_f64_arg(evaluated_args.get(1), "REPT")?;
                match crate::core::text::rept(&text, cnt) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "SEARCH" | "SEARCHB" => {
                let find_text = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let within_text = evaluated_args
                    .get(1)
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let start_num = evaluated_args.get(2).and_then(|v| self.to_f64(v));
                res_to_rd(crate::core::text::search(
                    &find_text,
                    &within_text,
                    start_num,
                ))
            }
            "SUBSTITUTE" => {
                let text = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let old_text = evaluated_args
                    .get(1)
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let new_text = evaluated_args
                    .get(2)
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let instance = evaluated_args.get(3).and_then(|v| self.to_f64(v));
                match crate::core::text::substitute(&text, &old_text, &new_text, instance) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "T" => {
                let is_str = matches!(evaluated_args.first(), Some(ResultData::String(_)));
                let val = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                Ok(ResultData::String(crate::core::text::t_fn(&val, is_str)))
            }
            "TEXT" => {
                let num = self.to_f64_arg(evaluated_args.first(), "TEXT")?;
                let fmt = evaluated_args
                    .get(1)
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                match crate::core::text::text_fn(num, &fmt) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "TEXTAFTER" => {
                let text = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let delim = evaluated_args
                    .get(1)
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let instance = evaluated_args.get(2).and_then(|v| self.to_f64(v));
                match crate::core::text::textafter(&text, &delim, instance) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "TEXTBEFORE" => {
                let text = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let delim = evaluated_args
                    .get(1)
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let instance = evaluated_args.get(2).and_then(|v| self.to_f64(v));
                match crate::core::text::textbefore(&text, &delim, instance) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "TEXTJOIN" => {
                let delim = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let ignore = evaluated_args
                    .get(1)
                    .map(|v| self.to_bool(v))
                    .unwrap_or(true);
                let texts: Vec<String> = evaluated_args
                    .iter()
                    .skip(2)
                    .map(|v| v.to_string())
                    .collect();
                match crate::core::text::textjoin(&delim, ignore, &texts) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "TEXTSPLIT" => {
                let text = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let delim = evaluated_args
                    .get(1)
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                match crate::core::text::textsplit(&text, &delim) {
                    Ok(parts) => Ok(ResultData::List(
                        parts.into_iter().map(ResultData::String).collect(),
                    )),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "TRANSLATE" => {
                let text = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let from = evaluated_args
                    .get(1)
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let to = evaluated_args
                    .get(2)
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                match crate::core::text::translate(&text, &from, &to) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "UNICHAR" => {
                let num = self.to_f64_arg(evaluated_args.first(), "UNICHAR")?;
                match crate::core::text::unichar(num) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "UNICODE" => {
                let text = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                res_to_rd(crate::core::text::unicode(&text))
            }
            "VALUE" => {
                let text = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                res_to_rd(crate::core::text::value(&text))
            }
            "VALUETOTEXT" => {
                let val = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let fmt = evaluated_args.get(1).and_then(|v| self.to_f64(v));
                match crate::core::text::valuetotext(&val, fmt) {
                    Ok(s) => Ok(ResultData::String(s)),
                    Err(e) => Ok(ResultData::Error(e)),
                }
            }
            "LEFTB" => {
                let text = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let count = evaluated_args
                    .get(1)
                    .and_then(|v| self.to_f64(v))
                    .unwrap_or(1.0)
                    .floor() as usize;
                let res: String = text.chars().take(count).collect();
                Ok(ResultData::String(res))
            }
            "RIGHTB" => {
                let text = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let count = evaluated_args
                    .get(1)
                    .and_then(|v| self.to_f64(v))
                    .unwrap_or(1.0)
                    .floor() as usize;
                let chars: Vec<char> = text.chars().collect();
                let skip = chars.len().saturating_sub(count);
                let res: String = chars.into_iter().skip(skip).collect();
                Ok(ResultData::String(res))
            }
            "LENB" => {
                let text = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                Ok(ResultData::Float(text.len() as f64))
            }
            "MIDB" => {
                let text = evaluated_args
                    .first()
                    .map(|v| v.to_string())
                    .unwrap_or_default();
                let start = self.to_f64_arg(evaluated_args.get(1), "MIDB")?.floor() as usize;
                let count = self.to_f64_arg(evaluated_args.get(2), "MIDB")?.floor() as usize;
                if start < 1 {
                    Ok(ResultData::Error("#VALUE!".to_string()))
                } else {
                    let chars: Vec<char> = text.chars().collect();
                    let start_idx = (start - 1).min(chars.len());
                    let res: String = chars.into_iter().skip(start_idx).take(count).collect();
                    Ok(ResultData::String(res))
                }
            }
            _ => {
                *owned = false;
                Ok(ResultData::None)
            }
        }
    }
}
