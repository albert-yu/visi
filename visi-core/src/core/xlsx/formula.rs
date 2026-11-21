pub(crate) fn ooxml_function_prefix(upper_name: &str) -> Option<&'static str> {
    match upper_name {
        "FILTER" | "SORT" => Some("_xlfn._xlws."),
        "ACOT"
        | "ACOTH"
        | "AGGREGATE"
        | "ANCHORARRAY"
        | "ARABIC"
        | "ARRAYTOTEXT"
        | "BASE"
        | "BETA.DIST"
        | "BETA.INV"
        | "BINOM.DIST"
        | "BINOM.DIST.RANGE"
        | "BINOM.INV"
        | "BITAND"
        | "BITLSHIFT"
        | "BITOR"
        | "BITRSHIFT"
        | "BITXOR"
        | "BYCOL"
        | "BYROW"
        | "CEILING.MATH"
        | "CEILING.PRECISE"
        | "CHISQ.DIST"
        | "CHISQ.DIST.RT"
        | "CHISQ.INV"
        | "CHISQ.INV.RT"
        | "CHISQ.TEST"
        | "CHOOSECOLS"
        | "CHOOSEROWS"
        | "COMBINA"
        | "CONCAT"
        | "CONFIDENCE.NORM"
        | "CONFIDENCE.T"
        | "COT"
        | "COTH"
        | "COVARIANCE.P"
        | "COVARIANCE.S"
        | "CSC"
        | "CSCH"
        | "DAYS"
        | "DECIMAL"
        | "DROP"
        | "ENCODEURL"
        | "ERF.PRECISE"
        | "ERFC.PRECISE"
        | "EXPAND"
        | "EXPON.DIST"
        | "F.DIST"
        | "F.DIST.RT"
        | "F.INV"
        | "F.INV.RT"
        | "F.TEST"
        | "FILTERXML"
        | "FLOOR.MATH"
        | "FLOOR.PRECISE"
        | "FORECAST.ETS"
        | "FORECAST.ETS.CONFINT"
        | "FORECAST.ETS.SEASONALITY"
        | "FORECAST.ETS.STAT"
        | "FORECAST.LINEAR"
        | "FORMULATEXT"
        | "GAMMA"
        | "GAMMA.DIST"
        | "GAMMA.INV"
        | "GAUSS"
        | "HSTACK"
        | "HYPGEOM.DIST"
        | "IFNA"
        | "IFS"
        | "IMCOSH"
        | "IMCOT"
        | "IMCSC"
        | "IMCSCH"
        | "IMSEC"
        | "IMSECH"
        | "IMSINH"
        | "IMTAN"
        | "ISFORMULA"
        | "ISO.CEILING"
        | "ISOMITTED"
        | "ISOWEEKNUM"
        | "LAMBDA"
        | "LET"
        | "LOGNORM.DIST"
        | "LOGNORM.INV"
        | "MAKEARRAY"
        | "MAP"
        | "MAXIFS"
        | "MINIFS"
        | "MODE.MULT"
        | "MODE.SNGL"
        | "MUNIT"
        | "NEGBINOM.DIST"
        | "NETWORKDAYS.INTL"
        | "NORM.DIST"
        | "NORM.INV"
        | "NORM.S.DIST"
        | "NORM.S.INV"
        | "NUMBERVALUE"
        | "PDURATION"
        | "PERCENTILE.EXC"
        | "PERCENTILE.INC"
        | "PERCENTRANK.EXC"
        | "PERCENTRANK.INC"
        | "PERMUTATIONA"
        | "PHI"
        | "POISSON.DIST"
        | "QUARTILE.EXC"
        | "QUARTILE.INC"
        | "RANDARRAY"
        | "RANK.AVG"
        | "RANK.EQ"
        | "REDUCE"
        | "RRI"
        | "SCAN"
        | "SEC"
        | "SECH"
        | "SEQUENCE"
        | "SHEET"
        | "SHEETS"
        | "SINGLE"
        | "SKEW.P"
        | "SORTBY"
        | "STDEV.P"
        | "STDEV.S"
        | "SWITCH"
        | "T.DIST"
        | "T.DIST.2T"
        | "T.DIST.RT"
        | "T.INV"
        | "T.INV.2T"
        | "T.TEST"
        | "TAKE"
        | "TEXTAFTER"
        | "TEXTBEFORE"
        | "TEXTJOIN"
        | "TEXTSPLIT"
        | "TOCOL"
        | "TOROW"
        | "UNICHAR"
        | "UNICODE"
        | "UNIQUE"
        | "VALUETOTEXT"
        | "VAR.P"
        | "VAR.S"
        | "VSTACK"
        | "WEBSERVICE"
        | "WEIBULL.DIST"
        | "WORKDAY.INTL"
        | "WRAPCOLS"
        | "WRAPROWS"
        | "XLOOKUP"
        | "XMATCH"
        | "XOR"
        | "Z.TEST" => Some("_xlfn."),
        _ => None,
    }
}

pub(crate) fn encode_xlsx_formula(formula_src: &str) -> String {
    let with_spills = crate::core::parser::encode_xlsx_spill_references(formula_src);
    let with_xlfn = prefix_future_functions(&with_spills);
    prefix_xlpm_parameters(&with_xlfn)
}

fn strip_xlpm_prefix(ident: &str) -> &str {
    if ident.len() >= 6 && ident.as_bytes()[..6].eq_ignore_ascii_case(b"_xlpm.") {
        &ident[6..]
    } else {
        ident
    }
}

fn is_xlpm_prefixed(ident: &str) -> bool {
    ident.len() >= 6 && ident.as_bytes()[..6].eq_ignore_ascii_case(b"_xlpm.")
}

fn is_ident_start(b: u8) -> bool {
    b.is_ascii_alphabetic() || b == b'_'
}

fn is_ident_cont(b: u8) -> bool {
    b.is_ascii_alphanumeric() || b == b'_' || b == b'.'
}

fn find_matching_paren(s: &str, open_idx: usize) -> Option<usize> {
    let bytes = s.as_bytes();
    let len = bytes.len();
    let mut depth = 0usize;
    let mut i = open_idx;
    while i < len {
        match bytes[i] {
            b'"' => {
                i += 1;
                while i < len {
                    if bytes[i] == b'"' {
                        i += 1;
                        if i < len && bytes[i] == b'"' {
                            i += 1;
                        } else {
                            break;
                        }
                    } else {
                        i += 1;
                    }
                }
            }
            b'\'' => {
                i += 1;
                while i < len {
                    if bytes[i] == b'\'' {
                        i += 1;
                        if i < len && bytes[i] == b'\'' {
                            i += 1;
                        } else {
                            break;
                        }
                    } else {
                        i += 1;
                    }
                }
            }
            b'(' => {
                depth += 1;
                i += 1;
            }
            b')' => {
                depth -= 1;
                if depth == 0 {
                    return Some(i);
                }
                i += 1;
            }
            _ => i += 1,
        }
    }
    None
}

fn split_top_level_args(s: &str) -> Vec<&str> {
    let bytes = s.as_bytes();
    let len = bytes.len();
    let mut args = Vec::new();
    let mut start = 0usize;
    let mut paren_depth = 0usize;
    let mut bracket_depth = 0usize;
    let mut brace_depth = 0usize;
    let mut i = 0usize;
    while i < len {
        match bytes[i] {
            b'"' => {
                i += 1;
                while i < len {
                    if bytes[i] == b'"' {
                        i += 1;
                        if i < len && bytes[i] == b'"' {
                            i += 1;
                        } else {
                            break;
                        }
                    } else {
                        i += 1;
                    }
                }
            }
            b'\'' => {
                i += 1;
                while i < len {
                    if bytes[i] == b'\'' {
                        i += 1;
                        if i < len && bytes[i] == b'\'' {
                            i += 1;
                        } else {
                            break;
                        }
                    } else {
                        i += 1;
                    }
                }
            }
            b'(' => {
                paren_depth += 1;
                i += 1;
            }
            b')' => {
                paren_depth = paren_depth.saturating_sub(1);
                i += 1;
            }
            b'[' => {
                bracket_depth += 1;
                i += 1;
            }
            b']' => {
                bracket_depth = bracket_depth.saturating_sub(1);
                i += 1;
            }
            b'{' => {
                brace_depth += 1;
                i += 1;
            }
            b'}' => {
                brace_depth = brace_depth.saturating_sub(1);
                i += 1;
            }
            b',' if paren_depth == 0 && bracket_depth == 0 && brace_depth == 0 => {
                args.push(&s[start..i]);
                i += 1;
                start = i;
            }
            _ => i += 1,
        }
    }
    args.push(&s[start..len]);
    args
}

fn extract_single_param_ident(arg: &str) -> Option<(&str, &str, &str)> {
    let trimmed_start = arg.trim_start();
    let lead_len = arg.len() - trimmed_start.len();
    let trimmed = trimmed_start.trim_end();
    let trail_start = lead_len + trimmed.len();
    if trimmed.is_empty() {
        return None;
    }
    let bytes = trimmed.as_bytes();
    if !is_ident_start(bytes[0]) || !bytes.iter().all(|&b| is_ident_cont(b)) {
        return None;
    }
    Some((&arg[..lead_len], trimmed, &arg[trail_start..]))
}

fn prefix_xlpm_parameters(formula: &str) -> String {
    let mut out = String::with_capacity(formula.len() + 16);
    rewrite_xlpm_scope(formula, &[], &mut out);
    out
}

fn rewrite_xlpm_scope(expr: &str, bound: &[String], out: &mut String) {
    let bytes = expr.as_bytes();
    let len = bytes.len();
    let mut i = 0usize;
    let mut bracket_depth = 0usize;

    while i < len {
        match bytes[i] {
            b'"' => {
                let start = i;
                i += 1;
                while i < len {
                    if bytes[i] == b'"' {
                        i += 1;
                        if i < len && bytes[i] == b'"' {
                            i += 1;
                        } else {
                            break;
                        }
                    } else {
                        i += 1;
                    }
                }
                out.push_str(&expr[start..i]);
            }
            b'\'' => {
                let start = i;
                i += 1;
                while i < len {
                    if bytes[i] == b'\'' {
                        i += 1;
                        if i < len && bytes[i] == b'\'' {
                            i += 1;
                        } else {
                            break;
                        }
                    } else {
                        i += 1;
                    }
                }
                out.push_str(&expr[start..i]);
            }
            b'[' => {
                bracket_depth += 1;
                out.push('[');
                i += 1;
            }
            b']' => {
                bracket_depth = bracket_depth.saturating_sub(1);
                out.push(']');
                i += 1;
            }
            b if bracket_depth == 0 && is_ident_start(b) => {
                let start = i;
                while i < len && is_ident_cont(bytes[i]) {
                    i += 1;
                }
                let ident = &expr[start..i];
                let prev_is_ident_or_bang =
                    start > 0 && (is_ident_cont(bytes[start - 1]) || bytes[start - 1] == b'!');
                let next_is_paren = i < len && bytes[i] == b'(';
                let next_is_bang = i < len && bytes[i] == b'!';

                if !prev_is_ident_or_bang && next_is_paren {
                    let bare_fn = if ident.len() >= 6
                        && ident.as_bytes()[..6].eq_ignore_ascii_case(b"_xlfn.")
                    {
                        &ident[6..]
                    } else {
                        ident
                    };
                    if let Some(close_idx) = find_matching_paren(expr, i) {
                        let inner = &expr[i + 1..close_idx];
                        if bare_fn.eq_ignore_ascii_case("LET") {
                            let args = split_top_level_args(inner);
                            if args.len() >= 3 && args.len() % 2 == 1 {
                                out.push_str(ident);
                                out.push('(');
                                let mut local_bound = bound.to_vec();
                                for pair_idx in (0..args.len() - 1).step_by(2) {
                                    if pair_idx > 0 {
                                        out.push(',');
                                    }
                                    let name_arg = args[pair_idx];
                                    let val_arg = args[pair_idx + 1];
                                    let mut newly_bound = None;
                                    if let Some((lead, param, trail)) =
                                        extract_single_param_ident(name_arg)
                                    {
                                        out.push_str(lead);
                                        if !is_xlpm_prefixed(param) {
                                            out.push_str("_xlpm.");
                                        }
                                        out.push_str(param);
                                        out.push_str(trail);
                                        newly_bound =
                                            Some(strip_xlpm_prefix(param).to_ascii_uppercase());
                                    } else {
                                        rewrite_xlpm_scope(name_arg, &local_bound, out);
                                    }
                                    out.push(',');
                                    rewrite_xlpm_scope(val_arg, &local_bound, out);
                                    if let Some(bname) = newly_bound {
                                        local_bound.push(bname);
                                    }
                                }
                                out.push(',');
                                rewrite_xlpm_scope(args[args.len() - 1], &local_bound, out);
                                out.push(')');
                                i = close_idx + 1;
                                continue;
                            }
                        } else if bare_fn.eq_ignore_ascii_case("LAMBDA") {
                            let args = split_top_level_args(inner);
                            if args.len() >= 2 {
                                out.push_str(ident);
                                out.push('(');
                                let mut local_bound = bound.to_vec();
                                for (idx, param_arg) in args[..args.len() - 1].iter().enumerate() {
                                    if idx > 0 {
                                        out.push(',');
                                    }
                                    if let Some((lead, param, trail)) =
                                        extract_single_param_ident(param_arg)
                                    {
                                        out.push_str(lead);
                                        if !is_xlpm_prefixed(param) {
                                            out.push_str("_xlpm.");
                                        }
                                        out.push_str(param);
                                        out.push_str(trail);
                                        local_bound
                                            .push(strip_xlpm_prefix(param).to_ascii_uppercase());
                                    } else {
                                        rewrite_xlpm_scope(param_arg, &local_bound, out);
                                    }
                                }
                                out.push(',');
                                rewrite_xlpm_scope(args[args.len() - 1], &local_bound, out);
                                out.push(')');
                                i = close_idx + 1;
                                continue;
                            }
                        }
                    }

                    let upper_bare = strip_xlpm_prefix(ident).to_ascii_uppercase();
                    if !is_xlpm_prefixed(ident)
                        && bound.iter().any(|b| b == &upper_bare)
                        && !ident.contains('.')
                    {
                        out.push_str("_xlpm.");
                    }
                    out.push_str(ident);
                } else if !prev_is_ident_or_bang && !next_is_bang {
                    let upper_bare = strip_xlpm_prefix(ident).to_ascii_uppercase();
                    if !is_xlpm_prefixed(ident) && bound.iter().any(|b| b == &upper_bare) {
                        out.push_str("_xlpm.");
                    }
                    out.push_str(ident);
                } else {
                    out.push_str(ident);
                }
            }
            _ => {
                let ch = expr[i..].chars().next().unwrap();
                out.push(ch);
                i += ch.len_utf8();
            }
        }
    }
}

fn prefix_future_functions(formula: &str) -> String {
    let bytes = formula.as_bytes();
    let len = bytes.len();
    let mut out = String::with_capacity(len + 16);
    let mut i = 0usize;

    while i < len {
        match bytes[i] {
            b'"' => {
                out.push('"');
                i += 1;
                while i < len {
                    let ch = formula[i..].chars().next().unwrap();
                    out.push(ch);
                    i += ch.len_utf8();
                    if ch == '"' {
                        if i < len && bytes[i] == b'"' {
                            out.push('"');
                            i += 1;
                        } else {
                            break;
                        }
                    }
                }
            }
            b'\'' => {
                out.push('\'');
                i += 1;
                while i < len {
                    let ch = formula[i..].chars().next().unwrap();
                    out.push(ch);
                    i += ch.len_utf8();
                    if ch == '\'' {
                        if i < len && bytes[i] == b'\'' {
                            out.push('\'');
                            i += 1;
                        } else {
                            break;
                        }
                    }
                }
            }
            b if b.is_ascii_alphabetic() || b == b'_' => {
                let start = i;
                while i < len
                    && (bytes[i].is_ascii_alphanumeric() || bytes[i] == b'_' || bytes[i] == b'.')
                {
                    i += 1;
                }
                let ident = &formula[start..i];
                let prev_is_ident = start > 0
                    && (bytes[start - 1].is_ascii_alphanumeric()
                        || bytes[start - 1] == b'_'
                        || bytes[start - 1] == b'.');
                if !prev_is_ident && i < len && bytes[i] == b'(' {
                    let upper = ident.to_ascii_uppercase();
                    if !upper.starts_with("_XLFN.")
                        && !upper.starts_with("_XLWS.")
                        && let Some(prefix) = ooxml_function_prefix(&upper)
                    {
                        out.push_str(prefix);
                    }
                }
                out.push_str(ident);
            }
            _ => {
                let ch = formula[i..].chars().next().unwrap();
                out.push(ch);
                i += ch.len_utf8();
            }
        }
    }

    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn prefixes_modern_and_xlws_functions_without_touching_legacy_or_strings() {
        assert_eq!(
            encode_xlsx_formula("XLOOKUP(A1,B1:B5,C1:C5)"),
            "_xlfn.XLOOKUP(A1,B1:B5,C1:C5)"
        );
        assert_eq!(
            encode_xlsx_formula("SORT(FILTER(A1:A5,A1:A5>0))"),
            "_xlfn._xlws.SORT(_xlfn._xlws.FILTER(A1:A5,A1:A5>0))"
        );
        assert_eq!(
            encode_xlsx_formula("SUM(A1:A5)+IFERROR(1,0)"),
            "SUM(A1:A5)+IFERROR(1,0)"
        );
        assert_eq!(
            encode_xlsx_formula("CONCAT(\"XLOOKUP(\",_xlfn.XLOOKUP(A1,B1,C1))"),
            "_xlfn.CONCAT(\"XLOOKUP(\",_xlfn.XLOOKUP(A1,B1,C1))"
        );
    }

    #[test]
    fn prefixes_xlpm_parameters_in_let_and_lambda_scopes() {
        assert_eq!(
            encode_xlsx_formula("LET(x, A1*2, x + TaxRate)"),
            "_xlfn.LET(_xlpm.x, A1*2, _xlpm.x + TaxRate)"
        );
        assert_eq!(
            encode_xlsx_formula("LET(x, 10, y, x + 5, x * y)"),
            "_xlfn.LET(_xlpm.x, 10, _xlpm.y, _xlpm.x + 5, _xlpm.x * _xlpm.y)"
        );
        assert_eq!(
            encode_xlsx_formula("MAP(A1:A3, LAMBDA(v, v * 2))"),
            "_xlfn.MAP(A1:A3, _xlfn.LAMBDA(_xlpm.v, _xlpm.v * 2))"
        );
        assert_eq!(
            encode_xlsx_formula("LET(x, \"x\", y, Sales[x], x & y)"),
            "_xlfn.LET(_xlpm.x, \"x\", _xlpm.y, Sales[x], _xlpm.x & _xlpm.y)"
        );
        assert_eq!(
            encode_xlsx_formula("_xlfn.LET(_xlpm.x, 1, _xlpm.x + 2)"),
            "_xlfn.LET(_xlpm.x, 1, _xlpm.x + 2)"
        );
        assert_eq!(
            encode_xlsx_formula("LET(f, LAMBDA(n, n + 1), f(5))"),
            "_xlfn.LET(_xlpm.f, _xlfn.LAMBDA(_xlpm.n, _xlpm.n + 1), _xlpm.f(5))"
        );
    }
}
