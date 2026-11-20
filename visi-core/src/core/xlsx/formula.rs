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
    prefix_future_functions(&with_spills)
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
}
