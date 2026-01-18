const XL_FN_PREFIX: &str = "_xlfn.";
const XL_FN_PREFIX_BYTES: &[u8] = b"_xlfn.";
const XL_WS_PREFIX: &str = "_xlws.";
const XL_UDF_PREFIX: &str = "_xludf.";

// Functions that Excel stores in OOXML formulas with an `_xlfn.` prefix for forward
// compatibility (typically Excel 365 "future functions").
//
// Keep this list sorted (ASCII) for maintainability.
const XL_FN_REQUIRED_FUNCTIONS: &[&str] = &[
    "ACOT",
    "ACOTH",
    "AGGREGATE",
    "ARABIC",
    "BASE",
    "BETA.DIST",
    "BETA.INV",
    "BINOM.DIST",
    "BINOM.DIST.RANGE",
    "BINOM.INV",
    "BITAND",
    "BITLSHIFT",
    "BITOR",
    "BITRSHIFT",
    "BITXOR",
    "BYCOL",
    "BYROW",
    "CEILING.MATH",
    "CEILING.PRECISE",
    "CHISQ.DIST",
    "CHISQ.DIST.RT",
    "CHISQ.INV",
    "CHISQ.INV.RT",
    "CHISQ.TEST",
    "CHOOSECOLS",
    "CHOOSEROWS",
    "COMBINA",
    "CONCAT",
    "CONFIDENCE.NORM",
    "CONFIDENCE.T",
    "COT",
    "COTH",
    "COVARIANCE.P",
    "COVARIANCE.S",
    "CSC",
    "CSCH",
    "DAYS",
    "DECIMAL",
    "DROP",
    "EXPAND",
    "EXPON.DIST",
    "F.DIST",
    "F.DIST.RT",
    "F.INV",
    "F.INV.RT",
    "F.TEST",
    "FILTER",
    "FLOOR.MATH",
    "FLOOR.PRECISE",
    "FORECAST.ETS",
    "FORECAST.ETS.CONFINT",
    "FORECAST.ETS.SEASONALITY",
    "FORECAST.ETS.STAT",
    "FORECAST.LINEAR",
    "FORMULATEXT",
    "GAMMA",
    "GAMMA.DIST",
    "GAMMA.INV",
    "GAMMALN.PRECISE",
    "GAUSS",
    "HSTACK",
    "HYPGEOM.DIST",
    "IFNA",
    "IFS",
    "IMAGE",
    "ISFORMULA",
    "ISO.CEILING",
    "ISO.WEEKNUM",
    "ISOMITTED",
    "ISOWEEKNUM",
    "LAMBDA",
    "LET",
    "LOGNORM.DIST",
    "LOGNORM.INV",
    "MAKEARRAY",
    "MAP",
    "MAXIFS",
    "MINIFS",
    "MODE.MULT",
    "MODE.SNGL",
    "MUNIT",
    "NEGBINOM.DIST",
    "NETWORKDAYS.INTL",
    "NORM.DIST",
    "NORM.INV",
    "NORM.S.DIST",
    "NORM.S.INV",
    "NUMBERVALUE",
    "PDURATION",
    "PERCENTILE.EXC",
    "PERCENTILE.INC",
    "PERCENTRANK.EXC",
    "PERCENTRANK.INC",
    "PERMUTATIONA",
    "PHI",
    "POISSON.DIST",
    "QUARTILE.EXC",
    "QUARTILE.INC",
    "RANDARRAY",
    "RANK.AVG",
    "RANK.EQ",
    "REDUCE",
    "RRI",
    "SCAN",
    "SEC",
    "SECH",
    "SEQUENCE",
    "SHEET",
    "SHEETS",
    "SKEW.P",
    "SORT",
    "SORTBY",
    "STDEV.P",
    "STDEV.S",
    "SWITCH",
    "T.DIST",
    "T.DIST.2T",
    "T.DIST.RT",
    "T.INV",
    "T.INV.2T",
    "T.TEST",
    "TAKE",
    "TEXTAFTER",
    "TEXTBEFORE",
    "TEXTJOIN",
    "TEXTSPLIT",
    "TOCOL",
    "TOROW",
    "UNICHAR",
    "UNICODE",
    "UNIQUE",
    "VALUETOTEXT",
    "VAR.P",
    "VAR.S",
    "VSTACK",
    "WEIBULL.DIST",
    "WORKDAY.INTL",
    "WRAPCOLS",
    "WRAPROWS",
    "XLOOKUP",
    "XMATCH",
    "XOR",
    "Z.TEST",
];

pub(crate) fn normalize_display_formula(input: &str) -> String {
    let Some(normalized) = formula_model::normalize_formula_text(input) else {
        return String::new();
    };
    strip_xlfn_prefixes(&normalized)
}

pub(crate) fn strip_xlfn_prefixes(formula: &str) -> String {
    let bytes = formula.as_bytes();
    let mut out: Vec<u8> = Vec::new();
    if out.try_reserve_exact(bytes.len()).is_err() {
        return formula.to_string();
    }

    let mut i = 0;
    let mut in_string = false;

    while i < bytes.len() {
        match bytes[i] {
            b'"' => {
                out.push(b'"');
                if in_string {
                    // Excel escapes `"` within strings by doubling it (`""`).
                    if i + 1 < bytes.len() && bytes[i + 1] == b'"' {
                        out.push(b'"');
                        i += 2;
                        continue;
                    }
                    in_string = false;
                } else {
                    in_string = true;
                }
                i += 1;
                continue;
            }
            _ if !in_string && has_xlfn_prefix_at(bytes, i) => {
                // Only treat `_xlfn.` as a prefix when it occurs at an identifier boundary.
                // This avoids stripping `_xlfn.` when it appears in the middle of some other
                // identifier (e.g. a user-defined function name containing `_xlfn`).
                if i > 0 && is_ident_byte(bytes[i - 1]) {
                    // Fall through: emit the current byte verbatim.
                } else {
                    let after_prefix = i + XL_FN_PREFIX.len();
                    let mut j = after_prefix;
                    while j < bytes.len() && is_ident_byte(bytes[j]) {
                        j += 1;
                    }

                    if j > after_prefix {
                        let mut k = j;
                        while k < bytes.len() && bytes[k].is_ascii_whitespace() {
                            k += 1;
                        }
                        if k < bytes.len() && bytes[k] == b'(' {
                            // Only strip when it prefixes a function call.
                            i = after_prefix;
                            continue;
                        }
                    }
                }
            }
            _ => {}
        }

        out.push(bytes[i]);
        i += 1;
    }

    match String::from_utf8(out) {
        Ok(s) => s,
        Err(_) => {
            debug_assert!(false, "strip_xlfn_prefixes produced invalid utf-8");
            formula.to_string()
        }
    }
}

fn has_xlfn_prefix_at(bytes: &[u8], i: usize) -> bool {
    let prefix_len = XL_FN_PREFIX_BYTES.len();
    let Some(end) = i.checked_add(prefix_len) else {
        return false;
    };
    let Some(slice) = bytes.get(i..end) else {
        return false;
    };
    slice.eq_ignore_ascii_case(XL_FN_PREFIX_BYTES)
}

pub(crate) fn add_xlfn_prefixes(formula: &str) -> String {
    let bytes = formula.as_bytes();
    let mut out: Vec<u8> = Vec::new();
    if out.try_reserve_exact(bytes.len()).is_err() {
        return formula.to_string();
    }

    let mut i = 0;
    let mut in_string = false;

    while i < bytes.len() {
        match bytes[i] {
            b'"' => {
                out.push(b'"');
                if in_string {
                    if i + 1 < bytes.len() && bytes[i + 1] == b'"' {
                        out.push(b'"');
                        i += 2;
                        continue;
                    }
                    in_string = false;
                } else {
                    in_string = true;
                }
                i += 1;
                continue;
            }
            _ if !in_string && is_ident_start_byte(bytes[i]) => {
                let start = i;
                let mut end = start + 1;
                while end < bytes.len() && is_ident_byte(bytes[end]) {
                    end += 1;
                }

                let ident = &formula[start..end];
                let mut k = end;
                while k < bytes.len() && bytes[k].is_ascii_whitespace() {
                    k += 1;
                }

                let is_func_call = k < bytes.len() && bytes[k] == b'(';
                if is_func_call && needs_xlfn_prefix(ident) {
                    let has_prefix = start >= XL_FN_PREFIX.len()
                        && &formula[start - XL_FN_PREFIX.len()..start] == XL_FN_PREFIX;
                    if !has_prefix {
                        out.extend_from_slice(XL_FN_PREFIX.as_bytes());
                    }
                }

                out.extend_from_slice(&bytes[start..end]);
                i = end;
                continue;
            }
            _ => {}
        }

        out.push(bytes[i]);
        i += 1;
    }

    match String::from_utf8(out) {
        Ok(s) => s,
        Err(_) => {
            debug_assert!(false, "add_xlfn_prefixes produced invalid utf-8");
            formula.to_string()
        }
    }
}

fn needs_xlfn_prefix(ident: &str) -> bool {
    // Excel may store certain function namespaces without the leading `_xlfn.` prefix in the
    // UI-facing formula text we operate on here (because we strip `_xlfn.` when normalizing
    // display formulas). When writing back to OOXML we must restore `_xlfn.` for these
    // namespace-qualified functions so round-trips preserve the original file form.
    if ident
        .get(..XL_WS_PREFIX.len())
        .is_some_and(|p| p.eq_ignore_ascii_case(XL_WS_PREFIX))
        || ident
            .get(..XL_UDF_PREFIX.len())
            .is_some_and(|p| p.eq_ignore_ascii_case(XL_UDF_PREFIX))
    {
        return true;
    }
    XL_FN_REQUIRED_FUNCTIONS
        .iter()
        .any(|required| ident.eq_ignore_ascii_case(required))
}

fn is_ident_start_byte(b: u8) -> bool {
    b.is_ascii_alphabetic() || b == b'_'
}

fn is_ident_continue_byte(b: u8) -> bool {
    b.is_ascii_alphanumeric() || b == b'_'
}

fn is_ident_byte(b: u8) -> bool {
    is_ident_continue_byte(b) || b == b'.'
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::HashSet;

    #[test]
    fn normalize_display_formula_strips_leading_equals_and_trims() {
        assert_eq!(normalize_display_formula("=1+1"), "1+1");
        assert_eq!(normalize_display_formula("   = 1+1  "), "1+1");
        assert_eq!(normalize_display_formula("   "), "");
        assert_eq!(normalize_display_formula("="), "");
    }

    #[test]
    fn strip_xlfn_prefixes_ignores_string_literals() {
        let input = r#"_xlfn.CONCAT("_xlfn.",_xlfn.SEQUENCE(1))"#;
        assert_eq!(
            strip_xlfn_prefixes(input),
            r#"CONCAT("_xlfn.",SEQUENCE(1))"#
        );
    }

    #[test]
    fn strip_xlfn_prefixes_is_case_insensitive_for_prefix() {
        assert_eq!(strip_xlfn_prefixes("_XLFN.SEQUENCE(1)"), "SEQUENCE(1)");
    }

    #[test]
    fn strip_xlfn_prefixes_does_not_strip_mid_identifier() {
        let input = "FOO_xlfn.SEQUENCE(1)";
        assert_eq!(strip_xlfn_prefixes(input), input);
    }

    #[test]
    fn xlfn_roundtrip_preserves_xlws_namespace_functions() {
        let file = r#"_xlfn._xlws.WEBSERVICE("https://example.com")"#;
        let display = strip_xlfn_prefixes(file);
        assert_eq!(display, r#"_xlws.WEBSERVICE("https://example.com")"#);
        assert_eq!(add_xlfn_prefixes(&display), file);
    }

    #[test]
    fn xlfn_roundtrip_preserves_xludf_namespace_functions() {
        let file = "_xlfn._xludf.MYFUNC(1)";
        let display = strip_xlfn_prefixes(file);
        assert_eq!(display, "_xludf.MYFUNC(1)");
        assert_eq!(add_xlfn_prefixes(&display), file);
    }

    #[test]
    fn xlfn_required_functions_are_sorted_and_unique() {
        let mut prev: Option<&str> = None;
        let mut seen = HashSet::new();

        for &name in XL_FN_REQUIRED_FUNCTIONS {
            if let Some(prev) = prev {
                assert!(
                    prev < name,
                    "XL_FN_REQUIRED_FUNCTIONS must be ASCII-sorted; found out-of-order entries: {prev} then {name}"
                );
            }
            assert!(
                seen.insert(name),
                "XL_FN_REQUIRED_FUNCTIONS must not contain duplicates; duplicate entry: {name}"
            );
            prev = Some(name);
        }
    }

    #[test]
    fn add_xlfn_prefixes_roundtrips_known_functions() {
        let input = r#"CONCAT("_xlfn.",SEQUENCE(1))"#;
        let expected = r#"_xlfn.CONCAT("_xlfn.",_xlfn.SEQUENCE(1))"#;
        assert_eq!(add_xlfn_prefixes(input), expected);
    }

    #[test]
    fn xlfn_roundtrip_preserves_image_function() {
        let input = r#"IMAGE("https://example.com/x.png")"#;
        let expected = r#"_xlfn.IMAGE("https://example.com/x.png")"#;
        assert_eq!(add_xlfn_prefixes(input), expected);
        assert_eq!(strip_xlfn_prefixes(expected), input);
    }

    #[test]
    fn add_xlfn_prefixes_prefixes_textsplit() {
        let input = r#"TEXTSPLIT("a,b",",")"#;
        let expected = r#"_xlfn.TEXTSPLIT("a,b",",")"#;
        assert_eq!(add_xlfn_prefixes(input), expected);
    }

    #[test]
    fn add_xlfn_prefixes_prefixes_textafter() {
        let input = r#"TEXTAFTER("a_b","_")"#;
        let expected = r#"_xlfn.TEXTAFTER("a_b","_")"#;
        assert_eq!(add_xlfn_prefixes(input), expected);
    }

    #[test]
    fn add_xlfn_prefixes_prefixes_textbefore() {
        let input = r#"TEXTBEFORE("a_b","_")"#;
        let expected = r#"_xlfn.TEXTBEFORE("a_b","_")"#;
        assert_eq!(add_xlfn_prefixes(input), expected);
    }

    #[test]
    fn add_xlfn_prefixes_prefixes_valuetotext() {
        let input = r#"VALUETOTEXT(1)"#;
        let expected = r#"_xlfn.VALUETOTEXT(1)"#;
        assert_eq!(add_xlfn_prefixes(input), expected);
    }

    #[test]
    fn add_xlfn_prefixes_prefixes_maxifs_and_minifs() {
        let input = r#"MAXIFS(A1:A3,B1:B3,1)+MINIFS(A1:A3,B1:B3,1)"#;
        let expected = r#"_xlfn.MAXIFS(A1:A3,B1:B3,1)+_xlfn.MINIFS(A1:A3,B1:B3,1)"#;
        assert_eq!(add_xlfn_prefixes(input), expected);
    }

    #[test]
    fn add_xlfn_prefixes_prefixes_xor() {
        let input = r#"XOR(TRUE,FALSE)"#;
        let expected = r#"_xlfn.XOR(TRUE,FALSE)"#;
        assert_eq!(add_xlfn_prefixes(input), expected);
    }

    #[test]
    fn add_xlfn_prefixes_prefixes_multiple_modern_functions() {
        let input = r#"TAKE(SEQUENCE(3),2)+HSTACK({1;2},{3;4})"#;
        let expected = r#"_xlfn.TAKE(_xlfn.SEQUENCE(3),2)+_xlfn.HSTACK({1;2},{3;4})"#;
        assert_eq!(add_xlfn_prefixes(input), expected);
    }

    #[test]
    fn add_xlfn_prefixes_handles_dynamic_array_helpers() {
        let input = "TAKE(SEQUENCE(1),1)";
        let expected = "_xlfn.TAKE(_xlfn.SEQUENCE(1),1)";
        assert_eq!(add_xlfn_prefixes(input), expected);
    }

    #[test]
    fn add_xlfn_prefixes_handles_shape_functions() {
        let input = "DROP(A1:A3,1)+TAKE(B1:B3,1)+CHOOSECOLS(C1:E1,1)+CHOOSEROWS(C1:C3,1)+EXPAND(D1:E2,3,4)";
        let expected = "_xlfn.DROP(A1:A3,1)+_xlfn.TAKE(B1:B3,1)+_xlfn.CHOOSECOLS(C1:E1,1)+_xlfn.CHOOSEROWS(C1:C3,1)+_xlfn.EXPAND(D1:E2,3,4)";
        assert_eq!(add_xlfn_prefixes(input), expected);
    }

    #[test]
    fn add_xlfn_prefixes_handles_lambda_helpers() {
        let input = "MAP(SEQUENCE(3),LAMBDA(x,x))";
        let expected = "_xlfn.MAP(_xlfn.SEQUENCE(3),_xlfn.LAMBDA(x,x))";
        assert_eq!(add_xlfn_prefixes(input), expected);
    }

    #[test]
    fn add_xlfn_prefixes_handles_isomitted() {
        let input = "ISOMITTED(x)";
        let expected = "_xlfn.ISOMITTED(x)";
        assert_eq!(add_xlfn_prefixes(input), expected);
    }

    #[test]
    fn add_xlfn_prefixes_handles_dotted_function_names() {
        let input = "ISO.WEEKNUM(1)+WORKDAY.INTL(1,2)";
        let expected = "_xlfn.ISO.WEEKNUM(1)+_xlfn.WORKDAY.INTL(1,2)";
        assert_eq!(add_xlfn_prefixes(input), expected);
    }

    #[test]
    fn add_xlfn_prefixes_handles_statistical_dot_functions() {
        let input = "STDEV.S(A1:A3)+VAR.P(A1:A3)";
        let expected = "_xlfn.STDEV.S(A1:A3)+_xlfn.VAR.P(A1:A3)";
        assert_eq!(add_xlfn_prefixes(input), expected);
    }

    #[test]
    fn xlfn_roundtrip_preserves_forecast_ets_functions() {
        let display = "FORECAST.ETS(1,2,3)+FORECAST.ETS.CONFINT(1,2,3)+FORECAST.ETS.SEASONALITY(1,2,3)+FORECAST.ETS.STAT(1,2,3)";
        let file = "_xlfn.FORECAST.ETS(1,2,3)+_xlfn.FORECAST.ETS.CONFINT(1,2,3)+_xlfn.FORECAST.ETS.SEASONALITY(1,2,3)+_xlfn.FORECAST.ETS.STAT(1,2,3)";
        assert_eq!(add_xlfn_prefixes(display), file);
        assert_eq!(strip_xlfn_prefixes(file), display);
    }

    #[test]
    fn xl_fn_required_functions_is_sorted_and_unique() {
        for window in XL_FN_REQUIRED_FUNCTIONS.windows(2) {
            assert!(
                window[0] < window[1],
                "XL_FN_REQUIRED_FUNCTIONS must be sorted (ASCII) and unique: {:?}",
                window
            );
        }
    }

    #[test]
    fn add_and_strip_xlfn_prefixes_roundtrip_required_functions() {
        for func in XL_FN_REQUIRED_FUNCTIONS {
            let display = format!("{func}(1)");
            let file = format!("{XL_FN_PREFIX}{func}(1)");
            assert_eq!(
                add_xlfn_prefixes(&display),
                file,
                "add_xlfn_prefixes should prefix {func}"
            );
            assert_eq!(
                strip_xlfn_prefixes(&file),
                display,
                "strip_xlfn_prefixes should strip {func}"
            );
            assert_eq!(
                add_xlfn_prefixes(&strip_xlfn_prefixes(&file)),
                file,
                "add/strip roundtrip should preserve {func}"
            );
        }
    }

    #[test]
    fn formula_display_roundtrip() {
        let file = r#"_xlfn.CONCAT("_xlfn.",_xlfn.SEQUENCE(1))"#;
        let display = strip_xlfn_prefixes(file);
        assert_eq!(display, r#"CONCAT("_xlfn.",SEQUENCE(1))"#);
        assert_eq!(add_xlfn_prefixes(&display), file);
    }
}
