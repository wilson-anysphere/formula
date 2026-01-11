const XL_FN_PREFIX: &str = "_xlfn.";
const XL_FN_PREFIX_BYTES: &[u8] = b"_xlfn.";
const XL_WS_PREFIX: &str = "_xlws.";
const XL_UDF_PREFIX: &str = "_xludf.";

// Functions that Excel stores in OOXML formulas with an `_xlfn.` prefix for forward
// compatibility (typically Excel 365 "future functions").
//
// Keep this list sorted (ASCII) for maintainability.
const XL_FN_REQUIRED_FUNCTIONS: &[&str] = &[
    "AGGREGATE",
    "BYCOL",
    "BYROW",
    "CEILING.MATH",
    "CEILING.PRECISE",
    "CHOOSECOLS",
    "CHOOSEROWS",
    "CONCAT",
    "DROP",
    "EXPAND",
    "FILTER",
    "FLOOR.MATH",
    "FLOOR.PRECISE",
    "HSTACK",
    "IFNA",
    "IFS",
    "ISO.CEILING",
    "ISO.WEEKNUM",
    "ISOMITTED",
    "ISOWEEKNUM",
    "LAMBDA",
    "LET",
    "MAKEARRAY",
    "MAP",
    "NETWORKDAYS.INTL",
    "NUMBERVALUE",
    "RANDARRAY",
    "REDUCE",
    "SCAN",
    "SEQUENCE",
    "SORT",
    "SORTBY",
    "SWITCH",
    "TAKE",
    "TEXTJOIN",
    "TEXTSPLIT",
    "TOCOL",
    "TOROW",
    "UNIQUE",
    "VSTACK",
    "WORKDAY.INTL",
    "WRAPCOLS",
    "WRAPROWS",
    "XLOOKUP",
    "XMATCH",
];

pub(crate) fn normalize_display_formula(input: &str) -> String {
    let Some(normalized) = formula_model::normalize_formula_text(input) else {
        return String::new();
    };
    strip_xlfn_prefixes(&normalized)
}

pub(crate) fn strip_xlfn_prefixes(formula: &str) -> String {
    let bytes = formula.as_bytes();
    let mut out: Vec<u8> = Vec::with_capacity(bytes.len());

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
            _ => {}
        }

        out.push(bytes[i]);
        i += 1;
    }

    String::from_utf8(out).expect("formula rewrite should preserve utf-8")
}

fn has_xlfn_prefix_at(bytes: &[u8], i: usize) -> bool {
    let prefix_len = XL_FN_PREFIX_BYTES.len();
    if i.saturating_add(prefix_len) > bytes.len() {
        return false;
    }
    bytes[i..i + prefix_len]
        .iter()
        .zip(XL_FN_PREFIX_BYTES)
        .all(|(&b, &p)| b.to_ascii_lowercase() == p)
}

pub(crate) fn add_xlfn_prefixes(formula: &str) -> String {
    let bytes = formula.as_bytes();
    let mut out: Vec<u8> = Vec::with_capacity(bytes.len());

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

    String::from_utf8(out).expect("formula rewrite should preserve utf-8")
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
    fn add_xlfn_prefixes_roundtrips_known_functions() {
        let input = r#"CONCAT("_xlfn.",SEQUENCE(1))"#;
        let expected = r#"_xlfn.CONCAT("_xlfn.",_xlfn.SEQUENCE(1))"#;
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
    fn formula_display_roundtrip() {
        let file = r#"_xlfn.CONCAT("_xlfn.",_xlfn.SEQUENCE(1))"#;
        let display = strip_xlfn_prefixes(file);
        assert_eq!(display, r#"CONCAT("_xlfn.",SEQUENCE(1))"#);
        assert_eq!(add_xlfn_prefixes(&display), file);
    }
}
