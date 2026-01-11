const XL_FN_PREFIX: &str = "_xlfn.";

const XL_FN_REQUIRED_FUNCTIONS: &[&str] = &[
    "FILTER",
    "UNIQUE",
    "SORT",
    "SORTBY",
    "SEQUENCE",
    "XLOOKUP",
    "XMATCH",
    "LET",
    "LAMBDA",
    "RANDARRAY",
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
            _ if !in_string && bytes[i..].starts_with(XL_FN_PREFIX.as_bytes()) => {
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
                while end < bytes.len() && is_ident_continue_byte(bytes[end]) {
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
        let input = r#"CONCAT("_xlfn.",_xlfn.SEQUENCE(1))"#;
        assert_eq!(strip_xlfn_prefixes(input), r#"CONCAT("_xlfn.",SEQUENCE(1))"#);
    }

    #[test]
    fn add_xlfn_prefixes_roundtrips_known_functions() {
        let input = r#"CONCAT("_xlfn.",SEQUENCE(1))"#;
        let expected = r#"CONCAT("_xlfn.",_xlfn.SEQUENCE(1))"#;
        assert_eq!(add_xlfn_prefixes(input), expected);
    }

    #[test]
    fn formula_display_roundtrip() {
        let file = r#"CONCAT("_xlfn.",_xlfn.SEQUENCE(1))"#;
        let display = strip_xlfn_prefixes(file);
        assert_eq!(display, r#"CONCAT("_xlfn.",SEQUENCE(1))"#);
        assert_eq!(add_xlfn_prefixes(&display), file);
    }
}
