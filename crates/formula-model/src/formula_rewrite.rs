use unicode_normalization::UnicodeNormalization;

/// Excel compares sheet names case-insensitively.
///
/// We approximate Excel's behavior by:
/// - normalizing with Unicode NFKC (compatibility normalization)
/// - applying Unicode uppercasing
///
/// This is deterministic and locale-independent. It is not a byte-level ASCII-only compare,
/// and it handles common Unicode edge cases (e.g. compatibility characters).
pub(crate) fn sheet_name_eq_case_insensitive(a: &str, b: &str) -> bool {
    a.nfkc()
        .flat_map(|c| c.to_uppercase())
        .eq(b.nfkc().flat_map(|c| c.to_uppercase()))
}

fn looks_like_a1_cell_reference(name: &str) -> bool {
    // If an unquoted sheet name looks like a cell reference (e.g. "A1" or "XFD1048576"),
    // Excel requires quoting to disambiguate.
    let mut chars = name.chars().peekable();
    let Some(first) = chars.peek().copied() else {
        return false;
    };
    if !first.is_ascii_alphabetic() {
        return false;
    }

    let mut letters = String::new();
    while let Some(c) = chars.peek().copied() {
        if c.is_ascii_alphabetic() {
            if letters.len() >= 3 {
                return false;
            }
            letters.push(c);
            chars.next();
        } else {
            break;
        }
    }

    let mut digits = String::new();
    while let Some(c) = chars.peek().copied() {
        if c.is_ascii_digit() {
            digits.push(c);
            chars.next();
        } else {
            break;
        }
    }

    if letters.is_empty() || digits.is_empty() || chars.peek().is_some() {
        return false;
    }

    // Reject impossible columns (beyond XFD). This keeps the check cheap and avoids quoting
    // names like "SHEET1" where the "letters" part is > 3 and already returned false above.
    let col = letters.chars().fold(0u32, |acc, c| {
        acc * 26 + (c.to_ascii_uppercase() as u32 - 'A' as u32 + 1)
    });
    col <= 16_384
}

fn is_valid_unquoted_sheet_name(name: &str) -> bool {
    let mut chars = name.chars();
    let Some(first) = chars.next() else {
        return false;
    };

    // Excel allows many Unicode letters in unquoted names, but still requires an identifier-like
    // start character (can't start with a digit).
    if first.is_ascii_digit() {
        return false;
    }
    if !(first == '_' || first.is_alphabetic()) {
        return false;
    }

    if !chars.all(|ch| ch == '_' || ch == '.' || ch.is_alphanumeric()) {
        return false;
    }

    !looks_like_a1_cell_reference(name)
}

fn needs_quoting_for_sheet_reference(
    workbook_prefix: Option<&str>,
    start: &str,
    end: Option<&str>,
) -> bool {
    // For safety (and because we cannot know if the external workbook is open/closed), always
    // quote external workbook references. This is accepted by Excel and avoids subtle parser
    // differences for `[Book]Sheet` vs `'[Book]Sheet'`.
    if workbook_prefix.is_some() {
        return true;
    }

    if !is_valid_unquoted_sheet_name(start) {
        return true;
    }

    end.is_some_and(|end| !is_valid_unquoted_sheet_name(end))
}

fn escape_single_quotes(s: &str) -> String {
    s.replace('\'', "''")
}

fn format_sheet_reference(workbook_prefix: Option<&str>, start: &str, end: Option<&str>) -> String {
    let mut content = String::new();
    if let Some(prefix) = workbook_prefix {
        content.push_str(prefix);
    }
    content.push_str(start);
    if let Some(end) = end {
        content.push(':');
        content.push_str(end);
    }

    if needs_quoting_for_sheet_reference(workbook_prefix, start, end) {
        format!("'{}'", escape_single_quotes(&content))
    } else {
        content
    }
}

fn split_workbook_prefix(sheet_spec: &str) -> (Option<&str>, &str) {
    let Some(open) = sheet_spec.find('[') else {
        return (None, sheet_spec);
    };
    let Some(close_rel) = sheet_spec[open..].find(']') else {
        return (None, sheet_spec);
    };
    let close = open + close_rel;
    let prefix_end = close + 1;
    if prefix_end >= sheet_spec.len() {
        return (None, sheet_spec);
    }
    let (prefix, remainder) = sheet_spec.split_at(prefix_end);
    (Some(prefix), remainder)
}

fn rewrite_sheet_spec(spec: &str, old_name: &str, new_name: &str) -> Option<String> {
    let (workbook_prefix, remainder) = split_workbook_prefix(spec);
    let mut parts = remainder.splitn(2, ':');
    let start = parts.next().unwrap_or_default();
    let end = parts.next();

    let changed_start = sheet_name_eq_case_insensitive(start, old_name);
    let renamed_start = if changed_start { new_name } else { start };

    let (renamed_end, changed_end) = match end {
        Some(end) => {
            let changed = sheet_name_eq_case_insensitive(end, old_name);
            let renamed = if changed { new_name } else { end };
            (Some(renamed), changed)
        }
        None => (None, false),
    };

    if !changed_start && !changed_end {
        return None;
    }

    Some(format_sheet_reference(
        workbook_prefix,
        renamed_start,
        renamed_end,
    ))
}

fn parse_quoted_sheet_spec(formula: &str, start: usize) -> Option<(usize, &str, String)> {
    let bytes = formula.as_bytes();
    if bytes.get(start) != Some(&b'\'') {
        return None;
    }

    let mut i = start + 1;
    let mut unescaped = String::new();

    while i < bytes.len() {
        match bytes[i] {
            b'\'' => {
                if bytes.get(i + 1) == Some(&b'\'') {
                    unescaped.push('\'');
                    i += 2;
                } else {
                    i += 1;
                    break;
                }
            }
            _ => {
                let ch = formula[i..].chars().next()?;
                unescaped.push(ch);
                i += ch.len_utf8();
            }
        }
    }

    if i >= bytes.len() || bytes[i] != b'!' {
        return None;
    }

    let next = i + 1;
    Some((next, &formula[start..next], unescaped))
}

fn parse_unquoted_sheet_spec(formula: &str, start: usize) -> Option<(usize, &str, &str)> {
    let bytes = formula.as_bytes();
    let mut i = start;

    let first = formula[i..].chars().next()?;
    if first != '[' && first != '_' && !first.is_alphabetic() {
        return None;
    }

    // External workbook prefix: `[Book1.xlsx]Sheet1!A1`
    if first == '[' {
        // Scan until the closing `]` (workbook name can contain many characters).
        i += 1;
        while i < bytes.len() {
            if bytes[i] == b']' {
                i += 1;
                break;
            }
            let ch = formula[i..].chars().next()?;
            i += ch.len_utf8();
        }

        if i >= bytes.len() {
            return None;
        }

        let next_ch = formula[i..].chars().next()?;
        if next_ch != '_' && !next_ch.is_alphabetic() {
            // Not a workbook+sheet reference; likely a structured reference.
            return None;
        }
    } else {
        i += first.len_utf8();
    }

    while i < bytes.len() {
        if bytes[i] == b'!' {
            let next = i + 1;
            return Some((next, &formula[start..next], &formula[start..i]));
        }
        let ch = formula[i..].chars().next()?;
        if ch == '_' || ch == '.' || ch == ':' || ch.is_alphanumeric() {
            i += ch.len_utf8();
            continue;
        }
        break;
    }

    None
}

fn parse_error_literal(formula: &str, start: usize) -> Option<(usize, &str)> {
    let bytes = formula.as_bytes();
    if bytes.get(start) != Some(&b'#') {
        return None;
    }

    let mut i = start + 1;
    while i < bytes.len() {
        let ch = formula[i..].chars().next()?;
        match ch {
            'A'..='Z' | 'a'..='z' | '0'..='9' | '/' | '_' | '.' => {
                i += ch.len_utf8();
            }
            '!' | '?' => {
                i += ch.len_utf8();
                break;
            }
            _ => break,
        }
    }

    if i == start + 1 {
        return None;
    }

    Some((i, &formula[start..i]))
}

/// Rewrite all sheet references inside a formula when a sheet is renamed.
///
/// This is intentionally conservative: it only rewrites tokens that *parse* as
/// sheet references (`Sheet!A1`, `'My Sheet'!A1`, `Sheet1:Sheet3!A1`, etc) and it
/// does not touch string literals.
pub fn rewrite_sheet_names_in_formula(formula: &str, old_name: &str, new_name: &str) -> String {
    let mut out = String::with_capacity(formula.len());
    let mut i = 0;
    let mut in_string = false;
    let bytes = formula.as_bytes();

    while i < bytes.len() {
        if in_string {
            let ch = formula[i..]
                .chars()
                .next()
                .expect("i always at char boundary");
            out.push(ch);
            if ch == '"' {
                if bytes.get(i + 1) == Some(&b'"') {
                    out.push('"');
                    i += 2;
                    continue;
                }
                in_string = false;
            }
            i += ch.len_utf8();
            continue;
        }

        if bytes[i] == b'"' {
            in_string = true;
            out.push('"');
            i += 1;
            continue;
        }

        if bytes[i] == b'#' {
            if let Some((next, raw)) = parse_error_literal(formula, i) {
                out.push_str(raw);
                i = next;
                continue;
            }
        }

        if bytes[i] == b'\'' {
            if let Some((next, raw, sheet_spec)) = parse_quoted_sheet_spec(formula, i) {
                if let Some(rewritten) = rewrite_sheet_spec(&sheet_spec, old_name, new_name) {
                    out.push_str(&rewritten);
                    out.push('!');
                } else {
                    out.push_str(raw);
                }
                i = next;
                continue;
            }
        }

        if let Some((next, raw, sheet_spec)) = parse_unquoted_sheet_spec(formula, i) {
            if let Some(rewritten) = rewrite_sheet_spec(sheet_spec, old_name, new_name) {
                out.push_str(&rewritten);
                out.push('!');
            } else {
                out.push_str(raw);
            }
            i = next;
            continue;
        }

        let ch = formula[i..]
            .chars()
            .next()
            .expect("i always at char boundary");
        out.push(ch);
        i += ch.len_utf8();
    }

    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn format_sheet_reference_rules() {
        assert_eq!(format_sheet_reference(None, "Sheet1", None), "Sheet1");
        assert_eq!(format_sheet_reference(None, "My Sheet", None), "'My Sheet'");
        assert_eq!(format_sheet_reference(None, "O'Brien", None), "'O''Brien'");
        assert_eq!(
            format_sheet_reference(None, "Sheet1", Some("Sheet3")),
            "Sheet1:Sheet3"
        );
        assert_eq!(
            format_sheet_reference(None, "Sheet 1", Some("Sheet 3")),
            "'Sheet 1:Sheet 3'"
        );
        assert_eq!(
            format_sheet_reference(Some("[Book1.xlsx]"), "Sheet1", None),
            "'[Book1.xlsx]Sheet1'"
        );
        assert_eq!(
            format_sheet_reference(Some("C:\\path\\[Book1.xlsx]"), "Sheet1", None),
            "'C:\\path\\[Book1.xlsx]Sheet1'"
        );
    }

    #[test]
    fn rewrite_simple_sheet_ref() {
        assert_eq!(
            rewrite_sheet_names_in_formula("=Sheet1!A1", "Sheet1", "Summary"),
            "=Summary!A1"
        );
    }

    #[test]
    fn rewrite_quoted_sheet_ref() {
        assert_eq!(
            rewrite_sheet_names_in_formula("='Sheet 1'!A1", "Sheet 1", "My Sheet"),
            "='My Sheet'!A1"
        );
    }

    #[test]
    fn rewrite_sheet_ref_with_apostrophe() {
        assert_eq!(
            rewrite_sheet_names_in_formula("='O''Brien'!A1", "O'Brien", "Data"),
            "=Data!A1"
        );
        assert_eq!(
            rewrite_sheet_names_in_formula("=Data!A1", "Data", "O'Brien"),
            "='O''Brien'!A1"
        );
    }

    #[test]
    fn rewrite_does_not_touch_string_literals() {
        assert_eq!(
            rewrite_sheet_names_in_formula("=\"Sheet1!A1\"", "Sheet1", "Data"),
            "=\"Sheet1!A1\""
        );
    }

    #[test]
    fn rewrite_is_utf8_safe() {
        assert_eq!(
            rewrite_sheet_names_in_formula("=1+\"😀\"", "Sheet1", "Data"),
            "=1+\"😀\""
        );
    }

    #[test]
    fn rewrite_3d_reference() {
        assert_eq!(
            rewrite_sheet_names_in_formula("=Sheet1:Sheet3!A1", "Sheet1", "Data"),
            "=Data:Sheet3!A1"
        );
    }

    #[test]
    fn rewrite_quoted_3d_reference() {
        assert_eq!(
            rewrite_sheet_names_in_formula("='Sheet 1:Sheet 3'!A1", "Sheet 1", "Data"),
            "='Data:Sheet 3'!A1"
        );
    }

    #[test]
    fn rewrite_external_workbook_reference() {
        assert_eq!(
            rewrite_sheet_names_in_formula("='[Book1.xlsx]Sheet1'!A1", "Sheet1", "Data"),
            "='[Book1.xlsx]Data'!A1"
        );
    }

    #[test]
    fn rewrite_unquoted_external_workbook_reference() {
        assert_eq!(
            rewrite_sheet_names_in_formula("=[Book1.xlsx]Sheet1!A1", "Sheet1", "Data"),
            "='[Book1.xlsx]Data'!A1"
        );
    }

    #[test]
    fn rewrite_external_reference_with_path() {
        assert_eq!(
            rewrite_sheet_names_in_formula("='C:\\path\\[Book1.xlsx]Sheet1'!A1", "Sheet1", "Data",),
            "='C:\\path\\[Book1.xlsx]Data'!A1"
        );
    }

    #[test]
    fn rewrite_unicode_sheet_names_case_insensitive() {
        assert_eq!(
            rewrite_sheet_names_in_formula("='Résumé'!A1+résumé!B2", "Résumé", "Data"),
            "=Data!A1+Data!B2"
        );
        assert_eq!(
            rewrite_sheet_names_in_formula("='📊'!A1", "📊", "Data"),
            "=Data!A1"
        );
    }

    #[test]
    fn rewrite_does_not_touch_non_references() {
        assert_eq!(
            rewrite_sheet_names_in_formula("=Sheet1+1", "Sheet1", "Data"),
            "=Sheet1+1"
        );
        assert_eq!(
            rewrite_sheet_names_in_formula("=SHEET1(A1)", "Sheet1", "Data"),
            "=SHEET1(A1)"
        );
        assert_eq!(
            rewrite_sheet_names_in_formula("=Table1[Sheet1]", "Sheet1", "Data"),
            "=Table1[Sheet1]"
        );
    }

    #[test]
    fn rewrite_does_not_touch_error_literals() {
        assert_eq!(
            rewrite_sheet_names_in_formula("=#REF!", "REF", "Data"),
            "=#REF!"
        );
        assert_eq!(
            rewrite_sheet_names_in_formula("=#VALUE!+VALUE!A1", "VALUE", "Data"),
            "=#VALUE!+Data!A1"
        );
        assert_eq!(
            rewrite_sheet_names_in_formula("=#SPILL!+SPILL!A1", "SPILL", "Data"),
            "=#SPILL!+Data!A1"
        );
    }

    #[test]
    fn fuzz_unicode_sheet_names_roundtrip() {
        // Deterministic "fuzz" test: generates a variety of Unicode sheet names to ensure
        // rewriting never panics, remains valid UTF-8, and only rewrites true sheet refs.
        struct Lcg(u64);
        impl Lcg {
            fn next_u32(&mut self) -> u32 {
                self.0 = self.0.wrapping_mul(6364136223846793005).wrapping_add(1);
                (self.0 >> 32) as u32
            }

            fn gen_range(&mut self, max: u32) -> u32 {
                self.next_u32() % max
            }
        }

        const CHARSET: &[char] = &[
            'A', 'b', 'Z', '0', '9', '_', '.', '-', ' ', '\'', 'é', 'ß', 'İ', 'ı', '中', 'デ',
            '😀', '📊',
        ];

        fn gen_name(rng: &mut Lcg) -> String {
            let len = (rng.gen_range(10) + 1) as usize;
            let mut s = String::new();
            for _ in 0..len {
                let ch = CHARSET[rng.gen_range(CHARSET.len() as u32) as usize];
                s.push(ch);
            }
            // Avoid empty/whitespace-only names which can't exist in Excel.
            if s.trim().is_empty() {
                "Sheet".to_string()
            } else {
                s
            }
        }

        let mut rng = Lcg(0x1234_5678_9ABC_DEF0);
        for _ in 0..250 {
            let old_name = gen_name(&mut rng);
            let mut new_name = gen_name(&mut rng);
            if sheet_name_eq_case_insensitive(&old_name, &new_name) {
                new_name.push('X');
            }

            let old_ref = format_sheet_reference(None, &old_name, None);
            let new_ref = format_sheet_reference(None, &new_name, None);
            let formula = format!("={}!A1+\"{}!A1\"+{}!B2", old_ref, old_name, old_ref);
            let expected = format!("={}!A1+\"{}!A1\"+{}!B2", new_ref, old_name, new_ref);

            let rewritten = rewrite_sheet_names_in_formula(&formula, &old_name, &new_name);
            assert_eq!(rewritten, expected);
        }
    }
}
