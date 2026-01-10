fn is_valid_unquoted_sheet_name(name: &str) -> bool {
    let mut chars = name.chars();
    let Some(first) = chars.next() else {
        return false;
    };

    if first.is_ascii_digit() {
        return false;
    }

    if !(first.is_ascii_alphabetic() || first == '_') {
        return false;
    }

    chars.all(|ch| ch.is_ascii_alphanumeric() || ch == '_')
}

fn needs_quoting_for_sheet_reference(name: &str) -> bool {
    if let Some((start, end)) = name.split_once(':') {
        return !(is_valid_unquoted_sheet_name(start) && is_valid_unquoted_sheet_name(end));
    }

    !is_valid_unquoted_sheet_name(name)
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

    if needs_quoting_for_sheet_reference(&content) {
        format!("'{}'", escape_single_quotes(&content))
    } else {
        content
    }
}

fn split_workbook_prefix(sheet_spec: &str) -> (Option<&str>, &str) {
    if let Some(rest) = sheet_spec.strip_prefix('[') {
        if let Some(close_idx) = rest.find(']') {
            let prefix_len = close_idx + 2;
            let (prefix, remainder) = sheet_spec.split_at(prefix_len);
            return (Some(prefix), remainder);
        }
    }
    (None, sheet_spec)
}

fn rewrite_sheet_spec(spec: &str, old_name: &str, new_name: &str) -> Option<String> {
    let (workbook_prefix, remainder) = split_workbook_prefix(spec);
    let mut parts = remainder.splitn(2, ':');
    let start = parts.next().unwrap_or_default();
    let end = parts.next();

    let changed_start = start.eq_ignore_ascii_case(old_name);
    let renamed_start = if changed_start { new_name } else { start };

    let (renamed_end, changed_end) = match end {
        Some(end) => {
            let changed = end.eq_ignore_ascii_case(old_name);
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
            b => {
                unescaped.push(b as char);
                i += 1;
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

    let first = *bytes.get(i)?;
    if !(first.is_ascii_alphabetic() || first == b'_') {
        return None;
    }

    while i < bytes.len() {
        let b = bytes[i];
        if b == b'!' {
            let next = i + 1;
            return Some((next, &formula[start..next], &formula[start..i]));
        }
        if b.is_ascii_alphanumeric() || b == b'_' || b == b'.' || b == b':' {
            i += 1;
            continue;
        }
        break;
    }

    None
}

/// Rewrite all sheet references inside a formula when a sheet is renamed.
///
/// This is intentionally conservative: it only rewrites tokens that *parse* as
/// sheet references (`Sheet!A1`, `'My Sheet'!A1`, `Sheet1:Sheet3!A1`, etc) and it
/// does not touch string literals.
pub fn rewrite_sheet_names_in_formula(formula: &str, old_name: &str, new_name: &str) -> String {
    let mut out = String::with_capacity(formula.len());
    let mut i = 0;
    let bytes = formula.as_bytes();
    let mut in_string = false;

    while i < bytes.len() {
        let ch = bytes[i] as char;
        if in_string {
            out.push(ch);
            if ch == '"' {
                if bytes.get(i + 1) == Some(&b'"') {
                    out.push('"');
                    i += 2;
                    continue;
                }
                in_string = false;
            }
            i += 1;
            continue;
        }

        if ch == '"' {
            in_string = true;
            out.push('"');
            i += 1;
            continue;
        }

        if ch == '\'' {
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

        out.push(ch);
        i += 1;
    }

    out
}

#[cfg(test)]
mod tests {
    use super::*;

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
    fn rewrite_3d_reference() {
        assert_eq!(
            rewrite_sheet_names_in_formula("=Sheet1:Sheet3!A1", "Sheet1", "Data"),
            "=Data:Sheet3!A1"
        );
    }

    #[test]
    fn rewrite_external_workbook_reference() {
        assert_eq!(
            rewrite_sheet_names_in_formula("='[Book1.xlsx]Sheet1'!A1", "Sheet1", "Data"),
            "='[Book1.xlsx]Data'!A1"
        );
    }
}
