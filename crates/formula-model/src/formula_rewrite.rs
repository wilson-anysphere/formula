use unicode_normalization::UnicodeNormalization;

/// Excel compares sheet names case-insensitively across Unicode.
///
/// We approximate Excel's behavior by normalizing both names with Unicode NFKC
/// (compatibility normalization) and then applying Unicode uppercasing. This is
/// deterministic and locale-independent.
pub fn sheet_name_eq_case_insensitive(a: &str, b: &str) -> bool {
    a.nfkc()
        .flat_map(|c| c.to_uppercase())
        .eq(b.nfkc().flat_map(|c| c.to_uppercase()))
}

/// Returns a canonical "case folded" representation of a sheet name that matches
/// [`sheet_name_eq_case_insensitive`].
///
/// This is useful when building hash map keys for sheet-name lookups that need
/// to behave like Excel (e.g. treating `Straße` and `STRASSE` as equal).
pub fn sheet_name_casefold(name: &str) -> String {
    if name.is_ascii() {
        return name.to_ascii_uppercase();
    }
    name.nfkc().flat_map(|c| c.to_uppercase()).collect()
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

fn looks_like_r1c1_cell_reference(name: &str) -> bool {
    // In R1C1 notation, `R`/`C` are valid relative references. Excel may also treat
    // `R123C456` as a cell reference even when the workbook is in A1 mode.
    if name.eq_ignore_ascii_case("r") || name.eq_ignore_ascii_case("c") {
        return true;
    }

    let bytes = name.as_bytes();
    if bytes.first().copied().map(|b| b.to_ascii_uppercase()) != Some(b'R') {
        return false;
    }

    let mut i = 1;
    while i < bytes.len() && bytes[i].is_ascii_digit() {
        i += 1;
    }

    if i >= bytes.len() || bytes[i].to_ascii_uppercase() != b'C' {
        return false;
    }

    i += 1;
    while i < bytes.len() && bytes[i].is_ascii_digit() {
        i += 1;
    }

    i == bytes.len()
}

fn is_reserved_unquoted_sheet_name(name: &str) -> bool {
    // Excel boolean literals are tokenized as keywords; quoting avoids ambiguity in formulas.
    name.eq_ignore_ascii_case("true") || name.eq_ignore_ascii_case("false")
}

fn is_valid_unquoted_sheet_name(name: &str) -> bool {
    let mut chars = name.chars();
    let Some(first) = chars.next() else {
        return false;
    };

    // Excel allows many Unicode sheet names, but it also permits quoting for all of them.
    // We keep the unquoted form conservative and ASCII-only; see note below.
    if first.is_ascii_digit() {
        return false;
    }
    // NOTE: We intentionally restrict unquoted sheet names to an ASCII identifier subset.
    // - This is always accepted by Excel (more quoting is still valid).
    // - It keeps our output compatible with the current `formula-engine` lexer, which treats
    //   non-ASCII identifiers as parse errors unless they are quoted.
    if !(first == '_' || first.is_ascii_alphabetic()) {
        return false;
    }

    if !chars.all(|ch| ch == '_' || ch == '.' || ch.is_ascii_alphanumeric()) {
        return false;
    }

    if is_reserved_unquoted_sheet_name(name) {
        return false;
    }

    !(looks_like_a1_cell_reference(name) || looks_like_r1c1_cell_reference(name))
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
    match crate::external_refs::split_external_workbook_prefix(sheet_spec) {
        Some((prefix, remainder)) => (Some(prefix), remainder),
        None => (None, sheet_spec),
    }
}

fn rewrite_sheet_spec(spec: &str, old_name: &str, new_name: &str) -> Option<String> {
    let (workbook_prefix, remainder) = split_workbook_prefix(spec);

    // Ambiguous case: in valid Excel formulas, an unquoted `:` inside a sheet spec denotes a 3D
    // sheet span (`Sheet1:Sheet3!A1`). However, legacy/invalid files can contain sheet names with
    // characters Excel normally forbids (including `:`). When importing such workbooks we still
    // want to rewrite references to the original sheet name after sanitization.
    //
    // If the *entire* unquoted sheet spec matches `old_name`, treat it as a single sheet name
    // rather than a 3D span.
    if remainder.contains(':') && sheet_name_eq_case_insensitive(remainder, old_name) {
        return Some(format_sheet_reference(workbook_prefix, new_name, None));
    }

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

#[derive(Clone, Debug, PartialEq, Eq)]
enum DeleteSheetSpecRewrite {
    Unchanged,
    Adjusted(String),
    Invalidate,
}

fn sheet_index_in_order(sheet_order: &[String], name: &str) -> Option<usize> {
    sheet_order
        .iter()
        .position(|sheet_name| sheet_name_eq_case_insensitive(sheet_name, name))
}

fn rewrite_sheet_spec_for_delete(
    spec: &str,
    deleted_sheet: &str,
    sheet_order: &[String],
) -> DeleteSheetSpecRewrite {
    let (workbook_prefix, remainder) = split_workbook_prefix(spec);
    // Deleting a local sheet must not rewrite references that explicitly target an external
    // workbook (e.g. `[Book.xlsx]Sheet1!A1`). Those are independent from the current workbook's
    // sheet list, even if the sheet name happens to match.
    if workbook_prefix.is_some() {
        return DeleteSheetSpecRewrite::Unchanged;
    }
    let mut parts = remainder.splitn(2, ':');
    let start = parts.next().unwrap_or_default();
    let end = parts.next();

    let Some(end) = end else {
        return if sheet_name_eq_case_insensitive(start, deleted_sheet) {
            DeleteSheetSpecRewrite::Invalidate
        } else {
            DeleteSheetSpecRewrite::Unchanged
        };
    };

    let start_matches = sheet_name_eq_case_insensitive(start, deleted_sheet);
    let end_matches = sheet_name_eq_case_insensitive(end, deleted_sheet);

    if !start_matches && !end_matches {
        return DeleteSheetSpecRewrite::Unchanged;
    }

    let Some(start_idx) = sheet_index_in_order(sheet_order, start) else {
        return DeleteSheetSpecRewrite::Invalidate;
    };
    let Some(end_idx) = sheet_index_in_order(sheet_order, end) else {
        return DeleteSheetSpecRewrite::Invalidate;
    };

    // The span references only the deleted sheet.
    if start_idx == end_idx {
        return DeleteSheetSpecRewrite::Invalidate;
    }

    let dir = if end_idx > start_idx { 1isize } else { -1isize };
    let mut new_start_idx = start_idx as isize;
    let mut new_end_idx = end_idx as isize;

    // When deleting a 3D boundary, Excel shifts it one sheet toward the other boundary.
    if start_matches {
        new_start_idx += dir;
    }
    if end_matches {
        new_end_idx -= dir;
    }

    let Some(new_start) = new_start_idx
        .try_into()
        .ok()
        .and_then(|idx: usize| sheet_order.get(idx))
    else {
        return DeleteSheetSpecRewrite::Invalidate;
    };
    let Some(new_end) = new_end_idx
        .try_into()
        .ok()
        .and_then(|idx: usize| sheet_order.get(idx))
    else {
        return DeleteSheetSpecRewrite::Invalidate;
    };

    let end = (!sheet_name_eq_case_insensitive(new_start, new_end)).then_some(new_end.as_str());

    DeleteSheetSpecRewrite::Adjusted(format_sheet_reference(
        workbook_prefix,
        new_start.as_str(),
        end,
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
        i = crate::external_refs::find_external_workbook_prefix_end_if_followed_by_sheet_or_name_token(
            formula, start,
        )?;

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

fn sheet_ref_tail_end(formula: &str, start: usize) -> usize {
    let bytes = formula.as_bytes();
    let mut i = start;
    let mut bracket_depth: u32 = 0;
    let mut paren_depth: u32 = 0;
    let mut in_string = false;

    while i < bytes.len() {
        let b = bytes[i];

        if in_string {
            if b == b'"' {
                if bytes.get(i + 1) == Some(&b'"') {
                    i += 2;
                    continue;
                }
                in_string = false;
            }
            i += 1;
            continue;
        }

        match b {
            b'"' => {
                in_string = true;
                i += 1;
            }
            b'[' => {
                bracket_depth = bracket_depth.saturating_add(1);
                i += 1;
            }
            b']' => {
                bracket_depth = bracket_depth.saturating_sub(1);
                i += 1;
            }
            b'(' => {
                paren_depth = paren_depth.saturating_add(1);
                i += 1;
            }
            b')' => {
                if paren_depth == 0 {
                    break;
                }
                paren_depth = paren_depth.saturating_sub(1);
                i += 1;
            }
            _ => {
                if bracket_depth == 0
                    && paren_depth == 0
                    && matches!(
                        b,
                        b' ' | b'\t'
                            | b'\n'
                            | b'\r'
                            | b','
                            | b';'
                            | b'+'
                            | b'-'
                            | b'*'
                            | b'/'
                            | b'^'
                            | b'&'
                            | b'='
                            | b'<'
                            | b'>'
                            | b'{'
                            | b'}'
                            | b'%'
                    )
                {
                    break;
                }
                i += 1;
            }
        }
    }

    i
}

/// Rewrite all sheet references inside a formula when a sheet is renamed.
///
/// This is intentionally conservative: it only rewrites tokens that *parse* as
/// sheet references (`Sheet!A1`, `'My Sheet'!A1`, `Sheet1:Sheet3!A1`, etc) and it
/// does not touch string literals.
///
/// This does **not** rewrite references that include an explicit workbook prefix
/// (e.g. `='[Book1.xlsx]Sheet1'!A1`), since those refer to an external workbook and
/// should not change when renaming a sheet inside the current workbook.
pub fn rewrite_sheet_names_in_formula(formula: &str, old_name: &str, new_name: &str) -> String {
    rewrite_sheet_names_in_formula_impl(formula, old_name, new_name, false)
}

/// Rewrite sheet references inside `formula` that refer to the current workbook only.
///
/// This is equivalent to [`rewrite_sheet_names_in_formula`]. It is retained as a separate helper
/// because some rewrite surfaces (e.g. sheet duplication) want to be explicit about not touching
/// external workbook references.
pub(crate) fn rewrite_sheet_names_in_formula_internal_refs_only(
    formula: &str,
    old_name: &str,
    new_name: &str,
) -> String {
    rewrite_sheet_names_in_formula_impl(formula, old_name, new_name, false)
}

fn rewrite_sheet_names_in_formula_impl(
    formula: &str,
    old_name: &str,
    new_name: &str,
    rewrite_external_workbooks: bool,
) -> String {
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
                if !rewrite_external_workbooks && split_workbook_prefix(&sheet_spec).0.is_some() {
                    out.push_str(raw);
                    i = next;
                    continue;
                }
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
            if !rewrite_external_workbooks && split_workbook_prefix(sheet_spec).0.is_some() {
                out.push_str(raw);
                i = next;
                continue;
            }
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

/// Rewrite sheet references inside `formula` when `deleted_sheet` is removed.
///
/// Excel treats direct references to a deleted sheet (`Sheet1!A1`) as `#REF!`.
/// For 3D references (`Sheet1:Sheet3!A1`), the span is adjusted using the sheet
/// order captured in `sheet_order`.
///
/// This routine is intentionally conservative: it only rewrites tokens that
/// parse as sheet references and it does not touch string literals.
///
/// This does **not** rewrite references that include an explicit workbook prefix
/// (e.g. `='[Book1.xlsx]Sheet1'!A1`), since those refer to an external workbook and
/// should not change when deleting a sheet inside the current workbook.
pub fn rewrite_deleted_sheet_references_in_formula(
    formula: &str,
    deleted_sheet: &str,
    sheet_order: &[String],
) -> String {
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
                // Ignore external workbook references.
                if split_workbook_prefix(&sheet_spec).0.is_some() {
                    out.push_str(raw);
                    i = next;
                    continue;
                }
                match rewrite_sheet_spec_for_delete(&sheet_spec, deleted_sheet, sheet_order) {
                    DeleteSheetSpecRewrite::Unchanged => {
                        out.push_str(raw);
                        i = next;
                        continue;
                    }
                    DeleteSheetSpecRewrite::Adjusted(rewritten) => {
                        out.push_str(&rewritten);
                        out.push('!');
                        i = next;
                        continue;
                    }
                    DeleteSheetSpecRewrite::Invalidate => {
                        out.push_str("#REF!");
                        i = sheet_ref_tail_end(formula, next);
                        continue;
                    }
                }
            }
        }

        if let Some((next, raw, sheet_spec)) = parse_unquoted_sheet_spec(formula, i) {
            // Ignore external workbook references.
            if split_workbook_prefix(sheet_spec).0.is_some() {
                out.push_str(raw);
                i = next;
                continue;
            }
            match rewrite_sheet_spec_for_delete(sheet_spec, deleted_sheet, sheet_order) {
                DeleteSheetSpecRewrite::Unchanged => {
                    out.push_str(raw);
                    i = next;
                    continue;
                }
                DeleteSheetSpecRewrite::Adjusted(rewritten) => {
                    out.push_str(&rewritten);
                    out.push('!');
                    i = next;
                    continue;
                }
                DeleteSheetSpecRewrite::Invalidate => {
                    out.push_str("#REF!");
                    i = sheet_ref_tail_end(formula, next);
                    continue;
                }
            }
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

fn is_name_char(b: u8) -> bool {
    b.is_ascii_alphanumeric() || b == b'_' || b == b'.'
}

fn bytes_eq_ignore_ascii_case(a: &[u8], b: &[u8]) -> bool {
    if a.len() != b.len() {
        return false;
    }
    a.iter()
        .zip(b.iter())
        .all(|(a, b)| a.to_ascii_lowercase() == b.to_ascii_lowercase())
}

/// Rewrite table names in structured references inside a formula.
///
/// This is intentionally conservative:
/// - It rewrites occurrences that look like a table token, either:
///   - a structured reference (`TableName[...]`), or
///   - a whole-table range reference (`TableName`).
/// - It does not touch string literals.
/// - It avoids rewriting sheet references (`Sheet1!A1`) and function calls (`Foo(...)`).
pub fn rewrite_table_names_in_formula(formula: &str, renames: &[(String, String)]) -> String {
    if renames.is_empty() {
        return formula.to_string();
    }

    fn is_external_sheet_qualified_prefix(formula: &str, bytes: &[u8], bang: usize) -> bool {
        // Determine whether the sheet reference immediately before `!` refers to an external
        // workbook (e.g. `[Book.xlsx]Sheet1!` or `'C:\path\[Book.xlsx]Sheet1'!`).
        if bang == 0 {
            return false;
        }

        // Quoted sheet reference: `'...'!`
        if bytes.get(bang.wrapping_sub(1)) == Some(&b'\'') {
            let end_quote = bang - 1;
            let mut i = end_quote;
            while i > 0 {
                i -= 1;
                if bytes[i] != b'\'' {
                    continue;
                }
                // Escaped quote inside a quoted sheet name is represented as `''`.
                if i > 0 && bytes[i - 1] == b'\'' {
                    i -= 1;
                    continue;
                }
                let start_quote = i;
                let token = &formula[start_quote + 1..end_quote];
                return crate::external_refs::find_external_workbook_prefix_span_in_sheet_spec(
                    token,
                )
                .is_some();
            }
            return false;
        }

        // Unquoted sheet reference: scan backwards over a conservative set of characters that can
        // appear in an unquoted external sheet key (`[Book.xlsx]Sheet1` / `Sheet1:Sheet3`).
        let mut start = bang;
        while start > 0 {
            let b = bytes[start - 1];
            let allowed = b.is_ascii_alphanumeric()
                || b == b'_'
                || b == b'.'
                || b == b'['
                || b == b']'
                || b == b':';
            if !allowed {
                break;
            }
            start -= 1;
        }
        let token = &formula[start..bang];
        crate::external_refs::find_external_workbook_prefix_span_in_sheet_spec(token).is_some()
    }

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

        // Table token:
        // - Structured reference: `TableName[...]`
        // - Whole-table reference: `TableName`
        //
        // We treat `.` and `_` as identifier characters, and we require the match
        // to be token-boundary aligned to avoid rewriting substrings.
        if i == 0 || !is_name_char(bytes[i - 1]) {
            let mut matched = false;
            for (old, new) in renames {
                let old_bytes = old.as_bytes();
                let Some(window) = bytes.get(i..i + old_bytes.len()) else {
                    continue;
                };
                if !bytes_eq_ignore_ascii_case(window, old_bytes) {
                    continue;
                }
                // Avoid rewriting external workbook structured references (e.g.
                // `"[Book.xlsx]Sheet1!Table1[Col]"` or `"[Book.xlsx]Table1[Col]"`) when renaming a
                // local table.
                //
                // Table names are workbook-scoped, so external workbook references should not be
                // impacted by local table renames.
                if i > 0 {
                    match bytes[i - 1] {
                        // Workbook-only structured ref prefix: `"[Book.xlsx]Table1[...]"`
                        b']' => continue,
                        // Sheet-qualified: `"...!Table1[...]"`.
                        b'!' if is_external_sheet_qualified_prefix(formula, bytes, i - 1) => {
                            continue;
                        }
                        _ => {}
                    }
                }
                // Ensure the match ends at an identifier boundary.
                match bytes.get(i + old_bytes.len()) {
                    Some(next) if is_name_char(*next) => continue,
                    // Avoid rewriting sheet references (`Sheet1!A1`).
                    Some(b'!') => continue,
                    // Avoid rewriting function calls (`Foo(...)`).
                    Some(b'(') => continue,
                    _ => {}
                }

                out.push_str(new);
                i += old_bytes.len();
                matched = true;
                break;
            }
            if matched {
                continue;
            }
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
        assert_eq!(format_sheet_reference(None, "Résumé", None), "'Résumé'");
        assert_eq!(format_sheet_reference(None, "数据", None), "'数据'");
        assert_eq!(format_sheet_reference(None, "TRUE", None), "'TRUE'");
        assert_eq!(format_sheet_reference(None, "R1C1", None), "'R1C1'");
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
    fn does_not_rewrite_external_workbook_reference() {
        assert_eq!(
            rewrite_sheet_names_in_formula("='[Book1.xlsx]Sheet1'!A1", "Sheet1", "Data"),
            "='[Book1.xlsx]Sheet1'!A1"
        );
    }

    #[test]
    fn does_not_rewrite_unquoted_external_workbook_reference() {
        assert_eq!(
            rewrite_sheet_names_in_formula("=[Book1.xlsx]Sheet1!A1", "Sheet1", "Data"),
            "=[Book1.xlsx]Sheet1!A1"
        );
    }

    #[test]
    fn does_not_rewrite_unquoted_external_workbook_reference_with_bracketed_path() {
        assert_eq!(
            rewrite_sheet_names_in_formula("=[C:\\[foo]\\Book1.xlsx]Sheet1!A1", "Sheet1", "Data"),
            "=[C:\\[foo]\\Book1.xlsx]Sheet1!A1"
        );
    }

    #[test]
    fn does_not_rewrite_external_workbook_reference_with_escaped_brackets() {
        assert_eq!(
            rewrite_sheet_names_in_formula("=[Book]]Name.xlsx]Sheet1!A1", "Sheet1", "Data"),
            "=[Book]]Name.xlsx]Sheet1!A1"
        );
    }

    #[test]
    fn does_not_rewrite_external_workbook_reference_with_nested_brackets_in_name() {
        // Workbook name: `Book[Name].xlsx` (note: the literal `]` must be escaped as `]]`).
        assert_eq!(
            rewrite_sheet_names_in_formula("=[Book[Name]].xlsx]Sheet1!A1", "Sheet1", "Data"),
            "=[Book[Name]].xlsx]Sheet1!A1"
        );
    }

    #[test]
    fn does_not_rewrite_external_reference_with_path() {
        assert_eq!(
            rewrite_sheet_names_in_formula("='C:\\path\\[Book1.xlsx]Sheet1'!A1", "Sheet1", "Data",),
            "='C:\\path\\[Book1.xlsx]Sheet1'!A1"
        );
    }

    #[test]
    fn does_not_rewrite_external_reference_with_brackets_in_path() {
        assert_eq!(
            rewrite_sheet_names_in_formula("='C:\\[foo]\\[Book1.xlsx]Sheet1'!A1", "Sheet1", "Data",),
            "='C:\\[foo]\\[Book1.xlsx]Sheet1'!A1"
        );
    }

    #[test]
    fn rewrite_structured_refs_for_renamed_table() {
        let renames = vec![("Table1".to_string(), "Table1_1".to_string())];
        assert_eq!(
            rewrite_table_names_in_formula("=SUM(Table1[Amount])", &renames),
            "=SUM(Table1_1[Amount])"
        );
    }

    #[test]
    fn rewrite_table_names_does_not_touch_external_workbook_structured_refs() {
        let renames = vec![("Table1".to_string(), "Sales".to_string())];
        assert_eq!(
            rewrite_table_names_in_formula("=SUM([Book.xlsx]Sheet1!Table1[Amount])", &renames),
            "=SUM([Book.xlsx]Sheet1!Table1[Amount])"
        );
        assert_eq!(
            rewrite_table_names_in_formula("=SUM('[Book.xlsx]Sheet1'!Table1[Amount])", &renames),
            "=SUM('[Book.xlsx]Sheet1'!Table1[Amount])"
        );
        assert_eq!(
            rewrite_table_names_in_formula(
                r"=SUM([C:\[foo]\Book.xlsx]Sheet1!Table1[Amount])",
                &renames
            ),
            r"=SUM([C:\[foo]\Book.xlsx]Sheet1!Table1[Amount])"
        );
        assert_eq!(
            rewrite_table_names_in_formula("=SUM([Book.xlsx]Table1[Amount])", &renames),
            "=SUM([Book.xlsx]Table1[Amount])"
        );
    }

    #[test]
    fn rewrite_table_name_token_without_brackets() {
        let renames = vec![("Table1".to_string(), "Table1_1".to_string())];
        assert_eq!(
            rewrite_table_names_in_formula("=SUM(Table1)", &renames),
            "=SUM(Table1_1)"
        );
    }

    #[test]
    fn rewrite_structured_refs_avoids_string_literals() {
        let renames = vec![("Table1".to_string(), "Table1_1".to_string())];
        assert_eq!(
            rewrite_table_names_in_formula("=\"Table1[Amount]\"", &renames),
            "=\"Table1[Amount]\""
        );
    }

    #[test]
    fn rewrite_structured_refs_does_not_match_substrings() {
        let renames = vec![("Table1".to_string(), "Table1_1".to_string())];
        assert_eq!(
            rewrite_table_names_in_formula("=SUM(MyTable1[Amount])", &renames),
            "=SUM(MyTable1[Amount])"
        );
    }

    #[test]
    fn rewrite_table_names_does_not_touch_sheet_references() {
        let renames = vec![("Sheet1".to_string(), "Sheet1_1".to_string())];
        assert_eq!(
            rewrite_table_names_in_formula("=Sheet1!A1", &renames),
            "=Sheet1!A1"
        );
    }

    #[test]
    fn rewrite_table_names_does_not_touch_function_calls() {
        let renames = vec![("Table1".to_string(), "Table1_1".to_string())];
        assert_eq!(
            rewrite_table_names_in_formula("=Table1(A1)", &renames),
            "=Table1(A1)"
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
    fn rewrite_quotes_unicode_sheet_names_in_output() {
        assert_eq!(
            rewrite_sheet_names_in_formula("=Sheet1!A1", "Sheet1", "数据"),
            "='数据'!A1"
        );
    }

    #[test]
    fn rewrite_quotes_reserved_sheet_names_in_output() {
        assert_eq!(
            rewrite_sheet_names_in_formula("=Sheet1!A1", "Sheet1", "TRUE"),
            "='TRUE'!A1"
        );
        assert_eq!(
            rewrite_sheet_names_in_formula("=Sheet1!A1", "Sheet1", "R1C1"),
            "='R1C1'!A1"
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

    #[test]
    fn delete_rewrites_simple_sheet_ref() {
        let order = vec!["Sheet1".to_string(), "Sheet2".to_string()];
        assert_eq!(
            rewrite_deleted_sheet_references_in_formula("=Sheet1!A1", "Sheet1", &order),
            "=#REF!"
        );
    }

    #[test]
    fn delete_adjusts_3d_boundary() {
        let order = vec![
            "Sheet1".to_string(),
            "Sheet2".to_string(),
            "Sheet3".to_string(),
        ];
        assert_eq!(
            rewrite_deleted_sheet_references_in_formula("=SUM(Sheet1:Sheet3!A1)", "Sheet1", &order),
            "=SUM(Sheet2:Sheet3!A1)"
        );
        assert_eq!(
            rewrite_deleted_sheet_references_in_formula("=SUM(Sheet1:Sheet3!A1)", "Sheet3", &order),
            "=SUM(Sheet1:Sheet2!A1)"
        );
    }

    #[test]
    fn delete_does_not_rewrite_external_workbook_refs() {
        let order = vec!["Sheet1".to_string(), "Sheet2".to_string()];
        assert_eq!(
            rewrite_deleted_sheet_references_in_formula(
                "='[Book.xlsx]Sheet1'!A1+Sheet1!A1",
                "Sheet1",
                &order
            ),
            "='[Book.xlsx]Sheet1'!A1+#REF!"
        );
        assert_eq!(
            rewrite_deleted_sheet_references_in_formula(
                "=[Book.xlsx]Sheet1!A1+Sheet1!A1",
                "Sheet1",
                &order
            ),
            "=[Book.xlsx]Sheet1!A1+#REF!"
        );
    }
}
