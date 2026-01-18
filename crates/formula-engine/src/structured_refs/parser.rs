use super::{StructuredColumn, StructuredColumns, StructuredRef, StructuredRefItem};
use formula_model::external_refs::unescape_bracketed_identifier_content;
use std::borrow::Cow;

fn scan_table_name_end(input: &str, pos: usize) -> Option<usize> {
    let bytes = input.as_bytes();
    if pos >= bytes.len() {
        return None;
    }

    // Fast path for the common ASCII-only table name case.
    //
    // If we encounter any non-ASCII bytes we fall back to the Unicode-capable scan below so
    // names like `Tâble1[Col]` continue to parse correctly.
    if bytes.get(pos).is_some_and(|b| b.is_ascii()) {
        let b0 = bytes[pos];
        if b0 != b'_' && !b0.is_ascii_alphabetic() {
            return None;
        }
        let mut i = pos + 1;
        while i < bytes.len() {
            let b = bytes[i];
            if b >= 0x80 {
                break;
            }
            if b == b'_' || b == b'.' || b.is_ascii_alphanumeric() {
                i += 1;
                continue;
            }
            return Some(i);
        }
        if i == bytes.len() || bytes[i] < 0x80 {
            return Some(i);
        }
    }

    let mut chars = input[pos..].char_indices();
    let (_, first) = chars.next()?;
    if !is_name_start(first) {
        return None;
    }
    let mut end = pos + first.len_utf8();
    for (off, ch) in chars {
        if !is_name_continue(ch) {
            break;
        }
        end = pos + off + ch.len_utf8();
    }
    Some(end)
}

fn parse_table_name_span_and_bracket_start(
    input: &str,
    pos: usize,
) -> Option<(Option<(usize, usize)>, usize)> {
    let bytes = input.as_bytes();
    if bytes.get(pos) == Some(&b'[') {
        return Some((None, pos));
    }

    let end = scan_table_name_end(input, pos)?;
    if bytes.get(end) != Some(&b'[') {
        return None;
    }
    Some((Some((pos, end)), end))
}

/// Find the end position of a structured reference starting at `pos`.
///
/// This is intended for **disambiguation** (e.g. lexer/parser bracket handling). It is designed to
/// be allocation-free and returns only the chosen end position (the byte index immediately after
/// the chosen closing `]`).
pub(crate) fn find_structured_ref_end(input: &str, pos: usize) -> Option<usize> {
    let (_, _, end_exclusive) = scan_structured_ref(input, pos)?;
    Some(end_exclusive)
}

/// Find the end position of a structured reference starting at `pos`, without validating the inner
/// spec.
///
/// This is intended for **parser recovery**: we still want to pick a stable closing `]` even when
/// the inner spec is invalid (so the parser can continue and later stages can surface Excel-like
/// errors). The `]]` ambiguity is still resolved by selecting the longest close position.
pub(crate) fn find_structured_ref_end_lenient(input: &str, pos: usize) -> Option<usize> {
    if pos >= input.len() {
        return None;
    }

    let (_, bracket_start) = parse_table_name_span_and_bracket_start(input, pos)?;
    let (_, end_exclusive) = parse_bracketed_lenient(input, bracket_start)?;
    Some(end_exclusive)
}

/// Structured-ref scan for disambiguation.
///
/// Returns:
/// - `table_name`: the optional table-name slice (if the ref is `Table1[...]`)
/// - `spec`: the inner spec slice (trimmed) inside the outermost brackets, excluding the brackets
/// - `end_exclusive`: the byte index immediately after the chosen closing `]`
///
/// This does **not** build a `StructuredRef`; it only validates that the syntax is acceptable and
/// resolves the `]]` ambiguity by selecting the longest valid end position.
pub(crate) fn scan_structured_ref<'a>(
    input: &'a str,
    pos: usize,
) -> Option<(Option<&'a str>, &'a str, usize)> {
    if pos >= input.len() {
        return None;
    }

    let (table_name_span, bracket_start) = parse_table_name_span_and_bracket_start(input, pos)?;
    let (inner, end_exclusive) = parse_bracketed_strict(input, bracket_start)?;

    let spec = inner.trim();
    let table_name = table_name_span.map(|(start, end)| &input[start..end]);
    Some((table_name, spec, end_exclusive))
}

pub(crate) fn parse_structured_ref_parts(
    table_name: Option<&str>,
    spec: &str,
) -> Option<StructuredRef> {
    // This helper is for callers that already have the table name and spec as separate strings
    // (e.g. from the eval AST). It avoids reconstructing `Table1[spec]` and re-scanning it.
    if let Some(name) = table_name {
        if !is_valid_table_name(name) {
            return None;
        }
    }

    parse_structured_ref_parts_unchecked(table_name, spec)
}

/// Like `parse_structured_ref_parts`, but assumes `table_name` (if present) is already validated.
///
/// Callers should only use this after a successful `scan_structured_ref` or when the table name
/// originates from a trusted source (e.g. table metadata).
pub(crate) fn parse_structured_ref_parts_unchecked(
    table_name: Option<&str>,
    spec: &str,
) -> Option<StructuredRef> {
    let (items, columns) = parse_inner_spec(spec)?;
    Some(StructuredRef {
        table_name: table_name.map(str::to_string),
        items,
        columns,
    })
}

pub fn parse_structured_ref(input: &str, pos: usize) -> Option<(StructuredRef, usize)> {
    let (table_name, spec, end_exclusive) = scan_structured_ref(input, pos)?;
    let sref = parse_structured_ref_parts_unchecked(table_name, spec)?;
    Some((sref, end_exclusive))
}

fn parse_bracketed_strict(input: &str, start: usize) -> Option<(&str, usize)> {
    parse_bracketed_select(input, start, &|inner| validate_inner_spec_raw(inner.trim()).is_some())
}

fn parse_bracketed_lenient(input: &str, start: usize) -> Option<(&str, usize)> {
    parse_bracketed_select(input, start, &|_inner| true)
}

fn parse_bracketed_select<'a, F>(
    input: &'a str,
    start: usize,
    accept: &F,
) -> Option<(&'a str, usize)>
where
    F: Fn(&str) -> bool,
{
    let bytes = input.as_bytes();
    if bytes.get(start) != Some(&b'[') {
        return None;
    }

    // Excel escapes `]` inside structured references as `]]`. Unfortunately `]]` is also used to
    // close nested bracket groups (e.g. `Table1[[Col]]` ends with `]]`).
    //
    // To disambiguate without relying on surrounding formula context, we treat `]]` as *either*
    // an escape (depth unchanged, consume 2 chars) or a close bracket (depth - 1, consume 1 char)
    // and then pick the longest end position that yields an acceptable inner spec.
    let input_len = bytes.len();

    // Fast path for the overwhelmingly common case where there are no ambiguous `]]` sequences:
    // do normal bracket matching and bail out to the state-machine only if `]]` appears.
    let mut depth: u32 = 1;
    let mut has_double_close = false;
    let mut pos = start + 1;
    while pos < input_len {
        match bytes[pos] {
            b'[' => {
                depth += 1;
                pos += 1;
            }
            b']' => {
                if bytes.get(pos + 1) == Some(&b']') {
                    has_double_close = true;
                    break;
                }
                depth = depth.saturating_sub(1);
                pos += 1;
                if depth == 0 {
                    let inner = &input[start + 1..pos - 1];
                    return accept(inner).then(|| (inner, pos));
                }
            }
            _ => {
                pos += 1;
            }
        }
    }
    if !has_double_close {
        return None;
    }

    parse_bracketed_with_double_close_select(input, start, accept)
}

fn parse_bracketed_with_double_close_select<'a, F>(
    input: &'a str,
    start: usize,
    accept: &F,
) -> Option<(&'a str, usize)>
where
    F: Fn(&str) -> bool,
{
    let bytes = input.as_bytes();
    let input_len = bytes.len();

    // State machine indexed relative to `base`.
    //
    // We grow the state vectors on demand rather than preallocating to `input.len()`. In valid
    // inputs, structured refs tend to be short compared to the full formula, and we can stop once
    // no states remain reachable.
    let base = start + 1;
    if base > input_len {
        return None;
    }

    let mut states: Vec<Vec<u32>> = vec![Vec::new()];
    let mut can_close: Vec<bool> = vec![false];
    push_depth(&mut states, &mut can_close, 0, 1);

    let mut abs_pos = base;
    while abs_pos < input_len {
        let rel_pos = abs_pos - base;
        if rel_pos >= states.len() {
            break;
        }
        let depths = std::mem::take(&mut states[rel_pos]);
        if depths.is_empty() {
            abs_pos += 1;
            continue;
        }
        for depth in depths {
            if depth == 0 {
                continue;
            }
            match bytes[abs_pos] {
                b'[' => push_depth(&mut states, &mut can_close, rel_pos + 1, depth + 1),
                b']' => {
                    // Treat as closing bracket.
                    push_depth(
                        &mut states,
                        &mut can_close,
                        rel_pos + 1,
                        depth.saturating_sub(1),
                    );
                    // Treat as escaped literal `]`.
                    if bytes.get(abs_pos + 1) == Some(&b']') {
                        push_depth(&mut states, &mut can_close, rel_pos + 2, depth);
                    }
                }
                _ => push_depth(&mut states, &mut can_close, rel_pos + 1, depth),
            }
        }
        abs_pos += 1;
    }

    let max_end_pos = (base + can_close.len().saturating_sub(1)).min(input_len);
    for abs_end_pos in (base..=max_end_pos).rev() {
        let rel_end_pos = abs_end_pos - base;
        if rel_end_pos >= can_close.len() || !can_close[rel_end_pos] {
            continue;
        }
        let inner = &input[base..abs_end_pos.saturating_sub(1)];
        if accept(inner) {
            return Some((inner, abs_end_pos));
        }
    }

    None
}

fn push_depth(states: &mut Vec<Vec<u32>>, can_close: &mut Vec<bool>, pos: usize, depth: u32) {
    if pos >= states.len() {
        states.resize_with(pos + 1, Vec::new);
        can_close.resize(pos + 1, false);
    }
    debug_assert_eq!(states.len(), can_close.len());
    if depth == 0 {
        can_close[pos] = true;
        return;
    }
    let entry = &mut states[pos];
    if !entry.contains(&depth) {
        entry.push(depth);
    }
}

#[derive(Debug, Clone, PartialEq)]
enum StructuredColumnRaw<'a> {
    Single(Cow<'a, str>),
    Range {
        start: Cow<'a, str>,
        end: Cow<'a, str>,
    },
}

#[derive(Debug, Clone, PartialEq)]
enum StructuredColumnsRaw<'a> {
    All,
    Single(Cow<'a, str>),
    Range {
        start: Cow<'a, str>,
        end: Cow<'a, str>,
    },
    Multi(Vec<StructuredColumnRaw<'a>>),
}

fn columns_raw_into_owned(columns: StructuredColumnsRaw<'_>) -> Option<StructuredColumns> {
    match columns {
        StructuredColumnsRaw::All => Some(StructuredColumns::All),
        StructuredColumnsRaw::Single(name) => Some(StructuredColumns::Single(name.into_owned())),
        StructuredColumnsRaw::Range { start, end } => Some(StructuredColumns::Range {
            start: start.into_owned(),
            end: end.into_owned(),
        }),
        StructuredColumnsRaw::Multi(parts) => {
            let mut out: Vec<StructuredColumn> = Vec::new();
            if out.try_reserve_exact(parts.len()).is_err() {
                debug_assert!(
                    false,
                    "allocation failed (structured columns, len={})",
                    parts.len()
                );
                return None;
            }
            for part in parts {
                out.push(match part {
                    StructuredColumnRaw::Single(name) => StructuredColumn::Single(name.into_owned()),
                    StructuredColumnRaw::Range { start, end } => StructuredColumn::Range {
                        start: start.into_owned(),
                        end: end.into_owned(),
                    },
                });
            }
            Some(StructuredColumns::Multi(out))
        }
    }
}

fn parse_inner_spec(inner: &str) -> Option<(Vec<StructuredRefItem>, StructuredColumns)> {
    let (items, columns) = parse_inner_spec_raw(inner)?;
    let columns = columns_raw_into_owned(columns)?;
    Some((items, columns))
}

fn validate_inner_spec_raw(inner: &str) -> Option<()> {
    if inner.is_empty() {
        return None;
    }

    // Nested form like "[[#Headers],[Column]]" or "[[Col1]:[Col2]]"
    if inner.trim_start().starts_with('[') {
        let saw_part = for_each_top_level_part(inner, ',', |part| {
            validate_bracket_group_or_range(part)?;
            Some(())
        })?;
        if !saw_part {
            return None;
        }
        return Some(());
    }

    // Simple form: "@Col", "Col", "#Headers", etc.
    let trimmed = inner.trim();
    if let Some(stripped) = trimmed.strip_prefix('@') {
        let stripped = stripped.trim();
        if stripped.is_empty() {
            return Some(());
        }
        if stripped.starts_with('[') {
            return validate_columns_only(stripped);
        }
        return Some(());
    }

    if parse_item(trimmed).is_some() {
        return Some(());
    }

    Some(())
}

fn parse_inner_spec_raw(inner: &str) -> Option<(Vec<StructuredRefItem>, StructuredColumnsRaw<'_>)> {
    if inner.is_empty() {
        return None;
    }

    // Nested form like "[[#Headers],[Column]]" or "[[Col1]:[Col2]]"
    if inner.trim_start().starts_with('[') {
        let mut items: Vec<StructuredRefItem> = Vec::new();
        let mut columns: Vec<StructuredColumnRaw<'_>> = Vec::new();
        let saw_part = for_each_top_level_part(inner, ',', |part| {
            let (maybe_item, cols) = parse_bracket_group_or_range(part)?;
            if let Some(it) = maybe_item {
                items.push(it);
                return Some(());
            }

            match cols {
                StructuredColumnsRaw::Single(name) => {
                    columns.push(StructuredColumnRaw::Single(name));
                }
                StructuredColumnsRaw::Range { start, end } => {
                    columns.push(StructuredColumnRaw::Range { start, end });
                }
                StructuredColumnsRaw::All => {
                    return None;
                }
                StructuredColumnsRaw::Multi(_) => {
                    return None;
                }
            }
            Some(())
        })?;
        if !saw_part {
            return None;
        }

        let cols = match columns.len() {
            0 => StructuredColumnsRaw::All,
            1 => match columns.pop()? {
                StructuredColumnRaw::Single(name) => StructuredColumnsRaw::Single(name),
                StructuredColumnRaw::Range { start, end } => StructuredColumnsRaw::Range { start, end },
            },
            _ => StructuredColumnsRaw::Multi(columns),
        };

        Some((items, cols))
    } else {
        // Simple form: "@Col", "Col", "#Headers", etc.
        let trimmed = inner.trim();
        if let Some(stripped) = trimmed.strip_prefix('@') {
            let stripped = stripped.trim();
            if stripped.is_empty() {
                return Some((vec![StructuredRefItem::ThisRow], StructuredColumnsRaw::All));
            }
            if stripped.starts_with('[') {
                let cols = parse_columns_only(stripped)?;
                return Some((vec![StructuredRefItem::ThisRow], cols));
            }
            return Some((
                vec![StructuredRefItem::ThisRow],
                StructuredColumnsRaw::Single(unescape_column_name(stripped)),
            ));
        }

        if let Some(item) = parse_item(trimmed) {
            return Some((vec![item], StructuredColumnsRaw::All));
        }

        Some((
            Vec::new(),
            StructuredColumnsRaw::Single(unescape_column_name(trimmed)),
        ))
    }
}

fn validate_columns_only(spec: &str) -> Option<()> {
    let spec = spec.trim();
    if spec.is_empty() {
        return Some(());
    }
    if !spec.starts_with('[') {
        return Some(());
    }

    let saw_part = for_each_top_level_part(spec, ',', |part| {
        if let Some((left, right)) = split_top_level_once(part, ':') {
            strip_wrapping_brackets(left)?;
            strip_wrapping_brackets(right)?;
            return Some(());
        }

        strip_wrapping_brackets(part)?;
        Some(())
    })?;
    if !saw_part {
        return None;
    }
    Some(())
}

fn parse_columns_only(spec: &str) -> Option<StructuredColumnsRaw<'_>> {
    let spec = spec.trim();
    if spec.is_empty() {
        return Some(StructuredColumnsRaw::All);
    }
    if !spec.starts_with('[') {
        return Some(StructuredColumnsRaw::Single(unescape_column_name(spec)));
    }

    let mut columns: Vec<StructuredColumnRaw<'_>> = Vec::new();
    let saw_part = for_each_top_level_part(spec, ',', |part| {
        if let Some((left, right)) = split_top_level_once(part, ':') {
            let start = strip_wrapping_brackets(left)?;
            let end = strip_wrapping_brackets(right)?;
            columns.push(StructuredColumnRaw::Range {
                start: unescape_column_name(start),
                end: unescape_column_name(end),
            });
            return Some(());
        }

        let inner = strip_wrapping_brackets(part)?;
        columns.push(StructuredColumnRaw::Single(unescape_column_name(inner)));
        Some(())
    })?;
    if !saw_part {
        return None;
    }

    let cols = match columns.len() {
        0 => StructuredColumnsRaw::All,
        1 => match columns.pop()? {
            StructuredColumnRaw::Single(name) => StructuredColumnsRaw::Single(name),
            StructuredColumnRaw::Range { start, end } => StructuredColumnsRaw::Range { start, end },
        },
        _ => StructuredColumnsRaw::Multi(columns),
    };

    Some(cols)
}

fn validate_bracket_group_or_range(part: &str) -> Option<()> {
    let part = part.trim();
    if part.is_empty() {
        return None;
    }

    // A range is something like "[Col1]:[Col2]" (top-level colon, not inside brackets).
    if let Some((left, right)) = split_top_level_once(part, ':') {
        strip_wrapping_brackets(left)?;
        strip_wrapping_brackets(right)?;
        return Some(());
    }

    strip_wrapping_brackets(part)?;
    Some(())
}

fn parse_bracket_group_or_range(
    part: &str,
) -> Option<(Option<StructuredRefItem>, StructuredColumnsRaw<'_>)> {
    let part = part.trim();
    if part.is_empty() {
        return None;
    }

    // A range is something like "[Col1]:[Col2]" (top-level colon, not inside brackets).
    if let Some((left, right)) = split_top_level_once(part, ':') {
        let start = strip_wrapping_brackets(left)?;
        let end = strip_wrapping_brackets(right)?;
        return Some((
            None,
            StructuredColumnsRaw::Range {
                start: unescape_column_name(start),
                end: unescape_column_name(end),
            },
        ));
    }

    let inner = strip_wrapping_brackets(part)?;
    if let Some(item) = parse_item(inner) {
        return Some((Some(item), StructuredColumnsRaw::All));
    }
    Some((None, StructuredColumnsRaw::Single(unescape_column_name(inner))))
}

fn for_each_top_level_part<'a>(
    input: &'a str,
    delimiter: char,
    mut f: impl FnMut(&'a str) -> Option<()>,
) -> Option<bool> {
    debug_assert!(delimiter.is_ascii(), "delimiter must be ASCII");
    let bytes = input.as_bytes();
    let delim = delimiter as u8;

    let mut start = 0usize;
    let mut any = false;
    for_each_top_level_delimiter_pos(bytes, delim, |i| {
        let part = input[start..i].trim();
        if !part.is_empty() {
            f(part)?;
            any = true;
        }
        start = i + 1;
        Some(())
    })?;
    let tail = input[start..].trim();
    if !tail.is_empty() {
        f(tail)?;
        any = true;
    }
    Some(any)
}

fn split_top_level_once(input: &str, delimiter: char) -> Option<(&str, &str)> {
    debug_assert!(delimiter.is_ascii(), "delimiter must be ASCII");
    let bytes = input.as_bytes();
    let delim = delimiter as u8;

    let mut split_at: Option<usize> = None;
    for_each_top_level_delimiter_pos(bytes, delim, |i| {
        if split_at.is_some() {
            return None;
        }
        split_at = Some(i);
        Some(())
    })?;

    let split_at = split_at?;
    Some((input[..split_at].trim(), input[split_at + 1..].trim()))
}

fn for_each_top_level_delimiter_pos(
    bytes: &[u8],
    delim: u8,
    mut f: impl FnMut(usize) -> Option<()>,
) -> Option<()> {
    let mut depth: u32 = 0;
    let mut i = 0usize;
    while i < bytes.len() {
        match bytes[i] {
            b'[' => {
                depth += 1;
                i += 1;
            }
            b']' => {
                // Excel escapes ']' inside structured references as ']]'. At the current nesting
                // depth, treat a double ']]' as a literal ']' so we don't break bracket matching
                // while splitting.
                if depth == 1 && bytes.get(i + 1) == Some(&b']') {
                    i += 2;
                    continue;
                }
                depth = depth.saturating_sub(1);
                i += 1;
            }
            b if depth == 0 && b == delim => {
                f(i)?;
                i += 1;
            }
            _ => {
                i += 1;
            }
        }
    }
    Some(())
}

fn strip_wrapping_brackets(s: &str) -> Option<&str> {
    let s = s.trim();
    if s.starts_with('[') && s.ends_with(']') && s.len() >= 2 {
        Some(&s[1..s.len() - 1])
    } else {
        None
    }
}

fn parse_item(item: &str) -> Option<StructuredRefItem> {
    fn matches_item_ignoring_ws_ascii_case(input: &str, expected_lower: &[u8]) -> bool {
        if input.is_ascii() {
            let mut i = 0usize;
            for &b in input.as_bytes() {
                if b.is_ascii_whitespace() {
                    continue;
                }
                let lower = b.to_ascii_lowercase();
                if expected_lower.get(i) != Some(&lower) {
                    return false;
                }
                i += 1;
            }
            return i == expected_lower.len();
        }

        let mut i = 0usize;
        for ch in input.chars() {
            if ch.is_whitespace() {
                continue;
            }

            let lower = ch.to_ascii_lowercase();
            if !lower.is_ascii() {
                return false;
            }
            let b = lower as u8;
            if expected_lower.get(i) != Some(&b) {
                return false;
            }
            i += 1;
        }
        i == expected_lower.len()
    }

    let item = item.trim();
    let item = item.strip_prefix('#').unwrap_or(item).trim();
    if matches_item_ignoring_ws_ascii_case(item, b"all") {
        Some(StructuredRefItem::All)
    } else if matches_item_ignoring_ws_ascii_case(item, b"data") {
        Some(StructuredRefItem::Data)
    } else if matches_item_ignoring_ws_ascii_case(item, b"headers") {
        Some(StructuredRefItem::Headers)
    } else if matches_item_ignoring_ws_ascii_case(item, b"totals") {
        Some(StructuredRefItem::Totals)
    } else if matches_item_ignoring_ws_ascii_case(item, b"thisrow") {
        Some(StructuredRefItem::ThisRow)
    } else {
        None
    }
}

fn unescape_column_name(name: &str) -> Cow<'_, str> {
    // Excel escapes ']' as ']]' in structured references.
    unescape_bracketed_identifier_content(name.trim())
}

fn is_name_start(ch: char) -> bool {
    ch == '_' || (!ch.is_ascii() && ch.is_alphabetic()) || ch.is_ascii_alphabetic()
}

fn is_name_continue(ch: char) -> bool {
    ch == '_' || ch == '.' || (!ch.is_ascii() && ch.is_alphanumeric()) || ch.is_ascii_alphanumeric()
}

fn is_valid_table_name(name: &str) -> bool {
    scan_table_name_end(name, 0).is_some_and(|end| end == name.len())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::structured_refs::{StructuredColumn, StructuredColumns, StructuredRefItem};

    #[test]
    fn for_each_top_level_part_reports_any() {
        let any = for_each_top_level_part(" ,  , ", ',', |_part| Some(()));
        assert_eq!(any, Some(false));

        let mut seen = Vec::new();
        let any = for_each_top_level_part(" a , , b ", ',', |part| {
            seen.push(part.to_string());
            Some(())
        });
        assert_eq!(any, Some(true));
        assert_eq!(seen, vec!["a", "b"]);
    }

    #[test]
    fn for_each_top_level_part_ignores_delims_inside_brackets_and_escapes() {
        let mut seen = Vec::new();
        let any = for_each_top_level_part("[a,b],[A]]B],c", ',', |part| {
            seen.push(part.to_string());
            Some(())
        });
        assert_eq!(any, Some(true));
        assert_eq!(seen, vec!["[a,b]", "[A]]B]", "c"]);
    }

    #[test]
    fn for_each_top_level_part_handles_nested_brackets() {
        let mut seen = Vec::new();
        let any = for_each_top_level_part("[[A,B],C],D", ',', |part| {
            seen.push(part.to_string());
            Some(())
        });
        assert_eq!(any, Some(true));
        assert_eq!(seen, vec!["[[A,B],C]", "D"]);
    }

    #[test]
    fn for_each_top_level_part_handles_stray_rbracket() {
        let mut seen = Vec::new();
        let any = for_each_top_level_part("A],B", ',', |part| {
            seen.push(part.to_string());
            Some(())
        });
        assert_eq!(any, Some(true));
        assert_eq!(seen, vec!["A]", "B"]);
    }

    #[test]
    fn split_top_level_once_handles_stray_rbracket() {
        // Malformed inputs can contain stray `]` characters; delimiter scanning should remain
        // robust and not underflow its depth tracking.
        assert_eq!(split_top_level_once("A]:B", ':'), Some(("A]", "B")));
    }

    #[test]
    fn split_top_level_once_ignores_delims_inside_brackets() {
        assert_eq!(
            split_top_level_once("[A:B]:[C]", ':'),
            Some(("[A:B]", "[C]"))
        );
    }

    #[test]
    fn split_top_level_once_rejects_multiple_top_level_delims() {
        assert_eq!(split_top_level_once("A:B:C", ':'), None);
    }

    #[test]
    fn split_top_level_once_treats_escaped_rbracket_as_literal() {
        // `]]` inside bracket depth 1 is a literal `]` and should not break depth tracking.
        assert_eq!(
            split_top_level_once("[A]]B]:C", ':'),
            Some(("[A]]B]", "C"))
        );
    }

    #[test]
    fn parses_non_ascii_table_name() {
        let input = "Tâble1[Qty]";
        let (sref, end) = parse_structured_ref(input, 0).unwrap();
        assert_eq!(end, input.len());
        assert_eq!(sref.table_name.as_deref(), Some("Tâble1"));
        assert_eq!(sref.items, Vec::<StructuredRefItem>::new());
        assert_eq!(sref.columns, StructuredColumns::Single("Qty".into()));
    }

    #[test]
    fn parses_ascii_table_name_with_underscore_and_dot() {
        let input = "Table_1.2[Qty]";
        let (sref, end) = parse_structured_ref(input, 0).unwrap();
        assert_eq!(end, input.len());
        assert_eq!(sref.table_name.as_deref(), Some("Table_1.2"));
        assert_eq!(sref.items, Vec::<StructuredRefItem>::new());
        assert_eq!(sref.columns, StructuredColumns::Single("Qty".into()));
    }

    #[test]
    fn find_structured_ref_end_matches_parse_structured_ref_end() {
        for input in [
            "Table1[Qty]",
            "Table1[[#Headers],[A]]B]]",
            "[@[Col1]:[Col3]]",
            "Table1[A]]B]",
        ] {
            let end = find_structured_ref_end(input, 0).unwrap();
            let (_, parse_end) = parse_structured_ref(input, 0).unwrap();
            assert_eq!(end, parse_end, "input={input}");
        }
    }

    #[test]
    fn find_structured_ref_end_rejects_invalid_specs() {
        assert!(find_structured_ref_end("Table1[]", 0).is_none());
        assert!(find_structured_ref_end("Table1[Qty", 0).is_none());
        assert!(find_structured_ref_end("1Table[Qty]", 0).is_none());
    }

    #[test]
    fn find_structured_ref_end_lenient_accepts_invalid_specs() {
        let input = "Table1[]";
        assert_eq!(find_structured_ref_end_lenient(input, 0), Some(input.len()));
        assert!(find_structured_ref_end_lenient("Table1[Qty", 0).is_none());
        assert!(find_structured_ref_end_lenient("1Table[Qty]", 0).is_none());
    }

    #[test]
    fn scan_structured_ref_returns_slices_and_end_pos() {
        let input = "Table1[Qty]+1";
        let (table, spec, end) = scan_structured_ref(input, 0).unwrap();
        assert_eq!(table, Some("Table1"));
        assert_eq!(spec, "Qty");
        assert_eq!(end, "Table1[Qty]".len());
    }

    #[test]
    fn scan_structured_ref_handles_escaped_rbracket_column_name() {
        let input = "Table1[A]]B]+1";
        let (table, spec, end) = scan_structured_ref(input, 0).unwrap();
        assert_eq!(table, Some("Table1"));
        assert_eq!(spec, "A]]B");
        assert_eq!(end, "Table1[A]]B]".len());
    }

    #[test]
    fn scan_structured_ref_handles_this_row_range_when_embedded_in_formula() {
        let input = "SUM([@[Col1]:[Col3]])";
        let start = "SUM(".len();
        let (table, spec, end) = scan_structured_ref(input, start).unwrap();
        assert_eq!(table, None);
        assert_eq!(spec, "@[Col1]:[Col3]");
        assert_eq!(end, start + "[@[Col1]:[Col3]]".len());
    }

    #[test]
    fn scan_structured_ref_trims_outer_whitespace() {
        let input = "Table1[  Qty \t]";
        let (table, spec, end) = scan_structured_ref(input, 0).unwrap();
        assert_eq!(table, Some("Table1"));
        assert_eq!(spec, "Qty");
        assert_eq!(end, input.len());
    }

    #[test]
    fn scan_structured_ref_reports_nested_and_multi_column_specs() {
        let input = "Table1[[#Headers],[Col1],[Col3]]";
        let (table, spec, end) = scan_structured_ref(input, 0).unwrap();
        assert_eq!(table, Some("Table1"));
        assert_eq!(spec, "[#Headers],[Col1],[Col3]");
        assert_eq!(end, input.len());
    }

    #[test]
    fn scan_structured_ref_reports_simple_this_row_spec() {
        let input = "[@Qty]";
        let (table, spec, end) = scan_structured_ref(input, 0).unwrap();
        assert_eq!(table, None);
        assert_eq!(spec, "@Qty");
        assert_eq!(end, input.len());
    }

    #[test]
    fn scan_structured_ref_rejects_invalid() {
        assert!(scan_structured_ref("Table1[]", 0).is_none());
        assert!(scan_structured_ref("Table1[Qty", 0).is_none());
        assert!(scan_structured_ref("1Table[Qty]", 0).is_none());
    }

    #[test]
    fn validate_inner_spec_raw_matches_parse_inner_spec_for_examples() {
        for spec in [
            "Qty",
            "@Qty",
            "@",
            "#Headers",
            "#\u{00A0}Headers",
            "[#Headers],[Col1],[Col3]",
            "[Col1]:[Col3]",
            "@[Col1]:[Col3]",
            "[[@Col1]:[Col3]]",
            "[]",
            "[#Headers],[A]]B]",
        ] {
            let validated = validate_inner_spec_raw(spec).is_some();
            let parsed = parse_inner_spec(spec).is_some();
            assert_eq!(validated, parsed, "spec={spec:?}");
        }
    }

    #[test]
    fn parse_structured_ref_parts_matches_parse_structured_ref() {
        fn parse_parts_from_full(input: &str) -> StructuredRef {
            let (table, spec, end) = scan_structured_ref(input, 0).unwrap();
            assert_eq!(end, input.len());
            parse_structured_ref_parts(table, spec).unwrap()
        }

        for input in [
            "Table1[Qty]",
            "Table1[[#Headers],[A]]B]]",
            "Table_1.2[Qty]",
            "Tâble1[Qty]",
            "[@[Col1]:[Col3]]",
            "A1.[Field]",
            "Table1[[#\u{00A0}Headers],[Qty]]",
        ] {
            let (full, end) = parse_structured_ref(input, 0).unwrap();
            assert_eq!(end, input.len());
            let parts = parse_parts_from_full(input);
            assert_eq!(parts, full, "input={input}");
        }
    }

    #[test]
    fn parse_structured_ref_parts_rejects_invalid_table_name() {
        assert!(parse_structured_ref_parts(Some("1Table"), "Qty").is_none());
        assert!(parse_structured_ref_parts(Some("Table-1"), "Qty").is_none());
        assert!(parse_structured_ref_parts(Some(""), "Qty").is_none());
    }

    #[test]
    fn rejects_table_name_starting_with_digit() {
        assert!(parse_structured_ref("1Table[Qty]", 0).is_none());
    }

    #[test]
    fn parses_table_column_ref() {
        let (sref, end) = parse_structured_ref("Table1[Qty]", 0).unwrap();
        assert_eq!(end, "Table1[Qty]".len());
        assert_eq!(sref.table_name.as_deref(), Some("Table1"));
        assert_eq!(sref.items, Vec::<StructuredRefItem>::new());
        assert_eq!(sref.columns, StructuredColumns::Single("Qty".into()));
    }

    #[test]
    fn parses_simple_column_name_with_escaped_rbracket() {
        let (sref, end) = parse_structured_ref("Table1[A]]B]", 0).unwrap();
        assert_eq!(end, "Table1[A]]B]".len());
        assert_eq!(sref.table_name.as_deref(), Some("Table1"));
        assert_eq!(sref.items, Vec::<StructuredRefItem>::new());
        assert_eq!(sref.columns, StructuredColumns::Single("A]B".into()));
    }

    #[test]
    fn rejects_empty_bracket_payload() {
        assert!(parse_structured_ref("Table1[]", 0).is_none());
    }

    #[test]
    fn rejects_unterminated_bracket_payload() {
        assert!(parse_structured_ref("Table1[Qty", 0).is_none());
    }

    #[test]
    fn parses_simple_structured_ref_prefix_when_followed_by_more_text() {
        let input = "Table1[Qty]+1";
        let (sref, end) = parse_structured_ref(input, 0).unwrap();
        assert_eq!(end, "Table1[Qty]".len());
        assert_eq!(sref.table_name.as_deref(), Some("Table1"));
        assert_eq!(sref.items, Vec::<StructuredRefItem>::new());
        assert_eq!(sref.columns, StructuredColumns::Single("Qty".into()));
    }

    #[test]
    fn parses_this_row_ref() {
        let (sref, _) = parse_structured_ref("[@Qty]", 0).unwrap();
        assert_eq!(sref.table_name, None);
        assert_eq!(sref.items, vec![StructuredRefItem::ThisRow]);
        assert_eq!(sref.columns, StructuredColumns::Single("Qty".into()));
    }

    #[test]
    fn parses_headers_ref() {
        let (sref, _) = parse_structured_ref("Table1[[#Headers],[Qty]]", 0).unwrap();
        assert_eq!(sref.items, vec![StructuredRefItem::Headers]);
        assert_eq!(sref.columns, StructuredColumns::Single("Qty".into()));
    }

    #[test]
    fn parses_headers_ref_with_unicode_whitespace() {
        // Ensure `# Headers` (with non-ASCII whitespace) still parses by taking the Unicode
        // whitespace fallback path.
        let (sref, _) = parse_structured_ref("Table1[[#\u{00A0}Headers],[Qty]]", 0).unwrap();
        assert_eq!(sref.items, vec![StructuredRefItem::Headers]);
        assert_eq!(sref.columns, StructuredColumns::Single("Qty".into()));
    }

    #[test]
    fn parses_nested_column_name_with_escaped_bracket() {
        let (sref, _) = parse_structured_ref("Table1[[#Headers],[A]]B]]", 0).unwrap();
        assert_eq!(sref.table_name.as_deref(), Some("Table1"));
        assert_eq!(sref.items, vec![StructuredRefItem::Headers]);
        assert_eq!(sref.columns, StructuredColumns::Single("A]B".into()));
    }

    #[test]
    fn parses_structured_ref_prefix_when_followed_by_another_ref() {
        let input = "Table1[[#Headers],[Qty]]+Table1[Qty]";
        let first = "Table1[[#Headers],[Qty]]";

        let (sref, end) = parse_structured_ref(input, 0).unwrap();
        assert_eq!(end, first.len());
        assert_eq!(sref.table_name.as_deref(), Some("Table1"));
        assert_eq!(sref.items, vec![StructuredRefItem::Headers]);
        assert_eq!(sref.columns, StructuredColumns::Single("Qty".into()));
    }

    #[test]
    fn parses_structured_ref_prefix_with_escaped_rbracket_column() {
        let input = "Table1[A]]B]+Table1[Qty]";
        let first = "Table1[A]]B]";

        let (sref, end) = parse_structured_ref(input, 0).unwrap();
        assert_eq!(end, first.len());
        assert_eq!(sref.table_name.as_deref(), Some("Table1"));
        assert_eq!(sref.items, Vec::<StructuredRefItem>::new());
        assert_eq!(sref.columns, StructuredColumns::Single("A]B".into()));
    }

    #[test]
    fn parses_multi_column_ref() {
        let (sref, _) = parse_structured_ref("Table1[[Col1],[Col3]]", 0).unwrap();
        assert_eq!(sref.items, Vec::<StructuredRefItem>::new());
        assert_eq!(
            sref.columns,
            StructuredColumns::Multi(vec![
                StructuredColumn::Single("Col1".into()),
                StructuredColumn::Single("Col3".into()),
            ])
        );
    }

    #[test]
    fn parses_this_row_column_range_ref() {
        let (sref, _) = parse_structured_ref("[@[Col1]:[Col3]]", 0).unwrap();
        assert_eq!(sref.items, vec![StructuredRefItem::ThisRow]);
        assert_eq!(
            sref.columns,
            StructuredColumns::Range {
                start: "Col1".into(),
                end: "Col3".into(),
            }
        );
    }

    #[test]
    fn parses_this_row_column_range_ref_when_embedded_in_formula() {
        let input = "SUM([@[Col1]:[Col3]])";
        let start = "SUM(".len();
        let (sref, end) = parse_structured_ref(input, start).unwrap();
        assert_eq!(end, start + "[@[Col1]:[Col3]]".len());
        assert_eq!(sref.items, vec![StructuredRefItem::ThisRow]);
        assert_eq!(
            sref.columns,
            StructuredColumns::Range {
                start: "Col1".into(),
                end: "Col3".into(),
            }
        );
    }
}
