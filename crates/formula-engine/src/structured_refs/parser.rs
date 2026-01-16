use super::{StructuredColumn, StructuredColumns, StructuredRef, StructuredRefItem};
use formula_model::external_refs::unescape_bracketed_identifier_content;
use std::borrow::Cow;

pub fn parse_structured_ref(input: &str, pos: usize) -> Option<(StructuredRef, usize)> {
    if pos >= input.len() {
        return None;
    }

    let bytes = input.as_bytes();
    let (table_name_span, bracket_start) = if bytes[pos] == b'[' {
        (None, pos)
    } else {
        let mut i = pos;
        let mut chars = input[pos..].char_indices();
        let (_, first) = chars.next()?;
        if !is_name_start(first) {
            return None;
        }
        i += first.len_utf8();
        while i < input.len() {
            let ch = input[i..].chars().next()?;
            if is_name_continue(ch) {
                i += ch.len_utf8();
            } else {
                break;
            }
        }
        if i >= input.len() || input.as_bytes()[i] != b'[' {
            return None;
        }
        (Some((pos, i)), i)
    };

    let (inner, end_pos) = parse_bracketed(input, bracket_start)?;
    let (items, columns) = parse_inner_spec(inner.trim())?;
    let table_name = table_name_span.map(|(start, end)| input[start..end].to_string());

    Some((
        StructuredRef {
            table_name,
            items,
            columns,
        },
        end_pos,
    ))
}

fn parse_bracketed(input: &str, start: usize) -> Option<(&str, usize)> {
    let bytes = input.as_bytes();
    if bytes.get(start) != Some(&b'[') {
        return None;
    }

    // Excel escapes `]` inside structured references as `]]`. Unfortunately `]]` is also used to
    // close nested bracket groups (e.g. `Table1[[Col]]` ends with `]]`).
    //
    // To disambiguate without relying on surrounding formula context, we treat `]]` as *either*
    // an escape (depth unchanged, consume 2 chars) or a close bracket (depth - 1, consume 1 char)
    // and then pick the longest end position that yields a syntactically valid inner spec.
    let len = bytes.len();
    if start + 1 > len {
        return None;
    }

    let mut states: Vec<Vec<u32>> = vec![Vec::new(); len + 1];
    let mut can_close: Vec<bool> = vec![false; len + 1];
    push_depth(&mut states, &mut can_close, start + 1, 1);

    for pos in start + 1..len {
        let depths = std::mem::take(&mut states[pos]);
        if depths.is_empty() {
            continue;
        }
        for depth in depths {
            if depth == 0 {
                continue;
            }
            match bytes[pos] {
                b'[' => push_depth(&mut states, &mut can_close, pos + 1, depth + 1),
                b']' => {
                    // Treat as closing bracket.
                    push_depth(&mut states, &mut can_close, pos + 1, depth.saturating_sub(1));
                    // Treat as escaped literal `]`.
                    if bytes.get(pos + 1) == Some(&b']') {
                        push_depth(&mut states, &mut can_close, pos + 2, depth);
                    }
                }
                _ => push_depth(&mut states, &mut can_close, pos + 1, depth),
            }
        }
    }

    // Choose the longest candidate end position that forms a valid structured-ref spec.
    for end_pos in (start + 1..=len).rev() {
        if !can_close[end_pos] {
            continue;
        }
        let inner = &input[start + 1..end_pos.saturating_sub(1)];
        if parse_inner_spec_raw(inner.trim()).is_some() {
            return Some((inner, end_pos));
        }
    }
    None
}

fn push_depth(states: &mut [Vec<u32>], can_close: &mut [bool], pos: usize, depth: u32) {
    if pos >= states.len() {
        return;
    }
    if depth == 0 && pos < can_close.len() {
        can_close[pos] = true;
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

fn columns_raw_into_owned(columns: StructuredColumnsRaw<'_>) -> StructuredColumns {
    match columns {
        StructuredColumnsRaw::All => StructuredColumns::All,
        StructuredColumnsRaw::Single(name) => StructuredColumns::Single(name.into_owned()),
        StructuredColumnsRaw::Range { start, end } => StructuredColumns::Range {
            start: start.into_owned(),
            end: end.into_owned(),
        },
        StructuredColumnsRaw::Multi(parts) => StructuredColumns::Multi(
            parts
                .into_iter()
                .map(|part| match part {
                    StructuredColumnRaw::Single(name) => StructuredColumn::Single(name.into_owned()),
                    StructuredColumnRaw::Range { start, end } => StructuredColumn::Range {
                        start: start.into_owned(),
                        end: end.into_owned(),
                    },
                })
                .collect(),
        ),
    }
}

fn parse_inner_spec(inner: &str) -> Option<(Vec<StructuredRefItem>, StructuredColumns)> {
    let (items, columns) = parse_inner_spec_raw(inner)?;
    Some((items, columns_raw_into_owned(columns)))
}

fn parse_inner_spec_raw(inner: &str) -> Option<(Vec<StructuredRefItem>, StructuredColumnsRaw<'_>)> {
    if inner.is_empty() {
        return None;
    }

    // Nested form like "[[#Headers],[Column]]" or "[[Col1]:[Col2]]"
    if inner.trim_start().starts_with('[') {
        let parts = split_top_level(inner, ',');
        if parts.is_empty() {
            return None;
        }

        let mut items: Vec<StructuredRefItem> = Vec::new();
        let mut columns: Vec<StructuredColumnRaw<'_>> = Vec::new();
        for part in parts.iter() {
            let (maybe_item, cols) = parse_bracket_group_or_range(part)?;
            if let Some(it) = maybe_item {
                items.push(it);
                continue;
            }

            match cols {
                StructuredColumnsRaw::Single(name) => columns.push(StructuredColumnRaw::Single(name)),
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

fn parse_columns_only(spec: &str) -> Option<StructuredColumnsRaw<'_>> {
    let spec = spec.trim();
    if spec.is_empty() {
        return Some(StructuredColumnsRaw::All);
    }
    if !spec.starts_with('[') {
        return Some(StructuredColumnsRaw::Single(unescape_column_name(spec)));
    }

    let parts = split_top_level(spec, ',');
    if parts.is_empty() {
        return None;
    }

    let mut columns: Vec<StructuredColumnRaw<'_>> = Vec::new();
    for part in parts {
        let part = part.trim();
        if part.is_empty() {
            continue;
        }

        let range_parts = split_top_level(part, ':');
        if range_parts.len() == 2 {
            let start = strip_wrapping_brackets(range_parts[0])?;
            let end = strip_wrapping_brackets(range_parts[1])?;
            columns.push(StructuredColumnRaw::Range {
                start: unescape_column_name(start),
                end: unescape_column_name(end),
            });
            continue;
        }

        let inner = strip_wrapping_brackets(part)?;
        columns.push(StructuredColumnRaw::Single(unescape_column_name(inner)));
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
fn parse_bracket_group_or_range(
    part: &str,
) -> Option<(Option<StructuredRefItem>, StructuredColumnsRaw<'_>)> {
    let part = part.trim();
    if part.is_empty() {
        return None;
    }

    // A range is something like "[Col1]:[Col2]" (top-level colon, not inside brackets).
    let range_parts = split_top_level(part, ':');
    if range_parts.len() == 2 {
        let start = strip_wrapping_brackets(range_parts[0])?;
        let end = strip_wrapping_brackets(range_parts[1])?;
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

fn split_top_level(input: &str, delimiter: char) -> Vec<&str> {
    let bytes = input.as_bytes();
    let delim = delimiter as u8;

    let mut parts = Vec::new();
    let mut depth: i32 = 0;
    let mut start = 0usize;
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
                depth -= 1;
                i += 1;
            }
            b if depth == 0 && b == delim => {
                parts.push(input[start..i].trim());
                i += 1;
                start = i;
            }
            _ => {
                i += 1;
            }
        }
    }
    parts.push(input[start..].trim());
    parts.into_iter().filter(|p| !p.is_empty()).collect()
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::structured_refs::{StructuredColumn, StructuredColumns, StructuredRefItem};

    #[test]
    fn parses_table_column_ref() {
        let (sref, end) = parse_structured_ref("Table1[Qty]", 0).unwrap();
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
}
