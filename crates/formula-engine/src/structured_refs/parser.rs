use super::{StructuredColumn, StructuredColumns, StructuredRef, StructuredRefItem};
use formula_model::external_refs::unescape_bracketed_identifier_content;

pub fn parse_structured_ref(input: &str, pos: usize) -> Option<(StructuredRef, usize)> {
    if pos >= input.len() {
        return None;
    }

    let bytes = input.as_bytes();
    let (table_name, bracket_start) = if bytes[pos] == b'[' {
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
        (Some(input[pos..i].to_string()), i)
    };

    let (inner, end_pos) = parse_bracketed(input, bracket_start)?;
    let (items, columns) = parse_inner_spec(inner.trim())?;

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
    states[start + 1].push(1);

    for pos in start + 1..len {
        let depths = states[pos].clone();
        if depths.is_empty() {
            continue;
        }
        for depth in depths {
            if depth == 0 {
                continue;
            }
            match bytes[pos] {
                b'[' => push_depth(&mut states, pos + 1, depth + 1),
                b']' => {
                    // Treat as closing bracket.
                    push_depth(&mut states, pos + 1, depth.saturating_sub(1));
                    // Treat as escaped literal `]`.
                    if bytes.get(pos + 1) == Some(&b']') {
                        push_depth(&mut states, pos + 2, depth);
                    }
                }
                _ => push_depth(&mut states, pos + 1, depth),
            }
        }
    }

    // Choose the longest candidate end position that forms a valid structured-ref spec.
    for end_pos in (start + 1..=len).rev() {
        if !states[end_pos].contains(&0) {
            continue;
        }
        let inner = &input[start + 1..end_pos.saturating_sub(1)];
        if parse_inner_spec(inner.trim()).is_some() {
            return Some((inner, end_pos));
        }
    }
    None
}

fn push_depth(states: &mut [Vec<u32>], pos: usize, depth: u32) {
    if pos >= states.len() {
        return;
    }
    let entry = &mut states[pos];
    if !entry.contains(&depth) {
        entry.push(depth);
    }
}

fn parse_inner_spec(inner: &str) -> Option<(Vec<StructuredRefItem>, StructuredColumns)> {
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
        let mut columns: Vec<StructuredColumn> = Vec::new();
        for part in parts.iter() {
            let (maybe_item, cols) = parse_bracket_group_or_range(part)?;
            if let Some(it) = maybe_item {
                items.push(it);
                continue;
            }

            match cols {
                StructuredColumns::Single(name) => columns.push(StructuredColumn::Single(name)),
                StructuredColumns::Range { start, end } => {
                    columns.push(StructuredColumn::Range { start, end });
                }
                StructuredColumns::All => {
                    return None;
                }
                StructuredColumns::Multi(_) => {
                    return None;
                }
            }
        }

        let cols = match columns.as_slice() {
            [] => StructuredColumns::All,
            [StructuredColumn::Single(name)] => StructuredColumns::Single(name.clone()),
            [StructuredColumn::Range { start, end }] => StructuredColumns::Range {
                start: start.clone(),
                end: end.clone(),
            },
            _ => StructuredColumns::Multi(columns),
        };

        Some((items, cols))
    } else {
        // Simple form: "@Col", "Col", "#Headers", etc.
        let trimmed = inner.trim();
        if let Some(stripped) = trimmed.strip_prefix('@') {
            let stripped = stripped.trim();
            if stripped.is_empty() {
                return Some((vec![StructuredRefItem::ThisRow], StructuredColumns::All));
            }
            if stripped.starts_with('[') {
                let cols = parse_columns_only(stripped)?;
                return Some((vec![StructuredRefItem::ThisRow], cols));
            }
            return Some((
                vec![StructuredRefItem::ThisRow],
                StructuredColumns::Single(unescape_column_name(stripped)),
            ));
        }

        if let Some(item) = parse_item(trimmed) {
            return Some((vec![item], StructuredColumns::All));
        }

        Some((
            Vec::new(),
            StructuredColumns::Single(unescape_column_name(trimmed)),
        ))
    }
}

fn parse_columns_only(spec: &str) -> Option<StructuredColumns> {
    let spec = spec.trim();
    if spec.is_empty() {
        return Some(StructuredColumns::All);
    }
    if !spec.starts_with('[') {
        return Some(StructuredColumns::Single(unescape_column_name(spec)));
    }

    let parts = split_top_level(spec, ',');
    if parts.is_empty() {
        return None;
    }

    let mut columns: Vec<StructuredColumn> = Vec::new();
    for part in parts {
        let part = part.trim();
        if part.is_empty() {
            continue;
        }

        let range_parts = split_top_level(part, ':');
        if range_parts.len() == 2 {
            let start = strip_wrapping_brackets(range_parts[0])?;
            let end = strip_wrapping_brackets(range_parts[1])?;
            columns.push(StructuredColumn::Range {
                start: unescape_column_name(start),
                end: unescape_column_name(end),
            });
            continue;
        }

        let inner = strip_wrapping_brackets(part)?;
        columns.push(StructuredColumn::Single(unescape_column_name(inner)));
    }

    let cols = match columns.as_slice() {
        [] => StructuredColumns::All,
        [StructuredColumn::Single(name)] => StructuredColumns::Single(name.clone()),
        [StructuredColumn::Range { start, end }] => StructuredColumns::Range {
            start: start.clone(),
            end: end.clone(),
        },
        _ => StructuredColumns::Multi(columns),
    };

    Some(cols)
}
fn parse_bracket_group_or_range(
    part: &str,
) -> Option<(Option<StructuredRefItem>, StructuredColumns)> {
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
            StructuredColumns::Range {
                start: unescape_column_name(start),
                end: unescape_column_name(end),
            },
        ));
    }

    let inner = strip_wrapping_brackets(part)?;
    if let Some(item) = parse_item(inner) {
        return Some((Some(item), StructuredColumns::All));
    }
    Some((None, StructuredColumns::Single(unescape_column_name(inner))))
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
    let item = item.trim();
    let item = item.strip_prefix('#').unwrap_or(item);
    let mut norm = String::new();
    for ch in item.chars() {
        if !ch.is_whitespace() {
            norm.push(ch.to_ascii_lowercase());
        }
    }
    match norm.as_str() {
        "all" => Some(StructuredRefItem::All),
        "data" => Some(StructuredRefItem::Data),
        "headers" => Some(StructuredRefItem::Headers),
        "totals" => Some(StructuredRefItem::Totals),
        "thisrow" => Some(StructuredRefItem::ThisRow),
        _ => None,
    }
}

fn unescape_column_name(name: &str) -> String {
    // Excel escapes ']' as ']]' in structured references.
    unescape_bracketed_identifier_content(name.trim()).into_owned()
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
