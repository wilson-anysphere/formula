use super::{StructuredColumn, StructuredColumns, StructuredRef, StructuredRefItem};

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
    let (item, columns) = parse_inner_spec(inner.trim())?;

    Some((
        StructuredRef {
            table_name,
            item,
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

    let mut depth = 0i32;
    let mut end = None;
    let mut i = start;
    while i < bytes.len() {
        match bytes[i] {
            b'[' => {
                depth += 1;
                i += 1;
            }
            b']' => {
                // Excel escapes ']' inside structured references as ']]'. At the outermost depth,
                // treat a double ']]' as a literal ']' rather than the end of the bracketed segment.
                if depth == 1 && bytes.get(i + 1) == Some(&b']') {
                    i += 2;
                    continue;
                }
                depth -= 1;
                if depth == 0 {
                    end = Some(i);
                    i += 1;
                    break;
                }
                i += 1;
            }
            _ => {
                i += 1;
            }
        }
    }

    let end = end?;
    let inner = &input[start + 1..end];
    Some((inner, i))
}

fn parse_inner_spec(inner: &str) -> Option<(Option<StructuredRefItem>, StructuredColumns)> {
    if inner.is_empty() {
        return None;
    }

    // Nested form like "[[#Headers],[Column]]" or "[[Col1]:[Col2]]"
    if inner.trim_start().starts_with('[') {
        let parts = split_top_level(inner, ',');
        if parts.is_empty() {
            return None;
        }

        let mut item: Option<StructuredRefItem> = None;
        let mut columns: Vec<StructuredColumn> = Vec::new();
        for (idx, part) in parts.iter().enumerate() {
            let (maybe_item, cols) = parse_bracket_group_or_range(part)?;
            if let Some(it) = maybe_item {
                // Excel allows an optional item (e.g. [#Headers]) followed by one or more column selectors.
                if idx == 0 && item.is_none() && columns.is_empty() {
                    item = Some(it);
                    continue;
                }
                // Multiple items in a structured ref (e.g. [[#Headers],[#Data],[Col]]) are not supported.
                return None;
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

        Some((item, cols))
    } else {
        // Simple form: "@Col", "Col", "#Headers", etc.
        let trimmed = inner.trim();
        if let Some(stripped) = trimmed.strip_prefix('@') {
            return Some((
                Some(StructuredRefItem::ThisRow),
                StructuredColumns::Single(unescape_column_name(stripped.trim())),
            ));
        }

        if let Some(item) = parse_item(trimmed) {
            return Some((Some(item), StructuredColumns::All));
        }

        Some((None, StructuredColumns::Single(unescape_column_name(trimmed))))
    }
}

fn parse_bracket_group_or_range(part: &str) -> Option<(Option<StructuredRefItem>, StructuredColumns)> {
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
    name.trim().replace("]]", "]")
}

fn is_name_start(ch: char) -> bool {
    ch.is_ascii_alphabetic() || ch == '_'
}

fn is_name_continue(ch: char) -> bool {
    ch.is_ascii_alphanumeric() || ch == '_' || ch == '.'
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
        assert_eq!(sref.item, None);
        assert_eq!(sref.columns, StructuredColumns::Single("Qty".into()));
    }

    #[test]
    fn parses_this_row_ref() {
        let (sref, _) = parse_structured_ref("[@Qty]", 0).unwrap();
        assert_eq!(sref.table_name, None);
        assert_eq!(sref.item, Some(StructuredRefItem::ThisRow));
        assert_eq!(sref.columns, StructuredColumns::Single("Qty".into()));
    }

    #[test]
    fn parses_headers_ref() {
        let (sref, _) = parse_structured_ref("Table1[[#Headers],[Qty]]", 0).unwrap();
        assert_eq!(sref.item, Some(StructuredRefItem::Headers));
        assert_eq!(sref.columns, StructuredColumns::Single("Qty".into()));
    }

    #[test]
    fn parses_multi_column_ref() {
        let (sref, _) = parse_structured_ref("Table1[[Col1],[Col3]]", 0).unwrap();
        assert_eq!(sref.item, None);
        assert_eq!(
            sref.columns,
            StructuredColumns::Multi(vec![
                StructuredColumn::Single("Col1".into()),
                StructuredColumn::Single("Col3".into()),
            ])
        );
    }
}
