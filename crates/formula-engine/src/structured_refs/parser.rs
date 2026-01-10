use super::{StructuredColumns, StructuredRef, StructuredRefItem};

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
    if input.as_bytes().get(start) != Some(&b'[') {
        return None;
    }

    let mut depth = 0i32;
    let mut end = None;
    let mut i = start;
    while i < input.len() {
        let ch = input[i..].chars().next()?;
        match ch {
            '[' => depth += 1,
            ']' => {
                depth -= 1;
                if depth == 0 {
                    end = Some(i);
                    i += ch.len_utf8();
                    break;
                }
            }
            _ => {}
        }
        i += ch.len_utf8();
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

        if parts.len() == 1 {
            return parse_bracket_group_or_range(parts[0]).map(|(item, cols)| (item, cols));
        }

        if parts.len() == 2 {
            let first = strip_wrapping_brackets(parts[0])?;
            let first_item = parse_item(first);
            if let Some(item) = first_item {
                let (_, cols) = parse_bracket_group_or_range(parts[1])?;
                return Some((Some(item), cols));
            }

            // If the first part isn't an item, treat the whole thing as a multi-column selection (unsupported for now).
            return None;
        }

        None
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
    let mut parts = Vec::new();
    let mut depth: i32 = 0;
    let mut start = 0usize;
    for (idx, ch) in input.char_indices() {
        match ch {
            '[' => depth += 1,
            ']' => depth -= 1,
            _ => {}
        }
        if depth == 0 && ch == delimiter {
            parts.push(input[start..idx].trim());
            start = idx + ch.len_utf8();
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
    use crate::structured_refs::{StructuredColumns, StructuredRefItem};

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
}

