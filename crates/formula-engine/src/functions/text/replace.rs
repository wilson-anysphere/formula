use crate::error::{ExcelError, ExcelResult};

/// SUBSTITUTE(text, old_text, new_text, [instance_num])
pub fn substitute(
    text: &str,
    old_text: &str,
    new_text: &str,
    instance_num: Option<i32>,
) -> ExcelResult<String> {
    if old_text.is_empty() {
        return Ok(text.to_string());
    }

    match instance_num {
        None => Ok(text.replace(old_text, new_text)),
        Some(n) => {
            if n < 1 {
                return Err(ExcelError::Value);
            }

            let mut count = 0i32;
            let mut last_idx = 0usize;
            for (idx, _) in text.match_indices(old_text) {
                count += 1;
                if count == n {
                    let out_len = text.len().saturating_sub(old_text.len()).saturating_add(new_text.len());
                    let mut out = String::new();
                    if out.try_reserve_exact(out_len).is_err() {
                        debug_assert!(false, "allocation failed (substitute nth match, len={out_len})");
                        return Err(ExcelError::Num);
                    }
                    out.push_str(&text[..idx]);
                    out.push_str(new_text);
                    out.push_str(&text[idx + old_text.len()..]);
                    return Ok(out);
                }
                last_idx = idx;
            }

            // No nth match: return original.
            let _ = last_idx;
            Ok(text.to_string())
        }
    }
}

/// REPLACE(old_text, start_num, num_chars, new_text)
pub fn replace(
    old_text: &str,
    start_num: i32,
    num_chars: i32,
    new_text: &str,
) -> ExcelResult<String> {
    if start_num < 1 || num_chars < 0 {
        return Err(ExcelError::Value);
    }

    let start = (start_num - 1) as usize;
    let len_chars = old_text.chars().count();
    let start = start.min(len_chars);
    let end = (start + num_chars as usize).min(len_chars);

    let start_byte = char_pos_to_byte(old_text, start);
    let end_byte = char_pos_to_byte(old_text, end);

    let out_len = old_text
        .len()
        .saturating_sub(end_byte.saturating_sub(start_byte))
        .saturating_add(new_text.len());
    let mut out = String::new();
    if out.try_reserve_exact(out_len).is_err() {
        debug_assert!(false, "allocation failed (replace, len={out_len})");
        return Err(ExcelError::Num);
    }
    out.push_str(&old_text[..start_byte]);
    out.push_str(new_text);
    out.push_str(&old_text[end_byte..]);
    Ok(out)
}

fn char_pos_to_byte(s: &str, char_pos: usize) -> usize {
    if char_pos == 0 {
        return 0;
    }
    if char_pos >= s.chars().count() {
        return s.len();
    }
    s.char_indices()
        .nth(char_pos)
        .map(|(idx, _)| idx)
        .unwrap_or_else(|| s.len())
}
