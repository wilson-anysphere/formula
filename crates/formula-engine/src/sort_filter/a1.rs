use crate::sort_filter::types::RangeRef;
use thiserror::Error;

#[derive(Debug, Error, PartialEq, Eq)]
pub enum A1ParseError {
    #[error("empty reference")]
    Empty,
    #[error("invalid A1 reference: {0}")]
    Invalid(String),
}

/// Parse an A1-style range like `A1:D10` into a 0-based `RangeRef` (inclusive bounds).
pub fn parse_a1_range(a1: &str) -> Result<RangeRef, A1ParseError> {
    let a1 = a1.trim();
    if a1.is_empty() {
        return Err(A1ParseError::Empty);
    }

    let (start, end) = if let Some((left, right)) = a1.split_once(':') {
        (left, right)
    } else {
        (a1, a1)
    };

    let (start_col, start_row) = parse_a1_cell(start)?;
    let (end_col, end_row) = parse_a1_cell(end)?;

    let (start_row, end_row) = if start_row <= end_row {
        (start_row, end_row)
    } else {
        (end_row, start_row)
    };

    let (start_col, end_col) = if start_col <= end_col {
        (start_col, end_col)
    } else {
        (end_col, start_col)
    };

    Ok(RangeRef {
        start_row,
        start_col,
        end_row,
        end_col,
    })
}

pub fn to_a1_range(range: RangeRef) -> String {
    let start = to_a1_cell(range.start_col, range.start_row);
    let end = to_a1_cell(range.end_col, range.end_row);
    if start == end {
        start
    } else {
        format!("{start}:{end}")
    }
}

fn parse_a1_cell(a1: &str) -> Result<(usize, usize), A1ParseError> {
    let a1 = a1.trim();
    if a1.is_empty() {
        return Err(A1ParseError::Empty);
    }

    let mut col_part = String::new();
    let mut row_part = String::new();
    for ch in a1.chars() {
        if ch.is_ascii_alphabetic() {
            if !row_part.is_empty() {
                return Err(A1ParseError::Invalid(a1.to_string()));
            }
            col_part.push(ch.to_ascii_uppercase());
        } else if ch.is_ascii_digit() {
            row_part.push(ch);
        } else if ch == '$' {
            // Ignore absolute markers.
            continue;
        } else {
            return Err(A1ParseError::Invalid(a1.to_string()));
        }
    }

    if col_part.is_empty() || row_part.is_empty() {
        return Err(A1ParseError::Invalid(a1.to_string()));
    }

    let col = column_letters_to_index(&col_part).ok_or_else(|| A1ParseError::Invalid(a1.to_string()))?;
    let row_1based: usize = row_part
        .parse()
        .map_err(|_| A1ParseError::Invalid(a1.to_string()))?;
    if row_1based == 0 {
        return Err(A1ParseError::Invalid(a1.to_string()));
    }
    Ok((col, row_1based - 1))
}

fn to_a1_cell(col_index: usize, row_index: usize) -> String {
    format!("{}{}", index_to_column_letters(col_index), row_index + 1)
}

fn column_letters_to_index(letters: &str) -> Option<usize> {
    let mut col: usize = 0;
    for ch in letters.chars() {
        if !ch.is_ascii_uppercase() {
            return None;
        }
        let v = (ch as u8 - b'A' + 1) as usize;
        col = col.checked_mul(26)?;
        col = col.checked_add(v)?;
    }
    Some(col - 1)
}

fn index_to_column_letters(mut index: usize) -> String {
    // 0 -> A, 25 -> Z, 26 -> AA
    let mut out = String::new();
    index += 1;
    while index > 0 {
        let rem = (index - 1) % 26;
        out.push((b'A' + rem as u8) as char);
        index = (index - 1) / 26;
    }
    out.chars().rev().collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_single_cell_range() {
        let range = parse_a1_range("C3").unwrap();
        assert_eq!(
            range,
            RangeRef {
                start_row: 2,
                start_col: 2,
                end_row: 2,
                end_col: 2
            }
        );
        assert_eq!(to_a1_range(range), "C3");
    }

    #[test]
    fn parse_multi_cell_range() {
        let range = parse_a1_range("A1:D10").unwrap();
        assert_eq!(
            range,
            RangeRef {
                start_row: 0,
                start_col: 0,
                end_row: 9,
                end_col: 3
            }
        );
        assert_eq!(to_a1_range(range), "A1:D10");
    }

    #[test]
    fn parse_abs_ref() {
        let range = parse_a1_range("$B$2:$B$4").unwrap();
        assert_eq!(
            range,
            RangeRef {
                start_row: 1,
                start_col: 1,
                end_row: 3,
                end_col: 1
            }
        );
    }
}

