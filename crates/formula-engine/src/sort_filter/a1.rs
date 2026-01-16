use crate::sort_filter::types::RangeRef;
use formula_model::column_label_to_index;
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

    let col = column_label_to_index(&col_part)
        .map_err(|_| A1ParseError::Invalid(a1.to_string()))? as usize;
    let row_1based: usize = row_part
        .parse()
        .map_err(|_| A1ParseError::Invalid(a1.to_string()))?;
    if row_1based == 0 {
        return Err(A1ParseError::Invalid(a1.to_string()));
    }
    Ok((col, row_1based - 1))
}

fn to_a1_cell(col_index: usize, row_index: usize) -> String {
    // Prefer the shared CellRef formatter for consistency and to avoid overflow in `row + 1`
    // arithmetic for large row indices.
    match (u32::try_from(row_index), u32::try_from(col_index)) {
        (Ok(row), Ok(col)) => {
            let mut out = String::new();
            formula_model::push_a1_cell_ref(row, col, false, false, &mut out);
            out
        }
        _ => {
            let row_1_based = u64::try_from(row_index)
                .unwrap_or(u64::MAX)
                .saturating_add(1);
            format!("{}{}", index_to_column_letters(col_index), row_1_based)
        }
    }
}

fn index_to_column_letters(index: usize) -> String {
    // 0 -> A, 25 -> Z, 26 -> AA
    //
    // Do arithmetic in u64 so we don't overflow on large `usize` indices (e.g. wasm32 where
    // `usize == u32` and callers may use `u32::MAX`).
    let mut n = u64::try_from(index).unwrap_or(u64::MAX).saturating_add(1);
    let mut out = Vec::<u8>::new();
    while n > 0 {
        let rem = (n - 1) % 26;
        out.push(b'A' + rem as u8);
        n = (n - 1) / 26;
    }
    out.reverse();
    String::from_utf8(out).expect("column letters are always valid UTF-8")
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

    #[test]
    fn formats_large_rows_without_overflow() {
        let row = u32::MAX as usize;
        let range = RangeRef {
            start_row: row,
            start_col: 0,
            end_row: row,
            end_col: 0,
        };
        assert_eq!(to_a1_range(range), "A4294967296");
    }
}
