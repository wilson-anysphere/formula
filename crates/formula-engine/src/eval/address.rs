use thiserror::Error;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct CellAddr {
    pub row: u32,
    pub col: u32,
}

#[derive(Debug, Error, Clone, PartialEq, Eq)]
pub enum AddressParseError {
    #[error("invalid A1 address: {0}")]
    InvalidA1(String),
    #[error("column out of range")]
    ColumnOutOfRange,
    #[error("row out of range")]
    RowOutOfRange,
}

/// Parse an A1-style address like `A1` or `$B$12` into a 0-indexed [`CellAddr`].
pub fn parse_a1(input: &str) -> Result<CellAddr, AddressParseError> {
    let input = input.trim();
    if input.is_empty() {
        return Err(AddressParseError::InvalidA1(input.to_string()));
    }

    let mut chars = input.chars().peekable();
    // Optional absolute marker.
    if matches!(chars.peek(), Some('$')) {
        chars.next();
    }

    let mut col: u32 = 0;
    let mut col_len = 0;
    while let Some(ch) = chars.peek().copied() {
        if ch.is_ascii_alphabetic() {
            let up = ch.to_ascii_uppercase();
            let digit = (up as u8 - b'A' + 1) as u32;
            col = col
                .checked_mul(26)
                .and_then(|v| v.checked_add(digit))
                .ok_or(AddressParseError::ColumnOutOfRange)?;
            col_len += 1;
            chars.next();
        } else {
            break;
        }
    }

    if col_len == 0 {
        return Err(AddressParseError::InvalidA1(input.to_string()));
    }

    // Optional absolute marker.
    if matches!(chars.peek(), Some('$')) {
        chars.next();
    }

    let mut row: u32 = 0;
    let mut row_len = 0;
    while let Some(ch) = chars.peek().copied() {
        if ch.is_ascii_digit() {
            row = row
                .checked_mul(10)
                .and_then(|v| v.checked_add((ch as u8 - b'0') as u32))
                .ok_or(AddressParseError::RowOutOfRange)?;
            row_len += 1;
            chars.next();
        } else {
            break;
        }
    }

    if row_len == 0 || chars.next().is_some() {
        return Err(AddressParseError::InvalidA1(input.to_string()));
    }

    if col == 0 {
        return Err(AddressParseError::ColumnOutOfRange);
    }
    if row == 0 {
        return Err(AddressParseError::RowOutOfRange);
    }

    // Excel max is XFD (16384) columns and 1,048,576 rows. Enforce for sanity.
    if col > 16_384 {
        return Err(AddressParseError::ColumnOutOfRange);
    }
    if row > 1_048_576 {
        return Err(AddressParseError::RowOutOfRange);
    }

    Ok(CellAddr {
        row: row - 1,
        col: col - 1,
    })
}

