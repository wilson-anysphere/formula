use formula_model::CellRef;
use thiserror::Error;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct CellAddr {
    pub row: u32,
    pub col: u32,
}

impl CellAddr {
    /// Sentinel value used in range endpoints to indicate "the last row/column of the sheet".
    ///
    /// This value is never produced by [`parse_a1`]: the largest valid A1 row is `i32::MAX`
    /// (1-indexed), which parses to `i32::MAX - 1` when stored 0-indexed. This matches the
    /// eval IR (`eval::ast::Ref`) which stores absolute coordinates in `i32` with `i32::MAX`
    /// reserved as a sheet-end sentinel.
    pub const SHEET_END: u32 = u32::MAX;

    /// Formats this 0-indexed address into an Excel-style A1 string (e.g. `A1`, `BC32`).
    #[must_use]
    pub fn to_a1(self) -> String {
        CellRef::new(self.row, self.col).to_a1()
    }
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
    if row > i32::MAX as u32 {
        return Err(AddressParseError::RowOutOfRange);
    }

    // Excel max is XFD (16,384) columns and 1,048,576 rows.
    //
    // We continue to enforce the Excel column bound because the engine data model (and
    // `formula-model::CellKey`) assumes a fixed 16,384-column grid.
    //
    // Rows are capped at `i32::MAX` (1-indexed). This matches the engine's internal row limit and
    // avoids overflow in internal row/col arithmetic which relies on signed `i32` offsets.
    if col > formula_model::EXCEL_MAX_COLS {
        return Err(AddressParseError::ColumnOutOfRange);
    }

    Ok(CellAddr {
        row: row - 1,
        col: col - 1,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn cell_addr_formats_to_a1() {
        assert_eq!(parse_a1("A1").unwrap().to_a1(), "A1");
        assert_eq!(parse_a1("$BC$32").unwrap().to_a1(), "BC32");
    }

    #[test]
    fn cell_addr_formats_large_rows_without_overflow() {
        let addr = CellAddr {
            row: u32::MAX,
            col: 0,
        };
        assert_eq!(addr.to_a1(), "A4294967296");
    }

    #[test]
    fn parse_a1_caps_rows_at_i32_max() {
        assert!(parse_a1("A2147483647").is_ok());
        assert!(matches!(
            parse_a1("A2147483648"),
            Err(AddressParseError::RowOutOfRange)
        ));
    }
}
