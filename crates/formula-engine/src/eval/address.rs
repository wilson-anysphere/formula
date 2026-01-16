use formula_model::{parse_a1_endpoint, A1Endpoint, A1ParseError};
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
        let mut out = String::new();
        formula_model::push_a1_cell_ref(self.row, self.col, false, false, &mut out);
        out
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
    let endpoint = parse_a1_endpoint(input).map_err(|e| match e {
        A1ParseError::InvalidColumn => AddressParseError::ColumnOutOfRange,
        A1ParseError::InvalidRow => AddressParseError::RowOutOfRange,
        _ => AddressParseError::InvalidA1(input.to_string()),
    })?;

    let A1Endpoint::Cell(cell) = endpoint else {
        return Err(AddressParseError::InvalidA1(input.to_string()));
    };

    let row = cell.row;
    if row >= i32::MAX as u32 {
        return Err(AddressParseError::RowOutOfRange);
    }

    Ok(CellAddr { row, col: cell.col })
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
