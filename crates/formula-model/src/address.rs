use core::fmt;

use serde::{Deserialize, Serialize};

/// A reference to a single cell within a worksheet.
///
/// Rows and columns are **0-indexed**:
/// - `row = 0` is Excel row `1`
/// - `col = 0` is Excel column `A`
#[derive(Copy, Clone, Debug, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub struct CellRef {
    /// 0-indexed row.
    pub row: u32,
    /// 0-indexed column.
    pub col: u32,
}

impl CellRef {
    /// Construct a new [`CellRef`].
    #[inline]
    pub const fn new(row: u32, col: u32) -> Self {
        Self { row, col }
    }

    /// Convert to Excel A1 notation (e.g. `A1`, `BC32`).
    pub fn to_a1(self) -> String {
        format!("{}{}", col_to_name(self.col), self.row + 1)
    }

    /// Parse an Excel A1-style reference (e.g. `A1`, `$B$2`).
    pub fn from_a1(a1: &str) -> Result<Self, A1ParseError> {
        let s = a1.trim();
        if s.is_empty() {
            return Err(A1ParseError::Empty);
        }

        // Accept optional `$` markers.
        let mut idx = 0usize;
        let bytes = s.as_bytes();
        if bytes.get(idx) == Some(&b'$') {
            idx += 1;
        }

        let col_start = idx;
        while idx < bytes.len() && bytes[idx].is_ascii_alphabetic() {
            idx += 1;
        }

        if idx == col_start {
            return Err(A1ParseError::MissingColumn);
        }

        let col_str = &s[col_start..idx];
        if bytes.get(idx) == Some(&b'$') {
            idx += 1;
        }

        let row_start = idx;
        while idx < bytes.len() && bytes[idx].is_ascii_digit() {
            idx += 1;
        }

        if idx == row_start {
            return Err(A1ParseError::MissingRow);
        }
        if idx != bytes.len() {
            return Err(A1ParseError::TrailingCharacters);
        }

        let col = name_to_col(col_str)?;
        if col >= crate::cell::EXCEL_MAX_COLS {
            return Err(A1ParseError::InvalidColumn);
        }
        let row_1_based: u32 = s[row_start..idx]
            .parse()
            .map_err(|_| A1ParseError::InvalidRow)?;
        if row_1_based == 0 {
            return Err(A1ParseError::InvalidRow);
        }
        if row_1_based > crate::cell::EXCEL_MAX_ROWS {
            return Err(A1ParseError::InvalidRow);
        }

        Ok(Self {
            row: row_1_based - 1,
            col,
        })
    }
}

impl fmt::Display for CellRef {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(&self.to_a1())
    }
}

/// A rectangular region within a worksheet.
///
/// The range is inclusive and always normalized such that:
/// - `start.row <= end.row`
/// - `start.col <= end.col`
#[derive(Copy, Clone, Debug, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub struct Range {
    pub start: CellRef,
    pub end: CellRef,
}

impl Range {
    /// Construct a new range, normalizing coordinates if needed.
    pub const fn new(a: CellRef, b: CellRef) -> Self {
        let start_row = if a.row <= b.row { a.row } else { b.row };
        let end_row = if a.row <= b.row { b.row } else { a.row };
        let start_col = if a.col <= b.col { a.col } else { b.col };
        let end_col = if a.col <= b.col { b.col } else { a.col };
        Self {
            start: CellRef::new(start_row, start_col),
            end: CellRef::new(end_row, end_col),
        }
    }

    /// Returns true if `cell` lies within this range.
    #[inline]
    pub const fn contains(&self, cell: CellRef) -> bool {
        cell.row >= self.start.row
            && cell.row <= self.end.row
            && cell.col >= self.start.col
            && cell.col <= self.end.col
    }

    /// Number of columns in the range.
    #[inline]
    pub const fn width(&self) -> u32 {
        self.end.col - self.start.col + 1
    }

    /// Number of rows in the range.
    #[inline]
    pub const fn height(&self) -> u32 {
        self.end.row - self.start.row + 1
    }

    /// Returns true if the range is exactly one cell.
    #[inline]
    pub const fn is_single_cell(&self) -> bool {
        self.start.row == self.end.row && self.start.col == self.end.col
    }

    /// Parse an Excel A1-style range like `A1:B2` or a single-cell reference like `C3`.
    pub fn from_a1(a1: &str) -> Result<Self, RangeParseError> {
        let s = a1.trim();
        if s.is_empty() {
            return Err(RangeParseError::Empty);
        }

        match s.split_once(':') {
            None => {
                let cell = CellRef::from_a1(s).map_err(RangeParseError::Cell)?;
                Ok(Range::new(cell, cell))
            }
            Some((a, b)) => {
                let start = CellRef::from_a1(a).map_err(RangeParseError::Cell)?;
                let end = CellRef::from_a1(b).map_err(RangeParseError::Cell)?;
                Ok(Range::new(start, end))
            }
        }
    }
}

impl fmt::Display for Range {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        if self.is_single_cell() {
            write!(f, "{}", self.start)
        } else {
            write!(f, "{}:{}", self.start, self.end)
        }
    }
}

/// Errors that can occur when parsing an A1 cell reference.
#[derive(Copy, Clone, Debug, PartialEq, Eq)]
pub enum A1ParseError {
    Empty,
    MissingColumn,
    MissingRow,
    InvalidColumn,
    InvalidRow,
    TrailingCharacters,
}

impl fmt::Display for A1ParseError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        let msg = match self {
            A1ParseError::Empty => "empty A1 reference",
            A1ParseError::MissingColumn => "missing column in A1 reference",
            A1ParseError::MissingRow => "missing row in A1 reference",
            A1ParseError::InvalidColumn => "invalid column in A1 reference",
            A1ParseError::InvalidRow => "invalid row in A1 reference",
            A1ParseError::TrailingCharacters => "trailing characters in A1 reference",
        };
        f.write_str(msg)
    }
}

impl std::error::Error for A1ParseError {}

/// Errors that can occur when parsing an A1 range.
#[derive(Debug)]
pub enum RangeParseError {
    Empty,
    Cell(A1ParseError),
}

impl fmt::Display for RangeParseError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            RangeParseError::Empty => f.write_str("empty A1 range"),
            RangeParseError::Cell(e) => write!(f, "invalid cell reference in range: {e}"),
        }
    }
}

impl std::error::Error for RangeParseError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            RangeParseError::Empty => None,
            RangeParseError::Cell(e) => Some(e),
        }
    }
}

fn col_to_name(col: u32) -> String {
    // Excel columns are 1-based in A1 notation. We store 0-based internally.
    let mut n = col + 1;
    let mut out = Vec::<u8>::new();
    while n > 0 {
        let rem = (n - 1) % 26;
        out.push(b'A' + rem as u8);
        n = (n - 1) / 26;
    }
    out.reverse();
    String::from_utf8(out).expect("column letters are always valid UTF-8")
}

fn name_to_col(s: &str) -> Result<u32, A1ParseError> {
    let mut col: u32 = 0;
    for b in s.bytes() {
        if !b.is_ascii_alphabetic() {
            return Err(A1ParseError::InvalidColumn);
        }
        let v = (b.to_ascii_uppercase() - b'A') as u32 + 1;
        col = col
            .checked_mul(26)
            .and_then(|c| c.checked_add(v))
            .ok_or(A1ParseError::InvalidColumn)?;
    }
    if col == 0 {
        return Err(A1ParseError::InvalidColumn);
    }
    Ok(col - 1)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a1_roundtrip() {
        let c = CellRef::new(0, 0);
        assert_eq!(c.to_a1(), "A1");
        assert_eq!(CellRef::from_a1("A1").unwrap(), c);
        assert_eq!(CellRef::from_a1("$A$1").unwrap(), c);

        let c2 = CellRef::new(31, 54); // BC32
        assert_eq!(c2.to_a1(), "BC32");
        assert_eq!(CellRef::from_a1("bc32").unwrap(), c2);
    }

    #[test]
    fn a1_range_parsing() {
        let r = Range::from_a1("A1:B2").unwrap();
        assert_eq!(r.start, CellRef::new(0, 0));
        assert_eq!(r.end, CellRef::new(1, 1));

        let single = Range::from_a1("C3").unwrap();
        assert!(single.is_single_cell());
        assert_eq!(single.start, CellRef::new(2, 2));
    }

    #[test]
    fn a1_bounds_are_excel_compatible() {
        assert!(CellRef::from_a1("XFD1048576").is_ok());
        assert!(CellRef::from_a1("XFE1").is_err()); // col 16385 is out of bounds
        assert!(CellRef::from_a1("A1048577").is_err()); // row 1,048,577 is out of bounds
    }
}
