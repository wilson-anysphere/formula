use core::fmt;

use serde::de::Error as _;
use serde::{Deserialize, Deserializer, Serialize};

/// A reference to a single cell within a worksheet.
///
/// Rows and columns are **0-indexed**:
/// - `row = 0` is Excel row `1`
/// - `col = 0` is Excel column `A`
#[derive(Copy, Clone, Debug, PartialEq, Eq, Hash, Serialize)]
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

    /// Checked constructor for a [`CellRef`].
    ///
    /// Returns `None` if `col` is outside Excel's sheet bounds.
    #[inline]
    pub fn try_new(row: u32, col: u32) -> Option<Self> {
        if col < crate::cell::EXCEL_MAX_COLS {
            Some(Self::new(row, col))
        } else {
            None
        }
    }

    /// Convert to Excel A1 notation (e.g. `A1`, `BC32`).
    pub fn to_a1(self) -> String {
        // Rows are stored as 0-based u32s and intentionally allow values beyond Excel's
        // 1,048,576 limit. Do row arithmetic in u64 so formatting is robust even for large
        // internal sentinel values (e.g. u32::MAX) without debug overflow panics.
        let row_1_based = u64::from(self.row) + 1;
        format!("{}{}", col_to_name(self.col), row_1_based)
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

        Ok(Self {
            row: row_1_based - 1,
            col,
        })
    }
}

impl<'de> Deserialize<'de> for CellRef {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        #[derive(Deserialize)]
        struct Helper {
            row: u32,
            col: u32,
        }

        let helper = Helper::deserialize(deserializer)?;
        if helper.col >= crate::cell::EXCEL_MAX_COLS {
            return Err(D::Error::custom(format!(
                "col out of Excel bounds: {}",
                helper.col
            )));
        }
        Ok(CellRef::new(helper.row, helper.col))
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
#[derive(Copy, Clone, Debug, PartialEq, Eq, Hash, Serialize)]
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

    /// Returns true if this range overlaps `other` (i.e., their intersection is non-empty).
    #[inline]
    pub const fn intersects(&self, other: &Range) -> bool {
        self.start.row <= other.end.row
            && self.end.row >= other.start.row
            && self.start.col <= other.end.col
            && self.end.col >= other.start.col
    }

    /// Returns the intersection of this range with `other`, or `None` if they do not overlap.
    pub const fn intersection(&self, other: &Range) -> Option<Range> {
        if !self.intersects(other) {
            return None;
        }
        let start = CellRef::new(
            max_u32(self.start.row, other.start.row),
            max_u32(self.start.col, other.start.col),
        );
        let end = CellRef::new(
            min_u32(self.end.row, other.end.row),
            min_u32(self.end.col, other.end.col),
        );
        Some(Range::new(start, end))
    }

    /// Returns the bounding box that contains both this range and `other`.
    pub const fn bounding_box(&self, other: &Range) -> Range {
        let start = CellRef::new(
            min_u32(self.start.row, other.start.row),
            min_u32(self.start.col, other.start.col),
        );
        let end = CellRef::new(
            max_u32(self.end.row, other.end.row),
            max_u32(self.end.col, other.end.col),
        );
        Range::new(start, end)
    }

    /// Total number of cells in the range (`width * height`).
    pub const fn cell_count(&self) -> u64 {
        (self.width() as u64) * (self.height() as u64)
    }

    /// Iterate over all cells in the range in row-major order (top-to-bottom, left-to-right).
    pub fn iter(self) -> RangeIter {
        RangeIter::new(self)
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

const fn min_u32(a: u32, b: u32) -> u32 {
    if a < b {
        a
    } else {
        b
    }
}

const fn max_u32(a: u32, b: u32) -> u32 {
    if a > b {
        a
    } else {
        b
    }
}

impl<'de> Deserialize<'de> for Range {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        #[derive(Deserialize)]
        struct Helper {
            start: CellRef,
            end: CellRef,
        }

        let helper = Helper::deserialize(deserializer)?;
        Ok(Range::new(helper.start, helper.end))
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

/// Iterator over all cells in a [`Range`] in row-major order.
#[derive(Clone, Debug)]
pub struct RangeIter {
    range: Range,
    next: Option<CellRef>,
}

impl RangeIter {
    fn new(range: Range) -> Self {
        Self {
            range,
            next: Some(range.start),
        }
    }
}

impl Iterator for RangeIter {
    type Item = CellRef;

    fn next(&mut self) -> Option<Self::Item> {
        let current = self.next?;
        if current.row == self.range.end.row && current.col == self.range.end.col {
            self.next = None;
            return Some(current);
        }

        let mut next_row = current.row;
        let mut next_col = current.col + 1;
        if next_col > self.range.end.col {
            next_col = self.range.start.col;
            next_row = next_row.saturating_add(1);
        }

        self.next = Some(CellRef::new(next_row, next_col));
        Some(current)
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

/// Convert a 0-based column index to an Excel column label and append it to `out`.
pub fn push_column_label(col: u32, out: &mut String) {
    // Excel column labels are 1-based.
    let mut col = u64::from(col) + 1;
    let mut buf = [0u8; 10];
    let mut i = 0usize;
    while col > 0 {
        let rem = ((col - 1) % 26) as u8;
        buf[i] = b'A' + rem;
        i += 1;
        col = (col - 1) / 26;
    }
    for ch in buf[..i].iter().rev() {
        out.push(*ch as char);
    }
}

fn col_to_name(col: u32) -> String {
    let mut out = String::new();
    push_column_label(col, &mut out);
    out
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
        assert_eq!(
            CellRef::from_a1("A1048577").unwrap(),
            CellRef::new(1_048_576, 0)
        );
    }

    #[test]
    fn a1_formats_excel_max_row() {
        // Excel row 1,048,576 is stored as 0-based row 1,048,575.
        assert_eq!(CellRef::new(1_048_575, 0).to_a1(), "A1048576");
    }

    #[test]
    fn a1_formats_u32_max_row_without_overflow() {
        let a1 = CellRef::new(u32::MAX, 0).to_a1();
        assert!(a1.starts_with('A'));
        assert!(a1.contains("4294967296"));
        assert_eq!(a1, "A4294967296");
    }

    #[test]
    fn range_iterates_row_major() {
        let r = Range::new(CellRef::new(0, 0), CellRef::new(1, 1));
        let cells: Vec<_> = r.iter().collect();
        assert_eq!(
            cells,
            vec![
                CellRef::new(0, 0),
                CellRef::new(0, 1),
                CellRef::new(1, 0),
                CellRef::new(1, 1),
            ]
        );
    }

    #[test]
    fn cell_ref_deserialize_validates_bounds() {
        let json = serde_json::json!({ "row": 0, "col": crate::cell::EXCEL_MAX_COLS });
        let err = serde_json::from_value::<CellRef>(json).unwrap_err();
        assert!(err.to_string().contains("out of Excel bounds"));
    }

    #[test]
    fn range_deserialize_normalizes() {
        let json = serde_json::json!({
            "start": { "row": 5, "col": 5 },
            "end": { "row": 2, "col": 2 }
        });

        let range = serde_json::from_value::<Range>(json).unwrap();
        assert_eq!(range.start, CellRef::new(2, 2));
        assert_eq!(range.end, CellRef::new(5, 5));
    }

    #[test]
    fn range_intersection_and_bbox() {
        let a = Range::from_a1("A1:C3").unwrap();
        let b = Range::from_a1("B2:D4").unwrap();

        assert!(a.intersects(&b));
        assert_eq!(a.intersection(&b), Some(Range::from_a1("B2:C3").unwrap()));
        assert_eq!(a.bounding_box(&b), Range::from_a1("A1:D4").unwrap());
        assert_eq!(a.cell_count(), 9);
    }
}
