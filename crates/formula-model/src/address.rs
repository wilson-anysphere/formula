use core::fmt;
use core::fmt::Write as _;

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
        let mut out = String::new();
        push_column_label(self.col, &mut out);
        let _ = write!(out, "{row_1_based}");
        out
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

/// A best-effort A1 endpoint used by print/defined-name parsing.
///
/// This accepts:
/// - cell endpoints (e.g. `A1`, `$B$2`)
/// - whole-row endpoints (e.g. `$1`)
/// - whole-column endpoints (e.g. `$A`)
///
/// Rows and columns are 0-indexed in the returned values.
#[derive(Copy, Clone, Debug, PartialEq, Eq)]
pub enum A1Endpoint {
    Cell(CellRef),
    Row(u32),
    Col(u32),
}

/// Parse an A1 endpoint (cell / whole row / whole column).
///
/// This is stricter than the formula lexer:
/// - columns must be within Excel bounds (`A..=XFD`)
/// - `0` row numbers are rejected
/// - any non-ASCII letter/digit characters are rejected
pub fn parse_a1_endpoint(s: &str) -> Result<A1Endpoint, A1ParseError> {
    let s = s.trim();
    if s.is_empty() {
        return Err(A1ParseError::Empty);
    }

    let mut col_1_based: u32 = 0;
    let mut col_len = 0usize;
    let mut row_1_based: u32 = 0;
    let mut row_len = 0usize;
    let mut saw_digit = false;

    for &b in s.as_bytes() {
        if b == b'$' {
            continue;
        }

        if b.is_ascii_alphabetic() {
            if saw_digit {
                return Err(A1ParseError::TrailingCharacters);
            }
            col_len += 1;
            let v = (b.to_ascii_uppercase() - b'A') as u32 + 1;
            col_1_based = col_1_based
                .checked_mul(26)
                .and_then(|c| c.checked_add(v))
                .ok_or(A1ParseError::InvalidColumn)?;
            continue;
        }

        if b.is_ascii_digit() {
            saw_digit = true;
            row_len += 1;
            let v = (b - b'0') as u32;
            row_1_based = row_1_based
                .checked_mul(10)
                .and_then(|r| r.checked_add(v))
                .ok_or(A1ParseError::InvalidRow)?;
            continue;
        }

        return Err(A1ParseError::TrailingCharacters);
    }

    if col_len == 0 && row_len == 0 {
        return Err(A1ParseError::Empty);
    }

    if col_len > 0 {
        let col0 = col_1_based.saturating_sub(1);
        if col0 >= crate::cell::EXCEL_MAX_COLS {
            return Err(A1ParseError::InvalidColumn);
        }

        if row_len == 0 {
            return Ok(A1Endpoint::Col(col0));
        }

        if row_1_based == 0 {
            return Err(A1ParseError::InvalidRow);
        }

        return Ok(A1Endpoint::Cell(CellRef::new(row_1_based - 1, col0)));
    }

    // Row-only endpoint.
    if row_1_based == 0 {
        return Err(A1ParseError::InvalidRow);
    }
    Ok(A1Endpoint::Row(row_1_based - 1))
}

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

/// Append an A1-style cell reference (e.g. `A1`, `$B$2`) to `out`.
///
/// - `row1` is 1-based (Excel row numbering).
/// - `col0` is 0-based (Excel column `A` is `0`).
/// - `abs_col` / `abs_row` control whether `$` markers are included.
pub fn push_a1_cell_ref_row1(row1: u64, col0: u32, abs_col: bool, abs_row: bool, out: &mut String) {
    if abs_col {
        out.push('$');
    }
    push_column_label(col0, out);
    if abs_row {
        out.push('$');
    }
    let _ = write!(out, "{row1}");
}

/// Append an A1-style cell reference (e.g. `A1`, `$B$2`) to `out`.
///
/// - `row0` is 0-based (Excel row `1` is `0`).
/// - `col0` is 0-based (Excel column `A` is `0`).
/// - `abs_col` / `abs_row` control whether `$` markers are included.
pub fn push_a1_cell_ref(row0: u32, col0: u32, abs_col: bool, abs_row: bool, out: &mut String) {
    push_a1_cell_ref_row1(u64::from(row0) + 1, col0, abs_col, abs_row, out);
}

/// Append an A1-style cell range (e.g. `A1:B2`, `$A$1:$B$2`) to `out`.
///
/// If the endpoints are identical, this emits a single-cell reference.
///
/// - `start_row1` / `end_row1` are 1-based (Excel row numbering).
/// - `start_col0` / `end_col0` are 0-based (Excel column `A` is `0`).
/// - `abs_col` / `abs_row` control whether `$` markers are included.
pub fn push_a1_cell_range_row1(
    start_row1: u64,
    start_col0: u32,
    end_row1: u64,
    end_col0: u32,
    abs_col: bool,
    abs_row: bool,
    out: &mut String,
) {
    push_a1_cell_ref_row1(start_row1, start_col0, abs_col, abs_row, out);
    if start_row1 == end_row1 && start_col0 == end_col0 {
        return;
    }
    out.push(':');
    push_a1_cell_ref_row1(end_row1, end_col0, abs_col, abs_row, out);
}

/// Append an A1-style cell range (e.g. `A1:B2`, `$A$1:$B$2`) to `out`.
///
/// If the endpoints are identical, this emits a single-cell reference.
///
/// - `*_row0` are 0-based (Excel row `1` is `0`).
/// - `*_col0` are 0-based (Excel column `A` is `0`).
/// - `abs_col` / `abs_row` control whether `$` markers are included.
pub fn push_a1_cell_range(
    start_row0: u32,
    start_col0: u32,
    end_row0: u32,
    end_col0: u32,
    abs_col: bool,
    abs_row: bool,
    out: &mut String,
) {
    push_a1_cell_range_row1(
        u64::from(start_row0) + 1,
        start_col0,
        u64::from(end_row0) + 1,
        end_col0,
        abs_col,
        abs_row,
        out,
    );
}

/// Append an A1-style cell range (e.g. `A1:B2`, `$A1:B$2`) to `out`.
///
/// This is the low-level "area" formatter used by BIFF/XLS formula decoders where the absolute
/// markers are stored per endpoint. If the endpoints are identical, this emits a single-cell
/// reference using the first endpoint's absolute flags.
///
/// - `*_row1` are 1-based (Excel row numbering).
/// - `*_col0` are 0-based (Excel column `A` is `0`).
/// - `abs_*` control whether `$` markers are included.
pub fn push_a1_cell_area_row1(
    start_row1: u64,
    start_col0: u32,
    abs_start_col: bool,
    abs_start_row: bool,
    end_row1: u64,
    end_col0: u32,
    abs_end_col: bool,
    abs_end_row: bool,
    out: &mut String,
) {
    push_a1_cell_ref_row1(start_row1, start_col0, abs_start_col, abs_start_row, out);
    if start_row1 == end_row1 && start_col0 == end_col0 {
        return;
    }
    out.push(':');
    push_a1_cell_ref_row1(end_row1, end_col0, abs_end_col, abs_end_row, out);
}

/// Append an A1-style whole-row reference (e.g. `$1`) to `out`.
///
/// - `row1` is 1-based (Excel row numbering).
/// - `abs_row` controls whether the `$` marker is included.
pub fn push_a1_row_ref_row1(row1: u64, abs_row: bool, out: &mut String) {
    if abs_row {
        out.push('$');
    }
    let _ = write!(out, "{row1}");
}

/// Append an A1-style whole-row range (e.g. `$1:$10`) to `out`.
///
/// - `start_row1` / `end_row1` are 1-based (Excel row numbering).
/// - `abs_row` controls whether `$` markers are included on both endpoints.
pub fn push_a1_row_range_row1(start_row1: u64, end_row1: u64, abs_row: bool, out: &mut String) {
    push_a1_row_ref_row1(start_row1, abs_row, out);
    out.push(':');
    push_a1_row_ref_row1(end_row1, abs_row, out);
}

/// Append an A1-style whole-column reference (e.g. `$A`) to `out`.
///
/// - `col0` is 0-based (Excel column `A` is `0`).
/// - `abs_col` controls whether the `$` marker is included.
pub fn push_a1_col_ref(col0: u32, abs_col: bool, out: &mut String) {
    if abs_col {
        out.push('$');
    }
    push_column_label(col0, out);
}

/// Append an A1-style whole-column range (e.g. `$A:$D`) to `out`.
///
/// - `start_col0` / `end_col0` are 0-based (Excel column `A` is `0`).
/// - `abs_col` controls whether `$` markers are included on both endpoints.
pub fn push_a1_col_range(start_col0: u32, end_col0: u32, abs_col: bool, out: &mut String) {
    push_a1_col_ref(start_col0, abs_col, out);
    out.push(':');
    push_a1_col_ref(end_col0, abs_col, out);
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

/// Convert an Excel column label (e.g. `A`, `XFD`) into a 0-based column index.
///
/// This is a strict A1 helper:
/// - Only ASCII letters are accepted.
/// - The result must be within Excel's column bounds (`A..=XFD`).
pub fn column_label_to_index(label: &str) -> Result<u32, A1ParseError> {
    let col = name_to_col(label)?;
    if col >= crate::cell::EXCEL_MAX_COLS {
        return Err(A1ParseError::InvalidColumn);
    }
    Ok(col)
}

/// Convert an Excel column label (e.g. `A`, `XFE`) into a 0-based column index.
///
/// This is a **lenient** A1 helper:
/// - Only ASCII letters are accepted.
/// - The label must be 1-3 letters (Excel's A1 grammar).
/// - The result is **not** restricted to Excel's `A..=XFD` column bound.
///
/// This is used by formula lexers/parsers that need to recognize out-of-bounds references like
/// `XFE1` so they can be represented as references and later evaluate to `#REF!`, rather than
/// being tokenized as identifiers.
pub fn column_label_to_index_lenient(label: &str) -> Result<u32, A1ParseError> {
    if label.is_empty() || label.len() > 3 {
        return Err(A1ParseError::InvalidColumn);
    }
    name_to_col(label)
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

    #[test]
    fn push_a1_cell_ref_respects_absolute_flags() {
        let mut out = String::new();
        push_a1_cell_ref_row1(2, 1, true, true, &mut out);
        assert_eq!(out, "$B$2");

        let mut out = String::new();
        push_a1_cell_ref_row1(2, 1, true, false, &mut out);
        assert_eq!(out, "$B2");

        let mut out = String::new();
        push_a1_cell_ref_row1(2, 1, false, true, &mut out);
        assert_eq!(out, "B$2");

        let mut out = String::new();
        push_a1_cell_ref_row1(2, 1, false, false, &mut out);
        assert_eq!(out, "B2");

        let mut out = String::new();
        push_a1_cell_ref(0, 0, false, false, &mut out);
        assert_eq!(out, "A1");
    }

    #[test]
    fn push_a1_row_and_col_refs_respect_absolute_flags() {
        let mut out = String::new();
        push_a1_row_ref_row1(7, true, &mut out);
        assert_eq!(out, "$7");

        let mut out = String::new();
        push_a1_row_ref_row1(7, false, &mut out);
        assert_eq!(out, "7");

        let mut out = String::new();
        push_a1_col_ref(0, true, &mut out);
        assert_eq!(out, "$A");

        let mut out = String::new();
        push_a1_col_ref(0, false, &mut out);
        assert_eq!(out, "A");
    }

    #[test]
    fn push_a1_row_and_col_ranges_respect_absolute_flags() {
        let mut out = String::new();
        push_a1_row_range_row1(1, 10, true, &mut out);
        assert_eq!(out, "$1:$10");

        let mut out = String::new();
        push_a1_row_range_row1(1, 10, false, &mut out);
        assert_eq!(out, "1:10");

        let mut out = String::new();
        push_a1_col_range(0, 3, true, &mut out);
        assert_eq!(out, "$A:$D");

        let mut out = String::new();
        push_a1_col_range(0, 3, false, &mut out);
        assert_eq!(out, "A:D");
    }

    #[test]
    fn push_a1_cell_range_row1_omits_colon_for_single_cell() {
        let mut out = String::new();
        push_a1_cell_range_row1(1, 0, 1, 0, true, true, &mut out);
        assert_eq!(out, "$A$1");
    }

    #[test]
    fn push_a1_cell_range_uses_row0_and_formats_like_to_a1() {
        let mut out = String::new();
        push_a1_cell_range(0, 0, 0, 0, false, false, &mut out);
        assert_eq!(out, "A1");

        let mut out = String::new();
        push_a1_cell_range(0, 0, 1, 1, false, false, &mut out);
        assert_eq!(out, "A1:B2");
    }

    #[test]
    fn push_a1_cell_range_row1_formats_two_cell_endpoints() {
        let mut out = String::new();
        push_a1_cell_range_row1(1, 0, 2, 1, true, true, &mut out);
        assert_eq!(out, "$A$1:$B$2");
    }

    #[test]
    fn push_a1_cell_area_row1_supports_per_endpoint_absolute_flags() {
        let mut out = String::new();
        push_a1_cell_area_row1(1, 0, true, false, 2, 1, false, true, &mut out);
        assert_eq!(out, "$A1:B$2");
    }

    #[test]
    fn column_label_to_index_accepts_excel_bounds() {
        assert_eq!(column_label_to_index("A").unwrap(), 0);
        assert_eq!(column_label_to_index("XFD").unwrap(), 16_383);
        assert!(column_label_to_index("XFE").is_err());
        assert!(column_label_to_index("").is_err());
        assert!(column_label_to_index("A0").is_err());
    }

    #[test]
    fn column_label_to_index_lenient_accepts_out_of_bounds_labels() {
        assert_eq!(column_label_to_index_lenient("A").unwrap(), 0);
        assert_eq!(column_label_to_index_lenient("XFD").unwrap(), 16_383);
        assert_eq!(column_label_to_index_lenient("XFE").unwrap(), 16_384);
        assert_eq!(column_label_to_index_lenient("ZZZ").unwrap(), 18_277);
        assert!(column_label_to_index_lenient("AAAA").is_err());
        assert!(column_label_to_index_lenient("").is_err());
        assert!(column_label_to_index_lenient("A0").is_err());
    }

    #[test]
    fn parse_a1_endpoint_parses_cell_row_and_col_refs() {
        assert_eq!(
            parse_a1_endpoint("A1").unwrap(),
            A1Endpoint::Cell(CellRef::new(0, 0))
        );
        assert_eq!(
            parse_a1_endpoint("$B$2").unwrap(),
            A1Endpoint::Cell(CellRef::new(1, 1))
        );
        assert_eq!(parse_a1_endpoint("$A").unwrap(), A1Endpoint::Col(0));
        assert_eq!(parse_a1_endpoint("1").unwrap(), A1Endpoint::Row(0));
        assert_eq!(parse_a1_endpoint("$7").unwrap(), A1Endpoint::Row(6));
    }

    #[test]
    fn parse_a1_endpoint_rejects_invalid_inputs() {
        assert!(parse_a1_endpoint("").is_err());
        assert!(parse_a1_endpoint("A0").is_err());
        assert!(parse_a1_endpoint("XFE1").is_err());
        assert!(parse_a1_endpoint("A1B").is_err());
        assert!(parse_a1_endpoint("1A").is_err());
    }
}
