use serde::de::Error as _;
use serde::{Deserialize, Deserializer, Serialize};

use crate::{CellRef, CellValue};

/// Excel-compatible maximum rows per worksheet (1,048,576).
pub const EXCEL_MAX_ROWS: u32 = 1_048_576;

/// Excel-compatible maximum columns per worksheet (16,384).
pub const EXCEL_MAX_COLS: u32 = 16_384;

const COL_BITS: u32 = 14; // 2^14 = 16,384 columns.
const COL_MASK: u64 = (1u64 << COL_BITS) - 1;

/// Compact key used for sparse cell storage.
///
/// The key is a packed `(row, col)` pair into a `u64`:
///
/// ```text
/// key = (row << 14) | col
/// ```
///
/// This supports Excel's maximum dimensions while keeping the key within 34 bits
/// (JSON-safe for JavaScript numbers).
#[derive(Copy, Clone, Debug, PartialEq, Eq, Hash, PartialOrd, Ord, Serialize)]
#[repr(transparent)]
pub struct CellKey(u64);

impl CellKey {
    /// Encode a `(row, col)` coordinate into a compact [`CellKey`].
    #[inline]
    pub fn new(row: u32, col: u32) -> Self {
        assert!(row < EXCEL_MAX_ROWS, "row out of Excel bounds: {row}");
        assert!(col < EXCEL_MAX_COLS, "col out of Excel bounds: {col}");
        Self(((row as u64) << COL_BITS) | (col as u64))
    }

    /// Decode the row component (0-indexed).
    #[inline]
    pub const fn row(self) -> u32 {
        (self.0 >> COL_BITS) as u32
    }

    /// Decode the column component (0-indexed).
    #[inline]
    pub const fn col(self) -> u32 {
        (self.0 & COL_MASK) as u32
    }

    /// Raw packed value.
    #[inline]
    pub const fn as_u64(self) -> u64 {
        self.0
    }

    /// Convert to a [`CellRef`].
    #[inline]
    pub const fn to_ref(self) -> CellRef {
        CellRef::new(self.row(), self.col())
    }

    /// Create a key from a [`CellRef`].
    #[inline]
    pub fn from_ref(cell: CellRef) -> Self {
        Self::new(cell.row, cell.col)
    }
}

impl<'de> Deserialize<'de> for CellKey {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        let raw = u64::deserialize(deserializer)?;
        let row = raw >> COL_BITS;
        let col = raw & COL_MASK;

        if row >= EXCEL_MAX_ROWS as u64 {
            return Err(D::Error::custom(format!(
                "CellKey row out of Excel bounds: {row}"
            )));
        }
        if col >= EXCEL_MAX_COLS as u64 {
            return Err(D::Error::custom(format!(
                "CellKey col out of Excel bounds: {col}"
            )));
        }

        Ok(CellKey(raw))
    }
}

impl From<CellKey> for u64 {
    fn from(value: CellKey) -> Self {
        value.0
    }
}

impl From<CellRef> for CellKey {
    fn from(value: CellRef) -> Self {
        Self::from_ref(value)
    }
}

/// Address of a cell within a workbook.
#[derive(Copy, Clone, Debug, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub struct CellId {
    /// Worksheet identifier.
    pub sheet_id: u32,
    /// Cell coordinates within the worksheet.
    pub cell: CellRef,
}

impl CellId {
    /// Create a new [`CellId`].
    pub const fn new(sheet_id: u32, row: u32, col: u32) -> Self {
        Self {
            sheet_id,
            cell: CellRef::new(row, col),
        }
    }
}

/// A single cell record.
///
/// Cells are stored sparsely: when a cell is "truly empty" (no value, no formula,
/// default style), it is removed from the worksheet map.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct Cell {
    /// The cell's computed (or literal) value.
    #[serde(default)]
    pub value: CellValue,

    /// Formula text, if the cell contains a formula.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formula: Option<String>,

    /// Index into the workbook style table.
    #[serde(default)]
    pub style_id: u32,
}

impl Default for Cell {
    fn default() -> Self {
        Self {
            value: CellValue::Empty,
            formula: None,
            style_id: 0,
        }
    }
}

impl Cell {
    /// Create a new cell with the given value.
    pub fn new(value: CellValue) -> Self {
        Self {
            value,
            ..Self::default()
        }
    }

    /// Returns true if this cell has no observable content or formatting.
    ///
    /// Such cells should not be stored in the sparse map.
    pub fn is_truly_empty(&self) -> bool {
        self.value == CellValue::Empty && self.formula.is_none() && self.style_id == 0
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn cell_key_roundtrip() {
        let key = CellKey::new(0, 0);
        assert_eq!(key.row(), 0);
        assert_eq!(key.col(), 0);
        assert_eq!(key.to_ref(), CellRef::new(0, 0));

        let key2 = CellKey::new(EXCEL_MAX_ROWS - 1, EXCEL_MAX_COLS - 1);
        assert_eq!(key2.row(), EXCEL_MAX_ROWS - 1);
        assert_eq!(key2.col(), EXCEL_MAX_COLS - 1);
    }

    #[test]
    fn cell_key_deserialize_validates_bounds() {
        let too_large = ((EXCEL_MAX_ROWS as u64) << COL_BITS) | 0;
        let json = too_large.to_string();
        let err = serde_json::from_str::<CellKey>(&json).unwrap_err();
        let msg = err.to_string();
        assert!(msg.contains("out of Excel bounds"));
    }
}
