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
/// This supports up to `u32::MAX` rows and Excel's maximum column count while
/// keeping the key within JavaScript's safe integer range.
#[derive(Copy, Clone, Debug, PartialEq, Eq, Hash, PartialOrd, Ord, Serialize)]
#[repr(transparent)]
pub struct CellKey(u64);

impl CellKey {
    /// Encode a `(row, col)` coordinate into a compact [`CellKey`].
    #[inline]
    pub fn new(row: u32, col: u32) -> Self {
        assert!(col < EXCEL_MAX_COLS, "col out of Excel bounds: {col}");
        Self(((row as u64) << COL_BITS) | (col as u64))
    }

    /// Checked version of [`CellKey::new`].
    ///
    /// Returns `None` if `col` is outside Excel's sheet bounds.
    #[inline]
    pub fn try_new(row: u32, col: u32) -> Option<Self> {
        if col < EXCEL_MAX_COLS {
            Some(Self(((row as u64) << COL_BITS) | (col as u64)))
        } else {
            None
        }
    }

    /// Attempt to construct a [`CellKey`] from a packed `u64`.
    ///
    /// Returns `None` if:
    /// - the decoded row does not fit within `u32`
    /// - the decoded column is outside Excel's sheet bounds
    #[inline]
    pub fn try_from_u64(raw: u64) -> Option<Self> {
        let row = raw >> COL_BITS;
        let col = raw & COL_MASK;
        if row <= u32::MAX as u64 && col < EXCEL_MAX_COLS as u64 {
            Some(CellKey(raw))
        } else {
            None
        }
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

        if row > u32::MAX as u64 {
            return Err(D::Error::custom(format!(
                "CellKey row out of bounds: {row}"
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
    ///
    /// Invariant: stored **without** a leading `'='` (e.g. `"SUM(A1:A3)"`, not
    /// `"=SUM(A1:A3)"`). Use [`crate::display_formula_text`] when rendering a
    /// formula for UI display.
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        deserialize_with = "deserialize_formula_opt"
    )]
    pub formula: Option<String>,

    /// Excel phonetic guide (furigana) text associated with this cell.
    ///
    /// This is used to power Excel-compatible functions like `PHONETIC()`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub phonetic: Option<String>,

    /// Index into the workbook style table.
    #[serde(default)]
    pub style_id: u32,
}

impl Default for Cell {
    fn default() -> Self {
        Self {
            value: CellValue::Empty,
            formula: None,
            phonetic: None,
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
        self.value == CellValue::Empty
            && self.formula.is_none()
            && self.phonetic.is_none()
            && self.style_id == 0
    }

    /// Returns the phonetic guide text (furigana) for this cell, if any.
    pub fn phonetic_text(&self) -> Option<&str> {
        self.phonetic.as_deref()
    }
}

fn deserialize_formula_opt<'de, D>(deserializer: D) -> Result<Option<String>, D::Error>
where
    D: Deserializer<'de>,
{
    let raw = Option::<String>::deserialize(deserializer)?;
    Ok(raw.and_then(|s| crate::normalize_formula_text(&s)))
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

        let key3 = CellKey::new(EXCEL_MAX_ROWS + 5, 0);
        assert_eq!(key3.row(), EXCEL_MAX_ROWS + 5);
    }

    #[test]
    fn cell_key_deserialize_validates_bounds() {
        let too_large = (((u32::MAX as u64) + 1) << COL_BITS) | 0;
        let json = too_large.to_string();
        let err = serde_json::from_str::<CellKey>(&json).unwrap_err();
        let msg = err.to_string();
        assert!(msg.contains("row out of bounds"));
    }

    #[test]
    fn cell_key_try_new_enforces_excel_bounds() {
        assert!(CellKey::try_new(0, EXCEL_MAX_COLS).is_none());
        assert!(CellKey::try_new(0, 0).is_some());
        assert!(CellKey::try_new(u32::MAX, 0).is_some());
    }

    #[test]
    fn cell_key_try_from_u64_roundtrips() {
        let key = CellKey::new(123, 456);
        assert_eq!(CellKey::try_from_u64(key.as_u64()), Some(key));
        assert_eq!(
            CellKey::try_from_u64(((EXCEL_MAX_ROWS as u64) << COL_BITS) | 0),
            Some(CellKey::new(EXCEL_MAX_ROWS, 0))
        );
        assert!(CellKey::try_from_u64((((u32::MAX as u64) + 1) << COL_BITS) | 0).is_none());
    }
}
