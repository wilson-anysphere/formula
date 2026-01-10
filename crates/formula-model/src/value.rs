use serde::{Deserialize, Serialize};

use crate::{CellRef, ErrorValue};

/// Versioned, JSON-friendly representation of a cell value.
///
/// The enum uses an explicit `{type, value}` tagged layout for stable IPC.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(tag = "type", content = "value", rename_all = "snake_case")]
pub enum CellValue {
    /// Empty / unset cell value.
    Empty,
    /// IEEE-754 double precision number.
    Number(f64),
    /// Plain string (not rich text).
    String(String),
    /// Boolean.
    Boolean(bool),
    /// Excel error value.
    Error(ErrorValue),
    /// Rich text value (stub).
    RichText(RichText),
    /// Array result (stub).
    Array(ArrayValue),
    /// Marker for a cell that is part of a spilled array (stub).
    Spill(SpillValue),
}

impl Default for CellValue {
    fn default() -> Self {
        CellValue::Empty
    }
}

impl CellValue {
    /// Returns true if the value is [`CellValue::Empty`].
    pub fn is_empty(&self) -> bool {
        matches!(self, CellValue::Empty)
    }
}

/// Stub representation of rich text.
///
/// Excel supports multiple formatting "runs" inside a single cell string.
/// This struct is a placeholder until full fidelity is implemented.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct RichText {
    pub text: String,
}

/// Stub representation of a dynamic array result.
///
/// For now this stores a 2D matrix. The calculation engine may later choose a
/// more compact representation.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct ArrayValue {
    /// 2D array in row-major order.
    pub data: Vec<Vec<CellValue>>,
}

/// Stub marker for cells that belong to a spilled range.
#[derive(Copy, Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub struct SpillValue {
    /// Origin cell containing the spilling formula.
    pub origin: CellRef,
}
