use serde::{Deserialize, Serialize};

pub use crate::rich_text::RichText;
use crate::{CellRef, ErrorValue};
use std::fmt;

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
    /// Excel rich value (data type) representing an entity.
    ///
    /// This is a JSON-friendly representation of Excel's "entity" rich values
    /// (e.g. Stocks / Geography) where the primary scalar representation is a
    /// display string.
    Entity(EntityValue),
    /// Excel rich value representing a record.
    ///
    /// Record values are treated similarly to entities for scalar export: they
    /// degrade to their display string in legacy IO paths.
    Record(RecordValue),
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

impl From<f64> for CellValue {
    fn from(value: f64) -> Self {
        CellValue::Number(value)
    }
}

impl From<bool> for CellValue {
    fn from(value: bool) -> Self {
        CellValue::Boolean(value)
    }
}

impl From<String> for CellValue {
    fn from(value: String) -> Self {
        CellValue::String(value)
    }
}

impl From<&str> for CellValue {
    fn from(value: &str) -> Self {
        CellValue::String(value.to_string())
    }
}

impl From<ErrorValue> for CellValue {
    fn from(value: ErrorValue) -> Self {
        CellValue::Error(value)
    }
}

impl From<RichText> for CellValue {
    fn from(value: RichText) -> Self {
        CellValue::RichText(value)
    }
}

/// Stub representation of an Excel rich "entity" value.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct EntityValue {
    /// User-visible string representation (what Excel renders in the grid).
    pub display: String,
}

impl EntityValue {
    pub fn new(display: impl Into<String>) -> Self {
        Self {
            display: display.into(),
        }
    }
}

impl fmt::Display for EntityValue {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(&self.display)
    }
}

impl From<EntityValue> for CellValue {
    fn from(value: EntityValue) -> Self {
        CellValue::Entity(value)
    }
}

/// Stub representation of an Excel rich "record" value.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct RecordValue {
    /// User-visible string representation (what Excel renders in the grid).
    pub display: String,
}

impl RecordValue {
    pub fn new(display: impl Into<String>) -> Self {
        Self {
            display: display.into(),
        }
    }
}

impl fmt::Display for RecordValue {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(&self.display)
    }
}

impl From<RecordValue> for CellValue {
    fn from(value: RecordValue) -> Self {
        CellValue::Record(value)
    }
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
