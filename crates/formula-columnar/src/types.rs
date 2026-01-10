#![forbid(unsafe_code)]

use std::sync::Arc;

/// Logical column type, matching the sheet/data-model layer.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum ColumnType {
    Number,
    String,
    Boolean,
    DateTime,
    Currency { scale: u8 },
    Percentage { scale: u8 },
}

impl Default for ColumnType {
    fn default() -> Self {
        Self::String
    }
}

/// A cell/scalar value used by the columnar engine.
///
/// Note: `String` values use `Arc<str>` to avoid per-cell allocations when backed by a dictionary.
#[derive(Clone, Debug, PartialEq)]
pub enum Value {
    Null,
    /// A generic number (float).
    Number(f64),
    Boolean(bool),
    String(Arc<str>),
    DateTime(i64),
    Currency(i64),
    Percentage(i64),
}

impl Value {
    pub fn is_null(&self) -> bool {
        matches!(self, Value::Null)
    }
}
