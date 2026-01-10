use std::collections::HashMap;
use std::fmt;

use serde::{Deserialize, Serialize};

/// Spreadsheet cell reference (currently an A1-style address string).
#[derive(Clone, Debug, PartialEq, Eq, Hash, PartialOrd, Ord, Serialize, Deserialize)]
#[serde(transparent)]
pub struct CellRef(String);

impl CellRef {
    pub fn new(address: impl Into<String>) -> Self {
        Self(address.into())
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl fmt::Display for CellRef {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.0.fmt(f)
    }
}

impl From<&str> for CellRef {
    fn from(value: &str) -> Self {
        Self::new(value)
    }
}

/// Minimal cell value representation used by What‑If tools.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(tag = "type", content = "value", rename_all = "snake_case")]
pub enum CellValue {
    Number(f64),
    Text(String),
    Bool(bool),
    Blank,
}

impl CellValue {
    pub fn as_number(&self) -> Option<f64> {
        match self {
            CellValue::Number(v) => Some(*v),
            _ => None,
        }
    }
}

/// Abstraction over the calculation engine used by What‑If tools.
///
/// In the full application this will be implemented by the spreadsheet model
/// (dependency graph + calculation). For unit tests we provide a minimal
/// in-memory implementation.
pub trait WhatIfModel {
    type Error;

    fn get_cell_value(&self, cell: &CellRef) -> Result<CellValue, Self::Error>;
    fn set_cell_value(&mut self, cell: &CellRef, value: CellValue) -> Result<(), Self::Error>;
    fn recalculate(&mut self) -> Result<(), Self::Error>;
}

#[derive(Debug)]
pub enum WhatIfError<E> {
    Model(E),
    NonNumericCell { cell: CellRef, value: CellValue },
    InvalidParams(&'static str),
    NoBracketFound,
    NumericalFailure(&'static str),
}

impl<E> From<E> for WhatIfError<E> {
    fn from(value: E) -> Self {
        WhatIfError::Model(value)
    }
}

impl<E: fmt::Display> fmt::Display for WhatIfError<E> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            WhatIfError::Model(err) => write!(f, "model error: {err}"),
            WhatIfError::NonNumericCell { cell, value } => {
                write!(f, "cell {cell} is not numeric: {value:?}")
            }
            WhatIfError::InvalidParams(msg) => write!(f, "invalid parameters: {msg}"),
            WhatIfError::NoBracketFound => write!(f, "could not bracket a solution"),
            WhatIfError::NumericalFailure(msg) => write!(f, "numerical failure: {msg}"),
        }
    }
}

impl<E: std::error::Error + 'static> std::error::Error for WhatIfError<E> {}

/// Convenience: a lightweight in-memory model for tests and examples.
#[derive(Clone, Debug, Default)]
pub struct InMemoryModel {
    values: HashMap<CellRef, CellValue>,
}

impl InMemoryModel {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn with_value(mut self, cell: impl Into<CellRef>, value: CellValue) -> Self {
        self.values.insert(cell.into(), value);
        self
    }

    pub fn values(&self) -> &HashMap<CellRef, CellValue> {
        &self.values
    }

    pub fn values_mut(&mut self) -> &mut HashMap<CellRef, CellValue> {
        &mut self.values
    }
}

impl WhatIfModel for InMemoryModel {
    type Error = &'static str;

    fn get_cell_value(&self, cell: &CellRef) -> Result<CellValue, Self::Error> {
        Ok(self.values.get(cell).cloned().unwrap_or(CellValue::Blank))
    }

    fn set_cell_value(&mut self, cell: &CellRef, value: CellValue) -> Result<(), Self::Error> {
        self.values.insert(cell.clone(), value);
        Ok(())
    }

    fn recalculate(&mut self) -> Result<(), Self::Error> {
        Ok(())
    }
}
