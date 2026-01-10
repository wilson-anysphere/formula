use chrono::NaiveDateTime;
use serde::{Deserialize, Serialize};
use thiserror::Error;

/// A minimal cell value representation used by the sort/filter subsystem.
///
/// In the full spreadsheet engine this will likely be replaced by a richer value type that
/// carries formatting metadata (important for Excel-compatible filtering on displayed values).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum CellValue {
    Blank,
    Number(f64),
    Text(String),
    Bool(bool),
    DateTime(NaiveDateTime),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub struct RangeRef {
    pub start_row: usize,
    pub start_col: usize,
    pub end_row: usize,
    pub end_col: usize,
}

impl RangeRef {
    pub fn width(&self) -> usize {
        self.end_col.saturating_sub(self.start_col) + 1
    }

    pub fn height(&self) -> usize {
        self.end_row.saturating_sub(self.start_row) + 1
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HeaderOption {
    None,
    HasHeader,
    /// Attempt to detect a header row using a heuristic similar to Excel's "My data has headers"
    /// suggestion.
    Auto,
}

#[derive(Debug, Clone, PartialEq)]
pub struct RangeData {
    pub range: RangeRef,
    /// Row-major data for the range.
    pub rows: Vec<Vec<CellValue>>,
}

#[derive(Debug, Error, PartialEq, Eq)]
pub enum RangeDataError {
    #[error("range width is {range_width} but row {row_index} has width {row_width}")]
    RowWidthMismatch {
        range_width: usize,
        row_index: usize,
        row_width: usize,
    },
    #[error("range height is {range_height} but got {row_count} rows")]
    RowCountMismatch {
        range_height: usize,
        row_count: usize,
    },
}

impl RangeData {
    pub fn new(range: RangeRef, rows: Vec<Vec<CellValue>>) -> Result<Self, RangeDataError> {
        let range_width = range.width();
        let range_height = range.height();

        if rows.len() != range_height {
            return Err(RangeDataError::RowCountMismatch {
                range_height,
                row_count: rows.len(),
            });
        }

        for (row_index, row) in rows.iter().enumerate() {
            if row.len() != range_width {
                return Err(RangeDataError::RowWidthMismatch {
                    range_width,
                    row_index,
                    row_width: row.len(),
                });
            }
        }

        Ok(Self { range, rows })
    }

    pub fn width(&self) -> usize {
        self.range.width()
    }

    pub fn height(&self) -> usize {
        self.rows.len()
    }

    pub fn get(&self, row: usize, col: usize) -> Option<&CellValue> {
        self.rows.get(row).and_then(|r| r.get(col))
    }
}

