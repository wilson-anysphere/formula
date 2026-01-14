use core::fmt;

use serde::{Deserialize, Serialize};

use crate::{A1ParseError, CellRef, Range, RangeParseError};

pub(crate) fn default_zoom() -> f32 {
    1.0
}

pub(crate) fn is_default_zoom(z: &f32) -> bool {
    (*z - 1.0).abs() < f32::EPSILON
}

fn is_true(b: &bool) -> bool {
    *b
}

/// Workbook-level view state (Excel `workbookView`).
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize, Default)]
pub struct WorkbookView {
    /// Currently active sheet tab.
    ///
    /// Excel stores this as `activeTab` (0-based sheet index). We store the stable sheet id.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub active_sheet_id: Option<u32>,

    /// Optional workbook window metadata.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub window: Option<WorkbookWindow>,
}

impl WorkbookView {
    pub(crate) fn is_default(&self) -> bool {
        self == &Self::default()
    }
}

/// Workbook window geometry/state.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize, Default)]
pub struct WorkbookWindow {
    /// Window x position (units are implementation-defined; XLSX uses twips).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub x: Option<i32>,
    /// Window y position (units are implementation-defined; XLSX uses twips).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub y: Option<i32>,
    /// Window width (units are implementation-defined; XLSX uses twips).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width: Option<u32>,
    /// Window height (units are implementation-defined; XLSX uses twips).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub height: Option<u32>,
    /// Normal / minimized / maximized.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub state: Option<WorkbookWindowState>,
}

#[derive(Copy, Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum WorkbookWindowState {
    Normal,
    Minimized,
    Maximized,
}

/// Worksheet view state (Excel `sheetView`).
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct SheetView {
    /// Sheet pane / frozen/split state.
    #[serde(default)]
    pub pane: SheetPane,

    /// Current selection (active cell + ranges).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub selection: Option<SheetSelection>,

    /// Display gridlines.
    #[serde(
        default = "crate::serde_defaults::default_true",
        skip_serializing_if = "is_true"
    )]
    pub show_grid_lines: bool,

    /// Display row/column headings.
    #[serde(
        default = "crate::serde_defaults::default_true",
        skip_serializing_if = "is_true"
    )]
    pub show_headings: bool,

    /// Display zeros.
    #[serde(
        default = "crate::serde_defaults::default_true",
        skip_serializing_if = "is_true"
    )]
    pub show_zeros: bool,

    /// View zoom level (1.0 = 100%).
    #[serde(default = "default_zoom", skip_serializing_if = "is_default_zoom")]
    pub zoom: f32,
}

impl Default for SheetView {
    fn default() -> Self {
        Self {
            pane: SheetPane::default(),
            selection: None,
            show_grid_lines: true,
            show_headings: true,
            show_zeros: true,
            zoom: default_zoom(),
        }
    }
}

impl SheetView {
    pub(crate) fn is_default(&self) -> bool {
        self == &Self::default()
    }
}

/// Worksheet pane state (Excel `pane`).
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize, Default)]
pub struct SheetPane {
    /// Frozen pane row count (top).
    #[serde(default)]
    pub frozen_rows: u32,

    /// Frozen pane column count (left).
    #[serde(default)]
    pub frozen_cols: u32,

    /// Horizontal split position (non-freeze panes).
    ///
    /// XLSX stores split positions as a floating point offset in twips (1/20th of a point).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub x_split: Option<f32>,

    /// Vertical split position (non-freeze panes).
    ///
    /// XLSX stores split positions as a floating point offset in twips (1/20th of a point).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub y_split: Option<f32>,

    /// Top-left visible cell for the bottom-right pane (Excel `topLeftCell`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub top_left_cell: Option<CellRef>,
}

/// Active cell + selected ranges (Excel `selection`).
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct SheetSelection {
    /// Active cell (caret).
    #[serde(default = "default_active_cell")]
    pub active_cell: CellRef,

    /// Selected ranges. When empty, the selection is implied to be `active_cell`.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub ranges: Vec<Range>,
}

fn default_active_cell() -> CellRef {
    CellRef::new(0, 0)
}

impl Default for SheetSelection {
    fn default() -> Self {
        Self {
            active_cell: default_active_cell(),
            ranges: Vec::new(),
        }
    }
}

impl SheetSelection {
    pub fn new(active_cell: CellRef, ranges: Vec<Range>) -> Self {
        Self {
            active_cell,
            ranges,
        }
    }

    /// Excel `sqref` encoding (space-separated A1 ranges).
    pub fn sqref(&self) -> String {
        if self.ranges.is_empty() {
            return self.active_cell.to_a1();
        }
        format_sqref(&self.ranges)
    }

    /// Parse an Excel `sqref` string into a selection.
    pub fn from_sqref(active_cell: CellRef, sqref: &str) -> Result<Self, SqrefParseError> {
        let ranges = parse_sqref(sqref)?;
        Ok(Self {
            active_cell,
            ranges,
        })
    }
}

/// Convert 0-based model coordinates into Excel's A1 string (1-based rows, A..XFD columns).
pub fn cell_to_a1(row: u32, col: u32) -> String {
    CellRef::new(row, col).to_a1()
}

/// Parse an Excel A1 string (e.g. `B2`) into 0-based model coordinates.
pub fn a1_to_cell(a1: &str) -> Result<(u32, u32), A1ParseError> {
    let cell = CellRef::from_a1(a1)?;
    Ok((cell.row, cell.col))
}

/// Convert selection ranges into Excel's `sqref` payload (space-separated A1 ranges).
pub fn format_sqref(ranges: &[Range]) -> String {
    ranges
        .iter()
        .map(|r| r.to_string())
        .collect::<Vec<_>>()
        .join(" ")
}

/// Errors that can occur when parsing an Excel `sqref` selection range list.
#[derive(Debug)]
pub enum SqrefParseError {
    Empty,
    Range {
        input: String,
        source: RangeParseError,
    },
}

impl fmt::Display for SqrefParseError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            SqrefParseError::Empty => f.write_str("empty sqref"),
            SqrefParseError::Range { input, source } => {
                write!(f, "invalid range in sqref ({input}): {source}")
            }
        }
    }
}

impl std::error::Error for SqrefParseError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            SqrefParseError::Empty => None,
            SqrefParseError::Range { source, .. } => Some(source),
        }
    }
}

/// Parse an Excel `sqref` selection range list.
pub fn parse_sqref(sqref: &str) -> Result<Vec<Range>, SqrefParseError> {
    let s = sqref.trim();
    if s.is_empty() {
        return Err(SqrefParseError::Empty);
    }
    s.split_whitespace()
        .map(|part| {
            Range::from_a1(part).map_err(|e| SqrefParseError::Range {
                input: part.to_string(),
                source: e,
            })
        })
        .collect()
}
