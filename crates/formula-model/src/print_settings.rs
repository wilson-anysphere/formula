//! Print and page setup settings (Excel-compatible).
//!
//! In the XLSX file format these settings are stored in two different places:
//! - Print area and print titles are represented as workbook defined names:
//!   - `_xlnm.Print_Area`
//!   - `_xlnm.Print_Titles`
//!   with a `localSheetId` that points at the worksheet index.
//! - Page setup, margins, scaling, and manual page breaks live in each worksheet XML.
//!
//! `formula-model` treats the structs in this module as the canonical in-memory
//! representation. Import/export layers should translate between these structs
//! and the underlying XLSX representation.
//!
//! ## Indexing / conversion notes
//! The core model uses **0-based** coordinates (see [`crate::CellRef`] and
//! [`crate::Range`]):
//! - row `0` = Excel row `1`
//! - col `0` = Excel column `A`
//!
//! XLSX stores print settings using **1-based** row/column numbers. When
//! round-tripping:
//! - Print areas: add 1 to each row/col when formatting A1 references.
//! - Print titles (`RowRange`/`ColRange`): add 1 when formatting A1 row/col ranges.
//! - Manual page breaks: XLSX stores `brk/@id` as the **1-based** row/col number
//!   after which the break occurs. The model stores these values as **0-based**
//!   indices; add 1 when writing and subtract 1 when reading.

use std::collections::BTreeSet;

use serde::{Deserialize, Serialize};

use crate::Range;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum Orientation {
    Portrait,
    Landscape,
}

impl Default for Orientation {
    fn default() -> Self {
        Self::Portrait
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(default)]
pub struct PaperSize {
    /// OpenXML `ST_PaperSize` numeric code (e.g. `1` = Letter, `9` = A4).
    pub code: u16,
}

impl PaperSize {
    pub const LETTER: Self = Self { code: 1 };
    pub const A4: Self = Self { code: 9 };

    pub fn dimensions_in_inches(self) -> (f64, f64) {
        match self.code {
            1 => (8.5, 11.0),                    // Letter
            9 => (8.267_716_535, 11.692_913_39), // A4 (210mm x 297mm)
            _ => (8.5, 11.0),                    // fallback
        }
    }
}

impl Default for PaperSize {
    fn default() -> Self {
        Self::LETTER
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize)]
#[serde(default)]
pub struct PageMargins {
    /// Inches.
    pub left: f64,
    pub right: f64,
    pub top: f64,
    pub bottom: f64,
    pub header: f64,
    pub footer: f64,
}

impl Default for PageMargins {
    fn default() -> Self {
        Self {
            left: 0.7,
            right: 0.7,
            top: 0.75,
            bottom: 0.75,
            header: 0.3,
            footer: 0.3,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum Scaling {
    Percent(u16),
    FitTo { width: u16, height: u16 },
}

impl Default for Scaling {
    fn default() -> Self {
        Self::Percent(100)
    }
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(default)]
pub struct PageSetup {
    pub orientation: Orientation,
    pub paper_size: PaperSize,
    pub margins: PageMargins,
    pub scaling: Scaling,
}

impl Default for PageSetup {
    fn default() -> Self {
        Self {
            orientation: Orientation::default(),
            paper_size: PaperSize::default(),
            margins: PageMargins::default(),
            scaling: Scaling::default(),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(default)]
pub struct RowRange {
    /// 0-based row index (inclusive).
    pub start: u32,
    /// 0-based row index (inclusive).
    pub end: u32,
}

impl RowRange {
    pub fn normalized(self) -> Self {
        Self {
            start: self.start.min(self.end),
            end: self.start.max(self.end),
        }
    }
}

impl Default for RowRange {
    fn default() -> Self {
        Self { start: 0, end: 0 }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(default)]
pub struct ColRange {
    /// 0-based column index (inclusive).
    pub start: u32,
    /// 0-based column index (inclusive).
    pub end: u32,
}

impl ColRange {
    pub fn normalized(self) -> Self {
        Self {
            start: self.start.min(self.end),
            end: self.start.max(self.end),
        }
    }
}

impl Default for ColRange {
    fn default() -> Self {
        Self { start: 0, end: 0 }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default, Serialize, Deserialize)]
#[serde(default)]
pub struct PrintTitles {
    pub repeat_rows: Option<RowRange>,
    pub repeat_cols: Option<ColRange>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(default)]
pub struct ManualPageBreaks {
    /// 0-based row indices after which a manual break occurs.
    #[serde(default, skip_serializing_if = "BTreeSet::is_empty")]
    pub row_breaks_after: BTreeSet<u32>,
    /// 0-based column indices after which a manual break occurs.
    #[serde(default, skip_serializing_if = "BTreeSet::is_empty")]
    pub col_breaks_after: BTreeSet<u32>,
}

impl ManualPageBreaks {
    pub fn is_empty(&self) -> bool {
        self.row_breaks_after.is_empty() && self.col_breaks_after.is_empty()
    }
}

impl Default for ManualPageBreaks {
    fn default() -> Self {
        Self {
            row_breaks_after: BTreeSet::new(),
            col_breaks_after: BTreeSet::new(),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(default)]
pub struct SheetPrintSettings {
    pub sheet_name: String,

    /// Print area for the sheet.
    ///
    /// This is stored as a list of rectangular ranges. Coordinates are 0-based
    /// (`Range`/`CellRef`), unlike XLSX which uses 1-based A1 references.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub print_area: Option<Vec<Range>>,

    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub print_titles: Option<PrintTitles>,

    pub page_setup: PageSetup,

    pub manual_page_breaks: ManualPageBreaks,
}

impl SheetPrintSettings {
    pub fn new(sheet_name: impl Into<String>) -> Self {
        Self {
            sheet_name: sheet_name.into(),
            print_area: None,
            print_titles: None,
            page_setup: PageSetup::default(),
            manual_page_breaks: ManualPageBreaks::default(),
        }
    }

    pub fn is_default(&self) -> bool {
        self.print_area.is_none()
            && self.print_titles.is_none()
            && self.page_setup == PageSetup::default()
            && self.manual_page_breaks.is_empty()
    }
}

impl Default for SheetPrintSettings {
    fn default() -> Self {
        Self::new(String::new())
    }
}

#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
#[serde(default)]
pub struct WorkbookPrintSettings {
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub sheets: Vec<SheetPrintSettings>,
}

impl WorkbookPrintSettings {
    pub fn is_empty(&self) -> bool {
        self.sheets.is_empty()
    }
}
