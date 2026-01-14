mod a1;
mod convert;
mod page_breaks;
mod pdf;
pub(crate) mod xlsx;

pub use a1::{
    format_print_area_defined_name, format_print_titles_defined_name,
    parse_print_area_defined_name, parse_print_titles_defined_name, CellRange, ColRange,
    PrintTitles, RowRange,
};
pub use page_breaks::{calculate_pages, Page};
pub use pdf::export_range_to_pdf_bytes;
pub(crate) use xlsx::parse_worksheet_print_settings;
pub use xlsx::{
    read_workbook_print_settings, read_workbook_print_settings_from_reader, write_workbook_print_settings,
    write_workbook_print_settings_with_limit,
};

use std::collections::BTreeSet;

/// Default column width in points when no explicit width is available.
///
/// This matches the desktop shell's prior behavior when it padded `col_widths_points`.
pub const DEFAULT_COL_WIDTH_POINTS: f64 = 64.0;

/// Default row height in points when no explicit height is available.
///
/// This matches the desktop shell's prior behavior when it padded `row_heights_points`.
pub const DEFAULT_ROW_HEIGHT_POINTS: f64 = 20.0;

#[derive(Debug, thiserror::Error)]
pub enum PrintError {
    #[error("io error: {0}")]
    Io(#[from] std::io::Error),

    #[error("zip error: {0}")]
    Zip(#[from] zip::result::ZipError),

    #[error("xml error: {0}")]
    Xml(#[from] quick_xml::Error),

    #[error("xml attribute error: {0}")]
    XmlAttr(#[from] quick_xml::events::attributes::AttrError),

    #[error("utf8 error: {0}")]
    Utf8(#[from] std::str::Utf8Error),

    #[error("invalid A1 reference: {0}")]
    InvalidA1(String),

    #[error("missing required xlsx part: {0}")]
    MissingPart(&'static str),

    #[error("xlsx part '{part}' is too large ({size} bytes, max {max})")]
    PartTooLarge { part: String, size: u64, max: u64 },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Orientation {
    Portrait,
    Landscape,
}

impl Default for Orientation {
    fn default() -> Self {
        Self::Portrait
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
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

#[derive(Debug, Clone, Copy, PartialEq)]
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

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Scaling {
    Percent(u16),
    FitTo { width: u16, height: u16 },
}

impl Default for Scaling {
    fn default() -> Self {
        Self::Percent(100)
    }
}

#[derive(Debug, Clone, PartialEq)]
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

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ManualPageBreaks {
    /// Row numbers (1-based) after which a manual break occurs (OpenXML `brk/@id`).
    pub row_breaks_after: BTreeSet<u32>,
    /// Column numbers (1-based) after which a manual break occurs (OpenXML `brk/@id`).
    pub col_breaks_after: BTreeSet<u32>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct SheetPrintSettings {
    pub sheet_name: String,
    pub print_area: Option<Vec<CellRange>>,
    pub print_titles: Option<PrintTitles>,
    pub page_setup: PageSetup,
    pub manual_page_breaks: ManualPageBreaks,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct WorkbookPrintSettings {
    pub sheets: Vec<SheetPrintSettings>,
}
