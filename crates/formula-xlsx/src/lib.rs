//! Minimal XLSX/XLSM package handling focused on macro preservation.
//!
//! The long-term project goal is a full-fidelity Excel compatibility layer.
//! For now we implement just enough OPC plumbing to:
//! - Load an XLSX/XLSM ZIP archive.
//! - Preserve unknown parts byte-for-byte.
//! - Preserve `xl/vbaProject.bin` exactly on write.
//! - Optionally parse `vbaProject.bin` to expose modules for UI display.
//!
//! This crate also includes helpers for worksheet metadata needed for an
//! Excel-like sheet tab experience (sheet order, visibility, tab colors).

mod package;
mod path;
mod relationships;
mod workbook;

pub mod charts;
pub mod comments;
pub mod drawingml;
pub mod outline;
mod sheet_metadata;
pub mod pivots;
pub mod print;
pub mod shared_strings;
pub mod vba;

pub mod conditional_formatting;
pub mod styles;

pub use conditional_formatting::*;
pub use package::{XlsxError, XlsxPackage};
pub use pivots::{PivotCacheDefinitionPart, PivotCacheRecordsPart, PivotTablePart, XlsxPivots};
pub use sheet_metadata::{
    parse_sheet_tab_color, parse_workbook_sheets, write_sheet_tab_color, write_workbook_sheets,
    WorkbookSheetInfo,
};
pub use styles::*;
pub use workbook::ChartExtractionError;
