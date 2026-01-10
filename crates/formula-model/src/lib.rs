//! `formula-model` defines the core in-memory spreadsheet data structures.
//!
//! The crate is intentionally self-contained so it can be reused by:
//! - the calculation engine (dependency graph, evaluation, etc.)
//! - `.xlsx` import/export layers
//! - Tauri/IPC and WASM boundaries via `serde` (JSON-safe schema)

mod address;
mod cell;
pub mod charts;
mod comments;
mod display;
pub mod conditional_formatting;
mod error;
pub mod rich_text;
mod outline;
mod formula_rewrite;
mod style;
pub mod table;
mod value;
mod workbook;
mod worksheet;

pub use address::{A1ParseError, CellRef, Range, RangeIter, RangeParseError};
pub use cell::{Cell, CellId, CellKey, EXCEL_MAX_COLS, EXCEL_MAX_ROWS};
pub use comments::{Comment, CommentAuthor, CommentKind, Mention, Reply, TimestampMs};
pub use display::{format_cell_display, CellDisplay};
pub use conditional_formatting::*;
pub use error::ErrorValue;
pub use outline::{HiddenState, Outline, OutlineAxis, OutlineEntry, OutlinePr};
pub use formula_rewrite::rewrite_sheet_names_in_formula;
pub use style::{
    Alignment, Border, BorderStyle, Color, Fill, Font, HorizontalAlignment, Style, StyleTable,
    VerticalAlignment,
};
pub use table::{
    AutoFilter, FilterColumn, SortCondition, SortState, Table, TableArea, TableColumn, TableStyleInfo,
};
pub use value::{ArrayValue, CellValue, RichText, SpillValue};
pub use workbook::{RenameSheetError, Workbook, WorkbookId};
pub use worksheet::{
    ColProperties, RowProperties, SheetVisibility, TabColor, Worksheet, WorksheetId,
};

/// Current serialization schema version.
///
/// This is embedded into [`Workbook`] to enable forward-compatible IPC payloads.
pub const SCHEMA_VERSION: u32 = 1;

