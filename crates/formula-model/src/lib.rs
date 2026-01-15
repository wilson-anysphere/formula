//! `formula-model` defines the core in-memory spreadsheet data structures.
//!
//! The crate is intentionally self-contained so it can be reused by:
//! - the calculation engine (dependency graph, evaluation, etc.)
//! - `.xlsx` import/export layers
//! - Tauri/IPC and WASM boundaries via `serde` (JSON-safe schema)

mod address;
pub mod autofilter;
pub mod calc_settings;
mod cell;
pub mod charts;
mod comments;
pub mod conditional_formatting;
mod data_validation;
mod date_system;
mod display;
/// Drawing primitives (images, shapes, charts, etc.).
pub mod drawings;
mod error;
pub mod external_refs;
pub mod formula_rewrite;
mod formula_text;
mod hyperlinks;
pub mod import;
mod merge;
mod names;
mod outline;
pub mod pivots;
mod print_settings;
mod protection;
pub mod rich_text;
mod serde_defaults;
mod sheet_name;
mod style;
pub mod table;
mod theme;
mod value;
mod view;
mod workbook;
mod worksheet;

pub use address::{A1ParseError, CellRef, Range, RangeIter, RangeParseError};
pub use autofilter::{
    DateComparison, FilterCriterion, FilterJoin, FilterValue, NumberComparison, OpaqueCustomFilter,
    OpaqueDynamicFilter, SheetAutoFilter, TextMatch, TextMatchKind,
};
pub use calc_settings::{CalcSettings, CalculationMode, IterativeCalculationSettings};
pub use cell::{Cell, CellId, CellKey, EXCEL_MAX_COLS, EXCEL_MAX_ROWS};
pub use comments::{
    Comment, CommentAuthor, CommentError, CommentKind, CommentPatch, Mention, Reply, TimestampMs,
};
pub use conditional_formatting::*;
pub use data_validation::*;
pub use date_system::DateSystem;
pub use display::{format_cell_display, format_cell_display_in_workbook, CellDisplay};
pub use error::ErrorValue;
pub use formula_rewrite::{
    rewrite_deleted_sheet_references_in_formula, rewrite_sheet_names_in_formula,
    rewrite_table_names_in_formula,
};
pub use formula_text::{display_formula_text, normalize_formula_text};
pub use hyperlinks::{Hyperlink, HyperlinkTarget};
pub use merge::{MergeError, MergedRegion, MergedRegions};
pub use names::{
    validate_defined_name, DefinedName, DefinedNameError, DefinedNameId, DefinedNameScope,
    DefinedNameValidationError, EXCEL_DEFINED_NAME_MAX_LEN, XLNM_FILTER_DATABASE, XLNM_PRINT_AREA,
    XLNM_PRINT_TITLES,
};
pub use outline::{HiddenState, Outline, OutlineAxis, OutlineEntry, OutlinePr};
pub use print_settings::{
    ColRange, ManualPageBreaks, Orientation, PageMargins, PageSetup, PaperSize, PrintTitles,
    RowRange, Scaling, SheetPrintSettings, WorkbookPrintSettings,
};
pub use protection::{
    hash_legacy_password, verify_legacy_password, SheetProtection, SheetProtectionAction,
    WorkbookProtection,
};
pub use sheet_name::{
    sanitize_sheet_name, sheet_name_casefold, sheet_name_eq_case_insensitive, validate_sheet_name,
    SheetNameError, EXCEL_MAX_SHEET_NAME_LEN,
};
pub use style::{
    Alignment, Border, BorderEdge, BorderStyle, Color, Fill, FillPattern, Font,
    HorizontalAlignment, Protection, Style, StyleTable, VerticalAlignment,
};
pub use table::{
    validate_table_name, AutoFilter, FilterColumn, SortCondition, SortState, Table, TableArea,
    TableColumn, TableError, TableIdentifier, TableStyleInfo,
};
pub use theme::{
    indexed_color_argb, number_format_color, parse_number_format_color_token, resolve_color,
    resolve_color_in_context, resolve_number_format_color, ArgbColor, ColorContext, ThemeColorSlot,
    ThemePalette,
};
pub use value::{
    ArrayValue, CellValue, EntityValue, ImageValue, LinkedEntityValue, RecordValue, RichText,
    SpillValue,
};
pub use view::{
    a1_to_cell, cell_to_a1, format_sqref, parse_sqref, SheetPane, SheetSelection, SheetView,
    SqrefParseError, WorkbookView, WorkbookWindow, WorkbookWindowState,
};
pub use workbook::{DeleteSheetError, DuplicateSheetError, RenameSheetError, Workbook, WorkbookId};
pub use worksheet::{
    ColProperties, RangeBatch, RangeBatchBuffer, RangeBatchRef, RowProperties, SheetVisibility,
    TabColor, Worksheet, WorksheetId,
};

/// Current serialization schema version.
///
/// This is embedded into [`Workbook`] to enable forward-compatible IPC payloads.
pub const SCHEMA_VERSION: u32 = 4;

fn new_uuid() -> uuid::Uuid {
    #[cfg(not(target_arch = "wasm32"))]
    {
        uuid::Uuid::new_v4()
    }

    #[cfg(target_arch = "wasm32")]
    {
        use std::sync::atomic::{AtomicU64, Ordering};

        static COUNTER: AtomicU64 = AtomicU64::new(1);
        uuid::Uuid::from_u128(COUNTER.fetch_add(1, Ordering::Relaxed) as u128)
    }
}
