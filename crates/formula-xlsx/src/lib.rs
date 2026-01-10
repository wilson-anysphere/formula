//! XLSX/XLSM compatibility layer.
//!
//! The long-term project goal is a full-fidelity Excel compatibility layer. The
//! crate currently exposes multiple APIs:
//!
//! - [`XlsxPackage`]: low-level Open Packaging Convention (OPC) ZIP handling
//!   that preserves unknown parts and binary payloads like `xl/vbaProject.bin`
//!   byte-for-byte.
//! - [`read_workbook`]/[`write_workbook`]: a semantic importer/exporter for
//!   [`formula_model::Workbook`].
//! - [`XlsxDocument`]: a higher-fidelity round-trip representation that pairs a
//!   [`formula_model::Workbook`] with preserved parts plus enough metadata to
//!   rewrite core SpreadsheetML files without breaking relationship IDs or
//!   cached values.
//!
//! The module surface also contains focused parsers/writers for some other Excel
//! parts (shared strings with rich text, sheet metadata for tab order/colors,
//! pivot table metadata, etc.).

pub mod autofilter;
pub mod charts;
pub mod comments;
pub mod conditional_formatting;
mod compare;
mod openxml;
pub mod drawingml;
pub mod drawings;
pub mod hyperlinks;
pub mod outline;
mod package;
mod path;
pub mod pivots;
pub mod print;
mod read;
mod reader;
mod relationships;
pub mod shared_strings;
pub mod merge_cells;
pub mod minimal;
mod sheet_metadata;
pub mod styles;
pub mod tables;
pub mod vba;
mod workbook;
pub mod write;
mod writer;

use std::collections::{BTreeMap, HashMap};

pub use compare::*;
pub use conditional_formatting::*;
pub use hyperlinks::{
    parse_worksheet_hyperlinks, update_worksheet_relationships, update_worksheet_xml,
};
pub use package::{XlsxError, XlsxPackage};
pub use pivots::{
    pivot_charts::PivotChartPart,
    slicers::{PivotSlicerParts, SlicerDefinition, TimelineDefinition},
    PivotCacheDefinitionPart, PivotCacheRecordsPart, PivotTablePart, XlsxPivots,
};
pub use read::{load_from_bytes, load_from_path};
pub use reader::{read_workbook, read_workbook_from_reader};
pub use sheet_metadata::{
    parse_sheet_tab_color, parse_workbook_sheets, write_sheet_tab_color, write_workbook_sheets,
    WorkbookSheetInfo,
};
pub use styles::*;
pub use crate::minimal::write_minimal_xlsx;
pub use workbook::ChartExtractionError;
pub use writer::{write_workbook, write_workbook_to_writer, XlsxWriteError};

use formula_model::rich_text::RichText;
use formula_model::{CellRef, Workbook, WorksheetId};

/// Excel date system used to interpret serialized dates.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DateSystem {
    /// The default Excel 1900 date system (with the Lotus 1-2-3 leap year bug).
    V1900,
    /// The Excel 1904 date system.
    V1904,
}

impl Default for DateSystem {
    fn default() -> Self {
        Self::V1900
    }
}

#[derive(Debug, Clone, Default)]
pub struct CalcPr {
    pub calc_id: Option<String>,
    pub calc_mode: Option<String>,
    pub full_calc_on_load: Option<bool>,
}

#[derive(Debug, Clone)]
pub struct SheetMeta {
    pub worksheet_id: WorksheetId,
    pub sheet_id: u32,
    pub relationship_id: String,
    pub state: Option<String>,
    pub path: String,
}

#[derive(Debug, Clone, Default)]
pub struct FormulaMeta {
    pub file_text: String,
    pub t: Option<String>,
    pub reference: Option<String>,
    pub shared_index: Option<u32>,
    pub always_calc: Option<bool>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum CellValueKind {
    Number,
    SharedString { index: u32 },
    InlineString,
    Bool,
    Error,
    Str,
}

#[derive(Debug, Clone, Default)]
pub struct CellMeta {
    pub value_kind: Option<CellValueKind>,
    pub raw_value: Option<String>,
    pub formula: Option<FormulaMeta>,
}

#[derive(Debug, Clone, Default)]
pub struct XlsxMeta {
    pub date_system: DateSystem,
    pub calc_pr: CalcPr,
    pub sheets: Vec<SheetMeta>,
    pub cell_meta: HashMap<(WorksheetId, CellRef), CellMeta>,
}

/// A workbook paired with the original OPC package parts needed for high-fidelity round-trip.
#[derive(Debug, Clone)]
pub struct XlsxDocument {
    pub workbook: Workbook,
    /// Uncompressed bytes for every part in the OPC package.
    parts: BTreeMap<String, Vec<u8>>,
    /// Shared strings in the order they appeared in the file (if present).
    shared_strings: Vec<RichText>,
    meta: XlsxMeta,
}

impl XlsxDocument {
    pub fn new(workbook: Workbook) -> Self {
        let sheets = workbook
            .sheets
            .iter()
            .enumerate()
            .map(|(idx, sheet)| SheetMeta {
                worksheet_id: sheet.id,
                sheet_id: (idx + 1) as u32,
                relationship_id: format!("rId{}", idx + 1),
                state: None,
                path: format!("xl/worksheets/sheet{}.xml", idx + 1),
            })
            .collect();

        Self {
            workbook,
            parts: BTreeMap::new(),
            shared_strings: Vec::new(),
            meta: XlsxMeta {
                sheets,
                ..XlsxMeta::default()
            },
        }
    }

    pub fn parts(&self) -> &BTreeMap<String, Vec<u8>> {
        &self.parts
    }

    pub fn save_to_vec(&self) -> Result<Vec<u8>, write::WriteError> {
        write::write_to_vec(self)
    }
}
