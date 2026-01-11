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
//! - [`WorkbookPackage`]: a focused round-trip wrapper used by the style
//!   pipeline (`styles.xml` + cell `s` indices).
//!
//! The module surface also contains focused parsers/writers for some other Excel
//! parts (shared strings with rich text, sheet metadata for tab order/colors,
//! pivot table metadata, etc.).

pub mod autofilter;
pub mod calc_settings;
pub mod charts;
pub mod comments;
mod compare;
mod content_types;
pub mod conditional_formatting;
mod formula_text;
mod model_package;
mod openxml;
mod preserve;
mod xml;
pub mod drawingml;
pub mod drawings;
pub mod hyperlinks;
pub mod merge_cells;
pub mod minimal;
pub mod outline;
pub mod theme;
mod package;
pub mod patch;
mod path;
pub mod pivots;
pub mod print;
mod recalc_policy;
mod read;
mod reader;
mod relationships;
pub mod shared_strings;
mod sheet_metadata;
pub mod streaming;
pub mod styles;
pub mod tables;
pub mod vba;
mod workbook;
pub mod write;
mod writer;

use std::collections::{BTreeMap, HashMap};

pub use crate::minimal::write_minimal_xlsx;
pub use calc_settings::CalcSettingsError;
pub use compare::*;
pub use conditional_formatting::*;
pub use hyperlinks::{
    parse_worksheet_hyperlinks, update_worksheet_relationships, update_worksheet_xml,
};
pub use model_package::{WorkbookPackage, WorkbookPackageError};
pub use package::{CellPatch as PackageCellPatch, CellPatchSheet, WorksheetPartInfo, XlsxError, XlsxPackage};
pub use patch::{CellPatch, CellStyleRef, WorkbookCellPatches, WorksheetCellPatches};
pub use pivots::{
    cache_records::{pivot_cache_datetime_to_naive_date, PivotCacheRecordsReader, PivotCacheValue},
    graph::{PivotTableInstance, XlsxPivotGraph},
    pivot_charts::PivotChartPart,
    slicers::{
        slicer_selection_to_row_filter, slicer_selection_to_row_filter_with_resolver,
        timeline_selection_to_row_filter, PivotSlicerParts, SlicerDefinition, SlicerSelectionState,
        TimelineDefinition, TimelineSelectionState,
    },
    PivotCacheDefinition, PivotCacheDefinitionPart, PivotCacheField, PivotCacheRecordsPart,
    PivotCacheSourceType, PivotTableDataField, PivotTableDefinition, PivotTableField, PivotTablePart,
    PivotTableStyleInfo, PreservedPivotParts, RelationshipStub, XlsxPivots,
};
pub use recalc_policy::RecalcPolicy;
pub use read::{load_from_bytes, load_from_path, read_workbook_model_from_bytes};
pub use reader::{read_workbook, read_workbook_from_reader};
pub use sheet_metadata::{
    parse_sheet_tab_color, parse_workbook_sheets, write_sheet_tab_color, write_workbook_sheets,
    WorkbookSheetInfo,
};
pub use streaming::{
    patch_xlsx_streaming, patch_xlsx_streaming_workbook_cell_patches,
    patch_xlsx_streaming_workbook_cell_patches_with_styles, StreamingPatchError,
    WorksheetCellPatch,
};
pub use styles::*;
pub use workbook::ChartExtractionError;
pub use writer::{write_workbook, write_workbook_to_writer, XlsxWriteError};
pub use xml::XmlDomError;

use formula_model::rich_text::RichText;
use formula_model::{CellRef, CellValue, Workbook, WorksheetId};

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

impl DateSystem {
    pub fn to_engine_date_system(self) -> formula_engine::date::ExcelDateSystem {
        match self {
            DateSystem::V1900 => formula_engine::date::ExcelDateSystem::EXCEL_1900,
            DateSystem::V1904 => formula_engine::date::ExcelDateSystem::Excel1904,
        }
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
    SharedString {
        index: u32,
    },
    InlineString,
    Bool,
    Error,
    Str,
    /// Cell value types that Formula does not interpret but should preserve on round-trip.
    ///
    /// SpreadsheetML `c` elements use a `t=` attribute to describe how to interpret the `<v>`
    /// payload. Excel emits additional values beyond the common `s/b/e/str/inlineStr` set (for
    /// example `t="d"` for ISO-8601 dates). When we don't understand the type, we keep the `t`
    /// string and the raw `<v>` text so we can rewrite `sheetData` without corrupting the file.
    Other {
        t: String,
    },
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
    calc_affecting_edits: bool,
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
            calc_affecting_edits: false,
        }
    }

    pub fn parts(&self) -> &BTreeMap<String, Vec<u8>> {
        &self.parts
    }

    pub fn save_to_vec(&self) -> Result<Vec<u8>, write::WriteError> {
        self.save_to_vec_with_recalc_policy(RecalcPolicy::default())
    }

    pub fn save_to_vec_with_recalc_policy(
        &self,
        recalc_policy: RecalcPolicy,
    ) -> Result<Vec<u8>, write::WriteError> {
        write::write_to_vec_with_recalc_policy(self, recalc_policy)
    }

    pub fn set_cell_value(
        &mut self,
        sheet_id: WorksheetId,
        cell: CellRef,
        value: CellValue,
    ) -> bool {
        let Some(sheet) = self.workbook.sheet_mut(sheet_id) else {
            return false;
        };
        sheet.set_value(cell, value.clone());

        let Some(cell_record) = sheet.cell(cell) else {
            self.meta.cell_meta.remove(&(sheet_id, cell));
            return true;
        };

        let meta = self.meta.cell_meta.entry((sheet_id, cell)).or_default();
        match (&meta.value_kind, &cell_record.value) {
            // Preserve less-common/unknown `t=` values by keeping the original type while the
            // model stores the cell value as a string (e.g. `t="d"` uses an ISO-8601 `<v>`).
            (Some(CellValueKind::Other { t }), CellValue::String(s)) => {
                meta.value_kind = Some(CellValueKind::Other { t: t.clone() });
                meta.raw_value = Some(s.clone());
            }
            _ => {
                let (value_kind, raw_value) = cell_meta_from_value(&cell_record.value);
                meta.value_kind = value_kind;
                meta.raw_value = raw_value;
            }
        }

        if meta.value_kind.is_none() && meta.raw_value.is_none() && meta.formula.is_none() {
            self.meta.cell_meta.remove(&(sheet_id, cell));
        }

        true
    }

    pub fn set_cell_formula(
        &mut self,
        sheet_id: WorksheetId,
        cell: CellRef,
        formula_display: Option<String>,
    ) -> bool {
        let Some(sheet) = self.workbook.sheet_mut(sheet_id) else {
            return false;
        };

        let had_formula_before = sheet.formula(cell).is_some()
            || self
                .meta
                .cell_meta
                .get(&(sheet_id, cell))
                .and_then(|m| m.formula.as_ref())
                .is_some_and(formula_meta_has_semantics);

        let Some(formula_display) = formula_display else {
            if had_formula_before {
                self.calc_affecting_edits = true;
            }
            sheet.set_formula(cell, None);

            // Preserve `FormulaMeta.file_text` for master formulas so the writer can detect
            // formula removals and apply recalculation safety as needed.
            let remove_meta = match self.meta.cell_meta.get_mut(&(sheet_id, cell)) {
                Some(meta) => {
                    if let Some(formula_meta) = meta.formula.as_mut() {
                        if formula_meta.file_text.is_empty() {
                            // Shared formula follower (no inline text) - clearing should remove the
                            // formula metadata entirely so the writer doesn't keep it.
                            meta.formula = None;
                        }
                    }

                    meta.formula.is_none() && meta.value_kind.is_none() && meta.raw_value.is_none()
                }
                None => false,
            };

            if remove_meta {
                self.meta.cell_meta.remove(&(sheet_id, cell));
            }

            // If the cell became truly empty, keep formula metadata (if any) so we can still
            // detect that a formula was removed later.
            if sheet.cell(cell).is_none() {
                let keep = self
                    .meta
                    .cell_meta
                    .get(&(sheet_id, cell))
                    .and_then(|m| m.formula.as_ref())
                    .is_some_and(|f| !f.file_text.is_empty());
                if !keep {
                    self.meta.cell_meta.remove(&(sheet_id, cell));
                }
            }

            return true;
        };

        let display = crate::formula_text::normalize_display_formula(&formula_display);
        if sheet.formula(cell).map(crate::formula_text::normalize_display_formula).as_deref()
            != Some(display.as_str())
        {
            self.calc_affecting_edits = true;
        }
        sheet.set_formula(cell, Some(display.clone()));

        let meta = self.meta.cell_meta.entry((sheet_id, cell)).or_default();
        if let Some(existing) = meta.formula.as_mut() {
            if existing.file_text.is_empty() {
                // Textless shared formulas become standalone formulas when edited.
                existing.t = None;
                existing.reference = None;
                existing.shared_index = None;
                existing.always_calc = None;
            }
            // Keep `file_text` unchanged so it can act as a baseline for detecting formula edits.
        }

        if let Some(cell_record) = sheet.cell(cell) {
            match (&meta.value_kind, &cell_record.value) {
                (Some(CellValueKind::Other { t }), CellValue::String(s)) => {
                    meta.value_kind = Some(CellValueKind::Other { t: t.clone() });
                    meta.raw_value = Some(s.clone());
                }
                _ => {
                    let (value_kind, raw_value) = cell_meta_from_value(&cell_record.value);
                    meta.value_kind = value_kind;
                    meta.raw_value = raw_value;
                }
            }
        }

        true
    }

    pub fn clear_cell(&mut self, sheet_id: WorksheetId, cell: CellRef) -> bool {
        let Some(sheet) = self.workbook.sheet_mut(sheet_id) else {
            return false;
        };
        let had_formula_before = sheet.formula(cell).is_some()
            || self
                .meta
                .cell_meta
                .get(&(sheet_id, cell))
                .and_then(|m| m.formula.as_ref())
                .is_some_and(formula_meta_has_semantics);
        sheet.clear_cell(cell);

        // Keep formula metadata for cleared master formulas so the writer can detect formula
        // removals and apply recalculation safety as needed.
        let keep_formula_meta = self
            .meta
            .cell_meta
            .get(&(sheet_id, cell))
            .and_then(|m| m.formula.as_ref())
            .is_some_and(|f| !f.file_text.is_empty());
        if keep_formula_meta {
            if let Some(meta) = self.meta.cell_meta.get_mut(&(sheet_id, cell)) {
                meta.value_kind = None;
                meta.raw_value = None;
            }
        } else {
            self.meta.cell_meta.remove(&(sheet_id, cell));
        }

        if had_formula_before {
            self.calc_affecting_edits = true;
        }
        true
    }
}

fn cell_meta_from_value(value: &CellValue) -> (Option<CellValueKind>, Option<String>) {
    match value {
        CellValue::Empty => (None, None),
        CellValue::Number(n) => (Some(CellValueKind::Number), Some(n.to_string())),
        CellValue::Boolean(b) => (
            Some(CellValueKind::Bool),
            Some(if *b { "1" } else { "0" }.to_string()),
        ),
        CellValue::Error(err) => (Some(CellValueKind::Error), Some(err.as_str().to_string())),
        CellValue::String(s) => (
            Some(CellValueKind::SharedString { index: 0 }),
            Some(s.clone()),
        ),
        CellValue::RichText(rich) => (
            Some(CellValueKind::SharedString { index: 0 }),
            Some(rich.text.clone()),
        ),
        _ => (Some(CellValueKind::Number), None),
    }
}

fn formula_meta_has_semantics(meta: &FormulaMeta) -> bool {
    !meta.file_text.is_empty()
        || meta.t.is_some()
        || meta.reference.is_some()
        || meta.shared_index.is_some()
        || meta.always_calc.is_some()
}
