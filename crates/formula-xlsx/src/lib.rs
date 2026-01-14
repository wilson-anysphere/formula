//! XLSX/XLSM compatibility layer.
//!
//! The long-term project goal is a full-fidelity Excel compatibility layer. The
//! crate currently exposes multiple APIs:
//!
//! - [`XlsxPackage`]: low-level Open Packaging Convention (OPC) ZIP handling
//!   that inflates the full ZIP into memory (part name -> bytes). This preserves
//!   part payloads like `xl/vbaProject.bin` byte-for-byte, but writing generally
//!   re-packs the ZIP container.
//! - [`XlsxLazyPackage`]: a lazy/streaming OPC package wrapper that avoids
//!   inflating every ZIP entry into memory and writes via the streaming rewrite
//!   pipeline (raw-copying untouched ZIP entries for performance and fidelity).
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
pub mod cell_images;
pub mod charts;
pub mod comments;
mod compare;
pub mod conditional_formatting;
mod content_types;
pub mod data_validations;
pub mod drawingml;
pub mod drawings;
mod encrypted;
pub mod embedded_cell_images;
pub mod embedded_images;
#[cfg(not(target_arch = "wasm32"))]
mod encrypted_ole;
mod formula_text;
pub mod hyperlinks;
mod lazy_package;
mod macro_repair;
mod macro_strip;
pub mod merge_cells;
pub mod metadata;
pub mod minimal;
pub mod offcrypto;
mod model_package;
pub mod openxml;
pub mod outline;
mod package;
pub mod patch;
mod path;
pub mod pivots;
mod preserve;
pub mod print;
#[cfg(not(target_arch = "wasm32"))]
mod office_crypto;
mod read;
#[cfg(not(target_arch = "wasm32"))]
mod reader;
mod recalc_policy;
mod relationships;
pub mod rich_data;
pub mod shared_strings;
mod sheet_metadata;
pub mod streaming;
pub mod styles;
pub mod tables;
pub mod theme;
mod package_stream;
#[cfg(feature = "vba")]
pub mod vba;
mod workbook;
pub mod write;
#[cfg(not(target_arch = "wasm32"))]
mod writer;
mod xml;
mod zip_util;

pub use crate::macro_strip::validate_opc_relationships;

use std::collections::{BTreeMap, HashMap};

pub use crate::minimal::write_minimal_xlsx;
pub use calc_settings::CalcSettingsError;
pub use compare::*;
pub use conditional_formatting::*;
pub use embedded_cell_images::EmbeddedCellImage;
pub use embedded_images::{extract_embedded_images, EmbeddedImageCell};
pub use hyperlinks::{
    parse_worksheet_hyperlinks, update_worksheet_relationships, update_worksheet_xml,
};
pub use offcrypto::{
    decrypt_agile_encrypted_package, decrypt_agile_encrypted_package_with_warnings,
    decrypt_agile_ooxml_from_cfb, decrypt_agile_ooxml_from_ole_bytes, decrypt_agile_ooxml_from_ole_reader,
    decrypt_ooxml_from_cfb, decrypt_ooxml_from_ole_bytes, decrypt_ooxml_from_ole_reader, OffCryptoError,
    OffCryptoWarning,
};
pub use lazy_package::XlsxLazyPackage;
pub use model_package::{WorkbookPackage, WorkbookPackageError};
pub use package::{
    read_part_from_reader, read_part_from_reader_limited, rewrite_content_types_workbook_content_type,
    rewrite_content_types_workbook_kind, theme_palette_from_reader, theme_palette_from_reader_limited,
    worksheet_parts_from_reader, worksheet_parts_from_reader_limited, CellPatch as PackageCellPatch,
    CellPatchSheet, MacroPresence, WorkbookKind, WorksheetPartInfo, XlsxError, XlsxPackage,
    XlsxPackageLimits, MAX_XLSX_PACKAGE_PART_BYTES, MAX_XLSX_PACKAGE_TOTAL_BYTES,
};
pub use package_stream::StreamingXlsxPackage;
pub use patch::{CellPatch, CellStyleRef, WorkbookCellPatches, WorksheetCellPatches};
pub use pivots::{
    cache_records::{pivot_cache_datetime_to_naive_date, PivotCacheRecordsReader, PivotCacheValue},
    graph::{PivotTableInstance, XlsxPivotGraph},
    pivot_charts::{PivotChartPart, PivotChartWithPlacement},
    slicers::{
        slicer_selection_to_engine_filter_field, slicer_selection_to_engine_filter_field_with_resolver,
        slicer_selection_to_row_filter, slicer_selection_to_row_filter_with_resolver,
        timeline_selection_to_engine_filter_field, timeline_selection_to_engine_filter_field_with_cache,
        timeline_selection_to_row_filter, PivotSlicerParts, SlicerDefinition, SlicerSelectionState,
        TimelineDefinition, TimelineSelectionState,
    },
    ux_graph::XlsxPivotUxGraph,
    PivotCacheDefinition, PivotCacheDefinitionPart, PivotCacheField, PivotCacheRecordsPart,
    PivotCacheSourceType, PivotTableDataField, PivotTableDefinition, PivotTableField,
    PivotTableFieldItem, PivotTablePageField, PivotTablePart, PivotTableStyleInfo,
    PreservedPivotParts, RelationshipStub, XlsxPivots,
};
#[cfg(not(target_arch = "wasm32"))]
pub use read::load_from_path;
pub use read::ReadError;
pub use read::{
    load_from_bytes, load_from_bytes_with_password, load_from_reader,
    read_workbook_model_from_bytes, read_workbook_model_from_bytes_with_password,
    read_workbook_model_from_reader,
};
#[cfg(not(target_arch = "wasm32"))]
pub use reader::{read_workbook, read_workbook_from_reader};
#[cfg(not(target_arch = "wasm32"))]
pub use encrypted_ole::{load_from_encrypted_ole_bytes, read_workbook_from_encrypted_reader};
pub use recalc_policy::RecalcPolicy;
pub use rich_data::metadata::parse_value_metadata_vm_to_rich_value_index_map;
pub use rich_data::resolve_rich_value_image_targets;
pub use rich_data::rich_value_structure::{
    parse_rich_value_structure_xml, RichValueStructure, RichValueStructureMember,
    RichValueStructures,
};
pub use rich_data::rich_value_types::{parse_rich_value_types_xml, RichValueType, RichValueTypes};
pub use rich_data::scan_cells_with_metadata_indices;
pub use rich_data::{
    discover_rich_data_part_names, discover_rich_data_part_names_from_metadata_rels,
    extract_rich_cell_images, RichDataError,
};
pub use rich_data::{ExtractedRichValueImages, RichValueEntry, RichValueIndex, RichValueWarning};
pub use sheet_metadata::{
    parse_sheet_tab_color, parse_workbook_sheets, write_sheet_tab_color, write_workbook_sheets,
    WorkbookSheetInfo,
};
pub use streaming::{
    patch_xlsx_streaming, patch_xlsx_streaming_with_recalc_policy,
    patch_xlsx_streaming_workbook_cell_patches,
    patch_xlsx_streaming_workbook_cell_patches_with_part_overrides,
    patch_xlsx_streaming_workbook_cell_patches_with_part_overrides_and_recalc_policy,
    patch_xlsx_streaming_workbook_cell_patches_with_recalc_policy,
    patch_xlsx_streaming_workbook_cell_patches_with_styles,
    patch_xlsx_streaming_workbook_cell_patches_with_styles_and_part_overrides,
    patch_xlsx_streaming_workbook_cell_patches_with_styles_and_part_overrides_and_recalc_policy,
    patch_xlsx_streaming_workbook_cell_patches_with_styles_and_recalc_policy,
    strip_vba_project_streaming, strip_vba_project_streaming_with_kind, PartOverride,
    StreamingPatchError, WorksheetCellPatch,
};
pub use styles::*;
pub use workbook::ChartExtractionError;
#[cfg(not(target_arch = "wasm32"))]
pub use writer::{
    write_workbook, write_workbook_to_writer, write_workbook_to_writer_encrypted,
    write_workbook_to_writer_with_kind, XlsxWriteError,
};
pub use xml::XmlDomError;

use formula_model::drawings::DrawingObject;
use formula_model::rich_text::RichText;
use formula_model::{CellRef, CellValue, Comment, ErrorValue, Workbook, WorksheetId};

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
    /// SpreadsheetML cell metadata indices (`c/@vm` and `c/@cm`).
    ///
    /// Excel emits `vm`/`cm` attributes on `<c>` elements to reference value metadata and cell
    /// metadata records (used for modern features like linked data types / rich values).
    ///
    /// These are typically integer indices into `xl/metadata.xml`, but we keep the raw attribute
    /// text so we can round-trip the file without normalizing the formatting (e.g. leading zeros).
    pub vm: Option<String>,
    pub cm: Option<String>,
    pub formula: Option<FormulaMeta>,
}

#[derive(Debug, Clone, Default)]
pub struct XlsxMeta {
    pub date_system: DateSystem,
    pub calc_pr: CalcPr,
    pub sheets: Vec<SheetMeta>,
    pub cell_meta: HashMap<(WorksheetId, CellRef), CellMeta>,
    /// Baseline conditional formatting blocks extracted from each worksheet XML payload.
    ///
    /// This stores the original SpreadsheetML/x14 `<conditionalFormatting>` XML fragments so
    /// fidelity diagnostics can detect when conditional formatting has been rewritten during
    /// round-trip operations.
    ///
    /// Notes:
    /// - Only the extracted conditional formatting blocks are stored (not the full worksheet XML)
    ///   to keep memory usage reasonable.
    /// - A worksheet key is present only when conditional formatting blocks were detected.
    pub conditional_formatting: HashMap<WorksheetId, Vec<RawConditionalFormattingBlock>>,
    /// Mapping from worksheet cells to rich value record indices (e.g. images-in-cell backed by
    /// `xl/richData/richValue.xml`).
    pub rich_value_cells: HashMap<(WorksheetId, CellRef), u32>,
    /// Per-worksheet mapping of existing comment-related XML part names discovered on load.
    ///
    /// This is populated only for workbooks loaded via [`load_from_bytes`]. It is used by the
    /// `XlsxDocument` writer to update existing comment parts in-place while preserving unknown
    /// comment-related parts and worksheet relationship IDs.
    pub comment_part_names: HashMap<WorksheetId, WorksheetCommentPartNames>,
    /// Snapshot of worksheet comments as loaded from the original workbook, normalized to a stable
    /// ordering.
    ///
    /// This is used to detect when comments have been edited in-memory so we can rewrite only the
    /// affected comment XML parts on save.
    pub comment_snapshot: HashMap<WorksheetId, Vec<Comment>>,
    /// Snapshot of worksheet drawings as loaded from the original workbook.
    ///
    /// This is used by the `XlsxDocument` writer to detect edits to `Worksheet.drawings` without
    /// having to reparse `xl/drawings/*.xml` parts at save time.
    pub drawings_snapshot: HashMap<WorksheetId, Vec<DrawingObject>>,
    /// Snapshot of the workbook print settings as they were originally loaded into the
    /// in-memory model.
    ///
    /// The writer uses this to detect no-op saves: when the model's
    /// [`formula_model::Workbook::print_settings`] matches this snapshot, we avoid rewriting
    /// `xl/workbook.xml` (print-related defined names) and `xl/worksheets/sheetN.xml`
    /// (page setup/margins/breaks).
    pub print_settings_snapshot: formula_model::WorkbookPrintSettings,
}

/// OPC part names for comment XML parts referenced from a worksheet's `.rels`.
#[derive(Debug, Clone, Default)]
pub struct WorksheetCommentPartNames {
    /// Legacy note comments (e.g. `xl/comments1.xml`).
    pub legacy_comments: Option<String>,
    /// Modern threaded comments (e.g. `xl/threadedComments/threadedComments1.xml`).
    pub threaded_comments: Option<String>,
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
    workbook_kind: WorkbookKind,
}

impl XlsxDocument {
    pub fn new(workbook: Workbook) -> Self {
        Self::new_with_kind(workbook, WorkbookKind::Workbook)
    }

    pub fn new_with_kind(workbook: Workbook, workbook_kind: WorkbookKind) -> Self {
        let date_system = match workbook.date_system {
            formula_model::DateSystem::Excel1900 => DateSystem::V1900,
            formula_model::DateSystem::Excel1904 => DateSystem::V1904,
        };

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
                date_system,
                sheets,
                ..XlsxMeta::default()
            },
            calc_affecting_edits: false,
            workbook_kind,
        }
    }

    pub fn parts(&self) -> &BTreeMap<String, Vec<u8>> {
        &self.parts
    }

    /// Returns the parsed XLSX round-trip metadata captured while loading the workbook.
    ///
    /// This is an advanced API intended for round-trip diagnostics and integration
    /// tests that need to assert on preserved SpreadsheetML details (for example
    /// unknown cell value types, `vm/cm`-derived cell metadata, or future
    /// `xl/metadata.xml` parsing).
    ///
    /// Most callers should treat [`XlsxDocument::workbook`] as the primary source
    /// of truth and only consult this metadata when working on fidelity / OPC
    /// preservation issues.
    pub fn xlsx_meta(&self) -> &XlsxMeta {
        &self.meta
    }

    /// Returns a mutable view of the parsed XLSX round-trip metadata captured while loading the workbook.
    ///
    /// This is an advanced API intended for fidelity-focused tooling and tests.
    ///
    /// Mutating the metadata can easily produce invalid XLSX output (for example by
    /// inserting inconsistent indices or referencing non-existent cells/sheets), and
    /// it does not automatically keep the associated [`Workbook`] model in sync.
    ///
    /// Most callers should treat [`XlsxDocument::workbook`] as the primary source
    /// of truth and only mutate this metadata when they explicitly need to control
    /// low-level SpreadsheetML round-trip behavior.
    pub fn xlsx_meta_mut(&mut self) -> &mut XlsxMeta {
        &mut self.meta
    }

    /// Returns the baseline conditional formatting XML blocks extracted from the worksheet when
    /// this document was loaded (if any).
    ///
    /// This is a convenience wrapper over [`XlsxMeta::conditional_formatting`]. The returned
    /// blocks preserve the original SpreadsheetML schema (including x14 extensions) so
    /// round-trip tooling can detect when conditional formatting has been rewritten.
    pub fn conditional_formatting_blocks(
        &self,
        sheet_id: WorksheetId,
    ) -> Option<&[RawConditionalFormattingBlock]> {
        self.meta
            .conditional_formatting
            .get(&sheet_id)
            .map(|v| v.as_slice())
    }

    /// Returns metadata captured for a specific cell (if any).
    ///
    /// This is a convenience wrapper over [`XlsxMeta::cell_meta`] and exists
    /// primarily for round-trip oriented tooling/tests. Many cells have no
    /// associated metadata entry.
    ///
    /// Note: Excel treats merged regions as a single cell anchored at the
    /// region's top-left. To match workbook semantics, this helper resolves any
    /// cell inside a merged region to that anchor cell before looking up
    /// metadata.
    pub fn cell_meta(&self, sheet_id: WorksheetId, cell: CellRef) -> Option<&CellMeta> {
        let cell = self
            .workbook
            .sheet(sheet_id)
            .map(|sheet| sheet.merged_regions.resolve_cell(cell))
            .unwrap_or(cell);
        self.meta.cell_meta.get(&(sheet_id, cell))
    }

    pub fn rich_value_index(&self, sheet: WorksheetId, cell: CellRef) -> Option<u32> {
        self.rich_value_index_for_cell(sheet, cell).ok().flatten()
    }

    pub fn workbook_kind(&self) -> WorkbookKind {
        self.workbook_kind
    }

    pub fn set_workbook_kind(&mut self, workbook_kind: WorkbookKind) {
        self.workbook_kind = workbook_kind;
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

        // Match workbook semantics for merged regions: any cell inside a merge resolves to the
        // anchor (top-left) cell.
        let cell = sheet.merged_regions.resolve_cell(cell);

        // Treat absent cells as equivalent to `CellValue::Empty` when detecting value changes. This
        // lets callers round-trip metadata-only cells (e.g. `c/@vm`) without losing those
        // attributes when they set the value to empty.
        let old_value = sheet
            .cell(cell)
            .map(|record| record.value.clone())
            .unwrap_or(CellValue::Empty);
        let value_changed = old_value != value;

        // `vm="..."` is a SpreadsheetML value-metadata pointer (typically into `xl/metadata*.xml`).
        // If the cached value changes we can't update the corresponding metadata records, so drop
        // `vm` to avoid leaving dangling rich-data references.
        //
        // Note: this must happen before `sheet.set_value` so clearing a cell can also remove its
        // associated `CellMeta` entry (since `vm` participates in `keep_due_to_metadata` below).
        if value_changed {
            if let Some(meta) = self.meta.cell_meta.get_mut(&(sheet_id, cell)) {
                meta.vm = None;
            }
            self.meta.rich_value_cells.remove(&(sheet_id, cell));
        }
        sheet.set_value(cell, value.clone());

        let Some(cell_record) = sheet.cell(cell) else {
            let keep_due_to_metadata = self
                .meta
                .cell_meta
                .get(&(sheet_id, cell))
                .is_some_and(|meta| meta.cm.is_some() || meta.vm.is_some());
            if !keep_due_to_metadata {
                self.meta.cell_meta.remove(&(sheet_id, cell));
            }
            return true;
        };

        let meta = self.meta.cell_meta.entry((sheet_id, cell)).or_default();
        if value_changed {
            meta.vm = None;
            self.meta.rich_value_cells.remove(&(sheet_id, cell));
        }
        match (&meta.value_kind, &cell_record.value) {
            // Preserve less-common/unknown `t=` values by keeping the original type while the
            // model stores the cell value as a string (e.g. `t="d"` uses an ISO-8601 `<v>`).
            (Some(CellValueKind::Other { t }), CellValue::String(s)) => {
                meta.value_kind = Some(CellValueKind::Other { t: t.clone() });
                meta.raw_value = Some(s.clone());
            }
            // If the cell previously referenced a specific shared string index, preserve it
            // when the visible text is unchanged. This avoids flipping between duplicate shared
            // string entries that may differ only in unsupported substructures (phonetic/ruby,
            // extLst, etc.).
            (Some(CellValueKind::SharedString { index }), CellValue::String(s))
                if self
                    .shared_strings
                    .get(*index as usize)
                    .is_some_and(|rt| rt.text.as_str() == s.as_str()) =>
            {
                let idx = *index;
                let raw_matches = meta
                    .raw_value
                    .as_deref()
                    .and_then(|raw| raw.trim().parse::<u32>().ok())
                    .is_some_and(|raw| raw == idx);
                meta.value_kind = Some(CellValueKind::SharedString { index: idx });
                if !raw_matches {
                    meta.raw_value = Some(idx.to_string());
                }
            }
            (Some(CellValueKind::SharedString { index }), CellValue::Entity(entity))
                if self
                    .shared_strings
                    .get(*index as usize)
                    .is_some_and(|rt| rt.text.as_str() == entity.display_value.as_str()) =>
            {
                let idx = *index;
                let raw_matches = meta
                    .raw_value
                    .as_deref()
                    .and_then(|raw| raw.trim().parse::<u32>().ok())
                    .is_some_and(|raw| raw == idx);
                meta.value_kind = Some(CellValueKind::SharedString { index: idx });
                if !raw_matches {
                    meta.raw_value = Some(idx.to_string());
                }
            }
            (Some(CellValueKind::SharedString { index }), CellValue::Record(record))
                if {
                    let display = record.to_string();
                    self.shared_strings
                        .get(*index as usize)
                        .is_some_and(|rt| rt.text.as_str() == display.as_str())
                } =>
            {
                let idx = *index;
                let raw_matches = meta
                    .raw_value
                    .as_deref()
                    .and_then(|raw| raw.trim().parse::<u32>().ok())
                    .is_some_and(|raw| raw == idx);
                meta.value_kind = Some(CellValueKind::SharedString { index: idx });
                if !raw_matches {
                    meta.raw_value = Some(idx.to_string());
                }
            }
            (Some(CellValueKind::SharedString { index }), CellValue::Image(image))
                if image
                    .alt_text
                    .as_deref()
                    .filter(|s| !s.is_empty())
                    .is_some_and(|alt| {
                        self.shared_strings
                            .get(*index as usize)
                            .is_some_and(|rt| rt.text.as_str() == alt)
                    }) =>
            {
                // Preserve the original shared string index when the degraded display text is
                // unchanged. This avoids switching between duplicate shared string entries that
                // may differ only in unsupported substructures (phonetic runs, extLst, etc.).
                let idx = *index;
                let raw_matches = meta
                    .raw_value
                    .as_deref()
                    .and_then(|raw| raw.trim().parse::<u32>().ok())
                    .is_some_and(|raw| raw == idx);
                meta.value_kind = Some(CellValueKind::SharedString { index: idx });
                if !raw_matches {
                    meta.raw_value = Some(idx.to_string());
                }
            }
            (Some(CellValueKind::SharedString { index }), CellValue::RichText(rich))
                if self
                    .shared_strings
                    .get(*index as usize)
                    .is_some_and(|rt| rt == rich) =>
            {
                let idx = *index;
                let raw_matches = meta
                    .raw_value
                    .as_deref()
                    .and_then(|raw| raw.trim().parse::<u32>().ok())
                    .is_some_and(|raw| raw == idx);
                meta.value_kind = Some(CellValueKind::SharedString { index: idx });
                if !raw_matches {
                    meta.raw_value = Some(idx.to_string());
                }
            }
            _ => {
                let (value_kind, raw_value) = cell_meta_from_value(&cell_record.value);
                meta.value_kind = value_kind;
                meta.raw_value = raw_value;
            }
        }

        // `vm` (value metadata) indices point into `xl/metadata.xml` and are tied to the stored
        // cell value. When the caller edits the cell value we do not currently update
        // `xl/metadata.xml`, so keep `vm` only when the cell remains a rich-value placeholder
        // (`#VALUE!`).
        if !matches!(
            cell_record.value,
            CellValue::Error(ErrorValue::Value)
        ) {
            meta.vm = None;
        }

        if meta.value_kind.is_none()
            && meta.raw_value.is_none()
            && meta.formula.is_none()
            && meta.vm.is_none()
            && meta.cm.is_none()
        {
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

        // Match workbook semantics for merged regions: any cell inside a merge
        // resolves to the anchor (top-left) cell.
        let cell = sheet.merged_regions.resolve_cell(cell);

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

                    meta.formula.is_none()
                        && meta.value_kind.is_none()
                        && meta.raw_value.is_none()
                        && meta.vm.is_none()
                        && meta.cm.is_none()
                }
                None => false,
            };

            if remove_meta {
                self.meta.cell_meta.remove(&(sheet_id, cell));
            }

            // If the cell became truly empty, keep formula metadata (if any) so we can still
            // detect that a formula was removed later.
            if sheet.cell(cell).is_none() {
                let keep = self.meta.cell_meta.get(&(sheet_id, cell)).is_some_and(|m| {
                    m.cm.is_some()
                        || m.vm.is_some()
                        || m.formula.as_ref().is_some_and(|f| !f.file_text.is_empty())
                });
                if !keep {
                    self.meta.cell_meta.remove(&(sheet_id, cell));
                }
            }

            return true;
        };

        let display = crate::formula_text::normalize_display_formula(&formula_display);
        if sheet
            .formula(cell)
            .map(crate::formula_text::normalize_display_formula)
            .as_deref()
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
                (Some(CellValueKind::SharedString { index }), CellValue::String(s))
                    if self
                        .shared_strings
                        .get(*index as usize)
                        .is_some_and(|rt| rt.text.as_str() == s.as_str()) =>
                {
                    let idx = *index;
                    let raw_matches = meta
                        .raw_value
                        .as_deref()
                        .and_then(|raw| raw.trim().parse::<u32>().ok())
                        .is_some_and(|raw| raw == idx);
                    meta.value_kind = Some(CellValueKind::SharedString { index: idx });
                    if !raw_matches {
                        meta.raw_value = Some(idx.to_string());
                    }
                }
                (Some(CellValueKind::SharedString { index }), CellValue::Entity(entity))
                    if self
                        .shared_strings
                        .get(*index as usize)
                        .is_some_and(|rt| rt.text.as_str() == entity.display_value.as_str()) =>
                {
                    let idx = *index;
                    let raw_matches = meta
                        .raw_value
                        .as_deref()
                        .and_then(|raw| raw.trim().parse::<u32>().ok())
                        .is_some_and(|raw| raw == idx);
                    meta.value_kind = Some(CellValueKind::SharedString { index: idx });
                    if !raw_matches {
                        meta.raw_value = Some(idx.to_string());
                    }
                }
                (Some(CellValueKind::SharedString { index }), CellValue::Record(record))
                    if {
                        let display = record.to_string();
                        self.shared_strings
                            .get(*index as usize)
                            .is_some_and(|rt| rt.text.as_str() == display.as_str())
                    } =>
                {
                    let idx = *index;
                    let raw_matches = meta
                        .raw_value
                        .as_deref()
                        .and_then(|raw| raw.trim().parse::<u32>().ok())
                        .is_some_and(|raw| raw == idx);
                    meta.value_kind = Some(CellValueKind::SharedString { index: idx });
                    if !raw_matches {
                        meta.raw_value = Some(idx.to_string());
                    }
                }
                (Some(CellValueKind::SharedString { index }), CellValue::Image(image))
                    if image
                        .alt_text
                        .as_deref()
                        .filter(|s| !s.is_empty())
                        .is_some_and(|alt| {
                            self.shared_strings
                                .get(*index as usize)
                                .is_some_and(|rt| rt.text.as_str() == alt)
                        }) =>
                {
                    let idx = *index;
                    let raw_matches = meta
                        .raw_value
                        .as_deref()
                        .and_then(|raw| raw.trim().parse::<u32>().ok())
                        .is_some_and(|raw| raw == idx);
                    meta.value_kind = Some(CellValueKind::SharedString { index: idx });
                    if !raw_matches {
                        meta.raw_value = Some(idx.to_string());
                    }
                }
                (Some(CellValueKind::SharedString { index }), CellValue::RichText(rich))
                    if self
                        .shared_strings
                        .get(*index as usize)
                        .is_some_and(|rt| rt == rich) =>
                {
                    let idx = *index;
                    let raw_matches = meta
                        .raw_value
                        .as_deref()
                        .and_then(|raw| raw.trim().parse::<u32>().ok())
                        .is_some_and(|raw| raw == idx);
                    meta.value_kind = Some(CellValueKind::SharedString { index: idx });
                    if !raw_matches {
                        meta.raw_value = Some(idx.to_string());
                    }
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

        // Match workbook semantics for merged regions: any cell inside a merge
        // resolves to the anchor (top-left) cell.
        let cell = sheet.merged_regions.resolve_cell(cell);

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
        let keep_vm_cm = self
            .meta
            .cell_meta
            .get(&(sheet_id, cell))
            .is_some_and(|m| m.vm.is_some() || m.cm.is_some());

        if keep_formula_meta || keep_vm_cm {
            if let Some(meta) = self.meta.cell_meta.get_mut(&(sheet_id, cell)) {
                meta.value_kind = None;
                meta.raw_value = None;
                if !keep_formula_meta {
                    meta.formula = None;
                }
            }
        } else {
            self.meta.cell_meta.remove(&(sheet_id, cell));
        }

        if had_formula_before {
            self.calc_affecting_edits = true;
        }
        true
    }

    /// Resolve a worksheet cell's `c/@vm` (value metadata index) to a rich value index.
    ///
    /// This is a best-effort helper for Excel rich values (including images-in-cell).
    ///
    /// Returns:
    /// - `Ok(None)` if the cell has no stored `CellMeta.vm` (no `c/@vm`), or if the workbook has no
    ///   `xl/metadata.xml` part.
    /// - `Ok(Some(index))` if the cell's `vm` can be resolved to an `xl/richData/richValue.xml`
    ///   record index.
    ///
    /// Note: this is currently best-effort and will not return an error (the `Result` is for API
    /// consistency / future extensibility).
    pub fn rich_value_index_for_cell(
        &self,
        sheet_id: WorksheetId,
        cell: CellRef,
    ) -> Result<Option<u32>, XlsxError> {
        // Match workbook semantics for merged regions: any cell inside a merge resolves to the
        // anchor (top-left) cell.
        let cell = self
            .workbook
            .sheet(sheet_id)
            .map(|sheet| sheet.merged_regions.resolve_cell(cell))
            .unwrap_or(cell);

        // Only report rich values for cells that still have stored `vm` metadata.
        // (e.g. a cell cleared after load may retain a stale `rich_value_cells` entry.)
        let has_vm = self
            .meta
            .cell_meta
            .get(&(sheet_id, cell))
            .is_some_and(|m| m.vm.is_some());
        if !has_vm {
            return Ok(None);
        };

        Ok(self.meta.rich_value_cells.get(&(sheet_id, cell)).copied())
    }

    /// Resolve all cells with a stored `c/@vm` (value metadata index) to rich value indices.
    ///
    /// This is a best-effort helper that returns an empty map if `xl/metadata.xml` is missing.
    pub fn rich_value_indices(&self) -> Result<HashMap<(WorksheetId, CellRef), u32>, XlsxError> {
        let mut out = HashMap::new();
        for (&(worksheet_id, cell_ref), &idx) in &self.meta.rich_value_cells {
            // Only include cells that still have stored `vm` metadata.
            let has_vm = self
                .meta
                .cell_meta
                .get(&(worksheet_id, cell_ref))
                .is_some_and(|m| m.vm.is_some());
            if !has_vm {
                continue;
            };
            out.insert((worksheet_id, cell_ref), idx);
        }

        Ok(out)
    }

    /// Resolve an "image in cell" rich value to the underlying `xl/media/*` target part.
    ///
    /// This is a best-effort helper:
    /// - Returns `Ok(None)` when the cell has no rich value, or when the rich value is not an image
    ///   (or cannot be resolved).
    /// - Returns `Err(_)` only for invalid XML/UTF-8 in the relevant rich-data parts.
    pub fn image_target_for_cell(
        &self,
        sheet_id: WorksheetId,
        cell: CellRef,
    ) -> Result<Option<String>, XlsxError> {
        let Some(rich_value_index) = self.rich_value_index_for_cell(sheet_id, cell)? else {
            return Ok(None);
        };

        // Resolve rich value index -> relationship index via `xl/richData/richValue*.xml`.
        let Some(rel_index) = self.rich_value_rel_index(rich_value_index)? else {
            return Ok(None);
        };

        // Resolve relationship index -> rId via `xl/richData/richValueRel*.xml`.
        let Some((rich_value_rel_part, rel_id)) = self.rich_value_rel_id(rel_index)? else {
            return Ok(None);
        };

        // Resolve rId -> target part via the richValueRel part's `.rels`.
        let rels_part = crate::openxml::rels_part_name(&rich_value_rel_part);
        let Some(rels_bytes) = self.parts.get(&rels_part) else {
            return Ok(None);
        };

        for rel in crate::openxml::parse_relationships(rels_bytes)? {
            if rel.id != rel_id {
                continue;
            }
            if rel
                .target_mode
                .as_deref()
                .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
            {
                return Ok(None);
            }
            if rel.type_uri != crate::drawings::REL_TYPE_IMAGE {
                return Ok(None);
            }

            // Relationship targets may include URI fragments (`../media/image.png#foo`); those are
            // not part names.
            let target = rel
                .target
                .split_once('#')
                .map(|(t, _)| t)
                .unwrap_or(&rel.target);
            if target.is_empty() {
                return Ok(None);
            }
            // Some producers emit rich-data relationship targets that are relative to `xl/`
            // (e.g. `Target="media/image1.png"`) or omit the leading `/` for package-root targets
            // (e.g. `Target="xl/media/image1.png"`). Use the same best-effort target normalization
            // as the other rich-data extractors.
            let target_part =
                crate::rich_data::resolve_rich_value_rel_target_part(&rich_value_rel_part, target);
            if !self.parts.contains_key(&target_part) {
                return Ok(None);
            }
            return Ok(Some(target_part));
        }

        Ok(None)
    }

    fn rich_value_rel_index(&self, rich_value_index: u32) -> Result<Option<u32>, XlsxError> {
        #[derive(Debug, Clone)]
        struct ParsedRv {
            explicit_index: Option<u32>,
            rel_index: Option<u32>,
        }

        fn rich_value_part_suffix(part_name: &str) -> Option<u32> {
            if !part_name.starts_with("xl/richData/") {
                return None;
            }
            let file_name = part_name.rsplit('/').next()?;
            let file_name_lower = file_name.to_ascii_lowercase();
            if !file_name_lower.ends_with(".xml") {
                return None;
            }

            let stem_lower = &file_name_lower[..file_name_lower.len() - ".xml".len()];
            // Check the plural prefix first: `richvalues` starts with `richvalue`.
            let suffix = if let Some(rest) = stem_lower.strip_prefix("richvalues") {
                rest
            } else if let Some(rest) = stem_lower.strip_prefix("richvalue") {
                rest
            } else {
                return None;
            };

            if suffix.is_empty() {
                return Some(0);
            }
            if !suffix.chars().all(|c| c.is_ascii_digit()) {
                return None;
            }
            suffix.parse::<u32>().ok()
        }

        fn parse_rv_explicit_index(rv: roxmltree::Node<'_, '_>) -> Option<u32> {
            rv.attribute("i")
                .or_else(|| rv.attribute("id"))
                .or_else(|| rv.attribute("idx"))
                .and_then(|v| v.trim().parse::<u32>().ok())
        }

        fn parse_rv_rel_index(rv: roxmltree::Node<'_, '_>) -> Option<u32> {
            let v_elems: Vec<_> = rv
                .descendants()
                .filter(|n| n.is_element() && n.tag_name().name() == "v")
                .collect();

            // Prefer `<v t="rel">` / `<v t="r">`.
            for v in &v_elems {
                let Some(t) = v.attribute("t") else {
                    continue;
                };
                if t == "rel" || t == "r" {
                    if let Some(text) = v.text() {
                        if let Ok(idx) = text.trim().parse::<u32>() {
                            return Some(idx);
                        }
                    }
                }
            }

            // Fall back to a numeric `<v>` without a type marker.
            for v in &v_elems {
                if v.attribute("t").is_some() {
                    continue;
                }
                if let Some(text) = v.text() {
                    if let Ok(idx) = text.trim().parse::<u32>() {
                        return Some(idx);
                    }
                }
            }

            // Last-ditch: any numeric `<v>`.
            for v in &v_elems {
                if let Some(text) = v.text() {
                    if let Ok(idx) = text.trim().parse::<u32>() {
                        return Some(idx);
                    }
                }
            }

            None
        }

        let mut parsed: Vec<ParsedRv> = Vec::new();

        // Deterministic part ordering (numeric suffix; not lexicographic).
        //
        // Excel can split rich value stores across many parts (e.g. richValue.xml, richValue1.xml, ...,
        // richValue10.xml). A lexicographic sort puts richValue10 before richValue2, corrupting the
        // implicit index assignment.
        let mut part_names: Vec<(u32, &str)> = self
            .parts
            .keys()
            .filter_map(|name| rich_value_part_suffix(name).map(|idx| (idx, name.as_str())))
            .collect();
        part_names.sort_by(|(a_idx, a_name), (b_idx, b_name)| {
            a_idx.cmp(b_idx).then_with(|| a_name.cmp(b_name))
        });

        for (_idx, part_name) in part_names {
            let Some(bytes) = self.parts.get(part_name) else {
                continue;
            };
            let xml = std::str::from_utf8(bytes)
                .map_err(|e| XlsxError::Invalid(format!("{part_name} is not valid UTF-8: {e}")))?;
            let doc = roxmltree::Document::parse(xml)?;

            for rv in doc
                .root_element()
                .descendants()
                .filter(|n| n.is_element() && n.tag_name().name() == "rv")
            {
                parsed.push(ParsedRv {
                    explicit_index: parse_rv_explicit_index(rv),
                    rel_index: parse_rv_rel_index(rv),
                });
            }
        }

        if parsed.is_empty() {
            return Ok(None);
        }

        // Build a global index -> relationship index map that honors explicit `<rv id="...">` indices
        // when present.
        let mut out: HashMap<u32, Option<u32>> = HashMap::new();
        let mut max_explicit: Option<u32> = None;
        for rv in &parsed {
            let Some(idx) = rv.explicit_index else {
                continue;
            };
            if out.contains_key(&idx) {
                // Deterministic: first wins.
                continue;
            }
            max_explicit = Some(max_explicit.map(|m| m.max(idx)).unwrap_or(idx));
            out.insert(idx, rv.rel_index);
        }

        let mut next = match max_explicit {
            Some(max) => max.saturating_add(1),
            None => 0,
        };
        for rv in &parsed {
            if rv.explicit_index.is_some() {
                continue;
            }
            while out.contains_key(&next) {
                next = next.saturating_add(1);
            }
            out.insert(next, rv.rel_index);
            next = next.saturating_add(1);
        }

        Ok(out.get(&rich_value_index).copied().flatten())
    }

    fn rich_value_rel_id(&self, rel_index: u32) -> Result<Option<(String, String)>, XlsxError> {
        fn rich_value_rel_part_suffix(part_name: &str) -> Option<u32> {
            const PREFIX: &str = "xl/richData/richValueRel";
            const SUFFIX: &str = ".xml";
            if !part_name.starts_with(PREFIX) || !part_name.ends_with(SUFFIX) {
                return None;
            }
            if part_name.contains("/_rels/") {
                return None;
            }
            let mid = &part_name[PREFIX.len()..part_name.len() - SUFFIX.len()];
            if mid.is_empty() {
                return Some(0);
            }
            if !mid.chars().all(|c| c.is_ascii_digit()) {
                return None;
            }
            mid.parse::<u32>().ok()
        }

        fn parse_rel_ids(xml: &str) -> Result<Vec<String>, XlsxError> {
            let doc = roxmltree::Document::parse(xml)?;
            let mut out = Vec::new();
            for rel in doc
                .descendants()
                .filter(|n| n.is_element() && n.tag_name().name() == "rel")
            {
                let rid = rel
                    .attribute((
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                        "id",
                    ))
                    .or_else(|| rel.attribute("r:id"))
                    .or_else(|| rel.attribute("id"));
                let Some(rid) = rid else {
                    continue;
                };
                out.push(rid.to_string());
            }
            Ok(out)
        }

        // Prefer the canonical name, but fall back to any `richValueRel*.xml`.
        let rich_value_rel_part = if self.parts.contains_key("xl/richData/richValueRel.xml") {
            "xl/richData/richValueRel.xml".to_string()
        } else {
            let mut candidates: Vec<(u32, &str)> = self
                .parts
                .keys()
                .filter_map(|name| rich_value_rel_part_suffix(name).map(|idx| (idx, name.as_str())))
                .collect();
            candidates.sort_by(|(a_idx, a_name), (b_idx, b_name)| {
                a_idx.cmp(b_idx).then_with(|| a_name.cmp(b_name))
            });
            candidates
                .first()
                .map(|(_, name)| (*name).to_string())
                .unwrap_or_default()
        };

        if rich_value_rel_part.is_empty() {
            return Ok(None);
        }

        let Some(bytes) = self.parts.get(&rich_value_rel_part) else {
            return Ok(None);
        };
        let xml = std::str::from_utf8(bytes).map_err(|e| {
            XlsxError::Invalid(format!("{rich_value_rel_part} is not valid UTF-8: {e}"))
        })?;
        let rel_ids = parse_rel_ids(xml)?;
        let Some(rel_id) = crate::rich_data::rel_slot_get(&rel_ids, rel_index as usize).cloned()
        else {
            return Ok(None);
        };

        Ok(Some((rich_value_rel_part, rel_id)))
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
        CellValue::Entity(entity) => (
            Some(CellValueKind::SharedString { index: 0 }),
            Some(entity.display_value.clone()),
        ),
        CellValue::Record(record) => (
            Some(CellValueKind::SharedString { index: 0 }),
            Some(record.to_string()),
        ),
        CellValue::Image(image) => match image.alt_text.as_deref().filter(|s| !s.is_empty()) {
            Some(alt) => (
                Some(CellValueKind::SharedString { index: 0 }),
                Some(alt.to_string()),
            ),
            None => (None, None),
        },
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
