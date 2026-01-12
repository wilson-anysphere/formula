//! Legacy Excel 97-2003 `.xls` (BIFF) import support.
//!
//! This importer is intentionally best-effort: BIFF contains many features that
//! aren't representable in [`formula_model`]. We load sheets, cell values,
//! formulas (as text), merged-cell regions, basic row/column size/visibility
//! metadata, hyperlinks, basic outline (grouping) metadata, and legacy cell
//! comments/notes ([`formula_model::CommentKind::Note`]) on worksheets where
//! available. We also attempt to preserve workbook/worksheet view state where
//! possible (active tab, frozen panes, zoom, selection, and grid/headings/zero
//! visibility flags).
//!
//! When present, workbook- and sheet-scoped defined names (named ranges) are also
//! imported. Defined-name formula (`rgce`) decoding is best-effort and may emit
//! warnings for unsupported tokens.
//!
//! Note import is best-effort and intentionally lossy:
//! - Comment box geometry/visibility/formatting is not preserved
//! - Only plain text + author (when available) are imported
//! - Threaded comments are not supported in `.xls`
//! - Notes inside merged regions are anchored to the merged region's top-left cell
//!
//! We also extract workbook-global styles (including number format codes) and the
//! workbook date system (1900 vs 1904) when possible. Anything else is preserved
//! as metadata/warnings. In particular, comment parsing may emit warnings when
//! BIFF NOTE/OBJ/TXO records are malformed or incomplete (and such comments may
//! be skipped).

use std::collections::{HashMap, HashSet};
use std::io::{Cursor, Write};
use std::path::{Path, PathBuf};

use calamine::{Data, Reader, Sheet, SheetType, SheetVisible, Xls};
use formula_model::{
    normalize_formula_text, CellRef, CellValue, ColRange, Comment, CommentAuthor, CommentKind,
    DefinedNameScope, ErrorValue, HyperlinkTarget, PrintTitles, Range, RowRange, SheetAutoFilter,
    SheetVisibility, Style, TabColor, Workbook, XLNM_FILTER_DATABASE, EXCEL_MAX_COLS, EXCEL_MAX_ROWS,
    EXCEL_MAX_SHEET_NAME_LEN,
};
use thiserror::Error;

mod biff;
mod formula_rewrite;

#[derive(Clone, Copy, Debug)]
struct DateTimeStyleIds {
    date: u32,
    datetime: u32,
    time: u32,
    duration: u32,
}

impl DateTimeStyleIds {
    fn new(workbook: &mut Workbook) -> Self {
        let date = workbook.intern_style(Style {
            number_format: Some("m/d/yy".to_string()),
            ..Default::default()
        });
        let datetime = workbook.intern_style(Style {
            number_format: Some("m/d/yy h:mm:ss".to_string()),
            ..Default::default()
        });
        let time = workbook.intern_style(Style {
            number_format: Some("h:mm:ss".to_string()),
            ..Default::default()
        });
        let duration = workbook.intern_style(Style {
            number_format: Some("[h]:mm:ss".to_string()),
            ..Default::default()
        });

        Self {
            date,
            datetime,
            time,
            duration,
        }
    }

    fn style_for_excel_datetime(self, dt: &calamine::ExcelDateTime) -> u32 {
        if dt.is_duration() {
            return self.duration;
        }

        // Calamine tells us the cell should be interpreted as a date/time (as
        // opposed to a raw number) but does not preserve the exact number format
        // string. We attempt to recover the precise format code via BIFF parsing,
        // but fall back to a best-effort heuristic when it isn't available.
        let serial = dt.as_f64();
        let frac = serial.abs().fract();

        if serial.abs() < 1.0 && frac != 0.0 {
            return self.time;
        }

        if frac == 0.0 {
            return self.date;
        }

        self.datetime
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SourceFormat {
    Xls,
}

impl SourceFormat {
    /// Default format to use when saving a workbook opened from this source.
    ///
    /// Writing legacy BIFF `.xls` files is out of scope; legacy imports default
    /// to `.xlsx`.
    pub const fn default_save_extension(self) -> &'static str {
        match self {
            SourceFormat::Xls => "xlsx",
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ImportSource {
    pub path: PathBuf,
    pub format: SourceFormat,
}

impl ImportSource {
    pub fn default_save_extension(&self) -> &'static str {
        self.format.default_save_extension()
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
/// Non-fatal warning emitted during `.xls` import.
///
/// The legacy BIFF importer is best-effort and continues when it encounters
/// malformed or unsupported records. These warnings are intended for user
/// visibility / diagnostics (for example, partial NOTE/OBJ/TXO comment records).
pub struct ImportWarning {
    pub message: String,
}

impl ImportWarning {
    fn new(message: impl Into<String>) -> Self {
        Self {
            message: message.into(),
        }
    }
}

/// A merged cell range in the source workbook.
///
/// Merged regions are also imported into [`formula_model::Worksheet::merged_regions`]. This type
/// is retained for backward compatibility with callers that still expect merged-range metadata
/// from [`XlsImportResult`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MergedRange {
    pub sheet_name: String,
    pub range: Range,
}

#[derive(Debug)]
/// Result of importing a legacy `.xls` workbook.
///
/// The returned [`Workbook`] contains the successfully imported data (values,
/// formulas, merged regions, row/column/outline metadata, hyperlinks, defined
/// names, and legacy comments/notes where available). Any non-fatal issues
/// encountered during import are returned in [`XlsImportResult::warnings`].
pub struct XlsImportResult {
    pub workbook: Workbook,
    pub source: ImportSource,
    pub merged_ranges: Vec<MergedRange>,
    pub warnings: Vec<ImportWarning>,
}

#[derive(Debug, Error)]
pub enum ImportError {
    #[error("failed to read `.xls`: {0}")]
    Xls(#[from] calamine::XlsError),
    #[error("encrypted workbook not supported: workbook is password-protected/encrypted; remove password protection in Excel and try again")]
    EncryptedWorkbook,
    #[error("invalid worksheet name: {0}")]
    InvalidSheetName(#[from] formula_model::SheetNameError),
}

/// Import a legacy `.xls` workbook from disk.
///
/// The importer is intentionally best-effort and attempts to load a subset of
/// BIFF features into [`formula_model`], including legacy cell notes/comments
/// (`NOTE/OBJ/TXO`, imported as [`formula_model::CommentKind::Note`]). Any
/// malformed or unsupported records may produce warnings rather than failing
/// the import.
pub fn import_xls_path(path: impl AsRef<Path>) -> Result<XlsImportResult, ImportError> {
    import_xls_path_with_biff_reader(path.as_ref(), biff::read_workbook_stream_from_xls)
}

/// Import a legacy `.xls` workbook from disk while treating BIFF workbook-stream parsing as
/// unavailable.
///
/// This is intended for testing the importer's best-effort fallback paths.
#[doc(hidden)]
pub fn import_xls_path_without_biff(path: impl AsRef<Path>) -> Result<XlsImportResult, ImportError> {
    import_xls_path_with_biff_reader(path.as_ref(), |_| Err("BIFF parsing disabled".to_string()))
}

fn import_xls_path_with_biff_reader(
    path: &Path,
    read_biff_workbook_stream: impl FnOnce(&Path) -> Result<Vec<u8>, String>,
) -> Result<XlsImportResult, ImportError> {
    let path = path.as_ref();
    // Best-effort: read the raw BIFF workbook stream up-front so we can detect
    // legacy `.xls` encryption (BIFF `FILEPASS`) before handing off to calamine.
    // Calamine does not support BIFF encryption and may return opaque parse
    // errors for password-protected workbooks.
    let mut warnings = Vec::new();
    let workbook_stream = match read_biff_workbook_stream(path) {
        Ok(bytes) => Some(bytes),
        Err(err) => {
            warnings.push(ImportWarning::new(format!(
                "failed to read `.xls` workbook stream: {err}"
            )));
            None
        }
    };
    let mut biff_version: Option<biff::BiffVersion> = None;
    let mut biff_codepage: Option<u16> = None;
    let mut biff_globals: Option<biff::globals::BiffWorkbookGlobals> = None;

    if let Some(workbook_stream) = workbook_stream.as_deref() {
        // Detect encrypted/password-protected `.xls` files before attempting to parse workbook
        // globals or opening via calamine. Encrypted BIFF streams contain a `FILEPASS` record in
        // the workbook globals substream.
        if biff::records::workbook_globals_has_filepass_record(workbook_stream) {
            return Err(ImportError::EncryptedWorkbook);
        }

        let detected_biff_version = biff::detect_biff_version(workbook_stream);
        let codepage = biff::parse_biff_codepage(workbook_stream);
        biff_version = Some(detected_biff_version);
        biff_codepage = Some(codepage);

        match biff::parse_biff_workbook_globals(workbook_stream, detected_biff_version, codepage) {
            Ok(globals) => {
                if globals.is_encrypted {
                    return Err(ImportError::EncryptedWorkbook);
                }
                biff_globals = Some(globals);
            }
            Err(err) => warnings.push(ImportWarning::new(format!(
                "failed to import `.xls` workbook globals: {err}"
            ))),
        }
    }

    // `calamine` can panic when parsing BIFF8 defined-name `NAME` (0x0018) records that are split
    // across `CONTINUE` records. Calamine reads the `cce` formula length from the NAME header, but
    // (incorrectly) assumes the entire token stream lives in the *first* physical record fragment.
    //
    // When a NAME record is continued, `cce` can exceed the first fragment length and calamine
    // panics while slicing `rgce`.
    //
    // Work around this by sanitizing *continued* NAME records in the workbook stream before
    // handing it to calamine:
    // - Zero out `NAME.cce` so calamine doesn't attempt to slice past the first fragment.
    // - (Best-effort) clamp `NAME.cch` if the name string itself would not fit in the first
    //   fragment.
    //
    // We still import defined names ourselves via BIFF parsing (including CONTINUE handling), so
    // calamine's defined-name formulas are not used for correctness here. However, calamine *does*
    // need the defined-name *names* table to decode `PtgName` tokens in worksheet formulas, so we
    // avoid masking out the NAME records entirely.
    let mut workbook: Xls<_> = match workbook_stream.as_deref() {
        Some(stream) => {
            let sanitized = sanitize_biff8_continued_name_records_for_calamine(stream);
            let xls_bytes = build_in_memory_xls(sanitized.as_deref().unwrap_or(stream))?;
            Xls::new(Cursor::new(xls_bytes))?
        }
        None => {
            let bytes =
                std::fs::read(path).map_err(|err| ImportError::Xls(calamine::XlsError::Io(err)))?;
            Xls::new(Cursor::new(bytes))?
        }
    };

    // We need to snapshot metadata (names, visibility, type) up-front because we
    // need mutable access to the workbook while iterating over ranges.
    let sheets: Vec<Sheet> = workbook.sheets_metadata().to_vec();
    // Snapshot defined names up-front because we need mutable access to the workbook while
    // iterating over ranges.
    let calamine_defined_names = workbook.defined_names().to_vec();

    let mut out = Workbook::new();
    let mut used_sheet_names: Vec<String> = Vec::new();
    let mut merged_ranges = Vec::new();

    let mut xf_style_ids: Option<Vec<u32>> = None;
    let mut xf_is_interesting: Option<Vec<bool>> = None;
    let mut sheet_tab_colors: Option<Vec<Option<TabColor>>> = None;
    let mut workbook_active_tab: Option<u16> = None;
    let mut biff_sheets: Option<Vec<biff::BoundSheetInfo>> = None;
    let mut row_col_props: Option<Vec<biff::SheetRowColProperties>> = None;
    let mut cell_xf_indices: Option<Vec<HashMap<CellRef, u16>>> = None;
    let mut cell_xf_parse_failed: Option<Vec<bool>> = None;
    let mut filter_database_ranges: Option<HashMap<usize, Range>> = None;

    if let Some(workbook_stream) = workbook_stream.as_deref() {
        let detected_biff_version =
            biff_version.unwrap_or_else(|| biff::detect_biff_version(workbook_stream));
        let codepage = biff_codepage.unwrap_or_else(|| biff::parse_biff_codepage(workbook_stream));
        biff_version.get_or_insert(detected_biff_version);
        biff_codepage.get_or_insert(codepage);

        if let Some(mut globals) = biff_globals.take() {
            out.date_system = globals.date_system;
            if let Some(mode) = globals.calculation_mode {
                out.calc_settings.calculation_mode = mode;
            }
            if let Some(value) = globals.calculate_before_save {
                out.calc_settings.calculate_before_save = value;
            }
            if let Some(enabled) = globals.iterative_enabled {
                out.calc_settings.iterative.enabled = enabled;
            }
            if let Some(max_iterations) = globals.iterative_max_iterations {
                out.calc_settings.iterative.max_iterations = max_iterations;
            }
            if let Some(max_change) = globals.iterative_max_change {
                out.calc_settings.iterative.max_change = max_change;
            }
            if let Some(full_precision) = globals.full_precision {
                out.calc_settings.full_precision = full_precision;
            }
            out.workbook_protection = std::mem::take(&mut globals.workbook_protection);
            workbook_active_tab = globals.active_tab_index;
            out.view.window = globals.workbook_window.take();
            warnings.extend(globals.warnings.drain(..).map(ImportWarning::new));
            sheet_tab_colors = Some(std::mem::take(&mut globals.sheet_tab_colors));

            let interesting = globals.xf_is_interesting_mask();
            xf_style_ids = Some(vec![0; interesting.len()]);
            xf_is_interesting = Some(interesting);
            biff_globals = Some(globals);
        }

        match biff::parse_biff_bound_sheets(workbook_stream, detected_biff_version, codepage) {
            Ok(sheets) => biff_sheets = Some(sheets),
            Err(err) => warnings.push(ImportWarning::new(format!(
                "failed to import `.xls` sheet metadata: {err}"
            ))),
        }

        // AutoFilter ranges are stored in a built-in workbook/worksheet defined name
        // (`_FilterDatabase`). Excel files in the wild use both workbook-scope and local-scope
        // definitions; decode them from the BIFF workbook globals stream.
        let sheet_count_for_autofilter = biff_sheets.as_ref().map(|s| s.len());
        match biff::parse_biff_filter_database_ranges(
            workbook_stream,
            detected_biff_version,
            codepage,
            sheet_count_for_autofilter,
        ) {
            Ok(parsed) => {
                let biff::ParsedFilterDatabaseRanges {
                    by_sheet,
                    warnings: biff_warnings,
                } = parsed;
                if !by_sheet.is_empty() {
                    filter_database_ranges = Some(by_sheet);
                }
                warnings.extend(biff_warnings.into_iter().map(|w| {
                    ImportWarning::new(format!("failed to import `.xls` autofilter: {w}"))
                }));
            }
            Err(err) => warnings.push(ImportWarning::new(format!(
                "failed to import `.xls` autofilter ranges: {err}"
            ))),
        }

        if let Some(sheets) = biff_sheets.as_ref() {
            let mut props_by_sheet = Vec::with_capacity(sheets.len());
            for sheet in sheets {
                if sheet.offset >= workbook_stream.len() {
                    warnings.push(ImportWarning::new(format!(
                        "failed to import `.xls` row/column properties for BIFF sheet index {} (`{}`): out-of-bounds stream offset {}",
                        props_by_sheet.len(),
                        sheet.name,
                        sheet.offset
                    )));
                    props_by_sheet.push(biff::SheetRowColProperties::default());
                    continue;
                }

                match biff::parse_biff_sheet_row_col_properties(workbook_stream, sheet.offset) {
                    Ok(mut props) => {
                        warnings.extend(props.warnings.drain(..).map(|warning| {
                            ImportWarning::new(format!(
                                "failed to fully import `.xls` row/column properties for BIFF sheet index {} (`{}`): {warning}",
                                props_by_sheet.len(),
                                sheet.name
                            ))
                        }));
                        props_by_sheet.push(props);
                    }
                    Err(parse_err) => {
                        warnings.push(ImportWarning::new(format!(
                            "failed to import `.xls` row/column properties for BIFF sheet index {} (`{}`): {parse_err}",
                            props_by_sheet.len(),
                            sheet.name
                        )));
                        props_by_sheet.push(biff::SheetRowColProperties::default());
                    }
                }
            }
            row_col_props = Some(props_by_sheet);

            if let Some(mask) = xf_is_interesting.as_deref() {
                // Even if the workbook contains no non-General number formats, scan for
                // out-of-range XF indices so corrupt files still surface a warning.
                if !mask.is_empty() {
                    let mut cell_xfs_by_sheet = Vec::with_capacity(sheets.len());
                    let mut parse_failed_by_sheet = Vec::with_capacity(sheets.len());
                    for sheet in sheets {
                        if sheet.offset >= workbook_stream.len() {
                            warnings.push(ImportWarning::new(format!(
                                "failed to import `.xls` cell styles for BIFF sheet index {} (`{}`): out-of-bounds stream offset {}",
                                cell_xfs_by_sheet.len(),
                                sheet.name,
                                sheet.offset
                            )));
                            cell_xfs_by_sheet.push(HashMap::new());
                            parse_failed_by_sheet.push(true);
                            continue;
                        }

                        match biff::parse_biff_sheet_cell_xf_indices_filtered(
                            workbook_stream,
                            sheet.offset,
                            Some(mask),
                        ) {
                            Ok(xfs) => {
                                cell_xfs_by_sheet.push(xfs);
                                parse_failed_by_sheet.push(false);
                            }
                            Err(parse_err) => {
                                warnings.push(ImportWarning::new(format!(
                                    "failed to import `.xls` cell styles for BIFF sheet index {} (`{}`): {parse_err}",
                                    cell_xfs_by_sheet.len(),
                                    sheet.name
                                )));
                                cell_xfs_by_sheet.push(HashMap::new());
                                parse_failed_by_sheet.push(true);
                            }
                        }
                    }

                    cell_xf_indices = Some(cell_xfs_by_sheet);
                    cell_xf_parse_failed = Some(parse_failed_by_sheet);
                }
            }
        }

        // Intern styles only for XF indices that were referenced by at least one "interesting" cell
        // record (including BLANK/MUL* records). This is best-effort; out-of-range XF indices are
        // ignored here but still preserved in `cell_xf_indices` so callers can surface warnings.
        if let (Some(globals), Some(style_ids), Some(cell_xfs_by_sheet)) = (
            biff_globals.as_ref(),
            xf_style_ids.as_mut(),
            cell_xf_indices.as_ref(),
        ) {
            let mut used = vec![false; style_ids.len()];
            for sheet_xfs in cell_xfs_by_sheet {
                for &xf_idx in sheet_xfs.values() {
                    let idx = xf_idx as usize;
                    if idx < used.len() {
                        used[idx] = true;
                    }
                }
            }

            for (xf_idx, style) in globals.resolve_styles_for_used_mask(&used) {
                if style == Style::default() {
                    continue;
                }
                style_ids[xf_idx] = out.intern_style(style);
            }
        }

        // Drop the BIFF globals cache once we've built the XF -> style_id mapping.
        drop(biff_globals.take());
    }

    // `calamine` may surface date/time serials as `Data::DateTime`, but it does not preserve the
    // exact BIFF number format code. We attempt to recover the precise format via BIFF parsing; if
    // that fails we fall back to a small set of heuristics (interned lazily).
    let mut date_time_styles: Option<DateTimeStyleIds> = None;

    let (sheet_mapping, mapping_warning) =
        reconcile_biff_sheet_mapping(&sheets, biff_sheets.as_deref());
    if let Some(message) = mapping_warning {
        warnings.push(ImportWarning::new(message));
    }

    let mut final_sheet_names_by_idx: Vec<String> = Vec::with_capacity(sheets.len());
    // Track worksheet ids so BIFF `itab` scopes can be mapped to the output model.
    let mut sheet_ids_by_calamine_idx: Vec<formula_model::WorksheetId> =
        Vec::with_capacity(sheets.len());

    for (sheet_idx, sheet_meta) in sheets.iter().enumerate() {
        let source_sheet_name = sheet_meta.name.clone();
        let biff_idx = sheet_mapping.get(sheet_idx).copied().flatten();

        let sheet_cell_xfs_raw =
            biff_idx.and_then(|biff_idx| cell_xf_indices.as_ref().and_then(|v| v.get(biff_idx)));

        let sheet_cell_xfs = sheet_cell_xfs_raw.filter(|map| !map.is_empty());

        // If BIFF styles are unavailable (or corrupt), fall back to heuristic date/time formats for
        // any `Data::DateTime` cells that would otherwise have no style.
        let sheet_has_out_of_range_xf = match (xf_style_ids.as_deref(), sheet_cell_xfs_raw) {
            (Some(style_ids), Some(map)) => map
                .values()
                .any(|&xf_idx| xf_idx as usize >= style_ids.len()),
            _ => false,
        };

        let sheet_xf_parse_failed = biff_idx
            .and_then(|biff_idx| cell_xf_parse_failed.as_ref().and_then(|v| v.get(biff_idx)))
            .copied()
            .unwrap_or(false);

        let sheet_needs_datetime_fallback = xf_style_ids.is_none()
            || sheet_cell_xfs_raw.is_none()
            || sheet_has_out_of_range_xf
            || sheet_xf_parse_failed;

        let value_range = match workbook.worksheet_range(&source_sheet_name) {
            Ok(range) => Some(range),
            Err(err) => {
                warnings.push(ImportWarning::new(format!(
                    "failed to read cell values for sheet `{source_sheet_name}`: {err}"
                )));
                None
            }
        };

        if sheet_needs_datetime_fallback && date_time_styles.is_none() {
            date_time_styles = Some(DateTimeStyleIds::new(&mut out));
        }

        let sheet_date_time_styles = if sheet_needs_datetime_fallback {
            date_time_styles
        } else {
            None
        };

        let (sheet_id, sheet_name) = match out.add_sheet(source_sheet_name.clone()) {
            Ok(sheet_id) => {
                used_sheet_names.push(source_sheet_name.clone());
                (sheet_id, source_sheet_name.clone())
            }
            Err(err) => {
                let mut candidate =
                    sanitize_sheet_name(&source_sheet_name, sheet_idx + 1, &used_sheet_names);
                let sheet_id = loop {
                    match out.add_sheet(candidate.clone()) {
                        Ok(sheet_id) => break sheet_id,
                        Err(_) => {
                            // If our best-effort sanitization still collides (e.g. due to
                            // case-insensitive comparisons), treat the candidate as taken and
                            // generate another.
                            let mut augmented = used_sheet_names.clone();
                            augmented.push(candidate);
                            candidate =
                                sanitize_sheet_name(&source_sheet_name, sheet_idx + 1, &augmented);
                        }
                    }
                };

                warnings.push(ImportWarning::new(format!(
                    "sanitized sheet name `{source_sheet_name}` -> `{candidate}` ({err})"
                )));
                used_sheet_names.push(candidate.clone());
                (sheet_id, candidate)
            }
        };
        final_sheet_names_by_idx.push(sheet_name.clone());
        sheet_ids_by_calamine_idx.push(sheet_id);
        let sheet = out
            .sheet_mut(sheet_id)
            .expect("sheet id should exist immediately after add");

        sheet.visibility = sheet_visible_to_visibility(sheet_meta.visible);
        sheet.tab_color = biff_idx
            .and_then(|idx| sheet_tab_colors.as_ref().and_then(|v| v.get(idx)).cloned())
            .flatten();

        if sheet_meta.typ != SheetType::WorkSheet {
            warnings.push(ImportWarning::new(format!(
                "sheet `{sheet_name}` has unsupported type {:?}; importing as worksheet",
                sheet_meta.typ
            )));
        }

        if let (Some(workbook_stream), Some(biff_idx)) = (workbook_stream.as_deref(), biff_idx) {
            if let Some(sheet_info) = biff_sheets.as_ref().and_then(|v| v.get(biff_idx)) {
                if sheet_info.offset >= workbook_stream.len() {
                    warnings.push(ImportWarning::new(format!(
                        "failed to import `.xls` view state for BIFF sheet index {} (`{}`): out-of-bounds stream offset {}",
                        biff_idx, sheet_name, sheet_info.offset
                    )));
                } else {
                    match biff::parse_biff_sheet_view_state(workbook_stream, sheet_info.offset) {
                        Ok(mut view_state) => {
                            warnings.extend(view_state.warnings.drain(..).map(ImportWarning::new));

                            if let Some(show) = view_state.show_grid_lines {
                                sheet.view.show_grid_lines = show;
                            }
                            if let Some(show) = view_state.show_headings {
                                sheet.view.show_headings = show;
                            }
                            if let Some(show) = view_state.show_zeros {
                                sheet.view.show_zeros = show;
                            }
                            if let Some(zoom) = view_state.zoom {
                                sheet.zoom = zoom;
                                sheet.view.zoom = zoom;
                            }
                            if let Some(pane) = view_state.pane {
                                sheet.frozen_rows = pane.frozen_rows;
                                sheet.frozen_cols = pane.frozen_cols;
                                sheet.view.pane = pane;
                            }
                            if let Some(selection) = view_state.selection {
                                sheet.view.selection = Some(selection);
                            }
                        }
                        Err(err) => warnings.push(ImportWarning::new(format!(
                            "failed to import `.xls` view state for BIFF sheet index {} (`{}`): {err}",
                            biff_idx, sheet_name
                        ))),
                    }

                    match biff::parse_biff_sheet_protection(workbook_stream, sheet_info.offset) {
                        Ok(mut protection) => {
                            warnings.extend(protection.warnings.drain(..).map(ImportWarning::new));
                            sheet.sheet_protection = protection.protection;
                        }
                        Err(err) => warnings.push(ImportWarning::new(format!(
                            "failed to import `.xls` sheet protection for BIFF sheet index {} (`{}`): {err}",
                            biff_idx, sheet_name
                        ))),
                    }
                }
            }
        }

        if let Some(biff_idx) = biff_idx {
            if let Some(props) = row_col_props
                .as_ref()
                .and_then(|props_by_sheet| props_by_sheet.get(biff_idx))
            {
                apply_row_col_properties(sheet, props);
                apply_outline_properties(sheet, props);
            }

            if sheet.auto_filter.is_none() {
                if let Some(range) = filter_database_ranges
                    .as_ref()
                    .and_then(|ranges| ranges.get(&biff_idx))
                    .copied()
                {
                    sheet.auto_filter = Some(SheetAutoFilter {
                        range,
                        filter_columns: Vec::new(),
                        sort_state: None,
                        raw_xml: Vec::new(),
                    });
                }
            }
        }

        // Merged regions: prefer calamine's parsed merge metadata, but fall back to scanning the
        // worksheet BIFF substream for `MERGEDCELLS` records when calamine provides none.
        let mut merge_ranges: Vec<Range> = Vec::new();
        if let Some(merge_cells) = workbook.worksheet_merge_cells(&source_sheet_name) {
            for dim in merge_cells {
                merge_ranges.push(Range::new(
                    CellRef::new(dim.start.0, dim.start.1),
                    CellRef::new(dim.end.0, dim.end.1),
                ));
            }
        }

        // Best-effort fallback when calamine does not surface any merged-cell ranges.
        if merge_ranges.is_empty() {
            if let (Some(workbook_stream), Some(biff_idx)) = (workbook_stream.as_deref(), biff_idx) {
                if let Some(sheet_info) = biff_sheets.as_ref().and_then(|s| s.get(biff_idx)) {
                    if sheet_info.offset >= workbook_stream.len() {
                        warnings.push(ImportWarning::new(format!(
                            "failed to import `.xls` merged cells for sheet `{sheet_name}`: out-of-bounds stream offset {}",
                            sheet_info.offset
                        )));
                    } else {
                        match biff::parse_biff_sheet_merged_cells(
                            workbook_stream,
                            sheet_info.offset,
                        ) {
                            Ok(mut ranges) => {
                                if !ranges.is_empty() {
                                    merge_ranges.append(&mut ranges);
                                }
                            }
                            Err(err) => warnings.push(ImportWarning::new(format!(
                                "failed to import `.xls` merged cells for sheet `{sheet_name}`: {err}"
                            ))),
                        }
                    }
                }
            }
        }

        if !merge_ranges.is_empty() {
            let mut seen: HashSet<Range> = HashSet::new();
            for range in merge_ranges {
                if !seen.insert(range) {
                    continue;
                }

                // Populate the model's merged-region table so cell resolution matches Excel.
                if let Err(err) = sheet.merged_regions.add(range) {
                    warnings.push(ImportWarning::new(format!(
                        "failed to add merged region `{range}` in sheet `{sheet_name}`: {err}"
                    )));
                }

                // Back-compat: preserve merged metadata on the import result.
                merged_ranges.push(MergedRange {
                    sheet_name: sheet_name.clone(),
                    range,
                });
            }
        }

        if let Some(range) = value_range.as_ref() {
            let range_start = range.start().unwrap_or((0, 0));

            for (row, col, value) in range.used_cells() {
                let Some(cell_ref) = to_cell_ref(range_start, row, col) else {
                    warnings.push(ImportWarning::new(format!(
                        "skipping out-of-bounds cell in sheet `{sheet_name}` at ({row},{col})"
                    )));
                    continue;
                };

                let anchor = sheet.merged_regions.resolve_cell(cell_ref);
                let Some((value, mut style_id)) = convert_value(value, sheet_date_time_styles)
                else {
                    continue;
                };

                if let Some(resolved) =
                    style_id_for_cell_xf(xf_style_ids.as_deref(), sheet_cell_xfs, anchor)
                {
                    style_id = Some(resolved);
                }

                sheet.set_value(anchor, value);
                if let Some(style_id) = style_id {
                    sheet.set_style_id(anchor, style_id);
                }
            }
        }

        // Extract BIFF hyperlinks after merged regions have been populated so callers can resolve
        // anchors consistently with the model's merged-cell semantics.
        if let (Some(workbook_stream), Some(codepage), Some(biff_idx)) =
            (workbook_stream.as_deref(), biff_codepage, biff_idx)
        {
            if let Some(sheet_info) = biff_sheets.as_ref().and_then(|s| s.get(biff_idx)) {
                if sheet_info.offset >= workbook_stream.len() {
                    warnings.push(ImportWarning::new(format!(
                        "failed to import `.xls` hyperlinks for sheet `{sheet_name}`: out-of-bounds stream offset {}",
                        sheet_info.offset
                    )));
                } else {
                    match biff::parse_biff_sheet_hyperlinks(
                        workbook_stream,
                        sheet_info.offset,
                        codepage,
                    ) {
                        Ok(parsed) => {
                            let hyperlink_start = sheet.hyperlinks.len();
                            sheet.hyperlinks.extend(parsed.hyperlinks);
                            // Excel treats hyperlinks applied to a merged cell as applying to the
                            // entire merged region. When BIFF stores the hyperlink as a single-cell
                            // anchor inside a merge, expand it so `Worksheet::hyperlink_at` works
                            // for any cell inside the merged region.
                            if !sheet.merged_regions.is_empty() && hyperlink_start < sheet.hyperlinks.len() {
                                for link in sheet.hyperlinks[hyperlink_start..].iter_mut() {
                                    if link.range.is_single_cell() {
                                        if let Some(merged) = sheet.merged_regions.containing_range(link.range.start) {
                                            link.range = merged;
                                        }
                                    }
                                }
                            }
                            warnings.extend(parsed.warnings.into_iter().map(|w| {
                                ImportWarning::new(format!(
                                    "failed to import `.xls` hyperlink in sheet `{sheet_name}`: {w}"
                                ))
                            }));
                        }
                        Err(err) => warnings.push(ImportWarning::new(format!(
                            "failed to import `.xls` hyperlinks for sheet `{sheet_name}`: {err}"
                        ))),
                    }
                }
            }
        }

        // Extract legacy cell comments ("notes") encoded via NOTE/OBJ/TXO records in the sheet
        // BIFF substream. This runs after merged regions have been populated so
        // `Worksheet::add_comment` can normalize anchors inside merged regions.
        if let (Some(workbook_stream), Some(codepage), Some(biff_version), Some(biff_idx)) = (
            workbook_stream.as_deref(),
            biff_codepage,
            biff_version,
            biff_idx,
        ) {
            if let Some(sheet_info) = biff_sheets.as_ref().and_then(|s| s.get(biff_idx)) {
                if sheet_info.offset >= workbook_stream.len() {
                    warnings.push(ImportWarning::new(format!(
                        "failed to import `.xls` note comments for sheet `{sheet_name}`: out-of-bounds stream offset {}",
                        sheet_info.offset
                    )));
                } else {
                    match biff::parse_biff_sheet_notes(
                        workbook_stream,
                        sheet_info.offset,
                        biff_version,
                        codepage,
                    ) {
                        Ok(parsed) => {
                            warnings.extend(parsed.warnings.into_iter().map(|w| {
                                ImportWarning::new(format!(
                                    "failed to import `.xls` note comment in sheet `{sheet_name}`: {w}"
                                ))
                            }));
                            let notes = parsed.notes;

                            // Generate deterministic comment ids for BIFF NOTE comments so repeated
                            // imports of the same `.xls` file yield stable ids.
                            //
                            // NOTE records are linked to TXO text payloads via `obj_id`, but some
                            // files contain duplicate object ids or other collisions. In those
                            // cases, fall back to `Worksheet::add_comment`'s UUID generation by
                            // leaving the id empty and surface a warning.
                            let mut seen_obj_ids: HashMap<u16, CellRef> = HashMap::new();
                            let mut seen_ids: HashSet<String> = HashSet::new();
                            for note in notes {
                                let anchor = sheet.merged_regions.resolve_cell(note.cell);
                                let candidate_id =
                                    format!("xls-note:{}:{}", anchor.to_a1(), note.obj_id);

                                let mut collision = false;
                                if let Some(prev_cell) = seen_obj_ids.get(&note.obj_id).copied() {
                                    collision = true;
                                    warnings.push(ImportWarning::new(format!(
                                        "duplicate `.xls` NOTE obj_id {} in sheet `{sheet_name}` (index {sheet_idx}) (cell {} already used at {}); generating random comment id",
                                        note.obj_id,
                                        anchor.to_a1(),
                                        prev_cell.to_a1(),
                                    )));
                                }
                                if seen_ids.contains(&candidate_id) {
                                    collision = true;
                                    warnings.push(ImportWarning::new(format!(
                                        "duplicate `.xls` NOTE id `{candidate_id}` in sheet `{sheet_name}` (index {sheet_idx}); generating random comment id",
                                    )));
                                }

                                let id = if collision {
                                    String::new()
                                } else {
                                    seen_obj_ids.insert(note.obj_id, anchor);
                                    seen_ids.insert(candidate_id.clone());
                                    candidate_id
                                };

                                let mut comment = Comment {
                                    id,
                                    kind: CommentKind::Note,
                                    content: note.text,
                                    author: CommentAuthor {
                                        name: note.author,
                                        ..Default::default()
                                    },
                                    ..Default::default()
                                };

                                match sheet.add_comment(anchor, comment.clone()) {
                                    Ok(_) => {}
                                    Err(formula_model::CommentError::DuplicateCommentId(dup_id)) => {
                                        warnings.push(ImportWarning::new(format!(
                                            "duplicate `.xls` comment id `{dup_id}` in sheet `{sheet_name}` (index {sheet_idx}); generating random comment id",
                                        )));
                                    comment.id.clear();
                                        if let Err(err) = sheet.add_comment(anchor, comment) {
                                        warnings.push(ImportWarning::new(format!(
                                            "failed to import `.xls` note comment for sheet `{sheet_name}` (index {sheet_idx}) at {}: {err}",
                                                anchor.to_a1(),
                                        )));
                                    }
                                }
                                    Err(err) => warnings.push(ImportWarning::new(format!(
                                        "failed to import `.xls` note comment for sheet `{sheet_name}` (index {sheet_idx}) at {}: {err}",
                                        anchor.to_a1(),
                                    ))),
                                }
                            }
                        }
                        Err(err) => warnings.push(ImportWarning::new(format!(
                            "failed to import `.xls` note comments for sheet `{sheet_name}`: {err}"
                        ))),
                    }
                }
            }
        }
        match workbook.worksheet_formula(&source_sheet_name) {
            Ok(formula_range) => {
                let formula_start = formula_range.start().unwrap_or((0, 0));
                for (row, col, formula) in formula_range.used_cells() {
                    let Some(cell_ref) = to_cell_ref(formula_start, row, col) else {
                        warnings.push(ImportWarning::new(format!(
                            "skipping out-of-bounds formula in sheet `{sheet_name}` at ({row},{col})"
                        )));
                        continue;
                    };

                    // Calamine may surface BIFF8 Unicode strings with embedded NUL bytes (notably
                    // defined-name references via `PtgName`). Strip them so the formula text is
                    // parseable and stable across import paths.
                    let formula_clean;
                    let formula = if formula.contains('\0') {
                        formula_clean = formula.replace('\0', "");
                        formula_clean.as_str()
                    } else {
                        formula
                    };

                    let Some(normalized) = normalize_formula_text(formula) else {
                        continue;
                    };

                    let anchor = sheet.merged_regions.resolve_cell(cell_ref);
                    sheet.set_formula(anchor, Some(normalized));

                    if let Some(resolved) =
                        style_id_for_cell_xf(xf_style_ids.as_deref(), sheet_cell_xfs, anchor)
                    {
                        sheet.set_style_id(anchor, resolved);
                    }
                }
            }
            Err(err) => warnings.push(ImportWarning::new(format!(
                "failed to read formulas for sheet `{sheet_name}`: {err}"
            ))),
        }

        // `calamine` does not surface `BLANK` records via `used_cells()`, but Excel
        // allows formatting empty cells. Apply any XF-derived number formats to
        // the sheet even when the value is empty so those cells round-trip.
        if let (Some(xf_style_ids), Some(sheet_cell_xfs)) =
            (xf_style_ids.as_deref(), sheet_cell_xfs)
        {
            let mut out_of_range_xf_count: usize = 0;
            if sheet.merged_regions.is_empty() {
                for (&cell_ref, &xf_idx) in sheet_cell_xfs {
                    if cell_ref.row >= EXCEL_MAX_ROWS || cell_ref.col >= EXCEL_MAX_COLS {
                        continue;
                    }

                    let Some(&style_id) = xf_style_ids.get(xf_idx as usize) else {
                        out_of_range_xf_count = out_of_range_xf_count.saturating_add(1);
                        continue;
                    };
                    if style_id == 0 {
                        continue;
                    };

                    if sheet
                        .cell(cell_ref)
                        .is_some_and(|cell| cell.style_id == style_id)
                    {
                        continue;
                    }
                    sheet.set_style_id(cell_ref, style_id);
                }
            } else {
                // We potentially see XF indices for multiple cells that resolve to the same merged-cell
                // anchor. When that happens, prefer the anchor cell’s own resolvable style; otherwise
                // choose the first cell (row/col order) within the merged region to keep this pass
                // deterministic (independent of HashMap iteration order).
                let mut best_style_for_anchor: HashMap<CellRef, (u32, CellRef, bool)> =
                    HashMap::new();

                for (&cell_ref, &xf_idx) in sheet_cell_xfs {
                    if cell_ref.row >= EXCEL_MAX_ROWS || cell_ref.col >= EXCEL_MAX_COLS {
                        continue;
                    }

                    let maybe_anchor = sheet.merged_regions.anchor_for(cell_ref);
                    // Normalise style assignments to merged-cell anchors so formatting inside a merged
                    // region round-trips consistently with the importer’s value/formula semantics.
                    let anchor = maybe_anchor.unwrap_or(cell_ref);

                    let Some(&style_id) = xf_style_ids.get(xf_idx as usize) else {
                        out_of_range_xf_count = out_of_range_xf_count.saturating_add(1);
                        continue;
                    };
                    if style_id == 0 {
                        continue;
                    };
                    if maybe_anchor.is_none() {
                        if sheet
                            .cell(anchor)
                            .is_some_and(|cell| cell.style_id == style_id)
                        {
                            continue;
                        }
                        sheet.set_style_id(anchor, style_id);
                        continue;
                    }

                    let source_is_anchor = cell_ref == anchor;
                    best_style_for_anchor
                        .entry(anchor)
                        .and_modify(|(existing_style_id, existing_source, existing_is_anchor)| {
                            if *existing_is_anchor {
                                // Anchor-derived styles always win.
                                return;
                            }
                            if source_is_anchor {
                                *existing_style_id = style_id;
                                *existing_source = cell_ref;
                                *existing_is_anchor = true;
                                return;
                            }
                            if (cell_ref.row, cell_ref.col)
                                < (existing_source.row, existing_source.col)
                            {
                                *existing_style_id = style_id;
                                *existing_source = cell_ref;
                            }
                        })
                        .or_insert((style_id, cell_ref, source_is_anchor));
                }

                for (anchor, (style_id, _, _)) in best_style_for_anchor {
                    if sheet
                        .cell(anchor)
                        .is_some_and(|cell| cell.style_id == style_id)
                    {
                        continue;
                    }
                    sheet.set_style_id(anchor, style_id);
                }
            }

            if out_of_range_xf_count > 0 {
                warnings.push(ImportWarning::new(format!(
                    "skipped {out_of_range_xf_count} cells in sheet `{sheet_name}` with out-of-range XF indices"
                )));
            }
        }
    }

    // Track per-sheet rename pairs (old sheet name -> final imported sheet name) so we can rewrite
    // sheet-qualified strings (cell formulas, internal hyperlinks, calamine-defined-name fallback)
    // after best-effort sheet-name sanitization.
    let mut sheet_rename_pairs: Vec<(String, String)> = Vec::new();

    // If we had to sanitize sheet names, internal hyperlinks and cell formulas may still
    // reference the original (invalid) sheet names. Rewrite those references to point at the
    // final imported sheet names so navigation and formulas remain correct after import and
    // round-trips to XLSX.
    if !final_sheet_names_by_idx.is_empty() {
        let mut resolved_sheet_names: HashMap<String, String> = HashMap::new();

        for (idx, sheet_meta) in sheets.iter().enumerate() {
            let Some(final_name) = final_sheet_names_by_idx.get(idx) else {
                continue;
            };

            // For hyperlink targets we resolve case-insensitively and strip embedded NULs so we
            // can match calamine's decoded sheet names against BIFF's BoundSheet names.
            resolved_sheet_names.insert(
                normalize_sheet_name_for_match(&sheet_meta.name),
                final_name.clone(),
            );

            // For formula rewriting we use exact old sheet name strings (the rewrite helper
            // handles case-insensitive matching internally).
            if sheet_meta.name != *final_name {
                sheet_rename_pairs.push((sheet_meta.name.clone(), final_name.clone()));
            }

            // Add a BIFF BoundSheet name alias when available: calamine sheet metadata and BIFF
            // BoundSheet names can diverge due to encoding issues or malformed files.
            if let Some(biff_idx) = sheet_mapping.get(idx).copied().flatten() {
                if let Some(biff_name) = biff_sheets
                    .as_ref()
                    .and_then(|sheets| sheets.get(biff_idx))
                    .map(|s| s.name.as_str())
                {
                    resolved_sheet_names
                        .entry(normalize_sheet_name_for_match(biff_name))
                        .or_insert_with(|| final_name.clone());

                    if biff_name != final_name && biff_name != sheet_meta.name {
                        sheet_rename_pairs.push((biff_name.to_string(), final_name.clone()));
                    }
                }
            }
        }

        if !resolved_sheet_names.is_empty() {
            // Rewrite internal hyperlink targets that refer to source sheet names.
            for sheet in &mut out.sheets {
                for link in &mut sheet.hyperlinks {
                    let HyperlinkTarget::Internal { sheet, .. } = &mut link.target else {
                        continue;
                    };
                    let key = normalize_sheet_name_for_match(sheet);
                    if let Some(resolved) = resolved_sheet_names.get(&key) {
                        *sheet = resolved.clone();
                    }
                }
            }
        }

        if !sheet_rename_pairs.is_empty() {
            formula_rewrite::rewrite_workbook_formulas_for_sheet_renames(
                &mut out,
                &sheet_rename_pairs,
            );
        }
    }

    if let Some(i_tab_cur) = workbook_active_tab {
        let biff_tab_idx = i_tab_cur as usize;

        // `WINDOW1.iTabCur` is a BIFF sheet index (BoundSheet order). Prefer mapping it back to the
        // imported sheet order (calamine metadata order) using the same reconciliation mapping we
        // apply for other per-sheet BIFF metadata.
        let imported_idx = sheet_mapping
            .iter()
            .position(|mapped| mapped.is_some_and(|biff_idx| biff_idx == biff_tab_idx))
            .unwrap_or(biff_tab_idx);

        if let Some(sheet) = out.sheets.get(imported_idx) {
            out.view.active_sheet_id = Some(sheet.id);
        } else {
            warnings.push(ImportWarning::new(format!(
                "skipping `.xls` active tab index {i_tab_cur}: workbook contains {} imported sheets",
                out.sheets.len()
            )));
        }
    }

    // Import defined names (workbook- and sheet-scoped).
    if let (Some(workbook_stream), Some(biff_version), Some(codepage)) =
        (workbook_stream.as_deref(), biff_version, biff_codepage)
    {
        // Resolve BIFF sheet indices to the sheet names used by our output workbook.
        let mut sheet_names_by_biff_idx: Vec<String> = biff_sheets
            .as_deref()
            .unwrap_or_default()
            .iter()
            .map(|s| s.name.clone())
            .collect();

        for (cal_idx, maybe_biff_idx) in sheet_mapping.iter().enumerate() {
            let Some(biff_idx) = *maybe_biff_idx else {
                continue;
            };
            let Some(final_name) = final_sheet_names_by_idx.get(cal_idx) else {
                continue;
            };
            if biff_idx < sheet_names_by_biff_idx.len() {
                sheet_names_by_biff_idx[biff_idx] = final_name.clone();
            }
        }

        // Resolve BIFF sheet indices to WorksheetIds.
        let mut sheet_ids_by_biff_idx: Vec<Option<formula_model::WorksheetId>> =
            vec![None; sheet_names_by_biff_idx.len()];
        for (cal_idx, maybe_biff_idx) in sheet_mapping.iter().enumerate() {
            let Some(biff_idx) = *maybe_biff_idx else {
                continue;
            };
            let Some(sheet_id) = sheet_ids_by_calamine_idx.get(cal_idx).copied() else {
                continue;
            };
            if biff_idx < sheet_ids_by_biff_idx.len() {
                sheet_ids_by_biff_idx[biff_idx] = Some(sheet_id);
            }
        }

        // Resolve BIFF sheet indices to XLSX `localSheetId` values (0-based in workbook sheet
        // order). This is preserved in the model for round-trip fidelity when converting
        // `.xls` -> `.xlsx`.
        let mut local_sheet_ids_by_biff_idx: Vec<Option<u32>> =
            vec![None; sheet_ids_by_biff_idx.len()];
        for (cal_idx, maybe_biff_idx) in sheet_mapping.iter().enumerate() {
            let Some(biff_idx) = *maybe_biff_idx else {
                continue;
            };
            let cal_idx_u32: u32 = match cal_idx.try_into() {
                Ok(v) => v,
                Err(_) => continue,
            };
            if biff_idx < local_sheet_ids_by_biff_idx.len() {
                local_sheet_ids_by_biff_idx[biff_idx] = Some(cal_idx_u32);
            }
        }

        match biff::parse_biff_defined_names(
            workbook_stream,
            biff_version,
            codepage,
            &sheet_names_by_biff_idx,
        ) {
            Ok(mut parsed) => {
                warnings.extend(parsed.warnings.drain(..).map(ImportWarning::new));

                for name in parsed.names.drain(..) {
                    let (scope, xlsx_local_sheet_id) = match name.scope_sheet {
                        None => (DefinedNameScope::Workbook, None),
                        Some(biff_idx) => match sheet_ids_by_biff_idx.get(biff_idx).copied().flatten() {
                            Some(sheet_id) => (
                                DefinedNameScope::Sheet(sheet_id),
                                local_sheet_ids_by_biff_idx.get(biff_idx).copied().flatten(),
                            ),
                            None => {
                                warnings.push(ImportWarning::new(format!(
                                    "defined name `{}` has out-of-range sheet scope itab={} (sheet count={}); importing as workbook-scoped",
                                    name.name,
                                    biff_idx.saturating_add(1),
                                    sheet_ids_by_biff_idx.len()
                                )));
                                (DefinedNameScope::Workbook, None)
                            }
                        },
                    };

                    if let Err(err) = out.create_defined_name(
                        scope,
                        name.name.clone(),
                        name.refers_to.clone(),
                        name.comment.clone(),
                        name.hidden,
                        xlsx_local_sheet_id,
                    ) {
                        warnings.push(ImportWarning::new(format!(
                            "skipping defined name `{}`: {err}",
                            name.name
                        )));
                    }
                }
            }
            Err(err) => warnings.push(ImportWarning::new(format!(
                "failed to import `.xls` defined names: {err}"
            ))),
        }
    }

    // Calamine fallback: defined names.
    //
    // Calamine can surface defined names even when BIFF workbook parsing isn't available (or the
    // BIFF NAME parser fails). The `.xls` API does not expose sheet scope metadata (itab), so we
    // treat these names as workbook-scoped.
    //
    // Deterministic precedence: if a name already exists (e.g. imported via BIFF), skip the
    // calamine definition.
    let defined_names_before_calamine = out.defined_names.len();
    let mut imported_count: usize = 0;
    let mut skipped_count: usize = 0;
    for (name, refers_to) in calamine_defined_names {
        let name = normalize_calamine_defined_name_name(&name);
        let refers_to = refers_to.trim();
        let refers_to = refers_to
            .strip_prefix('=')
            .unwrap_or(refers_to)
            .to_string();

        // Defined names can contain sheet references, and those can point at sheet names that were
        // later sanitized during import. Rewrite any sheet-qualified references using the same
        // sheet rename mapping we apply to cells.
        let mut refers_to = refers_to;
        if !sheet_rename_pairs.is_empty() {
            let original = refers_to.clone();
            let mut rewritten = refers_to.clone();
            // Apply rename pairs in reverse order so we don't cascade rewrites when a sanitized
            // name collides with another sheet's original name (see comment in the cell-formula
            // rewrite pass above).
            for (old_name, new_name) in sheet_rename_pairs.iter().rev() {
                rewritten =
                    formula_model::rewrite_sheet_names_in_formula(&rewritten, old_name, new_name);
            }

            // `refers_to` is stored without a leading `=`; normalize defensively after rewriting.
            // If rewriting produces an empty string, keep the original.
            if let Some(normalized) = normalize_formula_text(&rewritten) {
                refers_to = normalized;
            } else {
                refers_to = original;
            }
        }

        // When BIFF defined names were imported successfully, prefer them over calamine’s
        // best-effort string representation.
        if defined_names_before_calamine != 0
            && out
                .defined_names
                .iter()
                .any(|existing| existing.name.eq_ignore_ascii_case(&name))
        {
            continue;
        }

        match out.create_defined_name(
            DefinedNameScope::Workbook,
            name.clone(),
            refers_to,
            None,
            false,
            None,
        ) {
            Ok(_) => imported_count = imported_count.saturating_add(1),
            Err(err) => {
                skipped_count = skipped_count.saturating_add(1);
                warnings.push(ImportWarning::new(format!(
                    "skipping `.xls` defined name `{name}` from calamine fallback: {err}"
                )));
            }
        }
    }

    if defined_names_before_calamine == 0 && imported_count > 0 {
        warnings.push(ImportWarning::new(
            "imported `.xls` defined names via calamine fallback; defined name metadata may be incomplete (scope/hidden/comment may be missing)",
        ));
    }

    if skipped_count > 0 {
        warnings.push(ImportWarning::new(format!(
            "skipped {skipped_count} `.xls` defined names from calamine fallback due to invalid/duplicate names"
        )));
    }
    populate_print_settings_from_defined_names(&mut out, &mut warnings);

    // Best-effort import of worksheet AutoFilter ranges (phase 1).
    //
    // In BIFF, the filtered range is typically stored as a hidden built-in defined name
    // `_xlnm._FilterDatabase`, scoped to the worksheet that owns the AutoFilter.
    //
    // We only import the presence + range in phase 1 (filter criteria and sort state are not yet
    // supported). Never fail import due to AutoFilter parsing.
    let mut autofilters: Vec<(formula_model::WorksheetId, Range)> = Vec::new();
    for name in &out.defined_names {
        if !name.name.eq_ignore_ascii_case(XLNM_FILTER_DATABASE) {
            continue;
        }

        let range = match parse_autofilter_range_from_defined_name(&name.refers_to) {
            Ok(range) => range,
            Err(err) => {
                warnings.push(ImportWarning::new(format!(
                    "failed to parse `.xls` AutoFilter range `{}` from defined name `{}`: {err}",
                    name.refers_to, name.name
                )));
                continue;
            }
        };

        // AutoFilter ranges are expected to be sheet-scoped. When BIFF scope metadata is
        // unavailable (calamine fallback), attempt to infer the owning sheet from a sheet-qualified
        // `refers_to` string like `Sheet1!$A$1:$C$10`.
        let sheet_id = match name.scope {
            DefinedNameScope::Sheet(sheet_id) => Some(sheet_id),
            DefinedNameScope::Workbook => {
                let mut refers_to = name.refers_to.trim();
                if let Some(rest) = refers_to.strip_prefix('=') {
                    refers_to = rest.trim();
                }
                if let Some(rest) = refers_to.strip_prefix('@') {
                    refers_to = rest.trim();
                }
                match refers_to.rsplit_once('!') {
                    Some((lhs, _)) => {
                        // Best-effort parse of sheet name, handling quoted identifiers (`'My Sheet'!A1`)
                        // and external-workbook prefixes (`[Book]Sheet1!A1`).
                        let mut sheet_name = lhs.trim();
                        if let Some((_, rest)) = sheet_name.rsplit_once(']') {
                            sheet_name = rest.trim();
                        }
                        let mut sheet_name = if let Some(inner) = sheet_name
                            .strip_prefix('\'')
                            .and_then(|s| s.strip_suffix('\''))
                        {
                            inner.replace("''", "'")
                        } else {
                            sheet_name.to_string()
                        };
                        if let Some((first, _)) = sheet_name.split_once(':') {
                            // Excel forbids `:` in sheet names, so this can only be a 3D sheet span.
                            sheet_name = first.to_string();
                        }

                        let sheet_id = out
                            .sheets
                            .iter()
                            .find(|s| sheet_name_eq_case_insensitive(&s.name, &sheet_name))
                            .map(|s| s.id);
                        if sheet_id.is_none() {
                            warnings.push(ImportWarning::new(format!(
                                "skipping `.xls` AutoFilter defined name `{}`: unknown sheet `{}` in `{}`",
                                name.name, sheet_name, name.refers_to
                            )));
                        }

                        sheet_id
                    }
                    None => {
                        if out.sheets.len() == 1 {
                            Some(out.sheets[0].id)
                        } else {
                            warnings.push(ImportWarning::new(format!(
                                "skipping `.xls` AutoFilter defined name `{}`: expected sheet-qualified range, got `{}`",
                                name.name, name.refers_to
                            )));
                            None
                        }
                    }
                }
            }
        };

        let Some(sheet_id) = sheet_id else {
            continue;
        };

        autofilters.push((sheet_id, range));
    }

    // Calamine's `.xls` defined-name support does not handle BIFF8 built-in `NAME` records (the
    // `fBuiltin` flag) correctly and also cannot decode non-3D area refs like `PtgArea` into A1
    // text. This means AutoFilter ranges stored as `_xlnm._FilterDatabase` can be lost when BIFF
    // workbook-stream parsing is unavailable and we fall back to calamine for defined names.
    //
    // Best-effort: if calamine surfaced (and we skipped) any invalid defined names, attempt to
    // recover the AutoFilter range directly from the workbook stream via our BIFF parser.
    if autofilters.is_empty() && skipped_count > 0 {
        let workbook_stream_fallback = if workbook_stream.is_none() {
            match biff::read_workbook_stream_from_xls(path) {
                Ok(bytes) => Some(bytes),
                Err(err) => {
                    warnings.push(ImportWarning::new(format!(
                        "failed to recover `.xls` AutoFilter ranges from workbook stream: {err}"
                    )));
                    None
                }
            }
        } else {
            None
        };
        let workbook_stream_bytes = workbook_stream
            .as_deref()
            .or(workbook_stream_fallback.as_deref());

        if let Some(workbook_stream_bytes) = workbook_stream_bytes {
            let biff_version = biff::detect_biff_version(workbook_stream_bytes);
            let codepage = biff::parse_biff_codepage(workbook_stream_bytes);

            let sheet_names_by_biff_idx =
                biff::parse_biff_bound_sheets(workbook_stream_bytes, biff_version, codepage)
                    .ok()
                    .unwrap_or_default()
                    .into_iter()
                    .map(|s| s.name)
                    .collect::<Vec<_>>();

            let sheet_names_by_biff_idx = if sheet_names_by_biff_idx.is_empty() {
                sheets.iter().map(|s| s.name.clone()).collect::<Vec<_>>()
            } else {
                sheet_names_by_biff_idx
            };

            match biff::parse_biff_defined_names(
                workbook_stream_bytes,
                biff_version,
                codepage,
                &sheet_names_by_biff_idx,
            ) {
                Ok(mut parsed) => {
                    for name in parsed.names.drain(..) {
                        if name.name != XLNM_FILTER_DATABASE {
                            continue;
                        }
                        let Some(biff_sheet_idx) = name.scope_sheet else {
                            continue;
                        };
                        let Some(&sheet_id) = sheet_ids_by_calamine_idx.get(biff_sheet_idx) else {
                            warnings.push(ImportWarning::new(format!(
                                "skipping `.xls` AutoFilter defined name `{}`: out-of-range sheet index {} (sheet count={})",
                                name.name,
                                biff_sheet_idx.saturating_add(1),
                                sheet_ids_by_calamine_idx.len()
                            )));
                            continue;
                        };

                        let mut a1 = name.refers_to.trim();
                        if let Some(rest) = a1.strip_prefix('=') {
                            a1 = rest.trim();
                        }
                        if let Some(rest) = a1.strip_prefix('@') {
                            a1 = rest.trim();
                        }
                        if let Some((_, rhs)) = a1.rsplit_once('!') {
                            a1 = rhs;
                        }

                        match Range::from_a1(a1) {
                            Ok(range) => autofilters.push((sheet_id, range)),
                            Err(err) => warnings.push(ImportWarning::new(format!(
                                "failed to parse `.xls` AutoFilter range `{}` from defined name `{}`: {err}",
                                name.refers_to, name.name
                            ))),
                        }
                    }
                }
                Err(err) => warnings.push(ImportWarning::new(format!(
                    "failed to recover `.xls` AutoFilter ranges from workbook stream: {err}"
                ))),
            }
        }
    }

    for (sheet_id, range) in autofilters {
        let Some(sheet) = out.sheet_mut(sheet_id) else {
            warnings.push(ImportWarning::new(format!(
                "skipping `.xls` AutoFilter range for missing sheet id {sheet_id}"
            )));
            continue;
        };
        if sheet.auto_filter.is_some() {
            continue;
        }

        sheet.auto_filter = Some(SheetAutoFilter {
            range,
            filter_columns: Vec::new(),
            sort_state: None,
            raw_xml: Vec::new(),
        });
    }
    Ok(XlsImportResult {
        workbook: out,
        source: ImportSource {
            path: path.to_path_buf(),
            format: SourceFormat::Xls,
        },
        merged_ranges,
        warnings,
    })
}

fn to_cell_ref(start: (u32, u32), row: usize, col: usize) -> Option<CellRef> {
    // NOTE: calamine `Range` iterators return coordinates relative to `range.start()`
    // rather than absolute worksheet coordinates.
    let row: u32 = row.try_into().ok()?;
    let col: u32 = col.try_into().ok()?;

    let row = start.0.checked_add(row)?;
    let col = start.1.checked_add(col)?;

    if row >= EXCEL_MAX_ROWS || col >= EXCEL_MAX_COLS {
        return None;
    }

    Some(CellRef::new(row, col))
}

fn sheet_visible_to_visibility(visible: SheetVisible) -> SheetVisibility {
    match visible {
        SheetVisible::Visible => SheetVisibility::Visible,
        SheetVisible::Hidden => SheetVisibility::Hidden,
        SheetVisible::VeryHidden => SheetVisibility::VeryHidden,
    }
}

fn convert_value(
    value: &Data,
    date_time_styles: Option<DateTimeStyleIds>,
) -> Option<(CellValue, Option<u32>)> {
    match value {
        Data::Empty => None,
        Data::Bool(v) => Some((CellValue::Boolean(*v), None)),
        Data::Int(v) => Some((CellValue::Number(*v as f64), None)),
        Data::Float(v) => Some((CellValue::Number(*v), None)),
        Data::String(v) => Some((CellValue::String(v.clone()), None)),
        Data::Error(e) => Some((CellValue::Error(cell_error_to_error_value(e.clone())), None)),
        Data::DateTime(v) => Some((
            CellValue::Number(v.as_f64()),
            date_time_styles.map(|styles| styles.style_for_excel_datetime(v)),
        )),
        Data::DateTimeIso(v) => Some((CellValue::String(v.clone()), None)),
        Data::DurationIso(v) => Some((CellValue::String(v.clone()), None)),
    }
}

fn style_id_for_cell_xf(
    xf_style_ids: Option<&[u32]>,
    sheet_cell_xfs: Option<&HashMap<CellRef, u16>>,
    cell_ref: CellRef,
) -> Option<u32> {
    let xf_index = sheet_cell_xfs?.get(&cell_ref).copied()? as usize;
    let style_id = *xf_style_ids?.get(xf_index)?;
    (style_id != 0).then_some(style_id)
}

fn cell_error_to_error_value(err: calamine::CellErrorType) -> ErrorValue {
    use calamine::CellErrorType;

    match err {
        CellErrorType::Div0 => ErrorValue::Div0,
        CellErrorType::NA => ErrorValue::NA,
        CellErrorType::Name => ErrorValue::Name,
        CellErrorType::Null => ErrorValue::Null,
        CellErrorType::Num => ErrorValue::Num,
        CellErrorType::Ref => ErrorValue::Ref,
        CellErrorType::Value => ErrorValue::Value,
        CellErrorType::GettingData => ErrorValue::GettingData,
    }
}

fn apply_row_col_properties(
    sheet: &mut formula_model::Worksheet,
    props: &biff::SheetRowColProperties,
) {
    for (&row, row_props) in &props.rows {
        if row >= EXCEL_MAX_ROWS {
            continue;
        }
        if row_props.height.is_some() {
            sheet.set_row_height(row, row_props.height);
        }
    }

    for (&col, col_props) in &props.cols {
        if col >= EXCEL_MAX_COLS {
            continue;
        }
        if col_props.width.is_some() {
            sheet.set_col_width(col, col_props.width);
        }
    }
}
fn apply_outline_properties(sheet: &mut formula_model::Worksheet, props: &biff::SheetRowColProperties) {
    sheet.outline.pr = props.outline_pr;

    for (&row, row_props) in &props.rows {
        if row >= EXCEL_MAX_ROWS {
            continue;
        }
        if row_props.outline_level == 0 && !row_props.collapsed {
            continue;
        }
        let row_1based = row.saturating_add(1);
        let entry = sheet.outline.rows.entry_mut(row_1based);
        entry.level = row_props.outline_level.min(7);
        entry.collapsed = row_props.collapsed;
    }

    for (&col, col_props) in &props.cols {
        if col >= EXCEL_MAX_COLS {
            continue;
        }
        if col_props.outline_level == 0 && !col_props.collapsed {
            continue;
        }
        let col_1based = col.saturating_add(1);
        let entry = sheet.outline.cols.entry_mut(col_1based);
        entry.level = col_props.outline_level.min(7);
        entry.collapsed = col_props.collapsed;
    }

    // Derive which rows/columns are hidden because they live inside a collapsed outline group.
    sheet.outline.recompute_outline_hidden_rows();
    sheet.outline.recompute_outline_hidden_cols();

    // BIFF uses the same hidden bit for user-hidden and outline-hidden rows/columns. Prefer the
    // derived outline state, and treat any remaining hidden flags as user-hidden.
    for (&row, row_props) in &props.rows {
        if !row_props.hidden || row >= EXCEL_MAX_ROWS {
            continue;
        }
        let row_1based = row.saturating_add(1);
        if sheet.outline.rows.entry(row_1based).hidden.outline {
            continue;
        }
        sheet.set_row_hidden(row, true);
    }

    for (&col, col_props) in &props.cols {
        if !col_props.hidden || col >= EXCEL_MAX_COLS {
            continue;
        }
        let col_1based = col.saturating_add(1);
        if sheet.outline.cols.entry(col_1based).hidden.outline {
            continue;
        }
        sheet.set_col_hidden(col, true);
    }

    // Ensure outline-hidden flags are up to date after any user-hidden state mutations.
    sheet.outline.recompute_outline_hidden_rows();
    sheet.outline.recompute_outline_hidden_cols();
}

fn truncate_to_utf16_len(value: &str, max_len: usize) -> String {
    if value.encode_utf16().count() <= max_len {
        return value.to_string();
    }

    let mut out = String::new();
    let mut used = 0usize;
    for ch in value.chars() {
        let ch_len = ch.len_utf16();
        if used.saturating_add(ch_len) > max_len {
            break;
        }
        out.push(ch);
        used = used.saturating_add(ch_len);
    }
    out
}

fn normalize_calamine_defined_name_name(name: &str) -> String {
    // Calamine can surface BIFF8 Unicode strings with embedded NUL bytes; strip them so the
    // imported name matches Excel’s visible name semantics.
    name.replace('\0', "")
}

fn sheet_name_eq_case_insensitive(a: &str, b: &str) -> bool {
    a.chars()
        .flat_map(|ch| ch.to_uppercase())
        .eq(b.chars().flat_map(|ch| ch.to_uppercase()))
}

fn infer_sheet_name_from_workbook_scoped_print_name(
    workbook: &Workbook,
    name: &str,
    refers_to: &str,
    warnings: &mut Vec<ImportWarning>,
) -> Option<String> {
    let refers_to = refers_to.trim();
    if refers_to.is_empty() {
        return None;
    }

    let areas = match split_print_name_areas(refers_to) {
        Ok(v) => v,
        Err(err) => {
            warnings.push(ImportWarning::new(format!(
                "failed to infer sheet scope for workbook-scoped `{name}`: {err}"
            )));
            return None;
        }
    };

    let mut inferred: Option<String> = None;
    for area in areas {
        let (sheet_name, _) = match split_print_name_sheet_ref(area) {
            Ok(v) => v,
            Err(err) => {
                warnings.push(ImportWarning::new(format!(
                    "failed to infer sheet scope for workbook-scoped `{name}` entry {area:?}: {err}"
                )));
                return None;
            }
        };

        let Some(sheet_name) = sheet_name else {
            // We only infer sheet scope from explicit `Sheet!` prefixes.
            return None;
        };
        let sheet_name = strip_workbook_prefix_from_sheet_ref(&sheet_name).to_string();

        match inferred.as_ref() {
            None => inferred = Some(sheet_name),
            Some(existing) => {
                if !sheet_name_eq_case_insensitive(existing, &sheet_name) {
                    warnings.push(ImportWarning::new(format!(
                        "skipping workbook-scoped `{name}`: references multiple sheets (`{existing}` and `{sheet_name}`)"
                    )));
                    return None;
                }
            }
        }
    }

    let inferred = inferred?;
    workbook.sheet_by_name(&inferred).map(|s| s.name.clone())
}

fn populate_print_settings_from_defined_names(workbook: &mut Workbook, warnings: &mut Vec<ImportWarning>) {
    // We need to snapshot the defined names up-front so we can mutably update print settings while
    // iterating.
    let builtins: Vec<(DefinedNameScope, String, String)> = workbook
        .defined_names
        .iter()
        .filter(|n| {
            n.name
                .eq_ignore_ascii_case(formula_model::XLNM_PRINT_AREA)
                || n.name
                    .eq_ignore_ascii_case(formula_model::XLNM_PRINT_TITLES)
        })
        .map(|n| (n.scope, n.name.clone(), n.refers_to.clone()))
        .collect();

    // Pass 1: sheet-scoped print names (canonical Excel encoding).
    // Pass 2: workbook-scoped print names (calamine fallback loses sheet scope).
    for pass in 0u8..=1u8 {
        for (scope, name, refers_to) in &builtins {
            let sheet_name = match (pass, scope) {
                (0, DefinedNameScope::Sheet(sheet_id)) => workbook.sheet(*sheet_id).map(|s| s.name.clone()),
                (1, DefinedNameScope::Workbook) => infer_sheet_name_from_workbook_scoped_print_name(
                    workbook,
                    name,
                    refers_to,
                    warnings,
                ),
                _ => None,
            };

            let Some(sheet_name) = sheet_name else {
                continue;
            };

            if name.eq_ignore_ascii_case(formula_model::XLNM_PRINT_AREA) {
                if workbook
                    .sheet_print_settings_by_name(&sheet_name)
                    .print_area
                    .is_some()
                {
                    continue;
                }
                if let Some(ranges) = parse_print_area_refers_to(&sheet_name, refers_to, warnings) {
                    workbook.set_sheet_print_area_by_name(&sheet_name, Some(ranges));
                }
            } else if name.eq_ignore_ascii_case(formula_model::XLNM_PRINT_TITLES) {
                if workbook
                    .sheet_print_settings_by_name(&sheet_name)
                    .print_titles
                    .is_some()
                {
                    continue;
                }
                if let Some(titles) = parse_print_titles_refers_to(&sheet_name, refers_to, warnings) {
                    workbook.set_sheet_print_titles_by_name(&sheet_name, Some(titles));
                }
            }
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ParsedA1Range {
    Cell(Range),
    Row(RowRange),
    Col(ColRange),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ParsedEndpoint {
    Cell(CellRef),
    Row(u32), // 0-based
    Col(u32), // 0-based
}

fn parse_print_area_refers_to(
    expected_sheet_name: &str,
    refers_to: &str,
    warnings: &mut Vec<ImportWarning>,
) -> Option<Vec<Range>> {
    let refers_to = refers_to.trim();
    if refers_to.is_empty() {
        return None;
    }

    let areas = match split_print_name_areas(refers_to) {
        Ok(areas) => areas,
        Err(err) => {
            warnings.push(ImportWarning::new(format!(
                "failed to parse `{}` for sheet `{expected_sheet_name}`: {err}",
                formula_model::XLNM_PRINT_AREA
            )));
            return None;
        }
    };

    let mut ranges: Vec<Range> = Vec::new();
    for area in areas {
        let (sheet_name, range_str) = match split_print_name_sheet_ref(area) {
            Ok(v) => v,
            Err(err) => {
                warnings.push(ImportWarning::new(format!(
                    "skipping `{}` entry {area:?} for sheet `{expected_sheet_name}`: {err}",
                    formula_model::XLNM_PRINT_AREA
                )));
                continue;
            }
        };

        if let Some(found_sheet_name) = sheet_name.as_deref() {
            let found_sheet_name = strip_workbook_prefix_from_sheet_ref(found_sheet_name);
            if !sheet_name_eq_case_insensitive(found_sheet_name, expected_sheet_name) {
                warnings.push(ImportWarning::new(format!(
                    "skipping `{}` entry {area:?} for sheet `{expected_sheet_name}`: references different sheet `{found_sheet_name}`",
                    formula_model::XLNM_PRINT_AREA
                )));
                continue;
            }
        }

        match parse_print_name_range(range_str) {
            Ok(ParsedA1Range::Cell(range)) => ranges.push(range),
            Ok(ParsedA1Range::Row(_) | ParsedA1Range::Col(_)) => {
                warnings.push(ImportWarning::new(format!(
                    "skipping `{}` entry {area:?} for sheet `{expected_sheet_name}`: print area must be a cell range",
                    formula_model::XLNM_PRINT_AREA
                )));
            }
            Err(err) => warnings.push(ImportWarning::new(format!(
                "skipping `{}` entry {area:?} for sheet `{expected_sheet_name}`: {err}",
                formula_model::XLNM_PRINT_AREA
            ))),
        }
    }

    (!ranges.is_empty()).then_some(ranges)
}

fn parse_print_titles_refers_to(
    expected_sheet_name: &str,
    refers_to: &str,
    warnings: &mut Vec<ImportWarning>,
) -> Option<PrintTitles> {
    let refers_to = refers_to.trim();
    if refers_to.is_empty() {
        return None;
    }

    let areas = match split_print_name_areas(refers_to) {
        Ok(areas) => areas,
        Err(err) => {
            warnings.push(ImportWarning::new(format!(
                "failed to parse `{}` for sheet `{expected_sheet_name}`: {err}",
                formula_model::XLNM_PRINT_TITLES
            )));
            return None;
        }
    };

    let mut titles = PrintTitles::default();
    for area in areas {
        let (sheet_name, range_str) = match split_print_name_sheet_ref(area) {
            Ok(v) => v,
            Err(err) => {
                warnings.push(ImportWarning::new(format!(
                    "skipping `{}` entry {area:?} for sheet `{expected_sheet_name}`: {err}",
                    formula_model::XLNM_PRINT_TITLES
                )));
                continue;
            }
        };

        if let Some(found_sheet_name) = sheet_name.as_deref() {
            let found_sheet_name = strip_workbook_prefix_from_sheet_ref(found_sheet_name);
            if !sheet_name_eq_case_insensitive(found_sheet_name, expected_sheet_name) {
                warnings.push(ImportWarning::new(format!(
                    "skipping `{}` entry {area:?} for sheet `{expected_sheet_name}`: references different sheet `{found_sheet_name}`",
                    formula_model::XLNM_PRINT_TITLES
                )));
                continue;
            }
        }

        match parse_print_name_range(range_str) {
            Ok(ParsedA1Range::Row(rows)) => titles.repeat_rows = Some(rows),
            Ok(ParsedA1Range::Col(cols)) => titles.repeat_cols = Some(cols),
            Ok(ParsedA1Range::Cell(range)) => {
                // Some decoders (e.g. calamine `.xls` defined-name fallback) represent whole-row/
                // whole-column print titles as explicit cell ranges (`$A$1:$IV$1`, `$A$1:$A$65536`)
                // rather than row/col-only references (`$1:$1`, `$A:$A`).
                if range.start.row == range.end.row && range.start.col != range.end.col {
                    titles.repeat_rows = Some(RowRange {
                        start: range.start.row,
                        end: range.end.row,
                    });
                } else if range.start.col == range.end.col && range.start.row != range.end.row {
                    titles.repeat_cols = Some(ColRange {
                        start: range.start.col,
                        end: range.end.col,
                    });
                } else {
                    warnings.push(ImportWarning::new(format!(
                        "skipping `{}` entry {area:?} for sheet `{expected_sheet_name}`: print titles must be a row or column range",
                        formula_model::XLNM_PRINT_TITLES
                    )));
                }
            }
            Err(err) => warnings.push(ImportWarning::new(format!(
                "skipping `{}` entry {area:?} for sheet `{expected_sheet_name}`: {err}",
                formula_model::XLNM_PRINT_TITLES
            ))),
        }
    }

    (titles.repeat_rows.is_some() || titles.repeat_cols.is_some()).then_some(titles)
}

fn split_print_name_areas(formula: &str) -> Result<Vec<&str>, String> {
    // Sheet names may be quoted (single quotes) and can contain commas. Split on commas only when
    // not inside quotes.
    let mut parts = Vec::new();
    let mut start = 0usize;
    let mut in_quotes = false;
    let bytes = formula.as_bytes();
    let mut i = 0usize;

    while i < bytes.len() {
        match bytes[i] {
            b'\'' => {
                if in_quotes {
                    if i + 1 < bytes.len() && bytes[i + 1] == b'\'' {
                        // Escaped quote in a sheet name.
                        i += 1;
                    } else {
                        in_quotes = false;
                    }
                } else {
                    in_quotes = true;
                }
            }
            b',' if !in_quotes => {
                let part = formula[start..i].trim();
                if !part.is_empty() {
                    parts.push(part);
                }
                start = i + 1;
            }
            _ => {}
        }

        i += 1;
    }

    let part = formula[start..].trim();
    if !part.is_empty() {
        parts.push(part);
    }

    Ok(parts)
}

/// Split an area reference like:
/// - `Sheet1!$A$1:$B$2`
/// - `'Sheet One'!$1:$1`
///
/// into `(sheet_name, range_str)`.
///
/// Returns `sheet_name=None` when the reference has no explicit `Sheet!` prefix.
fn split_print_name_sheet_ref(input: &str) -> Result<(Option<String>, &str), String> {
    let input = input.trim();
    if input.is_empty() {
        return Err("empty reference".to_string());
    }

    let bytes = input.as_bytes();
    if bytes.first() == Some(&b'\'') {
        let mut sheet = String::new();
        let mut i = 1usize;
        while i < bytes.len() {
            match bytes[i] {
                b'\'' => {
                    if i + 1 < bytes.len() && bytes[i + 1] == b'\'' {
                        sheet.push('\'');
                        i += 2;
                        continue;
                    }

                    // End of quoted sheet name.
                    if i + 1 >= bytes.len() || bytes[i + 1] != b'!' {
                        return Err(format!(
                            "expected ! after quoted sheet name in {input:?}"
                        ));
                    }

                    let rest = &input[(i + 2)..];
                    return Ok((Some(sheet), rest));
                }
                _ => sheet.push(bytes[i] as char),
            }
            i += 1;
        }

        return Err(format!("unterminated quoted sheet name in {input:?}"));
    }

    let Some(idx) = input.find('!') else {
        return Ok((None, input));
    };

    Ok((Some(input[..idx].to_string()), &input[(idx + 1)..]))
}

fn strip_workbook_prefix_from_sheet_ref(sheet_name: &str) -> &str {
    // Best-effort: Excel can serialize sheet references as `'[Book1.xlsx]Sheet1'!A1`.
    // If a workbook prefix exists, keep only the portion after the last `]`.
    sheet_name
        .rfind(']')
        .and_then(|idx| (idx + 1 <= sheet_name.len()).then_some(&sheet_name[(idx + 1)..]))
        .filter(|s| !s.is_empty())
        .unwrap_or(sheet_name)
}

fn parse_print_name_range(ref_str: &str) -> Result<ParsedA1Range, String> {
    let ref_str = ref_str.trim();
    if ref_str.is_empty() {
        return Err("empty range".to_string());
    }

    let (start, end) = match ref_str.split_once(':') {
        Some((a, b)) => (a, b),
        None => (ref_str, ref_str),
    };

    let start = parse_print_name_endpoint(start)?;
    let end = parse_print_name_endpoint(end)?;

    match (start, end) {
        (ParsedEndpoint::Cell(a), ParsedEndpoint::Cell(b)) => Ok(ParsedA1Range::Cell(Range::new(a, b))),
        (ParsedEndpoint::Row(a), ParsedEndpoint::Row(b)) => Ok(ParsedA1Range::Row(RowRange { start: a, end: b })),
        (ParsedEndpoint::Col(a), ParsedEndpoint::Col(b)) => Ok(ParsedA1Range::Col(ColRange { start: a, end: b })),
        _ => Err(format!("mismatched range endpoints in {ref_str:?}")),
    }
}

fn parse_print_name_endpoint(s: &str) -> Result<ParsedEndpoint, String> {
    let trimmed = s.trim().trim_matches('$');
    if trimmed.is_empty() {
        return Err("empty endpoint".to_string());
    }

    let mut letters = String::new();
    let mut digits = String::new();

    for ch in trimmed.chars() {
        if ch == '$' {
            continue;
        }
        if ch.is_ascii_alphabetic() && digits.is_empty() {
            letters.push(ch);
        } else if ch.is_ascii_digit() {
            digits.push(ch);
        } else {
            return Err(format!("invalid character {ch:?} in endpoint {s:?}"));
        }
    }

    match (letters.is_empty(), digits.is_empty()) {
        (false, false) => {
            let cell_ref = format!("{letters}{digits}");
            let cell = CellRef::from_a1(&cell_ref).map_err(|err| {
                format!("invalid cell reference in endpoint {s:?}: {err}")
            })?;
            Ok(ParsedEndpoint::Cell(cell))
        }
        (false, true) => {
            let col = parse_col_letters_to_index(&letters)?;
            Ok(ParsedEndpoint::Col(col))
        }
        (true, false) => {
            let row_1_based: u32 = digits
                .parse()
                .map_err(|_| format!("invalid row number in endpoint {s:?}"))?;
            if row_1_based == 0 {
                return Err(format!("invalid row number in endpoint {s:?}"));
            }
            Ok(ParsedEndpoint::Row(row_1_based - 1))
        }
        (true, true) => Err(format!("invalid endpoint {s:?}")),
    }
}

fn parse_col_letters_to_index(letters: &str) -> Result<u32, String> {
    let mut col: u32 = 0;
    for ch in letters.chars() {
        if !ch.is_ascii_alphabetic() {
            return Err(format!("invalid column letters {letters:?}"));
        }
        let v = (ch.to_ascii_uppercase() as u8 - b'A') as u32 + 1;
        col = col
            .checked_mul(26)
            .and_then(|c| c.checked_add(v))
            .ok_or_else(|| format!("invalid column letters {letters:?}"))?;
    }
    if col == 0 {
        return Err(format!("invalid column letters {letters:?}"));
    }
    Ok(col - 1)
}

fn sheet_name_taken(candidate: &str, existing_names: &[String]) -> bool {
    existing_names
        .iter()
        .any(|existing| sheet_name_eq_case_insensitive(existing, candidate))
}

/// Best-effort sanitization for legacy `.xls` sheet names.
///
/// Excel sheet names have a number of restrictions (see [`formula_model::validate_sheet_name`]).
/// Calamine may still surface corrupt/non-compliant names from malformed BIFF files; this helper
/// attempts to produce a deterministic, valid, unique name for the destination workbook.
///
/// This is part of the public API only so it can be tested from `crates/formula-xls/tests/`.
#[doc(hidden)]
pub fn sanitize_sheet_name(
    original: &str,
    sheet_number: usize,
    existing_names: &[String],
) -> String {
    let without_nuls = original.replace('\0', "");
    let trimmed = without_nuls.trim();

    let mut cleaned = String::with_capacity(trimmed.len());
    for ch in trimmed.chars() {
        if matches!(ch, ':' | '\\' | '/' | '?' | '*' | '[' | ']') {
            cleaned.push('_');
        } else {
            cleaned.push(ch);
        }
    }

    // `formula_model::validate_sheet_name` forbids leading/trailing apostrophes.
    let cleaned = cleaned.trim_matches('\'');

    let mut base = if cleaned.trim().is_empty() {
        format!("Sheet{sheet_number}")
    } else {
        truncate_to_utf16_len(cleaned, EXCEL_MAX_SHEET_NAME_LEN)
    };

    // Truncation can re-introduce a trailing apostrophe (`foo'bar...` → `foo'`).
    // Excel forbids sheet names that begin or end with `'`.
    base = truncate_to_utf16_len(base.trim_matches('\''), EXCEL_MAX_SHEET_NAME_LEN);
    if base.trim().is_empty() {
        base = truncate_to_utf16_len(&format!("Sheet{sheet_number}"), EXCEL_MAX_SHEET_NAME_LEN);
    }

    if !sheet_name_taken(&base, existing_names) {
        return base;
    }

    for suffix_index in 2usize.. {
        let suffix = format!(" ({suffix_index})");
        let suffix_len = suffix.encode_utf16().count();
        let max_base_len = EXCEL_MAX_SHEET_NAME_LEN.saturating_sub(suffix_len);

        let mut candidate = truncate_to_utf16_len(&base, max_base_len);
        candidate.push_str(&suffix);
        let candidate = truncate_to_utf16_len(&candidate, EXCEL_MAX_SHEET_NAME_LEN);

        if !sheet_name_taken(&candidate, existing_names) {
            return candidate;
        }
    }

    unreachable!("suffix loop should always return a unique sheet name");
}

/// Parse worksheet merged regions from BIFF `MERGEDCELLS` records.
///
/// This helper is part of the public API only so it can be exercised from integration tests in
/// `crates/formula-xls/tests/`. Most callers should use [`import_xls_path`] instead.
#[doc(hidden)]
pub fn parse_biff_sheet_merged_cells(
    workbook_stream: &[u8],
    start: usize,
) -> Result<Vec<Range>, String> {
    biff::parse_biff_sheet_merged_cells(workbook_stream, start)
}

fn normalize_sheet_name_for_match(name: &str) -> String {
    name.replace('\0', "").trim().to_lowercase()
}

fn reconcile_biff_sheet_mapping(
    calamine_sheets: &[Sheet],
    biff_sheets: Option<&[biff::BoundSheetInfo]>,
) -> (Vec<Option<usize>>, Option<String>) {
    let Some(biff_sheets) = biff_sheets else {
        return (vec![None; calamine_sheets.len()], None);
    };
    if biff_sheets.is_empty() {
        return (vec![None; calamine_sheets.len()], None);
    }

    let calamine_count = calamine_sheets.len();
    let biff_count = biff_sheets.len();

    // Primary mapping: BIFF BoundSheet order (workbook order) should align with calamine.
    let mut index_mapping = vec![None; calamine_count];
    let mut index_used_biff = vec![false; biff_count];
    for idx in 0..calamine_count.min(biff_count) {
        index_mapping[idx] = Some(idx);
        index_used_biff[idx] = true;
    }

    // Secondary mapping: normalized, case-insensitive name match.
    let mut name_mapping = vec![None; calamine_count];
    let mut name_used_biff = vec![false; biff_count];

    let mut biff_by_name: HashMap<String, Vec<usize>> = HashMap::new();
    for (idx, sheet) in biff_sheets.iter().enumerate() {
        biff_by_name
            .entry(normalize_sheet_name_for_match(&sheet.name))
            .or_default()
            .push(idx);
    }

    for (cal_idx, sheet) in calamine_sheets.iter().enumerate() {
        let key = normalize_sheet_name_for_match(&sheet.name);
        let Some(candidates) = biff_by_name.get(&key) else {
            continue;
        };

        if let Some(&biff_idx) = candidates.iter().find(|&&idx| !name_used_biff[idx]) {
            name_mapping[cal_idx] = Some(biff_idx);
            name_used_biff[biff_idx] = true;
        }
    }

    let index_mapped = index_mapping.iter().filter(|m| m.is_some()).count();
    let name_mapped = name_mapping.iter().filter(|m| m.is_some()).count();
    let index_score = index_mapping
        .iter()
        .enumerate()
        .filter_map(|(cal_idx, mapped)| {
            let biff_idx = (*mapped)?;
            let cal_name = normalize_sheet_name_for_match(&calamine_sheets[cal_idx].name);
            let biff_name = normalize_sheet_name_for_match(&biff_sheets[biff_idx].name);
            (cal_name == biff_name).then_some(())
        })
        .count();
    let name_score = name_mapping
        .iter()
        .enumerate()
        .filter_map(|(cal_idx, mapped)| {
            let biff_idx = (*mapped)?;
            let cal_name = normalize_sheet_name_for_match(&calamine_sheets[cal_idx].name);
            let biff_name = normalize_sheet_name_for_match(&biff_sheets[biff_idx].name);
            (cal_name == biff_name).then_some(())
        })
        .count();

    let (mapping, used_biff, strategy) = match name_mapped.cmp(&index_mapped) {
        std::cmp::Ordering::Greater => (name_mapping, name_used_biff, "name"),
        std::cmp::Ordering::Less => (index_mapping, index_used_biff, "index"),
        std::cmp::Ordering::Equal => {
            if name_score > index_score {
                (name_mapping, name_used_biff, "name")
            } else {
                (index_mapping, index_used_biff, "index")
            }
        }
    };

    let mapped_pairs: Vec<(usize, usize)> = mapping
        .iter()
        .enumerate()
        .filter_map(|(cal_idx, &biff_idx)| biff_idx.map(|biff_idx| (cal_idx, biff_idx)))
        .collect();
    let unmapped_calamine: Vec<usize> = mapping
        .iter()
        .enumerate()
        .filter_map(|(idx, mapped)| mapped.is_none().then_some(idx))
        .collect();
    let unmapped_biff: Vec<usize> = used_biff
        .iter()
        .enumerate()
        .filter_map(|(idx, used)| (!*used).then_some(idx))
        .collect();

    let calamine_names: Vec<&str> = calamine_sheets.iter().map(|s| s.name.as_str()).collect();
    let biff_names: Vec<&str> = biff_sheets.iter().map(|s| s.name.as_str()).collect();

    let should_warn = calamine_count != biff_count
        || strategy != "index"
        || !unmapped_calamine.is_empty()
        || !unmapped_biff.is_empty();

    let warning = should_warn.then(|| {
        format!(
            "failed to reconcile `.xls` sheet metadata (strategy={strategy}): calamine sheets ({calamine_count}) {calamine_names:?}; BIFF BoundSheet records ({biff_count}) {biff_names:?}; mapped indices (calamine->BIFF) {mapped_pairs:?}; unmapped calamine indices {unmapped_calamine:?}; unmapped BIFF indices {unmapped_biff:?}"
        )
    });

    (mapping, warning)
}

fn sanitize_biff8_continued_name_records_for_calamine(stream: &[u8]) -> Option<Vec<u8>> {
    const RECORD_NAME: u16 = 0x0018;
    const RECORD_CONTINUE: u16 = 0x003C;

    // Calamine's NAME parser reads:
    // - cch (u8) at offset 3 in the NAME payload
    // - cce (u16) at offset 4 in the NAME payload
    //
    // It can panic when `cce` exceeds the first physical fragment length. To avoid that, patch
    // `cce` to 0 for any NAME record that is continued.
    //
    // This keeps the name string available (so `PtgName` tokens can still resolve) while making
    // calamine skip parsing the formula payload.
    let mut name_header_offsets: Vec<usize> = Vec::new();
    let mut offset: usize = 0;

    while offset + 4 <= stream.len() {
        let record_id = u16::from_le_bytes([stream[offset], stream[offset + 1]]);
        let len = u16::from_le_bytes([stream[offset + 2], stream[offset + 3]]) as usize;

        let data_start = match offset.checked_add(4) {
            Some(v) => v,
            None => break,
        };
        let next = match data_start.checked_add(len) {
            Some(v) => v,
            None => break,
        };
        if next > stream.len() {
            break;
        }

        if record_id == RECORD_NAME && next + 4 <= stream.len() {
            let next_id = u16::from_le_bytes([stream[next], stream[next + 1]]);
            if next_id == RECORD_CONTINUE {
                name_header_offsets.push(offset);
            }
        }

        offset = next;
    }

    if name_header_offsets.is_empty() {
        return None;
    }

    let mut out = stream.to_vec();
    for header_offset in name_header_offsets {
        let len = u16::from_le_bytes([out[header_offset + 2], out[header_offset + 3]]) as usize;
        let data_start = header_offset + 4;

        // Patch `cce` (u16) at payload offset 4.
        if len >= 6 && data_start + 6 <= out.len() {
            out[data_start + 4..data_start + 6].copy_from_slice(&0u16.to_le_bytes());
        }

        // Best-effort: clamp `cch` if the name string cannot fit in the first fragment.
        //
        // Calamine slices `&payload[14..]` then indexes `buf[1..=cch]`, which requires
        // `payload.len() > 14 + cch`.
        if len >= 4 && len > 14 {
            let cch = out[data_start + 3] as usize;
            let available = len - 14;
            if available <= cch {
                // Ensure `available > cch` (or set to 0 if we can't even fit the flags byte).
                let new_cch = available.saturating_sub(1).min(u8::MAX as usize) as u8;
                out[data_start + 3] = new_cch;
            }
        }
    }

    Some(out)
}

fn build_in_memory_xls(workbook_stream: &[u8]) -> Result<Vec<u8>, ImportError> {
    // Construct a minimal CFB container containing only the Workbook stream. This is sufficient
    // for calamine's `Xls` parser and avoids needing to modify the original file on disk.
    //
    // Calamine's internal CFB parser rejects v4 compound files whose root directory has no
    // mini-stream (root start sector = ENDOFCHAIN). To keep `.xls` import best-effort for large
    // workbooks, always build a v3 CFB container here.
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create_with_version(cfb::Version::V3, cursor)
        .map_err(|err| ImportError::Xls(calamine::XlsError::Io(err)))?;

    {
        let mut stream = ole
            .create_stream("Workbook")
            .map_err(|err| ImportError::Xls(calamine::XlsError::Io(err)))?;
        stream
            .write_all(workbook_stream)
            .map_err(|err| ImportError::Xls(calamine::XlsError::Io(err)))?;
    }

    Ok(ole.into_inner().into_inner())
}

fn parse_autofilter_range_from_defined_name(refers_to: &str) -> Result<Range, String> {
    let refers_to = refers_to.trim();
    let refers_to = refers_to.strip_prefix('=').unwrap_or(refers_to).trim();
    let refers_to = refers_to.strip_prefix('@').unwrap_or(refers_to).trim();
    if refers_to.is_empty() {
        return Err("empty `refers_to`".to_string());
    }

    // Reject union formulas like `Sheet1!$A$1:$A$2,Sheet1!$C$1:$C$2` (comma at the top level).
    // We track quote and parenthesis nesting so commas inside quoted sheet names or function
    // argument lists do not get misclassified.
    let mut in_single_quote = false;
    let mut in_double_quote = false;
    let mut paren_depth: u32 = 0;

    let bytes = refers_to.as_bytes();
    let mut idx = 0usize;
    while idx < bytes.len() {
        match bytes[idx] {
            b'"' => {
                in_double_quote = !in_double_quote;
                idx += 1;
            }
            b'\'' if !in_double_quote => {
                if in_single_quote {
                    // Inside sheet-name quoting, `''` escapes a single quote.
                    if bytes.get(idx + 1) == Some(&b'\'') {
                        idx += 2;
                    } else {
                        in_single_quote = false;
                        idx += 1;
                    }
                } else {
                    in_single_quote = true;
                    idx += 1;
                }
            }
            b'(' if !in_single_quote && !in_double_quote => {
                paren_depth = paren_depth.saturating_add(1);
                idx += 1;
            }
            b')' if !in_single_quote && !in_double_quote => {
                paren_depth = paren_depth.saturating_sub(1);
                idx += 1;
            }
            b',' if !in_single_quote && !in_double_quote && paren_depth == 0 => {
                return Err("unsupported union formula (multiple areas)".to_string());
            }
            _ => idx += 1,
        }
    }

    // Strip any sheet qualifier, leaving the trailing A1 range part.
    // This handles `Sheet1!$A$1:$C$10` and `'My Sheet'!$A$1:$C$10`.
    let a1 = refers_to
        .rsplit_once('!')
        .map(|(_, tail)| tail)
        .unwrap_or(refers_to)
        .trim();

    Range::from_a1(a1).map_err(|err| format!("invalid A1 range `{a1}`: {err}"))
}
