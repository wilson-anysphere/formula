//! Legacy Excel 97-2003 `.xls` (BIFF) import support.
//!
//! This importer is intentionally best-effort: BIFF contains many features that
//! aren't representable in [`formula_model`]. We load sheets, cell values,
//! formulas (as text), merged-cell regions, and basic row/column size/visibility
//! metadata where available. We also extract workbook-global styles needed to
//! preserve Excel number format codes and the workbook date system (1900 vs 1904)
//! when possible. Anything else is preserved as metadata/warnings.

use std::collections::{HashMap, HashSet};
use std::path::{Path, PathBuf};

use calamine::{open_workbook, Data, Reader, Sheet, SheetType, SheetVisible, Xls};
use formula_model::{
    normalize_formula_text, CellRef, CellValue, ErrorValue, HyperlinkTarget, Range, SheetVisibility,
    Style, TabColor, Workbook, EXCEL_MAX_COLS, EXCEL_MAX_ROWS, EXCEL_MAX_SHEET_NAME_LEN,
};
use thiserror::Error;

mod biff;

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
    #[error("invalid worksheet name: {0}")]
    InvalidSheetName(#[from] formula_model::SheetNameError),
}

/// Import a legacy `.xls` workbook from disk.
pub fn import_xls_path(path: impl AsRef<Path>) -> Result<XlsImportResult, ImportError> {
    let path = path.as_ref();
    let mut workbook: Xls<_> = open_workbook(path)?;

    // We need to snapshot metadata (names, visibility, type) up-front because we
    // need mutable access to the workbook while iterating over ranges.
    let sheets: Vec<Sheet> = workbook.sheets_metadata().to_vec();

    let mut out = Workbook::new();
    let mut used_sheet_names: Vec<String> = Vec::new();
    let mut warnings = Vec::new();
    let mut merged_ranges = Vec::new();
    let workbook_stream = match biff::read_workbook_stream_from_xls(path) {
        Ok(bytes) => Some(bytes),
        Err(err) => {
            warnings.push(ImportWarning::new(format!(
                "failed to read `.xls` workbook stream: {err}"
            )));
            None
        }
    };

    let mut xf_style_ids: Option<Vec<Option<u32>>> = None;
    let mut xf_has_number_format: Option<Vec<bool>> = None;
    let mut sheet_tab_colors: Option<Vec<Option<TabColor>>> = None;
    let mut biff_sheets: Option<Vec<biff::BoundSheetInfo>> = None;
    let mut row_col_props: Option<Vec<biff::SheetRowColProperties>> = None;
    let mut cell_xf_indices: Option<Vec<HashMap<CellRef, u16>>> = None;
    let mut cell_xf_parse_failed: Option<Vec<bool>> = None;
    let mut biff_codepage: Option<u16> = None;

    if let Some(workbook_stream) = workbook_stream.as_deref() {
        let biff_version = biff::detect_biff_version(workbook_stream);
        let codepage = biff::parse_biff_codepage(workbook_stream);
        biff_codepage = Some(codepage);

        match biff::parse_biff_workbook_globals(workbook_stream, biff_version, codepage) {
            Ok(mut globals) => {
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
                warnings.extend(globals.warnings.drain(..).map(ImportWarning::new));
                sheet_tab_colors = Some(std::mem::take(&mut globals.sheet_tab_colors));

                let mut cache: HashMap<String, u32> = HashMap::new();
                let mut style_ids = Vec::with_capacity(globals.xf_count());
                let mut has_number_format = Vec::with_capacity(globals.xf_count());
                for xf_index in 0..globals.xf_count() as u32 {
                    let style_id = match globals.resolve_number_format_code(xf_index) {
                        Some(code) => {
                            has_number_format.push(true);
                            if let Some(existing) = cache.get(&code) {
                                Some(*existing)
                            } else {
                                let style_id = out.intern_style(Style {
                                    number_format: Some(code.clone()),
                                    ..Default::default()
                                });
                                cache.insert(code, style_id);
                                Some(style_id)
                            }
                        }
                        None => {
                            has_number_format.push(false);
                            None
                        }
                    };
                    style_ids.push(style_id);
                }

                xf_style_ids = Some(style_ids);
                xf_has_number_format = Some(has_number_format);
            }
            Err(err) => warnings.push(ImportWarning::new(format!(
                "failed to import `.xls` workbook globals: {err}"
            ))),
        }

        match biff::parse_biff_bound_sheets(workbook_stream, biff_version, codepage) {
            Ok(sheets) => biff_sheets = Some(sheets),
            Err(err) => warnings.push(ImportWarning::new(format!(
                "failed to import `.xls` sheet metadata: {err}"
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

            if let Some(mask) = xf_has_number_format.as_deref() {
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

        if let Some(biff_idx) = biff_idx {
            if let Some(props) = row_col_props
                .as_ref()
                .and_then(|props_by_sheet| props_by_sheet.get(biff_idx))
            {
                apply_row_col_properties(sheet, props);
                apply_outline_properties(sheet, props);
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
                let Some((value, mut style_id)) = convert_value(value, sheet_date_time_styles) else {
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
        if let (Some(workbook_stream), Some(codepage), Some(biff_idx)) = (
            workbook_stream.as_deref(),
            biff_codepage,
            biff_idx,
        ) {
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
        if let (Some(xf_style_ids), Some(sheet_cell_xfs)) = (xf_style_ids.as_deref(), sheet_cell_xfs)
        {
            let mut out_of_range_xf_count: usize = 0;
            if sheet.merged_regions.is_empty() {
                for (&cell_ref, &xf_idx) in sheet_cell_xfs {
                    if cell_ref.row >= EXCEL_MAX_ROWS || cell_ref.col >= EXCEL_MAX_COLS {
                        continue;
                    }

                    let Some(style_id) = xf_style_ids.get(xf_idx as usize).copied() else {
                        out_of_range_xf_count = out_of_range_xf_count.saturating_add(1);
                        continue;
                    };
                    let Some(style_id) = style_id else {
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

                    let Some(style_id) = xf_style_ids.get(xf_idx as usize).copied() else {
                        out_of_range_xf_count = out_of_range_xf_count.saturating_add(1);
                        continue;
                    };
                    let Some(style_id) = style_id else {
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

    // If we had to sanitize sheet names, internal hyperlinks may still reference the original
    // (invalid) sheet names. Rewrite internal hyperlink targets to point at the final imported
    // sheet names so navigation remains correct after import and round-trips to XLSX.
    if !final_sheet_names_by_idx.is_empty() {
        let mut resolved_sheet_names: HashMap<String, String> = HashMap::new();
        for (idx, sheet_meta) in sheets.iter().enumerate() {
            let Some(final_name) = final_sheet_names_by_idx.get(idx) else {
                continue;
            };
            resolved_sheet_names.insert(
                normalize_sheet_name_for_match(&sheet_meta.name),
                final_name.clone(),
            );

            if let Some(biff_idx) = sheet_mapping.get(idx).copied().flatten() {
                if let Some(biff_name) = biff_sheets
                    .as_ref()
                    .and_then(|sheets| sheets.get(biff_idx))
                    .map(|s| s.name.as_str())
                {
                    resolved_sheet_names
                        .entry(normalize_sheet_name_for_match(biff_name))
                        .or_insert_with(|| final_name.clone());
                }
            }
        }

        if !resolved_sheet_names.is_empty() {
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
    xf_style_ids: Option<&[Option<u32>]>,
    sheet_cell_xfs: Option<&HashMap<CellRef, u16>>,
    cell_ref: CellRef,
) -> Option<u32> {
    let xf_index = sheet_cell_xfs?.get(&cell_ref).copied()? as usize;
    xf_style_ids?.get(xf_index).copied().flatten()
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

fn sheet_name_eq_case_insensitive(a: &str, b: &str) -> bool {
    a.chars()
        .flat_map(|ch| ch.to_uppercase())
        .eq(b.chars().flat_map(|ch| ch.to_uppercase()))
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
