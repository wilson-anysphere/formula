//! Legacy Excel 97-2003 `.xls` (BIFF) import support.
//!
//! This importer is intentionally best-effort: BIFF contains many features that
//! aren't representable in [`formula_model`]. We load sheets, cell values,
//! formulas (as text), merged-cell regions, and basic row/column size/visibility
//! metadata where available. We also extract workbook-global styles needed to
//! preserve Excel number format codes and the workbook date system (1900 vs 1904)
//! when possible. Anything else is preserved as metadata/warnings.

use std::collections::HashMap;
use std::path::{Path, PathBuf};

use calamine::{open_workbook, Data, Reader, Sheet, SheetType, SheetVisible, Xls};
use formula_model::{
    normalize_formula_text, CellRef, CellValue, ErrorValue, Range, SheetVisibility, Style,
    Workbook, EXCEL_MAX_COLS, EXCEL_MAX_ROWS,
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
    let mut biff_sheets: Option<Vec<biff::BoundSheetInfo>> = None;
    let mut row_col_props: Option<Vec<biff::SheetRowColProperties>> = None;
    let mut cell_xf_indices: Option<Vec<HashMap<CellRef, u16>>> = None;
    let mut cell_xf_parse_failed: Option<Vec<bool>> = None;

    if let Some(workbook_stream) = workbook_stream.as_deref() {
        let biff_version = biff::detect_biff_version(workbook_stream);
        let codepage = biff::parse_biff_codepage(workbook_stream);

        match biff::parse_biff_workbook_globals(workbook_stream, biff_version, codepage) {
            Ok(mut globals) => {
                out.date_system = globals.date_system;
                warnings.extend(globals.warnings.drain(..).map(ImportWarning::new));

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
                    Ok(props) => props_by_sheet.push(props),
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

    for (sheet_idx, sheet_meta) in sheets.iter().enumerate() {
        let sheet_name = sheet_meta.name.clone();
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

        let value_range = match workbook.worksheet_range(&sheet_name) {
            Ok(range) => Some(range),
            Err(err) => {
                warnings.push(ImportWarning::new(format!(
                    "failed to read cell values for sheet `{sheet_name}`: {err}"
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

        let sheet_id = out.add_sheet(sheet_name.clone())?;
        let sheet = out
            .sheet_mut(sheet_id)
            .expect("sheet id should exist immediately after add");

        sheet.visibility = sheet_visible_to_visibility(sheet_meta.visible);

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
            }
        }

        if let Some(merge_cells) = workbook.worksheet_merge_cells(&sheet_name) {
            for dim in merge_cells {
                let range = Range::new(
                    CellRef::new(dim.start.0, dim.start.1),
                    CellRef::new(dim.end.0, dim.end.1),
                );

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

        match workbook.worksheet_formula(&sheet_name) {
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
        if row_props.height.is_some() {
            sheet.set_row_height(row, row_props.height);
        }
        if row_props.hidden {
            sheet.set_row_hidden(row, true);
        }
    }

    for (&col, col_props) in &props.cols {
        if col_props.width.is_some() {
            sheet.set_col_width(col, col_props.width);
        }
        if col_props.hidden {
            sheet.set_col_hidden(col, true);
        }
    }
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
