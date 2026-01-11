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
    let mut sheet_offsets: Option<Vec<(String, usize)>> = None;
    let mut cell_xf_indices: Option<HashMap<String, HashMap<CellRef, u16>>> = None;

    if let Some(workbook_stream) = workbook_stream.as_deref() {
        let biff_version = biff::detect_biff_version(workbook_stream);

        match biff::parse_biff_workbook_globals(workbook_stream, biff_version) {
            Ok(globals) => {
                out.date_system = globals.date_system;

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

        match biff::parse_biff_bound_sheets(workbook_stream, biff_version) {
            Ok(offsets) => sheet_offsets = Some(offsets),
            Err(err) => warnings.push(ImportWarning::new(format!(
                "failed to import `.xls` sheet metadata: {err}"
            ))),
        }

        if let (Some(sheet_offsets), Some(mask)) =
            (sheet_offsets.as_ref(), xf_has_number_format.as_deref())
        {
            if mask.iter().any(|v| *v) {
                let mut out_map = HashMap::new();

                for (sheet_name, offset) in sheet_offsets {
                    if *offset >= workbook_stream.len() {
                        warnings.push(ImportWarning::new(format!(
                            "failed to import `.xls` cell styles for sheet `{sheet_name}`: out-of-bounds stream offset {offset}"
                        )));
                        continue;
                    }

                    match biff::parse_biff_sheet_cell_xf_indices_filtered(
                        workbook_stream,
                        *offset,
                        Some(mask),
                    ) {
                        Ok(xfs) => {
                            if !xfs.is_empty() {
                                out_map.insert(sheet_name.clone(), xfs);
                            }
                        }
                        Err(parse_err) => warnings.push(ImportWarning::new(format!(
                            "failed to import `.xls` cell styles for sheet `{sheet_name}`: {parse_err}"
                        ))),
                    }
                }

                if !out_map.is_empty() {
                    cell_xf_indices = Some(out_map);
                }
            }
        }
    }

    let date_time_styles = DateTimeStyleIds::new(&mut out);

    let row_col_props =
        if let (Some(workbook_stream), Some(sheet_offsets)) =
            (workbook_stream.as_deref(), sheet_offsets.as_ref())
        {
            let mut out_map = HashMap::new();
            for (sheet_name, offset) in sheet_offsets {
                if *offset >= workbook_stream.len() {
                    warnings.push(ImportWarning::new(format!(
                        "failed to import `.xls` row/column properties for sheet `{sheet_name}`: out-of-bounds stream offset {offset}"
                    )));
                    continue;
                }

                match biff::parse_biff_sheet_row_col_properties(workbook_stream, *offset) {
                    Ok(props) => {
                        if !props.rows.is_empty() || !props.cols.is_empty() {
                            out_map.insert(sheet_name.clone(), props);
                        }
                    }
                    Err(parse_err) => warnings.push(ImportWarning::new(format!(
                        "failed to import `.xls` row/column properties for sheet `{sheet_name}`: {parse_err}"
                    ))),
                }
            }
            out_map
        } else {
            HashMap::new()
        };

    for sheet_meta in sheets {
        let sheet_name = sheet_meta.name.clone();
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

        if let Some(props) = row_col_props.get(&sheet_name) {
            apply_row_col_properties(sheet, props);
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

        let sheet_cell_xfs = cell_xf_indices
            .as_ref()
            .and_then(|map| map.get(&sheet_name));

        match workbook.worksheet_range(&sheet_name) {
            Ok(range) => {
                let range_start = range.start().unwrap_or((0, 0));

                for (row, col, value) in range.used_cells() {
                    let Some(cell_ref) = to_cell_ref(range_start, row, col) else {
                        warnings.push(ImportWarning::new(format!(
                            "skipping out-of-bounds cell in sheet `{sheet_name}` at ({row},{col})"
                        )));
                        continue;
                    };

                    let anchor = sheet.merged_regions.resolve_cell(cell_ref);
                    let Some((value, mut style_id)) = convert_value(value, date_time_styles) else {
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
            Err(err) => warnings.push(ImportWarning::new(format!(
                "failed to read cell values for sheet `{sheet_name}`: {err}"
            ))),
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
            for (&cell_ref, &xf_idx) in sheet_cell_xfs {
                if cell_ref.row >= EXCEL_MAX_ROWS || cell_ref.col >= EXCEL_MAX_COLS {
                    continue;
                }
                let style_id = xf_style_ids.get(xf_idx as usize).copied().flatten();
                let Some(style_id) = style_id else {
                    continue;
                };
                let anchor = sheet.merged_regions.resolve_cell(cell_ref);
                sheet.set_style_id(anchor, style_id);
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
    date_time_styles: DateTimeStyleIds,
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
            Some(date_time_styles.style_for_excel_datetime(v)),
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

/// Read the BIFF per-cell XF indices for every worksheet substream in a legacy `.xls` workbook.
///
/// The returned map is keyed by sheet name (as stored in the workbook global stream) and then by
/// absolute 0-based `(row, col)` coordinates.
///
/// This is intentionally style-only: it does *not* attempt to parse cell values or formulas.
pub fn read_cell_xfs_from_xls(
    path: impl AsRef<Path>,
) -> Result<HashMap<String, HashMap<(u32, u32), u16>>, String> {
    let workbook_stream = biff::read_workbook_stream_from_xls(path.as_ref())?;
    let biff_version = biff::detect_biff_version(&workbook_stream);
    let sheets = biff::parse_biff_bound_sheets(&workbook_stream, biff_version)?;

    let mut out = HashMap::new();
    for (sheet_name, offset) in sheets {
        if offset >= workbook_stream.len() {
            return Err(format!(
                "sheet `{sheet_name}` has out-of-bounds stream offset {offset}"
            ));
        }

        let xfs = parse_biff_sheet_cell_xfs(&workbook_stream, offset)?;
        out.insert(sheet_name, xfs);
    }

    Ok(out)
}

/// Parse a BIFF worksheet substream starting at `start` and return a mapping from cell coordinates
/// to the XF index (`ixfe`) referenced by the last record encountered for that cell.
///
/// Excel treats later records as overwriting earlier ones; we mirror that behaviour by always
/// preferring the last record seen for a given `(row, col)` key.
pub fn parse_biff_sheet_cell_xfs(
    workbook_stream: &[u8],
    start: usize,
) -> Result<HashMap<(u32, u32), u16>, String> {
    let xfs = biff::parse_biff_sheet_cell_xf_indices_filtered(workbook_stream, start, None)?;
    Ok(xfs
        .into_iter()
        .map(|(cell_ref, xf)| ((cell_ref.row, cell_ref.col), xf))
        .collect())
}
