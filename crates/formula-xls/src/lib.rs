//! Legacy Excel 97-2003 `.xls` (BIFF) import support.
//!
//! This importer is intentionally best-effort: BIFF contains many features that
//! aren't representable in [`formula_model`]. We load sheets, cell values,
//! formulas (as text), merged-cell regions, and basic row/column size/visibility
//! metadata where available. Anything else is preserved as metadata/warnings.

use std::collections::{BTreeMap, HashMap};
use std::io::Read;
use std::path::{Path, PathBuf};

use calamine::{open_workbook, Data, Reader, Sheet, SheetType, SheetVisible, Xls};
use formula_model::{
    normalize_formula_text, CellRef, CellValue, ColProperties, ErrorValue, Range, RowProperties,
    SheetVisibility, Style, Workbook, EXCEL_MAX_COLS, EXCEL_MAX_ROWS,
};
use thiserror::Error;

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
        // string. Use a best-effort heuristic to pick a reasonable default.
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
    let date_time_styles = DateTimeStyleIds::new(&mut out);
    let mut warnings = Vec::new();
    let mut merged_ranges = Vec::new();
    let row_col_props = match read_row_col_properties_from_xls(path) {
        Ok(props) => props,
        Err(err) => {
            warnings.push(ImportWarning::new(format!(
                "failed to import `.xls` row/column properties: {err}"
            )));
            HashMap::new()
        }
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

                    let Some((value, style_id)) = convert_value(value, date_time_styles) else {
                        continue;
                    };

                    let anchor = sheet.merged_regions.resolve_cell(cell_ref);
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

                    let normalized = normalize_formula_text(formula);
                    if normalized.is_empty() {
                        continue;
                    }

                    let anchor = sheet.merged_regions.resolve_cell(cell_ref);
                    sheet.set_formula(anchor, Some(normalized));
                }
            }
            Err(err) => warnings.push(ImportWarning::new(format!(
                "failed to read formulas for sheet `{sheet_name}`: {err}"
            ))),
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

#[derive(Debug, Default)]
struct SheetRowColProperties {
    rows: BTreeMap<u32, RowProperties>,
    cols: BTreeMap<u32, ColProperties>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum BiffVersion {
    Biff5,
    Biff8,
}

fn read_row_col_properties_from_xls(
    path: &Path,
) -> Result<HashMap<String, SheetRowColProperties>, String> {
    let mut comp = cfb::open(path).map_err(|err| err.to_string())?;
    let mut stream = open_xls_workbook_stream(&mut comp)?;

    let mut workbook_stream = Vec::new();
    stream
        .read_to_end(&mut workbook_stream)
        .map_err(|err| err.to_string())?;

    let biff = detect_biff_version(&workbook_stream);
    let sheets = parse_biff_bound_sheets(&workbook_stream, biff)?;

    let mut out = HashMap::new();
    for (sheet_name, offset) in sheets {
        if offset >= workbook_stream.len() {
            return Err(format!(
                "sheet `{sheet_name}` has out-of-bounds stream offset {offset}"
            ));
        }
        let props = parse_biff_sheet_row_col_properties(&workbook_stream, offset)?;
        out.insert(sheet_name, props);
    }

    Ok(out)
}

fn open_xls_workbook_stream<R: Read + std::io::Seek>(
    comp: &mut cfb::CompoundFile<R>,
) -> Result<cfb::Stream<R>, String> {
    for candidate in ["/Workbook", "/Book", "Workbook", "Book"] {
        if let Ok(stream) = comp.open_stream(candidate) {
            return Ok(stream);
        }
    }
    Err("missing workbook stream (expected `Workbook` or `Book`)".to_string())
}

fn detect_biff_version(workbook_stream: &[u8]) -> BiffVersion {
    let Some((record_id, data)) = read_biff_record(workbook_stream, 0) else {
        return BiffVersion::Biff8;
    };

    // BOF record type. Use BIFF8 heuristics compatible with calamine.
    if record_id != 0x0809 && record_id != 0x0009 {
        return BiffVersion::Biff8;
    }

    let Some(biff_version) = data.get(0..2).map(|v| u16::from_le_bytes([v[0], v[1]])) else {
        return BiffVersion::Biff8;
    };

    let dt = data
        .get(2..4)
        .map(|v| u16::from_le_bytes([v[0], v[1]]))
        .unwrap_or(0);

    match biff_version {
        0x0500 => BiffVersion::Biff5,
        0x0600 => BiffVersion::Biff8,
        0 => {
            if dt == 0x1000 {
                BiffVersion::Biff5
            } else {
                BiffVersion::Biff8
            }
        }
        _ => BiffVersion::Biff8,
    }
}

fn parse_biff_bound_sheets(
    workbook_stream: &[u8],
    biff: BiffVersion,
) -> Result<Vec<(String, usize)>, String> {
    let mut offset = 0usize;
    let mut out = Vec::new();

    loop {
        let Some((record_id, data)) = read_biff_record(workbook_stream, offset) else {
            break;
        };
        offset = offset
            .checked_add(4)
            .and_then(|o| o.checked_add(data.len()))
            .ok_or_else(|| "BIFF record offset overflow".to_string())?;

        match record_id {
            // BoundSheet8 [MS-XLS 2.4.28]
            0x0085 => {
                if data.len() < 7 {
                    return Err("BoundSheet8 record too short".to_string());
                }

                let sheet_offset = u32::from_le_bytes([data[0], data[1], data[2], data[3]]) as usize;
                let (name, _) = parse_biff_short_string(&data[6..], biff)?;
                let name = name.replace('\0', "");
                out.push((name, sheet_offset));
            }
            // EOF terminates the workbook global substream.
            0x000A => break,
            _ => {}
        }
    }

    Ok(out)
}

fn parse_biff_sheet_row_col_properties(
    workbook_stream: &[u8],
    start: usize,
) -> Result<SheetRowColProperties, String> {
    let mut props = SheetRowColProperties::default();
    let mut offset = start;

    loop {
        let Some((record_id, data)) = read_biff_record(workbook_stream, offset) else {
            break;
        };
        offset = offset
            .checked_add(4)
            .and_then(|o| o.checked_add(data.len()))
            .ok_or_else(|| "BIFF record offset overflow".to_string())?;

        match record_id {
            // ROW [MS-XLS 2.4.184]
            0x0208 => {
                if data.len() < 16 {
                    return Err("ROW record too short".to_string());
                }
                let row = u16::from_le_bytes([data[0], data[1]]) as u32;
                let height_options = u16::from_le_bytes([data[6], data[7]]);
                let height_twips = height_options & 0x7FFF;
                let default_height = (height_options & 0x8000) != 0;
                let options = u32::from_le_bytes([data[12], data[13], data[14], data[15]]);
                let hidden = (options & 0x0000_0020) != 0;

                let height = (!default_height && height_twips > 0)
                    .then_some(height_twips as f32 / 20.0);

                if hidden || height.is_some() {
                    let entry = props.rows.entry(row).or_default();
                    if let Some(height) = height {
                        entry.height = Some(height);
                    }
                    if hidden {
                        entry.hidden = true;
                    }
                }
            }
            // COLINFO [MS-XLS 2.4.48]
            0x007D => {
                if data.len() < 12 {
                    return Err("COLINFO record too short".to_string());
                }
                let first_col = u16::from_le_bytes([data[0], data[1]]) as u32;
                let last_col = u16::from_le_bytes([data[2], data[3]]) as u32;
                let width_raw = u16::from_le_bytes([data[4], data[5]]);
                let options = u16::from_le_bytes([data[8], data[9]]);
                let hidden = (options & 0x0001) != 0;

                let width = (width_raw > 0).then_some(width_raw as f32 / 256.0);

                if hidden || width.is_some() {
                    for col in first_col..=last_col {
                        let entry = props.cols.entry(col).or_default();
                        if let Some(width) = width {
                            entry.width = Some(width);
                        }
                        if hidden {
                            entry.hidden = true;
                        }
                    }
                }
            }
            // EOF terminates the sheet substream.
            0x000A => break,
            _ => {}
        }
    }

    Ok(props)
}

fn read_biff_record(workbook_stream: &[u8], offset: usize) -> Option<(u16, &[u8])> {
    let header = workbook_stream.get(offset..offset + 4)?;
    let record_id = u16::from_le_bytes([header[0], header[1]]);
    let len = u16::from_le_bytes([header[2], header[3]]) as usize;
    let data_start = offset + 4;
    let data_end = data_start.checked_add(len)?;
    let data = workbook_stream.get(data_start..data_end)?;
    Some((record_id, data))
}

fn parse_biff_short_string(input: &[u8], biff: BiffVersion) -> Result<(String, usize), String> {
    match biff {
        BiffVersion::Biff5 => parse_biff5_short_string(input),
        BiffVersion::Biff8 => parse_biff8_short_string(input),
    }
}

fn parse_biff5_short_string(input: &[u8]) -> Result<(String, usize), String> {
    let Some((&len, rest)) = input.split_first() else {
        return Err("unexpected end of string".to_string());
    };
    let len = len as usize;
    let bytes = rest
        .get(0..len)
        .ok_or_else(|| "unexpected end of string".to_string())?;
    Ok((String::from_utf8_lossy(bytes).into_owned(), 1 + len))
}

fn parse_biff8_short_string(input: &[u8]) -> Result<(String, usize), String> {
    if input.len() < 2 {
        return Err("unexpected end of string".to_string());
    }
    let cch = input[0] as usize;
    let flags = input[1];
    let mut offset = 2usize;

    let richtext_runs = if flags & 0x08 != 0 {
        if input.len() < offset + 2 {
            return Err("unexpected end of string".to_string());
        }
        let runs = u16::from_le_bytes([input[offset], input[offset + 1]]) as usize;
        offset += 2;
        runs
    } else {
        0
    };

    let ext_size = if flags & 0x04 != 0 {
        if input.len() < offset + 4 {
            return Err("unexpected end of string".to_string());
        }
        let size = u32::from_le_bytes([
            input[offset],
            input[offset + 1],
            input[offset + 2],
            input[offset + 3],
        ]) as usize;
        offset += 4;
        size
    } else {
        0
    };

    let is_unicode = (flags & 0x01) != 0;
    let char_bytes = if is_unicode { cch * 2 } else { cch };

    let chars = input
        .get(offset..offset + char_bytes)
        .ok_or_else(|| "unexpected end of string".to_string())?;
    offset += char_bytes;

    let name = if is_unicode {
        let mut u16s = Vec::with_capacity(cch);
        for chunk in chars.chunks_exact(2) {
            u16s.push(u16::from_le_bytes([chunk[0], chunk[1]]));
        }
        String::from_utf16_lossy(&u16s)
    } else {
        String::from_utf8_lossy(chars).into_owned()
    };

    let richtext_bytes = richtext_runs
        .checked_mul(4)
        .ok_or_else(|| "rich text run count overflow".to_string())?;
    if input.len() < offset + richtext_bytes + ext_size {
        return Err("unexpected end of string".to_string());
    }
    offset += richtext_bytes + ext_size;

    Ok((name, offset))
}

fn apply_row_col_properties(sheet: &mut formula_model::Worksheet, props: &SheetRowColProperties) {
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
