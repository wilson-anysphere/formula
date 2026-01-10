//! Legacy Excel 97-2003 `.xls` (BIFF) import support.
//!
//! This importer is intentionally best-effort: BIFF contains many features that
//! aren't (yet) representable in [`formula_model`]. We load sheets, cell values,
//! and formulas (as text) and preserve anything else as metadata/warnings.

use std::path::{Path, PathBuf};

use calamine::{open_workbook, Data, Reader, Sheet, SheetType, SheetVisible, Xls};
use formula_model::{
    CellRef, CellValue, ErrorValue, Range, SheetVisibility, Style, Workbook, EXCEL_MAX_COLS,
    EXCEL_MAX_ROWS,
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
/// `formula-model` does not yet have a dedicated representation for merged
/// cells, so we preserve them as metadata.
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

    for sheet_meta in sheets {
        let sheet_name = sheet_meta.name.clone();
        let sheet_id = out.add_sheet(sheet_name.clone());
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

                    sheet.set_value(cell_ref, value);
                    if let Some(style_id) = style_id {
                        sheet.set_style_id(cell_ref, style_id);
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

                    let normalized = normalize_formula(formula);
                    if normalized.is_empty() {
                        continue;
                    }

                    sheet.set_formula(cell_ref, Some(normalized));
                }
            }
            Err(err) => warnings.push(ImportWarning::new(format!(
                "failed to read formulas for sheet `{sheet_name}`: {err}"
            ))),
        }

        if let Some(merge_cells) = workbook.worksheet_merge_cells(&sheet_name) {
            for dim in merge_cells {
                let range = Range::new(
                    CellRef::new(dim.start.0, dim.start.1),
                    CellRef::new(dim.end.0, dim.end.1),
                );
                merged_ranges.push(MergedRange {
                    sheet_name: sheet_name.clone(),
                    range,
                });
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

fn normalize_formula(formula: &str) -> String {
    let trimmed = formula.trim();
    if trimmed.is_empty() {
        return String::new();
    }

    if trimmed.starts_with('=') {
        trimmed.to_owned()
    } else {
        format!("={trimmed}")
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
