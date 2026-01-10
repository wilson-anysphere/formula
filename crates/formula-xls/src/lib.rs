//! Legacy Excel 97-2003 `.xls` (BIFF) import support.
//!
//! This importer is intentionally best-effort: BIFF contains many features that
//! aren't (yet) representable in [`formula_model`]. We load sheets, cell values,
//! and formulas (as text) and preserve anything else as metadata/warnings.

use std::path::{Path, PathBuf};

use calamine::{open_workbook, Data, Reader, Xls};
use formula_model::{
    CellRef, CellValue, ErrorValue, Range, Workbook, EXCEL_MAX_COLS, EXCEL_MAX_ROWS,
};
use thiserror::Error;

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

    let sheet_names = workbook.sheet_names().to_owned();

    let mut out = Workbook::new();
    let mut warnings = Vec::new();
    let mut merged_ranges = Vec::new();

    for sheet_name in sheet_names {
        let sheet_id = out.add_sheet(sheet_name.clone());
        let sheet = out
            .sheet_mut(sheet_id)
            .expect("sheet id should exist immediately after add");

        let range = workbook.worksheet_range(&sheet_name)?;
        let range_start = range.start().unwrap_or((0, 0));

        for (row, col, value) in range.used_cells() {
            let Some(cell_ref) = to_cell_ref(range_start, row, col) else {
                warnings.push(ImportWarning::new(format!(
                    "skipping out-of-bounds cell in sheet `{sheet_name}` at ({row},{col})"
                )));
                continue;
            };

            let Some(value) = convert_value(value) else {
                continue;
            };

            sheet.set_value(cell_ref, value);
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

fn convert_value(value: &Data) -> Option<CellValue> {
    match value {
        Data::Empty => None,
        Data::Bool(v) => Some(CellValue::Boolean(*v)),
        Data::Int(v) => Some(CellValue::Number(*v as f64)),
        Data::Float(v) => Some(CellValue::Number(*v)),
        Data::String(v) => Some(CellValue::String(v.clone())),
        Data::Error(e) => Some(CellValue::Error(cell_error_to_error_value(e.clone()))),
        Data::DateTime(v) => Some(CellValue::Number(v.as_f64())),
        Data::DateTimeIso(v) => Some(CellValue::String(v.clone())),
        Data::DurationIso(v) => Some(CellValue::String(v.clone())),
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
