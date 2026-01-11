use std::path::{Path, PathBuf};

pub use formula_xlsx as xlsx;
pub use formula_xls as xls;
pub use formula_xlsb as xlsb;

#[derive(Debug, thiserror::Error)]
pub enum Error {
    #[error("unsupported extension `{extension}` for workbook `{path}`")]
    UnsupportedExtension { path: PathBuf, extension: String },
    #[error("failed to open workbook `{path}`: {source}")]
    OpenIo {
        path: PathBuf,
        #[source]
        source: std::io::Error,
    },
    #[error("failed to open `.xlsx` workbook `{path}`: {source}")]
    OpenXlsx {
        path: PathBuf,
        #[source]
        source: xlsx::XlsxError,
    },
    #[error("failed to open `.xls` workbook `{path}`: {source}")]
    OpenXls {
        path: PathBuf,
        #[source]
        source: xls::ImportError,
    },
    #[error("failed to open `.xlsb` workbook `{path}`: {source}")]
    OpenXlsb {
        path: PathBuf,
        #[source]
        source: xlsb::Error,
    },
    #[error("failed to save workbook `{path}`: {source}")]
    SaveIo {
        path: PathBuf,
        #[source]
        source: std::io::Error,
    },
    #[error("failed to save workbook package to `{path}`: {source}")]
    SaveXlsxPackage {
        path: PathBuf,
        #[source]
        source: xlsx::XlsxError,
    },
    #[error("failed to export workbook as `.xlsx` to `{path}`: {source}")]
    SaveXlsxExport {
        path: PathBuf,
        #[source]
        source: xlsx::XlsxWriteError,
    },
    #[error("failed to export `.xlsb` workbook as `.xlsx` to `{path}`: {source}")]
    SaveXlsbExport {
        path: PathBuf,
        #[source]
        source: xlsb::Error,
    },
}

/// A workbook opened from disk.
#[derive(Debug)]
pub enum Workbook {
    /// XLSX/XLSM opened as an Open Packaging Convention (OPC) package.
    ///
    /// This preserves unknown parts (e.g. `customXml/`, `xl/vbaProject.bin`) byte-for-byte.
    Xlsx(xlsx::XlsxPackage),
    Xls(xls::XlsImportResult),
    Xlsb(xlsb::XlsbWorkbook),
}

/// Open a spreadsheet workbook based on file extension.
///
/// Currently supports:
/// - `.xls` (via `formula-xls`)
/// - `.xlsb` (via `formula-xlsb`)
/// - `.xlsx` / `.xlsm` (via `formula-xlsx`)
pub fn open_workbook(path: impl AsRef<Path>) -> Result<Workbook, Error> {
    let path = path.as_ref();
    let ext = path
        .extension()
        .and_then(|s| s.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase();
    match ext.as_str() {
        "xlsx" | "xlsm" => {
            let bytes = std::fs::read(path).map_err(|source| Error::OpenIo {
                path: path.to_path_buf(),
                source,
            })?;
            let package =
                xlsx::XlsxPackage::from_bytes(&bytes).map_err(|source| Error::OpenXlsx {
                    path: path.to_path_buf(),
                    source,
                })?;
            Ok(Workbook::Xlsx(package))
        }
        "xls" => xls::import_xls_path(path)
            .map(Workbook::Xls)
            .map_err(|source| Error::OpenXls {
                path: path.to_path_buf(),
                source,
            }),
        "xlsb" => xlsb::XlsbWorkbook::open(path)
            .map(Workbook::Xlsb)
            .map_err(|source| Error::OpenXlsb {
                path: path.to_path_buf(),
                source,
            }),
        other => Err(Error::UnsupportedExtension {
            path: path.to_path_buf(),
            extension: other.to_string(),
        }),
    }
}

/// Save a workbook to disk.
///
/// Notes:
/// - [`Workbook::Xlsx`] is saved by writing the underlying OPC package back out,
///   preserving unknown parts.
/// - [`Workbook::Xls`] is exported as `.xlsx` (writing `.xls` is out of scope).
/// - [`Workbook::Xlsb`] is exported as a best-effort `.xlsx` (writing `.xlsb` is
///   out of scope).
pub fn save_workbook(workbook: &Workbook, path: impl AsRef<Path>) -> Result<(), Error> {
    let path = path.as_ref();
    match workbook {
        Workbook::Xlsx(package) => {
            let file = std::fs::File::create(path).map_err(|source| Error::SaveIo {
                path: path.to_path_buf(),
                source,
            })?;
            package
                .write_to(file)
                .map_err(|source| Error::SaveXlsxPackage {
                    path: path.to_path_buf(),
                    source,
                })?;
            Ok(())
        }
        Workbook::Xls(result) => xlsx::write_workbook(&result.workbook, path).map_err(|source| {
            Error::SaveXlsxExport {
                path: path.to_path_buf(),
                source,
            }
        }),
        Workbook::Xlsb(wb) => {
            let model = xlsb_to_model_workbook(wb).map_err(|source| Error::SaveXlsbExport {
                path: path.to_path_buf(),
                source,
            })?;
            xlsx::write_workbook(&model, path).map_err(|source| Error::SaveXlsxExport {
                path: path.to_path_buf(),
                source,
            })
        }
    }
}

fn xlsb_to_model_workbook(wb: &xlsb::XlsbWorkbook) -> Result<formula_model::Workbook, xlsb::Error> {
    use formula_model::{CellRef, CellValue, ErrorValue, Workbook as ModelWorkbook};

    fn normalize_formula(formula: &str) -> Option<String> {
        let trimmed = formula.trim();
        if trimmed.is_empty() {
            return None;
        }
        if trimmed.starts_with('=') {
            Some(trimmed.to_owned())
        } else {
            Some(format!("={trimmed}"))
        }
    }

    let mut out = ModelWorkbook::new();

    for (sheet_index, meta) in wb.sheet_metas().iter().enumerate() {
        let sheet_id = out.add_sheet(meta.name.clone());
        let sheet = out
            .sheet_mut(sheet_id)
            .expect("sheet id should exist immediately after add");

        wb.for_each_cell(sheet_index, |cell| {
            let cell_ref = CellRef::new(cell.row, cell.col);
            match cell.value {
                xlsb::CellValue::Blank => {}
                xlsb::CellValue::Number(v) => sheet.set_value(cell_ref, CellValue::Number(v)),
                xlsb::CellValue::Bool(v) => sheet.set_value(cell_ref, CellValue::Boolean(v)),
                xlsb::CellValue::Text(s) => sheet.set_value(cell_ref, CellValue::String(s)),
                xlsb::CellValue::Error(code) => sheet.set_value(
                    cell_ref,
                    CellValue::Error(ErrorValue::from_code(code).unwrap_or(ErrorValue::Unknown)),
                ),
            }

            if let Some(formula) = cell.formula.and_then(|f| f.text) {
                if let Some(normalized) = normalize_formula(&formula) {
                    sheet.set_formula(cell_ref, Some(normalized));
                }
            }
        })?;
    }

    Ok(out)
}
