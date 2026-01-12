use std::path::{Path, PathBuf};

pub use formula_xls as xls;
pub use formula_xlsb as xlsb;
pub use formula_xlsx as xlsx;

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
    #[error("failed to save workbook as `.xlsb` package to `{path}`: {source}")]
    SaveXlsbPackage {
        path: PathBuf,
        #[source]
        source: xlsb::Error,
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

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum WorkbookFormat {
    Xlsx,
    Xls,
    Xlsb,
}

fn workbook_format(path: &Path) -> Result<WorkbookFormat, Error> {
    let ext = path
        .extension()
        .and_then(|s| s.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase();

    match ext.as_str() {
        "xlsx" | "xlsm" => Ok(WorkbookFormat::Xlsx),
        "xls" => Ok(WorkbookFormat::Xls),
        "xlsb" => Ok(WorkbookFormat::Xlsb),
        other => Err(Error::UnsupportedExtension {
            path: path.to_path_buf(),
            extension: other.to_string(),
        }),
    }
}

/// Open a spreadsheet workbook from disk directly into a [`formula_model::Workbook`].
///
/// This is a faster, lower-memory alternative to [`open_workbook`] for read-only/import
/// workflows:
/// - For `.xlsx`/`.xlsm`, this uses the streaming reader in `formula-xlsx` and avoids inflating
///   the entire OPC package into memory.
/// - For `.xls`, this returns the imported model workbook from `formula-xls`.
/// - For `.xlsb`, this converts the parsed workbook into a model workbook.
pub fn open_workbook_model(path: impl AsRef<Path>) -> Result<formula_model::Workbook, Error> {
    use std::fs::File;

    let path = path.as_ref();
    match workbook_format(path)? {
        WorkbookFormat::Xlsx => {
            let file = File::open(path).map_err(|source| Error::OpenIo {
                path: path.to_path_buf(),
                source,
            })?;
            xlsx::read_workbook_from_reader(file).map_err(|source| Error::OpenXlsx {
                path: path.to_path_buf(),
                source,
            })
        }
        WorkbookFormat::Xls => xls::import_xls_path(path)
            .map(|result| result.workbook)
            .map_err(|source| Error::OpenXls {
                path: path.to_path_buf(),
                source,
            }),
        WorkbookFormat::Xlsb => {
            let wb = xlsb::XlsbWorkbook::open(path).map_err(|source| Error::OpenXlsb {
                path: path.to_path_buf(),
                source,
            })?;
            xlsb_to_model_workbook(&wb).map_err(|source| Error::OpenXlsb {
                path: path.to_path_buf(),
                source,
            })
        }
    }
}

/// Open a spreadsheet workbook based on file extension.
///
/// Currently supports:
/// - `.xls` (via `formula-xls`)
/// - `.xlsb` (via `formula-xlsb`)
/// - `.xlsx` / `.xlsm` (via `formula-xlsx`)
pub fn open_workbook(path: impl AsRef<Path>) -> Result<Workbook, Error> {
    let path = path.as_ref();
    match workbook_format(path)? {
        WorkbookFormat::Xlsx => {
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
        WorkbookFormat::Xls => xls::import_xls_path(path)
            .map(Workbook::Xls)
            .map_err(|source| Error::OpenXls {
                path: path.to_path_buf(),
                source,
            }),
        WorkbookFormat::Xlsb => xlsb::XlsbWorkbook::open(path)
            .map(Workbook::Xlsb)
            .map_err(|source| Error::OpenXlsb {
                path: path.to_path_buf(),
                source,
            }),
    }
}

/// Save a workbook to disk.
///
/// Notes:
/// - [`Workbook::Xlsx`] is saved by writing the underlying OPC package back out,
///   preserving unknown parts.
/// - [`Workbook::Xls`] is exported as `.xlsx` (writing `.xls` is out of scope).
/// - [`Workbook::Xlsb`] can be saved losslessly back to `.xlsb` (package copy),
///   or exported to `.xlsx` depending on the output extension.
pub fn save_workbook(workbook: &Workbook, path: impl AsRef<Path>) -> Result<(), Error> {
    let path = path.as_ref();
    let ext = path
        .extension()
        .and_then(|s| s.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase();

    match workbook {
        Workbook::Xlsx(package) => match ext.as_str() {
            "xlsx" | "xlsm" => {
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
            other => Err(Error::UnsupportedExtension {
                path: path.to_path_buf(),
                extension: other.to_string(),
            }),
        },
        Workbook::Xls(result) => match ext.as_str() {
            "xlsx" => xlsx::write_workbook(&result.workbook, path).map_err(|source| {
                Error::SaveXlsxExport {
                    path: path.to_path_buf(),
                    source,
                }
            }),
            other => Err(Error::UnsupportedExtension {
                path: path.to_path_buf(),
                extension: other.to_string(),
            }),
        },
        Workbook::Xlsb(wb) => match ext.as_str() {
            "xlsb" => wb.save_as(path).map_err(|source| Error::SaveXlsbPackage {
                path: path.to_path_buf(),
                source,
            }),
            "xlsx" => {
                let model = xlsb_to_model_workbook(wb).map_err(|source| Error::SaveXlsbExport {
                    path: path.to_path_buf(),
                    source,
                })?;
                xlsx::write_workbook(&model, path).map_err(|source| Error::SaveXlsxExport {
                    path: path.to_path_buf(),
                    source,
                })
            }
            other => Err(Error::UnsupportedExtension {
                path: path.to_path_buf(),
                extension: other.to_string(),
            }),
        },
    }
}

fn xlsb_to_model_workbook(wb: &xlsb::XlsbWorkbook) -> Result<formula_model::Workbook, xlsb::Error> {
    use formula_model::{
        normalize_formula_text, CalculationMode, CellRef, CellValue, DateSystem, DefinedNameScope,
        ErrorValue, SheetVisibility, Style, Workbook as ModelWorkbook,
    };

    let mut out = ModelWorkbook::new();
    out.date_system = if wb.workbook_properties().date_system_1904 {
        DateSystem::Excel1904
    } else {
        DateSystem::Excel1900
    };
    if let Some(calc_mode) = wb.workbook_properties().calc_mode {
        out.calc_settings.calculation_mode = match calc_mode {
            xlsb::CalcMode::Auto => CalculationMode::Automatic,
            xlsb::CalcMode::Manual => CalculationMode::Manual,
            xlsb::CalcMode::AutoExceptTables => CalculationMode::AutomaticNoTable,
        };
    }
    if let Some(full_calc_on_load) = wb.workbook_properties().full_calc_on_load {
        out.calc_settings.full_calc_on_load = full_calc_on_load;
    }

    // Best-effort style mapping: XLSB cell records reference an XF index.
    //
    // We preserve number formats for now (fonts/fills/etc are not yet exposed by
    // `formula-xlsb::Styles`). When a built-in `numFmtId` is used, prefer a
    // `__builtin_numFmtId:<id>` placeholder for ids that would otherwise be
    // canonicalized to a *different* built-in id when exporting as XLSX.
    let mut xf_to_style_id: Vec<u32> = Vec::with_capacity(wb.styles().len());
    for xf_idx in 0..wb.styles().len() {
        let info = wb
            .styles()
            .get(xf_idx as u32)
            .expect("xf index within wb.styles().len()");
        if info.num_fmt_id == 0 {
            // The default "General" format doesn't need a distinct style id in
            // `formula-model` and would otherwise cause us to store many
            // formatting-only blank cells that we can't faithfully reproduce
            // (fonts/fills/etc are not yet exposed by `formula-xlsb::Styles`).
            xf_to_style_id.push(0);
            continue;
        }
        let number_format = match info.number_format.as_deref() {
            Some(fmt) if fmt.starts_with(formula_format::BUILTIN_NUM_FMT_ID_PLACEHOLDER_PREFIX) => {
                Some(fmt.to_string())
            }
            Some(fmt) => {
                if let Some(builtin) = formula_format::builtin_format_code(info.num_fmt_id) {
                    // Guard against (rare) custom formats that reuse a built-in id.
                    if fmt == builtin {
                        let canonical = formula_format::builtin_format_id(builtin);
                        if canonical == Some(info.num_fmt_id) {
                            Some(builtin.to_string())
                        } else {
                            Some(format!(
                                "{}{}",
                                formula_format::BUILTIN_NUM_FMT_ID_PLACEHOLDER_PREFIX,
                                info.num_fmt_id
                            ))
                        }
                    } else {
                        Some(fmt.to_string())
                    }
                } else {
                    Some(fmt.to_string())
                }
            }
            None => {
                // If we don't know the code but the id is in the reserved built-in range,
                // preserve it for round-trip.
                if info.num_fmt_id != 0 && info.num_fmt_id < 164 {
                    Some(format!(
                        "{}{}",
                        formula_format::BUILTIN_NUM_FMT_ID_PLACEHOLDER_PREFIX,
                        info.num_fmt_id
                    ))
                } else {
                    None
                }
            }
        };

        let style_id = number_format
            .as_ref()
            .map(|fmt| {
                out.intern_style(Style {
                    number_format: Some(fmt.clone()),
                    ..Default::default()
                })
            })
            .unwrap_or(0);
        xf_to_style_id.push(style_id);
    }

    let mut worksheet_ids_by_index: Vec<formula_model::WorksheetId> =
        Vec::with_capacity(wb.sheet_metas().len());

    for (sheet_index, meta) in wb.sheet_metas().iter().enumerate() {
        let sheet_id = out
            .add_sheet(meta.name.clone())
            .map_err(|err| xlsb::Error::InvalidSheetName(format!("{}: {err}", meta.name)))?;
        worksheet_ids_by_index.push(sheet_id);
        let sheet = out
            .sheet_mut(sheet_id)
            .expect("sheet id should exist immediately after add");
        sheet.visibility = match meta.visibility {
            xlsb::SheetVisibility::Visible => SheetVisibility::Visible,
            xlsb::SheetVisibility::Hidden => SheetVisibility::Hidden,
            xlsb::SheetVisibility::VeryHidden => SheetVisibility::VeryHidden,
        };

        wb.for_each_cell(sheet_index, |cell| {
            let cell_ref = CellRef::new(cell.row, cell.col);
            let style_id = xf_to_style_id
                .get(cell.style as usize)
                .copied()
                .unwrap_or(0);

            match cell.value {
                xlsb::CellValue::Blank => {}
                xlsb::CellValue::Number(v) => sheet.set_value(cell_ref, CellValue::Number(v)),
                xlsb::CellValue::Bool(v) => sheet.set_value(cell_ref, CellValue::Boolean(v)),
                xlsb::CellValue::Text(s) => sheet.set_value(cell_ref, CellValue::String(s)),
                xlsb::CellValue::Error(code) => sheet.set_value(
                    cell_ref,
                    CellValue::Error(ErrorValue::from_code(code).unwrap_or(ErrorValue::Unknown)),
                ),
            };

            // Cells with non-zero style ids must be stored, even if blank, matching
            // Excel's ability to format empty cells.
            if style_id != 0 {
                sheet.set_style_id(cell_ref, style_id);
            }

            if let Some(formula) = cell.formula.and_then(|f| f.text) {
                if let Some(normalized) = normalize_formula_text(&formula) {
                    sheet.set_formula(cell_ref, Some(normalized));
                }
            }
        })?;
    }

    // Defined names: parsed from `xl/workbook.bin` `BrtName` records.
    for name in wb.defined_names() {
        let Some(formula) = name.formula.as_ref().and_then(|f| f.text.as_deref()) else {
            continue;
        };
        let Some(refers_to) = normalize_formula_text(formula) else {
            continue;
        };

        let (scope, local_sheet_id) = match name.scope_sheet.and_then(|idx| {
            worksheet_ids_by_index
                .get(idx as usize)
                .copied()
                .map(|id| (idx, id))
        }) {
            Some((local_sheet_id, sheet_id)) => {
                (DefinedNameScope::Sheet(sheet_id), Some(local_sheet_id))
            }
            None => (DefinedNameScope::Workbook, None),
        };

        // Best-effort: ignore invalid/duplicate names so we can still export the workbook.
        let _ = out.create_defined_name(
            scope,
            name.name.clone(),
            refers_to,
            name.comment.clone(),
            name.hidden,
            local_sheet_id,
        );
    }

    Ok(out)
}

#[cfg(test)]
mod tests {
    use super::xlsb_to_model_workbook;
    use formula_model::{CellRef, DateSystem};
    use std::path::Path;

    #[test]
    fn xlsb_to_model_strips_leading_equals_from_formulas() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../formula-xlsb/tests/fixtures/simple.xlsb"
        ));

        let wb = crate::xlsb::XlsbWorkbook::open(fixture_path).expect("open xlsb fixture");
        let model = xlsb_to_model_workbook(&wb).expect("convert to model");
        let sheet = model.sheet_by_name("Sheet1").expect("Sheet1 missing");

        let cell = CellRef::from_a1("C1").expect("valid ref");
        let formula = sheet.formula(cell).expect("expected formula in C1");
        assert!(
            !formula.starts_with('='),
            "formula should be stored without leading '=' (got {formula:?})"
        );
        assert_eq!(formula, "B1*2");
    }

    #[test]
    fn xlsb_to_model_preserves_date_system() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../formula-xlsb/tests/fixtures/date1904.xlsb"
        ));

        let wb = crate::xlsb::XlsbWorkbook::open(fixture_path).expect("open xlsb fixture");
        let model = xlsb_to_model_workbook(&wb).expect("convert to model");
        assert_eq!(model.date_system, DateSystem::Excel1904);
    }

    #[test]
    fn xlsb_to_model_preserves_number_formats_from_styles() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../formula-xlsb/tests/fixtures_styles/date.xlsb"
        ));

        let wb = crate::xlsb::XlsbWorkbook::open(fixture_path).expect("open xlsb fixture");
        let model = xlsb_to_model_workbook(&wb).expect("convert to model");

        let sheet_name = &wb.sheet_metas()[0].name;
        let sheet = model.sheet_by_name(sheet_name).expect("sheet missing");

        let a1 = CellRef::from_a1("A1").expect("valid ref");
        let cell = sheet.cell(a1).expect("A1 missing");
        assert_ne!(cell.style_id, 0, "expected XLSB style to be preserved");

        let style = model
            .styles
            .get(cell.style_id)
            .expect("style id should exist");
        assert_eq!(style.number_format.as_deref(), Some("m/d/yyyy"));
    }
}
