use std::path::Path;

pub use formula_xls as xls;
pub use formula_xlsb as xlsb;

#[derive(Debug, thiserror::Error)]
pub enum Error {
    #[error("unsupported extension: {0}")]
    UnsupportedExtension(String),
    #[error(transparent)]
    Xls(#[from] xls::ImportError),
    #[error(transparent)]
    Xlsb(#[from] xlsb::Error),
}

/// A workbook opened from disk.
#[derive(Debug)]
pub enum Workbook {
    Xls(xls::XlsImportResult),
    Xlsb(xlsb::XlsbWorkbook),
}

/// Open a spreadsheet workbook based on file extension.
///
/// Currently supports:
/// - `.xls` (via `formula-xls`)
/// - `.xlsb` (via `formula-xlsb`)
pub fn open_workbook(path: impl AsRef<Path>) -> Result<Workbook, Error> {
    let path = path.as_ref();
    match path
        .extension()
        .and_then(|s| s.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase()
        .as_str()
    {
        "xls" => Ok(Workbook::Xls(xls::import_xls_path(path)?)),
        "xlsb" => Ok(Workbook::Xlsb(xlsb::XlsbWorkbook::open(path)?)),
        other => Err(Error::UnsupportedExtension(other.to_string())),
    }
}
