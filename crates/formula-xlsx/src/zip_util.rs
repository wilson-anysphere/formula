use std::io::{Read, Seek};

use zip::read::ZipFile;
use zip::result::ZipError;
use zip::ZipArchive;

/// Open a ZIP entry by name, tolerating a leading `/` mismatch.
///
/// Valid XLSX/XLSM files should *not* include a leading `/` in the underlying ZIP entry names,
/// but some producers do. Since SpreadsheetML relationship targets and other part-name handling in
/// this crate assume canonical names like `xl/workbook.xml`, we try both `name` and `/{name}` (or
/// the stripped variant when `name` itself starts with `/`).
///
/// Note: `ZipFile` borrows `ZipArchive`. We can't naively call `archive.by_name()` twice and
/// return the borrowed `ZipFile` because that triggers borrow-checker errors. Instead, this helper
/// inspects `archive.file_names()` first to decide which entry name to open, then calls
/// `archive.by_name()` exactly once.
pub(crate) fn open_zip_part<'a, R: Read + Seek>(
    archive: &'a mut ZipArchive<R>,
    name: &str,
) -> Result<ZipFile<'a>, ZipError> {
    let alt = if let Some(stripped) = name.strip_prefix('/') {
        stripped.to_string()
    } else {
        let mut with_slash = String::with_capacity(name.len() + 1);
        with_slash.push('/');
        with_slash.push_str(name);
        with_slash
    };

    let mut candidate = None::<String>;
    for entry in archive.file_names() {
        if entry == name {
            candidate = Some(name.to_string());
            break;
        }
        if entry == alt.as_str() {
            candidate = Some(alt.clone());
        }
    }

    match candidate {
        Some(name) => archive.by_name(&name),
        None => Err(ZipError::FileNotFound),
    }
}

