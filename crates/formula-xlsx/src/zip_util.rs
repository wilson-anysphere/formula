use std::io::{Read, Seek};

use zip::read::ZipFile;
use zip::result::ZipError;
use zip::ZipArchive;

pub(crate) fn zip_part_names_equivalent(a: &str, b: &str) -> bool {
    fn strip_leading_separators(mut bytes: &[u8]) -> &[u8] {
        while matches!(bytes.first(), Some(b'/' | b'\\')) {
            bytes = &bytes[1..];
        }
        bytes
    }

    let a = strip_leading_separators(a.as_bytes());
    let b = strip_leading_separators(b.as_bytes());
    if a.len() != b.len() {
        return false;
    }

    for (&a, &b) in a.iter().zip(b.iter()) {
        let a = if a == b'\\' { b'/' } else { a.to_ascii_lowercase() };
        let b = if b == b'\\' { b'/' } else { b.to_ascii_lowercase() };
        if a != b {
            return false;
        }
    }

    true
}

/// Open a ZIP entry by name, tolerating common producer mistakes:
/// - leading `/` mismatch
/// - Windows-style `\` path separators
/// - ASCII case differences
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

    let mut candidate = None::<(String, u8)>;
    for entry in archive.file_names() {
        if entry == name {
            candidate = Some((entry.to_string(), 3));
            break;
        }
        if entry == alt.as_str() {
            candidate = Some((entry.to_string(), 2));
            continue;
        }

        if zip_part_names_equivalent(entry, name) {
            if candidate.as_ref().map_or(true, |(_, score)| *score < 1) {
                candidate = Some((entry.to_string(), 1));
            }
        }
    }

    match candidate {
        Some((name, _)) => archive.by_name(&name),
        None => Err(ZipError::FileNotFound),
    }
}
