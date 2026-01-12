use std::io::{Read, Seek};

use zip::result::ZipError;
use zip::ZipArchive;

fn normalize_zip_entry_name(name: &str) -> String {
    let mut normalized = name.trim_start_matches('/');
    let replaced;
    if normalized.contains('\\') {
        replaced = normalized.replace('\\', "/");
        normalized = &replaced;
    }
    normalized.to_string()
}

pub(crate) fn find_zip_entry_case_insensitive<R: Read + Seek>(
    zip: &ZipArchive<R>,
    name: &str,
) -> Option<String> {
    let target = normalize_zip_entry_name(name);
    for candidate in zip.file_names() {
        let normalized = normalize_zip_entry_name(candidate);
        if normalized.eq_ignore_ascii_case(&target) {
            return Some(candidate.to_string());
        }
    }
    None
}

/// Read a ZIP entry by name, tolerating a leading `/`, Windows separators, and case mismatches.
pub(crate) fn read_zip_entry_bytes<R: Read + Seek>(
    zip: &mut ZipArchive<R>,
    name: &str,
) -> Result<Option<Vec<u8>>, ZipError> {
    let Some(actual) = find_zip_entry_case_insensitive(zip, name) else {
        return Ok(None);
    };

    let mut entry = zip.by_name(&actual)?;
    if entry.is_dir() {
        return Ok(None);
    }

    let mut buf = Vec::with_capacity(entry.size() as usize);
    entry.read_to_end(&mut buf).map_err(ZipError::Io)?;
    Ok(Some(buf))
}

