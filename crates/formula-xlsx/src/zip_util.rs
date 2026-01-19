use std::io::{Read, Seek};

use zip::read::ZipFile;
use zip::result::ZipError;
use zip::ZipArchive;

use crate::XlsxError;

/// Default maximum uncompressed size permitted for any single ZIP part inflated into memory.
///
/// This is a defense-in-depth guardrail against ZIP bombs (tiny compressed size, huge uncompressed
/// size) and forged ZIP metadata (e.g. an incorrect `uncompressed_size` field).
pub(crate) const DEFAULT_MAX_ZIP_PART_BYTES: u64 = 256 * 1024 * 1024; // 256MiB

/// Default maximum total uncompressed bytes permitted across multi-part ZIP inflation APIs.
///
/// This applies to APIs that may read many ZIP parts into memory, such as
/// [`crate::XlsxPackage::from_bytes`].
pub(crate) const DEFAULT_MAX_ZIP_TOTAL_BYTES: u64 = 512 * 1024 * 1024; // 512MiB

pub(crate) fn zip_part_names_equivalent(a: &str, b: &str) -> bool {
    fn hex_val(b: u8) -> Option<u8> {
        match b {
            b'0'..=b'9' => Some(b - b'0'),
            b'a'..=b'f' => Some(b - b'a' + 10),
            b'A'..=b'F' => Some(b - b'A' + 10),
            _ => None,
        }
    }

    struct Normalized<'a> {
        bytes: &'a [u8],
        in_leading_separators: bool,
    }

    impl<'a> Normalized<'a> {
        fn new(s: &'a str) -> Self {
            Self {
                bytes: s.as_bytes(),
                in_leading_separators: true,
            }
        }

        fn next_byte(&mut self) -> Option<u8> {
            loop {
                let b = *self.bytes.first()?;
                let decoded = if b == b'%' && self.bytes.len() >= 3 {
                    let hi = self.bytes[1];
                    let lo = self.bytes[2];
                    if let (Some(hi), Some(lo)) = (hex_val(hi), hex_val(lo)) {
                        self.bytes = &self.bytes[3..];
                        (hi << 4) | lo
                    } else {
                        self.bytes = &self.bytes[1..];
                        b
                    }
                } else {
                    self.bytes = &self.bytes[1..];
                    b
                };

                // Skip any number of leading `/` or `\` separators, even when percent-encoded.
                if self.in_leading_separators && matches!(decoded, b'/' | b'\\') {
                    continue;
                }
                self.in_leading_separators = false;

                let normalized = if decoded == b'\\' {
                    b'/'
                } else {
                    decoded.to_ascii_lowercase()
                };
                return Some(normalized);
            }
        }
    }

    let mut a = Normalized::new(a);
    let mut b = Normalized::new(b);
    loop {
        match (a.next_byte(), b.next_byte()) {
            (Some(a), Some(b)) if a == b => continue,
            (None, None) => return true,
            _ => return false,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    use std::io::{Cursor, Write};

    use zip::write::FileOptions;
    use zip::ZipWriter;

    #[test]
    fn equivalent_handles_case_separators_and_leading_slashes() {
        assert!(zip_part_names_equivalent("XL\\Workbook.xml", "xl/workbook.xml"));
        assert!(zip_part_names_equivalent("/xl/workbook.xml", "xl/workbook.xml"));
        assert!(zip_part_names_equivalent("\\xl\\workbook.xml", "xl/workbook.xml"));
    }

    #[test]
    fn equivalent_handles_percent_encoded_names() {
        assert!(zip_part_names_equivalent(
            "xl/worksheets/sheet 1.xml",
            "xl/worksheets/sheet%201.xml"
        ));
        assert!(zip_part_names_equivalent(
            "xl/worksheets/sheet1.xml",
            "xl%2Fworksheets%2Fsheet1.xml"
        ));
        assert!(zip_part_names_equivalent(
            "xl/worksheets/sheet1.xml",
            "%2Fxl%2Fworksheets%2Fsheet1.xml"
        ));
    }

    fn build_zip(entries: &[(&str, &[u8])]) -> Vec<u8> {
        let cursor = Cursor::new(Vec::new());
        let mut zip = ZipWriter::new(cursor);
        let options =
            FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);
        for (name, bytes) in entries {
            zip.start_file(*name, options).unwrap();
            zip.write_all(bytes).unwrap();
        }
        zip.finish().unwrap().into_inner()
    }

    #[test]
    fn read_zip_part_optional_with_limit_allows_within_limit() {
        let bytes = build_zip(&[("a.txt", b"hello world")]); // 11 bytes
        let mut archive = ZipArchive::new(Cursor::new(bytes)).unwrap();

        let part = read_zip_part_optional_with_limit(&mut archive, "a.txt", 11)
            .unwrap()
            .unwrap();
        assert_eq!(part, b"hello world");
    }

    #[test]
    fn read_zip_part_optional_with_limit_errors_when_too_large() {
        let bytes = build_zip(&[("a.txt", b"hello world")]); // 11 bytes
        let mut archive = ZipArchive::new(Cursor::new(bytes)).unwrap();

        let err = read_zip_part_optional_with_limit(&mut archive, "a.txt", 10).unwrap_err();
        match err {
            XlsxError::PartTooLarge { part, .. } => {
                assert_eq!(part, "a.txt");
            }
            other => panic!("expected PartTooLarge, got {other:?}"),
        }
    }

    #[test]
    fn open_zip_part_prefers_exact_over_equivalent() {
        let bytes = build_zip(&[
            ("XL\\Workbook.xml", b"equivalent"),
            ("xl/workbook.xml", b"exact"),
        ]);
        let mut archive = ZipArchive::new(Cursor::new(bytes)).unwrap();
        let mut file = open_zip_part(&mut archive, "xl/workbook.xml").unwrap();
        let mut out = String::new();
        file.read_to_string(&mut out).unwrap();
        assert_eq!(out, "exact");
    }

    #[test]
    fn open_zip_part_handles_leading_slash_variant() {
        let bytes = build_zip(&[("/xl/workbook.xml", b"with_slash")]);
        let mut archive = ZipArchive::new(Cursor::new(bytes)).unwrap();
        let mut file = open_zip_part(&mut archive, "xl/workbook.xml").unwrap();
        let mut out = String::new();
        file.read_to_string(&mut out).unwrap();
        assert_eq!(out, "with_slash");
    }
}

pub(crate) fn zip_part_name_starts_with(name: &str, canonical_prefix: &[u8]) -> bool {
    fn hex_val(b: u8) -> Option<u8> {
        match b {
            b'0'..=b'9' => Some(b - b'0'),
            b'a'..=b'f' => Some(b - b'a' + 10),
            b'A'..=b'F' => Some(b - b'A' + 10),
            _ => None,
        }
    }

    struct Normalized<'a> {
        bytes: &'a [u8],
        in_leading_separators: bool,
    }

    impl<'a> Normalized<'a> {
        fn new(s: &'a str) -> Self {
            Self {
                bytes: s.as_bytes(),
                in_leading_separators: true,
            }
        }

        fn next_byte(&mut self) -> Option<u8> {
            loop {
                let b = *self.bytes.first()?;
                let decoded = if b == b'%' && self.bytes.len() >= 3 {
                    let hi = self.bytes[1];
                    let lo = self.bytes[2];
                    if let (Some(hi), Some(lo)) = (hex_val(hi), hex_val(lo)) {
                        self.bytes = &self.bytes[3..];
                        (hi << 4) | lo
                    } else {
                        self.bytes = &self.bytes[1..];
                        b
                    }
                } else {
                    self.bytes = &self.bytes[1..];
                    b
                };

                if self.in_leading_separators && matches!(decoded, b'/' | b'\\') {
                    continue;
                }
                self.in_leading_separators = false;

                let normalized = if decoded == b'\\' {
                    b'/'
                } else {
                    decoded.to_ascii_lowercase()
                };
                return Some(normalized);
            }
        }
    }

    let mut n = Normalized::new(name);
    for &b in canonical_prefix {
        match n.next_byte() {
            Some(got) if got == b => {}
            _ => return false,
        }
    }
    true
}

/// Compute a canonicalized key for a ZIP entry/part name suitable for case- and separator-insensitive
/// lookup.
///
/// The normalization matches [`zip_part_names_equivalent`]:
/// - percent-decodes valid `%xx` sequences
/// - strips leading path separators (`/` or `\`), including when percent-encoded
/// - normalizes `\` to `/`
/// - ASCII-lowercases
///
/// We return a byte vector (not a `String`) so we can represent arbitrary percent-decoded bytes
/// without requiring valid UTF-8.
pub(crate) fn zip_part_name_lookup_key(name: &str) -> Result<Vec<u8>, XlsxError> {
    fn hex_val(b: u8) -> Option<u8> {
        match b {
            b'0'..=b'9' => Some(b - b'0'),
            b'a'..=b'f' => Some(b - b'a' + 10),
            b'A'..=b'F' => Some(b - b'A' + 10),
            _ => None,
        }
    }

    let mut bytes = name.as_bytes();
    let mut out = Vec::new();
    out.try_reserve(bytes.len())
        .map_err(|_| XlsxError::AllocationFailure("zip_part_name_lookup_key"))?;
    let mut in_leading_separators = true;
    while let Some(&b) = bytes.first() {
        let decoded = if b == b'%' && bytes.len() >= 3 {
            let hi = bytes[1];
            let lo = bytes[2];
            if let (Some(hi), Some(lo)) = (hex_val(hi), hex_val(lo)) {
                bytes = &bytes[3..];
                (hi << 4) | lo
            } else {
                bytes = &bytes[1..];
                b
            }
        } else {
            bytes = &bytes[1..];
            b
        };

        // Skip any number of leading `/` or `\` separators, even when percent-encoded.
        if in_leading_separators && matches!(decoded, b'/' | b'\\') {
            continue;
        }
        in_leading_separators = false;

        let normalized = if decoded == b'\\' {
            b'/'
        } else {
            decoded.to_ascii_lowercase()
        };
        out.push(normalized);
    }
    Ok(out)
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
/// inspects `archive.file_names()` first to decide which entry index to open, then calls
/// `archive.by_index()` exactly once.
pub(crate) fn open_zip_part<'a, R: Read + Seek>(
    archive: &'a mut ZipArchive<R>,
    name: &str,
) -> Result<ZipFile<'a>, ZipError> {
    fn is_alt_slash_variant(entry: &str, name: &str) -> bool {
        if let Some(stripped) = name.strip_prefix('/') {
            entry == stripped
        } else {
            entry.strip_prefix('/').is_some_and(|rest| rest == name)
        }
    }

    let mut candidate = None::<(usize, u8)>;
    for (idx, entry) in archive.file_names().enumerate() {
        if entry == name {
            candidate = Some((idx, 3));
            break;
        }
        if is_alt_slash_variant(entry, name) {
            candidate = Some((idx, 2));
            continue;
        }

        if zip_part_names_equivalent(entry, name) {
            if candidate.as_ref().map_or(true, |(_, score)| *score < 1) {
                candidate = Some((idx, 1));
            }
        }
    }

    match candidate {
        Some((idx, _)) => archive.by_index(idx),
        None => Err(ZipError::FileNotFound),
    }
}

#[derive(Debug, Clone)]
pub(crate) struct ZipInflateBudget {
    max_total_bytes: u64,
    used_bytes: u64,
}

impl ZipInflateBudget {
    pub(crate) fn new(max_total_bytes: u64) -> Self {
        Self {
            max_total_bytes,
            used_bytes: 0,
        }
    }

    pub(crate) fn remaining_bytes(&self) -> u64 {
        self.max_total_bytes.saturating_sub(self.used_bytes)
    }

    pub(crate) fn used_bytes(&self) -> u64 {
        self.used_bytes
    }

    pub(crate) fn max_total_bytes(&self) -> u64 {
        self.max_total_bytes
    }

    pub(crate) fn consume(&mut self, _part: &str, bytes: u64) -> Result<(), XlsxError> {
        self.used_bytes = self.used_bytes.checked_add(bytes).unwrap_or(u64::MAX);
        if self.used_bytes > self.max_total_bytes {
            return Err(XlsxError::PackageTooLarge {
                total: self.used_bytes,
                max: self.max_total_bytes,
            });
        }
        Ok(())
    }
}

/// Read a ZIP entry into memory with an uncompressed size limit.
///
/// This helper does **not** trust ZIP metadata alone. It:
/// - checks the declared uncompressed size (`ZipFile::size()`) as a fast-path;
/// - reads via `Read::take(max + 1)` to guard against forged metadata;
/// - and errors deterministically if more than `max_bytes` are observed.
pub(crate) fn read_zip_file_bytes_with_limit(
    file: &mut ZipFile<'_>,
    part: &str,
    max_bytes: u64,
) -> Result<Vec<u8>, XlsxError> {
    read_zip_file_bytes_with_optional_budget(file, part, max_bytes, None)
}

pub(crate) fn read_zip_file_bytes_with_budget(
    file: &mut ZipFile<'_>,
    part: &str,
    max_part_bytes: u64,
    budget: &mut ZipInflateBudget,
) -> Result<Vec<u8>, XlsxError> {
    read_zip_file_bytes_with_optional_budget(file, part, max_part_bytes, Some(budget))
}

fn read_zip_file_bytes_with_optional_budget(
    file: &mut ZipFile<'_>,
    part: &str,
    max_part_bytes: u64,
    mut budget: Option<&mut ZipInflateBudget>,
) -> Result<Vec<u8>, XlsxError> {
    fn add_or_max(a: u64, b: u64) -> u64 {
        a.checked_add(b).unwrap_or(u64::MAX)
    }

    let declared_size = file.size();
    let used_before = budget.as_ref().map(|b| b.used_bytes()).unwrap_or(0);
    let remaining_total = budget
        .as_ref()
        .map(|b| b.remaining_bytes())
        .unwrap_or(u64::MAX);

    let effective_max = max_part_bytes.min(remaining_total);
    let limit_is_total = effective_max < max_part_bytes;
    if budget.is_some() && effective_max == 0 {
        let max_total = budget.as_ref().map(|b| b.max_total_bytes()).unwrap_or(0);
        return Err(XlsxError::PackageTooLarge {
            total: add_or_max(used_before, 1),
            max: max_total,
        });
    }

    // Fast-path: reject based on declared uncompressed size.
    if declared_size > max_part_bytes {
        return Err(XlsxError::PartTooLarge {
            part: part.to_string(),
            size: declared_size,
            max: max_part_bytes,
        });
    }
    if limit_is_total && declared_size > effective_max {
        let max_total = budget.as_ref().map(|b| b.max_total_bytes()).unwrap_or(0);
        return Err(XlsxError::PackageTooLarge {
            total: add_or_max(used_before, declared_size),
            max: max_total,
        });
    }

    // Don't trust ZIP metadata alone. Guard against incorrect/forged size fields by limiting reads
    // to `effective_max + 1` and erroring if we see more than `effective_max` bytes.
    let mut buf = Vec::new();
    let read_limit = effective_max.checked_add(1).unwrap_or(u64::MAX);
    let mut reader = file.take(read_limit);
    reader.read_to_end(&mut buf)?;

    let observed = buf.len() as u64;
    if observed > effective_max {
        if limit_is_total {
            let max_total = budget.as_ref().map(|b| b.max_total_bytes()).unwrap_or(0);
            return Err(XlsxError::PackageTooLarge {
                total: add_or_max(used_before, observed),
                max: max_total,
            });
        }
        return Err(XlsxError::PartTooLarge {
            part: part.to_string(),
            size: observed,
            max: max_part_bytes,
        });
    }

    if let Some(budget) = budget.as_mut() {
        budget.consume(part, observed)?;
    }

    Ok(buf)
}

/// Read a ZIP part by name, returning `Ok(None)` when the entry does not exist.
pub(crate) fn read_zip_part_optional_with_limit<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    name: &str,
    max_part_bytes: u64,
) -> Result<Option<Vec<u8>>, XlsxError> {
    match open_zip_part(archive, name) {
        Ok(mut file) => {
            if file.is_dir() {
                return Ok(None);
            }
            let buf = read_zip_file_bytes_with_limit(&mut file, name, max_part_bytes)?;
            Ok(Some(buf))
        }
        Err(zip::result::ZipError::FileNotFound) => Ok(None),
        Err(err) => Err(err.into()),
    }
}

/// Read a ZIP part by name while also consuming from a shared "total inflated bytes" budget.
pub(crate) fn read_zip_part_optional_with_budget<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    name: &str,
    max_part_bytes: u64,
    budget: &mut ZipInflateBudget,
) -> Result<Option<Vec<u8>>, XlsxError> {
    match open_zip_part(archive, name) {
        Ok(mut file) => {
            if file.is_dir() {
                return Ok(None);
            }
            let buf = read_zip_file_bytes_with_budget(&mut file, name, max_part_bytes, budget)?;
            Ok(Some(buf))
        }
        Err(zip::result::ZipError::FileNotFound) => Ok(None),
        Err(err) => Err(err.into()),
    }
}
