use std::io::{Read, Seek};

use zip::read::ZipFile;
use zip::result::ZipError;
use zip::ZipArchive;

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
