#[cfg(test)]
mod cryptoapi_rc4;

use std::borrow::Cow;
use std::path::{Path, PathBuf};

use encoding_rs::{UTF_16BE, UTF_16LE, WINDOWS_1252};
use formula_fs::{atomic_write, AtomicWriteError};
use formula_model::import::{import_csv_into_workbook, CsvImportError, CsvOptions};
use formula_model::sanitize_sheet_name;
pub use formula_xls as xls;
pub use formula_xlsb as xlsb;
pub use formula_xlsx as xlsx;
use std::io::{Read, Seek};

mod encryption_info;
pub mod offcrypto;
pub use encryption_info::{extract_agile_encryption_info_xml, EncryptionInfoXmlError};
mod rc4_cryptoapi;
pub use rc4_cryptoapi::{HashAlg, Rc4CryptoApiDecryptReader, Rc4CryptoApiEncryptedPackageError};
#[cfg(any(test, feature = "offcrypto"))]
mod ms_offcrypto;
mod rc4_encrypted_package;
mod encrypted_package;
pub use encrypted_package::StandardAesEncryptedPackageReader;

#[cfg(feature = "encrypted-workbooks")]
mod encrypted_ooxml;
#[cfg(any(test, feature = "encrypted-workbooks"))]
mod encrypted_package_reader;
const OLE_MAGIC: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
const PARQUET_MAGIC: [u8; 4] = *b"PAR1";

pub(crate) fn parse_encrypted_package_size_prefix_bytes(prefix: [u8; 8], ciphertext_len: Option<u64>) -> u64 {
    // MS-OFFCRYPTO describes this field as a `u64le`, but some producers/libraries treat it as
    // `(u32 size, u32 reserved)` (often with `reserved = 0`). When the high DWORD is non-zero but
    // the combined 64-bit value is not plausible for the available ciphertext, fall back to the
    // low DWORD for compatibility.
    //
    // Note: avoid falling back when the low DWORD is zero. Some real files may have true 64-bit
    // sizes that are exact multiples of 2^32, so `lo=0, hi!=0` must be treated as a 64-bit size.
    let len_lo = u32::from_le_bytes([prefix[0], prefix[1], prefix[2], prefix[3]]) as u64;
    let len_hi = u32::from_le_bytes([prefix[4], prefix[5], prefix[6], prefix[7]]) as u64;
    let size_u64 = len_lo | (len_hi << 32);

    match ciphertext_len {
        Some(ciphertext_len) => {
            if len_lo != 0 && len_hi != 0 && size_u64 > ciphertext_len && len_lo <= ciphertext_len {
                len_lo
            } else {
                size_u64
            }
        }
        None => {
            // Without ciphertext length (e.g. streaming readers), prefer compatibility with
            // producers that treat the high DWORD as reserved.
            if len_hi != 0 && len_lo != 0 {
                len_lo
            } else {
                size_u64
            }
        }
    }
}

pub(crate) fn parse_encrypted_package_original_size(encrypted_package: &[u8]) -> Option<u64> {
    if encrypted_package.len() < 8 {
        return None;
    }
    let mut prefix = [0u8; 8];
    prefix.copy_from_slice(&encrypted_package[..8]);
    let ciphertext_len = encrypted_package.len().saturating_sub(8) as u64;
    Some(parse_encrypted_package_size_prefix_bytes(prefix, Some(ciphertext_len)))
}

#[cfg(test)]
mod encrypted_package_size_prefix_tests {
    use super::*;

    #[test]
    fn does_not_fall_back_when_low_dword_is_zero() {
        // Some files may store a true 64-bit size that is an exact multiple of 2^32 (lo=0).
        // The reserved-high-DWORD compatibility fallback must not misinterpret that as 0.
        let mut prefix = [0u8; 8];
        prefix[4..].copy_from_slice(&1u32.to_le_bytes()); // hi=1, lo=0 => 2^32

        assert_eq!(
            parse_encrypted_package_size_prefix_bytes(prefix, Some(0)),
            1u64 << 32
        );
        assert_eq!(parse_encrypted_package_size_prefix_bytes(prefix, None), 1u64 << 32);
    }
}
// BIFF record ids for legacy `.xls` encryption detection.
//
// Presence of `FILEPASS` in the workbook globals substream indicates the workbook stream is
// encrypted/password-protected.
const BIFF_RECORD_FILEPASS: u16 = 0x002F;
const BIFF_RECORD_EOF: u16 = 0x000A;
const BIFF_RECORD_BOF_BIFF8: u16 = 0x0809;
const BIFF_RECORD_BOF_BIFF5: u16 = 0x0009;

// Maximum bytes to inspect for text/CSV sniffing.
const TEXT_SNIFF_LEN: usize = 16 * 1024;

#[derive(Debug, thiserror::Error)]
pub enum Error {
    #[error("unsupported extension `{extension}` for workbook `{path}`")]
    UnsupportedExtension { path: PathBuf, extension: String },
    #[error(
        "password required: workbook `{path}` is password-protected/encrypted; supply a password via `open_workbook_with_password(..)` / `open_workbook_model_with_password(..)`"
    )]
    PasswordRequired { path: PathBuf },
    #[error("invalid password for workbook `{path}`")]
    InvalidPassword { path: PathBuf },
    #[error(
        "unsupported encrypted OOXML workbook `{path}`: EncryptionInfo version {version_major}.{version_minor} is invalid or not supported"
    )]
    UnsupportedOoxmlEncryption {
        path: PathBuf,
        version_major: u16,
        version_minor: u16,
    },
    #[error("unsupported encryption for workbook `{path}`: {kind}")]
    UnsupportedEncryption { path: PathBuf, kind: String },
    #[error(
        "unsupported encrypted workbook `{path}`: decrypted workbook kind `{kind}` is not supported"
    )]
    UnsupportedEncryptedWorkbookKind { path: PathBuf, kind: &'static str },
    #[error(
        "password required: workbook `{path}` is password-protected/encrypted (legacy `.xls` encryption); supply a password via `open_workbook_with_password(..)` / `open_workbook_model_with_password(..)`"
    )]
    EncryptedWorkbook { path: PathBuf },
    #[error(
        "parquet support not enabled: workbook `{path}` appears to be a `.parquet` file; rebuild with the `formula-io/parquet` feature"
    )]
    ParquetSupportNotEnabled { path: PathBuf },
    #[error("failed to open workbook `{path}`: {source}")]
    OpenIo {
        path: PathBuf,
        #[source]
        source: std::io::Error,
    },
    #[error("failed to detect workbook format for `{path}`: {source}")]
    DetectIo {
        path: PathBuf,
        #[source]
        source: std::io::Error,
    },
    #[error("failed to detect workbook format for `{path}`: {source}")]
    DetectZip {
        path: PathBuf,
        #[source]
        source: zip::result::ZipError,
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
    #[error("failed to open `.csv` workbook `{path}`: {source}")]
    OpenCsv {
        path: PathBuf,
        #[source]
        source: CsvImportError,
    },
    #[error("failed to open `.parquet` workbook `{path}`: {source}")]
    OpenParquet {
        path: PathBuf,
        #[source]
        source: Box<dyn std::error::Error + Send + Sync>,
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
    #[cfg(feature = "encrypted-workbooks")]
    #[error("failed to encrypt workbook `{path}`: {source}")]
    SaveOoxmlEncryption {
        path: PathBuf,
        #[source]
        source: formula_office_crypto::OfficeCryptoError,
    },
}

/// A workbook opened from disk.
#[derive(Debug)]
pub enum Workbook {
    /// XLSX/XLSM opened as an Open Packaging Convention (OPC) package.
    ///
    /// This preserves unknown parts (e.g. `customXml/`, `xl/vbaProject.bin`) byte-for-byte.
    Xlsx(xlsx::XlsxLazyPackage),
    Xls(xls::XlsImportResult),
    Xlsb(xlsb::XlsbWorkbook),
    /// A workbook represented as an in-memory model (e.g. imported from a non-OPC format like CSV/Parquet).
    Model(formula_model::Workbook),
}

/// Best-effort workbook format classification.
#[derive(Debug, Copy, Clone, PartialEq, Eq)]
pub enum WorkbookFormat {
    Xlsx,
    Xlsm,
    Xlsb,
    Xls,
    Csv,
    Parquet,
    Unknown,
}

/// Best-effort workbook encryption classification.
///
/// This is intended for UI preflight (e.g. prompting for a password) and corpus tooling. It does
/// not attempt to decrypt or open the workbook.
#[derive(Debug, Copy, Clone, PartialEq, Eq)]
pub enum WorkbookEncryption {
    /// Workbook does not appear to be encrypted / password-protected.
    None,
    /// Office-encrypted OOXML package stored in an OLE compound file via the `EncryptionInfo` +
    /// `EncryptedPackage` streams (e.g. a password-protected `.xlsx`).
    OoxmlEncryptedPackage {
        /// Optional scheme details (e.g. Standard vs Agile) once `EncryptionInfo` parsing is
        /// implemented for this helper.
        scheme: Option<OoxmlEncryptedPackageScheme>,
    },
    /// Legacy BIFF `.xls` workbook encryption indicated by a `FILEPASS` record in the workbook
    /// stream.
    LegacyXlsFilePass {
        /// Optional scheme details (best-effort) derived from the `FILEPASS` record payload when
        /// available.
        scheme: Option<LegacyXlsFilePassScheme>,
    },
}

/// OOXML `EncryptedPackage` encryption scheme (from `EncryptionInfo`).
///
/// This is currently a placeholder until `EncryptionInfo` parsing is implemented for
/// [`detect_workbook_encryption`].
#[derive(Debug, Copy, Clone, PartialEq, Eq)]
pub enum OoxmlEncryptedPackageScheme {
    Unknown,
}

/// Legacy `.xls` `FILEPASS` encryption scheme.
///
/// This is best-effort and is derived from the BIFF8 `FILEPASS` header when available:
/// - `wEncryptionType = 0x0000` => XOR
/// - `wEncryptionType = 0x0001, wEncryptionSubType = 0x0001` => RC4 ("standard")
/// - `wEncryptionType = 0x0001, wEncryptionSubType = 0x0002` => RC4 CryptoAPI
#[derive(Debug, Copy, Clone, PartialEq, Eq)]
pub enum LegacyXlsFilePassScheme {
    Xor,
    Rc4,
    Rc4CryptoApi,
    Unknown,
}

/// Options controlling workbook open behavior.
#[derive(Debug, Clone, Default)]
pub struct OpenOptions {
    /// Optional password for encrypted workbooks.
    pub password: Option<String>,
}

/// Default maximum plaintext size allowed when decrypting an OOXML `EncryptedPackage`.
///
/// This is a defensive guardrail against malicious/corrupt encrypted workbooks that claim a
/// pathological decrypted size in the `EncryptedPackage` header.
///
/// Note: the decrypted bytes are a ZIP/OPC package (already compressed), so the plaintext size
/// should be close to the ciphertext size. We allow some slack and also apply an absolute cap.
const DEFAULT_MAX_OFFCRYPTO_OUTPUT_SIZE: u64 = 1024 * 1024 * 1024; // 1GiB

fn default_offcrypto_max_output_size_u64(encrypted_package_len: u64) -> u64 {
    let scaled = encrypted_package_len.saturating_mul(4);
    scaled.min(DEFAULT_MAX_OFFCRYPTO_OUTPUT_SIZE)
}

fn default_offcrypto_max_output_size(encrypted_package_len: usize) -> u64 {
    default_offcrypto_max_output_size_u64(encrypted_package_len as u64)
}

fn encrypted_package_plaintext_len_is_plausible(plaintext_len: u64, ciphertext_len: u64) -> bool {
    // The `EncryptedPackage` ciphertext should be >= plaintext length (encryption does not compress
    // and typically adds padding/metadata). Reject obviously inconsistent or pathological headers up
    // front so we can surface a stable "unsupported encryption" error instead of attempting a ZIP
    // parse over a truncated reader (and potentially allocating attacker-controlled buffers).
    if plaintext_len > ciphertext_len {
        return false;
    }

    // Also apply a conservative absolute cap so we don't attempt to open extremely large encrypted
    // packages (which would require holding decrypted ZIP bytes in memory in some paths, or at
    // least incur very expensive IO).
    let encrypted_package_len = ciphertext_len.saturating_add(8);
    plaintext_len <= default_offcrypto_max_output_size_u64(encrypted_package_len)
}

/// Decrypt an ECMA-376 Standard-encrypted `EncryptedPackage` stream with a derived AES key.
///
/// This is a low-level helper used by encrypted workbook implementations. Callers that do not
/// explicitly supply size limits get a conservative default derived from `encrypted_package.len()`.
pub fn standard_decrypt_encrypted_package(
    key: &[u8],
    encrypted_package: &[u8],
) -> Result<Vec<u8>, formula_offcrypto::OffcryptoError> {
    let mut opts = formula_offcrypto::DecryptOptions::default();
    opts.limits.max_output_size = Some(default_offcrypto_max_output_size(encrypted_package.len()));
    formula_offcrypto::standard_decrypt_encrypted_package(key, encrypted_package, &opts)
}

fn looks_like_text_csv_prefix(prefix: &[u8]) -> bool {
    if prefix.is_empty() {
        return false;
    }

    // Excel can export delimited text as UTF-16 (e.g. via "Unicode Text"). Those files contain
    // NUL bytes, so detect a UTF-16 BOM and apply the heuristics to a decoded UTF-8 preview instead
    // of rejecting the file as binary.
    if prefix.starts_with(&[0xFF, 0xFE]) || prefix.starts_with(&[0xFE, 0xFF]) {
        let (encoding, rest) = if prefix.starts_with(&[0xFF, 0xFE]) {
            (UTF_16LE, &prefix[2..])
        } else {
            (UTF_16BE, &prefix[2..])
        };
        // UTF-16 requires an even number of bytes; ignore a trailing odd byte in the preview.
        let rest = &rest[..rest.len().saturating_sub(rest.len() % 2)];
        let (cow, _) = encoding.decode_without_bom_handling(rest);
        return looks_like_text_csv_str(cow.as_ref());
    }

    // Avoid misclassifying binary formats as text.
    if prefix.iter().any(|b| *b == 0) {
        // Best-effort: handle UTF-16 inputs that lack a BOM by detecting the "ASCII UTF-16" NUL
        // byte pattern and running heuristics on a decoded preview instead of rejecting as binary.
        if let Some(encoding) = detect_utf16_bomless_encoding(prefix) {
            let prefix = &prefix[..prefix.len().saturating_sub(prefix.len() % 2)];
            let (cow, _) = encoding.decode_without_bom_handling(prefix);
            return looks_like_text_csv_str(cow.as_ref());
        }
        return false;
    }

    // Reject disallowed control bytes (keep common whitespace used in delimited text).
    for &b in prefix {
        if b < 0x20 && !matches!(b, b'\t' | b'\n' | b'\r') {
            return false;
        }
        if b == 0x7F {
            return false;
        }
    }

    let prefix = prefix
        .strip_prefix(&[0xEF, 0xBB, 0xBF]) // UTF-8 BOM (common in Excel-exported CSVs)
        .unwrap_or(prefix);

    // Prefer UTF-8, but accept Windows-1252 (matching CSV import behavior).
    let decoded: Cow<'_, str> = match std::str::from_utf8(prefix) {
        Ok(s) => Cow::Borrowed(s),
        Err(_) => {
            let (cow, _, _) = WINDOWS_1252.decode(prefix);
            cow
        }
    };
    let decoded = decoded.as_ref();

    looks_like_text_csv_str(decoded)
}

fn detect_utf16_bomless_encoding(prefix: &[u8]) -> Option<&'static encoding_rs::Encoding> {
    let len = prefix.len().min(TEXT_SNIFF_LEN);
    let len = len - (len % 2);
    if len < 4 {
        return None;
    }
    let sample = &prefix[..len];

    let mut le_markers = 0usize;
    let mut be_markers = 0usize;
    let mut even_zero = 0usize;
    let mut odd_zero = 0usize;

    for (idx, b) in sample.iter().enumerate() {
        if *b == 0 {
            if idx % 2 == 0 {
                even_zero += 1;
            } else {
                odd_zero += 1;
            }
        }
    }

    const MARKERS: [u8; 6] = [b',', b';', b'\t', b'|', b'\r', b'\n'];
    for pair in sample.chunks_exact(2) {
        let a = pair[0];
        let b = pair[1];
        if b == 0 && MARKERS.contains(&a) {
            le_markers += 1;
        }
        if a == 0 && MARKERS.contains(&b) {
            be_markers += 1;
        }
    }

    const MIN_MARKERS: usize = 2;
    if le_markers >= MIN_MARKERS || be_markers >= MIN_MARKERS {
        if le_markers > be_markers {
            return Some(UTF_16LE);
        }
        if be_markers > le_markers {
            return Some(UTF_16BE);
        }
    }

    if odd_zero > even_zero.saturating_mul(3) {
        Some(UTF_16LE)
    } else if even_zero > odd_zero.saturating_mul(3) {
        Some(UTF_16BE)
    } else {
        None
    }
}

fn looks_like_text_csv_str(decoded: &str) -> bool {
    // Require a plausible delimiter.
    if !(decoded.contains(',')
        || decoded.contains(';')
        || decoded.contains('\t')
        || decoded.contains('|'))
    {
        return false;
    }

    // Prefer to see at least one record terminator, but allow single-line delimited text. This is
    // useful for extension-less temp files or single-row exports where the trailing newline is
    // missing.
    let has_newline = decoded.contains('\n') || decoded.contains('\r');
    if !has_newline {
        // Be conservative for single-line inputs to avoid misclassifying ordinary prose:
        // - allow single occurrences of rarer delimiters (`;`, tab, `|`)
        // - for comma, allow one or more commas but reject the common prose pattern ", " (so
        //   "Hello, world" is not treated as CSV).
        let mut commas = 0usize;
        let mut non_comma = 0usize;
        for b in decoded.as_bytes() {
            match *b {
                b',' => commas += 1,
                b';' | b'\t' | b'|' => non_comma += 1,
                _ => {}
            }
        }
        if non_comma == 0 {
            if commas == 0 {
                return false;
            }
            if commas == 1 && decoded.contains(", ") {
                return false;
            }
        }
    }

    // Conservative: reject inputs with a high proportion of control characters (excluding common
    // whitespace used in delimited text).
    let mut control = 0usize;
    let mut total = 0usize;
    for ch in decoded.chars().take(2048) {
        total += 1;
        if ch.is_control() && !matches!(ch, '\n' | '\r' | '\t') {
            control += 1;
        }
    }
    if total > 0 && control * 100 > total {
        // >1% control chars is unlikely for a CSV/text file.
        return false;
    }

    true
}

fn sniff_text_csv<R: Read + Seek>(reader: &mut R) -> Result<bool, std::io::Error> {
    reader.seek(std::io::SeekFrom::Start(0))?;
    let mut buf = vec![0u8; TEXT_SNIFF_LEN];
    let n = reader.read(&mut buf)?;
    buf.truncate(n);
    Ok(looks_like_text_csv_prefix(&buf))
}

/// Detect the workbook format based on file signatures (and ZIP part names when needed).
///
/// This is intended to support opening workbooks even when the file extension is missing
/// or incorrect.
pub fn detect_workbook_format(path: impl AsRef<Path>) -> Result<WorkbookFormat, Error> {
    let path = path.as_ref();

    let mut file = std::fs::File::open(path).map_err(|source| Error::DetectIo {
        path: path.to_path_buf(),
        source,
    })?;

    let mut header = [0u8; 8];
    let n = file.read(&mut header).map_err(|source| Error::DetectIo {
        path: path.to_path_buf(),
        source,
    })?;
    let header = &header[..n];

    if header.len() >= OLE_MAGIC.len() && header[..OLE_MAGIC.len()] == OLE_MAGIC {
        // OLE compound files can either be legacy `.xls` BIFF workbooks, or Office-encrypted
        // OOXML packages (e.g. password-protected `.xlsx`) that wrap the real workbook in an
        // `EncryptedPackage` stream.
        //
        // We don't support decryption here; detect and return a user-friendly error so callers
        // don't try to route it through the legacy `.xls` importer.
        //
        // Decryption framing note (MS-OFFCRYPTO): `EncryptedPackage` begins with an 8-byte
        // little-endian plaintext size prefix, followed by block-aligned ciphertext that can include
        // padding beyond the declared size.
        //
        // Cipher mode differs by scheme:
        // - Standard/CryptoAPI AES: AES-ECB (no IV)
        // - Agile (4.4): AES-CBC with per-segment IV derivation
        // See `docs/offcrypto-standard-encryptedpackage.md` (Standard) and `docs/22-ooxml-encryption.md`
        // (Agile).
        file.rewind().map_err(|source| Error::DetectIo {
            path: path.to_path_buf(),
            source,
        })?;
        if let Ok(mut ole) = cfb::CompoundFile::open(file) {
            if let Some(err) = encrypted_ooxml_error(&mut ole, path, None) {
                return Err(err);
            }

            // Only treat OLE compound files as legacy `.xls` workbooks when they contain the BIFF
            // workbook stream. Other Office document types (and arbitrary OLE containers) should
            // not be misclassified as spreadsheets.
            let has_workbook_stream =
                stream_exists(&mut ole, "Workbook") || stream_exists(&mut ole, "Book");
            if !has_workbook_stream {
                return Ok(WorkbookFormat::Unknown);
            }
            // Some arbitrary OLE containers can contain a stream named `Workbook`/`Book`. Only
            // treat the file as a legacy Excel workbook when that stream actually looks like a
            // BIFF workbook stream (starts with a BOF record).
            if matches!(
                ole_workbook_stream_starts_with_biff_bof(&mut ole),
                Some(false)
            ) {
                return Ok(WorkbookFormat::Unknown);
            }

            if ole_workbook_has_biff_filepass_record(&mut ole) {
                return Err(Error::EncryptedWorkbook {
                    path: path.to_path_buf(),
                });
            }

            return Ok(WorkbookFormat::Xls);
        }

        // If we can't parse the compound file structure, fall back to the legacy `.xls`
        // classification (the downstream importer will still surface an error).
        return Ok(WorkbookFormat::Xls);
    }
    if header.len() >= PARQUET_MAGIC.len() && header[..PARQUET_MAGIC.len()] == PARQUET_MAGIC {
        return Ok(WorkbookFormat::Parquet);
    }

    // ZIP-based formats (XLSX/XLSM/XLSB) all begin with a `PK` signature.
    if header.len() >= 2 && header[..2] == *b"PK" {
        file.rewind().map_err(|source| Error::DetectIo {
            path: path.to_path_buf(),
            source,
        })?;

        let archive = zip::ZipArchive::new(file).map_err(|source| Error::DetectZip {
            path: path.to_path_buf(),
            source,
        })?;

        let mut has_workbook_bin = false;
        let mut has_workbook_xml = false;
        let mut has_vba_project = false;

        for name in archive.file_names() {
            let mut normalized = name.trim_start_matches('/');
            let replaced;
            if normalized.contains('\\') {
                replaced = normalized.replace('\\', "/");
                normalized = &replaced;
            }

            if normalized.eq_ignore_ascii_case("xl/workbook.bin") {
                has_workbook_bin = true;
            } else if normalized.eq_ignore_ascii_case("xl/workbook.xml") {
                has_workbook_xml = true;
            } else if normalized.eq_ignore_ascii_case("xl/vbaProject.bin") {
                has_vba_project = true;
            }

            if has_workbook_bin || (has_workbook_xml && has_vba_project) {
                break;
            }
        }

        if has_workbook_bin {
            return Ok(WorkbookFormat::Xlsb);
        }
        if has_workbook_xml {
            return Ok(if has_vba_project {
                WorkbookFormat::Xlsm
            } else {
                WorkbookFormat::Xlsx
            });
        }
        return Ok(WorkbookFormat::Unknown);
    }

    // Only consider CSV/text once we have ruled out known binary formats.
    if sniff_text_csv(&mut file).map_err(|source| Error::DetectIo {
        path: path.to_path_buf(),
        source,
    })? {
        return Ok(WorkbookFormat::Csv);
    }
    let ext = path
        .extension()
        .and_then(|s| s.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase();
    if ext == "csv" {
        return Ok(WorkbookFormat::Csv);
    }
    if ext == "parquet" {
        return Ok(WorkbookFormat::Parquet);
    }

    Ok(WorkbookFormat::Unknown)
}

/// Detect whether a workbook is password-protected/encrypted without attempting to open it.
///
/// This is best-effort and conservative:
/// - Returns [`WorkbookEncryption::None`] for non-OLE files, OLE files without encryption markers,
///   and malformed OLE containers that can't be parsed.
/// - Does **not** return [`Error::EncryptedWorkbook`] just because encryption is present.
pub fn detect_workbook_encryption(path: impl AsRef<Path>) -> Result<WorkbookEncryption, Error> {
    let path = path.as_ref();

    let mut file = std::fs::File::open(path).map_err(|source| Error::DetectIo {
        path: path.to_path_buf(),
        source,
    })?;

    let mut header = [0u8; 8];
    let n = file.read(&mut header).map_err(|source| Error::DetectIo {
        path: path.to_path_buf(),
        source,
    })?;
    let header = &header[..n];

    if header.len() < OLE_MAGIC.len() || header[..OLE_MAGIC.len()] != OLE_MAGIC {
        return Ok(WorkbookEncryption::None);
    }

    file.rewind().map_err(|source| Error::DetectIo {
        path: path.to_path_buf(),
        source,
    })?;

    let Ok(mut ole) = cfb::CompoundFile::open(file) else {
        // Malformed OLE container; we can't reliably sniff streams.
        return Ok(WorkbookEncryption::None);
    };

    // Office-encrypted OOXML workbooks are stored in an OLE container with both of these streams.
    // Be a little permissive and treat the presence of either stream as a signal that the workbook
    // is in the OOXML `EncryptedPackage` framing.
    let has_encryption_info = stream_exists(&mut ole, "EncryptionInfo");
    let has_encrypted_package = stream_exists(&mut ole, "EncryptedPackage");
    if has_encryption_info || has_encrypted_package {
        return Ok(WorkbookEncryption::OoxmlEncryptedPackage { scheme: None });
    }

    if let Some(scheme) = ole_workbook_filepass_scheme(&mut ole) {
        return Ok(WorkbookEncryption::LegacyXlsFilePass { scheme });
    }

    Ok(WorkbookEncryption::None)
}

/// Inspect a workbook on disk and, when it is an OLE-encrypted OOXML container, return a
/// best-effort MS-OFFCRYPTO `EncryptionInfo` summary.
///
/// This does **not** require a password and does not attempt to decrypt the workbook.
///
/// Returns:
/// - `Ok(None)` when the file is not an OLE encrypted OOXML container.
/// - `Ok(Some(summary))` when it is.
pub fn inspect_ooxml_encryption(
    path: impl AsRef<Path>,
) -> Result<Option<formula_offcrypto::EncryptionInfoSummary>, Error> {
    use std::io::Read as _;

    let path = path.as_ref();

    let mut file = std::fs::File::open(path).map_err(|source| Error::DetectIo {
        path: path.to_path_buf(),
        source,
    })?;

    // Fast-path: check the OLE magic header before attempting to parse the compound file structure.
    let mut header = [0u8; 8];
    let n = file.read(&mut header).map_err(|source| Error::DetectIo {
        path: path.to_path_buf(),
        source,
    })?;
    let header = &header[..n];
    if header.len() < OLE_MAGIC.len() || header[..OLE_MAGIC.len()] != OLE_MAGIC {
        return Ok(None);
    }

    file.rewind().map_err(|source| Error::DetectIo {
        path: path.to_path_buf(),
        source,
    })?;

    let Ok(mut ole) = cfb::CompoundFile::open(file) else {
        // Malformed OLE container; we can't reliably sniff streams.
        return Ok(None);
    };

    if !(stream_exists(&mut ole, "EncryptionInfo") && stream_exists(&mut ole, "EncryptedPackage")) {
        return Ok(None);
    }

    let encryption_info = read_ole_stream_best_effort(&mut ole, "EncryptionInfo").map_err(|source| {
        Error::DetectIo {
            path: path.to_path_buf(),
            source,
        }
    })?;
    let Some(encryption_info) = encryption_info else {
        return Ok(None);
    };

    let summary = match formula_offcrypto::inspect_encryption_info(&encryption_info) {
        Ok(summary) => summary,
        Err(formula_offcrypto::OffcryptoError::UnsupportedVersion { major, minor }) => {
            return Err(Error::UnsupportedOoxmlEncryption {
                path: path.to_path_buf(),
                version_major: major,
                version_minor: minor,
            });
        }
        Err(err) => {
            return Err(Error::DetectIo {
                path: path.to_path_buf(),
                source: std::io::Error::new(std::io::ErrorKind::InvalidData, err),
            });
        }
    };

    Ok(Some(summary))
}

fn workbook_format_impl(path: &Path, allow_encrypted_xls: bool) -> Result<WorkbookFormat, Error> {
    use std::fs::File;
    use std::io::{Read, Seek, SeekFrom};

    let ext = path
        .extension()
        .and_then(|s| s.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase();

    let ext_format = match ext.as_str() {
        // `.xltx`/`.xltm`/`.xlam` are all OOXML ZIP containers and should be treated as XLSX
        // packages for extension-based fallback dispatch.
        "xlsx" | "xltx" | "xltm" | "xlam" => Some(WorkbookFormat::Xlsx),
        "xlsm" => Some(WorkbookFormat::Xlsm),
        // `.xlt`/`.xla` are legacy BIFF8 OLE compound files, so treat them like `.xls` for fallback
        // dispatch (when sniffing can't run due to an I/O error).
        "xls" | "xlt" | "xla" => Some(WorkbookFormat::Xls),
        "xlsb" => Some(WorkbookFormat::Xlsb),
        "csv" => Some(WorkbookFormat::Csv),
        "parquet" => Some(WorkbookFormat::Parquet),
        _ => None,
    };

    // Best-effort content sniffing.
    //
    // This enables:
    // - extension-less spreadsheet files
    // - spreadsheet files with the wrong extension (common in temp-file workflows)
    //
    // If the file doesn't exist (or can't be opened), fall back to extension-based dispatch so
    // the downstream open path produces the most specific error variant (`OpenXls`, `OpenXlsb`,
    // etc).
    let mut file = match File::open(path) {
        Ok(file) => file,
        Err(source) => {
            if let Some(fmt) = ext_format {
                return Ok(fmt);
            }
            return Err(Error::OpenIo {
                path: path.to_path_buf(),
                source,
            });
        }
    };

    let mut header = [0u8; 8];
    let n = file.read(&mut header).map_err(|source| Error::OpenIo {
        path: path.to_path_buf(),
        source,
    })?;

    if n >= OLE_MAGIC.len() && header[..OLE_MAGIC.len()] == OLE_MAGIC {
        // OLE compound files can either be legacy `.xls` BIFF workbooks, or Office-encrypted
        // OOXML packages (e.g. password-protected `.xlsx`) that wrap the real workbook in an
        // `EncryptedPackage` stream.
        //
        // We don't support decryption here; detect and return a user-friendly error.
        file.seek(SeekFrom::Start(0))
            .map_err(|source| Error::OpenIo {
                path: path.to_path_buf(),
                source,
            })?;
        if let Ok(mut ole) = cfb::CompoundFile::open(file) {
            if let Some(err) = encrypted_ooxml_error(&mut ole, path, None) {
                return Err(err);
            }

            // Only treat OLE compound files as legacy `.xls` workbooks when they contain the BIFF
            // workbook stream. Other Office document types (and arbitrary OLE containers) should
            // not be misclassified as `.xls`.
            let has_workbook_stream =
                stream_exists(&mut ole, "Workbook") || stream_exists(&mut ole, "Book");
            if !has_workbook_stream {
                // Not an Excel BIFF workbook. Fall back to extension-based dispatch so callers get
                // the most specific open error (e.g. `OpenXlsx` for a `.xlsx` file that is actually
                // some other OLE container).
                if let Some(fmt) = ext_format {
                    return Ok(fmt);
                }
                return Err(Error::UnsupportedExtension {
                    path: path.to_path_buf(),
                    extension: ext,
                });
            }
            // Some arbitrary OLE containers can contain a stream named `Workbook`/`Book`. Only
            // treat the file as a legacy Excel workbook when that stream actually looks like a
            // BIFF workbook stream (starts with a BOF record).
            if matches!(
                ole_workbook_stream_starts_with_biff_bof(&mut ole),
                Some(false)
            ) {
                if let Some(fmt) = ext_format {
                    return Ok(fmt);
                }
                return Err(Error::UnsupportedExtension {
                    path: path.to_path_buf(),
                    extension: ext,
                });
            }

            if ole_workbook_has_biff_filepass_record(&mut ole) {
                // Legacy `.xls` BIFF encryption is signalled via a `FILEPASS` record in the workbook
                // globals substream. Most open paths don't support legacy decryption and should
                // surface a clear error early. However, password-capable callers may want to route
                // encrypted `.xls` files to the `.xls` importer so it can attempt decryption.
                if !allow_encrypted_xls {
                    return Err(Error::EncryptedWorkbook {
                        path: path.to_path_buf(),
                    });
                }
            }

            return Ok(WorkbookFormat::Xls);
        }
        // If we can't parse the compound file structure, fall back to the legacy `.xls`
        // classification (the downstream importer will still surface an error).
        return Ok(WorkbookFormat::Xls);
    }

    if n >= PARQUET_MAGIC.len() && header[..PARQUET_MAGIC.len()] == PARQUET_MAGIC {
        return Ok(WorkbookFormat::Parquet);
    }

    // ZIP-based formats (XLSX/XLSM/XLSB) all begin with a `PK` signature.
    if n >= 2 && header[..2] == *b"PK" {
        // Rewind so ZipArchive can read from the start.
        file.seek(SeekFrom::Start(0))
            .map_err(|source| Error::OpenIo {
                path: path.to_path_buf(),
                source,
            })?;

        if let Ok(zip) = zip::ZipArchive::new(file) {
            let mut has_workbook_xml = false;
            let mut has_workbook_bin = false;
            let mut has_vba_project = false;

            for name in zip.file_names() {
                let mut normalized = name.trim_start_matches('/');
                let replaced;
                if normalized.contains('\\') {
                    replaced = normalized.replace('\\', "/");
                    normalized = &replaced;
                }

                if normalized.eq_ignore_ascii_case("xl/workbook.xml") {
                    has_workbook_xml = true;
                } else if normalized.eq_ignore_ascii_case("xl/workbook.bin") {
                    has_workbook_bin = true;
                    break;
                } else if normalized.eq_ignore_ascii_case("xl/vbaProject.bin") {
                    has_vba_project = true;
                }

                if has_workbook_xml && has_vba_project {
                    break;
                }
            }

            if has_workbook_bin {
                return Ok(WorkbookFormat::Xlsb);
            }
            if has_workbook_xml {
                return Ok(if has_vba_project {
                    WorkbookFormat::Xlsm
                } else {
                    WorkbookFormat::Xlsx
                });
            }
        }

        // ZIP signatures must win even if we can't classify the archive: fall back to
        // extension-based dispatch (or an unsupported-extension error) rather than treating the
        // input as text/CSV.
        return match ext_format {
            Some(fmt) => Ok(fmt),
            None => Err(Error::UnsupportedExtension {
                path: path.to_path_buf(),
                extension: ext,
            }),
        };
    }

    // Only consider CSV/text once we have ruled out known binary formats.
    if sniff_text_csv(&mut file).map_err(|source| Error::OpenIo {
        path: path.to_path_buf(),
        source,
    })? {
        return Ok(WorkbookFormat::Csv);
    }

    match ext_format {
        Some(fmt) => Ok(fmt),
        None => Err(Error::UnsupportedExtension {
            path: path.to_path_buf(),
            extension: ext,
        }),
    }
}

fn workbook_format(path: &Path) -> Result<WorkbookFormat, Error> {
    workbook_format_impl(path, false)
}

fn workbook_format_allow_encrypted_xls(path: &Path) -> Result<WorkbookFormat, Error> {
    workbook_format_impl(path, true)
}

/// Open a spreadsheet workbook from disk directly into a [`formula_model::Workbook`] with options.
///
/// This is the password-aware variant of [`open_workbook_model`]. When a password is provided and
/// the input is a legacy `.xls` workbook using BIFF `FILEPASS` encryption, this will attempt to
/// decrypt and import the workbook.
pub fn open_workbook_model_with_options(
    path: impl AsRef<Path>,
    opts: OpenOptions,
) -> Result<formula_model::Workbook, Error> {
    use std::fs::File;
    use std::io::BufReader;

    let path = path.as_ref();

    // First, handle password-protected OOXML workbooks that are stored as OLE compound files
    // (`EncryptionInfo` + `EncryptedPackage` streams).
    //
    // When the `encrypted-workbooks` feature is enabled, attempt in-memory decryption. Otherwise,
    // surface an "unsupported encryption" error so callers don't assume a password will work.
    #[cfg(feature = "encrypted-workbooks")]
    {
        let is_xlsb = path
            .extension()
            .and_then(|s| s.to_str())
            .is_some_and(|ext| ext.eq_ignore_ascii_case("xlsb"));
        if opts.password.is_some() && !is_xlsb {
            if let Some(workbook) =
                try_open_standard_aes_encrypted_ooxml_model_workbook(path, opts.password.as_deref())?
            {
                return Ok(workbook);
            }
        }
        if let Some(bytes) =
            try_decrypt_ooxml_encrypted_package_from_path(path, opts.password.as_deref())?
        {
            return open_workbook_model_from_decrypted_ooxml_zip_bytes(path, bytes);
        }
    }

    if let Some(package_bytes) =
        maybe_read_plaintext_ooxml_package_from_encrypted_ole(path, opts.password.as_deref())?
    {
        return open_workbook_model_from_decrypted_ooxml_zip_bytes(path, package_bytes);
    }
    let ext = path
        .extension()
        .and_then(|s| s.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase();

    let format = if opts.password.is_some() {
        workbook_format_allow_encrypted_xls(path)?
    } else {
        match workbook_format(path) {
            Ok(fmt) => fmt,
            Err(Error::EncryptedWorkbook { .. }) => {
                return Err(Error::PasswordRequired {
                    path: path.to_path_buf(),
                });
            }
            Err(err) => return Err(err),
        }
    };

    match format {
        WorkbookFormat::Xlsx | WorkbookFormat::Xlsm => {
            let file = File::open(path).map_err(|source| Error::OpenIo {
                path: path.to_path_buf(),
                source,
            })?;
            xlsx::read_workbook_from_reader(file).map_err(|source| Error::OpenXlsx {
                path: path.to_path_buf(),
                source,
            })
        }
        WorkbookFormat::Xls => {
            match xls::import_xls_path_with_password(path, opts.password.as_deref()) {
                Ok(result) => Ok(result.workbook),
                Err(xls::ImportError::EncryptedWorkbook) => Err(Error::PasswordRequired {
                    path: path.to_path_buf(),
                }),
                Err(xls::ImportError::InvalidPassword) => Err(Error::InvalidPassword {
                    path: path.to_path_buf(),
                }),
                Err(xls::ImportError::UnsupportedEncryption(scheme)) => {
                    Err(Error::UnsupportedEncryption {
                        path: path.to_path_buf(),
                        kind: scheme,
                    })
                }
                Err(xls::ImportError::Decrypt(message)) => Err(Error::UnsupportedEncryption {
                    path: path.to_path_buf(),
                    kind: format!(
                        "legacy `.xls` FILEPASS encryption metadata is invalid: {message}"
                    ),
                }),
                Err(source) => Err(Error::OpenXls {
                    path: path.to_path_buf(),
                    source,
                }),
            }
        }
        WorkbookFormat::Xlsb => {
            let wb = xlsb::XlsbWorkbook::open_with_options(
                path,
                xlsb::OpenOptions {
                    preserve_unknown_parts: false,
                    preserve_parsed_parts: false,
                    preserve_worksheets: false,
                    decode_formulas: true,
                    ..Default::default()
                },
            )
            .map_err(|source| Error::OpenXlsb {
                path: path.to_path_buf(),
                source,
            })?;
            xlsb_to_model_workbook(&wb).map_err(|source| Error::OpenXlsb {
                path: path.to_path_buf(),
                source,
            })
        }
        WorkbookFormat::Csv => {
            let file = File::open(path).map_err(|source| Error::OpenIo {
                path: path.to_path_buf(),
                source,
            })?;
            let reader = BufReader::new(file);

            let stem = path.file_stem().and_then(|s| s.to_str()).unwrap_or("");
            let sheet_name = formula_model::validate_sheet_name(stem)
                .ok()
                .map(|_| stem.to_string())
                .unwrap_or_else(|| sanitize_sheet_name(stem));

            let mut workbook = formula_model::Workbook::new();
            import_csv_into_workbook(
                &mut workbook,
                sheet_name,
                reader,
                CsvOptions::default(),
            )
            .map_err(|source| Error::OpenCsv {
                path: path.to_path_buf(),
                source,
            })?;

            Ok(workbook)
        }
        WorkbookFormat::Parquet => open_parquet_model_workbook(path),
        _ => Err(Error::UnsupportedExtension {
            path: path.to_path_buf(),
            extension: ext.to_string(),
        }),
    }
}

/// Open a spreadsheet workbook from disk directly into a [`formula_model::Workbook`].
///
/// This is a faster, lower-memory alternative to [`open_workbook`] for read-only/import
/// workflows:
/// - For `.xlsx`/`.xlsm`/`.xltx`/`.xltm`/`.xlam`, this uses the streaming reader in `formula-xlsx`
///   and avoids inflating the entire OPC package into memory.
/// - For `.xls`/`.xlt`/`.xla`, this returns the imported model workbook from `formula-xls`.
/// - For `.xlsb`, this converts the parsed workbook into a model workbook.
/// - For `.csv`, this imports the CSV into a columnar-backed worksheet.
/// - For `.parquet`, this imports the Parquet file into a columnar-backed worksheet (requires the
///   `formula-io` crate feature `parquet`).
pub fn open_workbook_model(path: impl AsRef<Path>) -> Result<formula_model::Workbook, Error> {
    use std::fs::File;
    use std::io::BufReader;

    let path = path.as_ref();
    let ext = path
        .extension()
        .and_then(|s| s.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase();

    match workbook_format(path)? {
        WorkbookFormat::Xlsx | WorkbookFormat::Xlsm => {
            let file = File::open(path).map_err(|source| Error::OpenIo {
                path: path.to_path_buf(),
                source,
            })?;
            xlsx::read_workbook_from_reader(file).map_err(|source| Error::OpenXlsx {
                path: path.to_path_buf(),
                source,
            })
        }
        WorkbookFormat::Xls => match xls::import_xls_path(path) {
            Ok(result) => Ok(result.workbook),
            Err(xls::ImportError::EncryptedWorkbook) => Err(Error::EncryptedWorkbook {
                path: path.to_path_buf(),
            }),
            Err(source) => Err(Error::OpenXls {
                path: path.to_path_buf(),
                source,
            }),
        },
        WorkbookFormat::Xlsb => {
            let wb = xlsb::XlsbWorkbook::open_with_options(
                path,
                xlsb::OpenOptions {
                    preserve_unknown_parts: false,
                    preserve_parsed_parts: false,
                    preserve_worksheets: false,
                    decode_formulas: true,
                    ..Default::default()
                },
            )
            .map_err(|source| Error::OpenXlsb {
                path: path.to_path_buf(),
                source,
            })?;
            xlsb_to_model_workbook(&wb).map_err(|source| Error::OpenXlsb {
                path: path.to_path_buf(),
                source,
            })
        }
        WorkbookFormat::Csv => {
            let file = File::open(path).map_err(|source| Error::OpenIo {
                path: path.to_path_buf(),
                source,
            })?;
            let reader = BufReader::new(file);

            let stem = path.file_stem().and_then(|s| s.to_str()).unwrap_or("");
            let sheet_name = formula_model::validate_sheet_name(stem)
                .ok()
                .map(|_| stem.to_string())
                .unwrap_or_else(|| sanitize_sheet_name(stem));

            let mut workbook = formula_model::Workbook::new();
            import_csv_into_workbook(&mut workbook, sheet_name, reader, CsvOptions::default())
                .map_err(|source| Error::OpenCsv {
                    path: path.to_path_buf(),
                    source,
                })?;

            Ok(workbook)
        }
        WorkbookFormat::Parquet => open_parquet_model_workbook(path),
        _ => Err(Error::UnsupportedExtension {
            path: path.to_path_buf(),
            extension: ext.to_string(),
        }),
    }
}

fn open_workbook_from_decrypted_ooxml_zip_bytes(
    path: &Path,
    decrypted_bytes: Vec<u8>,
) -> Result<Workbook, Error> {
    match sniff_ooxml_zip_workbook_kind(&decrypted_bytes) {
        Some(WorkbookFormat::Xlsb) => {
            let wb = xlsb::XlsbWorkbook::open_from_vec(decrypted_bytes).map_err(|source| {
                Error::OpenXlsb {
                    path: path.to_path_buf(),
                    source,
                }
            })?;
            Ok(Workbook::Xlsb(wb))
        }
        _ => {
            let package = xlsx::XlsxLazyPackage::from_vec(decrypted_bytes).map_err(|source| {
                Error::OpenXlsx {
                    path: path.to_path_buf(),
                    source,
                }
            })?;
            Ok(Workbook::Xlsx(package))
        }
    }
}

fn open_workbook_model_from_decrypted_ooxml_zip_bytes(
    path: &Path,
    decrypted_bytes: Vec<u8>,
) -> Result<formula_model::Workbook, Error> {
    if zip_contains_workbook_bin(&decrypted_bytes) {
        let wb = xlsb::XlsbWorkbook::open_from_vec_with_options(
            decrypted_bytes,
            xlsb::OpenOptions {
                preserve_unknown_parts: false,
                preserve_parsed_parts: false,
                preserve_worksheets: false,
                decode_formulas: true,
                ..Default::default()
            },
        )
        .map_err(|source| Error::OpenXlsb {
            path: path.to_path_buf(),
            source,
        })?;
        xlsb_to_model_workbook(&wb).map_err(|source| Error::OpenXlsb {
            path: path.to_path_buf(),
            source,
        })
    } else {
        xlsx::read_workbook_from_reader(std::io::Cursor::new(decrypted_bytes)).map_err(
            |source| Error::OpenXlsx {
                path: path.to_path_buf(),
                source,
            },
        )
    }
}

/// Open a spreadsheet workbook from disk directly into a [`formula_model::Workbook`], optionally
/// providing a password for encrypted workbooks.
///
/// - For password-protected legacy `.xls` workbooks (BIFF `FILEPASS`), this will attempt to decrypt
///   the workbook stream using `formula-xls` when `password` is provided.
/// - For Office-encrypted OOXML workbooks stored in an OLE container (`EncryptionInfo` +
///   `EncryptedPackage`):
///   - when the `formula-io/encrypted-workbooks` feature is enabled, this will attempt to decrypt
///     the workbook in-memory (and may return [`Error::PasswordRequired`] / [`Error::InvalidPassword`]
///     on failure).
///   - otherwise, this returns [`Error::UnsupportedEncryption`].
///
/// With the `formula-io/encrypted-workbooks` feature enabled, this function will attempt to decrypt
/// and open supported encrypted OOXML workbooks in memory (without persisting plaintext to disk).
pub fn open_workbook_model_with_password(
    path: impl AsRef<Path>,
    password: Option<&str>,
) -> Result<formula_model::Workbook, Error> {
    let path = path.as_ref();
    // Handle the special-case where an `EncryptedPackage` stream already contains a plaintext ZIP
    // payload (e.g. synthetic fixtures or already-decrypted pipelines). This does not require
    // decryption support, and it must run *before* attempting decryption so we don't misclassify a
    // plaintext payload as an "invalid password" error.
    if password.is_some() {
        if let Some(bytes) = maybe_read_plaintext_ooxml_package_from_encrypted_ole_if_plaintext(path)?
        {
            return open_workbook_model_from_decrypted_ooxml_zip_bytes(path, bytes);
        }
    }

    // Attempt to decrypt Office-encrypted OOXML workbooks (OLE container with `EncryptionInfo` +
    // `EncryptedPackage`) when the feature is enabled.
    #[cfg(feature = "encrypted-workbooks")]
    {
        let is_xlsb = path
            .extension()
            .and_then(|s| s.to_str())
            .is_some_and(|ext| ext.eq_ignore_ascii_case("xlsb"));
        if password.is_some() && !is_xlsb {
            if let Some(workbook) = try_open_standard_aes_encrypted_ooxml_model_workbook(path, password)?
            {
                return Ok(workbook);
            }
        }

        if let Some(bytes) = try_decrypt_ooxml_encrypted_package_from_path(path, password)? {
            return open_workbook_model_from_decrypted_ooxml_zip_bytes(path, bytes);
        }
    }

    if let Some(err) = encrypted_ooxml_error_from_path(path, password) {
        return Err(err);
    }

    // If no password was provided, preserve the existing open path for non-encrypted files.
    //
    // For legacy `.xls` BIFF encryption (FILEPASS), provide a more actionable error: callers using
    // the password-capable API likely want to prompt the user for a password rather than being told
    // to remove encryption.
    let Some(password) = password else {
        if let Ok(encryption) = detect_workbook_encryption(path) {
            if matches!(encryption, WorkbookEncryption::LegacyXlsFilePass { .. }) {
                return Err(Error::PasswordRequired {
                    path: path.to_path_buf(),
                });
            }
        }
        return open_workbook_model(path);
    };
    // For non-OOXML formats (including legacy `.xls` FILEPASS), reuse the options-based open path.
    open_workbook_model_with_options(
        path,
        OpenOptions {
            password: Some(password.to_string()),
            ..Default::default()
        },
    )
}

/// Open a spreadsheet workbook from disk, optionally providing a password for encrypted
/// workbooks.
///
/// - For password-protected legacy `.xls` workbooks (BIFF `FILEPASS`), this will attempt to decrypt
///   the workbook stream using `formula-xls` when `password` is provided.
/// - For Office-encrypted OOXML workbooks stored in an OLE container (`EncryptionInfo` +
///   `EncryptedPackage`):
///   - when the `formula-io/encrypted-workbooks` feature is enabled, this will attempt to decrypt
///     the workbook in-memory (and may return [`Error::PasswordRequired`] / [`Error::InvalidPassword`]
///     on failure).
///   - otherwise, this returns [`Error::UnsupportedEncryption`].
///
/// With the `formula-io/encrypted-workbooks` feature enabled, this function will attempt to decrypt
/// and open supported encrypted OOXML workbooks in memory (without persisting plaintext to disk).
pub fn open_workbook_with_password(
    path: impl AsRef<Path>,
    password: Option<&str>,
) -> Result<Workbook, Error> {
    let path = path.as_ref();
    // Handle the special-case where an `EncryptedPackage` stream already contains a plaintext ZIP
    // payload (e.g. synthetic fixtures or already-decrypted pipelines). This does not require
    // decryption support, and it must run *before* attempting decryption so we don't misclassify a
    // plaintext payload as an "invalid password" error.
    if password.is_some() {
        if let Some(bytes) = maybe_read_plaintext_ooxml_package_from_encrypted_ole_if_plaintext(path)?
        {
            return open_workbook_from_decrypted_ooxml_zip_bytes(path, bytes);
        }
    }

    // Attempt to decrypt Office-encrypted OOXML workbooks (OLE container with `EncryptionInfo` +
    // `EncryptedPackage`) when the feature is enabled.
    #[cfg(feature = "encrypted-workbooks")]
    if let Some(bytes) = try_decrypt_ooxml_encrypted_package_from_path(path, password)? {
        return open_workbook_from_decrypted_ooxml_zip_bytes(path, bytes);
    }

    if let Some(err) = encrypted_ooxml_error_from_path(path, password) {
        return Err(err);
    }

    // Delegate to the options-based password open path (this handles legacy `.xls` FILEPASS).
    open_workbook_with_options(
        path,
        OpenOptions {
            password: password.map(ToString::to_string),
            ..Default::default()
        },
    )
}
/// A workbook opened from disk with optional preserved OLE metadata streams.
///
/// When an Office-encrypted OOXML workbook (OLE/CFB wrapper with `EncryptionInfo` +
/// `EncryptedPackage`) is opened with password support enabled (`formula-io/encrypted-workbooks`),
/// Formula decrypts the underlying ZIP package into memory and additionally captures any other OLE
/// streams/storages (e.g. `\u{0005}SummaryInformation`) so they can be re-emitted when saving back
/// as an encrypted workbook.
#[cfg(feature = "encrypted-workbooks")]
#[derive(Debug)]
pub struct OpenedWorkbookWithPreservedOle {
    pub workbook: Workbook,
    pub preserved_ole: Option<formula_office_crypto::OleEntries>,
}

#[cfg(feature = "encrypted-workbooks")]
impl OpenedWorkbookWithPreservedOle {
    /// Save the workbook, preserving Office encryption when the input was an encrypted OOXML OLE
    /// wrapper.
    ///
    /// - For non-encrypted inputs, this falls back to [`save_workbook`].
    /// - For encrypted OOXML inputs, this writes an OLE/CFB `EncryptionInfo` + `EncryptedPackage`
    ///   wrapper and copies preserved non-encryption OLE streams byte-for-byte.
    ///
    /// Note: the password is required at save time because this type does not store it.
    pub fn save_preserving_encryption(
        &self,
        path: impl AsRef<Path>,
        password: &str,
    ) -> Result<(), Error> {
        let path = path.as_ref();
        if let Some(preserved) = &self.preserved_ole {
            save_workbook_encrypted_ooxml(&self.workbook, path, password, preserved)
        } else {
            save_workbook(&self.workbook, path)
        }
    }
}

/// Open a workbook and, when it is an Office-encrypted OOXML OLE container, also preserve any
/// additional non-encryption OLE streams/storages for round-trip.
#[cfg(feature = "encrypted-workbooks")]
pub fn open_workbook_with_password_and_preserved_ole(
    path: impl AsRef<Path>,
    password: Option<&str>,
) -> Result<OpenedWorkbookWithPreservedOle, Error> {
    let path = path.as_ref();

    if let Some((bytes, preserved)) =
        try_decrypt_ooxml_encrypted_package_from_path_with_preserved_ole(path, password)?
    {
        let workbook = open_workbook_from_decrypted_ooxml_zip_bytes(path, bytes)?;
        return Ok(OpenedWorkbookWithPreservedOle {
            workbook,
            preserved_ole: Some(preserved),
        });
    }

    let workbook = open_workbook_with_password(path, password)?;
    Ok(OpenedWorkbookWithPreservedOle {
        workbook,
        preserved_ole: None,
    })
}
fn encrypted_ooxml_error_from_path(path: &Path, password: Option<&str>) -> Option<Error> {
    use std::io::{Read as _, Seek as _, SeekFrom};

    let mut file = std::fs::File::open(path).ok()?;
    let mut header = [0u8; 8];
    let n = file.read(&mut header).ok()?;
    if n < OLE_MAGIC.len() || header[..OLE_MAGIC.len()] != OLE_MAGIC {
        return None;
    }
    file.seek(SeekFrom::Start(0)).ok()?;
    let mut ole = cfb::CompoundFile::open(file).ok()?;
    encrypted_ooxml_error(&mut ole, path, password)
}

fn stream_exists<R: std::io::Read + std::io::Write + std::io::Seek>(
    ole: &mut cfb::CompoundFile<R>,
    name: &str,
) -> bool {
    open_stream_best_effort(ole, name).is_some()
}

fn open_stream_best_effort<R: std::io::Read + std::io::Write + std::io::Seek>(
    ole: &mut cfb::CompoundFile<R>,
    name: &str,
) -> Option<cfb::Stream<R>> {
    if let Ok(stream) = ole.open_stream(name) {
        return Some(stream);
    }

    let trimmed = name.trim_start_matches('/');
    if trimmed != name {
        if let Ok(stream) = ole.open_stream(trimmed) {
            return Some(stream);
        }
    }

    let with_leading_slash = format!("/{trimmed}");
    if let Ok(stream) = ole.open_stream(&with_leading_slash) {
        return Some(stream);
    }

    // Some real-world producers vary casing for the `EncryptionInfo`/`EncryptedPackage` streams (and
    // some `cfb` implementations appear to treat `open_stream` as case-sensitive). Walk the
    // directory tree and locate a matching entry case-insensitively, then open the *exact*
    // discovered path so downstream reads are deterministic.
    let mut found_path: Option<String> = None;
    for entry in ole.walk() {
        if !entry.is_stream() {
            continue;
        }
        let path = entry.path().to_string_lossy();
        let normalized = path.as_ref().strip_prefix('/').unwrap_or(path.as_ref());
        if normalized.eq_ignore_ascii_case(trimmed) {
            found_path = Some(path.into_owned());
            break;
        }
    }

    let found_path = found_path?;
    if let Ok(stream) = ole.open_stream(&found_path) {
        return Some(stream);
    }

    // Be defensive: some implementations accept the walk()-returned path but reject a leading slash
    // (or vice versa).
    let stripped = found_path.strip_prefix('/').unwrap_or(found_path.as_str());
    if stripped != found_path {
        if let Ok(stream) = ole.open_stream(stripped) {
            return Some(stream);
        }
        let with_slash = format!("/{stripped}");
        if let Ok(stream) = ole.open_stream(&with_slash) {
            return Some(stream);
        }
    }

    None
}

fn read_ole_stream_best_effort<R: std::io::Read + std::io::Write + std::io::Seek>(
    ole: &mut cfb::CompoundFile<R>,
    name: &str,
) -> Result<Option<Vec<u8>>, std::io::Error> {
    use std::io::Read as _;

    fn read_candidate<R: std::io::Read + std::io::Write + std::io::Seek>(
        ole: &mut cfb::CompoundFile<R>,
        path: &str,
    ) -> Result<Option<Vec<u8>>, std::io::Error> {
        let mut stream = match ole.open_stream(path) {
            Ok(s) => s,
            Err(err) if err.kind() == std::io::ErrorKind::NotFound => return Ok(None),
            Err(err) => return Err(err),
        };
        let mut buf = Vec::new();
        stream.read_to_end(&mut buf)?;
        Ok(Some(buf))
    }

    if let Some(buf) = read_candidate(ole, name)? {
        return Ok(Some(buf));
    }
    let target = name.trim_start_matches('/');
    if target != name {
        if let Some(buf) = read_candidate(ole, target)? {
            return Ok(Some(buf));
        }
    }
    let with_leading_slash = format!("/{target}");
    if let Some(buf) = read_candidate(ole, &with_leading_slash)? {
        return Ok(Some(buf));
    }

    let candidate = {
        ole.walk().find_map(|entry| {
            if !entry.is_stream() {
                return None;
            }
            let path = entry.path().to_string_lossy().into_owned();
            let normalized = path.strip_prefix('/').unwrap_or(&path);
            if normalized.eq_ignore_ascii_case(target) {
                Some(path)
            } else {
                None
            }
        })
    };

    if let Some(path) = candidate {
        if let Some(buf) = read_candidate(ole, &path)? {
            return Ok(Some(buf));
        }
        // Some callers use paths without a leading slash; try again stripped.
        let stripped = path.strip_prefix('/').unwrap_or(&path);
        if stripped != path {
            if let Some(buf) = read_candidate(ole, stripped)? {
                return Ok(Some(buf));
            }
        }
    }

    Ok(None)
}

#[cfg(feature = "encrypted-workbooks")]
fn open_stream_case_insensitive<R: std::io::Read + std::io::Write + std::io::Seek>(
    ole: &mut cfb::CompoundFile<R>,
    name: &str,
) -> std::io::Result<cfb::Stream<R>> {
    open_stream_best_effort(ole, name).ok_or_else(|| {
        std::io::Error::new(
            std::io::ErrorKind::NotFound,
            format!("stream not found: {name}"),
        )
    })
}

#[cfg(feature = "encrypted-workbooks")]
fn read_stream_bytes_case_insensitive<R: std::io::Read + std::io::Write + std::io::Seek>(
    ole: &mut cfb::CompoundFile<R>,
    name: &str,
) -> std::io::Result<Vec<u8>> {
    use std::io::Read as _;

    let mut stream = open_stream_case_insensitive(ole, name)?;
    let mut buf = Vec::new();
    stream.read_to_end(&mut buf)?;
    Ok(buf)
}

#[cfg(feature = "encrypted-workbooks")]
fn decode_utf16le_z_lossy(bytes: &[u8]) -> Result<String, formula_offcrypto::OffcryptoError> {
    if bytes.is_empty() {
        return Ok(String::new());
    }
    if bytes.len() % 2 != 0 {
        return Err(formula_offcrypto::OffcryptoError::InvalidCspNameUtf16);
    }

    let mut units: Vec<u16> = Vec::with_capacity(bytes.len() / 2);
    for chunk in bytes.chunks_exact(2) {
        units.push(u16::from_le_bytes([chunk[0], chunk[1]]));
    }
    let end = units.iter().position(|u| *u == 0).unwrap_or(units.len());
    String::from_utf16(&units[..end]).map_err(|_| formula_offcrypto::OffcryptoError::InvalidCspNameUtf16)
}

/// Parse a Standard (CryptoAPI) `EncryptionInfo` stream while being tolerant of missing/incorrect
/// header flags.
///
/// Some real-world producers omit `fCryptoAPI` / `fAES` even though the rest of the header matches
/// the Standard/CryptoAPI schema. `formula-offcrypto` intentionally rejects those for AES to avoid
/// false positives when parsing arbitrary bytes, but for Formula's workbook open path we already
/// know we're inside an OOXML-encrypted OLE container (`EncryptionInfo` + `EncryptedPackage`
/// streams), so we can safely fall back to a lenient parser.
#[cfg(feature = "encrypted-workbooks")]
fn parse_standard_encryption_info_lenient(
    encryption_info: &[u8],
) -> Result<formula_offcrypto::StandardEncryptionInfo, formula_offcrypto::OffcryptoError> {
    fn read_u16_le(
        bytes: &[u8],
        pos: &mut usize,
        context: &'static str,
    ) -> Result<u16, formula_offcrypto::OffcryptoError> {
        let end = pos.saturating_add(2);
        let slice = bytes
            .get(*pos..end)
            .ok_or(formula_offcrypto::OffcryptoError::Truncated { context })?;
        *pos = end;
        Ok(u16::from_le_bytes([slice[0], slice[1]]))
    }

    fn read_u32_le(
        bytes: &[u8],
        pos: &mut usize,
        context: &'static str,
    ) -> Result<u32, formula_offcrypto::OffcryptoError> {
        let end = pos.saturating_add(4);
        let slice = bytes
            .get(*pos..end)
            .ok_or(formula_offcrypto::OffcryptoError::Truncated { context })?;
        *pos = end;
        Ok(u32::from_le_bytes([slice[0], slice[1], slice[2], slice[3]]))
    }

    let mut pos = 0usize;
    let major = read_u16_le(encryption_info, &mut pos, "EncryptionVersionInfo.major")?;
    let minor = read_u16_le(encryption_info, &mut pos, "EncryptionVersionInfo.minor")?;
    let _version_flags = read_u32_le(encryption_info, &mut pos, "EncryptionVersionInfo.flags")?;

    // Standard encryption uses `versionMinor == 2` with major typically 2/3/4.
    if minor != 2 || !matches!(major, 2 | 3 | 4) {
        return Err(formula_offcrypto::OffcryptoError::UnsupportedVersion { major, minor });
    }

    let header_size = read_u32_le(encryption_info, &mut pos, "EncryptionInfo.header_size")? as usize;
    const MIN_STANDARD_HEADER_SIZE: usize = 8 * 4;
    const MAX_STANDARD_HEADER_SIZE: usize = 1024 * 1024;
    if header_size < MIN_STANDARD_HEADER_SIZE || header_size > MAX_STANDARD_HEADER_SIZE {
        return Err(formula_offcrypto::OffcryptoError::InvalidEncryptionInfo {
            context: "EncryptionInfo.header_size is out of bounds",
        });
    }

    let header_bytes = encryption_info
        .get(pos..pos + header_size)
        .ok_or(formula_offcrypto::OffcryptoError::Truncated {
            context: "EncryptionHeader",
        })?;
    pos += header_size;

    let mut hpos = 0usize;
    let raw_flags = read_u32_le(header_bytes, &mut hpos, "EncryptionHeader.flags")?;
    let flags = formula_offcrypto::StandardEncryptionHeaderFlags::from_raw(raw_flags);
    if flags.f_external {
        return Err(formula_offcrypto::OffcryptoError::UnsupportedExternalEncryption);
    }

    let header = formula_offcrypto::StandardEncryptionHeader {
        flags,
        size_extra: read_u32_le(header_bytes, &mut hpos, "EncryptionHeader.sizeExtra")?,
        alg_id: read_u32_le(header_bytes, &mut hpos, "EncryptionHeader.algId")?,
        alg_id_hash: read_u32_le(header_bytes, &mut hpos, "EncryptionHeader.algIdHash")?,
        key_size_bits: read_u32_le(header_bytes, &mut hpos, "EncryptionHeader.keySize")?,
        provider_type: read_u32_le(header_bytes, &mut hpos, "EncryptionHeader.providerType")?,
        reserved1: read_u32_le(header_bytes, &mut hpos, "EncryptionHeader.reserved1")?,
        reserved2: read_u32_le(header_bytes, &mut hpos, "EncryptionHeader.reserved2")?,
        csp_name: decode_utf16le_z_lossy(&header_bytes[hpos..])?,
    };

    let salt_size = read_u32_le(encryption_info, &mut pos, "EncryptionVerifier.saltSize")? as usize;
    let salt = encryption_info
        .get(pos..pos + salt_size)
        .ok_or(formula_offcrypto::OffcryptoError::Truncated {
            context: "EncryptionVerifier.salt",
        })?
        .to_vec();
    pos += salt_size;

    let encrypted_verifier_bytes = encryption_info
        .get(pos..pos + 16)
        .ok_or(formula_offcrypto::OffcryptoError::Truncated {
            context: "EncryptionVerifier.encryptedVerifier",
        })?;
    let mut encrypted_verifier = [0u8; 16];
    encrypted_verifier.copy_from_slice(encrypted_verifier_bytes);
    pos += 16;

    let verifier_hash_size =
        read_u32_le(encryption_info, &mut pos, "EncryptionVerifier.verifierHashSize")?;
    let encrypted_verifier_hash = encryption_info.get(pos..).unwrap_or_default().to_vec();

    Ok(formula_offcrypto::StandardEncryptionInfo {
        header,
        verifier: formula_offcrypto::StandardEncryptionVerifier {
            salt,
            encrypted_verifier,
            verifier_hash_size,
            encrypted_verifier_hash,
        },
    })
}

#[cfg(feature = "encrypted-workbooks")]
fn try_decrypt_ooxml_encrypted_package_from_path(
    path: &Path,
    password: Option<&str>,
) -> Result<Option<Vec<u8>>, Error> {
    use std::io::{Read as _, Seek as _};

    let mut file = std::fs::File::open(path).map_err(|source| Error::OpenIo {
        path: path.to_path_buf(),
        source,
    })?;

    let mut header = [0u8; 8];
    let n = file.read(&mut header).map_err(|source| Error::OpenIo {
        path: path.to_path_buf(),
        source,
    })?;
    if n < OLE_MAGIC.len() || header[..OLE_MAGIC.len()] != OLE_MAGIC {
        return Ok(None);
    }

    file.rewind().map_err(|source| Error::OpenIo {
        path: path.to_path_buf(),
        source,
    })?;
    let Ok(mut ole) = cfb::CompoundFile::open(file) else {
        // Malformed OLE container; fall back to non-encrypted open paths.
        return Ok(None);
    };

    // Attempt to read both required streams (best-effort + case-insensitive lookup).
    let encryption_info = match read_stream_bytes_case_insensitive(&mut ole, "EncryptionInfo") {
        Ok(bytes) => bytes,
        Err(err) if err.kind() == std::io::ErrorKind::NotFound => return Ok(None),
        Err(_) => {
            return Err(Error::UnsupportedOoxmlEncryption {
                path: path.to_path_buf(),
                version_major: 0,
                version_minor: 0,
            })
        }
    };
    let encrypted_package = match read_stream_bytes_case_insensitive(&mut ole, "EncryptedPackage") {
        Ok(bytes) => bytes,
        Err(err) if err.kind() == std::io::ErrorKind::NotFound => return Ok(None),
        Err(_) => {
            return Err(Error::UnsupportedOoxmlEncryption {
                path: path.to_path_buf(),
                version_major: 0,
                version_minor: 0,
            })
        }
    };

    // Some synthetic fixtures (and some pipelines) may already contain a plaintext ZIP payload in
    // `EncryptedPackage`. Let callers handle that via the plaintext open path so this helper only
    // yields bytes when we actually decrypted an encrypted payload.
    if maybe_extract_ooxml_package_bytes(&encrypted_package).is_some() {
        return Ok(None);
    }

    if encryption_info.len() < 4 {
        return Err(Error::UnsupportedOoxmlEncryption {
            path: path.to_path_buf(),
            version_major: 0,
            version_minor: 0,
        });
    }
    let version_major = u16::from_le_bytes([encryption_info[0], encryption_info[1]]);
    let version_minor = u16::from_le_bytes([encryption_info[2], encryption_info[3]]);

    // Decryption support is limited to the common modern schemes:
    // - Agile encryption (4.4)
    // - Standard/CryptoAPI encryption (`versionMinor == 2`; commonly 3.2, but 2.2/4.2 are observed)
    //
    // Fail early on other versions so callers get a precise error even if a password is missing.
    let supported = (version_major == 4 && version_minor == 4)
        || (version_minor == 2 && matches!(version_major, 2 | 3 | 4));
    if !supported {
        return Err(Error::UnsupportedOoxmlEncryption {
            path: path.to_path_buf(),
            version_major,
            version_minor,
        });
    }
    let Some(password) = password else {
        return Err(Error::PasswordRequired {
            path: path.to_path_buf(),
        });
    };

    // `EncryptedPackage` streams should start with an 8-byte plaintext length header.
    //
    // If the stream is too short (and we didn't already detect a plaintext ZIP payload above),
    // treat it as a malformed/unsupported encryption container rather than an invalid password.
    // A wrong password should still surface as `InvalidPassword` once we can actually attempt a
    // verifier/integrity check.
    if encrypted_package.len() <= 8 {
        return Err(Error::UnsupportedOoxmlEncryption {
            path: path.to_path_buf(),
            version_major,
            version_minor,
        });
    }
    let decrypted = if (version_major, version_minor) == (4, 4) {
        // Agile (4.4): prefer the strict decryptor in `formula-xlsx` because it validates
        // `dataIntegrity` when present. Some producers omit `dataIntegrity`; fall back to a more
        // tolerant decryption path in that case.
        match xlsx::offcrypto::decrypt_ooxml_encrypted_package(
            &encryption_info,
            &encrypted_package,
            password,
        ) {
            Ok(bytes) => bytes,
            Err(err) => match err {
                xlsx::OffCryptoError::WrongPassword | xlsx::OffCryptoError::IntegrityMismatch => {
                    return Err(Error::InvalidPassword {
                        path: path.to_path_buf(),
                    })
                }
                xlsx::OffCryptoError::UnsupportedEncryptionVersion { major, minor } => {
                    return Err(Error::UnsupportedOoxmlEncryption {
                        path: path.to_path_buf(),
                        version_major: major,
                        version_minor: minor,
                    })
                }
                xlsx::OffCryptoError::MissingRequiredElement { ref element }
                    if element.eq_ignore_ascii_case("dataIntegrity") =>
                {
                    // The `EncryptedPackage` stream starts with an 8-byte plaintext length prefix.
                    if encrypted_package.len() < 8 {
                        return Err(Error::UnsupportedOoxmlEncryption {
                            path: path.to_path_buf(),
                            version_major,
                            version_minor,
                        });
                    }
                    let mut len_bytes = [0u8; 8];
                    len_bytes.copy_from_slice(&encrypted_package[..8]);
                    let ciphertext_len = encrypted_package.len().saturating_sub(8) as u64;
                    let plaintext_len =
                        parse_encrypted_package_size_prefix_bytes(len_bytes, Some(ciphertext_len));
                    if !encrypted_package_plaintext_len_is_plausible(plaintext_len, ciphertext_len) {
                        return Err(Error::UnsupportedOoxmlEncryption {
                            path: path.to_path_buf(),
                            version_major,
                            version_minor,
                        });
                    }
                    let ciphertext = &encrypted_package[8..];

                    let reader = encrypted_ooxml::decrypted_package_reader(
                        std::io::Cursor::new(ciphertext),
                        plaintext_len,
                        &encryption_info,
                        password,
                    )
                    .map_err(|err| match err {
                        encrypted_ooxml::DecryptError::InvalidPassword => Error::InvalidPassword {
                            path: path.to_path_buf(),
                        },
                        encrypted_ooxml::DecryptError::UnsupportedVersion { major, minor } => {
                            Error::UnsupportedOoxmlEncryption {
                                path: path.to_path_buf(),
                                version_major: major,
                                version_minor: minor,
                            }
                        }
                        // Preserve historical "unsupported encryption" semantics for malformed/partial
                        // encrypted containers.
                        encrypted_ooxml::DecryptError::InvalidInfo(_)
                        | encrypted_ooxml::DecryptError::Io(_) => Error::UnsupportedOoxmlEncryption {
                            path: path.to_path_buf(),
                            version_major,
                            version_minor,
                        },
                    })?;

                    let mut buf = Vec::new();
                    let mut reader = reader;
                    reader.read_to_end(&mut buf).map_err(|_source| Error::UnsupportedOoxmlEncryption {
                        path: path.to_path_buf(),
                        version_major,
                        version_minor,
                    })?;
                    buf
                }
                _ => {
                    return Err(Error::UnsupportedOoxmlEncryption {
                        path: path.to_path_buf(),
                        version_major,
                        version_minor,
                    })
                }
            },
        }
    } else {
        // Standard (CryptoAPI) encryption. There are multiple key derivation variants in the wild;
        // attempt the common CryptoAPI derivation first, then fall back to a truncated variant used
        // by some producers (notably for AES-128).
        //
        // If parsing fails due to unsupported flags (e.g. missing `fCryptoAPI`), fall back to
        // `formula-xlsx`'s legacy Standard decryptor, which is more permissive and covers some
        // fixtures.
        let decrypt_with_offcrypto = || -> Result<Vec<u8>, formula_offcrypto::OffcryptoError> {
            use sha1::{Digest as _, Sha1};

            let info = match formula_offcrypto::parse_encryption_info(&encryption_info) {
                Ok(formula_offcrypto::EncryptionInfo::Standard {
                    header, verifier, ..
                }) => formula_offcrypto::StandardEncryptionInfo { header, verifier },
                Ok(formula_offcrypto::EncryptionInfo::Agile { .. }) => {
                    // Mismatched schema: treat as unsupported for this decryptor.
                    return Err(formula_offcrypto::OffcryptoError::UnsupportedEncryption {
                        encryption_type: formula_offcrypto::EncryptionType::Agile,
                    });
                }
                Ok(formula_offcrypto::EncryptionInfo::Unsupported { version }) => {
                    return Err(formula_offcrypto::OffcryptoError::UnsupportedVersion {
                        major: version.major,
                        minor: version.minor,
                    });
                }
                // Some producers omit Standard header flags (notably `fCryptoAPI`/`fAES`) even
                // though the rest of the header follows the CryptoAPI schema. Fall back to a
                // lenient parser so we can still decrypt such workbooks.
                Err(formula_offcrypto::OffcryptoError::UnsupportedNonCryptoApiStandardEncryption)
                | Err(formula_offcrypto::OffcryptoError::InvalidFlags { .. }) => {
                    parse_standard_encryption_info_lenient(&encryption_info)?
                }
                Err(err) => return Err(err),
            };

            // Standard/CryptoAPI RC4 (CALG_RC4) uses a different key derivation than Standard AES.
            const CALG_RC4: u32 = 0x0000_6801;
            if info.header.alg_id == CALG_RC4 {
                let decrypted = formula_offcrypto::standard_rc4::decrypt_encrypted_package(
                    &info,
                    &encrypted_package,
                    password,
                )?;
                if !decrypted.starts_with(b"PK") {
                    return Err(formula_offcrypto::OffcryptoError::InvalidPassword);
                }
                return Ok(decrypted);
            }

            // --- Derive iterated SHA-1 hash (shared by both key variants) ---
            let key_len_u32 = info
                .header
                .key_size_bits
                .checked_div(8)
                .filter(|_| info.header.key_size_bits % 8 == 0)
                .ok_or_else(|| formula_offcrypto::OffcryptoError::InvalidKeySizeBits {
                    key_size_bits: info.header.key_size_bits,
                })?;
            let key_len = usize::try_from(key_len_u32).unwrap_or(0);
            if key_len == 0 {
                return Err(formula_offcrypto::OffcryptoError::InvalidKeySizeBits {
                    key_size_bits: info.header.key_size_bits,
                });
            }

            let mut password_utf16 = Vec::with_capacity(password.len().saturating_mul(2));
            for ch in password.encode_utf16() {
                password_utf16.extend_from_slice(&ch.to_le_bytes());
            }

            // h = sha1(salt || password_utf16)
            let mut hasher = Sha1::new();
            hasher.update(&info.verifier.salt);
            hasher.update(&password_utf16);
            let mut h: [u8; 20] = hasher.finalize().into();

            // for i in 0..50_000: h = sha1(u32le(i) || h)
            let mut buf = [0u8; 4 + 20];
            for i in 0..50_000u32 {
                buf[..4].copy_from_slice(&i.to_le_bytes());
                buf[4..].copy_from_slice(&h);
                h = Sha1::digest(&buf).into();
            }

            // hfinal = sha1(h || u32le(0))
            let mut buf0 = [0u8; 20 + 4];
            buf0[..20].copy_from_slice(&h);
            buf0[20..].copy_from_slice(&0u32.to_le_bytes());
            let hfinal: [u8; 20] = Sha1::digest(&buf0).into();

            let key_cryptoapi = {
                // CryptoAPI `CryptDeriveKey` semantics (as used by `msoffcrypto-tool`).
                let mut ipad = [0x36u8; 64];
                let mut opad = [0x5Cu8; 64];
                for i in 0..20 {
                    ipad[i] ^= hfinal[i];
                    opad[i] ^= hfinal[i];
                }
                let x1: [u8; 20] = Sha1::digest(&ipad).into();
                let x2: [u8; 20] = Sha1::digest(&opad).into();
                let mut out = [0u8; 40];
                out[..20].copy_from_slice(&x1);
                out[20..].copy_from_slice(&x2);
                if key_len > out.len() {
                    return Err(formula_offcrypto::OffcryptoError::DerivedKeyTooLong {
                        key_size_bits: info.header.key_size_bits,
                        required_bytes: key_len,
                        available_bytes: out.len(),
                    });
                }
                out[..key_len].to_vec()
            };

            // Alternate derivation: truncate `hfinal` directly (only meaningful for key sizes up to
            // the SHA-1 digest length).
            let key_trunc = if key_len <= hfinal.len() {
                Some(hfinal[..key_len].to_vec())
            } else {
                None
            };

            // Standard encryption is a ZIP/OPC container. Some producers vary how the verifier
            // fields and `EncryptedPackage` stream are encrypted (ECB vs CBC-segmented); since the
            // decrypted bytes must form a valid ZIP archive, attempt decryption with a small set of
            // key derivations + EncryptedPackage layouts and validate the output as a ZIP archive.
            //
            // Note: `standard_verify_key` is intentionally strict (ECB verifier scheme). For
            // compatibility with non-Excel producers, do not require verifier success before
            // attempting package decryption.
            fn looks_like_zip_container(bytes: &[u8]) -> bool {
                if !bytes.starts_with(b"PK") {
                    return false;
                }
                zip::ZipArchive::new(std::io::Cursor::new(bytes)).is_ok()
            }

            let mut key_candidates: Vec<Vec<u8>> = Vec::new();
            key_candidates.push(key_cryptoapi);
            if let Some(key) = key_trunc {
                if !key_candidates.iter().any(|k| k.as_slice() == key.as_slice()) {
                    key_candidates.push(key);
                }
            }

            for key in key_candidates {
                match formula_offcrypto::encrypted_package::decrypt_standard_encrypted_package_auto(
                    &key,
                    &info.verifier.salt,
                    &encrypted_package,
                ) {
                    Ok(decrypted) => {
                        if looks_like_zip_container(&decrypted) {
                            return Ok(decrypted);
                        }
                    }
                    Err(formula_offcrypto::OffcryptoError::InvalidPassword) => {}
                    Err(other) => return Err(other),
                }
            }

            Err(formula_offcrypto::OffcryptoError::InvalidPassword)
        };

        match decrypt_with_offcrypto() {
            Ok(bytes) => bytes,
            Err(formula_offcrypto::OffcryptoError::InvalidPassword) => {
                return Err(Error::InvalidPassword {
                    path: path.to_path_buf(),
                })
            }
            Err(formula_offcrypto::OffcryptoError::UnsupportedNonCryptoApiStandardEncryption)
            | Err(formula_offcrypto::OffcryptoError::InvalidFlags { .. })
            | Err(formula_offcrypto::OffcryptoError::UnsupportedExternalEncryption)
            | Err(formula_offcrypto::OffcryptoError::UnsupportedAlgorithm(_)) => {
                // Fall back to a more permissive Standard decryptor for non-standard/malformed
                // `EncryptionInfo` headers (e.g. some producers omit `fCryptoAPI` / `fAES` flags).
                //
                // Prefer the internal decryptor first because it supports additional Standard
                // variants (including alternative key derivations + CBC framing differences).
                match encrypted_ooxml::decrypt_encrypted_package(
                    &encryption_info,
                    &encrypted_package,
                    password,
                ) {
                    Ok(bytes) => bytes,
                    Err(err) => match err {
                        encrypted_ooxml::DecryptError::InvalidPassword => {
                            return Err(Error::InvalidPassword {
                                path: path.to_path_buf(),
                            })
                        }
                        encrypted_ooxml::DecryptError::UnsupportedVersion { major, minor } => {
                            return Err(Error::UnsupportedOoxmlEncryption {
                                path: path.to_path_buf(),
                                version_major: major,
                                version_minor: minor,
                            })
                        }
                        encrypted_ooxml::DecryptError::InvalidInfo(_)
                        | encrypted_ooxml::DecryptError::Io(_) => {
                            // As a last resort, fall back to `formula-xlsx`'s legacy Standard decryptor.
                            match xlsx::offcrypto::decrypt_ooxml_encrypted_package(
                                &encryption_info,
                                &encrypted_package,
                                password,
                            ) {
                                Ok(bytes) => bytes,
                                Err(err) => match err {
                                    xlsx::OffCryptoError::WrongPassword
                                    | xlsx::OffCryptoError::IntegrityMismatch => {
                                        return Err(Error::InvalidPassword {
                                            path: path.to_path_buf(),
                                        })
                                    }
                                    xlsx::OffCryptoError::UnsupportedEncryptionVersion {
                                        major,
                                        minor,
                                    } => {
                                        return Err(Error::UnsupportedOoxmlEncryption {
                                            path: path.to_path_buf(),
                                            version_major: major,
                                            version_minor: minor,
                                        })
                                    }
                                    _ => {
                                        return Err(Error::UnsupportedOoxmlEncryption {
                                            path: path.to_path_buf(),
                                            version_major,
                                            version_minor,
                                        })
                                    }
                                },
                            }
                        }
                    },
                }
            }
            Err(_) => {
                return Err(Error::UnsupportedOoxmlEncryption {
                    path: path.to_path_buf(),
                    version_major,
                    version_minor,
                })
            }
        }
    };

    Ok(Some(decrypted))
}

fn sniff_ooxml_zip_workbook_kind(decrypted_bytes: &[u8]) -> Option<WorkbookFormat> {
    let archive = zip::ZipArchive::new(std::io::Cursor::new(decrypted_bytes)).ok()?;

    let mut has_workbook_bin = false;
    let mut has_workbook_xml = false;
    let mut has_vba_project = false;

    for name in archive.file_names() {
        let mut normalized = name.trim_start_matches('/');
        let replaced;
        if normalized.contains('\\') {
            replaced = normalized.replace('\\', "/");
            normalized = &replaced;
        }

        if normalized.eq_ignore_ascii_case("xl/workbook.bin") {
            has_workbook_bin = true;
        } else if normalized.eq_ignore_ascii_case("xl/workbook.xml") {
            has_workbook_xml = true;
        } else if normalized.eq_ignore_ascii_case("xl/vbaProject.bin") {
            has_vba_project = true;
        }

        if has_workbook_bin || (has_workbook_xml && has_vba_project) {
            break;
        }
    }

    if has_workbook_bin {
        return Some(WorkbookFormat::Xlsb);
    }
    if has_workbook_xml {
        return Some(if has_vba_project {
            WorkbookFormat::Xlsm
        } else {
            WorkbookFormat::Xlsx
        });
    }
    None
}

fn zip_contains_workbook_bin(package_bytes: &[u8]) -> bool {
    matches!(
        sniff_ooxml_zip_workbook_kind(package_bytes),
        Some(WorkbookFormat::Xlsb)
    )
}
#[cfg(feature = "encrypted-workbooks")]
fn try_decrypt_ooxml_encrypted_package_from_path_with_preserved_ole(
    path: &Path,
    password: Option<&str>,
) -> Result<Option<(Vec<u8>, formula_office_crypto::OleEntries)>, Error> {
    use std::io::{Read as _, Seek as _};

    let mut file = std::fs::File::open(path).map_err(|source| Error::OpenIo {
        path: path.to_path_buf(),
        source,
    })?;

    let mut header = [0u8; 8];
    let n = file.read(&mut header).map_err(|source| Error::OpenIo {
        path: path.to_path_buf(),
        source,
    })?;
    if n < OLE_MAGIC.len() || header[..OLE_MAGIC.len()] != OLE_MAGIC {
        return Ok(None);
    }

    file.rewind().map_err(|source| Error::OpenIo {
        path: path.to_path_buf(),
        source,
    })?;
    let Ok(mut ole) = cfb::CompoundFile::open(file) else {
        // Malformed OLE container; fall back to non-encrypted open paths.
        return Ok(None);
    };

    // Read the required encryption streams first so we can fail fast on non-encrypted inputs and
    // wrong-password errors before doing any preservation work.
    let encryption_info = match read_stream_bytes_case_insensitive(&mut ole, "EncryptionInfo") {
        Ok(bytes) => bytes,
        Err(err) if err.kind() == std::io::ErrorKind::NotFound => return Ok(None),
        Err(_) => {
            return Err(Error::UnsupportedOoxmlEncryption {
                path: path.to_path_buf(),
                version_major: 0,
                version_minor: 0,
            })
        }
    };
    let encrypted_package = match read_stream_bytes_case_insensitive(&mut ole, "EncryptedPackage") {
        Ok(bytes) => bytes,
        Err(err) if err.kind() == std::io::ErrorKind::NotFound => return Ok(None),
        Err(_) => {
            return Err(Error::UnsupportedOoxmlEncryption {
                path: path.to_path_buf(),
                version_major: 0,
                version_minor: 0,
            })
        }
    };

    let decrypted = if let Some(package_bytes) = maybe_extract_ooxml_package_bytes(&encrypted_package) {
        if password.is_none() {
            return Err(Error::PasswordRequired {
                path: path.to_path_buf(),
            });
        }
        package_bytes.to_vec()
    } else {
        if encryption_info.len() < 4 {
            return Err(Error::UnsupportedOoxmlEncryption {
                path: path.to_path_buf(),
                version_major: 0,
                version_minor: 0,
            });
        }
        let version_major = u16::from_le_bytes([encryption_info[0], encryption_info[1]]);
        let version_minor = u16::from_le_bytes([encryption_info[2], encryption_info[3]]);

        let is_agile = version_major == 4 && version_minor == 4;
        let is_standard = version_minor == 2 && matches!(version_major, 2 | 3 | 4);
        if !is_agile && !is_standard {
            return Err(Error::UnsupportedOoxmlEncryption {
                path: path.to_path_buf(),
                version_major,
                version_minor,
            });
        }

        let Some(password) = password else {
            return Err(Error::PasswordRequired {
                path: path.to_path_buf(),
            });
        };

        // `EncryptedPackage` streams should start with an 8-byte plaintext length header followed by
        // ciphertext. If the stream is too short, treat it as a malformed/unsupported encryption
        // container rather than an invalid password.
        if encrypted_package.len() <= 8 {
            return Err(Error::UnsupportedOoxmlEncryption {
                path: path.to_path_buf(),
                version_major,
                version_minor,
            });
        }

        if is_standard {
            match formula_office_crypto::decrypt_standard_encrypted_package(
                &encryption_info,
                &encrypted_package,
                password,
            ) {
                Ok(bytes) => bytes,
                Err(err) => match err {
                    formula_office_crypto::OfficeCryptoError::InvalidPassword
                    | formula_office_crypto::OfficeCryptoError::IntegrityCheckFailed => {
                        return Err(Error::InvalidPassword {
                            path: path.to_path_buf(),
                        })
                    }
                    formula_office_crypto::OfficeCryptoError::Io(source) => {
                        return Err(Error::OpenIo {
                            path: path.to_path_buf(),
                            source,
                        })
                    }
                    _ => {
                        return Err(Error::UnsupportedOoxmlEncryption {
                            path: path.to_path_buf(),
                            version_major,
                            version_minor,
                        })
                    }
                },
            }
        } else {
            match xlsx::offcrypto::decrypt_ooxml_encrypted_package(
                &encryption_info,
                &encrypted_package,
                password,
            ) {
                Ok(bytes) => bytes,
                Err(err) => match err {
                    xlsx::OffCryptoError::WrongPassword | xlsx::OffCryptoError::IntegrityMismatch => {
                        return Err(Error::InvalidPassword {
                            path: path.to_path_buf(),
                        })
                    }
                    xlsx::OffCryptoError::UnsupportedEncryptionVersion { major, minor } => {
                        return Err(Error::UnsupportedOoxmlEncryption {
                            path: path.to_path_buf(),
                            version_major: major,
                            version_minor: minor,
                        })
                    }
                    xlsx::OffCryptoError::MissingRequiredElement { ref element }
                        if element.eq_ignore_ascii_case("dataIntegrity") =>
                    {
                        // The `EncryptedPackage` stream starts with an 8-byte plaintext length prefix.
                        if encrypted_package.len() < 8 {
                            return Err(Error::UnsupportedOoxmlEncryption {
                                path: path.to_path_buf(),
                                version_major,
                                version_minor,
                            });
                        }
                        let mut len_bytes = [0u8; 8];
                        len_bytes.copy_from_slice(&encrypted_package[..8]);
                        let ciphertext_len = encrypted_package.len().saturating_sub(8) as u64;
                        let plaintext_len =
                            parse_encrypted_package_size_prefix_bytes(len_bytes, Some(ciphertext_len));
                        if !encrypted_package_plaintext_len_is_plausible(plaintext_len, ciphertext_len) {
                            return Err(Error::UnsupportedOoxmlEncryption {
                                path: path.to_path_buf(),
                                version_major,
                                version_minor,
                            });
                        }
                        let ciphertext = &encrypted_package[8..];

                        let reader = encrypted_ooxml::decrypted_package_reader(
                            std::io::Cursor::new(ciphertext),
                            plaintext_len,
                            &encryption_info,
                            password,
                        )
                        .map_err(|err| match err {
                            encrypted_ooxml::DecryptError::InvalidPassword => Error::InvalidPassword {
                                path: path.to_path_buf(),
                            },
                            encrypted_ooxml::DecryptError::UnsupportedVersion { major, minor } => {
                                Error::UnsupportedOoxmlEncryption {
                                    path: path.to_path_buf(),
                                    version_major: major,
                                    version_minor: minor,
                                }
                            }
                            encrypted_ooxml::DecryptError::InvalidInfo(_)
                            | encrypted_ooxml::DecryptError::Io(_) => Error::UnsupportedOoxmlEncryption {
                                path: path.to_path_buf(),
                                version_major,
                                version_minor,
                            },
                        })?;

                        let mut buf = Vec::new();
                        let mut reader = reader;
                        reader
                            .read_to_end(&mut buf)
                            .map_err(|_source| Error::UnsupportedOoxmlEncryption {
                                path: path.to_path_buf(),
                                version_major,
                                version_minor,
                            })?;
                        buf
                    }
                    _ => {
                        return Err(Error::UnsupportedOoxmlEncryption {
                            path: path.to_path_buf(),
                            version_major,
                            version_minor,
                        })
                    }
                },
            }
        }
    };

    // Best-effort: preserve all other OLE streams/storages so they can be re-emitted on save.
    let preserved = match formula_office_crypto::extract_ole_entries(&mut ole) {
        Ok(entries) => entries,
        Err(err) => {
            let source = match err {
                formula_office_crypto::OfficeCryptoError::Io(e) => e,
                other => std::io::Error::new(std::io::ErrorKind::Other, other),
            };
            return Err(Error::OpenIo {
                path: path.to_path_buf(),
                source,
            });
        }
    };

    Ok(Some((decrypted, preserved)))
}
fn maybe_extract_ooxml_package_bytes(encrypted_package: &[u8]) -> Option<&[u8]> {
    // Most XLSX/ZIP containers start with `PK`.
    if encrypted_package.starts_with(b"PK") {
        return Some(encrypted_package);
    }

    // MS-OFFCRYPTO encrypted OOXML files store `EncryptedPackage` as:
    //   [u64le plaintext_size][encrypted_bytes...]
    //
    // When decryption has already been applied upstream (or when a synthetic test fixture is used),
    // the bytes after the size prefix may already be a ZIP file. Support that shape so callers can
    // still open such inputs via the "password-open" path.
    if encrypted_package.len() >= 8 {
        let declared_len = parse_encrypted_package_original_size(encrypted_package)?;
        let declared_len = usize::try_from(declared_len).ok()?;
        let rest = &encrypted_package[8..];
        if rest.starts_with(b"PK") {
            return Some(&rest[..declared_len.min(rest.len())]);
        }
    }

    None
}

#[cfg(not(feature = "encrypted-workbooks"))]
fn unsupported_office_ooxml_encryption(path: &Path) -> Error {
    Error::UnsupportedEncryption {
        path: path.to_path_buf(),
        kind: "Office-encrypted OOXML workbook (OLE EncryptionInfo + EncryptedPackage) is not supported; save a decrypted copy in Excel and try again".to_string(),
    }
}

/// Attempt to read the `EncryptedPackage` stream as a plaintext OOXML ZIP payload (without
/// performing any cryptography).
///
/// This is primarily used for:
/// - synthetic fixtures that wrap a plaintext package in an OLE container, and
/// - callers that have already decrypted the payload upstream but still provide the OOXML-in-OLE
///   wrapper structure.
///
/// Returns:
/// - `Ok(Some(zip_bytes))` when the `EncryptedPackage` stream appears to contain a valid ZIP payload
///   (either directly or after the usual 8-byte size header),
/// - `Ok(None)` when the input is not an encrypted OOXML wrapper or the package does not appear to
///   be plaintext,
/// - `Err(..)` for I/O errors while reading the file.
fn maybe_read_plaintext_ooxml_package_from_encrypted_ole_if_plaintext(
    path: &Path,
) -> Result<Option<Vec<u8>>, Error> {
    use std::io::{Read as _, Seek as _, SeekFrom};
    use std::io::{self, Cursor};

    let mut file = std::fs::File::open(path).map_err(|source| Error::OpenIo {
        path: path.to_path_buf(),
        source,
    })?;

    let mut header = [0u8; 8];
    let n = file.read(&mut header).map_err(|source| Error::OpenIo {
        path: path.to_path_buf(),
        source,
    })?;
    if n < OLE_MAGIC.len() || header[..OLE_MAGIC.len()] != OLE_MAGIC {
        return Ok(None);
    }

    file.seek(SeekFrom::Start(0))
        .map_err(|source| Error::OpenIo {
            path: path.to_path_buf(),
            source,
        })?;

    let mut ole = match cfb::CompoundFile::open(file) {
        Ok(ole) => ole,
        Err(_) => return Ok(None),
    };

    if !(stream_exists(&mut ole, "EncryptionInfo") && stream_exists(&mut ole, "EncryptedPackage")) {
        return Ok(None);
    }

    let mut stream = match open_stream_best_effort(&mut ole, "EncryptedPackage") {
        Some(s) => s,
        None => return Ok(None),
    };

    // Read a small prefix to determine if this looks like a plaintext ZIP payload.
    let mut prefix = [0u8; 10];
    let prefix_len = match stream.read(&mut prefix) {
        Ok(n) => n,
        Err(err) if err.kind() == io::ErrorKind::UnexpectedEof => 0,
        Err(err) => {
            return Err(Error::OpenIo {
                path: path.to_path_buf(),
                source: err,
            })
        }
    };
    if prefix_len < 2 {
        return Ok(None);
    }

    let looks_like_zip_direct = prefix[..2] == *b"PK";
    let looks_like_zip_after_len = prefix_len >= 10 && prefix[8..10] == *b"PK";
    if !looks_like_zip_direct && !looks_like_zip_after_len {
        return Ok(None);
    }

    // Read the full stream bytes only when the prefix suggests a plaintext package.
    let mut encrypted_package = Vec::new();
    encrypted_package.extend_from_slice(&prefix[..prefix_len]);
    stream.read_to_end(&mut encrypted_package).map_err(|source| Error::OpenIo {
        path: path.to_path_buf(),
        source,
    })?;

    let Some(package_bytes) = maybe_extract_ooxml_package_bytes(&encrypted_package) else {
        return Ok(None);
    };

    // Validate ZIP structure to avoid false positives on ciphertext that happens to start with `PK`.
    if zip::ZipArchive::new(Cursor::new(package_bytes)).is_err() {
        return Ok(None);
    }

    Ok(Some(package_bytes.to_vec()))
}

fn maybe_read_plaintext_ooxml_package_from_encrypted_ole(
    path: &Path,
    password: Option<&str>,
) -> Result<Option<Vec<u8>>, Error> {
    use std::io::{Read as _, Seek as _, SeekFrom};

    let mut file = std::fs::File::open(path).map_err(|source| Error::OpenIo {
        path: path.to_path_buf(),
        source,
    })?;

    let mut header = [0u8; 8];
    let n = file.read(&mut header).map_err(|source| Error::OpenIo {
        path: path.to_path_buf(),
        source,
    })?;
    if n < OLE_MAGIC.len() || header[..OLE_MAGIC.len()] != OLE_MAGIC {
        return Ok(None);
    }

    file.seek(SeekFrom::Start(0))
        .map_err(|source| Error::OpenIo {
            path: path.to_path_buf(),
            source,
        })?;

    let mut ole = match cfb::CompoundFile::open(file) {
        Ok(ole) => ole,
        Err(_) => {
            // Malformed OLE container; fall back to the non-encrypted open path.
            return Ok(None);
        }
    };

    if !(stream_exists(&mut ole, "EncryptionInfo") && stream_exists(&mut ole, "EncryptedPackage"))
    {
        return Ok(None);
    }

    if password.is_none() {
        #[cfg(feature = "encrypted-workbooks")]
        {
            return Err(Error::PasswordRequired {
                path: path.to_path_buf(),
            });
        }
        #[cfg(not(feature = "encrypted-workbooks"))]
        {
            return Err(unsupported_office_ooxml_encryption(path));
        }
    }

    // Read the required streams. Some producers/library combinations require a leading slash in the
    // `cfb::CompoundFile::open_stream` path, so try both forms (mirroring `stream_exists`).
    let _encryption_info = read_ole_stream_best_effort(&mut ole, "EncryptionInfo")
        .map_err(|source| Error::OpenIo {
            path: path.to_path_buf(),
            source,
        })?
        .ok_or_else(|| Error::UnsupportedOoxmlEncryption {
            path: path.to_path_buf(),
            version_major: 0,
            version_minor: 0,
        })?;
    let encrypted_package = read_ole_stream_best_effort(&mut ole, "EncryptedPackage")
        .map_err(|source| Error::OpenIo {
            path: path.to_path_buf(),
            source,
        })?
        .ok_or_else(|| Error::UnsupportedOoxmlEncryption {
            path: path.to_path_buf(),
            version_major: 0,
            version_minor: 0,
        })?;

    if let Some(package_bytes) = maybe_extract_ooxml_package_bytes(&encrypted_package) {
        // Avoid copying when the EncryptedPackage stream is already a ZIP file (rare, but useful
        // for synthetic fixtures and already-decrypted pipelines).
        if package_bytes.as_ptr() == encrypted_package.as_ptr()
            && package_bytes.len() == encrypted_package.len()
        {
            return Ok(Some(encrypted_package));
        }
        return Ok(Some(package_bytes.to_vec()));
    }

    // Not a plaintext ZIP payload; surface the usual encryption-related error.
    Err(
        encrypted_ooxml_error(&mut ole, path, password).unwrap_or(
            Error::UnsupportedOoxmlEncryption {
                path: path.to_path_buf(),
                version_major: 0,
                version_minor: 0,
            },
        ),
    )
}

/// Returns an OOXML-encryption related error when the given OLE compound file is an encrypted
/// OOXML container (`EncryptionInfo` + `EncryptedPackage` streams).
///
/// This is used to provide user-friendly error reporting for password-protected `.xlsx` files
/// (which are *not* ZIP archives; they are OLE containers that wrap an encrypted ZIP package).
fn encrypted_ooxml_error<R: std::io::Read + std::io::Write + std::io::Seek>(
    ole: &mut cfb::CompoundFile<R>,
    path: &Path,
    password: Option<&str>,
) -> Option<Error> {
    if !(stream_exists(ole, "EncryptionInfo") && stream_exists(ole, "EncryptedPackage")) {
        return None;
    }

    // EncryptionInfo begins with:
    // - VersionMajor (u16 LE)
    // - VersionMinor (u16 LE)
    //
    // See MS-OFFCRYPTO / ECMA-376. Excel commonly uses Agile encryption (4.4).
    let (version_major, version_minor) = {
        let Some(encryption_info) = read_ole_stream_best_effort(ole, "EncryptionInfo")
            .ok()
            .flatten()
        else {
            // Streams exist but can't be opened; treat as unsupported.
            return Some(Error::UnsupportedOoxmlEncryption {
                path: path.to_path_buf(),
                version_major: 0,
                version_minor: 0,
            });
        };
        if encryption_info.len() < 4 {
            return Some(Error::UnsupportedOoxmlEncryption {
                path: path.to_path_buf(),
                version_major: 0,
                version_minor: 0,
            });
        }
        (
            u16::from_le_bytes([encryption_info[0], encryption_info[1]]),
            u16::from_le_bytes([encryption_info[2], encryption_info[3]]),
        )
    };

    // Most real-world Excel files use either:
    // - Agile encryption (4.4; XML descriptor payload)
    // - Standard/CryptoAPI encryption (major in {2,3,4}, minor=2; binary header/verifier)
    //
    // Be defensive around malformed/synthetic fixtures: if the "version" header doesn't look like
    // a plausible small integer pair, fall back to generic "password required" semantics instead of
    // reporting a nonsense version.
    if version_major > 10 || version_minor > 10 {
        if password.is_none() {
            return Some(Error::PasswordRequired {
                path: path.to_path_buf(),
            });
        }
        return Some(Error::InvalidPassword {
            path: path.to_path_buf(),
        });
    }

    let is_agile = version_major == 4 && version_minor == 4;
    let is_standard = version_minor == 2 && matches!(version_major, 2 | 3 | 4);
    if !is_agile && !is_standard {
        return Some(Error::UnsupportedOoxmlEncryption {
            path: path.to_path_buf(),
            version_major,
            version_minor,
        });
    }

    #[cfg(not(feature = "encrypted-workbooks"))]
    {
        return Some(unsupported_office_ooxml_encryption(path));
    }

    #[cfg(feature = "encrypted-workbooks")]
    {
        if password.is_none() {
            return Some(Error::PasswordRequired {
                path: path.to_path_buf(),
            });
        }

        // We don't attempt to decrypt in this helper; it exists to provide UX-friendly error
        // classification for encrypted OOXML wrappers (`EncryptionInfo` + `EncryptedPackage`
        // streams).
        //
        // When the `encrypted-workbooks` feature is enabled, the password-aware open paths attempt
        // in-memory decryption earlier and only fall back to this classification logic when
        // decryption is not in play.
        Some(Error::InvalidPassword {
            path: path.to_path_buf(),
        })
    }
}

#[cfg(feature = "encrypted-workbooks")]
#[allow(dead_code)]
fn open_encrypted_ooxml_model_workbook(
    path: &Path,
    password: &str,
) -> Result<Option<formula_model::Workbook>, Error> {
    let Some(decrypted) = try_decrypt_ooxml_encrypted_package_from_path(path, Some(password))?
    else {
        return Ok(None);
    };
    open_workbook_model_from_decrypted_ooxml_zip_bytes(path, decrypted).map(Some)
}

#[cfg(feature = "encrypted-workbooks")]
#[allow(dead_code)]
fn open_encrypted_ooxml_workbook(path: &Path, password: &str) -> Result<Option<Workbook>, Error> {
    let Some(decrypted) = try_decrypt_ooxml_encrypted_package_from_path(path, Some(password))?
    else {
        return Ok(None);
    };
    open_workbook_from_decrypted_ooxml_zip_bytes(path, decrypted).map(Some)
}

/// Decrypt an Office-encrypted OOXML workbook (OLE `EncryptionInfo` + `EncryptedPackage`) into the
/// plaintext OOXML ZIP bytes, along with the detected workbook format.
///
/// Returns `Ok(None)` when the file does not appear to be an encrypted OOXML container.
#[cfg(feature = "encrypted-workbooks")]
#[allow(dead_code)]
fn decrypt_encrypted_ooxml_package(
    path: &Path,
    password: &str,
) -> Result<Option<(WorkbookFormat, Vec<u8>)>, Error> {
    use std::io::Read as _;

    let mut file = std::fs::File::open(path).map_err(|source| Error::OpenIo {
        path: path.to_path_buf(),
        source,
    })?;

    let mut header = [0u8; 8];
    let n = file.read(&mut header).map_err(|source| Error::OpenIo {
        path: path.to_path_buf(),
        source,
    })?;
    if n < OLE_MAGIC.len() || header[..OLE_MAGIC.len()] != OLE_MAGIC {
        return Ok(None);
    }

    file.rewind().map_err(|source| Error::OpenIo {
        path: path.to_path_buf(),
        source,
    })?;
    let Ok(mut ole) = cfb::CompoundFile::open(file) else {
        // Malformed OLE container; treat as "not an encrypted OOXML workbook" and let normal open
        // code paths surface the parsing error.
        return Ok(None);
    };

    if !(stream_exists(&mut ole, "EncryptionInfo") && stream_exists(&mut ole, "EncryptedPackage")) {
        return Ok(None);
    }

    // Read full `EncryptionInfo` and `EncryptedPackage` streams (best-effort for leading slash and
    // casing variations).
    let encryption_info = read_ole_stream_best_effort(&mut ole, "EncryptionInfo")
        .map_err(|source| Error::OpenIo {
            path: path.to_path_buf(),
            source,
        })?
        .ok_or_else(|| Error::UnsupportedOoxmlEncryption {
            path: path.to_path_buf(),
            version_major: 0,
            version_minor: 0,
        })?;
    let encrypted_package = read_ole_stream_best_effort(&mut ole, "EncryptedPackage")
        .map_err(|source| Error::OpenIo {
            path: path.to_path_buf(),
            source,
        })?
        .ok_or_else(|| Error::UnsupportedOoxmlEncryption {
            path: path.to_path_buf(),
            version_major: 0,
            version_minor: 0,
        })?;

    // Some synthetic fixtures (and some pipelines) may already contain a plaintext OOXML ZIP payload
    // in `EncryptedPackage`. When that happens, skip cryptography entirely and treat the bytes as
    // the "decrypted" package payload.
    if let Some(package_bytes) = maybe_extract_ooxml_package_bytes(&encrypted_package) {
        // Validate ZIP structure before treating the stream as plaintext to avoid false positives
        // on ciphertext that happens to begin with `PK`.
        if let Some(format) = workbook_format_from_ooxml_zip_bytes(package_bytes) {
            let format = if matches!(format, WorkbookFormat::Unknown) {
                WorkbookFormat::Xlsx
            } else {
                format
            };
            let reuse_full_buffer = package_bytes.as_ptr() == encrypted_package.as_ptr()
                && package_bytes.len() == encrypted_package.len();
            if reuse_full_buffer {
                return Ok(Some((format, encrypted_package)));
            }
            return Ok(Some((format, package_bytes.to_vec())));
        }
    }

    // `EncryptionInfo` version header is the first 4 bytes:
    // - VersionMajor (u16 LE)
    // - VersionMinor (u16 LE)
    let (version_major, version_minor) = match encryption_info.get(0..4) {
        Some(header) => (
            u16::from_le_bytes([header[0], header[1]]),
            u16::from_le_bytes([header[2], header[3]]),
        ),
        None => {
            return Err(Error::UnsupportedOoxmlEncryption {
                path: path.to_path_buf(),
                version_major: 0,
                version_minor: 0,
            });
        }
    };

    // Mirror the defensive behavior in `encrypted_ooxml_error`: treat implausible version headers
    // as a generic "encrypted container" so we surface `InvalidPassword` instead of nonsense.
    if version_major > 10 || version_minor > 10 {
        return Err(Error::InvalidPassword {
            path: path.to_path_buf(),
        });
    }

    let is_agile = version_major == 4 && version_minor == 4;
    let is_standard = version_minor == 2 && matches!(version_major, 2 | 3 | 4);
    if !is_agile && !is_standard {
        return Err(Error::UnsupportedOoxmlEncryption {
            path: path.to_path_buf(),
            version_major,
            version_minor,
        });
    }

    let decrypted = if is_agile {
        let unsupported = || Error::UnsupportedOoxmlEncryption {
            path: path.to_path_buf(),
            version_major,
            version_minor,
        };

        // Real-world producers vary in how they encode/wrap the Agile `EncryptionInfo` XML
        // (UTF-8/UTF-16, length prefixes, padding). Normalize to a strict UTF-8 XML payload so we
        // can reuse the `formula-xlsx` offcrypto implementation.
        //
        // Treat malformed Agile descriptors as unsupported encryption (not a wrong password).
        let xml = extract_agile_encryption_info_xml(&encryption_info).map_err(|_| unsupported())?;
        let mut normalized_info = Vec::with_capacity(8 + xml.len());
        normalized_info.extend_from_slice(encryption_info.get(..8).ok_or_else(unsupported)?);
        normalized_info.extend_from_slice(xml.as_bytes());

        match xlsx::decrypt_agile_encrypted_package(&normalized_info, &encrypted_package, password) {
            Ok(bytes) => bytes,
            Err(err) => match err {
                xlsx::OffCryptoError::WrongPassword | xlsx::OffCryptoError::IntegrityMismatch => {
                    return Err(Error::InvalidPassword {
                        path: path.to_path_buf(),
                    });
                }
                xlsx::OffCryptoError::UnsupportedEncryptionVersion { major, minor } => {
                    return Err(Error::UnsupportedOoxmlEncryption {
                        path: path.to_path_buf(),
                        version_major: major,
                        version_minor: minor,
                    });
                }
                _ => {
                    return Err(Error::UnsupportedOoxmlEncryption {
                        path: path.to_path_buf(),
                        version_major,
                        version_minor,
                    });
                }
            },
        }
    } else {
        // Standard/CryptoAPI encryption has multiple variants in the wild.
        //
        // In particular, some workbooks use the "CryptoAPI RC4" cipher (CALG_RC4) instead of AES.
        // Excel-generated Standard encryption typically uses AES, but we keep the implementation
        // compatible with the committed corpus fixtures (including RC4).
        //
        // References:
        // - `docs/offcrypto-standard-cryptoapi-rc4.md`
        // - `crates/formula-io/tests/offcrypto_standard_rc4_vectors.rs`
        use crate::offcrypto::cryptoapi::{
            hash_password_fixed_spin,
            password_to_utf16le,
            HashAlg as CryptoApiHashAlg,
        };

        // Parse the Standard EncryptionInfo stream to read algId/algIdHash/keySize + verifier
        // fields. For the RC4 variant, we need the salt + key size to derive per-block keys.
        //
        // Stream layout (MS-OFFCRYPTO):
        // [0..8)   EncryptionVersionInfo (major, minor, flags)
        // [8..12)  EncryptionHeaderSize (u32 LE)
        // [..]     EncryptionHeader (header_size bytes)
        // [..]     EncryptionVerifier (salt/verifier/verifierHash)
        let header_size = encryption_info
            .get(8..12)
            .and_then(|b| b.try_into().ok())
            .map(u32::from_le_bytes)
            .ok_or_else(|| Error::InvalidPassword {
                path: path.to_path_buf(),
            })? as usize;

        let header_start = 12usize;
        let header_end = header_start.saturating_add(header_size);
        let header_bytes = encryption_info
            .get(header_start..header_end)
            .ok_or_else(|| Error::InvalidPassword {
                path: path.to_path_buf(),
            })?;
        if header_bytes.len() < 32 {
            return Err(Error::InvalidPassword {
                path: path.to_path_buf(),
            });
        }

        let alg_id = u32::from_le_bytes(header_bytes[8..12].try_into().unwrap());
        let alg_id_hash = u32::from_le_bytes(header_bytes[12..16].try_into().unwrap());
        let key_size_bits = u32::from_le_bytes(header_bytes[16..20].try_into().unwrap());

        let mut offset = header_end;
        let salt_size: usize = encryption_info
            .get(offset..offset + 4)
            .and_then(|b| b.try_into().ok())
            .map(u32::from_le_bytes)
            .ok_or_else(|| Error::InvalidPassword {
                path: path.to_path_buf(),
            })? as usize;
        offset += 4;
        let salt = encryption_info
            .get(offset..offset + salt_size)
            .ok_or_else(|| Error::InvalidPassword {
                path: path.to_path_buf(),
            })?
            .to_vec();
        offset += salt_size;

        let encrypted_verifier = encryption_info
            .get(offset..offset + 16)
            .ok_or_else(|| Error::InvalidPassword {
                path: path.to_path_buf(),
            })?
            .to_vec();
        offset += 16;

        let verifier_hash_size = encryption_info
            .get(offset..offset + 4)
            .and_then(|b| b.try_into().ok())
            .map(u32::from_le_bytes)
            .ok_or_else(|| Error::InvalidPassword {
                path: path.to_path_buf(),
            })?;
        offset += 4;

        let encrypted_verifier_hash = encryption_info
            .get(offset..)
            .unwrap_or_default()
            .to_vec();

        const CALG_RC4: u32 = 0x0000_6801;

        if alg_id == CALG_RC4 {
            // --- Standard CryptoAPI RC4 ---------------------------------------------------------
            let hash_alg = CryptoApiHashAlg::from_calg_id(alg_id_hash).map_err(|_| {
                Error::InvalidPassword {
                    path: path.to_path_buf(),
                }
            })?;

            let key_size_bits_raw = key_size_bits;
            // MS-OFFCRYPTO specifies that `keySize=0` MUST be interpreted as 40-bit.
            let key_size_bits = if key_size_bits == 0 { 40 } else { key_size_bits };
            if key_size_bits % 8 != 0 {
                return Err(Error::InvalidPassword {
                    path: path.to_path_buf(),
                });
            }
            let key_len = (key_size_bits / 8) as usize;

            // Derive the spun password hash H (Hash(salt || password_utf16le) then 50k rounds).
            let pw_utf16le = password_to_utf16le(password);
            let h_vec = hash_password_fixed_spin(&pw_utf16le, &salt, hash_alg);

            // Verify password using EncryptionVerifier:
            // decrypt `encryptedVerifier || encryptedVerifierHash` using block 0 key.
            let mut verifier_cipher = Vec::new();
            verifier_cipher.extend_from_slice(&encrypted_verifier);
            verifier_cipher.extend_from_slice(&encrypted_verifier_hash);
            let mut verifier_reader = Rc4CryptoApiDecryptReader::new_with_hash_alg(
                std::io::Cursor::new(verifier_cipher.as_slice()),
                verifier_cipher.len() as u64,
                h_vec.clone(),
                key_len,
                hash_alg,
            )
            .map_err(|_| Error::InvalidPassword {
                path: path.to_path_buf(),
            })?;
            let mut verifier_plain = Vec::new();
            verifier_reader
                .read_to_end(&mut verifier_plain)
                .map_err(|_| Error::InvalidPassword {
                    path: path.to_path_buf(),
                })?;
            if verifier_plain.len() < 16 {
                return Err(Error::InvalidPassword {
                    path: path.to_path_buf(),
                });
            }
            let verifier = &verifier_plain[..16];
            let hash_len = verifier_hash_size as usize;
            let decrypted_hash = verifier_plain
                .get(16..16 + hash_len)
                .ok_or_else(|| Error::InvalidPassword {
                    path: path.to_path_buf(),
                })?;
            let expected_hash = match hash_alg {
                CryptoApiHashAlg::Sha1 => {
                    use sha1::Digest as _;
                    sha1::Sha1::digest(verifier).to_vec()
                }
                CryptoApiHashAlg::Md5 => {
                    use md5::Digest as _;
                    md5::Md5::digest(verifier).to_vec()
                }
            };
            let expected_hash = expected_hash
                .get(..hash_len)
                .ok_or_else(|| Error::InvalidPassword {
                    path: path.to_path_buf(),
                })?;
            if !crate::offcrypto::standard::ct_eq(expected_hash, decrypted_hash) {
                return Err(Error::InvalidPassword {
                    path: path.to_path_buf(),
                });
            }

            // Decrypt the EncryptedPackage stream (RC4 in 0x200-byte blocks) and truncate to the
            // plaintext `package_size` prefix.
            let cursor = std::io::Cursor::new(encrypted_package.as_slice());
            let mut reader = Rc4CryptoApiDecryptReader::from_encrypted_package_stream(
                cursor,
                h_vec,
                key_size_bits_raw,
                alg_id_hash,
            )
            .map_err(|_| Error::InvalidPassword {
                path: path.to_path_buf(),
            })?;
            let mut out = Vec::new();
            reader.read_to_end(&mut out).map_err(|_| Error::InvalidPassword {
                path: path.to_path_buf(),
            })?;
            out
        } else {
            // --- Standard CryptoAPI AES ---------------------------------------------------------
            // Prefer the Standard decryptor in `formula-office-crypto` because it supports a wider
            // range of hash algorithms (and performs verifier validation). However, some producers
            // omit or mis-set `EncryptionHeader.Flags` (e.g. missing `fCryptoAPI`/`fAES`), which the
            // strict parser rejects. Fall back to `formula-xlsx`'s Standard decryptor
            // (`formula-offcrypto`) for compatibility in those cases.
            match formula_office_crypto::decrypt_standard_encrypted_package(
                &encryption_info,
                &encrypted_package,
                password,
            ) {
                Ok(bytes) => bytes,
                Err(formula_office_crypto::OfficeCryptoError::InvalidPassword) => {
                    return Err(Error::InvalidPassword {
                        path: path.to_path_buf(),
                    });
                }
                Err(_) => match xlsx::offcrypto::decrypt_ooxml_encrypted_package(
                    &encryption_info,
                    &encrypted_package,
                    password,
                ) {
                    Ok(bytes) => bytes,
                    Err(err) => match err {
                        xlsx::OffCryptoError::WrongPassword
                        | xlsx::OffCryptoError::IntegrityMismatch => {
                            return Err(Error::InvalidPassword {
                                path: path.to_path_buf(),
                            });
                        }
                        xlsx::OffCryptoError::UnsupportedEncryptionVersion { major, minor } => {
                            return Err(Error::UnsupportedOoxmlEncryption {
                                path: path.to_path_buf(),
                                version_major: major,
                                version_minor: minor,
                            });
                        }
                        _ => {
                            return Err(Error::UnsupportedOoxmlEncryption {
                                path: path.to_path_buf(),
                                version_major,
                                version_minor,
                            });
                        }
                    },
                },
            }
        }
    };
    let format = match workbook_format_from_ooxml_zip_bytes(&decrypted) {
        Some(WorkbookFormat::Xlsb) => WorkbookFormat::Xlsb,
        Some(WorkbookFormat::Xlsm) => WorkbookFormat::Xlsm,
        Some(WorkbookFormat::Xlsx) => WorkbookFormat::Xlsx,
        // The decrypted bytes are a ZIP file but don't clearly contain either an XLSX/XLSM or XLSB
        // workbook. Treat as an XLSX package so callers can still access the decrypted bytes via
        // `XlsxLazyPackage` (useful for synthetic fixtures).
        Some(WorkbookFormat::Unknown) => WorkbookFormat::Xlsx,
        // A successful OOXML `EncryptedPackage` decryption must yield a ZIP payload. If it does not,
        // treat the result as an invalid password (or corrupt/malformed encrypted package).
        None => {
            return Err(Error::InvalidPassword {
                path: path.to_path_buf(),
            });
        }
        // Non-OOXML formats are not expected from an OOXML `EncryptedPackage` payload.
        Some(_) => WorkbookFormat::Xlsx,
    };
    Ok(Some((format, decrypted)))
}

#[cfg(feature = "encrypted-workbooks")]
#[allow(dead_code)]
fn workbook_format_from_ooxml_zip_bytes(bytes: &[u8]) -> Option<WorkbookFormat> {
    use std::io::Cursor;

    if bytes.len() < 2 || bytes[..2] != *b"PK" {
        return None;
    }
    let cursor = Cursor::new(bytes);
    let Ok(archive) = zip::ZipArchive::new(cursor) else {
        return None;
    };

    let mut has_workbook_bin = false;
    let mut has_workbook_xml = false;
    let mut has_vba_project = false;

    for name in archive.file_names() {
        let mut normalized = name.trim_start_matches('/');
        let replaced;
        if normalized.contains('\\') {
            replaced = normalized.replace('\\', "/");
            normalized = &replaced;
        }

        if normalized.eq_ignore_ascii_case("xl/workbook.bin") {
            has_workbook_bin = true;
        } else if normalized.eq_ignore_ascii_case("xl/workbook.xml") {
            has_workbook_xml = true;
        } else if normalized.eq_ignore_ascii_case("xl/vbaProject.bin") {
            has_vba_project = true;
        }

        if has_workbook_bin || (has_workbook_xml && has_vba_project) {
            break;
        }
    }

    if has_workbook_bin {
        return Some(WorkbookFormat::Xlsb);
    }
    if has_workbook_xml {
        return Some(if has_vba_project {
            WorkbookFormat::Xlsm
        } else {
            WorkbookFormat::Xlsx
        });
    }

    Some(WorkbookFormat::Unknown)
}
/// Return `Some(true)` when the OLE `Workbook`/`Book` stream starts with a BIFF `BOF` record.
///
/// Returns:
/// - `Some(true)`  => stream looks like BIFF (starts with BOF)
/// - `Some(false)` => stream is present but does *not* look like BIFF
/// - `None`        => stream couldn't be opened or read (treat as "unknown"; callers should be
///   conservative and avoid rejecting potentially-corrupt but otherwise valid `.xls` files)
fn ole_workbook_stream_starts_with_biff_bof<R: std::io::Read + std::io::Write + std::io::Seek>(
    ole: &mut cfb::CompoundFile<R>,
) -> Option<bool> {
    use std::io::Read as _;

    let mut stream = None;
    for candidate in ["Workbook", "/Workbook", "Book", "/Book"] {
        if let Ok(s) = ole.open_stream(candidate) {
            stream = Some(s);
            break;
        }
    }
    let mut stream = stream?;

    let mut header = [0u8; 4];
    if stream.read_exact(&mut header).is_err() {
        return None;
    }

    let record_id = u16::from_le_bytes([header[0], header[1]]);
    Some(matches!(
        record_id,
        BIFF_RECORD_BOF_BIFF8 | BIFF_RECORD_BOF_BIFF5
    ))
}

fn ole_workbook_has_biff_filepass_record<R: std::io::Read + std::io::Write + std::io::Seek>(
    ole: &mut cfb::CompoundFile<R>,
) -> bool {
    use std::io::{Read as _, Seek as _, SeekFrom};

    let mut stream = None;
    for candidate in ["Workbook", "/Workbook", "Book", "/Book"] {
        if let Ok(s) = ole.open_stream(candidate) {
            stream = Some(s);
            break;
        }
    }
    let mut stream = match stream {
        Some(s) => s,
        None => return false,
    };

    // Best-effort scan over BIFF records in the workbook globals substream.
    //
    // This is intentionally minimal and defensive:
    // - stop at the first `EOF`
    // - stop at the next `BOF` after the first record (indicates the next substream; some truncated
    //   files omit the expected `EOF`)
    // - stop after a small byte budget to avoid scanning huge streams during format detection
    let mut first_record = true;
    let mut scanned: usize = 0;
    const MAX_SCAN_BYTES: usize = 4 * 1024 * 1024; // 4MiB should comfortably cover workbook globals headers

    loop {
        let mut header = [0u8; 4];
        if stream.read_exact(&mut header).is_err() {
            return false;
        }
        scanned = scanned.saturating_add(4);

        let record_id = u16::from_le_bytes([header[0], header[1]]);
        let len = u16::from_le_bytes([header[2], header[3]]) as usize;

        // BIFF streams always begin with a BOF record. If the workbook stream doesn't look like
        // BIFF, don't attempt to interpret it as record headers (avoids false positives on other
        // OLE payloads that happen to contain a `0x002F` word early).
        if first_record && !matches!(record_id, BIFF_RECORD_BOF_BIFF8 | BIFF_RECORD_BOF_BIFF5) {
            return false;
        }

        if record_id == BIFF_RECORD_FILEPASS {
            return true;
        }
        if record_id == BIFF_RECORD_EOF {
            break;
        }
        if !first_record && matches!(record_id, BIFF_RECORD_BOF_BIFF8 | BIFF_RECORD_BOF_BIFF5) {
            break;
        }
        first_record = false;

        // Skip record payload bytes.
        if stream.seek(SeekFrom::Current(len as i64)).is_err() {
            // Fallback when seeking isn't supported: read + discard.
            let mut remaining = len;
            let mut buf = [0u8; 4096];
            while remaining > 0 {
                let to_read = remaining.min(buf.len());
                if stream.read_exact(&mut buf[..to_read]).is_err() {
                    return false;
                }
                remaining -= to_read;
            }
        }
        scanned = scanned.saturating_add(len);
        if scanned >= MAX_SCAN_BYTES {
            break;
        }
    }

    false
}

/// Best-effort parse of a legacy `.xls` BIFF `FILEPASS` record inside the workbook globals stream.
///
/// Returns:
/// - `None`       => no `FILEPASS` record found (or the workbook stream isn't a BIFF stream)
/// - `Some(None)` => `FILEPASS` record found, but the payload was missing/truncated or couldn't be
///   interpreted
/// - `Some(Some(scheme))` => `FILEPASS` record found and scheme was classified
fn ole_workbook_filepass_scheme<R: std::io::Read + std::io::Write + std::io::Seek>(
    ole: &mut cfb::CompoundFile<R>,
) -> Option<Option<LegacyXlsFilePassScheme>> {
    use std::io::{Read as _, Seek as _, SeekFrom};

    let mut stream = None;
    for candidate in ["Workbook", "/Workbook", "Book", "/Book"] {
        if let Ok(s) = ole.open_stream(candidate) {
            stream = Some(s);
            break;
        }
    }
    let mut stream = stream?;

    // Best-effort scan over BIFF records in the workbook globals substream.
    //
    // Mirrors `ole_workbook_has_biff_filepass_record`, but reads and interprets the FILEPASS payload
    // when present.
    let mut first_record = true;
    let mut bof_record_id: u16 = 0;
    let mut scanned: usize = 0;
    const MAX_SCAN_BYTES: usize = 4 * 1024 * 1024; // 4MiB

    loop {
        let mut header = [0u8; 4];
        if stream.read_exact(&mut header).is_err() {
            return None;
        }
        scanned = scanned.saturating_add(4);

        let record_id = u16::from_le_bytes([header[0], header[1]]);
        let len = u16::from_le_bytes([header[2], header[3]]) as usize;

        // BIFF streams always begin with a BOF record. If the workbook stream doesn't look like
        // BIFF, don't attempt to interpret it as record headers (avoids false positives on other
        // OLE payloads that happen to contain a `0x002F` word early).
        if first_record && !matches!(record_id, BIFF_RECORD_BOF_BIFF8 | BIFF_RECORD_BOF_BIFF5) {
            return None;
        }
        if first_record {
            bof_record_id = record_id;
        }

        if record_id == BIFF_RECORD_FILEPASS {
            let mut payload = vec![0u8; len];
            if stream.read_exact(&mut payload).is_err() {
                return Some(None);
            }

            // BIFF5/BIFF8 distinguish their BOF record ids. BIFF5 encryption is XOR-only.
            if bof_record_id == BIFF_RECORD_BOF_BIFF5 {
                return Some(Some(LegacyXlsFilePassScheme::Xor));
            }

            // BIFF8 FILEPASS starts with:
            // - wEncryptionType (u16)
            // - wEncryptionSubType (u16) when wEncryptionType == 0x0001 (RC4)
            if payload.len() < 2 {
                return Some(None);
            }
            let encryption_type = u16::from_le_bytes([payload[0], payload[1]]);
            let scheme = match encryption_type {
                0x0000 => Some(LegacyXlsFilePassScheme::Xor),
                0x0001 => {
                    if payload.len() < 4 {
                        Some(LegacyXlsFilePassScheme::Unknown)
                    } else {
                        let sub_type = u16::from_le_bytes([payload[2], payload[3]]);
                        match sub_type {
                            0x0001 => Some(LegacyXlsFilePassScheme::Rc4),
                            0x0002 => Some(LegacyXlsFilePassScheme::Rc4CryptoApi),
                            _ => Some(LegacyXlsFilePassScheme::Unknown),
                        }
                    }
                }
                _ => Some(LegacyXlsFilePassScheme::Unknown),
            };
            return Some(scheme);
        }

        if record_id == BIFF_RECORD_EOF {
            break;
        }
        if !first_record && matches!(record_id, BIFF_RECORD_BOF_BIFF8 | BIFF_RECORD_BOF_BIFF5) {
            break;
        }
        first_record = false;

        // Skip record payload bytes.
        if stream.seek(SeekFrom::Current(len as i64)).is_err() {
            // Fallback when seeking isn't supported: read + discard.
            let mut remaining = len;
            let mut buf = [0u8; 4096];
            while remaining > 0 {
                let to_read = remaining.min(buf.len());
                if stream.read_exact(&mut buf[..to_read]).is_err() {
                    return None;
                }
                remaining -= to_read;
            }
        }
        scanned = scanned.saturating_add(len);
        if scanned >= MAX_SCAN_BYTES {
            break;
        }
    }

    None
}

/// Open a spreadsheet workbook based on file extension.
///
/// For `.xlsx` / `.xlsm` inputs, this returns [`Workbook::Xlsx`] backed by
/// [`formula_xlsx::XlsxLazyPackage`], which keeps the underlying OPC ZIP container as a lazy source
/// (file path or bytes) and only inflates individual parts on demand. Saving uses a streaming ZIP
/// rewrite path that **raw-copies** untouched entries (`zip::ZipWriter::raw_copy_file`) for
/// performance and round-trip fidelity.
///
/// If you only need a [`formula_model::Workbook`] (data + formulas) and do not need OPC-level
/// round-trip preservation, prefer [`open_workbook_model`], which parses only the parts needed to
/// build the model.
///
/// Currently supports:
/// - `.xls` / `.xlt` / `.xla` (via `formula-xls`)
/// - `.xlsb` (via `formula-xlsb`)
/// - `.xlsx` / `.xlsm` / `.xltx` / `.xltm` / `.xlam` (via `formula-xlsx`)
/// - `.csv` (via `formula-model` CSV import)
/// - `.parquet` (via `formula-columnar`, requires the `formula-io` crate feature `parquet`)
pub fn open_workbook(path: impl AsRef<Path>) -> Result<Workbook, Error> {
    let path = path.as_ref();
    let ext = path
        .extension()
        .and_then(|s| s.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase();

    match workbook_format(path)? {
        WorkbookFormat::Xlsx | WorkbookFormat::Xlsm => {
            let file = std::fs::File::open(path).map_err(|source| Error::OpenIo {
                path: path.to_path_buf(),
                source,
            })?;
            let package = xlsx::XlsxLazyPackage::from_file(path.to_path_buf(), file).map_err(
                |source| Error::OpenXlsx {
                    path: path.to_path_buf(),
                    source,
                },
            )?;
            Ok(Workbook::Xlsx(package))
        }
        WorkbookFormat::Xls => match xls::import_xls_path(path) {
            Ok(result) => Ok(Workbook::Xls(result)),
            Err(xls::ImportError::EncryptedWorkbook) => Err(Error::EncryptedWorkbook {
                path: path.to_path_buf(),
            }),
            Err(source) => Err(Error::OpenXls {
                path: path.to_path_buf(),
                source,
            }),
        },
        WorkbookFormat::Xlsb => {
            xlsb::XlsbWorkbook::open(path)
                .map(Workbook::Xlsb)
                .map_err(|source| Error::OpenXlsb {
                    path: path.to_path_buf(),
                    source,
                })
        }
        WorkbookFormat::Csv => {
            let file = std::fs::File::open(path).map_err(|source| Error::OpenIo {
                path: path.to_path_buf(),
                source,
            })?;
            let reader = std::io::BufReader::new(file);

            let stem = path.file_stem().and_then(|s| s.to_str()).unwrap_or("");
            let sheet_name = formula_model::validate_sheet_name(stem)
                .ok()
                .map(|_| stem.to_string())
                .unwrap_or_else(|| sanitize_sheet_name(stem));

            let mut workbook = formula_model::Workbook::new();
            import_csv_into_workbook(&mut workbook, sheet_name, reader, CsvOptions::default())
                .map_err(|source| Error::OpenCsv {
                    path: path.to_path_buf(),
                    source,
                })?;

            Ok(Workbook::Model(workbook))
        }
        WorkbookFormat::Parquet => open_parquet_model_workbook(path).map(Workbook::Model),
        _ => Err(Error::UnsupportedExtension {
            path: path.to_path_buf(),
            extension: ext.to_string(),
        }),
    }
}

#[cfg(feature = "encrypted-workbooks")]
#[allow(dead_code)]
fn try_open_standard_aes_encrypted_ooxml_model_workbook(
    path: &Path,
    password: Option<&str>,
) -> Result<Option<formula_model::Workbook>, Error> {
    use std::io::{Read as _, Seek as _, SeekFrom};

    // Only handle Office-encrypted OOXML OLE containers.
    let mut file = std::fs::File::open(path).map_err(|source| Error::OpenIo {
        path: path.to_path_buf(),
        source,
    })?;
    let mut header = [0u8; 8];
    let n = file.read(&mut header).map_err(|source| Error::OpenIo {
        path: path.to_path_buf(),
        source,
    })?;
    if n < OLE_MAGIC.len() || header[..OLE_MAGIC.len()] != OLE_MAGIC {
        return Ok(None);
    }
    file.rewind().map_err(|source| Error::OpenIo {
        path: path.to_path_buf(),
        source,
    })?;

    let Ok(mut ole) = cfb::CompoundFile::open(file) else {
        // Malformed OLE container; let the normal open path surface errors.
        return Ok(None);
    };

    // Read `EncryptionInfo` so we can decide whether this is Standard/CryptoAPI AES.
    let encryption_info = match read_stream_bytes_case_insensitive(&mut ole, "EncryptionInfo") {
        Ok(bytes) => bytes,
        Err(err) if err.kind() == std::io::ErrorKind::NotFound => return Ok(None),
        Err(_) => {
            return Err(Error::UnsupportedOoxmlEncryption {
                path: path.to_path_buf(),
                version_major: 0,
                version_minor: 0,
            })
        }
    };

    if encryption_info.len() < 4 {
        return Err(Error::UnsupportedOoxmlEncryption {
            path: path.to_path_buf(),
            version_major: 0,
            version_minor: 0,
        });
    }
    let version_major = u16::from_le_bytes([encryption_info[0], encryption_info[1]]);
    let version_minor = u16::from_le_bytes([encryption_info[2], encryption_info[3]]);

    // Only handle Standard/CryptoAPI encryption (minor == 2). Agile decryption uses different open
    // semantics (we preserve the ZIP package for Workbook::Xlsx).
    let is_standard = version_minor == 2 && matches!(version_major, 2 | 3 | 4);
    if !is_standard {
        return Ok(None);
    }

    // Parse Standard header to detect RC4 vs AES.
    let info = crate::offcrypto::parse_encryption_info_standard(&encryption_info).map_err(|err| {
        match err {
            crate::offcrypto::OffcryptoError::UnsupportedEncryptionInfoVersion { major, minor, .. } => {
                Error::UnsupportedOoxmlEncryption {
                    path: path.to_path_buf(),
                    version_major: major,
                    version_minor: minor,
                }
            }
            _ => Error::UnsupportedOoxmlEncryption {
                path: path.to_path_buf(),
                version_major,
                version_minor,
            },
        }
    })?;

    if info.header.alg_id == crate::offcrypto::CALG_RC4 {
        // Standard/CryptoAPI RC4 is still opened as a preserved ZIP package via the in-memory
        // decrypt path (it is not supported by the streaming decrypt reader).
        return Ok(None);
    }

    if !matches!(
        info.header.alg_id,
        crate::offcrypto::CALG_AES_128 | crate::offcrypto::CALG_AES_192 | crate::offcrypto::CALG_AES_256
    ) {
        return Err(Error::UnsupportedOoxmlEncryption {
            path: path.to_path_buf(),
            version_major,
            version_minor,
        });
    }

    let Some(password) = password else {
        return Err(Error::PasswordRequired {
            path: path.to_path_buf(),
        });
    };

    let mut encrypted_package_stream = match open_stream_case_insensitive(&mut ole, "EncryptedPackage")
    {
        Ok(stream) => stream,
        Err(err) if err.kind() == std::io::ErrorKind::NotFound => return Ok(None),
        Err(_) => {
            return Err(Error::UnsupportedOoxmlEncryption {
                path: path.to_path_buf(),
                version_major,
                version_minor,
            })
        }
    };

    // The `EncryptedPackage` stream begins with an 8-byte plaintext length prefix.
    let mut len_bytes = [0u8; 8];
    encrypted_package_stream
        .read_exact(&mut len_bytes)
        .map_err(|source| Error::OpenIo {
            path: path.to_path_buf(),
            source,
        })?;
    // Some producers treat the 8-byte prefix as `(u32 size, u32 reserved)` and may write a
    // non-zero reserved high DWORD. Compute ciphertext length for a plausibility check and fall
    // back to the low DWORD when the combined u64 is not sensible.
    let base = encrypted_package_stream
        .seek(SeekFrom::Current(0))
        .map_err(|source| Error::OpenIo {
            path: path.to_path_buf(),
            source,
        })?;
    let end = encrypted_package_stream
        .seek(SeekFrom::End(0))
        .map_err(|source| Error::OpenIo {
            path: path.to_path_buf(),
            source,
        })?;
    encrypted_package_stream
        .seek(SeekFrom::Start(base))
        .map_err(|source| Error::OpenIo {
            path: path.to_path_buf(),
            source,
        })?;
    let ciphertext_len = end.saturating_sub(base);
    let plaintext_len = parse_encrypted_package_size_prefix_bytes(len_bytes, Some(ciphertext_len));
    if !encrypted_package_plaintext_len_is_plausible(plaintext_len, ciphertext_len) {
        return Err(Error::UnsupportedOoxmlEncryption {
            path: path.to_path_buf(),
            version_major,
            version_minor,
        });
    }

    // Wrap the OLE stream so offset 0 corresponds to the start of ciphertext (after the length
    // header).
    struct CiphertextStream<R> {
        inner: R,
        base: u64,
    }

    impl<R: std::io::Read> std::io::Read for CiphertextStream<R> {
        fn read(&mut self, buf: &mut [u8]) -> std::io::Result<usize> {
            self.inner.read(buf)
        }
    }

    impl<R: std::io::Read + std::io::Seek> std::io::Seek for CiphertextStream<R> {
        fn seek(&mut self, pos: SeekFrom) -> std::io::Result<u64> {
            let end_inner = match pos {
                SeekFrom::End(_) => Some(self.inner.seek(SeekFrom::End(0))?),
                _ => None,
            };

            let cur_inner = self.inner.seek(SeekFrom::Current(0))?;
            let cur = cur_inner
                .checked_sub(self.base)
                .ok_or_else(|| std::io::Error::new(std::io::ErrorKind::InvalidInput, "invalid ciphertext base offset"))?;

            let new_pos: i128 = match pos {
                SeekFrom::Start(n) => n as i128,
                SeekFrom::Current(off) => cur as i128 + off as i128,
                SeekFrom::End(off) => {
                    let end = end_inner
                        .expect("end_inner computed above")
                        .checked_sub(self.base)
                        .ok_or_else(|| {
                            std::io::Error::new(
                                std::io::ErrorKind::InvalidInput,
                                "invalid ciphertext end offset",
                            )
                        })?;
                    end as i128 + off as i128
                }
            };
            if new_pos < 0 {
                return Err(std::io::Error::new(
                    std::io::ErrorKind::InvalidInput,
                    "invalid seek to a negative position",
                ));
            }
            let new_pos_u64 = new_pos as u64;
            self.inner.seek(SeekFrom::Start(self.base + new_pos_u64))?;
            Ok(new_pos_u64)
        }
    }

    let ciphertext_reader = CiphertextStream {
        inner: encrypted_package_stream,
        base,
    };

    let reader = encrypted_ooxml::decrypted_package_reader(
        ciphertext_reader,
        plaintext_len,
        &encryption_info,
        password,
    )
    .map_err(|err| match err {
        encrypted_ooxml::DecryptError::InvalidPassword => Error::InvalidPassword {
            path: path.to_path_buf(),
        },
        encrypted_ooxml::DecryptError::UnsupportedVersion { major, minor } => {
            Error::UnsupportedOoxmlEncryption {
                path: path.to_path_buf(),
                version_major: major,
                version_minor: minor,
            }
        }
        encrypted_ooxml::DecryptError::InvalidInfo(_) | encrypted_ooxml::DecryptError::Io(_) => {
            Error::UnsupportedOoxmlEncryption {
                path: path.to_path_buf(),
                version_major,
                version_minor,
            }
        }
    })?;

    let workbook = xlsx::read_workbook_from_reader(reader).map_err(|source| Error::OpenXlsx {
        path: path.to_path_buf(),
        source,
    })?;
    Ok(Some(workbook))
}

/// Open a spreadsheet workbook with options.
///
/// This is the password-aware variant of [`open_workbook`]. When a password is provided and the
/// input is a legacy `.xls` workbook using BIFF `FILEPASS` encryption, this will attempt to decrypt
/// and import the workbook.
pub fn open_workbook_with_options(
    path: impl AsRef<Path>,
    opts: OpenOptions,
) -> Result<Workbook, Error> {
    let path = path.as_ref();

    // First, handle password-protected OOXML workbooks that are stored as OLE compound files
    // (`EncryptionInfo` + `EncryptedPackage` streams).
    //
    // When the `encrypted-workbooks` feature is enabled, attempt in-memory decryption. Otherwise,
    // surface an "unsupported encryption" error so callers don't assume a password will work.
    #[cfg(feature = "encrypted-workbooks")]
    {
        if let Some(bytes) =
            try_decrypt_ooxml_encrypted_package_from_path(path, opts.password.as_deref())?
        {
            return open_workbook_from_decrypted_ooxml_zip_bytes(path, bytes);
        }
    }

    if let Some(package_bytes) =
        maybe_read_plaintext_ooxml_package_from_encrypted_ole(path, opts.password.as_deref())?
    {
        return open_workbook_from_decrypted_ooxml_zip_bytes(path, package_bytes);
    }
    let ext = path
        .extension()
        .and_then(|s| s.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase();

    let format = if opts.password.is_some() {
        workbook_format_allow_encrypted_xls(path)?
    } else {
        match workbook_format(path) {
            Ok(fmt) => fmt,
            Err(Error::EncryptedWorkbook { .. }) => {
                return Err(Error::PasswordRequired {
                    path: path.to_path_buf(),
                });
            }
            Err(err) => return Err(err),
        }
    };

    match format {
        WorkbookFormat::Xlsx | WorkbookFormat::Xlsm => {
            let file = std::fs::File::open(path).map_err(|source| Error::OpenIo {
                path: path.to_path_buf(),
                source,
            })?;
            let package = xlsx::XlsxLazyPackage::from_file(path.to_path_buf(), file).map_err(
                |source| Error::OpenXlsx {
                    path: path.to_path_buf(),
                    source,
                },
            )?;
            Ok(Workbook::Xlsx(package))
        }
        WorkbookFormat::Xls => {
            match xls::import_xls_path_with_password(path, opts.password.as_deref()) {
                Ok(result) => Ok(Workbook::Xls(result)),
                Err(xls::ImportError::EncryptedWorkbook) => Err(Error::PasswordRequired {
                    path: path.to_path_buf(),
                }),
                Err(xls::ImportError::InvalidPassword) => Err(Error::InvalidPassword {
                    path: path.to_path_buf(),
                }),
                Err(xls::ImportError::UnsupportedEncryption(scheme)) => {
                    Err(Error::UnsupportedEncryption {
                        path: path.to_path_buf(),
                        kind: scheme,
                    })
                }
                Err(xls::ImportError::Decrypt(message)) => Err(Error::UnsupportedEncryption {
                    path: path.to_path_buf(),
                    kind: format!(
                        "legacy `.xls` FILEPASS encryption metadata is invalid: {message}"
                    ),
                }),
                Err(source) => Err(Error::OpenXls {
                    path: path.to_path_buf(),
                    source,
                }),
            }
        }
        WorkbookFormat::Xlsb => {
            xlsb::XlsbWorkbook::open(path)
                .map(Workbook::Xlsb)
                .map_err(|source| Error::OpenXlsb {
                    path: path.to_path_buf(),
                    source,
                })
        }
        WorkbookFormat::Csv => {
            let file = std::fs::File::open(path).map_err(|source| Error::OpenIo {
                path: path.to_path_buf(),
                source,
            })?;
            let reader = std::io::BufReader::new(file);

            let stem = path.file_stem().and_then(|s| s.to_str()).unwrap_or("");
            let sheet_name = formula_model::validate_sheet_name(stem)
                .ok()
                .map(|_| stem.to_string())
                .unwrap_or_else(|| sanitize_sheet_name(stem));

            let mut workbook = formula_model::Workbook::new();
            import_csv_into_workbook(
                &mut workbook,
                sheet_name,
                reader,
                CsvOptions::default(),
            )
            .map_err(|source| Error::OpenCsv {
                path: path.to_path_buf(),
                source,
            })?;

            Ok(Workbook::Model(workbook))
        }
        WorkbookFormat::Parquet => open_parquet_model_workbook(path).map(Workbook::Model),
        _ => Err(Error::UnsupportedExtension {
            path: path.to_path_buf(),
            extension: ext.to_string(),
        }),
    }
}

#[cfg(not(feature = "parquet"))]
fn open_parquet_model_workbook(path: &Path) -> Result<formula_model::Workbook, Error> {
    Err(Error::ParquetSupportNotEnabled {
        path: path.to_path_buf(),
    })
}

#[cfg(feature = "parquet")]
fn open_parquet_model_workbook(path: &Path) -> Result<formula_model::Workbook, Error> {
    use formula_model::CellRef;
    use std::sync::Arc;

    let table = formula_columnar::parquet::read_parquet_to_columnar(path).map_err(|source| {
        Error::OpenParquet {
            path: path.to_path_buf(),
            source: Box::new(source),
        }
    })?;

    let mut workbook = formula_model::Workbook::new();

    // Match CSV behavior: prefer the file stem if it is already a valid Excel sheet name,
    // otherwise sanitize it.
    let stem = path.file_stem().and_then(|s| s.to_str()).unwrap_or("");
    let sheet_name = formula_model::validate_sheet_name(stem)
        .ok()
        .map(|_| stem.to_string())
        .unwrap_or_else(|| sanitize_sheet_name(stem));

    let sheet_id = workbook
        .add_sheet(sheet_name)
        .or_else(|_| workbook.add_sheet("Sheet1"))
        .expect("Sheet1 is always a valid sheet name");

    let sheet = workbook
        .sheet_mut(sheet_id)
        .expect("sheet must exist immediately after add_sheet");
    sheet.set_columnar_table(CellRef::new(0, 0), Arc::new(table));

    Ok(workbook)
}

/// Save a workbook to disk.
///
/// Notes:
/// - [`Workbook::Xlsx`] is saved by writing the underlying OPC package back out (via
///   [`formula_xlsx::XlsxLazyPackage`]), preserving unknown parts. Unchanged ZIP entries are
///   typically preserved via a streaming raw-copy path rather than being regenerated.
/// - [`Workbook::Xls`] is exported as `.xlsx` (writing `.xls` is out of scope).
/// - [`Workbook::Xlsb`] can be saved losslessly back to `.xlsb` (package copy),
///   or exported to `.xlsx` depending on the output extension.
/// - [`Workbook::Model`] is exported as `.xlsx` via [`formula_xlsx::write_workbook`].
pub fn save_workbook(workbook: &Workbook, path: impl AsRef<Path>) -> Result<(), Error> {
    let path = path.as_ref();
    let ext = path
        .extension()
        .and_then(|s| s.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase();

    match workbook {
        Workbook::Xlsx(package) => match ext.as_str() {
            "xlsx" | "xlsm" | "xltx" | "xltm" | "xlam" => {
                let kind =
                    xlsx::WorkbookKind::from_extension(&ext).expect("handled by match arm above");

                let mut out = package.clone();

                // If we're saving a workbook with any macro-capable content to a macro-free
                // extension (e.g. `.xlsx`/`.xltx`), strip macro parts/relationships so we don't
                // produce a macro-enabled workbook in disguise (which Excel refuses to open).
                //
                // Macro-capable surfaces include:
                // - classic VBA (`xl/vbaProject.bin`)
                // - Excel 4.0 macro sheets (`xl/macrosheets/**`)
                // - legacy dialog sheets (`xl/dialogsheets/**`)
                let should_strip_macros = kind.is_macro_free() && out.macro_presence().any();
                if should_strip_macros {
                    // When stripping macros, use the macro-stripper's built-in `[Content_Types].xml`
                    // rewrite logic (including removing macro part overrides) rather than layering
                    // a separate workbook-kind patch on top of the original macro-enabled content
                    // types.
                    out.remove_vba_project_with_kind(kind)
                        .map_err(|source| Error::SaveXlsxPackage {
                            path: path.to_path_buf(),
                            source,
                        })?;
                } else {
                    out.enforce_workbook_kind(kind)
                        .map_err(|source| Error::SaveXlsxPackage {
                            path: path.to_path_buf(),
                            source,
                        })?;
                }

                let res = atomic_write(path, |file| out.write_to(file));
                match res {
                    Ok(()) => Ok(()),
                    Err(AtomicWriteError::Io(source)) => Err(Error::SaveIo {
                        path: path.to_path_buf(),
                        source,
                    }),
                    Err(AtomicWriteError::Writer(source)) => Err(Error::SaveXlsxPackage {
                        path: path.to_path_buf(),
                        source,
                    }),
                }
            }
            other => Err(Error::UnsupportedExtension {
                path: path.to_path_buf(),
                extension: other.to_string(),
            }),
        },
        Workbook::Xls(result) => match ext.as_str() {
            "xlsx" | "xltx" | "xltm" | "xlam" => {
                let kind =
                    xlsx::WorkbookKind::from_extension(&ext).expect("handled by match arm above");
                let res = atomic_write(path, |file| {
                    xlsx::write_workbook_to_writer_with_kind(&result.workbook, file, kind)
                });
                match res {
                    Ok(()) => Ok(()),
                    Err(AtomicWriteError::Io(source)) => Err(Error::SaveIo {
                        path: path.to_path_buf(),
                        source,
                    }),
                    Err(AtomicWriteError::Writer(source)) => Err(Error::SaveXlsxExport {
                        path: path.to_path_buf(),
                        source,
                    }),
                }
            }
            other => Err(Error::UnsupportedExtension {
                path: path.to_path_buf(),
                extension: other.to_string(),
            }),
        },
        Workbook::Xlsb(wb) => match ext.as_str() {
            "xlsb" => {
                let res = atomic_write(path, |file| wb.save_as_to_writer(file));
                match res {
                    Ok(()) => Ok(()),
                    Err(AtomicWriteError::Io(source)) => Err(Error::SaveIo {
                        path: path.to_path_buf(),
                        source,
                    }),
                    Err(AtomicWriteError::Writer(source)) => Err(Error::SaveXlsbPackage {
                        path: path.to_path_buf(),
                        source,
                    }),
                }
            }
            "xlsx" | "xltx" | "xltm" | "xlam" => {
                let kind =
                    xlsx::WorkbookKind::from_extension(&ext).expect("handled by match arm above");
                let model = xlsb_to_model_workbook(wb).map_err(|source| Error::SaveXlsbExport {
                    path: path.to_path_buf(),
                    source,
                })?;
                let res = atomic_write(path, |file| {
                    xlsx::write_workbook_to_writer_with_kind(&model, file, kind)
                });
                match res {
                    Ok(()) => Ok(()),
                    Err(AtomicWriteError::Io(source)) => Err(Error::SaveIo {
                        path: path.to_path_buf(),
                        source,
                    }),
                    Err(AtomicWriteError::Writer(source)) => Err(Error::SaveXlsxExport {
                        path: path.to_path_buf(),
                        source,
                    }),
                }
            }
            other => Err(Error::UnsupportedExtension {
                path: path.to_path_buf(),
                extension: other.to_string(),
            }),
        },
        Workbook::Model(model) => match ext.as_str() {
            "xlsx" | "xltx" | "xltm" | "xlam" => {
                let kind =
                    xlsx::WorkbookKind::from_extension(&ext).expect("handled by match arm above");
                let res = atomic_write(path, |file| {
                    xlsx::write_workbook_to_writer_with_kind(model, file, kind)
                });
                match res {
                    Ok(()) => Ok(()),
                    Err(AtomicWriteError::Io(source)) => Err(Error::SaveIo {
                        path: path.to_path_buf(),
                        source,
                    }),
                    Err(AtomicWriteError::Writer(source)) => Err(Error::SaveXlsxExport {
                        path: path.to_path_buf(),
                        source,
                    }),
                }
            }
            other => Err(Error::UnsupportedExtension {
                path: path.to_path_buf(),
                extension: other.to_string(),
            }),
        },
    }
}

#[cfg(feature = "encrypted-workbooks")]
fn save_workbook_encrypted_ooxml(
    workbook: &Workbook,
    path: &Path,
    password: &str,
    preserved_ole: &formula_office_crypto::OleEntries,
) -> Result<(), Error> {
    use std::io::{Cursor, Write as _};

    let ext = path
        .extension()
        .and_then(|s| s.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase();

    // Serialize the workbook to an OOXML ZIP package in memory (avoid writing plaintext to disk).
    let zip_bytes: Vec<u8> = match workbook {
        Workbook::Xlsx(package) => match ext.as_str() {
            "xlsx" | "xlsm" | "xltx" | "xltm" | "xlam" => {
                let kind =
                    xlsx::WorkbookKind::from_extension(&ext).expect("handled by match arm above");

                let mut out = package.clone();
                let should_strip_macros = kind.is_macro_free() && out.macro_presence().any();
                if should_strip_macros {
                    out.remove_vba_project_with_kind(kind)
                        .map_err(|source| Error::SaveXlsxPackage {
                            path: path.to_path_buf(),
                            source,
                        })?;
                } else {
                    out.enforce_workbook_kind(kind)
                        .map_err(|source| Error::SaveXlsxPackage {
                            path: path.to_path_buf(),
                            source,
                        })?;
                }

                out.write_to_bytes().map_err(|source| Error::SaveXlsxPackage {
                    path: path.to_path_buf(),
                    source,
                })?
            }
            other => {
                return Err(Error::UnsupportedExtension {
                    path: path.to_path_buf(),
                    extension: other.to_string(),
                })
            }
        },
        Workbook::Xls(result) => match ext.as_str() {
            "xlsx" | "xltx" | "xltm" | "xlam" => {
                let kind = xlsx::WorkbookKind::from_extension(&ext)
                    .expect("handled by match arm above");
                let mut cursor = Cursor::new(Vec::new());
                xlsx::write_workbook_to_writer_with_kind(&result.workbook, &mut cursor, kind)
                    .map_err(|source| Error::SaveXlsxExport {
                        path: path.to_path_buf(),
                        source,
                    })?;
                cursor.into_inner()
            }
            other => {
                return Err(Error::UnsupportedExtension {
                    path: path.to_path_buf(),
                    extension: other.to_string(),
                })
            }
        },
        Workbook::Xlsb(wb) => match ext.as_str() {
            "xlsb" => {
                let mut cursor = Cursor::new(Vec::new());
                wb.save_as_to_writer(&mut cursor)
                    .map_err(|source| Error::SaveXlsbPackage {
                        path: path.to_path_buf(),
                        source,
                    })?;
                cursor.into_inner()
            }
            "xlsx" | "xltx" | "xltm" | "xlam" => {
                let kind = xlsx::WorkbookKind::from_extension(&ext)
                    .expect("handled by match arm above");
                let model = xlsb_to_model_workbook(wb).map_err(|source| Error::SaveXlsbExport {
                    path: path.to_path_buf(),
                    source,
                })?;
                let mut cursor = Cursor::new(Vec::new());
                xlsx::write_workbook_to_writer_with_kind(&model, &mut cursor, kind).map_err(
                    |source| Error::SaveXlsxExport {
                        path: path.to_path_buf(),
                        source,
                    },
                )?;
                cursor.into_inner()
            }
            other => {
                return Err(Error::UnsupportedExtension {
                    path: path.to_path_buf(),
                    extension: other.to_string(),
                })
            }
        },
        Workbook::Model(model) => match ext.as_str() {
            "xlsx" | "xltx" | "xltm" | "xlam" => {
                let kind = xlsx::WorkbookKind::from_extension(&ext)
                    .expect("handled by match arm above");
                let mut cursor = Cursor::new(Vec::new());
                xlsx::write_workbook_to_writer_with_kind(model, &mut cursor, kind).map_err(
                    |source| Error::SaveXlsxExport {
                        path: path.to_path_buf(),
                        source,
                    },
                )?;
                cursor.into_inner()
            }
            other => {
                return Err(Error::UnsupportedExtension {
                    path: path.to_path_buf(),
                    extension: other.to_string(),
                })
            }
        },
    };

    let ole_bytes = formula_office_crypto::encrypt_package_to_ole_with_entries(
        &zip_bytes,
        password,
        formula_office_crypto::EncryptOptions::default(),
        Some(preserved_ole),
    )
    .map_err(|source| Error::SaveOoxmlEncryption {
        path: path.to_path_buf(),
        source,
    })?;

    let res = atomic_write(path, |file| file.write_all(&ole_bytes));
    match res {
        Ok(()) => Ok(()),
        Err(AtomicWriteError::Io(source)) => Err(Error::SaveIo {
            path: path.to_path_buf(),
            source,
        }),
        Err(AtomicWriteError::Writer(source)) => Err(Error::SaveIo {
            path: path.to_path_buf(),
            source,
        }),
    }
}

fn xlsb_error_code_to_model_error(code: u8) -> formula_model::ErrorValue {
    use core::str::FromStr;
    use formula_model::ErrorValue;

    xlsb::errors::xlsb_error_literal(code)
        .and_then(|lit| ErrorValue::from_str(lit).ok())
        .unwrap_or(ErrorValue::Unknown)
}

fn xlsb_to_model_workbook(wb: &xlsb::XlsbWorkbook) -> Result<formula_model::Workbook, xlsb::Error> {
    use formula_model::{
        normalize_formula_text, CalculationMode, CellRef, CellValue, DateSystem,
        DefinedNameScope, SheetVisibility, Style, Workbook as ModelWorkbook,
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
                    CellValue::Error(xlsb_error_code_to_model_error(code)),
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

            // Best-effort phonetic guide (furigana) extraction. This metadata is stored in XLSB
            // "wide strings" as a trailing phonetic/extended block that `formula-xlsb` preserves
            // on the parsed cell record.
            if let Some(phonetic) = cell
                .preserved_string
                .as_ref()
                .and_then(|s| s.phonetic_text())
            {
                let mut model_cell = sheet.cell(cell_ref).cloned().unwrap_or_default();
                model_cell.phonetic = Some(phonetic);
                sheet.set_cell(cell_ref, model_cell);
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
    use super::parse_encrypted_package_size_prefix_bytes;
    use super::xlsb_error_code_to_model_error;
    use super::xlsb_to_model_workbook;
    use formula_model::{CellRef, CellValue, DateSystem, ErrorValue};
    use std::io::{Cursor, Write};
    use std::path::Path;
    use zip::write::FileOptions;
    use zip::{CompressionMethod, ZipWriter};

    fn biff12_record(id: u32, payload: &[u8]) -> Vec<u8> {
        let mut out = Vec::new();
        crate::xlsb::biff12_varint::write_record_id(&mut out, id).expect("write record id");
        crate::xlsb::biff12_varint::write_record_len(&mut out, payload.len() as u32)
            .expect("write record len");
        out.extend_from_slice(payload);
        out
    }

    fn write_utf16_string(out: &mut Vec<u8>, s: &str) {
        let units: Vec<u16> = s.encode_utf16().collect();
        out.extend_from_slice(&(units.len() as u32).to_le_bytes());
        for u in units {
            out.extend_from_slice(&u.to_le_bytes());
        }
    }

    fn build_minimal_xlsb_with_parts(sheet_bin: &[u8], shared_strings_bin: &[u8]) -> Vec<u8> {
        // workbook.bin containing only a single BrtSheet followed by BrtEndSheets.
        // BrtSheet record data:
        //   [state_flags:u32][sheet_id:u32][relId:XLWideString][name:XLWideString]
        const BRT_SHEET: u32 = 0x009C;
        const BRT_END_SHEETS: u32 = 0x0090;

        let mut sheet_rec = Vec::new();
        sheet_rec.extend_from_slice(&0u32.to_le_bytes()); // flags/state
        sheet_rec.extend_from_slice(&1u32.to_le_bytes()); // sheet id
        write_utf16_string(&mut sheet_rec, "rId1");
        write_utf16_string(&mut sheet_rec, "Sheet1");

        let workbook_bin = [
            biff12_record(BRT_SHEET, &sheet_rec),
            biff12_record(BRT_END_SHEETS, &[]),
        ]
        .concat();

        // Minimal workbook relationships: Id->Target mapping (Type omitted; parser tolerates this).
        let workbook_rels = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.bin"/></Relationships>"#;

        let cursor = Cursor::new(Vec::new());
        let mut zip = ZipWriter::new(cursor);
        let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

        zip.start_file("xl/workbook.bin", options.clone()).unwrap();
        zip.write_all(&workbook_bin).unwrap();

        zip.start_file("xl/_rels/workbook.bin.rels", options.clone())
            .unwrap();
        zip.write_all(workbook_rels).unwrap();

        zip.start_file("xl/worksheets/sheet1.bin", options.clone())
            .unwrap();
        zip.write_all(sheet_bin).unwrap();

        zip.start_file("xl/sharedStrings.bin", options).unwrap();
        zip.write_all(shared_strings_bin).unwrap();

        zip.finish().unwrap().into_inner()
    }

    #[test]
    fn parse_encrypted_package_size_prefix_prefers_low_dword_when_u64_is_implausible() {
        // Some producers encode the 8-byte `EncryptedPackage` size prefix as:
        //   [u32 size][u32 reserved]
        // which yields a non-zero high DWORD but should still be interpreted as the low DWORD when
        // the combined u64 is not plausible for the ciphertext length.
        let len_lo: u32 = 100;
        let len_hi: u32 = 1;
        let mut prefix = [0u8; 8];
        prefix[..4].copy_from_slice(&len_lo.to_le_bytes());
        prefix[4..].copy_from_slice(&len_hi.to_le_bytes());

        let parsed = parse_encrypted_package_size_prefix_bytes(prefix, Some(200));
        assert_eq!(parsed, len_lo as u64);
    }

    #[test]
    fn parse_encrypted_package_size_prefix_uses_full_u64_when_plausible() {
        let len_lo: u32 = 100;
        let len_hi: u32 = 1;
        let mut prefix = [0u8; 8];
        prefix[..4].copy_from_slice(&len_lo.to_le_bytes());
        prefix[4..].copy_from_slice(&len_hi.to_le_bytes());

        let parsed = parse_encrypted_package_size_prefix_bytes(prefix, Some(5_000_000_000));
        assert_eq!(parsed, (len_lo as u64) | ((len_hi as u64) << 32));
    }

    #[test]
    fn parse_encrypted_package_size_prefix_without_ciphertext_len_prefers_low_dword_for_compat() {
        let len_lo: u32 = 1234;
        let len_hi: u32 = 1;
        let mut prefix = [0u8; 8];
        prefix[..4].copy_from_slice(&len_lo.to_le_bytes());
        prefix[4..].copy_from_slice(&len_hi.to_le_bytes());

        let parsed = parse_encrypted_package_size_prefix_bytes(prefix, None);
        assert_eq!(parsed, len_lo as u64);
    }

    #[test]
    fn parse_encrypted_package_size_prefix_returns_u64_when_high_dword_is_zero() {
        let len_lo: u32 = 42;
        let len_hi: u32 = 0;
        let mut prefix = [0u8; 8];
        prefix[..4].copy_from_slice(&len_lo.to_le_bytes());
        prefix[4..].copy_from_slice(&len_hi.to_le_bytes());

        let parsed = parse_encrypted_package_size_prefix_bytes(prefix, Some(10));
        assert_eq!(parsed, len_lo as u64);
    }

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
    fn xlsb_to_model_maps_biff12_error_codes() {
        // Patch an existing XLSB fixture in-place (at the worksheet record level) so we can test
        // both value cells (`BrtCellBoolErr`) and cached formula results (`BrtFmlaError`).
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../formula-xlsb/tests/fixtures/simple.xlsb"
        ));
        let wb = crate::xlsb::XlsbWorkbook::open(fixture_path).expect("open xlsb fixture");

        let dir = tempfile::tempdir().expect("temp dir");
        let patched_path = dir.path().join("errors.xlsb");

        let mut edits = Vec::new();
        let base_row = 50u32;
        for (i, (code, expected)) in [
            (0x00, ErrorValue::Null),
            (0x07, ErrorValue::Div0),
            (0x0F, ErrorValue::Value),
            (0x17, ErrorValue::Ref),
            (0x1D, ErrorValue::Name),
            (0x24, ErrorValue::Num),
            (0x2A, ErrorValue::NA),
            (0x2B, ErrorValue::GettingData),
            (0x2C, ErrorValue::Spill),
            (0x2D, ErrorValue::Calc),
            (0x2E, ErrorValue::Field),
            (0x2F, ErrorValue::Connect),
            (0x30, ErrorValue::Blocked),
            (0x31, ErrorValue::Unknown),
        ]
        .into_iter()
        .enumerate()
        {
            let row = base_row + i as u32;

            // Plain error value.
            edits.push(crate::xlsb::CellEdit {
                row,
                col: 0,
                new_value: crate::xlsb::CellValue::Error(code),
                new_style: None,
                clear_formula: false,
                new_formula: None,
                new_rgcb: None,
                new_formula_flags: None,
                shared_string_index: None,
            });

            // Formula that evaluates to the error literal as a constant (PtgErr).
            edits.push(crate::xlsb::CellEdit {
                row,
                col: 1,
                new_value: crate::xlsb::CellValue::Error(code),
                new_style: None,
                clear_formula: false,
                new_formula: Some(vec![0x1C, code]),
                new_rgcb: None,
                new_formula_flags: None,
                shared_string_index: None,
            });

            // Validate after conversion.
            let _ = expected;
        }

        wb.save_with_cell_edits(&patched_path, 0, &edits)
            .expect("save patched xlsb");

        let patched = crate::xlsb::XlsbWorkbook::open(&patched_path).expect("open patched xlsb");
        let model = xlsb_to_model_workbook(&patched).expect("convert to model");
        let sheet = model.sheet_by_name("Sheet1").expect("Sheet1 missing");

        for (i, (_code, expected)) in [
            (0x00, ErrorValue::Null),
            (0x07, ErrorValue::Div0),
            (0x0F, ErrorValue::Value),
            (0x17, ErrorValue::Ref),
            (0x1D, ErrorValue::Name),
            (0x24, ErrorValue::Num),
            (0x2A, ErrorValue::NA),
            (0x2B, ErrorValue::GettingData),
            (0x2C, ErrorValue::Spill),
            (0x2D, ErrorValue::Calc),
            (0x2E, ErrorValue::Field),
            (0x2F, ErrorValue::Connect),
            (0x30, ErrorValue::Blocked),
            (0x31, ErrorValue::Unknown),
        ]
        .into_iter()
        .enumerate()
        {
            let row = base_row + i as u32;
            assert_eq!(
                sheet.value(CellRef::new(row, 0)),
                CellValue::Error(expected),
                "value cell row={row} expected={expected}"
            );
            assert_eq!(
                sheet.value(CellRef::new(row, 1)),
                CellValue::Error(expected),
                "formula cached value row={row} expected={expected}"
            );

            // Formula text should also round-trip through decode (no leading '=').
            let formula = sheet
                .formula(CellRef::new(row, 1))
                .expect("formula cell should have formula text");
            assert_eq!(formula, expected.as_str());
        }
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

    #[test]
    fn xlsb_to_model_preserves_shared_string_phonetic_bytes() {
        // Shared strings (`sharedStrings.bin`) record ids (subset):
        // - BrtSST    0x009F
        // - BrtSI     0x0013
        // - BrtSSTEnd 0x00A0
        const BRT_SST: u32 = 0x009F;
        const BRT_SI: u32 = 0x0013;
        const BRT_SST_END: u32 = 0x00A0;

        // Worksheet record ids (subset):
        // - BrtBeginSheetData 0x0091
        // - BrtEndSheetData   0x0092
        // - BrtRow            0x0000
        // - BrtCellIsst       0x0007
        const BRT_SHEETDATA: u32 = 0x0091;
        const BRT_SHEETDATA_END: u32 = 0x0092;
        const BRT_ROW: u32 = 0x0000;
        const BRT_STRING: u32 = 0x0007;

        let phonetic_text = "PHO_MARKER_123";
        let mut phonetic_bytes = Vec::new();
        write_utf16_string(&mut phonetic_bytes, phonetic_text);

        // Build sharedStrings.bin with a single SI that has the phonetic bit set.
        // BrtSI payload:
        //   [flags:u8][text:XLWideString][phonetic tail bytes...]
        let mut si_payload = Vec::new();
        si_payload.push(0x02); // phonetic flag
        write_utf16_string(&mut si_payload, "Base");
        si_payload.extend_from_slice(&phonetic_bytes);

        let mut shared_strings_bin = Vec::new();
        shared_strings_bin.extend_from_slice(&biff12_record(
            BRT_SST,
            &[1u32.to_le_bytes(), 1u32.to_le_bytes()].concat(),
        ));
        shared_strings_bin.extend_from_slice(&biff12_record(BRT_SI, &si_payload));
        shared_strings_bin.extend_from_slice(&biff12_record(BRT_SST_END, &[]));

        // Build sheet1.bin with A1 as a shared string (`isst=0`).
        let mut sheet_bin = Vec::new();
        sheet_bin.extend_from_slice(&biff12_record(BRT_SHEETDATA, &[]));
        sheet_bin.extend_from_slice(&biff12_record(BRT_ROW, &0u32.to_le_bytes()));

        let mut cell_payload = Vec::new();
        cell_payload.extend_from_slice(&0u32.to_le_bytes()); // col
        cell_payload.extend_from_slice(&0u32.to_le_bytes()); // style
        cell_payload.extend_from_slice(&0u32.to_le_bytes()); // isst
        sheet_bin.extend_from_slice(&biff12_record(BRT_STRING, &cell_payload));

        sheet_bin.extend_from_slice(&biff12_record(BRT_SHEETDATA_END, &[]));

        let xlsb_bytes = build_minimal_xlsb_with_parts(&sheet_bin, &shared_strings_bin);
        let tmp = tempfile::NamedTempFile::new().expect("temp file");
        std::fs::write(tmp.path(), xlsb_bytes).expect("write temp xlsb");

        let wb = crate::xlsb::XlsbWorkbook::open(tmp.path()).expect("open xlsb");
        let model = xlsb_to_model_workbook(&wb).expect("convert to model");
        let sheet = model.sheet_by_name("Sheet1").expect("Sheet1 missing");

        let cell = sheet
            .cell(CellRef::from_a1("A1").unwrap())
            .expect("A1 missing");
        assert_eq!(cell.value, CellValue::String("Base".to_string()));
        assert_eq!(cell.phonetic.as_deref(), Some(phonetic_text));
    }

    #[test]
    fn xlsb_error_cell_values_map_to_model_error_values() {
        for (code, expected) in [
            (0x00, ErrorValue::Null),
            (0x07, ErrorValue::Div0),
            (0x0F, ErrorValue::Value),
            (0x17, ErrorValue::Ref),
            (0x1D, ErrorValue::Name),
            (0x24, ErrorValue::Num),
            (0x2A, ErrorValue::NA),
            (0x2B, ErrorValue::GettingData),
            (0x2C, ErrorValue::Spill),
            (0x2D, ErrorValue::Calc),
            (0x2E, ErrorValue::Field),
            (0x2F, ErrorValue::Connect),
            (0x30, ErrorValue::Blocked),
            (0x31, ErrorValue::Unknown),
        ] {
            assert_eq!(
                xlsb_error_code_to_model_error(code),
                expected,
                "xlsb error code {code:#04x} should map to {expected:?}"
            );
        }
    }
}
