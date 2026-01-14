use std::borrow::Cow;
use std::path::{Path, PathBuf};

use encoding_rs::{UTF_16BE, UTF_16LE, WINDOWS_1252};
use formula_fs::{atomic_write, AtomicWriteError};
use formula_model::import::{
    import_csv_into_workbook, CsvImportError, CsvOptions,
};
use formula_model::sanitize_sheet_name;
use std::io::{Read, Seek};
pub use formula_xls as xls;
pub use formula_xlsb as xlsb;
pub use formula_xlsx as xlsx;

pub mod offcrypto;
mod encryption_info;
pub use encryption_info::{extract_agile_encryption_info_xml, EncryptionInfoXmlError};
mod rc4_cryptoapi;
pub use rc4_cryptoapi::{Rc4CryptoApiDecryptReader, Rc4CryptoApiEncryptedPackageError};
mod ms_offcrypto;

const OLE_MAGIC: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
const PARQUET_MAGIC: [u8; 4] = *b"PAR1";
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
        "unsupported encrypted OOXML workbook `{path}`: EncryptionInfo version {version_major}.{version_minor} is not supported"
    )]
    UnsupportedOoxmlEncryption {
        path: PathBuf,
        version_major: u16,
        version_minor: u16,
    },
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
#[derive(Debug, Copy, Clone, PartialEq, Eq)]
pub enum WorkbookEncryptionKind {
    /// An Office-encrypted OOXML package stored in an OLE container via the `EncryptionInfo` +
    /// `EncryptedPackage` streams (e.g. a password-protected `.xlsx`).
    OoxmlOleEncryptedPackage,
    /// A legacy BIFF workbook stream containing a `FILEPASS` record (open password).
    XlsFilepass,
    /// An OLE compound file that appears to be encrypted, but doesn't match the known workbook
    /// encryption patterns.
    UnknownOleEncrypted,
}

#[derive(Debug, Copy, Clone, PartialEq, Eq)]
pub struct WorkbookEncryptionInfo {
    pub kind: WorkbookEncryptionKind,
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
        // Decryption spec note (MS-OFFCRYPTO): `EncryptedPackage` is segmented (0x1000 plaintext
        // chunks) with a per-segment IV derived from the salt + segment index; see
        // `docs/offcrypto-standard-encryptedpackage.md`.
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
            if matches!(ole_workbook_stream_starts_with_biff_bof(&mut ole), Some(false)) {
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

/// Detect whether a workbook is password-protected/encrypted.
///
/// This is best-effort and conservative:
/// - Returns `Ok(Some(..))` only when encryption markers are found.
/// - Returns `Ok(None)` for non-OLE files, OLE files without encryption markers, and malformed OLE
///   containers that can't be parsed.
pub fn detect_workbook_encryption(
    path: impl AsRef<Path>,
) -> Result<Option<WorkbookEncryptionInfo>, Error> {
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

    let has_encryption_info = stream_exists(&mut ole, "EncryptionInfo");
    let has_encrypted_package = stream_exists(&mut ole, "EncryptedPackage");

    if has_encryption_info || has_encrypted_package {
        let kind = if has_encryption_info && has_encrypted_package {
            WorkbookEncryptionKind::OoxmlOleEncryptedPackage
        } else {
            WorkbookEncryptionKind::UnknownOleEncrypted
        };
        return Ok(Some(WorkbookEncryptionInfo { kind }));
    }

    if ole_workbook_has_biff_filepass_record(&mut ole) {
        return Ok(Some(WorkbookEncryptionInfo {
            kind: WorkbookEncryptionKind::XlsFilepass,
        }));
    }

    Ok(None)
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
            if matches!(ole_workbook_stream_starts_with_biff_bof(&mut ole), Some(false)) {
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

/// Open a spreadsheet workbook from disk directly into a [`formula_model::Workbook`], providing a
/// password for encrypted OOXML workbooks when needed.
///
/// This behaves like [`open_workbook_model`], but can surface an [`Error::InvalidPassword`] when a
/// password is provided but incorrect.
pub fn open_workbook_model_with_password(
    path: impl AsRef<Path>,
    password: Option<&str>,
) -> Result<formula_model::Workbook, Error> {
    use std::fs::File;
    use std::io::BufReader;

    let path = path.as_ref();
    if let Some(err) = encrypted_ooxml_error_from_path(path, password) {
        return Err(err);
    }

    // If no password was provided, preserve the existing open path (including the current format
    // classification behaviour for legacy `.xls`).
    let Some(password) = password else {
        return open_workbook_model(path);
    };

    let ext = path
        .extension()
        .and_then(|s| s.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase();

    match workbook_format_allow_encrypted_xls(path)? {
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
        WorkbookFormat::Xls => match xls::import_xls_path_with_password(path, password) {
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

/// Open a spreadsheet workbook from disk, providing a password for encrypted OOXML workbooks when
/// needed.
///
/// This behaves like [`open_workbook`], but can surface an [`Error::InvalidPassword`] when a
/// password is provided but incorrect.
pub fn open_workbook_with_password(
    path: impl AsRef<Path>,
    password: Option<&str>,
) -> Result<Workbook, Error> {
    let path = path.as_ref();
    if let Some(err) = encrypted_ooxml_error_from_path(path, password) {
        return Err(err);
    }

    // If no password was provided, preserve the existing open path (including the current format
    // classification behaviour for legacy `.xls`).
    let Some(password) = password else {
        return open_workbook(path);
    };

    let ext = path
        .extension()
        .and_then(|s| s.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase();

    match workbook_format_allow_encrypted_xls(path)? {
        WorkbookFormat::Xlsx | WorkbookFormat::Xlsm => {
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
        WorkbookFormat::Xls => match xls::import_xls_path_with_password(path, password) {
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
    if ole.open_stream(name).is_ok() {
        return true;
    }
    let with_leading_slash = format!("/{name}");
    if ole.open_stream(&with_leading_slash).is_ok() {
        return true;
    }

    // Best-effort: some producers (or intermediate tools) can vary stream casing. Walk the
    // directory tree and compare paths case-insensitively.
    let target = name.trim_start_matches('/');
    ole.walk().any(|entry| {
        if !entry.is_stream() {
            return false;
        }
        let path = entry.path().to_string_lossy();
        let normalized = path.strip_prefix('/').unwrap_or(&path);
        normalized.eq_ignore_ascii_case(target)
    })
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
    let with_leading_slash = format!("/{name}");
    if let Some(buf) = read_candidate(ole, &with_leading_slash)? {
        return Ok(Some(buf));
    }

    let candidate = {
        let target = name.trim_start_matches('/');
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

    if password.is_none() {
        return Some(Error::PasswordRequired {
            path: path.to_path_buf(),
        });
    }

    // We currently don't implement OOXML decryption in `formula-io`, but we still want callers to
    // be able to surface a dedicated "wrong password" error when the user *did* provide one.
    Some(Error::InvalidPassword {
        path: path.to_path_buf(),
    })
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

/// Open a spreadsheet workbook based on file extension.
///
/// If you only need a [`formula_model::Workbook`] (data + formulas) and do not need full-fidelity
/// round-trip preservation, prefer [`open_workbook_model`] which uses streaming readers and avoids
/// inflating every package part into memory.
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
/// - [`Workbook::Xlsx`] is saved by writing the underlying OPC package back out,
///   preserving unknown parts.
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
                if kind.is_macro_free() && out.macro_presence().any() {
                    out.remove_vba_project()
                        .map_err(|source| Error::SaveXlsxPackage {
                            path: path.to_path_buf(),
                            source,
                        })?;
                }

                out.enforce_workbook_kind(kind)
                    .map_err(|source| Error::SaveXlsxPackage {
                        path: path.to_path_buf(),
                        source,
                    })?;

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
                let kind = xlsx::WorkbookKind::from_extension(&ext)
                    .expect("handled by match arm above");
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
                let kind = xlsx::WorkbookKind::from_extension(&ext)
                    .expect("handled by match arm above");
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
                let kind = xlsx::WorkbookKind::from_extension(&ext)
                    .expect("handled by match arm above");
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

fn xlsb_error_code_to_model_error(code: u8) -> formula_model::ErrorValue {
    use core::str::FromStr;
    use formula_model::ErrorValue;

    xlsb::errors::xlsb_error_literal(code)
        .and_then(|lit| ErrorValue::from_str(lit).ok())
        .unwrap_or(ErrorValue::Unknown)
}

fn xlsb_to_model_workbook(wb: &xlsb::XlsbWorkbook) -> Result<formula_model::Workbook, xlsb::Error> {
    use formula_model::{
        normalize_formula_text, CalculationMode, CellRef, CellValue, DateSystem, DefinedNameScope,
        SheetVisibility, Style, Workbook as ModelWorkbook,
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
    use super::xlsb_error_code_to_model_error;
    use super::xlsb_to_model_workbook;
    use formula_model::{CellRef, CellValue, DateSystem, ErrorValue};
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
                new_formula: None,
                new_rgcb: None,
                new_formula_flags: None,
                shared_string_index: None,
                new_style: None,
            });

            // Formula that evaluates to the error literal as a constant (PtgErr).
            edits.push(crate::xlsb::CellEdit {
                row,
                col: 1,
                new_value: crate::xlsb::CellValue::Error(code),
                new_formula: Some(vec![0x1C, code]),
                new_rgcb: None,
                new_formula_flags: None,
                shared_string_index: None,
                new_style: None,
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
