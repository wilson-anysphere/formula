//! Legacy Excel `.xls` (BIFF5/BIFF8) import support.
//!
//! This importer is intentionally best-effort: BIFF contains many features that
//! aren't representable in [`formula_model`]. We load sheets, cell values,
//! formulas (as text), merged-cell regions, basic row/column size/visibility
//! metadata, hyperlinks, basic outline (grouping) metadata, and legacy cell
//! comments/notes ([`formula_model::CommentKind::Note`]) on worksheets where
//! available. We also attempt to preserve workbook/worksheet view state where
//! possible (active tab, workbook window geometry/state, frozen panes, zoom,
//! selection, and grid/headings/zero visibility flags). Basic workbook- and
//! worksheet-level protection settings are also imported when present.
//!
//! When present, workbook- and sheet-scoped defined names (named ranges) are also
//! imported. Defined-name formula (`rgce`) decoding is best-effort and may emit
//! warnings for unsupported tokens.
//!
//! Note import is best-effort and intentionally lossy:
//! - Comment box geometry/visibility/formatting is not preserved
//! - Only plain text + author (when available) are imported
//! - Threaded comments are not supported in `.xls`
//! - Notes inside merged regions are anchored to the merged region's top-left cell
//!
//! We also extract workbook-global styles (including number format codes) and the
//! workbook date system (1900 vs 1904) when possible. Anything else is preserved
//! as metadata/warnings. In particular, comment parsing may emit warnings when
//! BIFF NOTE/OBJ/TXO records are malformed or incomplete (and such comments may
//! be skipped).

use std::collections::{HashMap, HashSet};
use std::io::{Cursor, Read, Write};
use std::path::{Path, PathBuf};

use calamine::{Data, Reader, Sheet, SheetType, SheetVisible, Xls};
use formula_model::{
    normalize_formula_text, sheet_name_eq_case_insensitive, CellRef, CellValue, ColRange, Comment,
    CommentAuthor, CommentKind, DefinedNameScope, ErrorValue, HyperlinkTarget, PrintTitles, Range,
    RowRange, SheetAutoFilter, SheetVisibility, Style, TabColor, Workbook, EXCEL_MAX_COLS,
    EXCEL_MAX_ROWS, EXCEL_MAX_SHEET_NAME_LEN, XLNM_FILTER_DATABASE,
};
use thiserror::Error;

mod biff;
mod ct;
mod decrypt;
pub use decrypt::DecryptError;
pub mod diagnostics;
mod formula_rewrite;

#[derive(Clone, Copy, Debug)]
struct DateTimeStyleIds {
    date: u32,
    datetime: u32,
    time: u32,
    duration: u32,
}

impl DateTimeStyleIds {
    fn new(workbook: &mut Workbook) -> Self {
        let date = workbook.intern_style(Style {
            number_format: Some("m/d/yy".to_string()),
            ..Default::default()
        });
        let datetime = workbook.intern_style(Style {
            number_format: Some("m/d/yy h:mm:ss".to_string()),
            ..Default::default()
        });
        let time = workbook.intern_style(Style {
            number_format: Some("h:mm:ss".to_string()),
            ..Default::default()
        });
        let duration = workbook.intern_style(Style {
            number_format: Some("[h]:mm:ss".to_string()),
            ..Default::default()
        });

        Self {
            date,
            datetime,
            time,
            duration,
        }
    }

    fn style_for_excel_datetime(self, dt: &calamine::ExcelDateTime) -> u32 {
        if dt.is_duration() {
            return self.duration;
        }

        // Calamine tells us the cell should be interpreted as a date/time (as
        // opposed to a raw number) but does not preserve the exact number format
        // string. We attempt to recover the precise format code via BIFF parsing,
        // but fall back to a best-effort heuristic when it isn't available.
        let serial = dt.as_f64();
        let frac = serial.abs().fract();

        if serial.abs() < 1.0 && frac != 0.0 {
            return self.time;
        }

        if frac == 0.0 {
            return self.date;
        }

        self.datetime
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SourceFormat {
    Xls,
}

impl SourceFormat {
    /// Default format to use when saving a workbook opened from this source.
    ///
    /// Writing legacy BIFF `.xls` files is out of scope; legacy imports default
    /// to `.xlsx`.
    pub const fn default_save_extension(self) -> &'static str {
        match self {
            SourceFormat::Xls => "xlsx",
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ImportSource {
    pub path: PathBuf,
    pub format: SourceFormat,
}

impl ImportSource {
    pub fn default_save_extension(&self) -> &'static str {
        self.format.default_save_extension()
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
/// Non-fatal warning emitted during `.xls` import.
///
/// The legacy BIFF importer is best-effort and continues when it encounters
/// malformed or unsupported records. These warnings are intended for user
/// visibility / diagnostics (for example, partial NOTE/OBJ/TXO comment records).
pub struct ImportWarning {
    pub message: String,
}

impl ImportWarning {
    fn new(message: impl Into<String>) -> Self {
        Self {
            message: message.into(),
        }
    }
}

/// Cap the total number of non-fatal warnings surfaced during `.xls` import.
///
/// Individual BIFF parsers already cap warnings in some cases, but the importer aggregates warnings
/// across many phases and sheets. A pathological workbook could otherwise produce an excessively
/// large `Vec<ImportWarning>` and very noisy UX.
#[cfg(not(test))]
const MAX_IMPORT_WARNINGS_TOTAL: usize = 1000;
// Keep unit tests fast by using a smaller cap.
#[cfg(test)]
const MAX_IMPORT_WARNINGS_TOTAL: usize = 200;

const IMPORT_WARNINGS_SUPPRESSED_MESSAGE: &str = "additional `.xls` import warnings suppressed";

/// Global warning accumulator for `.xls` import.
///
/// This enforces a hard cap on the number of warnings stored in memory across the entire workbook,
/// while still tracking how many warnings were *attempted* (for conservative heuristics).
#[derive(Debug, Default)]
struct ImportWarnings {
    warnings: Vec<ImportWarning>,
    total_count: usize,
    suppressed: bool,
}

impl ImportWarnings {
    fn new() -> Self {
        Self::default()
    }

    fn push(&mut self, warning: ImportWarning) {
        self.total_count = self.total_count.saturating_add(1);

        if self.suppressed {
            return;
        }

        if self.warnings.len() < MAX_IMPORT_WARNINGS_TOTAL {
            self.warnings.push(warning);
            return;
        }

        // First overflow: record a single sentinel warning and stop storing further warnings.
        self.warnings
            .push(ImportWarning::new(IMPORT_WARNINGS_SUPPRESSED_MESSAGE));
        self.suppressed = true;
    }

    fn push_msg(&mut self, msg: impl Into<String>) {
        self.push(ImportWarning::new(msg));
    }

    fn extend<I: IntoIterator<Item = ImportWarning>>(&mut self, iter: I) {
        for w in iter {
            self.push(w);
        }
    }

    /// Number of warnings *attempted* (including suppressed ones).
    fn count(&self) -> usize {
        self.total_count
    }

    fn into_vec(self) -> Vec<ImportWarning> {
        self.warnings
    }
}

fn push_import_warning(warnings: &mut ImportWarnings, msg: impl Into<String>, suppressed: &mut bool) {
    // Preserve existing call sites that still thread a `suppressed` flag: callers may set
    // `suppressed=true` to model the post-cap state in tests.
    if *suppressed {
        warnings.suppressed = true;
    }
    warnings.push_msg(msg);
    if warnings.suppressed {
        *suppressed = true;
    }
}

/// A merged cell range in the source workbook.
///
/// Merged regions are also imported into [`formula_model::Worksheet::merged_regions`]. This type
/// is retained for backward compatibility with callers that still expect merged-range metadata
/// from [`XlsImportResult`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MergedRange {
    pub sheet_name: String,
    pub range: Range,
}

#[derive(Debug)]
/// Result of importing a legacy `.xls` workbook.
///
/// The returned [`Workbook`] contains the successfully imported data (values,
/// formulas, merged regions, row/column/outline metadata, hyperlinks, defined
/// names, and legacy comments/notes where available). Any non-fatal issues
/// encountered during import are returned in [`XlsImportResult::warnings`].
pub struct XlsImportResult {
    pub workbook: Workbook,
    pub source: ImportSource,
    pub merged_ranges: Vec<MergedRange>,
    pub warnings: Vec<ImportWarning>,
}

#[derive(Debug, Error)]
pub enum ImportError {
    #[error("failed to read `.xls`: {0}")]
    Xls(#[from] calamine::XlsError),
    #[error("internal panic while reading `.xls`: {0}")]
    CalaminePanic(String),
    #[error(
        "password required: workbook is password-protected/encrypted; provide a password via `import_xls_path_with_password(.., Some(password))` or `import_xls_bytes_with_password(.., password)`"
    )]
    EncryptedWorkbook,
    #[error("invalid password")]
    InvalidPassword,
    #[error("unsupported `.xls` encryption scheme: {0}")]
    UnsupportedEncryption(String),
    #[error("failed to decrypt `.xls`: {0}")]
    Decrypt(DecryptError),
    #[error("invalid worksheet name: {0}")]
    InvalidSheetName(#[from] formula_model::SheetNameError),
}

impl From<decrypt::DecryptError> for ImportError {
    fn from(err: decrypt::DecryptError) -> Self {
        match err {
            decrypt::DecryptError::WrongPassword => ImportError::InvalidPassword,
            decrypt::DecryptError::UnsupportedEncryption(scheme) => {
                ImportError::UnsupportedEncryption(scheme)
            }
            other => ImportError::Decrypt(other),
        }
    }
}

fn panic_payload_to_string(payload: &(dyn std::any::Any + Send)) -> String {
    if let Some(s) = payload.downcast_ref::<&str>() {
        (*s).to_string()
    } else if let Some(s) = payload.downcast_ref::<String>() {
        s.clone()
    } else {
        "non-string panic payload".to_string()
    }
}

/// Catch panics from `calamine` (or other third-party parsing paths) and convert them into a
/// structured [`ImportError`].
///
/// This should only be used at third-party boundaries: a panic inside our own importer logic is
/// still a bug and should crash in tests.
fn catch_calamine_panic<T>(f: impl FnOnce() -> T) -> Result<T, ImportError> {
    match std::panic::catch_unwind(std::panic::AssertUnwindSafe(f)) {
        Ok(v) => Ok(v),
        Err(payload) => Err(ImportError::CalaminePanic(panic_payload_to_string(
            payload.as_ref(),
        ))),
    }
}

fn catch_calamine_panic_with_context<T>(
    context: &str,
    f: impl FnOnce() -> T,
) -> Result<T, ImportError> {
    catch_calamine_panic(f).map_err(|err| match err {
        ImportError::CalaminePanic(message) => {
            ImportError::CalaminePanic(format!("{context}: {message}"))
        }
        other => other,
    })
}

/// Import a legacy `.xls` workbook from disk.
///
/// The importer is intentionally best-effort and attempts to load a subset of
/// BIFF features into [`formula_model`], including legacy cell notes/comments
/// (`NOTE/OBJ/TXO`, imported as [`formula_model::CommentKind::Note`]). Any
/// malformed or unsupported records may produce warnings rather than failing
/// the import.
pub fn import_xls_path(path: impl AsRef<Path>) -> Result<XlsImportResult, ImportError> {
    import_xls_with_biff_reader(path.as_ref(), None, None, biff::read_workbook_stream_from_xls)
}

/// Import a legacy `.xls` workbook from disk with an optional password for BIFF `FILEPASS`
/// encryption.
///
/// This supports BIFF5/BIFF8 `.xls` workbooks that use the legacy `FILEPASS` record for workbook
/// password protection, including:
/// - BIFF8 XOR obfuscation (`wEncryptionType=0x0000`)
/// - BIFF8 RC4 "Standard Encryption" (`wEncryptionType=0x0001`, `wEncryptionSubType=0x0001`)
/// - BIFF8 RC4 CryptoAPI (`wEncryptionType=0x0001`, `wEncryptionSubType=0x0002`)
///
/// Notes on password handling:
/// - BIFF8 RC4 *Standard* uses only the first 15 UTF-16 code units of the password (Excel truncation).
/// - BIFF8 RC4 CryptoAPI uses the full password string (no 15-character truncation).
/// - Some writers may emit an empty password; pass `""` to decrypt such files.
pub fn import_xls_path_with_password(
    path: impl AsRef<Path>,
    password: Option<&str>,
) -> Result<XlsImportResult, ImportError> {
    import_xls_with_biff_reader(
        path.as_ref(),
        password,
        None,
        biff::read_workbook_stream_from_xls,
    )
}

/// Import a legacy `.xls` workbook from in-memory bytes.
///
/// This is intended for non-filesystem contexts (e.g. when a workbook is already loaded into
/// memory). The returned [`XlsImportResult::source`] path is set to `"<memory>"`.
pub fn import_xls_bytes(bytes: &[u8]) -> Result<XlsImportResult, ImportError> {
    import_xls_with_biff_reader(Path::new("<memory>"), None, Some(bytes), |_| {
        let cursor = Cursor::new(bytes);
        let mut comp = cfb::CompoundFile::open(cursor).map_err(|err| err.to_string())?;
        let mut stream = biff::open_xls_workbook_stream(&mut comp)?;
        let mut workbook_stream = Vec::new();
        stream
            .read_to_end(&mut workbook_stream)
            .map_err(|err| err.to_string())?;
        Ok(workbook_stream)
    })
}

/// Import a password-protected legacy `.xls` workbook from in-memory bytes.
///
/// This is intended for non-filesystem contexts (e.g. when a workbook is already loaded into
/// memory). The returned [`XlsImportResult::source`] path is set to `"<memory>"`.
pub fn import_xls_bytes_with_password(
    bytes: &[u8],
    password: &str,
) -> Result<XlsImportResult, ImportError> {
    import_xls_with_biff_reader(Path::new("<memory>"), Some(password), Some(bytes), |_| {
        let cursor = Cursor::new(bytes);
        let mut comp = cfb::CompoundFile::open(cursor).map_err(|err| err.to_string())?;
        let mut stream = biff::open_xls_workbook_stream(&mut comp)?;
        let mut workbook_stream = Vec::new();
        stream
            .read_to_end(&mut workbook_stream)
            .map_err(|err| err.to_string())?;
        Ok(workbook_stream)
    })
}

/// Import a legacy `.xls` workbook from disk while treating BIFF workbook-stream parsing as
/// unavailable.
///
/// This is intended for testing the importer's best-effort fallback paths.
#[doc(hidden)]
pub fn import_xls_path_without_biff(
    path: impl AsRef<Path>,
) -> Result<XlsImportResult, ImportError> {
    import_xls_with_biff_reader(path.as_ref(), None, None, |_| {
        Err("BIFF parsing disabled".to_string())
    })
}

fn import_xls_with_biff_reader(
    path: &Path,
    password: Option<&str>,
    fallback_xls_bytes: Option<&[u8]>,
    read_biff_workbook_stream: impl FnOnce(&Path) -> Result<Vec<u8>, String>,
) -> Result<XlsImportResult, ImportError> {
    let path = path.as_ref();
    // Best-effort: read the raw BIFF workbook stream up-front so we can detect
    // legacy `.xls` encryption (BIFF `FILEPASS`) before handing off to calamine.
    // Calamine does not support BIFF encryption and may return opaque parse
    // errors for password-protected workbooks.
    let mut warnings = ImportWarnings::new();
    let mut warnings_suppressed = false;
    let mut workbook_stream = match catch_calamine_panic_with_context(
        "reading `.xls` workbook stream",
        || read_biff_workbook_stream(path),
    ) {
        Ok(Ok(bytes)) => Some(bytes),
        Ok(Err(err)) => {
            push_import_warning(
                &mut warnings,
                format!("failed to read `.xls` workbook stream: {err}"),
                &mut warnings_suppressed,
            );
            None
        }
        Err(ImportError::CalaminePanic(message)) => {
            push_import_warning(
                &mut warnings,
                format!("panic while reading `.xls` workbook stream: {message}"),
                &mut warnings_suppressed,
            );
            None
        }
        Err(other) => return Err(other),
    };

    // Attempt to decrypt BIFF `FILEPASS` records when a password is provided. We do this before
    // running any BIFF record parsers so downstream metadata scans see plaintext.
    let needs_decrypt = workbook_stream
        .as_deref()
        .is_some_and(biff::records::workbook_globals_has_filepass_record);
    if needs_decrypt {
        let Some(password) = password else {
            return Err(ImportError::EncryptedWorkbook);
        };

        if let Some(workbook_stream) = workbook_stream.as_mut() {
            catch_calamine_panic_with_context("decrypting `.xls` workbook stream", || {
                decrypt::decrypt_biff_workbook_stream(workbook_stream, password)
            })??;

            // Decryption produces a plaintext workbook stream, but the `FILEPASS` record header
            // remains. Mask it so downstream BIFF parsers (and `calamine`) don't treat the stream
            // as encrypted and stop scanning at `FILEPASS`.
            //
            // Note: `decrypt_biff_workbook_stream` already masks `FILEPASS`, but keep this call
            // here as a safety net so future decryptors don't regress this requirement.
            biff::records::mask_workbook_globals_filepass_record_id_in_place(workbook_stream);
        }
    }

    let mut biff_version: Option<biff::BiffVersion> = None;
    let mut biff_codepage: Option<u16> = None;
    let mut biff_globals: Option<biff::globals::BiffWorkbookGlobals> = None;
    if let Some(workbook_stream) = workbook_stream.as_deref() {
        let detected_biff_version = biff::detect_biff_version(workbook_stream);
        let codepage = biff::parse_biff_codepage(workbook_stream);
        biff_version = Some(detected_biff_version);
        biff_codepage = Some(codepage);

        match biff::parse_biff_workbook_globals(
            workbook_stream,
            detected_biff_version,
            codepage,
        ) {
            Ok(globals) => {
                biff_globals = Some(globals);
            }
            Err(err) => push_import_warning(
                &mut warnings,
                format!("failed to import `.xls` workbook globals: {err}"),
                &mut warnings_suppressed,
            ),
        };
    }

    // `calamine` can panic when parsing BIFF8 defined-name `NAME` (0x0018) records that are split
    // across `CONTINUE` records. Calamine reads the `cce` formula length from the NAME header, but
    // (incorrectly) assumes the entire token stream lives in the *first* physical record fragment.
    //
    // When a NAME record is continued, `cce` can exceed the first fragment length and calamine
    // panics while slicing `rgce`.
    //
    // Work around this by sanitizing BIFF8 NAME records in the workbook stream before handing it
    // to calamine:
    // - Coalesce consecutive `CONTINUE` fragments into the NAME record and compact away the
    //   embedded CONTINUE headers so `payload[14..]` is contiguous.
    // - Best-effort: strip the extra 1-byte "continued string segment" flags prefix that appears
    //   at the start of continued Unicode-string segments, so calamine can read the full name
    //   string without truncation.
    // - Zero out `NAME.cce` so calamine doesn't attempt to slice/parse a potentially-continued
    //   `rgce` token stream.
    // - Clamp `NAME.cch` based on the bytes available (and the string encoding flag) to prevent
    //   unchecked slice panics on corrupt files.
    //
    // We still import defined names ourselves via BIFF parsing (including CONTINUE handling), so
    // calamine's defined-name formulas are not used for correctness here. However, calamine *does*
    // need the defined-name *names* table to decode `PtgName` tokens in worksheet formulas.
    //
    // The sanitizer therefore tries hard to keep NAME records parseable (by clamping fields and
    // coalescing CONTINUE fragments) rather than masking them. In rare cases where a NAME record
    // is so truncated that it can't be repaired safely without corrupting the stream, we mask it
    // to prevent panics (at the cost of potentially shifting name indices in calamine's table).
    let mut calamine_formulas_disabled = false;

    let mut workbook: Xls<_> = match workbook_stream.as_deref() {
        Some(stream) => {
            let detected_biff_version =
                biff_version.unwrap_or_else(|| biff::detect_biff_version(stream));

            // Calamine's continued-NAME panic workaround only applies to BIFF8 NAME records. Avoid
            // patching BIFF5 streams (different NAME layout) to keep `.xls` import best-effort.
            let sanitized = match detected_biff_version {
                biff::BiffVersion::Biff8 => sanitize_biff8_stream_for_calamine(stream),
                biff::BiffVersion::Biff5 => None,
            };
            let stream_for_calamine = sanitized.as_deref().unwrap_or(stream);

            let open_with_stream = |workbook_stream: &[u8]| -> Result<Xls<_>, ImportError> {
                let xls_bytes = build_in_memory_xls(workbook_stream)?;
                match catch_calamine_panic_with_context("opening `.xls` via calamine", || {
                    Xls::new(Cursor::new(xls_bytes))
                })? {
                    Ok(workbook) => Ok(workbook),
                    Err(err) => Err(ImportError::Xls(err)),
                }
            };

            match open_with_stream(stream_for_calamine) {
                Ok(workbook) => workbook,
                Err(ImportError::CalaminePanic(message))
                    if detected_biff_version == biff::BiffVersion::Biff8 =>
                {
                    // Calamine can panic while parsing `.xls` workbook streams when decoding BIFF8
                    // formulas (notably shared-formula materialization near BIFF8 row/col bounds).
                    //
                    // Prefer a best-effort recovery path: retry with BIFF8 formula token streams
                    // sanitized/disabled for calamine (we still import formulas from the raw BIFF
                    // stream ourselves).
                    push_import_warning(
                        &mut warnings,
                        format!(
                            "panic while opening `.xls` via calamine: {message}; retrying with BIFF8 formula decoding disabled"
                        ),
                        &mut warnings_suppressed,
                    );

                    let sanitized_formulas =
                        sanitize_biff8_formula_records_for_calamine(stream_for_calamine);
                    if sanitized_formulas.is_some() {
                        calamine_formulas_disabled = true;
                    }
                    let stream_for_calamine =
                        sanitized_formulas.as_deref().unwrap_or(stream_for_calamine);
                    open_with_stream(stream_for_calamine)?
                }
                Err(other) => return Err(other),
            }
        }
        None => {
            let bytes = match fallback_xls_bytes {
                Some(bytes) => bytes.to_vec(),
                None => std::fs::read(path)
                    .map_err(|err| ImportError::Xls(calamine::XlsError::Io(err)))?,
            };
            catch_calamine_panic_with_context("opening `.xls` via calamine", || {
                Xls::new(Cursor::new(bytes))
            })?
            .map_err(ImportError::Xls)?
        }
    };

    // We need to snapshot metadata (names, visibility, type) up-front because we
    // need mutable access to the workbook while iterating over ranges.
    let sheets: Vec<Sheet> = catch_calamine_panic_with_context("reading sheet metadata", || {
        workbook.sheets_metadata().to_vec()
    })?;
    // Snapshot defined names up-front because we need mutable access to the workbook while
    // iterating over ranges.
    let calamine_defined_names =
        catch_calamine_panic_with_context("reading defined names", || {
            workbook.defined_names().to_vec()
        })?;

    let mut out = Workbook::new();
    let mut used_sheet_names: Vec<String> = Vec::new();
    let mut merged_ranges = Vec::new();

    let mut xf_style_ids: Option<Vec<u32>> = None;
    let mut xf_is_interesting: Option<Vec<bool>> = None;
    let mut sheet_tab_colors: Option<Vec<Option<TabColor>>> = None;
    let mut workbook_active_tab: Option<u16> = None;
    let mut biff_sheets: Option<Vec<biff::BoundSheetInfo>> = None;
    let mut row_col_props: Option<Vec<biff::SheetRowColProperties>> = None;
    let mut cell_xf_indices: Option<Vec<HashMap<CellRef, u16>>> = None;
    let mut cell_xf_parse_failed: Option<Vec<bool>> = None;
    let mut filter_database_ranges: Option<HashMap<usize, Range>> = None;
    // Map output worksheet ids to their BIFF worksheet substream offsets. Used for best-effort
    // parsing of sheet-local metadata like AutoFilter criteria.
    let mut sheet_stream_offsets_by_sheet_id: HashMap<formula_model::WorksheetId, usize> =
        HashMap::new();
    // Map output worksheet ids to whether the sheet stream contained FILTERMODE. Used to decide
    // warning policy when criteria records are missing/unsupported.
    let mut sheet_filter_mode_by_sheet_id: HashMap<formula_model::WorksheetId, bool> =
        HashMap::new();
    let mut sst_phonetics: Option<Vec<Option<String>>> = None;

    if let Some(workbook_stream) = workbook_stream.as_deref() {
        let detected_biff_version =
            biff_version.unwrap_or_else(|| biff::detect_biff_version(workbook_stream));
        let codepage = biff_codepage.unwrap_or_else(|| biff::parse_biff_codepage(workbook_stream));
        biff_version.get_or_insert(detected_biff_version);
        biff_codepage.get_or_insert(codepage);
        out.codepage = codepage;

        if detected_biff_version == biff::BiffVersion::Biff8 {
            match biff::sst::parse_biff8_sst_phonetics(workbook_stream, codepage) {
                Ok(phonetics) => {
                    if !phonetics.is_empty() {
                        sst_phonetics = Some(phonetics);
                    }
                }
                Err(err) => push_import_warning(
                    &mut warnings,
                    format!("failed to parse `.xls` shared string phonetics: {err}"),
                    &mut warnings_suppressed,
                ),
            }
        }

        if let Some(mut globals) = biff_globals.take() {
            out.date_system = globals.date_system;
            if let Some(mode) = globals.calculation_mode {
                out.calc_settings.calculation_mode = mode;
            }
            if let Some(value) = globals.calculate_before_save {
                out.calc_settings.calculate_before_save = value;
            }
            if let Some(enabled) = globals.iterative_enabled {
                out.calc_settings.iterative.enabled = enabled;
            }
            if let Some(max_iterations) = globals.iterative_max_iterations {
                out.calc_settings.iterative.max_iterations = max_iterations;
            }
            if let Some(max_change) = globals.iterative_max_change {
                out.calc_settings.iterative.max_change = max_change;
            }
            if let Some(full_precision) = globals.full_precision {
                out.calc_settings.full_precision = full_precision;
            }
            out.workbook_protection = std::mem::take(&mut globals.workbook_protection);
            workbook_active_tab = globals.active_tab_index;
            // `Workbook.view.window` is optional metadata. Prefer any value already populated on the
            // model (e.g. future calamine support) over our best-effort BIFF parsing.
            if out.view.window.is_none() {
                out.view.window = globals.workbook_window.take();
            }
            warnings.extend(globals.warnings.drain(..).map(|w| ImportWarning::new(w)));
            if warnings.suppressed {
                warnings_suppressed = true;
            }
            sheet_tab_colors = Some(std::mem::take(&mut globals.sheet_tab_colors));

            let interesting = globals.xf_is_interesting_mask();
            xf_style_ids = Some(vec![0; interesting.len()]);
            xf_is_interesting = Some(interesting);
            biff_globals = Some(globals);
        }

        match biff::parse_biff_bound_sheets(workbook_stream, detected_biff_version, codepage) {
            Ok(sheets) => biff_sheets = Some(sheets),
            Err(err) => push_import_warning(
                &mut warnings,
                format!("failed to import `.xls` sheet metadata: {err}"),
                &mut warnings_suppressed,
            ),
        }

        // AutoFilter ranges are stored in a built-in workbook/worksheet defined name
        // (`_FilterDatabase`). Excel files in the wild use both workbook-scope and local-scope
        // definitions; decode them from the BIFF workbook globals stream.
        let sheet_count_for_autofilter = biff_sheets.as_ref().map(|s| s.len());
        match biff::parse_biff_filter_database_ranges(
            workbook_stream,
            detected_biff_version,
            codepage,
            sheet_count_for_autofilter,
        ) {
            Ok(parsed) => {
                let biff::ParsedFilterDatabaseRanges {
                    by_sheet,
                    warnings: biff_warnings,
                } = parsed;
                if !by_sheet.is_empty() {
                    filter_database_ranges = Some(by_sheet);
                }
                for w in biff_warnings {
                    push_import_warning(
                        &mut warnings,
                        format!("failed to import `.xls` autofilter: {w}"),
                        &mut warnings_suppressed,
                    );
                }
            }
            Err(err) => push_import_warning(
                &mut warnings,
                format!("failed to import `.xls` autofilter ranges: {err}"),
                &mut warnings_suppressed,
            ),
        }

        if let Some(sheets) = biff_sheets.as_ref() {
            let mut props_by_sheet = Vec::with_capacity(sheets.len());
            for sheet in sheets {
                if sheet.offset >= workbook_stream.len() {
                    push_import_warning(
                        &mut warnings,
                        format!(
                            "failed to import `.xls` row/column properties for BIFF sheet index {} (`{}`): out-of-bounds stream offset {}",
                            props_by_sheet.len(),
                            sheet.name,
                            sheet.offset
                        ),
                        &mut warnings_suppressed,
                    );
                    props_by_sheet.push(biff::SheetRowColProperties::default());
                    continue;
                }

                match biff::parse_biff_sheet_row_col_properties(
                    workbook_stream,
                    sheet.offset,
                    codepage,
                ) {
                    Ok(mut props) => {
                        for warning in props.warnings.drain(..) {
                            push_import_warning(
                                &mut warnings,
                                format!(
                                    "failed to fully import `.xls` row/column properties for BIFF sheet index {} (`{}`): {warning}",
                                    props_by_sheet.len(),
                                    sheet.name
                                ),
                                &mut warnings_suppressed,
                            );
                        }
                        props_by_sheet.push(props);
                    }
                    Err(parse_err) => {
                        push_import_warning(
                            &mut warnings,
                            format!(
                                "failed to import `.xls` row/column properties for BIFF sheet index {} (`{}`): {parse_err}",
                                props_by_sheet.len(),
                                sheet.name
                            ),
                            &mut warnings_suppressed,
                        );
                        props_by_sheet.push(biff::SheetRowColProperties::default());
                    }
                }
            }
            row_col_props = Some(props_by_sheet);

            if let Some(mask) = xf_is_interesting.as_deref() {
                // Even if the workbook contains no non-General number formats, scan for
                // out-of-range XF indices so corrupt files still surface a warning.
                if !mask.is_empty() {
                    let mut cell_xfs_by_sheet = Vec::with_capacity(sheets.len());
                    let mut parse_failed_by_sheet = Vec::with_capacity(sheets.len());
                    for sheet in sheets {
                        if sheet.offset >= workbook_stream.len() {
                            push_import_warning(
                                &mut warnings,
                                format!(
                                    "failed to import `.xls` cell styles for BIFF sheet index {} (`{}`): out-of-bounds stream offset {}",
                                    cell_xfs_by_sheet.len(),
                                    sheet.name,
                                    sheet.offset
                                ),
                                &mut warnings_suppressed,
                            );
                            cell_xfs_by_sheet.push(HashMap::new());
                            parse_failed_by_sheet.push(true);
                            continue;
                        }

                        match biff::parse_biff_sheet_cell_xf_indices_filtered(
                            workbook_stream,
                            sheet.offset,
                            Some(mask),
                        ) {
                            Ok(xfs) => {
                                cell_xfs_by_sheet.push(xfs);
                                parse_failed_by_sheet.push(false);
                            }
                            Err(parse_err) => {
                                push_import_warning(
                                    &mut warnings,
                                    format!(
                                        "failed to import `.xls` cell styles for BIFF sheet index {} (`{}`): {parse_err}",
                                        cell_xfs_by_sheet.len(),
                                        sheet.name
                                    ),
                                    &mut warnings_suppressed,
                                );
                                cell_xfs_by_sheet.push(HashMap::new());
                                parse_failed_by_sheet.push(true);
                            }
                        }
                    }

                    cell_xf_indices = Some(cell_xfs_by_sheet);
                    cell_xf_parse_failed = Some(parse_failed_by_sheet);
                }
            }
        }

        // Intern styles only for XF indices that were referenced by at least one "interesting" cell
        // record (including BLANK/MUL* records). This is best-effort; out-of-range XF indices are
        // ignored here but still preserved in `cell_xf_indices` so callers can surface warnings.
        if let (Some(globals), Some(style_ids), Some(cell_xfs_by_sheet)) = (
            biff_globals.as_ref(),
            xf_style_ids.as_mut(),
            cell_xf_indices.as_ref(),
        ) {
            let mut used = vec![false; style_ids.len()];
            for sheet_xfs in cell_xfs_by_sheet {
                for &xf_idx in sheet_xfs.values() {
                    let idx = xf_idx as usize;
                    if idx < used.len() {
                        used[idx] = true;
                    }
                }
            }

            // ROW/COLINFO records can reference default formats (ixfe) even when no individual cell
            // record uses those XF indices. Include them in the "used" set so their styles are
            // interned and can be applied to `Worksheet.row_properties[*].style_id` /
            // `Worksheet.col_properties[*].style_id`.
            if let Some(props_by_sheet) = row_col_props.as_ref() {
                for props in props_by_sheet {
                    for row_props in props.rows.values() {
                        if let Some(xf_idx) = row_props.xf_index {
                            let idx = xf_idx as usize;
                            if idx < used.len() {
                                used[idx] = true;
                            }
                        }
                    }
                    for col_props in props.cols.values() {
                        if let Some(xf_idx) = col_props.xf_index {
                            let idx = xf_idx as usize;
                            if idx < used.len() {
                                used[idx] = true;
                            }
                        }
                    }
                }
            }

            for (xf_idx, style) in globals.resolve_styles_for_used_mask(&used) {
                if style == Style::default() {
                    continue;
                }
                style_ids[xf_idx] = out.intern_style(style);
            }
        }

        // Drop the BIFF globals cache once we've built the XF -> style_id mapping.
        drop(biff_globals.take());
    }

    // `calamine` may surface date/time serials as `Data::DateTime`, but it does not preserve the
    // exact BIFF number format code. We attempt to recover the precise format via BIFF parsing; if
    // that fails we fall back to a small set of heuristics (interned lazily).
    let mut date_time_styles: Option<DateTimeStyleIds> = None;

    let (sheet_mapping, mapping_warning) =
        reconcile_biff_sheet_mapping(&sheets, biff_sheets.as_deref());
    if let Some(message) = mapping_warning {
        push_import_warning(&mut warnings, message, &mut warnings_suppressed);
    }

    // Precompute BIFF sheet names (BoundSheet order) for `rgce` decoding. These may differ from the
    // final imported sheet names if we sanitize, but those references are rewritten later.
    let sheet_names_by_biff_idx_for_rgce: Vec<String> = biff_sheets
        .as_ref()
        .map(|sheets| sheets.iter().map(|s| s.name.clone()).collect())
        .unwrap_or_default();

    // Best-effort workbook-global context for BIFF8 formula decoding (SUPBOOK + EXTERNSHEET + NAME
    // metadata for `PtgName`).
    let workbook_rgce_decoder: Option<biff::worksheet_formulas::WorksheetFormulaDecoder> =
        if let (Some(workbook_stream), Some(biff_version), Some(codepage)) =
            (workbook_stream.as_deref(), biff_version, biff_codepage)
        {
            (biff_version == biff::BiffVersion::Biff8).then(|| {
                biff::worksheet_formulas::WorksheetFormulaDecoder::new(
                    workbook_stream,
                    biff_version,
                    codepage,
                    &sheet_names_by_biff_idx_for_rgce,
                )
            })
        } else {
            None
        };

    let mut final_sheet_names_by_idx: Vec<String> = Vec::with_capacity(sheets.len());
    // Track worksheet ids so BIFF `itab` scopes can be mapped to the output model.
    let mut sheet_ids_by_calamine_idx: Vec<formula_model::WorksheetId> =
        Vec::with_capacity(sheets.len());

    // Best-effort BIFF8 formula decoding context used to override calamine formula strings in a
    // handful of tricky cases (notably shared formulas and 3D areas with relative flags).
    let mut biff_rgce_sheet_names: Vec<String> = Vec::new();
    let mut biff_rgce_externsheet: Vec<biff::externsheet::ExternSheetEntry> = Vec::new();
    let mut biff_rgce_supbooks: Vec<biff::supbook::SupBookInfo> = Vec::new();
    let mut biff_rgce_defined_names: Vec<biff::rgce::DefinedNameMeta> = Vec::new();

    if let (Some(workbook_stream), Some(biff_version), Some(codepage)) =
        (workbook_stream.as_deref(), biff_version, biff_codepage)
    {
        if biff_version == biff::BiffVersion::Biff8 {
            // Prefer BIFF BoundSheet names (BIFF sheet index order) when available; fall back to
            // calamine metadata order.
            if let Some(biff_sheets) = biff_sheets.as_deref() {
                biff_rgce_sheet_names = biff_sheets.iter().map(|s| s.name.clone()).collect();
            }
            if biff_rgce_sheet_names.is_empty() {
                biff_rgce_sheet_names = sheets.iter().map(|s| s.name.clone()).collect();
            }

            let externsheet_table = biff::externsheet::parse_biff_externsheet(
                workbook_stream,
                biff_version,
                codepage,
            );
            biff_rgce_externsheet = externsheet_table.entries;

            let supbook_table = biff::supbook::parse_biff8_supbook_table(workbook_stream, codepage);
            biff_rgce_supbooks = supbook_table.supbooks;

            let (metas, _) = biff::defined_names::parse_biff_defined_name_metas(
                workbook_stream,
                biff_version,
                codepage,
                &biff_rgce_sheet_names,
            );
            biff_rgce_defined_names = metas;
        }
    }
    for (sheet_idx, sheet_meta) in sheets.iter().enumerate() {
        let source_sheet_name = sheet_meta.name.clone();
        let biff_idx = sheet_mapping.get(sheet_idx).copied().flatten();

        let sheet_cell_xfs_raw =
            biff_idx.and_then(|biff_idx| cell_xf_indices.as_ref().and_then(|v| v.get(biff_idx)));

        let sheet_cell_xfs = sheet_cell_xfs_raw.filter(|map| !map.is_empty());

        // If BIFF styles are unavailable (or corrupt), fall back to heuristic date/time formats for
        // any `Data::DateTime` cells that would otherwise have no style.
        let sheet_has_out_of_range_xf = match (xf_style_ids.as_deref(), sheet_cell_xfs_raw) {
            (Some(style_ids), Some(map)) => map
                .values()
                .any(|&xf_idx| xf_idx as usize >= style_ids.len()),
            _ => false,
        };

        let sheet_xf_parse_failed = biff_idx
            .and_then(|biff_idx| cell_xf_parse_failed.as_ref().and_then(|v| v.get(biff_idx)))
            .copied()
            .unwrap_or(false);

        let sheet_needs_datetime_fallback = xf_style_ids.is_none()
            || sheet_cell_xfs_raw.is_none()
            || sheet_has_out_of_range_xf
            || sheet_xf_parse_failed;

        // Best-effort: calamine can panic on malformed `.xls` records. Treat panics as warnings
        // and continue importing other sheets/metadata.
        let value_range = match catch_calamine_panic_with_context(
            &format!("reading cell values for sheet `{source_sheet_name}`"),
            || workbook.worksheet_range(&source_sheet_name),
        ) {
            Ok(Ok(range)) => Some(range),
            Ok(Err(err)) => {
                push_import_warning(
                    &mut warnings,
                    format!("failed to read cell values for sheet `{source_sheet_name}`: {err}"),
                    &mut warnings_suppressed,
                );
                None
            }
            Err(ImportError::CalaminePanic(message)) => {
                push_import_warning(
                    &mut warnings,
                    format!(
                        "failed to read cell values for sheet `{source_sheet_name}`: calamine panicked: {message}"
                    ),
                    &mut warnings_suppressed,
                );
                None
            }
            Err(other) => return Err(other),
        };

        if sheet_needs_datetime_fallback && date_time_styles.is_none() {
            date_time_styles = Some(DateTimeStyleIds::new(&mut out));
        }

        let sheet_date_time_styles = if sheet_needs_datetime_fallback {
            date_time_styles
        } else {
            None
        };

        let (sheet_id, sheet_name) = match out.add_sheet(source_sheet_name.clone()) {
            Ok(sheet_id) => {
                used_sheet_names.push(source_sheet_name.clone());
                (sheet_id, source_sheet_name.clone())
            }
            Err(err) => {
                let mut candidate =
                    sanitize_sheet_name(&source_sheet_name, sheet_idx + 1, &used_sheet_names);
                let sheet_id = loop {
                    match out.add_sheet(candidate.clone()) {
                        Ok(sheet_id) => break sheet_id,
                        Err(_) => {
                            // If our best-effort sanitization still collides (e.g. due to
                            // case-insensitive comparisons), treat the candidate as taken and
                            // generate another.
                            let mut augmented = used_sheet_names.clone();
                            augmented.push(candidate);
                            candidate =
                                sanitize_sheet_name(&source_sheet_name, sheet_idx + 1, &augmented);
                        }
                    }
                };

                push_import_warning(
                    &mut warnings,
                    format!("sanitized sheet name `{source_sheet_name}` -> `{candidate}` ({err})"),
                    &mut warnings_suppressed,
                );
                used_sheet_names.push(candidate.clone());
                (sheet_id, candidate)
            }
        };
        final_sheet_names_by_idx.push(sheet_name.clone());
        sheet_ids_by_calamine_idx.push(sheet_id);

        // Worksheet print settings (page setup + margins + manual page breaks) are stored in the
        // worksheet BIFF substream. Parse and apply them before borrowing `sheet_mut` so we can
        // call `Workbook::set_*` helpers without running into borrow conflicts.
        //
        // Only implemented for BIFF8 currently; BIFF5 worksheet record layouts differ.
        if biff_version == Some(biff::BiffVersion::Biff8) {
            if let (Some(workbook_stream), Some(biff_idx)) = (workbook_stream.as_deref(), biff_idx)
            {
                if let Some(sheet_info) = biff_sheets.as_ref().and_then(|v| v.get(biff_idx)) {
                    if sheet_info.offset >= workbook_stream.len() {
                        push_import_warning(
                            &mut warnings,
                            format!(
                                "failed to import `.xls` print settings for BIFF sheet index {} (`{}`): out-of-bounds stream offset {}",
                                biff_idx, sheet_name, sheet_info.offset
                            ),
                            &mut warnings_suppressed,
                        );
                    } else {
                        match biff::parse_biff_sheet_print_settings(
                            workbook_stream,
                            sheet_info.offset,
                        ) {
                            Ok(mut parsed) => {
                                for warning in parsed.warnings.drain(..) {
                                    push_import_warning(
                                        &mut warnings,
                                        format!(
                                            "failed to fully import `.xls` print settings for BIFF sheet index {} (`{}`): {warning}",
                                            biff_idx, sheet_name
                                        ),
                                        &mut warnings_suppressed,
                                    );
                                }

                                // Even if the sheet has default page setup and no breaks, updating
                                // is harmless and ensures later print settings (print area/titles)
                                // share a single `SheetPrintSettings` entry when needed.
                                out.set_sheet_page_setup(
                                    sheet_id,
                                    parsed.page_setup.unwrap_or_default(),
                                );
                                out.set_manual_page_breaks(sheet_id, parsed.manual_page_breaks);
                            }
                            Err(err) => push_import_warning(
                                &mut warnings,
                                format!(
                                    "failed to import `.xls` print settings for BIFF sheet index {} (`{}`): {err}",
                                    biff_idx, sheet_name
                                ),
                                &mut warnings_suppressed,
                            ),
                        }
                    }
                }
            }
        }
        let sheet = out
            .sheet_mut(sheet_id)
            .expect("sheet id should exist immediately after add");

        let calamine_visibility = sheet_visible_to_visibility(sheet_meta.visible);
        let biff_visibility = biff_idx
            .and_then(|idx| biff_sheets.as_ref().and_then(|v| v.get(idx)))
            .and_then(|info| info.sheet_visibility);
        sheet.visibility = biff_visibility.unwrap_or(calamine_visibility);
        sheet.tab_color = biff_idx
            .and_then(|idx| sheet_tab_colors.as_ref().and_then(|v| v.get(idx)).cloned())
            .flatten();

        if sheet_meta.typ != SheetType::WorkSheet {
            push_import_warning(
                &mut warnings,
                format!(
                    "sheet `{sheet_name}` has unsupported type {:?}; importing as worksheet",
                    sheet_meta.typ
                ),
                &mut warnings_suppressed,
            );
        }

        if let (Some(workbook_stream), Some(biff_idx)) = (workbook_stream.as_deref(), biff_idx) {
            if let Some(sheet_info) = biff_sheets.as_ref().and_then(|v| v.get(biff_idx)) {
                if sheet_info.offset >= workbook_stream.len() {
                    push_import_warning(
                        &mut warnings,
                        format!(
                            "failed to import `.xls` view state for BIFF sheet index {} (`{}`): out-of-bounds stream offset {}",
                            biff_idx, sheet_name, sheet_info.offset
                        ),
                        &mut warnings_suppressed,
                    );
                } else {
                    match biff::parse_biff_sheet_view_state(workbook_stream, sheet_info.offset) {
                        Ok(mut view_state) => {
                            for warning in view_state.warnings.drain(..) {
                                push_import_warning(
                                    &mut warnings,
                                    warning,
                                    &mut warnings_suppressed,
                                );
                            }

                            if let Some(show) = view_state.show_grid_lines {
                                sheet.view.show_grid_lines = show;
                            }
                            if let Some(show) = view_state.show_headings {
                                sheet.view.show_headings = show;
                            }
                            if let Some(show) = view_state.show_zeros {
                                sheet.view.show_zeros = show;
                            }
                            if let Some(zoom) = view_state.zoom {
                                sheet.zoom = zoom;
                                sheet.view.zoom = zoom;
                            }
                            if let Some(pane) = view_state.pane {
                                sheet.frozen_rows = pane.frozen_rows;
                                sheet.frozen_cols = pane.frozen_cols;
                                sheet.view.pane = pane;
                            }
                            if let Some(selection) = view_state.selection {
                                sheet.view.selection = Some(selection);
                            }
                        }
                        Err(err) => push_import_warning(
                            &mut warnings,
                            format!(
                                "failed to import `.xls` view state for BIFF sheet index {} (`{}`): {err}",
                                biff_idx, sheet_name
                            ),
                            &mut warnings_suppressed,
                        ),
                    }

                    match biff::parse_biff_sheet_protection(workbook_stream, sheet_info.offset) {
                        Ok(mut protection) => {
                            for warning in protection.warnings.drain(..) {
                                push_import_warning(
                                    &mut warnings,
                                    warning,
                                    &mut warnings_suppressed,
                                );
                            }
                            sheet.sheet_protection = protection.protection;
                        }
                        Err(err) => push_import_warning(
                            &mut warnings,
                            format!(
                                "failed to import `.xls` sheet protection for BIFF sheet index {} (`{}`): {err}",
                                biff_idx, sheet_name
                            ),
                            &mut warnings_suppressed,
                        ),
                    }
                }
            }
        }

        if let Some(biff_idx) = biff_idx {
            // Cache the worksheet substream offset (BoundSheetInfo.offset) so later best-effort
            // parsers (e.g. AutoFilter criteria) can scan the sheet stream.
            if let Some(offset) = biff_sheets
                .as_ref()
                .and_then(|sheets| sheets.get(biff_idx))
                .map(|info| info.offset)
            {
                sheet_stream_offsets_by_sheet_id.insert(sheet.id, offset);
            }

            let filter_database_range = filter_database_ranges
                .as_ref()
                .and_then(|ranges| ranges.get(&biff_idx))
                .copied();

            let mut sheet_stream_autofilter_range: Option<Range> = None;

            if let Some(props) = row_col_props
                .as_ref()
                .and_then(|props_by_sheet| props_by_sheet.get(biff_idx))
            {
                sheet_filter_mode_by_sheet_id.insert(sheet.id, props.filter_mode);
                apply_row_col_properties(sheet, props);
                apply_outline_properties(sheet, props);
                apply_row_col_style_ids(
                    sheet,
                    props,
                    xf_style_ids.as_deref(),
                    &mut warnings,
                    &mut warnings_suppressed,
                    &sheet_name,
                );

                sheet_stream_autofilter_range = props.auto_filter_range;

                if props.filter_mode {
                    // BIFF `FILTERMODE` indicates that some rows are currently hidden by a filter.
                    //
                    // BIFF row metadata does not distinguish between user-hidden and filter-hidden
                    // rows. We classify hidden rows best-effort after the canonical AutoFilter
                    // range is finalized.
                    push_import_warning(
                        &mut warnings,
                        format!(
                            "sheet `{sheet_name}` has FILTERMODE record at offset {} (filtered rows); filter criteria and row visibility may not round-trip exactly on import",
                            props
                                .filter_mode_offset
                                .map(|o| o.to_string())
                                .unwrap_or_else(|| "unknown".to_string())
                        ),
                        &mut warnings_suppressed,
                    );
                }
            }

            if sheet.auto_filter.is_none() {
                // Prefer the canonical AutoFilter range from the `_FilterDatabase` defined name when
                // available, but fall back to best-effort inference from the worksheet substream
                // (AUTOFILTERINFO / FILTERMODE + DIMENSIONS).
                if let Some(range) = filter_database_range.or(sheet_stream_autofilter_range) {
                    sheet.auto_filter = Some(SheetAutoFilter {
                        range,
                        // Any filter criteria / sort state from the sheet stream is applied later,
                        // after the canonical AutoFilter range is finalized via defined-name based
                        // recovery.
                        filter_columns: Vec::new(),
                        sort_state: None,
                        raw_xml: Vec::new(),
                    });
                }
            }
        }

        // Merged regions: prefer calamine's parsed merge metadata, but fall back to scanning the
        // worksheet BIFF substream for `MERGEDCELLS` records when calamine provides none.
        let mut merge_ranges: Vec<Range> = Vec::new();
        if let Some(merge_cells) = catch_calamine_panic_with_context(
            &format!("reading merged cells for sheet `{source_sheet_name}`"),
            || workbook.worksheet_merge_cells(&source_sheet_name),
        )? {
            for dim in merge_cells {
                merge_ranges.push(Range::new(
                    CellRef::new(dim.start.0, dim.start.1),
                    CellRef::new(dim.end.0, dim.end.1),
                ));
            }
        }

        // Best-effort fallback when calamine does not surface any merged-cell ranges.
        if merge_ranges.is_empty() {
            if let (Some(workbook_stream), Some(biff_idx)) = (workbook_stream.as_deref(), biff_idx)
            {
                if let Some(sheet_info) = biff_sheets.as_ref().and_then(|s| s.get(biff_idx)) {
                    if sheet_info.offset >= workbook_stream.len() {
                        push_import_warning(
                            &mut warnings,
                            format!(
                                "failed to import `.xls` merged cells for sheet `{sheet_name}`: out-of-bounds stream offset {}",
                                sheet_info.offset
                            ),
                            &mut warnings_suppressed,
                        );
                    } else {
                        match biff::parse_biff_sheet_merged_cells(
                            workbook_stream,
                            sheet_info.offset,
                        ) {
                            Ok(mut parsed) => {
                                if !parsed.ranges.is_empty() {
                                    merge_ranges.append(&mut parsed.ranges);
                                }
                                for w in parsed.warnings.drain(..) {
                                    push_import_warning(
                                        &mut warnings,
                                        format!(
                                            "failed to import `.xls` merged cells for sheet `{sheet_name}`: {w}"
                                        ),
                                        &mut warnings_suppressed,
                                    );
                                }
                            }
                            Err(err) => push_import_warning(
                                &mut warnings,
                                format!(
                                    "failed to import `.xls` merged cells for sheet `{sheet_name}`: {err}"
                                ),
                                &mut warnings_suppressed,
                            ),
                        }
                    }
                }
            }
        }

        if !merge_ranges.is_empty() {
            let mut seen: HashSet<Range> = HashSet::new();
            for range in merge_ranges {
                if !seen.insert(range) {
                    continue;
                }

                // Populate the model's merged-region table so cell resolution matches Excel.
                if let Err(err) = sheet.merged_regions.add(range) {
                    push_import_warning(
                        &mut warnings,
                        format!(
                            "failed to add merged region `{range}` in sheet `{sheet_name}`: {err}"
                        ),
                        &mut warnings_suppressed,
                    );
                }

                // Back-compat: preserve merged metadata on the import result.
                merged_ranges.push(MergedRange {
                    sheet_name: sheet_name.clone(),
                    range,
                });
            }
        }

        if let Some(range) = value_range.as_ref() {
            let range_start = range.start().unwrap_or((0, 0));

            for (row, col, value) in range.used_cells() {
                let Some(cell_ref) = to_cell_ref(range_start, row, col) else {
                    push_import_warning(
                        &mut warnings,
                        format!(
                            "skipping out-of-bounds cell in sheet `{sheet_name}` at ({row},{col})"
                        ),
                        &mut warnings_suppressed,
                    );
                    continue;
                };

                let anchor = sheet.merged_regions.resolve_cell(cell_ref);
                let Some((value, mut style_id)) = convert_value(value, sheet_date_time_styles)
                else {
                    continue;
                };

                if let Some(resolved) =
                    style_id_for_cell_xf(xf_style_ids.as_deref(), sheet_cell_xfs, anchor)
                {
                    style_id = Some(resolved);
                }

                sheet.set_value(anchor, value);
                if let Some(style_id) = style_id {
                    sheet.set_style_id(anchor, style_id);
                }
            }
        }

        // Populate per-cell phonetic guide text (furigana) from the workbook SST `ExtRst` blocks.
        //
        // BIFF stores the phonetic metadata on the *shared string* entry (SST) and references it
        // from the worksheet via `LABELSST.isst`. We apply the extracted phonetic string to the
        // corresponding model cell, anchored to merged regions like other cell metadata.
        if let (Some(workbook_stream), Some(biff_idx), Some(sst_phonetics)) = (
            workbook_stream.as_deref(),
            biff_idx,
            sst_phonetics.as_ref(),
        ) {
            if let Some(sheet_info) = biff_sheets.as_ref().and_then(|s| s.get(biff_idx)) {
                if sheet_info.offset >= workbook_stream.len() {
                    push_import_warning(
                        &mut warnings,
                        format!(
                            "failed to import `.xls` phonetic guides for sheet `{sheet_name}`: out-of-bounds stream offset {}",
                            sheet_info.offset
                        ),
                        &mut warnings_suppressed,
                    );
                } else {
                    match biff::parse_biff_sheet_labelsst_indices(workbook_stream, sheet_info.offset)
                    {
                        Ok(map) => {
                            // Deterministic merge/anchor preference: apply anchor cells first.
                            let mut entries: Vec<(CellRef, u32)> = map.into_iter().collect();
                            entries.sort_by_key(|(cell, _)| {
                                let anchor = sheet.merged_regions.resolve_cell(*cell);
                                // (is_non_anchor, row, col)
                                (anchor != *cell, cell.row, cell.col)
                            });

                            for (cell_ref, isst) in entries {
                                let Some(phonetic) =
                                    sst_phonetics.get(isst as usize).and_then(|p| p.as_ref())
                                else {
                                    continue;
                                };

                                let anchor = sheet.merged_regions.resolve_cell(cell_ref);

                                // Do not create new cells purely for phonetic metadata.
                                let Some(cell) = sheet.cell_mut(anchor) else {
                                    continue;
                                };

                                // Prefer the anchor's own record over non-anchor cells inside a merge.
                                if anchor != cell_ref && cell.phonetic.is_some() {
                                    continue;
                                }

                                cell.phonetic = Some(phonetic.clone());
                            }
                        }
                        Err(err) => push_import_warning(
                            &mut warnings,
                            format!(
                                "failed to import `.xls` phonetic guides for sheet `{sheet_name}`: {err}"
                            ),
                            &mut warnings_suppressed,
                        ),
                    }
                }
            }
        }

        // Extract BIFF hyperlinks after merged regions have been populated so callers can resolve
        // anchors consistently with the model's merged-cell semantics.
        if let (Some(workbook_stream), Some(codepage), Some(biff_idx)) =
            (workbook_stream.as_deref(), biff_codepage, biff_idx)
        {
            if let Some(sheet_info) = biff_sheets.as_ref().and_then(|s| s.get(biff_idx)) {
                if sheet_info.offset >= workbook_stream.len() {
                    push_import_warning(
                        &mut warnings,
                        format!(
                            "failed to import `.xls` hyperlinks for sheet `{sheet_name}`: out-of-bounds stream offset {}",
                            sheet_info.offset
                        ),
                        &mut warnings_suppressed,
                    );
                } else {
                    match biff::parse_biff_sheet_hyperlinks(
                        workbook_stream,
                        sheet_info.offset,
                        codepage,
                    ) {
                        Ok(parsed) => {
                            let hyperlink_start = sheet.hyperlinks.len();
                            sheet.hyperlinks.extend(parsed.hyperlinks);
                            // Excel treats hyperlinks applied to a merged cell as applying to the
                            // entire merged region. When BIFF stores the hyperlink as a single-cell
                            // anchor inside a merge, expand it so `Worksheet::hyperlink_at` works
                            // for any cell inside the merged region.
                            if !sheet.merged_regions.is_empty()
                                && hyperlink_start < sheet.hyperlinks.len()
                            {
                                for link in sheet.hyperlinks[hyperlink_start..].iter_mut() {
                                    if link.range.is_single_cell() {
                                        if let Some(merged) =
                                            sheet.merged_regions.containing_range(link.range.start)
                                        {
                                            link.range = merged;
                                        }
                                    }
                                }
                            }
                            for w in parsed.warnings {
                                push_import_warning(
                                    &mut warnings,
                                    format!(
                                        "failed to import `.xls` hyperlink in sheet `{sheet_name}`: {w}"
                                    ),
                                    &mut warnings_suppressed,
                                );
                            }
                        }
                        Err(err) => push_import_warning(
                            &mut warnings,
                            format!("failed to import `.xls` hyperlinks for sheet `{sheet_name}`: {err}"),
                            &mut warnings_suppressed,
                        ),
                    }
                }
            }
        }

        // Extract legacy cell comments ("notes") encoded via NOTE/OBJ/TXO records in the sheet
        // BIFF substream. This runs after merged regions have been populated so
        // `Worksheet::add_comment` can normalize anchors inside merged regions.
        if let (Some(workbook_stream), Some(codepage), Some(biff_version), Some(biff_idx)) = (
            workbook_stream.as_deref(),
            biff_codepage,
            biff_version,
            biff_idx,
        ) {
            if let Some(sheet_info) = biff_sheets.as_ref().and_then(|s| s.get(biff_idx)) {
                if sheet_info.offset >= workbook_stream.len() {
                    push_import_warning(
                        &mut warnings,
                        format!(
                            "failed to import `.xls` note comments for sheet `{sheet_name}`: out-of-bounds stream offset {}",
                            sheet_info.offset
                        ),
                        &mut warnings_suppressed,
                    );
                } else {
                    match biff::parse_biff_sheet_notes(
                        workbook_stream,
                        sheet_info.offset,
                        biff_version,
                        codepage,
                    ) {
                        Ok(parsed) => {
                            for w in parsed.warnings {
                                push_import_warning(
                                    &mut warnings,
                                    format!(
                                        "failed to import `.xls` note comment in sheet `{sheet_name}`: {w}"
                                    ),
                                    &mut warnings_suppressed,
                                );
                            }
                            let notes = parsed.notes;

                            // Generate deterministic comment ids for BIFF NOTE comments so repeated
                            // imports of the same `.xls` file yield stable ids.
                            //
                            // NOTE records are linked to TXO text payloads via `obj_id`, but some
                            // files contain duplicate object ids or other collisions. In those
                            // cases, fall back to `Worksheet::add_comment`'s UUID generation by
                            // leaving the id empty and surface a warning.
                            let mut seen_obj_ids: HashMap<u16, CellRef> = HashMap::new();
                            let mut seen_ids: HashSet<String> = HashSet::new();
                            for note in notes {
                                let anchor = sheet.merged_regions.resolve_cell(note.cell);
                                let candidate_id =
                                    format!("xls-note:{}:{}", anchor.to_a1(), note.obj_id);

                                let mut collision = false;
                                if let Some(prev_cell) = seen_obj_ids.get(&note.obj_id).copied() {
                                    collision = true;
                                    push_import_warning(
                                        &mut warnings,
                                        format!(
                                            "duplicate `.xls` NOTE obj_id {} in sheet `{sheet_name}` (index {sheet_idx}) (cell {} already used at {}); generating random comment id",
                                            note.obj_id,
                                            anchor.to_a1(),
                                            prev_cell.to_a1(),
                                        ),
                                        &mut warnings_suppressed,
                                    );
                                }
                                if seen_ids.contains(&candidate_id) {
                                    collision = true;
                                    push_import_warning(
                                        &mut warnings,
                                        format!(
                                            "duplicate `.xls` NOTE id `{candidate_id}` in sheet `{sheet_name}` (index {sheet_idx}); generating random comment id",
                                        ),
                                        &mut warnings_suppressed,
                                    );
                                }

                                let id = if collision {
                                    String::new()
                                } else {
                                    seen_obj_ids.insert(note.obj_id, anchor);
                                    seen_ids.insert(candidate_id.clone());
                                    candidate_id
                                };

                                let mut comment = Comment {
                                    id,
                                    kind: CommentKind::Note,
                                    content: note.text,
                                    author: CommentAuthor {
                                        name: note.author,
                                        ..Default::default()
                                    },
                                    ..Default::default()
                                };

                                match sheet.add_comment(anchor, comment.clone()) {
                                    Ok(_) => {}
                                    Err(formula_model::CommentError::DuplicateCommentId(dup_id)) => {
                                        push_import_warning(
                                            &mut warnings,
                                            format!(
                                                "duplicate `.xls` comment id `{dup_id}` in sheet `{sheet_name}` (index {sheet_idx}); generating random comment id",
                                            ),
                                            &mut warnings_suppressed,
                                        );
                                        comment.id.clear();
                                        if let Err(err) = sheet.add_comment(anchor, comment) {
                                            push_import_warning(
                                                &mut warnings,
                                                format!(
                                                    "failed to import `.xls` note comment for sheet `{sheet_name}` (index {sheet_idx}) at {}: {err}",
                                                    anchor.to_a1(),
                                                ),
                                                &mut warnings_suppressed,
                                            );
                                        }
                                    }
                                    Err(err) => push_import_warning(
                                        &mut warnings,
                                        format!(
                                            "failed to import `.xls` note comment for sheet `{sheet_name}` (index {sheet_idx}) at {}: {err}",
                                            anchor.to_a1(),
                                        ),
                                        &mut warnings_suppressed,
                                    ),
                                }
                            }
                        }
                        Err(err) => push_import_warning(
                            &mut warnings,
                            format!("failed to import `.xls` note comments for sheet `{sheet_name}`: {err}"),
                            &mut warnings_suppressed,
                        ),
                    }
                }
            }
        }

        let mut calamine_formula_failed = calamine_formulas_disabled;

        if !calamine_formulas_disabled {
            // Best-effort: calamine's `.xls` formula decoding can error or panic on token streams it
            // doesn't understand. Treat both errors and panics as non-fatal and fall back to BIFF
            // parsing when possible.
            match catch_calamine_panic_with_context(
                &format!("reading formulas for sheet `{source_sheet_name}`"),
                || workbook.worksheet_formula(&source_sheet_name),
            ) {
                Ok(Ok(formula_range)) => {
                    let formula_start = formula_range.start().unwrap_or((0, 0));
                    for (row, col, formula) in formula_range.used_cells() {
                        let Some(cell_ref) = to_cell_ref(formula_start, row, col) else {
                            push_import_warning(
                                &mut warnings,
                                format!(
                                    "skipping out-of-bounds formula in sheet `{sheet_name}` at ({row},{col})"
                                ),
                                &mut warnings_suppressed,
                            );
                            continue;
                        };

                        // Calamine may surface BIFF8 Unicode strings with embedded NUL bytes (notably
                        // defined-name references via `PtgName`). Strip them so the formula text is
                        // parseable and stable across import paths.
                        let formula_clean;
                        let formula = if formula.contains('\0') {
                            formula_clean = formula.replace('\0', "");
                            formula_clean.as_str()
                        } else {
                            formula
                        };

                        let Some(normalized) = normalize_formula_text(formula) else {
                            continue;
                        };

                        let anchor = sheet.merged_regions.resolve_cell(cell_ref);
                        sheet.set_formula(anchor, Some(normalized));

                        if let Some(resolved) =
                            style_id_for_cell_xf(xf_style_ids.as_deref(), sheet_cell_xfs, anchor)
                        {
                            sheet.set_style_id(anchor, resolved);
                        }
                    }
                }
                Ok(Err(err)) => {
                    calamine_formula_failed = true;
                    push_import_warning(
                        &mut warnings,
                        format!("failed to read formulas for sheet `{sheet_name}`: {err}"),
                        &mut warnings_suppressed,
                    );
                }
                Err(ImportError::CalaminePanic(message)) => {
                    calamine_formula_failed = true;
                    push_import_warning(
                        &mut warnings,
                        format!(
                            "failed to read formulas for sheet `{sheet_name}`: calamine panicked: {message}"
                        ),
                        &mut warnings_suppressed,
                    );
                }
                Err(other) => return Err(other),
            }
        } else {
            // Calamine workbook open succeeded only after we sanitized formula record payloads; do
            // not trust `worksheet_formula` output.
            calamine_formula_failed = true;
        }

        // BIFF8 fallback/merge: calamine's `.xls` formula decoding can fail for token streams it
        // doesn't understand. When the raw workbook stream is available, decode formulas directly
        // from BIFF8 `FORMULA` records and merge results so a calamine failure doesn't drop all
        // formulas for the sheet.
        if let (Some(workbook_stream), Some(biff_version), Some(codepage), Some(biff_idx)) = (
            workbook_stream.as_deref(),
            biff_version,
            biff_codepage,
            biff_idx,
        ) {
            if biff_version == biff::BiffVersion::Biff8 {
                if let Some(sheet_info) = biff_sheets.as_ref().and_then(|s| s.get(biff_idx)) {
                    if sheet_info.offset >= workbook_stream.len() {
                        // Only surface this warning if we need BIFF formula import due to a calamine
                        // failure; otherwise keep `.xls` import warning behavior stable.
                        if calamine_formula_failed {
                            push_import_warning(
                                &mut warnings,
                                format!(
                                    "failed to import `.xls` formulas for sheet `{sheet_name}` via BIFF fallback: out-of-bounds stream offset {}",
                                    sheet_info.offset
                                ),
                                &mut warnings_suppressed,
                            );
                        }
                    } else {
                        // Provide the full workbook-global rgce decode context so BIFF8 worksheet
                        // formulas can resolve `PtgName` (NAME table) and 3D/external references
                        // (EXTERNSHEET/SUPBOOK) consistently with defined-name decoding.
                        let ctx = if let Some(decoder) = workbook_rgce_decoder.as_ref() {
                            decoder.rgce_decode_context(&sheet_names_by_biff_idx_for_rgce)
                        } else {
                            let empty_sheet_names: &[String] = &[];
                            let empty_externsheet: &[biff::externsheet::ExternSheetEntry] = &[];
                            let empty_supbooks: &[biff::supbook::SupBookInfo] = &[];
                            let empty_defined_names: &[biff::rgce::DefinedNameMeta] = &[];
                            biff::rgce::RgceDecodeContext {
                                codepage,
                                sheet_names: empty_sheet_names,
                                externsheet: empty_externsheet,
                                supbooks: empty_supbooks,
                                defined_names: empty_defined_names,
                            }
                        };

                        match biff::parse_biff8_sheet_formulas(
                            workbook_stream,
                            sheet_info.offset,
                            &ctx,
                        ) {
                            Ok(parsed) => {
                                for (cell_ref, formula) in parsed.formulas {
                                    let anchor = sheet.merged_regions.resolve_cell(cell_ref);
                                    if !calamine_formula_failed {
                                        if let Some(existing) = sheet.formula(anchor) {
                                            // Prefer calamine's string representation when it is
                                            // available, but fill any gaps (or entire sheets) from BIFF.
                                            //
                                            // Calamine renders some formulas it cannot fully resolve
                                            // (notably shared-formula `PtgExp` cases) as `#UNKNOWN!`.
                                            // Treat that as missing so BIFF parsing can replace it with a
                                            // real formula when possible.
                                            //
                                            // Calamine can also render some formulas as `#REF!` when it
                                            // cannot represent the underlying token stream (e.g. invalid
                                            // sheet names in 3D references). Prefer BIFF parsing in those
                                            // cases as well, but avoid overriding a `#REF!` sentinel with a
                                            // less-informative `#UNKNOWN!`.
                                            if !existing.contains("#UNKNOWN!")
                                                && existing != ErrorValue::Ref.as_str()
                                                && !is_calamine_unrecognised_formula_placeholder(
                                                    existing,
                                                )
                                            {
                                                continue;
                                            }
                                        }
                                    }

                                    let Some(normalized) = normalize_formula_text(&formula) else {
                                        continue;
                                    };
                                    if !calamine_formula_failed {
                                        if sheet.formula(anchor) == Some(ErrorValue::Ref.as_str())
                                            && normalized == ErrorValue::Unknown.as_str()
                                        {
                                            // Don't replace a (potentially meaningful) `#REF!` with
                                            // `#UNKNOWN!` if BIFF parsing couldn't decode the token
                                            // stream either.
                                            continue;
                                        }
                                    }
                                    sheet.set_formula(anchor, Some(normalized));

                                    if let Some(resolved) = style_id_for_cell_xf(
                                        xf_style_ids.as_deref(),
                                        sheet_cell_xfs,
                                        anchor,
                                    ) {
                                        sheet.set_style_id(anchor, resolved);
                                    }
                                }

                                if calamine_formula_failed {
                                    for w in parsed.warnings {
                                        push_import_warning(
                                            &mut warnings,
                                            format!(
                                                "failed to import `.xls` formula in sheet `{sheet_name}` via BIFF fallback: {w}"
                                            ),
                                            &mut warnings_suppressed,
                                        );
                                    }
                                }
                            }
                            Err(err) => {
                                if calamine_formula_failed {
                                    push_import_warning(
                                        &mut warnings,
                                        format!(
                                            "failed to import `.xls` formulas for sheet `{sheet_name}` via BIFF fallback: {err}"
                                        ),
                                        &mut warnings_suppressed,
                                    );
                                }
                            }
                        }
                    }
                }
            }
        }

        // BIFF8 `PtgExp`/`PtgTbl` formulas (shared formulas + array/CSE formulas + legacy table formulas).
        //
        // Calamine does not resolve `PtgExp`/`PtgTbl` references. Recover formulas by scanning the
        // raw worksheet substream for `SHRFMLA` (shared formula), `ARRAY` (array formula), and
        // `TABLE` definition records.
        //
        // Fallback behaviors:
        // - If definition records are missing/corrupt but the base cell still stores a full
        //   `FORMULA.rgce` token stream, recover follower formulas by materializing from the base
        //   cell's rgce across the row/col delta.
        // - Some producers store the shared rgce only in `SHRFMLA` and omit the base cell's full
        //   formula token stream (or the base cell `FORMULA` record entirely). In that case,
        //   recover formulas by decoding `SHRFMLA` rgce relative to each cell in the shared range.
        if let (
            Some(workbook_stream),
            Some(codepage),
            Some(biff_idx),
            Some(biff_version),
            Some(biff_sheets),
        ) = (
            workbook_stream.as_deref(),
            biff_codepage,
            biff_idx,
            biff_version,
            biff_sheets.as_ref(),
        ) {
            if biff_version == biff::BiffVersion::Biff8 {
                if let Some(sheet_info) = biff_sheets.get(biff_idx) {
                    // Provide the full workbook-global rgce decode context so recovered formulas
                    // can resolve `PtgName` and 3D/external references.
                    let ctx = if let Some(decoder) = workbook_rgce_decoder.as_ref() {
                        decoder.rgce_decode_context(&sheet_names_by_biff_idx_for_rgce)
                    } else {
                        let empty_sheet_names: &[String] = &[];
                        let empty_externsheet: &[biff::externsheet::ExternSheetEntry] = &[];
                        let empty_supbooks: &[biff::supbook::SupBookInfo] = &[];
                        let empty_defined_names: &[biff::rgce::DefinedNameMeta] = &[];
                            biff::rgce::RgceDecodeContext {
                                codepage,
                                sheet_names: empty_sheet_names,
                                externsheet: empty_externsheet,
                                supbooks: empty_supbooks,
                                defined_names: empty_defined_names,
                            }
                    };

                    // Best-effort: resolve `PtgExp`/`PtgTbl` formulas directly from the worksheet
                    // substream, including non-standard base-cell coordinate widths.
                    match biff::worksheet_formulas::parse_biff8_worksheet_ptgexp_formulas(
                        workbook_stream,
                        sheet_info.offset,
                        &ctx,
                    ) {
                        Ok(parsed) => {
                            for warning in parsed.warnings {
                                push_import_warning(
                                    &mut warnings,
                                    format!(
                                        "failed to fully decode `.xls` formulas in sheet `{sheet_name}`: {warning}"
                                    ),
                                    &mut warnings_suppressed,
                                );
                            }
                            for (cell_ref, formula) in parsed.formulas {
                                let anchor = sheet.merged_regions.resolve_cell(cell_ref);
                                if sheet
                                    .formula(anchor)
                                    .is_some_and(|f| {
                                        f != ErrorValue::Unknown.as_str()
                                            && !is_calamine_unrecognised_formula_placeholder(f)
                                    })
                                {
                                    continue;
                                }
                                let Some(normalized) = normalize_formula_text(&formula) else {
                                    continue;
                                };
                                sheet.set_formula(anchor, Some(normalized));
                                if let Some(resolved) = style_id_for_cell_xf(
                                    xf_style_ids.as_deref(),
                                    sheet_cell_xfs,
                                    anchor,
                                ) {
                                    sheet.set_style_id(anchor, resolved);
                                }
                            }
                        }
                        Err(err) => push_import_warning(
                            &mut warnings,
                            format!(
                                "failed to parse `.xls` PtgExp/PtgTbl formulas for sheet `{sheet_name}` from workbook stream: {err}"
                            ),
                            &mut warnings_suppressed,
                            ),
                    }

                    let mut apply_recovered_formulas =
                        |mut recovered: biff::formulas::PtgExpFallbackResult,
                          override_existing: bool,
                          warnings: &mut ImportWarnings,
                          warnings_suppressed: &mut bool| {
                            for warning in recovered.warnings.drain(..) {
                                push_import_warning(warnings, warning, &mut *warnings_suppressed);
                            }

                            for (cell_ref, formula_text) in recovered.formulas {
                                let anchor = sheet.merged_regions.resolve_cell(cell_ref);
                                if !override_existing {
                                    // Best-effort fallback only: do not override formulas that were
                                    // already resolved by calamine or earlier recovery passes.
                                    //
                                    // Calamine can return the `#UNKNOWN!` error sentinel for `PtgExp`
                                    // formulas it cannot resolve; treat that as "unresolved" so we can
                                    // replace it with a recovered materialized formula when possible.
                                    if sheet
                                        .formula(anchor)
                                        .is_some_and(|f| {
                                            f != ErrorValue::Unknown.as_str()
                                                && !is_calamine_unrecognised_formula_placeholder(f)
                                        })
                                    {
                                        continue;
                                    }
                                }
                            // Strip NUL bytes so formula text is parseable/stable (matches the
                            // calamine formula cleanup path).
                            let formula_text_clean;
                            let formula_text = if formula_text.contains('\0') {
                                formula_text_clean = formula_text.replace('\0', "");
                                formula_text_clean.as_str()
                            } else {
                                formula_text.as_str()
                            };

                            if let Some(normalized) = normalize_formula_text(formula_text) {
                                sheet.set_formula(anchor, Some(normalized));
                                if let Some(resolved) = style_id_for_cell_xf(
                                    xf_style_ids.as_deref(),
                                    sheet_cell_xfs,
                                    anchor,
                                ) {
                                    sheet.set_style_id(anchor, resolved);
                                }
                            }
                        }
                    };

                    match biff::formulas::recover_ptgexp_formulas_from_shrfmla_and_array(
                        workbook_stream,
                        sheet_info.offset,
                        &ctx,
                    ) {
                        Ok(recovered) => apply_recovered_formulas(
                            recovered,
                            true,
                            &mut warnings,
                            &mut warnings_suppressed,
                        ),
                        Err(err) => push_import_warning(
                            &mut warnings,
                            format!(
                                "failed to recover shared/array formulas for sheet `{sheet_name}`: {err}"
                            ),
                            &mut warnings_suppressed,
                        ),
                    }

                    match biff::formulas::recover_ptgexp_formulas_from_base_cell(
                        workbook_stream,
                        sheet_info.offset,
                        &ctx,
                    ) {
                        Ok(recovered) => apply_recovered_formulas(
                            recovered,
                            false,
                            &mut warnings,
                            &mut warnings_suppressed,
                        ),
                        Err(err) => push_import_warning(
                            &mut warnings,
                            format!(
                                "failed to recover shared formulas for sheet `{sheet_name}`: {err}"
                            ),
                            &mut warnings_suppressed,
                        ),
                    }
                }
            }
        }

        // Best-effort override path for BIFF8 formulas that calamine can mis-decode:
        // - shared formulas (`PtgExp` + `SHRFMLA`)
        // - 3D area references (`PtgArea3d`) that carry relative flags in the high bits of the
        //   column fields.
        if let (Some(workbook_stream), Some(codepage), Some(biff_version), Some(biff_idx)) = (
            workbook_stream.as_deref(),
            biff_codepage,
            biff_version,
            biff_idx,
        ) {
            if biff_version == biff::BiffVersion::Biff8 {
                if let Some(sheet_info) = biff_sheets.as_ref().and_then(|s| s.get(biff_idx)) {
                    let rgce_ctx = biff::rgce::RgceDecodeContext {
                        codepage,
                        sheet_names: &biff_rgce_sheet_names,
                        externsheet: &biff_rgce_externsheet,
                        supbooks: &biff_rgce_supbooks,
                        defined_names: &biff_rgce_defined_names,
                    };

                    if sheet_info.offset < workbook_stream.len() {
                        match biff::formulas::parse_biff8_sheet_formula_overrides(
                            workbook_stream,
                            sheet_info.offset,
                            &rgce_ctx,
                        ) {
                            Ok(parsed) => {
                                for (cell_ref, formula) in parsed.formulas {
                                    let anchor = sheet.merged_regions.resolve_cell(cell_ref);
                                    sheet.set_formula(anchor, Some(formula));
                                }
                                for warning in parsed.warnings {
                                    push_import_warning(
                                        &mut warnings,
                                        format!(
                                            "failed to fully decode `.xls` formula(s) for sheet `{sheet_name}`: {warning}"
                                        ),
                                        &mut warnings_suppressed,
                                    );
                                }
                            }
                            Err(err) => push_import_warning(
                                &mut warnings,
                                format!(
                                    "failed to import `.xls` BIFF8 formulas for sheet `{sheet_name}`: {err}"
                                ),
                                &mut warnings_suppressed,
                            ),
                        }
                    }

                    match biff::parse_biff_sheet_shared_formulas(workbook_stream, sheet_info.offset) {
                        Ok(mut parsed) => {
                            for w in parsed.warnings.drain(..) {
                                push_import_warning(
                                    &mut warnings,
                                    format!(
                                        "failed to import `.xls` shared formulas for sheet `{sheet_name}`: {w}"
                                    ),
                                    &mut warnings_suppressed,
                                );
                            }

                            for shared in parsed.shared_formulas.drain(..) {
                                let base_cell = biff::rgce::CellCoord::new(
                                    shared.row_first as u32,
                                    shared.col_first as u32,
                                );
                                let analysis =
                                    biff::rgce::analyze_biff8_shared_formula_rgce(&shared.rgce)
                                        .ok();
                                let mut materialization_failed = false;

                                for row_u16 in shared.row_first..=shared.row_last {
                                    for col_u16 in shared.col_first..=shared.col_last {
                                        let row: u32 = row_u16 as u32;
                                        let col: u32 = col_u16 as u32;
                                        if row >= EXCEL_MAX_ROWS || col >= EXCEL_MAX_COLS {
                                            continue;
                                        }

                                        let cell_ref = CellRef::new(row, col);
                                        let anchor = sheet.merged_regions.resolve_cell(cell_ref);

                                        if sheet.formula(anchor).is_some_and(|existing| {
                                            existing != ErrorValue::Unknown.as_str()
                                                && !is_calamine_unrecognised_formula_placeholder(
                                                    existing,
                                                )
                                        }) {
                                            continue;
                                        }

                                        let target_cell = biff::rgce::CellCoord::new(row, col);
                                        let delta_is_zero = target_cell == base_cell;
                                        let needs_materialization = !delta_is_zero
                                            && !materialization_failed
                                            && (analysis
                                                .as_ref()
                                                .is_some_and(|a| a.has_abs_refs_with_relative_flags)
                                                // Best-effort: if we cannot analyze the token
                                                // stream, still attempt materialization for
                                                // follower cells and fall back on failure.
                                                || analysis.is_none());

                                        let rgce_to_decode: std::borrow::Cow<'_, [u8]> =
                                            if needs_materialization {
                                                match biff::rgce::materialize_biff8_shared_formula_rgce(
                                                    &shared.rgce,
                                                    base_cell,
                                                    target_cell,
                                                ) {
                                                    Ok(v) => std::borrow::Cow::Owned(v),
                                                    Err(err) => {
                                                        materialization_failed = true;
                                                        push_import_warning(
                                                            &mut warnings,
                                                            format!(
                                                                "failed to materialize shared formula in sheet `{sheet_name}` at {}: {err}",
                                                                cell_ref.to_a1()
                                                            ),
                                                            &mut warnings_suppressed,
                                                        );
                                                        std::borrow::Cow::Borrowed(&shared.rgce)
                                                    }
                                                }
                                            } else {
                                                std::borrow::Cow::Borrowed(&shared.rgce)
                                            };

                                        let decoded = if shared.rgcb.is_empty() {
                                            biff::rgce::decode_biff8_rgce_with_base(
                                                &rgce_to_decode,
                                                &rgce_ctx,
                                                Some(target_cell),
                                            )
                                        } else {
                                            biff::rgce::decode_biff8_rgce_with_base_and_rgcb(
                                                &rgce_to_decode,
                                                &shared.rgcb,
                                                &rgce_ctx,
                                                Some(target_cell),
                                            )
                                        };

                                        for warning in decoded.warnings {
                                            push_import_warning(
                                                &mut warnings,
                                                format!(
                                                    "failed to decode shared formula in sheet `{sheet_name}` at {}: {warning}",
                                                    cell_ref.to_a1()
                                                ),
                                                &mut warnings_suppressed,
                                            );
                                        }

                                        let Some(normalized) = normalize_formula_text(&decoded.text)
                                        else {
                                            continue;
                                        };
                                        sheet.set_formula(anchor, Some(normalized));
                                        if let Some(resolved) = style_id_for_cell_xf(
                                            xf_style_ids.as_deref(),
                                            sheet_cell_xfs,
                                            anchor,
                                        ) {
                                            sheet.set_style_id(anchor, resolved);
                                        }
                                    }
                                }
                            }
                        }
                        Err(err) => push_import_warning(
                            &mut warnings,
                            format!(
                                "failed to import `.xls` shared formulas for sheet `{sheet_name}`: {err}"
                            ),
                            &mut warnings_suppressed,
                        ),
                    }
                }
            }
        }

        // Best-effort: recover What-If Analysis data-table formulas (`TABLE(...)`) that are encoded
        // in BIFF8 as `PtgTbl` tokens plus a worksheet-level `TABLE` record.
        //
        // Calamine’s formula extraction may omit these formulas because the `TABLE` record context
        // lives outside the `rgce` token stream. When BIFF workbook bytes are available, decode
        // them directly from the worksheet substream so we do not silently drop cell formulas.
        if let (Some(workbook_stream), Some(biff_version), Some(biff_idx)) =
            (workbook_stream.as_deref(), biff_version, biff_idx)
        {
            if biff_version == biff::BiffVersion::Biff8 {
                if let Some(sheet_info) = biff_sheets.as_ref().and_then(|s| s.get(biff_idx)) {
                    match biff::parse_biff8_sheet_table_formulas(workbook_stream, sheet_info.offset)
                    {
                        Ok(mut parsed) => {
                            for (cell_ref, formula) in parsed.formulas.drain() {
                                let anchor = sheet.merged_regions.resolve_cell(cell_ref);
                                sheet.set_formula(anchor, Some(formula));
                            }
                            for w in parsed.warnings.drain(..) {
                                push_import_warning(
                                    &mut warnings,
                                    format!(
                                        "failed to fully import `.xls` TABLE formulas for sheet `{sheet_name}`: {w}"
                                    ),
                                    &mut warnings_suppressed,
                                );
                            }
                        }
                        Err(err) => push_import_warning(
                            &mut warnings,
                            format!(
                                "failed to import `.xls` TABLE formulas for sheet `{sheet_name}`: {err}"
                            ),
                            &mut warnings_suppressed,
                        ),
                    }
                }
            }
        }

        // `calamine` does not surface `BLANK` records via `used_cells()`, but Excel
        // allows formatting empty cells. Apply any XF-derived number formats to
        // the sheet even when the value is empty so those cells round-trip.
        if let (Some(xf_style_ids), Some(sheet_cell_xfs)) =
            (xf_style_ids.as_deref(), sheet_cell_xfs)
        {
            let mut out_of_range_xf_count: usize = 0;
            if sheet.merged_regions.is_empty() {
                for (&cell_ref, &xf_idx) in sheet_cell_xfs {
                    if cell_ref.row >= EXCEL_MAX_ROWS || cell_ref.col >= EXCEL_MAX_COLS {
                        continue;
                    }

                    let Some(&style_id) = xf_style_ids.get(xf_idx as usize) else {
                        out_of_range_xf_count = out_of_range_xf_count.saturating_add(1);
                        continue;
                    };
                    if style_id == 0 {
                        continue;
                    };

                    if sheet
                        .cell(cell_ref)
                        .is_some_and(|cell| cell.style_id == style_id)
                    {
                        continue;
                    }
                    sheet.set_style_id(cell_ref, style_id);
                }
            } else {
                // We potentially see XF indices for multiple cells that resolve to the same merged-cell
                // anchor. When that happens, prefer the anchor cell’s own resolvable style; otherwise
                // choose the first cell (row/col order) within the merged region to keep this pass
                // deterministic (independent of HashMap iteration order).
                let mut best_style_for_anchor: HashMap<CellRef, (u32, CellRef, bool)> =
                    HashMap::new();

                for (&cell_ref, &xf_idx) in sheet_cell_xfs {
                    if cell_ref.row >= EXCEL_MAX_ROWS || cell_ref.col >= EXCEL_MAX_COLS {
                        continue;
                    }

                    let maybe_anchor = sheet.merged_regions.anchor_for(cell_ref);
                    // Normalise style assignments to merged-cell anchors so formatting inside a merged
                    // region round-trips consistently with the importer’s value/formula semantics.
                    let anchor = maybe_anchor.unwrap_or(cell_ref);

                    let Some(&style_id) = xf_style_ids.get(xf_idx as usize) else {
                        out_of_range_xf_count = out_of_range_xf_count.saturating_add(1);
                        continue;
                    };
                    if style_id == 0 {
                        continue;
                    };
                    if maybe_anchor.is_none() {
                        if sheet
                            .cell(anchor)
                            .is_some_and(|cell| cell.style_id == style_id)
                        {
                            continue;
                        }
                        sheet.set_style_id(anchor, style_id);
                        continue;
                    }

                    let source_is_anchor = cell_ref == anchor;
                    best_style_for_anchor
                        .entry(anchor)
                        .and_modify(|(existing_style_id, existing_source, existing_is_anchor)| {
                            if *existing_is_anchor {
                                // Anchor-derived styles always win.
                                return;
                            }
                            if source_is_anchor {
                                *existing_style_id = style_id;
                                *existing_source = cell_ref;
                                *existing_is_anchor = true;
                                return;
                            }
                            if (cell_ref.row, cell_ref.col)
                                < (existing_source.row, existing_source.col)
                            {
                                *existing_style_id = style_id;
                                *existing_source = cell_ref;
                            }
                        })
                        .or_insert((style_id, cell_ref, source_is_anchor));
                }

                for (anchor, (style_id, _, _)) in best_style_for_anchor {
                    if sheet
                        .cell(anchor)
                        .is_some_and(|cell| cell.style_id == style_id)
                    {
                        continue;
                    }
                    sheet.set_style_id(anchor, style_id);
                }
            }

            if out_of_range_xf_count > 0 {
                push_import_warning(
                    &mut warnings,
                    format!(
                        "skipped {out_of_range_xf_count} cells in sheet `{sheet_name}` with out-of-range XF indices"
                    ),
                    &mut warnings_suppressed,
                );
            }
        }
    }

    // Track per-sheet rename pairs (old sheet name -> final imported sheet name) so we can rewrite
    // sheet-qualified strings (cell formulas, internal hyperlinks, calamine-defined-name fallback)
    // after best-effort sheet-name sanitization.
    let mut sheet_rename_pairs: Vec<(String, String)> = Vec::new();

    // If we had to sanitize sheet names, internal hyperlinks and cell formulas may still
    // reference the original (invalid) sheet names. Rewrite those references to point at the
    // final imported sheet names so navigation and formulas remain correct after import and
    // round-trips to XLSX.
    if !final_sheet_names_by_idx.is_empty() {
        let mut resolved_sheet_names: HashMap<String, String> = HashMap::new();

        for (idx, sheet_meta) in sheets.iter().enumerate() {
            let Some(final_name) = final_sheet_names_by_idx.get(idx) else {
                continue;
            };

            // For hyperlink targets we resolve case-insensitively and strip embedded NULs so we
            // can match calamine's decoded sheet names against BIFF's BoundSheet names.
            resolved_sheet_names.insert(
                normalize_sheet_name_for_match(&sheet_meta.name),
                final_name.clone(),
            );

            // For formula rewriting we use exact old sheet name strings (the rewrite helper
            // handles case-insensitive matching internally).
            if sheet_meta.name != *final_name {
                sheet_rename_pairs.push((sheet_meta.name.clone(), final_name.clone()));
            }

            // Add a BIFF BoundSheet name alias when available: calamine sheet metadata and BIFF
            // BoundSheet names can diverge due to encoding issues or malformed files.
            if let Some(biff_idx) = sheet_mapping.get(idx).copied().flatten() {
                if let Some(biff_name) = biff_sheets
                    .as_ref()
                    .and_then(|sheets| sheets.get(biff_idx))
                    .map(|s| s.name.as_str())
                {
                    resolved_sheet_names
                        .entry(normalize_sheet_name_for_match(biff_name))
                        .or_insert_with(|| final_name.clone());

                    if biff_name != final_name && biff_name != sheet_meta.name {
                        sheet_rename_pairs.push((biff_name.to_string(), final_name.clone()));
                    }
                }
            }
        }

        if !resolved_sheet_names.is_empty() {
            // Rewrite internal hyperlink targets that refer to source sheet names.
            for sheet in &mut out.sheets {
                for link in &mut sheet.hyperlinks {
                    let HyperlinkTarget::Internal { sheet, .. } = &mut link.target else {
                        continue;
                    };
                    let key = normalize_sheet_name_for_match(sheet);
                    if let Some(resolved) = resolved_sheet_names.get(&key) {
                        *sheet = resolved.clone();
                    }
                }
            }
        }

        if !sheet_rename_pairs.is_empty() {
            formula_rewrite::rewrite_workbook_formulas_for_sheet_renames(
                &mut out,
                &sheet_rename_pairs,
            );
        }
    }

    if let Some(i_tab_cur) = workbook_active_tab {
        let biff_tab_idx = i_tab_cur as usize;

        // `WINDOW1.iTabCur` is a BIFF sheet index (BoundSheet order). Prefer mapping it back to the
        // imported sheet order (calamine metadata order) using the same reconciliation mapping we
        // apply for other per-sheet BIFF metadata.
        let imported_idx = sheet_mapping
            .iter()
            .position(|mapped| mapped.is_some_and(|biff_idx| biff_idx == biff_tab_idx))
            .unwrap_or(biff_tab_idx);

        if let Some(sheet) = out.sheets.get(imported_idx) {
            out.view.active_sheet_id = Some(sheet.id);
        } else {
            push_import_warning(
                &mut warnings,
                format!(
                    "skipping `.xls` active tab index {i_tab_cur}: workbook contains {} imported sheets",
                    out.sheets.len()
                ),
                &mut warnings_suppressed,
            );
        }
    }

    // Import defined names (workbook- and sheet-scoped).
    if let (Some(workbook_stream), Some(biff_version), Some(codepage)) =
        (workbook_stream.as_deref(), biff_version, biff_codepage)
    {
        // Resolve BIFF sheet indices to the sheet names used by our output workbook.
        let sheet_names_by_biff_idx = build_sheet_names_by_biff_idx(
            biff_sheets.as_deref(),
            &sheet_mapping,
            &final_sheet_names_by_idx,
        );

        // Resolve BIFF sheet indices to WorksheetIds.
        let mut sheet_ids_by_biff_idx: Vec<Option<formula_model::WorksheetId>> =
            vec![None; sheet_names_by_biff_idx.len()];
        for (cal_idx, maybe_biff_idx) in sheet_mapping.iter().enumerate() {
            let Some(biff_idx) = *maybe_biff_idx else {
                continue;
            };
            let Some(sheet_id) = sheet_ids_by_calamine_idx.get(cal_idx).copied() else {
                continue;
            };
            if biff_idx < sheet_ids_by_biff_idx.len() {
                sheet_ids_by_biff_idx[biff_idx] = Some(sheet_id);
            }
        }

        // Best-effort: recover shared/array formulas that are stored as BIFF token streams but are
        // not surfaced by calamine's `.xls` formula API (e.g. `SHRFMLA` + `PtgExp`).
        if biff_version == biff::BiffVersion::Biff8 {
            if let (Some(biff_sheets), Some(decoder)) =
                (biff_sheets.as_deref(), workbook_rgce_decoder.as_ref())
            {
                import_biff8_shared_formulas(
                    &mut out,
                    workbook_stream,
                    biff_sheets,
                    &sheet_names_by_biff_idx,
                    &sheet_ids_by_biff_idx,
                    decoder,
                    &mut warnings,
                    &mut warnings_suppressed,
                );
            }
        }

        // Resolve BIFF sheet indices to XLSX `localSheetId` values (0-based in workbook sheet
        // order). This is preserved in the model for round-trip fidelity when converting
        // `.xls` -> `.xlsx`.
        let mut local_sheet_ids_by_biff_idx: Vec<Option<u32>> =
            vec![None; sheet_ids_by_biff_idx.len()];
        for (cal_idx, maybe_biff_idx) in sheet_mapping.iter().enumerate() {
            let Some(biff_idx) = *maybe_biff_idx else {
                continue;
            };
            let cal_idx_u32: u32 = match cal_idx.try_into() {
                Ok(v) => v,
                Err(_) => continue,
            };
            if biff_idx < local_sheet_ids_by_biff_idx.len() {
                local_sheet_ids_by_biff_idx[biff_idx] = Some(cal_idx_u32);
            }
        }

        match biff::parse_biff_defined_names(
            workbook_stream,
            biff_version,
            codepage,
            &sheet_names_by_biff_idx,
        ) {
            Ok(mut parsed) => {
                for warning in parsed.warnings.drain(..) {
                    push_import_warning(&mut warnings, warning, &mut warnings_suppressed);
                }

                for name in parsed.names.drain(..) {
                    let (scope, xlsx_local_sheet_id) = match name.scope_sheet {
                        None => (DefinedNameScope::Workbook, None),
                        Some(biff_idx) => {
                            match sheet_ids_by_biff_idx.get(biff_idx).copied().flatten() {
                                Some(sheet_id) => (
                                    DefinedNameScope::Sheet(sheet_id),
                                    local_sheet_ids_by_biff_idx.get(biff_idx).copied().flatten(),
                                ),
                                None => {
                                    push_import_warning(
                                        &mut warnings,
                                        format!(
                                            "defined name `{}` has out-of-range sheet scope itab={} (sheet count={}); importing as workbook-scoped",
                                            name.name,
                                            biff_idx.saturating_add(1),
                                            sheet_ids_by_biff_idx.len()
                                        ),
                                        &mut warnings_suppressed,
                                    );
                                    (DefinedNameScope::Workbook, None)
                                }
                            }
                        }
                    };

                    if let Err(err) = out.create_defined_name(
                        scope,
                        name.name.clone(),
                        name.refers_to.clone(),
                        name.comment.clone(),
                        name.hidden,
                        xlsx_local_sheet_id,
                    ) {
                        push_import_warning(
                            &mut warnings,
                            format!("skipping defined name `{}`: {err}", name.name),
                            &mut warnings_suppressed,
                        );
                    }
                }
            }
            Err(err) => push_import_warning(
                &mut warnings,
                format!("failed to import `.xls` defined names: {err}"),
                &mut warnings_suppressed,
            ),
        }
    }

    // Calamine fallback: defined names.
    //
    // Calamine can surface defined names even when BIFF workbook parsing isn't available (or the
    // BIFF NAME parser fails). The `.xls` API does not expose sheet scope metadata (itab), so we
    // treat these names as workbook-scoped.
    //
    // Deterministic precedence: if a name already exists (e.g. imported via BIFF), skip the
    // calamine definition.
    let defined_names_before_calamine = out.defined_names.len();
    let mut imported_count: usize = 0;
    let mut skipped_count: usize = 0;
    for (name, refers_to) in calamine_defined_names {
        let name = normalize_calamine_defined_name_name(&name);
        let refers_to = refers_to.trim();
        let refers_to = refers_to.strip_prefix('=').unwrap_or(refers_to);
        // Calamine can surface BIFF8 formula/Unicode strings with embedded NUL bytes (notably
        // defined-name references via `PtgName`). Strip them so the formula text is parseable and
        // stable across import paths.
        let refers_to = if refers_to.contains('\0') {
            refers_to.replace('\0', "")
        } else {
            refers_to.to_string()
        };

        // Defined names can contain sheet references, and those can point at sheet names that were
        // later sanitized during import. Rewrite any sheet-qualified references using the same
        // sheet rename mapping we apply to cells.
        let mut refers_to = refers_to;
        if !sheet_rename_pairs.is_empty() {
            let original = refers_to.clone();
            let mut rewritten = refers_to.clone();
            // Apply rename pairs in reverse order so we don't cascade rewrites when a sanitized
            // name collides with another sheet's original name (see comment in the cell-formula
            // rewrite pass above).
            for (old_name, new_name) in sheet_rename_pairs.iter().rev() {
                rewritten =
                    formula_model::rewrite_sheet_names_in_formula(&rewritten, old_name, new_name);
            }

            // `refers_to` is stored without a leading `=`; normalize defensively after rewriting.
            // If rewriting produces an empty string, keep the original.
            if let Some(normalized) = normalize_formula_text(&rewritten) {
                refers_to = normalized;
            } else {
                refers_to = original;
            }
        }

        // When BIFF defined names were imported successfully, prefer them over calamine’s
        // best-effort string representation.
        if defined_names_before_calamine != 0
            && out
                .defined_names
                .iter()
                .any(|existing| existing.name.eq_ignore_ascii_case(&name))
        {
            continue;
        }

        match out.create_defined_name(
            DefinedNameScope::Workbook,
            name.clone(),
            refers_to,
            None,
            false,
            None,
        ) {
            Ok(_) => imported_count = imported_count.saturating_add(1),
            Err(err) => {
                skipped_count = skipped_count.saturating_add(1);
                push_import_warning(
                    &mut warnings,
                    format!("skipping `.xls` defined name `{name}` from calamine fallback: {err}"),
                    &mut warnings_suppressed,
                );
            }
        }
    }

    if defined_names_before_calamine == 0 && imported_count > 0 {
        push_import_warning(
            &mut warnings,
            "imported `.xls` defined names via calamine fallback; defined name metadata may be incomplete (scope/hidden/comment may be missing)",
            &mut warnings_suppressed,
        );
    }

    if skipped_count > 0 {
        push_import_warning(
            &mut warnings,
            format!(
                "skipped {skipped_count} `.xls` defined names from calamine fallback due to invalid/duplicate names"
            ),
            &mut warnings_suppressed,
        );
    }
    populate_print_settings_from_defined_names(&mut out, &mut warnings, &mut warnings_suppressed);

    // Best-effort import of worksheet AutoFilter ranges (phase 1).
    //
    // In BIFF, the filtered range is typically stored as a hidden built-in defined name
    // `_xlnm._FilterDatabase`, scoped to the worksheet that owns the AutoFilter.
    //
    // We only import the presence + range in phase 1 (filter criteria are not yet supported).
    //
    // Sort state is imported separately from worksheet BIFF `SORT` records when available.
    // Never fail import due to AutoFilter parsing.
    let mut autofilters: Vec<(formula_model::WorksheetId, Range)> = Vec::new();
    for name in &out.defined_names {
        if !is_filter_database_defined_name(&name.name) {
            continue;
        }

        let range = match parse_autofilter_range_from_defined_name(&name.refers_to) {
            Ok(range) => range,
            Err(err) => {
                push_import_warning(
                    &mut warnings,
                    format!(
                        "failed to parse `.xls` AutoFilter range `{}` from defined name `{}`: {err}",
                        name.refers_to, name.name
                    ),
                    &mut warnings_suppressed,
                );
                continue;
            }
        };

        let sheet_id = match name.scope {
            DefinedNameScope::Sheet(sheet_id) => sheet_id,
            DefinedNameScope::Workbook => {
                // Calamine's `.xls` defined-name API does not expose sheet scope. When a workbook-
                // scoped `_xlnm._FilterDatabase` is encountered, attempt to infer its sheet from an
                // explicit `Sheet!` prefix in the definition formula.
                let (sheet_name, inference_failed) =
                    infer_autofilter_sheet_name_from_workbook_scoped_defined_name(
                        &out,
                        &name.name,
                        &name.refers_to,
                        &mut warnings,
                        &mut warnings_suppressed,
                    );
                let Some(sheet_name) = sheet_name else {
                    if !inference_failed {
                        push_import_warning(
                            &mut warnings,
                            format!(
                                "skipping `.xls` AutoFilter defined name `{}`: workbook-scoped and sheet scope could not be inferred from `{}`",
                                name.name, name.refers_to
                            ),
                            &mut warnings_suppressed,
                        );
                    }
                    continue;
                };
                let Some(sheet_id) = out.sheet_by_name(&sheet_name).map(|s| s.id) else {
                    continue;
                };
                sheet_id
            }
        };

        autofilters.push((sheet_id, range));
    }

    // Calamine's `.xls` defined-name support does not handle BIFF8 built-in `NAME` records (the
    // `fBuiltin` flag) correctly and also cannot decode non-3D area refs like `PtgArea` into A1
    // text. This means AutoFilter ranges stored as `_xlnm._FilterDatabase` can be lost when BIFF
    // workbook-stream parsing is unavailable and we fall back to calamine for defined names.
    //
    // Best-effort: attempt to recover the AutoFilter range directly from the workbook stream via
    // our BIFF parser.
    //
    // We do this when:
    // - BIFF parsing was unavailable (`workbook_stream=None`), or
    // - calamine surfaced invalid/duplicate defined names (a strong signal that built-in names like
    //   `_FilterDatabase` were not imported correctly).
    //
    // This keeps `.xls` AutoFilter import resilient to future calamine behavior changes (e.g. if it
    // stops returning malformed built-in names and instead omits them entirely).
    let should_recover_autofilter_from_workbook_stream =
        workbook_stream.is_none() || (autofilters.is_empty() && skipped_count > 0);
    if should_recover_autofilter_from_workbook_stream {
        let workbook_stream_fallback = if workbook_stream.is_none() {
            match biff::read_workbook_stream_from_xls(path) {
                Ok(bytes) => Some(bytes),
                Err(err) => {
                    push_import_warning(
                        &mut warnings,
                        format!("failed to recover `.xls` AutoFilter ranges from workbook stream: {err}"),
                        &mut warnings_suppressed,
                    );
                    None
                }
            }
        } else {
            None
        };
        let workbook_stream_bytes = workbook_stream
            .as_deref()
            .or(workbook_stream_fallback.as_deref());

        if let Some(workbook_stream_bytes) = workbook_stream_bytes {
            let biff_version = biff::detect_biff_version(workbook_stream_bytes);
            let codepage = biff::parse_biff_codepage(workbook_stream_bytes);

            let biff_bound_sheets =
                biff::parse_biff_bound_sheets(workbook_stream_bytes, biff_version, codepage)
                    .ok()
                    .unwrap_or_default();

            let (sheet_names_by_biff_idx, sheet_offsets_by_biff_idx): (Vec<String>, Vec<usize>) =
                if biff_bound_sheets.is_empty() {
                    (sheets.iter().map(|s| s.name.clone()).collect(), Vec::new())
                } else {
                    (
                        biff_bound_sheets
                            .iter()
                            .map(|s| s.name.clone())
                            .collect::<Vec<_>>(),
                        biff_bound_sheets
                            .iter()
                            .map(|s| s.offset)
                            .collect::<Vec<_>>(),
                    )
                };

            let resolve_sheet_id_for_biff_idx = |biff_sheet_idx: usize| {
                // Best-effort mapping of BIFF sheet index -> output WorksheetId.
                //
                // Prefer sheet-name match (more robust when sheet orders differ), but fall
                // back to assuming BIFF sheet indices align with calamine's sheet order.
                sheet_names_by_biff_idx
                    .get(biff_sheet_idx)
                    .and_then(|biff_name| {
                        out.sheets
                            .iter()
                            .find(|s| sheet_name_eq_case_insensitive(&s.name, biff_name))
                            .map(|s| s.id)
                    })
                    .or_else(|| sheet_ids_by_calamine_idx.get(biff_sheet_idx).copied())
            };

            // Best-effort: infer which BIFF sheet contains AutoFilter metadata from its worksheet
            // substream. This is used to map workbook-scoped `_FilterDatabase` names that do not
            // specify a sheet (e.g. `=$A$1:$B$3`).
            //
            // We only use this heuristic when exactly one sheet has AutoFilter records.
            let biff_sheet_idx_with_sheet_stream_autofilter =
                if sheet_offsets_by_biff_idx.is_empty() {
                    None
                } else {
                    const RECORD_AUTOFILTERINFO: u16 = 0x009D;
                    const RECORD_FILTERMODE: u16 = 0x009B;

                    let mut matches = Vec::<usize>::new();
                    for (idx, &offset) in sheet_offsets_by_biff_idx.iter().enumerate() {
                        if matches.len() > 1 {
                            break;
                        }
                        if offset >= workbook_stream_bytes.len() {
                            continue;
                        }
                        let Ok(iter) = biff::records::BestEffortSubstreamIter::from_offset(
                            workbook_stream_bytes,
                            offset,
                        ) else {
                            continue;
                        };
                        for record in iter {
                            match record.record_id {
                                RECORD_AUTOFILTERINFO | RECORD_FILTERMODE => {
                                    matches.push(idx);
                                    break;
                                }
                                biff::records::RECORD_EOF => break,
                                _ => {}
                            }
                        }
                    }

                    if matches.len() == 1 {
                        Some(matches[0])
                    } else {
                        None
                    }
                };

            // Attempt to recover AutoFilter ranges even when `_FilterDatabase` is workbook-scoped.
            //
            // Some `.xls` files store the FilterDatabase NAME with workbook scope (`itab==0`) and
            // reference the target sheet via a 3D token (`PtgArea3d` / `PtgRef3d`) that requires
            // resolving through `EXTERNSHEET` (and sometimes `SUPBOOK`). This helper recovers a
            // BIFF-sheet-index -> range mapping directly from the BIFF workbook stream.
            match biff::parse_biff_filter_database_ranges(
                workbook_stream_bytes,
                biff_version,
                codepage,
                Some(sheet_names_by_biff_idx.len()),
            ) {
                Ok(parsed) => {
                    for warning in parsed.warnings {
                        push_import_warning(
                            &mut warnings,
                            format!(
                                "failed to fully recover `.xls` AutoFilter ranges from workbook stream: {warning}"
                            ),
                            &mut warnings_suppressed,
                        );
                    }

                    for (biff_sheet_idx, range) in parsed.by_sheet {
                        let sheet_id = resolve_sheet_id_for_biff_idx(biff_sheet_idx);

                        let Some(sheet_id) = sheet_id else {
                            push_import_warning(
                                &mut warnings,
                                format!(
                                    "skipping `.xls` AutoFilter range `{range}`: out-of-range sheet index {} (sheet count={})",
                                    biff_sheet_idx.saturating_add(1),
                                    out.sheets.len()
                                ),
                                &mut warnings_suppressed,
                            );
                            continue;
                        };

                        autofilters.push((sheet_id, range));
                    }
                }
                Err(err) => push_import_warning(
                    &mut warnings,
                    format!("failed to recover `.xls` AutoFilter ranges from workbook stream: {err}"),
                    &mut warnings_suppressed,
                ),
            }

            match biff::parse_biff_defined_names(
                workbook_stream_bytes,
                biff_version,
                codepage,
                &sheet_names_by_biff_idx,
            ) {
                Ok(mut parsed) => {
                    for name in parsed.names.drain(..) {
                        if !is_filter_database_defined_name(&name.name) {
                            continue;
                        }
                        let range = match parse_autofilter_range_from_defined_name(&name.refers_to)
                        {
                            Ok(range) => range,
                            Err(err) => {
                                push_import_warning(
                                    &mut warnings,
                                    format!(
                                        "failed to parse `.xls` AutoFilter range `{}` from defined name `{}`: {err}",
                                        name.refers_to, name.name
                                    ),
                                    &mut warnings_suppressed,
                                );
                                continue;
                            }
                        };

                        let sheet_id = match name.scope_sheet {
                            Some(biff_sheet_idx) => {
                                let sheet_id = resolve_sheet_id_for_biff_idx(biff_sheet_idx);
                                if sheet_id.is_none() {
                                    push_import_warning(
                                        &mut warnings,
                                        format!(
                                            "skipping `.xls` AutoFilter defined name `{}`: out-of-range sheet index {} (sheet count={})",
                                            name.name,
                                            biff_sheet_idx.saturating_add(1),
                                            out.sheets.len()
                                        ),
                                        &mut warnings_suppressed,
                                    );
                                }
                                sheet_id
                            }
                            None => {
                                // Attempt to infer the sheet target of workbook-scoped AutoFilter
                                // names. Some `.xls` files store `_FilterDatabase` as workbook-scope
                                // but use an unqualified 2D range formula.
                                let (sheet_name, inference_failed) =
                                    infer_autofilter_sheet_name_from_workbook_scoped_defined_name(
                                        &out,
                                        &name.name,
                                        &name.refers_to,
                                        &mut warnings,
                                        &mut warnings_suppressed,
                                    );

                                let inferred = sheet_name
                                    .and_then(|sheet_name| {
                                        out.sheet_by_name(&sheet_name).map(|s| s.id)
                                    })
                                    .or_else(|| {
                                        // Best-effort: infer sheet from the worksheet substream's
                                        // AutoFilter metadata (AUTOFILTERINFO / FILTERMODE).
                                        if inference_failed {
                                            return None;
                                        }
                                        let biff_sheet_idx =
                                            biff_sheet_idx_with_sheet_stream_autofilter?;
                                        resolve_sheet_id_for_biff_idx(biff_sheet_idx)
                                    });

                                if inferred.is_none() && !inference_failed {
                                    push_import_warning(
                                        &mut warnings,
                                        format!(
                                            "skipping `.xls` AutoFilter defined name `{}`: workbook-scoped and sheet scope could not be inferred from `{}`",
                                            name.name, name.refers_to
                                        ),
                                        &mut warnings_suppressed,
                                    );
                                }
                                inferred
                            }
                        };

                        let Some(sheet_id) = sheet_id else {
                            continue;
                        };
                        autofilters.push((sheet_id, range));
                    }
                }
                Err(err) => push_import_warning(
                    &mut warnings,
                    format!("failed to recover `.xls` AutoFilter ranges from workbook stream: {err}"),
                    &mut warnings_suppressed,
                ),
            }
        }
    }

    for (sheet_id, range) in autofilters {
        let Some(sheet) = out.sheet_mut(sheet_id) else {
            push_import_warning(
                &mut warnings,
                format!("skipping `.xls` AutoFilter range for missing sheet id {sheet_id}"),
                &mut warnings_suppressed,
            );
            continue;
        };
        match sheet.auto_filter.as_mut() {
            Some(existing) => {
                // Prefer ranges derived from `_FilterDatabase` NAME records (or other defined-name
                // based recovery) over any earlier best-effort inference from worksheet DIMENSIONS
                // / AUTOFILTERINFO / FILTERMODE.
                //
                // We only update the range and preserve any existing (future) filter state fields.
                if existing.range != range {
                    existing.range = range;
                }
            }
            None => {
                sheet.auto_filter = Some(SheetAutoFilter {
                    range,
                    filter_columns: Vec::new(),
                    sort_state: None,
                    raw_xml: Vec::new(),
                });
            }
        }
    }

    // Post-process AutoFilter-related state that depends on the final canonical filter range.
    //
    // The AutoFilter range itself may be corrected later via `_FilterDatabase` defined-name parsing
    // and workbook-stream recovery. Any downstream state that depends on the range (e.g. SORT key
    // ranges and FILTERMODE row classification) must therefore happen after the final range is
    // applied.
    if let Some(props_by_sheet) = row_col_props.as_ref() {
        for (calamine_idx, &sheet_id) in sheet_ids_by_calamine_idx.iter().enumerate() {
            let biff_idx = sheet_mapping.get(calamine_idx).copied().flatten();
            let Some(biff_idx) = biff_idx else {
                continue;
            };

            let Some(props) = props_by_sheet.get(biff_idx) else {
                continue;
            };

            let Some(sheet) = out.sheet_mut(sheet_id) else {
                push_import_warning(
                    &mut warnings,
                    format!("skipping `.xls` AutoFilter post-processing for missing sheet id {sheet_id}"),
                    &mut warnings_suppressed,
                );
                continue;
            };
            let sheet_name = sheet.name.clone();

            let Some(af_range) = sheet.auto_filter.as_ref().map(|af| af.range) else {
                continue;
            };

            // Best-effort: attach AutoFilter filter columns recovered from BIFF8 future records
            // (AutoFilter12) if we don't already have criteria from legacy AUTOFILTER records.
            if let Some(af) = sheet.auto_filter.as_mut() {
                if af.filter_columns.is_empty() && !props.auto_filter_columns.is_empty() {
                    let max_col_id = af_range.end.col.saturating_sub(af_range.start.col);
                    let cols: Vec<_> = props
                        .auto_filter_columns
                        .iter()
                        .filter(|col| col.col_id <= max_col_id)
                        .cloned()
                        .collect();
                    if !cols.is_empty() {
                        af.filter_columns = cols;
                    }
                }
            }

            // Best-effort: attach worksheet SORT state to the AutoFilter if the key ranges fall
            // within the final AutoFilter range.
            if sheet
                .auto_filter
                .as_ref()
                .and_then(|af| af.sort_state.as_ref())
                .is_none()
            {
                if let Some(sort_state) = props.sort_state.clone() {
                    if sort_state.conditions.iter().all(|cond| {
                        cond.range.start.row >= af_range.start.row
                            && cond.range.end.row <= af_range.end.row
                            && cond.range.start.col >= af_range.start.col
                            && cond.range.end.col <= af_range.end.col
                    }) {
                        if let Some(af) = sheet.auto_filter.as_mut() {
                            af.sort_state = Some(sort_state);
                        }
                    }
                }
            }

            // Best-effort: recover AutoFilter sort state from BIFF `SORT` metadata when the sheet
            // stream scan did not yield a supported `SORT` layout.
            if let (Some(workbook_stream_bytes), Some(af)) =
                (workbook_stream.as_deref(), sheet.auto_filter.as_mut())
            {
                if af.sort_state.is_none() {
                    if let Some(&offset) = sheet_stream_offsets_by_sheet_id.get(&sheet.id) {
                        if offset >= workbook_stream_bytes.len() {
                            push_import_warning(
                                &mut warnings,
                                format!(
                                    "failed to import `.xls` sort state for sheet `{sheet_name}`: out-of-bounds stream offset {offset}",
                                ),
                                &mut warnings_suppressed,
                            );
                        } else {
                            match biff::parse_biff_sheet_sort_state(
                                workbook_stream_bytes,
                                offset,
                                af.range,
                            ) {
                                Ok(mut parsed) => {
                                    for w in parsed.warnings.drain(..) {
                                        push_import_warning(
                                            &mut warnings,
                                            format!(
                                                "failed to import `.xls` sort state for sheet `{sheet_name}`: {w}"
                                            ),
                                            &mut warnings_suppressed,
                                        );
                                    }
                                    if af.sort_state.is_none() {
                                        af.sort_state = parsed.sort_state;
                                    }
                                }
                                Err(err) => push_import_warning(
                                    &mut warnings,
                                    format!(
                                        "failed to import `.xls` sort state for sheet `{sheet_name}`: {err}"
                                    ),
                                    &mut warnings_suppressed,
                                ),
                            }
                        }
                    }
                }
            }

            if !props.filter_mode {
                continue;
            }

            // BIFF `FILTERMODE` indicates that some rows are hidden by an active AutoFilter. BIFF
            // ROW records do not distinguish between user-hidden rows and filter-hidden rows, so
            // during the initial sheet import we treat hidden ROW records as user-hidden. After
            // the final AutoFilter range is applied, reclassify hidden data rows within that range
            // as filter-hidden.
            let start_row = af_range.start.row;
            let end_row = af_range.end.row;
            if end_row <= start_row {
                continue;
            }

            // FILTERMODE only applies to rows below the header row.
            let data_start_1based = start_row.saturating_add(2);
            let data_end_1based = end_row.saturating_add(1);
            if data_start_1based <= data_end_1based {
                sheet.clear_filter_hidden_range(data_start_1based, data_end_1based);
            }

            for (&row, row_props) in &props.rows {
                if !row_props.hidden {
                    continue;
                }
                // Skip the header row: Excel filters apply to data rows below the header.
                if row <= start_row || row > end_row {
                    continue;
                }
                let row_1based = row.saturating_add(1);
                if sheet.row_outline_entry(row_1based).hidden.outline {
                    // Outline-hidden rows remain hidden regardless of filter state.
                    continue;
                }
                // Convert BIFF hidden rows within the AutoFilter data range to filter-hidden rows.
                sheet.set_row_hidden(row, false);
                sheet.set_filter_hidden_row(row_1based, true);
            }
        }
    }

    // Best-effort import of AutoFilter criteria from worksheet AUTOFILTER records.
    //
    // This is intentionally resilient: malformed records are surfaced as warnings but do not fail
    // the overall import.
    if let (Some(workbook_stream_bytes), Some(biff_version), Some(codepage)) =
        (workbook_stream.as_deref(), biff_version, biff_codepage)
    {
        for sheet in out.sheets.iter_mut() {
            let Some(af) = sheet.auto_filter.as_mut() else {
                continue;
            };
            let Some(&offset) = sheet_stream_offsets_by_sheet_id.get(&sheet.id) else {
                continue;
            };

            match biff::parse_biff_sheet_autofilter_criteria(
                workbook_stream_bytes,
                offset,
                biff_version,
                codepage,
                af.range,
            ) {
                Ok(mut parsed) => {
                    let parsed_columns_empty = parsed.filter_columns.is_empty();
                    let parsed_warnings_empty = parsed.warnings.is_empty();
                    // Prefer AutoFilter criteria parsed from AutoFilter12 future records (when
                    // available) and only fall back to legacy AUTOFILTER criteria when no other
                    // filter-column metadata was recovered.
                    if af.filter_columns.is_empty() && !parsed.filter_columns.is_empty() {
                        af.filter_columns = std::mem::take(&mut parsed.filter_columns);
                    }

                    let filter_mode = sheet_filter_mode_by_sheet_id
                        .get(&sheet.id)
                        .copied()
                        .unwrap_or(false);
                    if filter_mode
                        && af.filter_columns.is_empty()
                        && parsed_columns_empty
                        && parsed_warnings_empty
                    {
                        push_import_warning(
                            &mut warnings,
                            format!(
                                "failed to fully import `.xls` autofilter criteria for sheet `{}`: FILTERMODE is set but no AutoFilter criteria records were found",
                                sheet.name
                            ),
                            &mut warnings_suppressed,
                        );
                    }

                    for w in parsed.warnings.drain(..) {
                        push_import_warning(
                            &mut warnings,
                            format!(
                                "failed to fully import `.xls` autofilter criteria for sheet `{}`: {w}",
                                sheet.name
                            ),
                            &mut warnings_suppressed,
                        );
                    }
                }
                Err(err) => push_import_warning(
                    &mut warnings,
                    format!(
                        "failed to fully import `.xls` autofilter criteria for sheet `{}`: {err}",
                        sheet.name
                    ),
                    &mut warnings_suppressed,
                ),
            }
        }
    }
    Ok(XlsImportResult {
        workbook: out,
        source: ImportSource {
            path: path.to_path_buf(),
            format: SourceFormat::Xls,
        },
        merged_ranges,
        warnings: warnings.into_vec(),
    })
}

fn to_cell_ref(start: (u32, u32), row: usize, col: usize) -> Option<CellRef> {
    // NOTE: calamine `Range` iterators return coordinates relative to `range.start()`
    // rather than absolute worksheet coordinates.
    let row: u32 = row.try_into().ok()?;
    let col: u32 = col.try_into().ok()?;

    let row = start.0.checked_add(row)?;
    let col = start.1.checked_add(col)?;

    if row >= EXCEL_MAX_ROWS || col >= EXCEL_MAX_COLS {
        return None;
    }

    Some(CellRef::new(row, col))
}

fn sheet_visible_to_visibility(visible: SheetVisible) -> SheetVisibility {
    match visible {
        SheetVisible::Visible => SheetVisibility::Visible,
        SheetVisible::Hidden => SheetVisibility::Hidden,
        SheetVisible::VeryHidden => SheetVisibility::VeryHidden,
    }
}

fn convert_value(
    value: &Data,
    date_time_styles: Option<DateTimeStyleIds>,
) -> Option<(CellValue, Option<u32>)> {
    match value {
        Data::Empty => None,
        Data::Bool(v) => Some((CellValue::Boolean(*v), None)),
        Data::Int(v) => Some((CellValue::Number(*v as f64), None)),
        Data::Float(v) => Some((CellValue::Number(*v), None)),
        Data::String(v) => {
            let text = if v.contains('\0') {
                v.replace('\0', "")
            } else {
                v.clone()
            };
            Some((CellValue::String(text), None))
        }
        Data::Error(e) => Some((CellValue::Error(cell_error_to_error_value(e.clone())), None)),
        Data::DateTime(v) => Some((
            CellValue::Number(v.as_f64()),
            date_time_styles.map(|styles| styles.style_for_excel_datetime(v)),
        )),
        Data::DateTimeIso(v) => {
            let text = if v.contains('\0') {
                v.replace('\0', "")
            } else {
                v.clone()
            };
            Some((CellValue::String(text), None))
        }
        Data::DurationIso(v) => {
            let text = if v.contains('\0') {
                v.replace('\0', "")
            } else {
                v.clone()
            };
            Some((CellValue::String(text), None))
        }
    }
}

fn style_id_for_cell_xf(
    xf_style_ids: Option<&[u32]>,
    sheet_cell_xfs: Option<&HashMap<CellRef, u16>>,
    cell_ref: CellRef,
) -> Option<u32> {
    let xf_index = sheet_cell_xfs?.get(&cell_ref).copied()? as usize;
    let style_id = *xf_style_ids?.get(xf_index)?;
    (style_id != 0).then_some(style_id)
}

fn cell_error_to_error_value(err: calamine::CellErrorType) -> ErrorValue {
    use calamine::CellErrorType;

    match err {
        CellErrorType::Div0 => ErrorValue::Div0,
        CellErrorType::NA => ErrorValue::NA,
        CellErrorType::Name => ErrorValue::Name,
        CellErrorType::Null => ErrorValue::Null,
        CellErrorType::Num => ErrorValue::Num,
        CellErrorType::Ref => ErrorValue::Ref,
        CellErrorType::Value => ErrorValue::Value,
        CellErrorType::GettingData => ErrorValue::GettingData,
    }
}

fn is_calamine_unrecognised_formula_placeholder(formula: &str) -> bool {
    let trimmed = formula.trim_start();
    trimmed.starts_with("Unrecognised formula") || trimmed.starts_with("Unrecognized formula")
}

fn apply_row_col_properties(
    sheet: &mut formula_model::Worksheet,
    props: &biff::SheetRowColProperties,
) {
    for (&row, row_props) in &props.rows {
        if row >= EXCEL_MAX_ROWS {
            continue;
        }
        if row_props.height.is_some() {
            sheet.set_row_height(row, row_props.height);
        }
    }

    for (&col, col_props) in &props.cols {
        if col >= EXCEL_MAX_COLS {
            continue;
        }
        if col_props.width.is_some() {
            sheet.set_col_width(col, col_props.width);
        }
    }
}

fn apply_row_col_style_ids(
    sheet: &mut formula_model::Worksheet,
    props: &biff::SheetRowColProperties,
    xf_style_ids: Option<&[u32]>,
    warnings: &mut ImportWarnings,
    suppressed: &mut bool,
    sheet_name: &str,
) {
    let Some(xf_style_ids) = xf_style_ids else {
        return;
    };

    let mut out_of_range_rows: usize = 0;
    for (&row, row_props) in &props.rows {
        if row >= EXCEL_MAX_ROWS {
            continue;
        }
        let Some(xf_idx) = row_props.xf_index else {
            continue;
        };
        let Some(&style_id) = xf_style_ids.get(xf_idx as usize) else {
            out_of_range_rows = out_of_range_rows.saturating_add(1);
            continue;
        };
        if style_id != 0 {
            sheet.set_row_style_id(row, Some(style_id));
        }
    }

    let mut out_of_range_cols: usize = 0;
    for (&col, col_props) in &props.cols {
        if col >= EXCEL_MAX_COLS {
            continue;
        }
        let Some(xf_idx) = col_props.xf_index else {
            continue;
        };
        let Some(&style_id) = xf_style_ids.get(xf_idx as usize) else {
            out_of_range_cols = out_of_range_cols.saturating_add(1);
            continue;
        };
        if style_id != 0 {
            sheet.set_col_style_id(col, Some(style_id));
        }
    }

    if out_of_range_rows > 0 || out_of_range_cols > 0 {
        let mut parts = Vec::new();
        if out_of_range_rows > 0 {
            parts.push(format!("{out_of_range_rows} rows"));
        }
        if out_of_range_cols > 0 {
            parts.push(format!("{out_of_range_cols} columns"));
        }
        push_import_warning(
            warnings,
            format!(
                "skipped {} in sheet `{sheet_name}` with out-of-range XF indices",
                parts.join(" and ")
            ),
            suppressed,
        );
    }
}

fn apply_outline_properties(
    sheet: &mut formula_model::Worksheet,
    props: &biff::SheetRowColProperties,
) {
    sheet.outline.pr = props.outline_pr;

    for (&row, row_props) in &props.rows {
        if row >= EXCEL_MAX_ROWS {
            continue;
        }
        if row_props.outline_level == 0 && !row_props.collapsed {
            continue;
        }
        let row_1based = row.saturating_add(1);
        let entry = sheet.outline.rows.entry_mut(row_1based);
        entry.level = row_props.outline_level.min(7);
        entry.collapsed = row_props.collapsed;
    }

    for (&col, col_props) in &props.cols {
        if col >= EXCEL_MAX_COLS {
            continue;
        }
        if col_props.outline_level == 0 && !col_props.collapsed {
            continue;
        }
        let col_1based = col.saturating_add(1);
        let entry = sheet.outline.cols.entry_mut(col_1based);
        entry.level = col_props.outline_level.min(7);
        entry.collapsed = col_props.collapsed;
    }

    // Derive which rows/columns are hidden because they live inside a collapsed outline group.
    sheet.outline.recompute_outline_hidden_rows();
    sheet.outline.recompute_outline_hidden_cols();

    // BIFF uses the same hidden bit for user-hidden and outline-hidden rows/columns. Prefer the
    // derived outline state, and treat any remaining hidden flags as user-hidden.
    for (&row, row_props) in &props.rows {
        if !row_props.hidden || row >= EXCEL_MAX_ROWS {
            continue;
        }
        let row_1based = row.saturating_add(1);
        if sheet.outline.rows.entry(row_1based).hidden.outline {
            continue;
        }
        sheet.set_row_hidden(row, true);
    }

    for (&col, col_props) in &props.cols {
        if !col_props.hidden || col >= EXCEL_MAX_COLS {
            continue;
        }
        let col_1based = col.saturating_add(1);
        if sheet.outline.cols.entry(col_1based).hidden.outline {
            continue;
        }
        sheet.set_col_hidden(col, true);
    }

    // Ensure outline-hidden flags are up to date after any user-hidden state mutations.
    sheet.outline.recompute_outline_hidden_rows();
    sheet.outline.recompute_outline_hidden_cols();
}

fn truncate_to_utf16_len(value: &str, max_len: usize) -> String {
    if value.encode_utf16().count() <= max_len {
        return value.to_string();
    }

    let mut out = String::new();
    let mut used = 0usize;
    for ch in value.chars() {
        let ch_len = ch.len_utf16();
        if used.saturating_add(ch_len) > max_len {
            break;
        }
        out.push(ch);
        used = used.saturating_add(ch_len);
    }
    out
}

fn normalize_calamine_defined_name_name(name: &str) -> String {
    // Calamine can surface BIFF8 Unicode strings with embedded NUL bytes; strip them so the
    // imported name matches Excel’s visible name semantics.
    name.replace('\0', "")
}

fn is_filter_database_defined_name(name: &str) -> bool {
    // In XLSX, the AutoFilter range is stored in `_xlnm._FilterDatabase`. Some `.xls` writers (and
    // some decoders) omit the `_xlnm.` prefix and use Excel's visible built-in name
    // `_FilterDatabase` instead. Treat both spellings as the AutoFilter name.
    //
    // Calamine has been observed to surface `_FilterDatabase` / `_xlnm._FilterDatabase` with the
    // final `e` truncated (i.e. `_FilterDatabas` / `_xlnm._FilterDatabas`) for some BIFF NAME
    // encodings; accept those variants as well so we still preserve AutoFilter ranges when BIFF
    // parsing is unavailable.
    name.eq_ignore_ascii_case(XLNM_FILTER_DATABASE)
        || name.eq_ignore_ascii_case("_xlnm._FilterDatabas")
        || name.eq_ignore_ascii_case("_FilterDatabase")
        || name.eq_ignore_ascii_case("_FilterDatabas")
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum WorkbookScopedDefinedNameSheetInference {
    /// The defined name's `refers_to` formula does not contain an explicit `Sheet!` prefix, or the
    /// inferred sheet does not exist in the workbook.
    ///
    /// This outcome is not warning-worthy; callers may choose best-effort heuristics.
    NotApplicable,
    /// The defined name's `refers_to` formula was malformed or ambiguous in a way that should be
    /// surfaced to the user as an import warning.
    ///
    /// Callers must not run heuristic fallbacks based on warning suppression state.
    Failure,
    /// The sheet scope was inferred from an explicit `Sheet!` prefix.
    Inferred(String),
}

fn infer_sheet_name_from_workbook_scoped_defined_name(
    workbook: &Workbook,
    name: &str,
    refers_to: &str,
    warnings: &mut ImportWarnings,
    suppressed: &mut bool,
) -> WorkbookScopedDefinedNameSheetInference {
    // `Workbook::create_defined_name` strips leading `=` but not `@`. Strip both defensively so we
    // can infer sheet scope from dynamic-array era implicit intersection prefixes as well.
    let refers_to = refers_to.trim();
    let refers_to = refers_to.strip_prefix('=').unwrap_or(refers_to).trim();
    let refers_to = strip_wrapping_parentheses(refers_to);
    let refers_to = refers_to.strip_prefix('@').unwrap_or(refers_to).trim();
    let refers_to = strip_wrapping_parentheses(refers_to);
    if refers_to.is_empty() {
        return WorkbookScopedDefinedNameSheetInference::NotApplicable;
    }

    let areas = match split_print_name_areas(refers_to) {
        Ok(v) => v,
        Err(err) => {
            push_import_warning(
                warnings,
                format!("failed to infer sheet scope for workbook-scoped `{name}`: {err}"),
                suppressed,
            );
            return WorkbookScopedDefinedNameSheetInference::Failure;
        }
    };

    let mut inferred: Option<String> = None;
    for area in areas {
        let (sheet_name, _) = match split_print_name_sheet_ref(area) {
            Ok(v) => v,
            Err(err) => {
                push_import_warning(
                    warnings,
                    format!(
                        "failed to infer sheet scope for workbook-scoped `{name}` entry {area:?}: {err}"
                    ),
                    suppressed,
                );
                return WorkbookScopedDefinedNameSheetInference::Failure;
            }
        };

        let Some(sheet_name) = sheet_name else {
            // We only infer sheet scope from explicit `Sheet!` prefixes.
            return WorkbookScopedDefinedNameSheetInference::NotApplicable;
        };
        let sheet_name = strip_workbook_prefix_from_sheet_ref(&sheet_name).to_string();

        match inferred.as_ref() {
            None => inferred = Some(sheet_name),
            Some(existing) => {
                if !sheet_name_eq_case_insensitive(existing, &sheet_name) {
                    push_import_warning(
                        warnings,
                        format!(
                            "skipping workbook-scoped `{name}`: references multiple sheets (`{existing}` and `{sheet_name}`)"
                        ),
                        suppressed,
                    );
                    return WorkbookScopedDefinedNameSheetInference::Failure;
                }
            }
        }
    }

    let Some(inferred) = inferred else {
        return WorkbookScopedDefinedNameSheetInference::NotApplicable;
    };
    match workbook.sheet_by_name(&inferred) {
        Some(s) => WorkbookScopedDefinedNameSheetInference::Inferred(s.name.clone()),
        None => WorkbookScopedDefinedNameSheetInference::NotApplicable,
    }
}

fn infer_autofilter_sheet_name_from_workbook_scoped_defined_name(
    workbook: &Workbook,
    name: &str,
    refers_to: &str,
    warnings: &mut ImportWarnings,
    suppressed: &mut bool,
) -> (Option<String>, bool) {
    let warnings_before_infer = warnings.count();
    let inferred = infer_sheet_name_from_workbook_scoped_defined_name(
        workbook, name, refers_to, warnings, suppressed,
    );
    let inference_failed = matches!(inferred, WorkbookScopedDefinedNameSheetInference::Failure)
        || warnings.count() != warnings_before_infer;

    let sheet_name = match inferred {
        WorkbookScopedDefinedNameSheetInference::Inferred(name) => Some(name),
        WorkbookScopedDefinedNameSheetInference::NotApplicable
        | WorkbookScopedDefinedNameSheetInference::Failure => None,
    }
    .or_else(|| (workbook.sheets.len() == 1).then(|| workbook.sheets[0].name.clone()))
    .or_else(|| {
        // Some `.xls` files store a workbook-scoped `_FilterDatabase` name whose `refers_to` is an
        // unqualified 2D range (e.g. `=$A$1:$B$3`), even when the workbook contains multiple
        // sheets. In this case, we have no explicit sheet scope to apply.
        //
        // Best-effort: if exactly one sheet already has AutoFilter metadata inferred from the
        // worksheet substream (AUTOFILTERINFO / FILTERMODE + DIMENSIONS), assume this FilterDatabase
        // range belongs to that sheet.
        //
        // Do not guess when multiple sheets have AutoFilter metadata, or when sheet inference
        // failed in a warning-worthy way.
        if inference_failed {
            return None;
        }

        let refers_to = refers_to.trim();
        let refers_to = refers_to.strip_prefix('=').unwrap_or(refers_to).trim();
        let refers_to = refers_to.strip_prefix('@').unwrap_or(refers_to).trim();
        if refers_to.contains('!') {
            return None;
        }

        let mut sheets_with_autofilter = workbook.sheets.iter().filter(|s| s.auto_filter.is_some());
        let only = sheets_with_autofilter.next()?;
        sheets_with_autofilter
            .next()
            .is_none()
            .then(|| only.name.clone())
    });

    (sheet_name, inference_failed)
}

fn populate_print_settings_from_defined_names(
    workbook: &mut Workbook,
    warnings: &mut ImportWarnings,
    suppressed: &mut bool,
) {
    // We need to snapshot the defined names up-front so we can mutably update print settings while
    // iterating.
    let builtins: Vec<(DefinedNameScope, String, String)> = workbook
        .defined_names
        .iter()
        .filter(|n| {
            n.name.eq_ignore_ascii_case(formula_model::XLNM_PRINT_AREA)
                || n.name
                    .eq_ignore_ascii_case(formula_model::XLNM_PRINT_TITLES)
        })
        .map(|n| (n.scope, n.name.clone(), n.refers_to.clone()))
        .collect();

    // Pass 1: sheet-scoped print names (canonical Excel encoding).
    // Pass 2: workbook-scoped print names (calamine fallback loses sheet scope).
    for pass in 0u8..=1u8 {
        for (scope, name, refers_to) in &builtins {
            let sheet_name = match (pass, scope) {
                (0, DefinedNameScope::Sheet(sheet_id)) => {
                    workbook.sheet(*sheet_id).map(|s| s.name.clone())
                }
                (1, DefinedNameScope::Workbook) => {
                    match infer_sheet_name_from_workbook_scoped_defined_name(
                        workbook, name, refers_to, warnings, suppressed,
                    ) {
                        WorkbookScopedDefinedNameSheetInference::Inferred(sheet_name) => {
                            Some(sheet_name)
                        }
                        WorkbookScopedDefinedNameSheetInference::NotApplicable
                        | WorkbookScopedDefinedNameSheetInference::Failure => None,
                    }
                }
                _ => None,
            };

            let Some(sheet_name) = sheet_name else {
                continue;
            };

            if name.eq_ignore_ascii_case(formula_model::XLNM_PRINT_AREA) {
                if workbook
                    .sheet_print_settings_by_name(&sheet_name)
                    .print_area
                    .is_some()
                {
                    continue;
                }
                if let Some(ranges) =
                    parse_print_area_refers_to(&sheet_name, refers_to, warnings, suppressed)
                {
                    workbook.set_sheet_print_area_by_name(&sheet_name, Some(ranges));
                }
            } else if name.eq_ignore_ascii_case(formula_model::XLNM_PRINT_TITLES) {
                if workbook
                    .sheet_print_settings_by_name(&sheet_name)
                    .print_titles
                    .is_some()
                {
                    continue;
                }
                if let Some(titles) =
                    parse_print_titles_refers_to(&sheet_name, refers_to, warnings, suppressed)
                {
                    workbook.set_sheet_print_titles_by_name(&sheet_name, Some(titles));
                }
            }
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ParsedA1Range {
    Cell(Range),
    Row(RowRange),
    Col(ColRange),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ParsedEndpoint {
    Cell(CellRef),
    Row(u32), // 0-based
    Col(u32), // 0-based
}

fn parse_print_area_refers_to(
    expected_sheet_name: &str,
    refers_to: &str,
    warnings: &mut ImportWarnings,
    suppressed: &mut bool,
) -> Option<Vec<Range>> {
    let refers_to = refers_to.trim();
    if refers_to.is_empty() {
        return None;
    }

    let areas = match split_print_name_areas(refers_to) {
        Ok(areas) => areas,
        Err(err) => {
            push_import_warning(
                warnings,
                format!(
                    "failed to parse `{}` for sheet `{expected_sheet_name}`: {err}",
                    formula_model::XLNM_PRINT_AREA
                ),
                suppressed,
            );
            return None;
        }
    };

    let mut ranges: Vec<Range> = Vec::new();
    for area in areas {
        let (sheet_name, range_str) = match split_print_name_sheet_ref(area) {
            Ok(v) => v,
            Err(err) => {
                push_import_warning(
                    warnings,
                    format!(
                        "skipping `{}` entry {area:?} for sheet `{expected_sheet_name}`: {err}",
                        formula_model::XLNM_PRINT_AREA
                    ),
                    suppressed,
                );
                continue;
            }
        };

        if let Some(found_sheet_name) = sheet_name.as_deref() {
            let found_sheet_name = strip_workbook_prefix_from_sheet_ref(found_sheet_name);
            if !sheet_name_eq_case_insensitive(found_sheet_name, expected_sheet_name) {
                push_import_warning(
                    warnings,
                    format!(
                        "skipping `{}` entry {area:?} for sheet `{expected_sheet_name}`: references different sheet `{found_sheet_name}`",
                        formula_model::XLNM_PRINT_AREA
                    ),
                    suppressed,
                );
                continue;
            }
        }

        match parse_print_name_range(range_str) {
            Ok(ParsedA1Range::Cell(range)) => ranges.push(range),
            Ok(ParsedA1Range::Row(_) | ParsedA1Range::Col(_)) => {
                push_import_warning(
                    warnings,
                    format!(
                        "skipping `{}` entry {area:?} for sheet `{expected_sheet_name}`: print area must be a cell range",
                        formula_model::XLNM_PRINT_AREA
                    ),
                    suppressed,
                );
            }
            Err(err) => push_import_warning(
                warnings,
                format!(
                    "skipping `{}` entry {area:?} for sheet `{expected_sheet_name}`: {err}",
                    formula_model::XLNM_PRINT_AREA
                ),
                suppressed,
            ),
        }
    }

    (!ranges.is_empty()).then_some(ranges)
}

fn parse_print_titles_refers_to(
    expected_sheet_name: &str,
    refers_to: &str,
    warnings: &mut ImportWarnings,
    suppressed: &mut bool,
) -> Option<PrintTitles> {
    let refers_to = refers_to.trim();
    if refers_to.is_empty() {
        return None;
    }

    let areas = match split_print_name_areas(refers_to) {
        Ok(areas) => areas,
        Err(err) => {
            push_import_warning(
                warnings,
                format!(
                    "failed to parse `{}` for sheet `{expected_sheet_name}`: {err}",
                    formula_model::XLNM_PRINT_TITLES
                ),
                suppressed,
            );
            return None;
        }
    };

    let mut titles = PrintTitles::default();
    for area in areas {
        let (sheet_name, range_str) = match split_print_name_sheet_ref(area) {
            Ok(v) => v,
            Err(err) => {
                push_import_warning(
                    warnings,
                    format!(
                        "skipping `{}` entry {area:?} for sheet `{expected_sheet_name}`: {err}",
                        formula_model::XLNM_PRINT_TITLES
                    ),
                    suppressed,
                );
                continue;
            }
        };

        if let Some(found_sheet_name) = sheet_name.as_deref() {
            let found_sheet_name = strip_workbook_prefix_from_sheet_ref(found_sheet_name);
            if !sheet_name_eq_case_insensitive(found_sheet_name, expected_sheet_name) {
                push_import_warning(
                    warnings,
                    format!(
                        "skipping `{}` entry {area:?} for sheet `{expected_sheet_name}`: references different sheet `{found_sheet_name}`",
                        formula_model::XLNM_PRINT_TITLES
                    ),
                    suppressed,
                );
                continue;
            }
        }

        match parse_print_name_range(range_str) {
            Ok(ParsedA1Range::Row(rows)) => titles.repeat_rows = Some(rows),
            Ok(ParsedA1Range::Col(cols)) => titles.repeat_cols = Some(cols),
            Ok(ParsedA1Range::Cell(range)) => {
                // Some decoders (e.g. calamine `.xls` defined-name fallback) represent whole-row/
                // whole-column print titles as explicit cell ranges (`$A$1:$IV$1`, `$A$1:$A$65536`)
                // rather than row/col-only references (`$1:$1`, `$A:$A`).
                if range.start.row == range.end.row && range.start.col != range.end.col {
                    titles.repeat_rows = Some(RowRange {
                        start: range.start.row,
                        end: range.end.row,
                    });
                } else if range.start.col == range.end.col && range.start.row != range.end.row {
                    titles.repeat_cols = Some(ColRange {
                        start: range.start.col,
                        end: range.end.col,
                    });
                } else {
                    push_import_warning(
                        warnings,
                        format!(
                            "skipping `{}` entry {area:?} for sheet `{expected_sheet_name}`: print titles must be a row or column range",
                            formula_model::XLNM_PRINT_TITLES
                        ),
                        suppressed,
                    );
                }
            }
            Err(err) => push_import_warning(
                warnings,
                format!(
                    "skipping `{}` entry {area:?} for sheet `{expected_sheet_name}`: {err}",
                    formula_model::XLNM_PRINT_TITLES
                ),
                suppressed,
            ),
        }
    }

    (titles.repeat_rows.is_some() || titles.repeat_cols.is_some()).then_some(titles)
}

fn split_print_name_areas(formula: &str) -> Result<Vec<&str>, String> {
    // Sheet names may be quoted (single quotes) and can contain commas. Split on commas only when
    // not inside quotes.
    let mut parts = Vec::new();
    let mut start = 0usize;
    let mut in_quotes = false;
    let bytes = formula.as_bytes();
    let mut i = 0usize;

    while i < bytes.len() {
        match bytes[i] {
            b'\'' => {
                if in_quotes {
                    if i + 1 < bytes.len() && bytes[i + 1] == b'\'' {
                        // Escaped quote in a sheet name.
                        i += 1;
                    } else {
                        in_quotes = false;
                    }
                } else {
                    in_quotes = true;
                }
            }
            b',' if !in_quotes => {
                let part = formula[start..i].trim();
                if !part.is_empty() {
                    parts.push(part);
                }
                start = i + 1;
            }
            _ => {}
        }

        i += 1;
    }

    let part = formula[start..].trim();
    if !part.is_empty() {
        parts.push(part);
    }

    Ok(parts)
}

/// Split an area reference like:
/// - `Sheet1!$A$1:$B$2`
/// - `'Sheet One'!$1:$1`
///
/// into `(sheet_name, range_str)`.
///
/// Returns `sheet_name=None` when the reference has no explicit `Sheet!` prefix.
fn split_print_name_sheet_ref(input: &str) -> Result<(Option<String>, &str), String> {
    let input = input.trim();
    if input.is_empty() {
        return Err("empty reference".to_string());
    }

    // Some rgce decoders (and modern Excel) can include an explicit implicit-intersection operator
    // `@` before a reference (e.g. `@Sheet1!A1:B2`). Built-in defined names like Print_Area /
    // Print_Titles should still be recognized in this form, so strip a leading `@` if present.
    let input = input.strip_prefix('@').unwrap_or(input).trim();

    let bytes = input.as_bytes();
    if bytes.first() == Some(&b'\'') {
        // Quoted sheet names may contain escaped quotes (`''` represents a literal `'`).
        //
        // Avoid interpreting raw bytes as chars here: sheet names can contain non-ASCII UTF-8.
        // Stitch together UTF-8 slices instead.
        let mut sheet = String::new();
        let mut i = 1usize;
        let mut seg_start = i;

        while i < bytes.len() {
            if bytes[i] == b'\'' {
                if i + 1 < bytes.len() && bytes[i + 1] == b'\'' {
                    // Escaped quote in a sheet name.
                    sheet.push_str(&input[seg_start..i]);
                    sheet.push('\'');
                    i += 2;
                    seg_start = i;
                    continue;
                }

                // End of quoted sheet name.
                sheet.push_str(&input[seg_start..i]);
                if i + 1 >= bytes.len() || bytes[i + 1] != b'!' {
                    return Err(format!("expected ! after quoted sheet name in {input:?}"));
                }

                let rest = &input[(i + 2)..];
                return Ok((Some(sheet), rest));
            }

            i += 1;
        }

        return Err(format!("unterminated quoted sheet name in {input:?}"));
    }

    let Some(idx) = input.find('!') else {
        return Ok((None, input));
    };

    Ok((Some(input[..idx].to_string()), &input[(idx + 1)..]))
}

fn strip_workbook_prefix_from_sheet_ref(sheet_name: &str) -> &str {
    // Best-effort: Excel can serialize sheet references as `'[Book1.xlsx]Sheet1'!A1`.
    // If a workbook prefix exists, keep only the portion after the last `]`.
    sheet_name
        .rfind(']')
        .and_then(|idx| (idx + 1 <= sheet_name.len()).then_some(&sheet_name[(idx + 1)..]))
        .filter(|s| !s.is_empty())
        .unwrap_or(sheet_name)
}

fn parse_print_name_range(ref_str: &str) -> Result<ParsedA1Range, String> {
    // Allow an explicit implicit-intersection prefix (`@A1:B2`, `@Sheet1!A1:B2`), as produced by
    // some rgce decoders for value-class range tokens.
    let ref_str = ref_str.trim().strip_prefix('@').unwrap_or(ref_str.trim());
    if ref_str.is_empty() {
        return Err("empty range".to_string());
    }

    let (start, end) = match ref_str.split_once(':') {
        Some((a, b)) => (a, b),
        None => (ref_str, ref_str),
    };

    let start = parse_print_name_endpoint(start)?;
    let end = parse_print_name_endpoint(end)?;

    match (start, end) {
        (ParsedEndpoint::Cell(a), ParsedEndpoint::Cell(b)) => {
            Ok(ParsedA1Range::Cell(Range::new(a, b)))
        }
        (ParsedEndpoint::Row(a), ParsedEndpoint::Row(b)) => {
            Ok(ParsedA1Range::Row(RowRange { start: a, end: b }))
        }
        (ParsedEndpoint::Col(a), ParsedEndpoint::Col(b)) => {
            Ok(ParsedA1Range::Col(ColRange { start: a, end: b }))
        }
        _ => Err(format!("mismatched range endpoints in {ref_str:?}")),
    }
}

fn parse_print_name_endpoint(s: &str) -> Result<ParsedEndpoint, String> {
    let trimmed = s.trim().trim_matches('$');
    if trimmed.is_empty() {
        return Err("empty endpoint".to_string());
    }

    let mut letters = String::new();
    let mut digits = String::new();

    for ch in trimmed.chars() {
        if ch == '$' {
            continue;
        }
        if ch.is_ascii_alphabetic() && digits.is_empty() {
            letters.push(ch);
        } else if ch.is_ascii_digit() {
            digits.push(ch);
        } else {
            return Err(format!("invalid character {ch:?} in endpoint {s:?}"));
        }
    }

    match (letters.is_empty(), digits.is_empty()) {
        (false, false) => {
            let cell_ref = format!("{letters}{digits}");
            let cell = CellRef::from_a1(&cell_ref)
                .map_err(|err| format!("invalid cell reference in endpoint {s:?}: {err}"))?;
            Ok(ParsedEndpoint::Cell(cell))
        }
        (false, true) => {
            let col = parse_col_letters_to_index(&letters)?;
            Ok(ParsedEndpoint::Col(col))
        }
        (true, false) => {
            let row_1_based: u32 = digits
                .parse()
                .map_err(|_| format!("invalid row number in endpoint {s:?}"))?;
            if row_1_based == 0 {
                return Err(format!("invalid row number in endpoint {s:?}"));
            }
            Ok(ParsedEndpoint::Row(row_1_based - 1))
        }
        (true, true) => Err(format!("invalid endpoint {s:?}")),
    }
}

fn parse_col_letters_to_index(letters: &str) -> Result<u32, String> {
    let mut col: u32 = 0;
    for ch in letters.chars() {
        if !ch.is_ascii_alphabetic() {
            return Err(format!("invalid column letters {letters:?}"));
        }
        let v = (ch.to_ascii_uppercase() as u8 - b'A') as u32 + 1;
        col = col
            .checked_mul(26)
            .and_then(|c| c.checked_add(v))
            .ok_or_else(|| format!("invalid column letters {letters:?}"))?;
    }
    if col == 0 {
        return Err(format!("invalid column letters {letters:?}"));
    }
    Ok(col - 1)
}

fn sheet_name_taken(candidate: &str, existing_names: &[String]) -> bool {
    existing_names
        .iter()
        .any(|existing| sheet_name_eq_case_insensitive(existing, candidate))
}

/// Best-effort sanitization for legacy `.xls` sheet names.
///
/// Excel sheet names have a number of restrictions (see [`formula_model::validate_sheet_name`]).
/// Calamine may still surface corrupt/non-compliant names from malformed BIFF files; this helper
/// attempts to produce a deterministic, valid, unique name for the destination workbook.
///
/// This is part of the public API only so it can be tested from `crates/formula-xls/tests/`.
#[doc(hidden)]
pub fn sanitize_sheet_name(
    original: &str,
    sheet_number: usize,
    existing_names: &[String],
) -> String {
    let without_nuls = original.replace('\0', "");
    let trimmed = without_nuls.trim();

    let mut cleaned = String::with_capacity(trimmed.len());
    for ch in trimmed.chars() {
        if matches!(ch, ':' | '\\' | '/' | '?' | '*' | '[' | ']') {
            cleaned.push('_');
        } else {
            cleaned.push(ch);
        }
    }

    // `formula_model::validate_sheet_name` forbids leading/trailing apostrophes.
    let cleaned = cleaned.trim_matches('\'');

    let mut base = if cleaned.trim().is_empty() {
        format!("Sheet{sheet_number}")
    } else {
        truncate_to_utf16_len(cleaned, EXCEL_MAX_SHEET_NAME_LEN)
    };

    // Truncation can re-introduce a trailing apostrophe (`foo'bar...` → `foo'`).
    // Excel forbids sheet names that begin or end with `'`.
    base = truncate_to_utf16_len(base.trim_matches('\''), EXCEL_MAX_SHEET_NAME_LEN);
    if base.trim().is_empty() {
        base = truncate_to_utf16_len(&format!("Sheet{sheet_number}"), EXCEL_MAX_SHEET_NAME_LEN);
    }

    if !sheet_name_taken(&base, existing_names) {
        return base;
    }

    for suffix_index in 2usize.. {
        let suffix = format!(" ({suffix_index})");
        let suffix_len = suffix.encode_utf16().count();
        let max_base_len = EXCEL_MAX_SHEET_NAME_LEN.saturating_sub(suffix_len);

        let mut candidate = truncate_to_utf16_len(&base, max_base_len);
        candidate.push_str(&suffix);
        let candidate = truncate_to_utf16_len(&candidate, EXCEL_MAX_SHEET_NAME_LEN);

        if !sheet_name_taken(&candidate, existing_names) {
            return candidate;
        }
    }

    unreachable!("suffix loop should always return a unique sheet name");
}

/// Mask the BIFF `FILEPASS` record id (0x002F) in the workbook globals substream.
///
/// This is intended for callers that have already decrypted an encrypted `.xls` workbook stream:
/// decrypted BIFF streams still contain the `FILEPASS` record header, but downstream BIFF parsers
/// (and `calamine`) treat `FILEPASS` as an encryption terminator and stop scanning.
///
/// Masking `FILEPASS` to a reserved/unknown record id allows parsers to skip it and continue.
///
/// Returns the number of record headers that were masked (normally 0 or 1).
///
/// This helper is part of the public API only so it can be exercised from integration tests in
/// `crates/formula-xls/tests/` and used by higher-level decryption plumbing.
#[doc(hidden)]
pub fn mask_biff_filepass_record_id(workbook_stream: &mut [u8]) -> usize {
    biff::records::mask_workbook_globals_filepass_record_id_in_place(workbook_stream)
}

/// Parse worksheet merged regions from BIFF `MERGEDCELLS` records.
///
/// This helper is part of the public API only so it can be exercised from integration tests in
/// `crates/formula-xls/tests/`. Most callers should use [`import_xls_path`] instead.
#[doc(hidden)]
pub fn parse_biff_sheet_merged_cells(
    workbook_stream: &[u8],
    start: usize,
) -> Result<Vec<Range>, String> {
    Ok(biff::parse_biff_sheet_merged_cells(workbook_stream, start)?.ranges)
}

fn normalize_sheet_name_for_match(name: &str) -> String {
    name.replace('\0', "").trim().to_lowercase()
}

fn reconcile_biff_sheet_mapping(
    calamine_sheets: &[Sheet],
    biff_sheets: Option<&[biff::BoundSheetInfo]>,
) -> (Vec<Option<usize>>, Option<String>) {
    let Some(biff_sheets) = biff_sheets else {
        return (vec![None; calamine_sheets.len()], None);
    };
    if biff_sheets.is_empty() {
        return (vec![None; calamine_sheets.len()], None);
    }

    let calamine_count = calamine_sheets.len();
    let biff_count = biff_sheets.len();

    // Primary mapping: BIFF BoundSheet order (workbook order) should align with calamine.
    let mut index_mapping = vec![None; calamine_count];
    let mut index_used_biff = vec![false; biff_count];
    for idx in 0..calamine_count.min(biff_count) {
        index_mapping[idx] = Some(idx);
        index_used_biff[idx] = true;
    }

    // Secondary mapping: normalized, case-insensitive name match.
    let mut name_mapping = vec![None; calamine_count];
    let mut name_used_biff = vec![false; biff_count];

    let mut biff_by_name: HashMap<String, Vec<usize>> = HashMap::new();
    for (idx, sheet) in biff_sheets.iter().enumerate() {
        biff_by_name
            .entry(normalize_sheet_name_for_match(&sheet.name))
            .or_default()
            .push(idx);
    }

    for (cal_idx, sheet) in calamine_sheets.iter().enumerate() {
        let key = normalize_sheet_name_for_match(&sheet.name);
        let Some(candidates) = biff_by_name.get(&key) else {
            continue;
        };

        if let Some(&biff_idx) = candidates.iter().find(|&&idx| !name_used_biff[idx]) {
            name_mapping[cal_idx] = Some(biff_idx);
            name_used_biff[biff_idx] = true;
        }
    }

    let index_mapped = index_mapping.iter().filter(|m| m.is_some()).count();
    let name_mapped = name_mapping.iter().filter(|m| m.is_some()).count();
    let index_score = index_mapping
        .iter()
        .enumerate()
        .filter_map(|(cal_idx, mapped)| {
            let biff_idx = (*mapped)?;
            let cal_name = normalize_sheet_name_for_match(&calamine_sheets[cal_idx].name);
            let biff_name = normalize_sheet_name_for_match(&biff_sheets[biff_idx].name);
            (cal_name == biff_name).then_some(())
        })
        .count();
    let name_score = name_mapping
        .iter()
        .enumerate()
        .filter_map(|(cal_idx, mapped)| {
            let biff_idx = (*mapped)?;
            let cal_name = normalize_sheet_name_for_match(&calamine_sheets[cal_idx].name);
            let biff_name = normalize_sheet_name_for_match(&biff_sheets[biff_idx].name);
            (cal_name == biff_name).then_some(())
        })
        .count();

    let (mapping, used_biff, strategy) = match name_mapped.cmp(&index_mapped) {
        std::cmp::Ordering::Greater => (name_mapping, name_used_biff, "name"),
        std::cmp::Ordering::Less => (index_mapping, index_used_biff, "index"),
        std::cmp::Ordering::Equal => {
            if name_score > index_score {
                (name_mapping, name_used_biff, "name")
            } else {
                (index_mapping, index_used_biff, "index")
            }
        }
    };

    let mapped_pairs: Vec<(usize, usize)> = mapping
        .iter()
        .enumerate()
        .filter_map(|(cal_idx, &biff_idx)| biff_idx.map(|biff_idx| (cal_idx, biff_idx)))
        .collect();
    let unmapped_calamine: Vec<usize> = mapping
        .iter()
        .enumerate()
        .filter_map(|(idx, mapped)| mapped.is_none().then_some(idx))
        .collect();
    let unmapped_biff: Vec<usize> = used_biff
        .iter()
        .enumerate()
        .filter_map(|(idx, used)| (!*used).then_some(idx))
        .collect();

    let calamine_names: Vec<&str> = calamine_sheets.iter().map(|s| s.name.as_str()).collect();
    let biff_names: Vec<&str> = biff_sheets.iter().map(|s| s.name.as_str()).collect();

    let should_warn = calamine_count != biff_count
        || strategy != "index"
        || !unmapped_calamine.is_empty()
        || !unmapped_biff.is_empty();

    let warning = should_warn.then(|| {
        format!(
            "failed to reconcile `.xls` sheet metadata (strategy={strategy}): calamine sheets ({calamine_count}) {calamine_names:?}; BIFF BoundSheet records ({biff_count}) {biff_names:?}; mapped indices (calamine->BIFF) {mapped_pairs:?}; unmapped calamine indices {unmapped_calamine:?}; unmapped BIFF indices {unmapped_biff:?}"
        )
    });

    (mapping, warning)
}

fn build_sheet_names_by_biff_idx(
    biff_sheets: Option<&[biff::BoundSheetInfo]>,
    sheet_mapping: &[Option<usize>],
    final_sheet_names_by_idx: &[String],
) -> Vec<String> {
    let mut sheet_names_by_biff_idx: Vec<String> = biff_sheets
        .unwrap_or_default()
        .iter()
        .map(|s| s.name.clone())
        .collect();

    for (cal_idx, maybe_biff_idx) in sheet_mapping.iter().enumerate() {
        let Some(biff_idx) = *maybe_biff_idx else {
            continue;
        };
        let Some(final_name) = final_sheet_names_by_idx.get(cal_idx) else {
            continue;
        };
        if biff_idx < sheet_names_by_biff_idx.len() {
            sheet_names_by_biff_idx[biff_idx] = final_name.clone();
        }
    }

    sheet_names_by_biff_idx
}

fn import_biff8_shared_formulas(
    workbook: &mut Workbook,
    workbook_stream: &[u8],
    biff_sheets: &[biff::BoundSheetInfo],
    sheet_names_by_biff_idx: &[String],
    sheet_ids_by_biff_idx: &[Option<formula_model::WorksheetId>],
    rgce_decoder: &biff::worksheet_formulas::WorksheetFormulaDecoder,
    warnings: &mut ImportWarnings,
    suppressed: &mut bool,
) {
    // BIFF8 shared formulas are encoded as:
    // - FORMULA records whose rgce is a `PtgExp` token referencing the anchor cell, and
    // - a SHRFMLA record following the anchor FORMULA record that contains the shared rgce stream.
    //
    // Calamine's `.xls` formula API (`worksheet_formula`) can omit formulas for cells that only
    // contain `PtgExp`, so we recover them directly from the workbook stream.
    //
    // This currently focuses on SHRFMLA-based shared formulas; unsupported tokens are rendered
    // best-effort by the rgce decoder.
    const RECORD_FORMULA: u16 = 0x0006;
    const RECORD_SHRFMLA: u16 = 0x04BC;
    const PTG_EXP: u8 = 0x01;

    let ctx = rgce_decoder.rgce_decode_context(sheet_names_by_biff_idx);

    for (biff_idx, sheet_info) in biff_sheets.iter().enumerate() {
        let Some(sheet_id) = sheet_ids_by_biff_idx.get(biff_idx).copied().flatten() else {
            continue;
        };
        let Some(sheet) = workbook.sheet_mut(sheet_id) else {
            continue;
        };

        if sheet_info.offset >= workbook_stream.len() {
            push_import_warning(
                warnings,
                format!(
                    "failed to import `.xls` shared formulas for sheet `{}`: out-of-bounds stream offset {}",
                    sheet.name, sheet_info.offset
                ),
                suppressed,
            );
            continue;
        }

        let allows_continuation = |id: u16| id == RECORD_SHRFMLA;
        let Ok(iter) = biff::records::LogicalBiffRecordIter::from_offset(
            workbook_stream,
            sheet_info.offset,
            allows_continuation,
        ) else {
            push_import_warning(
                warnings,
                format!(
                    "failed to import `.xls` shared formulas for sheet `{}`: invalid substream offset {}",
                    sheet.name, sheet_info.offset
                ),
                suppressed,
            );
            continue;
        };

        #[derive(Debug, Clone)]
        struct SharedFormulaDef {
            rgce: Vec<u8>,
            rgcb: Vec<u8>,
        }

        // Map anchor (row,col) -> shared formula definition (rgce + trailing rgcb bytes).
        let mut shared_by_anchor: HashMap<(u16, u16), SharedFormulaDef> = HashMap::new();
        // Collect (cell_row, cell_col, anchor_row, anchor_col) for PtgExp formulas.
        let mut ptgexp_cells: Vec<(u16, u16, u16, u16)> = Vec::new();
        let mut last_formula_cell: Option<(u16, u16)> = None;

        for record in iter {
            let record = match record {
                Ok(r) => r,
                Err(err) => {
                    push_import_warning(
                        warnings,
                        format!(
                            "failed to import `.xls` shared formulas for sheet `{}`: malformed BIFF record: {err}",
                            sheet.name
                        ),
                        suppressed,
                    );
                    break;
                }
            };

            // Stop at the next substream BOF; worksheet substream begins at `sheet_info.offset`.
            if record.offset != sheet_info.offset && biff::records::is_bof_record(record.record_id)
            {
                break;
            }

            match record.record_id {
                RECORD_FORMULA => {
                    let data = record.data.as_ref();
                    if data.len() < 22 {
                        continue;
                    }
                    let row = u16::from_le_bytes([data[0], data[1]]);
                    let col = u16::from_le_bytes([data[2], data[3]]);
                    last_formula_cell = Some((row, col));

                    let cce = u16::from_le_bytes([data[20], data[21]]) as usize;
                    if cce == 0 {
                        continue;
                    }
                    if data.len() < 22 + cce {
                        continue;
                    }
                    let rgce = &data[22..22 + cce];
                    if rgce.len() >= 5 && rgce[0] == PTG_EXP {
                        let anchor_row = u16::from_le_bytes([rgce[1], rgce[2]]);
                        let anchor_col = u16::from_le_bytes([rgce[3], rgce[4]]);
                        ptgexp_cells.push((row, col, anchor_row, anchor_col));
                    }
                }
                RECORD_SHRFMLA => {
                    let Some((anchor_row, anchor_col)) = last_formula_cell.take() else {
                        continue;
                    };
                    // Use the fragment-aware SHRFMLA parser so BIFF8 `PtgStr` literals continued
                    // across `CONTINUE` boundaries have the extra 1-byte continued-segment option
                    // flags stripped out before downstream rgce decoding.
                    let (rgce, rgcb) =
                        match biff::worksheet_formulas::parse_biff8_shrfmla_record_with_rgcb(&record)
                        {
                            Ok(parsed) => parsed,
                            Err(_) => continue,
                        };
                    shared_by_anchor.insert(
                        (anchor_row, anchor_col),
                        SharedFormulaDef { rgce, rgcb },
                    );
                }
                biff::records::RECORD_EOF => break,
                _ => {}
            }
        }

        if ptgexp_cells.is_empty() || shared_by_anchor.is_empty() {
            continue;
        }

        for (row, col, anchor_row, anchor_col) in ptgexp_cells {
            let cell_ref = CellRef::new(row as u32, col as u32);
            let anchor_cell = sheet.merged_regions.resolve_cell(cell_ref);
            let existing = sheet.formula(anchor_cell);

            // Best-effort: allow overriding formulas for PtgExp-backed cells when we can resolve the
            // backing SHRFMLA token stream.
            //
            // Calamine can omit formulas for PtgExp follower cells entirely, or surface placeholder
            // errors (`#UNKNOWN!` / `#REF!`) when it cannot resolve the shared formula definition.
            //
            // For merged regions, only overwrite the anchor when it is currently unset or just a
            // placeholder; otherwise preserve the existing anchor formula.
            if anchor_cell != cell_ref
                && existing.is_some_and(|f| {
                    f != ErrorValue::Unknown.as_str()
                        && f != ErrorValue::Ref.as_str()
                        && !is_calamine_unrecognised_formula_placeholder(f)
                })
            {
                continue;
            }
            let Some(shared) = shared_by_anchor.get(&(anchor_row, anchor_col)) else {
                continue;
            };
            let shared_rgce = &shared.rgce;
            let shared_rgcb = &shared.rgcb;

            // SHRFMLA stores the shared `rgce` stream in the anchor cell's coordinate space.
            // Materialize the token stream for the current cell by shifting any relative
            // `PtgRef`/`PtgArea` (and 3D variants) by the (row,col) delta.
            let materialized = if row == anchor_row && col == anchor_col {
                None
            } else {
                biff::formulas::materialize_biff8_rgce(
                    shared_rgce,
                    anchor_row as u32,
                    anchor_col as u32,
                    row as u32,
                    col as u32,
                )
            };

            let rgce_bytes = match materialized.as_deref() {
                Some(bytes) => bytes,
                None => {
                    // Best-effort: if we can't materialize the shared token stream (due to an
                    // unsupported/malformed token), fall back to decoding the anchor token stream
                    // as-is so the cell still has a recoverable formula.
                    if row != anchor_row || col != anchor_col {
                        push_import_warning(
                            warnings,
                            format!(
                                "failed to import `.xls` shared formula in sheet `{}` at {}: unsupported or malformed shared rgce token stream",
                                sheet.name,
                                cell_ref.to_a1()
                            ),
                            suppressed,
                        );
                    }
                    shared_rgce
                }
            };

            let decoded = biff::rgce::decode_biff8_rgce_with_base_and_rgcb(
                rgce_bytes,
                shared_rgcb,
                &ctx,
                Some(biff::rgce::CellCoord::new(row as u32, col as u32)),
            );

            for warning in decoded.warnings {
                push_import_warning(
                    warnings,
                    format!(
                        "failed to import `.xls` shared formula in sheet `{}` at {}: {warning}",
                        sheet.name,
                        cell_ref.to_a1()
                    ),
                    suppressed,
                );
            }

            if decoded.text.is_empty() {
                continue;
            }

            // Prefer our BIFF-derived shared formula materialization over calamine’s worksheet
            // formula string for `PtgExp` cells, but avoid overwriting an existing non-unknown
            // formula with our own `#UNKNOWN!` best-effort output.
            let decoded_is_unknown = decoded
                .text
                .trim_start()
                .starts_with(ErrorValue::Unknown.as_str());
            let should_set = match existing {
                None => true,
                Some(existing) => {
                    if existing == decoded.text {
                        false
                    } else if decoded_is_unknown
                        && !existing
                            .trim_start()
                            .starts_with(ErrorValue::Unknown.as_str())
                        && !is_calamine_unrecognised_formula_placeholder(existing)
                    {
                        false
                    } else {
                        true
                    }
                }
            };

            if should_set {
                if let Some(normalized) = normalize_formula_text(&decoded.text) {
                    sheet.set_formula(anchor_cell, Some(normalized));
                }
            }
        }
    }
}

fn sanitize_biff8_stream_for_calamine(stream: &[u8]) -> Option<Vec<u8>> {
    let mut sanitized = sanitize_biff8_continued_name_records_for_calamine(stream);
    let base = sanitized.as_deref().unwrap_or(stream);
    if let Some(patched) = sanitize_biff8_wide_ptgexp_formulas_for_calamine(base) {
        sanitized = Some(patched);
    }
    sanitized
}

fn sanitize_biff8_wide_ptgexp_formulas_for_calamine(stream: &[u8]) -> Option<Vec<u8>> {
    // Calamine assumes BIFF8 `PtgExp`/`PtgTbl` payloads are always 4 bytes (rw u16 + col u16).
    // Some producers emit wider row/col widths even in `.xls` files, which can cause calamine to
    // misalign the token stream and panic.
    //
    // Work around this by clamping `FORMULA.cce` to 5 bytes (ptg + 4-byte payload) for formulas
    // whose `rgce` begins with `PtgExp`/`PtgTbl` and whose `cce` exceeds the BIFF8 minimum. This
    // keeps record sizes stable (no shifting) while avoiding calamine's token-parser panic. We
    // still recover the correct formula text later via our own BIFF parser using the original
    // workbook stream.
    const RECORD_FORMULA: u16 = 0x0006;
    const PTG_EXP: u8 = 0x01;
    const PTG_TBL: u8 = 0x02;

    let mut out: Option<Vec<u8>> = None;
    let mut iter = biff::records::BiffRecordIter::from_offset(stream, 0).ok()?;

    while let Some(next) = iter.next() {
        let record = match next {
            Ok(r) => r,
            Err(_) => break,
        };
        if record.record_id != RECORD_FORMULA {
            continue;
        }
        let data = record.data;
        if data.len() < 22 {
            continue;
        }
        let cce = u16::from_le_bytes([data[20], data[21]]) as usize;
        if cce <= 5 {
            continue;
        }
        if data.len() < 22 + cce {
            continue;
        }
        let rgce = &data[22..22 + cce];
        let Some(&ptg) = rgce.first() else {
            continue;
        };
        if ptg != PTG_EXP && ptg != PTG_TBL {
            continue;
        }

        let out = out.get_or_insert_with(|| stream.to_vec());
        let cce_offset = record.offset + 4 + 20;
        if cce_offset + 2 <= out.len() {
            out[cce_offset..cce_offset + 2].copy_from_slice(&5u16.to_le_bytes());
        }
    }

    out
}

fn sanitize_biff8_formula_records_for_calamine(stream: &[u8]) -> Option<Vec<u8>> {
    // Calamine can panic while decoding BIFF8 formulas. One observed case is shared-formula
    // materialization near BIFF8 row/col bounds where shifting a `PtgRef`/`PtgArea` row overflows a
    // `u16` during `PtgExp` resolution.
    //
    // To keep `.xls` import best-effort, sanitize all BIFF8 FORMULA rgce streams for calamine by
    // replacing them with a minimal constant expression (`0`). This prevents calamine from parsing
    // or materializing any formula token streams while still preserving:
    // - cached formula results (cell values), and
    // - sheet structure (row/col coordinates).
    //
    // We still import formulas ourselves from the raw BIFF stream, so correctness does not depend
    // on calamine's formula strings when this sanitizer is applied.
    const RECORD_FORMULA: u16 = 0x0006;

    // Minimal rgce for `0`: PtgInt(0).
    const SAFE_RGCE: [u8; 3] = [0x1E, 0x00, 0x00];

    let mut out: Option<Vec<u8>> = None;
    let mut iter = biff::records::BiffRecordIter::from_offset(stream, 0).ok()?;

    while let Some(next) = iter.next() {
        let record = match next {
            Ok(r) => r,
            Err(_) => break,
        };
        if record.record_id != RECORD_FORMULA {
            continue;
        }
        let data = record.data;
        if data.len() < 22 {
            continue;
        }

        let out = out.get_or_insert_with(|| stream.to_vec());
        let cce_offset = record.offset + 4 + 20;
        let rgce_offset = record.offset + 4 + 22;
        if cce_offset + 2 > out.len() {
            continue;
        }

        // If the record is long enough, patch the rgce bytes; otherwise set cce=0 so calamine
        // skips parsing.
        if rgce_offset + SAFE_RGCE.len() <= out.len() && data.len() >= 22 + SAFE_RGCE.len() {
            out[cce_offset..cce_offset + 2].copy_from_slice(&(SAFE_RGCE.len() as u16).to_le_bytes());
            out[rgce_offset..rgce_offset + SAFE_RGCE.len()].copy_from_slice(&SAFE_RGCE);
        } else {
            out[cce_offset..cce_offset + 2].copy_from_slice(&0u16.to_le_bytes());
        }
    }

    out
}

fn sanitize_biff8_continued_name_records_for_calamine(stream: &[u8]) -> Option<Vec<u8>> {
    const RECORD_NAME: u16 = 0x0018;
    const RECORD_CONTINUE: u16 = 0x003C;
    const NAME_FLAG_BUILTIN: u16 = 0x0020;
    // BIFF record id reserved for "unknown" sanitization. Any value that calamine doesn't treat as
    // a special record is fine; we use 0xFFFF which is not a defined BIFF record id.
    const RECORD_MASKED: u16 = 0xFFFF;

    // Calamine's NAME parser reads:
    // - cch (u8) at offset 3 in the NAME payload
    // - cce (u16) at offset 4 in the NAME payload
    //
    // It can panic when:
    // - the NAME record is continued and `cce` exceeds the first physical fragment length, or
    // - `cch` / `cce` claims bytes that don't fit in the physical record payload (corrupt files).
    //
    // To avoid this, sanitize NAME records that are continued *or* appear malformed:
    // - Coalesce consecutive CONTINUE records into the NAME record length so calamine sees a
    //   single contiguous payload slice.
    // - Compact away the embedded CONTINUE headers (and best-effort strip the extra "continued
    //   string segment" flags byte) so calamine can read the full name string.
    // - Patch `cce` to 0 so calamine skips parsing the `rgce` token stream.
    // - Clamp `cch` based on available bytes to prevent unchecked slice panics.
    //
    // This keeps the name string available (so `PtgName` tokens can still resolve) while making
    // calamine skip parsing the formula payload.
    let mut name_header_offsets: Vec<usize> = Vec::new();
    let mut name_mask_offsets: Vec<usize> = Vec::new();
    let mut offset: usize = 0;

    while offset + 4 <= stream.len() {
        let record_id = u16::from_le_bytes([stream[offset], stream[offset + 1]]);
        let len = u16::from_le_bytes([stream[offset + 2], stream[offset + 3]]) as usize;

        let data_start = match offset.checked_add(4) {
            Some(v) => v,
            None => break,
        };
        let next = match data_start.checked_add(len) {
            Some(v) => v,
            None => break,
        };
        if next > stream.len() {
            break;
        }

        if record_id == RECORD_NAME {
            let next_is_continue = next + 4 <= stream.len()
                && u16::from_le_bytes([stream[next], stream[next + 1]]) == RECORD_CONTINUE;

            // Calamine unconditionally slices `&r.data[14..]` for NAME parsing. If the physical
            // record payload is shorter than 14 bytes (or exactly 14 bytes with no following
            // CONTINUE record), calamine can panic via out-of-bounds slice indexing.
            //
            // For truncated NAME records that are *not* continued, we can't safely "fix" them
            // without consuming bytes from unrelated subsequent records (which would corrupt the
            // workbook stream). Instead, best-effort: mask the record id so calamine ignores it.
            //
            // If the truncated NAME is followed by a CONTINUE record, we can repair it by
            // coalescing the CONTINUE payload into the physical record length during patching.
            if len < 14 && !next_is_continue {
                name_mask_offsets.push(offset);
                offset = next;
                continue;
            }
            if len == 14 && !next_is_continue {
                name_mask_offsets.push(offset);
                offset = next;
                continue;
            }

            // Determine whether this NAME record needs patching.
            //
            // Even if we can't read the full header, `len < 14` / `len == 14` + CONTINUE is enough
            // to justify patching because calamine would otherwise panic.
            let mut needs_patch = next_is_continue || len < 14;

            // Best-effort: detect obviously out-of-bounds `cch`/`cce` values that would cause
            // calamine slice panics even without CONTINUE records.
            if len >= 6 && data_start + 6 <= stream.len() {
                let grbit = if len >= 2 && data_start + 2 <= stream.len() {
                    u16::from_le_bytes([stream[data_start], stream[data_start + 1]])
                } else {
                    0
                };
                let is_builtin = (grbit & NAME_FLAG_BUILTIN) != 0;
                let cch = stream[data_start + 3] as usize;
                let cce =
                    u16::from_le_bytes([stream[data_start + 4], stream[data_start + 5]]) as usize;

                if len >= 14 {
                    let available = len.saturating_sub(14);
                    if is_builtin {
                        // Built-in NAME layout: `rgchName` is `cch` bytes (usually a single-byte id).
                        if available < cch {
                            needs_patch = true;
                        }
                    } else {
                        // User-defined NAME layout: `rgchName` is an XLUnicodeStringNoCch.
                        //
                        // The first byte after the 14-byte header is a flags byte (compressed vs
                        // UTF-16LE). The character payload is `cch` bytes (compressed) or `2*cch`
                        // bytes (uncompressed).
                        //
                        // Calamine can panic when the physical record does not contain enough bytes
                        // for the declared `cch`, so clamp defensively.
                        if available == 0 {
                            needs_patch = true;
                        } else {
                            let flags = stream[data_start + 14];
                            let is_unicode = (flags & 0x01) != 0;
                            let required = 1usize.saturating_add(if is_unicode {
                                2usize.saturating_mul(cch)
                            } else {
                                cch
                            });
                            if available < required {
                                needs_patch = true;
                            }
                        }
                    }

                    // Best-effort: guard against obviously out-of-bounds `cce` values as well.
                    // Calamine slices `rgce` from the end of the physical record using `cce`.
                    if cce > len {
                        needs_patch = true;
                    }
                }
            }

            if needs_patch {
                name_header_offsets.push(offset);
            }
        }

        offset = next;
    }

    if name_header_offsets.is_empty() && name_mask_offsets.is_empty() {
        return None;
    }

    let mut out = stream.to_vec();
    for header_offset in name_header_offsets {
        let original_len =
            u16::from_le_bytes([out[header_offset + 2], out[header_offset + 3]]) as usize;
        let data_start = header_offset + 4;
        if data_start.saturating_add(original_len) > out.len() {
            continue;
        }

        // Determine whether this NAME record is built-in (name id bytes) or user-defined
        // (XLUnicodeStringNoCch).
        let grbit = if original_len >= 2 && data_start + 2 <= out.len() {
            u16::from_le_bytes([out[data_start], out[data_start + 1]])
        } else {
            0
        };
        let is_builtin = (grbit & NAME_FLAG_BUILTIN) != 0;

        // Coalesce consecutive CONTINUE record(s) into the NAME record's physical length so
        // calamine sees a single contiguous byte slice.
        //
        // This is especially important for NAME records whose first fragment ends at (or before)
        // the fixed 14-byte header: calamine unconditionally slices `&payload[14..]` and can panic
        // if it is empty.
        let mut continue_payloads: Vec<(usize, usize)> = Vec::new();
        let mut len = original_len;
        let mut cursor = data_start.saturating_add(original_len);
        while cursor + 4 <= out.len() {
            let id = u16::from_le_bytes([out[cursor], out[cursor + 1]]);
            if id != RECORD_CONTINUE {
                break;
            }
            let cont_len = u16::from_le_bytes([out[cursor + 2], out[cursor + 3]]) as usize;
            let total = 4usize.saturating_add(cont_len);
            if cursor.saturating_add(total) > out.len() {
                break;
            }
            if len.saturating_add(total) > u16::MAX as usize {
                break;
            }
            continue_payloads.push((cursor + 4, cont_len));
            len = len.saturating_add(total);
            cursor = cursor.saturating_add(total);
        }

        if !continue_payloads.is_empty() {
            out[header_offset + 2..header_offset + 4].copy_from_slice(&(len as u16).to_le_bytes());

            // Compact the physical fragments into a contiguous logical payload, leaving the NAME
            // header (14 bytes) intact. This removes the embedded CONTINUE record headers and
            // (best-effort) strips the continued-string option flags byte from subsequent fragments
            // so calamine can read the full name string without truncation/panics.
            let mut payload = vec![0u8; len];

            // Copy the header bytes we have. If the first fragment is truncated (<14 bytes), fill
            // missing header bytes with zeros; this matches common "all lengths are zero" layouts.
            let header_copy_len = original_len.min(14).min(len);
            payload[..header_copy_len]
                .copy_from_slice(&out[data_start..data_start.saturating_add(header_copy_len)]);

            // Copy any bytes present after the fixed header in the first fragment.
            let mut write_cursor = 14usize.min(len);
            if original_len > 14 && write_cursor < len {
                let src_start = data_start + 14;
                let src_end = data_start + original_len;
                let copy_len = src_end.saturating_sub(src_start).min(len - write_cursor);
                payload[write_cursor..write_cursor + copy_len]
                    .copy_from_slice(&out[src_start..src_start + copy_len]);
                write_cursor = write_cursor.saturating_add(copy_len);
            }

            // For user-defined names, the name string begins with a single flags byte at offset 14.
            // When that string is continued, BIFF8 CONTINUE fragments typically begin with an
            // additional 1-byte "continued segment" flags prefix. Calamine's NAME parser does not
            // handle these extra flags bytes, so strip them best-effort.
            let has_initial_string_flags = !is_builtin && original_len > 14;
            let mut is_first_continue = true;
            for &(cont_start, cont_len) in &continue_payloads {
                if write_cursor >= len {
                    break;
                }
                let mut skip = 0usize;
                if !is_builtin {
                    if has_initial_string_flags {
                        skip = 1;
                    } else if !is_first_continue {
                        skip = 1;
                    }
                }
                skip = skip.min(cont_len);
                let src_start = cont_start.saturating_add(skip);
                let src_end = cont_start.saturating_add(cont_len);
                if src_end > out.len() || src_start > src_end {
                    break;
                }
                let copy_len = (src_end - src_start).min(len - write_cursor);
                payload[write_cursor..write_cursor + copy_len]
                    .copy_from_slice(&out[src_start..src_start + copy_len]);
                write_cursor = write_cursor.saturating_add(copy_len);
                is_first_continue = false;
            }

            let payload_end = data_start.saturating_add(len);
            if payload_end <= out.len() {
                out[data_start..payload_end].copy_from_slice(&payload);
            }
        }

        // If we still don't have any bytes at `payload[14..]`, we can't prevent calamine from
        // panicking while indexing `buf[0]`. Mask the record id so calamine ignores it.
        if len <= 14 || data_start.saturating_add(len) > out.len() {
            name_mask_offsets.push(header_offset);
            continue;
        }

        // Patch `cce` (u16) at payload offset 4.
        if len >= 6 && data_start + 6 <= out.len() {
            out[data_start + 4..data_start + 6].copy_from_slice(&0u16.to_le_bytes());
        }

        // Best-effort: clamp `cch` if the name bytes cannot fit in the physical record payload.
        if len >= 4 && data_start + 4 <= out.len() {
            let grbit = if len >= 2 && data_start + 2 <= out.len() {
                u16::from_le_bytes([out[data_start], out[data_start + 1]])
            } else {
                0
            };
            let is_builtin = (grbit & NAME_FLAG_BUILTIN) != 0;
            let cch = out[data_start + 3] as usize;
            let available = len.saturating_sub(14);
            if is_builtin {
                // Built-in NAME: `rgchName` is stored as raw bytes (usually a single-byte id), so
                // we only require `available >= cch`.
                if available < cch {
                    out[data_start + 3] = available.min(u8::MAX as usize) as u8;
                }
            } else if available == 0 {
                out[data_start + 3] = 0;
            } else {
                // User-defined NAME: `rgchName` is stored as XLUnicodeStringNoCch. It starts with a
                // flags byte (compressed vs UTF-16LE). Clamp `cch` based on the bytes available.
                let flags = out[data_start + 14];
                let is_unicode = (flags & 0x01) != 0;
                let max_cch = if is_unicode {
                    available.saturating_sub(1) / 2
                } else {
                    available.saturating_sub(1)
                };
                if cch > max_cch {
                    out[data_start + 3] = max_cch.min(u8::MAX as usize) as u8;
                }
            }
        }
    }

    for mask_offset in name_mask_offsets {
        out[mask_offset..mask_offset + 2].copy_from_slice(&RECORD_MASKED.to_le_bytes());
    }

    Some(out)
}

fn build_in_memory_xls(workbook_stream: &[u8]) -> Result<Vec<u8>, ImportError> {
    // Construct a minimal CFB container containing only the Workbook stream. This is sufficient
    // for calamine's `Xls` parser and avoids needing to modify the original file on disk.
    //
    // Calamine's internal CFB parser rejects v4 compound files whose root directory has no
    // mini-stream (root start sector = ENDOFCHAIN). To keep `.xls` import best-effort for large
    // workbooks, always build a v3 CFB container here.
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create_with_version(cfb::Version::V3, cursor)
        .map_err(|err| ImportError::Xls(calamine::XlsError::Io(err)))?;

    {
        let mut stream = ole
            .create_stream("Workbook")
            .map_err(|err| ImportError::Xls(calamine::XlsError::Io(err)))?;
        stream
            .write_all(workbook_stream)
            .map_err(|err| ImportError::Xls(calamine::XlsError::Io(err)))?;
    }

    Ok(ole.into_inner().into_inner())
}

fn strip_wrapping_parentheses(mut expr: &str) -> &str {
    // Strip a single layer of wrapping parentheses when they enclose the entire expression.
    //
    // This is needed because BIFF formulas can contain `PtgParen` tokens that preserve redundant
    // parentheses. When rendered to A1 text (either by our BIFF rgce decoder or by calamine), this
    // can produce formulas like `=(Sheet1!$A$1:$C$5)`, which we still want to recognize as a plain
    // range reference.
    loop {
        let trimmed = expr.trim();
        if trimmed.len() < 2 || !trimmed.starts_with('(') || !trimmed.ends_with(')') {
            return trimmed;
        }

        let mut in_single_quote = false;
        let mut in_double_quote = false;
        let mut depth: u32 = 0;

        let bytes = trimmed.as_bytes();
        let mut idx = 0usize;
        while idx < bytes.len() {
            match bytes[idx] {
                b'"' if !in_single_quote => {
                    in_double_quote = !in_double_quote;
                    idx += 1;
                }
                b'\'' if !in_double_quote => {
                    if in_single_quote {
                        // Inside sheet-name quoting, `''` escapes a single quote.
                        if bytes.get(idx + 1) == Some(&b'\'') {
                            idx += 2;
                        } else {
                            in_single_quote = false;
                            idx += 1;
                        }
                    } else {
                        in_single_quote = true;
                        idx += 1;
                    }
                }
                b'(' if !in_single_quote && !in_double_quote => {
                    depth = depth.saturating_add(1);
                    idx += 1;
                }
                b')' if !in_single_quote && !in_double_quote => {
                    depth = depth.saturating_sub(1);
                    // If the initial `(` closes before the end of the string, the outer
                    // parentheses do not wrap the entire expression.
                    if depth == 0 && idx != bytes.len().saturating_sub(1) {
                        return trimmed;
                    }
                    idx += 1;
                }
                _ => idx += 1,
            }
        }

        if depth != 0 {
            return trimmed;
        }

        expr = &trimmed[1..trimmed.len().saturating_sub(1)];
    }
}

fn parse_autofilter_range_from_defined_name(refers_to: &str) -> Result<Range, String> {
    let refers_to = refers_to.trim();
    let refers_to = refers_to.strip_prefix('=').unwrap_or(refers_to).trim();
    let refers_to = strip_wrapping_parentheses(refers_to);
    let refers_to = refers_to.strip_prefix('@').unwrap_or(refers_to).trim();
    let refers_to = strip_wrapping_parentheses(refers_to);
    if refers_to.is_empty() {
        return Err("empty `refers_to`".to_string());
    }

    // Reject union formulas like `Sheet1!$A$1:$A$2,Sheet1!$C$1:$C$2` (comma at the top level).
    // We track quote and parenthesis nesting so commas inside quoted sheet names or function
    // argument lists do not get misclassified.
    let mut in_single_quote = false;
    let mut in_double_quote = false;
    let mut paren_depth: u32 = 0;

    let bytes = refers_to.as_bytes();
    let mut idx = 0usize;
    while idx < bytes.len() {
        match bytes[idx] {
            b'"' if !in_single_quote => {
                in_double_quote = !in_double_quote;
                idx += 1;
            }
            b'\'' if !in_double_quote => {
                if in_single_quote {
                    // Inside sheet-name quoting, `''` escapes a single quote.
                    if bytes.get(idx + 1) == Some(&b'\'') {
                        idx += 2;
                    } else {
                        in_single_quote = false;
                        idx += 1;
                    }
                } else {
                    in_single_quote = true;
                    idx += 1;
                }
            }
            b'(' if !in_single_quote && !in_double_quote => {
                paren_depth = paren_depth.saturating_add(1);
                idx += 1;
            }
            b')' if !in_single_quote && !in_double_quote => {
                paren_depth = paren_depth.saturating_sub(1);
                idx += 1;
            }
            b',' if !in_single_quote && !in_double_quote && paren_depth == 0 => {
                return Err("unsupported union formula (multiple areas)".to_string());
            }
            _ => idx += 1,
        }
    }

    // Strip any sheet qualifier, leaving the trailing A1 range part.
    // This handles `Sheet1!$A$1:$C$10` and `'My Sheet'!$A$1:$C$10`.
    let a1 = refers_to
        .rsplit_once('!')
        .map(|(_, tail)| tail)
        .unwrap_or(refers_to)
        .trim();
    let a1 = strip_wrapping_parentheses(a1);

    Range::from_a1(a1).map_err(|err| format!("invalid A1 range `{a1}`: {err}"))
}

#[cfg(test)]
mod tests {
    use super::*;

    use std::io::{Cursor, Write};

    use calamine::Data;
    use calamine::Xls;

    const RECORD_BOF: u16 = 0x0809;
    const RECORD_EOF: u16 = 0x000A;
    const RECORD_CODEPAGE: u16 = 0x0042;
    const RECORD_WINDOW1: u16 = 0x003D;
    const RECORD_FONT: u16 = 0x0031;
    const RECORD_XF: u16 = 0x00E0;
    const RECORD_BOUNDSHEET: u16 = 0x0085;
    const RECORD_NAME: u16 = 0x0018;
    const RECORD_CONTINUE: u16 = 0x003C;
    const NAME_FLAG_BUILTIN: u16 = 0x0020;

    const RECORD_DIMENSIONS: u16 = 0x0200;
    const RECORD_WINDOW2: u16 = 0x023E;
    const RECORD_NUMBER: u16 = 0x0203;

    const BOF_VERSION_BIFF8: u16 = 0x0600;
    const BOF_DT_WORKBOOK_GLOBALS: u16 = 0x0005;
    const BOF_DT_WORKSHEET: u16 = 0x0010;

    const XF_FLAG_LOCKED: u16 = 0x0001;
    const XF_FLAG_STYLE: u16 = 0x0004;

    #[test]
    fn catch_calamine_panic_converts_panic_to_import_error() {
        let err = catch_calamine_panic(|| panic!("boom"))
            .expect_err("expected panic to be converted to ImportError");
        let ImportError::CalaminePanic(message) = err else {
            panic!("unexpected error variant: {err:?}");
        };
        assert!(
            message.contains("boom"),
            "expected panic payload in message, got: {message}"
        );
    }

    #[test]
    fn catch_calamine_panic_with_context_prefixes_message() {
        let err = catch_calamine_panic_with_context("some context", || panic!("boom"))
            .expect_err("expected panic to be converted to ImportError");
        let ImportError::CalaminePanic(message) = err else {
            panic!("unexpected error variant: {err:?}");
        };
        assert!(
            message.contains("some context"),
            "expected context prefix in message, got: {message}"
        );
        assert!(
            message.contains("boom"),
            "expected panic payload in message, got: {message}"
        );
    }

    #[test]
    fn biff_stream_read_panic_is_best_effort_warning() {
        // Use an on-disk `.xls` fixture and force the BIFF stream reader to panic. The importer
        // should treat that as non-fatal (warn + fall back to calamine's direct path) rather than
        // aborting the process.
        let bytes: &[u8] = include_bytes!("../tests/fixtures/basic.xls");

        let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
        tmp.write_all(bytes).expect("write xls bytes");

        let result = import_xls_with_biff_reader(tmp.path(), None, None, |_| {
            panic!("boom from biff reader");
        })
        .expect("expected import to succeed after BIFF reader panic");

        assert!(
            result.warnings.iter().any(|w| w
                .message
                .contains("panic while reading `.xls` workbook stream")),
            "expected warning about BIFF workbook stream panic, got: {:?}",
            result.warnings
        );
    }

    fn push_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
        out.extend_from_slice(&id.to_le_bytes());
        out.extend_from_slice(&(data.len() as u16).to_le_bytes());
        out.extend_from_slice(data);
    }

    fn bof(dt: u16) -> [u8; 16] {
        let mut out = [0u8; 16];
        out[0..2].copy_from_slice(&BOF_VERSION_BIFF8.to_le_bytes());
        out[2..4].copy_from_slice(&dt.to_le_bytes());
        out[4..6].copy_from_slice(&0x0DBBu16.to_le_bytes()); // build
        out[6..8].copy_from_slice(&0x07CCu16.to_le_bytes()); // year (1996)
        out
    }

    fn window1() -> [u8; 18] {
        let mut out = [0u8; 18];
        out[14..16].copy_from_slice(&1u16.to_le_bytes()); // cTabSel
        out[16..18].copy_from_slice(&600u16.to_le_bytes()); // wTabRatio
        out
    }

    fn window2() -> [u8; 18] {
        let mut out = [0u8; 18];
        let grbit: u16 = 0x02B6;
        out[0..2].copy_from_slice(&grbit.to_le_bytes());
        out
    }

    fn write_short_unicode_string(out: &mut Vec<u8>, s: &str) {
        // BIFF8 ShortXLUnicodeString: [cch: u8][flags: u8][chars]
        let bytes = s.as_bytes();
        let len: u8 = bytes
            .len()
            .try_into()
            .expect("string too long for u8 length");
        out.push(len);
        out.push(0); // compressed
        out.extend_from_slice(bytes);
    }

    fn font_arial() -> Vec<u8> {
        // Minimal BIFF8 FONT record payload.
        const COLOR_AUTOMATIC: u16 = 0x7FFF;
        let mut out = Vec::<u8>::new();
        out.extend_from_slice(&200u16.to_le_bytes()); // height
        out.extend_from_slice(&0u16.to_le_bytes()); // option flags
        out.extend_from_slice(&COLOR_AUTOMATIC.to_le_bytes()); // color
        out.extend_from_slice(&400u16.to_le_bytes()); // weight
        out.extend_from_slice(&0u16.to_le_bytes()); // escapement
        out.push(0); // underline
        out.push(0); // family
        out.push(0); // charset
        out.push(0); // reserved
        write_short_unicode_string(&mut out, "Arial");
        out
    }

    fn xf_record(is_style_xf: bool) -> [u8; 20] {
        let mut out = [0u8; 20];
        out[0..2].copy_from_slice(&0u16.to_le_bytes()); // font
        out[2..4].copy_from_slice(&0u16.to_le_bytes()); // fmt
        let flags: u16 = XF_FLAG_LOCKED | if is_style_xf { XF_FLAG_STYLE } else { 0 };
        out[4..6].copy_from_slice(&flags.to_le_bytes());
        out
    }

    fn number_cell(row: u16, col: u16, xf: u16, v: f64) -> [u8; 14] {
        let mut out = [0u8; 14];
        out[0..2].copy_from_slice(&row.to_le_bytes());
        out[2..4].copy_from_slice(&col.to_le_bytes());
        out[4..6].copy_from_slice(&xf.to_le_bytes());
        out[6..14].copy_from_slice(&v.to_le_bytes());
        out
    }

    fn build_minimal_workbook_stream_with_corrupt_name_oob_cch() -> Vec<u8> {
        // Build a minimal, calamine-parseable BIFF8 workbook stream that contains a malformed NAME
        // record where `cch` claims more bytes than exist in the physical record payload.
        //
        // Historically, such malformed records can panic calamine due to unchecked slice indexing.
        let mut globals = Vec::<u8>::new();

        push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
        push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
        push_record(&mut globals, RECORD_WINDOW1, &window1());
        push_record(&mut globals, RECORD_FONT, &font_arial());

        for _ in 0..16 {
            push_record(&mut globals, RECORD_XF, &xf_record(true));
        }
        let xf_general = 16u16;
        push_record(&mut globals, RECORD_XF, &xf_record(false));

        // Single worksheet.
        let boundsheet_start = globals.len();
        let mut boundsheet = Vec::<u8>::new();
        boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
        boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
        write_short_unicode_string(&mut boundsheet, "Sheet1");
        push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
        let boundsheet_offset_pos = boundsheet_start + 4;

        // Malformed NAME record:
        // - cch=200, but we only provide flags byte + one character byte.
        let mut name_payload = Vec::<u8>::new();
        name_payload.extend_from_slice(&0u16.to_le_bytes()); // grbit
        name_payload.push(0); // chKey
        name_payload.push(200); // cch (declared)
        name_payload.extend_from_slice(&0u16.to_le_bytes()); // cce
        name_payload.extend_from_slice(&0u16.to_le_bytes()); // ixals
        name_payload.extend_from_slice(&0u16.to_le_bytes()); // itab
        name_payload.extend_from_slice(&[0, 0, 0, 0]); // cchCustMenu, cchDescription, cchHelpTopic, cchStatusText
        name_payload.push(0); // name string flags (compressed)
        name_payload.push(b'A'); // only 1 byte of name data
        push_record(&mut globals, RECORD_NAME, &name_payload);

        push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

        // -- Sheet substream ----------------------------------------------------
        let sheet_offset = globals.len();
        let mut sheet = Vec::<u8>::new();
        push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

        // DIMENSIONS: rows [0, 1) cols [0, 1)
        let mut dims = Vec::<u8>::new();
        dims.extend_from_slice(&0u32.to_le_bytes());
        dims.extend_from_slice(&1u32.to_le_bytes());
        dims.extend_from_slice(&0u16.to_le_bytes());
        dims.extend_from_slice(&1u16.to_le_bytes());
        dims.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut sheet, RECORD_DIMENSIONS, &dims);
        push_record(&mut sheet, RECORD_WINDOW2, &window2());
        push_record(
            &mut sheet,
            RECORD_NUMBER,
            &number_cell(0, 0, xf_general, 0.0),
        );
        push_record(&mut sheet, RECORD_EOF, &[]);

        // Patch BoundSheet offset.
        globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
            .copy_from_slice(&(sheet_offset as u32).to_le_bytes());
        globals.extend_from_slice(&sheet);
        globals
    }

    fn build_minimal_workbook_stream_with_corrupt_name_oob_cce() -> Vec<u8> {
        // Build a minimal BIFF8 workbook stream containing a malformed NAME record whose `cce`
        // (formula byte count) claims more bytes than exist in the physical record payload.
        //
        // Calamine has historically panicked on such records due to unchecked slicing when it tries
        // to locate/parse the `rgce` token stream.
        let mut globals = Vec::<u8>::new();

        push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
        push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
        push_record(&mut globals, RECORD_WINDOW1, &window1());
        push_record(&mut globals, RECORD_FONT, &font_arial());

        for _ in 0..16 {
            push_record(&mut globals, RECORD_XF, &xf_record(true));
        }
        let xf_general = 16u16;
        push_record(&mut globals, RECORD_XF, &xf_record(false));

        // Single worksheet.
        let boundsheet_start = globals.len();
        let mut boundsheet = Vec::<u8>::new();
        boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
        boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
        write_short_unicode_string(&mut boundsheet, "Sheet1");
        push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
        let boundsheet_offset_pos = boundsheet_start + 4;

        // Malformed NAME record:
        // - cch=1, provide a valid one-byte name string ("A")
        // - cce=0xFFFF, but omit any rgce bytes.
        let mut name_payload = Vec::<u8>::new();
        name_payload.extend_from_slice(&0u16.to_le_bytes()); // grbit
        name_payload.push(0); // chKey
        name_payload.push(1); // cch (declared)
        name_payload.extend_from_slice(&0xFFFFu16.to_le_bytes()); // cce (declared)
        name_payload.extend_from_slice(&0u16.to_le_bytes()); // ixals
        name_payload.extend_from_slice(&0u16.to_le_bytes()); // itab
        name_payload.extend_from_slice(&[0, 0, 0, 0]); // cchCustMenu, cchDescription, cchHelpTopic, cchStatusText
        name_payload.push(0); // name string flags (compressed)
        name_payload.push(b'A'); // 1 byte of name data
        push_record(&mut globals, RECORD_NAME, &name_payload);

        push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

        // -- Sheet substream ----------------------------------------------------
        let sheet_offset = globals.len();
        let mut sheet = Vec::<u8>::new();
        push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

        // DIMENSIONS: rows [0, 1) cols [0, 1)
        let mut dims = Vec::<u8>::new();
        dims.extend_from_slice(&0u32.to_le_bytes());
        dims.extend_from_slice(&1u32.to_le_bytes());
        dims.extend_from_slice(&0u16.to_le_bytes());
        dims.extend_from_slice(&1u16.to_le_bytes());
        dims.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut sheet, RECORD_DIMENSIONS, &dims);
        push_record(&mut sheet, RECORD_WINDOW2, &window2());
        push_record(
            &mut sheet,
            RECORD_NUMBER,
            &number_cell(0, 0, xf_general, 0.0),
        );
        push_record(&mut sheet, RECORD_EOF, &[]);

        // Patch BoundSheet offset.
        globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
            .copy_from_slice(&(sheet_offset as u32).to_le_bytes());
        globals.extend_from_slice(&sheet);
        globals
    }

    fn build_minimal_workbook_stream_with_continued_builtin_name_header_only() -> Vec<u8> {
        // Build a minimal BIFF8 workbook stream containing a built-in NAME record whose `rgchName`
        // payload (the built-in id byte) is stored entirely in a `CONTINUE` record.
        //
        // Calamine does not always handle continued NAME records safely; the sanitizer should clamp
        // `cch` so calamine doesn't attempt to read bytes that don't exist in the first fragment.
        let mut globals = Vec::<u8>::new();

        push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
        push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
        push_record(&mut globals, RECORD_WINDOW1, &window1());
        push_record(&mut globals, RECORD_FONT, &font_arial());

        for _ in 0..16 {
            push_record(&mut globals, RECORD_XF, &xf_record(true));
        }
        let xf_general = 16u16;
        push_record(&mut globals, RECORD_XF, &xf_record(false));

        // Single worksheet.
        let boundsheet_start = globals.len();
        let mut boundsheet = Vec::<u8>::new();
        boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
        boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
        write_short_unicode_string(&mut boundsheet, "Sheet1");
        push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
        let boundsheet_offset_pos = boundsheet_start + 4;

        // Built-in NAME record header only (14 bytes); `cch=1` but `rgchName` byte lives in the
        // `CONTINUE` record that follows.
        let mut name_header = Vec::<u8>::new();
        name_header.extend_from_slice(&NAME_FLAG_BUILTIN.to_le_bytes()); // grbit (builtin)
        name_header.push(0); // chKey
        name_header.push(1); // cch (built-in id length)
        name_header.extend_from_slice(&0u16.to_le_bytes()); // cce (no formula)
        name_header.extend_from_slice(&0u16.to_le_bytes()); // ixals
        name_header.extend_from_slice(&0u16.to_le_bytes()); // itab (workbook scope)
        name_header.extend_from_slice(&[0, 0, 0, 0]); // cchCustMenu, cchDescription, cchHelpTopic, cchStatusText
        push_record(&mut globals, RECORD_NAME, &name_header);
        // CONTINUE payload contains the built-in id byte (e.g. Print_Area = 0x06).
        push_record(&mut globals, RECORD_CONTINUE, &[0x06u8]);

        push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

        // -- Sheet substream ----------------------------------------------------
        let sheet_offset = globals.len();
        let mut sheet = Vec::<u8>::new();
        push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

        // DIMENSIONS: rows [0, 1) cols [0, 1)
        let mut dims = Vec::<u8>::new();
        dims.extend_from_slice(&0u32.to_le_bytes());
        dims.extend_from_slice(&1u32.to_le_bytes());
        dims.extend_from_slice(&0u16.to_le_bytes());
        dims.extend_from_slice(&1u16.to_le_bytes());
        dims.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut sheet, RECORD_DIMENSIONS, &dims);
        push_record(&mut sheet, RECORD_WINDOW2, &window2());
        push_record(
            &mut sheet,
            RECORD_NUMBER,
            &number_cell(0, 0, xf_general, 0.0),
        );
        push_record(&mut sheet, RECORD_EOF, &[]);

        // Patch BoundSheet offset.
        globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
            .copy_from_slice(&(sheet_offset as u32).to_le_bytes());
        globals.extend_from_slice(&sheet);
        globals
    }

    #[test]
    fn strips_embedded_nuls_from_string_cell_values() {
        let data = Data::String("Hello\0World".to_string());
        let (value, style) = convert_value(&data, None).expect("expected value");
        assert_eq!(value, CellValue::String("HelloWorld".to_string()));
        assert_eq!(style, None);
    }

    fn build_minimal_workbook_stream_with_continued_user_defined_name_string() -> Vec<u8> {
        // Build a minimal BIFF8 workbook stream containing a user-defined NAME record whose name
        // string is split across a CONTINUE record.
        let mut globals = Vec::<u8>::new();

        push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
        push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
        push_record(&mut globals, RECORD_WINDOW1, &window1());
        push_record(&mut globals, RECORD_FONT, &font_arial());

        for _ in 0..16 {
            push_record(&mut globals, RECORD_XF, &xf_record(true));
        }
        let xf_general = 16u16;
        push_record(&mut globals, RECORD_XF, &xf_record(false));

        // Single worksheet.
        let boundsheet_start = globals.len();
        let mut boundsheet = Vec::<u8>::new();
        boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
        boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
        write_short_unicode_string(&mut boundsheet, "Sheet1");
        push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
        let boundsheet_offset_pos = boundsheet_start + 4;

        let name = "ABCDE";
        let rgce: [u8; 3] = [0x1E, 0x2A, 0x00]; // PtgInt 42

        // NAME record header (14 bytes).
        let mut header = Vec::<u8>::new();
        header.extend_from_slice(&0u16.to_le_bytes()); // grbit
        header.push(0); // chKey
        header.push(name.len() as u8); // cch
        header.extend_from_slice(&(rgce.len() as u16).to_le_bytes()); // cce
        header.extend_from_slice(&0u16.to_le_bytes()); // ixals
        header.extend_from_slice(&0u16.to_le_bytes()); // itab (workbook scope)
        header.extend_from_slice(&[0, 0, 0, 0]); // cchCustMenu, cchDescription, cchHelpTopic, cchStatusText

        // First fragment: header + flags + "AB"
        let mut first = Vec::<u8>::new();
        first.extend_from_slice(&header);
        first.push(0); // string flags (compressed)
        first.extend_from_slice(&name.as_bytes()[..2]); // "AB"
        push_record(&mut globals, RECORD_NAME, &first);

        // Second fragment: continued segment flags + "CDE" + rgce.
        let mut second = Vec::<u8>::new();
        second.push(0); // continued segment option flags (fHighByte=0)
        second.extend_from_slice(&name.as_bytes()[2..]); // "CDE"
        second.extend_from_slice(&rgce);
        push_record(&mut globals, RECORD_CONTINUE, &second);

        push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

        // -- Sheet substream ----------------------------------------------------
        let sheet_offset = globals.len();
        let mut sheet = Vec::<u8>::new();
        push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

        // DIMENSIONS: rows [0, 1) cols [0, 1)
        let mut dims = Vec::<u8>::new();
        dims.extend_from_slice(&0u32.to_le_bytes());
        dims.extend_from_slice(&1u32.to_le_bytes());
        dims.extend_from_slice(&0u16.to_le_bytes());
        dims.extend_from_slice(&1u16.to_le_bytes());
        dims.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut sheet, RECORD_DIMENSIONS, &dims);
        push_record(&mut sheet, RECORD_WINDOW2, &window2());
        push_record(
            &mut sheet,
            RECORD_NUMBER,
            &number_cell(0, 0, xf_general, 0.0),
        );
        push_record(&mut sheet, RECORD_EOF, &[]);

        // Patch BoundSheet offset.
        globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
            .copy_from_slice(&(sheet_offset as u32).to_le_bytes());
        globals.extend_from_slice(&sheet);
        globals
    }

    fn calamine_can_open_workbook_stream(workbook_stream: &[u8]) -> bool {
        std::panic::catch_unwind(|| {
            let xls_bytes = build_in_memory_xls(workbook_stream).expect("cfb");
            let workbook: Xls<_> = Xls::new(Cursor::new(xls_bytes)).expect("open xls");
            // Force defined-name parsing.
            let _ = workbook.defined_names();
        })
        .is_ok()
    }

    fn calamine_defined_name_names(workbook_stream: &[u8]) -> Option<Vec<String>> {
        std::panic::catch_unwind(|| {
            let xls_bytes = build_in_memory_xls(workbook_stream).expect("cfb");
            let workbook: Xls<_> = Xls::new(Cursor::new(xls_bytes)).expect("open xls");
            workbook
                .defined_names()
                .iter()
                .map(|(name, _)| name.clone())
                .collect::<Vec<_>>()
        })
        .ok()
    }

    fn first_record_payload<'a>(stream: &'a [u8], record_id: u16) -> Option<&'a [u8]> {
        let mut offset = 0usize;
        while offset + 4 <= stream.len() {
            let id = u16::from_le_bytes([stream[offset], stream[offset + 1]]);
            let len = u16::from_le_bytes([stream[offset + 2], stream[offset + 3]]) as usize;
            let data_start = offset.checked_add(4)?;
            let next = data_start.checked_add(len)?;
            if next > stream.len() {
                return None;
            }
            if id == record_id {
                return Some(&stream[data_start..next]);
            }
            offset = next;
        }
        None
    }

    fn first_record_header_offset(stream: &[u8], record_id: u16) -> Option<usize> {
        let mut offset = 0usize;
        while offset + 4 <= stream.len() {
            let id = u16::from_le_bytes([stream[offset], stream[offset + 1]]);
            let len = u16::from_le_bytes([stream[offset + 2], stream[offset + 3]]) as usize;
            let data_start = offset.checked_add(4)?;
            let next = data_start.checked_add(len)?;
            if next > stream.len() {
                return None;
            }
            if id == record_id {
                return Some(offset);
            }
            offset = next;
        }
        None
    }

    #[test]
    fn sanitizes_malformed_name_record_out_of_bounds_cch_for_calamine() {
        let stream = build_minimal_workbook_stream_with_corrupt_name_oob_cch();

        // Calamine has historically panicked on malformed NAME records due to unchecked slice
        // indexing. If a newer calamine version becomes resilient to this input, this test should
        // still pass; the sanitizer exists for compatibility and defense-in-depth.
        let _ = calamine_can_open_workbook_stream(&stream);

        let payload =
            first_record_payload(&stream, RECORD_NAME).expect("expected NAME record in fixture");
        assert_eq!(payload[3], 200, "expected corrupt cch=200");

        let sanitized =
            sanitize_biff8_continued_name_records_for_calamine(&stream).expect("expected sanitize");

        let sanitized_payload = first_record_payload(&sanitized, RECORD_NAME)
            .expect("expected NAME record in sanitized workbook stream");
        // Sanitizer should clamp cch down so the name string fits in the physical record.
        assert_eq!(sanitized_payload[3], 1, "expected clamped cch=1");
        // And should keep cce patched to 0 so calamine does not attempt to parse a potentially
        // continued/invalid rgce stream.
        assert_eq!(
            u16::from_le_bytes([sanitized_payload[4], sanitized_payload[5]]),
            0,
            "expected patched cce=0"
        );

        assert!(
            calamine_can_open_workbook_stream(&sanitized),
            "expected calamine to open after sanitizing malformed NAME record"
        );
    }

    #[test]
    fn sanitizes_malformed_name_record_out_of_bounds_cce_for_calamine() {
        let stream = build_minimal_workbook_stream_with_corrupt_name_oob_cce();

        // Calamine has historically panicked on malformed NAME records due to unchecked slice
        // indexing. If a newer calamine version becomes resilient to this input, this test should
        // still pass; the sanitizer exists for compatibility and defense-in-depth.
        let _ = calamine_can_open_workbook_stream(&stream);

        let payload =
            first_record_payload(&stream, RECORD_NAME).expect("expected NAME record in fixture");
        assert_eq!(
            u16::from_le_bytes([payload[4], payload[5]]),
            0xFFFF,
            "expected corrupt cce=0xFFFF"
        );

        let sanitized =
            sanitize_biff8_continued_name_records_for_calamine(&stream).expect("expected sanitize");

        let sanitized_payload = first_record_payload(&sanitized, RECORD_NAME)
            .expect("expected NAME record in sanitized workbook stream");
        // Sanitizer should patch cce to 0 so calamine does not attempt to slice/parse rgce.
        assert_eq!(
            u16::from_le_bytes([sanitized_payload[4], sanitized_payload[5]]),
            0,
            "expected patched cce=0"
        );

        assert!(
            calamine_can_open_workbook_stream(&sanitized),
            "expected calamine to open after sanitizing malformed NAME record"
        );
    }

    #[test]
    fn sanitizes_continued_builtin_name_when_id_bytes_are_missing_in_first_fragment() {
        let stream = build_minimal_workbook_stream_with_continued_builtin_name_header_only();

        // Even if calamine becomes resilient to this input in the future, the sanitizer should
        // still clamp cch/cce defensively so workbook import never panics.
        let _ = calamine_can_open_workbook_stream(&stream);

        let payload =
            first_record_payload(&stream, RECORD_NAME).expect("expected NAME record in fixture");
        assert_eq!(payload[0..2], NAME_FLAG_BUILTIN.to_le_bytes());
        assert_eq!(payload[3], 1, "expected built-in cch=1");

        let sanitized =
            sanitize_biff8_continued_name_records_for_calamine(&stream).expect("expected sanitize");

        let sanitized_payload = first_record_payload(&sanitized, RECORD_NAME)
            .expect("expected NAME record in sanitized workbook stream");
        // NAME record contains only the 14-byte header in its first fragment; sanitizer should
        // coalesce/compact the following CONTINUE record so the built-in id byte is available and
        // calamine can parse the name without panicking.
        assert_eq!(sanitized_payload[3], 1, "expected cch preserved");
        assert_eq!(sanitized_payload[14], 0x06, "expected built-in id byte");
        assert_eq!(
            u16::from_le_bytes([sanitized_payload[4], sanitized_payload[5]]),
            0,
            "expected patched cce=0"
        );

        assert!(
            calamine_can_open_workbook_stream(&sanitized),
            "expected calamine to open after sanitizing continued built-in NAME record"
        );
    }

    #[test]
    fn preserves_continued_user_defined_name_string_for_calamine() {
        let stream = build_minimal_workbook_stream_with_continued_user_defined_name_string();

        // If calamine becomes resilient to this input in the future, this should still pass.
        let _ = calamine_can_open_workbook_stream(&stream);

        let sanitized =
            sanitize_biff8_continued_name_records_for_calamine(&stream).expect("expected sanitize");

        let payload = first_record_payload(&sanitized, RECORD_NAME)
            .expect("expected NAME record in sanitized workbook stream");
        assert_eq!(payload[3], 5, "expected cch preserved");
        assert_eq!(payload[14], 0, "expected compressed string flags");
        assert_eq!(&payload[15..20], b"ABCDE", "expected compacted name bytes");

        let names = calamine_defined_name_names(&sanitized).expect("expected calamine open");
        assert!(
            names
                .iter()
                .any(|n| normalize_calamine_defined_name_name(n) == "ABCDE"),
            "expected calamine to surface full name string; names={names:?}"
        );
    }

    fn build_minimal_workbook_stream_with_truncated_name_record(payload_len: usize) -> Vec<u8> {
        // Build a minimal BIFF8 workbook stream containing a NAME record whose physical record
        // payload is too short for calamine's `&r.data[14..]` indexing.
        let mut globals = Vec::<u8>::new();

        push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
        push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
        push_record(&mut globals, RECORD_WINDOW1, &window1());
        push_record(&mut globals, RECORD_FONT, &font_arial());

        for _ in 0..16 {
            push_record(&mut globals, RECORD_XF, &xf_record(true));
        }
        let xf_general = 16u16;
        push_record(&mut globals, RECORD_XF, &xf_record(false));

        // Single worksheet.
        let boundsheet_start = globals.len();
        let mut boundsheet = Vec::<u8>::new();
        boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
        boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
        write_short_unicode_string(&mut boundsheet, "Sheet1");
        push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
        let boundsheet_offset_pos = boundsheet_start + 4;

        // Truncated NAME record payload (all zeros is fine; the sanitizer will mask the record id
        // so calamine ignores it).
        push_record(&mut globals, RECORD_NAME, &vec![0u8; payload_len]);

        push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

        // -- Sheet substream ----------------------------------------------------
        let sheet_offset = globals.len();
        let mut sheet = Vec::<u8>::new();
        push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

        // DIMENSIONS: rows [0, 1) cols [0, 1)
        let mut dims = Vec::<u8>::new();
        dims.extend_from_slice(&0u32.to_le_bytes());
        dims.extend_from_slice(&1u32.to_le_bytes());
        dims.extend_from_slice(&0u16.to_le_bytes());
        dims.extend_from_slice(&1u16.to_le_bytes());
        dims.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut sheet, RECORD_DIMENSIONS, &dims);
        push_record(&mut sheet, RECORD_WINDOW2, &window2());
        push_record(
            &mut sheet,
            RECORD_NUMBER,
            &number_cell(0, 0, xf_general, 0.0),
        );
        push_record(&mut sheet, RECORD_EOF, &[]);

        // Patch BoundSheet offset.
        globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
            .copy_from_slice(&(sheet_offset as u32).to_le_bytes());
        globals.extend_from_slice(&sheet);
        globals
    }

    fn build_minimal_workbook_stream_with_truncated_name_record_and_continue() -> Vec<u8> {
        // Build a minimal BIFF8 workbook stream containing a truncated NAME record whose header
        // bytes are completed by a following CONTINUE record.
        //
        // Calamine panics on `NAME.len < 14`, but we can repair this case by coalescing the
        // following CONTINUE record into the NAME's physical record length.
        let mut globals = Vec::<u8>::new();

        push_record(&mut globals, RECORD_BOF, &bof(BOF_DT_WORKBOOK_GLOBALS));
        push_record(&mut globals, RECORD_CODEPAGE, &1252u16.to_le_bytes());
        push_record(&mut globals, RECORD_WINDOW1, &window1());
        push_record(&mut globals, RECORD_FONT, &font_arial());

        for _ in 0..16 {
            push_record(&mut globals, RECORD_XF, &xf_record(true));
        }
        let xf_general = 16u16;
        push_record(&mut globals, RECORD_XF, &xf_record(false));

        // Single worksheet.
        let boundsheet_start = globals.len();
        let mut boundsheet = Vec::<u8>::new();
        boundsheet.extend_from_slice(&0u32.to_le_bytes()); // placeholder lbPlyPos
        boundsheet.extend_from_slice(&0u16.to_le_bytes()); // visible worksheet
        write_short_unicode_string(&mut boundsheet, "Sheet1");
        push_record(&mut globals, RECORD_BOUNDSHEET, &boundsheet);
        let boundsheet_offset_pos = boundsheet_start + 4;

        // Truncated NAME record payload (10 bytes). Set a non-zero cch/cce so the sanitizer has
        // something to patch.
        let mut name_payload = vec![0u8; 10];
        name_payload[3] = 1; // cch
        name_payload[4..6].copy_from_slice(&7u16.to_le_bytes()); // cce (bogus)
        push_record(&mut globals, RECORD_NAME, &name_payload);

        // CONTINUE payload supplies the bytes calamine expects at `NAME.payload[14..]`. Use a
        // minimal XLUnicodeStringNoCch encoding for the name: flags=0x00 (compressed) + "A".
        push_record(&mut globals, RECORD_CONTINUE, &[0x00, b'A']);

        push_record(&mut globals, RECORD_EOF, &[]); // EOF globals

        // -- Sheet substream ----------------------------------------------------
        let sheet_offset = globals.len();
        let mut sheet = Vec::<u8>::new();
        push_record(&mut sheet, RECORD_BOF, &bof(BOF_DT_WORKSHEET));

        // DIMENSIONS: rows [0, 1) cols [0, 1)
        let mut dims = Vec::<u8>::new();
        dims.extend_from_slice(&0u32.to_le_bytes());
        dims.extend_from_slice(&1u32.to_le_bytes());
        dims.extend_from_slice(&0u16.to_le_bytes());
        dims.extend_from_slice(&1u16.to_le_bytes());
        dims.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut sheet, RECORD_DIMENSIONS, &dims);
        push_record(&mut sheet, RECORD_WINDOW2, &window2());
        push_record(
            &mut sheet,
            RECORD_NUMBER,
            &number_cell(0, 0, xf_general, 0.0),
        );
        push_record(&mut sheet, RECORD_EOF, &[]);

        // Patch BoundSheet offset.
        globals[boundsheet_offset_pos..boundsheet_offset_pos + 4]
            .copy_from_slice(&(sheet_offset as u32).to_le_bytes());
        globals.extend_from_slice(&sheet);
        globals
    }

    #[test]
    fn masks_truncated_name_records_so_calamine_does_not_panic() {
        let stream = build_minimal_workbook_stream_with_truncated_name_record(10);

        // If calamine fixes the underlying panic in the future, this should still pass.
        let _ = calamine_can_open_workbook_stream(&stream);

        let name_header_offset =
            first_record_header_offset(&stream, RECORD_NAME).expect("expected NAME record");
        assert_eq!(
            u16::from_le_bytes([stream[name_header_offset], stream[name_header_offset + 1]]),
            RECORD_NAME
        );
        assert_eq!(
            u16::from_le_bytes([
                stream[name_header_offset + 2],
                stream[name_header_offset + 3]
            ]),
            10
        );

        let sanitized =
            sanitize_biff8_continued_name_records_for_calamine(&stream).expect("expected sanitize");

        // Sanitizer should mask the record id so calamine ignores it.
        assert_eq!(
            u16::from_le_bytes([
                sanitized[name_header_offset],
                sanitized[name_header_offset + 1]
            ]),
            0xFFFF
        );

        assert!(
            calamine_can_open_workbook_stream(&sanitized),
            "expected calamine to open after masking truncated NAME record"
        );
    }

    #[test]
    fn repairs_truncated_name_records_when_followed_by_continue() {
        let stream = build_minimal_workbook_stream_with_truncated_name_record_and_continue();

        // If calamine fixes the underlying panic in the future, this should still pass.
        let _ = calamine_can_open_workbook_stream(&stream);

        let name_header_offset =
            first_record_header_offset(&stream, RECORD_NAME).expect("expected NAME record");
        assert_eq!(
            u16::from_le_bytes([stream[name_header_offset], stream[name_header_offset + 1]]),
            RECORD_NAME
        );
        assert_eq!(
            u16::from_le_bytes([
                stream[name_header_offset + 2],
                stream[name_header_offset + 3]
            ]),
            10
        );

        let sanitized =
            sanitize_biff8_continued_name_records_for_calamine(&stream).expect("expected sanitize");

        // Sanitizer should *not* mask the record id; it can be repaired by coalescing the
        // following CONTINUE record into the physical record length.
        assert_eq!(
            u16::from_le_bytes([
                sanitized[name_header_offset],
                sanitized[name_header_offset + 1]
            ]),
            RECORD_NAME
        );
        assert_eq!(
            u16::from_le_bytes([
                sanitized[name_header_offset + 2],
                sanitized[name_header_offset + 3]
            ]),
            16,
            "expected NAME record length to be expanded to include CONTINUE record"
        );

        let sanitized_payload = first_record_payload(&sanitized, RECORD_NAME)
            .expect("expected NAME record in sanitized workbook stream");
        assert_eq!(sanitized_payload[3], 1, "expected cch preserved");
        assert_eq!(
            u16::from_le_bytes([sanitized_payload[4], sanitized_payload[5]]),
            0,
            "expected patched cce=0"
        );

        assert!(
            calamine_can_open_workbook_stream(&sanitized),
            "expected calamine to open after repairing truncated NAME record"
        );
    }

    #[test]
    fn parse_autofilter_range_strips_wrapping_parentheses() {
        let range = parse_autofilter_range_from_defined_name("=($A$1:$B$3)")
            .expect("expected parenthesized range to parse");
        assert_eq!(range, Range::from_a1("A1:B3").unwrap());

        let range = parse_autofilter_range_from_defined_name("=(Sheet1!$A$1:$B$3)")
            .expect("expected parenthesized sheet-qualified range to parse");
        assert_eq!(range, Range::from_a1("A1:B3").unwrap());
    }

    #[test]
    fn parse_autofilter_range_rejects_union_even_when_wrapped_in_parentheses() {
        let err = parse_autofilter_range_from_defined_name("=(Sheet1!$A$1:$A$2,Sheet1!$C$1:$C$2)")
            .expect_err("expected union formula to be rejected");
        assert!(err.contains("union"), "expected union error, got {err:?}");
    }

    #[test]
    fn parse_autofilter_range_rejects_union_when_sheet_name_contains_double_quotes() {
        // Sheet names can be quoted with `'` and may still contain `"`; ensure we don't treat
        // `"` as starting a string literal inside a quoted sheet name.
        let err = parse_autofilter_range_from_defined_name(
            r#"=('My"Sheet'!$A$1:$A$2,'My"Sheet'!$C$1:$C$2)"#,
        )
        .expect_err("expected union formula to be rejected");
        assert!(err.contains("union"), "expected union error, got {err:?}");
    }

    #[test]
    fn parse_autofilter_range_strips_wrapping_parentheses_around_implicit_intersection_prefix() {
        // Calamine (and some BIFF formula decoders) can render implicit intersection (`@`) with
        // redundant parentheses.
        let range = parse_autofilter_range_from_defined_name("=(@$A$1:$B$3)")
            .expect("expected parenthesized implicit intersection range to parse");
        assert_eq!(range, Range::from_a1("A1:B3").unwrap());

        let range = parse_autofilter_range_from_defined_name("=(@(Sheet1!$A$1:$B$3))")
            .expect("expected parenthesized sheet-qualified implicit intersection range to parse");
        assert_eq!(range, Range::from_a1("A1:B3").unwrap());
    }

    #[test]
    fn infer_sheet_scope_strips_wrapping_parentheses() {
        let mut workbook = Workbook::new();
        workbook.add_sheet("Sheet1").unwrap();
        workbook.add_sheet("Sheet2").unwrap();

        let mut warnings = ImportWarnings::new();
        let mut suppressed = false;
        let sheet = infer_sheet_name_from_workbook_scoped_defined_name(
            &workbook,
            XLNM_FILTER_DATABASE,
            "=(Sheet1!$A$1:$B$3)",
            &mut warnings,
            &mut suppressed,
        );
        assert!(matches!(
            sheet,
            WorkbookScopedDefinedNameSheetInference::Inferred(ref sheet_name)
                if sheet_name == "Sheet1"
        ));
        assert_eq!(warnings.count(), 0, "unexpected warnings attempted");
        let stored = warnings.into_vec();
        assert!(stored.is_empty(), "warnings={stored:?}");
    }

    #[test]
    fn import_warnings_caps_total_and_emits_sentinel() {
        let mut warnings = ImportWarnings::new();
        let attempted = MAX_IMPORT_WARNINGS_TOTAL + 50;
        warnings.extend((0..attempted).map(|i| ImportWarning::new(format!("warning {i}"))));

        assert_eq!(warnings.count(), attempted);

        let stored = warnings.into_vec();
        assert!(
            stored.len() <= MAX_IMPORT_WARNINGS_TOTAL + 1,
            "stored warning count should be capped; len={}; cap={}",
            stored.len(),
            MAX_IMPORT_WARNINGS_TOTAL
        );
        let sentinel_count = stored
            .iter()
            .filter(|w| w.message == IMPORT_WARNINGS_SUPPRESSED_MESSAGE)
            .count();
        assert_eq!(
            sentinel_count, 1,
            "expected exactly one sentinel warning; stored={stored:?}"
        );
    }

    #[test]
    fn infer_sheet_scope_reports_failure_even_when_warnings_are_suppressed() {
        let mut workbook = Workbook::new();
        workbook.add_sheet("Sheet1").unwrap();
        workbook.add_sheet("Sheet2").unwrap();

        let mut warnings = ImportWarnings::new();
        let mut suppressed = true;
        let sheet = infer_sheet_name_from_workbook_scoped_defined_name(
            &workbook,
            XLNM_FILTER_DATABASE,
            "='Sheet1$A$1:$B$3",
            &mut warnings,
            &mut suppressed,
        );

        assert_eq!(sheet, WorkbookScopedDefinedNameSheetInference::Failure);
        assert!(warnings.count() > 0, "expected at least one warning to be attempted");
        let stored = warnings.into_vec();
        assert!(stored.is_empty(), "expected warnings to be suppressed, got: {stored:?}");
    }

    #[test]
    fn autofilter_workbook_scoped_sheet_inference_does_not_guess_when_inference_failed_even_if_suppressed(
    ) {
        let mut workbook = Workbook::new();
        workbook.add_sheet("Sheet1").unwrap();
        workbook.add_sheet("Sheet2").unwrap();

        // Pre-seed AutoFilter metadata on exactly one sheet, matching the importer's best-effort
        // workbook-scoped `_FilterDatabase` heuristic.
        let sheet2_id = workbook
            .sheet_by_name("Sheet2")
            .map(|s| s.id)
            .expect("Sheet2");
        let sheet2 = workbook.sheet_mut(sheet2_id).expect("Sheet2 mut");
         sheet2.auto_filter = Some(SheetAutoFilter {
             range: Range::from_a1("A1:B3").unwrap(),
             filter_columns: Vec::new(),
             sort_state: None,
             raw_xml: Vec::new(),
         });

        let mut warnings = ImportWarnings::new();
        let mut suppressed = true;
        let (sheet_name, inference_failed) = infer_autofilter_sheet_name_from_workbook_scoped_defined_name(
            &workbook,
            XLNM_FILTER_DATABASE,
            "='Sheet1$A$1:$B$3",
            &mut warnings,
            &mut suppressed,
        );

        assert!(inference_failed, "expected inference to fail");
        assert!(
            sheet_name.is_none(),
            "expected no heuristic sheet guessing on inference failure, got: {sheet_name:?}"
        );
    }

    #[test]
    fn filter_database_defined_name_accepts_truncated_canonical_spelling() {
        assert!(is_filter_database_defined_name("_xlnm._FilterDatabas"));
    }
}
