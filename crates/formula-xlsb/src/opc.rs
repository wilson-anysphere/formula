use crate::biff12_varint;
use crate::parser::Error as ParseError;
use crate::parser::{
    biff12, parse_shared_strings, parse_sheet, parse_sheet_stream, parse_workbook, Cell, DefinedName,
    SheetData, SheetMeta, WorkbookProperties,
};
#[cfg(any(not(target_arch = "wasm32"), feature = "write"))]
use crate::parser::CellValue;
#[cfg(not(target_arch = "wasm32"))]
use crate::patch::{
    patch_sheet_bin, patch_sheet_bin_streaming, value_edit_is_noop_inline_string, CellEdit,
};
#[cfg(not(target_arch = "wasm32"))]
use crate::shared_strings_write::{
    reusable_plain_si_utf16_end, SharedStringsWriter, SharedStringsWriterStreaming,
};
use crate::styles::Styles;
use crate::workbook_context::WorkbookContext;
use crate::SharedString;
use formula_office_crypto as office_crypto;
use quick_xml::events::Event;
use quick_xml::Reader as XmlReader;
use quick_xml::Writer as XmlWriter;
use std::collections::{HashMap, HashSet};
use std::fmt;
use std::io::{self, Cursor, Read, Seek, SeekFrom, Write};
use std::ops::ControlFlow;
use std::path::PathBuf;
use std::sync::Arc;
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

#[cfg(not(target_arch = "wasm32"))]
use formula_fs::{atomic_write_with_path, AtomicWriteError};
#[cfg(not(target_arch = "wasm32"))]
use std::fs::File;
#[cfg(not(target_arch = "wasm32"))]
use std::collections::{BTreeMap, BTreeSet};
#[cfg(not(target_arch = "wasm32"))]
use std::path::Path;

const DEFAULT_SHARED_STRINGS_PART: &str = "xl/sharedStrings.bin";
const DEFAULT_STYLES_PART: &str = "xl/styles.bin";
const DEFAULT_WORKBOOK_PART: &str = "xl/workbook.bin";
const DEFAULT_WORKBOOK_RELS_PART: &str = "xl/_rels/workbook.bin.rels";

/// Maximum uncompressed size allowed for a single ZIP entry (OPC part) when reading XLSB files.
///
/// XLSB is a ZIP container; the per-entry `size` value is attacker-controlled. We enforce an
/// explicit bound to prevent immediate huge allocations / OOM when opening or saving.
const MAX_XLSB_ZIP_PART_BYTES: u64 = 256 * 1024 * 1024; // 256 MiB

/// Maximum total decoded bytes we will store in `preserved_parts` during workbook open.
///
/// This bounds memory usage when `preserve_unknown_parts=true` and a package contains many parts.
const MAX_XLSB_PRESERVED_TOTAL_BYTES: u64 = 512 * 1024 * 1024; // 512 MiB

/// Maximum total bytes allowed when opening an XLSB package from an in-memory buffer.
///
/// Opening from bytes (including decrypted Office-encrypted payloads) materializes the full ZIP
/// container in memory so the workbook can lazily re-open parts later. Avoid unbounded reads /
/// allocations for attacker-controlled inputs.
const MAX_XLSB_PACKAGE_BYTES: u64 = 512 * 1024 * 1024; // 512 MiB

/// Maximum number of ZIP entries we will process when opening an XLSB file.
///
/// This prevents pathological packages with millions of tiny entries from causing excessive CPU
/// time and memory overhead even when individual parts are size-limited.
const MAX_XLSB_ZIP_ENTRIES: usize = 100_000;

/// Cap initial allocation when reading a ZIP entry; do not trust `ZipFile::size()` for prealloc.
const ZIP_ENTRY_READ_PREALLOC_BYTES: usize = 64 * 1024; // 64 KiB

const ENV_MAX_XLSB_ZIP_PART_BYTES: &str = "FORMULA_XLSB_MAX_ZIP_PART_BYTES";
const ENV_MAX_XLSB_PRESERVED_TOTAL_BYTES: &str = "FORMULA_XLSB_MAX_PRESERVED_TOTAL_BYTES";
const ENV_MAX_XLSB_ZIP_ENTRIES: &str = "FORMULA_XLSB_MAX_ZIP_ENTRIES";
const ENV_MAX_XLSB_PACKAGE_BYTES: &str = "FORMULA_XLSB_MAX_PACKAGE_BYTES";

const OFFICE_DOCUMENT_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const SHARED_STRINGS_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
const STYLES_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const TABLE_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";

const SHARED_STRINGS_CONTENT_TYPE: &str = "application/vnd.ms-excel.sharedStrings";
const STYLES_CONTENT_TYPE: &str = "application/vnd.ms-excel.styles";

fn max_xlsb_zip_part_bytes() -> u64 {
    std::env::var(ENV_MAX_XLSB_ZIP_PART_BYTES)
        .ok()
        .and_then(|v| v.parse::<u64>().ok())
        .unwrap_or(MAX_XLSB_ZIP_PART_BYTES)
}

fn max_xlsb_preserved_total_bytes() -> u64 {
    std::env::var(ENV_MAX_XLSB_PRESERVED_TOTAL_BYTES)
        .ok()
        .and_then(|v| v.parse::<u64>().ok())
        .unwrap_or(MAX_XLSB_PRESERVED_TOTAL_BYTES)
}

fn max_xlsb_zip_entries() -> usize {
    std::env::var(ENV_MAX_XLSB_ZIP_ENTRIES)
        .ok()
        .and_then(|v| v.parse::<usize>().ok())
        .unwrap_or(MAX_XLSB_ZIP_ENTRIES)
}

fn max_xlsb_package_bytes() -> u64 {
    std::env::var(ENV_MAX_XLSB_PACKAGE_BYTES)
        .ok()
        .and_then(|v| v.parse::<u64>().ok())
        .unwrap_or(MAX_XLSB_PACKAGE_BYTES)
}

fn xlsb_package_too_large_error(size: u64, max: u64) -> ParseError {
    ParseError::Io(io::Error::new(
        io::ErrorKind::InvalidData,
        format!("XLSB package too large: {size} bytes exceeds limit {max} bytes"),
    ))
}

/// OLE/CFB file signature.
///
/// See: https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/
const OLE_MAGIC: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

/// Controls how much of the original package we keep around for round-trip preservation.
#[derive(Debug, Clone)]
pub struct OpenOptions {
    /// If true, read and store any ZIP entries that we do not currently parse.
    ///
    /// This enables future round-trip support by copying these parts back out unchanged.
    pub preserve_unknown_parts: bool,
    /// If true, also preserve raw bytes for parts we *do* parse (workbook and sharedStrings).
    ///
    /// This is useful for future round-tripping when the writer is still incomplete.
    /// Note that this can increase memory usage for workbooks with a very large shared string table.
    pub preserve_parsed_parts: bool,
    /// If true, also preserve the raw bytes for worksheet `.bin` parts.
    ///
    /// Worksheet parts can be very large. If you only need fast read access,
    /// leave this off and rely on re-reading the source workbook when writing.
    pub preserve_worksheets: bool,
    /// If true, decode parsed formula token streams (`rgce`/`rgcb`) into best-effort Excel formula
    /// text during parsing.
    ///
    /// This is enabled by default to preserve historical behavior. For very large XLSB files,
    /// decoding every formula can be expensive in CPU (and some allocations). Callers that only
    /// need raw `rgce`/`rgcb` bytes for evaluation or round-trip preservation can set this to
    /// `false` to skip decoding.
    ///
    /// When `false`, [`crate::Formula::text`] will always be `None` and
    /// [`crate::Formula::warnings`] will be empty.
    pub decode_formulas: bool,
}

impl Default for OpenOptions {
    fn default() -> Self {
        Self {
            preserve_unknown_parts: true,
            preserve_parsed_parts: true,
            preserve_worksheets: false,
            decode_formulas: true,
        }
    }
}

trait ReadSeek: Read + Seek {}
impl<T: Read + Seek> ReadSeek for T {}

#[cfg_attr(target_arch = "wasm32", allow(dead_code))]
#[derive(Clone)]
enum WorkbookSource {
    Path(PathBuf),
    Bytes(Arc<[u8]>),
}

impl fmt::Debug for WorkbookSource {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            WorkbookSource::Path(path) => f.debug_tuple("Path").field(path).finish(),
            WorkbookSource::Bytes(bytes) => f.debug_tuple("Bytes").field(&bytes.len()).finish(),
        }
    }
}

struct ParsedWorkbook {
    sheets: Vec<SheetMeta>,
    workbook_part: String,
    workbook_rels_part: String,
    shared_strings: Vec<String>,
    shared_strings_table: Vec<SharedString>,
    #[cfg_attr(target_arch = "wasm32", allow(dead_code))]
    shared_strings_part: Option<String>,
    workbook_context: WorkbookContext,
    workbook_properties: WorkbookProperties,
    defined_names: Vec<DefinedName>,
    styles: Styles,
    styles_part: Option<String>,
    preserved_parts: HashMap<String, Vec<u8>>,
    preserve_parsed_parts: bool,
    decode_formulas: bool,
}

/// An opened XLSB workbook.
///
/// This type keeps enough metadata to stream worksheets on demand. It also optionally stores
/// raw bytes for parts we do not understand, enabling round-trip preservation later.
#[derive(Debug)]
pub struct XlsbWorkbook {
    source: WorkbookSource,
    sheets: Vec<SheetMeta>,
    workbook_part: String,
    workbook_rels_part: String,
    shared_strings: Vec<String>,
    shared_strings_table: Vec<SharedString>,
    #[cfg_attr(target_arch = "wasm32", allow(dead_code))]
    shared_strings_part: Option<String>,
    workbook_context: WorkbookContext,
    workbook_properties: WorkbookProperties,
    defined_names: Vec<DefinedName>,
    styles: Styles,
    styles_part: Option<String>,
    preserved_parts: HashMap<String, Vec<u8>>,
    preserve_parsed_parts: bool,
    decode_formulas: bool,
}

/// A single formula edit expressed as Excel formula text.
///
/// This is a higher-level companion to [`crate::CellEdit`], intended for use with save APIs that
/// need to (re-)encode formulas.
#[cfg(feature = "write")]
#[derive(Debug, Clone, PartialEq)]
pub struct FormulaTextCellEdit {
    pub row: u32,
    pub col: u32,
    pub new_value: CellValue,
    /// Excel formula text, with or without a leading `=`.
    pub formula: String,
}

impl XlsbWorkbook {
    #[cfg(not(target_arch = "wasm32"))]
    pub fn open(path: impl AsRef<Path>) -> Result<Self, ParseError> {
        Self::open_with_options(path, OpenOptions::default())
    }

    /// Open an XLSB workbook, transparently handling Office-encrypted (OLE/CFB `EncryptedPackage`)
    /// wrappers using the provided `password`.
    ///
    /// If the input is a normal ZIP-based XLSB package, this behaves like [`Self::open`].
    #[cfg(not(target_arch = "wasm32"))]
    pub fn open_with_password(path: impl AsRef<Path>, password: &str) -> Result<Self, ParseError> {
        let path = path.as_ref().to_path_buf();
        let mut file = File::open(&path)?;

        let mut header = [0u8; 8];
        let n = file.read(&mut header)?;

        if n >= OLE_MAGIC.len() && header[..OLE_MAGIC.len()] == OLE_MAGIC {
            // Office-encrypted `.xlsb` files are stored as an OLE/CFB wrapper containing
            // `EncryptionInfo` + `EncryptedPackage`. Decrypt to raw ZIP bytes in memory.
            let max = max_xlsb_package_bytes();

            file.seek(SeekFrom::Start(0))?;
            let len = file.seek(SeekFrom::End(0))?;
            if len > max {
                return Err(xlsb_package_too_large_error(len, max));
            }
            file.seek(SeekFrom::Start(0))?;

            let mut ole_bytes = Vec::new();
            file.take(max.saturating_add(1)).read_to_end(&mut ole_bytes)?;
            if ole_bytes.len() as u64 > max {
                return Err(xlsb_package_too_large_error(ole_bytes.len() as u64, max));
            }

            let zip_bytes = office_crypto::decrypt_encrypted_package(&ole_bytes, password)
                .map_err(map_office_crypto_err)?;
            return Self::open_from_owned_bytes(zip_bytes.into(), OpenOptions::default());
        }

        Self::open_with_options(path, OpenOptions::default())
    }

    #[cfg(not(target_arch = "wasm32"))]
    pub fn open_with_options(
        path: impl AsRef<Path>,
        options: OpenOptions,
    ) -> Result<Self, ParseError> {
        let path = path.as_ref().to_path_buf();
        let mut file = File::open(&path)?;
        preflight_zip_entry_count(&mut file)?;
        file.seek(SeekFrom::Start(0))?;
        let mut zip = ZipArchive::new(file)?;
        let parsed = parse_xlsb_from_zip(&mut zip, options)?;
        Ok(Self::from_parsed(WorkbookSource::Path(path), parsed))
    }

    /// Open an XLSB workbook from in-memory ZIP bytes, without copying.
    pub fn from_bytes(bytes: Arc<[u8]>, options: OpenOptions) -> Result<Self, ParseError> {
        Self::open_from_owned_bytes(bytes, options)
    }

    /// Open an XLSB workbook from an in-memory reader.
    ///
    /// This is primarily intended for decrypted Office-encrypted XLSB files, where the decrypted
    /// OPC/ZIP package exists only in memory.
    ///
    /// `options` mirrors [`Self::open_with_options`], controlling which parts are preserved and
    /// whether formulas are decoded.
    ///
    /// Note: this implementation buffers the full ZIP container into memory so worksheet parts
    /// can be streamed on demand.
    pub fn open_from_reader<R: Read + Seek>(
        mut reader: R,
        options: OpenOptions,
    ) -> Result<Self, ParseError> {
        let max = max_xlsb_package_bytes();

        // Avoid unbounded reads/allocation for attacker-controlled input streams. We still
        // materialize the package into memory so the workbook can lazily re-open the ZIP later.
        reader.seek(SeekFrom::Start(0))?;
        let len = reader.seek(SeekFrom::End(0))?;
        if len > max {
            return Err(xlsb_package_too_large_error(len, max));
        }
        reader.seek(SeekFrom::Start(0))?;

        // Avoid buffering the entire ZIP package into memory only to discover that it contains an
        // excessive number of entries. This is especially important for attacker-controlled
        // streams where the ZIP may contain many small entries but still remain within the overall
        // package byte limit.
        preflight_zip_entry_count(&mut reader)?;
        reader.seek(SeekFrom::Start(0))?;

        let mut bytes = Vec::new();
        reader.take(max.saturating_add(1)).read_to_end(&mut bytes)?;
        if bytes.len() as u64 > max {
            return Err(xlsb_package_too_large_error(bytes.len() as u64, max));
        }
        let bytes: Arc<[u8]> = bytes.into();
        Self::open_from_owned_bytes(bytes, options)
    }

    /// Open an XLSB workbook from an in-memory reader, controlling preservation options.
    ///
    /// This is an alias for [`Self::open_from_reader`].
    pub fn open_from_reader_with_options<R: Read + Seek>(
        reader: R,
        options: OpenOptions,
    ) -> Result<Self, ParseError> {
        Self::open_from_reader(reader, options)
    }

    /// Open an XLSB workbook from an owned in-memory ZIP buffer.
    ///
    /// This avoids copying when the caller already has a `Vec<u8>` (e.g. decrypted `EncryptedPackage`
    /// bytes).
    pub fn open_from_vec(bytes: Vec<u8>) -> Result<Self, ParseError> {
        Self::open_from_vec_with_options(bytes, OpenOptions::default())
    }

    /// Open an XLSB workbook from an owned in-memory ZIP buffer, controlling preservation options.
    pub fn open_from_vec_with_options(
        bytes: Vec<u8>,
        options: OpenOptions,
    ) -> Result<Self, ParseError> {
        Self::open_from_owned_bytes(bytes.into(), options)
    }

    /// Open an XLSB workbook from in-memory ZIP bytes.
    pub fn open_from_bytes(bytes: &[u8]) -> Result<Self, ParseError> {
        Self::open_from_bytes_with_options(bytes, OpenOptions::default())
    }

    /// Open an XLSB workbook from in-memory ZIP bytes, controlling preservation options.
    pub fn open_from_bytes_with_options(
        bytes: &[u8],
        options: OpenOptions,
    ) -> Result<Self, ParseError> {
        let max = max_xlsb_package_bytes();
        if bytes.len() as u64 > max {
            return Err(xlsb_package_too_large_error(bytes.len() as u64, max));
        }
        Self::open_from_owned_bytes(Arc::from(bytes), options)
    }

    /// Open an XLSB workbook from either:
    /// - raw ZIP bytes (normal XLSB package), or
    /// - an Office-encrypted OLE/CFB container with `EncryptionInfo` + `EncryptedPackage`.
    ///
    /// This is primarily intended for callers that already have the input in memory (e.g. from an
    /// upload).
    pub fn open_from_bytes_with_password(
        bytes: &[u8],
        password: &str,
        options: OpenOptions,
    ) -> Result<Self, ParseError> {
        let max = max_xlsb_package_bytes();
        if bytes.len() as u64 > max {
            return Err(xlsb_package_too_large_error(bytes.len() as u64, max));
        }
        if bytes.starts_with(b"PK") {
            return Self::open_from_bytes_with_options(bytes, options);
        }

        if bytes.len() >= OLE_MAGIC.len() && bytes[..OLE_MAGIC.len()] == OLE_MAGIC {
            let zip_bytes = office_crypto::decrypt_encrypted_package(bytes, password)
                .map_err(map_office_crypto_err)?;
            return Self::open_from_owned_bytes(zip_bytes.into(), options);
        }

        // Fall back to the standard ZIP reader (which will surface a `ZipError` for non-ZIP input).
        Self::open_from_bytes_with_options(bytes, options)
    }

    fn open_from_owned_bytes(bytes: Arc<[u8]>, options: OpenOptions) -> Result<Self, ParseError> {
        let max = max_xlsb_package_bytes();
        if bytes.len() as u64 > max {
            return Err(xlsb_package_too_large_error(bytes.len() as u64, max));
        }
        let mut cursor = Cursor::new(bytes.clone());
        preflight_zip_entry_count(&mut cursor)?;
        cursor.seek(SeekFrom::Start(0))?;
        let mut zip = ZipArchive::new(cursor)?;
        let parsed = parse_xlsb_from_zip(&mut zip, options)?;
        Ok(Self::from_parsed(WorkbookSource::Bytes(bytes), parsed))
    }

    fn from_parsed(source: WorkbookSource, parsed: ParsedWorkbook) -> Self {
        Self {
            source,
            sheets: parsed.sheets,
            workbook_part: parsed.workbook_part,
            workbook_rels_part: parsed.workbook_rels_part,
            shared_strings: parsed.shared_strings,
            shared_strings_table: parsed.shared_strings_table,
            shared_strings_part: parsed.shared_strings_part,
            workbook_context: parsed.workbook_context,
            workbook_properties: parsed.workbook_properties,
            defined_names: parsed.defined_names,
            styles: parsed.styles,
            styles_part: parsed.styles_part,
            preserved_parts: parsed.preserved_parts,
            preserve_parsed_parts: parsed.preserve_parsed_parts,
            decode_formulas: parsed.decode_formulas,
        }
    }

    fn open_zip(&self) -> Result<ZipArchive<Box<dyn ReadSeek>>, ParseError> {
        match &self.source {
            WorkbookSource::Bytes(bytes) => {
                let reader: Box<dyn ReadSeek> = Box::new(Cursor::new(bytes.clone()));
                Ok(ZipArchive::new(reader)?)
            }
            #[cfg(not(target_arch = "wasm32"))]
            WorkbookSource::Path(path) => {
                let file = File::open(path)?;
                let reader: Box<dyn ReadSeek> = Box::new(file);
                Ok(ZipArchive::new(reader)?)
            }
            #[cfg(target_arch = "wasm32")]
            WorkbookSource::Path(_path) => Err(ParseError::Io(io::Error::new(
                io::ErrorKind::Unsupported,
                "cannot open XLSB from a filesystem path on wasm targets; use `open_from_bytes` or `open_from_reader`",
            ))),
        }
    }

    pub fn sheet_metas(&self) -> &[SheetMeta] {
        &self.sheets
    }

    pub fn shared_strings(&self) -> &[String] {
        &self.shared_strings
    }

    /// Shared strings with rich text / phonetic preservation.
    pub fn shared_strings_table(&self) -> &[SharedString] {
        &self.shared_strings_table
    }

    pub fn workbook_properties(&self) -> &WorkbookProperties {
        &self.workbook_properties
    }

    pub fn workbook_context(&self) -> &WorkbookContext {
        &self.workbook_context
    }

    pub fn defined_names(&self) -> &[DefinedName] {
        &self.defined_names
    }

    /// Workbook styles parsed from the styles part (typically `xl/styles.bin`).
    pub fn styles(&self) -> &Styles {
        &self.styles
    }

    /// Parse the styles part using locale-aware built-in number formats.
    ///
    /// This is a convenience wrapper around [`Styles::parse_with_locale`] that
    /// uses the preserved styles part bytes from this workbook.
    pub fn styles_with_locale(
        &self,
        locale: formula_format::Locale,
    ) -> Option<Result<Styles, ParseError>> {
        let bytes = self.styles_bin()?;
        Some(Styles::parse_with_locale(bytes, locale))
    }

    /// Raw bytes for parts that should be preserved on round-trip.
    ///
    /// Depending on [`OpenOptions`], this can include:
    /// - Parts we don't parse (unknown ZIP entries)
    /// - Parsed parts that we still want to keep byte-for-byte (e.g. `xl/workbook.bin`)
    /// - Worksheet parts (optional; can be large)
    pub fn preserved_parts(&self) -> &HashMap<String, Vec<u8>> {
        &self.preserved_parts
    }

    /// Raw styles part bytes, preserved for round-trip.
    pub fn styles_bin(&self) -> Option<&[u8]> {
        let part = self.styles_part.as_deref()?;
        self.preserved_parts.get(part).map(|v| v.as_slice())
    }

    /// Read a worksheet by index and return all discovered cells.
    ///
    /// For large sheets you likely want `for_each_cell` instead.
    pub fn read_sheet(&self, sheet_index: usize) -> Result<SheetData, ParseError> {
        let meta = self
            .sheets
            .get(sheet_index)
            .ok_or(ParseError::SheetIndexOutOfBounds(sheet_index))?;

        if let Some(bytes) = self.preserved_parts.get(&meta.part_path) {
            let mut cursor = Cursor::new(bytes.as_slice());
            return parse_sheet(
                &mut cursor,
                &self.shared_strings,
                Some(&self.shared_strings_table),
                &self.workbook_context,
                self.preserve_parsed_parts,
                self.decode_formulas,
            );
        }

        let mut zip = self.open_zip()?;
        let sheet = zip.by_name(&meta.part_path)?;
        let max = max_xlsb_zip_part_bytes();
        let size = sheet.size();
        if size > max {
            return Err(ParseError::PartTooLarge {
                part: meta.part_path.clone(),
                size,
                max,
            });
        }
        let mut sheet = sheet.take(max.saturating_add(1));

        let parsed = parse_sheet(
            &mut sheet,
            &self.shared_strings,
            Some(&self.shared_strings_table),
            &self.workbook_context,
            self.preserve_parsed_parts,
            self.decode_formulas,
        );
        if sheet.limit() == 0 {
            return Err(ParseError::PartTooLarge {
                part: meta.part_path.clone(),
                size: max.saturating_add(1),
                max,
            });
        }
        parsed
    }

    /// Read the raw worksheet `.bin` part bytes for the given sheet index.
    ///
    /// This is primarily intended for callers that want to patch a worksheet stream using
    /// [`crate::patch_sheet_bin`] and then write the workbook using
    /// [`XlsbWorkbook::save_with_part_overrides`], without forcing `preserve_worksheets=true` in
    /// [`OpenOptions`].
    pub fn worksheet_bin_bytes(&self, sheet_index: usize) -> Result<Vec<u8>, ParseError> {
        let meta = self
            .sheets
            .get(sheet_index)
            .ok_or(ParseError::SheetIndexOutOfBounds(sheet_index))?;
        let sheet_part = meta.part_path.clone();

        if let Some(bytes) = self.preserved_parts.get(&sheet_part) {
            return Ok(bytes.clone());
        }

        let mut zip = self.open_zip()?;
        read_zip_entry_required(&mut zip, &sheet_part)
    }

    /// Stream cells from a worksheet without materializing the whole sheet.
    ///
    /// This always scans the whole sheet. If you want to stop early (e.g. after finding N
    /// formulas), use [`XlsbWorkbook::for_each_cell_control_flow`].
    pub fn for_each_cell<F>(&self, sheet_index: usize, mut f: F) -> Result<(), ParseError>
    where
        F: FnMut(Cell),
    {
        self.for_each_cell_control_flow(sheet_index, |cell| {
            f(cell);
            ControlFlow::Continue(())
        })
    }

    /// Stream cells from a worksheet, allowing early exit.
    ///
    /// The callback controls iteration via [`ControlFlow`]:
    /// - `ControlFlow::Continue(())` to keep scanning
    /// - `ControlFlow::Break(())` to stop scanning the sheet early
    pub fn for_each_cell_control_flow<F>(
        &self,
        sheet_index: usize,
        mut f: F,
    ) -> Result<(), ParseError>
    where
        F: FnMut(Cell) -> ControlFlow<(), ()>,
    {
        let meta = self
            .sheets
            .get(sheet_index)
            .ok_or(ParseError::SheetIndexOutOfBounds(sheet_index))?;

        if let Some(bytes) = self.preserved_parts.get(&meta.part_path) {
            let mut cursor = Cursor::new(bytes.as_slice());
            parse_sheet_stream(
                &mut cursor,
                &self.shared_strings,
                Some(&self.shared_strings_table),
                &self.workbook_context,
                self.preserve_parsed_parts,
                self.decode_formulas,
                |cell| f(cell),
            )?;
            return Ok(());
        }

        let mut zip = self.open_zip()?;
        let sheet = zip.by_name(&meta.part_path)?;
        let max = max_xlsb_zip_part_bytes();
        let size = sheet.size();
        if size > max {
            return Err(ParseError::PartTooLarge {
                part: meta.part_path.clone(),
                size,
                max,
            });
        }
        let mut sheet = sheet.take(max.saturating_add(1));

        let parsed = parse_sheet_stream(
            &mut sheet,
            &self.shared_strings,
            Some(&self.shared_strings_table),
            &self.workbook_context,
            self.preserve_parsed_parts,
            self.decode_formulas,
            |cell| f(cell),
        );
        if sheet.limit() == 0 {
            return Err(ParseError::PartTooLarge {
                part: meta.part_path.clone(),
                size: max.saturating_add(1),
                max,
            });
        }
        parsed?;
        Ok(())
    }

    /// Save the workbook as a new `.xlsb` file.
    ///
    /// This is currently a *lossless* package writer: it repackages the original XLSB ZIP
    /// container by copying every entry's uncompressed payload byte-for-byte.
    ///
    /// The writer reads entries from the original source package (either a filesystem path or
    /// in-memory bytes). If an entry name exists in [`XlsbWorkbook::preserved_parts`], that byte
    /// payload is used as an override. This provides a forward-compatible hook for future code to
    /// patch individual parts (for example to write modified worksheets) while keeping the rest of
    /// the package intact.
    ///
    /// How [`OpenOptions`] affects `save_as`:
    /// - `preserve_unknown_parts`: stores raw bytes for unknown ZIP entries in `preserved_parts`,
    ///   but `save_as` will still copy them from the source package even when this is `false`.
    /// - `preserve_parsed_parts`: stores raw bytes for `xl/workbook.bin` and the shared strings
    ///   part (typically `xl/sharedStrings.bin`) so they can be re-emitted without re-reading
    ///   those ZIP entries.
    /// - `preserve_worksheets`: stores raw bytes for worksheet `.bin` parts (can be large). When
    ///   `false`, worksheets are streamed from the source ZIP during `save_as`.
    ///
    /// If you need to override specific parts (e.g. a patched worksheet stream), use
    /// [`XlsbWorkbook::save_with_part_overrides`].
    #[cfg(not(target_arch = "wasm32"))]
    pub fn save_as(&self, dest: impl AsRef<Path>) -> Result<(), ParseError> {
        self.save_with_part_overrides(dest, &HashMap::new())
    }

    /// Save the workbook as an `.xlsb` package written to an arbitrary writer.
    ///
    /// This is equivalent to [`Self::save_as`] but does not require a filesystem path.
    pub fn save_as_to_writer<W: Write + Seek>(&self, writer: W) -> Result<(), ParseError> {
        self.save_with_part_overrides_to_writer(writer, &HashMap::new())
    }

    /// Save the workbook as an `.xlsb` package and return the resulting ZIP bytes.
    pub fn save_as_to_bytes(&self) -> Result<Vec<u8>, ParseError> {
        let mut cursor = Cursor::new(Vec::new());
        self.save_as_to_writer(&mut cursor)?;
        Ok(cursor.into_inner())
    }

    /// Save the workbook as a password-protected/encrypted `.xlsb` file.
    ///
    /// This writes an OLE compound file wrapper containing:
    /// - `EncryptionInfo`
    /// - `EncryptedPackage` (the encrypted XLSB ZIP payload)
    #[cfg(not(target_arch = "wasm32"))]
    pub fn save_as_encrypted(
        &self,
        dest: impl AsRef<Path>,
        password: &str,
    ) -> Result<(), ParseError> {
        let dest = dest.as_ref();
        atomic_write_with_path(dest, |tmp_path| {
            let mut out = File::create(tmp_path)?;
            self.save_as_encrypted_to_writer(&mut out, password)
        })
        .map_err(|err| match err {
            AtomicWriteError::Io(err) => ParseError::Io(err),
            AtomicWriteError::Writer(err) => err,
        })
    }

    /// Save the workbook as an encrypted/password-protected `.xlsb` payload written to an
    /// arbitrary writer.
    pub fn save_as_encrypted_to_writer<W: Write>(
        &self,
        mut writer: W,
        password: &str,
    ) -> Result<(), ParseError> {
        let package_bytes = self.save_as_to_bytes()?;
        let ole_bytes = office_crypto::encrypt_package_to_ole(
            &package_bytes,
            password,
            office_crypto::EncryptOptions::default(),
        )
            .map_err(|err| ParseError::OfficeCrypto(err.to_string()))?;
        writer.write_all(&ole_bytes)?;
        Ok(())
    }

    /// Save the workbook with an updated numeric cell value.
    ///
    /// This is a convenience wrapper around the in-memory worksheet patcher ([`patch_sheet_bin`])
    /// plus the part override writer
    /// ([`XlsbWorkbook::save_with_part_overrides`]).
    ///
    /// Note: this may insert missing `BrtRow` / cell records inside `BrtSheetData` if the target
    /// cell does not already exist in the worksheet stream.
    #[cfg(not(target_arch = "wasm32"))]
    pub fn save_with_edits(
        &self,
        dest: impl AsRef<Path>,
        sheet_index: usize,
        row: u32,
        col: u32,
        value: f64,
    ) -> Result<(), ParseError> {
        self.save_with_cell_edits(
            dest,
            sheet_index,
            &[CellEdit {
                row,
                col,
                new_value: CellValue::Number(value),
                new_style: None,
                clear_formula: false,
                new_formula: None,
                new_rgcb: None,
                new_formula_flags: None,
                shared_string_index: None,
            }],
        )
    }

    /// Save the workbook with a set of edits for a single worksheet.
    ///
    /// This loads `xl/worksheets/sheetN.bin` into memory. For very large worksheets, consider
    /// [`XlsbWorkbook::save_with_cell_edits_streaming`].
    #[cfg(not(target_arch = "wasm32"))]
    pub fn save_with_cell_edits(
        &self,
        dest: impl AsRef<Path>,
        sheet_index: usize,
        edits: &[CellEdit],
    ) -> Result<(), ParseError> {
        let meta = self
            .sheets
            .get(sheet_index)
            .ok_or(ParseError::SheetIndexOutOfBounds(sheet_index))?;
        let sheet_part = meta.part_path.clone();

        let sheet_bytes = if let Some(bytes) = self.preserved_parts.get(&sheet_part) {
            bytes.clone()
        } else {
            let mut zip = self.open_zip()?;
            read_zip_entry_required(&mut zip, &sheet_part)?
        };

        let patched = patch_sheet_bin(&sheet_bytes, edits)?;
        self.save_with_part_overrides(dest, &HashMap::from([(sheet_part, patched)]))
    }

    /// Save the workbook with a set of formula edits for a single worksheet, expressed as Excel
    /// formula text.
    ///
    /// This is similar to [`XlsbWorkbook::save_with_cell_edits`], but:
    /// - accepts formula text instead of raw `rgce` bytes
    /// - when encountering forward-compatible / future functions (typically `_xlfn.*`) that map
    ///   to the BIFF UDF sentinel (255), automatically patches `xl/workbook.bin` to intern a
    ///   missing `ExternName` entry so the rgce encoder can emit a valid `PtgNameX` reference.
    #[cfg(all(feature = "write", not(target_arch = "wasm32")))]
    pub fn save_with_cell_formula_text_edits(
        &self,
        dest: impl AsRef<Path>,
        sheet_index: usize,
        edits: &[FormulaTextCellEdit],
    ) -> Result<(), ParseError> {
        use crate::ftab::{function_id_from_name, FTAB_USER_DEFINED};
        use crate::rgce::{encode_rgce_with_context_ast_in_sheet, CellCoord, EncodeError};
        use crate::workbook_bin_patch::patch_workbook_bin_intern_namex_functions;
        use crate::workbook_context::{ExternName, SupBook, SupBookKind};
        use formula_engine as fe;

        if edits.is_empty() {
            return self.save_as(dest);
        }

        let meta = self
            .sheets
            .get(sheet_index)
            .ok_or(ParseError::SheetIndexOutOfBounds(sheet_index))?;
        let sheet_part = meta.part_path.clone();
        let sheet_name = meta.name.clone();

        // Load worksheet bytes (in-memory patcher).
        let sheet_bytes = if let Some(bytes) = self.preserved_parts.get(&sheet_part) {
            bytes.clone()
        } else {
            let mut zip = self.open_zip()?;
            read_zip_entry_required(&mut zip, &sheet_part)?
        };

        // Load workbook.bin so we can patch it if we need to intern new NameX function entries.
        let workbook_bin = if let Some(bytes) = self.preserved_parts.get(&self.workbook_part) {
            bytes.clone()
        } else {
            let mut zip = self.open_zip()?;
            read_zip_entry_required(&mut zip, &self.workbook_part)?
        };

        let mut ctx = self.workbook_context.clone();

        // Collect any forward-compat / future functions (iftab=255) that are missing from the
        // workbook's NameX tables.
        let mut wanted: BTreeMap<String, String> = BTreeMap::new();
        for edit in edits {
            let ast =
                fe::parse_formula(&edit.formula, fe::ParseOptions::default()).map_err(|e| {
                    ParseError::UnsupportedFormulaText(format!(
                        "{} (span {}..{})",
                        e.message, e.span.start, e.span.end
                    ))
                })?;

            fn walk(expr: &fe::Expr, wanted: &mut BTreeMap<String, String>) {
                match expr {
                    fe::Expr::FunctionCall(call) => {
                        // Normalize `_xlfn.` prefix for stable ordering/dedup, but preserve the
                        // original name for the inserted ExternName so decoding round-trips.
                        let mut key = call.name.original.to_ascii_uppercase();
                        if let Some(stripped) = key.strip_prefix("_XLFN.") {
                            key = stripped.to_string();
                        }

                        wanted
                            .entry(key)
                            .and_modify(|existing| {
                                // Prefer a `_xlfn.`-prefixed spelling when the caller provided
                                // one, matching Excel's forward-compat namespace.
                                let existing_has_prefix =
                                    existing.to_ascii_uppercase().starts_with("_XLFN.");
                                let new_has_prefix = call
                                    .name
                                    .original
                                    .to_ascii_uppercase()
                                    .starts_with("_XLFN.");
                                if !existing_has_prefix && new_has_prefix {
                                    *existing = call.name.original.clone();
                                }
                            })
                            .or_insert_with(|| call.name.original.clone());

                        for arg in &call.args {
                            walk(arg, wanted);
                        }
                    }
                    fe::Expr::Call(call) => {
                        walk(&call.callee, wanted);
                        for arg in &call.args {
                            walk(arg, wanted);
                        }
                    }
                    fe::Expr::FieldAccess(access) => walk(&access.base, wanted),
                    fe::Expr::Array(arr) => {
                        for row in &arr.rows {
                            for el in row {
                                walk(el, wanted);
                            }
                        }
                    }
                    fe::Expr::Unary(u) => walk(&u.expr, wanted),
                    fe::Expr::Postfix(p) => walk(&p.expr, wanted),
                    fe::Expr::Binary(b) => {
                        walk(&b.left, wanted);
                        walk(&b.right, wanted);
                    }
                    _ => {}
                }
            }

            walk(&ast.expr, &mut wanted);
        }

        let mut missing: Vec<String> = Vec::new();
        for (_key, original) in wanted {
            if function_id_from_name(&original) != Some(FTAB_USER_DEFINED) {
                continue;
            }
            // `namex_function_ref` handles `_xlfn.` prefix normalization.
            if ctx.namex_function_ref(&original).is_some() {
                continue;
            }
            missing.push(original);
        }
        missing.sort_by(|a, b| a.to_ascii_uppercase().cmp(&b.to_ascii_uppercase()));
        missing.dedup_by(|a, b| a.eq_ignore_ascii_case(b));

        let patched_workbook_bin = if missing.is_empty() {
            None
        } else {
            let patch = match patch_workbook_bin_intern_namex_functions(&workbook_bin, &missing)? {
                Some(patch) => patch,
                None => {
                    return Err(ParseError::Io(io::Error::new(
                        io::ErrorKind::InvalidData,
                        "failed to patch workbook.bin for forward-compatible functions",
                    )));
                }
            };

            // Update the in-memory workbook context so formula encoding can reference the newly
            // inserted NameX entries.
            if patch.created_supbook {
                debug_assert!(
                    ctx.addin_supbook_index().is_none(),
                    "workbook.bin patch created an AddIn SupBook, but the context already had one"
                );
                let supbook_index = patch.inserted.first().map(|e| e.supbook_index).unwrap_or(0);

                // The patcher always appends the new AddIn SupBook.
                let created = ctx.push_namex_supbook(
                    SupBook {
                        raw_name: "\u{0001}".to_string(),
                        kind: SupBookKind::AddIn,
                    },
                    Vec::new(),
                );
                if created != supbook_index {
                    return Err(ParseError::Io(io::Error::new(
                        io::ErrorKind::InvalidData,
                        format!(
                            "workbook.bin patch produced unexpected supbook index: context={created}, patch={supbook_index}"
                        ),
                    )));
                }
            }

            for entry in &patch.inserted {
                ctx.insert_namex_extern_name(
                    entry.supbook_index,
                    entry.name_index,
                    ExternName {
                        name: entry.name.clone(),
                        is_function: true,
                        scope_sheet: None,
                    },
                );
            }

            Some(patch.workbook_bin)
        };

        // Encode the formula token streams with the updated context.
        let mut binary_edits: Vec<CellEdit> = Vec::with_capacity(edits.len());
        for edit in edits {
            let base = CellCoord::new(edit.row, edit.col);
            let encoded =
                encode_rgce_with_context_ast_in_sheet(&edit.formula, &ctx, &sheet_name, base)
                    .map_err(|e| match e {
                        EncodeError::Parse(msg) => ParseError::UnsupportedFormulaText(msg),
                        other => ParseError::UnsupportedFormulaText(other.to_string()),
                    })?;

            binary_edits.push(CellEdit {
                row: edit.row,
                col: edit.col,
                new_value: edit.new_value.clone(),
                new_style: None,
                clear_formula: false,
                new_formula: Some(encoded.rgce),
                new_rgcb: Some(encoded.rgcb),
                new_formula_flags: None,
                shared_string_index: None,
            });
        }

        let patched_sheet = patch_sheet_bin(&sheet_bytes, &binary_edits)?;

        let mut overrides: HashMap<String, Vec<u8>> = HashMap::new();
        overrides.insert(sheet_part, patched_sheet);
        if let Some(wb) = patched_workbook_bin {
            overrides.insert(self.workbook_part.clone(), wb);
        }

        self.save_with_part_overrides(dest, &overrides)
    }

    /// Save the workbook with a set of edits for a single worksheet, updating the shared strings
    /// part as needed so shared-string (`BrtCellIsst`) cells can stay as shared-string references.
    #[cfg(not(target_arch = "wasm32"))]
    pub fn save_with_cell_edits_shared_strings(
        &self,
        dest: impl AsRef<Path>,
        sheet_index: usize,
        edits: &[CellEdit],
    ) -> Result<(), ParseError> {
        let meta = self
            .sheets
            .get(sheet_index)
            .ok_or(ParseError::SheetIndexOutOfBounds(sheet_index))?;
        let sheet_part = meta.part_path.clone();

        let sheet_bytes = if let Some(bytes) = self.preserved_parts.get(&sheet_part) {
            bytes.clone()
        } else {
            let mut zip = self.open_zip()?;
            read_zip_entry_required(&mut zip, &sheet_part)?
        };

        let Some(shared_strings_part) = self.shared_strings_part.as_deref() else {
            // Workbook has no shared string table. Fall back to the generic patcher which may
            // convert shared-string cells to inline strings.
            return self.save_with_cell_edits(dest, sheet_index, edits);
        };

        let shared_strings_bytes = match self.preserved_parts.get(shared_strings_part) {
            Some(bytes) => bytes.clone(),
            None => {
                let mut zip = self.open_zip()?;
                match read_zip_entry(&mut zip, shared_strings_part)? {
                    Some(bytes) => bytes,
                    None => {
                        // Shared strings part went missing; fall back to the generic patcher.
                        return self.save_with_cell_edits(dest, sheet_index, edits);
                    }
                }
            }
        };

        let targets: HashSet<(u32, u32)> = edits.iter().map(|e| (e.row, e.col)).collect();
        let cell_records = if targets.is_empty() {
            HashMap::new()
        } else {
            sheet_cell_records(&sheet_bytes, &targets)?
        };

        let mut sst = SharedStringsWriter::new(shared_strings_bytes)?;

        let mut updated_edits = edits.to_vec();
        for edit in &mut updated_edits {
            let CellValue::Text(text) = &edit.new_value else {
                continue;
            };
            if edit.new_formula.is_some() || edit.new_rgcb.is_some() {
                // Formula cells store cached string results inline (BrtFmlaString). Even when the
                // workbook has a shared string table, cached formula strings do not reference it.
                continue;
            }

            let coord = (edit.row, edit.col);
            let record = cell_records.get(&coord);
            let record_id = record.map(|r| r.id);
            if record_id.is_some_and(is_formula_cell_record) && !edit.clear_formula {
                // Formula records store cached strings inline. Do not treat them as SST references
                // unless this edit explicitly clears the formula (turning the cell into a plain
                // value cell).
                continue;
            }

            if record_id == Some(biff12::CELL_ST) {
                if let Some(record) = record {
                    if value_edit_is_noop_inline_string(&record.payload, edit)? {
                        // Preserve a byte-identical worksheet stream for no-op inline-string edits.
                        // The workbook-level shared-string counts should also remain unchanged in
                        // this case, since the cell still uses `BrtCellSt` storage.
                        continue;
                    }
                }
            }

            if record_id == Some(biff12::STRING) {
                if let Some(record) = record {
                    // Preserve existing `BrtCellIsst` cells as shared-string references. When the
                    // edit is a no-op (text matches the existing shared string) keep the original
                    // `isst` so rich-text / phonetic shared strings stay byte-identical.
                    if record.payload.len() >= 12 {
                        let isst = u32::from_le_bytes(record.payload[8..12].try_into().unwrap());
                        if self
                            .shared_strings
                            .get(isst as usize)
                            .is_some_and(|s| s == text)
                            && edit.new_formula.is_none()
                            && edit.new_rgcb.is_none()
                        {
                            edit.shared_string_index = Some(isst);
                            continue;
                        }
                    }
                }
            }

            // When updating a workbook that already has a shared string table, prefer emitting
            // text cells as shared-string references (BrtCellIsst) so the worksheet stays scalable
            // and counts in the shared string table remain consistent.
            edit.shared_string_index = Some(sst.intern_plain(text)?);
        }

        let total_ref_delta: i64 = updated_edits
            .iter()
            .map(|edit| {
                let coord = (edit.row, edit.col);
                let old_id = cell_records.get(&coord).map(|r| r.id);
                let old_uses_sst = matches!(old_id, Some(biff12::STRING));
                let old_is_formula = old_id.is_some_and(is_formula_cell_record);
                let new_is_formula = if old_is_formula {
                    // Existing formula cell: it stays a formula unless explicitly cleared.
                    !edit.clear_formula
                } else {
                    // New/ non-formula cell: it's only a formula when the edit sets formula bytes.
                    edit.new_formula.is_some() || edit.new_rgcb.is_some()
                };
                let new_uses_sst = matches!(edit.new_value, CellValue::Text(_))
                    && edit.shared_string_index.is_some()
                    && !new_is_formula;
                match (old_uses_sst, new_uses_sst) {
                    (false, true) => 1,
                    (true, false) => -1,
                    _ => 0,
                }
            })
            .sum();
        sst.note_total_ref_delta(total_ref_delta)?;

        let updated_shared_strings_bytes = sst.into_bytes()?;
        let patched_sheet = patch_sheet_bin(&sheet_bytes, &updated_edits)?;

        self.save_with_part_overrides(
            dest,
            &HashMap::from([
                (sheet_part, patched_sheet),
                (
                    shared_strings_part.to_string(),
                    updated_shared_strings_bytes,
                ),
            ]),
        )
    }

    /// Save the workbook with a set of edits for a single worksheet, patching the worksheet part
    /// as a stream.
    ///
    /// This avoids loading `xl/worksheets/sheetN.bin` into memory (important for very large XLSB
    /// worksheets). Unchanged records are copied byte-for-byte, preserving varint header
    /// encodings and minimizing diffs.
    #[cfg(not(target_arch = "wasm32"))]
    pub fn save_with_cell_edits_streaming(
        &self,
        dest: impl AsRef<Path>,
        sheet_index: usize,
        edits: &[CellEdit],
    ) -> Result<(), ParseError> {
        let meta = self
            .sheets
            .get(sheet_index)
            .ok_or(ParseError::SheetIndexOutOfBounds(sheet_index))?;
        let sheet_part = meta.part_path.clone();

        self.save_with_part_overrides_streaming(
            dest,
            &HashMap::new(),
            &sheet_part,
            |input, output| patch_sheet_bin_streaming(input, output, edits),
        )
    }

    /// Save the workbook with a set of edits for a single worksheet, patching the worksheet part
    /// as a stream while updating the shared string table (if present).
    ///
    /// This avoids buffering `xl/worksheets/sheetN.bin` in memory (important for large sheets)
    /// while still preserving shared-string (`BrtCellIsst`) storage for edited text cells.
    #[cfg(not(target_arch = "wasm32"))]
    pub fn save_with_cell_edits_streaming_shared_strings(
        &self,
        dest: impl AsRef<Path>,
        sheet_index: usize,
        edits: &[CellEdit],
    ) -> Result<(), ParseError> {
        let meta = self
            .sheets
            .get(sheet_index)
            .ok_or(ParseError::SheetIndexOutOfBounds(sheet_index))?;
        let sheet_part = meta.part_path.clone();

        let Some(shared_strings_part) = self.shared_strings_part.as_deref() else {
            // Workbook has no shared string table. Fall back to the generic streaming patcher,
            // which may emit inline strings.
            return self.save_with_cell_edits_streaming(dest, sheet_index, edits);
        };

        let max_part = max_xlsb_zip_part_bytes();

        let targets: HashSet<(u32, u32)> = edits.iter().map(|e| (e.row, e.col)).collect();
        let cell_records = if targets.is_empty() {
            HashMap::new()
        } else if let Some(sheet_bytes) = self.preserved_parts.get(&sheet_part) {
            sheet_cell_records(sheet_bytes, &targets)?
        } else {
            let mut zip = self.open_zip()?;
            let entry = zip.by_name(&sheet_part)?;
            let size = entry.size();
            if size > max_part {
                return Err(ParseError::PartTooLarge {
                    part: sheet_part.clone(),
                    size,
                    max: max_part,
                });
            }
            let mut limited = entry.take(max_part.saturating_add(1));
            let parsed = sheet_cell_records_streaming(&mut limited, &targets);
            if limited.limit() == 0 {
                return Err(ParseError::PartTooLarge {
                    part: sheet_part.clone(),
                    size: max_part.saturating_add(1),
                    max: max_part,
                });
            }
            parsed?
        };

        // Build a mapping of *plain* shared strings to their indices.
        //
        // We only reuse strings that have no rich/phonetic payload:
        // - true plain `BrtSI` records (flags=0), and
        // - "effectively plain" `BrtSI` records where the rich/phonetic flag bits are set but the
        //   corresponding blocks are empty (`cRun=0` / `cb=0`).
        //
        // This avoids unintentionally applying rich/phonetic formatting to newly edited cells
        // while still deduplicating common real-world producer quirks.
        let base_si_count = u32::try_from(self.shared_strings_table.len())
            .map_err(|_| ParseError::UnexpectedEof)?;
        let mut existing_plain_to_index: HashMap<&str, u32> = HashMap::new();
        existing_plain_to_index.reserve(self.shared_strings_table.len());
        for (idx, si) in self.shared_strings_table.iter().enumerate() {
            let reusable_plain = if si.raw_si.is_none() {
                true
            } else {
                si.raw_si
                    .as_deref()
                    .is_some_and(|raw| reusable_plain_si_utf16_end(raw).is_some())
            };
            if reusable_plain {
                existing_plain_to_index
                    .entry(si.plain_text())
                    .or_insert(idx as u32);
            }
        }

        let mut new_plain_to_index: HashMap<String, u32> = HashMap::new();
        let mut appended_plain: Vec<String> = Vec::new();
        appended_plain.reserve(edits.len());

        let mut updated_edits = edits.to_vec();
        for edit in &mut updated_edits {
            let CellValue::Text(text) = &edit.new_value else {
                continue;
            };
            if edit.new_formula.is_some() || edit.new_rgcb.is_some() {
                // Formula cells store cached string results inline (BrtFmlaString). Even when the
                // workbook has a shared string table, cached formula strings do not reference it.
                continue;
            }

            let coord = (edit.row, edit.col);
            let record = cell_records.get(&coord);
            let record_id = record.map(|r| r.id);
            if record_id.is_some_and(is_formula_cell_record) && !edit.clear_formula {
                continue;
            }

            if record_id == Some(biff12::CELL_ST) {
                if let Some(record) = record {
                    if value_edit_is_noop_inline_string(&record.payload, edit)? {
                        // Preserve a byte-identical worksheet stream for no-op inline-string edits.
                        // The workbook-level shared-string counts should also remain unchanged in
                        // this case, since the cell still uses `BrtCellSt` storage.
                        continue;
                    }
                }
            }

            if record_id == Some(biff12::STRING) {
                if let Some(record) = record {
                    // No-op shared-string edit: keep the existing `isst` to avoid inserting a new
                    // (plain) `BrtSI` record when the original string has rich-text/phonetic data.
                    if record.payload.len() >= 12 {
                        let isst = u32::from_le_bytes(record.payload[8..12].try_into().unwrap());
                        if self
                            .shared_strings
                            .get(isst as usize)
                            .is_some_and(|s| s == text)
                            && edit.new_formula.is_none()
                            && edit.new_rgcb.is_none()
                        {
                            edit.shared_string_index = Some(isst);
                            continue;
                        }
                    }
                }
            }

            if let Some(&idx) = existing_plain_to_index.get(text.as_str()) {
                edit.shared_string_index = Some(idx);
            } else if let Some(&idx) = new_plain_to_index.get(text.as_str()) {
                edit.shared_string_index = Some(idx);
            } else {
                let idx = base_si_count
                    .checked_add(
                        u32::try_from(appended_plain.len())
                            .map_err(|_| ParseError::UnexpectedEof)?,
                    )
                    .ok_or(ParseError::UnexpectedEof)?;
                appended_plain.push(text.clone());
                new_plain_to_index.insert(text.clone(), idx);
                edit.shared_string_index = Some(idx);
            }
        }

        let total_ref_delta: i64 = updated_edits
            .iter()
            .map(|edit| {
                let coord = (edit.row, edit.col);
                let old_id = cell_records.get(&coord).map(|r| r.id);
                let old_uses_sst = matches!(old_id, Some(biff12::STRING));
                let old_is_formula = old_id.is_some_and(is_formula_cell_record);
                let new_is_formula = if old_is_formula {
                    !edit.clear_formula
                } else {
                    edit.new_formula.is_some() || edit.new_rgcb.is_some()
                };
                let new_uses_sst = matches!(edit.new_value, CellValue::Text(_))
                    && edit.shared_string_index.is_some()
                    && !new_is_formula;
                match (old_uses_sst, new_uses_sst) {
                    (false, true) => 1,
                    (true, false) => -1,
                    _ => 0,
                }
            })
            .sum();

        // If we don't need to change the shared string table, avoid streaming it through the
        // patcher entirely.
        if appended_plain.is_empty() && total_ref_delta == 0 {
            return self.save_with_part_overrides_streaming(
                dest,
                &HashMap::new(),
                &sheet_part,
                |input, output| patch_sheet_bin_streaming(input, output, &updated_edits),
            );
        }

        // If the shared strings part is missing from the source package, fall back to the generic
        // streaming patcher (which may convert edited text cells to inline strings).
        if !self.preserved_parts.contains_key(shared_strings_part) {
            let mut zip = self.open_zip()?;
            match zip.by_name(shared_strings_part) {
                Ok(_) => {}
                Err(zip::result::ZipError::FileNotFound) => {
                    return self.save_with_cell_edits_streaming(dest, sheet_index, edits);
                }
                Err(e) => return Err(e.into()),
            };
        }

        // Stream both the worksheet and the shared string table to avoid materializing
        // `xl/sharedStrings.bin` in memory when only a few strings are appended.
        let stream_parts = BTreeSet::from([sheet_part.clone(), shared_strings_part.to_string()]);
        self.save_with_part_overrides_streaming_multi(
            dest,
            &HashMap::new(),
            &stream_parts,
            |part_name, input, output| {
                if part_name == sheet_part.as_str() {
                    patch_sheet_bin_streaming(input, output, &updated_edits)
                } else if part_name == shared_strings_part {
                    SharedStringsWriterStreaming::patch(
                        input,
                        output,
                        &appended_plain,
                        base_si_count,
                        total_ref_delta,
                    )
                    .map_err(ParseError::from)
                } else {
                    Err(ParseError::Io(io::Error::new(
                        io::ErrorKind::InvalidInput,
                        format!("unexpected streamed part: {part_name}"),
                    )))
                }
            },
        )
    }

    /// Save the workbook with cell edits across multiple worksheets, streaming each worksheet part
    /// through the patcher while writing a single output ZIP.
    ///
    /// This avoids buffering full worksheet `.bin` payloads in memory, making it suitable for
    /// workbooks with very large sheets.
    #[cfg(not(target_arch = "wasm32"))]
    pub fn save_with_cell_edits_streaming_multi(
        &self,
        dest: impl AsRef<Path>,
        edits_by_sheet: &BTreeMap<usize, Vec<CellEdit>>,
    ) -> Result<(), ParseError> {
        let mut edits_by_part: BTreeMap<String, &[CellEdit]> = BTreeMap::new();
        for (&sheet_index, edits) in edits_by_sheet {
            if edits.is_empty() {
                continue;
            }
            let meta = self
                .sheets
                .get(sheet_index)
                .ok_or(ParseError::SheetIndexOutOfBounds(sheet_index))?;
            edits_by_part.insert(meta.part_path.clone(), edits.as_slice());
        }

        if edits_by_part.is_empty() {
            return self.save_as(dest);
        }

        let stream_parts: BTreeSet<String> = edits_by_part.keys().cloned().collect();
        self.save_with_part_overrides_streaming_multi(
            dest,
            &HashMap::new(),
            &stream_parts,
            |part_name, input, output| {
                let edits = edits_by_part.get(part_name).ok_or_else(|| {
                    ParseError::Io(io::Error::new(
                        io::ErrorKind::InvalidInput,
                        format!("missing worksheet edits for streamed part: {part_name}"),
                    ))
                })?;
                patch_sheet_bin_streaming(input, output, edits)
            },
        )
    }

    /// Save the workbook with cell edits across multiple worksheets, streaming each worksheet part
    /// while updating the shared string table (if present) so edited text cells can remain
    /// shared-string (`BrtCellIsst`) references.
    #[cfg(not(target_arch = "wasm32"))]
    pub fn save_with_cell_edits_streaming_multi_shared_strings(
        &self,
        dest: impl AsRef<Path>,
        edits_by_sheet: &BTreeMap<usize, Vec<CellEdit>>,
    ) -> Result<(), ParseError> {
        let Some(shared_strings_part) = self.shared_strings_part.as_deref() else {
            // Workbook has no shared string table. Fall back to the generic multi-sheet streaming
            // patcher, which may emit inline strings.
            return self.save_with_cell_edits_streaming_multi(dest, edits_by_sheet);
        };

        // We need a ZIP handle to stream worksheet parts while discovering existing cell record
        // types. Note that we do *not* need to read `xl/sharedStrings.bin` into memory: we can
        // intern against the parsed table and patch the binary part as a stream during the output
        // ZIP write.
        let mut zip = self.open_zip()?;
        let max_part = max_xlsb_zip_part_bytes();

        let base_si_count = u32::try_from(self.shared_strings_table.len())
            .map_err(|_| ParseError::UnexpectedEof)?;
        let mut existing_plain_to_index: HashMap<&str, u32> = HashMap::new();
        existing_plain_to_index.reserve(self.shared_strings_table.len());
        for (idx, si) in self.shared_strings_table.iter().enumerate() {
            let reusable_plain = if si.raw_si.is_none() {
                true
            } else {
                si.raw_si
                    .as_deref()
                    .is_some_and(|raw| reusable_plain_si_utf16_end(raw).is_some())
            };
            if reusable_plain {
                existing_plain_to_index
                    .entry(si.plain_text())
                    .or_insert(idx as u32);
            }
        }

        let mut new_plain_to_index: HashMap<String, u32> = HashMap::new();

        let total_possible_appends: usize = edits_by_sheet.values().map(|v| v.len()).sum();
        let mut appended_plain: Vec<String> = Vec::new();
        appended_plain.reserve(total_possible_appends);

        let mut updated_edits_by_part: BTreeMap<String, Vec<CellEdit>> = BTreeMap::new();
        let mut total_ref_delta: i64 = 0;

        for (&sheet_index, edits) in edits_by_sheet {
            if edits.is_empty() {
                continue;
            }

            let meta = self
                .sheets
                .get(sheet_index)
                .ok_or(ParseError::SheetIndexOutOfBounds(sheet_index))?;
            let sheet_part = meta.part_path.clone();

            let targets: HashSet<(u32, u32)> = edits.iter().map(|e| (e.row, e.col)).collect();
            let cell_records = if targets.is_empty() {
                HashMap::new()
            } else if let Some(sheet_bytes) = self.preserved_parts.get(&sheet_part) {
                sheet_cell_records(sheet_bytes, &targets)?
            } else {
                let entry = zip.by_name(&sheet_part)?;
                let size = entry.size();
                if size > max_part {
                    return Err(ParseError::PartTooLarge {
                        part: sheet_part.clone(),
                        size,
                        max: max_part,
                    });
                }
                let mut limited = entry.take(max_part.saturating_add(1));
                let parsed = sheet_cell_records_streaming(&mut limited, &targets);
                if limited.limit() == 0 {
                    return Err(ParseError::PartTooLarge {
                        part: sheet_part.clone(),
                        size: max_part.saturating_add(1),
                        max: max_part,
                    });
                }
                parsed?
            };

            let mut updated_edits = edits.clone();
            for edit in &mut updated_edits {
                let CellValue::Text(text) = &edit.new_value else {
                    continue;
                };
                if edit.new_formula.is_some() || edit.new_rgcb.is_some() {
                    // Formula cells store cached string results inline (BrtFmlaString). Even when
                    // the workbook has a shared string table, cached formula strings do not
                    // reference it.
                    continue;
                }

                let coord = (edit.row, edit.col);
                let record = cell_records.get(&coord);
                let record_id = record.map(|r| r.id);
                if record_id.is_some_and(is_formula_cell_record) && !edit.clear_formula {
                    continue;
                }

                if record_id == Some(biff12::CELL_ST) {
                    if let Some(record) = record {
                        if value_edit_is_noop_inline_string(&record.payload, edit)? {
                            // Preserve no-op inline-string edits without touching the shared string
                            // table.
                            continue;
                        }
                    }
                }

                if record_id == Some(biff12::STRING) {
                    if let Some(record) = record {
                        // No-op shared-string edit: keep the existing `isst` to avoid inserting a
                        // new plain `BrtSI` when the original string has rich-text/phonetic data.
                        if record.payload.len() >= 12 {
                            let isst =
                                u32::from_le_bytes(record.payload[8..12].try_into().unwrap());
                            if self
                                .shared_strings
                                .get(isst as usize)
                                .is_some_and(|s| s == text)
                            {
                                edit.shared_string_index = Some(isst);
                                continue;
                            }
                        }
                    }
                }

                if let Some(&idx) = existing_plain_to_index.get(text.as_str()) {
                    edit.shared_string_index = Some(idx);
                } else if let Some(&idx) = new_plain_to_index.get(text.as_str()) {
                    edit.shared_string_index = Some(idx);
                } else {
                    let idx = base_si_count
                        .checked_add(
                            u32::try_from(appended_plain.len())
                                .map_err(|_| ParseError::UnexpectedEof)?,
                        )
                        .ok_or(ParseError::UnexpectedEof)?;
                    appended_plain.push(text.clone());
                    new_plain_to_index.insert(text.clone(), idx);
                    edit.shared_string_index = Some(idx);
                }
            }

            let sheet_delta: i64 = updated_edits
                .iter()
                .map(|edit| {
                    let coord = (edit.row, edit.col);
                    let old_id = cell_records.get(&coord).map(|r| r.id);
                    let old_uses_sst = matches!(old_id, Some(biff12::STRING));
                    let old_is_formula = old_id.is_some_and(is_formula_cell_record);
                    let new_is_formula = if old_is_formula {
                        !edit.clear_formula
                    } else {
                        edit.new_formula.is_some() || edit.new_rgcb.is_some()
                    };
                    let new_uses_sst = matches!(edit.new_value, CellValue::Text(_))
                        && edit.shared_string_index.is_some()
                        && !new_is_formula;
                    match (old_uses_sst, new_uses_sst) {
                        (false, true) => 1,
                        (true, false) => -1,
                        _ => 0,
                    }
                })
                .sum();
            total_ref_delta = total_ref_delta
                .checked_add(sheet_delta)
                .ok_or(ParseError::UnexpectedEof)?;

            updated_edits_by_part.insert(sheet_part, updated_edits);
        }

        if updated_edits_by_part.is_empty() {
            return self.save_as(dest);
        }

        let needs_sst_patch = !appended_plain.is_empty() || total_ref_delta != 0;

        // If the shared strings part is missing from the source package, fall back to the generic
        // streaming patcher.
        if needs_sst_patch && !self.preserved_parts.contains_key(shared_strings_part) {
            match zip.by_name(shared_strings_part) {
                Ok(_) => {}
                Err(zip::result::ZipError::FileNotFound) => {
                    return self.save_with_cell_edits_streaming_multi(dest, edits_by_sheet);
                }
                Err(e) => return Err(e.into()),
            };
        }

        // Stream updated worksheet parts, and stream the shared string table only when we need to
        // patch counts / append strings.
        let mut stream_parts: BTreeSet<String> = updated_edits_by_part.keys().cloned().collect();
        if needs_sst_patch {
            stream_parts.insert(shared_strings_part.to_string());
        }

        self.save_with_part_overrides_streaming_multi(
            dest,
            &HashMap::new(),
            &stream_parts,
            |part_name, input, output| {
                if part_name == shared_strings_part {
                    SharedStringsWriterStreaming::patch(
                        input,
                        output,
                        &appended_plain,
                        base_si_count,
                        total_ref_delta,
                    )
                    .map_err(ParseError::from)
                } else {
                    let edits = updated_edits_by_part.get(part_name).ok_or_else(|| {
                        ParseError::Io(io::Error::new(
                            io::ErrorKind::InvalidInput,
                            format!("missing worksheet edits for streamed part: {part_name}"),
                        ))
                    })?;
                    patch_sheet_bin_streaming(input, output, edits)
                }
            },
        )
    }

    /// Save the workbook with cell edits across multiple worksheets.
    ///
    /// This is a convenience wrapper around the in-memory worksheet patcher
    /// ([`patch_sheet_bin`]) plus the part override writer
    /// ([`XlsbWorkbook::save_with_part_overrides`]).
    ///
    /// - Each affected worksheet part is read once (from `preserved_parts` if present, otherwise
    ///   streamed from the source ZIP).
    /// - Each sheet is patched once with all of its edits.
    /// - The workbook is then written with a single call to `save_with_part_overrides`.
    #[cfg(not(target_arch = "wasm32"))]
    pub fn save_with_cell_edits_multi(
        &self,
        dest: impl AsRef<Path>,
        edits_by_sheet: &BTreeMap<usize, Vec<CellEdit>>,
    ) -> Result<(), ParseError> {
        let mut overrides: HashMap<String, Vec<u8>> = HashMap::new();
        let mut zip: Option<ZipArchive<Box<dyn ReadSeek>>> = None;

        for (&sheet_index, edits) in edits_by_sheet {
            if edits.is_empty() {
                continue;
            }

            let meta = self
                .sheets
                .get(sheet_index)
                .ok_or(ParseError::SheetIndexOutOfBounds(sheet_index))?;
            let sheet_part = meta.part_path.clone();

            let patched = if let Some(bytes) = self.preserved_parts.get(&sheet_part) {
                patch_sheet_bin(bytes, edits)?
            } else {
                let zip = match zip.as_mut() {
                    Some(zip) => zip,
                    None => {
                        zip = Some(self.open_zip()?);
                        zip.as_mut().expect("zip just initialized")
                    }
                };
                let bytes = read_zip_entry_required(zip, &sheet_part)?;
                patch_sheet_bin(&bytes, edits)?
            };

            overrides.insert(sheet_part, patched);
        }

        self.save_with_part_overrides(dest, &overrides)
    }

    /// Save the workbook while overriding specific part payloads.
    ///
    /// `overrides` maps ZIP entry paths (e.g. `xl/worksheets/sheet1.bin`) to replacement bytes.
    /// All other parts are copied from the source workbook, except for any entry already present
    /// in [`XlsbWorkbook::preserved_parts`], which is emitted from that buffer.
    ///
    /// If any overridden worksheet part differs from the original package, we treat this as an
    /// edited save and remove `xl/calcChain.bin` (if present) and its references from:
    /// - `[Content_Types].xml`
    /// - the workbook relationships part (typically `xl/_rels/workbook.bin.rels`)
    ///
    /// A stale calcChain can cause Excel to open with incorrect cached results or spend time
    /// rebuilding the chain.
    #[cfg(not(target_arch = "wasm32"))]
    pub fn save_with_part_overrides(
        &self,
        dest: impl AsRef<Path>,
        overrides: &HashMap<String, Vec<u8>>,
    ) -> Result<(), ParseError> {
        let dest = dest.as_ref();
        atomic_write_with_path(dest, |tmp_path| {
            let out = File::create(tmp_path)?;
            self.save_with_part_overrides_to_writer(out, overrides)
        })
        .map_err(|err| match err {
            AtomicWriteError::Io(err) => ParseError::Io(err),
            AtomicWriteError::Writer(err) => err,
        })
    }

    /// Save the workbook while overriding specific part payloads, writing the output ZIP to
    /// `writer`.
    ///
    /// This is the writer-based equivalent of [`Self::save_with_part_overrides`].
    pub fn save_with_part_overrides_to_writer<W: Write + Seek>(
        &self,
        writer: W,
        overrides: &HashMap<String, Vec<u8>>,
    ) -> Result<(), ParseError> {
        let mut zip = self.open_zip()?;
        let max_part = max_xlsb_zip_part_bytes();

        let edited = worksheets_edited(&mut zip, &self.sheets, overrides)?;

        let ignored_overrides: HashSet<String> = if edited {
            overrides
                .keys()
                .filter(|key| is_calc_chain_part_name(key))
                .cloned()
                .collect()
        } else {
            HashSet::new()
        };

        // Compute updated plumbing parts if we need to invalidate calcChain.
        let mut updated_content_types: Option<Vec<u8>> = None;
        let mut updated_workbook_rels: Option<Vec<u8>> = None;
        let mut updated_workbook_bin: Option<Vec<u8>> = None;

        if edited {
            let content_types = get_part_bytes(
                &mut zip,
                &self.preserved_parts,
                overrides,
                "[Content_Types].xml",
            )?;
            if let Some(content_types) = content_types {
                updated_content_types = Some(remove_calc_chain_from_content_types(&content_types)?);
            }

            let workbook_rels = get_part_bytes(
                &mut zip,
                &self.preserved_parts,
                overrides,
                &self.workbook_rels_part,
            )?;
            if let Some(workbook_rels) = workbook_rels {
                updated_workbook_rels = Some(remove_calc_chain_from_workbook_rels(&workbook_rels)?);
            }

            let workbook_bin = get_part_bytes(
                &mut zip,
                &self.preserved_parts,
                overrides,
                &self.workbook_part,
            )?;
            if let Some(workbook_bin) = workbook_bin {
                if let Some(patched) = patch_workbook_bin_full_calc_on_load(&workbook_bin)? {
                    updated_workbook_bin = Some(patched);
                }
            }
        }

        let mut zip_writer = ZipWriter::new(writer);

        // Use a consistent compression method for output. This does *not* affect payload
        // preservation: we always copy/write the uncompressed part bytes.
        let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

        let mut used_overrides: HashSet<String> = HashSet::new();

        for i in 0..zip.len() {
            let entry = zip.by_index(i)?;
            let name = entry.name().to_string();

            if entry.is_dir() {
                // Directory entries are optional in ZIPs, but we recreate them when present to
                // preserve the package layout more closely.
                zip_writer.add_directory(name, options.clone())?;
                continue;
            }

            // Drop calcChain when any worksheet was edited.
            if edited && is_calc_chain_part_name(&name) {
                if overrides.contains_key(&name) {
                    used_overrides.insert(name);
                }
                continue;
            }

            zip_writer.start_file(name.as_str(), options.clone())?;

            // When invalidating calcChain, we may need to rewrite XML parts even if they're
            // present in `overrides`.
            if edited && is_content_types_part_name(&name) {
                if let Some(updated) = &updated_content_types {
                    if overrides.contains_key(&name) {
                        used_overrides.insert(name.clone());
                    }
                    zip_writer.write_all(updated)?;
                    continue;
                }
            }
            if edited && name == self.workbook_rels_part {
                if let Some(updated) = &updated_workbook_rels {
                    if overrides.contains_key(&name) {
                        used_overrides.insert(name.clone());
                    }
                    zip_writer.write_all(updated)?;
                    continue;
                }
            }
            if edited && name == self.workbook_part {
                if let Some(updated) = &updated_workbook_bin {
                    if overrides.contains_key(&name) {
                        used_overrides.insert(name.clone());
                    }
                    zip_writer.write_all(updated)?;
                    continue;
                }
            }

            if let Some(bytes) = overrides.get(&name) {
                used_overrides.insert(name.clone());
                zip_writer.write_all(bytes)?;
            } else if let Some(bytes) = self.preserved_parts.get(&name) {
                zip_writer.write_all(bytes)?;
            } else {
                let size = entry.size();
                if size > max_part {
                    return Err(ParseError::PartTooLarge {
                        part: name,
                        size,
                        max: max_part,
                    });
                }
                let mut limited = entry.take(max_part.saturating_add(1));
                let copied = io::copy(&mut limited, &mut zip_writer)?;
                if copied > max_part {
                    return Err(ParseError::PartTooLarge {
                        part: name,
                        size: copied,
                        max: max_part,
                    });
                }
            }
        }

        if used_overrides.len() + ignored_overrides.len() != overrides.len() {
            let mut missing = Vec::new();
            for key in overrides.keys() {
                if !used_overrides.contains(key) && !ignored_overrides.contains(key) {
                    missing.push(key.clone());
                }
            }
            missing.sort();
            return Err(ParseError::Io(io::Error::new(
                io::ErrorKind::InvalidInput,
                format!(
                    "override parts not found in source package: {}",
                    missing.join(", ")
                ),
            )));
        }

        zip_writer.finish()?;
        Ok(())
    }

    #[cfg(not(target_arch = "wasm32"))]
    fn save_with_part_overrides_streaming_multi<F>(
        &self,
        dest: impl AsRef<Path>,
        overrides: &HashMap<String, Vec<u8>>,
        stream_parts: &BTreeSet<String>,
        stream_override: F,
    ) -> Result<(), ParseError>
    where
        F: Fn(&str, &mut dyn Read, &mut dyn Write) -> Result<bool, ParseError>,
    {
        if stream_parts.is_empty() {
            return self.save_with_part_overrides(dest, overrides);
        }

        for part in stream_parts {
            if overrides.contains_key(part) {
                return Err(ParseError::Io(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    format!("streaming override conflicts with byte override for part: {part}"),
                )));
            }
        }

        let dest = dest.as_ref();
        atomic_write_with_path(dest, |tmp_path| {
            let mut zip = self.open_zip()?;
            let max_part = max_xlsb_zip_part_bytes();

            let edited_by_bytes = worksheets_edited(&mut zip, &self.sheets, overrides)?;

            let worksheet_paths: HashSet<&str> =
                self.sheets.iter().map(|s| s.part_path.as_str()).collect();

            let mut edited_by_stream = false;
            for stream_part in stream_parts {
                if !worksheet_paths.contains(stream_part.as_str()) {
                    continue;
                }

                let mut sink = io::sink();
                let edited = if let Some(bytes) = self.preserved_parts.get(stream_part) {
                    let mut cursor = Cursor::new(bytes);
                    stream_override(stream_part, &mut cursor, &mut sink)?
                } else {
                    let entry = zip.by_name(stream_part)?;
                    let size = entry.size();
                    if size > max_part {
                        return Err(ParseError::PartTooLarge {
                            part: stream_part.clone(),
                            size,
                            max: max_part,
                        });
                    }
                    let mut limited = entry.take(max_part.saturating_add(1));
                    let result = stream_override(stream_part, &mut limited, &mut sink);
                    if limited.limit() == 0 {
                        return Err(ParseError::PartTooLarge {
                            part: stream_part.clone(),
                            size: max_part.saturating_add(1),
                            max: max_part,
                        });
                    }
                    result?
                };
                edited_by_stream = edited_by_stream || edited;
            }

            let edited = edited_by_bytes || edited_by_stream;

            let ignored_overrides: HashSet<String> = if edited {
                overrides
                    .keys()
                    .filter(|key| is_calc_chain_part_name(key))
                    .cloned()
                    .collect()
            } else {
                HashSet::new()
            };

            // Compute updated plumbing parts if we need to invalidate calcChain.
            let mut updated_content_types: Option<Vec<u8>> = None;
            let mut updated_workbook_rels: Option<Vec<u8>> = None;
            let mut updated_workbook_bin: Option<Vec<u8>> = None;

            if edited {
                let content_types = get_part_bytes(
                    &mut zip,
                    &self.preserved_parts,
                    overrides,
                    "[Content_Types].xml",
                )?;
                if let Some(content_types) = content_types {
                    updated_content_types =
                        Some(remove_calc_chain_from_content_types(&content_types)?);
                }

                let workbook_rels = get_part_bytes(
                    &mut zip,
                    &self.preserved_parts,
                    overrides,
                    &self.workbook_rels_part,
                )?;
                if let Some(workbook_rels) = workbook_rels {
                    updated_workbook_rels =
                        Some(remove_calc_chain_from_workbook_rels(&workbook_rels)?);
                }

                let workbook_bin = get_part_bytes(
                    &mut zip,
                    &self.preserved_parts,
                    overrides,
                    &self.workbook_part,
                )?;
                if let Some(workbook_bin) = workbook_bin {
                    if let Some(patched) = patch_workbook_bin_full_calc_on_load(&workbook_bin)? {
                        updated_workbook_bin = Some(patched);
                    }
                }
            }

            let out = File::create(tmp_path)?;
            let mut writer = ZipWriter::new(out);

            // Use a consistent compression method for output. This does *not* affect payload
            // preservation: we always copy/write the uncompressed part bytes.
            let options =
                FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

            let mut used_overrides: HashSet<String> = HashSet::new();
            let mut used_stream_overrides: HashSet<String> = HashSet::new();

            for i in 0..zip.len() {
                let entry = zip.by_index(i)?;
                let name = entry.name().to_string();

                if entry.is_dir() {
                    writer.add_directory(name, options.clone())?;
                    continue;
                }

                // Drop calcChain when any worksheet was edited.
                if edited && is_calc_chain_part_name(&name) {
                    if overrides.contains_key(&name) {
                        used_overrides.insert(name);
                    }
                    continue;
                }

                writer.start_file(name.as_str(), options.clone())?;

                // When invalidating calcChain, we may need to rewrite XML parts even if they're
                // present in `overrides`.
                if edited && is_content_types_part_name(&name) {
                    if let Some(updated) = &updated_content_types {
                        if overrides.contains_key(&name) {
                            used_overrides.insert(name.clone());
                        }
                        writer.write_all(updated)?;
                        continue;
                    }
                }
                if edited && name == self.workbook_rels_part {
                    if let Some(updated) = &updated_workbook_rels {
                        if overrides.contains_key(&name) {
                            used_overrides.insert(name.clone());
                        }
                        writer.write_all(updated)?;
                        continue;
                    }
                }
                if edited && name == self.workbook_part {
                    if let Some(updated) = &updated_workbook_bin {
                        if overrides.contains_key(&name) {
                            used_overrides.insert(name.clone());
                        }
                        writer.write_all(updated)?;
                        continue;
                    }
                }

                if stream_parts.contains(&name) {
                    used_stream_overrides.insert(name.clone());
                    if let Some(bytes) = self.preserved_parts.get(&name) {
                        let mut cursor = Cursor::new(bytes);
                        stream_override(&name, &mut cursor, &mut writer)?;
                    } else {
                        let size = entry.size();
                        if size > max_part {
                            return Err(ParseError::PartTooLarge {
                                part: name,
                                size,
                                max: max_part,
                            });
                        }
                        let mut limited = entry.take(max_part.saturating_add(1));
                        let result = stream_override(&name, &mut limited, &mut writer);
                        if limited.limit() == 0 {
                            return Err(ParseError::PartTooLarge {
                                part: name,
                                size: max_part.saturating_add(1),
                                max: max_part,
                            });
                        }
                        result?;
                    }
                    continue;
                }

                if let Some(bytes) = overrides.get(&name) {
                    used_overrides.insert(name.clone());
                    writer.write_all(bytes)?;
                } else if let Some(bytes) = self.preserved_parts.get(&name) {
                    writer.write_all(bytes)?;
                } else {
                    let size = entry.size();
                    if size > max_part {
                        return Err(ParseError::PartTooLarge {
                            part: name,
                            size,
                            max: max_part,
                        });
                    }
                    let mut limited = entry.take(max_part.saturating_add(1));
                    let copied = io::copy(&mut limited, &mut writer)?;
                    if copied > max_part {
                        return Err(ParseError::PartTooLarge {
                            part: name,
                            size: copied,
                            max: max_part,
                        });
                    }
                }
            }

            if used_stream_overrides.len() != stream_parts.len() {
                let mut missing: Vec<String> = stream_parts
                    .iter()
                    .filter(|part| !used_stream_overrides.contains(part.as_str()))
                    .cloned()
                    .collect();
                missing.sort();
                return Err(ParseError::Io(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    format!(
                        "override parts not found in source package: {}",
                        missing.join(", ")
                    ),
                )));
            }

            if used_overrides.len() + ignored_overrides.len() != overrides.len() {
                let mut missing = Vec::new();
                for key in overrides.keys() {
                    if !used_overrides.contains(key) && !ignored_overrides.contains(key) {
                        missing.push(key.clone());
                    }
                }
                missing.sort();
                return Err(ParseError::Io(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    format!(
                        "override parts not found in source package: {}",
                        missing.join(", ")
                    ),
                )));
            }

            writer.finish()?;
            Ok(())
        })
        .map_err(|err| match err {
            AtomicWriteError::Io(err) => ParseError::Io(err),
            AtomicWriteError::Writer(err) => err,
        })
    }

    /// Save the workbook while overriding a single part via a streaming patch callback.
    ///
    /// This is similar to [`XlsbWorkbook::save_with_part_overrides`], but allows generating a
    /// replacement payload for `stream_part` without first buffering the entire part in memory.
    ///
    /// The callback is invoked twice when `stream_part` is a worksheet:
    /// 1) once writing to an `io::sink()` to determine whether the part would change, which drives
    ///    calcChain invalidation behavior,
    /// 2) once during the actual ZIP write.
    #[cfg(not(target_arch = "wasm32"))]
    pub fn save_with_part_overrides_streaming<F>(
        &self,
        dest: impl AsRef<Path>,
        overrides: &HashMap<String, Vec<u8>>,
        stream_part: &str,
        stream_override: F,
    ) -> Result<(), ParseError>
    where
        F: Fn(&mut dyn Read, &mut dyn Write) -> Result<bool, ParseError>,
    {
        let stream_parts = BTreeSet::from([stream_part.to_string()]);
        self.save_with_part_overrides_streaming_multi(
            dest,
            overrides,
            &stream_parts,
            |_part_name, input, output| stream_override(input, output),
        )
    }
}

/// Best-effort preflight to reject ZIP archives with an excessive number of entries *before*
/// constructing a `zip::ZipArchive`.
///
/// `zip::ZipArchive::new` eagerly reads the central directory and allocates metadata for every
/// entry. For pathological archives with millions of entries, this can consume substantial memory
/// even if we later bail out. We therefore parse the end-of-central-directory records ourselves to
/// obtain the entry count and enforce [`MAX_XLSB_ZIP_ENTRIES`] early.
fn preflight_zip_entry_count<R: Read + Seek>(reader: &mut R) -> Result<(), ParseError> {
    let max_entries = max_xlsb_zip_entries();

    // Seek to end so we can read the EOCD record from the tail.
    let end = reader.seek(SeekFrom::End(0))?;

    // Minimum EOCD record size is 22 bytes.
    if end < 22 {
        return Ok(());
    }

    // The EOCD record is located within the last 22 + 65535 bytes (max comment length).
    // Add 20 bytes to also include the Zip64 EOCD locator which immediately precedes the EOCD.
    const EOCD_MIN: u64 = 22;
    const EOCD_MAX_COMMENT: u64 = 65_535;
    const ZIP64_LOCATOR_LEN: u64 = 20;
    let tail_len = end.min(EOCD_MIN + EOCD_MAX_COMMENT + ZIP64_LOCATOR_LEN) as usize;

    // Read tail bytes into memory for signature search.
    reader.seek(SeekFrom::End(-(tail_len as i64)))?;
    let mut tail = vec![0u8; tail_len];
    reader.read_exact(&mut tail)?;

    // Search backwards for the EOCD signature, validating the comment length so we don't match a
    // signature embedded inside the ZIP comment itself.
    const EOCD_SIG: [u8; 4] = [0x50, 0x4B, 0x05, 0x06]; // PK\05\06
    let mut eocd_pos: Option<usize> = None;
    let max_start = tail.len().saturating_sub(22);
    for i in (0..=max_start).rev() {
        if tail.get(i..i + 4) != Some(&EOCD_SIG) {
            continue;
        }
        let comment_len = u16::from_le_bytes(tail[i + 20..i + 22].try_into().unwrap()) as usize;
        if i + 22 + comment_len == tail.len() {
            eocd_pos = Some(i);
            break;
        }
    }

    let Some(eocd_pos) = eocd_pos else {
        // Could not find EOCD; fall back to `zip::ZipArchive` which will return an error.
        return Ok(());
    };

    // EOCD total entry count is a u16 at offset 10.
    let total_entries_u16 =
        u16::from_le_bytes(tail[eocd_pos + 10..eocd_pos + 12].try_into().unwrap());

    let total_entries_u64 = if total_entries_u16 != 0xFFFF {
        total_entries_u16 as u64
    } else {
        // Zip64: read the Zip64 EOCD locator (20 bytes) immediately before the EOCD record.
        if eocd_pos < ZIP64_LOCATOR_LEN as usize {
            // EOCD indicates zip64 but we don't have the locator bytes. Treat this as a hard
            // failure to avoid constructing a potentially huge `ZipArchive`.
            return Err(ParseError::TooManyZipEntries {
                count: usize::MAX,
                max: max_entries,
            });
        }
        let locator_pos = eocd_pos - (ZIP64_LOCATOR_LEN as usize);
        const ZIP64_LOCATOR_SIG: [u8; 4] = [0x50, 0x4B, 0x06, 0x07]; // PK\06\07
        if tail.get(locator_pos..locator_pos + 4) != Some(&ZIP64_LOCATOR_SIG) {
            // Some ZIPs may legitimately have exactly 65535 entries without zip64. In that case,
            // treat 0xFFFF as the literal count.
            0xFFFFu64
        } else {
            let zip64_eocd_offset =
                u64::from_le_bytes(tail[locator_pos + 8..locator_pos + 16].try_into().unwrap());

            // Save current position so we can restore after reading the zip64 EOCD record.
            let cur = reader.seek(SeekFrom::Current(0))?;
            reader.seek(SeekFrom::Start(zip64_eocd_offset))?;

            // Zip64 EOCD record has a fixed header of at least 56 bytes; total entries is at
            // offset 32..40 in that header.
            let mut hdr = [0u8; 56];
            reader.read_exact(&mut hdr)?;
            reader.seek(SeekFrom::Start(cur))?;

            const ZIP64_EOCD_SIG: [u8; 4] = [0x50, 0x4B, 0x06, 0x06]; // PK\06\06
            if hdr[..4] != ZIP64_EOCD_SIG {
                return Err(ParseError::Zip(zip::result::ZipError::InvalidArchive(
                    "invalid ZIP64 EOCD signature",
                )));
            }

            u64::from_le_bytes(hdr[32..40].try_into().unwrap())
        }
    };

    let total_entries = usize::try_from(total_entries_u64).unwrap_or(usize::MAX);
    if total_entries > max_entries {
        return Err(ParseError::TooManyZipEntries {
            count: total_entries,
            max: max_entries,
        });
    }

    Ok(())
}

fn parse_xlsb_from_zip<R: Read + Seek>(
    zip: &mut ZipArchive<R>,
    options: OpenOptions,
) -> Result<ParsedWorkbook, ParseError> {
    let max_entries = max_xlsb_zip_entries();
    let entry_count = zip.len();
    if entry_count > max_entries {
        return Err(ParseError::TooManyZipEntries {
            count: entry_count,
            max: max_entries,
        });
    }

    let mut preserved_parts = HashMap::new();
    let mut preserved_total_bytes: u64 = 0;

    // Preserve package-level plumbing we don't parse but will need to re-emit on round-trip.
    preserve_part(
        zip,
        &mut preserved_parts,
        &mut preserved_total_bytes,
        "[Content_Types].xml",
    )?;
    preserve_part(
        zip,
        &mut preserved_parts,
        &mut preserved_total_bytes,
        "_rels/.rels",
    )?;

    let workbook_part = preserved_parts
        .get("_rels/.rels")
        .and_then(|root_rels| {
            match relationship_target_by_type(root_rels, OFFICE_DOCUMENT_REL_TYPE) {
                Ok(Some(target)) => root_target_candidates(&target)
                    .into_iter()
                    .find_map(|candidate| find_zip_entry_case_insensitive(zip, &candidate)),
                Ok(None) => None,
                // Be tolerant of malformed root relationships and fall back to the default
                // workbook location, matching historical behavior.
                Err(_) => None,
            }
        })
        .or_else(|| find_zip_entry_case_insensitive(zip, DEFAULT_WORKBOOK_PART))
        .ok_or_else(|| ParseError::Zip(zip::result::ZipError::FileNotFound))?;

    let workbook_rels_part = {
        let candidate = rels_part_name_for_part(&workbook_part);
        find_zip_entry_case_insensitive(zip, &candidate)
            .or_else(|| find_zip_entry_case_insensitive(zip, DEFAULT_WORKBOOK_RELS_PART))
            .ok_or_else(|| ParseError::Zip(zip::result::ZipError::FileNotFound))?
    };

    let workbook_rels_bytes = read_zip_entry(zip, &workbook_rels_part)?
        .ok_or_else(|| ParseError::Zip(zip::result::ZipError::FileNotFound))?;
    insert_preserved_part(
        &mut preserved_parts,
        &mut preserved_total_bytes,
        workbook_rels_part.clone(),
        workbook_rels_bytes.clone(),
    )?;
    let workbook_rels = parse_relationships(&workbook_rels_bytes)?;

    let content_types_xml = preserved_parts
        .get("[Content_Types].xml")
        .map(|bytes| bytes.as_slice());

    let shared_strings_part = resolve_workbook_part_name(
        zip,
        &workbook_rels_bytes,
        content_types_xml,
        SHARED_STRINGS_REL_TYPE,
        Some(SHARED_STRINGS_CONTENT_TYPE),
        DEFAULT_SHARED_STRINGS_PART,
    )?;

    let styles_part = resolve_workbook_part_name(
        zip,
        &workbook_rels_bytes,
        content_types_xml,
        STYLES_REL_TYPE,
        Some(STYLES_CONTENT_TYPE),
        DEFAULT_STYLES_PART,
    )?;

    // Styles are required for round-trip. We also parse `cellXfs` so callers can
    // resolve per-cell `style` indices to number formats (e.g. dates).
    let styles_bin = match styles_part.as_deref() {
        Some(part) => read_zip_entry(zip, part)?,
        None => None,
    };
    let styles = match styles_bin.as_deref() {
        Some(bytes) => Styles::parse(bytes).unwrap_or_default(),
        None => Styles::default(),
    };
    if let Some(bytes) = styles_bin {
        insert_preserved_part(
            &mut preserved_parts,
            &mut preserved_total_bytes,
            styles_part
                .clone()
                .unwrap_or_else(|| DEFAULT_STYLES_PART.to_string()),
            bytes,
        )?;
    }

    // `workbook.bin` can be large. When we don't need to preserve raw bytes, parse it directly
    // from the ZIP entry stream to avoid buffering the entire part into memory.
    let (mut sheets, workbook_context, workbook_properties, defined_names, workbook_bin) =
        if options.preserve_parsed_parts {
            let workbook_bin = read_zip_entry_required(zip, &workbook_part)?;
            let (sheets, ctx, props, defined_names) = parse_workbook(
                &mut Cursor::new(&workbook_bin),
                &workbook_rels,
                options.decode_formulas,
            )?;
            (sheets, ctx, props, defined_names, Some(workbook_bin))
        } else {
            let max = max_xlsb_zip_part_bytes();
            let wb = zip.by_name(&workbook_part)?;
            let size = wb.size();
            if size > max {
                return Err(ParseError::PartTooLarge {
                    part: workbook_part.clone(),
                    size,
                    max,
                });
            }
            let mut limited = wb.take(max.saturating_add(1));
            let parsed = parse_workbook(&mut limited, &workbook_rels, options.decode_formulas);
            if limited.limit() == 0 {
                return Err(ParseError::PartTooLarge {
                    part: workbook_part.clone(),
                    size: max.saturating_add(1),
                    max,
                });
            }
            let (sheets, ctx, props, defined_names) = parsed?;
            (sheets, ctx, props, defined_names, None)
        };
    let mut workbook_context = workbook_context;

    load_table_definitions(zip, &mut workbook_context, &sheets)?;

    // `workbook.bin.rels` targets are compared case-insensitively by Excel on Windows/macOS,
    // but ZIP entry names are case-sensitive. Normalize sheet part paths to the exact entry
    // name present in the archive so downstream reads/writes don't fail on case-only
    // mismatches.
    for sheet in &mut sheets {
        if let Some(actual) = find_zip_entry_case_insensitive(zip, &sheet.part_path) {
            sheet.part_path = actual;
        }
    }
    if let Some(workbook_bin) = workbook_bin.filter(|_| options.preserve_parsed_parts) {
        insert_preserved_part(
            &mut preserved_parts,
            &mut preserved_total_bytes,
            workbook_part.clone(),
            workbook_bin,
        )?;
    }

    let shared_strings = match shared_strings_part.as_deref() {
        Some(part) => {
            if options.preserve_parsed_parts {
                match read_zip_entry(zip, part)? {
                    Some(bytes) => {
                        let table = parse_shared_strings(&mut Cursor::new(&bytes))?;
                        let strings = table.iter().map(|s| s.plain_text().to_string()).collect();
                        insert_preserved_part(
                            &mut preserved_parts,
                            &mut preserved_total_bytes,
                            part.to_string(),
                            bytes,
                        )?;
                        (strings, table)
                    }
                    None => (Vec::new(), Vec::new()),
                }
            } else {
                // Like `workbook.bin`, shared strings can be very large. Stream parse when we don't
                // need raw bytes for round-trip preservation.
                match zip.by_name(part) {
                    Ok(sst) => {
                        let max = max_xlsb_zip_part_bytes();
                        let size = sst.size();
                        if size > max {
                            return Err(ParseError::PartTooLarge {
                                part: part.to_string(),
                                size,
                                max,
                            });
                        }
                        let mut limited = sst.take(max.saturating_add(1));
                        let parsed = parse_shared_strings(&mut limited);
                        if limited.limit() == 0 {
                            return Err(ParseError::PartTooLarge {
                                part: part.to_string(),
                                size: max.saturating_add(1),
                                max,
                            });
                        }
                        let table = parsed?;
                        let strings = table.iter().map(|s| s.plain_text().to_string()).collect();
                        (strings, table)
                    }
                    Err(zip::result::ZipError::FileNotFound) => (Vec::new(), Vec::new()),
                    Err(e) => return Err(e.into()),
                }
            }
        }
        None => (Vec::new(), Vec::new()),
    };

    let (shared_strings, shared_strings_table) = shared_strings;

    // Treat part names as case-insensitive and tolerate a leading `/` in ZIP entries (some
    // producers emit them). Use normalized lowercase names for comparison so we don't
    // accidentally preserve known parts twice (e.g. both `[Content_Types].xml` and
    // `/[Content_Types].xml`).
    let mut known_parts: HashSet<String> = [
        "[Content_Types].xml",
        "_rels/.rels",
        DEFAULT_WORKBOOK_PART,
        DEFAULT_WORKBOOK_RELS_PART,
        DEFAULT_SHARED_STRINGS_PART,
        DEFAULT_STYLES_PART,
    ]
    .into_iter()
    .map(|name| normalize_zip_part_name(name).to_ascii_lowercase())
    .collect();
    known_parts.insert(normalize_zip_part_name(&workbook_part).to_ascii_lowercase());
    known_parts.insert(normalize_zip_part_name(&workbook_rels_part).to_ascii_lowercase());
    if let Some(part) = &shared_strings_part {
        known_parts.insert(normalize_zip_part_name(part).to_ascii_lowercase());
    }
    if let Some(part) = &styles_part {
        known_parts.insert(normalize_zip_part_name(part).to_ascii_lowercase());
    }

    let worksheet_paths: HashSet<String> = sheets
        .iter()
        .map(|s| normalize_zip_part_name(&s.part_path).to_ascii_lowercase())
        .collect();
    if options.preserve_unknown_parts {
        for name in zip.file_names().map(str::to_string).collect::<Vec<_>>() {
            let normalized = normalize_zip_part_name(&name).to_ascii_lowercase();
            let is_known =
                known_parts.contains(&normalized) || worksheet_paths.contains(&normalized);
            if is_known {
                continue;
            }
            if let Some(bytes) = read_zip_entry(zip, &name)? {
                insert_preserved_part(
                    &mut preserved_parts,
                    &mut preserved_total_bytes,
                    name,
                    bytes,
                )?;
            }
        }
    }

    if options.preserve_worksheets {
        for sheet in &sheets {
            if let Some(bytes) = read_zip_entry(zip, &sheet.part_path)? {
                insert_preserved_part(
                    &mut preserved_parts,
                    &mut preserved_total_bytes,
                    sheet.part_path.clone(),
                    bytes,
                )?;
            }
        }
    }

    Ok(ParsedWorkbook {
        sheets,
        workbook_part,
        workbook_rels_part,
        shared_strings,
        shared_strings_table,
        shared_strings_part,
        workbook_context,
        workbook_properties,
        defined_names,
        styles,
        styles_part,
        preserved_parts,
        preserve_parsed_parts: options.preserve_parsed_parts,
        decode_formulas: options.decode_formulas,
    })
}

#[cfg(test)]
mod zip_guardrail_tests {
    use super::*;
    use std::io::{Cursor, Write};
    use std::sync::Mutex;
    use tempfile::NamedTempFile;

    static ENV_LOCK: Mutex<()> = Mutex::new(());

    struct EnvVarGuard {
        key: &'static str,
        old: Option<String>,
    }

    impl EnvVarGuard {
        fn set(key: &'static str, value: &str) -> Self {
            let old = std::env::var(key).ok();
            std::env::set_var(key, value);
            Self { key, old }
        }
    }

    impl Drop for EnvVarGuard {
        fn drop(&mut self) {
            match self.old.take() {
                Some(v) => std::env::set_var(self.key, v),
                None => std::env::remove_var(self.key),
            }
        }
    }

    fn build_minimal_xlsb_zip(extra_parts: &[(&str, &[u8])]) -> Vec<u8> {
        let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
        let options = FileOptions::<()>::default().compression_method(CompressionMethod::Stored);

        // Minimal workbook plumbing required by `parse_xlsb_from_zip`.
        writer
            .start_file(DEFAULT_WORKBOOK_PART, options.clone())
            .unwrap();
        writer.write_all(&[]).unwrap();

        writer
            .start_file(DEFAULT_WORKBOOK_RELS_PART, options.clone())
            .unwrap();
        writer.write_all(&[]).unwrap();

        for (name, bytes) in extra_parts {
            writer.start_file(*name, options.clone()).unwrap();
            writer.write_all(bytes).unwrap();
        }

        writer.finish().unwrap().into_inner()
    }

    fn build_xlsb_zip_with_workbook_parts(
        workbook_bin: &[u8],
        workbook_rels_bin: &[u8],
        extra_parts: &[(&str, &[u8])],
    ) -> Vec<u8> {
        let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
        let options = FileOptions::<()>::default().compression_method(CompressionMethod::Stored);

        writer
            .start_file(DEFAULT_WORKBOOK_PART, options.clone())
            .unwrap();
        writer.write_all(workbook_bin).unwrap();

        writer
            .start_file(DEFAULT_WORKBOOK_RELS_PART, options.clone())
            .unwrap();
        writer.write_all(workbook_rels_bin).unwrap();

        for (name, bytes) in extra_parts {
            writer.start_file(*name, options.clone()).unwrap();
            writer.write_all(bytes).unwrap();
        }

        writer.finish().unwrap().into_inner()
    }

    fn encode_utf16_string_payload(text: &str) -> Vec<u8> {
        let units: Vec<u16> = text.encode_utf16().collect();
        let mut out = Vec::with_capacity(4 + units.len() * 2);
        out.extend_from_slice(&(units.len() as u32).to_le_bytes());
        for unit in units {
            out.extend_from_slice(&unit.to_le_bytes());
        }
        out
    }

    fn build_workbook_bin_single_sheet(rel_id: &str, sheet_name: &str) -> Vec<u8> {
        let mut payload = Vec::new();
        payload.extend_from_slice(&0u32.to_le_bytes()); // state_flags (visible)
        payload.extend_from_slice(&1u32.to_le_bytes()); // sheet_id
        payload.extend_from_slice(&encode_utf16_string_payload(rel_id));
        payload.extend_from_slice(&encode_utf16_string_payload(sheet_name));

        let mut out = Vec::new();
        biff12_varint::write_record_id(&mut out, biff12::SHEET).unwrap();
        biff12_varint::write_record_len(&mut out, payload.len() as u32).unwrap();
        out.extend_from_slice(&payload);
        out
    }

    fn corrupt_zip_entry_uncompressed_size(zip_bytes: &mut [u8], entry_name: &str, new_size: u32) {
        // ZIP local file header signature `PK\x03\x04`.
        const LOCAL_SIG: [u8; 4] = [0x50, 0x4B, 0x03, 0x04];
        // ZIP central directory header signature `PK\x01\x02`.
        const CENTRAL_SIG: [u8; 4] = [0x50, 0x4B, 0x01, 0x02];

        let entry_name = entry_name.as_bytes();

        let mut patched_local = false;
        for i in 0..zip_bytes.len().saturating_sub(4) {
            if zip_bytes.get(i..i + 4) != Some(&LOCAL_SIG) {
                continue;
            }
            if i + 30 > zip_bytes.len() {
                continue;
            }
            let name_len =
                u16::from_le_bytes(zip_bytes[i + 26..i + 28].try_into().unwrap()) as usize;
            let extra_len =
                u16::from_le_bytes(zip_bytes[i + 28..i + 30].try_into().unwrap()) as usize;
            let name_start = i + 30;
            let name_end = name_start.saturating_add(name_len);
            let extra_end = name_end.saturating_add(extra_len);
            if extra_end > zip_bytes.len() {
                continue;
            }
            if zip_bytes.get(name_start..name_end) != Some(entry_name) {
                continue;
            }
            // Uncompressed size is at offset 22..26 in the local header.
            zip_bytes[i + 22..i + 26].copy_from_slice(&new_size.to_le_bytes());
            patched_local = true;
            break;
        }

        let mut patched_central = false;
        for i in 0..zip_bytes.len().saturating_sub(4) {
            if zip_bytes.get(i..i + 4) != Some(&CENTRAL_SIG) {
                continue;
            }
            if i + 46 > zip_bytes.len() {
                continue;
            }
            let name_len =
                u16::from_le_bytes(zip_bytes[i + 28..i + 30].try_into().unwrap()) as usize;
            let extra_len =
                u16::from_le_bytes(zip_bytes[i + 30..i + 32].try_into().unwrap()) as usize;
            let comment_len =
                u16::from_le_bytes(zip_bytes[i + 32..i + 34].try_into().unwrap()) as usize;
            let name_start = i + 46;
            let name_end = name_start.saturating_add(name_len);
            let extra_end = name_end.saturating_add(extra_len);
            let comment_end = extra_end.saturating_add(comment_len);
            if comment_end > zip_bytes.len() {
                continue;
            }
            if zip_bytes.get(name_start..name_end) != Some(entry_name) {
                continue;
            }
            // Uncompressed size is at offset 24..28 in the central directory header.
            zip_bytes[i + 24..i + 28].copy_from_slice(&new_size.to_le_bytes());
            patched_central = true;
            break;
        }

        assert!(patched_local, "did not find local header for entry: {entry_name:?}");
        assert!(
            patched_central,
            "did not find central directory header for entry: {entry_name:?}"
        );
    }

    #[test]
    fn rejects_oversized_unknown_zip_part() {
        let _lock = ENV_LOCK.lock().unwrap();
        let _max_part = EnvVarGuard::set(ENV_MAX_XLSB_ZIP_PART_BYTES, "10");
        let _max_preserved = EnvVarGuard::set(ENV_MAX_XLSB_PRESERVED_TOTAL_BYTES, "1024");

        let oversized = [0u8; 11];
        let bytes = build_minimal_xlsb_zip(&[("xl/unknown.bin", &oversized)]);

        let options = OpenOptions {
            preserve_unknown_parts: true,
            preserve_parsed_parts: false,
            preserve_worksheets: false,
            decode_formulas: false,
        };

        let err = XlsbWorkbook::open_from_vec_with_options(bytes, options)
            .err()
            .expect("expected oversized part error");

        match err {
            ParseError::PartTooLarge { part, size, max } => {
                assert_eq!(part, "xl/unknown.bin");
                assert_eq!(size, 11);
                assert_eq!(max, 10);
            }
            other => panic!("unexpected error: {other}"),
        }
    }

    #[test]
    fn rejects_oversized_workbook_part_when_zip_metadata_lies() {
        let _lock = ENV_LOCK.lock().unwrap();
        let _max_part = EnvVarGuard::set(ENV_MAX_XLSB_ZIP_PART_BYTES, "10");
        let _max_preserved = EnvVarGuard::set(ENV_MAX_XLSB_PRESERVED_TOTAL_BYTES, "1024");

        let workbook_bin = [0u8; 11];
        let mut bytes = build_xlsb_zip_with_workbook_parts(&workbook_bin, &[], &[]);

        // Corrupt the ZIP metadata so `ZipFile::size()` returns a value under the limit, while the
        // actual uncompressed entry payload is larger.
        corrupt_zip_entry_uncompressed_size(&mut bytes, DEFAULT_WORKBOOK_PART, 1);

        let options = OpenOptions {
            preserve_unknown_parts: false,
            preserve_parsed_parts: false,
            preserve_worksheets: false,
            decode_formulas: false,
        };

        let err = XlsbWorkbook::open_from_vec_with_options(bytes, options)
            .err()
            .expect("expected oversized part error");

        match err {
            ParseError::PartTooLarge { part, size, max } => {
                assert_eq!(part, DEFAULT_WORKBOOK_PART);
                assert_eq!(size, 11);
                assert_eq!(max, 10);
            }
            other => panic!("unexpected error: {other}"),
        }
    }

    #[test]
    fn rejects_oversized_shared_strings_part_when_zip_metadata_lies() {
        let _lock = ENV_LOCK.lock().unwrap();
        let _max_part = EnvVarGuard::set(ENV_MAX_XLSB_ZIP_PART_BYTES, "10");
        let _max_preserved = EnvVarGuard::set(ENV_MAX_XLSB_PRESERVED_TOTAL_BYTES, "1024");

        let workbook_bin = [];
        let shared_strings = [0u8; 11];
        let mut bytes = build_xlsb_zip_with_workbook_parts(
            &workbook_bin,
            &[],
            &[(DEFAULT_SHARED_STRINGS_PART, &shared_strings)],
        );

        // Corrupt ZIP metadata for `xl/sharedStrings.bin` to bypass the size() precheck.
        corrupt_zip_entry_uncompressed_size(&mut bytes, DEFAULT_SHARED_STRINGS_PART, 1);

        let options = OpenOptions {
            preserve_unknown_parts: false,
            preserve_parsed_parts: false,
            preserve_worksheets: false,
            decode_formulas: false,
        };

        let err = XlsbWorkbook::open_from_vec_with_options(bytes, options)
            .err()
            .expect("expected oversized part error");

        match err {
            ParseError::PartTooLarge { part, size, max } => {
                assert_eq!(part, DEFAULT_SHARED_STRINGS_PART);
                assert_eq!(size, 11);
                assert_eq!(max, 10);
            }
            other => panic!("unexpected error: {other}"),
        }
    }

    #[test]
    fn save_as_rejects_oversized_part_when_stream_copying() {
        let _lock = ENV_LOCK.lock().unwrap();
        let _max_part = EnvVarGuard::set(ENV_MAX_XLSB_ZIP_PART_BYTES, "10");

        let oversized = [0u8; 11];
        let bytes = build_minimal_xlsb_zip(&[("xl/unknown.bin", &oversized)]);

        let options = OpenOptions {
            preserve_unknown_parts: false,
            preserve_parsed_parts: false,
            preserve_worksheets: false,
            decode_formulas: false,
        };

        let wb = XlsbWorkbook::open_from_vec_with_options(bytes, options).expect("open workbook");
        let err = wb
            .save_as_to_bytes()
            .err()
            .expect("expected oversized part error during save");

        match err {
            ParseError::PartTooLarge { part, size, max } => {
                assert_eq!(part, "xl/unknown.bin");
                assert_eq!(size, 11);
                assert_eq!(max, 10);
            }
            other => panic!("unexpected error: {other}"),
        }
    }

    #[test]
    fn read_sheet_rejects_oversized_worksheet_part() {
        let _lock = ENV_LOCK.lock().unwrap();
        let _max_part = EnvVarGuard::set(ENV_MAX_XLSB_ZIP_PART_BYTES, "80");

        let workbook_bin = build_workbook_bin_single_sheet("rId1", "Sheet1");
        let workbook_rels = r#"<Relationship Id="rId1" Target="worksheets/sheet1.bin"/>"#;
        let oversized_sheet = vec![0u8; 81];
        let bytes = build_xlsb_zip_with_workbook_parts(
            &workbook_bin,
            workbook_rels.as_bytes(),
            &[("xl/worksheets/sheet1.bin", oversized_sheet.as_slice())],
        );

        let options = OpenOptions {
            preserve_unknown_parts: false,
            preserve_parsed_parts: false,
            preserve_worksheets: false,
            decode_formulas: false,
        };

        let wb = XlsbWorkbook::open_from_vec_with_options(bytes, options).expect("open workbook");
        let err = wb
            .read_sheet(0)
            .err()
            .expect("expected oversized worksheet part error");

        match err {
            ParseError::PartTooLarge { part, size, max } => {
                assert_eq!(part, "xl/worksheets/sheet1.bin");
                assert_eq!(size, 81);
                assert_eq!(max, 80);
            }
            other => panic!("unexpected error: {other}"),
        }
    }

    #[test]
    fn for_each_cell_rejects_oversized_worksheet_part() {
        let _lock = ENV_LOCK.lock().unwrap();
        let _max_part = EnvVarGuard::set(ENV_MAX_XLSB_ZIP_PART_BYTES, "80");

        let workbook_bin = build_workbook_bin_single_sheet("rId1", "Sheet1");
        let workbook_rels = r#"<Relationship Id="rId1" Target="worksheets/sheet1.bin"/>"#;
        let oversized_sheet = vec![0u8; 81];
        let bytes = build_xlsb_zip_with_workbook_parts(
            &workbook_bin,
            workbook_rels.as_bytes(),
            &[("xl/worksheets/sheet1.bin", oversized_sheet.as_slice())],
        );

        let options = OpenOptions {
            preserve_unknown_parts: false,
            preserve_parsed_parts: false,
            preserve_worksheets: false,
            decode_formulas: false,
        };

        let wb = XlsbWorkbook::open_from_vec_with_options(bytes, options).expect("open workbook");
        let err = wb
            .for_each_cell_control_flow(0, |_cell| ControlFlow::Continue(()))
            .err()
            .expect("expected oversized worksheet part error");

        match err {
            ParseError::PartTooLarge { part, size, max } => {
                assert_eq!(part, "xl/worksheets/sheet1.bin");
                assert_eq!(size, 81);
                assert_eq!(max, 80);
            }
            other => panic!("unexpected error: {other}"),
        }
    }

    #[test]
    fn read_sheet_rejects_oversized_worksheet_part_when_zip_metadata_lies() {
        let _lock = ENV_LOCK.lock().unwrap();
        let _max_part = EnvVarGuard::set(ENV_MAX_XLSB_ZIP_PART_BYTES, "34");

        let workbook_bin = build_workbook_bin_single_sheet("r", "S");
        let workbook_rels = r#"<Relationship Id="r" Target="s"/>"#;

        // 16 empty records (2 bytes each) + one 1-byte payload record (3 bytes) = 35 bytes.
        let mut oversized_sheet = Vec::new();
        oversized_sheet.extend(std::iter::repeat([0u8, 0u8]).take(16).flatten());
        oversized_sheet.extend_from_slice(&[0u8, 1u8, 0u8]);
        assert_eq!(oversized_sheet.len(), 35);

        let mut bytes = build_xlsb_zip_with_workbook_parts(
            &workbook_bin,
            workbook_rels.as_bytes(),
            &[("xl/s", oversized_sheet.as_slice())],
        );

        // Corrupt ZIP metadata for the worksheet part so `ZipFile::size()` is under the limit,
        // while the actual uncompressed payload exceeds it.
        corrupt_zip_entry_uncompressed_size(&mut bytes, "xl/s", 1);

        let options = OpenOptions {
            preserve_unknown_parts: false,
            preserve_parsed_parts: false,
            preserve_worksheets: false,
            decode_formulas: false,
        };
        let wb = XlsbWorkbook::open_from_vec_with_options(bytes, options).expect("open workbook");

        let err = wb
            .read_sheet(0)
            .err()
            .expect("expected oversized worksheet part error");

        match err {
            ParseError::PartTooLarge { part, size, max } => {
                assert_eq!(part, "xl/s");
                assert_eq!(size, 35);
                assert_eq!(max, 34);
            }
            other => panic!("unexpected error: {other}"),
        }
    }

    #[test]
    fn for_each_cell_rejects_oversized_worksheet_part_when_zip_metadata_lies() {
        let _lock = ENV_LOCK.lock().unwrap();
        let _max_part = EnvVarGuard::set(ENV_MAX_XLSB_ZIP_PART_BYTES, "34");

        let workbook_bin = build_workbook_bin_single_sheet("r", "S");
        let workbook_rels = r#"<Relationship Id="r" Target="s"/>"#;

        // 16 empty records (2 bytes each) + one 1-byte payload record (3 bytes) = 35 bytes.
        let mut oversized_sheet = Vec::new();
        oversized_sheet.extend(std::iter::repeat([0u8, 0u8]).take(16).flatten());
        oversized_sheet.extend_from_slice(&[0u8, 1u8, 0u8]);
        assert_eq!(oversized_sheet.len(), 35);

        let mut bytes = build_xlsb_zip_with_workbook_parts(
            &workbook_bin,
            workbook_rels.as_bytes(),
            &[("xl/s", oversized_sheet.as_slice())],
        );
        corrupt_zip_entry_uncompressed_size(&mut bytes, "xl/s", 1);

        let options = OpenOptions {
            preserve_unknown_parts: false,
            preserve_parsed_parts: false,
            preserve_worksheets: false,
            decode_formulas: false,
        };
        let wb = XlsbWorkbook::open_from_vec_with_options(bytes, options).expect("open workbook");

        let err = wb
            .for_each_cell_control_flow(0, |_cell| ControlFlow::Continue(()))
            .err()
            .expect("expected oversized worksheet part error");

        match err {
            ParseError::PartTooLarge { part, size, max } => {
                assert_eq!(part, "xl/s");
                assert_eq!(size, 35);
                assert_eq!(max, 34);
            }
            other => panic!("unexpected error: {other}"),
        }
    }

    #[test]
    fn enforces_total_preserved_parts_budget() {
        let _lock = ENV_LOCK.lock().unwrap();
        let _max_part = EnvVarGuard::set(ENV_MAX_XLSB_ZIP_PART_BYTES, "100");
        let _max_preserved = EnvVarGuard::set(ENV_MAX_XLSB_PRESERVED_TOTAL_BYTES, "50");

        let part_a = [0u8; 20];
        let part_b = [0u8; 20];
        let part_c = [0u8; 20];
        let bytes = build_minimal_xlsb_zip(&[
            ("xl/unk_a.bin", &part_a),
            ("xl/unk_b.bin", &part_b),
            ("xl/unk_c.bin", &part_c),
        ]);

        let options = OpenOptions {
            preserve_unknown_parts: true,
            preserve_parsed_parts: false,
            preserve_worksheets: false,
            decode_formulas: false,
        };

        let err = XlsbWorkbook::open_from_vec_with_options(bytes, options)
            .err()
            .expect("expected preserved parts budget error");

        match err {
            ParseError::PreservedPartsTooLarge { total, max } => {
                assert!(total > max);
                assert_eq!(max, 50);
            }
            other => panic!("unexpected error: {other}"),
        }
    }

    #[test]
    fn rejects_oversized_package_when_opening_from_reader() {
        let _lock = ENV_LOCK.lock().unwrap();
        let _max_pkg = EnvVarGuard::set(ENV_MAX_XLSB_PACKAGE_BYTES, "20");

        struct PanicOnReadCursor {
            inner: Cursor<Vec<u8>>,
        }

        impl Read for PanicOnReadCursor {
            fn read(&mut self, _buf: &mut [u8]) -> io::Result<usize> {
                panic!("unexpected read: open_from_reader should fail fast based on stream length");
            }
        }

        impl Seek for PanicOnReadCursor {
            fn seek(&mut self, pos: SeekFrom) -> io::Result<u64> {
                self.inner.seek(pos)
            }
        }

        // This does not need to be a valid ZIP. The size cap should fail fast before any parsing,
        // without reading from the stream.
        let bytes = vec![0u8; 21];
        let reader = PanicOnReadCursor {
            inner: Cursor::new(bytes),
        };
        let err = XlsbWorkbook::open_from_reader_with_options(reader, OpenOptions::default())
            .err()
            .expect("expected package too large error");

        match err {
            ParseError::Io(io_err) => {
                assert_eq!(io_err.kind(), std::io::ErrorKind::InvalidData);
                assert!(
                    io_err.to_string().contains("XLSB package too large"),
                    "unexpected error message: {io_err}"
                );
            }
            other => panic!("unexpected error: {other}"),
        }
    }

    #[test]
    fn rejects_oversized_package_when_opening_from_bytes() {
        let _lock = ENV_LOCK.lock().unwrap();
        let _max_pkg = EnvVarGuard::set(ENV_MAX_XLSB_PACKAGE_BYTES, "20");

        // This does not need to be a valid ZIP. The size cap should fail fast before any parsing.
        let bytes = vec![0u8; 21];
        let err = XlsbWorkbook::open_from_bytes_with_options(&bytes, OpenOptions::default())
            .err()
            .expect("expected package too large error");

        match err {
            ParseError::Io(io_err) => {
                assert_eq!(io_err.kind(), std::io::ErrorKind::InvalidData);
                assert!(
                    io_err.to_string().contains("XLSB package too large"),
                    "unexpected error message: {io_err}"
                );
            }
            other => panic!("unexpected error: {other}"),
        }
    }

    #[test]
    fn rejects_oversized_ole_wrapper_when_opening_with_password() {
        let _lock = ENV_LOCK.lock().unwrap();
        let _max_pkg = EnvVarGuard::set(ENV_MAX_XLSB_PACKAGE_BYTES, "20");

        // Construct an OLE header so `open_with_password` takes the decrypt path, but keep the
        // bytes invalid (size check should trigger before decryption/parsing).
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&OLE_MAGIC);
        bytes.resize(21, 0);

        let mut tmp = NamedTempFile::new().expect("tempfile");
        tmp.write_all(&bytes).expect("write temp ole");
        tmp.flush().expect("flush temp ole");

        let err = XlsbWorkbook::open_with_password(tmp.path(), "password")
            .err()
            .expect("expected package too large error");

        match err {
            ParseError::Io(io_err) => {
                assert_eq!(io_err.kind(), std::io::ErrorKind::InvalidData);
                assert!(
                    io_err.to_string().contains("XLSB package too large"),
                    "unexpected error message: {io_err}"
                );
            }
            other => panic!("unexpected error: {other}"),
        }
    }

    #[test]
    fn rejects_too_many_zip_entries_when_opening_from_reader_without_buffering_whole_package() {
        let _lock = ENV_LOCK.lock().unwrap();
        let _max_entries = EnvVarGuard::set(ENV_MAX_XLSB_ZIP_ENTRIES, "4");

        // Make the package large enough that a full `read_to_end` would exceed our threshold, but
        // still small enough that `preflight_zip_entry_count` only needs to read the tail.
        let large_payload = vec![0u8; 200_000];
        let bytes = build_minimal_xlsb_zip(&[
            ("xl/large.bin", large_payload.as_slice()),
            ("xl/extra1.bin", b"1"),
            ("xl/extra2.bin", b"2"),
        ]);

        struct PanicAfterReadThreshold {
            inner: Cursor<Vec<u8>>,
            bytes_read: usize,
            max_bytes: usize,
        }

        impl Read for PanicAfterReadThreshold {
            fn read(&mut self, buf: &mut [u8]) -> io::Result<usize> {
                let n = self.inner.read(buf)?;
                self.bytes_read = self.bytes_read.saturating_add(n);
                if self.bytes_read > self.max_bytes {
                    panic!(
                        "unexpectedly buffered too much of the package: {} bytes (limit {})",
                        self.bytes_read, self.max_bytes
                    );
                }
                Ok(n)
            }
        }

        impl Seek for PanicAfterReadThreshold {
            fn seek(&mut self, pos: SeekFrom) -> io::Result<u64> {
                self.inner.seek(pos)
            }
        }

        let reader = PanicAfterReadThreshold {
            inner: Cursor::new(bytes),
            bytes_read: 0,
            // `preflight_zip_entry_count` reads at most ~65KiB from the tail. If `open_from_reader`
            // were to buffer the whole package before preflighting, it would exceed this limit.
            max_bytes: 100_000,
        };

        let err = XlsbWorkbook::open_from_reader_with_options(reader, OpenOptions::default())
            .err()
            .expect("expected too many zip entries error");

        match err {
            ParseError::TooManyZipEntries { count, max } => {
                assert_eq!(max, 4);
                assert!(count > max);
            }
            other => panic!("unexpected error: {other}"),
        }
    }
}

fn map_office_crypto_err(err: office_crypto::OfficeCryptoError) -> ParseError {
    match err {
        office_crypto::OfficeCryptoError::InvalidPassword => ParseError::InvalidPassword,
        other => ParseError::OfficeCrypto(other.to_string()),
    }
}

fn parse_relationships(xml_bytes: &[u8]) -> Result<HashMap<String, String>, ParseError> {
    let xml = String::from_utf8_lossy(xml_bytes);
    let mut reader = XmlReader::from_str(&xml);
    reader.trim_text(true);
    let mut buf = Vec::new();
    let mut out = HashMap::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e))
                if e.name().as_ref().ends_with(b"Relationship") =>
            {
                let mut id = None;
                let mut target = None;
                for attr in e.attributes().flatten() {
                    match attr.key.as_ref() {
                        b"Id" => id = Some(attr.decode_and_unescape_value(&reader)?.into_owned()),
                        b"Target" => {
                            target = Some(attr.decode_and_unescape_value(&reader)?.into_owned())
                        }
                        _ => {}
                    }
                }
                if let (Some(id), Some(target)) = (id, target) {
                    out.insert(id, target);
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(ParseError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }
    Ok(out)
}

fn resolve_workbook_part_name<R: Read + Seek>(
    zip: &ZipArchive<R>,
    workbook_rels_xml: &[u8],
    content_types_xml: Option<&[u8]>,
    relationship_type: &str,
    content_type: Option<&str>,
    default_part: &str,
) -> Result<Option<String>, ParseError> {
    if let Some(target) = relationship_target_by_type(workbook_rels_xml, relationship_type)? {
        for candidate in workbook_target_candidates(&target) {
            if let Some(actual) = find_zip_entry_case_insensitive(zip, &candidate) {
                return Ok(Some(actual));
            }
        }
    }

    if let (Some(content_type), Some(content_types_xml)) = (content_type, content_types_xml) {
        for part_name in content_type_override_part_names(content_types_xml, content_type)? {
            let normalized = normalize_zip_part_name(&part_name);
            if let Some(actual) = find_zip_entry_case_insensitive(zip, &normalized) {
                return Ok(Some(actual));
            }
        }
    }

    Ok(find_zip_entry_case_insensitive(zip, default_part))
}

fn relationship_target_by_type(
    xml_bytes: &[u8],
    relationship_type: &str,
) -> Result<Option<String>, ParseError> {
    let mut reader = XmlReader::from_reader(std::io::BufReader::new(Cursor::new(xml_bytes)));
    reader.trim_text(true);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e))
                if e.name().as_ref().ends_with(b"Relationship") =>
            {
                if let Some(ty) = xml_attr_value(&e, &reader, b"Type")? {
                    if ty.eq_ignore_ascii_case(relationship_type) {
                        return Ok(xml_attr_value(&e, &reader, b"Target")?);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(ParseError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(None)
}

fn relationship_targets_by_type(
    xml_bytes: &[u8],
    relationship_type: &str,
) -> Result<Vec<String>, ParseError> {
    let mut reader = XmlReader::from_reader(std::io::BufReader::new(Cursor::new(xml_bytes)));
    reader.trim_text(true);
    let mut buf = Vec::new();
    let mut out = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e))
                if e.name().as_ref().ends_with(b"Relationship") =>
            {
                if let Some(ty) = xml_attr_value(&e, &reader, b"Type")? {
                    if ty.eq_ignore_ascii_case(relationship_type) {
                        if let Some(target) = xml_attr_value(&e, &reader, b"Target")? {
                            out.push(target);
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(ParseError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }
    Ok(out)
}

fn content_type_override_part_names(
    xml_bytes: &[u8],
    content_type: &str,
) -> Result<Vec<String>, ParseError> {
    let mut reader = XmlReader::from_reader(std::io::BufReader::new(Cursor::new(xml_bytes)));
    reader.trim_text(true);
    let mut buf = Vec::new();
    let mut out = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e))
                if e.name().as_ref().ends_with(b"Override") =>
            {
                let Some(ty) = xml_attr_value(&e, &reader, b"ContentType")? else {
                    continue;
                };
                if !ty.eq_ignore_ascii_case(content_type) {
                    continue;
                }
                if let Some(part) = xml_attr_value(&e, &reader, b"PartName")? {
                    out.push(part);
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(ParseError::Xml(e)),
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

fn root_target_candidates(target: &str) -> Vec<String> {
    let target = normalize_zip_part_name(target);
    let mut candidates = vec![target.clone()];
    if !target.to_ascii_lowercase().starts_with("xl/") {
        candidates.push(normalize_zip_part_name(&format!("xl/{target}")));
    }
    candidates
}

fn workbook_target_candidates(target: &str) -> Vec<String> {
    // workbook.bin is stored under `xl/`, so relationship targets are typically relative to `xl/`.
    // Some writers use absolute targets with a leading `/`.
    let target = target.replace('\\', "/");
    let target = target.trim_start_matches('/');

    let mut candidates = Vec::new();
    candidates.push(normalize_zip_part_name(&format!("xl/{target}")));
    candidates.push(normalize_zip_part_name(target));
    candidates
}

fn resolve_part_name_from_relationship(source_part: &str, target: &str) -> String {
    let source_part = normalize_zip_part_name(source_part);

    // Relationship targets use `/` separators, but some writers use Windows-style `\`.
    let mut target = target.replace('\\', "/");
    // Some writers use absolute targets with a leading `/`.
    let target_is_absolute = target.starts_with('/');
    target = target.trim_start_matches('/').to_string();
    // Relationship targets are URIs; internal targets may include a fragment (e.g. `foo.bin#bar`).
    // OPC part names do not include fragments, so strip them before resolving.
    let target = target
        .split_once('#')
        .map(|(base, _)| base)
        .unwrap_or(target.as_str());
    if target.is_empty() {
        // A target of just `#fragment` refers to the source part itself.
        return source_part;
    }

    if target_is_absolute {
        return normalize_zip_part_name(&target);
    }

    let base_dir = source_part
        .rsplit_once('/')
        .map(|(dir, _)| dir)
        .unwrap_or("");
    if base_dir.is_empty() {
        normalize_zip_part_name(&target)
    } else {
        normalize_zip_part_name(&format!("{base_dir}/{target}"))
    }
}

fn rels_part_name_for_part(part_name: &str) -> String {
    let part_name = normalize_zip_part_name(part_name);
    match part_name.rsplit_once('/') {
        Some((dir, file)) => format!("{dir}/_rels/{file}.rels"),
        None => format!("_rels/{part_name}.rels"),
    }
}

fn normalize_zip_part_name(part_name: &str) -> String {
    // Relationship targets are URIs; internal targets may include a fragment (e.g. `foo.bin#bar`).
    // OPC part names do not include fragments, so strip them before normalizing.
    let part_name = part_name
        .split_once('#')
        .map(|(base, _)| base)
        .unwrap_or(part_name);
    let part_name = part_name.replace('\\', "/");
    let part_name = part_name.trim_start_matches('/');
    let mut out: Vec<&str> = Vec::new();
    for seg in part_name.split('/') {
        match seg {
            "" | "." => {}
            ".." => {
                out.pop();
            }
            _ => out.push(seg),
        }
    }
    out.join("/")
}

#[cfg(test)]
mod relationship_target_tests {
    use super::resolve_part_name_from_relationship;

    #[test]
    fn strips_uri_fragments_from_relationship_targets() {
        assert_eq!(
            resolve_part_name_from_relationship(
                "xl/worksheets/sheet1.bin",
                "../tables/table1.xml#frag"
            ),
            "xl/tables/table1.xml"
        );
    }

    #[test]
    fn fragment_only_relationship_targets_resolve_to_source_part() {
        assert_eq!(
            resolve_part_name_from_relationship("xl/worksheets/sheet1.bin", "#frag"),
            "xl/worksheets/sheet1.bin"
        );
    }
}

fn find_zip_entry_case_insensitive<R: Read + Seek>(
    zip: &ZipArchive<R>,
    name: &str,
) -> Option<String> {
    let target = name.trim_start_matches('/').replace('\\', "/");

    for candidate in zip.file_names() {
        let mut normalized = candidate.trim_start_matches('/');
        let replaced;
        if normalized.contains('\\') {
            replaced = normalized.replace('\\', "/");
            normalized = &replaced;
        }
        if normalized.eq_ignore_ascii_case(&target) {
            return Some(candidate.to_string());
        }
    }

    None
}

fn read_zip_entry<R: Read + Seek>(
    zip: &mut ZipArchive<R>,
    name: &str,
) -> Result<Option<Vec<u8>>, ParseError> {
    let max = max_xlsb_zip_part_bytes();

    let try_read = |zip: &mut ZipArchive<R>, entry_name: &str| -> Result<Vec<u8>, ParseError> {
        let entry = zip.by_name(entry_name)?;
        let size = entry.size();
        if size > max {
            return Err(ParseError::PartTooLarge {
                part: entry_name.to_string(),
                size,
                max,
            });
        }

        // Don't trust the ZIP metadata for preallocation; use a small fixed buffer to keep
        // allocations modest even if `ZipFile::size()` is forged.
        let mut bytes = Vec::with_capacity(ZIP_ENTRY_READ_PREALLOC_BYTES);

        // Guard against ZIP metadata lies (or unknown sizes) by enforcing a hard cap on bytes read.
        entry.take(max.saturating_add(1)).read_to_end(&mut bytes)?;
        if bytes.len() as u64 > max {
            return Err(ParseError::PartTooLarge {
                part: entry_name.to_string(),
                size: bytes.len() as u64,
                max,
            });
        }
        Ok(bytes)
    };

    match try_read(zip, name) {
        Ok(bytes) => Ok(Some(bytes)),
        Err(ParseError::Zip(zip::result::ZipError::FileNotFound)) => {
            let Some(actual) = find_zip_entry_case_insensitive(zip, name) else {
                return Ok(None);
            };
            Ok(Some(try_read(zip, &actual)?))
        }
        Err(err) => Err(err),
    }
}

fn read_zip_entry_required<R: Read + Seek>(
    zip: &mut ZipArchive<R>,
    name: &str,
) -> Result<Vec<u8>, ParseError> {
    read_zip_entry(zip, name)?.ok_or_else(|| ParseError::Zip(zip::result::ZipError::FileNotFound))
}

fn insert_preserved_part(
    preserved_parts: &mut HashMap<String, Vec<u8>>,
    preserved_total_bytes: &mut u64,
    name: String,
    bytes: Vec<u8>,
) -> Result<(), ParseError> {
    let max_total = max_xlsb_preserved_total_bytes();
    let new_len = bytes.len() as u64;
    let old_len = preserved_parts
        .get(&name)
        .map(|v| v.len() as u64)
        .unwrap_or(0);

    let next_total = preserved_total_bytes
        .saturating_sub(old_len)
        .saturating_add(new_len);
    if next_total > max_total {
        return Err(ParseError::PreservedPartsTooLarge {
            total: next_total,
            max: max_total,
        });
    }

    preserved_parts.insert(name, bytes);
    *preserved_total_bytes = next_total;
    Ok(())
}

fn preserve_part<R: Read + Seek>(
    zip: &mut ZipArchive<R>,
    preserved: &mut HashMap<String, Vec<u8>>,
    preserved_total_bytes: &mut u64,
    name: &str,
) -> Result<(), ParseError> {
    if let Some(bytes) = read_zip_entry(zip, name)? {
        insert_preserved_part(preserved, preserved_total_bytes, name.to_string(), bytes)?;
    }
    Ok(())
}

fn load_table_definitions<R: Read + Seek>(
    zip: &mut ZipArchive<R>,
    ctx: &mut WorkbookContext,
    sheets: &[SheetMeta],
) -> Result<(), ParseError> {
    // Best-effort: tables are not required for basic parsing, but we use them to reconstruct
    // structured reference (Excel table) formulas with the correct display names.
    let table_parts: Vec<String> = zip
        .file_names()
        .filter(|name| {
            let normalized = normalize_zip_part_name(name).to_ascii_lowercase();
            normalized.starts_with("xl/tables/") && normalized.ends_with(".xml")
        })
        .map(str::to_string)
        .collect();

    for part in table_parts {
        let Some(bytes) = read_zip_entry(zip, &part)? else {
            continue;
        };
        let Some(info) = parse_table_xml_best_effort(&bytes) else {
            continue;
        };
        ctx.add_table(info.id, info.name);
        for (col_id, col_name) in info.columns {
            ctx.add_table_column(info.id, col_id, col_name);
        }
    }

    // Sheet association + range bounds (used to encode table-less structured refs like `[@Col]`).
    //
    // Table parts (`xl/tables/tableN.xml`) do not embed their owning sheet; that association is
    // defined by the worksheet relationships (`xl/worksheets/_rels/sheetN.bin.rels`).
    for sheet in sheets {
        let rels_candidate = rels_part_name_for_part(&sheet.part_path);
        let Some(rels_part) = find_zip_entry_case_insensitive(zip, &rels_candidate) else {
            continue;
        };

        let Some(rels_bytes) = read_zip_entry(zip, &rels_part)? else {
            continue;
        };

        let targets = match relationship_targets_by_type(&rels_bytes, TABLE_REL_TYPE) {
            Ok(v) => v,
            // Best-effort: malformed relationships should not block workbook open.
            Err(_) => continue,
        };

        for target in targets {
            let resolved = resolve_part_name_from_relationship(&sheet.part_path, &target);
            let Some(table_part) = find_zip_entry_case_insensitive(zip, &resolved) else {
                continue;
            };

            let Some(bytes) = read_zip_entry(zip, &table_part)? else {
                continue;
            };
            let Some(info) = parse_table_xml_best_effort(&bytes) else {
                continue;
            };

            ctx.add_table(info.id, info.name);
            for (col_id, col_name) in info.columns {
                ctx.add_table_column(info.id, col_id, col_name);
            }

            if let Some(a1_ref) = info.ref_a1 {
                if let Some((r1, c1, r2, c2)) = parse_a1_range_bounds(&a1_ref) {
                    ctx.add_table_range(info.id, sheet.name.clone(), r1, c1, r2, c2);
                }
            }
        }
    }

    Ok(())
}

#[derive(Debug)]
struct ParsedTableXml {
    id: u32,
    name: String,
    columns: Vec<(u32, String)>,
    ref_a1: Option<String>,
}

fn parse_table_xml_best_effort(xml_bytes: &[u8]) -> Option<ParsedTableXml> {
    let mut reader = XmlReader::from_reader(std::io::BufReader::new(Cursor::new(xml_bytes)));
    reader.trim_text(true);
    let mut buf = Vec::new();

    fn attr_value<B: std::io::BufRead>(
        e: &quick_xml::events::BytesStart<'_>,
        reader: &XmlReader<B>,
        key: &[u8],
    ) -> Option<String> {
        for attr in e.attributes().flatten() {
            if attr.key.as_ref() == key {
                return attr
                    .decode_and_unescape_value(reader)
                    .ok()
                    .map(|v| v.into_owned());
            }
        }
        None
    }

    let mut table_id: Option<u32> = None;
    let mut table_name: Option<String> = None;
    let mut ref_a1: Option<String> = None;
    let mut columns: Vec<(u32, String)> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) if e.name().as_ref().ends_with(b"table") => {
                if table_id.is_some() {
                    // Already parsed a table root.
                    buf.clear();
                    continue;
                }

                let id =
                    attr_value(&e, &reader, b"id").or_else(|| attr_value(&e, &reader, b"Id"))?;
                table_id = id.parse::<u32>().ok();

                // Excel stores both `name` and `displayName`. Formulas typically use displayName.
                let display = attr_value(&e, &reader, b"displayName")
                    .or_else(|| attr_value(&e, &reader, b"DisplayName"));
                let name =
                    attr_value(&e, &reader, b"name").or_else(|| attr_value(&e, &reader, b"Name"));
                table_name = display.or(name);

                // A1-style bounding box for the table.
                ref_a1 =
                    attr_value(&e, &reader, b"ref").or_else(|| attr_value(&e, &reader, b"Ref"));
            }
            Ok(Event::Start(e)) | Ok(Event::Empty(e))
                if e.name().as_ref().ends_with(b"tableColumn") =>
            {
                if table_id.is_none() {
                    buf.clear();
                    continue;
                };

                let Some(id) = attr_value(&e, &reader, b"id")
                    .or_else(|| attr_value(&e, &reader, b"Id"))
                    .and_then(|s| s.parse::<u32>().ok())
                else {
                    buf.clear();
                    continue;
                };
                let Some(name) =
                    attr_value(&e, &reader, b"name").or_else(|| attr_value(&e, &reader, b"Name"))
                else {
                    buf.clear();
                    continue;
                };
                columns.push((id, name));
            }
            Ok(Event::Eof) => break,
            Err(_) => return None,
            _ => {}
        }
        buf.clear();
    }

    Some(ParsedTableXml {
        id: table_id?,
        name: table_name?,
        columns,
        ref_a1,
    })
}

fn parse_a1_range_bounds(a1: &str) -> Option<(u32, u32, u32, u32)> {
    let a1 = a1.trim();
    if a1.is_empty() {
        return None;
    }

    // Be tolerant of unexpected sheet-qualified refs (e.g. `Sheet1!A1:B2`) by stripping the
    // prefix. Table `ref` attributes are typically unqualified.
    let a1 = a1.rsplit_once('!').map(|(_, tail)| tail).unwrap_or(a1);

    // Strip absolute markers.
    let a1 = a1.replace('$', "");

    let mut parts = a1.split(':');
    let start = parts.next()?.trim();
    let end = parts.next().unwrap_or(start).trim();
    if parts.next().is_some() {
        // Multiple ':' separators.
        return None;
    }

    let (r1, c1) = parse_a1_cell_ref(start)?;
    let (r2, c2) = parse_a1_cell_ref(end)?;

    let (min_row, max_row) = if r1 <= r2 { (r1, r2) } else { (r2, r1) };
    let (min_col, max_col) = if c1 <= c2 { (c1, c2) } else { (c2, c1) };
    Some((min_row, min_col, max_row, max_col))
}

fn parse_a1_cell_ref(cell: &str) -> Option<(u32, u32)> {
    let cell = cell.trim();
    if cell.is_empty() {
        return None;
    }

    let bytes = cell.as_bytes();
    let mut i = 0usize;
    while i < bytes.len() && bytes[i].is_ascii_alphabetic() {
        i += 1;
    }
    if i == 0 {
        return None;
    }
    let col_label = &cell[..i];
    let row_str = cell.get(i..)?.trim();
    if row_str.is_empty() {
        return None;
    }
    if !row_str.bytes().all(|b| b.is_ascii_digit()) {
        return None;
    }

    let col = a1_col_label_to_index(col_label)?;
    let row1: u32 = row_str.parse().ok()?;
    if row1 == 0 {
        return None;
    }
    let row = row1 - 1;

    // Clamp to Excel's grid limits.
    if col > 16_383 || row > 1_048_575 {
        return None;
    }

    Some((row, col))
}

fn a1_col_label_to_index(label: &str) -> Option<u32> {
    let mut col: u32 = 0;
    for ch in label.chars() {
        if !ch.is_ascii_alphabetic() {
            return None;
        }
        let upper = ch.to_ascii_uppercase() as u32;
        col = col.checked_mul(26)?;
        col = col.checked_add(upper.checked_sub('A' as u32)?.checked_add(1)?)?;
    }
    if col == 0 {
        return None;
    }
    Some(col - 1)
}

#[cfg(test)]
mod a1_range_tests {
    use super::parse_a1_range_bounds;

    #[test]
    fn parses_simple_a1_range() {
        assert_eq!(parse_a1_range_bounds("A1:B10"), Some((0, 0, 9, 1)));
    }

    #[test]
    fn parses_absolute_a1_range() {
        assert_eq!(parse_a1_range_bounds("$A$1:$B$10"), Some((0, 0, 9, 1)));
    }

    #[test]
    fn parses_single_cell_ref_as_range() {
        assert_eq!(parse_a1_range_bounds("C3"), Some((2, 2, 2, 2)));
    }

    #[test]
    fn parses_sheet_qualified_a1_range() {
        assert_eq!(parse_a1_range_bounds("Sheet1!A1:B2"), Some((0, 0, 1, 1)));
    }

    #[test]
    fn rejects_invalid_a1_range() {
        assert_eq!(parse_a1_range_bounds(""), None);
        assert_eq!(parse_a1_range_bounds("A0:B1"), None);
        assert_eq!(parse_a1_range_bounds("A1:B"), None);
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[derive(Debug)]
struct CellRecordInfo {
    id: u32,
    payload: Vec<u8>,
}

#[cfg(not(target_arch = "wasm32"))]
fn sheet_cell_records(
    sheet_bin: &[u8],
    targets: &HashSet<(u32, u32)>,
) -> Result<HashMap<(u32, u32), CellRecordInfo>, ParseError> {
    let mut cursor = Cursor::new(sheet_bin);
    let mut in_sheet_data = false;
    let mut current_row = 0u32;
    let mut found: HashMap<(u32, u32), CellRecordInfo> = HashMap::new();

    loop {
        let Some(id) = biff12_varint::read_record_id(&mut cursor)? else {
            break;
        };
        let Some(len) = biff12_varint::read_record_len(&mut cursor)? else {
            return Err(ParseError::UnexpectedEof);
        };
        let len = len as usize;

        let payload_start = cursor.position() as usize;
        let payload_end = payload_start
            .checked_add(len)
            .filter(|&end| end <= sheet_bin.len())
            .ok_or(ParseError::UnexpectedEof)?;
        let payload = &sheet_bin[payload_start..payload_end];
        cursor.set_position(payload_end as u64);

        match id {
            biff12::SHEETDATA => in_sheet_data = true,
            biff12::SHEETDATA_END => in_sheet_data = false,
            biff12::ROW if in_sheet_data => {
                if payload.len() >= 4 {
                    current_row = u32::from_le_bytes(payload[0..4].try_into().unwrap());
                }
            }
            _ if in_sheet_data => {
                if payload.len() >= 4 {
                    let col = u32::from_le_bytes(payload[0..4].try_into().unwrap());
                    let coord = (current_row, col);
                    if targets.contains(&coord) {
                        found.insert(
                            coord,
                            CellRecordInfo {
                                id,
                                payload: payload.to_vec(),
                            },
                        );
                        if found.len() == targets.len() {
                            break;
                        }
                    }
                }
            }
            _ => {}
        }
    }

    Ok(found)
}

#[cfg(not(target_arch = "wasm32"))]
fn sheet_cell_records_streaming<R: Read>(
    sheet_bin: &mut R,
    targets: &HashSet<(u32, u32)>,
) -> Result<HashMap<(u32, u32), CellRecordInfo>, ParseError> {
    const READ_CHUNK_BYTES: usize = 16 * 1024;
    let mut in_sheet_data = false;
    let mut current_row = 0u32;
    let mut found: HashMap<(u32, u32), CellRecordInfo> = HashMap::new();

    loop {
        let Some(id) = biff12_varint::read_record_id(sheet_bin)? else {
            break;
        };
        let Some(len) = biff12_varint::read_record_len(sheet_bin)? else {
            return Err(ParseError::UnexpectedEof);
        };
        let len = len as usize;

        match id {
            biff12::SHEETDATA => {
                in_sheet_data = true;
                skip_record_payload(sheet_bin, len)?;
            }
            biff12::SHEETDATA_END => {
                in_sheet_data = false;
                skip_record_payload(sheet_bin, len)?;
            }
            biff12::ROW if in_sheet_data => {
                if len >= 4 {
                    let mut buf = [0u8; 4];
                    sheet_bin.read_exact(&mut buf)?;
                    current_row = u32::from_le_bytes(buf);
                    skip_record_payload(sheet_bin, len - 4)?;
                } else {
                    skip_record_payload(sheet_bin, len)?;
                }
            }
            _ if in_sheet_data => {
                if len >= 4 {
                    let mut head = [0u8; 4];
                    sheet_bin.read_exact(&mut head)?;
                    let col = u32::from_le_bytes(head);
                    let coord = (current_row, col);
                    if targets.contains(&coord) {
                        // Record lengths are attacker-controlled; grow the buffer as we successfully
                        // read bytes rather than pre-reserving/zero-filling the full length.
                        let mut payload = Vec::new();
                        payload.extend_from_slice(&head);
                        if len > 4 {
                            let mut remaining = len - 4;
                            let mut buf = [0u8; READ_CHUNK_BYTES];
                            while remaining > 0 {
                                let chunk_len = buf.len().min(remaining);
                                sheet_bin.read_exact(&mut buf[..chunk_len])?;
                                payload.extend_from_slice(&buf[..chunk_len]);
                                remaining = remaining.saturating_sub(chunk_len);
                            }
                        }
                        found.insert(coord, CellRecordInfo { id, payload });
                        if found.len() == targets.len() {
                            break;
                        }
                        continue;
                    }
                    skip_record_payload(sheet_bin, len - 4)?;
                } else {
                    skip_record_payload(sheet_bin, len)?;
                }
            }
            _ => skip_record_payload(sheet_bin, len)?,
        }
    }

    Ok(found)
}

#[cfg(not(target_arch = "wasm32"))]
fn skip_record_payload<R: Read>(r: &mut R, mut len: usize) -> Result<(), ParseError> {
    let mut buf = [0u8; 16 * 1024];
    while len > 0 {
        let chunk_len = buf.len().min(len);
        r.read_exact(&mut buf[..chunk_len])?;
        len = len.saturating_sub(chunk_len);
    }
    Ok(())
}

#[cfg(not(target_arch = "wasm32"))]
fn is_formula_cell_record(id: u32) -> bool {
    matches!(
        id,
        biff12::FORMULA_FLOAT
            | biff12::FORMULA_STRING
            | biff12::FORMULA_BOOL
            | biff12::FORMULA_BOOLERR
    )
}

fn worksheets_edited<R: Read + Seek>(
    zip: &mut ZipArchive<R>,
    sheets: &[SheetMeta],
    overrides: &HashMap<String, Vec<u8>>,
) -> Result<bool, ParseError> {
    let worksheet_paths: HashSet<&str> = sheets.iter().map(|s| s.part_path.as_str()).collect();

    for (name, override_bytes) in overrides {
        if !worksheet_paths.contains(name.as_str()) {
            continue;
        }

        let Some(equal) = zip_entry_equals(zip, name, override_bytes)? else {
            // Treat missing original parts as edited; downstream the caller may be synthesizing a
            // sheet.
            return Ok(true);
        };
        if !equal {
            return Ok(true);
        }
    }

    Ok(false)
}

fn is_calc_chain_part_name(name: &str) -> bool {
    name.trim_start_matches('/')
        .eq_ignore_ascii_case("xl/calcChain.bin")
}

fn is_content_types_part_name(name: &str) -> bool {
    name.trim_start_matches('/')
        .eq_ignore_ascii_case("[Content_Types].xml")
}

fn patch_workbook_bin_full_calc_on_load(
    workbook_bin: &[u8],
) -> Result<Option<Vec<u8>>, ParseError> {
    let mut cursor = Cursor::new(workbook_bin);
    let mut out = Vec::with_capacity(workbook_bin.len());
    let mut changed = false;

    loop {
        let start = cursor.position() as usize;
        let Some(id) = biff12_varint::read_record_id(&mut cursor)? else {
            break;
        };
        let Some(len) = biff12_varint::read_record_len(&mut cursor)? else {
            return Err(ParseError::UnexpectedEof);
        };
        let len: usize = len as usize;
        let header_end = cursor.position() as usize;
        let payload_start = header_end;
        let payload_end = payload_start
            .checked_add(len)
            .filter(|&end| end <= workbook_bin.len())
            .ok_or(ParseError::UnexpectedEof)?;

        // Preserve the exact varint encoding for id/len.
        out.extend_from_slice(&workbook_bin[start..payload_start]);

        let payload = &workbook_bin[payload_start..payload_end];
        cursor.set_position(payload_end as u64);

        if id == biff12::CALC_PROP && payload.len() >= 6 {
            let mut patched = payload.to_vec();
            let flags_off = 4usize;
            let flags = u16::from_le_bytes([patched[flags_off], patched[flags_off + 1]]);
            let new_flags = flags | 0x0004;
            if new_flags != flags {
                patched[flags_off..flags_off + 2].copy_from_slice(&new_flags.to_le_bytes());
                changed = true;
            }
            out.extend_from_slice(&patched);
        } else {
            out.extend_from_slice(payload);
        }
    }

    Ok(changed.then_some(out))
}

fn zip_entry_equals<R: Read + Seek>(
    zip: &mut ZipArchive<R>,
    name: &str,
    expected: &[u8],
) -> Result<Option<bool>, ParseError> {
    let mut entry = match zip.by_name(name) {
        Ok(entry) => entry,
        Err(zip::result::ZipError::FileNotFound) => return Ok(None),
        Err(e) => return Err(e.into()),
    };

    let Ok(size) = usize::try_from(entry.size()) else {
        // An override can't match an entry whose uncompressed size doesn't fit in memory.
        return Ok(Some(false));
    };
    if size != expected.len() {
        return Ok(Some(false));
    }

    let mut buf = [0u8; 16 * 1024];
    let mut offset = 0usize;

    loop {
        let n = entry.read(&mut buf)?;
        if n == 0 {
            break;
        }

        let end = offset.checked_add(n).ok_or(ParseError::UnexpectedEof)?;
        if end > expected.len() {
            return Ok(Some(false));
        }

        if buf[..n] != expected[offset..end] {
            return Ok(Some(false));
        }

        offset = end;
    }

    Ok(Some(offset == expected.len()))
}

fn get_part_bytes<R: Read + Seek>(
    zip: &mut ZipArchive<R>,
    preserved_parts: &HashMap<String, Vec<u8>>,
    overrides: &HashMap<String, Vec<u8>>,
    name: &str,
) -> Result<Option<Vec<u8>>, ParseError> {
    if let Some(bytes) = overrides.get(name) {
        return Ok(Some(bytes.clone()));
    }
    if let Some(bytes) = preserved_parts.get(name) {
        return Ok(Some(bytes.clone()));
    }
    read_zip_entry(zip, name)
}

fn remove_calc_chain_from_content_types(xml_bytes: &[u8]) -> Result<Vec<u8>, ParseError> {
    let mut reader = XmlReader::from_reader(std::io::BufReader::new(Cursor::new(xml_bytes)));
    reader.trim_text(false);
    let mut writer = XmlWriter::new(Vec::new());
    let mut buf = Vec::new();
    let mut skip_depth = 0usize;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Eof) => break,
            Ok(event) => {
                if skip_depth > 0 {
                    match event {
                        Event::Start(_) => skip_depth += 1,
                        Event::End(_) => skip_depth = skip_depth.saturating_sub(1),
                        _ => {}
                    }
                } else if should_drop_content_type_event(&event, &reader)? {
                    if matches!(event, Event::Start(_)) {
                        skip_depth = 1;
                    }
                } else {
                    writer.write_event(event.into_owned())?;
                }
            }
            Err(e) => return Err(ParseError::Xml(e)),
        }
        buf.clear();
    }

    Ok(writer.into_inner())
}

fn should_drop_content_type_event<B: std::io::BufRead>(
    event: &Event<'_>,
    reader: &XmlReader<B>,
) -> Result<bool, ParseError> {
    let (Event::Start(e) | Event::Empty(e)) = event else {
        return Ok(false);
    };

    let qname = e.name();
    let name = qname.as_ref();
    if name.ends_with(b"Override") {
        if let Some(part) = xml_attr_value(e, reader, b"PartName")? {
            if part.eq_ignore_ascii_case("/xl/calcChain.bin")
                || part.eq_ignore_ascii_case("xl/calcChain.bin")
            {
                return Ok(true);
            }
        }
    } else if name.ends_with(b"Default") {
        // Some generators might (incorrectly) use a custom extension for calcChain.
        if let Some(ext) = xml_attr_value(e, reader, b"Extension")? {
            if ext.eq_ignore_ascii_case("calcchain") || ext.eq_ignore_ascii_case("calcchain.bin") {
                return Ok(true);
            }
        }
    }

    Ok(false)
}

fn remove_calc_chain_from_workbook_rels(xml_bytes: &[u8]) -> Result<Vec<u8>, ParseError> {
    let mut reader = XmlReader::from_reader(std::io::BufReader::new(Cursor::new(xml_bytes)));
    reader.trim_text(false);
    let mut writer = XmlWriter::new(Vec::new());
    let mut buf = Vec::new();
    let mut skip_depth = 0usize;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Eof) => break,
            Ok(event) => {
                if skip_depth > 0 {
                    match event {
                        Event::Start(_) => skip_depth += 1,
                        Event::End(_) => skip_depth = skip_depth.saturating_sub(1),
                        _ => {}
                    }
                } else if should_drop_workbook_rel_event(&event, &reader)? {
                    if matches!(event, Event::Start(_)) {
                        skip_depth = 1;
                    }
                } else {
                    writer.write_event(event.into_owned())?;
                }
            }
            Err(e) => return Err(ParseError::Xml(e)),
        }
        buf.clear();
    }

    Ok(writer.into_inner())
}

fn should_drop_workbook_rel_event<B: std::io::BufRead>(
    event: &Event<'_>,
    reader: &XmlReader<B>,
) -> Result<bool, ParseError> {
    let (Event::Start(e) | Event::Empty(e)) = event else {
        return Ok(false);
    };

    let qname = e.name();
    if !qname.as_ref().ends_with(b"Relationship") {
        return Ok(false);
    }

    if let Some(target) = xml_attr_value(e, reader, b"Target")? {
        let normalized = target.replace('\\', "/");
        if normalized.to_ascii_lowercase().ends_with("calcchain.bin") {
            return Ok(true);
        }
    }

    if let Some(ty) = xml_attr_value(e, reader, b"Type")? {
        if ty.to_ascii_lowercase().contains("relationships/calcchain") {
            return Ok(true);
        }
    }

    Ok(false)
}

fn xml_attr_value<B: std::io::BufRead>(
    e: &quick_xml::events::BytesStart<'_>,
    reader: &XmlReader<B>,
    key: &[u8],
) -> Result<Option<String>, ParseError> {
    for attr in e.attributes().flatten() {
        if attr.key.as_ref() == key {
            return Ok(Some(attr.decode_and_unescape_value(reader)?.into_owned()));
        }
    }
    Ok(None)
}
