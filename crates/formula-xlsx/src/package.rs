use std::collections::{BTreeMap, HashMap, HashSet};
use std::io::{Cursor, Read, Seek, SeekFrom, Write};

use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader as XmlReader, Writer as XmlWriter};
use thiserror::Error;

use crate::patch::{
    apply_cell_patches_to_package, apply_cell_patches_to_package_with_styles, WorkbookCellPatches,
};
use crate::pivots::cache_records::{PivotCacheRecordsReader, PivotCacheValue};
use crate::pivots::XlsxPivots;
use crate::recalc_policy::RecalcPolicyError;
use crate::sheet_metadata::{
    parse_sheet_tab_color, parse_workbook_sheets, write_sheet_tab_color, write_workbook_sheets,
    WorkbookSheetInfo,
};
use crate::theme::{parse_theme_palette, ThemePalette};
use crate::zip_util::open_zip_part;
use crate::{DateSystem, RecalcPolicy};
use formula_model::{CellRef, CellValue, SheetVisibility, StyleTable, TabColor};

const REL_TYPE_THEME: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";

/// Maximum allowed *inflated* bytes for a single ZIP entry in an XLSX package.
///
/// This is a safety limit to prevent loading ZIP bombs into memory when callers need to
/// materialize an entire XLSX/XLSM package (`XlsxPackage`) for preservation/repacking.
pub const MAX_XLSX_PACKAGE_PART_BYTES: u64 = 256 * 1024 * 1024; // 256 MiB

/// Maximum allowed *inflated* bytes across all ZIP entries in an XLSX package.
///
/// This is a safety limit to prevent loading ZIP bombs into memory when callers need to
/// materialize an entire XLSX/XLSM package (`XlsxPackage`) for preservation/repacking.
pub const MAX_XLSX_PACKAGE_TOTAL_BYTES: u64 = 512 * 1024 * 1024; // 512 MiB

/// Size limits enforced by [`XlsxPackage::from_bytes_limited`].
#[derive(Debug, Clone, Copy)]
pub struct XlsxPackageLimits {
    /// Maximum allowed uncompressed bytes for any single part.
    pub max_part_bytes: u64,
    /// Maximum allowed uncompressed bytes across the whole package.
    pub max_total_bytes: u64,
}

impl Default for XlsxPackageLimits {
    fn default() -> Self {
        Self {
            max_part_bytes: MAX_XLSX_PACKAGE_PART_BYTES,
            max_total_bytes: MAX_XLSX_PACKAGE_TOTAL_BYTES,
        }
    }
}

/// Excel workbook "kind" that drives the `/xl/workbook.xml` content type override in
/// `[Content_Types].xml`.
///
/// This is primarily used when exporting an XLSX package under a different extension (for example
/// `.xltx` templates or `.xlam` add-ins).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WorkbookKind {
    /// Standard workbook (`.xlsx`).
    Workbook,
    /// Macro-enabled workbook (`.xlsm`).
    MacroEnabledWorkbook,
    /// Workbook template (`.xltx`).
    Template,
    /// Macro-enabled workbook template (`.xltm`).
    MacroEnabledTemplate,
    /// Macro-enabled add-in (`.xlam`).
    MacroEnabledAddIn,
}

impl WorkbookKind {
    pub fn from_extension(ext: &str) -> Option<Self> {
        if ext.eq_ignore_ascii_case("xlsx") {
            return Some(Self::Workbook);
        }
        if ext.eq_ignore_ascii_case("xlsm") {
            return Some(Self::MacroEnabledWorkbook);
        }
        if ext.eq_ignore_ascii_case("xltx") {
            return Some(Self::Template);
        }
        if ext.eq_ignore_ascii_case("xltm") {
            return Some(Self::MacroEnabledTemplate);
        }
        if ext.eq_ignore_ascii_case("xlam") {
            return Some(Self::MacroEnabledAddIn);
        }
        None
    }

    pub fn workbook_content_type(self) -> &'static str {
        match self {
            WorkbookKind::Workbook => {
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
            }
            WorkbookKind::MacroEnabledWorkbook => {
                "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
            }
            WorkbookKind::Template => {
                "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml"
            }
            WorkbookKind::MacroEnabledTemplate => {
                "application/vnd.ms-excel.template.macroEnabled.main+xml"
            }
            WorkbookKind::MacroEnabledAddIn => {
                "application/vnd.ms-excel.addin.macroEnabled.main+xml"
            }
        }
    }

    pub fn is_macro_enabled(self) -> bool {
        matches!(
            self,
            WorkbookKind::MacroEnabledWorkbook
                | WorkbookKind::MacroEnabledTemplate
                | WorkbookKind::MacroEnabledAddIn
        )
    }

    pub fn is_macro_free(self) -> bool {
        matches!(self, WorkbookKind::Workbook | WorkbookKind::Template)
    }

    /// Detect a [`WorkbookKind`] from the `/xl/workbook.xml` "main" content type string.
    pub fn from_workbook_main_content_type(content_type: &str) -> Option<Self> {
        match content_type.trim() {
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" => {
                Some(Self::Workbook)
            }
            "application/vnd.ms-excel.sheet.macroEnabled.main+xml" => {
                Some(Self::MacroEnabledWorkbook)
            }
            "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml" => {
                Some(Self::Template)
            }
            "application/vnd.ms-excel.template.macroEnabled.main+xml" => {
                Some(Self::MacroEnabledTemplate)
            }
            "application/vnd.ms-excel.addin.macroEnabled.main+xml" => Some(Self::MacroEnabledAddIn),
            _ => None,
        }
    }

    /// Return the closest macro-free kind for this workbook kind.
    ///
    /// Note: Excel does not define a macro-free add-in extension; callers that strip macros from
    /// `.xlam` packages should treat the result as a standard `.xlsx` workbook.
    pub fn macro_free_kind(self) -> Self {
        match self {
            Self::MacroEnabledWorkbook => Self::Workbook,
            Self::MacroEnabledTemplate => Self::Template,
            Self::MacroEnabledAddIn => Self::Workbook,
            other => other,
        }
    }
}

/// Rewrite an existing `[Content_Types].xml` payload so that the `/xl/workbook.xml` override
/// advertises the workbook main content type corresponding to `kind`.
///
/// Returns `Ok(None)` when the input already matches `kind`.
pub fn rewrite_content_types_workbook_kind(
    content_types_xml: &[u8],
    kind: WorkbookKind,
) -> Result<Option<Vec<u8>>, XlsxError> {
    rewrite_content_types_workbook_content_type(content_types_xml, kind.workbook_content_type())
}

/// Rewrite an existing `[Content_Types].xml` payload so that the `/xl/workbook.xml` override
/// advertises `workbook_content_type`.
///
/// Returns `Ok(None)` when no change is required.
pub fn rewrite_content_types_workbook_content_type(
    content_types_xml: &[u8],
    workbook_content_type: &str,
) -> Result<Option<Vec<u8>>, XlsxError> {
    let mut reader = XmlReader::from_reader(content_types_xml);
    reader.config_mut().trim_text(false);
    let mut writer = XmlWriter::new(Vec::with_capacity(content_types_xml.len() + 128));
    let mut buf = Vec::new();

    let mut override_tag_name: Option<String> = None;
    let mut found = false;
    let mut changed = false;

    fn patch_workbook_override(
        e: &BytesStart<'_>,
        workbook_content_type: &str,
    ) -> Result<(bool, Option<BytesStart<'static>>), XlsxError> {
        let mut part_name = None;
        let mut existing_content_type = None;

        for attr in e.attributes().with_checks(false) {
            let attr = attr?;
            match crate::openxml::local_name(attr.key.as_ref()) {
                b"PartName" => part_name = Some(attr.unescape_value()?.into_owned()),
                b"ContentType" => existing_content_type = Some(attr.unescape_value()?.into_owned()),
                _ => {}
            }
        }

        let Some(part_name) = part_name else {
            return Ok((false, None));
        };
        let normalized = part_name.strip_prefix('/').unwrap_or(part_name.as_str());
        if normalized != "xl/workbook.xml" {
            return Ok((false, None));
        }

        if existing_content_type.as_deref() == Some(workbook_content_type) {
            return Ok((true, None));
        }

        let name = e.name();
        let tag_name = std::str::from_utf8(name.as_ref()).unwrap_or("Override");
        let mut patched = BytesStart::new(tag_name);
        let mut saw_content_type = false;
        for attr in e.attributes().with_checks(false) {
            let attr = attr?;
            if crate::openxml::local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"ContentType") {
                saw_content_type = true;
                patched.push_attribute((attr.key.as_ref(), workbook_content_type.as_bytes()));
            } else {
                patched.push_attribute((attr.key.as_ref(), attr.value.as_ref()));
            }
        }
        if !saw_content_type {
            patched.push_attribute(("ContentType", workbook_content_type));
        }

        Ok((true, Some(patched.into_owned())))
    }

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Start(ref e)
                if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Override") =>
            {
                if override_tag_name.is_none() {
                    override_tag_name =
                        Some(String::from_utf8_lossy(e.name().as_ref()).into_owned());
                }
                let (is_workbook, patched) = patch_workbook_override(e, workbook_content_type)?;
                if is_workbook {
                    found = true;
                }
                if let Some(patched) = patched {
                    changed = true;
                    writer.write_event(Event::Start(patched))?;
                } else {
                    writer.write_event(Event::Start(e.to_owned()))?;
                }
            }
            Event::Empty(ref e)
                if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Override") =>
            {
                if override_tag_name.is_none() {
                    override_tag_name =
                        Some(String::from_utf8_lossy(e.name().as_ref()).into_owned());
                }
                let (is_workbook, patched) = patch_workbook_override(e, workbook_content_type)?;
                if is_workbook {
                    found = true;
                }
                if let Some(patched) = patched {
                    changed = true;
                    writer.write_event(Event::Empty(patched))?;
                } else {
                    writer.write_event(Event::Empty(e.to_owned()))?;
                }
            }
            Event::End(ref e) if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Types") => {
                if !found {
                    changed = true;
                    let override_tag_name = override_tag_name
                        .clone()
                        .unwrap_or_else(|| prefixed_tag(e.name().as_ref(), "Override"));
                    let mut override_el = BytesStart::new(override_tag_name.as_str());
                    override_el.push_attribute(("PartName", "/xl/workbook.xml"));
                    override_el.push_attribute(("ContentType", workbook_content_type));
                    writer.write_event(Event::Empty(override_el))?;
                }
                writer.write_event(Event::End(e.to_owned()))?;
            }
            Event::Empty(ref e) if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Types") => {
                if !found {
                    changed = true;
                    let types_tag_name = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                    let override_tag_name = override_tag_name
                        .clone()
                        .unwrap_or_else(|| prefixed_tag(types_tag_name.as_bytes(), "Override"));

                    writer.write_event(Event::Start(e.to_owned()))?;

                    let mut override_el = BytesStart::new(override_tag_name.as_str());
                    override_el.push_attribute(("PartName", "/xl/workbook.xml"));
                    override_el.push_attribute(("ContentType", workbook_content_type));
                    writer.write_event(Event::Empty(override_el))?;

                    writer.write_event(Event::End(BytesEnd::new(types_tag_name.as_str())))?;
                } else {
                    writer.write_event(Event::Empty(e.to_owned()))?;
                }
            }
            Event::Eof => break,
            other => writer.write_event(other.into_owned())?,
        }

        buf.clear();
    }

    if changed {
        Ok(Some(writer.into_inner()))
    } else {
        Ok(None)
    }
}

#[derive(Debug, Error)]
pub enum XlsxError {
    #[error("zip error: {0}")]
    Zip(#[from] zip::result::ZipError),
    #[error("io error: {0}")]
    Io(#[from] std::io::Error),
    #[error("xml error: {0}")]
    Xml(#[from] quick_xml::Error),
    #[error("xml error: {0}")]
    RoXml(#[from] roxmltree::Error),
    #[error("utf-8 error: {0}")]
    Utf8(#[from] std::string::FromUtf8Error),
    #[error("xml attribute error: {0}")]
    Attr(#[from] quick_xml::events::attributes::AttrError),
    #[error("missing required attribute: {0}")]
    MissingAttr(&'static str),
    #[error("missing xlsx part: {0}")]
    MissingPart(String),
    #[error("invalid xlsx: {0}")]
    Invalid(String),
    #[error(
        "xlsx package part is too large to load safely: {part} is {size} bytes (max {max} bytes). \
Try reducing workbook size or saving without preserved parts."
    )]
    PartTooLarge { part: String, size: u64, max: u64 },
    #[error(
        "xlsx package is too large to load safely: {total} bytes uncompressed (max {max}). \
Try reducing workbook size or saving without preserved parts."
    )]
    PackageTooLarge { total: u64, max: u64 },
    #[error("invalid sheetId value")]
    InvalidSheetId,
    #[error("hyperlink error: {0}")]
    Hyperlink(String),
    #[error("invalid password for encrypted workbook")]
    InvalidPassword,
    #[error("unsupported encryption: {0}")]
    UnsupportedEncryption(String),
    #[error("invalid encrypted workbook: {0}")]
    InvalidEncryptedWorkbook(String),
    #[error(transparent)]
    StreamingPatch(#[from] Box<crate::streaming::StreamingPatchError>),
}

impl From<crate::streaming::StreamingPatchError> for XlsxError {
    fn from(err: crate::streaming::StreamingPatchError) -> Self {
        Self::StreamingPatch(Box::new(err))
    }
}

impl From<crate::encrypted::EncryptedOoxmlError> for XlsxError {
    fn from(err: crate::encrypted::EncryptedOoxmlError) -> Self {
        match err {
            crate::encrypted::EncryptedOoxmlError::InvalidPassword => Self::InvalidPassword,
            crate::encrypted::EncryptedOoxmlError::UnsupportedEncryption(msg) => {
                Self::UnsupportedEncryption(msg)
            }
            crate::encrypted::EncryptedOoxmlError::InvalidEncryptedWorkbook(msg) => {
                Self::InvalidEncryptedWorkbook(msg)
            }
            crate::encrypted::EncryptedOoxmlError::Io(err) => Self::Io(err),
        }
    }
}

/// Resolved metadata for a workbook sheet and its corresponding worksheet part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WorksheetPartInfo {
    pub name: String,
    pub sheet_id: u32,
    pub rel_id: String,
    pub visibility: SheetVisibility,
    /// ZIP entry name for the worksheet XML (e.g. `xl/worksheets/sheet1.xml`).
    pub worksheet_part: String,
}

/// Select a target worksheet for a cell patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum CellPatchSheet {
    /// Identify the sheet by workbook sheet name (e.g. `"Sheet1"`).
    SheetName(String),
    /// Identify the sheet by its worksheet XML part name (e.g. `"xl/worksheets/sheet1.xml"`).
    WorksheetPart(String),
}

/// A single cell edit to apply to an [`XlsxPackage`].
#[derive(Debug, Clone, PartialEq)]
pub struct CellPatch {
    pub sheet: CellPatchSheet,
    pub cell: CellRef,
    pub value: CellValue,
    /// Optional formula to write into the `<f>` element. Leading `=` is permitted.
    pub formula: Option<String>,
    /// Optional cell `vm` attribute override.
    ///
    /// SpreadsheetML uses `c/@vm` for RichData-backed cell content (e.g. images-in-cell).
    ///
    /// - `None`: preserve the existing attribute when patching an existing cell (and omit it when
    ///   inserting a new cell).
    /// - `Some(Some(n))`: set/overwrite `vm="n"`.
    /// - `Some(None)`: remove the attribute.
    pub vm: Option<Option<u32>>,
    /// Optional cell `cm` attribute override.
    ///
    /// Some RichData-backed cell content also requires `c/@cm`.
    ///
    /// - `None`: preserve the existing attribute when patching an existing cell (and omit it when
    ///   inserting a new cell).
    /// - `Some(Some(n))`: set/overwrite `cm="n"`.
    /// - `Some(None)`: remove the attribute.
    pub cm: Option<Option<u32>>,
}

impl CellPatch {
    pub fn new(
        sheet: CellPatchSheet,
        cell: CellRef,
        value: CellValue,
        formula: Option<String>,
    ) -> Self {
        Self {
            sheet,
            cell,
            value,
            formula,
            vm: None,
            cm: None,
        }
    }

    pub fn with_vm(mut self, vm: Option<Option<u32>>) -> Self {
        self.vm = vm;
        self
    }

    pub fn with_cm(mut self, cm: Option<Option<u32>>) -> Self {
        self.cm = cm;
        self
    }

    pub fn set_vm(self, vm: u32) -> Self {
        self.with_vm(Some(Some(vm)))
    }

    pub fn clear_vm(self) -> Self {
        self.with_vm(Some(None))
    }

    pub fn set_cm(self, cm: u32) -> Self {
        self.with_cm(Some(Some(cm)))
    }

    pub fn clear_cm(self) -> Self {
        self.with_cm(Some(None))
    }

    pub fn for_sheet_name(
        sheet_name: impl Into<String>,
        cell: CellRef,
        value: CellValue,
        formula: Option<String>,
    ) -> Self {
        Self::new(
            CellPatchSheet::SheetName(sheet_name.into()),
            cell,
            value,
            formula,
        )
    }

    pub fn for_worksheet_part(
        worksheet_part: impl Into<String>,
        cell: CellRef,
        value: CellValue,
        formula: Option<String>,
    ) -> Self {
        Self::new(
            CellPatchSheet::WorksheetPart(worksheet_part.into()),
            cell,
            value,
            formula,
        )
    }
}

impl From<RecalcPolicyError> for XlsxError {
    fn from(err: RecalcPolicyError) -> Self {
        match err {
            RecalcPolicyError::Io(err) => XlsxError::Io(err),
            RecalcPolicyError::Xml(err) => XlsxError::Xml(err),
            RecalcPolicyError::XmlAttr(err) => XlsxError::Attr(err),
        }
    }
}

/// Presence information for macro-capable workbook content.
///
/// This is intentionally more granular than a single `has_macros` flag so callers can distinguish
/// classic VBA projects from Excel 4.0 macrosheets and legacy dialog sheets.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct MacroPresence {
    pub has_vba: bool,
    pub has_xlm_macrosheets: bool,
    pub has_dialog_sheets: bool,
}

impl MacroPresence {
    pub fn any(self) -> bool {
        self.has_vba || self.has_xlm_macrosheets || self.has_dialog_sheets
    }
}

/// In-memory representation of an XLSX/XLSM package as a map of part name -> bytes.
///
/// We keep the API minimal to support macro preservation; a full model will
/// eventually build on top of this.
#[derive(Debug, Clone)]
pub struct XlsxPackage {
    parts: BTreeMap<String, Vec<u8>>,
}

/// Read a single ZIP part from an XLSX/XLSM container without inflating the entire package.
pub fn read_part_from_reader<R: Read + Seek>(
    reader: R,
    part_name: &str,
) -> Result<Option<Vec<u8>>, XlsxError> {
    read_part_from_reader_limited(reader, part_name, MAX_XLSX_PACKAGE_PART_BYTES)
}

/// Read a single ZIP part from an XLSX/XLSM container, enforcing a maximum uncompressed size.
///
/// This protects callers that need to extract a single part from untrusted workbooks without
/// risking unbounded allocations from ZIP metadata or decompression bombs.
pub fn read_part_from_reader_limited<R: Read + Seek>(
    mut reader: R,
    part_name: &str,
    max_bytes: u64,
) -> Result<Option<Vec<u8>>, XlsxError> {
    reader.seek(SeekFrom::Start(0))?;
    let mut zip = zip::ZipArchive::new(reader)?;

    let max_bytes = max_bytes.min(usize::MAX as u64);
    // `ZipFile` borrows `ZipArchive`; keep the `Result` in a local so it drops before `zip` to
    // avoid borrowck issues.
    let result = crate::zip_util::open_zip_part(&mut zip, part_name);
    match result {
        Ok(mut file) => {
            if file.is_dir() {
                return Ok(None);
            }
            Ok(Some(crate::zip_util::read_zip_file_bytes_with_limit(
                &mut file, part_name, max_bytes,
            )?))
        }
        Err(zip::result::ZipError::FileNotFound) => Ok(None),
        Err(err) => Err(err.into()),
    }
}

/// Parse the workbook theme palette from `xl/theme/theme1.xml` (if present) without inflating the
/// entire package.
pub fn theme_palette_from_reader<R: Read + Seek>(
    reader: R,
) -> Result<Option<ThemePalette>, XlsxError> {
    theme_palette_from_reader_limited(reader, MAX_XLSX_PACKAGE_PART_BYTES)
}

/// Parse the workbook theme palette from `xl/theme/theme1.xml` (if present) without inflating the
/// entire package, enforcing a maximum part size.
pub fn theme_palette_from_reader_limited<R: Read + Seek>(
    mut reader: R,
    max_bytes: u64,
) -> Result<Option<ThemePalette>, XlsxError> {
    reader.seek(SeekFrom::Start(0))?;
    let mut zip = zip::ZipArchive::new(reader)?;

    let rels_bytes = crate::zip_util::read_zip_part_optional_with_limit(
        &mut zip,
        "xl/_rels/workbook.xml.rels",
        max_bytes,
    )?;
    let theme_candidates: Vec<String> = rels_bytes
        .as_deref()
        .and_then(|bytes| crate::openxml::parse_relationships(bytes).ok())
        .and_then(|rels| {
            rels.into_iter().find(|rel| {
                rel.type_uri == REL_TYPE_THEME
                    && !rel
                        .target_mode
                        .as_deref()
                        .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
            })
        })
        .map(|rel| crate::path::resolve_target_candidates("xl/workbook.xml", &rel.target))
        .unwrap_or_else(|| vec!["xl/theme/theme1.xml".to_string()]);

    for candidate in theme_candidates {
        let Some(theme_xml) =
            crate::zip_util::read_zip_part_optional_with_limit(&mut zip, &candidate, max_bytes)?
        else {
            continue;
        };
        return Ok(Some(parse_theme_palette(&theme_xml)?));
    }

    Ok(None)
}

/// Resolve ordered workbook sheets to worksheet part names without inflating the entire package.
///
/// This is the streaming counterpart of [`XlsxPackage::worksheet_parts`]. It is intentionally
/// best-effort: if `xl/_rels/workbook.xml.rels` is missing or malformed, it falls back to the
/// common `xl/worksheets/sheet{sheetId}.xml` naming convention when possible.
pub fn worksheet_parts_from_reader<R: Read + Seek>(
    mut reader: R,
) -> Result<Vec<WorksheetPartInfo>, XlsxError> {
    reader.seek(SeekFrom::Start(0))?;
    let mut zip = zip::ZipArchive::new(reader)?;

    let mut part_names: HashSet<String> = HashSet::new();
    // Map a canonicalized lookup key (case/separator-insensitive, percent-decoding `%xx`) back to
    // the canonical part name we should return. This lets us resolve non-canonical relationship
    // targets (e.g. unescaped spaces) to percent-encoded ZIP entries while still returning
    // canonical part names (normalized separators, preserved `%xx` escapes).
    let mut part_name_keys: HashMap<Vec<u8>, String> = HashMap::new();
    for i in 0..zip.len() {
        let file = zip.by_index(i)?;
        if file.is_dir() {
            continue;
        }
        // ZIP entry names in valid XLSX/XLSM packages should not start with `/`, but tolerate
        // producers that include it (or use `\`) by normalizing to canonical part names.
        let name = file.name();
        let canonical = name.trim_start_matches(|c| c == '/' || c == '\\');
        let canonical = if canonical.contains('\\') {
            canonical.replace('\\', "/")
        } else {
            canonical.to_string()
        };
        part_names.insert(canonical.clone());
        part_name_keys
            .entry(crate::zip_util::zip_part_name_lookup_key(&canonical))
            .or_insert(canonical);
    }

    let workbook_xml = match open_zip_part(&mut zip, "xl/workbook.xml") {
        Ok(mut file) => crate::zip_util::read_zip_file_bytes_with_limit(
            &mut file,
            "xl/workbook.xml",
            MAX_XLSX_PACKAGE_PART_BYTES,
        )?,
        Err(zip::result::ZipError::FileNotFound) => {
            return Err(XlsxError::MissingPart("xl/workbook.xml".to_string()))
        }
        Err(err) => return Err(err.into()),
    };
    let workbook_xml = String::from_utf8(workbook_xml)?;
    let sheets = parse_workbook_sheets(&workbook_xml)?;

    let rels_bytes = match open_zip_part(&mut zip, "xl/_rels/workbook.xml.rels") {
        Ok(mut file) => crate::zip_util::read_zip_file_bytes_with_limit(
            &mut file,
            "xl/_rels/workbook.xml.rels",
            MAX_XLSX_PACKAGE_PART_BYTES,
        )
        .map(Some)?,
        Err(zip::result::ZipError::FileNotFound) => None,
        Err(err) => return Err(err.into()),
    };

    let relationships = match rels_bytes.as_deref() {
        Some(bytes) => crate::openxml::parse_relationships(bytes).unwrap_or_default(),
        None => Vec::new(),
    };
    let rel_by_id: HashMap<String, crate::openxml::Relationship> = relationships
        .into_iter()
        .map(|rel| (rel.id.clone(), rel))
        .collect();

    let workbook_part = "xl/workbook.xml";
    let mut out = Vec::with_capacity(sheets.len());

    for sheet in sheets {
        let resolved = rel_by_id
            .get(&sheet.rel_id)
            .filter(|rel| {
                !rel.target_mode
                    .as_deref()
                    .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
            })
            .and_then(|rel| {
                let candidates = crate::path::resolve_target_candidates(workbook_part, &rel.target);
                // Prefer exact matches to keep part-name strings canonical when possible (some
                // producers percent-encode relationship targets while storing ZIP entry names
                // unescaped, and vice versa).
                for candidate in &candidates {
                    if part_names.contains(candidate) {
                        return Some(candidate.clone());
                    }
                }
                candidates.into_iter().find_map(|candidate| {
                    part_name_keys
                        .get(&crate::zip_util::zip_part_name_lookup_key(&candidate))
                        .cloned()
                })
            })
            .or_else(|| {
                let candidate = format!("xl/worksheets/sheet{}.xml", sheet.sheet_id);
                if part_names.contains(&candidate) {
                    return Some(candidate);
                }
                part_name_keys
                    .get(&crate::zip_util::zip_part_name_lookup_key(&candidate))
                    .cloned()
            });

        let Some(worksheet_part) = resolved else {
            continue;
        };

        out.push(WorksheetPartInfo {
            name: sheet.name,
            sheet_id: sheet.sheet_id,
            rel_id: sheet.rel_id,
            visibility: sheet.visibility,
            worksheet_part,
        });
    }

    Ok(out)
}

/// Resolve ordered workbook sheets to worksheet part names without inflating the entire package,
/// enforcing a maximum size for any XML part that must be loaded (`xl/workbook.xml` and
/// `xl/_rels/workbook.xml.rels`).
pub fn worksheet_parts_from_reader_limited<R: Read + Seek>(
    mut reader: R,
    max_part_bytes: u64,
) -> Result<Vec<WorksheetPartInfo>, XlsxError> {
    reader.seek(SeekFrom::Start(0))?;
    let mut zip = zip::ZipArchive::new(reader)?;

    let mut part_names: HashSet<String> = HashSet::new();
    // Map a canonicalized lookup key (case/separator-insensitive, percent-decoding `%xx`) back to
    // the canonical part name we should return (normalized separators, preserved `%xx` escapes).
    let mut part_name_keys: HashMap<Vec<u8>, String> = HashMap::new();
    for i in 0..zip.len() {
        let file = zip.by_index(i)?;
        if file.is_dir() {
            continue;
        }
        // ZIP entry names in valid XLSX/XLSM packages should not start with `/`, but tolerate
        // producers that include it (or use `\`) by normalizing to canonical part names.
        let name = file.name();
        let canonical = name.trim_start_matches(|c| c == '/' || c == '\\');
        let canonical = if canonical.contains('\\') {
            canonical.replace('\\', "/")
        } else {
            canonical.to_string()
        };
        part_names.insert(canonical.clone());
        part_name_keys
            .entry(crate::zip_util::zip_part_name_lookup_key(&canonical))
            .or_insert(canonical);
    }

    fn read_zip_part_required<R: Read + Seek>(
        zip: &mut zip::ZipArchive<R>,
        name: &str,
        max_bytes: u64,
    ) -> Result<Vec<u8>, XlsxError> {
        crate::zip_util::read_zip_part_optional_with_limit(zip, name, max_bytes)?.ok_or_else(|| {
            XlsxError::MissingPart(name.strip_prefix('/').unwrap_or(name).to_string())
        })
    }

    fn read_zip_part_optional<R: Read + Seek>(
        zip: &mut zip::ZipArchive<R>,
        name: &str,
        max_bytes: u64,
    ) -> Result<Option<Vec<u8>>, XlsxError> {
        crate::zip_util::read_zip_part_optional_with_limit(zip, name, max_bytes)
    }

    let workbook_xml = read_zip_part_required(&mut zip, "xl/workbook.xml", max_part_bytes)?;
    let workbook_xml = String::from_utf8(workbook_xml)?;
    let sheets = parse_workbook_sheets(&workbook_xml)?;

    let rels_bytes =
        read_zip_part_optional(&mut zip, "xl/_rels/workbook.xml.rels", max_part_bytes)?;

    let relationships = match rels_bytes.as_deref() {
        Some(bytes) => crate::openxml::parse_relationships(bytes).unwrap_or_default(),
        None => Vec::new(),
    };
    let rel_by_id: HashMap<String, crate::openxml::Relationship> = relationships
        .into_iter()
        .map(|rel| (rel.id.clone(), rel))
        .collect();

    let workbook_part = "xl/workbook.xml";
    let mut out = Vec::with_capacity(sheets.len());

    for sheet in sheets {
        let resolved = rel_by_id
            .get(&sheet.rel_id)
            .filter(|rel| {
                !rel.target_mode
                    .as_deref()
                    .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
            })
            .and_then(|rel| {
                let candidates = crate::path::resolve_target_candidates(workbook_part, &rel.target);
                for candidate in &candidates {
                    if part_names.contains(candidate) {
                        return Some(candidate.clone());
                    }
                }
                candidates.into_iter().find_map(|candidate| {
                    part_name_keys
                        .get(&crate::zip_util::zip_part_name_lookup_key(&candidate))
                        .cloned()
                })
            })
            .or_else(|| {
                let candidate = format!("xl/worksheets/sheet{}.xml", sheet.sheet_id);
                if part_names.contains(&candidate) {
                    return Some(candidate);
                }
                part_name_keys
                    .get(&crate::zip_util::zip_part_name_lookup_key(&candidate))
                    .cloned()
            });

        let Some(worksheet_part) = resolved else {
            continue;
        };

        out.push(WorksheetPartInfo {
            name: sheet.name,
            sheet_id: sheet.sheet_id,
            rel_id: sheet.rel_id,
            visibility: sheet.visibility,
            worksheet_part,
        });
    }

    Ok(out)
}

impl XlsxPackage {
    pub fn from_bytes(bytes: &[u8]) -> Result<Self, XlsxError> {
        Self::from_bytes_limited(bytes, XlsxPackageLimits::default())
    }

    pub fn from_bytes_limited(bytes: &[u8], limits: XlsxPackageLimits) -> Result<Self, XlsxError> {
        let reader = Cursor::new(bytes);
        let mut zip = zip::ZipArchive::new(reader)?;

        let mut parts = BTreeMap::new();
        let mut budget = crate::zip_util::ZipInflateBudget::new(limits.max_total_bytes);
        for i in 0..zip.len() {
            let mut file = zip.by_index(i)?;
            if !file.is_file() {
                continue;
            }

            let name = file.name().to_string();
            let buf = crate::zip_util::read_zip_file_bytes_with_budget(
                &mut file,
                &name,
                limits.max_part_bytes,
                &mut budget,
            )?;
            parts.insert(name, buf);
        }

        Ok(Self { parts })
    }

    /// Load an [`XlsxPackage`] from bytes, transparently decrypting Office `EncryptedPackage`
    /// OLE wrappers when the input bytes are password-protected.
    pub fn from_bytes_with_password(bytes: &[u8], password: &str) -> Result<Self, XlsxError> {
        let bytes = crate::encrypted::maybe_decrypt_office_encrypted_package(bytes, password)?;
        Self::from_bytes(bytes.as_ref())
    }

    /// Construct an [`XlsxPackage`] from an already-inflated part map.
    ///
    /// This is crate-private and is primarily used by higher-level readers that already have all
    /// parts in memory (e.g. `load_from_bytes`) but want to reuse `XlsxPackage` helpers like pivot
    /// chart discovery without re-reading the ZIP container.
    pub(crate) fn from_parts_map(parts: BTreeMap<String, Vec<u8>>) -> Self {
        Self { parts }
    }

    pub(crate) fn into_parts_map(self) -> BTreeMap<String, Vec<u8>> {
        self.parts
    }

    pub fn part(&self, name: &str) -> Option<&[u8]> {
        if let Some(bytes) = self.parts.get(name) {
            return Some(bytes.as_slice());
        }

        if let Some(stripped) = name.strip_prefix('/') {
            return self.parts.get(stripped).map(|v| v.as_slice());
        }

        let mut with_slash = String::with_capacity(name.len() + 1);
        with_slash.push('/');
        with_slash.push_str(name);
        if let Some(bytes) = self.parts.get(with_slash.as_str()) {
            return Some(bytes.as_slice());
        }

        // Fall back to a linear scan for non-canonical producer output:
        // - Windows-style `\` separators (treated as `/`)
        // - ASCII case differences (e.g. `XL/Workbook.xml`)
        //
        // Note: We intentionally do *not* canonicalize the stored part names; this map is used for
        // higher-fidelity round-tripping where preserving unknown part names is valuable.
        self.parts
            .iter()
            .find(|(key, _)| crate::zip_util::zip_part_names_equivalent(key.as_str(), name))
            .map(|(_, bytes)| bytes.as_slice())
    }

    pub fn parts(&self) -> impl Iterator<Item = (&str, &[u8])> {
        self.parts
            .iter()
            .map(|(name, bytes)| (name.as_str(), bytes.as_slice()))
    }

    pub fn part_names(&self) -> impl Iterator<Item = &str> {
        self.parts.keys().map(String::as_str)
    }

    /// Borrow the raw part map (useful for higher-fidelity operations).
    pub fn parts_map(&self) -> &BTreeMap<String, Vec<u8>> {
        &self.parts
    }

    pub fn parts_map_mut(&mut self) -> &mut BTreeMap<String, Vec<u8>> {
        &mut self.parts
    }

    pub fn set_part(&mut self, name: impl Into<String>, bytes: Vec<u8>) {
        self.parts.insert(name.into(), bytes);
    }

    pub fn vba_project_bin(&self) -> Option<&[u8]> {
        self.part("xl/vbaProject.bin")
    }

    /// Returns the raw `xl/vbaProjectSignature.bin` payload when present.
    ///
    /// Some XLSM producers store VBA signature streams in a separate OPC part
    /// (`vbaProjectSignature.bin`) instead of embedding them inside
    /// `vbaProject.bin`.
    pub fn vba_project_signature_bin(&self) -> Option<&[u8]> {
        self.part("xl/vbaProjectSignature.bin")
    }

    /// Returns the optional `xl/vbaData.xml` payload when present.
    pub fn vba_data_xml(&self) -> Option<&[u8]> {
        self.part("xl/vbaData.xml")
    }

    /// Detect whether the package contains any macro-capable content (VBA, XLM macrosheets, or
    /// legacy dialog sheets).
    pub fn macro_presence(&self) -> MacroPresence {
        let mut presence = MacroPresence {
            has_vba: false,
            has_xlm_macrosheets: false,
            has_dialog_sheets: false,
        };

        for name in self.part_names() {
            let key = crate::zip_util::zip_part_name_lookup_key(name);
            if key == b"xl/vbaproject.bin" {
                presence.has_vba = true;
            }
            if key.starts_with(b"xl/macrosheets/") {
                presence.has_xlm_macrosheets = true;
            }
            if key.starts_with(b"xl/dialogsheets/") {
                presence.has_dialog_sheets = true;
            }

            if presence.has_vba && presence.has_xlm_macrosheets && presence.has_dialog_sheets {
                break;
            }
        }

        presence
    }

    /// Parse the workbook theme palette from the workbook theme part (if present).
    ///
    /// Prefer the theme part referenced from `xl/_rels/workbook.xml.rels` via relationship type
    /// `.../relationships/theme`, falling back to `xl/theme/theme1.xml` when the relationship is
    /// missing.
    pub fn theme_palette(&self) -> Result<Option<ThemePalette>, XlsxError> {
        let theme_candidates: Vec<String> = self
            .part("xl/_rels/workbook.xml.rels")
            .and_then(|bytes| crate::openxml::parse_relationships(bytes).ok())
            .and_then(|rels| {
                rels.into_iter().find(|rel| {
                    rel.type_uri == REL_TYPE_THEME
                        && !rel
                            .target_mode
                            .as_deref()
                            .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
                })
            })
            .map(|rel| crate::path::resolve_target_candidates("xl/workbook.xml", &rel.target))
            .unwrap_or_else(|| vec!["xl/theme/theme1.xml".to_string()]);

        for candidate in theme_candidates {
            let Some(theme_xml) = self.part(&candidate) else {
                continue;
            };
            return Ok(Some(parse_theme_palette(theme_xml)?));
        }

        Ok(None)
    }

    /// Extract in-cell images from `xl/cellImages.xml` (if present).
    ///
    /// This is a convenience helper that:
    /// - detects the `cellImages` part,
    /// - parses it to discover referenced relationship IDs,
    /// - resolves relationship targets to concrete package part names,
    /// - and returns the referenced image binaries.
    ///
    /// Only relationships with the standard image relationship type
    /// (`.../relationships/image`) are included.
    pub fn extract_cell_images(&self) -> Result<Vec<(String, Vec<u8>)>, XlsxError> {
        let Some(info) = self.cell_images_part_info()? else {
            return Ok(Vec::new());
        };

        Ok(info
            .embeds
            .into_iter()
            .map(|embed| (embed.target_part, embed.target_bytes))
            .collect())
    }

    pub fn write_to_bytes(&self) -> Result<Vec<u8>, XlsxError> {
        let mut buf = Vec::new();
        self.write_to(&mut buf)?;
        Ok(buf)
    }

    #[cfg(not(target_arch = "wasm32"))]
    pub fn write_to_encrypted_ole_bytes(&self, password: &str) -> Result<Vec<u8>, XlsxError> {
        let zip_bytes = self.write_to_bytes()?;
        crate::office_crypto::encrypt_package_to_ole(&zip_bytes, password)
            .map_err(|err| XlsxError::Invalid(format!("office encryption error: {err}")))
    }

    pub fn write_to<W: Write>(&self, mut w: W) -> Result<(), XlsxError> {
        let has_vba_project = self.vba_project_bin().is_some();
        let mut parts = self.parts.clone();
        if has_vba_project {
            crate::macro_repair::ensure_xlsm_content_types(&mut parts)?;
            crate::macro_repair::ensure_workbook_rels_has_vba(&mut parts)?;
            crate::macro_repair::ensure_vba_project_rels_has_signature(&mut parts)?;
        }

        // Ensure `[Content_Types].xml` contains `<Default>` entries for common image extensions
        // when the package includes image/media payloads (e.g. `xl/media/image1.png`).
        //
        // This is intentionally conservative: we only insert defaults for extensions that appear
        // in the package to avoid touching `[Content_Types].xml` unnecessarily.
        let mut needs_png = false;
        let mut needs_jpg = false;
        let mut needs_jpeg = false;
        let mut needs_gif = false;
        let mut needs_bmp = false;
        let mut needs_emf = false;
        let mut needs_wmf = false;
        let mut needs_svg = false;
        let mut needs_tif = false;
        let mut needs_tiff = false;
        let mut needs_webp = false;

        for name in parts.keys() {
            let name = name.strip_prefix('/').unwrap_or(name);
            if crate::ascii::ends_with_ignore_case(name, ".png") {
                needs_png = true;
            } else if crate::ascii::ends_with_ignore_case(name, ".jpg") {
                needs_jpg = true;
            } else if crate::ascii::ends_with_ignore_case(name, ".jpeg") {
                needs_jpeg = true;
            } else if crate::ascii::ends_with_ignore_case(name, ".gif") {
                needs_gif = true;
            } else if crate::ascii::ends_with_ignore_case(name, ".bmp") {
                needs_bmp = true;
            } else if crate::ascii::ends_with_ignore_case(name, ".emf") {
                needs_emf = true;
            } else if crate::ascii::ends_with_ignore_case(name, ".wmf") {
                needs_wmf = true;
            } else if crate::ascii::ends_with_ignore_case(name, ".svg") {
                needs_svg = true;
            } else if crate::ascii::ends_with_ignore_case(name, ".tif") {
                needs_tif = true;
            } else if crate::ascii::ends_with_ignore_case(name, ".tiff") {
                needs_tiff = true;
            } else if crate::ascii::ends_with_ignore_case(name, ".webp") {
                needs_webp = true;
            }
        }

        if needs_png {
            ensure_content_types_default(&mut parts, "png", "image/png")?;
        }
        if needs_jpg {
            ensure_content_types_default(&mut parts, "jpg", "image/jpeg")?;
        }
        if needs_jpeg {
            ensure_content_types_default(&mut parts, "jpeg", "image/jpeg")?;
        }
        if needs_gif {
            ensure_content_types_default(&mut parts, "gif", "image/gif")?;
        }
        if needs_bmp {
            ensure_content_types_default(&mut parts, "bmp", "image/bmp")?;
        }
        if needs_emf {
            ensure_content_types_default(&mut parts, "emf", "image/x-emf")?;
        }
        if needs_wmf {
            ensure_content_types_default(&mut parts, "wmf", "image/x-wmf")?;
        }
        if needs_svg {
            ensure_content_types_default(&mut parts, "svg", "image/svg+xml")?;
        }
        if needs_tif {
            ensure_content_types_default(&mut parts, "tif", "image/tiff")?;
        }
        if needs_tiff {
            ensure_content_types_default(&mut parts, "tiff", "image/tiff")?;
        }
        if needs_webp {
            ensure_content_types_default(&mut parts, "webp", "image/webp")?;
        }

        let cursor = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(cursor);
        let options = zip::write::FileOptions::<()>::default()
            .compression_method(zip::CompressionMethod::Deflated);

        for (name, bytes) in parts {
            zip.start_file(name, options)?;
            zip.write_all(&bytes)?;
        }

        let cursor = zip.finish()?;
        w.write_all(&cursor.into_inner())?;
        Ok(())
    }

    /// Ensure `[Content_Types].xml` advertises the correct workbook content type for the requested
    /// workbook kind.
    pub fn enforce_workbook_kind(&mut self, kind: WorkbookKind) -> Result<(), XlsxError> {
        ensure_workbook_content_type(&mut self.parts, kind.workbook_content_type())
    }

    /// Return the ordered workbook sheets with their resolved worksheet part paths.
    ///
    /// This reads `xl/workbook.xml` for the `<sheet>` list and `xl/_rels/workbook.xml.rels`
    /// to resolve each sheet's `r:id` relationship to a concrete worksheet XML part name.
    pub fn worksheet_parts(&self) -> Result<Vec<WorksheetPartInfo>, XlsxError> {
        let sheets = self.workbook_sheets()?;

        let rels_bytes = self
            .part("xl/_rels/workbook.xml.rels")
            .ok_or_else(|| XlsxError::MissingPart("xl/_rels/workbook.xml.rels".to_string()))?;
        let relationships = crate::openxml::parse_relationships(rels_bytes)?;
        let rel_by_id: HashMap<String, crate::openxml::Relationship> = relationships
            .into_iter()
            .map(|rel| (rel.id.clone(), rel))
            .collect();

        let mut out = Vec::with_capacity(sheets.len());
        for sheet in sheets {
            let rel = rel_by_id.get(&sheet.rel_id).ok_or_else(|| {
                XlsxError::Invalid(format!("missing relationship for {}", sheet.rel_id))
            })?;
            let worksheet_part = crate::path::resolve_target("xl/workbook.xml", &rel.target);
            out.push(WorksheetPartInfo {
                name: sheet.name,
                sheet_id: sheet.sheet_id,
                rel_id: sheet.rel_id,
                visibility: sheet.visibility,
                worksheet_part,
            });
        }
        Ok(out)
    }

    /// Apply a set of cell edits to the package and return the updated ZIP bytes.
    ///
    /// This uses the streaming patch pipeline, which rewrites the targeted worksheet parts and
    /// updates dependent workbook parts when needed (for example `xl/sharedStrings.xml` and the
    /// calcChain/full-calc settings after formula edits).
    pub fn apply_cell_patches_to_bytes(&self, patches: &[CellPatch]) -> Result<Vec<u8>, XlsxError> {
        self.apply_cell_patches_to_bytes_with_recalc_policy(patches, RecalcPolicy::default())
    }

    /// Apply a set of cell edits to the package and return the updated ZIP bytes using the
    /// provided [`RecalcPolicy`] (applied only when formulas change).
    pub fn apply_cell_patches_to_bytes_with_recalc_policy(
        &self,
        patches: &[CellPatch],
        policy_on_formula_change: RecalcPolicy,
    ) -> Result<Vec<u8>, XlsxError> {
        let mut sheet_name_to_part: HashMap<String, String> = HashMap::new();
        let mut worksheet_parts: Vec<WorksheetPartInfo> = Vec::new();
        if patches
            .iter()
            .any(|p| matches!(p.sheet, CellPatchSheet::SheetName(_)))
        {
            worksheet_parts = self.worksheet_parts()?;
            for entry in &worksheet_parts {
                sheet_name_to_part.insert(entry.name.clone(), entry.worksheet_part.clone());
            }
        }

        let mut patches_by_part: HashMap<
            String,
            BTreeMap<(u32, u32), crate::streaming::WorksheetCellPatch>,
        > = HashMap::new();

        for patch in patches {
            let worksheet_part = match &patch.sheet {
                CellPatchSheet::WorksheetPart(part) => part.clone(),
                CellPatchSheet::SheetName(name) => match sheet_name_to_part.get(name).cloned() {
                    Some(part) => part,
                    None => worksheet_parts
                        .iter()
                        .find(|s| formula_model::sheet_name_eq_case_insensitive(&s.name, name))
                        .map(|s| s.worksheet_part.clone())
                        .ok_or_else(|| XlsxError::Invalid(format!("unknown sheet name {name}")))?,
                },
            };

            patches_by_part
                .entry(worksheet_part.clone())
                .or_default()
                .insert(
                    (patch.cell.row, patch.cell.col),
                    crate::streaming::WorksheetCellPatch::new(
                        worksheet_part,
                        patch.cell,
                        patch.value.clone(),
                        patch.formula.clone(),
                    )
                    .with_vm(patch.vm)
                    .with_cm(patch.cm),
                );
        }

        let mut patches_by_part: HashMap<String, Vec<crate::streaming::WorksheetCellPatch>> =
            patches_by_part
                .into_iter()
                .map(|(part, cells)| (part, cells.into_values().collect()))
                .collect();
        for patches in patches_by_part.values_mut() {
            patches.sort_by_key(|p| (p.cell.row, p.cell.col));
        }

        for part in patches_by_part.keys() {
            // Be tolerant to non-canonical producer output:
            // - leading `/`
            // - Windows-style `\` separators
            // - ASCII case differences
            // - percent-encoded names
            //
            // `XlsxPackage::part` already implements these lookup semantics.
            if self.part(part).is_none() {
                return Err(crate::streaming::StreamingPatchError::MissingWorksheetPart(
                    part.clone(),
                )
                .into());
            }
        }

        // Route through the full streaming patch pipeline (sharedStrings-aware + recalc-policy
        // aware) rather than directly rewriting worksheet XML parts.
        let input_bytes = self.write_to_bytes()?;
        let mut streaming_patches = Vec::with_capacity(patches.len());
        for patches in patches_by_part.values() {
            streaming_patches.extend_from_slice(patches);
        }

        let mut out = Cursor::new(Vec::new());
        crate::streaming::patch_xlsx_streaming_with_recalc_policy(
            Cursor::new(input_bytes),
            &mut out,
            &streaming_patches,
            policy_on_formula_change,
        )?;
        Ok(out.into_inner())
    }

    /// Parse pivot-related parts (pivot tables + pivot caches) from the package.
    ///
    /// This is a lightweight metadata parser; the raw XML parts remain preserved
    /// verbatim in the package.
    pub fn pivots(&self) -> Result<XlsxPivots, XlsxError> {
        XlsxPivots::parse_from_entries(&self.parts)
    }

    /// Create a streaming reader for a pivot cache records part
    /// (e.g. `xl/pivotCache/pivotCacheRecords1.xml`).
    pub fn pivot_cache_records<'a>(
        &'a self,
        part_name: &str,
    ) -> Result<PivotCacheRecordsReader<'a>, XlsxError> {
        let part_name = part_name.strip_prefix('/').unwrap_or(part_name);
        let bytes = self
            .part(part_name)
            .ok_or_else(|| XlsxError::MissingPart(part_name.to_string()))?;
        Ok(PivotCacheRecordsReader::new(bytes))
    }

    /// Parse all `pivotCacheRecords*.xml` parts in the package into memory.
    ///
    /// Prefer [`Self::pivot_cache_records`] for large caches.
    pub fn pivot_cache_records_all(&self) -> BTreeMap<String, Vec<Vec<PivotCacheValue>>> {
        let mut out = BTreeMap::new();
        for (name, bytes) in &self.parts {
            if name.starts_with("xl/pivotCache/")
                && name.contains("pivotCacheRecords")
                && name.ends_with(".xml")
            {
                let mut reader = PivotCacheRecordsReader::new(bytes);
                out.insert(name.clone(), reader.parse_all_records());
            }
        }
        out
    }

    /// Parse the ordered list of workbook sheets from `xl/workbook.xml`.
    pub fn workbook_sheets(&self) -> Result<Vec<WorkbookSheetInfo>, XlsxError> {
        let workbook_xml = self
            .part("xl/workbook.xml")
            .ok_or_else(|| XlsxError::MissingPart("xl/workbook.xml".to_string()))?;
        let workbook_xml = String::from_utf8(workbook_xml.to_vec())?;
        parse_workbook_sheets(&workbook_xml)
    }

    /// Set the workbook date system (`workbookPr/@date1904`) inside `xl/workbook.xml`.
    ///
    /// This is required for correct serial date interpretation when opening the workbook in Excel
    /// and for aligning formula evaluation semantics (1900 vs 1904) during round-trip edits.
    pub fn set_workbook_date_system(&mut self, date_system: DateSystem) -> Result<(), XlsxError> {
        let workbook_xml = self
            .parts
            .get("xl/workbook.xml")
            .cloned()
            .ok_or_else(|| XlsxError::MissingPart("xl/workbook.xml".to_string()))?;
        let updated = workbook_xml_set_date_system(&workbook_xml, date_system)?;
        self.parts.insert("xl/workbook.xml".to_string(), updated);
        Ok(())
    }

    /// Rewrite the `<sheets>` list in `xl/workbook.xml` to match `sheets`.
    pub fn set_workbook_sheets(&mut self, sheets: &[WorkbookSheetInfo]) -> Result<(), XlsxError> {
        let workbook_xml = self
            .part("xl/workbook.xml")
            .ok_or_else(|| XlsxError::MissingPart("xl/workbook.xml".to_string()))?;
        let workbook_xml = String::from_utf8(workbook_xml.to_vec())?;
        let updated = write_workbook_sheets(&workbook_xml, sheets)?;
        self.set_part("xl/workbook.xml", updated.into_bytes());
        Ok(())
    }

    /// Read a worksheet tab color from a worksheet part (e.g. `xl/worksheets/sheet1.xml`).
    pub fn worksheet_tab_color(&self, worksheet_part: &str) -> Result<Option<TabColor>, XlsxError> {
        let xml = self
            .part(worksheet_part)
            .ok_or_else(|| XlsxError::MissingPart(worksheet_part.to_string()))?;
        let xml = String::from_utf8(xml.to_vec())?;
        parse_sheet_tab_color(&xml)
    }

    /// Update (or remove) a worksheet tab color in a worksheet part.
    pub fn set_worksheet_tab_color(
        &mut self,
        worksheet_part: &str,
        tab_color: Option<&TabColor>,
    ) -> Result<(), XlsxError> {
        let xml = self
            .part(worksheet_part)
            .ok_or_else(|| XlsxError::MissingPart(worksheet_part.to_string()))?;
        let xml = String::from_utf8(xml.to_vec())?;
        let updated = write_sheet_tab_color(&xml, tab_color)?;
        self.set_part(worksheet_part.to_string(), updated.into_bytes());
        Ok(())
    }

    /// Apply a set of cell edits to the existing workbook package.
    ///
    /// This is a high-fidelity edit pipeline intended for "edit existing XLSX/XLSM"
    /// scenarios (e.g. the desktop app save path). Only the affected worksheet parts
    /// (plus `xl/sharedStrings.xml` / `xl/workbook.xml` when required) are rewritten;
    /// every unrelated part is preserved byte-for-byte.
    pub fn apply_cell_patches(&mut self, patches: &WorkbookCellPatches) -> Result<(), XlsxError> {
        self.apply_cell_patches_with_recalc_policy(patches, RecalcPolicy::default())
    }

    /// Apply a set of cell edits to the existing workbook package using the provided
    /// [`RecalcPolicy`].
    pub fn apply_cell_patches_with_recalc_policy(
        &mut self,
        patches: &WorkbookCellPatches,
        recalc_policy: RecalcPolicy,
    ) -> Result<(), XlsxError> {
        apply_cell_patches_to_package(self, patches, recalc_policy)
    }

    /// Apply cell edits that reference `formula_model` style IDs.
    ///
    /// This behaves like [`Self::apply_cell_patches`], but allows patches to specify cell styles
    /// via `style_id` and updates `styles.xml` deterministically when new styles are introduced.
    pub fn apply_cell_patches_with_styles(
        &mut self,
        patches: &WorkbookCellPatches,
        style_table: &StyleTable,
    ) -> Result<(), XlsxError> {
        apply_cell_patches_to_package_with_styles(
            self,
            patches,
            style_table,
            RecalcPolicy::default(),
        )
    }

    /// Remove macro-related parts and relationships from the package.
    ///
    /// This is used when saving a macro-enabled workbook (`.xlsm`) as `.xlsx`.
    pub fn remove_vba_project(&mut self) -> Result<(), XlsxError> {
        crate::macro_strip::strip_macros(&mut self.parts)
    }

    /// Remove macro-related parts and relationships from the package, targeting a specific output
    /// workbook kind.
    ///
    /// This controls how the workbook "main" content type is rewritten in `[Content_Types].xml`
    /// after stripping macros.
    pub fn remove_vba_project_with_kind(
        &mut self,
        target_kind: WorkbookKind,
    ) -> Result<(), XlsxError> {
        crate::macro_strip::strip_macros_with_kind(&mut self.parts, target_kind)
    }
}

fn workbook_xml_set_date_system(
    workbook_xml: &[u8],
    date_system: DateSystem,
) -> Result<Vec<u8>, XlsxError> {
    let has_workbook_pr = workbook_xml
        .windows(b"workbookPr".len())
        .any(|w| w == b"workbookPr");

    let mut reader = XmlReader::from_reader(workbook_xml);
    reader.config_mut().trim_text(false);
    let mut writer = XmlWriter::new(Vec::with_capacity(workbook_xml.len() + 64));

    let mut buf = Vec::new();
    let mut skipping_workbook_pr = false;
    let mut workbook_ns: Option<crate::xml::WorkbookXmlNamespaces> = None;

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Empty(ref e) if local_name(e.name().as_ref()) == b"workbook" => {
                workbook_ns
                    .get_or_insert(crate::xml::workbook_xml_namespaces_from_workbook_start(e)?);

                // Degenerate/self-closing workbook roots can't contain child elements. If we need
                // to force the 1904 date system we must expand `<workbook/>` into
                // `<workbook>...<workbookPr/>...</workbook>`.
                if date_system == DateSystem::V1904 {
                    let workbook_tag_name = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                    writer.write_event(Event::Start(e.to_owned()))?;

                    let tag = workbook_ns
                        .as_ref()
                        .map(|ns| {
                            crate::xml::prefixed_tag(
                                ns.spreadsheetml_prefix.as_deref(),
                                "workbookPr",
                            )
                        })
                        .unwrap_or_else(|| "workbookPr".to_string());
                    let mut wb_pr = BytesStart::new(tag.as_str());
                    wb_pr.push_attribute(("date1904", "1"));
                    writer.write_event(Event::Empty(wb_pr))?;

                    writer.write_event(Event::End(BytesEnd::new(workbook_tag_name.as_str())))?;
                } else {
                    writer.write_event(Event::Empty(e.to_owned()))?;
                }
            }
            Event::Start(ref e) if local_name(e.name().as_ref()) == b"workbook" => {
                workbook_ns
                    .get_or_insert(crate::xml::workbook_xml_namespaces_from_workbook_start(e)?);
                writer.write_event(Event::Start(e.to_owned()))?;
                if date_system == DateSystem::V1904 && !has_workbook_pr {
                    let tag = workbook_ns
                        .as_ref()
                        .map(|ns| {
                            crate::xml::prefixed_tag(
                                ns.spreadsheetml_prefix.as_deref(),
                                "workbookPr",
                            )
                        })
                        .unwrap_or_else(|| "workbookPr".to_string());
                    let mut wb_pr = BytesStart::new(tag.as_str());
                    wb_pr.push_attribute(("date1904", "1"));
                    writer.write_event(Event::Empty(wb_pr))?;
                }
            }
            Event::Empty(ref e) if local_name(e.name().as_ref()) == b"workbookPr" => {
                writer.write_event(Event::Empty(patched_workbook_pr(e, date_system)?))?;
            }
            Event::Start(ref e) if local_name(e.name().as_ref()) == b"workbookPr" => {
                skipping_workbook_pr = true;
                writer.write_event(Event::Empty(patched_workbook_pr(e, date_system)?))?;
            }
            Event::End(ref e)
                if skipping_workbook_pr && local_name(e.name().as_ref()) == b"workbookPr" =>
            {
                skipping_workbook_pr = false;
            }
            Event::Eof => break,
            ev if skipping_workbook_pr => drop(ev),
            other => writer.write_event(other.into_owned())?,
        }

        buf.clear();
    }

    Ok(writer.into_inner())
}

fn patched_workbook_pr(
    e: &BytesStart<'_>,
    date_system: DateSystem,
) -> Result<BytesStart<'static>, XlsxError> {
    let name = e.name();
    let mut wb_pr = BytesStart::new(std::str::from_utf8(name.as_ref()).unwrap_or("workbookPr"));
    let mut had_date1904 = false;
    for attr in e.attributes() {
        let attr = attr?;
        if attr.key.as_ref() == b"date1904" {
            had_date1904 = true;
            continue;
        }
        wb_pr.push_attribute((attr.key.as_ref(), attr.value.as_ref()));
    }

    match date_system {
        DateSystem::V1900 => {
            if had_date1904 {
                wb_pr.push_attribute(("date1904", "0"));
            }
        }
        DateSystem::V1904 => wb_pr.push_attribute(("date1904", "1")),
    }

    Ok(wb_pr.into_owned())
}

fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|&b| b == b':').next().unwrap_or(name)
}

fn prefixed_tag(container_name: &[u8], local: &str) -> String {
    match container_name.iter().position(|&b| b == b':') {
        Some(idx) => {
            let prefix = std::str::from_utf8(&container_name[..idx]).unwrap_or_default();
            format!("{prefix}:{local}")
        }
        None => local.to_string(),
    }
}

fn ensure_workbook_content_type(
    parts: &mut BTreeMap<String, Vec<u8>>,
    workbook_content_type: &str,
) -> Result<(), XlsxError> {
    let ct_name = "[Content_Types].xml";
    let ct_key = if parts.contains_key(ct_name) {
        ct_name.to_string()
    } else {
        let Some(found) = parts
            .keys()
            .find(|name| crate::zip_util::zip_part_names_equivalent(name.as_str(), ct_name))
            .cloned()
        else {
            return Ok(());
        };
        found
    };

    let Some(existing) = parts.get(&ct_key).cloned() else {
        return Ok(());
    };

    let mut reader = XmlReader::from_reader(existing.as_slice());
    reader.config_mut().trim_text(false);
    let mut writer = XmlWriter::new(Vec::with_capacity(existing.len() + 128));
    let mut buf = Vec::new();

    let mut changed = false;
    let mut found = false;
    let mut skip_depth = 0usize;
    let mut override_tag_name: Option<String> = None;

    loop {
        let ev = reader.read_event_into(&mut buf)?;

        if skip_depth > 0 {
            match ev {
                Event::Start(_) => skip_depth += 1,
                Event::End(_) => skip_depth -= 1,
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
            continue;
        }

        match ev {
            Event::Eof => break,
            Event::Empty(ref e) if crate::openxml::local_name(e.name().as_ref()) == b"Override" => {
                if override_tag_name.is_none() {
                    override_tag_name =
                        Some(String::from_utf8_lossy(e.name().as_ref()).into_owned());
                }
                let (is_workbook, updated) = patched_workbook_override(e, workbook_content_type)?;
                if is_workbook {
                    found = true;
                }
                if let Some(updated) = updated {
                    writer.write_event(Event::Empty(updated))?;
                    changed = true;
                } else {
                    writer.write_event(Event::Empty(e.to_owned()))?;
                }
            }
            Event::Start(ref e) if crate::openxml::local_name(e.name().as_ref()) == b"Override" => {
                if override_tag_name.is_none() {
                    override_tag_name =
                        Some(String::from_utf8_lossy(e.name().as_ref()).into_owned());
                }
                let (is_workbook, updated) = patched_workbook_override(e, workbook_content_type)?;
                if is_workbook {
                    found = true;
                }
                if let Some(updated) = updated {
                    writer.write_event(Event::Empty(updated))?;
                    changed = true;
                    // Skip through the matching </Override>.
                    skip_depth = 1;
                } else {
                    writer.write_event(Event::Start(e.to_owned()))?;
                }
            }
            Event::End(e) if crate::openxml::local_name(e.name().as_ref()) == b"Types" => {
                if !found {
                    // No workbook override found; insert one before `</Types>`.
                    changed = true;
                    let override_tag_name = override_tag_name
                        .clone()
                        .unwrap_or_else(|| prefixed_tag(e.name().as_ref(), "Override"));
                    let mut override_el = BytesStart::new(override_tag_name.as_str());
                    override_el.push_attribute(("PartName", "/xl/workbook.xml"));
                    override_el.push_attribute(("ContentType", workbook_content_type));
                    writer.write_event(Event::Empty(override_el))?;
                }
                writer.write_event(Event::End(e))?;
            }
            Event::Empty(e) if crate::openxml::local_name(e.name().as_ref()) == b"Types" => {
                // Degenerate case: a self-closing `<Types/>` root. Expand it so we can inject
                // the required workbook override.
                if !found {
                    changed = true;
                    let types_tag_name = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                    let override_tag_name = override_tag_name
                        .clone()
                        .unwrap_or_else(|| prefixed_tag(types_tag_name.as_bytes(), "Override"));

                    writer.write_event(Event::Start(e))?;

                    let mut override_el = BytesStart::new(override_tag_name.as_str());
                    override_el.push_attribute(("PartName", "/xl/workbook.xml"));
                    override_el.push_attribute(("ContentType", workbook_content_type));
                    writer.write_event(Event::Empty(override_el))?;

                    writer.write_event(Event::End(BytesEnd::new(types_tag_name.as_str())))?;
                    found = true;
                } else {
                    writer.write_event(Event::Empty(e))?;
                }
            }
            other => writer.write_event(other.into_owned())?,
        }

        buf.clear();
    }

    if changed {
        parts.insert(ct_key, writer.into_inner());
    }
    Ok(())
}

pub(crate) fn ensure_content_types_default(
    parts: &mut BTreeMap<String, Vec<u8>>,
    ext: &str,
    content_type: &str,
) -> Result<(), XlsxError> {
    let ct_name = "[Content_Types].xml";
    let ct_key = if parts.contains_key(ct_name) {
        ct_name.to_string()
    } else {
        let Some(found) = parts
            .keys()
            .find(|name| crate::zip_util::zip_part_names_equivalent(name.as_str(), ct_name))
            .cloned()
        else {
            // Match `ensure_workbook_content_type` behavior: we don't synthesize a full content types
            // file for existing packages.
            return Ok(());
        };
        found
    };

    let Some(existing) = parts.get(&ct_key).cloned() else {
        // Match `ensure_workbook_content_type` behavior: we don't synthesize a full content types
        // file for existing packages.
        return Ok(());
    };

    let normalized_ext = crate::ascii::normalize_extension_ascii_lowercase(ext.trim().trim_start_matches('.'));
    let normalized_ext = normalized_ext.as_ref();
    if normalized_ext.is_empty() {
        return Ok(());
    }

    let mut reader = XmlReader::from_reader(existing.as_slice());
    reader.config_mut().trim_text(false);
    let mut writer = XmlWriter::new(Vec::with_capacity(existing.len() + 128));
    let mut buf = Vec::new();

    let mut default_tag_name: Option<String> = None;
    let mut found = false;
    let mut changed = false;

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Start(e) if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Default") => {
                if default_tag_name.is_none() {
                    default_tag_name =
                        Some(String::from_utf8_lossy(e.name().as_ref()).into_owned());
                }
                for attr in e.attributes().with_checks(false) {
                    let attr = attr?;
                    if local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"Extension") {
                        let ext = attr.unescape_value()?.into_owned();
                        if ext.trim().eq_ignore_ascii_case(normalized_ext) {
                            found = true;
                            break;
                        }
                    }
                }
                writer.write_event(Event::Start(e))?;
            }
            Event::Empty(e) if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Default") => {
                if default_tag_name.is_none() {
                    default_tag_name =
                        Some(String::from_utf8_lossy(e.name().as_ref()).into_owned());
                }
                for attr in e.attributes().with_checks(false) {
                    let attr = attr?;
                    if local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"Extension") {
                        let ext = attr.unescape_value()?.into_owned();
                        if ext.trim().eq_ignore_ascii_case(normalized_ext) {
                            found = true;
                            break;
                        }
                    }
                }
                writer.write_event(Event::Empty(e))?;
            }
            Event::End(e) if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Types") => {
                if !found {
                    changed = true;
                    let default_tag_name = default_tag_name
                        .clone()
                        .unwrap_or_else(|| prefixed_tag(e.name().as_ref(), "Default"));
                    let mut default_el = BytesStart::new(default_tag_name.as_str());
                    default_el.push_attribute(("Extension", normalized_ext));
                    default_el.push_attribute(("ContentType", content_type));
                    writer.write_event(Event::Empty(default_el))?;
                }
                writer.write_event(Event::End(e))?;
            }
            Event::Empty(e) if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Types") => {
                // Degenerate case: a self-closing `<Types/>` root. Expand it so we can inject the
                // required Default.
                if !found {
                    changed = true;
                    let types_tag_name = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                    let default_tag_name = default_tag_name
                        .clone()
                        .unwrap_or_else(|| prefixed_tag(types_tag_name.as_bytes(), "Default"));

                    writer.write_event(Event::Start(e))?;

                    let mut default_el = BytesStart::new(default_tag_name.as_str());
                    default_el.push_attribute(("Extension", normalized_ext));
                    default_el.push_attribute(("ContentType", content_type));
                    writer.write_event(Event::Empty(default_el))?;

                    writer.write_event(Event::End(BytesEnd::new(types_tag_name.as_str())))?;
                } else {
                    writer.write_event(Event::Empty(e))?;
                }
            }
            Event::Eof => break,
            other => writer.write_event(other)?,
        }

        buf.clear();
    }

    if changed {
        parts.insert(ct_key, writer.into_inner());
    }
    Ok(())
}

fn patched_workbook_override(
    e: &BytesStart<'_>,
    workbook_content_type: &str,
) -> Result<(bool, Option<BytesStart<'static>>), XlsxError> {
    let mut part_name = None;
    let mut existing_content_type = None;

    for attr in e.attributes().with_checks(false) {
        let attr = attr?;
        match crate::openxml::local_name(attr.key.as_ref()) {
            b"PartName" => part_name = Some(attr.unescape_value()?.into_owned()),
            b"ContentType" => existing_content_type = Some(attr.unescape_value()?.into_owned()),
            _ => {}
        }
    }

    let Some(part_name) = part_name else {
        return Ok((false, None));
    };

    let normalized = part_name.strip_prefix('/').unwrap_or(part_name.as_str());
    if normalized != "xl/workbook.xml" {
        return Ok((false, None));
    }

    if existing_content_type.as_deref() == Some(workbook_content_type) {
        return Ok((true, None));
    }

    let tag_name = e.name();
    let tag_name = std::str::from_utf8(tag_name.as_ref()).unwrap_or("Override");
    let mut updated = BytesStart::new(tag_name);
    updated.push_attribute(("PartName", "/xl/workbook.xml"));
    updated.push_attribute(("ContentType", workbook_content_type));
    Ok((true, Some(updated.into_owned())))
}

impl crate::XlsxDocument {
    /// Detect whether the document contains any macro-capable content (VBA, XLM macrosheets, or
    /// legacy dialog sheets).
    pub fn macro_presence(&self) -> MacroPresence {
        let mut presence = MacroPresence {
            has_vba: false,
            has_xlm_macrosheets: false,
            has_dialog_sheets: false,
        };

        for name in self.parts().keys() {
            let name = name.as_str();
            let key = crate::zip_util::zip_part_name_lookup_key(name);
            if key == b"xl/vbaproject.bin" {
                presence.has_vba = true;
            }
            if key.starts_with(b"xl/macrosheets/") {
                presence.has_xlm_macrosheets = true;
            }
            if key.starts_with(b"xl/dialogsheets/") {
                presence.has_dialog_sheets = true;
            }

            if presence.has_vba && presence.has_xlm_macrosheets && presence.has_dialog_sheets {
                break;
            }
        }

        presence
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use pretty_assertions::assert_eq;
    use roxmltree::Document;
    use std::collections::{BTreeMap, HashSet};

    fn build_package(files: &[(&str, &[u8])]) -> Vec<u8> {
        let cursor = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(cursor);
        let options = zip::write::FileOptions::<()>::default()
            .compression_method(zip::CompressionMethod::Deflated);

        for (name, bytes) in files {
            zip.start_file(*name, options).unwrap();
            zip.write_all(bytes).unwrap();
        }

        zip.finish().unwrap().into_inner()
    }

    #[test]
    fn from_bytes_limited_rejects_packages_exceeding_total_limit() {
        let bytes = build_package(&[("xl/a.xml", b"123456"), ("xl/b.xml", b"abcdef")]);

        let limits = XlsxPackageLimits {
            max_part_bytes: 10,
            max_total_bytes: 10,
        };

        match XlsxPackage::from_bytes_limited(&bytes, limits) {
            Err(XlsxError::PackageTooLarge { total, max }) => {
                assert_eq!(max, 10);
                assert!(
                    total > max,
                    "expected reported total ({total}) to exceed max ({max})"
                );
            }
            other => panic!("expected PackageTooLarge error, got {other:?}"),
        }
    }

    #[test]
    fn from_bytes_limited_rejects_parts_exceeding_part_limit() {
        let bytes = build_package(&[("xl/too-big.bin", b"0123456789A")]); // 11 bytes

        let limits = XlsxPackageLimits {
            max_part_bytes: 10,
            max_total_bytes: 100,
        };

        match XlsxPackage::from_bytes_limited(&bytes, limits) {
            Err(XlsxError::PartTooLarge { part, size, max }) => {
                assert_eq!(part, "xl/too-big.bin");
                assert_eq!(size, 11);
                assert_eq!(max, 10);
            }
            other => panic!("expected PartTooLarge error, got {other:?}"),
        }
    }

    #[test]
    fn read_part_from_reader_limited_rejects_oversized_part() {
        let max_bytes = 10u64;
        let payload = vec![0u8; (max_bytes + 1) as usize];
        let bytes = build_package(&[("xl/vbaProject.bin", payload.as_slice())]);

        let err = read_part_from_reader_limited(Cursor::new(bytes), "xl/vbaProject.bin", max_bytes)
            .unwrap_err();

        match err {
            XlsxError::PartTooLarge { part, size, max } => {
                assert_eq!(part, "xl/vbaProject.bin");
                assert_eq!(size, max_bytes + 1);
                assert_eq!(max, max_bytes);
            }
            other => panic!("expected PartTooLarge error, got {other:?}"),
        }
    }

    #[test]
    fn read_part_from_reader_limited_reads_part_within_limit() {
        let max_bytes = 10u64;
        let payload = b"0123456789";
        let bytes = build_package(&[("xl/vbaProject.bin", payload.as_slice())]);

        let extracted =
            read_part_from_reader_limited(Cursor::new(bytes), "xl/vbaProject.bin", max_bytes)
                .expect("read part")
                .expect("part exists");

        assert_eq!(extracted, payload);
    }

    #[test]
    fn read_part_from_reader_supports_backslash_separated_entry_names() {
        let workbook = b"workbook-bytes";
        let bytes = build_package(&[("xl\\workbook.xml", workbook.as_slice())]);

        let extracted = read_part_from_reader(Cursor::new(bytes), "xl/workbook.xml")
            .expect("read workbook.xml");

        assert_eq!(extracted, Some(workbook.to_vec()));
    }

    #[test]
    fn read_part_from_reader_supports_leading_slash_backslash_entry_names() {
        let workbook = b"workbook-bytes";
        let bytes = build_package(&[("/xl\\workbook.xml", workbook.as_slice())]);

        let extracted = read_part_from_reader(Cursor::new(bytes.clone()), "xl/workbook.xml")
            .expect("read workbook.xml");
        assert_eq!(extracted, Some(workbook.to_vec()));

        let extracted = read_part_from_reader(Cursor::new(bytes), "/xl/workbook.xml")
            .expect("read workbook.xml with leading slash");
        assert_eq!(extracted, Some(workbook.to_vec()));
    }

    #[test]
    fn macro_presence_tolerates_backslashes_and_case() {
        let bytes = build_package(&[
            ("XL\\vbaProject.bin", b"vba"),
            ("XL\\macrosheets\\sheet1.xml", b"macro"),
            ("XL\\dialogsheets\\sheet1.xml", b"dialog"),
        ]);

        let pkg = XlsxPackage::from_bytes(&bytes).expect("read test pkg");
        let presence = pkg.macro_presence();
        assert_eq!(
            presence,
            MacroPresence {
                has_vba: true,
                has_xlm_macrosheets: true,
                has_dialog_sheets: true,
            }
        );

        let mut doc = crate::XlsxDocument::new(formula_model::Workbook::new());
        doc.parts
            .insert("XL\\vbaProject.bin".to_string(), b"vba".to_vec());
        doc.parts
            .insert("XL\\macrosheets\\sheet1.xml".to_string(), b"macro".to_vec());
        doc.parts.insert(
            "XL\\dialogsheets\\sheet1.xml".to_string(),
            b"dialog".to_vec(),
        );
        assert_eq!(
            doc.macro_presence(),
            MacroPresence {
                has_vba: true,
                has_xlm_macrosheets: true,
                has_dialog_sheets: true,
            }
        );
    }

    #[test]
    fn apply_cell_patches_to_bytes_tolerates_backslash_worksheet_part_names() {
        let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#;

        let bytes = build_package(&[("xl\\worksheets\\sheet1.xml", worksheet_xml.as_bytes())]);
        let pkg = XlsxPackage::from_bytes(&bytes).expect("read test pkg");

        let patched = pkg
            .apply_cell_patches_to_bytes(&[CellPatch::for_worksheet_part(
                "xl/worksheets/sheet1.xml",
                CellRef::new(0, 0),
                CellValue::Number(42.0),
                None,
            )])
            .expect("apply cell patches");

        let mut zip = zip::ZipArchive::new(Cursor::new(patched)).expect("open patched zip");
        let mut file = zip
            .by_name("xl\\worksheets\\sheet1.xml")
            .expect("worksheet part preserved with backslashes");
        let mut out = Vec::new();
        file.read_to_end(&mut out).expect("read patched sheet");

        let xml = String::from_utf8(out).expect("sheet xml is utf-8");
        assert!(xml.contains(r#"r="A1""#), "expected cell A1 in {xml}");
        assert!(xml.contains("<v>42</v>"), "expected value 42 in {xml}");
    }

    #[test]
    fn extract_cell_images_resolves_media_relationships() {
        let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/cellImages.xml" ContentType="application/xml"/>
</Types>"#;

        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets/>
</workbook>"#;

        let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdCellImages" Type="http://example.com/relationships/unknown" Target="cellImages.xml"/>
</Relationships>"#;

        let cell_images_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cellImages xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <cellImage>
    <pic>
      <blipFill>
        <blip r:embed="rId1"/>
      </blipFill>
    </pic>
  </cellImage>
</cellImages>"#;

        let cell_images_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png#frag"/>
</Relationships>"#;

        let image_bytes = b"known-image-bytes";

        let bytes = build_package(&[
            ("[Content_Types].xml", content_types.as_bytes()),
            ("xl/workbook.xml", workbook_xml.as_bytes()),
            ("xl/_rels/workbook.xml.rels", workbook_rels.as_bytes()),
            ("xl/cellImages.xml", cell_images_xml.as_bytes()),
            ("xl/_rels/cellImages.xml.rels", cell_images_rels.as_bytes()),
            ("xl/media/image1.png", image_bytes.as_slice()),
        ]);

        let pkg = XlsxPackage::from_bytes(&bytes).expect("read test pkg");
        let extracted = pkg.extract_cell_images().expect("extract cell images");

        assert_eq!(
            extracted,
            vec![("xl/media/image1.png".to_string(), image_bytes.to_vec())]
        );
    }

    #[test]
    fn extract_cell_images_supports_cellimage_rid_and_parent_media_targets() {
        let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/cellImages.xml" ContentType="application/xml"/>
</Types>"#;

        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets/>
</workbook>"#;

        // Intentionally use an unknown relationship type; detection should be based on the target.
        let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdCellImages" Type="http://example.com/relationships/unknown" Target="cellImages.xml"/>
</Relationships>"#;

        // Some producers encode the relationship ID directly on `<cellImage r:id="...">`.
        let cell_images_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cellImages xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <cellImage r:id="rId1"/>
</cellImages>"#;

        // Some producers emit `../media/...` for workbook-level parts; ensure this resolves to `xl/media/...`.
        let cell_images_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png#frag"/>
</Relationships>"#;

        let image_bytes = b"known-image-bytes";

        let bytes = build_package(&[
            ("[Content_Types].xml", content_types.as_bytes()),
            ("xl/workbook.xml", workbook_xml.as_bytes()),
            ("xl/_rels/workbook.xml.rels", workbook_rels.as_bytes()),
            ("xl/cellImages.xml", cell_images_xml.as_bytes()),
            ("xl/_rels/cellImages.xml.rels", cell_images_rels.as_bytes()),
            ("xl/media/image1.png", image_bytes.as_slice()),
        ]);

        let pkg = XlsxPackage::from_bytes(&bytes).expect("read test pkg");
        let extracted = pkg.extract_cell_images().expect("extract cell images");

        assert_eq!(
            extracted,
            vec![("xl/media/image1.png".to_string(), image_bytes.to_vec())]
        );
    }

    #[test]
    fn extract_cell_images_discovers_cell_images_part_without_workbook_rels() {
        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets/>
</workbook>"#;

        // Use a numeric suffix to ensure we discover parts beyond the canonical `cellImages.xml`.
        let cell_images_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cellImages xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <cellImage r:id="rId1"/>
</cellImages>"#;

        let cell_images_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>"#;

        let image_bytes = b"known-image-bytes";

        let bytes = build_package(&[
            ("xl/workbook.xml", workbook_xml.as_bytes()),
            ("xl/cellImages1.xml", cell_images_xml.as_bytes()),
            ("xl/_rels/cellImages1.xml.rels", cell_images_rels.as_bytes()),
            ("xl/media/image1.png", image_bytes.as_slice()),
        ]);

        let pkg = XlsxPackage::from_bytes(&bytes).expect("read test pkg");
        let extracted = pkg.extract_cell_images().expect("extract cell images");

        assert_eq!(
            extracted,
            vec![("xl/media/image1.png".to_string(), image_bytes.to_vec())]
        );
    }

    fn build_minimal_package() -> Vec<u8> {
        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

        let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></worksheet>"#;

        let cursor = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(cursor);
        let options = zip::write::FileOptions::<()>::default()
            .compression_method(zip::CompressionMethod::Deflated);

        zip.start_file("xl/workbook.xml", options).unwrap();
        zip.write_all(workbook_xml.as_bytes()).unwrap();

        zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
        zip.write_all(worksheet_xml.as_bytes()).unwrap();

        zip.finish().unwrap().into_inner()
    }

    #[test]
    fn set_workbook_date_system_expands_prefixed_self_closing_workbook_root() {
        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>"#;

        let updated = workbook_xml_set_date_system(workbook_xml.as_bytes(), DateSystem::V1904)
            .expect("set date system");
        let updated = std::str::from_utf8(&updated).expect("utf8");
        let doc = Document::parse(updated).expect("updated workbook.xml parses");

        assert!(
            updated
                .contains(r#"xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main""#),
            "expected output to preserve SpreadsheetML namespace declaration, got:\n{updated}"
        );
        assert!(
            updated.contains(
                r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#
            ),
            "expected output to preserve relationships namespace declaration, got:\n{updated}"
        );
        assert!(
            updated.contains("<x:workbookPr"),
            "expected output to contain a prefixed workbookPr, got:\n{updated}"
        );
        assert!(
            updated.contains("</x:workbook>"),
            "expected output to expand the workbook root, got:\n{updated}"
        );

        let spreadsheetml = crate::xml::SPREADSHEETML_NS;
        let workbook_pr: Vec<_> = doc
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "workbookPr")
            .collect();
        assert_eq!(workbook_pr.len(), 1);
        assert_eq!(workbook_pr[0].tag_name().namespace(), Some(spreadsheetml));
        assert_eq!(workbook_pr[0].attribute("date1904"), Some("1"));
    }

    #[test]
    fn set_workbook_date_system_expands_default_ns_self_closing_workbook_root() {
        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;

        let updated = workbook_xml_set_date_system(workbook_xml.as_bytes(), DateSystem::V1904)
            .expect("set date system");
        let updated = std::str::from_utf8(&updated).expect("utf8");
        let doc = Document::parse(updated).expect("updated workbook.xml parses");

        assert!(
            updated.contains("<workbookPr"),
            "expected output to contain an unprefixed workbookPr, got:\n{updated}"
        );
        assert!(
            updated.contains("</workbook>"),
            "expected output to expand the workbook root, got:\n{updated}"
        );

        let spreadsheetml = crate::xml::SPREADSHEETML_NS;
        let workbook_pr: Vec<_> = doc
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "workbookPr")
            .collect();
        assert_eq!(workbook_pr.len(), 1);
        assert_eq!(workbook_pr[0].tag_name().namespace(), Some(spreadsheetml));
        assert_eq!(workbook_pr[0].attribute("date1904"), Some("1"));
    }

    fn load_fixture() -> Vec<u8> {
        std::fs::read(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../fixtures/xlsx/macros/basic.xlsm"
        ))
        .expect("fixture exists")
    }

    #[test]
    fn ensure_content_types_default_inserts_png() {
        let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
</Types>"#;

        let mut parts = BTreeMap::new();
        parts.insert(
            "[Content_Types].xml".to_string(),
            content_types.as_bytes().to_vec(),
        );

        ensure_content_types_default(&mut parts, "png", "image/png").expect("ensure png default");

        let updated = std::str::from_utf8(parts.get("[Content_Types].xml").unwrap()).expect("utf8");
        let doc = Document::parse(updated).expect("parse content types");
        assert!(
            doc.descendants().any(|n| {
                n.is_element()
                    && n.tag_name().name() == "Default"
                    && n.attribute("Extension") == Some("png")
                    && n.attribute("ContentType") == Some("image/png")
            }),
            "expected png Default to be inserted, got:\n{updated}"
        );
    }

    #[test]
    fn ensure_content_types_default_updates_slashed_content_types_key_in_place() {
        let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
</Types>"#;

        let mut parts = BTreeMap::new();
        parts.insert(
            "/[Content_Types].xml".to_string(),
            content_types.as_bytes().to_vec(),
        );

        ensure_content_types_default(&mut parts, "png", "image/png").expect("ensure png default");

        assert!(
            parts.contains_key("/[Content_Types].xml"),
            "expected content types to be patched in-place (preserving leading slash key)"
        );
        assert!(
            !parts.contains_key("[Content_Types].xml"),
            "expected ensure_content_types_default to not create a new canonical content types key"
        );

        let updated =
            std::str::from_utf8(parts.get("/[Content_Types].xml").unwrap()).expect("utf8");
        let doc = Document::parse(updated).expect("parse content types");
        assert!(
            doc.descendants().any(|n| {
                n.is_element()
                    && n.tag_name().name() == "Default"
                    && n.attribute("Extension") == Some("png")
                    && n.attribute("ContentType") == Some("image/png")
            }),
            "expected png Default to be inserted, got:\n{updated}"
        );
    }

    #[test]
    fn ensure_workbook_content_type_updates_slashed_content_types_key_in_place() {
        let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>"#;

        let mut parts = BTreeMap::new();
        parts.insert(
            "/[Content_Types].xml".to_string(),
            content_types.as_bytes().to_vec(),
        );

        ensure_workbook_content_type(
            &mut parts,
            WorkbookKind::MacroEnabledWorkbook.workbook_content_type(),
        )
        .expect("ensure workbook content type");

        assert!(
            parts.contains_key("/[Content_Types].xml"),
            "expected content types to be patched in-place (preserving leading slash key)"
        );
        assert!(
            !parts.contains_key("[Content_Types].xml"),
            "expected ensure_workbook_content_type to not create a new canonical content types key"
        );

        let updated =
            std::str::from_utf8(parts.get("/[Content_Types].xml").unwrap()).expect("utf8");
        assert!(
            updated.contains("application/vnd.ms-excel.sheet.macroEnabled.main+xml"),
            "expected workbook content type to be updated, got:\n{updated}"
        );
    }

    #[test]
    fn ensure_content_types_default_does_not_duplicate_existing_default() {
        let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
</Types>"#;

        let mut parts = BTreeMap::new();
        parts.insert(
            "[Content_Types].xml".to_string(),
            content_types.as_bytes().to_vec(),
        );

        ensure_content_types_default(&mut parts, "png", "image/png").expect("ensure png default");
        ensure_content_types_default(&mut parts, "png", "image/png").expect("ensure png default");

        let updated = std::str::from_utf8(parts.get("[Content_Types].xml").unwrap()).expect("utf8");
        let doc = Document::parse(updated).expect("parse content types");
        let count = doc
            .descendants()
            .filter(|n| {
                n.is_element()
                    && n.tag_name().name() == "Default"
                    && n.attribute("Extension") == Some("png")
            })
            .count();
        assert_eq!(
            count, 1,
            "expected png Default to not duplicate, got:\n{updated}"
        );
    }

    #[test]
    fn ensure_content_types_default_preserves_prefix_when_root_is_prefixed() {
        let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types">
  <ct:Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <ct:Default Extension="xml" ContentType="application/xml"/>
</ct:Types>"#;

        let mut parts = BTreeMap::new();
        parts.insert(
            "[Content_Types].xml".to_string(),
            content_types.as_bytes().to_vec(),
        );

        ensure_content_types_default(&mut parts, "png", "image/png").expect("ensure png default");

        let updated = std::str::from_utf8(parts.get("[Content_Types].xml").unwrap()).expect("utf8");
        let doc = Document::parse(updated).expect("parse content types");

        let ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types";
        let node = doc
            .descendants()
            .find(|n| {
                n.is_element()
                    && n.tag_name().name() == "Default"
                    && n.attribute("Extension") == Some("png")
            })
            .expect("inserted Default");
        assert_eq!(node.tag_name().namespace(), Some(ct_ns));
    }

    #[test]
    fn write_to_bytes_inserts_content_types_defaults_for_emf_wmf_svg_media() {
        let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
</Types>"#;

        let bytes = build_package(&[
            ("[Content_Types].xml", content_types.as_bytes()),
            ("xl/media/image1.emf", b"emf-bytes"),
            ("xl/media/image2.wmf", b"wmf-bytes"),
            ("xl/media/image3.svg", br#"<svg xmlns="http://www.w3.org/2000/svg"></svg>"#),
        ]);

        let pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
        let written = pkg.write_to_bytes().expect("write pkg");
        let pkg2 = XlsxPackage::from_bytes(&written).expect("read pkg2");

        let ct = std::str::from_utf8(pkg2.part("[Content_Types].xml").unwrap()).unwrap();
        let doc = Document::parse(ct).expect("parse content types");

        for (ext, content_type) in [
            ("emf", "image/x-emf"),
            ("wmf", "image/x-wmf"),
            ("svg", "image/svg+xml"),
        ] {
            assert!(
                doc.descendants().any(|n| {
                    n.is_element()
                        && n.tag_name().name() == "Default"
                        && n.attribute("Extension") == Some(ext)
                        && n.attribute("ContentType") == Some(content_type)
                }),
                "expected {ext} Default to be inserted, got:\n{ct}"
            );
        }
    }

    #[test]
    fn enforce_workbook_kind_inserts_prefixed_workbook_override_when_missing() {
        let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types">
  <ct:Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <ct:Default Extension="xml" ContentType="application/xml"/>
</ct:Types>"#;

        let bytes = build_package(&[("[Content_Types].xml", content_types.as_bytes())]);
        let mut pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
        pkg.enforce_workbook_kind(WorkbookKind::MacroEnabledWorkbook)
            .expect("enforce workbook kind");

        let updated = std::str::from_utf8(pkg.part("[Content_Types].xml").unwrap()).unwrap();
        let doc = Document::parse(updated).expect("parse content types");
        let ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types";

        let node = doc
            .descendants()
            .find(|n| {
                n.is_element()
                    && n.tag_name().name() == "Override"
                    && n.attribute("PartName") == Some("/xl/workbook.xml")
            })
            .expect("inserted workbook override");
        assert_eq!(node.tag_name().namespace(), Some(ct_ns));
        assert_eq!(
            node.attribute("ContentType"),
            Some("application/vnd.ms-excel.sheet.macroEnabled.main+xml")
        );
        assert!(
            updated.contains("<ct:Override"),
            "expected inserted Override to preserve root prefix, got:\n{updated}"
        );
        assert!(
            !updated.contains("<Override PartName=\"/xl/workbook.xml\""),
            "should not introduce unprefixed workbook Override, got:\n{updated}"
        );
    }

    #[test]
    fn enforce_workbook_kind_preserves_prefix_when_patching_existing_override() {
        let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types">
  <ct:Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <ct:Default Extension="xml" ContentType="application/xml"/>
  <ct:Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</ct:Types>"#;

        let bytes = build_package(&[("[Content_Types].xml", content_types.as_bytes())]);
        let mut pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
        pkg.enforce_workbook_kind(WorkbookKind::MacroEnabledWorkbook)
            .expect("enforce workbook kind");

        let updated = std::str::from_utf8(pkg.part("[Content_Types].xml").unwrap()).unwrap();
        let doc = Document::parse(updated).expect("parse content types");
        let ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types";

        let node = doc
            .descendants()
            .find(|n| {
                n.is_element()
                    && n.tag_name().name() == "Override"
                    && n.attribute("PartName") == Some("/xl/workbook.xml")
            })
            .expect("workbook override");
        assert_eq!(node.tag_name().namespace(), Some(ct_ns));
        assert_eq!(
            node.attribute("ContentType"),
            Some("application/vnd.ms-excel.sheet.macroEnabled.main+xml")
        );
        assert!(
            updated.contains("<ct:Override"),
            "expected patched Override to preserve prefix, got:\n{updated}"
        );
        assert!(
            !updated.contains("<Override PartName=\"/xl/workbook.xml\""),
            "should not rewrite workbook Override without a prefix in a prefix-only document, got:\n{updated}"
        );
    }

    #[test]
    fn round_trip_preserves_vba_project_bin_bytes() {
        let fixture = load_fixture();
        let pkg = XlsxPackage::from_bytes(&fixture).expect("read pkg");

        let original_bin = pkg
            .vba_project_bin()
            .expect("vbaProject.bin present")
            .to_vec();

        let written = pkg.write_to_bytes().expect("write pkg");
        let pkg2 = XlsxPackage::from_bytes(&written).expect("read pkg2");
        let roundtrip_bin = pkg2
            .vba_project_bin()
            .expect("vbaProject.bin present in roundtrip");

        assert_eq!(original_bin, roundtrip_bin);
    }

    #[test]
    fn ensures_content_types_and_relationships_for_xlsm() {
        let fixture = load_fixture();
        let pkg = XlsxPackage::from_bytes(&fixture).expect("read pkg");
        let written = pkg.write_to_bytes().expect("write");
        let pkg2 = XlsxPackage::from_bytes(&written).expect("read");

        let ct = std::str::from_utf8(pkg2.part("[Content_Types].xml").unwrap()).unwrap();
        assert!(ct.contains("application/vnd.ms-office.vbaProject"));

        let rels = std::str::from_utf8(pkg2.part("xl/_rels/workbook.xml.rels").unwrap()).unwrap();
        assert!(rels.contains("http://schemas.microsoft.com/office/2006/relationships/vbaProject"));
    }

    fn build_macro_package_needing_repair() -> Vec<u8> {
        let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
            PartName="/xl/workbook.xml"/>
</Types >"#;

        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

        let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></worksheet>"#;

        let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Target="worksheets/sheet1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Id = "rId1"/>
</Relationships >"#;

        build_package(&[
            ("[Content_Types].xml", content_types.as_bytes()),
            ("xl/workbook.xml", workbook_xml.as_bytes()),
            ("xl/_rels/workbook.xml.rels", workbook_rels.as_bytes()),
            ("xl/worksheets/sheet1.xml", worksheet_xml.as_bytes()),
            ("xl/vbaProject.bin", b"fake-vba-project"),
        ])
    }

    #[test]
    fn repairs_xlsm_content_types_and_workbook_rels_with_unusual_formatting() {
        let bytes = build_macro_package_needing_repair();
        let pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");

        let written = pkg.write_to_bytes().expect("write repaired pkg");
        let pkg2 = XlsxPackage::from_bytes(&written).expect("read pkg2");

        let ct = std::str::from_utf8(pkg2.part("[Content_Types].xml").unwrap()).unwrap();
        let doc = Document::parse(ct).expect("parse content types");

        let workbook_ct = doc
            .descendants()
            .find(|n| {
                n.is_element()
                    && n.tag_name().name() == "Override"
                    && n.attribute("PartName") == Some("/xl/workbook.xml")
            })
            .and_then(|n| n.attribute("ContentType"));
        assert_eq!(
            workbook_ct,
            Some("application/vnd.ms-excel.sheet.macroEnabled.main+xml")
        );

        let vba_ct = doc
            .descendants()
            .find(|n| {
                n.is_element()
                    && n.tag_name().name() == "Override"
                    && n.attribute("PartName") == Some("/xl/vbaProject.bin")
            })
            .and_then(|n| n.attribute("ContentType"));
        assert_eq!(vba_ct, Some("application/vnd.ms-office.vbaProject"));

        let rels_bytes = pkg2.part("xl/_rels/workbook.xml.rels").unwrap();
        let rels = crate::openxml::parse_relationships(rels_bytes).expect("parse rels");
        let vba_rel = rels
            .iter()
            .find(|rel| {
                rel.type_uri == "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
            })
            .expect("expected workbook.xml.rels to contain a vbaProject relationship");
        assert_eq!(vba_rel.target, "vbaProject.bin");
        assert_eq!(vba_rel.id, "rId2");

        let mut ids = HashSet::new();
        for rel in &rels {
            assert!(
                ids.insert(rel.id.clone()),
                "duplicate relationship id {}",
                rel.id
            );
        }
    }

    #[test]
    fn ensures_vba_signature_and_data_parts_for_signed_xlsm() {
        let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>"#;

        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

        // Intentionally omit the vbaProject relationship to ensure we add it.
        let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

        let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></worksheet>"#;

        // Intentionally omit `xl/_rels/vbaProject.bin.rels` to ensure we create it.
        let bytes = build_package(&[
            ("[Content_Types].xml", content_types.as_bytes()),
            ("xl/workbook.xml", workbook_xml.as_bytes()),
            ("xl/_rels/workbook.xml.rels", workbook_rels.as_bytes()),
            ("xl/worksheets/sheet1.xml", worksheet_xml.as_bytes()),
            ("xl/vbaProject.bin", b"fake-vba-project"),
            ("xl/vbaProjectSignature.bin", b"fake-signature"),
            ("xl/vbaData.xml", b"<vbaData/>"),
        ]);

        let pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
        let written = pkg.write_to_bytes().expect("write pkg");
        let pkg2 = XlsxPackage::from_bytes(&written).expect("read pkg2");

        let ct = std::str::from_utf8(pkg2.part("[Content_Types].xml").unwrap()).unwrap();
        assert!(ct.contains(
            r#"PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject""#
        ));
        assert!(ct.contains(
            r#"PartName="/xl/vbaProjectSignature.bin" ContentType="application/vnd.ms-office.vbaProjectSignature""#
        ));
        assert!(ct.contains(
            r#"PartName="/xl/vbaData.xml" ContentType="application/vnd.ms-office.vbaData+xml""#
        ));
        assert!(ct.contains("application/vnd.ms-excel.sheet.macroEnabled.main+xml"));

        let rels = std::str::from_utf8(pkg2.part("xl/_rels/workbook.xml.rels").unwrap()).unwrap();
        assert!(rels.contains("http://schemas.microsoft.com/office/2006/relationships/vbaProject"));
        assert!(rels.contains(r#"Target="vbaProject.bin""#));

        let vba_rels =
            std::str::from_utf8(pkg2.part("xl/_rels/vbaProject.bin.rels").unwrap()).unwrap();
        assert!(vba_rels.contains(
            "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature"
        ));
        assert!(vba_rels.contains(r#"Target="vbaProjectSignature.bin""#));
        assert!(vba_rels.contains(r#"Id="rId1""#));
    }

    #[test]
    fn repairs_vba_project_rels_signature_relationship_with_unusual_formatting() {
        let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>"#;

        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

        let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

        let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></worksheet>"#;

        // Tricky formatting:
        // - `Id = "rId1"` has spaces around `=`
        // - closing tag is `</Relationships >` (extra whitespace)
        // The old string-based implementation would fail to locate `</Relationships>` and would
        // not insert the signature relationship.
        let vba_project_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Target="vbaData.xml" Type="http://example.com/keep" Id = "rId1"/>
</Relationships >"#;

        let bytes = build_package(&[
            ("[Content_Types].xml", content_types.as_bytes()),
            ("xl/workbook.xml", workbook_xml.as_bytes()),
            ("xl/_rels/workbook.xml.rels", workbook_rels.as_bytes()),
            ("xl/worksheets/sheet1.xml", worksheet_xml.as_bytes()),
            ("xl/vbaProject.bin", b"fake-vba-project"),
            ("xl/vbaProjectSignature.bin", b"fake-signature"),
            ("xl/vbaData.xml", b"<vbaData/>"),
            ("xl/_rels/vbaProject.bin.rels", vba_project_rels.as_bytes()),
        ]);

        let pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
        let written = pkg.write_to_bytes().expect("write pkg");
        let pkg2 = XlsxPackage::from_bytes(&written).expect("read pkg2");

        let vba_rels_bytes = pkg2.part("xl/_rels/vbaProject.bin.rels").unwrap();
        let rels = crate::openxml::parse_relationships(vba_rels_bytes).expect("parse vba rels");

        let sig = rels
            .iter()
            .find(|rel| {
                rel.type_uri
                    == "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature"
            })
            .expect("expected vbaProject.bin.rels to contain a signature relationship");
        assert_eq!(sig.target, "vbaProjectSignature.bin");
        assert_eq!(sig.id, "rId2");

        let mut ids = HashSet::new();
        for rel in &rels {
            assert!(
                ids.insert(rel.id.clone()),
                "duplicate relationship id {}",
                rel.id
            );
        }
    }

    #[test]
    fn repairs_vba_project_rels_signature_relationship_patches_wrong_target() {
        let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>"#;

        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

        let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

        let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></worksheet>"#;

        // Relationship exists but has the wrong Target and unusual formatting.
        let vba_project_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Target="sig.bin" Id="rId9" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature"/>
</Relationships >"#;

        let bytes = build_package(&[
            ("[Content_Types].xml", content_types.as_bytes()),
            ("xl/workbook.xml", workbook_xml.as_bytes()),
            ("xl/_rels/workbook.xml.rels", workbook_rels.as_bytes()),
            ("xl/worksheets/sheet1.xml", worksheet_xml.as_bytes()),
            ("xl/vbaProject.bin", b"fake-vba-project"),
            ("xl/vbaProjectSignature.bin", b"fake-signature"),
            ("xl/_rels/vbaProject.bin.rels", vba_project_rels.as_bytes()),
        ]);

        let pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
        let written = pkg.write_to_bytes().expect("write pkg");
        let pkg2 = XlsxPackage::from_bytes(&written).expect("read pkg2");

        let vba_rels_bytes = pkg2.part("xl/_rels/vbaProject.bin.rels").unwrap();
        let rels = crate::openxml::parse_relationships(vba_rels_bytes).expect("parse vba rels");
        let sig = rels
            .iter()
            .find(|rel| {
                rel.type_uri
                    == "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature"
            })
            .expect("expected vbaProject.bin.rels to contain a signature relationship");
        assert_eq!(sig.target, "vbaProjectSignature.bin");
        assert_eq!(sig.id, "rId9");
    }

    #[test]
    fn repairs_xlsm_with_prefixed_content_types_xml() {
        let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types">
  <ct:Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <ct:Default Extension="xml" ContentType="application/xml"/>
  <ct:Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" PartName="/xl/workbook.xml"/>
</ct:Types>"#;

        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

        let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

        let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></worksheet>"#;

        let bytes = build_package(&[
            ("[Content_Types].xml", content_types.as_bytes()),
            ("xl/workbook.xml", workbook_xml.as_bytes()),
            ("xl/_rels/workbook.xml.rels", workbook_rels.as_bytes()),
            ("xl/worksheets/sheet1.xml", worksheet_xml.as_bytes()),
            ("xl/vbaProject.bin", b"fake-vba-project"),
        ]);

        let pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
        let written = pkg.write_to_bytes().expect("write pkg");
        let pkg2 = XlsxPackage::from_bytes(&written).expect("read pkg2");

        let ct = std::str::from_utf8(pkg2.part("[Content_Types].xml").unwrap()).unwrap();
        let doc = Document::parse(ct).expect("parse content types");

        let ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types";

        let workbook_override = doc
            .descendants()
            .find(|n| {
                n.is_element()
                    && n.tag_name().name() == "Override"
                    && n.attribute("PartName") == Some("/xl/workbook.xml")
            })
            .expect("workbook override");
        assert_eq!(workbook_override.tag_name().namespace(), Some(ct_ns));
        assert_eq!(
            workbook_override.attribute("ContentType"),
            Some("application/vnd.ms-excel.sheet.macroEnabled.main+xml")
        );

        let vba_override = doc
            .descendants()
            .find(|n| {
                n.is_element()
                    && n.tag_name().name() == "Override"
                    && n.attribute("PartName") == Some("/xl/vbaProject.bin")
            })
            .expect("vbaProject override");
        assert_eq!(vba_override.tag_name().namespace(), Some(ct_ns));
        assert_eq!(
            vba_override.attribute("ContentType"),
            Some("application/vnd.ms-office.vbaProject")
        );
    }

    #[test]
    fn repairs_xlsm_with_prefix_only_defaults_and_missing_overrides() {
        // A pathological but observed-in-the-wild shape: no default namespace on the root, but a
        // declared prefix used for child elements. If the document is missing `<Override>` entries
        // we must preserve the prefix when inserting them (otherwise the injected nodes land in no
        // namespace and Excel ignores them).
        let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types">
  <ct:Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <ct:Default Extension="xml" ContentType="application/xml"/>
</Types>"#;

        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

        let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

        let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></worksheet>"#;

        let bytes = build_package(&[
            ("[Content_Types].xml", content_types.as_bytes()),
            ("xl/workbook.xml", workbook_xml.as_bytes()),
            ("xl/_rels/workbook.xml.rels", workbook_rels.as_bytes()),
            ("xl/worksheets/sheet1.xml", worksheet_xml.as_bytes()),
            ("xl/vbaProject.bin", b"fake-vba-project"),
        ]);

        let pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
        let written = pkg.write_to_bytes().expect("write pkg");
        let pkg2 = XlsxPackage::from_bytes(&written).expect("read pkg2");

        let ct = std::str::from_utf8(pkg2.part("[Content_Types].xml").unwrap()).unwrap();
        let doc = Document::parse(ct).expect("parse content types");

        let ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types";

        let workbook_override = doc
            .descendants()
            .find(|n| {
                n.is_element()
                    && n.tag_name().name() == "Override"
                    && n.attribute("PartName") == Some("/xl/workbook.xml")
            })
            .expect("workbook override");
        assert_eq!(workbook_override.tag_name().namespace(), Some(ct_ns));
        assert_eq!(
            workbook_override.attribute("ContentType"),
            Some("application/vnd.ms-excel.sheet.macroEnabled.main+xml")
        );

        let vba_override = doc
            .descendants()
            .find(|n| {
                n.is_element()
                    && n.tag_name().name() == "Override"
                    && n.attribute("PartName") == Some("/xl/vbaProject.bin")
            })
            .expect("vbaProject override");
        assert_eq!(vba_override.tag_name().namespace(), Some(ct_ns));
        assert_eq!(
            vba_override.attribute("ContentType"),
            Some("application/vnd.ms-office.vbaProject")
        );

        assert!(
            ct.contains("<ct:Override"),
            "expected inserted Override nodes to use the ct: prefix, got:\n{ct}"
        );
    }

    #[test]
    fn repairs_workbook_rels_with_prefixed_self_closing_root() {
        let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>"#;

        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

        // Self-closing root with a prefix (no default namespace).
        let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pr:Relationships xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships"/>"#;

        let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></worksheet>"#;

        let bytes = build_package(&[
            ("[Content_Types].xml", content_types.as_bytes()),
            ("xl/workbook.xml", workbook_xml.as_bytes()),
            ("xl/_rels/workbook.xml.rels", workbook_rels.as_bytes()),
            ("xl/worksheets/sheet1.xml", worksheet_xml.as_bytes()),
            ("xl/vbaProject.bin", b"fake-vba-project"),
        ]);

        let pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
        let written = pkg.write_to_bytes().expect("write pkg");
        let pkg2 = XlsxPackage::from_bytes(&written).expect("read pkg2");

        let rels_xml =
            std::str::from_utf8(pkg2.part("xl/_rels/workbook.xml.rels").unwrap()).unwrap();
        let doc = Document::parse(rels_xml).expect("parse workbook rels");

        let rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships";
        let vba_rel = doc
            .descendants()
            .find(|n| {
                n.is_element()
                    && n.tag_name().name() == "Relationship"
                    && n.attribute("Type")
                        == Some("http://schemas.microsoft.com/office/2006/relationships/vbaProject")
            })
            .expect("vbaProject relationship");
        assert_eq!(vba_rel.tag_name().namespace(), Some(rel_ns));
        assert_eq!(vba_rel.attribute("Target"), Some("vbaProject.bin"));
        assert_eq!(vba_rel.attribute("Id"), Some("rId1"));
    }

    #[test]
    fn repairs_vba_project_rels_with_prefixed_self_closing_root() {
        let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>"#;

        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

        let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

        let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></worksheet>"#;

        let vba_project_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pr:Relationships xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships"/>"#;

        let bytes = build_package(&[
            ("[Content_Types].xml", content_types.as_bytes()),
            ("xl/workbook.xml", workbook_xml.as_bytes()),
            ("xl/_rels/workbook.xml.rels", workbook_rels.as_bytes()),
            ("xl/worksheets/sheet1.xml", worksheet_xml.as_bytes()),
            ("xl/vbaProject.bin", b"fake-vba-project"),
            ("xl/vbaProjectSignature.bin", b"fake-signature"),
            ("xl/_rels/vbaProject.bin.rels", vba_project_rels.as_bytes()),
        ]);

        let pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
        let written = pkg.write_to_bytes().expect("write pkg");
        let pkg2 = XlsxPackage::from_bytes(&written).expect("read pkg2");

        let rels_xml =
            std::str::from_utf8(pkg2.part("xl/_rels/vbaProject.bin.rels").unwrap()).unwrap();
        let doc = Document::parse(rels_xml).expect("parse vbaProject.bin.rels");

        let rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships";
        let sig_rel = doc
            .descendants()
            .find(|n| {
                n.is_element()
                    && n.tag_name().name() == "Relationship"
                    && n.attribute("Type")
                        == Some("http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature")
            })
            .expect("signature relationship");
        assert_eq!(sig_rel.tag_name().namespace(), Some(rel_ns));
        assert_eq!(sig_rel.attribute("Target"), Some("vbaProjectSignature.bin"));
        assert_eq!(sig_rel.attribute("Id"), Some("rId1"));
    }

    #[test]
    #[cfg(feature = "vba")]
    fn parses_vba_project_from_package() {
        let fixture = load_fixture();
        let pkg = XlsxPackage::from_bytes(&fixture).expect("read pkg");
        let project = pkg
            .vba_project()
            .expect("parse vba project")
            .expect("vba project present");

        assert_eq!(project.name.as_deref(), Some("VBAProject"));
        let module = project
            .modules
            .iter()
            .find(|m| m.name == "Module1")
            .expect("Module1 present");
        assert!(module.code.contains("Sub Hello"));
        assert_eq!(
            module.attributes.get("VB_Name").map(String::as_str),
            Some("Module1")
        );
    }

    #[test]
    fn package_exposes_sheet_list_and_tab_color_helpers() {
        let bytes = build_minimal_package();
        let mut pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");

        let sheets = pkg.workbook_sheets().expect("parse sheets");
        assert_eq!(sheets.len(), 1);
        assert_eq!(sheets[0].name, "Sheet1");
        assert_eq!(sheets[0].sheet_id, 1);
        assert_eq!(sheets[0].rel_id, "rId1");
        assert_eq!(
            sheets[0].visibility,
            formula_model::SheetVisibility::Visible
        );

        let mut updated = sheets.clone();
        updated[0].name = "Renamed".to_string();
        pkg.set_workbook_sheets(&updated).expect("write sheets");
        let renamed = pkg.workbook_sheets().expect("parse renamed sheets");
        assert_eq!(renamed[0].name, "Renamed");

        let color = TabColor::rgb("FFFF0000");
        pkg.set_worksheet_tab_color("xl/worksheets/sheet1.xml", Some(&color))
            .expect("set tab color");
        let parsed = pkg
            .worksheet_tab_color("xl/worksheets/sheet1.xml")
            .expect("parse tab color");
        assert_eq!(parsed.unwrap().rgb.as_deref(), Some("FFFF0000"));

        pkg.set_worksheet_tab_color("xl/worksheets/sheet1.xml", None)
            .expect("remove tab color");
        assert_eq!(
            pkg.worksheet_tab_color("xl/worksheets/sheet1.xml")
                .expect("parse tab color"),
            None
        );
    }

    #[test]
    fn remove_vba_project_strips_vba_parts() {
        let fixture = load_fixture();
        let mut pkg = XlsxPackage::from_bytes(&fixture).expect("read pkg");

        assert!(pkg.vba_project_bin().is_some());
        pkg.remove_vba_project().expect("remove vba project");

        let written = pkg.write_to_bytes().expect("write pkg");
        let pkg2 = XlsxPackage::from_bytes(&written).expect("read pkg2");

        assert!(pkg2.vba_project_bin().is_none());

        let ct = std::str::from_utf8(pkg2.part("[Content_Types].xml").unwrap()).unwrap();
        assert!(!ct.contains("vbaProject.bin"));
        assert!(!ct.contains("macroEnabled.main+xml"));

        let rels = std::str::from_utf8(pkg2.part("xl/_rels/workbook.xml.rels").unwrap()).unwrap();
        assert!(!rels.contains("relationships/vbaProject"));
    }

    fn build_synthetic_macro_package() -> Vec<u8> {
        let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>
  <Override PartName="/xl/vbaProjectSignature.bin" ContentType="application/vnd.ms-office.vbaProjectSignature"/>
  <Override PartName="/xl/vbaData.xml" ContentType="application/vnd.ms-office.vbaData+xml"/>
  <Override PartName="/customUI/customUI.xml" ContentType="application/xml"/>
  <Override PartName="/customUI/customUI14.xml" ContentType="application/xml"/>
  <Override PartName="/xl/activeX/activeX1.xml" ContentType="application/vnd.ms-office.activeX+xml"/>
  <Override PartName="/xl/ctrlProps/ctrlProp1.xml" ContentType="application/vnd.ms-office.activeX+xml"/>
  <Override PartName="/xl/embeddings/oleObject1.bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
</Types>"#;

        let root_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility" Target="customUI/customUI.xml"/>
  <Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>
</Relationships>"#;

        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

        let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>
</Relationships>"#;

        let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></worksheet>"#;

        let sheet_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/activeXControl" Target="../activeX/activeX1.xml#_x0000_s1025"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2006/relationships/ctrlProp" Target="../ctrlProps/ctrlProp1.xml"/>
</Relationships>"#;

        let vba_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vbaProjectSignature.bin"/>
</Relationships>"#;

        let custom_ui_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"></customUI>"#;

        let custom_ui_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="image1.png"/>
</Relationships>"#;

        let activex_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ax:ocx xmlns:ax="http://schemas.microsoft.com/office/2006/activeX"></ax:ocx>"#;

        let activex_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary" Target="activeX1.bin"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="../embeddings/oleObject1.bin"/>
</Relationships>"#;

        let ctrl_props_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ctrlProp xmlns="http://schemas.microsoft.com/office/2006/activeX"></ctrlProp>"#;

        let cursor = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(cursor);
        let options = zip::write::FileOptions::<()>::default()
            .compression_method(zip::CompressionMethod::Deflated);

        zip.start_file("[Content_Types].xml", options).unwrap();
        zip.write_all(content_types.as_bytes()).unwrap();

        zip.start_file("_rels/.rels", options).unwrap();
        zip.write_all(root_rels.as_bytes()).unwrap();

        zip.start_file("xl/workbook.xml", options).unwrap();
        zip.write_all(workbook_xml.as_bytes()).unwrap();

        zip.start_file("xl/_rels/workbook.xml.rels", options)
            .unwrap();
        zip.write_all(workbook_rels.as_bytes()).unwrap();

        zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
        zip.write_all(worksheet_xml.as_bytes()).unwrap();

        zip.start_file("xl/worksheets/_rels/sheet1.xml.rels", options)
            .unwrap();
        zip.write_all(sheet_rels.as_bytes()).unwrap();

        zip.start_file("customUI/customUI.xml", options).unwrap();
        zip.write_all(custom_ui_xml.as_bytes()).unwrap();

        zip.start_file("customUI/customUI14.xml", options).unwrap();
        zip.write_all(custom_ui_xml.as_bytes()).unwrap();

        zip.start_file("customUI/_rels/customUI.xml.rels", options)
            .unwrap();
        zip.write_all(custom_ui_rels.as_bytes()).unwrap();

        zip.start_file("customUI/image1.png", options).unwrap();
        zip.write_all(b"not-a-real-png").unwrap();

        zip.start_file("xl/vbaProject.bin", options).unwrap();
        zip.write_all(b"fake-vba-project").unwrap();

        zip.start_file("xl/_rels/vbaProject.bin.rels", options)
            .unwrap();
        zip.write_all(vba_rels.as_bytes()).unwrap();

        zip.start_file("xl/vbaProjectSignature.bin", options)
            .unwrap();
        zip.write_all(b"fake-signature").unwrap();

        zip.start_file("xl/vbaData.xml", options).unwrap();
        zip.write_all(b"<vbaData/>").unwrap();

        zip.start_file("xl/activeX/activeX1.xml", options).unwrap();
        zip.write_all(activex_xml.as_bytes()).unwrap();

        zip.start_file("xl/activeX/_rels/activeX1.xml.rels", options)
            .unwrap();
        zip.write_all(activex_rels.as_bytes()).unwrap();

        zip.start_file("xl/activeX/activeX1.bin", options).unwrap();
        zip.write_all(b"activex-binary").unwrap();

        zip.start_file("xl/embeddings/oleObject1.bin", options)
            .unwrap();
        zip.write_all(b"ole-embedding").unwrap();

        zip.start_file("xl/ctrlProps/ctrlProp1.xml", options)
            .unwrap();
        zip.write_all(ctrl_props_xml.as_bytes()).unwrap();

        zip.finish().unwrap().into_inner()
    }

    #[test]
    fn remove_vba_project_strips_macro_part_graph_and_repairs_relationships() {
        let bytes = build_synthetic_macro_package();
        let mut pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
        pkg.remove_vba_project().expect("strip macros");

        // Round-trip through ZIP writing to ensure we didn't leave any dangling references.
        let written = pkg.write_to_bytes().expect("write stripped pkg");
        let pkg2 = XlsxPackage::from_bytes(&written).expect("read stripped pkg");

        for removed in [
            "xl/vbaProject.bin",
            "xl/_rels/vbaProject.bin.rels",
            "xl/vbaProjectSignature.bin",
            "xl/vbaData.xml",
            "customUI/customUI.xml",
            "customUI/customUI14.xml",
            "customUI/_rels/customUI.xml.rels",
            "customUI/image1.png",
            "xl/activeX/activeX1.xml",
            "xl/activeX/_rels/activeX1.xml.rels",
            "xl/activeX/activeX1.bin",
            "xl/ctrlProps/ctrlProp1.xml",
            // Child part referenced only by the removed ActiveX graph.
            "xl/embeddings/oleObject1.bin",
        ] {
            assert!(
                pkg2.part(removed).is_none(),
                "expected {removed} to be removed"
            );
        }

        let ct = std::str::from_utf8(pkg2.part("[Content_Types].xml").unwrap()).unwrap();
        assert!(!ct.contains("macroEnabled.main+xml"));
        assert!(!ct.contains("vbaProject.bin"));
        assert!(!ct.contains("customUI/customUI.xml"));
        assert!(!ct.contains("customUI/customUI14.xml"));
        assert!(!ct.contains("activeX1.xml"));
        assert!(!ct.contains("ctrlProp1.xml"));

        // Relationship parts should no longer mention the stripped macro graph.
        for (name, bytes) in pkg2.parts() {
            if !name.ends_with(".rels") {
                continue;
            }
            let xml = std::str::from_utf8(bytes).unwrap();
            assert!(!xml.contains("vbaProject"));
            assert!(!xml.contains("customUI"));
            assert!(!xml.contains("activeX"));
            assert!(!xml.contains("ctrlProps"));
        }

        crate::macro_strip::validate_opc_relationships(pkg2.parts_map())
            .expect("stripped package relationships are consistent");
    }

    #[test]
    fn remove_vba_project_strips_worksheet_rid_references() {
        let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <controls>
    <control r:id="rIdA"/>
  </controls>
</worksheet>"#;

        let rels_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdA" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/control" Target="../activeX/activeX1.xml"/>
</Relationships>"#;

        let bytes = build_package(&[
            ("xl/worksheets/sheet1.xml", worksheet_xml.as_bytes()),
            ("xl/worksheets/_rels/sheet1.xml.rels", rels_xml.as_bytes()),
            ("xl/activeX/activeX1.xml", b"<activeX/>"),
        ]);

        let mut pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
        pkg.remove_vba_project().expect("strip macros");

        let updated_rels =
            std::str::from_utf8(pkg.part("xl/worksheets/_rels/sheet1.xml.rels").unwrap()).unwrap();
        assert!(!updated_rels.contains("rIdA"));

        let updated_sheet =
            std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
        assert!(!updated_sheet.contains("rIdA"));
        assert!(!updated_sheet.contains("<control r:id"));

        assert!(pkg.part("xl/activeX/activeX1.xml").is_none());
    }

    #[test]
    fn remove_vba_project_strips_vml_rid_references() {
        let vml_xml = r##"<?xml version="1.0" encoding="UTF-8"?>
<xml xmlns:v="urn:schemas-microsoft-com:vml"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <v:shape id="_x0000_s1025" type="#_x0000_t75">
    <o:OLEObject r:id="rIdOle"/>
  </v:shape>
  <v:shape id="_x0000_s1026" type="#_x0000_t75">
    <x:ClientData ObjectType="Note"></x:ClientData>
  </v:shape>
</xml>"##;

        let rels_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOle" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/activeXControlBinary" Target="../activeX/activeX1.bin"/>
</Relationships>"#;

        let bytes = build_package(&[
            ("xl/drawings/vmlDrawing1.vml", vml_xml.as_bytes()),
            (
                "xl/drawings/_rels/vmlDrawing1.vml.rels",
                rels_xml.as_bytes(),
            ),
            ("xl/activeX/activeX1.bin", b"dummy-bin"),
        ]);

        let mut pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
        pkg.remove_vba_project().expect("strip macros");

        let updated_rels =
            std::str::from_utf8(pkg.part("xl/drawings/_rels/vmlDrawing1.vml.rels").unwrap())
                .unwrap();
        assert!(!updated_rels.contains("rIdOle"));

        let updated_vml =
            std::str::from_utf8(pkg.part("xl/drawings/vmlDrawing1.vml").unwrap()).unwrap();
        assert!(!updated_vml.contains("rIdOle"));
        assert!(!updated_vml.contains("OLEObject"));
        assert!(updated_vml.contains("ObjectType=\"Note\""));

        assert!(pkg.part("xl/activeX/activeX1.bin").is_none());
    }

    #[test]
    fn remove_vba_project_preserves_vml_legacy_drawing_linkage_while_stripping_macro_shapes() {
        let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <legacyDrawing r:id="rIdVml"/>
</worksheet>"#;

        let sheet_rels_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdVml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>
</Relationships>"#;

        let vml_xml = r##"<?xml version="1.0" encoding="UTF-8"?>
<xml xmlns:v="urn:schemas-microsoft-com:vml"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <v:shape id="_x0000_s1025" type="#_x0000_t75">
    <o:OLEObject r:id="rIdOle"/>
  </v:shape>
  <v:shape id="_x0000_s1026" type="#_x0000_t75">
    <x:ClientData ObjectType="Note"></x:ClientData>
  </v:shape>
</xml>"##;

        let vml_rels_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOle" Type="http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary" Target="../activeX/activeX1.bin"/>
</Relationships>"#;

        let bytes = build_package(&[
            ("xl/worksheets/sheet1.xml", worksheet_xml.as_bytes()),
            (
                "xl/worksheets/_rels/sheet1.xml.rels",
                sheet_rels_xml.as_bytes(),
            ),
            ("xl/drawings/vmlDrawing1.vml", vml_xml.as_bytes()),
            (
                "xl/drawings/_rels/vmlDrawing1.vml.rels",
                vml_rels_xml.as_bytes(),
            ),
            ("xl/activeX/activeX1.bin", b"dummy-bin"),
        ]);

        let mut pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
        pkg.remove_vba_project().expect("strip macros");

        assert!(
            pkg.part("xl/drawings/vmlDrawing1.vml").is_some(),
            "expected vmlDrawing1.vml to be preserved (worksheet legacyDrawing reference)"
        );
        assert!(
            pkg.part("xl/worksheets/_rels/sheet1.xml.rels").is_some(),
            "expected worksheet relationship part to be preserved"
        );

        let updated_sheet =
            std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
        assert!(updated_sheet.contains("legacyDrawing"));
        assert!(updated_sheet.contains("rIdVml"));

        let updated_sheet_rels =
            std::str::from_utf8(pkg.part("xl/worksheets/_rels/sheet1.xml.rels").unwrap()).unwrap();
        assert!(updated_sheet_rels.contains("rIdVml"));
        assert!(updated_sheet_rels.contains("vmlDrawing1.vml"));

        let updated_vml =
            std::str::from_utf8(pkg.part("xl/drawings/vmlDrawing1.vml").unwrap()).unwrap();
        assert!(!updated_vml.contains("rIdOle"));
        assert!(!updated_vml.contains("OLEObject"));
        assert!(updated_vml.contains("ObjectType=\"Note\""));

        let updated_vml_rels =
            std::str::from_utf8(pkg.part("xl/drawings/_rels/vmlDrawing1.vml.rels").unwrap())
                .unwrap();
        assert!(!updated_vml_rels.contains("rIdOle"));

        assert!(pkg.part("xl/activeX/activeX1.bin").is_none());

        crate::macro_strip::validate_opc_relationships(pkg.parts_map())
            .expect("stripped package relationships are consistent");
    }

    #[test]
    fn remove_vba_project_strips_drawing_embed_references() {
        let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:pic>
      <xdr:blipFill>
        <a:blip r:embed="rIdImg"/>
      </xdr:blipFill>
    </xdr:pic>
  </xdr:twoCellAnchor>
  <xdr:twoCellAnchor>
    <xdr:pic>
      <xdr:blipFill>
        <a:blip r:embed="rIdKeep"/>
      </xdr:blipFill>
    </xdr:pic>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

        let rels_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../ctrlProps/image1.png"/>
  <Relationship Id="rIdKeep" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.png"/>
</Relationships>"#;

        let bytes = build_package(&[
            ("xl/drawings/drawing1.xml", drawing_xml.as_bytes()),
            ("xl/drawings/_rels/drawing1.xml.rels", rels_xml.as_bytes()),
            ("xl/ctrlProps/image1.png", b"macro-image"),
            ("xl/media/image2.png", b"keep-image"),
        ]);

        let mut pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
        pkg.remove_vba_project().expect("strip macros");

        let updated_rels =
            std::str::from_utf8(pkg.part("xl/drawings/_rels/drawing1.xml.rels").unwrap()).unwrap();
        assert!(!updated_rels.contains("rIdImg"));
        assert!(updated_rels.contains("rIdKeep"));

        let updated_drawing =
            std::str::from_utf8(pkg.part("xl/drawings/drawing1.xml").unwrap()).unwrap();
        assert!(!updated_drawing.contains("rIdImg"));
        assert!(updated_drawing.contains("rIdKeep"));

        assert!(pkg.part("xl/ctrlProps/image1.png").is_none());
        assert!(pkg.part("xl/media/image2.png").is_some());
    }

    #[test]
    fn remove_vba_project_strips_vml_relid_references() {
        let vml_xml = r##"<?xml version="1.0" encoding="UTF-8"?>
<xml xmlns:v="urn:schemas-microsoft-com:vml"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel">
  <v:shape id="_x0000_s1025" type="#_x0000_t75">
    <v:imagedata o:relid="rIdImg"/>
  </v:shape>
  <v:shape id="_x0000_s1026" type="#_x0000_t75">
    <v:imagedata o:relid="rIdKeep"/>
  </v:shape>
</xml>"##;

        let rels_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../ctrlProps/image1.png"/>
  <Relationship Id="rIdKeep" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.png"/>
</Relationships>"#;

        let bytes = build_package(&[
            ("xl/drawings/vmlDrawing1.vml", vml_xml.as_bytes()),
            (
                "xl/drawings/_rels/vmlDrawing1.vml.rels",
                rels_xml.as_bytes(),
            ),
            ("xl/ctrlProps/image1.png", b"macro-image"),
            ("xl/media/image2.png", b"keep-image"),
        ]);

        let mut pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
        pkg.remove_vba_project().expect("strip macros");

        let updated_rels =
            std::str::from_utf8(pkg.part("xl/drawings/_rels/vmlDrawing1.vml.rels").unwrap())
                .unwrap();
        assert!(!updated_rels.contains("rIdImg"));
        assert!(updated_rels.contains("rIdKeep"));

        let updated_vml =
            std::str::from_utf8(pkg.part("xl/drawings/vmlDrawing1.vml").unwrap()).unwrap();
        assert!(!updated_vml.contains("rIdImg"));
        assert!(updated_vml.contains("rIdKeep"));

        assert!(pkg.part("xl/ctrlProps/image1.png").is_none());
        assert!(pkg.part("xl/media/image2.png").is_some());
    }

    #[test]
    fn remove_vba_project_strips_drawing_link_references() {
        let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:pic>
      <xdr:blipFill>
        <a:blip r:link="rIdLink"/>
      </xdr:blipFill>
    </xdr:pic>
  </xdr:twoCellAnchor>
  <xdr:twoCellAnchor>
    <xdr:pic>
      <xdr:blipFill>
        <a:blip r:link="rIdKeep"/>
      </xdr:blipFill>
    </xdr:pic>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

        let rels_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../ctrlProps/image1.png"/>
  <Relationship Id="rIdKeep" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.png"/>
</Relationships>"#;

        let bytes = build_package(&[
            ("xl/drawings/drawing1.xml", drawing_xml.as_bytes()),
            ("xl/drawings/_rels/drawing1.xml.rels", rels_xml.as_bytes()),
            ("xl/ctrlProps/image1.png", b"macro-image"),
            ("xl/media/image2.png", b"keep-image"),
        ]);

        let mut pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
        pkg.remove_vba_project().expect("strip macros");

        let updated_rels =
            std::str::from_utf8(pkg.part("xl/drawings/_rels/drawing1.xml.rels").unwrap()).unwrap();
        assert!(!updated_rels.contains("rIdLink"));
        assert!(updated_rels.contains("rIdKeep"));

        let updated_drawing =
            std::str::from_utf8(pkg.part("xl/drawings/drawing1.xml").unwrap()).unwrap();
        assert!(!updated_drawing.contains("rIdLink"));
        assert!(updated_drawing.contains("rIdKeep"));

        assert!(pkg.part("xl/ctrlProps/image1.png").is_none());
        assert!(pkg.part("xl/media/image2.png").is_some());
    }
}
