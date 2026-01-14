use std::collections::BTreeMap;
use std::io::{Cursor, Read, Seek, SeekFrom, Write};
use std::sync::Arc;

use zip::ZipArchive;

use crate::package::{MacroPresence, WorkbookKind, XlsxError, MAX_XLSX_PACKAGE_PART_BYTES};
use crate::streaming::PartOverride;
use crate::zip_util::read_zip_file_bytes_with_limit;

#[cfg(not(target_arch = "wasm32"))]
use std::collections::HashMap;
#[cfg(not(target_arch = "wasm32"))]
use std::path::{Path, PathBuf};

#[cfg(not(target_arch = "wasm32"))]
use crate::WorkbookCellPatches;

#[cfg(not(target_arch = "wasm32"))]
use crate::streaming::strip_vba_project_streaming_with_kind;

#[cfg(not(target_arch = "wasm32"))]
use crate::streaming::patch_xlsx_streaming_workbook_cell_patches_with_part_overrides;

#[cfg(not(target_arch = "wasm32"))]
use tempfile::tempfile;

#[derive(Debug, Clone)]
enum Source {
    #[cfg(not(target_arch = "wasm32"))]
    Path(PathBuf),
    Bytes(Arc<Vec<u8>>),
}

impl Source {
    #[cfg(not(target_arch = "wasm32"))]
    fn open_reader(&self) -> Result<SourceReader<'_>, XlsxError> {
        match self {
            Source::Path(path) => Ok(SourceReader::File(std::fs::File::open(path)?)),
            Source::Bytes(bytes) => Ok(SourceReader::Bytes(Cursor::new(bytes.as_slice()))),
        }
    }

    #[cfg(target_arch = "wasm32")]
    fn open_reader(&self) -> Result<SourceReader<'_>, XlsxError> {
        match self {
            Source::Bytes(bytes) => Ok(SourceReader::Bytes(Cursor::new(bytes.as_slice()))),
        }
    }
}

/// A `Read + Seek` input reader for [`XlsxLazyPackage`] operations.
enum SourceReader<'a> {
    #[cfg(not(target_arch = "wasm32"))]
    File(std::fs::File),
    Bytes(Cursor<&'a [u8]>),
}

impl Read for SourceReader<'_> {
    fn read(&mut self, buf: &mut [u8]) -> std::io::Result<usize> {
        match self {
            #[cfg(not(target_arch = "wasm32"))]
            SourceReader::File(f) => f.read(buf),
            SourceReader::Bytes(c) => c.read(buf),
        }
    }
}

impl Seek for SourceReader<'_> {
    fn seek(&mut self, pos: SeekFrom) -> std::io::Result<u64> {
        match self {
            #[cfg(not(target_arch = "wasm32"))]
            SourceReader::File(f) => f.seek(pos),
            SourceReader::Bytes(c) => c.seek(pos),
        }
    }
}

/// A lazy/streaming XLSX/XLSM package wrapper.
///
/// Unlike [`crate::XlsxPackage`], this type does **not** inflate every ZIP entry into memory.
/// Instead, it keeps a reference to the underlying container (path or bytes) and uses the
/// streaming ZIP rewrite pipeline when writing.
///
/// This is intended for scenarios where callers want to preserve unknown parts (e.g. `customXml/`,
/// `xl/vbaProject.bin`) while keeping memory usage low for large workbooks.
#[derive(Debug, Clone)]
pub struct XlsxLazyPackage {
    source: Source,
    /// Deterministic map of part overrides keyed by canonical part name (`xl/workbook.xml`, without
    /// a leading `/`).
    overrides: BTreeMap<String, PartOverride>,
    /// When set, macros are stripped on write.
    strip_macros: bool,
    /// Optional workbook kind enforcement (used to update `[Content_Types].xml`).
    workbook_kind: Option<WorkbookKind>,
    /// Cached set of part names discovered at open time (canonical, without leading `/`).
    part_names: Vec<String>,
}

impl XlsxLazyPackage {
    /// Open an XLSX/XLSM file from disk without inflating all parts into memory.
    #[cfg(not(target_arch = "wasm32"))]
    pub fn open(path: impl AsRef<Path>) -> Result<Self, XlsxError> {
        let path = path.as_ref();
        let file = std::fs::File::open(path)?;
        Self::from_file(path.to_path_buf(), file)
    }

    /// Create a lazy package backed by a file on disk.
    ///
    /// This parses the ZIP central directory to discover part names, but does not inflate part
    /// contents into memory.
    #[cfg(not(target_arch = "wasm32"))]
    pub fn from_file(path: PathBuf, file: std::fs::File) -> Result<Self, XlsxError> {
        let part_names = list_part_names(file)?;
        Ok(Self {
            source: Source::Path(path),
            overrides: BTreeMap::new(),
            strip_macros: false,
            workbook_kind: None,
            part_names,
        })
    }

    /// Create a lazy package from owned ZIP bytes.
    pub fn from_vec(bytes: Vec<u8>) -> Result<Self, XlsxError> {
        let cursor = Cursor::new(bytes.as_slice());
        let part_names = list_part_names(cursor)?;
        Ok(Self {
            source: Source::Bytes(Arc::new(bytes)),
            overrides: BTreeMap::new(),
            strip_macros: false,
            workbook_kind: None,
            part_names,
        })
    }

    /// Create a lazy package by copying ZIP bytes.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self, XlsxError> {
        Self::from_vec(bytes.to_vec())
    }

    fn canonicalize_part_name(name: &str) -> String {
        let trimmed = name.trim_start_matches('/');
        // Normalize Windows-style separators.
        trimmed.replace('\\', "/")
    }

    /// Read a single part's bytes.
    ///
    /// This consults in-memory overrides first; if no override is present, it extracts the part
    /// from the underlying ZIP container.
    pub fn read_part(&self, name: &str) -> Result<Option<Vec<u8>>, XlsxError> {
        let canonical = Self::canonicalize_part_name(name);
        if let Some(override_op) = self.overrides.get(&canonical) {
            match override_op {
                PartOverride::Replace(bytes) | PartOverride::Add(bytes) => {
                    return Ok(Some(bytes.clone()));
                }
                PartOverride::Remove => return Ok(None),
            }
        }

        // Best-effort: if macros are being stripped, treat obvious macro surfaces as missing.
        if self.strip_macros && is_macro_part_name(&canonical) {
            return Ok(None);
        }

        let mut reader = self.source.open_reader()?;
        let mut archive = ZipArchive::new(&mut reader)?;
        let mut file = match crate::zip_util::open_zip_part(&mut archive, &canonical) {
            Ok(file) => file,
            Err(zip::result::ZipError::FileNotFound) => return Ok(None),
            Err(err) => return Err(err.into()),
        };
        let buf =
            read_zip_file_bytes_with_limit(&mut file, &canonical, MAX_XLSX_PACKAGE_PART_BYTES)?;
        Ok(Some(buf))
    }

    /// Detect whether the package contains macro-capable content (VBA, XLM macrosheets, or legacy
    /// dialog sheets).
    pub fn macro_presence(&self) -> MacroPresence {
        let mut presence = MacroPresence {
            has_vba: false,
            has_xlm_macrosheets: false,
            has_dialog_sheets: false,
        };

        for name in self.effective_part_names() {
            let lower = name.to_ascii_lowercase();
            if lower == "xl/vbaproject.bin" {
                presence.has_vba = true;
            } else if lower.starts_with("xl/macrosheets/") {
                presence.has_xlm_macrosheets = true;
            } else if lower.starts_with("xl/dialogsheets/") {
                presence.has_dialog_sheets = true;
            }

            if presence.has_vba && presence.has_xlm_macrosheets && presence.has_dialog_sheets {
                break;
            }
        }

        presence
    }

    /// Replace or insert a part in-memory.
    pub fn set_part(&mut self, name: &str, bytes: Vec<u8>) {
        let canonical = Self::canonicalize_part_name(name);
        self.overrides
            .insert(canonical.clone(), PartOverride::Replace(bytes));
        if !self.part_names.iter().any(|n| n == &canonical) {
            self.part_names.push(canonical);
        }
    }

    /// Remove macro-related parts and relationships from the package on write.
    pub fn remove_vba_project(&mut self) -> Result<(), XlsxError> {
        self.strip_macros = true;

        // Prevent callers from reintroducing macro surfaces via overrides.
        let macro_override_keys: Vec<String> = self
            .overrides
            .keys()
            .filter(|k| is_macro_part_name(k))
            .cloned()
            .collect();
        for key in macro_override_keys {
            self.overrides.insert(key, PartOverride::Remove);
        }

        // Best-effort: update cached names.
        self.part_names.retain(|n| !is_macro_part_name(n));
        Ok(())
    }

    /// Ensure `[Content_Types].xml` advertises the correct workbook content type for the requested
    /// workbook kind.
    pub fn enforce_workbook_kind(&mut self, kind: WorkbookKind) -> Result<(), XlsxError> {
        self.workbook_kind = Some(kind);
        Ok(())
    }

    /// Serialize the package to `.xlsx` bytes.
    pub fn write_to_bytes(&self) -> Result<Vec<u8>, XlsxError> {
        let mut cursor = Cursor::new(Vec::new());
        self.write_to(&mut cursor)?;
        Ok(cursor.into_inner())
    }

    /// Write the package to an output stream, preserving unchanged ZIP entries via streaming raw
    /// copy.
    #[cfg(not(target_arch = "wasm32"))]
    pub fn write_to<W: Write + Seek>(&self, output: W) -> Result<(), XlsxError> {
        let mut overrides = self.build_part_overrides(/*kind_already_handled=*/false)?;

        if self.strip_macros {
            // First pass: strip macros into a temporary file.
            let mut tmp = tempfile()?;
            let kind = self.workbook_kind.unwrap_or(WorkbookKind::Workbook);
            let input = self.source.open_reader()?;
            strip_vba_project_streaming_with_kind(input, &mut tmp, kind)?;

            // Second pass: apply any explicit part overrides on top of the macro-stripped package.
            // Note: workbook kind is already handled by the macro-strip streaming pass.
            overrides = self.build_part_overrides(/*kind_already_handled=*/true)?;
            tmp.seek(SeekFrom::Start(0))?;
            patch_xlsx_streaming_workbook_cell_patches_with_part_overrides(
                &mut tmp,
                output,
                &WorkbookCellPatches::default(),
                &overrides,
            )?;
            return Ok(());
        }

        let input = self.source.open_reader()?;
        patch_xlsx_streaming_workbook_cell_patches_with_part_overrides(
            input,
            output,
            &WorkbookCellPatches::default(),
            &overrides,
        )?;
        Ok(())
    }

    #[cfg(target_arch = "wasm32")]
    pub fn write_to<W: Write + Seek>(&self, _output: W) -> Result<(), XlsxError> {
        Err(XlsxError::Invalid(
            "XlsxLazyPackage::write_to is not supported on wasm32".to_string(),
        ))
    }

    fn effective_part_names(&self) -> impl Iterator<Item = &str> {
        self.part_names.iter().map(|s| s.as_str()).chain(
            self.overrides
                .iter()
                .filter_map(|(name, op)| match op {
                    PartOverride::Remove => None,
                    PartOverride::Replace(_) | PartOverride::Add(_) => Some(name.as_str()),
                }),
        )
    }

    #[cfg(not(target_arch = "wasm32"))]
    fn build_part_overrides(
        &self,
        kind_already_handled: bool,
    ) -> Result<HashMap<String, PartOverride>, XlsxError> {
        let mut out: HashMap<String, PartOverride> =
            self.overrides.iter().map(|(k, v)| (k.clone(), v.clone())).collect();
        if !kind_already_handled {
            if let Some(kind) = self.workbook_kind {
                if let Some(updated) = self.patched_content_types_for_kind(kind)? {
                    out.insert(
                        "[Content_Types].xml".to_string(),
                        PartOverride::Replace(updated),
                    );
                }
            }
        }
        Ok(out)
    }

    #[cfg(not(target_arch = "wasm32"))]
    fn patched_content_types_for_kind(&self, kind: WorkbookKind) -> Result<Option<Vec<u8>>, XlsxError> {
        let ct_name = "[Content_Types].xml";
        match self.overrides.get(ct_name) {
            Some(PartOverride::Remove) => return Ok(None),
            Some(PartOverride::Replace(bytes)) | Some(PartOverride::Add(bytes)) => {
                return Ok(crate::rewrite_content_types_workbook_kind(bytes, kind)?);
            }
            None => {}
        }

        let Some(existing) = self.read_part(ct_name)? else {
            return Ok(None);
        };
        Ok(crate::rewrite_content_types_workbook_kind(&existing, kind)?)
    }
}

fn is_macro_part_name(name: &str) -> bool {
    let name = name.strip_prefix('/').unwrap_or(name);
    let lower = name.to_ascii_lowercase();
    lower == "xl/vbaproject.bin"
        || lower == "xl/vbadata.xml"
        || lower == "xl/vbaprojectsignature.bin"
        || lower.starts_with("xl/macrosheets/")
        || lower.starts_with("xl/dialogsheets/")
        || lower.starts_with("customui/")
        || lower.starts_with("xl/activex/")
        || lower.starts_with("xl/ctrlprops/")
        || lower.starts_with("xl/controls/")
}

fn list_part_names<R: Read + Seek>(mut reader: R) -> Result<Vec<String>, XlsxError> {
    reader.seek(SeekFrom::Start(0))?;
    let mut zip = ZipArchive::new(reader)?;
    let mut names = Vec::new();
    for i in 0..zip.len() {
        let file = zip.by_index(i)?;
        if file.is_dir() {
            continue;
        }
        let name = file.name();
        let canonical = name.strip_prefix('/').unwrap_or(name).replace('\\', "/");
        names.push(canonical);
    }
    Ok(names)
}
