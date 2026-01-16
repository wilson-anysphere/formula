use std::collections::BTreeMap;
use std::borrow::Cow;
use std::fmt;
use std::io::{Cursor, Read, Seek, SeekFrom, Write};
use std::sync::Arc;

use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader as XmlReader, Writer as XmlWriter};
use zip::ZipArchive;

use crate::package::{MacroPresence, WorkbookKind, XlsxError, MAX_XLSX_PACKAGE_PART_BYTES};
use crate::streaming::PartOverride;
use crate::zip_util::read_zip_file_bytes_with_limit;

#[cfg(not(target_arch = "wasm32"))]
use std::collections::HashMap;
#[cfg(not(target_arch = "wasm32"))]
use std::path::{Path, PathBuf};

#[cfg(not(target_arch = "wasm32"))]
use formula_model::StyleTable;

#[cfg(not(target_arch = "wasm32"))]
use crate::WorkbookCellPatches;

#[cfg(not(target_arch = "wasm32"))]
use crate::streaming::strip_vba_project_streaming_with_kind;

#[cfg(not(target_arch = "wasm32"))]
use crate::streaming::patch_xlsx_streaming_workbook_cell_patches_with_part_overrides;

#[cfg(not(target_arch = "wasm32"))]
use crate::streaming::{
    patch_xlsx_streaming_workbook_cell_patches_with_part_overrides_and_recalc_policy,
    patch_xlsx_streaming_workbook_cell_patches_with_styles_and_part_overrides_and_recalc_policy,
};

#[cfg(not(target_arch = "wasm32"))]
use crate::RecalcPolicy;

#[cfg(not(target_arch = "wasm32"))]
use tempfile::tempfile;

#[derive(Clone)]
enum Source {
    #[cfg(not(target_arch = "wasm32"))]
    Path(Arc<PathBuf>),
    Bytes(Arc<Vec<u8>>),
}

impl fmt::Debug for Source {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Source::Bytes(bytes) => f
                .debug_struct("Bytes")
                .field("len", &bytes.len())
                .finish(),
            #[cfg(not(target_arch = "wasm32"))]
            Source::Path(path) => f.debug_tuple("Path").field(path).finish(),
        }
    }
}

impl Source {
    #[cfg(not(target_arch = "wasm32"))]
    fn open_reader(&self) -> Result<SourceReader<'_>, XlsxError> {
        match self {
            Source::Path(path) => Ok(SourceReader::File(std::fs::File::open(path.as_path())?)),
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
#[derive(Clone)]
pub struct XlsxLazyPackage {
    source: Source,
    /// Deterministic map of part overrides keyed by canonical part name (`xl/workbook.xml`, without
    /// a leading `/`).
    overrides: BTreeMap<String, PartOverride>,
    /// When set, macros are stripped on write, targeting the provided workbook kind.
    ///
    /// This matches the `target_kind` parameter of
    /// [`crate::streaming::strip_vba_project_streaming_with_kind`].
    strip_macros: Option<WorkbookKind>,
    /// Optional workbook kind enforcement (used to update `[Content_Types].xml`).
    workbook_kind: Option<WorkbookKind>,
    /// Cached set of part names discovered at open time (canonical, without leading `/`).
    part_names: Arc<Vec<String>>,
}

impl fmt::Debug for XlsxLazyPackage {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("XlsxLazyPackage")
            .field("source", &self.source)
            .field("parts", &self.part_names.len())
            .field("overrides", &self.overrides.len())
            .field("strip_macros", &self.strip_macros)
            .field("workbook_kind", &self.workbook_kind)
            .finish()
    }
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
        let part_names = Arc::new(list_part_names(file)?);
        Ok(Self {
            source: Source::Path(Arc::new(path)),
            overrides: BTreeMap::new(),
            strip_macros: None,
            workbook_kind: None,
            part_names,
        })
    }

    /// Create a lazy package from owned ZIP bytes.
    pub fn from_vec(bytes: Vec<u8>) -> Result<Self, XlsxError> {
        let cursor = Cursor::new(bytes.as_slice());
        let part_names = Arc::new(list_part_names(cursor)?);
        Ok(Self {
            source: Source::Bytes(Arc::new(bytes)),
            overrides: BTreeMap::new(),
            strip_macros: None,
            workbook_kind: None,
            part_names,
        })
    }

    /// Create a lazy package by copying ZIP bytes.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self, XlsxError> {
        Self::from_vec(bytes.to_vec())
    }

    fn canonicalize_part_name(name: &str) -> String {
        let trimmed = name.trim_start_matches(|c| c == '/' || c == '\\');
        // Normalize Windows-style separators.
        if trimmed.contains('\\') {
            trimmed.replace('\\', "/")
        } else {
            trimmed.to_string()
        }
    }

    /// Read a single part's bytes.
    ///
    /// This consults in-memory overrides first; if no override is present, it extracts the part
    /// from the underlying ZIP container.
    pub fn read_part(&self, name: &str) -> Result<Option<Vec<u8>>, XlsxError> {
        let canonical = Self::canonicalize_part_name(name);

        // If macro stripping is enabled, treat macro-related parts as missing even if a caller
        // added an override for them after calling `remove_vba_project`.
        if self.strip_macros.is_some() && is_macro_part_name(&canonical) {
            return Ok(None);
        }

        if let Some(override_op) = self.overrides.get(&canonical) {
            match override_op {
                PartOverride::Replace(bytes) | PartOverride::Add(bytes) => {
                    return Ok(Some(bytes.clone()));
                }
                PartOverride::Remove => return Ok(None),
            }
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
            let normalized = normalize_part_name_for_macro_match(name);
            let name = normalized.as_ref();
            if name.eq_ignore_ascii_case("xl/vbaProject.bin") {
                presence.has_vba = true;
            } else if crate::ascii::starts_with_ignore_case(name, "xl/macrosheets/") {
                presence.has_xlm_macrosheets = true;
            } else if crate::ascii::starts_with_ignore_case(name, "xl/dialogsheets/") {
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

        // Once macro stripping is enabled, never allow macro surfaces to be reintroduced into the
        // output via overrides.
        if self.strip_macros.is_some() && is_macro_part_name(&canonical) {
            self.overrides.insert(canonical, PartOverride::Remove);
            return;
        }

        self.overrides
            .insert(canonical, PartOverride::Replace(bytes));
    }

    /// Remove macro-related parts and relationships from the package on write.
    pub fn remove_vba_project(&mut self) -> Result<(), XlsxError> {
        self.remove_vba_project_with_kind(WorkbookKind::Workbook)
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
        self.strip_macros = Some(target_kind);
        // Match `XlsxPackage::{remove_vba_project,remove_vba_project_with_kind}` semantics: stripping
        // macros rewrites the workbook main content type to the requested macro-free kind.
        self.workbook_kind = Some(target_kind);
        // If callers already provided a `[Content_Types].xml` override (including via a prior
        // `enforce_workbook_kind` call), keep it consistent with the macro-strip target. Otherwise
        // we could strip macros and then re-apply a stale workbook content type override in the
        // second streaming pass.
        self.patch_content_types_override_for_kind(target_kind)?;

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

        Ok(())
    }

    fn patch_content_types_override_for_kind(
        &mut self,
        kind: WorkbookKind,
    ) -> Result<(), XlsxError> {
        let ct_name = "[Content_Types].xml";
        let Some(op) = self.overrides.get(ct_name).cloned() else {
            return Ok(());
        };

        let original = match op {
            PartOverride::Replace(bytes) | PartOverride::Add(bytes) => bytes,
            PartOverride::Remove => return Ok(()),
        };

        let mut updated = original.clone();
        if let Some(patched) = crate::rewrite_content_types_workbook_kind(&updated, kind)? {
            updated = patched;
        }
        // If macro stripping is enabled, ensure we don't reintroduce macro part overrides via an
        // existing `[Content_Types].xml` override.
        if self.strip_macros.is_some() {
            if let Some(stripped) = strip_content_types_macro_overrides(&updated)? {
                updated = stripped;
            }
        }
        if updated == original {
            return Ok(());
        };

        self.overrides
            .insert(ct_name.to_string(), PartOverride::Replace(updated));

        Ok(())
    }

    /// Ensure `[Content_Types].xml` advertises the correct workbook content type for the requested
    /// workbook kind.
    pub fn enforce_workbook_kind(&mut self, kind: WorkbookKind) -> Result<(), XlsxError> {
        self.workbook_kind = Some(kind);

        // When macro stripping is enabled, `[Content_Types].xml` will be rewritten as part of the
        // macro-strip streaming pass (using `self.workbook_kind`). Persisting an explicit override
        // here would be based on the *unstripped* content types file and can reintroduce
        // references to deleted macro parts (e.g. `xl/vbaProject.bin`) after the macro-strip pass.
        // It would also force the slower two-pass `strip -> temp file -> apply overrides` path even
        // when there are otherwise no non-macro overrides.
        //
        // If callers already overrode `[Content_Types].xml`, keep it in sync so it doesn't
        // reintroduce a stale workbook content type after macro stripping.
        if self.strip_macros.is_some() {
            self.patch_content_types_override_for_kind(kind)?;
            return Ok(());
        }

        let Some(existing) = self.read_part("[Content_Types].xml")? else {
            // Match `XlsxPackage` semantics: don't synthesize a missing content types file.
            return Ok(());
        };

        let Some(updated) = crate::rewrite_content_types_workbook_kind(&existing, kind)? else {
            // Avoid rewriting when no change is required.
            return Ok(());
        };

        self.overrides.insert(
            "[Content_Types].xml".to_string(),
            PartOverride::Replace(updated),
        );

        Ok(())
    }

    /// Serialize the package to `.xlsx`/`.xlsm` bytes.
    pub fn write_to_bytes(&self) -> Result<Vec<u8>, XlsxError> {
        let mut cursor = Cursor::new(Vec::new());
        self.write_to(&mut cursor)?;
        Ok(cursor.into_inner())
    }

    /// Write the package to an output stream, preserving unchanged ZIP entries via streaming raw
    /// copy.
    #[cfg(not(target_arch = "wasm32"))]
    pub fn write_to<W: Write + Seek>(&self, output: W) -> Result<(), XlsxError> {
        // If macro stripping is enabled, we can sometimes avoid the macro-stripping rewrite:
        // - When the *source* package is macro-free, macro stripping only needs to ensure we don't
        //   emit macro-capable overrides (e.g. added `xl/macrosheets/**` parts).
        // - When the source is macro-capable and there are non-macro overrides, we need an
        //   intermediate temp file so we can apply overrides on top of the macro-stripped base.
        //
        // Note: `part_names` is a snapshot of the *source* ZIP entries (it does not include
        // overrides), so we can use it to cheaply detect macro-capable sources.
        if let Some(_target_kind) = self.strip_macros {
            let source_has_macros = self.part_names.iter().any(|name| is_macro_part_name(name));

            if !source_has_macros {
                // Source is already macro-free; just apply overrides after forcing macro-related
                // overrides to `Remove` and patching workbook kind.
                let overrides = self.build_part_overrides(/*kind_already_handled=*/ false)?;
                let input = self.source.open_reader()?;
                patch_xlsx_streaming_workbook_cell_patches_with_part_overrides(
                    input,
                    output,
                    &WorkbookCellPatches::default(),
                    &overrides,
                )?;
                return Ok(());
            }

            // Source contains macro-capable parts; strip macros first.
            let kind = self.workbook_kind.unwrap_or(WorkbookKind::Workbook);

            // If there are no non-macro overrides to apply afterwards, we can stream the macro
            // stripping directly into the output.
            let has_non_macro_overrides =
                self.overrides.keys().any(|name| !is_macro_part_name(name));
            if !has_non_macro_overrides {
                let input = self.source.open_reader()?;
                strip_vba_project_streaming_with_kind(input, output, kind)?;
                return Ok(());
            }

            // Otherwise, use an intermediate temp file so we can apply overrides on top of the
            // macro-stripped base without inflating all parts into memory.
            let mut tmp = tempfile()?;
            let input = self.source.open_reader()?;
            strip_vba_project_streaming_with_kind(input, &mut tmp, kind)?;

            // Second pass: apply explicit part overrides on top of the macro-stripped package.
            // Note: workbook kind is already handled by the macro-strip streaming pass.
            let overrides = self.build_part_overrides(/*kind_already_handled=*/ true)?;
            tmp.seek(SeekFrom::Start(0))?;
            patch_xlsx_streaming_workbook_cell_patches_with_part_overrides(
                &mut tmp,
                output,
                &WorkbookCellPatches::default(),
                &overrides,
            )?;
            return Ok(());
        }

        let overrides = self.build_part_overrides(/*kind_already_handled=*/ false)?;
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

    /// Stream a patched workbook package to `output`, rewriting only affected worksheet XML parts
    /// (plus any required workbook plumbing like shared strings or styles).
    ///
    /// This composes cell patches with any existing [`PartOverride`] entries stored on the lazy
    /// package.
    #[cfg(not(target_arch = "wasm32"))]
    pub fn patch_cells_to_writer<W: Write + Seek>(
        &self,
        output: W,
        patches: &WorkbookCellPatches,
        recalc_policy: RecalcPolicy,
        style_table: Option<&StyleTable>,
    ) -> Result<(), XlsxError> {
        let needs_style_table = patches.sheets().any(|(_sheet, sheet_patches)| {
            sheet_patches.iter().any(|(_cell, patch)| {
                patch
                    .style_id()
                    .is_some_and(|style_id| style_id != 0)
            })
        });

        fn apply_streaming_patch<R: Read + Seek, W: Write + Seek>(
            input: R,
            output: W,
            patches: &WorkbookCellPatches,
            overrides: &HashMap<String, PartOverride>,
            recalc_policy: RecalcPolicy,
            needs_style_table: bool,
            style_table: Option<&StyleTable>,
        ) -> Result<(), XlsxError> {
            if needs_style_table {
                let style_table = style_table.ok_or_else(|| {
                    XlsxError::Invalid(
                        "style_table is required when WorkbookCellPatches contains non-zero style_id edits"
                            .to_string(),
                    )
                })?;

                patch_xlsx_streaming_workbook_cell_patches_with_styles_and_part_overrides_and_recalc_policy(
                    input,
                    output,
                    patches,
                    style_table,
                    overrides,
                    recalc_policy,
                )?;
            } else {
                patch_xlsx_streaming_workbook_cell_patches_with_part_overrides_and_recalc_policy(
                    input,
                    output,
                    patches,
                    overrides,
                    recalc_policy,
                )?;
            }
            Ok(())
        }

        // Match `write_to` macro stripping behavior:
        // - strip macros if requested,
        // - and (when needed) apply part overrides / cell patches on top of the stripped base using
        //   a second streaming rewrite pass.
        if let Some(_target_kind) = self.strip_macros {
            let source_has_macros = self.part_names.iter().any(|name| is_macro_part_name(name));

            if !source_has_macros {
                // Source is already macro-free; just apply patches + overrides after forcing
                // macro-related overrides to `Remove` and patching workbook kind.
                let overrides = self.build_part_overrides(/*kind_already_handled=*/ false)?;
                let input = self.source.open_reader()?;
                return apply_streaming_patch(
                    input,
                    output,
                    patches,
                    &overrides,
                    recalc_policy,
                    needs_style_table,
                    style_table,
                );
            }

            // Source contains macro-capable parts; strip macros first.
            let kind = self.workbook_kind.unwrap_or(WorkbookKind::Workbook);

            // If there are no non-macro changes (no patches + no non-macro overrides), we can stream
            // macro stripping directly into the output.
            let has_non_macro_overrides = !patches.is_empty()
                || self.overrides.keys().any(|name| !is_macro_part_name(name));
            if !has_non_macro_overrides {
                let input = self.source.open_reader()?;
                strip_vba_project_streaming_with_kind(input, output, kind)?;
                return Ok(());
            }

            // Otherwise, use an intermediate temp file so we can apply patches/overrides on top of
            // the macro-stripped base without inflating all parts into memory.
            let mut tmp = tempfile()?;
            let input = self.source.open_reader()?;
            strip_vba_project_streaming_with_kind(input, &mut tmp, kind)?;

            // Second pass: apply explicit part overrides + cell patches on top of the macro-stripped
            // package. Note: workbook kind is already handled by the macro-strip streaming pass.
            let overrides = self.build_part_overrides(/*kind_already_handled=*/ true)?;
            tmp.seek(SeekFrom::Start(0))?;
            return apply_streaming_patch(
                &mut tmp,
                output,
                patches,
                &overrides,
                recalc_policy,
                needs_style_table,
                style_table,
            );
        }

        let overrides = self.build_part_overrides(/*kind_already_handled=*/ false)?;
        let input = self.source.open_reader()?;
        apply_streaming_patch(
            input,
            output,
            patches,
            &overrides,
            recalc_policy,
            needs_style_table,
            style_table,
        )
    }

    fn effective_part_names(&self) -> impl Iterator<Item = &str> {
        self.part_names
            .iter()
            .map(|s| s.as_str())
            .chain(self.overrides.iter().filter_map(|(name, op)| match op {
                PartOverride::Remove => None,
                PartOverride::Replace(_) | PartOverride::Add(_) => Some(name.as_str()),
            }))
    }

    #[cfg(not(target_arch = "wasm32"))]
    fn build_part_overrides(
        &self,
        kind_already_handled: bool,
    ) -> Result<HashMap<String, PartOverride>, XlsxError> {
        let mut out: HashMap<String, PartOverride> = self
            .overrides
            .iter()
            .map(|(k, v)| (k.clone(), v.clone()))
            .collect();

        // If macro stripping is enabled, force all macro-related overrides to `Remove` so they
        // cannot reintroduce macro-capable surfaces into the output.
        if self.strip_macros.is_some() {
            for (name, op) in out.iter_mut() {
                if is_macro_part_name(name) {
                    *op = PartOverride::Remove;
                }
            }
        }

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
    fn patched_content_types_for_kind(
        &self,
        kind: WorkbookKind,
    ) -> Result<Option<Vec<u8>>, XlsxError> {
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
    let normalized = normalize_part_name_for_macro_match(name);
    let name = normalized.as_ref();

    fn is_macro_surface_part(name: &str) -> bool {
        name.eq_ignore_ascii_case("xl/vbaProject.bin")
            || name.eq_ignore_ascii_case("xl/vbaData.xml")
            || name.eq_ignore_ascii_case("xl/vbaProjectSignature.bin")
            || crate::ascii::starts_with_ignore_case(name, "xl/macrosheets/")
            || crate::ascii::starts_with_ignore_case(name, "xl/dialogsheets/")
            || crate::ascii::starts_with_ignore_case(name, "customui/")
            || crate::ascii::starts_with_ignore_case(name, "xl/activeX/")
            || crate::ascii::starts_with_ignore_case(name, "xl/ctrlProps/")
            || crate::ascii::starts_with_ignore_case(name, "xl/controls/")
    }

    if is_macro_surface_part(name) {
        return true;
    }

    // Relationship parts for deleted macro surfaces should also be deleted. For example, stripping
    // `xl/vbaProject.bin` also deletes `xl/_rels/vbaProject.bin.rels`.
    if crate::ascii::ends_with_ignore_case(name, ".rels") {
        if let Some(source) = source_part_from_rels_part(name) {
            if !source.is_empty() && is_macro_surface_part(&source) {
                return true;
            }
        }
    }

    false
}

fn strip_content_types_macro_overrides(xml: &[u8]) -> Result<Option<Vec<u8>>, XlsxError> {
    let mut reader = XmlReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut writer = XmlWriter::new(Vec::with_capacity(xml.len()));
    let mut buf = Vec::new();
    let mut changed = false;
    let mut skip_depth = 0usize;

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
            Event::Empty(ref e)
                if crate::openxml::local_name(e.name().as_ref())
                    .eq_ignore_ascii_case(b"Override") =>
            {
                if content_types_override_is_macro_part(e)? {
                    changed = true;
                    buf.clear();
                    continue;
                }
                writer.write_event(Event::Empty(e.to_owned()))?;
            }
            Event::Start(ref e)
                if crate::openxml::local_name(e.name().as_ref())
                    .eq_ignore_ascii_case(b"Override") =>
            {
                if content_types_override_is_macro_part(e)? {
                    changed = true;
                    skip_depth = 1;
                    buf.clear();
                    continue;
                }
                writer.write_event(Event::Start(e.to_owned()))?;
            }
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

fn content_types_override_is_macro_part(e: &BytesStart<'_>) -> Result<bool, XlsxError> {
    let mut part_name = None;
    for attr in e.attributes().with_checks(false) {
        let attr = attr?;
        if crate::openxml::local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"PartName") {
            part_name = Some(attr.unescape_value()?.into_owned());
            break;
        }
    }

    let Some(part_name) = part_name else {
        return Ok(false);
    };

    Ok(is_macro_part_name(&part_name))
}

fn source_part_from_rels_part(rels_part: &str) -> Option<String> {
    // Root relationships.
    if rels_part.eq_ignore_ascii_case("_rels/.rels") {
        return Some(String::new());
    }

    if let Some(rels_file) = crate::ascii::strip_prefix_ignore_case(rels_part, "_rels/") {
        let rels_file = crate::ascii::strip_suffix_ignore_case(rels_file, ".rels")?;
        return Some(rels_file.to_string());
    }

    let idx = crate::ascii::rfind_ignore_case(rels_part, "/_rels/")?;
    let dir = &rels_part[..idx];
    let rels_file = &rels_part[idx + "/_rels/".len()..];
    let rels_file = crate::ascii::strip_suffix_ignore_case(rels_file, ".rels")?;
    if dir.is_empty() {
        return Some(rels_file.to_string());
    }

    Some(format!("{dir}/{rels_file}"))
}

fn normalize_part_name_for_macro_match(name: &str) -> Cow<'_, str> {
    let name = name.strip_prefix('/').unwrap_or(name);
    if name.contains('\\') {
        Cow::Owned(name.replace('\\', "/"))
    } else {
        Cow::Borrowed(name)
    }
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
        let canonical = name.trim_start_matches(|c| c == '/' || c == '\\');
        if canonical.contains('\\') {
            names.push(canonical.replace('\\', "/"));
        } else {
            names.push(canonical.to_string());
        }
    }
    Ok(names)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn build_zip(files: &[(&str, &[u8])]) -> Vec<u8> {
        let cursor = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(cursor);
        let options = zip::write::FileOptions::<()>::default()
            .compression_method(zip::CompressionMethod::Stored);
        for (name, bytes) in files {
            zip.start_file(*name, options).unwrap();
            zip.write_all(bytes).unwrap();
        }
        zip.finish().unwrap().into_inner()
    }

    #[test]
    fn clone_shares_source_bytes_and_part_name_cache() {
        let bytes = build_zip(&[("xl/workbook.xml", b"<workbook/>")]);
        let pkg = XlsxLazyPackage::from_vec(bytes).expect("pkg");
        let cloned = pkg.clone();

        match (&pkg.source, &cloned.source) {
            (Source::Bytes(a), Source::Bytes(b)) => {
                assert!(
                    Arc::ptr_eq(a, b),
                    "expected clones to share the same backing ZIP bytes"
                );
            }
            _ => panic!("expected Source::Bytes"),
        }

        assert!(
            Arc::ptr_eq(&pkg.part_names, &cloned.part_names),
            "expected clones to share the cached part name list"
        );
    }

    #[test]
    fn debug_does_not_dump_zip_payload() {
        let bytes = build_zip(&[("xl/workbook.xml", b"<workbook/>")]);
        let pkg = XlsxLazyPackage::from_vec(bytes).expect("pkg");
        let dbg = format!("{pkg:?}");
        assert!(dbg.contains("XlsxLazyPackage"));
        // Avoid printing the entire ZIP payload.
        assert!(!dbg.contains("<workbook/>"));
    }
}
