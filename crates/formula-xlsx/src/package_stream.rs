use std::cell::RefCell;
use std::collections::{BTreeSet, HashMap};
use std::io::{Read, Seek, SeekFrom, Write};

use zip::ZipArchive;

use crate::patch::WorkbookCellPatches;
use crate::streaming::PartOverride;
use crate::zip_util::read_zip_file_bytes_with_limit;
use crate::{MacroPresence, RecalcPolicy, WorkbookKind, XlsxError, MAX_XLSX_PACKAGE_PART_BYTES};

struct PartNamesIter<'a, R: Read + Seek> {
    pkg: &'a StreamingXlsxPackage<R>,
    source_iter: std::collections::btree_set::Iter<'a, String>,
    added_iter: std::collections::btree_set::Iter<'a, String>,
    next_source: Option<&'a String>,
    next_added: Option<&'a String>,
}

impl<'a, R: Read + Seek> PartNamesIter<'a, R> {
    fn new(pkg: &'a StreamingXlsxPackage<R>) -> Self {
        let mut it = Self {
            pkg,
            source_iter: pkg.source_part_names.iter(),
            added_iter: pkg.added_part_names.iter(),
            next_source: None,
            next_added: None,
        };
        it.next_source = it.advance_source();
        it.next_added = it.added_iter.next();
        it
    }

    fn advance_source(&mut self) -> Option<&'a String> {
        while let Some(name) = self.source_iter.next() {
            if self.pkg.is_source_part_removed(name.as_str()) {
                continue;
            }
            return Some(name);
        }
        None
    }
}

impl<'a, R: Read + Seek> Iterator for PartNamesIter<'a, R> {
    type Item = &'a str;

    fn next(&mut self) -> Option<Self::Item> {
        match (self.next_source, self.next_added) {
            (None, None) => None,
            (Some(s), None) => {
                let out = s.as_str();
                self.next_source = self.advance_source();
                Some(out)
            }
            (None, Some(a)) => {
                let out = a.as_str();
                self.next_added = self.added_iter.next();
                Some(out)
            }
            (Some(s), Some(a)) => {
                // Deterministic merge to produce globally-sorted part names (matching
                // `XlsxPackage::part_names` ordering).
                if s <= a {
                    let out = s.as_str();
                    self.next_source = self.advance_source();
                    if s == a {
                        // Should not happen in valid usage (added parts are non-source), but be
                        // defensive.
                        self.next_added = self.added_iter.next();
                    }
                    Some(out)
                } else {
                    let out = a.as_str();
                    self.next_added = self.added_iter.next();
                    Some(out)
                }
            }
        }
    }
}

/// A lazy/streaming XLSX/XLSM package representation that avoids inflating every part into memory.
///
/// This type keeps the source ZIP reader plus an in-memory overlay of [`PartOverride`] mutations.
/// When saving via [`Self::write_to`], unchanged entries are preserved byte-for-byte via
/// `ZipWriter::raw_copy_file`.
pub struct StreamingXlsxPackage<R: Read + Seek> {
    reader: RefCell<R>,
    /// Canonical (normalized) part names present in the source archive.
    ///
    /// Canonicalization:
    /// - strip any leading `/`
    /// - treat `\\` as `/`
    source_part_names: BTreeSet<String>,
    /// Map canonical part name -> ZIP entry key used by the streaming rewriter.
    ///
    /// This is the ZIP entry name with only a leading `/` stripped (matching the streaming
    /// patcher's `canonical_name` computation). It may still contain `\\` if the original archive
    /// used them.
    source_part_name_to_zip_key: HashMap<String, String>,
    /// Map canonical part name -> source zip entry index.
    source_part_name_to_index: HashMap<String, usize>,
    /// Overlay of part overrides keyed by canonical part name.
    ///
    /// Keys are stored in the same form expected by the streaming ZIP rewriter:
    /// - if the part exists in the source ZIP, the key is the original ZIP entry name with only a
    ///   leading `/` stripped (i.e. it may contain `\\` path separators if the archive used them).
    /// - if the part does not exist in the source ZIP (new part), the key is the canonical name
    ///   (`/` separators, no leading `/`).
    part_overrides: HashMap<String, PartOverride>,
    /// Canonical part names that do not exist in the source archive but have been added via
    /// [`Self::set_part`].
    ///
    /// This is tracked separately so [`Self::part_names`] can produce the effective part-name view
    /// without cloning all source part names (important for large workbooks).
    added_part_names: BTreeSet<String>,
}

impl<R: Read + Seek> std::fmt::Debug for StreamingXlsxPackage<R> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("StreamingXlsxPackage")
            .field("source_part_names", &self.source_part_names)
            .field("part_overrides", &self.part_overrides)
            .field("added_part_names", &self.added_part_names)
            .finish()
    }
}

impl<R: Read + Seek> StreamingXlsxPackage<R> {
    fn resolve_source_part_name<'a>(&'a self, canonical: &str) -> Option<&'a str> {
        if let Some(name) = self.source_part_names.get(canonical) {
            return Some(name.as_str());
        }
        self.source_part_names
            .iter()
            .find(|name| crate::zip_util::zip_part_names_equivalent(name.as_str(), canonical))
            .map(String::as_str)
    }

    /// Open an XLSX/XLSM package from an arbitrary `Read + Seek` source.
    pub fn from_reader(mut reader: R) -> Result<Self, XlsxError> {
        // Scan the central directory once to build a canonical name index without inflating part
        // payloads.
        reader.seek(SeekFrom::Start(0))?;

        let mut source_part_names: BTreeSet<String> = BTreeSet::new();
        let mut source_part_name_to_zip_key: HashMap<String, String> = HashMap::new();
        let mut source_part_name_to_index: HashMap<String, usize> = HashMap::new();

        {
            let mut archive = ZipArchive::new(&mut reader)?;
            for i in 0..archive.len() {
                let file = archive.by_index(i)?;
                if file.is_dir() {
                    continue;
                }
                let raw_name = file.name().to_string();
                let zip_key = raw_name.strip_prefix('/').unwrap_or(raw_name.as_str());
                let canonical = canonical_part_name(&raw_name);

                source_part_names.insert(canonical.clone());

                // In degenerate cases where a ZIP contains duplicate names that normalize to the same
                // canonical value (e.g. `xl/workbook.xml` and `xl\\workbook.xml`), keep the first
                // index/key we saw. XLSX producers should not emit such archives.
                source_part_name_to_zip_key
                    .entry(canonical.clone())
                    .or_insert_with(|| zip_key.to_string());
                source_part_name_to_index.entry(canonical).or_insert(i);
            }
        }

        Ok(Self {
            reader: RefCell::new(reader),
            source_part_names,
            source_part_name_to_zip_key,
            source_part_name_to_index,
            part_overrides: HashMap::new(),
            added_part_names: BTreeSet::new(),
        })
    }

    /// Replace (or add) a part in the output package.
    ///
    /// When the part exists in the source package, this is represented as
    /// [`PartOverride::Replace`]. When it does not exist, this is represented as
    /// [`PartOverride::Add`].
    pub fn set_part(&mut self, name: &str, bytes: Vec<u8>) {
        let canonical_input = canonical_part_name(name);
        let (canonical, exists_in_source) = match self.resolve_source_part_name(&canonical_input) {
            Some(found) => (found.to_string(), true),
            None => (canonical_input, false),
        };
        let override_key = self
            .source_part_name_to_zip_key
            .get(&canonical)
            .cloned()
            .unwrap_or_else(|| canonical.clone());
        let op = if exists_in_source {
            PartOverride::Replace(bytes)
        } else {
            PartOverride::Add(bytes)
        };
        self.part_overrides.insert(override_key, op);
        if !exists_in_source {
            self.added_part_names.insert(canonical);
        } else {
            // Defensive: if callers previously added a part before we indexed the source correctly,
            // ensure we don't keep treating it as "added" once we know it's actually a source part.
            self.added_part_names.remove(&canonical);
        }
    }

    /// Remove a part from the output package.
    pub fn remove_part(&mut self, name: &str) {
        let canonical_input = canonical_part_name(name);
        let (canonical, exists_in_source) = match self.resolve_source_part_name(&canonical_input) {
            Some(found) => (found.to_string(), true),
            None => (canonical_input, false),
        };
        let override_key = self
            .source_part_name_to_zip_key
            .get(&canonical)
            .cloned()
            .unwrap_or_else(|| canonical.clone());
        self.part_overrides
            .insert(override_key, PartOverride::Remove);
        if !exists_in_source {
            self.added_part_names.remove(&canonical);
        }
    }

    /// Access the raw part override map (useful for debugging/testing).
    pub fn part_overrides(&self) -> &HashMap<String, PartOverride> {
        &self.part_overrides
    }

    /// Iterate the effective part names in the package (source parts plus overrides).
    ///
    /// Part names are returned in canonical form (no leading `/`, `/` separators).
    pub fn part_names(&self) -> impl Iterator<Item = &str> + '_ {
        PartNamesIter::new(self)
    }

    /// Detect whether the effective package contains any macro-capable content.
    ///
    /// Semantics match [`crate::XlsxPackage::macro_presence`].
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

    /// Read a single part, consulting overrides first and otherwise reading from the source ZIP.
    pub fn read_part(&self, name: &str) -> Result<Option<Vec<u8>>, XlsxError> {
        let canonical_input = canonical_part_name(name);

        // Added parts are keyed by canonical name directly.
        if let Some(override_op) = self.part_overrides.get(&canonical_input) {
            match override_op {
                PartOverride::Remove => return Ok(None),
                PartOverride::Replace(bytes) | PartOverride::Add(bytes) => {
                    return Ok(Some(bytes.clone()))
                }
            }
        }

        let Some(canonical) = self.resolve_source_part_name(&canonical_input) else {
            return Ok(None);
        };

        let override_key = self
            .source_part_name_to_zip_key
            .get(canonical)
            .map(String::as_str)
            .unwrap_or(canonical);

        if let Some(override_op) = self.part_overrides.get(override_key) {
            match override_op {
                PartOverride::Remove => return Ok(None),
                PartOverride::Replace(bytes) | PartOverride::Add(bytes) => {
                    return Ok(Some(bytes.clone()))
                }
            }
        }

        let Some(&idx) = self.source_part_name_to_index.get(canonical) else {
            return Ok(None);
        };

        let mut reader = self.reader.borrow_mut();
        reader.seek(SeekFrom::Start(0))?;
        let mut archive = ZipArchive::new(&mut *reader)?;
        let mut file = archive.by_index(idx)?;
        if file.is_dir() {
            return Ok(None);
        }
        let buf =
            read_zip_file_bytes_with_limit(&mut file, canonical, MAX_XLSX_PACKAGE_PART_BYTES)?;
        Ok(Some(buf))
    }

    /// Ensure `[Content_Types].xml` advertises the correct workbook content type for the requested
    /// workbook kind.
    ///
    /// This matches [`crate::XlsxPackage::enforce_workbook_kind`] but avoids materializing the full
    /// OPC package: only `[Content_Types].xml` is read and rewritten when needed.
    pub fn enforce_workbook_kind(&mut self, kind: WorkbookKind) -> Result<(), XlsxError> {
        let Some(content_types_xml) = self.read_part("[Content_Types].xml")? else {
            // Match `XlsxPackage` semantics: don't synthesize content types when missing.
            return Ok(());
        };

        let Some(updated) = crate::rewrite_content_types_workbook_kind(&content_types_xml, kind)?
        else {
            return Ok(());
        };

        // Always store as `Replace` so that existing parts are rewritten in-place (and missing
        // parts are appended), matching the streaming patcher semantics.
        let canonical_input = canonical_part_name("[Content_Types].xml");
        let canonical = self
            .resolve_source_part_name(&canonical_input)
            .unwrap_or(canonical_input.as_str())
            .to_string();
        let override_key = self
            .source_part_name_to_zip_key
            .get(&canonical)
            .cloned()
            .unwrap_or(canonical);
        self.part_overrides
            .insert(override_key, PartOverride::Replace(updated));

        Ok(())
    }

    /// Write the effective package to a new ZIP stream, raw-copying unchanged entries.
    pub fn write_to<W: Write + Seek>(&self, output: W) -> Result<(), XlsxError> {
        let patches = WorkbookCellPatches::default();

        let mut reader = self.reader.borrow_mut();
        reader.seek(SeekFrom::Start(0))?;
        crate::streaming::patch_xlsx_streaming_workbook_cell_patches_with_part_overrides_and_recalc_policy(
            &mut *reader,
            output,
            &patches,
            &self.part_overrides,
            RecalcPolicy::PRESERVE,
        )?;
        Ok(())
    }

    fn is_source_part_removed(&self, canonical_name: &str) -> bool {
        let override_key = self
            .source_part_name_to_zip_key
            .get(canonical_name)
            .map(String::as_str)
            .unwrap_or(canonical_name);
        matches!(
            self.part_overrides.get(override_key),
            Some(PartOverride::Remove)
        )
    }
}

/// Path-based constructor for non-wasm builds.
#[cfg(not(target_arch = "wasm32"))]
impl StreamingXlsxPackage<std::fs::File> {
    pub fn from_path(path: impl AsRef<std::path::Path>) -> Result<Self, XlsxError> {
        let file = std::fs::File::open(path)?;
        Self::from_reader(file)
    }
}

fn canonical_part_name(name: &str) -> String {
    // Normalize separators first, then strip any leading `/` (including those produced by
    // converting leading `\\` to `/`).
    let replaced = name.replace('\\', "/");
    replaced.trim_start_matches('/').to_string()
}
