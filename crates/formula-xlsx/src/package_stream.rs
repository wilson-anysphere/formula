use std::cell::RefCell;
use std::collections::{BTreeSet, HashMap};
use std::io::{Read, Seek, SeekFrom, Write};

use zip::ZipArchive;

use crate::patch::WorkbookCellPatches;
use crate::streaming::PartOverride;
use crate::{MacroPresence, RecalcPolicy, XlsxError};

trait ReadSeek: Read + Seek {}
impl<T: Read + Seek> ReadSeek for T {}

/// A lazy/streaming XLSX/XLSM package representation that avoids inflating every part into memory.
///
/// This type keeps the source ZIP reader plus an in-memory overlay of [`PartOverride`] mutations.
/// When saving via [`Self::write_to`], unchanged entries are preserved byte-for-byte via
/// `ZipWriter::raw_copy_file`.
pub struct StreamingXlsxPackage {
    reader: RefCell<Box<dyn ReadSeek>>,
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
    part_overrides: HashMap<String, PartOverride>,
}

impl std::fmt::Debug for StreamingXlsxPackage {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("StreamingXlsxPackage")
            .field("source_part_names", &self.source_part_names)
            .field("part_overrides", &self.part_overrides)
            .finish()
    }
}

impl StreamingXlsxPackage {
    /// Open an XLSX/XLSM package from an arbitrary `Read + Seek` source.
    pub fn from_reader<R: Read + Seek + 'static>(reader: R) -> Result<Self, XlsxError> {
        let mut boxed: Box<dyn ReadSeek> = Box::new(reader);

        // Scan the central directory once to build a canonical name index without inflating part
        // payloads.
        boxed.seek(SeekFrom::Start(0))?;
        let mut archive = ZipArchive::new(&mut boxed)?;

        let mut source_part_names: BTreeSet<String> = BTreeSet::new();
        let mut source_part_name_to_zip_key: HashMap<String, String> = HashMap::new();
        let mut source_part_name_to_index: HashMap<String, usize> = HashMap::new();

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

        Ok(Self {
            reader: RefCell::new(boxed),
            source_part_names,
            source_part_name_to_zip_key,
            source_part_name_to_index,
            part_overrides: HashMap::new(),
        })
    }

    /// Open an XLSX/XLSM package from a filesystem path (non-wasm).
    #[cfg(not(target_arch = "wasm32"))]
    pub fn from_path(path: impl AsRef<std::path::Path>) -> Result<Self, XlsxError> {
        use std::fs::File;
        let file = File::open(path)?;
        Self::from_reader(file)
    }

    /// Replace (or add) a part in the output package.
    ///
    /// When the part exists in the source package, this is represented as
    /// [`PartOverride::Replace`]. When it does not exist, this is represented as
    /// [`PartOverride::Add`].
    pub fn set_part(&mut self, name: &str, bytes: Vec<u8>) {
        let canonical = canonical_part_name(name);
        let op = if self.source_part_names.contains(&canonical) {
            PartOverride::Replace(bytes)
        } else {
            PartOverride::Add(bytes)
        };
        self.part_overrides.insert(canonical, op);
    }

    /// Remove a part from the output package.
    pub fn remove_part(&mut self, name: &str) {
        let canonical = canonical_part_name(name);
        self.part_overrides.insert(canonical, PartOverride::Remove);
    }

    /// Access the raw part override map (useful for debugging/testing).
    pub fn part_overrides(&self) -> &HashMap<String, PartOverride> {
        &self.part_overrides
    }

    /// Iterate the effective part names in the package (source parts plus overrides).
    ///
    /// Part names are returned in canonical form (no leading `/`, `/` separators).
    pub fn part_names(&self) -> impl Iterator<Item = String> {
        effective_part_names(&self.source_part_names, &self.part_overrides).into_iter()
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
            let name = name.strip_prefix('/').unwrap_or(name.as_str());
            let name = name.replace('\\', "/");
            if name == "xl/vbaProject.bin" {
                presence.has_vba = true;
            }
            if name.starts_with("xl/macrosheets/") {
                presence.has_xlm_macrosheets = true;
            }
            if name.starts_with("xl/dialogsheets/") {
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
        let canonical = canonical_part_name(name);

        if let Some(override_op) = self.part_overrides.get(&canonical) {
            match override_op {
                PartOverride::Remove => return Ok(None),
                PartOverride::Replace(bytes) | PartOverride::Add(bytes) => {
                    return Ok(Some(bytes.clone()))
                }
            }
        }

        let Some(&idx) = self.source_part_name_to_index.get(&canonical) else {
            return Ok(None);
        };

        let mut reader = self.reader.borrow_mut();
        reader.seek(SeekFrom::Start(0))?;
        let mut archive = ZipArchive::new(&mut *reader)?;
        let mut file = archive.by_index(idx)?;
        if file.is_dir() {
            return Ok(None);
        }
        let mut buf = Vec::with_capacity(file.size() as usize);
        file.read_to_end(&mut buf)?;
        Ok(Some(buf))
    }

    /// Write the effective package to a new ZIP stream, raw-copying unchanged entries.
    pub fn write_to<W: Write + Seek>(&self, output: W) -> Result<(), XlsxError> {
        let overrides = self.streaming_overrides();
        let patches = WorkbookCellPatches::default();

        let mut reader = self.reader.borrow_mut();
        reader.seek(SeekFrom::Start(0))?;
        crate::streaming::patch_xlsx_streaming_workbook_cell_patches_with_part_overrides_and_recalc_policy(
            &mut *reader,
            output,
            &patches,
            &overrides,
            RecalcPolicy::PRESERVE,
        )?;
        Ok(())
    }

    fn streaming_overrides(&self) -> HashMap<String, PartOverride> {
        let mut out = HashMap::with_capacity(self.part_overrides.len());
        for (canonical, op) in &self.part_overrides {
            // If the part exists in the source archive, translate to the ZIP entry key expected by
            // the streaming rewriter (which only strips leading `/`).
            if let Some(zip_key) = self.source_part_name_to_zip_key.get(canonical) {
                out.insert(zip_key.clone(), op.clone());
            } else {
                // New part: write using the canonical name (forward slashes, no leading `/`).
                out.insert(canonical.clone(), op.clone());
            }
        }
        out
    }
}

fn effective_part_names(
    source_part_names: &BTreeSet<String>,
    part_overrides: &HashMap<String, PartOverride>,
) -> BTreeSet<String> {
    let mut out = source_part_names.clone();
    for (name, op) in part_overrides {
        match op {
            PartOverride::Remove => {
                out.remove(name);
            }
            PartOverride::Replace(_) | PartOverride::Add(_) => {
                out.insert(name.clone());
            }
        }
    }
    out
}

fn canonical_part_name(name: &str) -> String {
    // Normalize separators first, then strip any leading `/` (including those produced by
    // converting leading `\\` to `/`).
    let replaced = name.replace('\\', "/");
    replaced.trim_start_matches('/').to_string()
}
