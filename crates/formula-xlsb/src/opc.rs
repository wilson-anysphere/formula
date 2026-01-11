use crate::biff12_varint;
use crate::parser::Error as ParseError;
use crate::parser::{
    biff12, parse_shared_strings, parse_sheet, parse_sheet_stream, parse_workbook, Cell, CellValue,
    SheetData, SheetMeta, WorkbookProperties,
};
use crate::patch::{patch_sheet_bin, patch_sheet_bin_streaming, CellEdit};
use crate::shared_strings_write::SharedStringsWriter;
use crate::styles::Styles;
use crate::workbook_context::WorkbookContext;
use crate::SharedString;
use quick_xml::events::Event;
use quick_xml::Reader as XmlReader;
use quick_xml::Writer as XmlWriter;
use std::collections::{HashMap, HashSet};
use std::fs::File;
use std::io::{self, Cursor, Read, Seek, Write};
use std::path::{Path, PathBuf};
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

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
    /// leave this off and rely on re-reading the source file when writing.
    pub preserve_worksheets: bool,
}

impl Default for OpenOptions {
    fn default() -> Self {
        Self {
            preserve_unknown_parts: true,
            preserve_parsed_parts: true,
            preserve_worksheets: false,
        }
    }
}

/// An opened XLSB workbook.
///
/// This type keeps enough metadata to stream worksheets on demand. It also optionally stores
/// raw bytes for parts we do not understand, enabling round-trip preservation later.
#[derive(Debug)]
pub struct XlsbWorkbook {
    path: PathBuf,
    sheets: Vec<SheetMeta>,
    shared_strings: Vec<String>,
    shared_strings_table: Vec<SharedString>,
    workbook_context: WorkbookContext,
    workbook_properties: WorkbookProperties,
    styles: Styles,
    preserved_parts: HashMap<String, Vec<u8>>,
    preserve_parsed_parts: bool,
}

impl XlsbWorkbook {
    pub fn open(path: impl AsRef<Path>) -> Result<Self, ParseError> {
        Self::open_with_options(path, OpenOptions::default())
    }

    pub fn open_with_options(
        path: impl AsRef<Path>,
        options: OpenOptions,
    ) -> Result<Self, ParseError> {
        let path = path.as_ref().to_path_buf();
        let file = File::open(&path)?;
        let mut zip = ZipArchive::new(file)?;

        let mut preserved_parts = HashMap::new();

        // Preserve package-level plumbing we don't parse but will need to re-emit on round-trip.
        preserve_part(&mut zip, &mut preserved_parts, "[Content_Types].xml")?;
        preserve_part(&mut zip, &mut preserved_parts, "_rels/.rels")?;

        let workbook_rels_bytes = read_zip_entry(&mut zip, "xl/_rels/workbook.bin.rels")?
            .ok_or_else(|| ParseError::Zip(zip::result::ZipError::FileNotFound))?;
        preserved_parts.insert(
            "xl/_rels/workbook.bin.rels".to_string(),
            workbook_rels_bytes.clone(),
        );
        let workbook_rels = parse_relationships(&workbook_rels_bytes)?;

        // Styles are required for round-trip. We also parse `cellXfs` so callers can
        // resolve per-cell `style` indices to number formats (e.g. dates).
        let styles_bin = read_zip_entry(&mut zip, "xl/styles.bin")?;
        let styles = match styles_bin.as_deref() {
            Some(bytes) => Styles::parse(bytes).unwrap_or_default(),
            None => Styles::default(),
        };
        if let Some(bytes) = styles_bin {
            preserved_parts.insert("xl/styles.bin".to_string(), bytes);
        }

        let workbook_bin = {
            let mut wb = zip.by_name("xl/workbook.bin")?;
            let mut bytes = Vec::with_capacity(wb.size() as usize);
            wb.read_to_end(&mut bytes)?;
            bytes
        };

        let (sheets, workbook_context, workbook_properties) =
            parse_workbook(&mut Cursor::new(&workbook_bin), &workbook_rels)?;
        if options.preserve_parsed_parts {
            preserved_parts.insert("xl/workbook.bin".to_string(), workbook_bin);
        }

        let shared_strings = match zip.by_name("xl/sharedStrings.bin") {
            Ok(mut sst) => {
                let mut bytes = Vec::with_capacity(sst.size() as usize);
                sst.read_to_end(&mut bytes)?;
                let table = parse_shared_strings(&mut Cursor::new(&bytes))?;
                let strings = table.iter().map(|s| s.plain_text().to_string()).collect();
                if options.preserve_parsed_parts {
                    preserved_parts.insert("xl/sharedStrings.bin".to_string(), bytes);
                }
                (strings, table)
            }
            Err(zip::result::ZipError::FileNotFound) => (Vec::new(), Vec::new()),
            Err(e) => return Err(e.into()),
        };

        let (shared_strings, shared_strings_table) = shared_strings;

        let known_parts: HashSet<&str> = [
            "[Content_Types].xml",
            "_rels/.rels",
            "xl/workbook.bin",
            "xl/_rels/workbook.bin.rels",
            "xl/sharedStrings.bin",
            "xl/styles.bin",
        ]
        .into_iter()
        .collect();

        let worksheet_paths: HashSet<String> = sheets.iter().map(|s| s.part_path.clone()).collect();
        if options.preserve_unknown_parts {
            for name in zip.file_names().map(str::to_string).collect::<Vec<_>>() {
                let is_known =
                    known_parts.contains(name.as_str()) || worksheet_paths.contains(&name);
                if is_known {
                    continue;
                }
                if let Ok(mut entry) = zip.by_name(&name) {
                    let mut bytes = Vec::with_capacity(entry.size() as usize);
                    entry.read_to_end(&mut bytes)?;
                    preserved_parts.insert(name, bytes);
                }
            }
        }

        if options.preserve_worksheets {
            for sheet in &sheets {
                if let Ok(mut entry) = zip.by_name(&sheet.part_path) {
                    let mut bytes = Vec::with_capacity(entry.size() as usize);
                    entry.read_to_end(&mut bytes)?;
                    preserved_parts.insert(sheet.part_path.clone(), bytes);
                }
            }
        }

        Ok(Self {
            path,
            sheets,
            shared_strings,
            shared_strings_table,
            workbook_context,
            workbook_properties,
            styles,
            preserved_parts,
            preserve_parsed_parts: options.preserve_parsed_parts,
        })
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

    /// Workbook styles parsed from `xl/styles.bin`.
    pub fn styles(&self) -> &Styles {
        &self.styles
    }

    /// Parse `xl/styles.bin` using locale-aware built-in number formats.
    ///
    /// This is a convenience wrapper around [`Styles::parse_with_locale`] that
    /// uses the preserved `xl/styles.bin` bytes from this workbook.
    pub fn styles_with_locale(&self, locale: formula_format::Locale) -> Option<Result<Styles, ParseError>> {
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

    /// Raw `xl/styles.bin`, preserved for round-trip.
    pub fn styles_bin(&self) -> Option<&[u8]> {
        self.preserved_parts
            .get("xl/styles.bin")
            .map(|v| v.as_slice())
    }

    /// Read a worksheet by index and return all discovered cells.
    ///
    /// For large sheets you likely want `for_each_cell` instead.
    pub fn read_sheet(&self, sheet_index: usize) -> Result<SheetData, ParseError> {
        let meta = self
            .sheets
            .get(sheet_index)
            .ok_or(ParseError::SheetIndexOutOfBounds(sheet_index))?;

        let file = File::open(&self.path)?;
        let mut zip = ZipArchive::new(file)?;
        let mut sheet = zip.by_name(&meta.part_path)?;

        parse_sheet(
            &mut sheet,
            &self.shared_strings,
            &self.workbook_context,
            self.preserve_parsed_parts,
        )
    }

    /// Stream cells from a worksheet without materializing the whole sheet.
    pub fn for_each_cell<F>(&self, sheet_index: usize, mut f: F) -> Result<(), ParseError>
    where
        F: FnMut(Cell),
    {
        let meta = self
            .sheets
            .get(sheet_index)
            .ok_or(ParseError::SheetIndexOutOfBounds(sheet_index))?;

        let file = File::open(&self.path)?;
        let mut zip = ZipArchive::new(file)?;
        let mut sheet = zip.by_name(&meta.part_path)?;

        parse_sheet_stream(
            &mut sheet,
            &self.shared_strings,
            &self.workbook_context,
            self.preserve_parsed_parts,
            |cell| f(cell),
        )?;
        Ok(())
    }

    /// Save the workbook as a new `.xlsb` file.
    ///
    /// This is currently a *lossless* package writer: it repackages the original XLSB ZIP
    /// container by copying every entry's uncompressed payload byte-for-byte.
    ///
    /// The writer always reads entries from the source workbook at `self.path`. If an entry name
    /// exists in [`XlsbWorkbook::preserved_parts`], that byte payload is used as an override. This
    /// provides a forward-compatible hook for future code to patch individual parts (for example
    /// to write modified worksheets) while keeping the rest of the package intact.
    ///
    /// How [`OpenOptions`] affects `save_as`:
    /// - `preserve_unknown_parts`: stores raw bytes for unknown ZIP entries in `preserved_parts`,
    ///   but `save_as` will still copy them from the source file even when this is `false`.
    /// - `preserve_parsed_parts`: stores raw bytes for `xl/workbook.bin` and `xl/sharedStrings.bin`
    ///   so they can be re-emitted without re-reading those ZIP entries.
    /// - `preserve_worksheets`: stores raw bytes for worksheet `.bin` parts (can be large). When
    ///   `false`, worksheets are streamed from the source ZIP during `save_as`.
    ///
    /// If you need to override specific parts (e.g. a patched worksheet stream), use
    /// [`XlsbWorkbook::save_with_part_overrides`].
    pub fn save_as(&self, dest: impl AsRef<Path>) -> Result<(), ParseError> {
        self.save_with_part_overrides(dest, &HashMap::new())
    }

    /// Save the workbook with an updated numeric cell value.
    ///
    /// This is a convenience wrapper around the in-memory worksheet patcher ([`patch_sheet_bin`])
    /// plus the part override writer
    /// ([`XlsbWorkbook::save_with_part_overrides`]).
    ///
    /// Note: this may insert missing `BrtRow` / cell records inside `BrtSheetData` if the target
    /// cell does not already exist in the worksheet stream.
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
                new_formula: None,
                shared_string_index: None,
            }],
        )
    }

    /// Save the workbook with a set of edits for a single worksheet.
    ///
    /// This loads `xl/worksheets/sheetN.bin` into memory. For very large worksheets, consider
    /// [`XlsbWorkbook::save_with_cell_edits_streaming`].
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
            let file = File::open(&self.path)?;
            let mut zip = ZipArchive::new(file)?;
            let mut entry = zip.by_name(&sheet_part)?;
            let mut bytes = Vec::with_capacity(entry.size() as usize);
            entry.read_to_end(&mut bytes)?;
            bytes
        };

        let patched = patch_sheet_bin(&sheet_bytes, edits)?;
        self.save_with_part_overrides(dest, &HashMap::from([(sheet_part, patched)]))
    }

    /// Save the workbook with a set of edits for a single worksheet, updating `xl/sharedStrings.bin`
    /// as needed so shared-string (`BrtCellIsst`) cells can stay as shared-string references.
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
            let file = File::open(&self.path)?;
            let mut zip = ZipArchive::new(file)?;
            let mut entry = zip.by_name(&sheet_part)?;
            let mut bytes = Vec::with_capacity(entry.size() as usize);
            entry.read_to_end(&mut bytes)?;
            bytes
        };

        let shared_strings_bytes = match self.preserved_parts.get("xl/sharedStrings.bin") {
            Some(bytes) => bytes.clone(),
            None => {
                let file = File::open(&self.path)?;
                let mut zip = ZipArchive::new(file)?;
                match read_zip_entry(&mut zip, "xl/sharedStrings.bin")? {
                    Some(bytes) => bytes,
                    None => {
                        // Workbook has no shared string table. Fall back to the generic patcher
                        // which may convert shared-string cells to inline strings.
                        return self.save_with_cell_edits(dest, sheet_index, edits);
                    }
                }
            }
        };

        // Identify which edits target existing `BrtCellIsst` (`0x0007`) cells so we can write a
        // shared-string index instead of rewriting them as inline strings.
        let mut text_targets = HashSet::new();
        for edit in edits {
            if matches!(edit.new_value, CellValue::Text(_)) {
                text_targets.insert((edit.row, edit.col));
            }
        }
        let cell_record_ids = if text_targets.is_empty() {
            HashMap::new()
        } else {
            sheet_cell_record_ids(&sheet_bytes, &text_targets)?
        };

        let mut sst = SharedStringsWriter::new(shared_strings_bytes)?;

        let mut updated_edits = edits.to_vec();
        for edit in &mut updated_edits {
            let CellValue::Text(text) = &edit.new_value else {
                continue;
            };

            let coord = (edit.row, edit.col);
            match cell_record_ids.get(&coord) {
                Some(&biff12::STRING) => {
                    // Preserve existing shared-string cells as shared-string references.
                    edit.shared_string_index = Some(sst.intern_plain(text)?);
                }
                None => {
                    // Newly inserted text cells can also use the shared string table, since we're
                    // already updating `xl/sharedStrings.bin`.
                    edit.shared_string_index = Some(sst.intern_plain(text)?);
                }
                _ => {}
            }
        }

        let updated_shared_strings_bytes = sst.into_bytes()?;
        let patched_sheet = patch_sheet_bin(&sheet_bytes, &updated_edits)?;

        self.save_with_part_overrides(
            dest,
            &HashMap::from([
                (sheet_part, patched_sheet),
                (
                    "xl/sharedStrings.bin".to_string(),
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

        self.save_with_part_overrides_streaming(dest, &HashMap::new(), &sheet_part, |input, output| {
            patch_sheet_bin_streaming(input, output, edits)
        })
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
    /// - `xl/_rels/workbook.bin.rels`
    ///
    /// A stale calcChain can cause Excel to open with incorrect cached results or spend time
    /// rebuilding the chain.
    pub fn save_with_part_overrides(
        &self,
        dest: impl AsRef<Path>,
        overrides: &HashMap<String, Vec<u8>>,
    ) -> Result<(), ParseError> {
        let dest = dest.as_ref();

        let file = File::open(&self.path)?;
        let mut zip = ZipArchive::new(file)?;

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
                "xl/_rels/workbook.bin.rels",
            )?;
            if let Some(workbook_rels) = workbook_rels {
                updated_workbook_rels = Some(remove_calc_chain_from_workbook_rels(&workbook_rels)?);
            }

            let workbook_bin = get_part_bytes(
                &mut zip,
                &self.preserved_parts,
                overrides,
                "xl/workbook.bin",
            )?;
            if let Some(workbook_bin) = workbook_bin {
                if let Some(patched) = patch_workbook_bin_full_calc_on_load(&workbook_bin)? {
                    updated_workbook_bin = Some(patched);
                }
            }
        }

        let out = File::create(dest)?;
        let mut writer = ZipWriter::new(out);

        // Use a consistent compression method for output. This does *not* affect payload
        // preservation: we always copy/write the uncompressed part bytes.
        let options = FileOptions::default().compression_method(CompressionMethod::Deflated);

        let mut used_overrides: HashSet<String> = HashSet::new();

        for i in 0..zip.len() {
            let mut entry = zip.by_index(i)?;
            let name = entry.name().to_string();

            if entry.is_dir() {
                // Directory entries are optional in ZIPs, but we recreate them when present to
                // preserve the package layout more closely.
                writer.add_directory(name, options)?;
                continue;
            }

            // Drop calcChain when any worksheet was edited.
            if edited && is_calc_chain_part_name(&name) {
                if overrides.contains_key(&name) {
                    used_overrides.insert(name);
                }
                continue;
            }

            writer.start_file(name.as_str(), options)?;

            // When invalidating calcChain, we may need to rewrite XML parts even if they're
            // present in `overrides`.
            if edited && name == "[Content_Types].xml" {
                if let Some(updated) = &updated_content_types {
                    if overrides.contains_key(&name) {
                        used_overrides.insert(name.clone());
                    }
                    writer.write_all(updated)?;
                    continue;
                }
            }
            if edited && name == "xl/_rels/workbook.bin.rels" {
                if let Some(updated) = &updated_workbook_rels {
                    if overrides.contains_key(&name) {
                        used_overrides.insert(name.clone());
                    }
                    writer.write_all(updated)?;
                    continue;
                }
            }
            if edited && name == "xl/workbook.bin" {
                if let Some(updated) = &updated_workbook_bin {
                    if overrides.contains_key(&name) {
                        used_overrides.insert(name.clone());
                    }
                    writer.write_all(updated)?;
                    continue;
                }
            }

            if let Some(bytes) = overrides.get(&name) {
                used_overrides.insert(name.clone());
                writer.write_all(bytes)?;
            } else if let Some(bytes) = self.preserved_parts.get(&name) {
                writer.write_all(bytes)?;
            } else {
                io::copy(&mut entry, &mut writer)?;
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

        writer.finish()?;
        Ok(())
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
        if overrides.contains_key(stream_part) {
            return Err(ParseError::Io(io::Error::new(
                io::ErrorKind::InvalidInput,
                format!(
                    "streaming override conflicts with byte override for part: {stream_part}"
                ),
            )));
        }

        let dest = dest.as_ref();

        let file = File::open(&self.path)?;
        let mut zip = ZipArchive::new(file)?;

        let edited_by_bytes = worksheets_edited(&mut zip, &self.sheets, overrides)?;
        let stream_is_worksheet = self.sheets.iter().any(|s| s.part_path == stream_part);

        let edited_by_stream = if stream_is_worksheet {
            let mut sink = io::sink();
            if let Some(bytes) = self.preserved_parts.get(stream_part) {
                let mut cursor = Cursor::new(bytes);
                stream_override(&mut cursor, &mut sink)?
            } else {
                let mut entry = zip.by_name(stream_part)?;
                stream_override(&mut entry, &mut sink)?
            }
        } else {
            false
        };

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
                updated_content_types = Some(remove_calc_chain_from_content_types(&content_types)?);
            }

            let workbook_rels = get_part_bytes(
                &mut zip,
                &self.preserved_parts,
                overrides,
                "xl/_rels/workbook.bin.rels",
            )?;
            if let Some(workbook_rels) = workbook_rels {
                updated_workbook_rels = Some(remove_calc_chain_from_workbook_rels(&workbook_rels)?);
            }

            let workbook_bin =
                get_part_bytes(&mut zip, &self.preserved_parts, overrides, "xl/workbook.bin")?;
            if let Some(workbook_bin) = workbook_bin {
                if let Some(patched) = patch_workbook_bin_full_calc_on_load(&workbook_bin)? {
                    updated_workbook_bin = Some(patched);
                }
            }
        }

        let out = File::create(dest)?;
        let mut writer = ZipWriter::new(out);

        // Use a consistent compression method for output. This does *not* affect payload
        // preservation: we always copy/write the uncompressed part bytes.
        let options = FileOptions::default().compression_method(CompressionMethod::Deflated);

        let mut used_overrides: HashSet<String> = HashSet::new();
        let mut used_stream_override = false;

        for i in 0..zip.len() {
            let mut entry = zip.by_index(i)?;
            let name = entry.name().to_string();

            if entry.is_dir() {
                writer.add_directory(name, options)?;
                continue;
            }

            // Drop calcChain when any worksheet was edited.
            if edited && is_calc_chain_part_name(&name) {
                if overrides.contains_key(&name) {
                    used_overrides.insert(name);
                }
                continue;
            }

            writer.start_file(name.as_str(), options)?;

            // When invalidating calcChain, we may need to rewrite XML parts even if they're
            // present in `overrides`.
            if edited && name == "[Content_Types].xml" {
                if let Some(updated) = &updated_content_types {
                    if overrides.contains_key(&name) {
                        used_overrides.insert(name.clone());
                    }
                    writer.write_all(updated)?;
                    continue;
                }
            }
            if edited && name == "xl/_rels/workbook.bin.rels" {
                if let Some(updated) = &updated_workbook_rels {
                    if overrides.contains_key(&name) {
                        used_overrides.insert(name.clone());
                    }
                    writer.write_all(updated)?;
                    continue;
                }
            }
            if edited && name == "xl/workbook.bin" {
                if let Some(updated) = &updated_workbook_bin {
                    if overrides.contains_key(&name) {
                        used_overrides.insert(name.clone());
                    }
                    writer.write_all(updated)?;
                    continue;
                }
            }

            if name == stream_part {
                used_stream_override = true;
                if let Some(bytes) = self.preserved_parts.get(&name) {
                    let mut cursor = Cursor::new(bytes);
                    stream_override(&mut cursor, &mut writer)?;
                } else {
                    stream_override(&mut entry, &mut writer)?;
                }
                continue;
            }

            if let Some(bytes) = overrides.get(&name) {
                used_overrides.insert(name.clone());
                writer.write_all(bytes)?;
            } else if let Some(bytes) = self.preserved_parts.get(&name) {
                writer.write_all(bytes)?;
            } else {
                io::copy(&mut entry, &mut writer)?;
            }
        }

        if !used_stream_override {
            return Err(ParseError::Io(io::Error::new(
                io::ErrorKind::InvalidInput,
                format!("override part not found in source package: {stream_part}"),
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

fn read_zip_entry<R: Read + Seek>(
    zip: &mut ZipArchive<R>,
    name: &str,
) -> Result<Option<Vec<u8>>, ParseError> {
    match zip.by_name(name) {
        Ok(mut entry) => {
            let mut bytes = Vec::with_capacity(entry.size() as usize);
            entry.read_to_end(&mut bytes)?;
            Ok(Some(bytes))
        }
        Err(zip::result::ZipError::FileNotFound) => Ok(None),
        Err(e) => Err(e.into()),
    }
}

fn preserve_part<R: Read + Seek>(
    zip: &mut ZipArchive<R>,
    preserved: &mut HashMap<String, Vec<u8>>,
    name: &str,
) -> Result<(), ParseError> {
    if let Some(bytes) = read_zip_entry(zip, name)? {
        preserved.insert(name.to_string(), bytes);
    }
    Ok(())
}

fn sheet_cell_record_ids(
    sheet_bin: &[u8],
    targets: &HashSet<(u32, u32)>,
) -> Result<HashMap<(u32, u32), u32>, ParseError> {
    let mut cursor = Cursor::new(sheet_bin);
    let mut in_sheet_data = false;
    let mut current_row = 0u32;
    let mut found: HashMap<(u32, u32), u32> = HashMap::new();

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
                        found.insert(coord, id);
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
