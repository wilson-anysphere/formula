use crate::parser::Error as ParseError;
use crate::parser::{
    parse_shared_strings, parse_sheet, parse_sheet_stream, parse_workbook, Cell, CellValue, SheetData,
    SheetMeta,
};
use crate::patch::{patch_sheet_bin, CellEdit};
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
    preserved_parts: HashMap<String, Vec<u8>>,
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

        // Styles are required for round-trip, but we don't parse them yet.
        if let Some(styles) = read_zip_entry(&mut zip, "xl/styles.bin")? {
            preserved_parts.insert("xl/styles.bin".to_string(), styles);
        }

        let workbook_bin = {
            let mut wb = zip.by_name("xl/workbook.bin")?;
            let mut bytes = Vec::with_capacity(wb.size() as usize);
            wb.read_to_end(&mut bytes)?;
            bytes
        };

        let (sheets, workbook_context) = parse_workbook(&mut Cursor::new(&workbook_bin), &workbook_rels)?;
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
            preserved_parts,
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

    pub fn workbook_context(&self) -> &WorkbookContext {
        &self.workbook_context
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

        parse_sheet(&mut sheet, &self.shared_strings, &self.workbook_context)
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

        parse_sheet_stream(&mut sheet, &self.shared_strings, &self.workbook_context, |cell| f(cell))?;
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
    /// This is a convenience wrapper around the streaming worksheet patcher
    /// ([`patch_sheet_bin`]) plus the part override writer
    /// ([`XlsbWorkbook::save_with_part_overrides`]).
    ///
    /// Note: this only supports updating an existing cell; it does not insert rows/columns.
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
            }],
        )
    }

    /// Save the workbook with a set of edits for a single worksheet.
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

        // Compute updated plumbing parts if we need to invalidate calcChain.
        let mut updated_content_types: Option<Vec<u8>> = None;
        let mut updated_workbook_rels: Option<Vec<u8>> = None;

        if edited {
            let content_types =
                get_part_bytes(&mut zip, &self.preserved_parts, overrides, "[Content_Types].xml")?;
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
            if edited && name.trim_start_matches('/').eq_ignore_ascii_case("xl/calcChain.bin") {
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

            if let Some(bytes) = overrides.get(&name) {
                used_overrides.insert(name.clone());
                writer.write_all(bytes)?;
            } else if let Some(bytes) = self.preserved_parts.get(&name) {
                writer.write_all(bytes)?;
            } else {
                io::copy(&mut entry, &mut writer)?;
            }
        }

        if used_overrides.len() != overrides.len() {
            let mut missing = Vec::new();
            for key in overrides.keys() {
                if !used_overrides.contains(key) {
                    missing.push(key.clone());
                }
            }
            missing.sort();
            return Err(ParseError::Io(io::Error::new(
                io::ErrorKind::InvalidInput,
                format!("override parts not found in source package: {}", missing.join(", ")),
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

    if entry.size() as usize != expected.len() {
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
