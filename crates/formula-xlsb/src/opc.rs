use crate::parser::Error as ParseError;
use crate::parser::{
    parse_shared_strings, parse_sheet, parse_sheet_stream, parse_workbook_sheets, Cell, SheetData,
    SheetMeta,
};
use quick_xml::events::Event;
use quick_xml::Reader as XmlReader;
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

        let sheets = parse_workbook_sheets(&mut Cursor::new(&workbook_bin), &workbook_rels)?;
        if options.preserve_parsed_parts {
            preserved_parts.insert("xl/workbook.bin".to_string(), workbook_bin);
        }

        let shared_strings = match zip.by_name("xl/sharedStrings.bin") {
            Ok(mut sst) => {
                let mut bytes = Vec::with_capacity(sst.size() as usize);
                sst.read_to_end(&mut bytes)?;
                let strings = parse_shared_strings(&mut Cursor::new(&bytes))?;
                if options.preserve_parsed_parts {
                    preserved_parts.insert("xl/sharedStrings.bin".to_string(), bytes);
                }
                strings
            }
            Err(zip::result::ZipError::FileNotFound) => Vec::new(),
            Err(e) => return Err(e.into()),
        };

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
            preserved_parts,
        })
    }

    pub fn sheet_metas(&self) -> &[SheetMeta] {
        &self.sheets
    }

    pub fn shared_strings(&self) -> &[String] {
        &self.shared_strings
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

        parse_sheet(&mut sheet, &self.shared_strings)
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

        parse_sheet_stream(&mut sheet, &self.shared_strings, |cell| f(cell))?;
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

    /// Save the workbook while overriding specific part payloads.
    ///
    /// `overrides` maps ZIP entry paths (e.g. `xl/worksheets/sheet1.bin`) to replacement bytes.
    /// All other parts are copied from the source workbook, except for any entry already present
    /// in [`XlsbWorkbook::preserved_parts`], which is emitted from that buffer.
    pub fn save_with_part_overrides(
        &self,
        dest: impl AsRef<Path>,
        overrides: &HashMap<String, Vec<u8>>,
    ) -> Result<(), ParseError> {
        let dest = dest.as_ref();

        let file = File::open(&self.path)?;
        let mut zip = ZipArchive::new(file)?;

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

            writer.start_file(name.as_str(), options)?;
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
