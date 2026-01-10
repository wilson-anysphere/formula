use crate::parser::{parse_shared_strings, parse_sheet, parse_workbook_sheets, Cell, SheetData, SheetMeta};
use crate::parser::Error as ParseError;
use quick_xml::events::Event;
use quick_xml::Reader as XmlReader;
use std::collections::{HashMap, HashSet};
use std::fs::File;
use std::io::{Cursor, Read, Seek};
use std::path::{Path, PathBuf};
use zip::ZipArchive;

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
    styles: Option<Vec<u8>>,
}

impl XlsbWorkbook {
    pub fn open(path: impl AsRef<Path>) -> Result<Self, ParseError> {
        Self::open_with_options(path, OpenOptions::default())
    }

    pub fn open_with_options(path: impl AsRef<Path>, options: OpenOptions) -> Result<Self, ParseError> {
        let path = path.as_ref().to_path_buf();
        let file = File::open(&path)?;
        let mut zip = ZipArchive::new(file)?;

        let workbook_rels = read_relationships(&mut zip, "xl/_rels/workbook.bin.rels")?;

        let mut preserved_parts = HashMap::new();

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

        let styles = match zip.by_name("xl/styles.bin") {
            Ok(mut styles) => {
                let mut bytes = Vec::with_capacity(styles.size() as usize);
                styles.read_to_end(&mut bytes)?;
                Some(bytes)
            }
            Err(zip::result::ZipError::FileNotFound) => None,
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
                let is_known = known_parts.contains(name.as_str()) || worksheet_paths.contains(&name);
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
            styles,
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
        self.styles.as_deref()
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
        let data = self.read_sheet(sheet_index)?;
        for cell in data.cells {
            f(cell);
        }
        Ok(())
    }
}

fn read_relationships<R: Read + Seek>(
    zip: &mut ZipArchive<R>,
    part: &str,
) -> Result<HashMap<String, String>, ParseError> {
    let mut rels = zip.by_name(part)?;
    let mut xml = String::new();
    rels.read_to_string(&mut xml)?;

    let mut reader = XmlReader::from_str(&xml);
    reader.trim_text(true);
    let mut buf = Vec::new();
    let mut out = HashMap::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) if e.name().as_ref().ends_with(b"Relationship") => {
                let mut id = None;
                let mut target = None;
                for attr in e.attributes().flatten() {
                    match attr.key.as_ref() {
                        b"Id" => id = Some(attr.decode_and_unescape_value(&reader)?.into_owned()),
                        b"Target" => target = Some(attr.decode_and_unescape_value(&reader)?.into_owned()),
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
