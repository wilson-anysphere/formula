use super::{
    format_print_area_defined_name, format_print_titles_defined_name,
    parse_print_area_defined_name, parse_print_titles_defined_name, ManualPageBreaks, Orientation,
    PageMargins, PageSetup, PaperSize, PrintError, Scaling, SheetPrintSettings,
    WorkbookPrintSettings,
};
use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use std::collections::{HashMap, HashSet};
use std::io::{Cursor, Read, Seek, SeekFrom, Write};
use zip::write::FileOptions;
use zip::{ZipArchive, ZipWriter};

use formula_model::sheet_name_casefold;

use crate::zip_util::open_zip_part;
use crate::zip_util::read_zip_file_bytes_with_limit;

/// Maximum uncompressed size permitted for any ZIP part inflated by the print-settings helpers.
///
/// The print-settings code only needs a handful of XML parts (`workbook.xml`, `workbook.xml.rels`,
/// and worksheet XML). These should remain small in legitimate workbooks; a very large part is
/// more likely to be a zip bomb (tiny compressed size, huge uncompressed size) that could OOM the
/// desktop backend while parsing print settings.
const MAX_PRINT_ZIP_PART_BYTES: u64 = 256 * 1024 * 1024; // 256MiB

#[derive(Debug, Clone)]
pub(crate) enum DefinedNameEdit {
    Set(String),
    Remove,
}

/// Workbook-level print settings that are stored as defined names in `xl/workbook.xml`.
///
/// This mirrors the XLSX representation:
/// - `print_area` corresponds to `_xlnm.Print_Area`
/// - `print_titles` corresponds to `_xlnm.Print_Titles`
///
/// These are 1-based ranges (A1 references) as they appear in the XLSX file.
#[derive(Debug, Clone)]
pub(crate) struct SheetDefinedPrintNames {
    pub(crate) sheet_name: String,
    pub(crate) r_id: String,
    pub(crate) print_area: Option<Vec<crate::print::CellRange>>,
    pub(crate) print_titles: Option<crate::print::PrintTitles>,
}

/// Parse worksheet names and the `_xlnm.Print_Area` / `_xlnm.Print_Titles` defined names from the
/// workbook XML part (`xl/workbook.xml`).
///
/// This is a part-based helper intended for higher-level readers that already have `workbook.xml`
/// bytes, avoiding the need to re-open or re-read the ZIP package.
pub(crate) fn parse_workbook_defined_print_names(
    workbook_xml: &[u8],
) -> Result<Vec<SheetDefinedPrintNames>, PrintError> {
    let workbook = parse_workbook_xml(workbook_xml)?;

    let mut out = Vec::with_capacity(workbook.sheets.len());
    for (sheet_index, sheet) in workbook.sheets.into_iter().enumerate() {
        let sheet_name = sheet.name;
        let print_area = workbook
            .defined_names
            .iter()
            .find(|dn| {
                dn.local_sheet_id == Some(sheet_index)
                    && dn.name.eq_ignore_ascii_case("_xlnm.Print_Area")
            })
            .map(|dn| parse_print_area_defined_name(&sheet_name, &dn.value))
            .transpose()?;

        let print_titles = workbook
            .defined_names
            .iter()
            .find(|dn| {
                dn.local_sheet_id == Some(sheet_index)
                    && dn.name.eq_ignore_ascii_case("_xlnm.Print_Titles")
            })
            .map(|dn| parse_print_titles_defined_name(&sheet_name, &dn.value))
            .transpose()?;

        out.push(SheetDefinedPrintNames {
            sheet_name,
            r_id: sheet.r_id,
            print_area,
            print_titles,
        });
    }

    Ok(out)
}

/// Parse the worksheet `pageSetup`/`pageMargins`/`pageSetUpPr` settings from an already-extracted
/// worksheet XML part (`xl/worksheets/sheetN.xml`).
pub(crate) fn parse_worksheet_page_setup(sheet_xml: &[u8]) -> Result<PageSetup, PrintError> {
    Ok(parse_worksheet_print_settings(sheet_xml)?.0)
}

/// Parse worksheet `rowBreaks` / `colBreaks` manual page breaks from an already-extracted
/// worksheet XML part (`xl/worksheets/sheetN.xml`).
pub(crate) fn parse_worksheet_manual_page_breaks(
    sheet_xml: &[u8],
) -> Result<ManualPageBreaks, PrintError> {
    Ok(parse_worksheet_print_settings(sheet_xml)?.1)
}

pub fn read_workbook_print_settings(
    xlsx_bytes: &[u8],
) -> Result<WorkbookPrintSettings, PrintError> {
    read_workbook_print_settings_from_reader(Cursor::new(xlsx_bytes))
}

/// Streaming variant of [`read_workbook_print_settings`].
///
/// This allows callers to extract print settings from an on-disk XLSX/XLSM package without first
/// reading the entire ZIP container into memory.
pub fn read_workbook_print_settings_from_reader<R: Read + Seek>(
    reader: R,
) -> Result<WorkbookPrintSettings, PrintError> {
    read_workbook_print_settings_from_reader_with_limit(reader, MAX_PRINT_ZIP_PART_BYTES)
}

#[cfg(test)]
pub(crate) fn read_workbook_print_settings_with_limit(
    xlsx_bytes: &[u8],
    max_part_bytes: u64,
) -> Result<WorkbookPrintSettings, PrintError> {
    read_workbook_print_settings_from_reader_with_limit(Cursor::new(xlsx_bytes), max_part_bytes)
}

fn read_workbook_print_settings_from_reader_with_limit<R: Read + Seek>(
    mut reader: R,
    max_part_bytes: u64,
) -> Result<WorkbookPrintSettings, PrintError> {
    reader.seek(SeekFrom::Start(0))?;
    let mut zip = ZipArchive::new(reader)?;
    let workbook_xml = read_zip_bytes(&mut zip, "xl/workbook.xml", max_part_bytes)?;
    let rels_xml = read_zip_bytes(&mut zip, "xl/_rels/workbook.xml.rels", max_part_bytes)?;

    let workbook_print_names = parse_workbook_defined_print_names(&workbook_xml)?;
    let rels = parse_workbook_rels(&rels_xml)?;

    let mut sheets = Vec::with_capacity(workbook_print_names.len());
    for sheet in workbook_print_names {
        let sheet_target = rels
            .get(&sheet.r_id)
            .ok_or(PrintError::MissingPart("worksheet relationship"))?;
        // Relationship targets can be relative to the workbook part's folder (e.g.
        // `worksheets/sheet1.xml`) or absolute (e.g. `/xl/worksheets/sheet1.xml`). Use the shared
        // OpenXML resolver to handle both.
        let sheet_path = crate::openxml::resolve_target("xl/workbook.xml", sheet_target);
        let sheet_xml = read_zip_bytes(&mut zip, &sheet_path, max_part_bytes)?;

        let page_setup = parse_worksheet_page_setup(&sheet_xml)?;
        let manual_page_breaks = parse_worksheet_manual_page_breaks(&sheet_xml)?;

        sheets.push(SheetPrintSettings {
            sheet_name: sheet.sheet_name,
            print_area: sheet.print_area,
            print_titles: sheet.print_titles,
            page_setup,
            manual_page_breaks,
        });
    }

    Ok(WorkbookPrintSettings { sheets })
}

pub fn write_workbook_print_settings(
    xlsx_bytes: &[u8],
    settings: &WorkbookPrintSettings,
) -> Result<Vec<u8>, PrintError> {
    write_workbook_print_settings_impl(xlsx_bytes, settings, MAX_PRINT_ZIP_PART_BYTES)
}

fn write_workbook_print_settings_impl(
    xlsx_bytes: &[u8],
    settings: &WorkbookPrintSettings,
    max_part_bytes: u64,
) -> Result<Vec<u8>, PrintError> {
    let mut zip = ZipArchive::new(Cursor::new(xlsx_bytes))?;
    let (workbook_part_name, workbook_xml) =
        read_zip_bytes_with_entry_name(&mut zip, "xl/workbook.xml", max_part_bytes)?;
    let rels_xml = read_zip_bytes(&mut zip, "xl/_rels/workbook.xml.rels", max_part_bytes)?;

    let workbook = parse_workbook_xml(&workbook_xml)?;
    let rels = parse_workbook_rels(&rels_xml)?;

    // Excel sheet names are case-insensitive across Unicode; accept settings keyed by any casing.
    let mut settings_by_sheet: HashMap<String, &SheetPrintSettings> = HashMap::new();
    for sheet in &settings.sheets {
        if sheet.sheet_name.is_empty() {
            continue;
        }
        settings_by_sheet.insert(sheet_name_casefold(&sheet.sheet_name), sheet);
    }

    let mut defined_name_edits: HashMap<(String, usize), DefinedNameEdit> = HashMap::new();
    for (sheet_index, sheet) in workbook.sheets.iter().enumerate() {
        let sheet_key = sheet_name_casefold(&sheet.name);
        if let Some(sheet_settings) = settings_by_sheet.get(&sheet_key) {
            match sheet_settings.print_area {
                Some(ref ranges) => {
                    defined_name_edits.insert(
                        ("_xlnm.Print_Area".to_string(), sheet_index),
                        DefinedNameEdit::Set(format_print_area_defined_name(&sheet.name, ranges)),
                    );
                }
                None => {
                    defined_name_edits.insert(
                        ("_xlnm.Print_Area".to_string(), sheet_index),
                        DefinedNameEdit::Remove,
                    );
                }
            }

            match sheet_settings.print_titles {
                Some(ref titles) => {
                    defined_name_edits.insert(
                        ("_xlnm.Print_Titles".to_string(), sheet_index),
                        DefinedNameEdit::Set(format_print_titles_defined_name(&sheet.name, titles)),
                    );
                }
                None => {
                    defined_name_edits.insert(
                        ("_xlnm.Print_Titles".to_string(), sheet_index),
                        DefinedNameEdit::Remove,
                    );
                }
            }
        }
    }

    let updated_workbook_xml = update_workbook_xml(&workbook_xml, &defined_name_edits)?;

    let mut updated_sheets: HashMap<String, Vec<u8>> = HashMap::new();
    for sheet in &workbook.sheets {
        let sheet_key = sheet_name_casefold(&sheet.name);
        let Some(sheet_settings) = settings_by_sheet.get(&sheet_key) else {
            continue;
        };
        let Some(sheet_target) = rels.get(&sheet.r_id) else {
            continue;
        };
        let sheet_path = crate::openxml::resolve_target("xl/workbook.xml", sheet_target);
        let (sheet_part_name, sheet_xml) =
            read_zip_bytes_with_entry_name(&mut zip, &sheet_path, max_part_bytes)?;
        let updated_xml = update_worksheet_xml(&sheet_xml, sheet_settings)?;
        updated_sheets.insert(sheet_part_name, updated_xml);
    }

    // Rewind zip to iterate entries again.
    let mut zip = ZipArchive::new(Cursor::new(xlsx_bytes))?;
    let mut out = ZipWriter::new(Cursor::new(Vec::new()));
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    for i in 0..zip.len() {
        let entry = zip.by_index(i)?;
        let name = entry.name().to_string();
        let canonical_name = name.strip_prefix('/').unwrap_or(name.as_str());
        if entry.is_dir() {
            out.add_directory(name, options.clone())?;
            continue;
        }

        let replacement = if canonical_name == workbook_part_name.as_str() {
            Some(updated_workbook_xml.as_slice())
        } else {
            updated_sheets.get(canonical_name).map(|v| v.as_slice())
        };

        if let Some(bytes) = replacement {
            out.start_file(name, options.clone())?;
            out.write_all(bytes)?;
        } else {
            // Preserve unchanged parts byte-for-byte (avoid decompression/recompression of large
            // binary assets like images).
            out.raw_copy_file(entry)?;
        }
    }

    Ok(out.finish()?.into_inner())
}

fn read_zip_bytes<R: Read + Seek>(
    zip: &mut ZipArchive<R>,
    name: &str,
    max_part_bytes: u64,
) -> Result<Vec<u8>, PrintError> {
    let mut file = open_zip_part(zip, name)?;
    read_zip_file_bytes_with_limit(&mut file, name, max_part_bytes).map_err(|err| match err {
        crate::XlsxError::PartTooLarge { part, size, max } => PrintError::PartTooLarge { part, size, max },
        crate::XlsxError::Io(err) => PrintError::Io(err),
        crate::XlsxError::Zip(err) => PrintError::Zip(err),
        other => PrintError::Io(std::io::Error::new(
            std::io::ErrorKind::InvalidData,
            other.to_string(),
        )),
    })
}

fn read_zip_bytes_with_entry_name<R: Read + Seek>(
    zip: &mut ZipArchive<R>,
    name: &str,
    max_part_bytes: u64,
) -> Result<(String, Vec<u8>), PrintError> {
    let mut file = open_zip_part(zip, name)?;
    let raw_name = file.name().to_string();
    let canonical_name = raw_name.strip_prefix('/').unwrap_or(raw_name.as_str()).to_string();
    let bytes = read_zip_file_bytes_with_limit(&mut file, &canonical_name, max_part_bytes)
        .map_err(|err| match err {
            crate::XlsxError::PartTooLarge { part, size, max } => {
                PrintError::PartTooLarge { part, size, max }
            }
            crate::XlsxError::Io(err) => PrintError::Io(err),
            crate::XlsxError::Zip(err) => PrintError::Zip(err),
            other => PrintError::Io(std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                other.to_string(),
            )),
        })?;
    Ok((canonical_name, bytes))
}

#[derive(Debug)]
struct WorkbookInfo {
    sheets: Vec<SheetInfo>,
    defined_names: Vec<DefinedName>,
}

#[derive(Debug)]
struct SheetInfo {
    name: String,
    r_id: String,
}

#[derive(Debug)]
struct DefinedName {
    name: String,
    local_sheet_id: Option<usize>,
    value: String,
}

fn parse_workbook_xml(workbook_xml: &[u8]) -> Result<WorkbookInfo, PrintError> {
    let mut reader = Reader::from_reader(workbook_xml);

    let mut buf = Vec::new();
    let mut sheets = Vec::new();
    let mut defined_names = Vec::new();

    let mut current_defined: Option<DefinedName> = None;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if e.local_name().as_ref() == b"sheet" => {
                sheets.push(parse_sheet_info(&reader, &e)?);
            }
            Event::Empty(e) if e.local_name().as_ref() == b"sheet" => {
                sheets.push(parse_sheet_info(&reader, &e)?);
            }
            Event::Start(e) if e.local_name().as_ref() == b"definedName" => {
                current_defined = Some(parse_defined_name_start(&reader, &e)?);
            }
            Event::Text(e) if current_defined.is_some() => {
                if let Some(ref mut dn) = current_defined {
                    dn.value.push_str(&e.unescape()?.to_string());
                }
            }
            Event::CData(e) if current_defined.is_some() => {
                if let Some(ref mut dn) = current_defined {
                    dn.value.push_str(std::str::from_utf8(e.as_ref())?);
                }
            }
            Event::End(e) if e.local_name().as_ref() == b"definedName" => {
                if let Some(dn) = current_defined.take() {
                    defined_names.push(dn);
                }
            }
            Event::Eof => break,
            _ => {}
        }

        buf.clear();
    }

    Ok(WorkbookInfo {
        sheets,
        defined_names,
    })
}

fn parse_sheet_info(reader: &Reader<&[u8]>, e: &BytesStart<'_>) -> Result<SheetInfo, PrintError> {
    let mut name: Option<String> = None;
    let mut r_id: Option<String> = None;
    for attr in e.attributes().with_checks(false) {
        let attr = attr?;
        // Ignore namespace declarations (`xmlns` / `xmlns:*`). Some producers may use prefixes like
        // `xmlns:id="..."` which would otherwise look like an `id` attribute when matching by
        // local-name.
        let key = attr.key.as_ref();
        if key.starts_with(b"xmlns") {
            continue;
        }

        match key {
            b"name" => name = Some(attr.unescape_value()?.to_string()),
            key if crate::openxml::local_name(key) == b"id" => {
                r_id = Some(attr.unescape_value()?.to_string())
            }
            _ => {}
        }
    }

    let name = name.ok_or_else(|| PrintError::InvalidA1("sheet missing name".to_string()))?;
    let r_id = r_id.ok_or_else(|| PrintError::InvalidA1("sheet missing r:id".to_string()))?;
    let _ = reader;
    Ok(SheetInfo { name, r_id })
}

fn parse_defined_name_start(
    _reader: &Reader<&[u8]>,
    e: &BytesStart<'_>,
) -> Result<DefinedName, PrintError> {
    let mut name: Option<String> = None;
    let mut local_sheet_id: Option<usize> = None;
    for attr in e.attributes().with_checks(false) {
        let attr = attr?;
        match attr.key.as_ref() {
            b"name" => name = Some(attr.unescape_value()?.to_string()),
            b"localSheetId" => {
                local_sheet_id = Some(
                    attr.unescape_value()?
                        .parse::<usize>()
                        .map_err(|_| PrintError::InvalidA1("invalid localSheetId".to_string()))?,
                )
            }
            _ => {}
        }
    }

    Ok(DefinedName {
        name: name.ok_or_else(|| PrintError::InvalidA1("definedName missing name".to_string()))?,
        local_sheet_id,
        value: String::new(),
    })
}

fn parse_workbook_rels(rels_xml: &[u8]) -> Result<HashMap<String, String>, PrintError> {
    let mut reader = Reader::from_reader(rels_xml);
    let mut buf = Vec::new();

    let mut map = HashMap::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e)
                if e.local_name()
                    .as_ref()
                    .eq_ignore_ascii_case(b"Relationship") =>
            {
                if let Some((id, target)) = parse_relationship(&e)? {
                    map.insert(id, target);
                }
            }
            Event::Empty(e)
                if e.local_name()
                    .as_ref()
                    .eq_ignore_ascii_case(b"Relationship") =>
            {
                if let Some((id, target)) = parse_relationship(&e)? {
                    map.insert(id, target);
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(map)
}

fn parse_relationship(e: &BytesStart<'_>) -> Result<Option<(String, String)>, PrintError> {
    let mut id: Option<String> = None;
    let mut target: Option<String> = None;
    for attr in e.attributes().with_checks(false) {
        let attr = attr?;
        let key_raw = attr.key.as_ref();
        if key_raw.starts_with(b"xmlns") {
            continue;
        }
        let key = crate::openxml::local_name(key_raw);
        if key.eq_ignore_ascii_case(b"Id") {
            id = Some(attr.unescape_value()?.to_string());
        } else if key.eq_ignore_ascii_case(b"Target") {
            target = Some(attr.unescape_value()?.to_string());
        }
    }

    match (id, target) {
        (Some(id), Some(target)) => Ok(Some((id, target))),
        _ => Ok(None),
    }
}

pub(crate) fn parse_worksheet_print_settings(
    sheet_xml: &[u8],
) -> Result<(PageSetup, ManualPageBreaks), PrintError> {
    let mut reader = Reader::from_reader(sheet_xml);
    let mut buf = Vec::new();

    let mut margins = None;
    let mut orientation: Option<Orientation> = None;
    let mut paper_size: Option<PaperSize> = None;
    let mut scale: Option<u16> = None;
    let mut fit_to_width: Option<u16> = None;
    let mut fit_to_height: Option<u16> = None;
    // In OOXML, `pageSetup` stores both percent scaling (`scale`) and fit-to-page dimensions
    // (`fitToWidth`/`fitToHeight`). The `sheetPr/pageSetUpPr/@fitToPage` flag determines which
    // mode is active.
    //
    // Some non-Excel producers omit `pageSetUpPr` while still providing `fitToWidth`/`fitToHeight`;
    // keep this optional so we can apply a best-effort inference only when the flag is absent.
    let mut fit_to_page: Option<bool> = None;

    let mut manual_breaks = ManualPageBreaks::default();
    let mut in_row_breaks = false;
    let mut in_col_breaks = false;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"pageMargins" => {
                margins = Some(parse_page_margins(&e)?);
            }
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"pageSetup" => {
                let (o, p, s, ftw, fth) = parse_page_setup(&e)?;
                orientation = o.or(orientation);
                paper_size = p.or(paper_size);
                scale = s.or(scale);
                fit_to_width = ftw.or(fit_to_width);
                fit_to_height = fth.or(fit_to_height);
            }
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"pageSetUpPr" => {
                fit_to_page = Some(parse_fit_to_page(&e)?);
            }
            Event::Start(e) if e.local_name().as_ref() == b"rowBreaks" => in_row_breaks = true,
            Event::End(e) if e.local_name().as_ref() == b"rowBreaks" => in_row_breaks = false,
            Event::Start(e) if e.local_name().as_ref() == b"colBreaks" => in_col_breaks = true,
            Event::End(e) if e.local_name().as_ref() == b"colBreaks" => in_col_breaks = false,
            Event::Start(e) | Event::Empty(e)
                if e.local_name().as_ref() == b"brk" && in_row_breaks =>
            {
                if let Some(id) = parse_break_id(&e)? {
                    manual_breaks.row_breaks_after.insert(id);
                }
            }
            Event::Start(e) | Event::Empty(e)
                if e.local_name().as_ref() == b"brk" && in_col_breaks =>
            {
                if let Some(id) = parse_break_id(&e)? {
                    manual_breaks.col_breaks_after.insert(id);
                }
            }
            Event::Eof => break,
            _ => {}
        }

        buf.clear();
    }

    fn normalize_scale(scale: u16) -> u16 {
        if scale == 0 { 100 } else { scale }
    }

    let scaling = match fit_to_page {
        Some(true) => Scaling::FitTo {
            width: fit_to_width.unwrap_or(0),
            height: fit_to_height.unwrap_or(0),
        },
        Some(false) => Scaling::Percent(normalize_scale(scale.unwrap_or(100))),
        None => {
            // Best-effort: if the sheet omits `pageSetUpPr` but supplies fit dimensions, treat it as
            // fit-to-page mode.
            if fit_to_width.is_some() || fit_to_height.is_some() {
                Scaling::FitTo {
                    width: fit_to_width.unwrap_or(0),
                    height: fit_to_height.unwrap_or(0),
                }
            } else if let Some(scale) = scale {
                Scaling::Percent(normalize_scale(scale))
            } else {
                Scaling::Percent(100)
            }
        }
    };

    Ok((
        PageSetup {
            orientation: orientation.unwrap_or_default(),
            paper_size: paper_size.unwrap_or_default(),
            margins: margins.unwrap_or_default(),
            scaling,
        },
        manual_breaks,
    ))
}

fn parse_page_margins(e: &BytesStart<'_>) -> Result<PageMargins, PrintError> {
    let mut margins = PageMargins::default();
    for attr in e.attributes().with_checks(false) {
        let attr = attr?;
        let value = attr.unescape_value()?;
        match attr.key.as_ref() {
            b"left" => margins.left = value.parse::<f64>().unwrap_or(margins.left),
            b"right" => margins.right = value.parse::<f64>().unwrap_or(margins.right),
            b"top" => margins.top = value.parse::<f64>().unwrap_or(margins.top),
            b"bottom" => margins.bottom = value.parse::<f64>().unwrap_or(margins.bottom),
            b"header" => margins.header = value.parse::<f64>().unwrap_or(margins.header),
            b"footer" => margins.footer = value.parse::<f64>().unwrap_or(margins.footer),
            _ => {}
        }
    }
    Ok(margins)
}

fn parse_page_setup(
    e: &BytesStart<'_>,
) -> Result<
    (
        Option<Orientation>,
        Option<PaperSize>,
        Option<u16>,
        Option<u16>,
        Option<u16>,
    ),
    PrintError,
> {
    let mut orientation = None;
    let mut paper_size = None;
    let mut scale = None;
    let mut fit_to_width = None;
    let mut fit_to_height = None;

    for attr in e.attributes().with_checks(false) {
        let attr = attr?;
        let value = attr.unescape_value()?;
        match attr.key.as_ref() {
            b"orientation" => {
                orientation = match value.as_ref() {
                    "landscape" => Some(Orientation::Landscape),
                    "portrait" => Some(Orientation::Portrait),
                    _ => None,
                }
            }
            b"paperSize" => {
                if let Ok(code) = value.parse::<u16>() {
                    paper_size = Some(PaperSize { code });
                }
            }
            b"scale" => {
                if let Ok(v) = value.parse::<u16>() {
                    scale = Some(v);
                }
            }
            b"fitToWidth" => {
                if let Ok(v) = value.parse::<u16>() {
                    fit_to_width = Some(v);
                }
            }
            b"fitToHeight" => {
                if let Ok(v) = value.parse::<u16>() {
                    fit_to_height = Some(v);
                }
            }
            _ => {}
        }
    }

    Ok((orientation, paper_size, scale, fit_to_width, fit_to_height))
}

fn parse_fit_to_page(e: &BytesStart<'_>) -> Result<bool, PrintError> {
    for attr in e.attributes().with_checks(false) {
        let attr = attr?;
        if attr.key.as_ref() == b"fitToPage" {
            let value = attr.unescape_value()?;
            return Ok(value.as_ref() == "1" || value.as_ref().eq_ignore_ascii_case("true"));
        }
    }
    Ok(false)
}

fn parse_break_id(e: &BytesStart<'_>) -> Result<Option<u32>, PrintError> {
    let mut id: Option<u32> = None;
    let mut man: Option<bool> = None;
    for attr in e.attributes().with_checks(false) {
        let attr = attr?;
        let value = attr.unescape_value()?;
        match attr.key.as_ref() {
            b"id" => id = value.parse::<u32>().ok(),
            b"man" => {
                man = Some(value.as_ref() == "1" || value.as_ref().eq_ignore_ascii_case("true"))
            }
            _ => {}
        }
    }

    if let Some(false) = man {
        return Ok(None);
    }

    Ok(id)
}

pub(crate) fn update_workbook_xml(
    workbook_xml: &[u8],
    edits: &HashMap<(String, usize), DefinedNameEdit>,
) -> Result<Vec<u8>, PrintError> {
    let mut reader = Reader::from_reader(workbook_xml);
    let mut writer = Writer::new(Vec::new());
    let mut buf = Vec::new();

    let mut workbook_prefix: Option<String> = None;
    let mut in_defined_names = false;
    let mut seen_defined_names = false;
    let mut skipping_defined_name = false;
    let mut current_defined_key: Option<(String, usize)> = None;
    let mut applied: HashSet<(String, usize)> = HashSet::new();
    let needs_defined_names = edits.values().any(|e| matches!(e, DefinedNameEdit::Set(_)));

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Start(ref e) if e.local_name().as_ref() == b"workbook" => {
                let ns = crate::xml::workbook_xml_namespaces_from_workbook_start(e)?;
                workbook_prefix = ns.spreadsheetml_prefix;
                writer.write_event(event)?;
            }
            Event::Empty(ref e) if e.local_name().as_ref() == b"workbook" => {
                let ns = crate::xml::workbook_xml_namespaces_from_workbook_start(e)?;
                workbook_prefix = ns.spreadsheetml_prefix;

                // Some workbooks are degenerate/self-closing (e.g. `<workbook .../>`). When we need
                // to insert defined names, expand the root to a proper start/end pair so we can
                // emit children.
                if needs_defined_names {
                    let workbook_tag =
                        String::from_utf8_lossy(e.name().as_ref()).into_owned();
                    writer.write_event(Event::Start(e.to_owned()))?;

                    let defined_names_tag =
                        crate::xml::prefixed_tag(workbook_prefix.as_deref(), "definedNames");
                    let defined_name_tag =
                        crate::xml::prefixed_tag(workbook_prefix.as_deref(), "definedName");

                    writer.write_event(Event::Start(BytesStart::new(defined_names_tag.as_str())))?;
                    for ((name, local_sheet_id), edit) in edits {
                        if let DefinedNameEdit::Set(value) = edit {
                            write_defined_name(
                                &mut writer,
                                defined_name_tag.as_str(),
                                name,
                                *local_sheet_id,
                                value,
                            )?;
                        }
                    }
                    writer.write_event(Event::End(BytesEnd::new(defined_names_tag.as_str())))?;

                    writer.write_event(Event::End(BytesEnd::new(workbook_tag.as_str())))?;
                    // The workbook was self-closing, so nothing else can appear inside it.
                    seen_defined_names = true;
                } else {
                    writer.write_event(event)?;
                }
            }

            Event::Empty(ref e) if e.local_name().as_ref() == b"definedNames" => {
                // Some producers (including Excel) serialize an empty `<definedNames/>` element
                // instead of omitting it entirely. If we need to insert new defined names, expand
                // the empty element rather than inserting a second `<definedNames>` block.
                seen_defined_names = true;

                if edits.values().any(|e| matches!(e, DefinedNameEdit::Set(_))) {
                    let tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                    writer.write_event(Event::Start(e.to_owned()))?;
                    for ((name, local_sheet_id), edit) in edits {
                        if let DefinedNameEdit::Set(value) = edit {
                            let defined_name_tag =
                                crate::xml::prefixed_tag(workbook_prefix.as_deref(), "definedName");
                            write_defined_name(
                                &mut writer,
                                defined_name_tag.as_str(),
                                name,
                                *local_sheet_id,
                                value,
                            )?;
                        }
                    }
                    writer.write_event(Event::End(BytesEnd::new(tag.as_str())))?;
                } else {
                    writer.write_event(event)?;
                }
            }
            Event::Start(ref e) if e.local_name().as_ref() == b"definedNames" => {
                in_defined_names = true;
                seen_defined_names = true;
                writer.write_event(event)?;
            }
            Event::End(ref e) if e.local_name().as_ref() == b"definedNames" => {
                if in_defined_names {
                    for ((name, local_sheet_id), edit) in edits {
                        if applied.contains(&(name.clone(), *local_sheet_id)) {
                            continue;
                        }
                        if let DefinedNameEdit::Set(value) = edit {
                            let tag =
                                crate::xml::prefixed_tag(workbook_prefix.as_deref(), "definedName");
                            write_defined_name(
                                &mut writer,
                                tag.as_str(),
                                name,
                                *local_sheet_id,
                                value,
                            )?;
                        }
                    }
                }
                in_defined_names = false;
                writer.write_event(event)?;
            }
            Event::Start(ref e)
                if in_defined_names && e.local_name().as_ref() == b"definedName" =>
            {
                let (name, local_sheet_id) = parse_defined_name_key(e)?;
                if let (Some(name), Some(local_sheet_id)) = (name, local_sheet_id) {
                    let key = (name.clone(), local_sheet_id);
                    if let Some(edit) = edits.get(&key) {
                        applied.insert(key.clone());
                        current_defined_key = Some(key);
                        match edit {
                            DefinedNameEdit::Set(value) => {
                                writer.write_event(Event::Start(e.to_owned()))?;
                                writer.write_event(Event::Text(BytesText::new(value)))?;
                                skipping_defined_name = true;
                                buf.clear();
                                continue;
                            }
                            DefinedNameEdit::Remove => {
                                skipping_defined_name = true;
                                buf.clear();
                                continue;
                            }
                        }
                    }
                }
                writer.write_event(event)?;
            }
            Event::Empty(ref e)
                if in_defined_names && e.local_name().as_ref() == b"definedName" =>
            {
                let (name, local_sheet_id) = parse_defined_name_key(e)?;
                if let (Some(name), Some(local_sheet_id)) = (name, local_sheet_id) {
                    let key = (name.clone(), local_sheet_id);
                    if let Some(edit) = edits.get(&key) {
                        applied.insert(key);
                        match edit {
                            DefinedNameEdit::Set(value) => {
                                // Expand the self-closing tag to allow inserting text content.
                                let tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                                writer.write_event(Event::Start(e.to_owned()))?;
                                writer.write_event(Event::Text(BytesText::new(value)))?;
                                writer.write_event(Event::End(BytesEnd::new(tag.as_str())))?;
                            }
                            DefinedNameEdit::Remove => {
                                // Skip.
                            }
                        }
                        buf.clear();
                        continue;
                    }
                }
                writer.write_event(event)?;
            }
            Event::End(ref e)
                if skipping_defined_name && e.local_name().as_ref() == b"definedName" =>
            {
                if let Some(key) = current_defined_key.take() {
                    if matches!(edits.get(&key), Some(DefinedNameEdit::Set(_))) {
                        writer.write_event(Event::End(e.to_owned()))?;
                    }
                }
                skipping_defined_name = false;
            }
            Event::End(ref e)
                if e.local_name().as_ref() == b"workbook"
                    && !seen_defined_names
                    && needs_defined_names =>
            {
                let defined_names_tag =
                    crate::xml::prefixed_tag(workbook_prefix.as_deref(), "definedNames");
                let defined_name_tag =
                    crate::xml::prefixed_tag(workbook_prefix.as_deref(), "definedName");

                writer.write_event(Event::Start(BytesStart::new(defined_names_tag.as_str())))?;
                for ((name, local_sheet_id), edit) in edits {
                    if let DefinedNameEdit::Set(value) = edit {
                        write_defined_name(
                            &mut writer,
                            defined_name_tag.as_str(),
                            name,
                            *local_sheet_id,
                            value,
                        )?;
                    }
                }
                writer.write_event(Event::End(BytesEnd::new(defined_names_tag.as_str())))?;
                writer.write_event(event)?;
            }
            Event::Eof => break,
            _ => {
                if !skipping_defined_name {
                    writer.write_event(event)?;
                }
            }
        }
        buf.clear();
    }

    Ok(writer.into_inner())
}

fn parse_defined_name_key(
    e: &BytesStart<'_>,
) -> Result<(Option<String>, Option<usize>), PrintError> {
    let mut name = None;
    let mut local_sheet_id = None;
    for attr in e.attributes().with_checks(false) {
        let attr = attr?;
        match attr.key.as_ref() {
            b"name" => name = Some(attr.unescape_value()?.to_string()),
            b"localSheetId" => {
                local_sheet_id = attr.unescape_value()?.parse::<usize>().ok();
            }
            _ => {}
        }
    }
    Ok((name, local_sheet_id))
}

fn write_defined_name(
    writer: &mut Writer<Vec<u8>>,
    tag: &str,
    name: &str,
    local_sheet_id: usize,
    value: &str,
) -> Result<(), PrintError> {
    let local_sheet_id_str = local_sheet_id.to_string();
    let mut start = BytesStart::new(tag).into_owned();
    start.push_attribute(("name", name));
    start.push_attribute(("localSheetId", local_sheet_id_str.as_str()));
    writer.write_event(Event::Start(start))?;
    writer.write_event(Event::Text(BytesText::new(value)))?;
    writer.write_event(Event::End(BytesEnd::new(tag)))?;
    Ok(())
}

pub(crate) fn update_worksheet_xml(
    sheet_xml: &[u8],
    settings: &SheetPrintSettings,
) -> Result<Vec<u8>, PrintError> {
    let sheet_xml_str = std::str::from_utf8(sheet_xml)?;
    let worksheet_prefix = crate::xml::worksheet_spreadsheetml_prefix(sheet_xml_str)?;

    let mut reader = Reader::from_reader(sheet_xml);
    let mut writer = Writer::new(Vec::new());
    let mut buf = Vec::new();

    let mut seen_sheet_pr = false;
    let mut in_sheet_pr = false;
    let mut sheet_pr_prefix: Option<String> = None;
    let mut seen_page_setup_pr = false;

    let mut seen_page_margins = false;
    let mut seen_page_setup = false;
    let mut seen_row_breaks = false;
    let mut seen_col_breaks = false;

    let mut skip_tag: Option<&'static [u8]> = None;
    let mut skip_depth = 0usize;

    loop {
        let event = reader.read_event_into(&mut buf)?;

        if let Some(tag) = skip_tag {
            match event {
                Event::Start(_) => skip_depth += 1,
                Event::End(ref e) => {
                    if skip_depth == 0 && e.local_name().as_ref() == tag {
                        skip_tag = None;
                    } else if skip_depth > 0 {
                        skip_depth -= 1;
                    }
                }
                Event::Eof => break,
                _ => {}
            }

            buf.clear();
            continue;
        }

        match event {
            Event::Start(ref e) if e.local_name().as_ref() == b"sheetPr" => {
                seen_sheet_pr = true;
                in_sheet_pr = true;
                let name = e.name();
                let name = name.as_ref();
                sheet_pr_prefix = name
                    .iter()
                    .rposition(|b| *b == b':')
                    .map(|idx| &name[..idx])
                    .and_then(|p| std::str::from_utf8(p).ok())
                    .map(|s| s.to_string());
                writer.write_event(event)?;
            }
            Event::End(ref e) if e.local_name().as_ref() == b"sheetPr" => {
                if settings.page_setup.scaling.is_fit_to() && !seen_page_setup_pr {
                    write_page_setup_pr(&mut writer, sheet_pr_prefix.as_deref(), true)?;
                }
                in_sheet_pr = false;
                sheet_pr_prefix = None;
                writer.write_event(event)?;
            }
            Event::Start(ref e) | Event::Empty(ref e)
                if in_sheet_pr && e.local_name().as_ref() == b"pageSetUpPr" =>
            {
                seen_page_setup_pr = true;
                let fit_to_page = settings.page_setup.scaling.is_fit_to();
                writer.write_event(update_page_setup_pr_event(&event, fit_to_page)?)?;
            }
            Event::Start(ref e) if e.local_name().as_ref() == b"pageMargins" => {
                seen_page_margins = true;
                let tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                writer.write_event(build_page_margins_event(
                    tag.as_str(),
                    &settings.page_setup.margins,
                )?)?;
                skip_tag = Some(b"pageMargins");
                skip_depth = 0;
            }
            Event::Empty(ref e) if e.local_name().as_ref() == b"pageMargins" => {
                seen_page_margins = true;
                let tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                writer.write_event(build_page_margins_event(
                    tag.as_str(),
                    &settings.page_setup.margins,
                )?)?;
            }
            Event::Start(ref e) if e.local_name().as_ref() == b"pageSetup" => {
                seen_page_setup = true;
                let tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                writer.write_event(build_page_setup_event(tag.as_str(), &settings.page_setup)?)?;
                skip_tag = Some(b"pageSetup");
                skip_depth = 0;
            }
            Event::Empty(ref e) if e.local_name().as_ref() == b"pageSetup" => {
                seen_page_setup = true;
                let tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                writer.write_event(build_page_setup_event(tag.as_str(), &settings.page_setup)?)?;
            }
            Event::Start(ref e) if e.local_name().as_ref() == b"rowBreaks" => {
                seen_row_breaks = true;
                if !settings.manual_page_breaks.row_breaks_after.is_empty() {
                    let name = e.name();
                    let name = name.as_ref();
                    let prefix = name
                        .iter()
                        .rposition(|b| *b == b':')
                        .map(|idx| &name[..idx])
                        .and_then(|p| std::str::from_utf8(p).ok());
                    write_row_breaks(&mut writer, prefix, &settings.manual_page_breaks)?;
                }
                skip_tag = Some(b"rowBreaks");
                skip_depth = 0;
            }
            Event::Empty(ref e) if e.local_name().as_ref() == b"rowBreaks" => {
                seen_row_breaks = true;
                if !settings.manual_page_breaks.row_breaks_after.is_empty() {
                    let name = e.name();
                    let name = name.as_ref();
                    let prefix = name
                        .iter()
                        .rposition(|b| *b == b':')
                        .map(|idx| &name[..idx])
                        .and_then(|p| std::str::from_utf8(p).ok());
                    write_row_breaks(&mut writer, prefix, &settings.manual_page_breaks)?;
                }
            }
            Event::Start(ref e) if e.local_name().as_ref() == b"colBreaks" => {
                seen_col_breaks = true;
                if !settings.manual_page_breaks.col_breaks_after.is_empty() {
                    let name = e.name();
                    let name = name.as_ref();
                    let prefix = name
                        .iter()
                        .rposition(|b| *b == b':')
                        .map(|idx| &name[..idx])
                        .and_then(|p| std::str::from_utf8(p).ok());
                    write_col_breaks(&mut writer, prefix, &settings.manual_page_breaks)?;
                }
                skip_tag = Some(b"colBreaks");
                skip_depth = 0;
            }
            Event::Empty(ref e) if e.local_name().as_ref() == b"colBreaks" => {
                seen_col_breaks = true;
                if !settings.manual_page_breaks.col_breaks_after.is_empty() {
                    let name = e.name();
                    let name = name.as_ref();
                    let prefix = name
                        .iter()
                        .rposition(|b| *b == b':')
                        .map(|idx| &name[..idx])
                        .and_then(|p| std::str::from_utf8(p).ok());
                    write_col_breaks(&mut writer, prefix, &settings.manual_page_breaks)?;
                }
            }
            Event::End(ref e) if e.local_name().as_ref() == b"worksheet" => {
                if !seen_sheet_pr && settings.page_setup.scaling.is_fit_to() {
                    let sheet_pr_tag =
                        crate::xml::prefixed_tag(worksheet_prefix.as_deref(), "sheetPr");
                    writer.write_event(Event::Start(BytesStart::new(sheet_pr_tag.as_str())))?;
                    write_page_setup_pr(&mut writer, worksheet_prefix.as_deref(), true)?;
                    writer.write_event(Event::End(BytesEnd::new(sheet_pr_tag.as_str())))?;
                }
                if !seen_page_margins {
                    let tag = crate::xml::prefixed_tag(worksheet_prefix.as_deref(), "pageMargins");
                    writer.write_event(build_page_margins_event(
                        tag.as_str(),
                        &settings.page_setup.margins,
                    )?)?;
                }
                if !seen_page_setup {
                    let tag = crate::xml::prefixed_tag(worksheet_prefix.as_deref(), "pageSetup");
                    writer
                        .write_event(build_page_setup_event(tag.as_str(), &settings.page_setup)?)?;
                }
                if !seen_row_breaks && !settings.manual_page_breaks.row_breaks_after.is_empty() {
                    write_row_breaks(
                        &mut writer,
                        worksheet_prefix.as_deref(),
                        &settings.manual_page_breaks,
                    )?;
                }
                if !seen_col_breaks && !settings.manual_page_breaks.col_breaks_after.is_empty() {
                    write_col_breaks(
                        &mut writer,
                        worksheet_prefix.as_deref(),
                        &settings.manual_page_breaks,
                    )?;
                }
                writer.write_event(event)?;
            }
            Event::Eof => break,
            _ => {
                writer.write_event(event)?;
            }
        }

        buf.clear();
    }

    Ok(writer.into_inner())
}

trait ScalingExt {
    fn is_fit_to(&self) -> bool;
}

impl ScalingExt for Scaling {
    fn is_fit_to(&self) -> bool {
        matches!(self, Scaling::FitTo { .. })
    }
}

fn update_page_setup_pr_event(
    event: &Event<'_>,
    fit_to_page: bool,
) -> Result<Event<'static>, PrintError> {
    let tag = match event {
        Event::Start(e) | Event::Empty(e) => {
            String::from_utf8_lossy(e.name().as_ref()).into_owned()
        }
        _ => unreachable!(),
    };
    let mut start = BytesStart::new(tag.as_str()).into_owned();

    let mut has_fit = false;
    if let Event::Start(e) | Event::Empty(e) = event {
        for attr in e.attributes().with_checks(false) {
            let attr = attr?;
            if attr.key.as_ref() == b"fitToPage" {
                has_fit = true;
                let v = if fit_to_page { "1" } else { "0" };
                start.push_attribute(("fitToPage", v));
            } else {
                start.push_attribute((attr.key.as_ref(), attr.value.as_ref()));
            }
        }
    }

    if fit_to_page && !has_fit {
        start.push_attribute(("fitToPage", "1"));
    }

    Ok(match event {
        Event::Start(_) => Event::Start(start),
        Event::Empty(_) => Event::Empty(start),
        _ => unreachable!(),
    })
}

fn write_page_setup_pr(
    writer: &mut Writer<Vec<u8>>,
    prefix: Option<&str>,
    fit_to_page: bool,
) -> Result<(), PrintError> {
    let fit = if fit_to_page { "1" } else { "0" };
    let tag = crate::xml::prefixed_tag(prefix, "pageSetUpPr");
    let mut start = BytesStart::new(tag.as_str());
    start.push_attribute(("fitToPage", fit));
    writer.write_event(Event::Empty(start))?;
    Ok(())
}

fn build_page_margins_event(
    tag: &str,
    margins: &PageMargins,
) -> Result<Event<'static>, PrintError> {
    let left = margins.left.to_string();
    let right = margins.right.to_string();
    let top = margins.top.to_string();
    let bottom = margins.bottom.to_string();
    let header = margins.header.to_string();
    let footer = margins.footer.to_string();
    let mut start = BytesStart::new(tag).into_owned();
    start.push_attribute(("left", left.as_str()));
    start.push_attribute(("right", right.as_str()));
    start.push_attribute(("top", top.as_str()));
    start.push_attribute(("bottom", bottom.as_str()));
    start.push_attribute(("header", header.as_str()));
    start.push_attribute(("footer", footer.as_str()));
    Ok(Event::Empty(start))
}

fn build_page_setup_event(tag: &str, page_setup: &PageSetup) -> Result<Event<'static>, PrintError> {
    let paper_size = page_setup.paper_size.code.to_string();
    let orientation = match page_setup.orientation {
        Orientation::Portrait => "portrait",
        Orientation::Landscape => "landscape",
    };

    let mut start = BytesStart::new(tag).into_owned();
    start.push_attribute(("paperSize", paper_size.as_str()));
    start.push_attribute(("orientation", orientation));

    match page_setup.scaling {
        Scaling::Percent(pct) => {
            let pct = pct.to_string();
            start.push_attribute(("scale", pct.as_str()));
        }
        Scaling::FitTo { width, height } => {
            let width = width.to_string();
            let height = height.to_string();
            start.push_attribute(("fitToWidth", width.as_str()));
            start.push_attribute(("fitToHeight", height.as_str()));
        }
    }

    Ok(Event::Empty(start))
}

fn write_row_breaks(
    writer: &mut Writer<Vec<u8>>,
    prefix: Option<&str>,
    breaks: &ManualPageBreaks,
) -> Result<(), PrintError> {
    let row_breaks_tag = crate::xml::prefixed_tag(prefix, "rowBreaks");
    let brk_tag = crate::xml::prefixed_tag(prefix, "brk");
    let count = breaks.row_breaks_after.len().to_string();
    let mut outer = BytesStart::new(row_breaks_tag.as_str());
    outer.push_attribute(("count", count.as_str()));
    outer.push_attribute(("manualBreakCount", count.as_str()));
    writer.write_event(Event::Start(outer))?;
    for id in &breaks.row_breaks_after {
        let id_str = id.to_string();
        let mut brk = BytesStart::new(brk_tag.as_str());
        brk.push_attribute(("id", id_str.as_str()));
        brk.push_attribute(("max", "16383"));
        brk.push_attribute(("man", "1"));
        writer.write_event(Event::Empty(brk))?;
    }
    writer.write_event(Event::End(BytesEnd::new(row_breaks_tag.as_str())))?;
    Ok(())
}

fn write_col_breaks(
    writer: &mut Writer<Vec<u8>>,
    prefix: Option<&str>,
    breaks: &ManualPageBreaks,
) -> Result<(), PrintError> {
    let col_breaks_tag = crate::xml::prefixed_tag(prefix, "colBreaks");
    let brk_tag = crate::xml::prefixed_tag(prefix, "brk");
    let count = breaks.col_breaks_after.len().to_string();
    let mut outer = BytesStart::new(col_breaks_tag.as_str());
    outer.push_attribute(("count", count.as_str()));
    outer.push_attribute(("manualBreakCount", count.as_str()));
    writer.write_event(Event::Start(outer))?;
    for id in &breaks.col_breaks_after {
        let id_str = id.to_string();
        let mut brk = BytesStart::new(brk_tag.as_str());
        brk.push_attribute(("id", id_str.as_str()));
        brk.push_attribute(("max", "1048575"));
        brk.push_attribute(("man", "1"));
        writer.write_event(Event::Empty(brk))?;
    }
    writer.write_event(Event::End(BytesEnd::new(col_breaks_tag.as_str())))?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::{Cursor, Write};

    const SPREADSHEETML_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    fn build_test_xlsx(entries: &[(&str, &[u8])]) -> Vec<u8> {
        let cursor = Cursor::new(Vec::new());
        let mut zip = ZipWriter::new(cursor);
        let options =
            FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);
        for (name, bytes) in entries {
            zip.start_file(*name, options).unwrap();
            zip.write_all(bytes).unwrap();
        }
        zip.finish().unwrap().into_inner()
    }

    #[test]
    fn parse_workbook_defined_print_names_matches_builtin_defined_names_case_insensitively() {
        let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
  <definedNames>
    <definedName name="_xlnm.print_area" localSheetId="0">Sheet1!$A$1:$A$1</definedName>
    <definedName name="_xlnm.PRINT_TITLES" localSheetId="0">Sheet1!$1:$1</definedName>
  </definedNames>
</workbook>"#;

        let parsed = parse_workbook_defined_print_names(workbook_xml).expect("parse print names");
        assert_eq!(parsed.len(), 1);
        assert_eq!(parsed[0].sheet_name, "Sheet1");
        assert_eq!(
            parsed[0].print_area.as_deref(),
            Some(
                &[crate::print::CellRange {
                    start_row: 1,
                    end_row: 1,
                    start_col: 1,
                    end_col: 1
                }][..]
            )
        );
        assert_eq!(
            parsed[0].print_titles,
            Some(crate::print::PrintTitles {
                repeat_rows: Some(crate::print::RowRange { start: 1, end: 1 }),
                repeat_cols: None,
            })
        );
    }

    #[test]
    fn update_workbook_xml_expands_self_closing_prefixed_root_and_inserts_defined_names() {
        let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>"#;

        let edits: HashMap<(String, usize), DefinedNameEdit> = HashMap::from([(
            ("_xlnm.Print_Area".to_string(), 0),
            DefinedNameEdit::Set("Sheet1!$A$1:$B$2".to_string()),
        )]);

        let updated = update_workbook_xml(workbook_xml, &edits).unwrap();
        let updated_str = std::str::from_utf8(&updated).unwrap();

        // Root should be expanded and remain prefixed.
        assert!(updated_str.contains("<x:workbook"));
        assert!(updated_str.contains("</x:workbook>"));

        // Inserted tags must use the SpreadsheetML prefix (`x:` here).
        assert!(updated_str.contains("<x:definedNames"));
        assert!(updated_str.contains("<x:definedName"));

        // Ensure output is well-formed and namespace-correct.
        let doc = roxmltree::Document::parse(updated_str).unwrap();
        let root = doc.root_element();
        assert_eq!(root.tag_name().name(), "workbook");
        assert_eq!(root.tag_name().namespace(), Some(SPREADSHEETML_NS));

        let defined_names = root
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "definedNames")
            .expect("missing definedNames");
        assert_eq!(defined_names.tag_name().namespace(), Some(SPREADSHEETML_NS));

        let defined_name = root
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "definedName")
            .expect("missing definedName");
        assert_eq!(defined_name.tag_name().namespace(), Some(SPREADSHEETML_NS));
    }

    #[test]
    fn update_workbook_xml_expands_self_closing_default_ns_root_and_inserts_unprefixed_defined_names()
    {
        let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;

        let edits: HashMap<(String, usize), DefinedNameEdit> = HashMap::from([(
            ("_xlnm.Print_Titles".to_string(), 0),
            DefinedNameEdit::Set("Sheet1!$1:$1".to_string()),
        )]);

        let updated = update_workbook_xml(workbook_xml, &edits).unwrap();
        let updated_str = std::str::from_utf8(&updated).unwrap();

        assert!(updated_str.contains("<workbook"));
        assert!(updated_str.contains("</workbook>"));

        // With SpreadsheetML as the default namespace, inserted tags should be unprefixed.
        assert!(updated_str.contains("<definedNames>"));
        assert!(updated_str.contains("<definedName"));

        let doc = roxmltree::Document::parse(updated_str).unwrap();
        let root = doc.root_element();
        assert_eq!(root.tag_name().name(), "workbook");
        assert_eq!(root.tag_name().namespace(), Some(SPREADSHEETML_NS));

        let defined_names = root
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "definedNames")
            .expect("missing definedNames");
        assert_eq!(defined_names.tag_name().namespace(), Some(SPREADSHEETML_NS));

        let defined_name = root
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "definedName")
            .expect("missing definedName");
        assert_eq!(defined_name.tag_name().namespace(), Some(SPREADSHEETML_NS));
    }

    #[test]
    fn update_workbook_xml_expands_self_closing_prefixed_defined_names_and_inserts_defined_name() {
        let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <x:definedNames/>
</x:workbook>"#;

        let edits: HashMap<(String, usize), DefinedNameEdit> = HashMap::from([(
            ("_xlnm.Print_Area".to_string(), 0),
            DefinedNameEdit::Set("Sheet1!$A$1:$B$2".to_string()),
        )]);

        let updated = update_workbook_xml(workbook_xml, &edits).unwrap();
        let updated_str = std::str::from_utf8(&updated).unwrap();

        // The existing self-closing `<x:definedNames/>` should be expanded, not duplicated.
        assert!(updated_str.contains("<x:definedNames>"));
        assert!(updated_str.contains("</x:definedNames>"));
        assert!(!updated_str.contains("<x:definedNames/>"));

        // Inserted tags must use the SpreadsheetML prefix (`x:` here).
        assert!(updated_str.contains("<x:definedName"));
        assert!(updated_str.contains("</x:definedName>"));
        assert!(!updated_str.contains("<definedNames"));
        assert!(!updated_str.contains("<definedName"));

        // Ensure output is well-formed and namespace-correct.
        let doc = roxmltree::Document::parse(updated_str).unwrap();
        let root = doc.root_element();
        assert_eq!(root.tag_name().name(), "workbook");
        assert_eq!(root.tag_name().namespace(), Some(SPREADSHEETML_NS));

        let defined_names: Vec<_> = root
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "definedNames")
            .collect();
        assert_eq!(defined_names.len(), 1, "expected exactly one <definedNames> element");
        assert_eq!(
            defined_names[0].tag_name().namespace(),
            Some(SPREADSHEETML_NS)
        );

        let defined_name = defined_names[0]
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "definedName")
            .expect("missing inserted definedName");
        assert_eq!(defined_name.tag_name().namespace(), Some(SPREADSHEETML_NS));
        assert_eq!(defined_name.attribute("name"), Some("_xlnm.Print_Area"));
        assert_eq!(defined_name.attribute("localSheetId"), Some("0"));
        assert_eq!(defined_name.text(), Some("Sheet1!$A$1:$B$2"));
    }

    #[test]
    fn update_workbook_xml_expands_self_closing_default_ns_defined_name_on_set() {
        let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <definedNames>
    <definedName name="_xlnm.Print_Titles" localSheetId="0"/>
  </definedNames>
</workbook>"#;

        let edits: HashMap<(String, usize), DefinedNameEdit> = HashMap::from([(
            ("_xlnm.Print_Titles".to_string(), 0),
            DefinedNameEdit::Set("Sheet1!$1:$1".to_string()),
        )]);

        let updated = update_workbook_xml(workbook_xml, &edits).unwrap();
        let updated_str = std::str::from_utf8(&updated).unwrap();

        // The self-closing element should now have an explicit end tag.
        assert!(updated_str.contains("</definedName>"));

        // With SpreadsheetML as the default namespace, tags should remain unprefixed.
        assert!(!updated_str.contains("<x:definedName"));
        assert!(!updated_str.contains("<x:definedNames"));

        let doc = roxmltree::Document::parse(updated_str).unwrap();
        let root = doc.root_element();
        assert_eq!(root.tag_name().name(), "workbook");
        assert_eq!(root.tag_name().namespace(), Some(SPREADSHEETML_NS));

        let count = root
            .descendants()
            .filter(|n| {
                n.is_element()
                    && n.tag_name().name() == "definedName"
                    && n.attribute("name") == Some("_xlnm.Print_Titles")
                    && n.attribute("localSheetId") == Some("0")
            })
            .count();
        assert_eq!(count, 1, "expected exactly one updated definedName, got {count}");

        let defined_name = root
            .descendants()
            .find(|n| {
                n.is_element()
                    && n.tag_name().name() == "definedName"
                    && n.attribute("name") == Some("_xlnm.Print_Titles")
                    && n.attribute("localSheetId") == Some("0")
            })
            .expect("missing definedName");
        assert_eq!(defined_name.tag_name().namespace(), Some(SPREADSHEETML_NS));
        assert_eq!(defined_name.text(), Some("Sheet1!$1:$1"));
    }

    #[test]
    fn update_workbook_xml_expands_self_closing_prefixed_defined_name_on_set() {
        let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <x:definedNames>
    <x:definedName name="_xlnm.Print_Titles" localSheetId="0"/>
  </x:definedNames>
</x:workbook>"#;

        let edits: HashMap<(String, usize), DefinedNameEdit> = HashMap::from([(
            ("_xlnm.Print_Titles".to_string(), 0),
            DefinedNameEdit::Set("Sheet1!$1:$1".to_string()),
        )]);

        let updated = update_workbook_xml(workbook_xml, &edits).unwrap();
        let updated_str = std::str::from_utf8(&updated).unwrap();

        // The self-closing element should now have an explicit end tag (and remain prefixed).
        assert!(updated_str.contains("</x:definedName>"));

        // Do not introduce unprefixed SpreadsheetML tags.
        assert!(!updated_str.contains("<definedNames"));
        assert!(!updated_str.contains("<definedName"));

        let doc = roxmltree::Document::parse(updated_str).unwrap();
        let root = doc.root_element();
        assert_eq!(root.tag_name().name(), "workbook");
        assert_eq!(root.tag_name().namespace(), Some(SPREADSHEETML_NS));

        let count = root
            .descendants()
            .filter(|n| {
                n.is_element()
                    && n.tag_name().name() == "definedName"
                    && n.attribute("name") == Some("_xlnm.Print_Titles")
                    && n.attribute("localSheetId") == Some("0")
            })
            .count();
        assert_eq!(count, 1, "expected exactly one updated definedName, got {count}");

        let defined_name = root
            .descendants()
            .find(|n| {
                n.is_element()
                    && n.tag_name().name() == "definedName"
                    && n.attribute("name") == Some("_xlnm.Print_Titles")
                    && n.attribute("localSheetId") == Some("0")
            })
            .expect("missing definedName");
        assert_eq!(defined_name.tag_name().namespace(), Some(SPREADSHEETML_NS));
        assert_eq!(defined_name.text(), Some("Sheet1!$1:$1"));
    }

    #[test]
    fn update_workbook_xml_removes_self_closing_prefixed_defined_name_on_remove() {
        let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <x:definedNames>
    <x:definedName name="_xlnm.Print_Area" localSheetId="0"/>
    <x:definedName name="_xlnm.Print_Titles" localSheetId="0">Sheet1!$1:$1</x:definedName>
  </x:definedNames>
</x:workbook>"#;

        let edits: HashMap<(String, usize), DefinedNameEdit> = HashMap::from([(
            ("_xlnm.Print_Area".to_string(), 0),
            DefinedNameEdit::Remove,
        )]);

        let updated = update_workbook_xml(workbook_xml, &edits).unwrap();
        let updated_str = std::str::from_utf8(&updated).unwrap();

        // Do not introduce unprefixed SpreadsheetML tags.
        assert!(!updated_str.contains("<definedNames"));
        assert!(!updated_str.contains("<definedName"));

        let doc = roxmltree::Document::parse(updated_str).unwrap();
        let root = doc.root_element();

        let removed = root.descendants().find(|n| {
            n.is_element()
                && n.tag_name().name() == "definedName"
                && n.attribute("name") == Some("_xlnm.Print_Area")
                && n.attribute("localSheetId") == Some("0")
        });
        assert!(removed.is_none(), "expected definedName to be removed");

        let kept = root.descendants().find(|n| {
            n.is_element()
                && n.tag_name().name() == "definedName"
                && n.attribute("name") == Some("_xlnm.Print_Titles")
                && n.attribute("localSheetId") == Some("0")
        });
        assert!(kept.is_some(), "expected unrelated definedName to remain");
    }

    #[test]
    fn read_workbook_print_settings_errors_when_worksheet_part_exceeds_limit() {
        let max = 64u64;
        let workbook_xml = br#"<sheet name="S" id="rId1"/>"#;
        let rels_xml = br#"<Relationship Id="rId1" Target="worksheets/sheet1.xml"/>"#;
        let sheet_xml = vec![b'a'; (max + 1) as usize];
        let xlsx_bytes = build_test_xlsx(&[
            ("xl/workbook.xml", workbook_xml),
            ("xl/_rels/workbook.xml.rels", rels_xml),
            ("xl/worksheets/sheet1.xml", sheet_xml.as_slice()),
        ]);

        let err = read_workbook_print_settings_with_limit(&xlsx_bytes, max).unwrap_err();
        match err {
            PrintError::PartTooLarge { part, size, max } => {
                assert_eq!(part, "xl/worksheets/sheet1.xml");
                assert_eq!(size, max + 1);
            }
            other => panic!("expected PartTooLarge error, got {other:?}"),
        }
    }

    #[test]
    fn update_workbook_xml_removes_self_closing_default_ns_defined_name_on_remove() {
        let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <definedNames>
    <definedName name="_xlnm.Print_Area" localSheetId="0"/>
    <definedName name="_xlnm.Print_Titles" localSheetId="0">Sheet1!$1:$1</definedName>
  </definedNames>
</workbook>"#;

        let edits: HashMap<(String, usize), DefinedNameEdit> = HashMap::from([(
            ("_xlnm.Print_Area".to_string(), 0),
            DefinedNameEdit::Remove,
        )]);

        let updated = update_workbook_xml(workbook_xml, &edits).unwrap();
        let updated_str = std::str::from_utf8(&updated).unwrap();

        // Do not introduce SpreadsheetML prefixes when the workbook uses the default namespace.
        assert!(!updated_str.contains("<x:definedNames"));
        assert!(!updated_str.contains("<x:definedName"));

        let doc = roxmltree::Document::parse(updated_str).unwrap();
        let root = doc.root_element();

        let removed = root.descendants().find(|n| {
            n.is_element()
                && n.tag_name().name() == "definedName"
                && n.attribute("name") == Some("_xlnm.Print_Area")
                && n.attribute("localSheetId") == Some("0")
        });
        assert!(removed.is_none(), "expected definedName to be removed");

        let kept = root.descendants().find(|n| {
            n.is_element()
                && n.tag_name().name() == "definedName"
                && n.attribute("name") == Some("_xlnm.Print_Titles")
                && n.attribute("localSheetId") == Some("0")
        });
        assert!(kept.is_some(), "expected unrelated definedName to remain");
    }
}
