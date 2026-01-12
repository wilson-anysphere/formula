use std::borrow::Cow;
use std::collections::{BTreeMap, HashMap};
use std::io::{Cursor, Read, Seek, SeekFrom};
#[cfg(not(target_arch = "wasm32"))]
use std::fs::File;
#[cfg(not(target_arch = "wasm32"))]
use std::path::Path;

use formula_engine::{parse_formula, CellAddr, ParseOptions, SerializeOptions};
use formula_model::rich_text::RichText;
use formula_model::{
    normalize_formula_text, Cell, CellRef, CellValue, DefinedNameScope, ErrorValue, Range,
    SheetVisibility, Workbook,
};
use quick_xml::events::attributes::AttrError;
use quick_xml::events::Event;
use quick_xml::Reader;
use thiserror::Error;
use zip::ZipArchive;

use crate::autofilter::{parse_worksheet_autofilter, AutoFilterParseError};
use crate::calc_settings::read_calc_settings_from_workbook_xml;
use crate::conditional_formatting::parse_worksheet_conditional_formatting_streaming;
use crate::path::{rels_for_part, resolve_target};
use crate::shared_strings::parse_shared_strings_xml;
use crate::sheet_metadata::parse_sheet_tab_color;
use crate::styles::StylesPart;
use crate::tables::{parse_table, TABLE_REL_TYPE};
use crate::{parse_worksheet_hyperlinks, XlsxError};
use crate::{
    CalcPr, CellMeta, CellValueKind, DateSystem, FormulaMeta, SheetMeta, XlsxDocument, XlsxMeta,
};

const WORKBOOK_PART: &str = "xl/workbook.xml";
const WORKBOOK_RELS_PART: &str = "xl/_rels/workbook.xml.rels";
const REL_TYPE_STYLES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const REL_TYPE_SHARED_STRINGS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";

#[derive(Debug, Error)]
pub enum ReadError {
    #[error("io error: {0}")]
    Io(#[from] std::io::Error),
    #[error("zip error: {0}")]
    Zip(#[from] zip::result::ZipError),
    #[error("xml error: {0}")]
    Xml(#[from] quick_xml::Error),
    #[error("xml attribute error: {0}")]
    XmlAttr(#[from] AttrError),
    #[error("utf-8 error: {0}")]
    Utf8(#[from] std::str::Utf8Error),
    #[error("sharedStrings.xml parse error: {0}")]
    SharedStrings(#[from] crate::shared_strings::SharedStringsError),
    #[error(transparent)]
    Styles(#[from] crate::styles::StylesPartError),
    #[error("invalid worksheet name: {0}")]
    InvalidSheetName(#[from] formula_model::SheetNameError),
    #[error(transparent)]
    Xlsx(#[from] XlsxError),
    #[error("missing required part: {0}")]
    MissingPart(&'static str),
    #[error("invalid cell reference: {0}")]
    InvalidCellRef(String),
    #[error("invalid range reference: {0}")]
    InvalidRangeRef(String),
}

impl From<crate::calc_settings::CalcSettingsError> for ReadError {
    fn from(err: crate::calc_settings::CalcSettingsError) -> Self {
        match err {
            crate::calc_settings::CalcSettingsError::MissingPart(part) => Self::MissingPart(part),
            crate::calc_settings::CalcSettingsError::Io(err) => Self::Io(err),
            crate::calc_settings::CalcSettingsError::Xml(err) => Self::Xml(err),
            crate::calc_settings::CalcSettingsError::XmlAttr(err) => Self::XmlAttr(err),
            crate::calc_settings::CalcSettingsError::Utf8(err) => Self::Utf8(err),
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
pub fn load_from_path(path: impl AsRef<Path>) -> Result<XlsxDocument, ReadError> {
    let mut file = File::open(path)?;
    let mut bytes = Vec::new();
    file.read_to_end(&mut bytes)?;
    load_from_bytes(&bytes)
}

/// Read an XLSX workbook model from in-memory bytes without materializing every
/// ZIP part into memory.
///
/// This is a lightweight alternative to [`load_from_bytes`] that only inflates
/// the parts required to build a [`formula_model::Workbook`] (workbook metadata,
/// styles, shared strings, and referenced worksheets).
pub fn read_workbook_model_from_bytes(bytes: &[u8]) -> Result<Workbook, ReadError> {
    read_workbook_model_from_reader(Cursor::new(bytes))
}

/// Read an XLSX workbook model directly from a seekable reader without inflating
/// the entire XLSX package (or every ZIP part) into memory.
pub fn read_workbook_model_from_reader<R: Read + Seek>(mut reader: R) -> Result<Workbook, ReadError> {
    // Ensure we start from the beginning; callers may pass a reused reader.
    reader.seek(SeekFrom::Start(0))?;
    let mut archive = ZipArchive::new(reader)?;
    read_workbook_model_from_zip(&mut archive)
}

fn read_workbook_model_from_zip<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
) -> Result<Workbook, ReadError> {
    let workbook_xml = read_zip_part_required(archive, WORKBOOK_PART)?;
    let workbook_rels = read_zip_part_required(archive, WORKBOOK_RELS_PART)?;

    let rels_info = parse_relationships(&workbook_rels)?;
    let (date_system, _calc_pr, sheets, defined_names) =
        parse_workbook_metadata(&workbook_xml, &rels_info.id_to_target)?;
    let calc_settings = read_calc_settings_from_workbook_xml(&workbook_xml)?;

    let mut workbook = Workbook::new();
    workbook.calc_settings = calc_settings;
    workbook.date_system = match date_system {
        DateSystem::V1900 => formula_model::DateSystem::Excel1900,
        DateSystem::V1904 => formula_model::DateSystem::Excel1904,
    };
    let mut worksheet_ids_by_index: Vec<formula_model::WorksheetId> =
        Vec::with_capacity(sheets.len());

    let styles_part_name = rels_info
        .styles_target
        .as_deref()
        .map(|target| resolve_target(WORKBOOK_PART, target))
        .unwrap_or_else(|| "xl/styles.xml".to_string());
    let styles_bytes = read_zip_part_optional(archive, &styles_part_name)?;
    let styles_part = StylesPart::parse_or_default(styles_bytes.as_deref(), &mut workbook.styles)?;
    let conditional_formatting_dxfs: Vec<formula_model::CfStyleOverride> = styles_bytes
        .as_deref()
        .and_then(|bytes| std::str::from_utf8(bytes).ok())
        .and_then(|xml| crate::styles::Styles::parse(xml).ok())
        .map(|s| s.dxfs)
        .unwrap_or_default();

    let shared_strings_part_name = rels_info
        .shared_strings_target
        .as_deref()
        .map(|target| resolve_target(WORKBOOK_PART, target))
        .unwrap_or_else(|| "xl/sharedStrings.xml".to_string());
    let shared_strings = match read_zip_part_optional(archive, &shared_strings_part_name)? {
        Some(bytes) => parse_shared_strings(&bytes)?,
        None => Vec::new(),
    };

    for sheet in sheets {
        let ws_id = workbook.add_sheet(sheet.name.clone())?;
        worksheet_ids_by_index.push(ws_id);
        let ws = workbook
            .sheet_mut(ws_id)
            .expect("sheet just inserted must exist");
        ws.xlsx_sheet_id = Some(sheet.sheet_id);
        ws.xlsx_rel_id = Some(sheet.relationship_id.clone());
        ws.visibility = match sheet.state.as_deref() {
            Some("hidden") => SheetVisibility::Hidden,
            Some("veryHidden") => SheetVisibility::VeryHidden,
            _ => SheetVisibility::Visible,
        };

        let sheet_xml = read_zip_part_optional(archive, &sheet.path)?
            .ok_or(ReadError::MissingPart(
                "worksheet part referenced from workbook.xml.rels",
            ))?;

        // Worksheet-level metadata lives inside the worksheet part (and sometimes its .rels).
        let sheet_xml_str = std::str::from_utf8(&sheet_xml)?;

        // Optional metadata: best-effort.
        ws.tab_color = parse_sheet_tab_color(sheet_xml_str).unwrap_or(None);

        // Conditional formatting: best-effort. This is parsed via a streaming extractor so we
        // don't DOM-parse the entire worksheet XML.
        if let Ok(parsed) = parse_worksheet_conditional_formatting_streaming(sheet_xml_str) {
            if !parsed.rules.is_empty() {
                ws.conditional_formatting_rules = parsed.rules;
                ws.conditional_formatting_dxfs = conditional_formatting_dxfs.clone();
            }
        }

        // Merged cells (must be parsed before cell content so we don't treat interior
        // cells as value-bearing).
        if let Ok(merges) = crate::merge_cells::read_merge_cells_from_worksheet_xml(sheet_xml_str) {
            for range in merges {
                let _ = ws.merged_regions.add(range);
            }
        }

        // Worksheet relationships are needed to resolve external hyperlink targets and table parts.
        let rels_part = rels_for_part(&sheet.path);
        let rels_xml_bytes = read_zip_part_optional(archive, &rels_part)?;
        let rels_xml = rels_xml_bytes
            .as_deref()
            .and_then(|bytes| std::str::from_utf8(bytes).ok());

        ws.hyperlinks = parse_worksheet_hyperlinks(sheet_xml_str, rels_xml).unwrap_or_default();

        ws.auto_filter = parse_worksheet_autofilter(sheet_xml_str).ok().flatten();

        attach_tables_from_parts(
            ws,
            &sheet.path,
            &sheet_xml,
            rels_xml_bytes.as_deref(),
            archive,
        );

        parse_worksheet_into_model(ws, ws_id, &sheet_xml, &shared_strings, &styles_part, None)?;
    }

    for defined in defined_names {
        let scope = match defined
            .local_sheet_id
            .and_then(|idx| worksheet_ids_by_index.get(idx as usize).copied())
        {
            Some(sheet_id) => DefinedNameScope::Sheet(sheet_id),
            None => DefinedNameScope::Workbook,
        };
        // Best-effort: ignore invalid/duplicate names so we can still load the workbook.
        let _ = workbook.create_defined_name(
            scope,
            defined.name,
            defined.value,
            defined.comment,
            defined.hidden,
            defined.local_sheet_id,
        );
    }

    Ok(workbook)
}

fn attach_tables_from_parts<R: Read + Seek>(
    worksheet: &mut formula_model::Worksheet,
    worksheet_part: &str,
    worksheet_xml: &[u8],
    worksheet_rels_xml: Option<&[u8]>,
    archive: &mut ZipArchive<R>,
) {
    attach_tables_from_part_getter(
        worksheet,
        worksheet_part,
        worksheet_xml,
        worksheet_rels_xml,
        |target| read_zip_part_optional(archive, target).ok().flatten().map(Cow::Owned),
    );
}

fn attach_tables_from_part_getter<'a, F>(
    worksheet: &mut formula_model::Worksheet,
    worksheet_part: &str,
    worksheet_xml: &[u8],
    worksheet_rels_xml: Option<&[u8]>,
    mut get_part: F,
) where
    F: FnMut(&str) -> Option<Cow<'a, [u8]>>,
{
    let table_rel_ids = match parse_table_part_ids(worksheet_xml) {
        Ok(ids) => ids,
        Err(_) => Vec::new(),
    };
    if table_rel_ids.is_empty() {
        return;
    }

    let Some(rels_xml) = worksheet_rels_xml else {
        return;
    };

    let relationships = match crate::openxml::parse_relationships(rels_xml) {
        Ok(rels) => rels,
        Err(_) => return,
    };

    let mut rels_by_id: HashMap<String, crate::openxml::Relationship> =
        HashMap::with_capacity(relationships.len());
    for rel in relationships {
        rels_by_id.insert(rel.id.clone(), rel);
    }

    let mut seen_rel_ids: std::collections::HashSet<String> = worksheet
        .tables
        .iter()
        .filter_map(|t| t.relationship_id.clone())
        .collect();

    for r_id in table_rel_ids {
        if !seen_rel_ids.insert(r_id.clone()) {
            continue;
        }

        let Some(rel) = rels_by_id.get(&r_id) else {
            continue;
        };
        if rel.type_uri != TABLE_REL_TYPE {
            continue;
        }

        let target = resolve_target(worksheet_part, &rel.target);
        let table_bytes = match get_part(&target) {
            Some(bytes) => bytes,
            None => continue,
        };

        let table_xml = match std::str::from_utf8(table_bytes.as_ref()) {
            Ok(xml) => xml,
            Err(_) => continue,
        };

        let mut table = match parse_table(table_xml) {
            Ok(t) => t,
            Err(_) => continue,
        };
        table.relationship_id = Some(r_id);
        table.part_path = Some(target);
        worksheet.tables.push(table);
    }
}

fn read_zip_part_required<R: Read + std::io::Seek>(
    archive: &mut ZipArchive<R>,
    name: &'static str,
) -> Result<Vec<u8>, ReadError> {
    read_zip_part_optional(archive, name)?.ok_or(ReadError::MissingPart(name))
}

fn read_zip_part_optional<R: Read + std::io::Seek>(
    archive: &mut ZipArchive<R>,
    name: &str,
) -> Result<Option<Vec<u8>>, ReadError> {
    match archive.by_name(name) {
        Ok(mut file) => {
            if file.is_dir() {
                return Ok(None);
            }
            let mut buf = Vec::with_capacity(file.size() as usize);
            file.read_to_end(&mut buf)?;
            Ok(Some(buf))
        }
        Err(zip::result::ZipError::FileNotFound) => Ok(None),
        Err(err) => Err(err.into()),
    }
}

pub fn load_from_bytes(bytes: &[u8]) -> Result<XlsxDocument, ReadError> {
    let cursor = Cursor::new(bytes);
    let mut archive = ZipArchive::new(cursor)?;

    let mut parts: BTreeMap<String, Vec<u8>> = BTreeMap::new();
    for i in 0..archive.len() {
        let mut file = archive.by_index(i)?;
        if file.is_dir() {
            continue;
        }
        let name = file.name().to_string();
        let mut buf = Vec::with_capacity(file.size() as usize);
        file.read_to_end(&mut buf)?;
        parts.insert(name, buf);
    }

    let workbook_xml = parts
        .get(WORKBOOK_PART)
        .ok_or(ReadError::MissingPart(WORKBOOK_PART))?;
    let workbook_rels = parts
        .get(WORKBOOK_RELS_PART)
        .ok_or(ReadError::MissingPart(WORKBOOK_RELS_PART))?;

    let rels_info = parse_relationships(workbook_rels)?;
    let (date_system, calc_pr, sheets, defined_names) =
        parse_workbook_metadata(workbook_xml, &rels_info.id_to_target)?;
    let calc_settings = read_calc_settings_from_workbook_xml(workbook_xml)?;

    let mut workbook = Workbook::new();
    workbook.calc_settings = calc_settings;
    workbook.date_system = match date_system {
        DateSystem::V1900 => formula_model::DateSystem::Excel1900,
        DateSystem::V1904 => formula_model::DateSystem::Excel1904,
    };
    let styles_part_name = rels_info
        .styles_target
        .as_deref()
        .map(|target| resolve_target(WORKBOOK_PART, target))
        .unwrap_or_else(|| "xl/styles.xml".to_string());
    let conditional_formatting_dxfs: Vec<formula_model::CfStyleOverride> = parts
        .get(&styles_part_name)
        .and_then(|bytes| std::str::from_utf8(bytes).ok())
        .and_then(|xml| crate::styles::Styles::parse(xml).ok())
        .map(|s| s.dxfs)
        .unwrap_or_default();
    let styles_part = StylesPart::parse_or_default(
        parts.get(&styles_part_name).map(|b| b.as_slice()),
        &mut workbook.styles,
    )?;

    let shared_strings_part_name = rels_info
        .shared_strings_target
        .as_deref()
        .map(|target| resolve_target(WORKBOOK_PART, target))
        .unwrap_or_else(|| "xl/sharedStrings.xml".to_string());
    let shared_strings = if let Some(bytes) = parts.get(&shared_strings_part_name) {
        parse_shared_strings(bytes)?
    } else {
        Vec::new()
    };
    let mut sheet_meta: Vec<SheetMeta> = Vec::with_capacity(sheets.len());
    let mut cell_meta = std::collections::HashMap::new();

    let mut worksheet_ids_by_index: Vec<formula_model::WorksheetId> = Vec::new();
    for sheet in sheets {
        let ws_id = workbook.add_sheet(sheet.name.clone())?;
        worksheet_ids_by_index.push(ws_id);
        let ws = workbook
            .sheet_mut(ws_id)
            .expect("sheet just inserted must exist");
        ws.xlsx_sheet_id = Some(sheet.sheet_id);
        ws.xlsx_rel_id = Some(sheet.relationship_id.clone());
        ws.visibility = match sheet.state.as_deref() {
            Some("hidden") => SheetVisibility::Hidden,
            Some("veryHidden") => SheetVisibility::VeryHidden,
            _ => SheetVisibility::Visible,
        };

        let sheet_xml = parts.get(&sheet.path).ok_or(ReadError::MissingPart(
            "worksheet part referenced from workbook.xml.rels",
        ))?;

        // Worksheet-level metadata lives inside the worksheet part (and sometimes its .rels).
        let sheet_xml_str = std::str::from_utf8(sheet_xml)?;

        ws.tab_color = parse_sheet_tab_color(sheet_xml_str)?;

        // Conditional formatting. Parsed via a streaming extractor so we don't DOM-parse the
        // full worksheet XML.
        let parsed_cf = parse_worksheet_conditional_formatting_streaming(sheet_xml_str)
            .unwrap_or_default();
        if !parsed_cf.rules.is_empty() {
            ws.conditional_formatting_rules = parsed_cf.rules;
            ws.conditional_formatting_dxfs = conditional_formatting_dxfs.clone();
        }

        // Merged cells (must be parsed before cell content so we don't treat interior
        // cells as value-bearing).
        let merges = crate::merge_cells::read_merge_cells_from_worksheet_xml(sheet_xml_str)
            .map_err(|err| match err {
                crate::merge_cells::MergeCellsError::Xml(e) => ReadError::Xml(e),
                crate::merge_cells::MergeCellsError::Attr(e) => ReadError::XmlAttr(e),
                crate::merge_cells::MergeCellsError::Utf8(e) => ReadError::Utf8(e),
                crate::merge_cells::MergeCellsError::InvalidRef(r) => ReadError::InvalidRangeRef(r),
                crate::merge_cells::MergeCellsError::Zip(e) => ReadError::Zip(e),
                crate::merge_cells::MergeCellsError::Io(e) => ReadError::Io(e),
            })?;
        for range in merges {
            ws.merged_regions
                .add(range)
                .map_err(|e| ReadError::InvalidRangeRef(e.to_string()))?;
        }

        // Hyperlinks.
        let rels_part = rels_for_part(&sheet.path);
        let rels_xml_bytes = parts.get(&rels_part).map(|bytes| bytes.as_slice());
        let rels_xml = rels_xml_bytes.map(std::str::from_utf8).transpose()?;
        ws.hyperlinks = parse_worksheet_hyperlinks(sheet_xml_str, rels_xml)?;

        // Worksheet autoFilter.
        ws.auto_filter = parse_worksheet_autofilter(sheet_xml_str).map_err(|err| match err {
            AutoFilterParseError::Xml(e) => ReadError::Xml(e),
            AutoFilterParseError::Attr(e) => ReadError::XmlAttr(e),
            AutoFilterParseError::MissingRef => {
                ReadError::InvalidRangeRef("missing worksheet autoFilter ref attribute".to_string())
            }
            AutoFilterParseError::InvalidRef(e) => ReadError::InvalidRangeRef(e.to_string()),
        })?;

        attach_tables_from_part_getter(
            ws,
            &sheet.path,
            sheet_xml,
            rels_xml_bytes,
            |target| parts.get(target).map(|bytes| Cow::Borrowed(bytes.as_slice())),
        );

        parse_worksheet_into_model(
            ws,
            ws_id,
            sheet_xml,
            &shared_strings,
            &styles_part,
            Some(&mut cell_meta),
        )?;

        expand_shared_formulas(ws, ws_id, &cell_meta);

        sheet_meta.push(SheetMeta {
            worksheet_id: ws_id,
            sheet_id: sheet.sheet_id,
            relationship_id: sheet.relationship_id,
            state: sheet.state,
            path: sheet.path,
        });
    }

    for defined in defined_names {
        let scope = match defined
            .local_sheet_id
            .and_then(|idx| worksheet_ids_by_index.get(idx as usize).copied())
        {
            Some(sheet_id) => DefinedNameScope::Sheet(sheet_id),
            None => DefinedNameScope::Workbook,
        };
        // Best-effort: ignore invalid/duplicate names so we can still load the workbook.
        let _ = workbook.create_defined_name(
            scope,
            defined.name,
            defined.value,
            defined.comment,
            defined.hidden,
            defined.local_sheet_id,
        );
    }

    Ok(XlsxDocument {
        workbook,
        parts,
        shared_strings,
        meta: XlsxMeta {
            date_system,
            calc_pr,
            sheets: sheet_meta,
            cell_meta,
        },
        calc_affecting_edits: false,
    })
}

fn parse_relationships(bytes: &[u8]) -> Result<RelationshipsInfo, ReadError> {
    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut id_to_target = BTreeMap::new();
    let mut styles_target = None;
    let mut shared_strings_target = None;
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == b"Relationship" => {
                let mut id = None;
                let mut type_ = None;
                let mut target = None;
                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"Id" => id = Some(attr.unescape_value()?.into_owned()),
                        b"Type" => type_ = Some(attr.unescape_value()?.into_owned()),
                        b"Target" => target = Some(attr.unescape_value()?.into_owned()),
                        _ => {}
                    }
                }
                if let (Some(id), Some(target)) = (id, target) {
                    if let Some(type_) = &type_ {
                        match type_.as_str() {
                            REL_TYPE_STYLES => {
                                styles_target.get_or_insert_with(|| target.clone());
                            }
                            REL_TYPE_SHARED_STRINGS => {
                                shared_strings_target.get_or_insert_with(|| target.clone());
                            }
                            _ => {}
                        }
                    }
                    id_to_target.insert(id, target);
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(RelationshipsInfo {
        id_to_target,
        styles_target,
        shared_strings_target,
    })
}

#[derive(Debug, Clone)]
struct RelationshipsInfo {
    id_to_target: BTreeMap<String, String>,
    styles_target: Option<String>,
    shared_strings_target: Option<String>,
}

fn parse_shared_strings(bytes: &[u8]) -> Result<Vec<RichText>, ReadError> {
    let xml = std::str::from_utf8(bytes)?;
    let parsed = parse_shared_strings_xml(xml)?;
    Ok(parsed.items)
}

#[derive(Debug, Clone)]
struct ParsedSheet {
    name: String,
    sheet_id: u32,
    relationship_id: String,
    state: Option<String>,
    path: String,
}

#[derive(Debug, Clone)]
struct ParsedDefinedName {
    name: String,
    local_sheet_id: Option<u32>,
    comment: Option<String>,
    hidden: bool,
    value: String,
}

fn parse_workbook_metadata(
    workbook_xml: &[u8],
    rels: &BTreeMap<String, String>,
) -> Result<(DateSystem, CalcPr, Vec<ParsedSheet>, Vec<ParsedDefinedName>), ReadError> {
    let mut reader = Reader::from_reader(workbook_xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    let mut date_system = DateSystem::V1900;
    let mut calc_pr = CalcPr::default();
    let mut sheets = Vec::new();
    let mut defined_names = Vec::new();
    let mut current_defined: Option<ParsedDefinedName> = None;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"workbookPr" => {
                for attr in e.attributes() {
                    let attr = attr?;
                    if attr.key.as_ref() == b"date1904" {
                        let val = attr.unescape_value()?.into_owned();
                        if val == "1" || val.eq_ignore_ascii_case("true") {
                            date_system = DateSystem::V1904;
                        }
                    }
                }
            }
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"calcPr" => {
                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"calcId" => calc_pr.calc_id = Some(attr.unescape_value()?.into_owned()),
                        b"calcMode" => {
                            calc_pr.calc_mode = Some(attr.unescape_value()?.into_owned())
                        }
                        b"fullCalcOnLoad" => {
                            let v = attr.unescape_value()?.into_owned();
                            calc_pr.full_calc_on_load =
                                Some(v == "1" || v.eq_ignore_ascii_case("true"))
                        }
                        _ => {}
                    }
                }
            }
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"sheet" => {
                let mut name = None;
                let mut sheet_id = None;
                let mut r_id = None;
                let mut state = None;
                for attr in e.attributes() {
                    let attr = attr?;
                    let key = attr.key.as_ref();
                    match key {
                        b"name" => name = Some(attr.unescape_value()?.into_owned()),
                        b"sheetId" => {
                            sheet_id =
                                Some(attr.unescape_value()?.into_owned().parse().unwrap_or(0))
                        }
                        b"state" => state = Some(attr.unescape_value()?.into_owned()),
                        _ if crate::openxml::local_name(key) == b"id" => {
                            r_id = Some(attr.unescape_value()?.into_owned())
                        }
                        _ => {}
                    }
                }
                let name = name.unwrap_or_else(|| "Sheet".to_string());
                let sheet_id = sheet_id.unwrap_or(0);
                let relationship_id = r_id.unwrap_or_else(|| "rId1".to_string());
                let target = rels
                    .get(&relationship_id)
                    .cloned()
                    .unwrap_or_else(|| "worksheets/sheet1.xml".to_string());
                let path = resolve_target(WORKBOOK_PART, &target);
                sheets.push(ParsedSheet {
                    name,
                    sheet_id,
                    relationship_id,
                    state,
                    path,
                });
            }
            Event::Start(e) if e.local_name().as_ref() == b"definedName" => {
                let mut name = None;
                let mut local_sheet_id = None;
                let mut comment = None;
                let mut hidden = false;
                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"name" => name = Some(attr.unescape_value()?.into_owned()),
                        b"localSheetId" => {
                            local_sheet_id =
                                attr.unescape_value()?.into_owned().parse::<u32>().ok();
                        }
                        b"comment" => comment = Some(attr.unescape_value()?.into_owned()),
                        b"hidden" => {
                            let v = attr.unescape_value()?.into_owned();
                            hidden = v == "1" || v.eq_ignore_ascii_case("true");
                        }
                        _ => {}
                    }
                }
                let Some(name) = name else {
                    current_defined = None;
                    continue;
                };
                current_defined = Some(ParsedDefinedName {
                    name,
                    local_sheet_id,
                    comment,
                    hidden,
                    value: String::new(),
                });
            }
            Event::Empty(e) if e.local_name().as_ref() == b"definedName" => {
                let mut name = None;
                let mut local_sheet_id = None;
                let mut comment = None;
                let mut hidden = false;
                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"name" => name = Some(attr.unescape_value()?.into_owned()),
                        b"localSheetId" => {
                            local_sheet_id =
                                attr.unescape_value()?.into_owned().parse::<u32>().ok();
                        }
                        b"comment" => comment = Some(attr.unescape_value()?.into_owned()),
                        b"hidden" => {
                            let v = attr.unescape_value()?.into_owned();
                            hidden = v == "1" || v.eq_ignore_ascii_case("true");
                        }
                        _ => {}
                    }
                }
                let Some(name) = name else {
                    continue;
                };
                defined_names.push(ParsedDefinedName {
                    name,
                    local_sheet_id,
                    comment,
                    hidden,
                    value: String::new(),
                });
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

    Ok((date_system, calc_pr, sheets, defined_names))
}

fn parse_worksheet_into_model(
    worksheet: &mut formula_model::Worksheet,
    worksheet_id: formula_model::WorksheetId,
    worksheet_xml: &[u8],
    shared_strings: &[RichText],
    styles_part: &StylesPart,
    mut cell_meta_map: Option<
        &mut std::collections::HashMap<(formula_model::WorksheetId, CellRef), CellMeta>,
    >,
) -> Result<(), ReadError> {
    let mut reader = Reader::from_reader(worksheet_xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    let mut in_sheet_data = false;
    let mut in_cols = false;
    let mut in_sheet_views = false;
    let mut in_sheet_view = false;

    // When we don't retain the full `cell_meta` map (fast reader), we still want to materialize
    // shared-formula followers into the worksheet model so formulas match the full reader.
    let mut shared_formula_groups: Option<HashMap<u32, SharedFormulaGroup>> =
        cell_meta_map.is_none().then(HashMap::new);

    let mut current_ref: Option<CellRef> = None;
    let mut current_t: Option<String> = None;
    let mut current_style: u32 = 0;
    let mut current_formula: Option<FormulaMeta> = None;
    let mut current_value_text: Option<String> = None;
    let mut current_inline_text: Option<String> = None;
    let mut in_v = false;
    let mut in_f = false;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if e.local_name().as_ref() == b"cols" => in_cols = true,
            Event::End(e) if e.local_name().as_ref() == b"cols" => in_cols = false,
            Event::Empty(e) if e.local_name().as_ref() == b"cols" => {
                in_cols = false;
                drop(e);
            }
            Event::Start(e) | Event::Empty(e) if in_cols && e.local_name().as_ref() == b"col" => {
                let mut min: Option<u32> = None;
                let mut max: Option<u32> = None;
                let mut width: Option<f32> = None;
                let mut custom_width: Option<bool> = None;
                let mut hidden = false;

                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"min" => {
                            min = Some(attr.unescape_value()?.into_owned().parse().unwrap_or(0))
                        }
                        b"max" => {
                            max = Some(attr.unescape_value()?.into_owned().parse().unwrap_or(0))
                        }
                        b"width" => {
                            width = attr.unescape_value()?.into_owned().parse::<f32>().ok();
                        }
                        b"customWidth" => {
                            let v = attr.unescape_value()?.into_owned();
                            custom_width = Some(parse_xml_bool(&v));
                        }
                        b"hidden" => {
                            let v = attr.unescape_value()?.into_owned();
                            hidden = parse_xml_bool(&v);
                        }
                        _ => {}
                    }
                }

                let Some(min) = min else {
                    continue;
                };
                let max = max.unwrap_or(min).min(formula_model::EXCEL_MAX_COLS);
                if min == 0 || max == 0 || min > formula_model::EXCEL_MAX_COLS {
                    continue;
                }

                for col_1_based in min..=max {
                    let col = col_1_based - 1;
                    if custom_width != Some(false) {
                        if let Some(width) = width {
                            worksheet.set_col_width(col, Some(width));
                        }
                    }
                    if hidden {
                        worksheet.set_col_hidden(col, true);
                    }
                }
            }

            Event::Start(e) if e.local_name().as_ref() == b"sheetData" => in_sheet_data = true,
            Event::End(e) if e.local_name().as_ref() == b"sheetData" => in_sheet_data = false,
            Event::Empty(e) if e.local_name().as_ref() == b"sheetData" => {
                in_sheet_data = false;
                drop(e);
            }

            Event::Start(e) if e.local_name().as_ref() == b"sheetViews" => in_sheet_views = true,
            Event::End(e) if e.local_name().as_ref() == b"sheetViews" => {
                in_sheet_views = false;
                in_sheet_view = false;
                drop(e);
            }
            Event::Empty(e) if e.local_name().as_ref() == b"sheetViews" => {
                in_sheet_views = false;
                in_sheet_view = false;
                drop(e);
            }

            Event::Start(e) if in_sheet_views && e.local_name().as_ref() == b"sheetView" => {
                in_sheet_view = true;
                for attr in e.attributes() {
                    let attr = attr?;
                    if attr.key.as_ref() == b"zoomScale" {
                        if let Ok(scale) = attr.unescape_value()?.into_owned().parse::<f32>() {
                            worksheet.zoom = scale / 100.0;
                        }
                    }
                }
            }
            Event::Empty(e) if in_sheet_views && e.local_name().as_ref() == b"sheetView" => {
                for attr in e.attributes() {
                    let attr = attr?;
                    if attr.key.as_ref() == b"zoomScale" {
                        if let Ok(scale) = attr.unescape_value()?.into_owned().parse::<f32>() {
                            worksheet.zoom = scale / 100.0;
                        }
                    }
                }
            }
            Event::End(e) if in_sheet_view && e.local_name().as_ref() == b"sheetView" => {
                in_sheet_view = false;
                drop(e);
            }

            Event::Start(e) | Event::Empty(e) if in_sheet_view && e.local_name().as_ref() == b"pane" => {
                let mut state: Option<String> = None;
                let mut x_split: Option<u32> = None;
                let mut y_split: Option<u32> = None;
                for attr in e.attributes() {
                    let attr = attr?;
                    let val = attr.unescape_value()?.into_owned();
                    match attr.key.as_ref() {
                        b"state" => state = Some(val),
                        b"xSplit" => x_split = val.parse().ok(),
                        b"ySplit" => y_split = val.parse().ok(),
                        _ => {}
                    }
                }
                if matches!(state.as_deref(), Some("frozen") | Some("frozenSplit")) {
                    worksheet.frozen_cols = x_split.unwrap_or(0);
                    worksheet.frozen_rows = y_split.unwrap_or(0);
                }
            }

            Event::Start(e) | Event::Empty(e) if in_sheet_data && e.local_name().as_ref() == b"row" => {
                let mut row_1_based: Option<u32> = None;
                let mut height: Option<f32> = None;
                let mut custom_height: Option<bool> = None;
                let mut hidden = false;

                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"r" => {
                            row_1_based =
                                Some(attr.unescape_value()?.into_owned().parse().unwrap_or(0));
                        }
                        b"ht" => {
                            height = attr.unescape_value()?.into_owned().parse::<f32>().ok();
                        }
                        b"customHeight" => {
                            let v = attr.unescape_value()?.into_owned();
                            custom_height = Some(parse_xml_bool(&v));
                        }
                        b"hidden" => {
                            let v = attr.unescape_value()?.into_owned();
                            hidden = parse_xml_bool(&v);
                        }
                        _ => {}
                    }
                }

                if let Some(row_1_based) = row_1_based {
                    // Accept any 1-based row index that fits in `u32` (the OOXML schema uses
                    // unsigned integers and our model supports sheets beyond Excel's UI limits).
                    if row_1_based > 0 {
                        let row = row_1_based - 1;
                        if custom_height != Some(false) {
                            if let Some(height) = height {
                                worksheet.set_row_height(row, Some(height));
                            }
                        }
                        if hidden {
                            worksheet.set_row_hidden(row, true);
                        }
                    }
                }
            }

            Event::Start(e) if in_sheet_data && e.local_name().as_ref() == b"c" => {
                current_ref = None;
                current_t = None;
                current_style = 0;
                current_formula = None;
                current_value_text = None;
                current_inline_text = None;
                in_v = false;
                in_f = false;

                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"r" => {
                            let a1 = attr.unescape_value()?.into_owned();
                            current_ref = Some(
                                CellRef::from_a1(&a1).map_err(|_| ReadError::InvalidCellRef(a1))?,
                            );
                        }
                        b"t" => current_t = Some(attr.unescape_value()?.into_owned()),
                        b"s" => {
                            let xf_index = attr.unescape_value()?.into_owned().parse().unwrap_or(0);
                            current_style = styles_part.style_id_for_xf(xf_index);
                        }
                        _ => {}
                    }
                }
            }
            Event::Empty(e) if in_sheet_data && e.local_name().as_ref() == b"c" => {
                let mut cell_ref = None;
                let mut style_id = 0u32;
                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"r" => {
                            let a1 = attr.unescape_value()?.into_owned();
                            cell_ref = Some(
                                CellRef::from_a1(&a1).map_err(|_| ReadError::InvalidCellRef(a1))?,
                            );
                        }
                        b"s" => {
                            let xf_index = attr.unescape_value()?.into_owned().parse().unwrap_or(0);
                            style_id = styles_part.style_id_for_xf(xf_index);
                        }
                        _ => {}
                    }
                }
                if let Some(cell_ref) = cell_ref {
                    // Skip non-anchor cells inside merged regions. Excel stores the value
                    // (and typically formatting) on the top-left cell only.
                    if worksheet.merged_regions.resolve_cell(cell_ref) == cell_ref && style_id != 0
                    {
                        let mut cell = Cell::default();
                        cell.style_id = style_id;
                        worksheet.set_cell(cell_ref, cell);
                    }
                }
            }

            Event::End(e) if in_sheet_data && e.local_name().as_ref() == b"c" => {
                if let Some(cell_ref) = current_ref {
                    if worksheet.merged_regions.resolve_cell(cell_ref) == cell_ref {
                        let (value, value_kind, raw_value) = if cell_meta_map.is_some() {
                            interpret_cell_value(
                                current_t.as_deref(),
                                &current_value_text,
                                &current_inline_text,
                                shared_strings,
                            )
                        } else {
                            (
                                interpret_cell_value_without_meta(
                                    current_t.as_deref(),
                                    &current_value_text,
                                    &current_inline_text,
                                    shared_strings,
                                ),
                                None,
                                None,
                            )
                        };

                        let formula_in_model = current_formula.as_ref().and_then(|f| {
                            let stripped = crate::formula_text::strip_xlfn_prefixes(&f.file_text);
                            normalize_formula_text(&stripped)
                        });

                        if let Some(groups) = shared_formula_groups.as_mut() {
                            let is_shared_master =
                                current_formula.as_ref().is_some_and(|formula| {
                                    formula.t.as_deref() == Some("shared")
                                        && formula.reference.is_some()
                                        && formula.shared_index.is_some()
                                        && !formula.file_text.is_empty()
                                });

                            if is_shared_master {
                                if let Some(formula) = current_formula.as_ref() {
                                    if let (Some(reference), Some(shared_index)) =
                                        (formula.reference.as_deref(), formula.shared_index)
                                    {
                                        if let Ok(range) = Range::from_a1(reference) {
                                            let master_display =
                                                crate::formula_text::strip_xlfn_prefixes(
                                                    &formula.file_text,
                                                );
                                            let mut opts = ParseOptions::default();
                                            opts.normalize_relative_to =
                                                Some(CellAddr::new(cell_ref.row, cell_ref.col));

                                            if let Ok(ast) = parse_formula(&master_display, opts) {
                                                groups.insert(
                                                    shared_index,
                                                    SharedFormulaGroup { range, ast },
                                                );
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        let mut cell = Cell::default();
                        cell.value = value;
                        cell.formula = formula_in_model;
                        cell.style_id = current_style;

                        if !cell.is_truly_empty() {
                            worksheet.set_cell(cell_ref, cell);
                        }

                        if let Some(cell_meta_map) = cell_meta_map.as_mut() {
                            let mut meta = CellMeta::default();
                            meta.value_kind = value_kind;
                            meta.raw_value = raw_value;
                            meta.formula = current_formula.take();

                            if meta.value_kind.is_some()
                                || meta.raw_value.is_some()
                                || meta.formula.is_some()
                                || current_style != 0
                            {
                                cell_meta_map.insert((worksheet_id, cell_ref), meta);
                            }
                        }
                    }
                }

                current_ref = None;
                current_t = None;
                current_style = 0;
                current_formula = None;
                current_value_text = None;
                current_inline_text = None;
                in_v = false;
                in_f = false;
            }

            Event::Start(e)
                if in_sheet_data && current_ref.is_some() && e.local_name().as_ref() == b"v" =>
            {
                in_v = true;
            }
            Event::End(e) if in_sheet_data && e.local_name().as_ref() == b"v" => in_v = false,
            Event::Text(e) if in_sheet_data && in_v => {
                current_value_text = Some(e.unescape()?.into_owned());
            }

            Event::Start(e)
                if in_sheet_data && current_ref.is_some() && e.local_name().as_ref() == b"f" =>
            {
                in_f = true;
                let mut formula = FormulaMeta::default();
                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"t" => formula.t = Some(attr.unescape_value()?.into_owned()),
                        b"ref" => formula.reference = Some(attr.unescape_value()?.into_owned()),
                        b"si" => {
                            formula.shared_index =
                                Some(attr.unescape_value()?.into_owned().parse().unwrap_or(0))
                        }
                        b"aca" => {
                            let v = attr.unescape_value()?.into_owned();
                            formula.always_calc = Some(v == "1" || v.eq_ignore_ascii_case("true"))
                        }
                        _ => {}
                    }
                }
                current_formula = Some(formula);
            }
            Event::Empty(e)
                if in_sheet_data && current_ref.is_some() && e.local_name().as_ref() == b"f" =>
            {
                let mut formula = FormulaMeta::default();
                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"t" => formula.t = Some(attr.unescape_value()?.into_owned()),
                        b"ref" => formula.reference = Some(attr.unescape_value()?.into_owned()),
                        b"si" => {
                            formula.shared_index =
                                Some(attr.unescape_value()?.into_owned().parse().unwrap_or(0))
                        }
                        b"aca" => {
                            let v = attr.unescape_value()?.into_owned();
                            formula.always_calc = Some(v == "1" || v.eq_ignore_ascii_case("true"))
                        }
                        _ => {}
                    }
                }
                current_formula = Some(formula);
            }
            Event::End(e) if in_sheet_data && e.local_name().as_ref() == b"f" => in_f = false,
            Event::Text(e) if in_sheet_data && in_f => {
                if let Some(formula) = current_formula.as_mut() {
                    formula.file_text = e.unescape()?.into_owned();
                }
            }

            Event::Start(e)
                if in_sheet_data
                    && current_ref.is_some()
                    && current_t.as_deref() == Some("inlineStr")
                    && e.local_name().as_ref() == b"is" =>
            {
                current_inline_text = Some(parse_inline_is_text(&mut reader)?);
            }
            Event::Empty(e)
                if in_sheet_data
                    && current_ref.is_some()
                    && current_t.as_deref() == Some("inlineStr")
                    && e.local_name().as_ref() == b"is" =>
            {
                current_inline_text = Some(String::new());
            }

            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    if let Some(groups) = shared_formula_groups {
        for group in groups.values() {
            for cell_ref in group.range.iter() {
                // Avoid overwriting any explicit formula already present in the worksheet. The goal is
                // to materialize formulas for shared-formula followers that are textless in the file.
                if worksheet.formula(cell_ref).is_some() {
                    continue;
                }

                let mut ser = SerializeOptions::default();
                ser.origin = Some(CellAddr::new(cell_ref.row, cell_ref.col));
                ser.omit_equals = true;

                let display = match group.ast.to_string(ser) {
                    Ok(s) => s,
                    Err(_) => continue,
                };

                worksheet.set_formula(cell_ref, Some(display));
            }
        }
    }

    Ok(())
}

#[derive(Debug, Clone)]
struct SharedFormulaGroup {
    range: Range,
    ast: formula_engine::Ast,
}

fn expand_shared_formulas(
    worksheet: &mut formula_model::Worksheet,
    worksheet_id: formula_model::WorksheetId,
    cell_meta_map: &HashMap<(formula_model::WorksheetId, CellRef), CellMeta>,
) {
    let mut groups: HashMap<u32, SharedFormulaGroup> = HashMap::new();

    for ((ws_id, cell_ref), meta) in cell_meta_map {
        if *ws_id != worksheet_id {
            continue;
        }
        let Some(formula) = meta.formula.as_ref() else {
            continue;
        };

        let is_shared_master = formula.t.as_deref() == Some("shared")
            && formula.reference.is_some()
            && formula.shared_index.is_some()
            && !formula.file_text.is_empty();
        if !is_shared_master {
            continue;
        }

        let Some(reference) = formula.reference.as_deref() else {
            continue;
        };
        let Some(shared_index) = formula.shared_index else {
            continue;
        };

        let range = match Range::from_a1(reference) {
            Ok(range) => range,
            Err(_) => continue,
        };

        let master_display = crate::formula_text::strip_xlfn_prefixes(&formula.file_text);
        let mut opts = ParseOptions::default();
        opts.normalize_relative_to = Some(CellAddr::new(cell_ref.row, cell_ref.col));

        let ast = match parse_formula(&master_display, opts) {
            Ok(ast) => ast,
            Err(_) => continue,
        };

        groups.insert(shared_index, SharedFormulaGroup { range, ast });
    }

    for group in groups.values() {
        for cell_ref in group.range.iter() {
            // Avoid overwriting any explicit formula already present in the worksheet. The goal is
            // to materialize formulas for shared-formula followers that are textless in the file.
            if worksheet.formula(cell_ref).is_some() {
                continue;
            }

            let mut ser = SerializeOptions::default();
            ser.origin = Some(CellAddr::new(cell_ref.row, cell_ref.col));
            ser.omit_equals = true;

            let display = match group.ast.to_string(ser) {
                Ok(s) => s,
                Err(_) => continue,
            };

            worksheet.set_formula(cell_ref, Some(display));
        }
    }
}

fn parse_xml_bool(val: &str) -> bool {
    val == "1" || val.eq_ignore_ascii_case("true")
}

fn parse_table_part_ids(xml: &[u8]) -> Result<Vec<String>, ReadError> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut out = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e)
                if crate::openxml::local_name(e.name().as_ref()) == b"tablePart" =>
            {
                for attr in e.attributes() {
                    let attr = attr?;
                    if crate::openxml::local_name(attr.key.as_ref()) == b"id" {
                        out.push(attr.unescape_value()?.into_owned());
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

fn parse_inline_is_text<R: std::io::BufRead>(reader: &mut Reader<R>) -> Result<String, ReadError> {
    let mut buf = Vec::new();
    let mut out = String::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if e.local_name().as_ref() == b"t" => {
                out.push_str(&read_text(reader, b"t")?);
            }
            Event::Start(e) if e.local_name().as_ref() == b"r" => {
                out.push_str(&parse_inline_r_text(reader)?);
            }
            Event::Start(e) => {
                reader.read_to_end_into(e.name(), &mut Vec::new())?;
            }
            Event::End(e) if e.local_name().as_ref() == b"is" => break,
            Event::Eof => {
                return Err(ReadError::Xlsx(XlsxError::Invalid(
                    "unexpected EOF while parsing inline string <is>".to_string(),
                )))
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(out)
}

fn parse_inline_r_text<R: std::io::BufRead>(reader: &mut Reader<R>) -> Result<String, ReadError> {
    let mut buf = Vec::new();
    let mut out = String::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if e.local_name().as_ref() == b"t" => {
                out.push_str(&read_text(reader, b"t")?);
            }
            Event::Start(e) => {
                reader.read_to_end_into(e.name(), &mut Vec::new())?;
            }
            Event::End(e) if e.local_name().as_ref() == b"r" => break,
            Event::Eof => {
                return Err(ReadError::Xlsx(XlsxError::Invalid(
                    "unexpected EOF while parsing inline string <r>".to_string(),
                )))
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(out)
}

fn read_text<R: std::io::BufRead>(
    reader: &mut Reader<R>,
    end_local: &[u8],
) -> Result<String, ReadError> {
    let mut buf = Vec::new();
    let mut text = String::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Text(e) => text.push_str(&e.unescape()?.into_owned()),
            Event::CData(e) => text.push_str(std::str::from_utf8(e.as_ref())?),
            Event::End(e) if e.local_name().as_ref() == end_local => break,
            Event::Eof => {
                return Err(ReadError::Xlsx(XlsxError::Invalid(
                    "unexpected EOF while parsing inline string <t>".to_string(),
                )))
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(text)
}

fn interpret_cell_value(
    t: Option<&str>,
    v_text: &Option<String>,
    inline_text: &Option<String>,
    shared_strings: &[RichText],
) -> (CellValue, Option<CellValueKind>, Option<String>) {
    match t {
        Some("s") => {
            let raw = v_text.clone().unwrap_or_default();
            let idx: u32 = raw.parse().unwrap_or(0);
            let text = shared_strings
                .get(idx as usize)
                .map(|rt| rt.text.clone())
                .unwrap_or_default();
            (
                CellValue::String(text),
                Some(CellValueKind::SharedString { index: idx }),
                Some(raw),
            )
        }
        Some("b") => {
            let raw = v_text.clone().unwrap_or_default();
            (
                CellValue::Boolean(raw == "1"),
                Some(CellValueKind::Bool),
                Some(raw),
            )
        }
        Some("e") => {
            let raw = v_text.clone().unwrap_or_default();
            let err = raw.parse::<ErrorValue>().unwrap_or(ErrorValue::Unknown);
            (CellValue::Error(err), Some(CellValueKind::Error), Some(raw))
        }
        Some("str") => {
            let raw = v_text.clone().unwrap_or_default();
            (
                CellValue::String(raw.clone()),
                Some(CellValueKind::Str),
                Some(raw),
            )
        }
        Some("inlineStr") => {
            let raw = inline_text.clone().unwrap_or_default();
            (
                CellValue::String(raw.clone()),
                Some(CellValueKind::InlineString),
                Some(raw),
            )
        }
        Some("n") | None => {
            if let Some(raw) = v_text.clone() {
                if let Ok(n) = raw.parse::<f64>() {
                    (CellValue::Number(n), Some(CellValueKind::Number), Some(raw))
                } else {
                    // A missing/number cell type with a non-numeric payload is invalid SpreadsheetML.
                    // Preserve as a plain string so we don't accidentally emit a numeric cell on write.
                    (
                        CellValue::String(raw.clone()),
                        Some(CellValueKind::Str),
                        Some(raw),
                    )
                }
            } else {
                (CellValue::Empty, None, None)
            }
        }
        Some(other) => {
            // Preserve unknown/less-common `t=` values (e.g. `t="d"` for ISO-8601 date text).
            if let Some(raw) = v_text.clone() {
                (
                    CellValue::String(raw.clone()),
                    Some(CellValueKind::Other {
                        t: other.to_string(),
                    }),
                    Some(raw),
                )
            } else {
                (
                    CellValue::Empty,
                    Some(CellValueKind::Other {
                        t: other.to_string(),
                    }),
                    None,
                )
            }
        }
    }
}

fn interpret_cell_value_without_meta(
    t: Option<&str>,
    v_text: &Option<String>,
    inline_text: &Option<String>,
    shared_strings: &[RichText],
) -> CellValue {
    match t {
        Some("s") => {
            let raw = v_text.as_deref().unwrap_or_default();
            let idx: u32 = raw.parse().unwrap_or(0);
            let text = shared_strings
                .get(idx as usize)
                .map(|rt| rt.text.clone())
                .unwrap_or_default();
            CellValue::String(text)
        }
        Some("b") => CellValue::Boolean(v_text.as_deref() == Some("1")),
        Some("e") => {
            let raw = v_text.as_deref().unwrap_or_default();
            let err = raw.parse::<ErrorValue>().unwrap_or(ErrorValue::Unknown);
            CellValue::Error(err)
        }
        Some("str") => CellValue::String(v_text.clone().unwrap_or_default()),
        Some("inlineStr") => CellValue::String(inline_text.clone().unwrap_or_default()),
        Some("n") | None => {
            if let Some(raw) = v_text.as_deref() {
                raw.parse::<f64>()
                    .map(CellValue::Number)
                    .unwrap_or_else(|_| CellValue::String(raw.to_string()))
            } else {
                CellValue::Empty
            }
        }
        Some(_) => {
            if let Some(raw) = v_text.as_deref() {
                CellValue::String(raw.to_string())
            } else {
                CellValue::Empty
            }
        }
    }
}
