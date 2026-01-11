use std::collections::{BTreeMap, HashMap};
use std::fs::File;
use std::io::{Cursor, Read};
use std::path::Path;

use formula_engine::{parse_formula, CellAddr, ParseOptions, SerializeOptions};
use formula_model::rich_text::RichText;
use formula_model::{
    normalize_formula_text, Cell, CellRef, CellValue, ErrorValue, Range, SheetVisibility, Workbook,
};
use quick_xml::events::Event;
use quick_xml::events::attributes::AttrError;
use quick_xml::Reader;
use thiserror::Error;
use zip::ZipArchive;

use crate::path::{rels_for_part, resolve_target};
use crate::shared_strings::parse_shared_strings_xml;
use crate::sheet_metadata::parse_sheet_tab_color;
use crate::styles::StylesPart;
use crate::{parse_worksheet_hyperlinks, XlsxError};
use crate::{CalcPr, CellMeta, CellValueKind, DateSystem, FormulaMeta, SheetMeta, XlsxDocument, XlsxMeta};

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
    let cursor = Cursor::new(bytes);
    let mut archive = ZipArchive::new(cursor)?;

    let workbook_xml = read_zip_part_required(&mut archive, WORKBOOK_PART)?;
    let workbook_rels = read_zip_part_required(&mut archive, WORKBOOK_RELS_PART)?;

    let rels_info = parse_relationships(&workbook_rels)?;
    let (_date_system, _calc_pr, sheets) =
        parse_workbook_metadata(&workbook_xml, &rels_info.id_to_target)?;

    let mut workbook = Workbook::new();

    let styles_part_name = rels_info
        .styles_target
        .as_deref()
        .map(|target| resolve_target(WORKBOOK_PART, target))
        .unwrap_or_else(|| "xl/styles.xml".to_string());
    let styles_bytes = read_zip_part_optional(&mut archive, &styles_part_name)?;
    let styles_part =
        StylesPart::parse_or_default(styles_bytes.as_deref(), &mut workbook.styles)?;

    let shared_strings_part_name = rels_info
        .shared_strings_target
        .as_deref()
        .map(|target| resolve_target(WORKBOOK_PART, target))
        .unwrap_or_else(|| "xl/sharedStrings.xml".to_string());
    let shared_strings = match read_zip_part_optional(&mut archive, &shared_strings_part_name)? {
        Some(bytes) => parse_shared_strings(&bytes)?,
        None => Vec::new(),
    };

    for sheet in sheets {
        let ws_id = workbook.add_sheet(sheet.name.clone())?;
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

        let sheet_xml = read_zip_part_optional(&mut archive, &sheet.path)?
            .ok_or(ReadError::MissingPart("worksheet part referenced from workbook.xml.rels"))?;

        // Worksheet-level metadata lives inside the worksheet part (and sometimes its .rels).
        let sheet_xml_str = std::str::from_utf8(&sheet_xml)?;

        ws.tab_color = parse_sheet_tab_color(sheet_xml_str)?;

        // Merged cells (must be parsed before cell content so we don't treat interior
        // cells as value-bearing).
        let merges =
            crate::merge_cells::read_merge_cells_from_worksheet_xml(sheet_xml_str).map_err(
                |err| match err {
                    crate::merge_cells::MergeCellsError::Xml(e) => ReadError::Xml(e),
                    crate::merge_cells::MergeCellsError::Attr(e) => ReadError::XmlAttr(e),
                    crate::merge_cells::MergeCellsError::Utf8(e) => ReadError::Utf8(e),
                    crate::merge_cells::MergeCellsError::InvalidRef(r) => {
                        ReadError::InvalidRangeRef(r)
                    }
                    crate::merge_cells::MergeCellsError::Zip(e) => ReadError::Zip(e),
                    crate::merge_cells::MergeCellsError::Io(e) => ReadError::Io(e),
                },
            )?;
        for range in merges {
            ws.merged_regions
                .add(range)
                .map_err(|e| ReadError::InvalidRangeRef(e.to_string()))?;
        }

        // Hyperlinks.
        let rels_part = rels_for_part(&sheet.path);
        let rels_xml_bytes = read_zip_part_optional(&mut archive, &rels_part)?;
        let rels_xml = rels_xml_bytes
            .as_deref()
            .map(|bytes| std::str::from_utf8(bytes))
            .transpose()?;
        ws.hyperlinks = parse_worksheet_hyperlinks(sheet_xml_str, rels_xml)?;

        parse_worksheet_into_model(
            ws,
            ws_id,
            &sheet_xml,
            &shared_strings,
            &styles_part,
            None,
        )?;
    }

    Ok(workbook)
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
    let (date_system, calc_pr, sheets) = parse_workbook_metadata(workbook_xml, &rels_info.id_to_target)?;

    let mut workbook = Workbook::new();
    let styles_part_name = rels_info
        .styles_target
        .as_deref()
        .map(|target| resolve_target(WORKBOOK_PART, target))
        .unwrap_or_else(|| "xl/styles.xml".to_string());
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

    for sheet in sheets {
        let ws_id = workbook.add_sheet(sheet.name.clone())?;
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

        let sheet_xml = parts
            .get(&sheet.path)
            .ok_or(ReadError::MissingPart("worksheet part referenced from workbook.xml.rels"))?;

        // Worksheet-level metadata lives inside the worksheet part (and sometimes its .rels).
        let sheet_xml_str = std::str::from_utf8(sheet_xml)?;

        ws.tab_color = parse_sheet_tab_color(sheet_xml_str)?;

        // Merged cells (must be parsed before cell content so we don't treat interior
        // cells as value-bearing).
        let merges =
            crate::merge_cells::read_merge_cells_from_worksheet_xml(sheet_xml_str).map_err(
                |err| match err {
                    crate::merge_cells::MergeCellsError::Xml(e) => ReadError::Xml(e),
                    crate::merge_cells::MergeCellsError::Attr(e) => ReadError::XmlAttr(e),
                    crate::merge_cells::MergeCellsError::Utf8(e) => ReadError::Utf8(e),
                    crate::merge_cells::MergeCellsError::InvalidRef(r) => {
                        ReadError::InvalidRangeRef(r)
                    }
                    crate::merge_cells::MergeCellsError::Zip(e) => ReadError::Zip(e),
                    crate::merge_cells::MergeCellsError::Io(e) => ReadError::Io(e),
                },
            )?;
        for range in merges {
            ws.merged_regions
                .add(range)
                .map_err(|e| ReadError::InvalidRangeRef(e.to_string()))?;
        }

        // Hyperlinks.
        let rels_part = rels_for_part(&sheet.path);
        let rels_xml = parts
            .get(&rels_part)
            .map(|bytes| std::str::from_utf8(bytes))
            .transpose()?;
        ws.hyperlinks = parse_worksheet_hyperlinks(sheet_xml_str, rels_xml)?;

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

fn parse_workbook_metadata(
    workbook_xml: &[u8],
    rels: &BTreeMap<String, String>,
) -> Result<(DateSystem, CalcPr, Vec<ParsedSheet>), ReadError> {
    let mut reader = Reader::from_reader(workbook_xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    let mut date_system = DateSystem::V1900;
    let mut calc_pr = CalcPr::default();
    let mut sheets = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == b"workbookPr" => {
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
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == b"calcPr" => {
                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"calcId" => calc_pr.calc_id = Some(attr.unescape_value()?.into_owned()),
                        b"calcMode" => calc_pr.calc_mode = Some(attr.unescape_value()?.into_owned()),
                        b"fullCalcOnLoad" => {
                            let v = attr.unescape_value()?.into_owned();
                            calc_pr.full_calc_on_load =
                                Some(v == "1" || v.eq_ignore_ascii_case("true"))
                        }
                        _ => {}
                    }
                }
            }
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == b"sheet" => {
                let mut name = None;
                let mut sheet_id = None;
                let mut r_id = None;
                let mut state = None;
                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"name" => name = Some(attr.unescape_value()?.into_owned()),
                        b"sheetId" => {
                            sheet_id =
                                Some(attr.unescape_value()?.into_owned().parse().unwrap_or(0))
                        }
                        b"r:id" => r_id = Some(attr.unescape_value()?.into_owned()),
                        b"state" => state = Some(attr.unescape_value()?.into_owned()),
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
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok((date_system, calc_pr, sheets))
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
    let mut in_inline_t = false;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if e.name().as_ref() == b"cols" => in_cols = true,
            Event::End(e) if e.name().as_ref() == b"cols" => in_cols = false,
            Event::Empty(e) if e.name().as_ref() == b"cols" => {
                in_cols = false;
                drop(e);
            }
            Event::Start(e) | Event::Empty(e) if in_cols && e.name().as_ref() == b"col" => {
                let mut min: Option<u32> = None;
                let mut max: Option<u32> = None;
                let mut width: Option<f32> = None;
                let mut hidden = false;

                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"min" => {
                            min = Some(
                                attr.unescape_value()?
                                    .into_owned()
                                    .parse()
                                    .unwrap_or(0),
                            )
                        }
                        b"max" => {
                            max = Some(
                                attr.unescape_value()?
                                    .into_owned()
                                    .parse()
                                    .unwrap_or(0),
                            )
                        }
                        b"width" => {
                            width = attr
                                .unescape_value()?
                                .into_owned()
                                .parse::<f32>()
                                .ok();
                        }
                        b"hidden" => {
                            let v = attr.unescape_value()?.into_owned();
                            hidden = v == "1" || v.eq_ignore_ascii_case("true");
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
                    if let Some(width) = width {
                        worksheet.set_col_width(col, Some(width));
                    }
                    if hidden {
                        worksheet.set_col_hidden(col, true);
                    }
                }
            }

            Event::Start(e) if e.name().as_ref() == b"sheetData" => in_sheet_data = true,
            Event::End(e) if e.name().as_ref() == b"sheetData" => in_sheet_data = false,
            Event::Empty(e) if e.name().as_ref() == b"sheetData" => {
                in_sheet_data = false;
                drop(e);
            }

            Event::Start(e) | Event::Empty(e) if in_sheet_data && e.name().as_ref() == b"row" => {
                let mut row_1_based: Option<u32> = None;
                let mut height: Option<f32> = None;
                let mut hidden = false;

                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"r" => {
                            row_1_based = Some(
                                attr.unescape_value()?
                                    .into_owned()
                                    .parse()
                                    .unwrap_or(0),
                            );
                        }
                        b"ht" => {
                            height = attr
                                .unescape_value()?
                                .into_owned()
                                .parse::<f32>()
                                .ok();
                        }
                        b"hidden" => {
                            let v = attr.unescape_value()?.into_owned();
                            hidden = v == "1" || v.eq_ignore_ascii_case("true");
                        }
                        _ => {}
                    }
                }

                if let Some(row_1_based) = row_1_based {
                    if row_1_based > 0 && row_1_based <= formula_model::EXCEL_MAX_ROWS {
                        let row = row_1_based - 1;
                        if let Some(height) = height {
                            worksheet.set_row_height(row, Some(height));
                        }
                        if hidden {
                            worksheet.set_row_hidden(row, true);
                        }
                    }
                }
            }

            Event::Start(e) if in_sheet_data && e.name().as_ref() == b"c" => {
                current_ref = None;
                current_t = None;
                current_style = 0;
                current_formula = None;
                current_value_text = None;
                current_inline_text = None;
                in_v = false;
                in_f = false;
                in_inline_t = false;

                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"r" => {
                            let a1 = attr.unescape_value()?.into_owned();
                            current_ref = Some(
                                CellRef::from_a1(&a1)
                                    .map_err(|_| ReadError::InvalidCellRef(a1))?,
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
            Event::Empty(e) if in_sheet_data && e.name().as_ref() == b"c" => {
                let mut cell_ref = None;
                let mut style_id = 0u32;
                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"r" => {
                            let a1 = attr.unescape_value()?.into_owned();
                            cell_ref = Some(
                                CellRef::from_a1(&a1)
                                    .map_err(|_| ReadError::InvalidCellRef(a1))?,
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
                    if worksheet.merged_regions.resolve_cell(cell_ref) == cell_ref && style_id != 0 {
                        let mut cell = Cell::default();
                        cell.style_id = style_id;
                        worksheet.set_cell(cell_ref, cell);
                    }
                }
            }

            Event::End(e) if in_sheet_data && e.name().as_ref() == b"c" => {
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
                            let is_shared_master = current_formula.as_ref().is_some_and(|formula| {
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
                in_inline_t = false;
            }

            Event::Start(e) if in_sheet_data && current_ref.is_some() && e.name().as_ref() == b"v" => {
                in_v = true;
            }
            Event::End(e) if in_sheet_data && e.name().as_ref() == b"v" => in_v = false,
            Event::Text(e) if in_sheet_data && in_v => {
                current_value_text = Some(e.unescape()?.into_owned());
            }

            Event::Start(e) if in_sheet_data && current_ref.is_some() && e.name().as_ref() == b"f" => {
                in_f = true;
                let mut formula = FormulaMeta::default();
                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"t" => formula.t = Some(attr.unescape_value()?.into_owned()),
                        b"ref" => formula.reference = Some(attr.unescape_value()?.into_owned()),
                        b"si" => formula.shared_index =
                            Some(attr.unescape_value()?.into_owned().parse().unwrap_or(0)),
                        b"aca" => {
                            let v = attr.unescape_value()?.into_owned();
                            formula.always_calc = Some(v == "1" || v.eq_ignore_ascii_case("true"))
                        }
                        _ => {}
                    }
                }
                current_formula = Some(formula);
            }
            Event::Empty(e) if in_sheet_data && current_ref.is_some() && e.name().as_ref() == b"f" => {
                let mut formula = FormulaMeta::default();
                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"t" => formula.t = Some(attr.unescape_value()?.into_owned()),
                        b"ref" => formula.reference = Some(attr.unescape_value()?.into_owned()),
                        b"si" => formula.shared_index =
                            Some(attr.unescape_value()?.into_owned().parse().unwrap_or(0)),
                        b"aca" => {
                            let v = attr.unescape_value()?.into_owned();
                            formula.always_calc = Some(v == "1" || v.eq_ignore_ascii_case("true"))
                        }
                        _ => {}
                    }
                }
                current_formula = Some(formula);
            }
            Event::End(e) if in_sheet_data && e.name().as_ref() == b"f" => in_f = false,
            Event::Text(e) if in_sheet_data && in_f => {
                if let Some(formula) = current_formula.as_mut() {
                    formula.file_text = e.unescape()?.into_owned();
                }
            }

            Event::Start(e)
                if in_sheet_data
                    && current_ref.is_some()
                    && current_t.as_deref() == Some("inlineStr")
                    && e.name().as_ref() == b"t" =>
            {
                in_inline_t = true;
            }
            Event::End(e)
                if in_sheet_data
                    && current_t.as_deref() == Some("inlineStr")
                    && e.name().as_ref() == b"t" =>
            {
                in_inline_t = false;
            }
            Event::Text(e) if in_sheet_data && in_inline_t => {
                let t = e.unescape()?.into_owned();
                match current_inline_text.as_mut() {
                    Some(existing) => existing.push_str(&t),
                    None => current_inline_text = Some(t),
                }
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
            (CellValue::Boolean(raw == "1"), Some(CellValueKind::Bool), Some(raw))
        }
        Some("e") => {
            let raw = v_text.clone().unwrap_or_default();
            let err = raw.parse::<ErrorValue>().unwrap_or(ErrorValue::Unknown);
            (CellValue::Error(err), Some(CellValueKind::Error), Some(raw))
        }
        Some("str") => {
            let raw = v_text.clone().unwrap_or_default();
            (CellValue::String(raw.clone()), Some(CellValueKind::Str), Some(raw))
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
                    Some(CellValueKind::Other { t: other.to_string() }),
                    Some(raw),
                )
            } else {
                (
                    CellValue::Empty,
                    Some(CellValueKind::Other { t: other.to_string() }),
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
