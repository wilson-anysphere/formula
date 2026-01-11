use std::collections::BTreeMap;
use std::fs::File;
use std::io::{Cursor, Read};
use std::path::Path;

use formula_model::rich_text::RichText;
use formula_model::{Cell, CellRef, CellValue, ErrorValue, SheetVisibility, Workbook};
use quick_xml::events::Event;
use quick_xml::events::attributes::AttrError;
use quick_xml::Reader;
use thiserror::Error;
use zip::ZipArchive;

use crate::shared_strings::parse_shared_strings_xml;
use crate::styles::StylesPart;
use crate::{CalcPr, CellMeta, CellValueKind, DateSystem, FormulaMeta, SheetMeta, XlsxDocument, XlsxMeta};

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
    #[error("missing required part: {0}")]
    MissingPart(&'static str),
    #[error("invalid cell reference: {0}")]
    InvalidCellRef(String),
}

pub fn load_from_path(path: impl AsRef<Path>) -> Result<XlsxDocument, ReadError> {
    let mut file = File::open(path)?;
    let mut bytes = Vec::new();
    file.read_to_end(&mut bytes)?;
    load_from_bytes(&bytes)
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

    let shared_strings = if let Some(bytes) = parts.get("xl/sharedStrings.xml") {
        parse_shared_strings(bytes)?
    } else {
        Vec::new()
    };

    let workbook_xml = parts
        .get("xl/workbook.xml")
        .ok_or(ReadError::MissingPart("xl/workbook.xml"))?;
    let workbook_rels = parts
        .get("xl/_rels/workbook.xml.rels")
        .ok_or(ReadError::MissingPart("xl/_rels/workbook.xml.rels"))?;

    let rels_map = parse_relationships(workbook_rels)?;
    let (date_system, calc_pr, sheets) = parse_workbook_metadata(workbook_xml, &rels_map)?;

    let mut workbook = Workbook::new();
    let styles_part = StylesPart::parse_or_default(parts.get("xl/styles.xml").map(|b| b.as_slice()), &mut workbook.styles)?;
    let mut sheet_meta: Vec<SheetMeta> = Vec::with_capacity(sheets.len());
    let mut cell_meta = std::collections::HashMap::new();

    for sheet in sheets {
        let ws_id = workbook.add_sheet(sheet.name.clone());
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

        parse_worksheet_into_model(
            ws,
            ws_id,
            sheet_xml,
            &shared_strings,
            &styles_part,
            &mut cell_meta,
        )?;

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

fn parse_relationships(bytes: &[u8]) -> Result<BTreeMap<String, String>, ReadError> {
    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut map = BTreeMap::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == b"Relationship" => {
                let mut id = None;
                let mut target = None;
                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"Id" => id = Some(attr.unescape_value()?.into_owned()),
                        b"Target" => target = Some(attr.unescape_value()?.into_owned()),
                        _ => {}
                    }
                }
                if let (Some(id), Some(target)) = (id, target) {
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
                let path = if target.starts_with('/') {
                    target.trim_start_matches('/').to_string()
                } else {
                    format!("xl/{target}")
                };
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
    cell_meta_map: &mut std::collections::HashMap<(formula_model::WorksheetId, CellRef), CellMeta>,
) -> Result<(), ReadError> {
    let mut reader = Reader::from_reader(worksheet_xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    let mut in_sheet_data = false;

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
            Event::Start(e) if e.name().as_ref() == b"sheetData" => in_sheet_data = true,
            Event::End(e) if e.name().as_ref() == b"sheetData" => in_sheet_data = false,
            Event::Empty(e) if e.name().as_ref() == b"sheetData" => {
                in_sheet_data = false;
                drop(e);
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
                    if style_id != 0 {
                        let mut cell = Cell::default();
                        cell.style_id = style_id;
                        worksheet.set_cell(cell_ref, cell);
                    }
                }
            }

            Event::End(e) if in_sheet_data && e.name().as_ref() == b"c" => {
                if let Some(cell_ref) = current_ref {
                    let (value, value_kind, raw_value) =
                        interpret_cell_value(current_t.as_deref(), &current_value_text, &current_inline_text, shared_strings);

                     let formula_in_model = current_formula.as_ref().and_then(|f| {
                         (!f.file_text.is_empty())
                             .then(|| crate::formula_text::strip_xlfn_prefixes(&f.file_text))
                     });

                    let mut cell = Cell::default();
                    cell.value = value;
                    cell.formula = formula_in_model;
                    cell.style_id = current_style;

                    if !cell.is_truly_empty() {
                        worksheet.set_cell(cell_ref, cell);
                    }

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

    Ok(())
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
        Some(_) | None => {
            if let Some(raw) = v_text.clone() {
                let value = raw.parse::<f64>().map(CellValue::Number).unwrap_or(CellValue::String(raw.clone()));
                (value, Some(CellValueKind::Number), Some(raw))
            } else {
                (CellValue::Empty, None, None)
            }
        }
    }
}
