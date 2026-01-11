use crate::tables::parse_table;
use crate::tables::TablePart;
use crate::tables::TABLE_REL_TYPE;
use crate::XlsxError;
use formula_model::{normalize_formula_text, Cell, CellRef, CellValue, Workbook, Worksheet};
use quick_xml::events::Event;
use quick_xml::Reader;
use std::collections::HashMap;
use std::fs::File;
use std::io::{Read, Seek};
use std::path::Path;
use zip::ZipArchive;

pub fn read_workbook(path: impl AsRef<Path>) -> Result<Workbook, XlsxError> {
    let file = File::open(path)?;
    read_workbook_from_reader(file)
}

pub fn read_workbook_from_reader<R: Read + Seek>(reader: R) -> Result<Workbook, XlsxError> {
    let mut archive = ZipArchive::new(reader)?;

    let workbook_xml = read_zip_string(&mut archive, "xl/workbook.xml")?;
    let workbook_rels = read_zip_string(&mut archive, "xl/_rels/workbook.xml.rels")?;

    let shared_strings = match read_zip_string(&mut archive, "xl/sharedStrings.xml") {
        Ok(xml) => parse_shared_strings(&xml)?,
        Err(XlsxError::MissingPart(_)) => Vec::new(),
        Err(err) => return Err(err),
    };

    let rels = parse_relationships(&workbook_rels)?;
    let sheets = crate::parse_workbook_sheets(&workbook_xml)?;

    let mut workbook = Workbook::new();
    for sheet in sheets {
        let rel = rels
            .get(&sheet.rel_id)
            .ok_or_else(|| XlsxError::Invalid(format!("missing relationship for {}", sheet.rel_id)))?;
        let sheet_path = workbook_part_name_for_target(&rel.target);
        let sheet_xml = read_zip_string(&mut archive, &sheet_path)?;

        let sheet_rels_path = sheet_path
            .rsplit_once('/')
            .map(|(dir, file)| format!("{dir}/_rels/{file}.rels"))
            .ok_or_else(|| XlsxError::Invalid(format!("unexpected worksheet path {sheet_path}")))?;
        let sheet_rels = read_zip_string(&mut archive, &sheet_rels_path).ok();
        let sheet_rels_map = if let Some(xml) = sheet_rels {
            parse_relationships(&xml)?
        } else {
            HashMap::new()
        };

        let sheet_id = workbook
            .add_sheet(sheet.name.clone())
            .map_err(|err| XlsxError::Invalid(format!("invalid worksheet name: {err}")))?;
        let sheet_model = workbook
            .sheet_mut(sheet_id)
            .ok_or_else(|| XlsxError::Invalid("failed to create worksheet".into()))?;
        sheet_model.xlsx_sheet_id = Some(sheet.sheet_id);
        sheet_model.xlsx_rel_id = Some(sheet.rel_id.clone());
        sheet_model.visibility = sheet.visibility;
        let table_parts = parse_sheet(&sheet_xml, &shared_strings, sheet_model)?;

        for part in table_parts {
            let rel = sheet_rels_map.get(&part.r_id).ok_or_else(|| {
                XlsxError::Invalid(format!(
                    "missing worksheet relationship {} for table part",
                    part.r_id
                ))
            })?;
            if rel.rel_type != TABLE_REL_TYPE {
                continue;
            }
            let target = normalize_relationship_target(&sheet_path, &rel.target)?;
            let table_xml = read_zip_string(&mut archive, &target)?;
            let mut table = parse_table(&table_xml).map_err(XlsxError::Invalid)?;
            table.relationship_id = Some(part.r_id.clone());
            table.part_path = Some(target);
            sheet_model.tables.push(table);
        }
    }

    Ok(workbook)
}

fn workbook_part_name_for_target(target: &str) -> String {
    if target.starts_with('/') {
        target.trim_start_matches('/').to_string()
    } else {
        // workbook.xml is under xl/, so most relationship targets are relative to that folder.
        crate::openxml::resolve_target("xl/workbook.xml", target)
    }
}

fn read_zip_string<R: Read + Seek>(archive: &mut ZipArchive<R>, path: &str) -> Result<String, XlsxError> {
    let mut file = archive.by_name(path).map_err(|e| {
        if matches!(e, zip::result::ZipError::FileNotFound) {
            XlsxError::MissingPart(path.to_string())
        } else {
            XlsxError::Zip(e)
        }
    })?;
    let mut buf = Vec::new();
    file.read_to_end(&mut buf)?;
    Ok(String::from_utf8(buf)?)
}

#[derive(Debug)]
struct Relationship {
    rel_type: String,
    target: String,
}

fn parse_relationships(xml: &str) -> Result<HashMap<String, Relationship>, XlsxError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut rels = HashMap::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == b"Relationship" => {
                let mut id = None;
                let mut rel_type = None;
                let mut target = None;
                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"Id" => id = Some(attr.unescape_value()?.into_owned()),
                        b"Type" => rel_type = Some(attr.unescape_value()?.into_owned()),
                        b"Target" => target = Some(attr.unescape_value()?.into_owned()),
                        _ => {}
                    }
                }
                let id = id.ok_or_else(|| XlsxError::Invalid("Relationship missing Id".into()))?;
                rels.insert(
                    id,
                    Relationship {
                        rel_type: rel_type.unwrap_or_default(),
                        target: target.unwrap_or_default(),
                    },
                );
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(rels)
}

fn parse_shared_strings(xml: &str) -> Result<Vec<String>, XlsxError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut strings = Vec::new();
    let mut in_t = false;
    let mut current = String::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if e.name().as_ref() == b"t" => {
                in_t = true;
                current.clear();
            }
            Event::Text(t) if in_t => {
                current.push_str(&t.unescape()?.into_owned());
            }
            Event::End(e) if e.name().as_ref() == b"t" => {
                in_t = false;
                strings.push(current.clone());
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(strings)
}

fn parse_sheet(
    xml: &str,
    shared_strings: &[String],
    sheet: &mut Worksheet,
) -> Result<Vec<TablePart>, XlsxError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut current_cell_ref: Option<CellRef> = None;
    let mut current_cell_type: Option<String> = None;
    let mut current_formula: Option<String> = None;
    let mut current_value: Option<String> = None;
    let mut current_inline_string: Option<String> = None;

    let mut in_f = false;
    let mut in_v = false;
    let mut in_is_t = false;

    let mut table_parts = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if e.name().as_ref() == b"c" => {
                current_cell_ref = None;
                current_cell_type = None;
                current_formula = None;
                current_value = None;
                current_inline_string = None;
                in_f = false;
                in_v = false;
                in_is_t = false;

                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"r" => {
                            let a1 = attr.unescape_value()?.into_owned();
                            current_cell_ref = Some(
                                CellRef::from_a1(&a1).map_err(|e| XlsxError::Invalid(e.to_string()))?,
                            );
                        }
                        b"t" => current_cell_type = Some(attr.unescape_value()?.into_owned()),
                        _ => {}
                    }
                }
            }
            Event::End(e) if e.name().as_ref() == b"c" => {
                if let Some(cell_ref) = current_cell_ref.take() {
                    let value = match current_cell_type.as_deref() {
                        Some("s") => {
                            let idx: usize = current_value
                                .as_deref()
                                .unwrap_or("0")
                                .parse()
                                .map_err(|_| XlsxError::Invalid("invalid shared string index".into()))?;
                            CellValue::String(shared_strings.get(idx).cloned().unwrap_or_default())
                        }
                        Some("b") => CellValue::Boolean(current_value.as_deref() == Some("1")),
                        Some("str") => CellValue::String(current_value.take().unwrap_or_default()),
                        Some("inlineStr") => {
                            CellValue::String(current_inline_string.take().unwrap_or_default())
                        }
                        _ => {
                            if let Some(v) = current_value.take() {
                                if v.is_empty() {
                                    CellValue::Empty
                                } else {
                                    CellValue::Number(v.parse::<f64>().map_err(|_| {
                                        XlsxError::Invalid(format!("invalid number '{v}'"))
                                    })?)
                                }
                            } else {
                                CellValue::Empty
                            }
                        }
                    };
                    let formula = current_formula.take().and_then(|f| normalize_formula_text(&f));
                    sheet.set_cell(
                        cell_ref,
                        Cell {
                            value,
                            formula,
                            style_id: 0,
                        },
                    );
                }
            }
            Event::Start(e) if e.name().as_ref() == b"f" => {
                in_f = true;
                current_formula = Some(String::new());
            }
            Event::Text(t) if in_f => {
                if let Some(f) = current_formula.as_mut() {
                    f.push_str(&t.unescape()?.into_owned());
                }
            }
            Event::End(e) if e.name().as_ref() == b"f" => {
                in_f = false;
            }
            Event::Start(e) if e.name().as_ref() == b"v" => {
                in_v = true;
                current_value = Some(String::new());
            }
            Event::Text(t) if in_v => {
                if let Some(v) = current_value.as_mut() {
                    v.push_str(&t.unescape()?.into_owned());
                }
            }
            Event::End(e) if e.name().as_ref() == b"v" => {
                in_v = false;
            }
            Event::Start(e) if e.name().as_ref() == b"is" => {
                current_inline_string = Some(String::new());
            }
            Event::Start(e) if e.name().as_ref() == b"t" && current_inline_string.is_some() => {
                in_is_t = true;
            }
            Event::Text(t) if in_is_t => {
                if let Some(s) = current_inline_string.as_mut() {
                    s.push_str(&t.unescape()?.into_owned());
                }
            }
            Event::End(e) if e.name().as_ref() == b"t" => {
                in_is_t = false;
            }
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == b"tablePart" => {
                let mut r_id = None;
                for attr in e.attributes() {
                    let attr = attr?;
                    if attr.key.as_ref() == b"r:id" {
                        r_id = Some(attr.unescape_value()?.into_owned());
                    }
                }
                if let Some(r_id) = r_id {
                    table_parts.push(TablePart { r_id });
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(table_parts)
}

fn normalize_relationship_target(sheet_path: &str, target: &str) -> Result<String, XlsxError> {
    // Relationship targets are relative to the directory of the .rels file (worksheets/_rels).
    // For worksheets, targets to tables are typically "../tables/table1.xml".
    if target.starts_with('/') {
        return Ok(target.trim_start_matches('/').to_string());
    }

    let base_dir = sheet_path
        .rsplit_once('/')
        .map(|(dir, _)| dir)
        .unwrap_or("xl");

    let mut parts: Vec<&str> = base_dir.split('/').collect();
    for segment in target.split('/') {
        match segment {
            "." | "" => {}
            ".." => {
                parts.pop();
            }
            other => parts.push(other),
        }
    }
    Ok(parts.join("/"))
}
