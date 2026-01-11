use crate::tables::{parse_table, TABLE_REL_TYPE};
use crate::{XlsxDocument, XlsxError};
use formula_model::Workbook;
use quick_xml::events::Event;
use quick_xml::Reader;
use std::collections::{HashMap, HashSet};
use std::fs::File;
use std::io::{Read, Seek, SeekFrom};
use std::path::Path;

pub fn read_workbook(path: impl AsRef<Path>) -> Result<Workbook, XlsxError> {
    let file = File::open(path)?;
    read_workbook_from_reader(file)
}

pub fn read_workbook_from_reader<R: Read + Seek>(reader: R) -> Result<Workbook, XlsxError> {
    read_workbook_model_from_reader(reader)
}

fn read_workbook_model_from_reader<R: Read + Seek>(mut reader: R) -> Result<Workbook, XlsxError> {
    // Ensure we read from the start; callers may pass a reused reader.
    reader.seek(SeekFrom::Start(0))?;
    let mut bytes = Vec::new();
    reader.read_to_end(&mut bytes)?;
    read_workbook_model_from_bytes(&bytes)
}

fn read_workbook_model_from_bytes(bytes: &[u8]) -> Result<Workbook, XlsxError> {
    let mut doc = crate::load_from_bytes(bytes).map_err(read_error_to_xlsx_error)?;
    attach_tables(&mut doc)?;
    Ok(doc.workbook)
}

fn attach_tables(doc: &mut XlsxDocument) -> Result<(), XlsxError> {
    for sheet_meta in &doc.meta.sheets {
        let Some(sheet) = doc.workbook.sheet_mut(sheet_meta.worksheet_id) else {
            continue;
        };

        let worksheet_xml = doc
            .parts
            .get(&sheet_meta.path)
            .ok_or_else(|| XlsxError::MissingPart(sheet_meta.path.clone()))?;

        let table_part_ids = parse_table_part_ids(worksheet_xml)?;
        if table_part_ids.is_empty() {
            continue;
        }

        let rels_part = crate::path::rels_for_part(&sheet_meta.path);
        let rels_xml = doc
            .parts
            .get(&rels_part)
            .ok_or_else(|| XlsxError::MissingPart(rels_part.clone()))?;
        let relationships = crate::openxml::parse_relationships(rels_xml)?;

        let mut rels_by_id: HashMap<String, crate::openxml::Relationship> =
            HashMap::with_capacity(relationships.len());
        for rel in relationships {
            rels_by_id.insert(rel.id.clone(), rel);
        }

        let mut seen_rel_ids: HashSet<String> = sheet
            .tables
            .iter()
            .filter_map(|t| t.relationship_id.clone())
            .collect();

        for r_id in table_part_ids {
            // Avoid duplicates if the underlying fast reader starts populating tables.
            if !seen_rel_ids.insert(r_id.clone()) {
                continue;
            }

            let rel = rels_by_id.get(&r_id).ok_or_else(|| {
                XlsxError::Invalid(format!(
                    "missing worksheet relationship {r_id} for table part"
                ))
            })?;
            if rel.type_uri != TABLE_REL_TYPE {
                continue;
            }

            let target = crate::path::resolve_target(&sheet_meta.path, &rel.target);
            let table_bytes = doc
                .parts
                .get(&target)
                .ok_or_else(|| XlsxError::MissingPart(target.clone()))?;

            let table_xml = std::str::from_utf8(table_bytes)
                .map_err(|e| XlsxError::Invalid(format!("invalid utf-8 in {target}: {e}")))?;

            let mut table = parse_table(table_xml).map_err(XlsxError::Invalid)?;
            table.relationship_id = Some(r_id);
            table.part_path = Some(target);
            sheet.tables.push(table);
        }
    }
    Ok(())
}

fn parse_table_part_ids(xml: &[u8]) -> Result<Vec<String>, XlsxError> {
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

fn read_error_to_xlsx_error(err: crate::read::ReadError) -> XlsxError {
    match err {
        crate::read::ReadError::Io(err) => XlsxError::Io(err),
        crate::read::ReadError::Zip(err) => XlsxError::Zip(err),
        crate::read::ReadError::Xml(err) => XlsxError::Xml(err),
        crate::read::ReadError::XmlAttr(err) => XlsxError::Attr(err),
        crate::read::ReadError::Utf8(err) => XlsxError::Invalid(format!("utf-8 error: {err}")),
        crate::read::ReadError::SharedStrings(err) => {
            XlsxError::Invalid(format!("shared strings error: {err}"))
        }
        crate::read::ReadError::Styles(err) => XlsxError::Invalid(format!("styles error: {err}")),
        crate::read::ReadError::InvalidSheetName(err) => {
            XlsxError::Invalid(format!("invalid worksheet name: {err}"))
        }
        crate::read::ReadError::Xlsx(err) => err,
        crate::read::ReadError::MissingPart(part) => XlsxError::MissingPart(part.to_string()),
        crate::read::ReadError::InvalidCellRef(a1) => {
            XlsxError::Invalid(format!("invalid cell reference: {a1}"))
        }
        crate::read::ReadError::InvalidRangeRef(range) => {
            XlsxError::Invalid(format!("invalid range reference: {range}"))
        }
    }
}
