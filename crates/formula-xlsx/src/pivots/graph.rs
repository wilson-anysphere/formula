use std::collections::{BTreeSet, HashMap};
use std::io::Cursor;

use quick_xml::events::Event;
use quick_xml::Reader;

use crate::openxml::{local_name, parse_relationships};
use crate::package::{XlsxError, XlsxPackage};
use crate::path::{rels_for_part, resolve_target};
use crate::sheet_metadata::parse_workbook_sheets;

const REL_TYPE_PIVOT_TABLE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";
const REL_TYPE_PIVOT_CACHE_RECORDS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords";

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotTableInstance {
    pub pivot_table_part: String,
    pub sheet_part: Option<String>,
    pub sheet_name: Option<String>,
    pub cache_id: Option<u32>,
    pub cache_definition_part: Option<String>,
    pub cache_records_part: Option<String>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct XlsxPivotGraph {
    pub pivot_tables: Vec<PivotTableInstance>,
}

#[derive(Debug, Clone)]
struct CacheParts {
    definition_part: Option<String>,
    records_part: Option<String>,
}

#[derive(Debug, Clone)]
struct WorkbookPivotCacheRef {
    cache_id: u32,
    rel_id: Option<String>,
}

impl XlsxPackage {
    /// Resolve the relationship graph between worksheets, pivot tables, and pivot caches.
    ///
    /// This helper is intentionally tolerant of missing parts and relationships: pivot tables
    /// are returned even when their sheet or cache relationships cannot be resolved.
    pub fn pivot_graph(&self) -> Result<XlsxPivotGraph, XlsxError> {
        let sheet_name_by_part = sheet_name_by_part(self)?;
        let cache_parts_by_id = cache_parts_by_id(self)?;

        let mut sheet_parts: BTreeSet<String> = BTreeSet::new();
        sheet_parts.extend(sheet_name_by_part.keys().cloned());
        sheet_parts.extend(
            self.part_names()
                .filter(|name| name.starts_with("xl/worksheets/") && name.ends_with(".xml"))
                .map(str::to_string),
        );

        let mut pivot_tables = Vec::new();
        let mut seen_pivot_parts: BTreeSet<String> = BTreeSet::new();

        for sheet_part in sheet_parts {
            let sheet_name = sheet_name_by_part.get(&sheet_part).cloned();
            let mut pivot_parts = BTreeSet::new();

            let rels_part = rels_for_part(&sheet_part);
            let relationships = match self.part(&rels_part) {
                Some(xml) => parse_relationships(xml)?,
                None => Vec::new(),
            };
            let rel_map: HashMap<_, _> = relationships
                .iter()
                .map(|rel| (rel.id.as_str(), rel))
                .collect();

            for rel in &relationships {
                if rel.type_uri == REL_TYPE_PIVOT_TABLE {
                    pivot_parts.insert(resolve_target(&sheet_part, &rel.target));
                }
            }

            if let Some(sheet_xml) = self.part(&sheet_part) {
                for rid in parse_sheet_pivot_table_relationship_ids(sheet_xml)? {
                    if let Some(rel) = rel_map.get(rid.as_str()) {
                        pivot_parts.insert(resolve_target(&sheet_part, &rel.target));
                    }
                }
            }

            for pivot_part in pivot_parts {
                seen_pivot_parts.insert(pivot_part.clone());
                pivot_tables.push(build_pivot_table_instance(
                    self,
                    pivot_part,
                    Some(sheet_part.clone()),
                    sheet_name.clone(),
                    &cache_parts_by_id,
                )?);
            }
        }

        // Fallback: include any pivot table parts that exist in the package but weren't
        // reachable from worksheet relationships.
        for pivot_part in self
            .part_names()
            .filter(|name| name.starts_with("xl/pivotTables/") && name.ends_with(".xml"))
        {
            if seen_pivot_parts.contains(pivot_part) {
                continue;
            }
            pivot_tables.push(build_pivot_table_instance(
                self,
                pivot_part.to_string(),
                None,
                None,
                &cache_parts_by_id,
            )?);
        }

        Ok(XlsxPivotGraph { pivot_tables })
    }
}

fn build_pivot_table_instance(
    package: &XlsxPackage,
    pivot_part: String,
    sheet_part: Option<String>,
    sheet_name: Option<String>,
    cache_parts_by_id: &HashMap<u32, CacheParts>,
) -> Result<PivotTableInstance, XlsxError> {
    let cache_id = match package.part(&pivot_part) {
        Some(xml) => parse_pivot_table_cache_id(xml)?,
        None => None,
    };

    let (cache_definition_part, cache_records_part) = cache_id
        .and_then(|id| cache_parts_by_id.get(&id))
        .map(|parts| (parts.definition_part.clone(), parts.records_part.clone()))
        .unwrap_or((None, None));

    Ok(PivotTableInstance {
        pivot_table_part: pivot_part,
        sheet_part,
        sheet_name,
        cache_id,
        cache_definition_part,
        cache_records_part,
    })
}

fn sheet_name_by_part(package: &XlsxPackage) -> Result<HashMap<String, String>, XlsxError> {
    let workbook_part = "xl/workbook.xml";
    let workbook_xml = match package.part(workbook_part) {
        Some(bytes) => bytes,
        None => return Ok(HashMap::new()),
    };
    let workbook_xml = String::from_utf8(workbook_xml.to_vec())?;
    let sheets = parse_workbook_sheets(&workbook_xml)?;

    let rels_part = rels_for_part(workbook_part);
    let workbook_rels = match package.part(&rels_part) {
        Some(bytes) => parse_relationships(bytes)?,
        None => return Ok(HashMap::new()),
    };
    let rel_map: HashMap<_, _> = workbook_rels
        .iter()
        .map(|rel| (rel.id.as_str(), rel))
        .collect();

    let mut out = HashMap::new();
    for sheet in sheets {
        let Some(rel) = rel_map.get(sheet.rel_id.as_str()) else {
            continue;
        };
        let sheet_part = resolve_target(workbook_part, &rel.target);
        out.insert(sheet_part, sheet.name);
    }

    Ok(out)
}

fn cache_parts_by_id(package: &XlsxPackage) -> Result<HashMap<u32, CacheParts>, XlsxError> {
    let workbook_part = "xl/workbook.xml";
    let workbook_xml = match package.part(workbook_part) {
        Some(bytes) => bytes,
        None => return Ok(HashMap::new()),
    };
    let cache_refs = parse_workbook_pivot_caches(workbook_xml)?;

    let rels_part = rels_for_part(workbook_part);
    let workbook_rels = match package.part(&rels_part) {
        Some(bytes) => parse_relationships(bytes)?,
        None => Vec::new(),
    };
    let rel_map: HashMap<_, _> = workbook_rels
        .iter()
        .map(|rel| (rel.id.as_str(), rel))
        .collect();

    let mut caches = HashMap::new();
    for cache in cache_refs {
        let definition_part = cache
            .rel_id
            .as_deref()
            .and_then(|rid| rel_map.get(rid))
            .map(|rel| resolve_target(workbook_part, &rel.target));

        let records_part = match definition_part.as_deref() {
            Some(def_part) => cache_records_part(package, def_part)?,
            None => None,
        };

        caches.insert(
            cache.cache_id,
            CacheParts {
                definition_part,
                records_part,
            },
        );
    }

    Ok(caches)
}

fn cache_records_part(
    package: &XlsxPackage,
    cache_definition_part: &str,
) -> Result<Option<String>, XlsxError> {
    let rels_part = rels_for_part(cache_definition_part);
    let rels_xml = match package.part(&rels_part) {
        Some(bytes) => bytes,
        None => return Ok(None),
    };

    let relationships = parse_relationships(rels_xml)?;
    for rel in relationships {
        if rel.type_uri == REL_TYPE_PIVOT_CACHE_RECORDS {
            return Ok(Some(resolve_target(cache_definition_part, &rel.target)));
        }
    }

    Ok(None)
}

fn parse_workbook_pivot_caches(xml: &[u8]) -> Result<Vec<WorkbookPivotCacheRef>, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut caches = Vec::new();

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Start(e) | Event::Empty(e) => {
                if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"pivotCache") {
                    let mut cache_id = None;
                    let mut rel_id = None;
                    for attr in e.attributes().with_checks(false) {
                        let attr = attr.map_err(quick_xml::Error::from)?;
                        let key = local_name(attr.key.as_ref());
                        let value = attr.unescape_value()?.into_owned();
                        if key.eq_ignore_ascii_case(b"cacheId") {
                            cache_id = value.parse::<u32>().ok();
                        } else if key.eq_ignore_ascii_case(b"id") {
                            rel_id = Some(value);
                        }
                    }
                    if let Some(cache_id) = cache_id {
                        caches.push(WorkbookPivotCacheRef { cache_id, rel_id });
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(caches)
}

fn parse_sheet_pivot_table_relationship_ids(xml: &[u8]) -> Result<Vec<String>, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut out = Vec::new();

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Start(e) | Event::Empty(e) => {
                if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"pivotTable") {
                    for attr in e.attributes().with_checks(false) {
                        let attr = attr.map_err(quick_xml::Error::from)?;
                        if local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"id") {
                            out.push(attr.unescape_value()?.into_owned());
                        }
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

fn parse_pivot_table_cache_id(xml: &[u8]) -> Result<Option<u32>, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Start(e) | Event::Empty(e) => {
                if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"pivotTableDefinition") {
                    for attr in e.attributes().with_checks(false) {
                        let attr = attr.map_err(quick_xml::Error::from)?;
                        if local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"cacheId") {
                            let value = attr.unescape_value()?.into_owned();
                            return Ok(value.parse::<u32>().ok());
                        }
                    }
                    return Ok(None);
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(None)
}
