use std::collections::{BTreeSet, HashMap};
use std::io::Cursor;

use quick_xml::events::Event;
use quick_xml::Reader;

use crate::openxml::{local_name, parse_relationships};
use crate::{XlsxDocument, XlsxError, XlsxPackage};
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
        pivot_graph_with(self.part_names(), |name| self.part(name))
    }

    /// Resolve the pivot cache parts backing a given pivot table part.
    ///
    /// Returns `Ok(None)` when:
    /// - `pivot_table_part` does not exist in the package,
    /// - the pivot table does not reference a `cacheId`,
    /// - the workbook does not define the corresponding pivot cache, or
    /// - the cache definition / records parts cannot be resolved (including missing parts).
    pub fn pivot_cache_parts_for_pivot_table(
        &self,
        pivot_table_part: &str,
    ) -> Result<Option<(String, String)>, XlsxError> {
        let Some(pivot_xml) = self.part(pivot_table_part) else {
            return Ok(None);
        };

        let Some(cache_id) = parse_pivot_table_cache_id(pivot_xml).ok().flatten() else {
            return Ok(None);
        };

        let part_names: Vec<String> = self.part_names().map(str::to_string).collect();
        let part = |name: &str| self.part(name);
        let cache_parts_by_id = cache_parts_by_id(&part, &part_names)?;
        let Some(parts) = cache_parts_by_id.get(&cache_id) else {
            return Ok(None);
        };

        let Some(definition_part) = parts.definition_part.clone() else {
            return Ok(None);
        };
        if self.part(&definition_part).is_none() {
            return Ok(None);
        }

        let Some(records_part) = parts.records_part.clone() else {
            return Ok(None);
        };
        if self.part(&records_part).is_none() {
            return Ok(None);
        }

        Ok(Some((definition_part, records_part)))
    }
}

impl XlsxDocument {
    /// Resolve the relationship graph between worksheets, pivot tables, and pivot caches from the
    /// preserved parts in an [`XlsxDocument`].
    pub fn pivot_graph(&self) -> Result<XlsxPivotGraph, XlsxError> {
        pivot_graph_with(self.parts().keys(), |name| {
            let name = name.strip_prefix('/').unwrap_or(name);
            self.parts().get(name).map(|bytes| bytes.as_slice())
        })
    }
}

pub(crate) fn pivot_graph_with<'a, PN, Part>(
    part_names: PN,
    part: Part,
) -> Result<XlsxPivotGraph, XlsxError>
where
    PN: IntoIterator,
    PN::Item: AsRef<str>,
    Part: Fn(&str) -> Option<&'a [u8]>,
{
    let part_names: Vec<String> = part_names
        .into_iter()
        .map(|name| name.as_ref().to_string())
        .collect();

    let sheet_name_by_part = sheet_name_by_part(&part)?;
    let cache_parts_by_id = cache_parts_by_id(&part, &part_names)?;

    let mut sheet_parts: BTreeSet<String> = BTreeSet::new();
    sheet_parts.extend(sheet_name_by_part.keys().cloned());
    sheet_parts.extend(
        part_names
            .iter()
            .filter(|name| name.starts_with("xl/worksheets/") && name.ends_with(".xml"))
            .cloned(),
    );

    let mut pivot_tables = Vec::new();
    let mut seen_pivot_parts: BTreeSet<String> = BTreeSet::new();

    for sheet_part in sheet_parts {
        let sheet_name = sheet_name_by_part.get(&sheet_part).cloned();
        let mut pivot_parts = BTreeSet::new();

        let rels_part = rels_for_part(&sheet_part);
        let relationships = match part(&rels_part) {
            Some(xml) => parse_relationships(xml).unwrap_or_default(),
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

        if let Some(sheet_xml) = part(&sheet_part) {
            // Best-effort: tolerate malformed worksheet XML by falling back to the `.rels` scan
            // above.
            if let Ok(rids) = parse_sheet_pivot_table_relationship_ids(sheet_xml) {
                for rid in rids {
                    if let Some(rel) = rel_map.get(rid.as_str()) {
                        pivot_parts.insert(resolve_target(&sheet_part, &rel.target));
                    }
                }
            }
        }

        for pivot_part in pivot_parts {
            seen_pivot_parts.insert(pivot_part.clone());
            pivot_tables.push(build_pivot_table_instance(
                &part,
                pivot_part,
                Some(sheet_part.clone()),
                sheet_name.clone(),
                &cache_parts_by_id,
            )?);
        }
    }

    // Fallback: include any pivot table parts that exist in the package but weren't reachable from
    // worksheet relationships.
    for pivot_part in part_names
        .iter()
        .filter(|name| name.starts_with("xl/pivotTables/") && name.ends_with(".xml"))
    {
        if seen_pivot_parts.contains(pivot_part.as_str()) {
            continue;
        }
        pivot_tables.push(build_pivot_table_instance(
            &part,
            pivot_part.to_string(),
            None,
            None,
            &cache_parts_by_id,
        )?);
    }

    Ok(XlsxPivotGraph { pivot_tables })
}

fn build_pivot_table_instance<'a>(
    part: &impl Fn(&str) -> Option<&'a [u8]>,
    pivot_part: String,
    sheet_part: Option<String>,
    sheet_name: Option<String>,
    cache_parts_by_id: &HashMap<u32, CacheParts>,
) -> Result<PivotTableInstance, XlsxError> {
    let cache_id = match part(&pivot_part) {
        Some(xml) => parse_pivot_table_cache_id(xml).ok().flatten(),
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

fn sheet_name_by_part<'a>(
    part: &impl Fn(&str) -> Option<&'a [u8]>,
) -> Result<HashMap<String, String>, XlsxError> {
    let workbook_part = "xl/workbook.xml";
    let workbook_xml = match part(workbook_part) {
        Some(bytes) => bytes,
        None => return Ok(HashMap::new()),
    };
    let workbook_xml = match String::from_utf8(workbook_xml.to_vec()) {
        Ok(xml) => xml,
        Err(_) => return Ok(HashMap::new()),
    };
    let sheets = match parse_workbook_sheets(&workbook_xml) {
        Ok(sheets) => sheets,
        Err(_) => return Ok(HashMap::new()),
    };

    let rels_part = rels_for_part(workbook_part);
    let workbook_rels = match part(&rels_part) {
        Some(bytes) => parse_relationships(bytes).unwrap_or_default(),
        None => Vec::new(),
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

fn cache_parts_by_id<'a>(
    part: &impl Fn(&str) -> Option<&'a [u8]>,
    part_names: &[String],
) -> Result<HashMap<u32, CacheParts>, XlsxError> {
    let workbook_part = "xl/workbook.xml";
    let cache_refs = match part(workbook_part) {
        Some(bytes) => parse_workbook_pivot_caches(bytes).unwrap_or_default(),
        None => Vec::new(),
    };

    let rels_part = rels_for_part(workbook_part);
    let workbook_rels = match part(&rels_part) {
        Some(bytes) => parse_relationships(bytes).unwrap_or_default(),
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
            Some(def_part) => cache_records_part(part, def_part)?,
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

    // Fallback for malformed workbooks: if `workbook.xml` omits `<pivotCaches>` (or the workbook
    // `.rels` is incomplete), attempt to resolve cache parts using the common `...DefinitionN.xml`
    // / `...RecordsN.xml` naming convention where `N` matches `cacheId`.
    for part_name in part_names
        .iter()
        .map(String::as_str)
        .filter(|name| name.starts_with("xl/pivotCache/") && name.ends_with(".xml"))
    {
        let Some(cache_id) = cache_id_from_pivot_cache_definition_part(part_name) else {
            continue;
        };
        caches.entry(cache_id).or_insert_with(|| CacheParts {
            definition_part: Some(part_name.to_string()),
            records_part: {
                let guess = format!("xl/pivotCache/pivotCacheRecords{cache_id}.xml");
                if part(&guess).is_some() {
                    Some(guess)
                } else {
                    None
                }
            },
        });
    }

    for (cache_id, parts) in caches.iter_mut() {
        if parts.definition_part.is_none() {
            let guess = format!("xl/pivotCache/pivotCacheDefinition{cache_id}.xml");
            if part(&guess).is_some() {
                parts.definition_part = Some(guess);
            }
        }

        if parts.records_part.is_none() {
            let guess = format!("xl/pivotCache/pivotCacheRecords{cache_id}.xml");
            if part(&guess).is_some() {
                parts.records_part = Some(guess);
            }
        }
    }

    Ok(caches)
}

fn cache_id_from_pivot_cache_definition_part(part_name: &str) -> Option<u32> {
    let file = part_name.rsplit('/').next()?;
    let digits = file
        .strip_prefix("pivotCacheDefinition")?
        .strip_suffix(".xml")?;
    digits.parse::<u32>().ok()
}

fn cache_records_part<'a>(
    part: &impl Fn(&str) -> Option<&'a [u8]>,
    cache_definition_part: &str,
) -> Result<Option<String>, XlsxError> {
    let rels_part = rels_for_part(cache_definition_part);
    let rels_xml = match part(&rels_part) {
        Some(bytes) => bytes,
        None => return Ok(None),
    };

    let relationships = parse_relationships(rels_xml).unwrap_or_default();
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
