use std::collections::{BTreeMap, BTreeSet};
use std::io::Cursor;

use quick_xml::events::Event;
use quick_xml::Reader;

use crate::XlsxError;

pub mod cache_definition;
pub mod pivot_charts;
pub mod slicers;
pub mod graph;

pub use graph::{PivotTableInstance, XlsxPivotGraph};

pub use cache_definition::{PivotCacheDefinition, PivotCacheField, PivotCacheSourceType};

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotTablePart {
    pub path: String,
    pub name: Option<String>,
    pub cache_id: Option<u32>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotCacheDefinitionPart {
    pub path: String,
    pub record_count: Option<u64>,
    pub fields: Vec<String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotCacheRecordsPart {
    pub path: String,
    pub count: Option<u64>,
}

#[derive(Debug, Clone, Default)]
pub struct XlsxPivots {
    pub pivot_tables: Vec<PivotTablePart>,
    pub pivot_cache_definitions: Vec<PivotCacheDefinitionPart>,
    pub pivot_cache_records: Vec<PivotCacheRecordsPart>,
}

impl XlsxPivots {
    pub fn parse_from_entries(entries: &BTreeMap<String, Vec<u8>>) -> Result<Self, XlsxError> {
        let mut pivots = XlsxPivots::default();

        let mut table_paths: BTreeSet<String> = BTreeSet::new();
        let mut cache_def_paths: BTreeSet<String> = BTreeSet::new();
        let mut cache_rec_paths: BTreeSet<String> = BTreeSet::new();

        for path in entries.keys() {
            if path.starts_with("xl/pivotTables/") && path.ends_with(".xml") {
                table_paths.insert(path.clone());
            } else if path.starts_with("xl/pivotCache/") && path.contains("pivotCacheDefinition") && path.ends_with(".xml") {
                cache_def_paths.insert(path.clone());
            } else if path.starts_with("xl/pivotCache/") && path.contains("pivotCacheRecords") && path.ends_with(".xml") {
                cache_rec_paths.insert(path.clone());
            }
        }

        for path in table_paths {
            let xml = entries.get(&path).unwrap();
            pivots.pivot_tables.push(parse_pivot_table_part(&path, xml)?);
        }
        for path in cache_def_paths {
            let xml = entries.get(&path).unwrap();
            pivots
                .pivot_cache_definitions
                .push(parse_pivot_cache_definition_part(&path, xml)?);
        }
        for path in cache_rec_paths {
            let xml = entries.get(&path).unwrap();
            pivots
                .pivot_cache_records
                .push(parse_pivot_cache_records_part(&path, xml)?);
        }

        Ok(pivots)
    }

    pub fn all_part_paths(&self) -> Vec<String> {
        let mut out = Vec::new();
        for p in &self.pivot_tables {
            out.push(p.path.clone());
        }
        for p in &self.pivot_cache_definitions {
            out.push(p.path.clone());
        }
        for p in &self.pivot_cache_records {
            out.push(p.path.clone());
        }
        out
    }
}

fn parse_pivot_table_part(path: &str, xml: &[u8]) -> Result<PivotTablePart, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut name = None;
    let mut cache_id = None;

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Start(e) => {
                if e.local_name().as_ref() == b"pivotTableDefinition" {
                    for attr in e.attributes().with_checks(false) {
                        let attr = attr.map_err(quick_xml::Error::from)?;
                        match attr.key.local_name().as_ref() {
                            b"name" => name = Some(attr.unescape_value()?.to_string()),
                            b"cacheId" => {
                                if let Ok(v) = attr.unescape_value()?.parse::<u32>() {
                                    cache_id = Some(v);
                                }
                            }
                            _ => {}
                        }
                    }
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(PivotTablePart {
        path: path.to_string(),
        name,
        cache_id,
    })
}

fn parse_pivot_cache_definition_part(path: &str, xml: &[u8]) -> Result<PivotCacheDefinitionPart, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut record_count = None;
    let mut fields = Vec::new();

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Start(e) | Event::Empty(e) => {
                if e.local_name().as_ref() == b"pivotCacheDefinition" {
                    for attr in e.attributes().with_checks(false) {
                        let attr = attr.map_err(quick_xml::Error::from)?;
                        if attr.key.local_name().as_ref() == b"recordCount" {
                            if let Ok(v) = attr.unescape_value()?.parse::<u64>() {
                                record_count = Some(v);
                            }
                        }
                    }
                } else if e.local_name().as_ref() == b"cacheField" {
                    for attr in e.attributes().with_checks(false) {
                        let attr = attr.map_err(quick_xml::Error::from)?;
                        if attr.key.local_name().as_ref() == b"name" {
                            fields.push(attr.unescape_value()?.to_string());
                        }
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(PivotCacheDefinitionPart {
        path: path.to_string(),
        record_count,
        fields,
    })
}

fn parse_pivot_cache_records_part(path: &str, xml: &[u8]) -> Result<PivotCacheRecordsPart, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut count = None;

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Start(e) => {
                if e.local_name().as_ref() == b"pivotCacheRecords" {
                    for attr in e.attributes().with_checks(false) {
                        let attr = attr.map_err(quick_xml::Error::from)?;
                        if attr.key.local_name().as_ref() == b"count" {
                            if let Ok(v) = attr.unescape_value()?.parse::<u64>() {
                                count = Some(v);
                            }
                        }
                    }
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(PivotCacheRecordsPart {
        path: path.to_string(),
        count,
    })
}

#[cfg(test)]
mod tests {
    use crate::XlsxPackage;

    use pretty_assertions::assert_eq;

    #[test]
    fn preserves_pivot_parts_on_round_trip() {
        let fixture = include_bytes!("../tests/fixtures/pivot-fixture.xlsx");
        let pkg = XlsxPackage::from_bytes(fixture).expect("read pkg");

        let pivots = pkg.pivots().expect("parse pivots");
        assert_eq!(pivots.pivot_tables.len(), 1);
        assert_eq!(pivots.pivot_tables[0].name.as_deref(), Some("PivotTable1"));
        assert_eq!(pivots.pivot_cache_definitions.len(), 1);
        assert_eq!(
            pivots.pivot_cache_definitions[0].fields,
            vec!["Region".to_string(), "Product".to_string(), "Sales".to_string()]
        );

        let original_parts: Vec<(String, Vec<u8>)> = pivots
            .all_part_paths()
            .into_iter()
            .map(|p| (p.clone(), pkg.part(&p).expect("part exists").to_vec()))
            .collect();

        let written = pkg.write_to_bytes().expect("write pkg");
        let pkg2 = XlsxPackage::from_bytes(&written).expect("read pkg2");

        for (path, bytes) in original_parts {
            assert_eq!(pkg2.part(&path), Some(bytes.as_slice()), "part {path} differs");
        }
    }
}
