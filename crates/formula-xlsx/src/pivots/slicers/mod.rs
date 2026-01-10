use crate::package::{XlsxError, XlsxPackage};
use crate::openxml::{local_name, parse_relationships, resolve_relationship_target, resolve_target};
use quick_xml::events::Event;
use quick_xml::Reader;
use std::collections::{BTreeMap, BTreeSet};
use std::io::Cursor;

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SlicerDefinition {
    pub part_name: String,
    pub name: Option<String>,
    pub uid: Option<String>,
    pub cache_part: Option<String>,
    pub connected_pivot_tables: Vec<String>,
    pub placed_on_drawings: Vec<String>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TimelineDefinition {
    pub part_name: String,
    pub name: Option<String>,
    pub uid: Option<String>,
    pub cache_part: Option<String>,
    pub connected_pivot_tables: Vec<String>,
    pub placed_on_drawings: Vec<String>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct PivotSlicerParts {
    pub slicers: Vec<SlicerDefinition>,
    pub timelines: Vec<TimelineDefinition>,
}

impl XlsxPackage {
    /// Parse slicers and timelines out of an XLSX package.
    ///
    /// This parser is intentionally conservative: it extracts the minimum metadata needed to
    /// wire up the UX layer, while leaving the XML untouched for round-trip fidelity.
    pub fn pivot_slicer_parts(&self) -> Result<PivotSlicerParts, XlsxError> {
        parse_pivot_slicer_parts(self)
    }
}

fn parse_pivot_slicer_parts(package: &XlsxPackage) -> Result<PivotSlicerParts, XlsxError> {
    let slicer_parts = package
        .part_names()
        .filter(|name| name.starts_with("xl/slicers/") && name.ends_with(".xml"))
        .map(str::to_string)
        .collect::<Vec<_>>();
    let timeline_parts = package
        .part_names()
        .filter(|name| name.starts_with("xl/timelines/") && name.ends_with(".xml"))
        .map(str::to_string)
        .collect::<Vec<_>>();

    let drawing_rels = package
        .part_names()
        .filter(|name| name.starts_with("xl/drawings/_rels/") && name.ends_with(".rels"))
        .map(str::to_string)
        .collect::<Vec<_>>();

    let mut slicer_to_drawings: BTreeMap<String, BTreeSet<String>> = BTreeMap::new();
    let mut timeline_to_drawings: BTreeMap<String, BTreeSet<String>> = BTreeMap::new();

    for rels_name in drawing_rels {
        let rels_bytes = package
            .part(&rels_name)
            .ok_or_else(|| XlsxError::MissingPart(rels_name.clone()))?;
        let relationships = parse_relationships(rels_bytes)?;
        let drawing_part = drawing_part_name_from_rels(&rels_name);
        for rel in relationships {
            let target = resolve_target(&drawing_part, &rel.target);
            if target.starts_with("xl/slicers/") {
                slicer_to_drawings
                    .entry(target)
                    .or_default()
                    .insert(drawing_part.clone());
            } else if target.starts_with("xl/timelines/") {
                timeline_to_drawings
                    .entry(target)
                    .or_default()
                    .insert(drawing_part.clone());
            }
        }
    }

    let mut slicers = Vec::with_capacity(slicer_parts.len());
    for part_name in slicer_parts {
        let xml = package
            .part(&part_name)
            .ok_or_else(|| XlsxError::MissingPart(part_name.clone()))?;
        let parsed = parse_slicer_xml(xml)?;

        let cache_part = match parsed.cache_rid.as_deref() {
            Some(rid) => resolve_relationship_target(package, &part_name, rid)?,
            None => None,
        };

        let connected_pivot_tables = if let Some(cache_part) = cache_part.as_deref() {
            resolve_slicer_cache_pivot_tables(package, cache_part)?
        } else {
            Vec::new()
        };

        let placed_on_drawings = slicer_to_drawings
            .get(&part_name)
            .map(|drawings| drawings.iter().cloned().collect::<Vec<_>>())
            .unwrap_or_default();

        slicers.push(SlicerDefinition {
            part_name: part_name.clone(),
            name: parsed.name,
            uid: parsed.uid,
            cache_part,
            connected_pivot_tables,
            placed_on_drawings,
        });
    }

    let mut timelines = Vec::with_capacity(timeline_parts.len());
    for part_name in timeline_parts {
        let xml = package
            .part(&part_name)
            .ok_or_else(|| XlsxError::MissingPart(part_name.clone()))?;
        let parsed = parse_timeline_xml(xml)?;

        let cache_part = match parsed.cache_rid.as_deref() {
            Some(rid) => resolve_relationship_target(package, &part_name, rid)?,
            None => None,
        };

        let connected_pivot_tables = if let Some(cache_part) = cache_part.as_deref() {
            resolve_timeline_cache_pivot_tables(package, cache_part)?
        } else {
            Vec::new()
        };

        let placed_on_drawings = timeline_to_drawings
            .get(&part_name)
            .map(|drawings| drawings.iter().cloned().collect::<Vec<_>>())
            .unwrap_or_default();

        timelines.push(TimelineDefinition {
            part_name: part_name.clone(),
            name: parsed.name,
            uid: parsed.uid,
            cache_part,
            connected_pivot_tables,
            placed_on_drawings,
        });
    }

    Ok(PivotSlicerParts { slicers, timelines })
}

fn resolve_slicer_cache_pivot_tables(
    package: &XlsxPackage,
    cache_part: &str,
) -> Result<Vec<String>, XlsxError> {
    let cache_bytes = package
        .part(cache_part)
        .ok_or_else(|| XlsxError::MissingPart(cache_part.to_string()))?;
    let pivot_rids = parse_slicer_cache_xml(cache_bytes)?;

    let mut pivot_tables = BTreeSet::new();
    for rid in pivot_rids {
        if let Some(target) = resolve_relationship_target(package, cache_part, &rid)? {
            pivot_tables.insert(target);
        }
    }
    Ok(pivot_tables.into_iter().collect())
}

fn resolve_timeline_cache_pivot_tables(
    package: &XlsxPackage,
    cache_part: &str,
) -> Result<Vec<String>, XlsxError> {
    let cache_bytes = package
        .part(cache_part)
        .ok_or_else(|| XlsxError::MissingPart(cache_part.to_string()))?;
    let pivot_rids = parse_timeline_cache_xml(cache_bytes)?;

    let mut pivot_tables = BTreeSet::new();
    for rid in pivot_rids {
        if let Some(target) = resolve_relationship_target(package, cache_part, &rid)? {
            pivot_tables.insert(target);
        }
    }
    Ok(pivot_tables.into_iter().collect())
}

fn drawing_part_name_from_rels(rels_name: &str) -> String {
    // Example: xl/drawings/_rels/drawing1.xml.rels -> xl/drawings/drawing1.xml
    let trimmed = rels_name
        .strip_prefix("xl/drawings/_rels/")
        .unwrap_or(rels_name);
    let trimmed = trimmed.strip_suffix(".rels").unwrap_or(trimmed);
    format!("xl/drawings/{trimmed}")
}

#[derive(Debug)]
struct ParsedSlicerXml {
    name: Option<String>,
    uid: Option<String>,
    cache_rid: Option<String>,
}

fn parse_slicer_xml(xml: &[u8]) -> Result<ParsedSlicerXml, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut name = None;
    let mut uid = None;
    let mut cache_rid = None;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(start) | Event::Empty(start) => {
                let element_name = start.name();
                let tag = local_name(element_name.as_ref());
                if tag.eq_ignore_ascii_case(b"slicer") {
                    for attr in start.attributes() {
                        let attr = attr?;
                        let key = local_name(attr.key.as_ref());
                        let value = attr.unescape_value()?.into_owned();
                        if key.eq_ignore_ascii_case(b"name") {
                            name = Some(value);
                        } else if key.eq_ignore_ascii_case(b"uid") {
                            uid = Some(value);
                        }
                    }
                } else if tag.eq_ignore_ascii_case(b"slicerCache") {
                    for attr in start.attributes() {
                        let attr = attr?;
                        if local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"id") {
                            cache_rid = Some(attr.unescape_value()?.into_owned());
                        }
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(ParsedSlicerXml { name, uid, cache_rid })
}

fn parse_slicer_cache_xml(xml: &[u8]) -> Result<Vec<String>, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut pivot_table_rids = Vec::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(start) | Event::Empty(start) => {
                if local_name(start.name().as_ref()).eq_ignore_ascii_case(b"slicerCachePivotTable")
                {
                    for attr in start.attributes() {
                        let attr = attr?;
                        if local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"id") {
                            pivot_table_rids.push(attr.unescape_value()?.into_owned());
                        }
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(pivot_table_rids)
}

#[derive(Debug)]
struct ParsedTimelineXml {
    name: Option<String>,
    uid: Option<String>,
    cache_rid: Option<String>,
}

fn parse_timeline_xml(xml: &[u8]) -> Result<ParsedTimelineXml, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();

    let mut name = None;
    let mut uid = None;
    let mut cache_rid = None;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(start) | Event::Empty(start) => {
                let element_name = start.name();
                let tag = local_name(element_name.as_ref());
                if tag.eq_ignore_ascii_case(b"timeline") {
                    for attr in start.attributes() {
                        let attr = attr?;
                        let key = local_name(attr.key.as_ref());
                        let value = attr.unescape_value()?.into_owned();
                        if key.eq_ignore_ascii_case(b"name") {
                            name = Some(value);
                        } else if key.eq_ignore_ascii_case(b"uid") {
                            uid = Some(value);
                        }
                    }
                } else if tag.eq_ignore_ascii_case(b"timelineCache") {
                    for attr in start.attributes() {
                        let attr = attr?;
                        if local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"id") {
                            cache_rid = Some(attr.unescape_value()?.into_owned());
                        }
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(ParsedTimelineXml { name, uid, cache_rid })
}

fn parse_timeline_cache_xml(xml: &[u8]) -> Result<Vec<String>, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut pivot_table_rids = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(start) | Event::Empty(start) => {
                if local_name(start.name().as_ref()).eq_ignore_ascii_case(b"pivotTable") {
                    for attr in start.attributes() {
                        let attr = attr?;
                        if local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"id") {
                            pivot_table_rids.push(attr.unescape_value()?.into_owned());
                        }
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(pivot_table_rids)
}
