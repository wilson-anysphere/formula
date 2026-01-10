use crate::package::{XlsxError, XlsxPackage};
use quick_xml::events::Event;
use quick_xml::Reader;
use std::io::Cursor;

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Relationship {
    pub id: String,
    pub type_uri: String,
    pub target: String,
}

pub fn rels_part_name(part_name: &str) -> String {
    let (dir, file) = part_name
        .rsplit_once('/')
        .unwrap_or(("", part_name));
    if dir.is_empty() {
        format!("_rels/{file}.rels")
    } else {
        format!("{dir}/_rels/{file}.rels")
    }
}

pub fn resolve_relationship_target(
    package: &XlsxPackage,
    part_name: &str,
    relationship_id: &str,
) -> Result<Option<String>, XlsxError> {
    let rels_name = rels_part_name(part_name);
    let rels_bytes = match package.part(&rels_name) {
        Some(bytes) => bytes,
        None => return Ok(None),
    };
    let relationships = parse_relationships(rels_bytes)?;
    for rel in relationships {
        if rel.id == relationship_id {
            return Ok(Some(resolve_target(part_name, &rel.target)));
        }
    }
    Ok(None)
}

pub fn resolve_target(base_part: &str, target: &str) -> String {
    let target = target.trim_start_matches('/');
    let base_dir = base_part.rsplit_once('/').map(|(dir, _)| dir).unwrap_or("");

    let mut components: Vec<&str> = if base_dir.is_empty() {
        Vec::new()
    } else {
        base_dir.split('/').collect()
    };

    for segment in target.split('/') {
        match segment {
            "" | "." => {}
            ".." => {
                components.pop();
            }
            _ => components.push(segment),
        }
    }

    components.join("/")
}

pub fn parse_relationships(xml: &[u8]) -> Result<Vec<Relationship>, XlsxError> {
    let mut reader = Reader::from_reader(Cursor::new(xml));
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut relationships = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(start) | Event::Empty(start) => {
                if local_name(start.name().as_ref()).eq_ignore_ascii_case(b"Relationship") {
                    let mut id = None;
                    let mut target = None;
                    let mut type_uri = None;
                    for attr in start.attributes() {
                        let attr = attr?;
                        let key = local_name(attr.key.as_ref());
                        let value = attr.unescape_value()?.into_owned();
                        if key.eq_ignore_ascii_case(b"Id") {
                            id = Some(value);
                        } else if key.eq_ignore_ascii_case(b"Target") {
                            target = Some(value);
                        } else if key.eq_ignore_ascii_case(b"Type") {
                            type_uri = Some(value);
                        }
                    }
                    if let (Some(id), Some(target), Some(type_uri)) = (id, target, type_uri) {
                        relationships.push(Relationship { id, target, type_uri });
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(relationships)
}

pub fn local_name(name: &[u8]) -> &[u8] {
    match name.iter().rposition(|b| *b == b':') {
        Some(idx) => &name[idx + 1..],
        None => name,
    }
}
