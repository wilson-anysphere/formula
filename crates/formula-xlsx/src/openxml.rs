use crate::package::{XlsxError, XlsxPackage};
use quick_xml::events::Event;
use quick_xml::Reader;
use std::io::Cursor;

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Relationship {
    pub id: String,
    pub type_uri: String,
    pub target: String,
    pub target_mode: Option<String>,
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
            if rel
                .target_mode
                .as_deref()
                .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
            {
                return Ok(None);
            }
            let target = strip_fragment(&rel.target);
            if target.is_empty() {
                return Ok(None);
            }
            return Ok(Some(resolve_target(part_name, target)));
        }
    }
    Ok(None)
}

fn strip_fragment(target: &str) -> &str {
    target
        .split_once('#')
        .map(|(base, _)| base)
        .unwrap_or(target)
}

pub fn resolve_target(base_part: &str, target: &str) -> String {
    // Relationship targets are URIs; some producers include a fragment (e.g. `foo.xml#bar`).
    // OPC part names do not include fragments, so strip them before resolving.
    let target = target
        .split_once('#')
        .map(|(base, _)| base)
        .unwrap_or(target);
    if target.is_empty() {
        // A target of just `#fragment` refers to the source part itself.
        return base_part
            .strip_prefix('/')
            .unwrap_or(base_part)
            .to_string();
    }

    // Relationship targets can be relative to the source part's folder (e.g. `worksheets/sheet1.xml`)
    // or absolute (e.g. `/xl/worksheets/sheet1.xml`). Absolute targets are rooted at the package
    // root and must not be prefixed with the source part directory.
    let (target, is_absolute) = match target.strip_prefix('/') {
        Some(target) => (target, true),
        None => (target, false),
    };
    let base_dir = if is_absolute {
        ""
    } else {
        base_part
            .rsplit_once('/')
            .map(|(dir, _)| dir)
            .unwrap_or("")
    };

    // `base_part` is typically an OPC part name without a leading slash, but be resilient and
    // ignore any empty segments so callers can pass `/xl/...` and still get normalized output.
    let mut components: Vec<&str> = if base_dir.is_empty() {
        Vec::new()
    } else {
        base_dir.split('/').filter(|s| !s.is_empty()).collect()
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
                    let mut target_mode = None;
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
                        } else if key.eq_ignore_ascii_case(b"TargetMode") {
                            target_mode = Some(value);
                        }
                    }
                    if let (Some(id), Some(target), Some(type_uri)) = (id, target, type_uri) {
                        relationships.push(Relationship {
                            id,
                            target,
                            type_uri,
                            target_mode,
                        });
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

#[cfg(test)]
mod tests {
    use super::*;

    use std::io::{Cursor, Write};

    use zip::write::FileOptions;
    use zip::ZipWriter;

    fn build_package(entries: &[(&str, &[u8])]) -> XlsxPackage {
        let cursor = Cursor::new(Vec::new());
        let mut zip = ZipWriter::new(cursor);
        let options =
            FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

        for (name, bytes) in entries {
            zip.start_file(*name, options).unwrap();
            zip.write_all(bytes).unwrap();
        }

        let bytes = zip.finish().unwrap().into_inner();
        XlsxPackage::from_bytes(&bytes).expect("read test pkg")
    }

    #[test]
    fn parse_relationships_captures_target_mode() {
        let rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
</Relationships>"#;

        let parsed = parse_relationships(rels).expect("parse relationships");
        assert_eq!(parsed.len(), 1);
        assert_eq!(parsed[0].id, "rId1");
        assert_eq!(parsed[0].target_mode.as_deref(), Some("External"));
    }

    #[test]
    fn resolve_relationship_target_skips_external() {
        let rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png#frag"/>
</Relationships>"#;

        let pkg = build_package(&[
            ("xl/worksheets/sheet1.xml", br#"<worksheet/>"#),
            ("xl/worksheets/_rels/sheet1.xml.rels", rels),
            ("xl/media/image1.png", b"png-bytes"),
        ]);

        assert_eq!(
            resolve_relationship_target(&pkg, "xl/worksheets/sheet1.xml", "rId1")
                .expect("resolve external"),
            None
        );
        assert_eq!(
            resolve_relationship_target(&pkg, "xl/worksheets/sheet1.xml", "rId2")
                .expect("resolve internal")
                .as_deref(),
            Some("xl/media/image1.png")
        );
    }

    #[test]
    fn resolve_target_strips_uri_fragments() {
        assert_eq!(
            resolve_target("xl/metadata.xml", "richData/rd1.xml#frag"),
            "xl/richData/rd1.xml"
        );
        assert_eq!(
            resolve_target("xl/_rels/workbook.xml.rels", "/xl/media/image1.png#frag"),
            "xl/media/image1.png"
        );
        assert_eq!(
            resolve_target("/xl/metadata.xml", "richData/rd1.xml#frag"),
            "xl/richData/rd1.xml"
        );
        assert_eq!(resolve_target("/xl/metadata.xml", "#frag"), "xl/metadata.xml");
    }
}
