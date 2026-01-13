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
            // Resolve the target using the same URI normalization as other code paths (including
            // fragment stripping). Note that a target of just `#fragment` refers to the source part.
            //
            // Some producers percent-encode relationship targets (e.g. `sheet%201.xml`) while
            // writing the actual ZIP entry name unescaped (or vice versa). Try the raw normalized
            // target first, then fall back to a best-effort percent-decoded candidate if the raw
            // part is missing.
            let candidates = crate::path::resolve_target_candidates(part_name, &rel.target);
            let mut resolved: Option<String> = None;
            for candidate in &candidates {
                // Prefer a candidate that matches an *actual* stored part name to keep part-name
                // strings canonical for downstream callers (some parts are stored unescaped while
                // relationships use percent-encoding, or vice versa).
                if package.parts_map().contains_key(candidate)
                    || package.parts_map().contains_key(format!("/{candidate}").as_str())
                {
                    resolved = Some(candidate.clone());
                    break;
                }
            }
            let resolved = resolved.unwrap_or_else(|| {
                candidates
                    .into_iter()
                    .next()
                    .unwrap_or_else(|| resolve_target(part_name, &rel.target))
            });
            if resolved.is_empty() {
                return Ok(None);
            }
            return Ok(Some(resolved));
        }
    }
    Ok(None)
}

/// Resolve an OPC relationship target using an arbitrary part getter.
///
/// This is the generic counterpart of [`resolve_relationship_target`]. It's used by higher-level
/// representations like [`crate::XlsxDocument`] that store raw parts but aren't backed by an
/// [`XlsxPackage`].
pub fn resolve_relationship_target_from_parts<'a, F>(
    get_part: F,
    part_name: &str,
    relationship_id: &str,
) -> Result<Option<String>, XlsxError>
where
    F: Fn(&str) -> Option<&'a [u8]>,
{
    // Most callers use canonical OPC names (`xl/...`), but tolerate non-standard inputs that still
    // show up in the wild.
    let part_name = part_name.strip_prefix('/').unwrap_or(part_name);

    let rels_name = rels_part_name(part_name);
    let rels_bytes = match get_part(&rels_name) {
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
            // Resolve the target using the same URI normalization as other code paths (including
            // fragment stripping). Note that a target of just `#fragment` refers to the source part.
            //
            // Some producers percent-encode relationship targets (e.g. `sheet%201.xml`) while
            // writing the actual ZIP entry name unescaped (or vice versa). Try the raw normalized
            // target first, then fall back to a best-effort percent-decoded candidate if the raw
            // part is missing.
            let candidates = crate::path::resolve_target_candidates(part_name, &rel.target);
            let resolved = candidates.iter().find(|candidate| {
                let candidate = candidate.as_str();
                if get_part(candidate).is_some() {
                    return true;
                }
                if let Some(stripped) = candidate.strip_prefix('/') {
                    return get_part(stripped).is_some();
                }
                let mut with_slash = String::with_capacity(candidate.len() + 1);
                with_slash.push('/');
                with_slash.push_str(candidate);
                get_part(with_slash.as_str()).is_some()
            });
            let resolved = resolved.cloned().unwrap_or_else(|| {
                candidates
                    .into_iter()
                    .next()
                    .unwrap_or_else(|| resolve_target(part_name, &rel.target))
            });
            if resolved.is_empty() {
                return Ok(None);
            }
            return Ok(Some(resolved));
        }
    }
    Ok(None)
}

pub fn resolve_target(base_part: &str, target: &str) -> String {
    // Keep relationship target resolution centralized in `path::resolve_target` so behavior stays
    // consistent across the codebase (including fragment/query stripping and Windows-path
    // tolerance).
    crate::path::resolve_target(base_part, target)
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
                    // Best-effort: some producers emit incomplete relationship entries (missing
                    // `Type`). For traversal/preservation, `Id` + `Target` are sufficient.
                    if let (Some(id), Some(target)) = (id, target) {
                        relationships.push(Relationship {
                            id,
                            target,
                            type_uri: type_uri.unwrap_or_default(),
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
    fn parse_relationships_tolerates_missing_type() {
        let rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Target="../media/image1.png"/>
</Relationships>"#;

        let parsed = parse_relationships(rels).expect("parse relationships");
        assert_eq!(parsed.len(), 1);
        assert_eq!(parsed[0].id, "rId1");
        assert_eq!(parsed[0].target, "../media/image1.png");
        assert_eq!(parsed[0].type_uri, "");
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
    fn resolve_relationship_target_handles_fragment_only_targets() {
        let rels = br##"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="urn:example:self" Target="#frag"/>
</Relationships>"##;

        let pkg = build_package(&[
            ("xl/worksheets/sheet1.xml", br#"<worksheet/>"#),
            ("xl/worksheets/_rels/sheet1.xml.rels", rels),
        ]);

        assert_eq!(
            resolve_relationship_target(&pkg, "xl/worksheets/sheet1.xml", "rId1")
                .expect("resolve fragment-only"),
            Some("xl/worksheets/sheet1.xml".to_string())
        );
    }

    #[test]
    fn resolve_relationship_target_prefers_percent_decoded_target_when_part_exists() {
        let rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image%201.png"/>
</Relationships>"#;

        let pkg = build_package(&[
            ("xl/worksheets/sheet1.xml", br#"<worksheet/>"#),
            ("xl/worksheets/_rels/sheet1.xml.rels", rels),
            ("xl/media/image 1.png", b"png-bytes"),
        ]);

        assert_eq!(
            resolve_relationship_target(&pkg, "xl/worksheets/sheet1.xml", "rId1")
                .expect("resolve percent-encoded")
                .as_deref(),
            Some("xl/media/image 1.png")
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

    #[test]
    fn resolve_target_normalizes_backslashes() {
        assert_eq!(
            resolve_target("xl/workbook.xml", "worksheets\\sheet1.xml#rId1"),
            "xl/worksheets/sheet1.xml"
        );
        assert_eq!(
            resolve_target("xl/worksheets/sheet1.xml", "..\\media\\image1.png"),
            "xl/media/image1.png"
        );
        assert_eq!(
            resolve_target("xl/_rels/workbook.xml.rels", "\\xl\\media\\image1.png#frag"),
            "xl/media/image1.png"
        );
    }
}
