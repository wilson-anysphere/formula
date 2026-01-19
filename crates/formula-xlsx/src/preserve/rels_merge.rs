use std::collections::{HashMap, HashSet};

use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader as XmlReader, Writer as XmlWriter};

use crate::path::resolve_target;
use crate::relationships::{parse_relationships, Relationship, Relationships, PACKAGE_REL_NS};
use crate::workbook::ChartExtractionError;

/// Minimal metadata needed to re-attach preserved relationships.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RelationshipStub {
    pub rel_id: String,
    pub target: String,
}

pub(crate) fn ensure_rels_has_relationships(
    rels_xml: Option<&[u8]>,
    part_name: &str,
    base_part: &str,
    rel_type: &str,
    relationships: &[RelationshipStub],
) -> Result<(Vec<u8>, HashMap<String, String>), ChartExtractionError> {
    if relationships.is_empty() {
        return Ok((rels_xml.unwrap_or_default().to_vec(), HashMap::new()));
    }

    let mut xml_bytes: Vec<u8> = match rels_xml {
        Some(bytes) => bytes.to_vec(),
        None => format!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"{PACKAGE_REL_NS}\">\n</Relationships>\n"
        )
        .into_bytes(),
    };

    let existing_rels = match rels_xml {
        Some(bytes) => parse_relationships(bytes, part_name)?,
        None => Vec::new(),
    };
    let mut rels = Relationships::new(existing_rels);

    let mut id_map: HashMap<String, String> = HashMap::new();
    let mut to_insert: Vec<Relationship> = Vec::new();

    for relationship in relationships {
        let desired_id = relationship.rel_id.as_str();
        let desired_target = relationship.target.as_str();

        if let Some(mapped) = id_map.get(desired_id) {
            // We've already allocated a stable replacement for this ID in this scope.
            // Ensure the relationship exists in the output but don't allocate again.
            if rels.get(mapped).is_none() {
                let rel = Relationship {
                    id: mapped.clone(),
                    type_: rel_type.to_string(),
                    target: desired_target.to_string(),
                    target_mode: None,
                };
                rels.push(rel.clone());
                to_insert.push(rel);
            }
            continue;
        }

        let final_id = match rels.get(desired_id) {
            None => desired_id.to_string(),
            Some(existing)
                if existing.type_ == rel_type
                    && resolve_target(base_part, &existing.target)
                        == resolve_target(base_part, desired_target) =>
            {
                desired_id.to_string()
            }
            Some(_) => {
                let new_id = rels.next_r_id();
                id_map.insert(desired_id.to_string(), new_id.clone());
                new_id
            }
        };

        if rels.get(&final_id).is_some() {
            continue;
        }

        let rel = Relationship {
            id: final_id.clone(),
            type_: rel_type.to_string(),
            target: desired_target.to_string(),
            target_mode: None,
        };
        rels.push(rel.clone());
        to_insert.push(rel);
    }

    if !to_insert.is_empty() {
        xml_bytes = insert_relationships_before_close(&xml_bytes, part_name, &to_insert)?;
    }

    Ok((xml_bytes, id_map))
}

fn insert_relationships_before_close(
    xml: &[u8],
    part_name: &str,
    to_insert: &[Relationship],
) -> Result<Vec<u8>, ChartExtractionError> {
    fn local_name(name: &[u8]) -> &[u8] {
        crate::openxml::local_name(name)
    }

    fn element_prefix(name: &[u8]) -> Option<&[u8]> {
        name.iter().rposition(|b| *b == b':').map(|idx| &name[..idx])
    }

    fn prefixed_tag(prefix: Option<&str>, local: &str) -> String {
        match prefix {
            Some(prefix) => format!("{prefix}:{local}"),
            None => local.to_string(),
        }
    }

    let mut reader = XmlReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut out = Vec::new();
    let Some(extra) = to_insert.len().checked_mul(128) else {
        return Err(ChartExtractionError::AllocationFailure(
            "insert_relationships_before_close output",
        ));
    };
    let Some(cap) = xml.len().checked_add(extra) else {
        return Err(ChartExtractionError::AllocationFailure(
            "insert_relationships_before_close output",
        ));
    };
    if out.try_reserve(cap).is_err() {
        return Err(ChartExtractionError::AllocationFailure(
            "insert_relationships_before_close output",
        ));
    }
    let mut writer = XmlWriter::new(out);
    let mut buf = Vec::new();

    let mut root_prefix: Option<String> = None;
    let mut root_has_default_ns = false;
    let mut root_declared_prefixes: HashSet<String> = HashSet::new();
    let mut relationship_prefix: Option<String> = None;

    loop {
        let event = reader
            .read_event_into(&mut buf)
            .map_err(|e| ChartExtractionError::XmlStructure(format!("{part_name}: {e}")))?;
        match event {
            Event::Start(e) if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Relationships") => {
                if root_prefix.is_none() {
                    root_prefix = element_prefix(e.name().as_ref())
                        .and_then(|p| std::str::from_utf8(p).ok())
                        .map(|s| s.to_string());
                }
                if !root_has_default_ns || root_declared_prefixes.is_empty() {
                    for attr in e.attributes() {
                        let attr = attr.map_err(|e| {
                            ChartExtractionError::XmlStructure(format!("{part_name}: {e}"))
                        })?;
                        let key = attr.key.as_ref();
                        if key == b"xmlns" && attr.value.as_ref() == PACKAGE_REL_NS.as_bytes() {
                            root_has_default_ns = true;
                        } else if let Some(prefix) = key.strip_prefix(b"xmlns:") {
                            if attr.value.as_ref() == PACKAGE_REL_NS.as_bytes() {
                                if let Ok(prefix) = std::str::from_utf8(prefix) {
                                    root_declared_prefixes.insert(prefix.to_string());
                                }
                            }
                        }
                    }
                }
                writer
                    .write_event(Event::Start(e))
                    .map_err(|e| ChartExtractionError::XmlStructure(format!("{part_name}: {e}")))?;
            }
            Event::Start(e) if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Relationship") => {
                if relationship_prefix.is_none() {
                    relationship_prefix = element_prefix(e.name().as_ref())
                        .and_then(|p| std::str::from_utf8(p).ok())
                        .map(|s| s.to_string());
                }
                writer
                    .write_event(Event::Start(e))
                    .map_err(|e| ChartExtractionError::XmlStructure(format!("{part_name}: {e}")))?;
            }
            Event::Empty(e) if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Relationship") => {
                if relationship_prefix.is_none() {
                    relationship_prefix = element_prefix(e.name().as_ref())
                        .and_then(|p| std::str::from_utf8(p).ok())
                        .map(|s| s.to_string());
                }
                writer
                    .write_event(Event::Empty(e))
                    .map_err(|e| ChartExtractionError::XmlStructure(format!("{part_name}: {e}")))?;
            }
            Event::End(e) if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Relationships") => {
                let prefix = relationship_prefix
                    .as_deref()
                    // Only reuse the existing Relationship element prefix if it is declared on the
                    // root element. Otherwise, we could emit a new sibling with an out-of-scope
                    // prefix (invalid XML), e.g. if the input declares `xmlns:pr` on each
                    // `<pr:Relationship>` element instead of the root.
                    .filter(|p| root_declared_prefixes.contains(*p))
                    .or_else(|| {
                        if root_has_default_ns {
                            None
                        } else {
                            root_prefix.as_deref()
                        }
                    });
                let relationship_tag_name = prefixed_tag(prefix, "Relationship");

                for rel in to_insert {
                    let mut e = BytesStart::new(relationship_tag_name.as_str());
                    e.push_attribute(("Id", rel.id.as_str()));
                    e.push_attribute(("Type", rel.type_.as_str()));
                    e.push_attribute(("Target", rel.target.as_str()));
                    if let Some(mode) = &rel.target_mode {
                        e.push_attribute(("TargetMode", mode.as_str()));
                    }
                    writer
                        .write_event(Event::Empty(e))
                        .map_err(|e| ChartExtractionError::XmlStructure(format!("{part_name}: {e}")))?;
                }

                writer
                    .write_event(Event::End(e))
                    .map_err(|e| ChartExtractionError::XmlStructure(format!("{part_name}: {e}")))?;
            }
            Event::Empty(e) if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Relationships") => {
                let root_name = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                if root_prefix.is_none() {
                    root_prefix = element_prefix(e.name().as_ref())
                        .and_then(|p| std::str::from_utf8(p).ok())
                        .map(|s| s.to_string());
                }
                if !root_has_default_ns || root_declared_prefixes.is_empty() {
                    for attr in e.attributes() {
                        let attr = attr.map_err(|e| {
                            ChartExtractionError::XmlStructure(format!("{part_name}: {e}"))
                        })?;
                        let key = attr.key.as_ref();
                        if key == b"xmlns" && attr.value.as_ref() == PACKAGE_REL_NS.as_bytes() {
                            root_has_default_ns = true;
                        } else if let Some(prefix) = key.strip_prefix(b"xmlns:") {
                            if attr.value.as_ref() == PACKAGE_REL_NS.as_bytes() {
                                if let Ok(prefix) = std::str::from_utf8(prefix) {
                                    root_declared_prefixes.insert(prefix.to_string());
                                }
                            }
                        }
                    }
                }
                let prefix = relationship_prefix
                    .as_deref()
                    .filter(|p| root_declared_prefixes.contains(*p))
                    .or_else(|| {
                        if root_has_default_ns {
                            None
                        } else {
                            root_prefix.as_deref()
                        }
                    });
                let relationship_tag_name = prefixed_tag(prefix, "Relationship");

                writer
                    .write_event(Event::Start(e))
                    .map_err(|e| ChartExtractionError::XmlStructure(format!("{part_name}: {e}")))?;

                for rel in to_insert {
                    let mut e = BytesStart::new(relationship_tag_name.as_str());
                    e.push_attribute(("Id", rel.id.as_str()));
                    e.push_attribute(("Type", rel.type_.as_str()));
                    e.push_attribute(("Target", rel.target.as_str()));
                    if let Some(mode) = &rel.target_mode {
                        e.push_attribute(("TargetMode", mode.as_str()));
                    }
                    writer
                        .write_event(Event::Empty(e))
                        .map_err(|e| ChartExtractionError::XmlStructure(format!("{part_name}: {e}")))?;
                }

                writer
                    .write_event(Event::End(BytesEnd::new(root_name.as_str())))
                    .map_err(|e| ChartExtractionError::XmlStructure(format!("{part_name}: {e}")))?;
            }
            Event::Eof => break,
            other => {
                writer
                    .write_event(other)
                    .map_err(|e| ChartExtractionError::XmlStructure(format!("{part_name}: {e}")))?;
            }
        }

        buf.clear();
    }

    Ok(writer.into_inner())
}

pub(crate) fn xml_escape(input: &str) -> String {
    input
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::relationships::PACKAGE_REL_NS;

    #[test]
    fn ensure_rels_inserts_before_relationships_close_with_whitespace() {
        let rels_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships >"#;

        let (updated, id_map) = ensure_rels_has_relationships(
            Some(rels_xml),
            "xl/_rels/workbook.xml.rels",
            "xl/workbook.xml",
            "http://schemas.microsoft.com/office/2006/relationships/vbaProject",
            &[RelationshipStub {
                rel_id: "rId2".to_string(),
                target: "vbaProject.bin".to_string(),
            }],
        )
        .expect("repair rels");

        assert!(id_map.is_empty());

        let rels =
            parse_relationships(&updated, "xl/_rels/workbook.xml.rels").expect("parse rels");
        assert_eq!(rels.len(), 2);
        assert!(rels
            .iter()
            .any(|r| r.id == "rId2" && r.type_ == "http://schemas.microsoft.com/office/2006/relationships/vbaProject"));
    }

    #[test]
    fn ensure_rels_preserves_prefix_when_inserting_into_prefixed_relationships() {
        let rels_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pr:Relationships xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships">
  <pr:Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</pr:Relationships >"#;

        let (updated, _) = ensure_rels_has_relationships(
            Some(rels_xml),
            "xl/_rels/workbook.xml.rels",
            "xl/workbook.xml",
            "http://schemas.microsoft.com/office/2006/relationships/vbaProject",
            &[RelationshipStub {
                rel_id: "rId2".to_string(),
                target: "vbaProject.bin".to_string(),
            }],
        )
        .expect("repair rels");

        let updated_str = std::str::from_utf8(&updated).unwrap();

        // Ensure the newly inserted element is still in the relationships namespace
        // (i.e. it used the same `pr:` prefix, not an unprefixed tag with no namespace).
        let doc = roxmltree::Document::parse(updated_str).unwrap();
        let inserted = doc
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "Relationship" && n.attribute("Id") == Some("rId2"))
            .expect("inserted relationship");
        assert_eq!(inserted.tag_name().namespace(), Some(PACKAGE_REL_NS));
        assert!(
            updated_str.contains(r#"<pr:Relationship Id="rId2""#),
            "expected inserted <pr:Relationship>, got:\n{updated_str}"
        );
        assert!(
            !updated_str.contains("<Relationship"),
            "should not introduce unprefixed <Relationship> tags in prefix-only .rels, got:\n{updated_str}"
        );
    }

    #[test]
    fn ensure_rels_does_not_reuse_out_of_scope_relationship_prefix() {
        // This `.rels` is valid XML: the existing `<x:Relationship>` element declares its prefix
        // on itself instead of the root. When inserting new siblings, we must not reuse that
        // out-of-scope `x:` prefix, or we'd produce invalid XML.
        let rels_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pr:Relationships xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships">
  <x:Relationship xmlns:x="http://schemas.openxmlformats.org/package/2006/relationships" Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</pr:Relationships>"#;

        let (updated, _) = ensure_rels_has_relationships(
            Some(rels_xml),
            "xl/_rels/workbook.xml.rels",
            "xl/workbook.xml",
            "http://schemas.microsoft.com/office/2006/relationships/vbaProject",
            &[RelationshipStub {
                rel_id: "rId2".to_string(),
                target: "vbaProject.bin".to_string(),
            }],
        )
        .expect("repair rels");

        let updated_str = std::str::from_utf8(&updated).unwrap();

        // Must still be valid XML.
        let doc = roxmltree::Document::parse(updated_str).unwrap();
        let inserted = doc
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "Relationship" && n.attribute("Id") == Some("rId2"))
            .expect("inserted relationship");
        assert_eq!(inserted.tag_name().namespace(), Some(PACKAGE_REL_NS));

        // Must use the root prefix (`pr:`), since `x:` is not in scope at the insertion point.
        assert!(
            updated_str.contains(r#"<pr:Relationship Id="rId2""#),
            "expected inserted <pr:Relationship>, got:\n{updated_str}"
        );
        assert!(
            !updated_str.contains(r#"<x:Relationship Id="rId2""#),
            "should not reuse out-of-scope Relationship prefix, got:\n{updated_str}"
        );
        assert!(
            !updated_str.contains("<Relationship"),
            "should not introduce unprefixed <Relationship> tags in prefix-only .rels, got:\n{updated_str}"
        );
    }

    #[test]
    fn ensure_rels_expands_prefixed_self_closing_root() {
        let rels_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pr:Relationships xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships"/>"#;

        let (updated, _) = ensure_rels_has_relationships(
            Some(rels_xml),
            "xl/_rels/workbook.xml.rels",
            "xl/workbook.xml",
            "http://schemas.microsoft.com/office/2006/relationships/vbaProject",
            &[RelationshipStub {
                rel_id: "rId1".to_string(),
                target: "vbaProject.bin".to_string(),
            }],
        )
        .expect("repair rels");

        let updated_str = std::str::from_utf8(&updated).unwrap();
        assert!(
            updated_str.contains("</pr:Relationships>"),
            "expected expanded end tag, got: {updated_str}"
        );
        assert!(
            updated_str.contains(r#"<pr:Relationship Id="rId1""#),
            "expected inserted <pr:Relationship>, got:\n{updated_str}"
        );
        assert!(
            !updated_str.contains("<Relationship"),
            "should not introduce unprefixed <Relationship> tags in prefix-only .rels, got:\n{updated_str}"
        );

        let doc = roxmltree::Document::parse(updated_str).unwrap();
        let inserted = doc
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "Relationship" && n.attribute("Id") == Some("rId1"))
            .expect("inserted relationship");
        assert_eq!(inserted.tag_name().namespace(), Some(PACKAGE_REL_NS));
    }

    #[test]
    fn ensure_rels_uses_root_prefix_when_prefixed_root_has_no_relationship_elements() {
        let rels_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pr:Relationships xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships">
</pr:Relationships>"#;

        let (updated, _) = ensure_rels_has_relationships(
            Some(rels_xml),
            "xl/_rels/workbook.xml.rels",
            "xl/workbook.xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
            &[RelationshipStub {
                rel_id: "rId1".to_string(),
                target: "worksheets/sheet1.xml".to_string(),
            }],
        )
        .expect("repair rels");

        let updated_str = std::str::from_utf8(&updated).unwrap();
        let doc = roxmltree::Document::parse(updated_str).unwrap();
        let inserted = doc
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "Relationship" && n.attribute("Id") == Some("rId1"))
            .expect("inserted relationship");
        assert_eq!(inserted.tag_name().namespace(), Some(PACKAGE_REL_NS));
        assert!(
            updated_str.contains(r#"<pr:Relationship Id="rId1""#),
            "expected inserted <pr:Relationship>, got:\n{updated_str}"
        );
        assert!(
            !updated_str.contains("<Relationship"),
            "should not introduce unprefixed <Relationship> tags in prefix-only .rels, got:\n{updated_str}"
        );
    }

    #[test]
    fn ensure_rels_conflicting_rid_allocates_new_id_and_returns_map() {
        // Note: the close tag contains whitespace (`</Relationships >`) which used to break the
        // naive string search used by relationship insertion.
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="typeA" Target="worksheets/sheet2.xml"/>
</Relationships >"#;

        let (updated, id_map) = ensure_rels_has_relationships(
            Some(xml),
            "xl/_rels/workbook.xml.rels",
            "xl/workbook.xml",
            "typeB",
            &[RelationshipStub {
                rel_id: "rId1".to_string(),
                target: "worksheets/sheet1.xml".to_string(),
            }],
        )
        .expect("insert rel");

        assert_eq!(
            id_map.get("rId1").map(String::as_str),
            Some("rId2"),
            "expected rId1 to be remapped due to conflict, got map: {id_map:?}"
        );

        let rels =
            parse_relationships(&updated, "xl/_rels/workbook.xml.rels").expect("parse rels");
        assert_eq!(rels.len(), 2);
        assert!(rels.iter().any(|r| {
            r.id == "rId1" && r.type_ == "typeA" && r.target == "worksheets/sheet2.xml"
        }));
        assert!(rels.iter().any(|r| {
            r.id == "rId2" && r.type_ == "typeB" && r.target == "worksheets/sheet1.xml"
        }));
    }
}
