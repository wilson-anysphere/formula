//! Parser for `xl/richData/richValueStructure.xml`.
//!
//! This file defines the ordered member layout for rich value structures. Members are stored in a
//! positional list in rich-value instances, so member ordering matters.

use std::collections::{BTreeMap, HashMap};

use roxmltree::{Document, Node};

use crate::{XlsxError, XlsxPackage};

pub type RichValueStructures = HashMap<String, RichValueStructure>;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RichValueStructure {
    pub members: Vec<RichValueStructureMember>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RichValueStructureMember {
    pub name: String,
    pub kind: Option<String>,
    /// Attributes not recognized by this parser (including namespaced attributes).
    pub attributes: BTreeMap<String, String>,
}

const RICH_VALUE_STRUCTURE_XML: &str = "xl/richData/richValueStructure.xml";

/// Read and parse `xl/richData/richValueStructure.xml` from an [`XlsxPackage`].
pub fn parse_rich_value_structure_from_package(
    pkg: &XlsxPackage,
) -> Result<Option<RichValueStructures>, XlsxError> {
    let Some(bytes) = pkg.part(RICH_VALUE_STRUCTURE_XML) else {
        return Ok(None);
    };
    Ok(Some(parse_rich_value_structure_xml(bytes)?))
}

/// Parse `xl/richData/richValueStructure.xml` into a map keyed by structure ID.
///
/// Parsing is namespace-tolerant for element names (matches by local-name).
pub fn parse_rich_value_structure_xml(xml: &[u8]) -> Result<RichValueStructures, XlsxError> {
    let xml = String::from_utf8(xml.to_vec())?;
    let doc = Document::parse(&xml)?;

    let mut out: RichValueStructures = HashMap::new();

    // Typical shape:
    // <rvStruct> <structures> <structure id="..."> <member .../>* </structure>* </structures> </rvStruct>
    //
    // Be tolerant: don't require a specific `<structures>` wrapper; scan for `structure` nodes
    // anywhere in the document.
    for structure_el in doc
        .descendants()
        .filter(|n| n.is_element() && matches_local_name(n.tag_name().name(), &["structure", "s"]))
    {
        let Some(id) = attr_local(structure_el, &["id", "s", "structureId", "structure_id"]) else {
            // Best-effort: ignore malformed/unrecognized <structure> entries.
            continue;
        };

        let mut members = Vec::new();
        for member_el in structure_el
            .descendants()
            .filter(|n| n.is_element() && matches_local_name(n.tag_name().name(), &["member", "m"]))
        {
            // Ensure this member belongs to the current structure (and not a nested structure).
            if member_el
                .ancestors()
                .filter(|n| n.is_element())
                .find(|n| matches_local_name(n.tag_name().name(), &["structure", "s"]))
                .is_some_and(|s| s != structure_el)
            {
                continue;
            }

            let Some(name) = attr_local(member_el, &["name", "n"]) else {
                // Best-effort: ignore malformed/unrecognized <member> entries.
                continue;
            };

            let kind = attr_local(member_el, &["kind", "k", "t", "type"]);
            let attributes =
                collect_unknown_attrs(member_el, &["name", "n", "kind", "k", "t", "type"]);

            members.push(RichValueStructureMember {
                name,
                kind,
                attributes,
            });
        }

        out.insert(id, RichValueStructure { members });
    }

    Ok(out)
}

const XMLNS_NS: &str = "http://www.w3.org/2000/xmlns/";

fn matches_local_name(name: &str, expected: &[&str]) -> bool {
    expected.iter().any(|n| name.eq_ignore_ascii_case(n))
}

fn attr_local(node: Node<'_, '_>, locals: &[&str]) -> Option<String> {
    for attr in node.attributes() {
        let local = attr.name().rsplit(':').next().unwrap_or(attr.name());
        if locals.iter().any(|n| local.eq_ignore_ascii_case(n)) {
            return Some(attr.value().to_string());
        }
    }
    None
}

fn collect_unknown_attrs(node: Node<'_, '_>, known_locals: &[&str]) -> BTreeMap<String, String> {
    let mut out = BTreeMap::new();

    for attr in node.attributes() {
        // Skip namespace declaration attributes.
        if attr.name() == "xmlns" || attr.namespace() == Some(XMLNS_NS) {
            continue;
        }

        let local = attr.name().rsplit(':').next().unwrap_or(attr.name());
        if known_locals.iter().any(|k| local.eq_ignore_ascii_case(k)) {
            continue;
        }

        let key = match attr.namespace() {
            Some(ns) => format!("{{{ns}}}{}", attr.name()),
            None => attr.name().to_string(),
        };

        out.insert(key, attr.value().to_string());
    }

    out
}

#[cfg(test)]
mod tests {
    use std::collections::{BTreeMap, HashMap};

    use pretty_assertions::assert_eq;

    use super::{parse_rich_value_structure_xml, RichValueStructure, RichValueStructureMember};

    #[test]
    fn parses_rich_value_structure() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rd:rvStruct xmlns:rd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <rd:structures>
    <rd:structure id="s_image">
      <rd:member name="imageRel" kind="rel" extra="1"/>
      <rd:member name="altText" kind="string"/>
      <!-- Missing name should be ignored. -->
      <rd:member kind="number"/>
    </rd:structure>
    <rd:structure id="s_empty"/>
    <!-- Missing id should be ignored. -->
    <rd:structure>
      <rd:member name="ignored"/>
    </rd:structure>
  </rd:structures>
</rd:rvStruct>"#;

        let structures = parse_rich_value_structure_xml(xml.as_bytes()).unwrap();
        assert_eq!(
            structures,
            HashMap::from([
                (
                    "s_image".to_string(),
                    RichValueStructure {
                        members: vec![
                            RichValueStructureMember {
                                name: "imageRel".to_string(),
                                kind: Some("rel".to_string()),
                                attributes: BTreeMap::from([(
                                    "extra".to_string(),
                                    "1".to_string()
                                )]),
                            },
                            RichValueStructureMember {
                                name: "altText".to_string(),
                                kind: Some("string".to_string()),
                                attributes: BTreeMap::new(),
                            }
                        ],
                    }
                ),
                (
                    "s_empty".to_string(),
                    RichValueStructure { members: vec![] }
                ),
            ])
        );
    }
}
