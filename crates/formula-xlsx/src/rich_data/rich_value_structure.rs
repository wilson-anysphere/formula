//! Parser for `xl/richData/richValueStructure.xml`.
//!
//! This file defines the ordered member layout for rich value structures. Members are stored in a
//! positional list in rich-value instances, so member ordering matters.

use std::collections::{BTreeMap, HashMap};

use roxmltree::{Document, Node};

use crate::XlsxError;

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

/// Parse `xl/richData/richValueStructure.xml` into a map keyed by structure ID.
///
/// Parsing is namespace-tolerant for element names (matches by local-name).
pub fn parse_rich_value_structure_xml(xml: &[u8]) -> Result<RichValueStructures, XlsxError> {
    let xml = String::from_utf8(xml.to_vec())?;
    let doc = Document::parse(&xml)?;

    let mut out: RichValueStructures = HashMap::new();

    // Typical shape:
    // <rvStruct> <structures> <structure id="..."> <member .../>* </structure>* </structures> </rvStruct>
    let Some(structures_el) = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "structures")
    else {
        return Ok(out);
    };

    for structure_el in structures_el
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "structure")
    {
        let Some(id) = attr_no_ns(structure_el, "id").map(|s| s.to_string()) else {
            // Best-effort: ignore malformed/unrecognized <structure> entries.
            continue;
        };

        let mut members = Vec::new();
        for member_el in structure_el
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "member")
        {
            let Some(name) = attr_no_ns(member_el, "name").map(|s| s.to_string()) else {
                // Best-effort: ignore malformed/unrecognized <member> entries.
                continue;
            };

            let kind = attr_no_ns(member_el, "kind").map(|s| s.to_string());
            let attributes = collect_unknown_attrs(member_el, &["name", "kind"]);

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

fn attr_no_ns<'a>(node: Node<'a, 'a>, local: &str) -> Option<&'a str> {
    for attr in node.attributes() {
        if attr.namespace().is_none() && attr.name() == local {
            return Some(attr.value());
        }
    }
    None
}

fn collect_unknown_attrs(node: Node<'_, '_>, known_unqualified: &[&str]) -> BTreeMap<String, String> {
    let mut out = BTreeMap::new();

    for attr in node.attributes() {
        // Skip namespace declaration attributes.
        if attr.name() == "xmlns" || attr.namespace() == Some(XMLNS_NS) {
            continue;
        }

        // Only treat *unqualified* known attrs as "known". Namespaced attributes are preserved,
        // even if their local name matches one of the known fields, since the namespace can be
        // semantically meaningful (e.g. `r:id`).
        if attr.namespace().is_none() && known_unqualified.iter().any(|k| *k == attr.name()) {
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

