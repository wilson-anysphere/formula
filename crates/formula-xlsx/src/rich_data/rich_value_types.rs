//! Parser for `xl/richData/richValueTypes.xml`.
//!
//! This file defines the mapping from a numeric rich-value type ID to a structure ID string
//! (defined in `richValueStructure.xml`). Even if the caller does not fully interpret rich-value
//! payloads yet, having access to these tables is useful for debugging and future decoding.

use std::collections::BTreeMap;

use roxmltree::{Document, Node};

use crate::{XlsxError, XlsxPackage};

pub type RichValueTypes = Vec<RichValueType>;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RichValueType {
    pub id: u32,
    pub name: Option<String>,
    pub structure_id: Option<String>,
    /// Attributes not recognized by this parser (including namespaced attributes).
    pub attributes: BTreeMap<String, String>,
}

const RICH_VALUE_TYPES_XML: &str = "xl/richData/richValueTypes.xml";

/// Read and parse `xl/richData/richValueTypes.xml` from an [`XlsxPackage`].
pub fn parse_rich_value_types_from_package(
    pkg: &XlsxPackage,
) -> Result<Option<RichValueTypes>, XlsxError> {
    let Some(bytes) = pkg.part(RICH_VALUE_TYPES_XML) else {
        return Ok(None);
    };
    Ok(Some(parse_rich_value_types_xml(bytes)?))
}

/// Parse `xl/richData/richValueTypes.xml` into a vector of rich value type definitions.
///
/// Parsing is namespace-tolerant for element names (matches by local-name).
pub fn parse_rich_value_types_xml(xml: &[u8]) -> Result<RichValueTypes, XlsxError> {
    let xml = String::from_utf8(xml.to_vec())?;
    let doc = Document::parse(&xml)?;

    let mut out = Vec::new();

    // Typical shape:
    // <rvTypes> <types> <type .../>* </types> </rvTypes>
    //
    // Be tolerant: allow additional wrapper/container nodes under `<types>`.
    // If we cannot find a `<types>` container, fall back to scanning the entire document.
    let mut type_nodes: Vec<Node<'_, '_>> = Vec::new();
    if let Some(types_el) = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name().eq_ignore_ascii_case("types"))
    {
        type_nodes.extend(types_el.descendants().filter(|n| {
            n.is_element()
                && matches_local_name(n.tag_name().name(), &["type", "richValueType", "rvType"])
        }));
    } else {
        type_nodes.extend(doc.descendants().filter(|n| {
            n.is_element()
                && matches_local_name(n.tag_name().name(), &["type", "richValueType", "rvType"])
        }));
    }

    for type_el in type_nodes {
        let id = attr_local(type_el, &["id", "t", "typeId", "type_id"])
            .and_then(|v| v.trim().parse::<u32>().ok());
        let Some(id) = id else {
            // Best-effort: ignore malformed/unrecognized <type> entries.
            continue;
        };

        let name = attr_local(type_el, &["name", "n"]);
        let structure_id = attr_local(type_el, &["structure", "s", "structureId", "structure_id"]);
        let attributes = collect_unknown_attrs(
            type_el,
            &[
                "id",
                "t",
                "typeId",
                "type_id",
                "name",
                "n",
                "structure",
                "s",
                "structureId",
                "structure_id",
            ],
        );

        out.push(RichValueType {
            id,
            name,
            structure_id,
            attributes,
        });
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
    use std::collections::BTreeMap;

    use pretty_assertions::assert_eq;

    use super::{parse_rich_value_types_xml, RichValueType};

    #[test]
    fn parses_rich_value_types() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rd:rvTypes xmlns:rd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rd:types>
    <rd:type id="0"
             name="com.microsoft.excel.image"
             structure="s_image"
             custom="x"
             r:ext="y"/>
    <rd:type id="1" structure="s_other"/>
  </rd:types>
</rd:rvTypes>"#;

        let types = parse_rich_value_types_xml(xml.as_bytes()).unwrap();
        assert_eq!(
            types,
            vec![
                RichValueType {
                    id: 0,
                    name: Some("com.microsoft.excel.image".to_string()),
                    structure_id: Some("s_image".to_string()),
                    attributes: BTreeMap::from([
                        ("custom".to_string(), "x".to_string()),
                        (
                            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}ext"
                                .to_string(),
                            "y".to_string()
                        ),
                    ]),
                },
                RichValueType {
                    id: 1,
                    name: None,
                    structure_id: Some("s_other".to_string()),
                    attributes: BTreeMap::new(),
                }
            ]
        );
    }
}
