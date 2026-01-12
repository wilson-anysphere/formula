//! Parser for `xl/richData/richValueTypes.xml`.
//!
//! This file defines the mapping from a numeric rich-value type ID to a structure ID string
//! (defined in `richValueStructure.xml`). Even if the caller does not fully interpret rich-value
//! payloads yet, having access to these tables is useful for debugging and future decoding.

use std::collections::BTreeMap;

use roxmltree::{Document, Node};

use crate::XlsxError;

pub type RichValueTypes = Vec<RichValueType>;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RichValueType {
    pub id: u32,
    pub name: Option<String>,
    pub structure_id: Option<String>,
    /// Attributes not recognized by this parser (including namespaced attributes).
    pub attributes: BTreeMap<String, String>,
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
    let Some(types_el) = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name().eq_ignore_ascii_case("types"))
    else {
        return Ok(out);
    };

    // Use `descendants()` (not `children()`) so we can tolerate additional wrapper/container nodes
    // under `<types>`.
    for type_el in types_el
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name().eq_ignore_ascii_case("type"))
    {
        let id = attr_no_ns(type_el, "id").and_then(|v| v.parse::<u32>().ok());
        let Some(id) = id else {
            // Best-effort: ignore malformed/unrecognized <type> entries.
            continue;
        };

        let name = attr_no_ns(type_el, "name").map(|s| s.to_string());
        let structure_id = attr_no_ns(type_el, "structure").map(|s| s.to_string());
        let attributes = collect_unknown_attrs(type_el, &["id", "name", "structure"]);

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
