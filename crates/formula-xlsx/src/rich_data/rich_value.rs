//! Best-effort parsing for `xl/richData/richValue.xml`.
//!
//! The full RichData schema is not yet implemented in this repo; this parser focuses on extracting
//! the minimum information needed for images-in-cells:
//! - a rich value record may contain a "relationship index" payload (0-based integer) that points
//!   into `xl/richData/richValueRel.xml`.

use roxmltree::Document;

use crate::{XlsxError, XlsxPackage};

/// Conventional part name for `xl/richData/richValue.xml`.
pub const RICH_VALUE_XML: &str = "xl/richData/richValue.xml";
/// Some producers use the pluralized `xl/richData/richValues.xml` naming pattern.
pub const RICH_VALUES_XML: &str = "xl/richData/richValues.xml";

/// Parse `xl/richData/richValue.xml` and return a vector where each entry corresponds to a rich
/// value record (0-based index).
///
/// For each record, we attempt to extract a "relationship index" integer, using a best-effort
/// heuristic:
/// - find the first descendant element (commonly `<v>`) with an attribute like `kind="rel"`
/// - parse its text content (or a value-like attribute) as an integer
///
/// If a record does not contain a recognizable relationship reference, its entry is `None`.
pub fn parse_rich_value_relationship_indices(
    xml_bytes: &[u8],
) -> Result<Vec<Option<usize>>, XlsxError> {
    let xml = std::str::from_utf8(xml_bytes)
        .map_err(|e| XlsxError::Invalid(format!("richValue.xml not utf-8: {e}")))?;
    let doc = Document::parse(xml)?;

    let mut out = Vec::new();
    // Typical shape:
    // <rvData> <values> <rv>...</rv>* </values> </rvData>
    //
    // Be tolerant: allow wrapper nodes under `<values>`. If there is a `<values>` container, prefer
    // scanning within it to avoid false-positive `<rv>` matches elsewhere in the document.
    let values_root = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name().eq_ignore_ascii_case("values"));
    let rv_nodes: Vec<roxmltree::Node<'_, '_>> = match values_root {
        Some(values) => values
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name().eq_ignore_ascii_case("rv"))
            // Avoid treating nested `<rv>` blocks as separate records. The rich value schema uses a
            // flat list of records under `<values>`.
            .filter(|rv| {
                !rv.ancestors()
                    .skip(1)
                    .filter(|n| n.is_element())
                    .any(|n| n.tag_name().name().eq_ignore_ascii_case("rv"))
            })
            .collect(),
        None => doc
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name().eq_ignore_ascii_case("rv"))
            .filter(|rv| {
                !rv.ancestors()
                    .skip(1)
                    .filter(|n| n.is_element())
                    .any(|n| n.tag_name().name().eq_ignore_ascii_case("rv"))
            })
            .collect(),
    };

    for rv in rv_nodes {
        out.push(parse_rv_relationship_index(&rv));
    }

    Ok(out)
}

fn parse_rv_relationship_index(rv: &roxmltree::Node<'_, '_>) -> Option<usize> {
    // Look for the first `<v ...>INT</v>` where the attribute value indicates a relationship index.
    for v in rv
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name().eq_ignore_ascii_case("v"))
    {
        // Ensure `v` belongs to this `rv` (and not a nested one).
        if v.ancestors()
            .filter(|n| n.is_element())
            .find(|n| n.tag_name().name().eq_ignore_ascii_case("rv"))
            .is_some_and(|closest_rv| closest_rv != *rv)
        {
            continue;
        }

        if !is_rel_kind(&v) {
            continue;
        }

        if let Some(idx) = parse_int_payload(&v) {
            return Some(idx);
        }
    }

    None
}

/// Parsed representation of `xl/richData/richValue.xml`.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct RichValues {
    pub values: Vec<RichValueInstance>,
}

/// A single `<rv>` entry.
#[derive(Debug, Clone, PartialEq)]
pub struct RichValueInstance {
    /// Best-effort parsed `type=` / `t=` attribute.
    pub type_id: Option<u32>,
    /// Best-effort raw `s=` / `structure=` attribute (if present).
    pub structure_id: Option<String>,
    /// Ordered field values discovered under this `<rv>` element.
    pub fields: Vec<RichValueFieldValue>,
}

/// A single field value (commonly a `<v>` element) under a `<rv>` entry.
#[derive(Debug, Clone, PartialEq)]
pub struct RichValueFieldValue {
    /// Best-effort type/kind discriminator (e.g. `kind="rel"`).
    pub kind: Option<String>,
    /// Raw value text payload.
    pub value: Option<String>,
}

impl RichValues {
    /// Parse `xl/richData/richValue.xml` from an [`XlsxPackage`].
    pub fn from_package(pkg: &XlsxPackage) -> Result<Option<Self>, XlsxError> {
        if let Some(bytes) = pkg.part(RICH_VALUE_XML) {
            return Ok(Some(parse_rich_values_xml(bytes)?));
        }
        if let Some(bytes) = pkg.part(RICH_VALUES_XML) {
            // Some producers use the pluralized naming pattern.
            return Ok(Some(parse_rich_values_xml(bytes)?));
        }
        Ok(None)
    }
}

/// Parse `xl/richData/richValue.xml` into a structured representation.
///
/// This parser is intentionally tolerant:
/// - element/attribute namespaces and prefixes are ignored (matching is done by local name)
/// - unknown nodes/attributes are ignored
/// - `<v>` elements are collected in document order even if wrapped in additional containers
pub fn parse_rich_values_xml(xml_bytes: &[u8]) -> Result<RichValues, XlsxError> {
    let xml = std::str::from_utf8(xml_bytes)
        .map_err(|e| XlsxError::Invalid(format!("richValue.xml not utf-8: {e}")))?;
    let doc = Document::parse(xml)?;

    let mut values = Vec::new();
    // Typical shape:
    // <rvData> <values> <rv>...</rv>* </values> </rvData>
    //
    // Be tolerant: allow wrapper nodes under `<values>`. If there is a `<values>` container, prefer
    // scanning within it to avoid false-positive `<rv>` matches elsewhere in the document.
    let values_root = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name().eq_ignore_ascii_case("values"));
    let rv_nodes: Vec<roxmltree::Node<'_, '_>> = match values_root {
        Some(values) => values
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name().eq_ignore_ascii_case("rv"))
            // Avoid treating nested `<rv>` blocks as separate records. The rich value schema uses a
            // flat list of records under `<values>`.
            .filter(|rv| {
                !rv.ancestors()
                    .skip(1)
                    .filter(|n| n.is_element())
                    .any(|n| n.tag_name().name().eq_ignore_ascii_case("rv"))
            })
            .collect(),
        None => doc
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name().eq_ignore_ascii_case("rv"))
            .filter(|rv| {
                !rv.ancestors()
                    .skip(1)
                    .filter(|n| n.is_element())
                    .any(|n| n.tag_name().name().eq_ignore_ascii_case("rv"))
            })
            .collect(),
    };

    for rv in rv_nodes {
        // Rich value instances typically use `t`/`type` to reference a type definition from
        // `richValueTypes.xml`.
        //
        // Some producers also use `id`/`idx` to assign a *global rich value index* (used when the
        // rich value store is split across multiple `richValue*.xml` parts). Do not treat those as
        // type IDs.
        let type_id = attr_local(&rv, &["t", "type", "typeId", "type_id"])
            .and_then(|v| v.trim().parse::<u32>().ok());
        let structure_id = attr_local(&rv, &["s", "structure", "structureId", "structure_id"]);

        let mut fields = Vec::new();
        for v in rv
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name().eq_ignore_ascii_case("v"))
        {
            // Ensure `v` belongs to this `rv` (and not a nested one).
            if v.ancestors()
                .filter(|n| n.is_element())
                .find(|n| n.tag_name().name().eq_ignore_ascii_case("rv"))
                .is_some_and(|closest_rv| closest_rv != rv)
            {
                continue;
            }

            let kind = attr_local(&v, &["kind", "k", "t", "type"]);
            let value = v
                .text()
                .map(|t| t.to_string())
                .or_else(|| attr_local(&v, &["v", "val", "value"]));

            fields.push(RichValueFieldValue { kind, value });
        }

        values.push(RichValueInstance {
            type_id,
            structure_id,
            fields,
        });
    }

    Ok(RichValues { values })
}

fn attr_local(node: &roxmltree::Node<'_, '_>, names: &[&str]) -> Option<String> {
    for attr in node.attributes() {
        let local = attr.name().rsplit(':').next().unwrap_or(attr.name());
        if names.iter().any(|n| local.eq_ignore_ascii_case(n)) {
            return Some(attr.value().to_string());
        }
    }
    None
}

fn is_rel_kind(node: &roxmltree::Node<'_, '_>) -> bool {
    // Common Excel-emitted attribute is `kind="rel"`, but allow other producers to use a
    // different name as long as the value is `rel`.
    for attr in node.attributes() {
        let name = attr.name();
        let local = name.rsplit(':').next().unwrap_or(name);
        if (local.eq_ignore_ascii_case("kind")
            || local.eq_ignore_ascii_case("k")
            || local.eq_ignore_ascii_case("t")
            || local.eq_ignore_ascii_case("type"))
            && (attr.value().eq_ignore_ascii_case("rel") || attr.value().eq_ignore_ascii_case("r"))
        {
            return true;
        }
    }
    false
}

fn parse_int_payload(node: &roxmltree::Node<'_, '_>) -> Option<usize> {
    if let Some(text) = node.text() {
        if let Ok(v) = text.trim().parse::<usize>() {
            return Some(v);
        }
    }

    // Fallback: some encodings may store the integer as an attribute instead of element text.
    for attr in node.attributes() {
        let name = attr.name();
        let local = name.rsplit(':').next().unwrap_or(name);
        if local.eq_ignore_ascii_case("v")
            || local.eq_ignore_ascii_case("val")
            || local.eq_ignore_ascii_case("value")
            || local.eq_ignore_ascii_case("i")
            || local.eq_ignore_ascii_case("idx")
            || local.eq_ignore_ascii_case("index")
        {
            if let Ok(v) = attr.value().trim().parse::<usize>() {
                return Some(v);
            }
        }
    }

    None
}

#[cfg(test)]
mod tests {
    use std::io::{Cursor, Write};

    use pretty_assertions::assert_eq;
    use zip::write::FileOptions;
    use zip::ZipWriter;

    use super::parse_rich_value_relationship_indices;
    use super::parse_rich_values_xml;
    use super::RichValueFieldValue;
    use super::RichValueInstance;
    use super::RichValues;
    use super::RICH_VALUES_XML;

    fn build_package(entries: &[(&str, &[u8])]) -> super::XlsxPackage {
        let cursor = Cursor::new(Vec::new());
        let mut zip = ZipWriter::new(cursor);
        let options =
            FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

        for (name, bytes) in entries {
            zip.start_file(*name, options).unwrap();
            zip.write_all(bytes).unwrap();
        }

        let bytes = zip.finish().unwrap().into_inner();
        super::XlsxPackage::from_bytes(&bytes).expect("read test pkg")
    }

    #[test]
    fn parses_kind_rel_values() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <values>
    <rv type="0">
      <v kind="rel">0</v>
      <v kind="string">Alt</v>
    </rv>
    <rv type="0">
      <v kind="string">No image</v>
    </rv>
  </values>
</rvData>"#;

        let parsed = parse_rich_value_relationship_indices(xml.as_bytes()).expect("parse");
        assert_eq!(parsed, vec![Some(0), None]);
    }

    #[test]
    fn parses_t_r_values() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <values>
    <rv type="0">
      <v t="r">12</v>
    </rv>
  </values>
</rvData>"#;

        let parsed = parse_rich_value_relationship_indices(xml.as_bytes()).expect("parse");
        assert_eq!(parsed, vec![Some(12)]);
    }

    #[test]
    fn ignores_non_integer_rel_values() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <values>
    <rv type="0">
      <v kind="rel">not-an-int</v>
    </rv>
    <rv type="0">
      <v kind="rel">  12 </v>
    </rv>
  </values>
</rvData>"#;

        let parsed = parse_rich_value_relationship_indices(xml.as_bytes()).expect("parse");
        assert_eq!(parsed, vec![None, Some(12)]);
    }

    #[test]
    fn rich_value_relationship_indices_prefers_values_container_when_present() {
        // Avoid false positives: if a document contains an unrelated `<rv>` outside `<values>`,
        // do not treat it as part of the rich value table.
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <values>
    <rv type="0">
      <v kind="rel">0</v>
    </rv>
  </values>
  <other>
    <rv type="0">
      <v kind="rel">1</v>
    </rv>
  </other>
</rvData>"#;

        let parsed = parse_rich_value_relationship_indices(xml.as_bytes()).expect("parse");
        assert_eq!(parsed, vec![Some(0)]);
    }

    #[test]
    fn rich_value_relationship_indices_ignores_nested_rv_values() {
        // Ensure relationship indices from nested `<rv>` blocks don't shadow the parent `<rv>`.
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <values>
    <rv type="0">
      <wrapper>
        <rv type="0">
          <v kind="rel">7</v>
        </rv>
      </wrapper>
      <v kind="rel">1</v>
    </rv>
  </values>
</rvData>"#;

        let parsed = parse_rich_value_relationship_indices(xml.as_bytes()).expect("parse");
        assert_eq!(parsed, vec![Some(1)]);
    }

    #[test]
    fn rich_values_xml_parses_fields_in_document_order() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rd:rvData xmlns:rd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <rd:values>
    <rd:wrapper>
      <rd:rv t="7" s="s_image">
        <rd:v kind="rel">12</rd:v>
        <rd:v kind="string">Alt</rd:v>
      </rd:rv>
    </rd:wrapper>
  </rd:values>
</rd:rvData>"#;

        let parsed = parse_rich_values_xml(xml.as_bytes()).expect("parse");
        assert_eq!(
            parsed,
            RichValues {
                values: vec![RichValueInstance {
                    type_id: Some(7),
                    structure_id: Some("s_image".to_string()),
                    fields: vec![
                        RichValueFieldValue {
                            kind: Some("rel".to_string()),
                            value: Some("12".to_string()),
                        },
                        RichValueFieldValue {
                            kind: Some("string".to_string()),
                            value: Some("Alt".to_string()),
                        },
                    ],
                }],
            }
        );
    }

    #[test]
    fn rich_values_xml_does_not_treat_id_as_type_id() {
        // Some producers assign a global rich value index via `id`/`idx`. Ensure we don't parse it
        // as the type id.
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <values>
    <rv id="10">
      <v kind="string">Hello</v>
    </rv>
    <rv id="11" t="2">
      <v kind="string">World</v>
    </rv>
  </values>
</rvData>"#;

        let parsed = parse_rich_values_xml(xml.as_bytes()).expect("parse");
        assert_eq!(parsed.values.len(), 2);
        assert_eq!(parsed.values[0].type_id, None);
        assert_eq!(parsed.values[1].type_id, Some(2));
    }

    #[test]
    fn rich_values_xml_ignores_nested_rv_fields() {
        // Ensure `<v>` values inside nested `<rv>` blocks are not attributed to the outer `<rv>`.
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <values>
    <rv t="1">
      <wrapper>
        <rv t="2">
          <v kind="string">inner</v>
        </rv>
      </wrapper>
      <v kind="string">outer</v>
    </rv>
  </values>
</rvData>"#;

        let parsed = parse_rich_values_xml(xml.as_bytes()).expect("parse");
        assert_eq!(parsed.values.len(), 1);
        assert_eq!(
            parsed.values[0].fields,
            vec![RichValueFieldValue {
                kind: Some("string".to_string()),
                value: Some("outer".to_string())
            }]
        );
    }

    #[test]
    fn rich_values_xml_prefers_values_container_when_present() {
        // If a document contains an unrelated `<rv>` element outside the `<values>` container,
        // avoid treating it as part of the rich value table.
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <values>
    <rv t="1">
      <v kind="string">in</v>
    </rv>
  </values>
  <other>
    <rv t="2">
      <v kind="string">out</v>
    </rv>
  </other>
</rvData>"#;

        let parsed = parse_rich_values_xml(xml.as_bytes()).expect("parse");
        assert_eq!(parsed.values.len(), 1);
        assert_eq!(parsed.values[0].type_id, Some(1));
    }

    #[test]
    fn rich_values_from_package_supports_plural_richvalues_part_name() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <values>
    <rv t="1"><v kind="string">x</v></rv>
  </values>
</rvData>"#;

        let pkg = build_package(&[(RICH_VALUES_XML, xml)]);
        let parsed = RichValues::from_package(&pkg)
            .expect("parse")
            .expect("should be present");
        assert_eq!(parsed.values.len(), 1);
        assert_eq!(parsed.values[0].type_id, Some(1));
    }
}
