//! Parsers for `xl/richData/richValueRel.xml` (rich-value relationship index table).
//!
//! Excel rich values avoid embedding raw `rId*` strings inside every rich value instance.
//! Instead, `xl/richData/richValue.xml` stores a 0-based integer index into
//! `xl/richData/richValueRel.xml`, which in turn stores `r:id="rIdN"` strings. Those are then
//! resolved using the standard OPC `.rels` part.

use roxmltree::Document;

use crate::XlsxError;

const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// Parse `xl/richData/richValueRel.xml` into a dense table of relationship IDs.
///
/// The returned vector uses:
/// - position: relationship index (0-based)
/// - value: `r:id` string (e.g. `"rId7"`)
///
/// This is intentionally best-effort: `<rel>` entries missing `r:id` are preserved as empty strings
/// to avoid shifting indices.
pub fn parse_rich_value_rel_table(xml_bytes: &[u8]) -> Result<Vec<String>, XlsxError> {
    let xml = std::str::from_utf8(xml_bytes)
        .map_err(|e| XlsxError::Invalid(format!("richValueRel.xml not utf-8: {e}")))?;
    let doc = Document::parse(xml)?;

    let mut out = Vec::new();
    for rel in doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name().eq_ignore_ascii_case("rel"))
    {
        out.push(get_rel_id(&rel).unwrap_or_default());
    }

    Ok(out)
}

fn get_rel_id(node: &roxmltree::Node<'_, '_>) -> Option<String> {
    node.attribute((REL_NS, "id"))
        .or_else(|| node.attribute("r:id"))
        .or_else(|| {
            // Some XML libraries represent namespaced attributes using Clark notation:
            // `{namespace}localname`.
            let clark = format!("{{{REL_NS}}}id");
            node.attribute(clark.as_str())
        })
        .map(|s| s.to_string())
}

#[cfg(test)]
mod tests {
    use pretty_assertions::assert_eq;

    use super::parse_rich_value_rel_table;

    #[test]
    fn parses_rel_table_in_document_order() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvRel xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rels>
    <rel r:id="rId1"/>
    <rel r:id="rId2"/>
  </rels>
</rvRel>"#;

        let parsed = parse_rich_value_rel_table(xml.as_bytes()).expect("parse");
        assert_eq!(parsed, vec!["rId1".to_string(), "rId2".to_string()]);
    }

    #[test]
    fn preserves_missing_ids_as_placeholders() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvRel xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rels>
    <rel/>
    <rel r:id="rId9"/>
  </rels>
</rvRel>"#;

        let parsed = parse_rich_value_rel_table(xml.as_bytes()).expect("parse");
        assert_eq!(parsed, vec!["".to_string(), "rId9".to_string()]);
    }
}

