//! Best-effort parsing for `xl/richData/richValue.xml`.
//!
//! The full RichData schema is not yet implemented in this repo; this parser focuses on extracting
//! the minimum information needed for images-in-cells:
//! - a rich value record may contain a "relationship index" payload (0-based integer) that points
//!   into `xl/richData/richValueRel.xml`.

use roxmltree::Document;

use crate::XlsxError;

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
    for rv in doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name().eq_ignore_ascii_case("rv"))
    {
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
        if !is_rel_kind(&v) {
            continue;
        }

        if let Some(idx) = parse_int_payload(&v) {
            return Some(idx);
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
        if matches!(
            local.to_ascii_lowercase().as_str(),
            "kind" | "k" | "t" | "type"
        ) && attr.value().eq_ignore_ascii_case("rel")
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
        if matches!(
            local.to_ascii_lowercase().as_str(),
            "v" | "val" | "value" | "i" | "idx" | "index"
        ) {
            if let Ok(v) = attr.value().trim().parse::<usize>() {
                return Some(v);
            }
        }
    }

    None
}

#[cfg(test)]
mod tests {
    use pretty_assertions::assert_eq;

    use super::parse_rich_value_relationship_indices;

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
}

