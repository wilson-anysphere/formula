//! Parsers for `xl/richData/richValueRel.xml` (rich-value relationship index table).
//!
//! Excel rich values avoid embedding raw `rId*` strings inside every rich value instance.
//! Instead, `xl/richData/richValue.xml` stores a 0-based integer index into
//! `xl/richData/richValueRel.xml`, which in turn stores `r:id="rIdN"` strings. Those are then
//! resolved using the standard OPC `.rels` part.

use roxmltree::Document;

use crate::{XlsxError, XlsxPackage};

const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
/// Conventional part name for `xl/richData/richValueRel.xml`.
pub const RICH_VALUE_REL_XML: &str = "xl/richData/richValueRel.xml";

/// A parsed `xl/richData/richValueRel.xml` part.
///
/// This is a thin wrapper over the dense relationship-id table returned by
/// [`parse_rich_value_rel_table`], with a convenience method for resolving a given `rId*` to its
/// OPC target part.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct RichValueRels {
    /// Ordered relationship IDs (`r:id="rId*"`), where the index corresponds to the 0-based
    /// relationship table index referenced by `richValue.xml`.
    pub r_ids: Vec<String>,
}

impl RichValueRels {
    /// Parse `xl/richData/richValueRel.xml` from an [`XlsxPackage`].
    pub fn from_package(pkg: &XlsxPackage) -> Result<Option<Self>, XlsxError> {
        let Some(bytes) = pkg.part(RICH_VALUE_REL_XML) else {
            return Ok(None);
        };
        Ok(Some(Self::parse(bytes)?))
    }

    /// Parse a `richValueRel.xml` payload.
    pub fn parse(xml_bytes: &[u8]) -> Result<Self, XlsxError> {
        Ok(Self {
            r_ids: parse_rich_value_rel_table(xml_bytes)?,
        })
    }

    /// Resolve the relationship target for `self.r_ids[idx]`.
    ///
    /// This resolves against the OPC relationships for `xl/richData/richValueRel.xml`:
    /// `xl/richData/_rels/richValueRel.xml.rels`.
    ///
    /// Returns `None` if:
    /// - `idx` is out of range
    /// - the entry is empty (`<rel/>`)
    /// - the `.rels` part is missing/malformed
    /// - the relationship is external
    pub fn resolve_target(&self, pkg: &XlsxPackage, idx: usize) -> Option<String> {
        let r_id = self.r_ids.get(idx)?;
        if r_id.is_empty() {
            return None;
        }

        // Use the existing relationship resolver, but strip URI fragments (Excel can emit
        // `Target="...#something"` for rich data).
        let target = crate::openxml::resolve_relationship_target(pkg, RICH_VALUE_REL_XML, r_id)
            .ok()
            .flatten()?;
        let target = strip_fragment(&target);
        if target.is_empty() {
            return None;
        }

        // Some producers emit `Target="media/image1.png"` (relative to `xl/`) rather than the more
        // common `Target="../media/image1.png"` (relative to `xl/richData/`). In that case the
        // standard relationship resolver produces `xl/richData/media/*`, which won't exist in the
        // package. Make a best-effort guess for this case when possible.
        if pkg.part(target).is_none() {
            if let Some(rest) = target.strip_prefix("xl/richData/") {
                if rest.starts_with("media/") {
                    let alt = format!("xl/{rest}");
                    if pkg.part(&alt).is_some() {
                        return Some(alt);
                    }
                } else if rest.starts_with("xl/") {
                    // Another common producer mistake is emitting `Target="xl/..."` without a
                    // leading `/`, which incorrectly resolves relative to `xl/richData/`.
                    if pkg.part(rest).is_some() {
                        return Some(rest.to_string());
                    }
                }
            }
        }

        Some(target.to_string())
    }
}

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
    // Typical shape:
    // <rvRel> <rels> <rel r:id="rId*"/>* </rels> </rvRel>
    //
    // Be tolerant: allow wrappers under `<rels>`, and fall back to scanning the full document if a
    // `<rels>` container isn't present.
    let rels_root = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name().eq_ignore_ascii_case("rels"));
    let rel_nodes = match rels_root {
        Some(rels) => rels
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name().eq_ignore_ascii_case("rel"))
            .collect::<Vec<_>>(),
        None => doc
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name().eq_ignore_ascii_case("rel"))
            .collect::<Vec<_>>(),
    };
    for rel in rel_nodes {
        out.push(get_rel_id(&rel).unwrap_or_default());
    }

    Ok(out)
}

fn get_rel_id(node: &roxmltree::Node<'_, '_>) -> Option<String> {
    // Prefer the correct relationships namespace (`r:id`), but be tolerant and fall back to any
    // attribute whose local name is `id` (ignoring namespace/prefix).
    node.attribute((REL_NS, "id"))
        .or_else(|| node.attribute("r:id"))
        .or_else(|| {
            // Some XML producers may omit namespaces/prefixes entirely.
            node.attribute("id")
        })
        .or_else(|| {
            // Some XML libraries represent namespaced attributes using Clark notation:
            // `{namespace}localname`.
            let clark = format!("{{{REL_NS}}}id");
            node.attribute(clark.as_str())
        })
        .or_else(|| {
            // Absolute fallback: scan by local name.
            node.attributes()
                .find(|attr| {
                    let local = attr.name().rsplit(':').next().unwrap_or(attr.name());
                    local.eq_ignore_ascii_case("id")
                })
                .map(|attr| attr.value())
        })
        .map(|s| s.to_string())
}

fn strip_fragment(target: &str) -> &str {
    target
        .split_once('#')
        .map(|(base, _)| base)
        .unwrap_or(target)
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

    #[test]
    fn parses_unqualified_id_attribute_for_tolerance() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvRel xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <rels>
    <rel id="rId3"/>
  </rels>
</rvRel>"#;

        let parsed = parse_rich_value_rel_table(xml.as_bytes()).expect("parse");
        assert_eq!(parsed, vec!["rId3".to_string()]);
    }

    #[test]
    fn parses_namespaced_id_attribute_with_weird_casing_for_tolerance() {
        // XML is case-sensitive, but be extra tolerant in case a producer emits `Id` instead of
        // `id` on a namespaced attribute.
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvRel xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
       xmlns:foo="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rels>
    <rel foo:Id="rId7"/>
  </rels>
</rvRel>"#;

        let parsed = parse_rich_value_rel_table(xml.as_bytes()).expect("parse");
        assert_eq!(parsed, vec!["rId7".to_string()]);
    }

    #[test]
    fn prefers_rels_container_when_present() {
        // Avoid false positives: if a document contains an unrelated `<rel>` element outside the
        // `<rels>` container, do not treat it as part of the relationship id table.
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvRel xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rels>
    <rel r:id="rId1"/>
  </rels>
  <other>
    <rel r:id="rId2"/>
  </other>
</rvRel>"#;

        let parsed = parse_rich_value_rel_table(xml.as_bytes()).expect("parse");
        assert_eq!(parsed, vec!["rId1".to_string()]);
    }
}
