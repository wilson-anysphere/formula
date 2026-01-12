//! RichData helpers for images-in-cells.
//!
//! The end-to-end wiring for images stored as rich values is:
//! `richValue.xml` (rich values) -> relationship-index -> `richValueRel.xml` (index -> rId) ->
//! `richValueRel.xml.rels` (rId -> Target) -> `xl/media/*`.

use std::collections::{BTreeMap, HashMap};

use crate::XlsxError;

use super::rich_value::parse_rich_value_relationship_indices;
use super::rich_value_rel::parse_rich_value_rel_table;

const RICH_VALUE_XML: &str = "xl/richData/richValue.xml";
const RICH_VALUE_REL_XML: &str = "xl/richData/richValueRel.xml";

/// Resolve workbook rich value indices to image target part paths (`xl/media/*`) when possible.
///
/// The returned vector uses:
/// - index: rich value index (0-based, corresponds to `xl/richData/richValue.xml` record order)
/// - value: resolved target part path (e.g. `xl/media/image1.png`) for image-rich-values, or `None`
///   if the record does not appear to reference an image.
///
/// This is intentionally best-effort:
/// - Missing parts return `Ok(Vec::new())` (if `richValue.xml` is missing) or a `Vec` of `None`s
///   (if the rich value table is present but supporting tables are missing).
/// - Unknown namespaces / extra elements are ignored.
/// - Individual records that don't match known shapes are treated as `None`.
pub fn resolve_rich_value_image_targets(
    parts: &BTreeMap<String, Vec<u8>>,
) -> Result<Vec<Option<String>>, XlsxError> {
    let Some(rich_value_xml) = parts.get(RICH_VALUE_XML) else {
        return Ok(Vec::new());
    };

    let rel_indices = parse_rich_value_relationship_indices(rich_value_xml)?;
    if rel_indices.is_empty() {
        return Ok(Vec::new());
    }

    let Some(rich_value_rel_xml) = parts.get(RICH_VALUE_REL_XML) else {
        return Ok(vec![None; rel_indices.len()]);
    };
    let rel_id_table = parse_rich_value_rel_table(rich_value_rel_xml)?;

    // Resolve relationship IDs (`rId*`) to concrete targets via the `.rels` part.
    let rels_part_name = crate::openxml::rels_part_name(RICH_VALUE_REL_XML);
    let Some(rich_value_rel_rels) = parts.get(&rels_part_name) else {
        return Ok(vec![None; rel_indices.len()]);
    };

    let relationships = crate::openxml::parse_relationships(rich_value_rel_rels)?;
    let mut targets_by_id: HashMap<String, String> = HashMap::with_capacity(relationships.len());
    for rel in relationships {
        if rel
            .target_mode
            .as_deref()
            .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
        {
            continue;
        }
        let target = crate::path::resolve_target(RICH_VALUE_REL_XML, &rel.target);
        targets_by_id.insert(rel.id, target);
    }

    let mut out = Vec::with_capacity(rel_indices.len());
    for rel_idx in rel_indices {
        let Some(rel_idx) = rel_idx else {
            out.push(None);
            continue;
        };

        let Some(r_id) = rel_id_table.get(rel_idx) else {
            out.push(None);
            continue;
        };
        if r_id.is_empty() {
            out.push(None);
            continue;
        }

        out.push(targets_by_id.get(r_id).cloned());
    }

    Ok(out)
}

#[cfg(test)]
mod tests {
    use pretty_assertions::assert_eq;
    use std::collections::BTreeMap;

    use super::resolve_rich_value_image_targets;

    #[test]
    fn resolves_image_targets_end_to_end() {
        let rich_value_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <values>
    <rv type="0">
      <v kind="rel">0</v>
      <v kind="string">Alt</v>
    </rv>
    <rv type="0">
      <v kind="string">No image</v>
    </rv>
    <rv type="0">
      <v kind="rel">1</v>
    </rv>
  </values>
</rvData>"#;

        let rich_value_rel_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvRel xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rels>
    <rel r:id="rId7"/>
    <rel r:id="rId8"/>
  </rels>
</rvRel>"#;

        let rels_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
  <Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.jpg"/>
</Relationships>"#;

        let mut parts: BTreeMap<String, Vec<u8>> = BTreeMap::new();
        parts.insert("xl/richData/richValue.xml".to_string(), rich_value_xml.to_vec());
        parts.insert(
            "xl/richData/richValueRel.xml".to_string(),
            rich_value_rel_xml.to_vec(),
        );
        parts.insert(
            "xl/richData/_rels/richValueRel.xml.rels".to_string(),
            rels_xml.to_vec(),
        );

        let resolved = resolve_rich_value_image_targets(&parts).expect("resolve");
        assert_eq!(
            resolved,
            vec![
                Some("xl/media/image1.png".to_string()),
                None,
                Some("xl/media/image2.jpg".to_string())
            ]
        );
    }

    #[test]
    fn missing_supporting_parts_returns_nones() {
        let rich_value_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <values>
    <rv><v kind="rel">0</v></rv>
  </values>
</rvData>"#;

        let mut parts: BTreeMap<String, Vec<u8>> = BTreeMap::new();
        parts.insert("xl/richData/richValue.xml".to_string(), rich_value_xml.to_vec());

        let resolved = resolve_rich_value_image_targets(&parts).expect("resolve");
        assert_eq!(resolved, vec![None]);
    }
}

