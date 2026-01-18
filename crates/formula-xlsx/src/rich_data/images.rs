//! RichData helpers for images-in-cells.
//!
//! The end-to-end wiring for images stored as rich values is:
//! `richValue.xml` (rich values) -> relationship-index -> `richValueRel.xml` (index -> rId) ->
//! `richValueRel.xml.rels` (rId -> Target) -> `xl/media/*`.
//!
//! Excel also emits an alternate rich value naming scheme for embedded images-in-cells:
//! - `xl/richData/rdrichvalue.xml` (positional values)
//! - `xl/richData/rdrichvaluestructure.xml` (key ordering; includes `_rvRel:LocalImageIdentifier`)

use std::collections::{BTreeMap, HashMap};

use roxmltree::Document;

use crate::XlsxError;

use super::rel_slot_get;
use super::rich_value::parse_rich_value_relationship_indices;
use super::rich_value_rel::parse_rich_value_rel_table;

const RD_RICH_VALUE_STRUCTURE_XML: &str = "xl/richData/rdrichvaluestructure.xml";
const RICH_VALUE_REL_XML: &str = "xl/richData/richValueRel.xml";

/// Resolve workbook rich value indices to image target part paths (`xl/media/*`) when possible.
///
/// The returned vector uses:
/// - index: rich value index (0-based, corresponds to the record order of the rich value table that
///   is present in the package. When multiple parts are present (e.g. `richValue.xml`,
///   `richValue1.xml`, ...), records are concatenated in numeric-suffix order.
/// - value: resolved target part path (e.g. `xl/media/image1.png`) for image-rich-values, or `None`
///   if the record does not appear to reference an image.
///
/// This is intentionally best-effort:
/// - Missing rich value parts return `Ok(Vec::new())`.
/// - If a rich value table is present but supporting tables are missing, returns a `Vec` of `None`s
///   with the same length as the parsed rich value records.
/// - Unknown namespaces / extra elements are ignored.
/// - Individual records that don't match known shapes are treated as `None`.
pub fn resolve_rich_value_image_targets(
    parts: &BTreeMap<String, Vec<u8>>,
) -> Result<Vec<Option<String>>, XlsxError> {
    let (rich_value_parts, rd_rich_value_parts) = split_rich_value_part_names(parts);

    // Determine which rich-value table to use. Excel has been observed to emit both the modern
    // `richValue*.xml` parts and the alternate `rdrichvalue*.xml` scheme for embedded images.
    // Prefer the modern table when it contains any `<rv>` entries; otherwise fall back to rd.
    let mut rel_indices = if !rich_value_parts.is_empty() {
        parse_rich_value_parts_relationship_indices(parts, &rich_value_parts)?
    } else {
        Vec::new()
    };

    if rel_indices.is_empty() && !rd_rich_value_parts.is_empty() {
        let structure_xml = parts.get(RD_RICH_VALUE_STRUCTURE_XML).map(|b| b.as_slice());
        rel_indices = parse_rdrichvalue_parts_relationship_indices(
            parts,
            &rd_rich_value_parts,
            structure_xml,
        )?;
    }

    if rel_indices.is_empty() {
        return Ok(Vec::new());
    }

    let Some(rich_value_rel_xml) = parts.get(RICH_VALUE_REL_XML) else {
        let mut out: Vec<Option<String>> = Vec::new();
        if out.try_reserve_exact(rel_indices.len()).is_err() {
            return Err(XlsxError::AllocationFailure(
                "resolve_rich_value_image_targets missing richValueRel.xml fallback",
            ));
        }
        out.resize_with(rel_indices.len(), || None);
        return Ok(out);
    };
    let rel_id_table = parse_rich_value_rel_table(rich_value_rel_xml)?;

    // Resolve relationship IDs (`rId*`) to concrete targets via the `.rels` part.
    let rels_part_name = crate::openxml::rels_part_name(RICH_VALUE_REL_XML);
    let Some(rich_value_rel_rels) = parts.get(&rels_part_name) else {
        let mut out: Vec<Option<String>> = Vec::new();
        if out.try_reserve_exact(rel_indices.len()).is_err() {
            return Err(XlsxError::AllocationFailure(
                "resolve_rich_value_image_targets missing .rels fallback",
            ));
        }
        out.resize_with(rel_indices.len(), || None);
        return Ok(out);
    };

    let relationships = crate::openxml::parse_relationships(rich_value_rel_rels)?;
    let mut targets_by_id: HashMap<String, String> = HashMap::new();
    if targets_by_id.try_reserve(relationships.len()).is_err() {
        return Err(XlsxError::AllocationFailure(
            "resolve_rich_value_image_targets targets_by_id",
        ));
    }
    for rel in relationships {
        if rel
            .target_mode
            .as_deref()
            .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
        {
            continue;
        }
        let target = strip_fragment(&rel.target);
        if target.is_empty() {
            continue;
        }
        let target = super::resolve_rich_value_rel_target_part(RICH_VALUE_REL_XML, target);
        // This helper is intended specifically for images-in-cells, which Excel stores under
        // `xl/media/*`. Ignore other relationship targets to avoid incorrectly returning e.g.
        // drawings/hyperlinks as "image" targets.
        if !target.starts_with("xl/media/") {
            continue;
        }
        targets_by_id.insert(rel.id, target);
    }

    let mut out: Vec<Option<String>> = Vec::new();
    if out.try_reserve_exact(rel_indices.len()).is_err() {
        return Err(XlsxError::AllocationFailure("resolve_rich_value_image_targets output"));
    }
    for rel_idx in rel_indices {
        let Some(rel_idx) = rel_idx else {
            out.push(None);
            continue;
        };

        let Some(r_id) = rel_slot_get(&rel_id_table, rel_idx) else {
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

fn split_rich_value_part_names<'a>(
    parts: &'a BTreeMap<String, Vec<u8>>,
) -> (Vec<&'a str>, Vec<&'a str>) {
    let mut rich_value_parts: Vec<&'a str> = Vec::new();
    let mut rd_rich_value_parts: Vec<&'a str> = Vec::new();

    for name in parts.keys() {
        let Some((family, _idx)) = super::parse_rich_value_part_name(name) else {
            continue;
        };
        match family {
            super::RichValuePartFamily::RichValue | super::RichValuePartFamily::RichValues => {
                rich_value_parts.push(name.as_str())
            }
            super::RichValuePartFamily::RdRichValue => rd_rich_value_parts.push(name.as_str()),
        }
    }

    rich_value_parts.sort_by(|a, b| super::cmp_rich_value_parts_by_numeric_suffix(a, b));
    rd_rich_value_parts.sort_by(|a, b| super::cmp_rich_value_parts_by_numeric_suffix(a, b));

    (rich_value_parts, rd_rich_value_parts)
}

fn parse_rich_value_parts_relationship_indices(
    parts: &BTreeMap<String, Vec<u8>>,
    part_names: &[&str],
) -> Result<Vec<Option<usize>>, XlsxError> {
    let mut out: Vec<Option<usize>> = Vec::new();
    for part_name in part_names {
        let Some(bytes) = parts.get(*part_name) else {
            continue;
        };
        out.extend(parse_rich_value_relationship_indices(bytes)?);
    }
    Ok(out)
}

fn parse_rdrichvalue_parts_relationship_indices(
    parts: &BTreeMap<String, Vec<u8>>,
    part_names: &[&str],
    structure_xml_bytes: Option<&[u8]>,
) -> Result<Vec<Option<usize>>, XlsxError> {
    let mut out: Vec<Option<usize>> = Vec::new();
    for part_name in part_names {
        let Some(bytes) = parts.get(*part_name) else {
            continue;
        };
        out.extend(parse_rdrichvalue_relationship_indices(bytes, structure_xml_bytes)?);
    }
    Ok(out)
}

pub(super) fn parse_rdrichvalue_relationship_indices(
    xml_bytes: &[u8],
    structure_xml_bytes: Option<&[u8]>,
) -> Result<Vec<Option<usize>>, XlsxError> {
    let structure_rel_positions = structure_xml_bytes
        .and_then(|bytes| parse_rdrichvaluestructure_local_image_positions(bytes).ok());

    let xml = std::str::from_utf8(xml_bytes)
        .map_err(|e| XlsxError::Invalid(format!("rdrichvalue.xml not utf-8: {e}")))?;
    let doc = Document::parse(xml)?;

    let mut out = Vec::new();
    for rv in doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name().eq_ignore_ascii_case("rv"))
    {
        // Prefer looking up the `_rvRel:LocalImageIdentifier` position via structure metadata.
        let pos = rv
            .attribute("s")
            .and_then(|s| s.parse::<usize>().ok())
            .and_then(|s_idx| {
                structure_rel_positions
                    .as_ref()
                    .and_then(|v| v.get(s_idx).copied())
                    .flatten()
            });

        let values: Vec<roxmltree::Node<'_, '_>> = rv
            .children()
            .filter(|n| n.is_element() && n.tag_name().name().eq_ignore_ascii_case("v"))
            .collect();

        if let Some(pos) = pos.and_then(|p| values.get(p).copied()) {
            let idx = pos.text().and_then(|t| t.trim().parse::<usize>().ok());
            out.push(idx);
            continue;
        }

        // Fallback: assume the first integer payload corresponds to the relationship index.
        let idx = values
            .iter()
            .find_map(|v| v.text().and_then(|t| t.trim().parse::<usize>().ok()));
        out.push(idx);
    }

    Ok(out)
}

fn parse_rdrichvaluestructure_local_image_positions(
    xml_bytes: &[u8],
) -> Result<Vec<Option<usize>>, XlsxError> {
    let xml = std::str::from_utf8(xml_bytes)
        .map_err(|e| XlsxError::Invalid(format!("rdrichvaluestructure.xml not utf-8: {e}")))?;
    let doc = Document::parse(xml)?;

    let mut out = Vec::new();
    for s in doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name().eq_ignore_ascii_case("s"))
    {
        let pos = s
            .children()
            .filter(|n| n.is_element() && n.tag_name().name().eq_ignore_ascii_case("k"))
            .enumerate()
            .find_map(|(idx, k)| {
                let name = k.attribute("n")?;
                // Excel uses `_rvRel:LocalImageIdentifier` for embedded local images.
                // Be tolerant of namespace/prefix changes by matching on suffix.
                name.ends_with("LocalImageIdentifier").then_some(idx)
            });

        out.push(pos);
    }

    Ok(out)
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
        parts.insert(
            "xl/richData/richValue.xml".to_string(),
            rich_value_xml.to_vec(),
        );
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
    fn resolves_image_targets_strip_uri_fragments() {
        let rich_value_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <values>
    <rv type="0">
      <v kind="rel">0</v>
    </rv>
  </values>
</rvData>"#;

        let rich_value_rel_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvRel xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rels>
    <rel r:id="rId7"/>
  </rels>
</rvRel>"#;

        let rels_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png#fragment"/>
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
        assert_eq!(resolved, vec![Some("xl/media/image1.png".to_string())]);
    }

    #[test]
    fn resolves_image_targets_when_relationship_target_is_media_relative() {
        let rich_value_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <values>
    <rv type="0">
      <v kind="rel">0</v>
    </rv>
  </values>
</rvData>"#;

        let rich_value_rel_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvRel xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rels>
    <rel r:id="rId7"/>
  </rels>
</rvRel>"#;

        let rels_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
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
        assert_eq!(resolved, vec![Some("xl/media/image1.png".to_string())]);
    }

    #[test]
    fn resolves_image_targets_when_relationship_target_is_xl_prefixed_without_root_slash() {
        let rich_value_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <values>
    <rv type="0">
      <v kind="rel">0</v>
    </rv>
  </values>
</rvData>"#;

        let rich_value_rel_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvRel xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rels>
    <rel r:id="rId7"/>
  </rels>
</rvRel>"#;

        // Target is missing a leading `/`, so naive resolution would yield `xl/richData/xl/media/...`.
        let rels_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId7"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    Target="xl/media/image1.png#fragment"/>
    Target="xl/media/image1.png#fragment"/>
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
        assert_eq!(resolved, vec![Some("xl/media/image1.png".to_string())]);
    }

    #[test]
    fn resolves_image_targets_when_relationship_target_is_media_relative_with_backslashes() {
        let rich_value_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <values>
    <rv type="0">
      <v kind="rel">0</v>
    </rv>
  </values>
</rvData>"#;

        let rich_value_rel_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvRel xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rels>
    <rel r:id="rId7"/>
  </rels>
</rvRel>"#;

        // Target is relative to `xl/`, not `xl/richData/`, and uses backslashes.
        let rels_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId7"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    Target="media\image1.png#fragment"/>
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
        assert_eq!(resolved, vec![Some("xl/media/image1.png".to_string())]);
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
        parts.insert(
            "xl/richData/richValue.xml".to_string(),
            rich_value_xml.to_vec(),
        );

        let resolved = resolve_rich_value_image_targets(&parts).expect("resolve");
        assert_eq!(resolved, vec![None]);
    }

    #[test]
    fn ignores_non_media_relationship_targets() {
        let rich_value_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <values>
    <rv><v kind="rel">0</v></rv>
  </values>
</rvData>"#;

        let rich_value_rel_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvRel xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rels>
    <rel r:id="rId1"/>
  </rels>
</rvRel>"#;

        // Target is not under `xl/media/*`, so should not be treated as an image target.
        let rels_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>"#;

        let mut parts: BTreeMap<String, Vec<u8>> = BTreeMap::new();
        parts.insert(
            "xl/richData/richValue.xml".to_string(),
            rich_value_xml.to_vec(),
        );
        parts.insert(
            "xl/richData/richValueRel.xml".to_string(),
            rich_value_rel_xml.to_vec(),
        );
        parts.insert(
            "xl/richData/_rels/richValueRel.xml.rels".to_string(),
            rels_xml.to_vec(),
        );

        let resolved = resolve_rich_value_image_targets(&parts).expect("resolve");
        assert_eq!(resolved, vec![None]);
    }

    #[test]
    fn resolves_rdrichvalue_image_targets_end_to_end() {
        // `rdrichvalue.xml` uses positional `<v>` values, with ordering described by
        // `rdrichvaluestructure.xml`. For embedded images-in-cells, the key
        // `_rvRel:LocalImageIdentifier` indexes into `richValueRel.xml`.
        let rd_rich_value_structure = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvStructures xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="1">
  <s t="_localImage">
    <k n="_rvRel:LocalImageIdentifier" t="i"/>
    <k n="CalcOrigin" t="i"/>
  </s>
</rvStructures>"#;

        let rd_rich_value_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="2">
  <rv s="0"><v>0</v><v>5</v></rv>
  <rv s="0"><v>1</v><v>5</v></rv>
</rvData>"#;

        let rich_value_rel_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueRels xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rel r:id="rId1"/>
  <rel r:id="rId2"/>
</richValueRels>"#;

        let rels_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.png"/>
</Relationships>"#;

        let mut parts: BTreeMap<String, Vec<u8>> = BTreeMap::new();
        parts.insert(
            "xl/richData/rdrichvaluestructure.xml".to_string(),
            rd_rich_value_structure.to_vec(),
        );
        parts.insert(
            "xl/richData/rdrichvalue.xml".to_string(),
            rd_rich_value_xml.to_vec(),
        );
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
                Some("xl/media/image2.png".to_string())
            ]
        );
    }

    #[test]
    fn resolves_image_targets_across_multiple_richvalue_parts() {
        // Excel can split rich values across multiple parts (richValue.xml, richValue1.xml, ...).
        // Ensure we concatenate them in numeric suffix order (not lexicographic).
        let rich_value_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <values>
    <rv type="0"><v kind="rel">0</v></rv>
  </values>
</rvData>"#;

        let rich_value2_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <values>
    <rv type="0"><v kind="rel">1</v></rv>
  </values>
</rvData>"#;

        let rich_value10_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <values>
    <rv type="0"><v kind="rel">2</v></rv>
  </values>
</rvData>"#;

        let rich_value_rel_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvRel xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rels>
    <rel r:id="rId1"/>
    <rel r:id="rId2"/>
    <rel r:id="rId3"/>
  </rels>
</rvRel>"#;

        let rels_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.png"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image3.png"/>
</Relationships>"#;

        let mut parts: BTreeMap<String, Vec<u8>> = BTreeMap::new();
        parts.insert("xl/richData/richValue.xml".to_string(), rich_value_xml.to_vec());
        parts.insert("xl/richData/richValue10.xml".to_string(), rich_value10_xml.to_vec());
        parts.insert("xl/richData/richValue2.xml".to_string(), rich_value2_xml.to_vec());
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
                Some("xl/media/image2.png".to_string()),
                Some("xl/media/image3.png".to_string())
            ]
        );
    }

    #[test]
    fn richvalue_xml_empty_but_other_richvalue_parts_present_is_not_treated_as_empty_table() {
        // Some workbooks include an empty `richValue.xml`, but store records in `richValue1.xml`.
        // We should treat the richValue family as present and non-empty (and not return an empty
        // result vector).
        let rich_value_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <values/>
</rvData>"#;

        let rich_value1_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <values>
    <rv type="0"><v kind="rel">0</v></rv>
  </values>
</rvData>"#;

        let rich_value_rel_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvRel xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rels>
    <rel r:id="rId1"/>
  </rels>
</rvRel>"#;

        let rels_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>"#;

        let mut parts: BTreeMap<String, Vec<u8>> = BTreeMap::new();
        parts.insert("xl/richData/richValue.xml".to_string(), rich_value_xml.to_vec());
        parts.insert("xl/richData/richValue1.xml".to_string(), rich_value1_xml.to_vec());
        parts.insert(
            "xl/richData/richValueRel.xml".to_string(),
            rich_value_rel_xml.to_vec(),
        );
        parts.insert(
            "xl/richData/_rels/richValueRel.xml.rels".to_string(),
            rels_xml.to_vec(),
        );

        let resolved = resolve_rich_value_image_targets(&parts).expect("resolve");
        assert_eq!(resolved, vec![Some("xl/media/image1.png".to_string())]);
    }

    #[test]
    fn falls_back_to_rdrichvalue_when_richvalue_table_is_empty() {
        // Some workbooks include `richValue.xml` (possibly empty) but store embedded images via
        // `rdrichvalue.xml`. If the legacy table contains no `<rv>` records, we should fall back to
        // rdRichValue.
        let rich_value_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <values/>
</rvData>"#;

        let rd_rich_value_structure = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvStructures xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="1">
  <s t="_localImage">
    <k n="_rvRel:LocalImageIdentifier" t="i"/>
    <k n="CalcOrigin" t="i"/>
  </s>
</rvStructures>"#;

        let rd_rich_value_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="2">
  <rv s="0"><v>0</v><v>5</v></rv>
  <rv s="0"><v>1</v><v>5</v></rv>
</rvData>"#;

        let rich_value_rel_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueRels xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rel r:id="rId1"/>
  <rel r:id="rId2"/>
</richValueRels>"#;

        let rels_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.png"/>
</Relationships>"#;

        let mut parts: BTreeMap<String, Vec<u8>> = BTreeMap::new();
        parts.insert(
            "xl/richData/richValue.xml".to_string(),
            rich_value_xml.to_vec(),
        );
        parts.insert(
            "xl/richData/rdrichvaluestructure.xml".to_string(),
            rd_rich_value_structure.to_vec(),
        );
        parts.insert(
            "xl/richData/rdrichvalue.xml".to_string(),
            rd_rich_value_xml.to_vec(),
        );
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
                Some("xl/media/image2.png".to_string())
            ]
        );
    }
}
