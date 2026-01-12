//! `xl/metadata.xml` parsing for Excel rich values.
//!
//! Excel's rich data types (including images-in-cell) store cell values indirectly:
//!
//! 1. Worksheet cells (`xl/worksheets/sheet*.xml`) reference a *value-metadata* record with
//!    `c/@vm` (1-based).
//! 2. `xl/metadata.xml` contains `<valueMetadata>` with a list of `<bk>` records. The `vm` value is
//!    a 1-based index into this list.
//! 3. Each `<valueMetadata><bk>` contains `<rc t="T" v="V"/>` where `t` is the 1-based index of
//!    `XLRICHVALUE` in `<metadataTypes>` (Excel has been observed to emit both 0-based and 1-based
//!    indices here), and `v` is a 0-based index into
//!    `<futureMetadata name="XLRICHVALUE">`'s `<bk>` list.
//! 4. Each `<futureMetadata><bk>` contains an extension element (commonly `xlrd:rvb`) with an
//!    `i="N"` attribute. This is the 0-based index into `xl/richData/richValue.xml`.
//!
//! This module resolves that chain and returns a `HashMap<vm, rich_value_index>`.

use std::collections::HashMap;

use roxmltree::Document;

use crate::xml::XmlDomError;

/// Parse `xl/metadata.xml` and return a mapping from worksheet `c/@vm` indices to rich-value
/// indices (`xl/richData/richValue.xml` records).
///
/// The returned map uses:
/// - key: `vm` (1-based index into `<valueMetadata>` `<bk>` records)
/// - value: rich value index (`rvb/@i`, 0-based index into `xl/richData/richValue.xml`)
///
/// This function is intentionally best-effort: if any intermediate linkage is missing for a given
/// `vm` entry (unknown metadata type index, out-of-bounds `v`, missing `rvb/@i`, etc.), that entry
/// is skipped. The only hard errors are invalid UTF-8 or invalid XML.
pub fn parse_value_metadata_vm_to_rich_value_index_map(
    metadata_xml: &[u8],
) -> Result<HashMap<u32, u32>, XmlDomError> {
    let xml = std::str::from_utf8(metadata_xml)?;
    let doc = Document::parse(xml)?;

    let Some(xlrichvalue_type_idx) = find_metadata_type_index(&doc, "XLRICHVALUE") else {
        return Ok(HashMap::new());
    };

    let future_bk_indices = parse_future_rich_value_indices(&doc, "XLRICHVALUE");
    if future_bk_indices.is_empty() {
        // Without the futureMetadata mapping we can't resolve any vm->rv index.
        return Ok(HashMap::new());
    }

    // Excel has been observed to encode `rc/@t` as either 0-based or 1-based indices into the
    // `<metadataTypes>` list. Accept both for robustness.
    let Ok(xlrichvalue_t_zero_based) = u32::try_from(xlrichvalue_type_idx) else {
        return Ok(HashMap::new());
    };
    let xlrichvalue_t_one_based = xlrichvalue_t_zero_based.saturating_add(1);

    Ok(parse_value_metadata_mappings(
        &doc,
        xlrichvalue_t_zero_based,
        xlrichvalue_t_one_based,
        &future_bk_indices,
    ))
}

fn find_metadata_type_index(doc: &Document<'_>, name: &str) -> Option<usize> {
    let metadata_types = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "metadataTypes")?;

    for (idx, node) in metadata_types
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "metadataType")
        .enumerate()
    {
        let Some(mt_name) = node.attribute("name") else {
            continue;
        };

        if mt_name.eq_ignore_ascii_case(name) {
            return Some(idx);
        }
    }

    None
}

fn parse_future_rich_value_indices(doc: &Document<'_>, name: &str) -> Vec<Option<u32>> {
    let future_metadata = doc.descendants().find(|n| {
        n.is_element()
            && n.tag_name().name() == "futureMetadata"
            && n.attribute("name").is_some_and(|n| n.eq_ignore_ascii_case(name))
    });

    let Some(future_metadata) = future_metadata else {
        return Vec::new();
    };

    future_metadata
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "bk")
        .map(|bk| {
            // Prefix/namespace can vary (`xlrd:rvb`, `rvb`, etc.). Match on local-name only.
            let rvb = bk
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "rvb")?;
            let i = rvb.attribute("i")?;
            i.parse::<u32>().ok()
        })
        .collect()
}

fn parse_value_metadata_mappings(
    doc: &Document<'_>,
    xlrichvalue_t_zero_based: u32,
    xlrichvalue_t_one_based: u32,
    future_bk_indices: &[Option<u32>],
) -> HashMap<u32, u32> {
    let Some(value_metadata) = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "valueMetadata")
    else {
        return HashMap::new();
    };

    let mut out = HashMap::new();

    for (bk_idx, bk) in value_metadata
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "bk")
        .enumerate()
    {
        let vm = (bk_idx + 1) as u32;

        let rc = bk.descendants().find(|n| {
            n.is_element()
                && n.tag_name().name() == "rc"
                && n.attribute("t")
                    .and_then(|t| t.parse::<u32>().ok())
                    .is_some_and(|t| t == xlrichvalue_t_zero_based || t == xlrichvalue_t_one_based)
        });

        let Some(rc) = rc else {
            continue;
        };

        let Some(v) = rc
            .attribute("v")
            .and_then(|v| v.parse::<usize>().ok())
            .and_then(|idx| future_bk_indices.get(idx).copied())
            .flatten()
        else {
            continue;
        };

        out.insert(vm, v);
    }

    out
}

#[cfg(test)]
mod tests {
    use pretty_assertions::assert_eq;

    use super::parse_value_metadata_vm_to_rich_value_index_map;

    #[test]
    fn parses_vm_to_rich_value_indices() {
        // Two metadataTypes to ensure `rc/@t` is interpreted as a 1-based index into the
        // `<metadataTypes>` list, not hard-coded to `1`.
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <metadataTypes count="2">
    <metadataType name="SOMEOTHERTYPE"/>
    <metadataType name="XLRICHVALUE"/>
  </metadataTypes>
  <futureMetadata name="XLRICHVALUE" count="2">
    <bk>
      <extLst>
        <ext uri="{00000000-0000-0000-0000-000000000000}">
          <xlrd:rvb i="5"/>
        </ext>
      </extLst>
    </bk>
    <bk>
      <extLst>
        <ext uri="{00000000-0000-0000-0000-000000000001}">
          <xlrd:rvb i="42"/>
        </ext>
      </extLst>
    </bk>
  </futureMetadata>
  <valueMetadata count="4">
    <bk><rc t="2" v="0"/></bk>
    <bk><rc t="2" v="1"/></bk>
    <!-- Wrong metadata type; should be ignored. -->
    <bk><rc t="1" v="0"/></bk>
    <!-- Out-of-bounds v; should be ignored. -->
    <bk><rc t="2" v="2"/></bk>
  </valueMetadata>
</metadata>
"#;

        let map = parse_value_metadata_vm_to_rich_value_index_map(xml.as_bytes()).unwrap();
        assert_eq!(map.get(&1), Some(&5));
        assert_eq!(map.get(&2), Some(&42));
        assert_eq!(map.len(), 2);
    }
}
