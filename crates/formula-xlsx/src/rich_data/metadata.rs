//! `xl/metadata.xml` parsing for Excel rich values.
//!
//! Excel's rich data types (including images-in-cell) store cell values indirectly:
//!
//! 1. Worksheet cells (`xl/worksheets/sheet*.xml`) reference a *value-metadata* record with
//!    `c/@vm` (**0-based or 1-based**; both have been observed).
//! 2. `xl/metadata.xml` contains `<valueMetadata>` with a list of `<bk>` records. The `vm` value
//!    selects a `<bk>` record, but the base varies by producer/version (treat it as opaque and
//!    resolve best-effort; many callers try both `vm` and `vm-1`).
//! 3. Each `<valueMetadata><bk>` contains `<rc t="T" v="V"/>` where:
//!    - `t` selects the `XLRICHVALUE` entry in `<metadataTypes>` (**0-based or 1-based** are both
//!      observed in the wild/tests; treat as ambiguous),
//!    - `v` selects an entry in `<futureMetadata name="XLRICHVALUE">`'s `<bk>` list (usually 0-based,
//!      but some producers appear to use 1-based indexing).
//! 4. Each `<futureMetadata><bk>` contains an extension element (commonly `xlrd:rvb`) with an
//!    `i="N"` attribute. This is the **0-based rich value index** into the rich value store
//!    (`xl/richData/richValue*.xml` or `xl/richData/rdrichvalue.xml`, depending on file/Excel build).
//!
//! This module resolves that chain and returns a `HashMap<vm, rich_value_index>`.

use std::collections::HashMap;

use roxmltree::Document;

use crate::xml::XmlDomError;

#[derive(Debug, Clone, Copy)]
struct BkRun<T> {
    count: u32,
    value: T,
}

/// Parse `xl/metadata.xml` and return a mapping from worksheet `c/@vm` indices to rich-value
/// indices (0-based records in the rich value store part, e.g. `xl/richData/richValue*.xml` or
/// `xl/richData/rdrichvalue.xml` depending on workbook).
///
/// The returned map uses:
/// - key: `vm` (the **1-based** `<valueMetadata>` `<bk>` record index; callers should consider
///   `vm-1` for files that use 0-based worksheet `c/@vm` values)
/// - value: rich value index (`rvb/@i`, 0-based index into the rich value store part)
///
/// This function is intentionally best-effort: if any intermediate linkage is missing for a given
/// `vm` entry (unknown metadata type index, out-of-bounds `v`, missing `rvb/@i`, etc.), that entry
/// is skipped. The only hard errors are invalid UTF-8 or invalid XML.
pub fn parse_value_metadata_vm_to_rich_value_index_map(
    metadata_xml: &[u8],
) -> Result<HashMap<u32, u32>, XmlDomError> {
    let xml = std::str::from_utf8(metadata_xml)?;
    let doc = Document::parse(xml)?;

    let Some((xlrichvalue_type_idx, metadata_types_count)) =
        find_metadata_type_index_and_count(&doc, "XLRICHVALUE")
    else {
        return Ok(HashMap::new());
    };

    // Excel has been observed to encode `rc/@t` as either 0-based or 1-based indices into the
    // `<metadataTypes>` list. Most files appear consistent, but some producers can mix the two
    // schemes (e.g. some `<bk>` entries use 1-based while others use 0-based).
    //
    // Accepting both schemes simultaneously can produce false positives when `<metadataTypes>`
    // contains multiple entries (the same `t` number can refer to different types depending on the
    // base). We therefore:
    // - Prefer a single inferred base when possible
    // - Accept both bases when we see strong evidence the file mixes them
    // - Otherwise fall back to "whichever yields more resolved vm -> richValue links" (and as a
    //   last resort, allow a mixed interpretation if it strictly increases how many entries we can
    //   resolve)
    //
    // Note: Some producers can omit the `<futureMetadata name="XLRICHVALUE">` indirection and store
    // the rich value index directly in `rc/@v`. This direct layout is supported best-effort, but is
    // not currently observed in the images-in-cell fixtures checked into this repo.
    let Ok(xlrichvalue_t_zero_based) = u32::try_from(xlrichvalue_type_idx) else {
        return Ok(HashMap::new());
    };
    let Some(xlrichvalue_t_one_based) = xlrichvalue_t_zero_based.checked_add(1) else {
        return Ok(HashMap::new());
    };
    let rc_t_indexing = infer_value_metadata_rc_t_indexing(&doc, metadata_types_count);
    let future_bk_indices = parse_future_rich_value_indices(&doc, "XLRICHVALUE");
    if future_bk_indices.is_empty() {
        // Direct mapping schema (not yet observed in this repo's images-in-cell fixtures):
        //
        // Some producers omit `<futureMetadata name="XLRICHVALUE">` entirely and instead store the
        // rich value index directly in `<valueMetadata><bk><rc ... v="..."/>`.
        let out = match rc_t_indexing {
            RcTIndexing::ZeroBased => {
                parse_value_metadata_direct_mappings(&doc, &[xlrichvalue_t_zero_based])
            }
            RcTIndexing::OneBased => {
                parse_value_metadata_direct_mappings(&doc, &[xlrichvalue_t_one_based])
            }
            RcTIndexing::Mixed => parse_value_metadata_direct_mappings(
                &doc,
                &[xlrichvalue_t_zero_based, xlrichvalue_t_one_based],
            ),
            RcTIndexing::Ambiguous => {
                let a = parse_value_metadata_direct_mappings(&doc, &[xlrichvalue_t_zero_based]);
                let b = parse_value_metadata_direct_mappings(&doc, &[xlrichvalue_t_one_based]);
                // Mixed mode: accept both candidates, but only if it increases how many entries we
                // can resolve.
                let c = parse_value_metadata_direct_mappings(
                    &doc,
                    &[xlrichvalue_t_zero_based, xlrichvalue_t_one_based],
                );
                if c.len() > a.len().max(b.len()) {
                    c
                } else if b.len() >= a.len() {
                    b
                } else {
                    a
                }
            }
        };
        return Ok(out);
    }

    let out = match rc_t_indexing {
        RcTIndexing::ZeroBased => {
            parse_value_metadata_mappings(&doc, &[xlrichvalue_t_zero_based], &future_bk_indices)
        }
        RcTIndexing::OneBased => {
            parse_value_metadata_mappings(&doc, &[xlrichvalue_t_one_based], &future_bk_indices)
        }
        RcTIndexing::Mixed => parse_value_metadata_mappings(
            &doc,
            &[xlrichvalue_t_zero_based, xlrichvalue_t_one_based],
            &future_bk_indices,
        ),
        RcTIndexing::Ambiguous => {
            let a = parse_value_metadata_mappings(&doc, &[xlrichvalue_t_zero_based], &future_bk_indices);
            let b = parse_value_metadata_mappings(&doc, &[xlrichvalue_t_one_based], &future_bk_indices);
            // Mixed mode: accept both candidates, but only if it increases how many entries we can
            // resolve.
            let c = parse_value_metadata_mappings(
                &doc,
                &[xlrichvalue_t_zero_based, xlrichvalue_t_one_based],
                &future_bk_indices,
            );
            if c.len() > a.len().max(b.len()) {
                c
            } else if b.len() >= a.len() {
                b
            } else {
                a
            }
        }
    };

    Ok(out)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum RcTIndexing {
    ZeroBased,
    OneBased,
    Mixed,
    Ambiguous,
}

fn find_metadata_type_index_and_count(doc: &Document<'_>, name: &str) -> Option<(usize, usize)> {
    let metadata_types = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "metadataTypes")?;

    let mut out_idx: Option<usize> = None;
    let mut count = 0usize;
    for (idx, node) in metadata_types
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "metadataType")
        .enumerate()
    {
        count += 1;
        let Some(mt_name) = node.attribute("name") else {
            continue;
        };

        if mt_name.eq_ignore_ascii_case(name) {
            out_idx = Some(idx);
        }
    }

    out_idx.map(|idx| (idx, count))
}

fn infer_value_metadata_rc_t_indexing(
    doc: &Document<'_>,
    metadata_types_count: usize,
) -> RcTIndexing {
    if metadata_types_count == 0 {
        return RcTIndexing::Ambiguous;
    }

    let Some(value_metadata) = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "valueMetadata")
    else {
        return RcTIndexing::Ambiguous;
    };

    // Collect min/max across all `<rc t="...">` to infer whether indices are 0-based (0..count-1)
    // or 1-based (1..count). In some workbooks the ranges overlap (e.g. only 1..count-1), in which
    // case we treat it as ambiguous and let the caller decide.
    let mut min_t: Option<u32> = None;
    let mut max_t: Option<u32> = None;
    for rc in value_metadata
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "rc")
    {
        let Some(t) = rc.attribute("t").and_then(|t| t.parse::<u32>().ok()) else {
            continue;
        };
        min_t = Some(min_t.map(|m| m.min(t)).unwrap_or(t));
        max_t = Some(max_t.map(|m| m.max(t)).unwrap_or(t));
    }

    let Some((min_t, max_t)) = min_t.zip(max_t) else {
        return RcTIndexing::Ambiguous;
    };

    let Ok(count_u32) = u32::try_from(metadata_types_count) else {
        return RcTIndexing::Ambiguous;
    };

    // If we see both 0 and `count` in the same file, that's strong evidence the producer mixed
    // 0-based and 1-based indexing schemes across `<bk>` entries.
    if min_t == 0 && max_t == count_u32 {
        return RcTIndexing::Mixed;
    }

    let zero_based_possible = max_t < count_u32;
    let one_based_possible = min_t >= 1 && max_t <= count_u32;

    match (zero_based_possible, one_based_possible) {
        (true, false) => RcTIndexing::ZeroBased,
        (false, true) => RcTIndexing::OneBased,
        // Either both are possible (range overlap), or neither is possible (corrupt input).
        _ => RcTIndexing::Ambiguous,
    }
}

fn parse_future_rich_value_indices(doc: &Document<'_>, name: &str) -> Vec<BkRun<Option<u32>>> {
    let future_metadata = doc.descendants().find(|n| {
        n.is_element()
            && n.tag_name().name() == "futureMetadata"
            && n.attribute("name")
                .is_some_and(|n| n.eq_ignore_ascii_case(name))
    });

    let Some(future_metadata) = future_metadata else {
        return Vec::new();
    };

    future_metadata
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "bk")
        .map(|bk| {
            let count = bk
                .attribute("count")
                .and_then(|c| c.trim().parse::<u32>().ok())
                .filter(|c| *c >= 1)
                .unwrap_or(1);

            // Prefix/namespace can vary (`xlrd:rvb`, `rvb`, etc.). Match on local-name only.
            let value = bk
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "rvb")
                .and_then(|rvb| rvb.attribute("i"))
                .and_then(|i| i.trim().parse::<u32>().ok());

            BkRun {
                count,
                value,
            }
        })
        .collect()
}

fn parse_value_metadata_mappings(
    doc: &Document<'_>,
    xlrichvalue_t_values: &[u32],
    future_bk_indices: &[BkRun<Option<u32>>],
) -> HashMap<u32, u32> {
    let Some(value_metadata) = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "valueMetadata")
    else {
        return HashMap::new();
    };

    let mut out = HashMap::new();

    let v_indexing_is_one_based =
        infer_value_metadata_rc_v_indexing(value_metadata, xlrichvalue_t_values, future_bk_indices);

    let mut vm_start_1_based: u32 = 1;

    for bk in value_metadata
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "bk")
    {
        let count = bk
            .attribute("count")
            .and_then(|c| c.trim().parse::<u32>().ok())
            .filter(|c| *c >= 1)
            .unwrap_or(1);

        // Some producers emit multiple `<rc>` entries per `<bk>` (one per metadata type). When we
        // allow multiple `rc/@t` candidates (mixed/ambiguous indexing), the first matching `t`
        // might not be the rich-value one. Prefer the first candidate that can be resolved to a
        // valid rich value index.
        let mut resolved_v: Option<u32> = None;
        for rc in bk
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "rc")
        {
            let Some(t) = rc.attribute("t").and_then(|t| t.parse::<u32>().ok()) else {
                continue;
            };
            if !xlrichvalue_t_values.contains(&t) {
                continue;
            }

            let Some(mut v_idx) = rc.attribute("v").and_then(|v| v.parse::<u32>().ok()) else {
                continue;
            };
            if v_indexing_is_one_based {
                let Some(one_based) = v_idx.checked_sub(1) else {
                    continue;
                };
                v_idx = one_based;
            }

            if let Some(v) = resolve_bk_run(future_bk_indices, v_idx).flatten() {
                resolved_v = Some(v);
                break;
            }
        }

        let Some(v) = resolved_v else {
            vm_start_1_based = vm_start_1_based.saturating_add(count);
            continue;
        };

        for offset in 0..count {
            out.insert(vm_start_1_based.saturating_add(offset), v);
        }
        vm_start_1_based = vm_start_1_based.saturating_add(count);
    }

    out
}

fn infer_value_metadata_rc_v_indexing(
    value_metadata: roxmltree::Node<'_, '_>,
    xlrichvalue_t_values: &[u32],
    future_bk_indices: &[BkRun<Option<u32>>],
) -> bool {
    let mut zero_based_matches = 0usize;
    let mut one_based_matches = 0usize;
    let mut saw_zero = false;

    for rc in value_metadata
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "rc")
    {
        let Some(t) = rc.attribute("t").and_then(|t| t.parse::<u32>().ok()) else {
            continue;
        };
        if !xlrichvalue_t_values.contains(&t) {
            continue;
        }

        let Some(v) = rc.attribute("v").and_then(|v| v.parse::<u32>().ok()) else {
            continue;
        };

        if v == 0 {
            saw_zero = true;
        }

        if resolve_bk_run(future_bk_indices, v).flatten().is_some() {
            zero_based_matches += 1;
        }
        if v > 0 && resolve_bk_run(future_bk_indices, v - 1).flatten().is_some() {
            one_based_matches += 1;
        }
    }

    // If we saw any `v="0"` for a rich value entry, treat indices as 0-based. Interpreting the
    // same file as 1-based would make those entries invalid and can also cause false positives
    // by "shifting" out-of-bounds 0-based values back into range.
    if saw_zero {
        return false;
    }

    // Prefer the indexing scheme that yields more resolved links.
    one_based_matches > zero_based_matches
}

fn resolve_bk_run<T: Copy>(runs: &[BkRun<T>], idx: u32) -> Option<T> {
    let mut cursor: u32 = 0;
    for run in runs {
        let count = run.count.max(1);
        let end = cursor.checked_add(count)?;
        if idx < end {
            return Some(run.value);
        }
        cursor = end;
    }
    None
}

fn parse_value_metadata_direct_mappings(
    doc: &Document<'_>,
    xlrichvalue_t_values: &[u32],
) -> HashMap<u32, u32> {
    let Some(value_metadata) = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "valueMetadata")
    else {
        return HashMap::new();
    };

    let mut out = HashMap::new();

    let mut vm_start_1_based: u32 = 1;

    for bk in value_metadata
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "bk")
    {
        let count = bk
            .attribute("count")
            .and_then(|c| c.trim().parse::<u32>().ok())
            .filter(|c| *c >= 1)
            .unwrap_or(1);

        let rc = bk.descendants().find(|n| {
            n.is_element()
                && n.tag_name().name() == "rc"
                && n.attribute("t")
                    .and_then(|t| t.parse::<u32>().ok())
                    .is_some_and(|t| xlrichvalue_t_values.contains(&t))
        });

        let Some(rc) = rc else {
            vm_start_1_based = match vm_start_1_based.checked_add(count) {
                Some(v) => v,
                None => break,
            };
            continue;
        };

        let Some(v) = rc.attribute("v").and_then(|v| v.parse::<u32>().ok()) else {
            vm_start_1_based = match vm_start_1_based.checked_add(count) {
                Some(v) => v,
                None => break,
            };
            continue;
        };

        for offset in 0..count {
            let Some(vm_idx) = vm_start_1_based.checked_add(offset) else {
                break;
            };
            out.insert(vm_idx, v);
        }
        vm_start_1_based = match vm_start_1_based.checked_add(count) {
            Some(v) => v,
            None => break,
        };
    }

    out
}

#[cfg(test)]
mod tests {
    use pretty_assertions::assert_eq;

    use super::parse_value_metadata_vm_to_rich_value_index_map;

    #[test]
    fn parses_vm_to_rich_value_indices_one_based_t() {
        // Two metadataTypes to ensure `rc/@t` is interpreted as an index into the `<metadataTypes>`
        // list (not hard-coded to `1`).
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
    <!-- Out-of-bounds v (even considering 1-based `v` indexing); should be ignored. -->
    <bk><rc t="2" v="3"/></bk>
  </valueMetadata>
</metadata>
"#;

        let map = parse_value_metadata_vm_to_rich_value_index_map(xml.as_bytes()).unwrap();
        assert_eq!(map.get(&1), Some(&5));
        assert_eq!(map.get(&2), Some(&42));
        assert_eq!(map.len(), 2);
    }

    #[test]
    fn parses_vm_to_rich_value_indices_zero_based_t() {
        // Same as above, but with `rc/@t` encoded as a 0-based index into `<metadataTypes>`.
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
    <bk><rc t="1" v="0"/></bk>
    <bk><rc t="1" v="1"/></bk>
    <!-- Wrong metadata type; should be ignored. -->
    <bk><rc t="0" v="0"/></bk>
    <!-- Out-of-bounds v (even considering 1-based `v` indexing); should be ignored. -->
    <bk><rc t="1" v="3"/></bk>
  </valueMetadata>
</metadata>
"#;

        let map = parse_value_metadata_vm_to_rich_value_index_map(xml.as_bytes()).unwrap();
        assert_eq!(map.get(&1), Some(&5));
        assert_eq!(map.get(&2), Some(&42));
        assert_eq!(map.len(), 2);
    }

    #[test]
    fn parses_vm_to_rich_value_indices_mixed_t_indexing() {
        // Some producers appear to mix 0-based and 1-based indexing for `rc/@t` in the same
        // document. When we can detect this case, we should still resolve rich values correctly.
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
  <valueMetadata count="3">
    <!-- `t` uses 1-based indexing. -->
    <bk><rc t="2" v="0"/></bk>
    <!-- `t` uses 0-based indexing. -->
    <bk><rc t="1" v="1"/></bk>
    <!-- Wrong metadata type; should be ignored. -->
    <bk><rc t="0" v="0"/></bk>
  </valueMetadata>
</metadata>
"#;

        let map = parse_value_metadata_vm_to_rich_value_index_map(xml.as_bytes()).unwrap();
        assert_eq!(map.get(&1), Some(&5));
        assert_eq!(map.get(&2), Some(&42));
        assert_eq!(map.len(), 2);
    }

    #[test]
    fn parses_vm_to_rich_value_indices_mixed_t_indexing_multiple_rc_entries() {
        // `<valueMetadata><bk>` can contain multiple `<rc>` entries. When we accept multiple `t`
        // candidates (mixed indexing), we should choose an rc that can be resolved to a rich value
        // index rather than relying on document order.
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <metadataTypes count="3">
    <metadataType name="SOMEOTHERTYPE"/>
    <metadataType name="XLRICHVALUE"/>
    <metadataType name="ANOTHERTYPE"/>
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
  <valueMetadata count="2">
    <bk>
      <!-- Force RcTIndexing::Mixed by ensuring min_t=0 and max_t=count -->
      <rc t="0" v="0"/>
      <!-- `t` matches one of the mixed candidates, but `v` is out-of-bounds and must be ignored. -->
      <rc t="2" v="99"/>
      <!-- This is the correct rich-value rc for this bk. -->
      <rc t="1" v="0"/>
      <rc t="3" v="0"/>
    </bk>
    <bk>
      <rc t="1" v="1"/>
    </bk>
  </valueMetadata>
</metadata>
"#;

        let map = parse_value_metadata_vm_to_rich_value_index_map(xml.as_bytes()).unwrap();
        assert_eq!(map.get(&1), Some(&5));
        assert_eq!(map.get(&2), Some(&42));
        assert_eq!(map.len(), 2);
    }

    #[test]
    fn parses_bk_count_run_length_encoding() {
        // valueMetadata uses bk/@count to compress repeated vm entries, and futureMetadata can do
        // the same for `rc/@v` indices. Ensure we resolve both without assuming 1:1 bk->index.
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <metadataTypes count="1">
    <metadataType name="XLRICHVALUE"/>
  </metadataTypes>
  <futureMetadata name="XLRICHVALUE" count="2">
    <bk count="2">
      <extLst>
        <ext uri="{00000000-0000-0000-0000-000000000000}">
          <xlrd:rvb i="5"/>
        </ext>
      </extLst>
    </bk>
  </futureMetadata>
  <valueMetadata count="3">
    <bk count="3"><rc t="1" v="1"/></bk>
  </valueMetadata>
</metadata>
"#;

        let map = parse_value_metadata_vm_to_rich_value_index_map(xml.as_bytes()).unwrap();
        // vm=2 should map into the repeated <bk count="3"> block.
        assert_eq!(map.get(&2), Some(&5));
        // All three vm indices should map to the same rich value.
        assert_eq!(map.get(&1), Some(&5));
        assert_eq!(map.get(&3), Some(&5));
        assert_eq!(map.len(), 3);
    }

    #[test]
    fn parses_vm_to_rich_value_indices_one_based_v() {
        // Some producers use 1-based indices into the `<futureMetadata name="XLRICHVALUE">` `<bk>`
        // list. This should still resolve to the `rvb/@i` rich value index.
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <metadataTypes count="1">
    <metadataType name="XLRICHVALUE"/>
  </metadataTypes>
  <futureMetadata name="XLRICHVALUE" count="1">
    <bk>
      <extLst>
        <ext uri="{00000000-0000-0000-0000-000000000000}">
          <xlrd:rvb i="7"/>
        </ext>
      </extLst>
    </bk>
  </futureMetadata>
  <valueMetadata count="1">
    <bk><rc t="1" v="1"/></bk>
  </valueMetadata>
</metadata>
"#;

        let map = parse_value_metadata_vm_to_rich_value_index_map(xml.as_bytes()).unwrap();
        assert_eq!(map.get(&1), Some(&7));
    }

    #[test]
    fn parses_direct_rich_value_indices_when_futuremetadata_missing() {
        // Some producers omit `<futureMetadata name="XLRICHVALUE">` entirely and store the rich
        // value index directly in `<valueMetadata><bk><rc ... v="..."/>`.
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <metadataTypes count="1">
    <metadataType name="XLRICHVALUE"/>
  </metadataTypes>
  <valueMetadata count="1">
    <bk><rc t="1" v="42"/></bk>
  </valueMetadata>
</metadata>
"#;

        let map = parse_value_metadata_vm_to_rich_value_index_map(xml.as_bytes()).unwrap();
        assert_eq!(map.get(&1), Some(&42));
        assert_eq!(map.len(), 1);
    }
}
