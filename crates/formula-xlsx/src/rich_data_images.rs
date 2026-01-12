use std::collections::HashMap;

use formula_model::CellRef;
use roxmltree::{Document, Node};

use crate::openxml;
use crate::path::{rels_for_part, resolve_target};
use crate::{XlsxError, XlsxPackage};

const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const METADATA_PART: &str = "xl/metadata.xml";
const RICH_VALUE_REL_PART: &str = "xl/richData/richValueRel.xml";
const XLRICHVALUE_TYPE_NAME: &str = "XLRICHVALUE";

#[derive(Debug, Clone)]
struct MetadataType {
    name: String,
}

#[derive(Debug, Clone, Default)]
struct FutureMetadataBk {
    /// All `rvb @i` values within the `<bk>` subtree, in document order.
    rvb_indices: Vec<u32>,
}

#[derive(Debug, Clone, Default)]
struct ParsedMetadata {
    /// `valueMetadata` bk index (vm) -> rich value index.
    vm_to_rich_value: HashMap<u32, u32>,
}

impl XlsxPackage {
    /// Extract in-cell images stored via the Excel `richData`/`richValue` mechanism.
    ///
    /// This is a best-effort extractor:
    /// - Missing optional parts result in an empty map rather than an error.
    /// - Malformed indices are skipped.
    ///
    /// Returns a map keyed by `(sheet_name, cell)` containing the referenced image bytes.
    pub fn extract_rich_data_images(
        &self,
    ) -> Result<HashMap<(String, CellRef), Vec<u8>>, XlsxError> {
        let Some(metadata_bytes) = self.part(METADATA_PART) else {
            return Ok(HashMap::new());
        };
        let Some(rich_value_rel_bytes) = self.part(RICH_VALUE_REL_PART) else {
            return Ok(HashMap::new());
        };

        let metadata_xml = String::from_utf8(metadata_bytes.to_vec())?;
        let rich_value_rel_xml = String::from_utf8(rich_value_rel_bytes.to_vec())?;

        let parsed_metadata = parse_metadata_xml(&metadata_xml)?;
        if parsed_metadata.vm_to_rich_value.is_empty() {
            return Ok(HashMap::new());
        }

        let rich_value_rel_doc = Document::parse(&rich_value_rel_xml)?;
        let rich_value_relationship_ids = parse_rich_value_relationship_ids(&rich_value_rel_doc);
        if rich_value_relationship_ids.is_empty() {
            return Ok(HashMap::new());
        }

        let rels_part = rels_for_part(RICH_VALUE_REL_PART);
        let Some(rels_bytes) = self.part(&rels_part) else {
            return Ok(HashMap::new());
        };
        let rels = openxml::parse_relationships(rels_bytes)?;
        let rel_target_by_id: HashMap<String, String> =
            rels.into_iter().map(|r| (r.id, r.target)).collect();

        let mut out: HashMap<(String, CellRef), Vec<u8>> = HashMap::new();

        for sheet in self.worksheet_parts()? {
            let Some(sheet_bytes) = self.part(&sheet.worksheet_part) else {
                continue;
            };
            let sheet_xml = String::from_utf8(sheet_bytes.to_vec())?;
            let doc = match Document::parse(&sheet_xml) {
                Ok(doc) => doc,
                Err(_) => continue,
            };

            for cell_node in doc
                .descendants()
                .filter(|n| n.is_element() && n.tag_name().name() == "c")
            {
                let Some(vm) = cell_node.attribute("vm") else {
                    continue;
                };
                let vm_idx: u32 = match vm.parse() {
                    Ok(v) => v,
                    Err(_) => continue,
                };
                let Some(rich_value_idx) = parsed_metadata.vm_to_rich_value.get(&vm_idx).copied()
                else {
                    continue;
                };

                let Some(rel_id) = rich_value_relationship_ids.get(rich_value_idx as usize) else {
                    continue;
                };

                let Some(target) = rel_target_by_id.get(rel_id) else {
                    continue;
                };
                let target_part = resolve_target(RICH_VALUE_REL_PART, target);
                let Some(image_bytes) = self.part(&target_part) else {
                    continue;
                };

                let Some(cell_ref) = cell_node
                    .attribute("r")
                    .and_then(|a1| CellRef::from_a1(a1).ok())
                else {
                    continue;
                };

                out.insert((sheet.name.clone(), cell_ref), image_bytes.to_vec());
            }
        }

        Ok(out)
    }
}

fn parse_metadata_xml(xml: &str) -> Result<ParsedMetadata, XlsxError> {
    let doc = Document::parse(xml)?;

    let metadata_types = parse_metadata_types(&doc);
    let future_metadata = parse_future_metadata(&doc);

    // Prefer the "metadataTypes/futureMetadata" indirection when it is available. Otherwise fall
    // back to the simpler "extLst-embedded rvb list" strategy.
    let use_future_metadata = !metadata_types.is_empty()
        && future_metadata
            .keys()
            .any(|k| k.eq_ignore_ascii_case(XLRICHVALUE_TYPE_NAME));

    if use_future_metadata {
        parse_value_metadata_using_future_metadata(&doc, &metadata_types, &future_metadata)
    } else {
        parse_value_metadata_using_extlst_bindings(&doc)
    }
}

fn parse_metadata_types(doc: &Document<'_>) -> Vec<MetadataType> {
    let mut out = Vec::new();
    let Some(node) = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "metadataTypes")
    else {
        return out;
    };

    for child in node
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "metadataType")
    {
        let Some(name) = child.attribute("name") else {
            continue;
        };
        out.push(MetadataType {
            name: name.to_string(),
        });
    }

    out
}

fn parse_future_metadata(doc: &Document<'_>) -> HashMap<String, Vec<FutureMetadataBk>> {
    let mut out: HashMap<String, Vec<FutureMetadataBk>> = HashMap::new();

    for fm_node in doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "futureMetadata")
    {
        let Some(name) = fm_node.attribute("name") else {
            continue;
        };

        let mut bks = Vec::new();
        for bk in fm_node
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "bk")
        {
            bks.push(parse_future_metadata_bk(&bk));
        }

        // Some producers wrap `<bk>` entries in a list element. If we didn't see any direct
        // children, scan descendants instead (best-effort).
        if bks.is_empty() {
            for bk in fm_node
                .descendants()
                .filter(|n| n.is_element() && n.tag_name().name() == "bk")
            {
                bks.push(parse_future_metadata_bk(&bk));
            }
        }

        if !bks.is_empty() {
            out.insert(name.to_string(), bks);
        }
    }

    out
}

fn parse_future_metadata_bk(bk: &Node<'_, '_>) -> FutureMetadataBk {
    let mut rvb_indices = Vec::new();
    for rvb in bk
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "rvb")
    {
        let Some(i) = rvb.attribute("i") else {
            continue;
        };
        let Ok(idx) = i.parse::<u32>() else {
            continue;
        };
        rvb_indices.push(idx);
    }

    FutureMetadataBk { rvb_indices }
}

fn parse_value_metadata_using_future_metadata(
    doc: &Document<'_>,
    metadata_types: &[MetadataType],
    future_metadata: &HashMap<String, Vec<FutureMetadataBk>>,
) -> Result<ParsedMetadata, XlsxError> {
    let mut out = ParsedMetadata::default();

    let Some(value_metadata) = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "valueMetadata")
    else {
        return Ok(out);
    };

    // Find the `futureMetadata` payload for XLRICHVALUE (case-insensitive).
    let future_name = future_metadata
        .keys()
        .find(|k| k.eq_ignore_ascii_case(XLRICHVALUE_TYPE_NAME))
        .cloned();
    let Some(future_name) = future_name else {
        return Ok(out);
    };
    let Some(future_bks) = future_metadata.get(&future_name) else {
        return Ok(out);
    };

    // Prefer direct children, but fall back to scanning descendants in case the producer wraps
    // `<bk>` entries in an intermediate list element.
    let mut bk_nodes: Vec<Node<'_, '_>> = value_metadata
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "bk")
        .collect();
    if bk_nodes.is_empty() {
        bk_nodes = value_metadata
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "bk")
            .collect();
    }

    for (vm_idx, bk) in bk_nodes.into_iter().enumerate() {
        let mut rich_value_idx = None;

        for rc in bk
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "rc")
        {
            let Some(t) = rc.attribute("t") else {
                continue;
            };
            let Ok(type_idx) = t.parse::<usize>() else {
                continue;
            };
            // Excel's `t` appears in the wild as both 0-based and 1-based.
            // Prefer 1-based (spec-ish), but fall back to 0-based for compatibility.
            let mut is_rich_type = false;
            for idx in if type_idx == 0 {
                [type_idx, type_idx]
            } else {
                [type_idx - 1, type_idx]
            } {
                if let Some(ty) = metadata_types.get(idx) {
                    if ty.name.eq_ignore_ascii_case(XLRICHVALUE_TYPE_NAME) {
                        is_rich_type = true;
                        break;
                    }
                }
            }
            if !is_rich_type {
                continue;
            }

            let Some(v) = rc.attribute("v") else {
                continue;
            };
            let Ok(future_idx) = v.parse::<usize>() else {
                continue;
            };
            let Some(future_bk) = future_bks.get(future_idx) else {
                continue;
            };
            let Some(i) = future_bk.rvb_indices.first().copied() else {
                continue;
            };

            rich_value_idx = Some(i);
            break;
        }

        if let Some(rich_value_idx) = rich_value_idx {
            out.vm_to_rich_value.insert(vm_idx as u32, rich_value_idx);
        }
    }

    Ok(out)
}

fn parse_value_metadata_using_extlst_bindings(
    doc: &Document<'_>,
) -> Result<ParsedMetadata, XlsxError> {
    let mut out = ParsedMetadata::default();

    let extlst_rvb = parse_extlst_rvb_bindings(doc);
    if extlst_rvb.is_empty() {
        return Ok(out);
    }

    let Some(value_metadata) = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "valueMetadata")
    else {
        return Ok(out);
    };

    // Prefer direct children, but fall back to scanning descendants in case the producer wraps
    // `<bk>` entries in an intermediate list element.
    let mut bk_nodes: Vec<Node<'_, '_>> = value_metadata
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "bk")
        .collect();
    if bk_nodes.is_empty() {
        bk_nodes = value_metadata
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "bk")
            .collect();
    }

    for (vm_idx, bk) in bk_nodes.into_iter().enumerate() {
        let Some(binding_idx) = bk
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "rc")
            .find_map(|rc| rc.attribute("v"))
            .and_then(|v| v.parse::<usize>().ok())
        else {
            continue;
        };

        let Some(rich_value_idx) = extlst_rvb.get(binding_idx).copied() else {
            continue;
        };

        out.vm_to_rich_value.insert(vm_idx as u32, rich_value_idx);
    }

    Ok(out)
}

fn parse_extlst_rvb_bindings(doc: &Document<'_>) -> Vec<u32> {
    let mut rvb = Vec::new();

    for extlst in doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "extLst")
    {
        for node in extlst
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "rvb")
        {
            let Some(i) = node.attribute("i") else {
                continue;
            };
            let Ok(idx) = i.parse::<u32>() else {
                continue;
            };
            rvb.push(idx);
        }
    }

    rvb
}

fn parse_rich_value_relationship_ids(doc: &Document<'_>) -> Vec<String> {
    let mut out = Vec::new();
    for node in doc.descendants().filter(|n| n.is_element()) {
        let id = node
            .attribute((REL_NS, "id"))
            .or_else(|| node.attribute("r:id"));
        let Some(id) = id else {
            continue;
        };
        out.push(id.to_string());
    }
    out
}
