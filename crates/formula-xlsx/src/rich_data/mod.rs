//! Parsers/utilities for Excel "rich data" parts.
//!
//! Excel stores cell-level rich values (data types, images-in-cells, etc.) via:
//! - `xl/worksheets/sheet*.xml` `c/@vm` (value-metadata index)
//! - `xl/metadata.xml` `<valueMetadata>` + rich-data extension payloads
//! - `xl/richData/richValue*.xml` (rich-value records)
//!
//! This module exposes best-effort parsing helpers without integrating into `formula-model` yet.

pub mod metadata;
pub mod rich_value_structure;
pub mod rich_value_types;
mod media_parts;
mod rich_value_images;

pub use rich_value_images::{ExtractedRichValueImages, RichValueEntry, RichValueIndex, RichValueWarning};

use std::collections::HashMap;

use formula_model::CellRef;
use thiserror::Error;

use crate::{path, XlsxError, XlsxPackage};

/// Errors returned by rich-data parsing helpers.
///
/// These parsers are intentionally "best effort": missing parts yield empty results, while
/// malformed XML returns an error.
#[derive(Debug, Error)]
pub enum RichDataError {
    #[error(transparent)]
    Xlsx(#[from] XlsxError),

    #[error("xml part {part} is not valid UTF-8: {source}")]
    XmlNonUtf8 {
        part: String,
        #[source]
        source: std::str::Utf8Error,
    },

    #[error("xml parse error in {part}: {source}")]
    XmlParse {
        part: String,
        #[source]
        source: roxmltree::Error,
    },
}

/// Best-effort extraction of "image in cell" rich values.
///
/// This follows the richData chain:
///
/// `worksheet <c vm=...>` → `xl/metadata.xml` → `xl/richData/richValue*.xml` →
/// `xl/richData/richValueRel.xml` → `xl/richData/_rels/richValueRel.xml.rels` → `xl/media/*`.
///
/// Missing parts return `Ok(empty)`. Malformed XML returns an error.
pub fn extract_rich_cell_images(
    pkg: &XlsxPackage,
) -> Result<HashMap<(String, CellRef), Vec<u8>>, RichDataError> {
    // If we can't resolve sheet names to parts, we can't provide stable (sheet name, cell) keys.
    // Treat missing workbook parts as "no richData".
    if pkg.part("xl/workbook.xml").is_none() || pkg.part("xl/_rels/workbook.xml.rels").is_none() {
        return Ok(HashMap::new());
    }

    // The workbook parsing stack can error for malformed workbook.xml; bubble that up.
    let worksheet_parts = pkg.worksheet_parts()?;

    let Some(metadata_bytes) = pkg.part("xl/metadata.xml") else {
        return Ok(HashMap::new());
    };
    let vm_to_rich_value = parse_vm_to_rich_value_index_map(metadata_bytes)?;
    if vm_to_rich_value.is_empty() {
        return Ok(HashMap::new());
    }

    let mut cells_with_rich_value: Vec<(String, CellRef, u32)> = Vec::new();
    for sheet in worksheet_parts {
        let Some(sheet_bytes) = pkg.part(&sheet.worksheet_part) else {
            continue;
        };
        let cells = parse_worksheet_vm_cells(sheet_bytes, &sheet.worksheet_part)?;
        for (cell, vm) in cells {
            let Some(&rich_value_idx) = vm_to_rich_value.get(&vm) else {
                continue;
            };
            cells_with_rich_value.push((sheet.name.clone(), cell, rich_value_idx));
        }
    }
    if cells_with_rich_value.is_empty() {
        return Ok(HashMap::new());
    }

    let rich_value_parts: Vec<&str> = pkg
        .part_names()
        .filter(|name| is_rich_value_part(name))
        .collect();
    if rich_value_parts.is_empty() {
        return Ok(HashMap::new());
    }

    let rich_value_to_rel_index = parse_rich_value_parts_rel_indices(pkg, &rich_value_parts)?;
    if rich_value_to_rel_index.is_empty() {
        return Ok(HashMap::new());
    }

    let Some(rich_value_rel_bytes) = pkg.part("xl/richData/richValueRel.xml") else {
        return Ok(HashMap::new());
    };
    let rel_index_to_rid = parse_rich_value_rel_rids(rich_value_rel_bytes)?;
    if rel_index_to_rid.is_empty() {
        return Ok(HashMap::new());
    }

    let Some(rich_value_rel_rels_bytes) = pkg.part("xl/richData/_rels/richValueRel.xml.rels")
    else {
        return Ok(HashMap::new());
    };
    let rid_to_target = parse_rich_value_rel_rels(rich_value_rel_rels_bytes)?;
    if rid_to_target.is_empty() {
        return Ok(HashMap::new());
    }

    let mut out: HashMap<(String, CellRef), Vec<u8>> = HashMap::new();
    for (sheet_name, cell, rich_value_idx) in cells_with_rich_value {
        let Some(&rel_index) = rich_value_to_rel_index.get(&rich_value_idx) else {
            continue;
        };
        let Some(rid) = rel_index_to_rid.get(rel_index as usize) else {
            continue;
        };
        let Some(target) = rid_to_target.get(rid.as_str()) else {
            continue;
        };
        let target_part = path::resolve_target("xl/richData/richValueRel.xml", target);
        let Some(bytes) = pkg.part(&target_part) else {
            continue;
        };
        out.insert((sheet_name, cell), bytes.to_vec());
    }

    Ok(out)
}

fn is_rich_value_part(name: &str) -> bool {
    const PREFIX: &str = "xl/richData/richValue";
    const SUFFIX: &str = ".xml";
    if !name.starts_with(PREFIX) || !name.ends_with(SUFFIX) {
        return false;
    }
    let mid = &name[PREFIX.len()..name.len() - SUFFIX.len()];
    mid.chars().all(|c| c.is_ascii_digit())
}

fn parse_vm_to_rich_value_index_map(bytes: &[u8]) -> Result<HashMap<u32, u32>, RichDataError> {
    // Prefer the structured metadata parser (which understands metadataTypes/futureMetadata),
    // but fall back to a looser `<xlrd:rvb i="..."/>` scan for forward/backward compatibility.
    let primary = metadata::parse_value_metadata_vm_to_rich_value_index_map(bytes)
        .map_err(|e| map_xml_dom_error("xl/metadata.xml", e))?;
    if !primary.is_empty() {
        // Excel's `vm` appears in the wild as both 0-based and 1-based. To be tolerant, insert
        // both the original key and its 0-based equivalent.
        let mut out = HashMap::new();
        for (vm, rv) in primary {
            out.entry(vm).or_insert(rv);
            if vm > 0 {
                out.entry(vm - 1).or_insert(rv);
            }
        }
        return Ok(out);
    }

    // Fallback parser: find all `<rvb i="...">` in document order and treat `<rc v="...">` as an
    // index into that list. This matches some simplified metadata.xml payloads.
    parse_metadata_vm_mapping_fallback(bytes)
}

fn parse_metadata_vm_mapping_fallback(bytes: &[u8]) -> Result<HashMap<u32, u32>, RichDataError> {
    let xml = std::str::from_utf8(bytes).map_err(|source| RichDataError::XmlNonUtf8 {
        part: "xl/metadata.xml".to_string(),
        source,
    })?;
    let doc = roxmltree::Document::parse(xml).map_err(|source| RichDataError::XmlParse {
        part: "xl/metadata.xml".to_string(),
        source,
    })?;

    let mut rvb_rich_value_indices: Vec<u32> = Vec::new();
    for node in doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "rvb")
    {
        let Some(i_attr) = node.attribute("i") else {
            continue;
        };
        let Ok(i) = i_attr.trim().parse::<u32>() else {
            continue;
        };
        rvb_rich_value_indices.push(i);
    }
    if rvb_rich_value_indices.is_empty() {
        return Ok(HashMap::new());
    }

    let Some(value_metadata) = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "valueMetadata")
    else {
        return Ok(HashMap::new());
    };

    let mut out: HashMap<u32, u32> = HashMap::new();
    for (vm_idx, bk) in value_metadata
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "bk")
        .enumerate()
    {
        for rc in bk
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "rc")
        {
            let Some(v_attr) = rc.attribute("v") else {
                continue;
            };
            let Ok(v_idx) = v_attr.trim().parse::<usize>() else {
                continue;
            };
            let Some(&rich_value_idx) = rvb_rich_value_indices.get(v_idx) else {
                continue;
            };
            if let Ok(vm_idx) = u32::try_from(vm_idx) {
                out.insert(vm_idx, rich_value_idx);
            }
            break;
        }
    }

    Ok(out)
}

fn map_xml_dom_error(part: &str, err: crate::xml::XmlDomError) -> RichDataError {
    match err {
        crate::xml::XmlDomError::Utf8(source) => RichDataError::XmlNonUtf8 {
            part: part.to_string(),
            source,
        },
        crate::xml::XmlDomError::Parse(source) => RichDataError::XmlParse {
            part: part.to_string(),
            source,
        },
    }
}

fn parse_worksheet_vm_cells(
    bytes: &[u8],
    part_name: &str,
) -> Result<Vec<(CellRef, u32)>, RichDataError> {
    let xml = std::str::from_utf8(bytes).map_err(|source| RichDataError::XmlNonUtf8 {
        part: part_name.to_string(),
        source,
    })?;
    let doc = roxmltree::Document::parse(xml).map_err(|source| RichDataError::XmlParse {
        part: part_name.to_string(),
        source,
    })?;

    let mut out: Vec<(CellRef, u32)> = Vec::new();
    for cell in doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "c")
    {
        let Some(vm) = cell.attribute("vm") else {
            continue;
        };
        let Ok(vm_idx) = vm.trim().parse::<u32>() else {
            continue;
        };
        let Some(r) = cell.attribute("r") else {
            continue;
        };
        let Ok(cell_ref) = CellRef::from_a1(r) else {
            continue;
        };
        out.push((cell_ref, vm_idx));
    }

    Ok(out)
}

fn parse_rich_value_parts_rel_indices(
    pkg: &XlsxPackage,
    part_names: &[&str],
) -> Result<HashMap<u32, u32>, RichDataError> {
    let mut out: HashMap<u32, u32> = HashMap::new();
    let mut global_idx: u32 = 0;

    for part_name in part_names {
        let Some(bytes) = pkg.part(part_name) else {
            continue;
        };
        let xml = std::str::from_utf8(bytes).map_err(|source| RichDataError::XmlNonUtf8 {
            part: (*part_name).to_string(),
            source,
        })?;
        let doc = roxmltree::Document::parse(xml).map_err(|source| RichDataError::XmlParse {
            part: (*part_name).to_string(),
            source,
        })?;

        for rv in doc
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "rv")
        {
            if let Some(rel_idx) = parse_rv_rel_index(rv) {
                out.insert(global_idx, rel_idx);
            }
            global_idx = global_idx.saturating_add(1);
        }
    }

    Ok(out)
}

fn parse_rv_rel_index(rv: roxmltree::Node<'_, '_>) -> Option<u32> {
    // The exact richValue schema depends on the rich value type. For cell images, the relationship
    // index seems to appear as a `<v>` whose text is a u32.
    let v_elems: Vec<_> = rv
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "v")
        .collect();

    // Prefer `<v t="rel">` / `<v t="r">`.
    for v in &v_elems {
        let Some(t) = v.attribute("t") else {
            continue;
        };
        if t == "rel" || t == "r" {
            if let Some(text) = v.text() {
                if let Ok(idx) = text.trim().parse::<u32>() {
                    return Some(idx);
                }
            }
        }
    }

    // Fall back to a numeric `<v>` without a type marker.
    for v in &v_elems {
        if v.attribute("t").is_some() {
            continue;
        }
        if let Some(text) = v.text() {
            if let Ok(idx) = text.trim().parse::<u32>() {
                return Some(idx);
            }
        }
    }

    // Last-ditch: any numeric `<v>`.
    for v in &v_elems {
        if let Some(text) = v.text() {
            if let Ok(idx) = text.trim().parse::<u32>() {
                return Some(idx);
            }
        }
    }

    None
}

fn parse_rich_value_rel_rids(bytes: &[u8]) -> Result<Vec<String>, RichDataError> {
    let part_name = "xl/richData/richValueRel.xml";
    let xml = std::str::from_utf8(bytes).map_err(|source| RichDataError::XmlNonUtf8 {
        part: part_name.to_string(),
        source,
    })?;
    let doc = roxmltree::Document::parse(xml).map_err(|source| RichDataError::XmlParse {
        part: part_name.to_string(),
        source,
    })?;

    let mut out: Vec<String> = Vec::new();
    for rel in doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "rel")
    {
        // Typically emitted as `r:id="rId..."`, but accept a few variants.
        let Some(rid) = rel
            .attribute((
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                "id",
            ))
            .or_else(|| rel.attribute("r:id"))
            .or_else(|| rel.attribute("id"))
        else {
            continue;
        };
        out.push(rid.to_string());
    }

    Ok(out)
}

fn parse_rich_value_rel_rels(bytes: &[u8]) -> Result<HashMap<String, String>, RichDataError> {
    let part_name = "xl/richData/_rels/richValueRel.xml.rels";
    let xml = std::str::from_utf8(bytes).map_err(|source| RichDataError::XmlNonUtf8 {
        part: part_name.to_string(),
        source,
    })?;
    let doc = roxmltree::Document::parse(xml).map_err(|source| RichDataError::XmlParse {
        part: part_name.to_string(),
        source,
    })?;

    let mut out: HashMap<String, String> = HashMap::new();
    for rel in doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "Relationship")
    {
        // Skip external relationships.
        if rel
            .attribute("TargetMode")
            .is_some_and(|m| m.trim().eq_ignore_ascii_case("External"))
        {
            continue;
        }
        let Some(id) = rel.attribute("Id") else {
            continue;
        };
        let Some(target) = rel.attribute("Target") else {
            continue;
        };
        out.insert(id.to_string(), target.to_string());
    }
    Ok(out)
}
