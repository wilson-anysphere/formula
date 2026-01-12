//! Parsers/utilities for Excel "rich data" parts.
//!
//! Excel stores cell-level rich values (data types, images-in-cells, etc.) via:
//! - `xl/worksheets/sheet*.xml` `c/@vm` (value-metadata index)
//! - `xl/metadata.xml` `<valueMetadata>` + rich-data extension payloads
//! - `xl/richData/richValue*.xml` (rich-value records)
//! - `xl/richData/richValueRel.xml` + `xl/richData/_rels/richValueRel.xml.rels` (relationship indirection)
//!
//! This module exposes best-effort parsing helpers without integrating into `formula-model` yet.

mod discovery;
pub mod metadata;
pub mod linked_data_types;
pub mod rich_value;
pub mod rich_value_parts;
pub mod rich_value_rel;
pub mod rich_value_structure;
pub mod rich_value_types;
mod images;
mod media_parts;
mod rich_value_images;
mod worksheet_scan;

pub use discovery::discover_rich_data_part_names;
pub use images::resolve_rich_value_image_targets;
pub use linked_data_types::{extract_linked_data_types, ExtractedLinkedDataType};
pub use rich_value::parse_rich_values_xml;
pub use rich_value::{RichValueFieldValue, RichValueInstance, RichValues};
pub use rich_value_images::{
    ExtractedRichValueImages, RichValueEntry, RichValueIndex, RichValueWarning,
};
pub use rich_value_parts::RichValueParts;
pub use rich_value_rel::RichValueRels;
pub use worksheet_scan::scan_cells_with_metadata_indices;

use std::cmp::Ordering;

/// One observed workbook relationship type for `xl/richData/richValueRel.xml`.
///
/// Note: Excel also emits other relationship type URIs for rich values depending on build/version;
/// see `docs/20-images-in-cells-richdata.md` for a fixture-backed summary.
pub const REL_TYPE_RICH_VALUE_REL: &str =
    "http://schemas.microsoft.com/office/2022/10/relationships/richValueRel";
/// Workbook relationship type for `xl/richData/rdrichvalue.xml`.
pub const REL_TYPE_RD_RICH_VALUE: &str =
    "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue";
/// Workbook relationship type for `xl/richData/rdrichvaluestructure.xml`.
pub const REL_TYPE_RD_RICH_VALUE_STRUCTURE: &str =
    "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure";
/// Workbook relationship type for `xl/richData/rdRichValueTypes.xml`.
pub const REL_TYPE_RD_RICH_VALUE_TYPES: &str =
    "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes";

/// `[Content_Types].xml` override content type for `xl/richData/rdRichValueTypes.xml`.
pub const CONTENT_TYPE_RDRICHVALUETYPES: &str = "application/vnd.ms-excel.rdrichvaluetypes+xml";
/// `[Content_Types].xml` override content type for `xl/richData/rdrichvalue.xml`.
pub const CONTENT_TYPE_RDRICHVALUE: &str = "application/vnd.ms-excel.rdrichvalue+xml";
/// `[Content_Types].xml` override content type for `xl/richData/rdrichvaluestructure.xml`.
pub const CONTENT_TYPE_RDRICHVALUESTRUCTURE: &str =
    "application/vnd.ms-excel.rdrichvaluestructure+xml";
/// `[Content_Types].xml` override content type for `xl/richData/richValueRel.xml`.
pub const CONTENT_TYPE_RICHVALUEREL: &str = "application/vnd.ms-excel.richvaluerel+xml";
use std::collections::HashMap;

use formula_model::CellRef;
use thiserror::Error;

use crate::{path, XlsxError, XlsxPackage};

/// Cached lookup tables for resolving worksheet `c/@vm` indices into rich value + media targets.
///
/// This is intended primarily for debugging tooling (like `dump_rich_data`) and is intentionally
/// best-effort:
/// - Missing parts yield empty lookups.
/// - Malformed XML yields an error.
#[derive(Debug, Clone, Default)]
pub struct RichDataVmIndex {
    vm_to_rich_value_index: HashMap<u32, u32>,
    rich_value_index_to_rel_index: HashMap<u32, u32>,
    rel_index_to_target_part: Vec<Option<String>>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct RichDataVmResolution {
    pub rich_value_index: Option<u32>,
    pub rel_index: Option<u32>,
    /// Resolved OPC part name (e.g. `xl/media/image1.png`).
    pub target_part: Option<String>,
}

impl RichDataVmIndex {
    /// Build the lookup tables used to resolve `vm` indices.
    pub fn build(pkg: &XlsxPackage) -> Result<Self, RichDataError> {
        let metadata_part = resolve_workbook_metadata_part_name(pkg)?;
        let vm_to_rich_value_index = pkg
            .part(&metadata_part)
            .map(|bytes| parse_vm_to_rich_value_index_map(bytes, &metadata_part))
            .transpose()?
            .unwrap_or_default();

        let mut rich_value_parts: Vec<&str> = pkg
            .part_names()
            .filter(|name| is_rich_value_part(name))
            .collect();
        rich_value_parts.sort_by(|a, b| cmp_rich_value_parts_by_numeric_suffix(a, b));

        let rich_value_index_to_rel_index = if rich_value_parts.is_empty() {
            HashMap::new()
        } else {
            parse_rich_value_parts_rel_indices(pkg, &rich_value_parts)?
        };

        let rel_index_to_target_part = build_rich_value_rel_index_to_target_part(pkg)?;

        Ok(Self {
            vm_to_rich_value_index,
            rich_value_index_to_rel_index,
            rel_index_to_target_part,
        })
    }

    /// Resolve a worksheet `c/@vm` value into rich value + relationship indices and a target part.
    pub fn resolve_vm(&self, vm: u32) -> RichDataVmResolution {
        let rich_value_index = self.vm_to_rich_value_index.get(&vm).copied();
        let rel_index = rich_value_index.and_then(|idx| {
            self.rich_value_index_to_rel_index
                .get(&idx)
                .copied()
                // Some legacy/degenerate packages omit `richValue*.xml` parts. In those workbooks
                // the rich value index often maps directly to an entry in `richValueRel.xml`.
                .or_else(|| {
                    if self.rich_value_index_to_rel_index.is_empty() {
                        Some(idx)
                    } else {
                        None
                    }
                })
        });
        let target_part = rel_index.and_then(|idx| {
            self.rel_index_to_target_part
                .get(idx as usize)
                .cloned()
                .flatten()
        });
        RichDataVmResolution {
            rich_value_index,
            rel_index,
            target_part,
        }
    }
}

// Note: `XlsxPackage::extract_rich_data_images` is defined in `media_parts.rs` as a legacy helper
// that returns raw `xl/media/*` parts referenced by the RichData graph.
//
// For per-cell in-cell image extraction, use `XlsxPackage::extract_rich_cell_images_by_cell`.
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
fn rich_data_error_to_xlsx(err: RichDataError) -> XlsxError {
    match err {
        RichDataError::Xlsx(err) => err,
        RichDataError::XmlNonUtf8 { part, source } => {
            XlsxError::Invalid(format!("xml part {part} is not valid UTF-8: {source}"))
        }
        RichDataError::XmlParse { part, source } => {
            XlsxError::Invalid(format!("xml parse error in {part}: {source}"))
        }
    }
}

impl XlsxPackage {
    /// Extract in-cell images stored via the Excel rich-data (`xl/metadata.xml` + `xl/richData/*`)
    /// mechanism.
    ///
    /// Returns a map keyed by `(sheet_name, cell)`.
    ///
    /// This is a convenience wrapper around [`crate::rich_data::extract_rich_cell_images`], with a
    /// fallback for legacy/simplified workbooks that omit the `xl/richData/richValue*.xml` parts
    /// and instead index directly into `xl/richData/richValueRel.xml`.
    pub fn extract_rich_cell_images_by_cell(
        &self,
    ) -> Result<HashMap<(String, CellRef), Vec<u8>>, XlsxError> {
        let extracted = if self.part_names().any(is_rich_value_part) {
            extract_rich_cell_images(self)
        } else {
            extract_rich_data_images_via_rel_table(self)
        };

        extracted.map_err(rich_data_error_to_xlsx)
    }
}

fn extract_rich_data_images_via_rel_table(
    pkg: &XlsxPackage,
) -> Result<HashMap<(String, CellRef), Vec<u8>>, RichDataError> {
    // If we can't resolve sheet names to parts, we can't provide stable (sheet name, cell) keys.
    // Treat missing workbook parts as "no richData".
    if pkg.part("xl/workbook.xml").is_none() || pkg.part("xl/_rels/workbook.xml.rels").is_none() {
        return Ok(HashMap::new());
    }

    // The workbook parsing stack can error for malformed workbook.xml; bubble that up.
    let worksheet_parts = pkg.worksheet_parts()?;

    let metadata_part = resolve_workbook_metadata_part_name(pkg)?;
    let Some(metadata_bytes) = pkg.part(&metadata_part) else {
        return Ok(HashMap::new());
    };
    let vm_to_rich_value = parse_vm_to_rich_value_index_map(metadata_bytes, &metadata_part)?;
    if vm_to_rich_value.is_empty() {
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
    for sheet in worksheet_parts {
        let Some(sheet_bytes) = pkg.part(&sheet.worksheet_part) else {
            continue;
        };
        let cells = parse_worksheet_vm_cells(sheet_bytes)?;
        for (cell, vm) in cells {
            let Some(&rich_value_idx) = vm_to_rich_value.get(&vm) else {
                continue;
            };
            let Some(rid) = rel_index_to_rid.get(rich_value_idx as usize) else {
                continue;
            };
            let Some(target) = rid_to_target.get(rid) else {
                continue;
            };
            let target_part = path::resolve_target("xl/richData/richValueRel.xml", target);
            let Some(bytes) = pkg.part(&target_part) else {
                continue;
            };
            out.insert((sheet.name.clone(), cell), bytes.to_vec());
        }
    }

    Ok(out)
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

    let metadata_part = resolve_workbook_metadata_part_name(pkg)?;
    let Some(metadata_bytes) = pkg.part(&metadata_part) else {
        return Ok(HashMap::new());
    };
    let vm_to_rich_value = parse_vm_to_rich_value_index_map(metadata_bytes, &metadata_part)?;
    if vm_to_rich_value.is_empty() {
        return Ok(HashMap::new());
    }

    let mut cells_with_rich_value: Vec<(String, CellRef, u32)> = Vec::new();
    for sheet in worksheet_parts {
        let Some(sheet_bytes) = pkg.part(&sheet.worksheet_part) else {
            continue;
        };
        let cells = parse_worksheet_vm_cells(sheet_bytes)?;
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

    let mut rich_value_parts: Vec<&str> = pkg
        .part_names()
        .filter(|name| is_rich_value_part(name))
        .collect();
    rich_value_parts.sort_by(|a, b| cmp_rich_value_parts_by_numeric_suffix(a, b));
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

fn resolve_workbook_metadata_part_name(pkg: &XlsxPackage) -> Result<String, RichDataError> {
    const DEFAULT: &str = "xl/metadata.xml";
    const REL_TYPE_SHEET_METADATA: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata";
    const REL_TYPE_METADATA: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata";

    let Some(rels_bytes) = pkg.part("xl/_rels/workbook.xml.rels") else {
        return Ok(DEFAULT.to_string());
    };

    let rels = crate::openxml::parse_relationships(rels_bytes)?;
    for rel in rels {
        if rel
            .target_mode
            .as_deref()
            .is_some_and(|m| m.trim().eq_ignore_ascii_case("External"))
        {
            continue;
        }

        let type_uri = rel.type_uri.trim();
        if type_uri == REL_TYPE_SHEET_METADATA || type_uri == REL_TYPE_METADATA {
            return Ok(path::resolve_target("xl/workbook.xml", &rel.target));
        }
    }

    Ok(DEFAULT.to_string())
}

fn build_rich_value_rel_index_to_target_part(
    pkg: &XlsxPackage,
) -> Result<Vec<Option<String>>, RichDataError> {
    const SOURCE_PART: &str = "xl/richData/richValueRel.xml";

    let Some(rich_value_rel_bytes) = pkg.part(SOURCE_PART) else {
        return Ok(Vec::new());
    };
    let rel_index_to_rid =
        rich_value_rel::parse_rich_value_rel_table(rich_value_rel_bytes).map_err(RichDataError::from)?;
    if rel_index_to_rid.is_empty() {
        return Ok(Vec::new());
    }

    let rels_part_name = crate::openxml::rels_part_name(SOURCE_PART);
    let Some(rels_bytes) = pkg.part(&rels_part_name) else {
        return Ok(vec![None; rel_index_to_rid.len()]);
    };

    let relationships = crate::openxml::parse_relationships(rels_bytes)?;
    let mut rid_to_target_part: HashMap<String, String> = HashMap::new();
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

        // Some producers emit `Target="media/image1.png"` (relative to `xl/`) rather than the more
        // common `Target="../media/image1.png"` (relative to `xl/richData/`). Make a best-effort
        // guess for this case.
        let target_part = if target.starts_with("media/") {
            format!("xl/{target}")
        } else {
            path::resolve_target(SOURCE_PART, target)
        };

        rid_to_target_part.insert(rel.id, target_part);
    }

    let mut out: Vec<Option<String>> = Vec::with_capacity(rel_index_to_rid.len());
    for rid in rel_index_to_rid {
        if rid.is_empty() {
            out.push(None);
        } else {
            out.push(rid_to_target_part.get(&rid).cloned());
        }
    }

    Ok(out)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
enum RichValuePartFamily {
    RichValue,
    RdRichValue,
}

/// Parse an `xl/richData/*` rich value part name and return its "family" plus numeric suffix.
///
/// Examples:
/// - `xl/richData/richValue.xml` -> (`RichValue`, 0)
/// - `xl/richData/richValue2.xml` -> (`RichValue`, 2)
/// - `xl/richData/rdrichvalue10.xml` -> (`RdRichValue`, 10)
///
/// This is case-insensitive for the filename (but not the containing directory).
fn parse_rich_value_part_name(part_path: &str) -> Option<(RichValuePartFamily, u32)> {
    if !part_path.starts_with("xl/richData/") {
        return None;
    }

    let file_name = part_path.rsplit('/').next()?;
    let file_name_lower = file_name.to_ascii_lowercase();
    if !file_name_lower.ends_with(".xml") {
        return None;
    }

    let stem = &file_name[..file_name.len() - ".xml".len()];
    let stem_lower = stem.to_ascii_lowercase();

    let (family, suffix) = if let Some(rest) = stem_lower.strip_prefix("richvalue") {
        (RichValuePartFamily::RichValue, rest)
    } else if let Some(rest) = stem_lower.strip_prefix("rdrichvalue") {
        (RichValuePartFamily::RdRichValue, rest)
    } else {
        return None;
    };

    let idx = if suffix.is_empty() {
        0
    } else if suffix.chars().all(|c| c.is_ascii_digit()) {
        suffix.parse::<u32>().ok()?
    } else {
        return None;
    };

    Some((family, idx))
}

fn is_rich_value_part(name: &str) -> bool {
    parse_rich_value_part_name(name).is_some()
}

fn cmp_rich_value_parts_by_numeric_suffix(a: &str, b: &str) -> Ordering {
    let Some((a_family, a_idx)) = parse_rich_value_part_name(a) else {
        return a.cmp(b);
    };
    let Some((b_family, b_idx)) = parse_rich_value_part_name(b) else {
        return a.cmp(b);
    };

    a_family
        .cmp(&b_family)
        .then(a_idx.cmp(&b_idx))
        .then_with(|| a.cmp(b))
}

fn parse_vm_to_rich_value_index_map(
    bytes: &[u8],
    part_name: &str,
) -> Result<HashMap<u32, u32>, RichDataError> {
    // Prefer the structured metadata parser (which understands metadataTypes/futureMetadata),
    // but fall back to a looser `<xlrd:rvb i="..."/>` scan for forward/backward compatibility.
    let primary = metadata::parse_value_metadata_vm_to_rich_value_index_map(bytes)
        .map_err(|e| map_xml_dom_error(part_name, e))?;
    if !primary.is_empty() {
        // Excel's `vm` appears in the wild as both 0-based and 1-based. To be tolerant, insert
        // both the original key and its 0-based equivalent.
        //
        // Note: `primary` is a `HashMap`, so iteration order is not deterministic. Insert in two
        // passes so the canonical (1-based) `vm` keys always win when the shifted `vm-1` entries
        // collide (e.g. vm=1 and vm=2 both attempt to populate key 1).
        let mut out = HashMap::new();

        // Pass 1: canonical keys.
        for (&vm, &rv) in primary.iter() {
            out.entry(vm).or_insert(rv);
        }

        // Pass 2: tolerate 0-based vm indices.
        for (vm, rv) in primary {
            if vm > 0 {
                out.entry(vm - 1).or_insert(rv);
            }
        }
        return Ok(out);
    }

    // Fallback parser: find all `<rvb i="...">` in document order and treat `<rc v="...">` as an
    // index into that list. This matches some simplified metadata.xml payloads.
    parse_metadata_vm_mapping_fallback(bytes, part_name)
}

fn parse_metadata_vm_mapping_fallback(
    bytes: &[u8],
    part_name: &str,
) -> Result<HashMap<u32, u32>, RichDataError> {
    let xml = std::str::from_utf8(bytes).map_err(|source| RichDataError::XmlNonUtf8 {
        part: part_name.to_string(),
        source,
    })?;
    let doc = roxmltree::Document::parse(xml).map_err(|source| RichDataError::XmlParse {
        part: part_name.to_string(),
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
    let mut vm_idx: u32 = 0;
    for bk in value_metadata
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "bk")
    {
        let count = bk
            .attribute("count")
            .and_then(|c| c.trim().parse::<u32>().ok())
            .filter(|c| *c >= 1)
            .unwrap_or(1);

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
            for offset in 0..count {
                out.insert(vm_idx.saturating_add(offset), rich_value_idx);
            }
            break;
        }

        vm_idx = vm_idx.saturating_add(count);
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

fn parse_worksheet_vm_cells(bytes: &[u8]) -> Result<Vec<(CellRef, u32)>, RichDataError> {
    // Use the streaming scan to avoid materializing a DOM for large worksheets.
    let cells = scan_cells_with_metadata_indices(bytes)?;
    Ok(cells
        .into_iter()
        .filter_map(|(cell, vm, _cm)| vm.map(|vm| (cell, vm)))
        .collect())
}

fn parse_rich_value_parts_rel_indices(
    pkg: &XlsxPackage,
    part_names: &[&str],
) -> Result<HashMap<u32, u32>, RichDataError> {
    let mut out: HashMap<u32, u32> = HashMap::new();
    let mut global_idx: u32 = 0;

    // Optional optimization: when present, `richValueTypes.xml` (and optionally
    // `richValueStructure.xml`) can describe the schema of each rich value type and which property
    // corresponds to a relationship (e.g. an image).
    //
    // This is best-effort: malformed/unexpected schemas are ignored and we fall back to the
    // heuristic `<v t="rel">` / first-numeric parsing.
    let rich_value_types = build_rich_value_type_schema_index_best_effort(pkg);

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
            if let Some(rel_idx) = parse_rv_rel_index(rv, rich_value_types.as_ref()) {
                out.insert(global_idx, rel_idx);
            }
            global_idx = global_idx.saturating_add(1);
        }
    }

    Ok(out)
}

#[derive(Debug, Clone)]
struct RichValueTypePropertyDescriptor {
    name: Option<String>,
    kind: Option<String>,
    position: usize,
}

fn build_rich_value_type_schema_index_best_effort(
    pkg: &XlsxPackage,
) -> Option<HashMap<u32, Vec<RichValueTypePropertyDescriptor>>> {
    let rich_value_types_bytes = pkg.part("xl/richData/richValueTypes.xml")?;

    let mut out: HashMap<u32, Vec<RichValueTypePropertyDescriptor>> =
        parse_rich_value_types_xml_best_effort(rich_value_types_bytes).unwrap_or_default();

    // `richValueTypes.xml` often maps type IDs to structure IDs, with the member layout defined in
    // `richValueStructure.xml`. When the inline property list is absent, derive property positions
    // from the structure definition.
    if let Some(type_to_structure) =
        parse_rich_value_type_structure_ids_best_effort(rich_value_types_bytes)
    {
        if let Some(structure_bytes) = pkg.part("xl/richData/richValueStructure.xml") {
            if let Ok(structures) =
                rich_value_structure::parse_rich_value_structure_xml(structure_bytes)
            {
                for (type_id, structure_id) in type_to_structure {
                    let Some(structure) = structures.get(&structure_id) else {
                        continue;
                    };
                    if structure.members.is_empty() {
                        continue;
                    }

                    // Only replace existing schemas when they don't appear to include a
                    // relationship property.
                    if out
                        .get(&type_id)
                        .is_some_and(|existing| existing.iter().any(descriptor_is_relationship))
                    {
                        continue;
                    }

                    let mut props = Vec::with_capacity(structure.members.len());
                    for (position, member) in structure.members.iter().enumerate() {
                        props.push(RichValueTypePropertyDescriptor {
                            name: Some(member.name.clone()),
                            kind: member.kind.clone(),
                            position,
                        });
                    }
                    out.insert(type_id, props);
                }
            }
        }
    }

    if out.is_empty() { None } else { Some(out) }
}

fn parse_rich_value_types_xml_best_effort(
    bytes: &[u8],
) -> Option<HashMap<u32, Vec<RichValueTypePropertyDescriptor>>> {
    let xml = std::str::from_utf8(bytes).ok()?;
    let doc = roxmltree::Document::parse(xml).ok()?;

    // Prefer explicit `<rvType>`/`<richValueType>` nodes, but fall back to `*Type` elements under
    // a `*Types` ancestor.
    let mut type_nodes: Vec<roxmltree::Node<'_, '_>> = doc
        .descendants()
        .filter(|n| n.is_element())
        .filter(|n| {
            let name = n.tag_name().name();
            name.eq_ignore_ascii_case("rvType") || name.eq_ignore_ascii_case("richValueType")
        })
        .collect();

    if type_nodes.is_empty() {
        type_nodes = doc
            .descendants()
            .filter(|n| n.is_element())
            .filter(|n| {
                let name = n.tag_name().name();
                if !name.to_ascii_lowercase().ends_with("type") {
                    return false;
                }
                n.parent_element()
                    .is_some_and(|p| p.tag_name().name().to_ascii_lowercase().ends_with("types"))
            })
            .collect();
    }

    if type_nodes.is_empty() {
        return None;
    }

    let mut out: HashMap<u32, Vec<RichValueTypePropertyDescriptor>> = HashMap::new();
    let mut next_fallback_idx: u32 = 0;
    for ty in type_nodes {
        let idx = parse_node_numeric_attr(ty, &["id", "i", "idx", "s"]).unwrap_or_else(|| {
            let idx = next_fallback_idx;
            next_fallback_idx = next_fallback_idx.wrapping_add(1);
            idx
        });

        let prop_container = ty
            .children()
            .find(|n| {
                n.is_element()
                    && matches!(
                        n.tag_name().name().to_ascii_lowercase().as_str(),
                        "props" | "properties" | "fields"
                    )
            })
            .unwrap_or(ty);

        let prop_nodes: Vec<roxmltree::Node<'_, '_>> = prop_container
            .children()
            .filter(|n| n.is_element() && looks_like_rich_value_type_property_def(*n))
            .collect();
        if prop_nodes.is_empty() {
            continue;
        }

        let mut props = Vec::with_capacity(prop_nodes.len());
        for (position, prop) in prop_nodes.into_iter().enumerate() {
            let name =
                find_node_attr_value(prop, &["name", "n", "id", "key", "k", "pid", "propId"]);
            let kind = find_node_attr_value(prop, &["t", "type", "kind", "vt", "valType"]);
            props.push(RichValueTypePropertyDescriptor {
                name,
                kind,
                position,
            });
        }
        out.insert(idx, props);
    }

    if out.is_empty() {
        None
    } else {
        Some(out)
    }
}

fn parse_rich_value_type_structure_ids_best_effort(bytes: &[u8]) -> Option<HashMap<u32, String>> {
    let xml = std::str::from_utf8(bytes).ok()?;
    let doc = roxmltree::Document::parse(xml).ok()?;

    // Reuse the same type-node detection strategy as the inline-property parser.
    let mut type_nodes: Vec<roxmltree::Node<'_, '_>> = doc
        .descendants()
        .filter(|n| n.is_element())
        .filter(|n| {
            let name = n.tag_name().name();
            name.eq_ignore_ascii_case("rvType") || name.eq_ignore_ascii_case("richValueType")
        })
        .collect();

    if type_nodes.is_empty() {
        type_nodes = doc
            .descendants()
            .filter(|n| n.is_element())
            .filter(|n| {
                let name = n.tag_name().name();
                if !name.to_ascii_lowercase().ends_with("type") {
                    return false;
                }
                n.parent_element()
                    .is_some_and(|p| p.tag_name().name().to_ascii_lowercase().ends_with("types"))
            })
            .collect();
    }

    if type_nodes.is_empty() {
        return None;
    }

    let mut out = HashMap::new();
    let mut next_fallback_idx: u32 = 0;
    for ty in type_nodes {
        let idx =
            parse_node_numeric_attr(ty, &["id", "i", "idx", "s"]).unwrap_or_else(|| {
                let idx = next_fallback_idx;
                next_fallback_idx = next_fallback_idx.wrapping_add(1);
                idx
            });
        let structure_id = find_node_attr_value(
            ty,
            &["structure", "struct", "structureId", "structureID", "sid", "sId"],
        );
        let Some(structure_id) = structure_id else {
            continue;
        };
        out.insert(idx, structure_id);
    }

    if out.is_empty() { None } else { Some(out) }
}

fn parse_rv_rel_index_with_schema(
    rv: roxmltree::Node<'_, '_>,
    rich_value_types: Option<&HashMap<u32, Vec<RichValueTypePropertyDescriptor>>>,
) -> Option<u32> {
    let rich_value_types = rich_value_types?;
    let type_idx = parse_node_numeric_attr(rv, &["s", "t", "type"])?;
    let schema = rich_value_types.get(&type_idx)?;
    let rel_prop = schema.iter().find(|d| descriptor_is_relationship(d))?;

    // Schema positions refer to the `<v>` children of `<rv>` in order.
    let v_children: Vec<_> = rv
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "v")
        .collect();
    let v = v_children.get(rel_prop.position)?;
    let text = v.text()?.trim();
    text.parse::<u32>().ok()
}

fn descriptor_is_relationship(desc: &RichValueTypePropertyDescriptor) -> bool {
    let kind = desc
        .kind
        .as_deref()
        .unwrap_or_default()
        .to_ascii_lowercase();
    if kind == "rel" || kind == "relationship" {
        return true;
    }

    let name = desc
        .name
        .as_deref()
        .unwrap_or_default()
        .to_ascii_lowercase();
    if name == "rel" || name == "relationship" {
        return true;
    }

    // Some schemas might use a property key like `image`/`img` for media relationships.
    name.contains("image") || name.contains("img")
}

fn local_name(name: &str) -> &str {
    name.rsplit_once(':')
        .map(|(_, local)| local)
        .unwrap_or(name)
}

fn parse_node_numeric_attr(node: roxmltree::Node<'_, '_>, keys: &[&str]) -> Option<u32> {
    for attr in node.attributes() {
        let k = local_name(attr.name());
        if keys.iter().any(|key| k.eq_ignore_ascii_case(key)) {
            if let Ok(v) = attr.value().trim().parse::<u32>() {
                return Some(v);
            }
        }
    }
    None
}

fn find_node_attr_value(node: roxmltree::Node<'_, '_>, keys: &[&str]) -> Option<String> {
    for attr in node.attributes() {
        let k = local_name(attr.name());
        if keys.iter().any(|key| k.eq_ignore_ascii_case(key)) {
            return Some(attr.value().to_string());
        }
    }
    None
}

fn looks_like_rich_value_type_property_def(node: roxmltree::Node<'_, '_>) -> bool {
    // Typical property nodes have either a name/id or a type/kind attribute. We treat this as a
    // heuristic so we don't accidentally include container nodes like `<props>`.
    for attr in node.attributes() {
        if matches!(
            local_name(attr.name()).to_ascii_lowercase().as_str(),
            "t" | "type" | "kind" | "name" | "n" | "id" | "key" | "k" | "pid" | "propid"
        ) {
            return true;
        }
    }
    false
}

fn parse_rv_rel_index(
    rv: roxmltree::Node<'_, '_>,
    rich_value_types: Option<&HashMap<u32, Vec<RichValueTypePropertyDescriptor>>>,
) -> Option<u32> {
    // If we can infer the rich value type schema, prefer schema-driven selection of the
    // relationship-index field. This avoids picking the wrong `<v>` when the rich value has
    // multiple numeric properties.
    if let Some(rel_idx) = parse_rv_rel_index_with_schema(rv, rich_value_types) {
        return Some(rel_idx);
    }

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
        let target = strip_fragment(target);
        if target.is_empty() {
            continue;
        }
        out.insert(id.to_string(), target.to_string());
    }
    Ok(out)
}

fn strip_fragment(target: &str) -> &str {
    target
        .split_once('#')
        .map(|(base, _)| base)
        .unwrap_or(target)
}
