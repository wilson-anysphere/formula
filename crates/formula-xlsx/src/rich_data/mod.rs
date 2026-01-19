//! Parsers/utilities for Excel "rich data" parts.
//!
//! Excel stores cell-level rich values (data types, images-in-cells, etc.) via:
//! - `xl/worksheets/sheet*.xml` `c/@vm` (value-metadata index)
//! - `xl/_rels/workbook.xml.rels` relationship type
//!   `http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata`
//!   (as emitted by Excel) to locate `xl/metadata.xml`
//! - `xl/metadata.xml` `<valueMetadata>` + rich-data extension payloads
//! - rich-value record parts, which appear in two naming families:
//!   - `xl/richData/richValue*.xml` (also observed as `xl/richData/richValues*.xml`)
//!   - `xl/richData/rdrichvalue*.xml`
//! - `xl/richData/richValueRel.xml` + `xl/richData/_rels/richValueRel.xml.rels` (relationship indirection)
//!
//! For a concrete schema walkthrough (including images-in-cells), see
//! `docs/xlsx-embedded-images-in-cells.md`.
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

pub use discovery::{discover_rich_data_part_names, discover_rich_data_part_names_from_metadata_rels};
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
use std::collections::HashMap;

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

use formula_model::CellRef;
use thiserror::Error;

use crate::{path, XlsxError, XlsxPackage};

/// Look up an entry in a rich-value relationship-slot indexed table.
///
/// Excel uses 0-based relationship-slot indices:
/// - `<v kind="rel">0</v>` points at the first `<rel>` in `richValueRel.xml`.
///
/// Some third-party generators emit 1-based slot indices. To be tolerant without breaking valid
/// Excel files, we first try the index as-is and only fall back to `index - 1` when the original
/// index is out of bounds.
pub(crate) fn rel_slot_get<T>(table: &[T], index: usize) -> Option<&T> {
    table
        .get(index)
        .or_else(|| index.checked_sub(1).and_then(|idx| table.get(idx)))
}
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
    vm_offset: u32,
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

        // If the metadata mapping appears to be 1-based (no key 0), detect whether the workbook's
        // worksheets use 0-based `c/@vm` values (by scanning for any `vm="0"` cells) and record the
        // offset needed to resolve those sheet values against the metadata map.
        let vm_offset: u32 = if !vm_to_rich_value_index.is_empty()
            && !vm_to_rich_value_index.contains_key(&0)
        {
            let worksheet_parts: Vec<String> = match pkg.worksheet_parts() {
                Ok(infos) => infos.into_iter().map(|info| info.worksheet_part).collect(),
                // Best-effort fallback when workbook sheet discovery fails.
                Err(_) => pkg
                    .part_names()
                    .filter(|name| name.starts_with("xl/worksheets/") && name.ends_with(".xml"))
                    .map(str::to_string)
                    .collect(),
            };

            let mut saw_zero = false;
            for worksheet_part in worksheet_parts {
                let Some(sheet_bytes) = pkg.part(&worksheet_part) else {
                    continue;
                };
                let cells = scan_cells_with_metadata_indices(sheet_bytes)?;
                if cells.iter().any(|(_, vm, _)| *vm == Some(0)) {
                    saw_zero = true;
                    break;
                }
            }
            if saw_zero { 1 } else { 0 }
        } else {
            0
        };

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
            vm_offset,
        })
    }

    /// Resolve a worksheet value-metadata index (`c/@vm`) into rich value + relationship indices
    /// and a target part.
    ///
    /// Note: Excel has been observed to encode worksheet `c/@vm` as either 1-based (canonical) or
    /// 0-based. [`Self::build`] records a best-effort `vm_offset` when the workbook appears to use
    /// 0-based worksheet indices, so callers should pass the raw `vm` value as it appears on the
    /// worksheet cell.
    pub fn resolve_vm(&self, vm: u32) -> RichDataVmResolution {
        let Some(vm) = vm.checked_add(self.vm_offset) else {
            return RichDataVmResolution {
                rich_value_index: None,
                rel_index: None,
                target_part: None,
            };
        };
        let rich_value_index = self.vm_to_rich_value_index.get(&vm).copied();
        let mut rel_index = rich_value_index.and_then(|idx| {
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
        let mut target_part = rel_index.and_then(|idx| {
            rel_slot_get(&self.rel_index_to_target_part, idx as usize)
                .cloned()
                .flatten()
        });

        // Additional best-effort fallback: some producers omit `xl/metadata.xml` and/or use `vm`
        // as a direct index into `xl/richData/richValueRel.xml`. In that case we can still resolve
        // targets even though we can't infer a `rich_value_index`.
        if target_part.is_none() && rich_value_index.is_none() {
            // Excel's `vm` is commonly 1-based; the RichData relationship table appears 0-based.
            // Try `vm-1` first, then `vm`.
            let candidates: [Option<u32>; 2] = [vm.checked_sub(1), Some(vm)];
            for candidate in candidates.into_iter().flatten() {
                if let Some(part) = self
                    .rel_index_to_target_part
                    .get(candidate as usize)
                    .cloned()
                    .flatten()
                {
                    rel_index = Some(candidate);
                    target_part = Some(part);
                    break;
                }
            }
        }

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
const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
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

    let Some(metadata_part) = find_metadata_part(pkg) else {
        return Ok(HashMap::new());
    };
    let Some(metadata_bytes) = pkg.part(&metadata_part) else {
        return Ok(HashMap::new());
    };
    let vm_to_rich_value = parse_vm_to_rich_value_index_map(metadata_bytes, &metadata_part)?;
    if vm_to_rich_value.is_empty() {
        return Ok(HashMap::new());
    }

    let rich_value_part_for_rels = pkg
        .part_names()
        .find(|name| is_rich_value_part(name))
        .unwrap_or("xl/richData/richValue.xml");
    let Some(rich_value_rel_part) = find_rich_value_rel_part(pkg, rich_value_part_for_rels) else {
        return Ok(HashMap::new());
    };
    let Some(rich_value_rel_bytes) = pkg.part(&rich_value_rel_part) else {
        return Ok(HashMap::new());
    };
    let rel_index_to_rid = parse_rich_value_rel_rids(rich_value_rel_bytes)?;
    if rel_index_to_rid.is_empty() {
        return Ok(HashMap::new());
    }

    let rich_value_rel_rels_part = path::rels_for_part(&rich_value_rel_part);
    let Some(rich_value_rel_rels_bytes) = pkg.part(&rich_value_rel_rels_part) else {
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

        // See `parse_vm_to_rich_value_index_map` for rationale: keep the metadata mapping in its
        // canonical 1-based form and infer whether this worksheet uses 0-based `vm` indices by
        // checking for any `vm="0"` cells.
        //
        // As a small safety guard, only apply the offset when the mapping doesn't already contain
        // 0-based keys.
        let vm_offset: u32 = if !vm_to_rich_value.contains_key(&0)
            && cells.iter().any(|(_, vm)| *vm == 0)
        {
            1
        } else {
            0
        };

        for (cell, vm) in cells {
            let Some(vm) = vm.checked_add(vm_offset) else {
                continue;
            };
            let Some(&rich_value_idx) = vm_to_rich_value.get(&vm) else {
                continue;
            };
            let Some(rid) = rel_slot_get(&rel_index_to_rid, rich_value_idx as usize) else {
                continue;
            };
            let Some(target) = rid_to_target.get(rid) else {
                continue;
            };
            let target_part = resolve_rich_value_rel_target_part(&rich_value_rel_part, target);
            let Some(bytes) = pkg.part(&target_part) else {
                continue;
            };
            out.insert((sheet.name.clone(), cell), bytes.to_vec());
        }
    }

    Ok(out)
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum RichValueImagePointer {
    /// Indirect reference via `xl/richData/richValueRel.xml` (`<v t="rel">N</v>`).
    RelIndex(u32),
    /// Direct relationship ID reference stored inside the rich value entry (e.g. `<v>rId1</v>`),
    /// resolved via the rich value part's `.rels` file (`xl/richData/_rels/richValue*.xml.rels`).
    DirectRelId { source_part: String, rel_id: String },
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

    let Some(metadata_part) = find_metadata_part(pkg) else {
        return Ok(HashMap::new());
    };
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

        // See `parse_vm_to_rich_value_index_map` for rationale: keep the metadata mapping in its
        // canonical 1-based form and infer whether this worksheet uses 0-based `vm` indices by
        // checking for any `vm="0"` cells.
        //
        // As a small safety guard, only apply the offset when the mapping doesn't already contain
        // 0-based keys.
        let vm_offset: u32 = if !vm_to_rich_value.contains_key(&0)
            && cells.iter().any(|(_, vm)| *vm == 0)
        {
            1
        } else {
            0
        };

        for (cell, vm) in cells {
            let Some(vm) = vm.checked_add(vm_offset) else {
                continue;
            };
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

    let rich_value_to_image_pointer = parse_rich_value_parts_image_pointers(pkg, &rich_value_parts)?;
    if rich_value_to_image_pointer.is_empty() {
        return Ok(HashMap::new());
    }

    // Optional indirection parts used by the most common Excel layout.
    let rich_value_part_for_rels = rich_value_parts
        .iter()
        .copied()
        .find(|name| *name == "xl/richData/richValue.xml")
        .unwrap_or(rich_value_parts[0]);
    let rich_value_rel_part = find_rich_value_rel_part(pkg, rich_value_part_for_rels);
    let rel_index_to_rid = match rich_value_rel_part
        .as_deref()
        .and_then(|part| pkg.part(part))
    {
        Some(bytes) => parse_rich_value_rel_rids(bytes)?,
        None => Vec::new(),
    };
    let rid_to_target = match rich_value_rel_part
        .as_deref()
        .map(path::rels_for_part)
        .as_deref()
        .and_then(|part| pkg.part(part))
    {
        Some(bytes) => parse_rich_value_rel_rels(bytes)?,
        None => HashMap::new(),
    };

    // Cache parsed relationships for richValue*.xml parts when resolving direct rId references.
    let mut rich_value_part_rels: HashMap<String, HashMap<String, String>> = HashMap::new();

    let mut out: HashMap<(String, CellRef), Vec<u8>> = HashMap::new();
    for (sheet_name, cell, rich_value_idx) in cells_with_rich_value {
        let Some(pointer) = rich_value_to_image_pointer.get(&rich_value_idx) else {
            continue;
        };
        let target_part = match pointer {
            RichValueImagePointer::RelIndex(rel_index) => {
                let Some(rich_value_rel_part) = rich_value_rel_part.as_deref() else {
                    continue;
                };
                if rel_index_to_rid.is_empty() || rid_to_target.is_empty() {
                    continue;
                }

                let Some(rid) = rel_slot_get(&rel_index_to_rid, *rel_index as usize) else {
                    continue;
                };
                let Some(target) = rid_to_target.get(rid.as_str()) else {
                    continue;
                };
                resolve_rich_value_rel_target_part(rich_value_rel_part, target)
            }
            RichValueImagePointer::DirectRelId {
                source_part,
                rel_id,
            } => {
                let rels = if let Some(cached) = rich_value_part_rels.get(source_part) {
                    cached
                } else {
                    let parsed = parse_rich_value_part_relationship_targets(pkg, source_part)?;
                    rich_value_part_rels.insert(source_part.clone(), parsed);
                    let Some(rels) = rich_value_part_rels.get(source_part) else {
                        // Best-effort: if we cannot read back what we just inserted (should be
                        // impossible in normal operation), skip this cell rather than panicking.
                        debug_assert!(false, "rich value rels cache insert could not be read back");
                        continue;
                    };
                    rels
                };

                let Some(target_part) = rels.get(rel_id.as_str()) else {
                    continue;
                };
                target_part.to_string()
            }
        };
        let Some(bytes) = pkg.part(&target_part) else {
            continue;
        };
        out.insert((sheet_name, cell), bytes.to_vec());
    }

    Ok(out)
}

/// Best-effort discovery of the workbook metadata part (`xl/metadata*.xml`).
///
/// Excel has historically used `xl/metadata.xml`, but real-world packages may:
/// - reference the metadata part via `xl/_rels/workbook.xml.rels`
/// - use numbered part names like `xl/metadata1.xml`
///
/// This helper prefers relationship-based resolution and falls back to filename scans.
pub(crate) fn find_metadata_part(pkg: &XlsxPackage) -> Option<String> {
    const WORKBOOK_PART: &str = "xl/workbook.xml";
    const WORKBOOK_RELS_PART: &str = "xl/_rels/workbook.xml.rels";

    if let Some(workbook_rels) = pkg.part(WORKBOOK_RELS_PART) {
        if let Ok(rels) = crate::openxml::parse_relationships(workbook_rels) {
            for rel in rels {
                if rel
                    .target_mode
                    .as_deref()
                    .is_some_and(|m| m.trim().eq_ignore_ascii_case("External"))
                {
                    continue;
                }

                // Prefer explicit metadata relationship types, but also accept any target that ends
                // with `metadata.xml` (e.g. `custom-metadata.xml` or `xl/metadata.xml`).
                if rel.target.ends_with("metadata.xml") || rel.type_uri.ends_with("/metadata") {
                    let resolved = path::resolve_target(WORKBOOK_PART, &rel.target);
                    if pkg.part(&resolved).is_some() {
                        return Some(resolved);
                    }
                }
            }
        }
    }

    if pkg.part("xl/metadata.xml").is_some() {
        return Some("xl/metadata.xml".to_string());
    }

    find_lowest_numbered_part(pkg, "xl/metadata", ".xml")
}

fn resolve_workbook_metadata_part_name(pkg: &XlsxPackage) -> Result<String, RichDataError> {
    Ok(find_metadata_part(pkg).unwrap_or_else(|| "xl/metadata.xml".to_string()))
}

/// Best-effort discovery of the richValueRel part (`xl/richData/richValueRel*.xml`) for a chosen
/// richValue part.
///
/// Excel typically stores this at `xl/richData/richValueRel.xml`, but it can also be referenced
/// via the richValue part's relationships (`xl/richData/_rels/richValue*.xml.rels`) and/or use a
/// numbered filename.
pub(crate) fn find_rich_value_rel_part(pkg: &XlsxPackage, rich_value_part: &str) -> Option<String> {
    if pkg.part("xl/richData/richValueRel.xml").is_some() {
        return Some("xl/richData/richValueRel.xml".to_string());
    }

    let rich_value_rels_part = path::rels_for_part(rich_value_part);
    if let Some(rich_value_rels_xml) = pkg.part(&rich_value_rels_part) {
        if let Ok(rels) = crate::openxml::parse_relationships(rich_value_rels_xml) {
            for rel in rels {
                if rel
                    .target_mode
                    .as_deref()
                    .is_some_and(|m| m.trim().eq_ignore_ascii_case("External"))
                {
                    continue;
                }
                if rel.target.contains("richValueRel") {
                    let resolved = path::resolve_target(rich_value_part, &rel.target);
                    if pkg.part(&resolved).is_some() {
                        return Some(resolved);
                    }
                }
            }
        }
    }

    find_lowest_numbered_part(pkg, "xl/richData/richValueRel", ".xml")
}

fn find_lowest_numbered_part(pkg: &XlsxPackage, prefix: &str, suffix: &str) -> Option<String> {
    let mut best: Option<(u32, String)> = None;

    for part in pkg.part_names() {
        let Some(num) = numeric_suffix(part, prefix, suffix) else {
            continue;
        };

        match &mut best {
            Some((best_num, best_name)) => {
                if num < *best_num || (num == *best_num && part < best_name.as_str()) {
                    *best_num = num;
                    *best_name = part.to_string();
                }
            }
            None => best = Some((num, part.to_string())),
        }
    }

    best.map(|(_, name)| name)
}

fn numeric_suffix(part_name: &str, prefix: &str, suffix: &str) -> Option<u32> {
    let mid = part_name.strip_prefix(prefix)?.strip_suffix(suffix)?;
    if !mid.chars().all(|c| c.is_ascii_digit()) {
        return None;
    }
    if mid.is_empty() {
        return Some(0);
    }
    mid.parse::<u32>().ok()
}

fn build_rich_value_rel_index_to_target_part(
    pkg: &XlsxPackage,
) -> Result<Vec<Option<String>>, RichDataError> {
    // `richValueRel.xml` is the canonical name for the relationship-slot table, but in the wild/tests
    // it may be numbered/custom (e.g. `richValueRel1.xml`, `customRichValueRel.xml`). Prefer the
    // canonical name when present and fall back to best-effort discovery.
    let Some(source_part) = find_any_rich_value_rel_part(pkg) else {
        return Ok(Vec::new());
    };
    let Some(rich_value_rel_bytes) = pkg.part(&source_part) else {
        return Ok(Vec::new());
    };
    let rel_index_to_rid = rich_value_rel::parse_rich_value_rel_table(rich_value_rel_bytes)
        .map_err(RichDataError::from)?;
    if rel_index_to_rid.is_empty() {
        return Ok(Vec::new());
    }

    let rels_part_name = path::rels_for_part(&source_part);
    let Some(rels_bytes) = pkg.part(&rels_part_name) else {
        let mut out: Vec<Option<String>> = Vec::new();
        if out.try_reserve_exact(rel_index_to_rid.len()).is_err() {
            return Err(XlsxError::AllocationFailure(
                "build_rich_value_rel_index_to_target_part missing rels fallback",
            )
            .into());
        }
        out.resize_with(rel_index_to_rid.len(), || None);
        return Ok(out);
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

        let mut target_part = resolve_rich_value_rel_target_part(&source_part, target);

        // Additional tolerance: if the resolved target does not exist in the package, attempt to
        // correct common relative-path mistakes.
        if pkg.part(&target_part).is_none() {
            if let Some(rest) = target_part.strip_prefix("xl/richData/") {
                if rest.starts_with("media/") {
                    let alt = format!("xl/{rest}");
                    if pkg.part(&alt).is_some() {
                        target_part = alt;
                    }
                } else if rest.starts_with("xl/") {
                    if pkg.part(rest).is_some() {
                        target_part = rest.to_string();
                    }
                }
            }
        }

        rid_to_target_part.insert(rel.id, target_part);
    }

    let mut out: Vec<Option<String>> = Vec::new();
    if out.try_reserve_exact(rel_index_to_rid.len()).is_err() {
        return Err(
            XlsxError::AllocationFailure("build_rich_value_rel_index_to_target_part output").into(),
        );
    }
    for rid in rel_index_to_rid {
        if rid.is_empty() {
            out.push(None);
        } else {
            out.push(rid_to_target_part.get(&rid).cloned());
        }
    }

    Ok(out)
}

pub(crate) fn resolve_rich_value_rel_target_part(source_part: &str, target: &str) -> String {
    // Relationship targets in `xl/richData/_rels/*.rels` are typically relative to `xl/richData/`
    // (e.g. `../media/image1.png`). Some producers instead emit targets relative to `xl/`
    // (e.g. `media/image1.png`), or emit `Target="xl/..."` without a leading `/` (which would
    // otherwise resolve to `xl/richData/xl/...`). Handle these as special-cases for robust
    // extraction.
    let target = strip_fragment(target);
    // Be resilient to invalid/unescaped Windows-style path separators.
    let target: std::borrow::Cow<'_, str> = if target.contains('\\') {
        std::borrow::Cow::Owned(target.replace('\\', "/"))
    } else {
        std::borrow::Cow::Borrowed(target)
    };
    let target = target.as_ref();
    let target = target.strip_prefix("./").unwrap_or(target);
    if target.starts_with("media/") {
        format!("xl/{target}")
    } else if target.starts_with("xl/") {
        target.to_string()
    } else {
        path::resolve_target(source_part, target)
    }
}

fn find_any_rich_value_rel_part(pkg: &XlsxPackage) -> Option<String> {
    // Canonical first.
    if pkg.part("xl/richData/richValueRel.xml").is_some() {
        return Some("xl/richData/richValueRel.xml".to_string());
    }

    // Next, prefer numbered variants like `xl/richData/richValueRel1.xml`.
    if let Some(part) = find_lowest_numbered_part(pkg, "xl/richData/richValueRel", ".xml") {
        return Some(part);
    }

    // Last-ditch fallback: any XML part under `xl/richData/` whose filename contains `richValueRel`
    // (case-insensitive). Some tests use custom names like `customRichValueRel.xml`.
    let mut candidates: Vec<&str> = pkg
        .part_names()
        .filter(|name| name.starts_with("xl/richData/") && crate::ascii::ends_with_ignore_case(name, ".xml"))
        .filter(|name| crate::ascii::contains_ignore_case(name, "richvaluerel"))
        .collect();
    if candidates.is_empty() {
        return None;
    }

    candidates.sort_by(|a, b| {
        fn rank(name: &str) -> (u8, usize) {
            let file = name.rsplit('/').next().unwrap_or(name);
            let prefix_rank = if crate::ascii::starts_with_ignore_case(file, "richvaluerel") {
                0
            } else {
                1
            };
            (prefix_rank, file.len())
        }

        let (a_rank, a_len) = rank(a);
        let (b_rank, b_len) = rank(b);
        a_rank
            .cmp(&b_rank)
            .then(a_len.cmp(&b_len))
            .then_with(|| a.cmp(b))
    });

    Some(candidates[0].to_string())
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
enum RichValuePartFamily {
    RichValue,
    RichValues,
    RdRichValue,
}

/// Parse an `xl/richData/*` rich value part name and return its "family" plus numeric suffix.
///
/// Examples:
/// - `xl/richData/richValue.xml` -> (`RichValue`, 0)
/// - `xl/richData/richValue2.xml` -> (`RichValue`, 2)
/// - `xl/richData/richValues.xml` -> (`RichValues`, 0)
/// - `xl/richData/richValues2.xml` -> (`RichValues`, 2)
/// - `xl/richData/rdrichvalue10.xml` -> (`RdRichValue`, 10)
///
/// This is case-insensitive for the filename (but not the containing directory).
fn parse_rich_value_part_name(part_path: &str) -> Option<(RichValuePartFamily, u32)> {
    if !part_path.starts_with("xl/richData/") {
        return None;
    }

    let file_name = part_path.rsplit('/').next()?;
    let stem = crate::ascii::strip_suffix_ignore_case(file_name, ".xml")?;

    // Check the plural prefix first: `richvalues` starts with `richvalue`.
    let (family, suffix) = if let Some(rest) = crate::ascii::strip_prefix_ignore_case(stem, "richvalues") {
        (RichValuePartFamily::RichValues, rest)
    } else if let Some(rest) = crate::ascii::strip_prefix_ignore_case(stem, "richvalue") {
        (RichValuePartFamily::RichValue, rest)
    } else if let Some(rest) = crate::ascii::strip_prefix_ignore_case(stem, "rdrichvalue") {
        (RichValuePartFamily::RdRichValue, rest)
    } else {
        return None;
    };

    let idx = if suffix.is_empty() {
        0
    } else if suffix.as_bytes().iter().all(u8::is_ascii_digit) {
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
        // The structured metadata parser returns a canonical mapping keyed by the `<valueMetadata>`
        // `<bk>` record index, which is **1-based**.
        //
        // Some worksheets encode `c/@vm` as 0-based in the wild. Attempting to store both 0-based
        // and 1-based keys in a single map is ambiguous when multiple `<valueMetadata>` records
        // exist (e.g. `vm="1"` could refer to the first record in a 1-based scheme or the second
        // record in a 0-based scheme).
        //
        // Callers that have access to the worksheet cells should infer that offset (e.g. by
        // checking for any `vm="0"` cells) and adjust worksheet `vm` values before using this map.
        return Ok(primary);
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
    // Canonical worksheet `c/@vm` indexing is 1-based (vm=1 refers to the first `<bk>` entry).
    // Some producers have been observed to emit 0-based worksheet indices; callers can handle that
    // via the same per-worksheet offset inference used for the structured parser.
    let mut vm_idx: u32 = 1;
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
                let Some(vm_key) = vm_idx.checked_add(offset) else {
                    break;
                };
                out.insert(vm_key, rich_value_idx);
            }
            break;
        }

        let Some(next_vm_idx) = vm_idx.checked_add(count) else {
            break;
        };
        vm_idx = next_vm_idx;
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

fn parse_rich_value_parts_image_pointers(
    pkg: &XlsxPackage,
    part_names: &[&str],
) -> Result<HashMap<u32, RichValueImagePointer>, RichDataError> {
    // Split part names by family so we can apply schema-aware parsing to `richValue*.xml` and
    // structure-aware parsing to `rdrichvalue*.xml`.
    let mut rich_value_parts: Vec<&str> = Vec::new();
    let mut rd_rich_value_parts: Vec<&str> = Vec::new();
    for part_name in part_names {
        let Some((family, _idx)) = parse_rich_value_part_name(part_name) else {
            continue;
        };
        match family {
            RichValuePartFamily::RichValue | RichValuePartFamily::RichValues => {
                rich_value_parts.push(*part_name)
            }
            RichValuePartFamily::RdRichValue => rd_rich_value_parts.push(*part_name),
        }
    }

    rich_value_parts.sort_by(|a, b| cmp_rich_value_parts_by_numeric_suffix(a, b));
    rd_rich_value_parts.sort_by(|a, b| cmp_rich_value_parts_by_numeric_suffix(a, b));

    let mut out: HashMap<u32, RichValueImagePointer> = HashMap::new();
    let mut global_idx: u32 = 0;

    // Optional optimization: when present, `richValueTypes.xml` (and optionally
    // `richValueStructure.xml`) can describe the schema of each rich value type and which property
    // corresponds to a relationship (e.g. an image).
    //
    // This is best-effort: malformed/unexpected schemas are ignored and we fall back to the
    // heuristic `<v t="rel">` / first-numeric parsing.
    let rich_value_types = build_rich_value_type_schema_index_best_effort(pkg);

    for part_name in rich_value_parts {
        let Some(bytes) = pkg.part(part_name) else {
            continue;
        };
        let xml = std::str::from_utf8(bytes).map_err(|source| RichDataError::XmlNonUtf8 {
            part: part_name.to_string(),
            source,
        })?;
        let doc = roxmltree::Document::parse(xml).map_err(|source| RichDataError::XmlParse {
            part: part_name.to_string(),
            source,
        })?;

        for rv in doc
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "rv")
        {
            if let Some(rel_idx) = parse_rv_rel_index(rv, rich_value_types.as_ref()) {
                out.insert(global_idx, RichValueImagePointer::RelIndex(rel_idx));
            } else if let Some(rel_id) = parse_rv_direct_rel_id(rv) {
                out.insert(
                    global_idx,
                    RichValueImagePointer::DirectRelId {
                        source_part: part_name.to_string(),
                        rel_id,
                    },
                );
            }
            let Some(next) = global_idx.checked_add(1) else {
                return Ok(out);
            };
            global_idx = next;
        }
    }

    if !rd_rich_value_parts.is_empty() {
        let rd_structure_xml = pkg.part("xl/richData/rdrichvaluestructure.xml");

        for part_name in rd_rich_value_parts {
            let Some(bytes) = pkg.part(part_name) else {
                continue;
            };
            let rel_indices = images::parse_rdrichvalue_relationship_indices(bytes, rd_structure_xml)?;
            for rel_idx in rel_indices {
                if let Some(rel_idx) = rel_idx {
                    if let Ok(rel_idx) = u32::try_from(rel_idx) {
                        out.insert(global_idx, RichValueImagePointer::RelIndex(rel_idx));
                    }
                }
                let Some(next) = global_idx.checked_add(1) else {
                    return Ok(out);
                };
                global_idx = next;
            }
        }
    }

    Ok(out)
}

fn parse_rich_value_parts_rel_indices(
    pkg: &XlsxPackage,
    part_names: &[&str],
) -> Result<HashMap<u32, u32>, RichDataError> {
    // Split part names by family so we can apply schema-aware parsing to `richValue*.xml` and
    // structure-aware parsing to `rdrichvalue*.xml`.
    //
    // The `rdRichValue` schema stores positional `<v>` elements, with the member ordering described
    // by `rdrichvaluestructure.xml`. In particular, embedded images use a `_rvRel:LocalImageIdentifier`
    // field that points into `richValueRel.xml`. Some workbooks reorder fields so the relationship
    // index is not the first numeric `<v>`, so we must consult the structure metadata when present.
    let mut rich_value_parts: Vec<&str> = Vec::new();
    let mut rd_rich_value_parts: Vec<&str> = Vec::new();
    for part_name in part_names {
        let Some((family, _idx)) = parse_rich_value_part_name(part_name) else {
            continue;
        };
        match family {
            RichValuePartFamily::RichValue | RichValuePartFamily::RichValues => {
                rich_value_parts.push(*part_name)
            }
            RichValuePartFamily::RdRichValue => rd_rich_value_parts.push(*part_name),
        }
    }

    rich_value_parts.sort_by(|a, b| cmp_rich_value_parts_by_numeric_suffix(a, b));
    rd_rich_value_parts.sort_by(|a, b| cmp_rich_value_parts_by_numeric_suffix(a, b));

    let mut out: HashMap<u32, u32> = HashMap::new();
    let mut global_idx: u32 = 0;

    // Optional optimization: when present, `richValueTypes.xml` (and optionally
    // `richValueStructure.xml`) can describe the schema of each rich value type and which property
    // corresponds to a relationship (e.g. an image).
    //
    // This is best-effort: malformed/unexpected schemas are ignored and we fall back to the
    // heuristic `<v t="rel">` / first-numeric parsing.
    let rich_value_types = build_rich_value_type_schema_index_best_effort(pkg);

    for part_name in rich_value_parts {
        let Some(bytes) = pkg.part(part_name) else {
            continue;
        };
        let xml = std::str::from_utf8(bytes).map_err(|source| RichDataError::XmlNonUtf8 {
            part: part_name.to_string(),
            source,
        })?;
        let doc = roxmltree::Document::parse(xml).map_err(|source| RichDataError::XmlParse {
            part: part_name.to_string(),
            source,
        })?;

        for rv in doc
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "rv")
        {
            if let Some(rel_idx) = parse_rv_rel_index(rv, rich_value_types.as_ref()) {
                out.insert(global_idx, rel_idx);
            }
            let Some(next) = global_idx.checked_add(1) else {
                return Ok(out);
            };
            global_idx = next;
        }
    }

    if !rd_rich_value_parts.is_empty() {
        let rd_structure_xml = pkg.part("xl/richData/rdrichvaluestructure.xml");

        for part_name in rd_rich_value_parts {
            let Some(bytes) = pkg.part(part_name) else {
                continue;
            };

            let rel_indices =
                images::parse_rdrichvalue_relationship_indices(bytes, rd_structure_xml)?;
            for rel_idx in rel_indices {
                if let Some(rel_idx) = rel_idx {
                    if let Ok(rel_idx) = u32::try_from(rel_idx) {
                        out.insert(global_idx, rel_idx);
                    }
                }
                let Some(next) = global_idx.checked_add(1) else {
                    return Ok(out);
                };
                global_idx = next;
            }
        }
    }

    Ok(out)
}

fn parse_rv_direct_rel_id(rv: roxmltree::Node<'_, '_>) -> Option<String> {
    // Prefer `<v>` descendant text.
    for v in rv
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "v")
    {
        let text = v.text()?.trim();
        if is_rid(text) {
            return Some(text.to_string());
        }
    }

    // Also search any attribute value in the `<rv>` subtree (some producers store rIds there).
    for node in rv.descendants().filter(|n| n.is_element()) {
        for attr in node.attributes() {
            let value = attr.value().trim();
            if is_rid(value) {
                return Some(value.to_string());
            }
        }
    }

    None
}

fn parse_rich_value_part_relationship_targets(
    pkg: &XlsxPackage,
    rich_value_part: &str,
) -> Result<HashMap<String, String>, RichDataError> {
    let rels_part = path::rels_for_part(rich_value_part);
    let Some(rels_bytes) = pkg.part(&rels_part) else {
        return Ok(HashMap::new());
    };

    let rels = crate::openxml::parse_relationships(rels_bytes)?;
    let mut out: HashMap<String, String> = HashMap::new();
    for rel in rels {
        if rel
            .target_mode
            .as_deref()
            .is_some_and(|m| m.trim().eq_ignore_ascii_case("External"))
        {
            continue;
        }
        if rel.type_uri != crate::drawings::REL_TYPE_IMAGE {
            continue;
        }

        let target = strip_fragment(&rel.target);
        if target.is_empty() {
            continue;
        }
        let target_part = resolve_rich_value_rel_target_part(rich_value_part, target);
        out.insert(rel.id, target_part);
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

                    let mut props = Vec::new();
                    if props.try_reserve_exact(structure.members.len()).is_err() {
                        continue;
                    }
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
                if !crate::ascii::ends_with_ignore_case(name, "type") {
                    return false;
                }
                n.parent_element()
                    .is_some_and(|p| {
                        let name = p.tag_name().name();
                        crate::ascii::ends_with_ignore_case(name, "types")
                    })
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
                    && {
                        let name = n.tag_name().name();
                        name.eq_ignore_ascii_case("props")
                            || name.eq_ignore_ascii_case("properties")
                            || name.eq_ignore_ascii_case("fields")
                    }
            })
            .unwrap_or(ty);

        let prop_nodes: Vec<roxmltree::Node<'_, '_>> = prop_container
            .children()
            .filter(|n| n.is_element() && looks_like_rich_value_type_property_def(*n))
            .collect();
        if prop_nodes.is_empty() {
            continue;
        }

        let mut props = Vec::new();
        if props.try_reserve_exact(prop_nodes.len()).is_err() {
            continue;
        }
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
                if !crate::ascii::ends_with_ignore_case(name, "type") {
                    return false;
                }
                n.parent_element()
                    .is_some_and(|p| crate::ascii::ends_with_ignore_case(p.tag_name().name(), "types"))
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
    let kind = desc.kind.as_deref().unwrap_or_default();
    if kind.eq_ignore_ascii_case("rel") || kind.eq_ignore_ascii_case("relationship") {
        return true;
    }

    let name = desc.name.as_deref().unwrap_or_default();
    if name.eq_ignore_ascii_case("rel") || name.eq_ignore_ascii_case("relationship") {
        return true;
    }

    // Some schemas might use a property key like `image`/`img` for media relationships.
    crate::ascii::contains_ignore_case(name, "image") || crate::ascii::contains_ignore_case(name, "img")
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
        let name = local_name(attr.name());
        if name.eq_ignore_ascii_case("t")
            || name.eq_ignore_ascii_case("type")
            || name.eq_ignore_ascii_case("kind")
            || name.eq_ignore_ascii_case("name")
            || name.eq_ignore_ascii_case("n")
            || name.eq_ignore_ascii_case("id")
            || name.eq_ignore_ascii_case("key")
            || name.eq_ignore_ascii_case("k")
            || name.eq_ignore_ascii_case("pid")
            || name.eq_ignore_ascii_case("propid")
        {
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
        // Typically emitted as `r:id="rId..."` (prefix varies), but be robust to prefix
        // differences and missing namespace declarations.
        let Some(rid) = rel.attribute((REL_NS, "id")).or_else(|| {
            rel.attributes()
                .find(|attr| attr.name() == "id" && is_rid(attr.value()))
                .map(|attr| attr.value())
        }) else {
            continue;
        };
        out.push(rid.to_string());
    }

    Ok(out)
}

fn is_rid(value: &str) -> bool {
    let Some(suffix) = value.strip_prefix("rId") else {
        return false;
    };
    !suffix.is_empty() && suffix.as_bytes().iter().all(|b| b.is_ascii_digit())
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

#[cfg(test)]
mod tests {
    use super::*;

    use std::io::{Cursor, Write};

    use zip::write::FileOptions;
    use zip::ZipWriter;

    fn build_package(entries: &[(&str, &[u8])]) -> XlsxPackage {
        let cursor = Cursor::new(Vec::new());
        let mut zip = ZipWriter::new(cursor);
        let options =
            FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

        for (name, bytes) in entries {
            zip.start_file(*name, options).unwrap();
            zip.write_all(bytes).unwrap();
        }

        let bytes = zip.finish().unwrap().into_inner();
        XlsxPackage::from_bytes(&bytes).expect("read test pkg")
    }

    #[test]
    fn vm_index_infers_one_based_metadata_offset_when_sheet_uses_vm_zero() {
        // Build a minimal package where:
        // - `xl/metadata.xml` uses the structured `metadataTypes`/`futureMetadata` layout, which
        //   yields a **1-based** vm -> rich value mapping.
        // - the worksheet uses `vm="0"` (observed in the wild for some producers).
        //
        // `RichDataVmIndex` should detect that mismatch and apply an offset so `resolve_vm(0)`
        // successfully resolves a media target.
        let sheet_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" vm="0"/>
    </row>
  </sheetData>
</worksheet>"#;

        let metadata_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <metadataTypes count="1">
    <metadataType name="XLRICHVALUE"/>
  </metadataTypes>
  <futureMetadata name="XLRICHVALUE" count="1">
    <bk>
      <xlrd:rvb i="0"/>
    </bk>
  </futureMetadata>
  <valueMetadata count="1">
    <bk>
      <!-- `t=\"1\"` selects the first metadataType entry using 1-based indexing. -->
      <rc t="1" v="0"/>
    </bk>
  </valueMetadata>
</metadata>"#;

        let rich_value_rel_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvRel xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rels>
    <rel r:id="rId1"/>
  </rels>
</rvRel>"#;

        let rich_value_rel_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>"#;

        let pkg = build_package(&[
            ("xl/worksheets/sheet1.xml", sheet_xml),
            ("xl/metadata.xml", metadata_xml),
            ("xl/richData/richValueRel.xml", rich_value_rel_xml),
            ("xl/richData/_rels/richValueRel.xml.rels", rich_value_rel_rels),
        ]);

        let index = RichDataVmIndex::build(&pkg).expect("build vm index");
        let resolved = index.resolve_vm(0);
        assert_eq!(
            resolved,
            RichDataVmResolution {
                rich_value_index: Some(0),
                rel_index: Some(0),
                target_part: Some("xl/media/image1.png".to_string()),
            }
        );
    }

    #[test]
    fn extracts_rich_cell_image_with_non_r_prefix_in_rich_value_rel() {
        let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

        let workbook_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

        let sheet_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" vm="1"/>
    </row>
  </sheetData>
</worksheet>"#;

        let metadata_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <metadataTypes count="1">
    <metadataType name="XLRICHVALUE"/>
  </metadataTypes>
  <futureMetadata name="XLRICHVALUE" count="1">
    <bk>
      <extLst>
        <ext uri="{00000000-0000-0000-0000-000000000000}">
          <xlrd:rvb i="0"/>
        </ext>
      </extLst>
    </bk>
  </futureMetadata>
  <valueMetadata count="1">
    <bk>
      <rc t="1" v="0"/>
    </bk>
  </valueMetadata>
</metadata>"#;

        let rich_value_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValue xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <rv>
    <v t="rel">0</v>
  </rv>
</richValue>"#;

        let rich_value_rel_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueRels xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rel rel:id="rId1"/>
</richValueRels>"#;

        let rich_value_rel_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>"#;

        let pkg = build_package(&[
            ("xl/workbook.xml", workbook_xml),
            ("xl/_rels/workbook.xml.rels", workbook_rels),
            ("xl/worksheets/sheet1.xml", sheet_xml),
            ("xl/metadata.xml", metadata_xml),
            ("xl/richData/richValue1.xml", rich_value_xml),
            ("xl/richData/richValueRel.xml", rich_value_rel_xml),
            ("xl/richData/_rels/richValueRel.xml.rels", rich_value_rel_rels),
            ("xl/media/image1.png", b"png-bytes"),
        ]);

        let images = extract_rich_cell_images(&pkg).expect("extract cell images");

        let key = ("Sheet1".to_string(), CellRef::from_a1("A1").unwrap());
        assert_eq!(
            images.get(&key).map(|bytes| bytes.as_slice()),
            Some(b"png-bytes".as_slice())
        );
    }
}
