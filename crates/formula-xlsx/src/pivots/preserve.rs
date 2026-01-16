use std::borrow::Cow;
use std::collections::{BTreeMap, HashMap, HashSet};
use std::io::{Read, Seek, SeekFrom};

use formula_model::sheet_name_eq_case_insensitive;
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};
use roxmltree::Document;
use zip::ZipArchive;

use crate::path::rels_for_part;
pub use crate::preserve::rels_merge::RelationshipStub;
use crate::preserve::rels_merge::{ensure_rels_has_relationships, xml_escape};
use crate::preserve::sheet_match::{
    match_sheet_by_name_or_index, workbook_sheet_parts, workbook_sheet_parts_from_workbook_xml,
};
use crate::relationships::parse_relationships;
use crate::workbook::ChartExtractionError;
use crate::zip_util::{ZipInflateBudget, DEFAULT_MAX_ZIP_PART_BYTES, DEFAULT_MAX_ZIP_TOTAL_BYTES};
use crate::XlsxPackage;

const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const SPREADSHEETML_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const PIVOT_CACHE_DEF_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";
const PIVOT_TABLE_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";
const SLICER_CACHE_REL_TYPE: &str =
    "http://schemas.microsoft.com/office/2007/relationships/slicerCache";
const TIMELINE_CACHE_DEF_REL_TYPE: &str =
    "http://schemas.microsoft.com/office/2007/relationships/timelineCacheDefinition";

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PreservedSheetPivotTables {
    pub sheet_index: usize,
    pub sheet_id: Option<u32>,
    /// The `<pivotTables>` subtree from the worksheet XML (outer XML).
    pub pivot_tables_xml: Vec<u8>,
    /// Relationships from the worksheet `.rels` required by `<pivotTables>`.
    pub pivot_table_rels: Vec<RelationshipStub>,
}

/// A slice of an XLSX package required to preserve pivot-related parts across a
/// "read -> write" pipeline.
///
/// Excel stores pivot tables, caches, slicers, and timelines in additional XML
/// parts that `rust_xlsxwriter` cannot generate. Call
/// [`XlsxPackage::preserve_pivot_parts`] before regenerating a workbook, then
/// re-apply the returned payload with
/// [`XlsxPackage::apply_preserved_pivot_parts`] to retain pivot artifacts.
///
/// The preserved data is keyed by the original worksheet name *and* stores the
/// worksheet's position in the workbook sheet list so attachments can be
/// restored after in-app sheet renames (without reordering).
#[derive(Debug, Clone)]
pub struct PreservedPivotParts {
    /// Source `[Content_Types].xml` for merging required overrides.
    pub content_types_xml: Vec<u8>,
    /// Raw pivot/slicer/timeline parts copied byte-for-byte.
    pub parts: BTreeMap<String, Vec<u8>>,
    /// Worksheet list from the source workbook (`xl/workbook.xml`) in display order.
    ///
    /// This enables name rewrites for pivot cache `worksheetSource` references when worksheets are
    /// renamed between preservation and re-application.
    ///
    /// Note: This intentionally excludes chart sheets (and other non-worksheet sheet types) so the
    /// stored `index` remains stable even when chart sheets are re-attached at a different position
    /// during regeneration-based saves.
    pub workbook_sheets: Vec<PreservedWorkbookSheet>,
    /// The `<pivotCaches>` subtree from `xl/workbook.xml` (outer XML).
    pub workbook_pivot_caches: Option<Vec<u8>>,
    /// Relationships from `xl/_rels/workbook.xml.rels` required by `<pivotCaches>`.
    pub workbook_pivot_cache_rels: Vec<RelationshipStub>,
    /// The `<slicerCaches>` subtree from `xl/workbook.xml` (outer XML).
    pub workbook_slicer_caches: Option<Vec<u8>>,
    /// Relationships from `xl/_rels/workbook.xml.rels` required by `<slicerCaches>`.
    pub workbook_slicer_cache_rels: Vec<RelationshipStub>,
    /// The `<timelineCaches>` subtree from `xl/workbook.xml` (outer XML).
    pub workbook_timeline_caches: Option<Vec<u8>>,
    /// Relationships from `xl/_rels/workbook.xml.rels` required by `<timelineCaches>`.
    pub workbook_timeline_cache_rels: Vec<RelationshipStub>,
    /// Preserved `<pivotTables>` subtrees and `.rels` metadata per worksheet.
    pub sheet_pivot_tables: BTreeMap<String, PreservedSheetPivotTables>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PreservedWorkbookSheet {
    pub name: String,
    pub index: usize,
}

/// Streaming variant of [`XlsxPackage::preserve_pivot_parts`].
///
/// This reads only the subset of ZIP parts required to retain pivot tables, caches, slicers, and
/// timelines for a later regeneration-based round-trip.
///
/// Unlike [`XlsxPackage::from_bytes`], this does **not** inflate every ZIP entry into memory.
pub fn preserve_pivot_parts_from_reader<R: Read + Seek>(
    reader: R,
) -> Result<PreservedPivotParts, ChartExtractionError> {
    preserve_pivot_parts_from_reader_limited(reader, DEFAULT_MAX_ZIP_PART_BYTES, DEFAULT_MAX_ZIP_TOTAL_BYTES)
}

/// Streaming variant of [`XlsxPackage::preserve_pivot_parts`] with configurable ZIP inflation
/// limits.
///
/// This is primarily useful for callers that treat the input as untrusted and want tighter bounds
/// than the crate defaults.
pub fn preserve_pivot_parts_from_reader_limited<R: Read + Seek>(
    mut reader: R,
    max_part_bytes: u64,
    max_total_bytes: u64,
) -> Result<PreservedPivotParts, ChartExtractionError> {
    reader
        .seek(SeekFrom::Start(0))
        .map_err(|e| ChartExtractionError::XmlStructure(format!("io error: {e}")))?;
    let mut archive = ZipArchive::new(reader)
        .map_err(|e| ChartExtractionError::XmlStructure(format!("zip error: {e}")))?;

    let mut part_names: HashSet<String> = HashSet::new();
    for i in 0..archive.len() {
        let file = archive
            .by_index(i)
            .map_err(|e| ChartExtractionError::XmlStructure(format!("zip error: {e}")))?;
        if file.is_dir() {
            continue;
        }
        let name = file.name();
        part_names.insert(name.strip_prefix('/').unwrap_or(name).to_string());
    }

    let mut budget = ZipInflateBudget::new(max_total_bytes);

    let content_types_xml =
        read_zip_part_required(&mut archive, "[Content_Types].xml", max_part_bytes, &mut budget)?;

    let mut parts = BTreeMap::new();
    for name in &part_names {
        if name.starts_with("xl/pivotTables/")
            || name.starts_with("xl/pivotCache/")
            || name.starts_with("xl/slicers/")
            || name.starts_with("xl/slicerCaches/")
            || name.starts_with("xl/timelines/")
            || name.starts_with("xl/timelineCaches/")
        {
            if let Some(bytes) = read_zip_part_optional(&mut archive, name, max_part_bytes, &mut budget)? {
                parts.insert(name.clone(), bytes);
            }
        }
    }

    let workbook_part = "xl/workbook.xml";
    let workbook_xml =
        read_zip_part_required(&mut archive, workbook_part, max_part_bytes, &mut budget)?;
    let workbook_xml_str = std::str::from_utf8(&workbook_xml)
        .map_err(|e| ChartExtractionError::XmlNonUtf8(workbook_part.to_string(), e))?;
    let workbook_doc = Document::parse(workbook_xml_str)
        .map_err(|e| ChartExtractionError::XmlParse(workbook_part.to_string(), e))?;

    let pivot_caches_node = workbook_doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "pivotCaches");
    let slicer_caches_node = workbook_doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "slicerCaches");
    let timeline_caches_node = workbook_doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "timelineCaches");

    let mut workbook_pivot_cache_rids: HashSet<String> = HashSet::new();
    if let Some(node) = pivot_caches_node {
        for rid in node
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "pivotCache")
            .filter_map(|n| {
                n.attribute((REL_NS, "id"))
                    .or_else(|| n.attribute("r:id"))
                    .or_else(|| n.attribute("id"))
            })
        {
            workbook_pivot_cache_rids.insert(rid.to_string());
        }
    }

    let mut workbook_slicer_cache_rids: HashSet<String> = HashSet::new();
    if let Some(node) = slicer_caches_node {
        for rid in node
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "slicerCache")
            .filter_map(|n| {
                n.attribute((REL_NS, "id"))
                    .or_else(|| n.attribute("r:id"))
                    .or_else(|| n.attribute("id"))
            })
        {
            workbook_slicer_cache_rids.insert(rid.to_string());
        }
    }

    let mut workbook_timeline_cache_rids: HashSet<String> = HashSet::new();
    if let Some(node) = timeline_caches_node {
        for rid in node
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "timelineCache")
            .filter_map(|n| {
                n.attribute((REL_NS, "id"))
                    .or_else(|| n.attribute("r:id"))
                    .or_else(|| n.attribute("id"))
            })
        {
            workbook_timeline_cache_rids.insert(rid.to_string());
        }
    }

    let workbook_rels_part = "xl/_rels/workbook.xml.rels";
    let workbook_rels_xml =
        read_zip_part_optional(&mut archive, workbook_rels_part, max_part_bytes, &mut budget)?;
    // Best-effort: some producers emit malformed rels. For pivot preservation we treat this as
    // "no relationships" instead of failing the whole extraction.
    let rel_map: HashMap<String, crate::relationships::Relationship> = match workbook_rels_xml
        .as_deref()
    {
        Some(workbook_rels_xml) => match parse_relationships(workbook_rels_xml, workbook_rels_part)
        {
            Ok(rels) => rels.into_iter().map(|r| (r.id.clone(), r)).collect(),
            Err(_) => HashMap::new(),
        },
        None => HashMap::new(),
    };

    // Only preserve the <pivotCaches> subtree when we can also preserve every referenced
    // pivotCacheDefinition relationship. Otherwise re-applying would introduce broken r:id
    // references in workbook.xml.
    let (workbook_pivot_caches, workbook_pivot_cache_rels) = match pivot_caches_node {
        Some(node) if !workbook_pivot_cache_rids.is_empty() => {
            let mut rels = Vec::new();
            let mut missing_rel = false;
            for rid in &workbook_pivot_cache_rids {
                match rel_map.get(rid) {
                    Some(rel) if rel.type_ == PIVOT_CACHE_DEF_REL_TYPE => {
                        rels.push(RelationshipStub {
                            rel_id: rid.clone(),
                            target: rel.target.clone(),
                        });
                    }
                    _ => {
                        missing_rel = true;
                        break;
                    }
                }
            }

            if missing_rel {
                (None, Vec::new())
            } else {
                (
                    Some(workbook_xml_str.as_bytes()[node.range()].to_vec()),
                    rels,
                )
            }
        }
        _ => (None, Vec::new()),
    };

    let (workbook_slicer_caches, workbook_slicer_cache_rels) = match slicer_caches_node {
        Some(node) if !workbook_slicer_cache_rids.is_empty() => {
            let mut rels = Vec::new();
            let mut missing_rel = false;
            for rid in &workbook_slicer_cache_rids {
                match rel_map.get(rid) {
                    Some(rel) if rel.type_ == SLICER_CACHE_REL_TYPE => {
                        rels.push(RelationshipStub {
                            rel_id: rid.clone(),
                            target: rel.target.clone(),
                        });
                    }
                    _ => {
                        missing_rel = true;
                        break;
                    }
                }
            }

            if missing_rel {
                (None, Vec::new())
            } else {
                (
                    Some(workbook_xml_str.as_bytes()[node.range()].to_vec()),
                    rels,
                )
            }
        }
        _ => (None, Vec::new()),
    };

    let (workbook_timeline_caches, workbook_timeline_cache_rels) = match timeline_caches_node {
        Some(node) if !workbook_timeline_cache_rids.is_empty() => {
            let mut rels = Vec::new();
            let mut missing_rel = false;
            for rid in &workbook_timeline_cache_rids {
                match rel_map.get(rid) {
                    Some(rel) if rel.type_ == TIMELINE_CACHE_DEF_REL_TYPE => {
                        rels.push(RelationshipStub {
                            rel_id: rid.clone(),
                            target: rel.target.clone(),
                        });
                    }
                    _ => {
                        missing_rel = true;
                        break;
                    }
                }
            }

            if missing_rel {
                (None, Vec::new())
            } else {
                (
                    Some(workbook_xml_str.as_bytes()[node.range()].to_vec()),
                    rels,
                )
            }
        }
        _ => (None, Vec::new()),
    };

    let sheets = workbook_sheet_parts_from_workbook_xml(
        &workbook_xml,
        workbook_rels_xml.as_deref(),
        |candidate| {
            part_names.contains(candidate)
                || part_names
                    .iter()
                    .any(|name| crate::zip_util::zip_part_names_equivalent(name, candidate))
        },
    )?;
    let workbook_sheets = sheets
        .iter()
        .filter(|sheet| sheet.part_name.starts_with("xl/worksheets/"))
        .enumerate()
        .map(|(index, sheet)| PreservedWorkbookSheet {
            name: sheet.name.clone(),
            index,
        })
        .collect::<Vec<_>>();
    let mut sheet_pivot_tables: BTreeMap<String, PreservedSheetPivotTables> = BTreeMap::new();

    let mut worksheet_index = 0usize;
    for sheet in &sheets {
        // Pivot tables can only live on worksheet parts. Skip chart sheets and other sheet types so
        // the stored `sheet_index` matches the worksheet-only order used for rename mapping.
        if !sheet.part_name.starts_with("xl/worksheets/") {
            continue;
        }
        let sheet_index = worksheet_index;
        worksheet_index += 1;

        let sheet_rels_part = rels_for_part(&sheet.part_name);
        let Some(sheet_rels_xml) =
            read_zip_part_optional(&mut archive, &sheet_rels_part, max_part_bytes, &mut budget)?
        else {
            continue;
        };
        let rels = match parse_relationships(&sheet_rels_xml, &sheet_rels_part) {
            Ok(rels) => rels,
            Err(_) => continue,
        };

        // Fast-path: if the sheet has no pivotTable relationships, it cannot contain a valid
        // `<pivotTables>` block we can re-apply.
        if !rels.iter().any(|rel| rel.type_ == PIVOT_TABLE_REL_TYPE) {
            continue;
        }

        let Some(sheet_xml) =
            read_zip_part_optional(&mut archive, &sheet.part_name, max_part_bytes, &mut budget)?
        else {
            continue;
        };
        let sheet_xml_str = std::str::from_utf8(&sheet_xml)
            .map_err(|e| ChartExtractionError::XmlNonUtf8(sheet.part_name.clone(), e))?;
        let sheet_doc = Document::parse(sheet_xml_str)
            .map_err(|e| ChartExtractionError::XmlParse(sheet.part_name.clone(), e))?;

        let pivot_tables_node = sheet_doc
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "pivotTables");
        let Some(pivot_tables_node) = pivot_tables_node else {
            continue;
        };

        let pivot_tables_xml = sheet_xml_str.as_bytes()[pivot_tables_node.range()].to_vec();

        let pivot_table_rids: Vec<String> = pivot_tables_node
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "pivotTable")
            .filter_map(|n| {
                n.attribute((REL_NS, "id"))
                    .or_else(|| n.attribute("r:id"))
                    .or_else(|| n.attribute("id"))
            })
            .map(|s| s.to_string())
            .collect();

        let rel_map: HashMap<_, _> = rels.into_iter().map(|r| (r.id.clone(), r)).collect();

        let mut pivot_table_rels = Vec::new();
        if !pivot_table_rids.is_empty() {
            let mut missing_rel = false;
            for rid in &pivot_table_rids {
                match rel_map.get(rid) {
                    Some(rel) if rel.type_ == PIVOT_TABLE_REL_TYPE => {
                        pivot_table_rels.push(RelationshipStub {
                            rel_id: rid.clone(),
                            target: rel.target.clone(),
                        });
                    }
                    _ => {
                        missing_rel = true;
                        break;
                    }
                }
            }

            if missing_rel {
                continue;
            }
        }

        sheet_pivot_tables.insert(
            sheet.name.clone(),
            PreservedSheetPivotTables {
                sheet_index,
                sheet_id: sheet.sheet_id,
                pivot_tables_xml,
                pivot_table_rels,
            },
        );
    }

    Ok(PreservedPivotParts {
        content_types_xml,
        parts,
        workbook_sheets,
        workbook_pivot_caches,
        workbook_pivot_cache_rels,
        workbook_slicer_caches,
        workbook_slicer_cache_rels,
        workbook_timeline_caches,
        workbook_timeline_cache_rels,
        sheet_pivot_tables,
    })
}

impl PreservedPivotParts {
    pub fn is_empty(&self) -> bool {
        self.parts.is_empty()
            && self.workbook_pivot_caches.is_none()
            && self.workbook_slicer_caches.is_none()
            && self.workbook_timeline_caches.is_none()
            && self.sheet_pivot_tables.is_empty()
            && self.workbook_pivot_cache_rels.is_empty()
            && self.workbook_slicer_cache_rels.is_empty()
            && self.workbook_timeline_cache_rels.is_empty()
    }
}

impl XlsxPackage {
    /// Extract pivot-related parts (pivot tables + pivot caches + slicers/timelines)
    /// so they can be re-applied to another package later.
    pub fn preserve_pivot_parts(&self) -> Result<PreservedPivotParts, ChartExtractionError> {
        let content_types_xml = self
            .part("[Content_Types].xml")
            .ok_or_else(|| ChartExtractionError::MissingPart("[Content_Types].xml".to_string()))?
            .to_vec();

        let mut parts = BTreeMap::new();
        for (name, bytes) in self.parts() {
            if name.starts_with("xl/pivotTables/")
                || name.starts_with("xl/pivotCache/")
                || name.starts_with("xl/slicers/")
                || name.starts_with("xl/slicerCaches/")
                || name.starts_with("xl/timelines/")
                || name.starts_with("xl/timelineCaches/")
            {
                parts.insert(name.to_string(), bytes.to_vec());
            }
        }

        let workbook_part = "xl/workbook.xml";
        let workbook_xml = self
            .part(workbook_part)
            .ok_or_else(|| ChartExtractionError::MissingPart(workbook_part.to_string()))?;
        let workbook_xml = std::str::from_utf8(workbook_xml)
            .map_err(|e| ChartExtractionError::XmlNonUtf8(workbook_part.to_string(), e))?;
        let workbook_doc = Document::parse(workbook_xml)
            .map_err(|e| ChartExtractionError::XmlParse(workbook_part.to_string(), e))?;

        let pivot_caches_node = workbook_doc
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "pivotCaches");
        let slicer_caches_node = workbook_doc
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "slicerCaches");
        let timeline_caches_node = workbook_doc
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "timelineCaches");

        let mut workbook_pivot_cache_rids: HashSet<String> = HashSet::new();
        if let Some(node) = pivot_caches_node {
            for rid in node
                .descendants()
                .filter(|n| n.is_element() && n.tag_name().name() == "pivotCache")
                .filter_map(|n| {
                    n.attribute((REL_NS, "id"))
                        .or_else(|| n.attribute("r:id"))
                        .or_else(|| n.attribute("id"))
                })
            {
                workbook_pivot_cache_rids.insert(rid.to_string());
            }
        }

        let mut workbook_slicer_cache_rids: HashSet<String> = HashSet::new();
        if let Some(node) = slicer_caches_node {
            for rid in node
                .descendants()
                .filter(|n| n.is_element() && n.tag_name().name() == "slicerCache")
                .filter_map(|n| {
                    n.attribute((REL_NS, "id"))
                        .or_else(|| n.attribute("r:id"))
                        .or_else(|| n.attribute("id"))
                })
            {
                workbook_slicer_cache_rids.insert(rid.to_string());
            }
        }

        let mut workbook_timeline_cache_rids: HashSet<String> = HashSet::new();
        if let Some(node) = timeline_caches_node {
            for rid in node
                .descendants()
                .filter(|n| n.is_element() && n.tag_name().name() == "timelineCache")
                .filter_map(|n| {
                    n.attribute((REL_NS, "id"))
                        .or_else(|| n.attribute("r:id"))
                        .or_else(|| n.attribute("id"))
                })
            {
                workbook_timeline_cache_rids.insert(rid.to_string());
            }
        }

        let workbook_rels_part = "xl/_rels/workbook.xml.rels";
        let rel_map: HashMap<String, crate::relationships::Relationship> = match self
            .part(workbook_rels_part)
        {
            Some(workbook_rels_xml) => parse_relationships(workbook_rels_xml, workbook_rels_part)?
                .into_iter()
                .map(|r| (r.id.clone(), r))
                .collect(),
            None => HashMap::new(),
        };

        // Only preserve the <pivotCaches> subtree when we can also preserve every referenced
        // pivotCacheDefinition relationship. Otherwise re-applying would introduce broken r:id
        // references in workbook.xml.
        let (workbook_pivot_caches, workbook_pivot_cache_rels) = match pivot_caches_node {
            Some(node) if !workbook_pivot_cache_rids.is_empty() => {
                let mut rels = Vec::new();
                let mut missing_rel = false;
                for rid in &workbook_pivot_cache_rids {
                    match rel_map.get(rid) {
                        Some(rel) if rel.type_ == PIVOT_CACHE_DEF_REL_TYPE => {
                            rels.push(RelationshipStub {
                                rel_id: rid.clone(),
                                target: rel.target.clone(),
                            });
                        }
                        _ => {
                            missing_rel = true;
                            break;
                        }
                    }
                }

                if missing_rel {
                    (None, Vec::new())
                } else {
                    (Some(workbook_xml.as_bytes()[node.range()].to_vec()), rels)
                }
            }
            _ => (None, Vec::new()),
        };

        let (workbook_slicer_caches, workbook_slicer_cache_rels) = match slicer_caches_node {
            Some(node) if !workbook_slicer_cache_rids.is_empty() => {
                let mut rels = Vec::new();
                let mut missing_rel = false;
                for rid in &workbook_slicer_cache_rids {
                    match rel_map.get(rid) {
                        Some(rel) if rel.type_ == SLICER_CACHE_REL_TYPE => {
                            rels.push(RelationshipStub {
                                rel_id: rid.clone(),
                                target: rel.target.clone(),
                            });
                        }
                        _ => {
                            missing_rel = true;
                            break;
                        }
                    }
                }

                if missing_rel {
                    (None, Vec::new())
                } else {
                    (Some(workbook_xml.as_bytes()[node.range()].to_vec()), rels)
                }
            }
            _ => (None, Vec::new()),
        };

        let (workbook_timeline_caches, workbook_timeline_cache_rels) = match timeline_caches_node {
            Some(node) if !workbook_timeline_cache_rids.is_empty() => {
                let mut rels = Vec::new();
                let mut missing_rel = false;
                for rid in &workbook_timeline_cache_rids {
                    match rel_map.get(rid) {
                        Some(rel) if rel.type_ == TIMELINE_CACHE_DEF_REL_TYPE => {
                            rels.push(RelationshipStub {
                                rel_id: rid.clone(),
                                target: rel.target.clone(),
                            });
                        }
                        _ => {
                            missing_rel = true;
                            break;
                        }
                    }
                }

                if missing_rel {
                    (None, Vec::new())
                } else {
                    (Some(workbook_xml.as_bytes()[node.range()].to_vec()), rels)
                }
            }
            _ => (None, Vec::new()),
        };

        let sheets = workbook_sheet_parts(self)?;
        let workbook_sheets = sheets
            .iter()
            .filter(|sheet| sheet.part_name.starts_with("xl/worksheets/"))
            .enumerate()
            .map(|(index, sheet)| PreservedWorkbookSheet {
                name: sheet.name.clone(),
                index,
            })
            .collect::<Vec<_>>();
        let mut sheet_pivot_tables: BTreeMap<String, PreservedSheetPivotTables> = BTreeMap::new();

        let mut worksheet_index = 0usize;
        for sheet in &sheets {
            // Pivot tables can only live on worksheet parts. Skip chart sheets and other sheet
            // types so the stored `sheet_index` matches the worksheet-only order used for rename
            // mapping.
            if !sheet.part_name.starts_with("xl/worksheets/") {
                continue;
            }
            let sheet_index = worksheet_index;
            worksheet_index += 1;

            let Some(sheet_xml) = self.part(&sheet.part_name) else {
                continue;
            };
            let sheet_xml_str = std::str::from_utf8(sheet_xml)
                .map_err(|e| ChartExtractionError::XmlNonUtf8(sheet.part_name.clone(), e))?;
            let sheet_doc = Document::parse(sheet_xml_str)
                .map_err(|e| ChartExtractionError::XmlParse(sheet.part_name.clone(), e))?;

            let pivot_tables_node = sheet_doc
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "pivotTables");
            let Some(pivot_tables_node) = pivot_tables_node else {
                continue;
            };

            let pivot_tables_xml = sheet_xml_str.as_bytes()[pivot_tables_node.range()].to_vec();

            let pivot_table_rids: Vec<String> = pivot_tables_node
                .descendants()
                .filter(|n| n.is_element() && n.tag_name().name() == "pivotTable")
                .filter_map(|n| {
                    n.attribute((REL_NS, "id"))
                        .or_else(|| n.attribute("r:id"))
                        .or_else(|| n.attribute("id"))
                })
                .map(|s| s.to_string())
                .collect();

            let mut pivot_table_rels = Vec::new();
            if !pivot_table_rids.is_empty() {
                let sheet_rels_part = rels_for_part(&sheet.part_name);
                let Some(sheet_rels_xml) = self.part(&sheet_rels_part) else {
                    continue;
                };
                let rels = parse_relationships(sheet_rels_xml, &sheet_rels_part)?;
                let rel_map: HashMap<_, _> = rels.into_iter().map(|r| (r.id.clone(), r)).collect();

                let mut missing_rel = false;
                for rid in &pivot_table_rids {
                    match rel_map.get(rid) {
                        Some(rel) if rel.type_ == PIVOT_TABLE_REL_TYPE => {
                            pivot_table_rels.push(RelationshipStub {
                                rel_id: rid.clone(),
                                target: rel.target.clone(),
                            });
                        }
                        _ => {
                            missing_rel = true;
                            break;
                        }
                    }
                }

                if missing_rel {
                    continue;
                }
            }

            sheet_pivot_tables.insert(
                sheet.name.clone(),
                PreservedSheetPivotTables {
                    sheet_index,
                    sheet_id: sheet.sheet_id,
                    pivot_tables_xml,
                    pivot_table_rels,
                },
            );
        }

        Ok(PreservedPivotParts {
            content_types_xml,
            parts,
            workbook_sheets,
            workbook_pivot_caches,
            workbook_pivot_cache_rels,
            workbook_slicer_caches,
            workbook_slicer_cache_rels,
            workbook_timeline_caches,
            workbook_timeline_cache_rels,
            sheet_pivot_tables,
        })
    }

    /// Apply previously captured pivot parts to this package.
    ///
    /// This function is intentionally conservative and only appends missing XML
    /// and relationships; it does not remove any existing destination data.
    pub fn apply_preserved_pivot_parts(
        &mut self,
        preserved: &PreservedPivotParts,
    ) -> Result<(), ChartExtractionError> {
        if preserved.is_empty() {
            return Ok(());
        }

        let sheets = workbook_sheet_parts(self)?;

        // Build a worksheet-only view of the workbook sheet list.
        //
        // Pivot cache worksheet sources can only point at worksheets (never chartsheets), and
        // regeneration-based saves may re-attach chartsheets at a different position in
        // `xl/workbook.xml`. Using worksheet-only indices prevents chart sheets from breaking the
        // preserved index mapping when worksheets are renamed.
        let worksheet_sheets = sheets
            .iter()
            .filter(|sheet| sheet.part_name.starts_with("xl/worksheets/"))
            .enumerate()
            .map(|(index, sheet)| crate::preserve::sheet_match::WorkbookSheetPart {
                name: sheet.name.clone(),
                index,
                sheet_id: sheet.sheet_id,
                part_name: sheet.part_name.clone(),
            })
            .collect::<Vec<_>>();

        let mut sheet_name_map: HashMap<String, String> = HashMap::new();
        for preserved_sheet in &preserved.workbook_sheets {
            let Some(matched) = match_sheet_by_name_or_index(
                &worksheet_sheets,
                &preserved_sheet.name,
                preserved_sheet.index,
            ) else {
                continue;
            };
            if matched.name != preserved_sheet.name {
                sheet_name_map.insert(preserved_sheet.name.clone(), matched.name.clone());
            }
        }

        for (name, bytes) in &preserved.parts {
            let bytes = if is_pivot_cache_definition_part(name) {
                rewrite_pivot_cache_definition_worksheet_source_sheets(
                    bytes,
                    name,
                    &sheet_name_map,
                )?
            } else {
                bytes.clone()
            };
            self.set_part(name.clone(), bytes);
        }

        self.merge_content_types(&preserved.content_types_xml, preserved.parts.keys())?;

        let workbook_part = "xl/workbook.xml";
        let workbook_rels_part = "xl/_rels/workbook.xml.rels";

        let workbook_pivot_cache_rid_map = if !preserved.workbook_pivot_cache_rels.is_empty() {
            let (updated_workbook_rels, rid_map) = ensure_rels_has_relationships(
                self.part(workbook_rels_part),
                workbook_rels_part,
                workbook_part,
                PIVOT_CACHE_DEF_REL_TYPE,
                &preserved.workbook_pivot_cache_rels,
            )?;
            self.set_part(workbook_rels_part, updated_workbook_rels);
            rid_map
        } else {
            HashMap::new()
        };

        let workbook_slicer_cache_rid_map = if !preserved.workbook_slicer_cache_rels.is_empty() {
            let (updated_workbook_rels, rid_map) = ensure_rels_has_relationships(
                self.part(workbook_rels_part),
                workbook_rels_part,
                workbook_part,
                SLICER_CACHE_REL_TYPE,
                &preserved.workbook_slicer_cache_rels,
            )?;
            self.set_part(workbook_rels_part, updated_workbook_rels);
            rid_map
        } else {
            HashMap::new()
        };

        let workbook_timeline_cache_rid_map = if !preserved.workbook_timeline_cache_rels.is_empty() {
            let (updated_workbook_rels, rid_map) = ensure_rels_has_relationships(
                self.part(workbook_rels_part),
                workbook_rels_part,
                workbook_part,
                TIMELINE_CACHE_DEF_REL_TYPE,
                &preserved.workbook_timeline_cache_rels,
            )?;
            self.set_part(workbook_rels_part, updated_workbook_rels);
            rid_map
        } else {
            HashMap::new()
        };

        if let Some(pivot_caches) = preserved.workbook_pivot_caches.as_deref() {
            if !preserved.workbook_pivot_cache_rels.is_empty() {
                let rewritten =
                    rewrite_relationship_ids(pivot_caches, "pivotCaches", &workbook_pivot_cache_rid_map)?;
                let workbook_xml = self
                    .part(workbook_part)
                    .ok_or_else(|| ChartExtractionError::MissingPart(workbook_part.to_string()))?;
                let updated =
                    ensure_workbook_xml_has_pivot_caches(workbook_xml, workbook_part, &rewritten)?;
                self.set_part(workbook_part, updated);
            }
        }
        if let Some(slicer_caches) = preserved.workbook_slicer_caches.as_deref() {
            if !preserved.workbook_slicer_cache_rels.is_empty() {
                let rewritten = rewrite_relationship_ids(
                    slicer_caches,
                    "slicerCaches",
                    &workbook_slicer_cache_rid_map,
                )?;
                let workbook_xml = self
                    .part(workbook_part)
                    .ok_or_else(|| ChartExtractionError::MissingPart(workbook_part.to_string()))?;
                let updated =
                    ensure_workbook_xml_has_slicer_caches(workbook_xml, workbook_part, &rewritten)?;
                self.set_part(workbook_part, updated);
            }
        }

        if let Some(timeline_caches) = preserved.workbook_timeline_caches.as_deref() {
            if !preserved.workbook_timeline_cache_rels.is_empty() {
                let rewritten = rewrite_relationship_ids(
                    timeline_caches,
                    "timelineCaches",
                    &workbook_timeline_cache_rid_map,
                )?;
                let workbook_xml = self
                    .part(workbook_part)
                    .ok_or_else(|| ChartExtractionError::MissingPart(workbook_part.to_string()))?;
                let updated = ensure_workbook_xml_has_timeline_caches(
                    workbook_xml,
                    workbook_part,
                    &rewritten,
                )?;
                self.set_part(workbook_part, updated);
            }
        }

        for (sheet_name, preserved_sheet) in &preserved.sheet_pivot_tables {
            let Some(sheet) =
                match_sheet_by_name_or_index(&worksheet_sheets, sheet_name, preserved_sheet.sheet_index)
            else {
                continue;
            };

            let rid_map = if !preserved_sheet.pivot_table_rels.is_empty() {
                let sheet_rels_part = rels_for_part(&sheet.part_name);
                let (updated_sheet_rels, rid_map) = ensure_rels_has_relationships(
                    self.part(&sheet_rels_part),
                    &sheet_rels_part,
                    &sheet.part_name,
                    PIVOT_TABLE_REL_TYPE,
                    &preserved_sheet.pivot_table_rels,
                )?;
                self.set_part(sheet_rels_part, updated_sheet_rels);
                rid_map
            } else {
                HashMap::new()
            };

            let rewritten = rewrite_relationship_ids(
                &preserved_sheet.pivot_tables_xml,
                "pivotTables",
                &rid_map,
            )?;
            let Some(sheet_xml) = self.part(&sheet.part_name) else {
                continue;
            };
            let updated_sheet_xml =
                ensure_sheet_xml_has_pivot_tables(sheet_xml, &sheet.part_name, &rewritten)?;
            self.set_part(sheet.part_name.clone(), updated_sheet_xml);
        }

        Ok(())
    }
}

fn is_pivot_cache_definition_part(part_name: &str) -> bool {
    part_name.starts_with("xl/pivotCache/")
        && !part_name.contains("/_rels/")
        && part_name.ends_with(".xml")
        && part_name.contains("pivotCacheDefinition")
}

fn lookup_sheet_name<'a>(
    name: &str,
    mapping: &'a HashMap<String, String>,
) -> Option<&'a str> {
    if let Some(v) = mapping.get(name) {
        return Some(v.as_str());
    }
    mapping
        .iter()
        .find(|(k, _)| sheet_name_eq_case_insensitive(k, name))
        .map(|(_, v)| v.as_str())
}

fn rewrite_sheet_name_in_ref(ref_value: &str, mapping: &HashMap<String, String>) -> Option<String> {
    let (sheet_token, range) = ref_value.rsplit_once('!')?;
    let (sheet_name, was_quoted) =
        if let Some(inner) = formula_model::unquote_excel_single_quoted_identifier_lenient(sheet_token)
        {
            (inner.into_owned(), true)
        } else {
            (sheet_token.to_string(), false)
        };

    let new_sheet = lookup_sheet_name(&sheet_name, mapping)?;
    let quote = was_quoted || formula_model::sheet_name_needs_quotes_a1(new_sheet);
    let new_token = if quote {
        let mut out = String::with_capacity(new_sheet.len().saturating_add(2));
        formula_model::push_excel_single_quoted_identifier(&mut out, new_sheet);
        out
    } else {
        new_sheet.to_string()
    };
    Some(format!("{new_token}!{range}"))
}

fn rewrite_pivot_cache_definition_worksheet_source_sheets(
    xml_bytes: &[u8],
    part_name: &str,
    sheet_name_map: &HashMap<String, String>,
) -> Result<Vec<u8>, ChartExtractionError> {
    if sheet_name_map.is_empty() {
        return Ok(xml_bytes.to_vec());
    }

    let xml = std::str::from_utf8(xml_bytes)
        .map_err(|e| ChartExtractionError::XmlNonUtf8(part_name.to_string(), e))?;

    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Vec::with_capacity(xml_bytes.len()));
    let mut buf = Vec::new();
    let mut rewritten = false;

    loop {
        let event = reader.read_event_into(&mut buf).map_err(|e| {
            ChartExtractionError::XmlStructure(format!("{part_name}: xml parse error: {e}"))
        })?;

        match event {
            Event::Eof => break,
            Event::Start(ref e) => {
                if e.local_name().as_ref().eq_ignore_ascii_case(b"worksheetSource") {
                    let qname = e.name();
                    let name = std::str::from_utf8(qname.as_ref()).map_err(|_| {
                        ChartExtractionError::XmlStructure(format!(
                            "{part_name}: non-utf8 element name"
                        ))
                    })?;
                    let mut out = BytesStart::new(name);
                    for attr in e.attributes().with_checks(false) {
                        let attr = attr.map_err(|e| {
                            ChartExtractionError::XmlStructure(format!(
                                "{part_name}: xml attribute error: {e}"
                            ))
                        })?;
                        let key = attr.key.as_ref();
                        let local = crate::openxml::local_name(key);
                        if local.eq_ignore_ascii_case(b"sheet") {
                            let value = attr.unescape_value().map_err(|e| {
                                ChartExtractionError::XmlStructure(format!(
                                    "{part_name}: xml attribute error: {e}"
                                ))
                            })?;
                            if let Some(new_sheet) = lookup_sheet_name(value.as_ref(), sheet_name_map)
                            {
                                rewritten = true;
                                let escaped = xml_escape(new_sheet);
                                out.push_attribute((key, escaped.as_bytes()));
                            } else {
                                out.push_attribute((key, attr.value.as_ref()));
                            }
                        } else if local.eq_ignore_ascii_case(b"ref") {
                            let value = attr.unescape_value().map_err(|e| {
                                ChartExtractionError::XmlStructure(format!(
                                    "{part_name}: xml attribute error: {e}"
                                ))
                            })?;
                            if let Some(new_ref) =
                                rewrite_sheet_name_in_ref(value.as_ref(), sheet_name_map)
                            {
                                rewritten = true;
                                let escaped = xml_escape(&new_ref);
                                out.push_attribute((key, escaped.as_bytes()));
                            } else {
                                out.push_attribute((key, attr.value.as_ref()));
                            }
                        } else {
                            out.push_attribute((key, attr.value.as_ref()));
                        }
                    }

                    writer.write_event(Event::Start(out)).map_err(|e| {
                        ChartExtractionError::XmlStructure(format!(
                            "{part_name}: xml write error: {e}"
                        ))
                    })?;
                } else {
                    writer.write_event(Event::Start(e.to_owned())).map_err(|e| {
                        ChartExtractionError::XmlStructure(format!(
                            "{part_name}: xml write error: {e}"
                        ))
                    })?;
                }
            }
            Event::Empty(ref e) => {
                if e.local_name().as_ref().eq_ignore_ascii_case(b"worksheetSource") {
                    let qname = e.name();
                    let name = std::str::from_utf8(qname.as_ref()).map_err(|_| {
                        ChartExtractionError::XmlStructure(format!(
                            "{part_name}: non-utf8 element name"
                        ))
                    })?;
                    let mut out = BytesStart::new(name);
                    for attr in e.attributes().with_checks(false) {
                        let attr = attr.map_err(|e| {
                            ChartExtractionError::XmlStructure(format!(
                                "{part_name}: xml attribute error: {e}"
                            ))
                        })?;
                        let key = attr.key.as_ref();
                        let local = crate::openxml::local_name(key);
                        if local.eq_ignore_ascii_case(b"sheet") {
                            let value = attr.unescape_value().map_err(|e| {
                                ChartExtractionError::XmlStructure(format!(
                                    "{part_name}: xml attribute error: {e}"
                                ))
                            })?;
                            if let Some(new_sheet) = lookup_sheet_name(value.as_ref(), sheet_name_map)
                            {
                                rewritten = true;
                                let escaped = xml_escape(new_sheet);
                                out.push_attribute((key, escaped.as_bytes()));
                            } else {
                                out.push_attribute((key, attr.value.as_ref()));
                            }
                        } else if local.eq_ignore_ascii_case(b"ref") {
                            let value = attr.unescape_value().map_err(|e| {
                                ChartExtractionError::XmlStructure(format!(
                                    "{part_name}: xml attribute error: {e}"
                                ))
                            })?;
                            if let Some(new_ref) =
                                rewrite_sheet_name_in_ref(value.as_ref(), sheet_name_map)
                            {
                                rewritten = true;
                                let escaped = xml_escape(&new_ref);
                                out.push_attribute((key, escaped.as_bytes()));
                            } else {
                                out.push_attribute((key, attr.value.as_ref()));
                            }
                        } else {
                            out.push_attribute((key, attr.value.as_ref()));
                        }
                    }

                    writer.write_event(Event::Empty(out)).map_err(|e| {
                        ChartExtractionError::XmlStructure(format!(
                            "{part_name}: xml write error: {e}"
                        ))
                    })?;
                } else {
                    writer.write_event(Event::Empty(e.to_owned())).map_err(|e| {
                        ChartExtractionError::XmlStructure(format!(
                            "{part_name}: xml write error: {e}"
                        ))
                    })?;
                }
            }
            _ => {
                writer.write_event(event.to_owned()).map_err(|e| {
                    ChartExtractionError::XmlStructure(format!("{part_name}: xml write error: {e}"))
                })?;
            }
        }

        buf.clear();
    }

    if !rewritten {
        return Ok(xml_bytes.to_vec());
    }

    Ok(writer.into_inner())
}

fn read_zip_part_optional<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    name: &str,
    max_part_bytes: u64,
    budget: &mut ZipInflateBudget,
) -> Result<Option<Vec<u8>>, ChartExtractionError> {
    crate::zip_util::read_zip_part_optional_with_budget(
        archive,
        name,
        max_part_bytes,
        budget,
    )
    .map_err(|e| ChartExtractionError::XmlStructure(e.to_string()))
}

fn read_zip_part_required<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    name: &str,
    max_part_bytes: u64,
    budget: &mut ZipInflateBudget,
) -> Result<Vec<u8>, ChartExtractionError> {
    read_zip_part_optional(archive, name, max_part_bytes, budget)?
        .ok_or_else(|| ChartExtractionError::MissingPart(name.to_string()))
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct PivotCacheEntry {
    cache_id: Option<u32>,
    raw_xml: String,
}

/// Merge the preserved `<pivotCaches>` block into an existing `xl/workbook.xml` document.
///
/// Excel enforces a strict ordering for direct children of `<workbook>`. In particular,
/// `<pivotCaches>` must appear **after** `<customWorkbookViews>` (if present) and **before**
/// elements like `<fileRecoveryPr>`/`<extLst>`. Inserting the section naively before
/// `</workbook>` can violate that ordering and cause Excel to repair/corrupt the file.
///
/// This helper:
/// - Inserts `<pivotCaches>` at a schema-safe location when missing.
/// - When `<pivotCaches>` already exists, merges only the missing `<pivotCache cacheId="…">`
///   entries (does not delete or reorder existing entries).
/// - Ensures the `<workbook>` element declares the relationships namespace (`{REL_NS}`) for any
///   prefixes used in inserted `*:id="..."` attributes (e.g. `r:id` or `rel:id`).
pub fn apply_preserved_pivot_caches_to_workbook_xml(
    workbook_xml: &str,
    preserved_pivot_caches_xml: &str,
) -> Result<String, ChartExtractionError> {
    apply_preserved_pivot_caches_to_workbook_xml_with_part(
        workbook_xml,
        "xl/workbook.xml",
        preserved_pivot_caches_xml,
    )
}

fn apply_preserved_pivot_caches_to_workbook_xml_with_part(
    workbook_xml: &str,
    part_name: &str,
    preserved_pivot_caches_xml: &str,
) -> Result<String, ChartExtractionError> {
    if preserved_pivot_caches_xml.trim().is_empty() {
        return Ok(workbook_xml.to_string());
    }

    let workbook_xml = expand_self_closing_workbook_root_if_needed(workbook_xml, part_name)?;
    let workbook_xml = workbook_xml.as_ref();

    let doc = Document::parse(workbook_xml)
        .map_err(|e| ChartExtractionError::XmlParse(part_name.to_string(), e))?;
    let workbook = doc.root_element();

    let mut updated = if let Some(pivot_caches) = workbook
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "pivotCaches")
    {
        merge_pivot_caches(workbook_xml, pivot_caches, preserved_pivot_caches_xml)?
    } else {
        insert_pivot_caches(workbook_xml, &workbook, preserved_pivot_caches_xml)?
    };

    // If we didn't insert anything, avoid mutating the workbook.
    if updated == workbook_xml {
        return Ok(updated);
    }

    // Ensure the `<workbook>` element declares the relationships namespace for any prefixes used
    // in the inserted fragment (e.g. `r:id` or `rel:id`).
    for prefix in detect_attr_prefixes(preserved_pivot_caches_xml, "id") {
        if prefix == "xmlns" {
            continue;
        }
        updated = ensure_workbook_has_namespace_prefix(&updated, part_name, &prefix, REL_NS)?;
    }

    Ok(updated)
}

fn merge_pivot_caches(
    workbook_xml: &str,
    existing_pivot_caches: roxmltree::Node<'_, '_>,
    preserved_pivot_caches_xml: &str,
) -> Result<String, ChartExtractionError> {
    let pivot_caches_range = existing_pivot_caches.range();
    let existing_section = &workbook_xml[pivot_caches_range.clone()];
    let pivot_caches_prefix = element_prefix_at(workbook_xml, pivot_caches_range.start);
    let pivot_caches_tag = crate::xml::prefixed_tag(pivot_caches_prefix, "pivotCaches");
    let close_tag = format!("</{pivot_caches_tag}>");

    let existing_cache_ids: HashSet<u32> = existing_pivot_caches
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "pivotCache")
        .filter_map(|n| n.attribute("cacheId").and_then(|v| v.parse::<u32>().ok()))
        .collect();
    let existing_pivot_cache_count = existing_pivot_caches
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "pivotCache")
        .count();

    let preserved_entries = parse_pivot_cache_entries(preserved_pivot_caches_xml)?;

    let mut to_insert = Vec::new();
    for entry in preserved_entries {
        match entry.cache_id {
            Some(id) if existing_cache_ids.contains(&id) => continue,
            Some(_) | None => to_insert.push(entry),
        }
    }

    if to_insert.is_empty() {
        return Ok(workbook_xml.to_string());
    }

    let mut inserted_xml = String::new();
    for entry in &to_insert {
        // Make sure inserted entries follow the existing SpreadsheetML prefix style.
        let rewritten = rewrite_spreadsheetml_prefix_in_fragment(
            &entry.raw_xml,
            pivot_caches_prefix,
            &["pivotCache"],
            "pivotCaches",
        )?;
        inserted_xml.push_str(&rewritten);
    }

    let mut new_section = if is_self_closing_element(existing_section) {
        let (start, trailing_ws) = split_trailing_whitespace(existing_section);
        let start = start.trim_end();
        let start = start.strip_suffix("/>").unwrap_or(start);
        let start = start.trim_end();
        let mut section = String::new();
        section.push_str(start);
        section.push('>');
        section.push_str(&inserted_xml);
        section.push_str(&close_tag);
        section.push_str(trailing_ws);
        section
    } else {
        let close_tag_pos = existing_section.rfind(&close_tag).ok_or_else(|| {
            ChartExtractionError::XmlStructure(format!(
                "workbook.xml: missing {close_tag} in <pivotCaches>"
            ))
        })?;
        let mut section =
            String::with_capacity(existing_section.len().saturating_add(inserted_xml.len()));
        section.push_str(&existing_section[..close_tag_pos]);
        section.push_str(&inserted_xml);
        section.push_str(&existing_section[close_tag_pos..]);
        section
    };

    // Keep `count="..."` in sync when present.
    let new_count = existing_pivot_cache_count + to_insert.len();
    new_section = update_pivot_caches_count_attr(&new_section, new_count);

    let mut out = String::with_capacity(workbook_xml.len() + inserted_xml.len());
    out.push_str(&workbook_xml[..pivot_caches_range.start]);
    out.push_str(&new_section);
    out.push_str(&workbook_xml[pivot_caches_range.end..]);
    Ok(out)
}

fn insert_pivot_caches(
    workbook_xml: &str,
    workbook_node: &roxmltree::Node<'_, '_>,
    preserved_pivot_caches_xml: &str,
) -> Result<String, ChartExtractionError> {
    let workbook_range = workbook_node.range();
    let root_start = workbook_range.start;
    let workbook_prefix = element_prefix_at(workbook_xml, root_start);
    let preserved_pivot_caches_xml = rewrite_spreadsheetml_prefix_in_fragment(
        preserved_pivot_caches_xml,
        workbook_prefix,
        &["pivotCaches", "pivotCache"],
        "pivotCaches",
    )?;

    // Some producers emit `xl/workbook.xml` with a self-closing root element:
    // `<workbook .../>` or `<x:workbook .../>`.
    //
    // When that happens we must expand the root in-place before we can insert
    // children like `<pivotCaches>`. Otherwise, inserting at "before </workbook>"
    // will compute an index inside the `<workbook .../>` start tag, producing
    // invalid XML.
    let workbook_root = &workbook_xml[workbook_range.clone()];
    if is_self_closing_element(workbook_root) {
        let workbook_tag = crate::xml::prefixed_tag(workbook_prefix, "workbook");
        let close_tag = format!("</{workbook_tag}>");

        let (start, trailing_ws) = split_trailing_whitespace(workbook_root);
        let start = start.trim_end();
        let start = start.strip_suffix("/>").unwrap_or(start);
        let start = start.trim_end();

        let mut expanded_root = String::with_capacity(
            start.len() + 1 + preserved_pivot_caches_xml.len() + close_tag.len() + trailing_ws.len(),
        );
        expanded_root.push_str(start);
        expanded_root.push('>');
        expanded_root.push_str(&preserved_pivot_caches_xml);
        expanded_root.push_str(&close_tag);
        expanded_root.push_str(trailing_ws);

        let mut out = String::with_capacity(
            workbook_xml.len() + preserved_pivot_caches_xml.len() + close_tag.len() + 1,
        );
        out.push_str(&workbook_xml[..workbook_range.start]);
        out.push_str(&expanded_root);
        out.push_str(&workbook_xml[workbook_range.end..]);
        return Ok(out);
    }

    let mut insert_idx = None;

    // 1) Prefer inserting after `<customWorkbookViews>` when present.
    for child in workbook_node.children().filter(|n| n.is_element()) {
        if child.tag_name().name() == "customWorkbookViews" {
            insert_idx = Some(child.range().end);
            break;
        }
    }

    // 2) Otherwise insert before the first "after" element.
    if insert_idx.is_none() {
        let after = [
            "smartTagPr",
            "smartTagTypes",
            "webPublishing",
            "fileRecoveryPr",
            "webPublishObjects",
            "extLst",
        ];
        for child in workbook_node.children().filter(|n| n.is_element()) {
            if after.contains(&child.tag_name().name()) {
                insert_idx = Some(child.range().start);
                break;
            }
        }
    }

    // 3) Fallback: insert before closing `</workbook>`.
    let insert_idx = match insert_idx {
        Some(idx) => idx,
        None => {
            let close_tag_len = crate::xml::prefixed_tag(workbook_prefix, "workbook").len() + 3;
            workbook_node
                .range()
                .end
                .checked_sub(close_tag_len)
                .ok_or_else(|| {
                    ChartExtractionError::XmlStructure(
                        "workbook.xml: invalid </workbook> close tag".to_string(),
                    )
                })?
        }
    };

    let mut out = String::with_capacity(workbook_xml.len() + preserved_pivot_caches_xml.len());
    out.push_str(&workbook_xml[..insert_idx]);
    out.push_str(&preserved_pivot_caches_xml);
    out.push_str(&workbook_xml[insert_idx..]);
    Ok(out)
}

fn parse_pivot_cache_entries(xml: &str) -> Result<Vec<PivotCacheEntry>, ChartExtractionError> {
    // The preserved `<pivotCaches>` section usually relies on `xmlns:r` declared on the
    // surrounding `<workbook>` element. When parsing it in isolation we need to provide
    // a namespace binding for the `r:` prefix so `r:id="..."` attributes remain valid.
    let spreadsheet_prefix = detect_prefix_in_fragment(xml, "pivotCaches")
        .or_else(|| detect_prefix_in_fragment(xml, "pivotCache"));
    let prefix_decl = spreadsheet_prefix
        .as_deref()
        .map(|p| format!(" xmlns:{p}=\"{SPREADSHEETML_NS}\""))
        .unwrap_or_default();

    // Workbook/worksheet fragments frequently inherit the relationships namespace declaration from
    // an ancestor element. Declare any prefixes we see on `*:id="..."` attributes so the fragment
    // is namespace-well-formed when parsed in isolation.
    let mut rel_decls = String::new();
    for prefix in detect_attr_prefixes(xml, "id") {
        if prefix == "xmlns" {
            continue;
        }
        if spreadsheet_prefix.as_deref() == Some(prefix.as_str()) {
            continue;
        }
        rel_decls.push_str(&format!(" xmlns:{prefix}=\"{REL_NS}\""));
    }

    let wrapped = format!(r#"<root{rel_decls}{prefix_decl}>{xml}</root>"#);
    let doc = Document::parse(&wrapped)
        .map_err(|e| ChartExtractionError::XmlParse("pivotCaches".to_string(), e))?;

    let pivot_caches = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "pivotCaches")
        .ok_or_else(|| {
            ChartExtractionError::XmlStructure("pivotCaches: missing <pivotCaches>".to_string())
        })?;

    let mut entries = Vec::new();
    for node in pivot_caches
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "pivotCache")
    {
        let cache_id = node
            .attribute("cacheId")
            .and_then(|v| v.parse::<u32>().ok());
        let raw_xml = wrapped[node.range()].to_string();
        entries.push(PivotCacheEntry { cache_id, raw_xml });
    }

    Ok(entries)
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct CacheRefEntry {
    rel_id: Option<String>,
    raw_xml: String,
}

fn parse_cache_ref_entries(
    xml: &str,
    context: &str,
    container: &str,
    child: &str,
) -> Result<Vec<CacheRefEntry>, ChartExtractionError> {
    // Workbook fragments frequently inherit the SpreadsheetML namespace declaration from an
    // ancestor element (the `<workbook>` root). If the fragment uses a prefix (e.g. `x:`), declare
    // it so `roxmltree` can parse the fragment in isolation.
    let spreadsheet_prefix =
        detect_prefix_in_fragment(xml, container).or_else(|| detect_prefix_in_fragment(xml, child));
    let prefix_decl = spreadsheet_prefix
        .as_deref()
        .map(|p| format!(" xmlns:{p}=\"{SPREADSHEETML_NS}\""))
        .unwrap_or_default();

    // Likewise, declare any relationship namespace prefixes used on `*:id="..."` attributes.
    let mut rel_decls = String::new();
    for prefix in detect_attr_prefixes(xml, "id") {
        if prefix == "xmlns" {
            continue;
        }
        if spreadsheet_prefix.as_deref() == Some(prefix.as_str()) {
            continue;
        }
        rel_decls.push_str(&format!(" xmlns:{prefix}=\"{REL_NS}\""));
    }

    let wrapped = format!(r#"<root{rel_decls}{prefix_decl}>{xml}</root>"#);
    let doc =
        Document::parse(&wrapped).map_err(|e| ChartExtractionError::XmlParse(context.to_string(), e))?;

    let container_node = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == container)
        .ok_or_else(|| {
            ChartExtractionError::XmlStructure(format!("{context}: missing <{container}>"))
        })?;

    let mut entries = Vec::new();
    for node in container_node
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == child)
    {
        let rel_id = node
            .attribute((REL_NS, "id"))
            .or_else(|| node.attribute("r:id"))
            .or_else(|| node.attribute("id"))
            .map(|s| s.to_string());
        let raw_xml = wrapped[node.range()].to_string();
        entries.push(CacheRefEntry { rel_id, raw_xml });
    }

    Ok(entries)
}

fn merge_workbook_cache_refs(
    workbook_xml: &str,
    existing_container: roxmltree::Node<'_, '_>,
    preserved_xml: &str,
    context: &str,
    container: &str,
    child: &str,
) -> Result<String, ChartExtractionError> {
    let range = existing_container.range();
    let existing_section = &workbook_xml[range.clone()];
    let container_prefix = element_prefix_at(workbook_xml, range.start);
    let container_tag = crate::xml::prefixed_tag(container_prefix, container);
    let close_tag = format!("</{container_tag}>");

    let existing_rids: HashSet<String> = existing_container
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == child)
        .filter_map(|n| {
            n.attribute((REL_NS, "id"))
                .or_else(|| n.attribute("r:id"))
                .or_else(|| n.attribute("id"))
        })
        .map(|s| s.to_string())
        .collect();
    let existing_count = existing_container
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == child)
        .count();

    let preserved_entries = parse_cache_ref_entries(preserved_xml, context, container, child)?;
    let mut to_insert: Vec<CacheRefEntry> = Vec::new();
    for entry in preserved_entries {
        match entry.rel_id.as_deref() {
            Some(rid) if existing_rids.contains(rid) => continue,
            Some(_) | None => to_insert.push(entry),
        }
    }

    if to_insert.is_empty() {
        return Ok(workbook_xml.to_string());
    }

    let mut inserted_xml = String::new();
    for entry in &to_insert {
        let rewritten = rewrite_spreadsheetml_prefix_in_fragment(
            &entry.raw_xml,
            container_prefix,
            &[child],
            context,
        )?;
        inserted_xml.push_str(&rewritten);
    }

    let mut new_section = if is_self_closing_element(existing_section) {
        let (start, trailing_ws) = split_trailing_whitespace(existing_section);
        let start = start.trim_end();
        let start = start.strip_suffix("/>").unwrap_or(start);
        let start = start.trim_end();
        let mut section = String::new();
        section.push_str(start);
        section.push('>');
        section.push_str(&inserted_xml);
        section.push_str(&close_tag);
        section.push_str(trailing_ws);
        section
    } else {
        let close_tag_pos = existing_section.rfind(&close_tag).ok_or_else(|| {
            ChartExtractionError::XmlStructure(format!(
                "workbook.xml: missing {close_tag} in <{container}>"
            ))
        })?;
        let mut section =
            String::with_capacity(existing_section.len().saturating_add(inserted_xml.len()));
        section.push_str(&existing_section[..close_tag_pos]);
        section.push_str(&inserted_xml);
        section.push_str(&existing_section[close_tag_pos..]);
        section
    };

    let new_count = existing_count + to_insert.len();
    new_section = update_pivot_caches_count_attr(&new_section, new_count);

    let mut out = String::with_capacity(workbook_xml.len() + inserted_xml.len());
    out.push_str(&workbook_xml[..range.start]);
    out.push_str(&new_section);
    out.push_str(&workbook_xml[range.end..]);
    Ok(out)
}

fn insert_workbook_cache_refs(
    workbook_xml: &str,
    workbook_node: &roxmltree::Node<'_, '_>,
    preserved_xml: &str,
    context: &str,
    container: &str,
    child: &str,
    prefer_after: &[&str],
    before: &[&str],
) -> Result<String, ChartExtractionError> {
    let workbook_range = workbook_node.range();
    let root_start = workbook_range.start;
    let workbook_prefix = element_prefix_at(workbook_xml, root_start);
    let preserved_xml = rewrite_spreadsheetml_prefix_in_fragment(
        preserved_xml,
        workbook_prefix,
        &[container, child],
        context,
    )?;

    let mut insert_idx = None;

    for name in prefer_after {
        for child_node in workbook_node.children().filter(|n| n.is_element()) {
            if child_node.tag_name().name() == *name {
                insert_idx = Some(child_node.range().end);
                break;
            }
        }
        if insert_idx.is_some() {
            break;
        }
    }

    if insert_idx.is_none() {
        for child_node in workbook_node.children().filter(|n| n.is_element()) {
            if before.contains(&child_node.tag_name().name()) {
                insert_idx = Some(child_node.range().start);
                break;
            }
        }
    }

    let insert_idx = match insert_idx {
        Some(idx) => idx,
        None => {
            let close_tag_len = crate::xml::prefixed_tag(workbook_prefix, "workbook").len() + 3;
            workbook_node
                .range()
                .end
                .checked_sub(close_tag_len)
                .ok_or_else(|| {
                    ChartExtractionError::XmlStructure(
                        "workbook.xml: invalid </workbook> close tag".to_string(),
                    )
                })?
        }
    };

    let mut out = String::with_capacity(workbook_xml.len() + preserved_xml.len());
    out.push_str(&workbook_xml[..insert_idx]);
    out.push_str(&preserved_xml);
    out.push_str(&workbook_xml[insert_idx..]);
    Ok(out)
}

fn apply_preserved_workbook_cache_refs_to_workbook_xml_with_part(
    workbook_xml: &str,
    part_name: &str,
    preserved_xml: &str,
    context: &str,
    container: &str,
    child: &str,
    prefer_after: &[&str],
    before: &[&str],
) -> Result<String, ChartExtractionError> {
    if preserved_xml.trim().is_empty() {
        return Ok(workbook_xml.to_string());
    }

    let workbook_xml = expand_self_closing_workbook_root_if_needed(workbook_xml, part_name)?;
    let workbook_xml = workbook_xml.as_ref();

    let doc = Document::parse(workbook_xml)
        .map_err(|e| ChartExtractionError::XmlParse(part_name.to_string(), e))?;
    let workbook = doc.root_element();

    let mut updated = if let Some(existing) = workbook
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == container)
    {
        merge_workbook_cache_refs(workbook_xml, existing, preserved_xml, context, container, child)?
    } else {
        insert_workbook_cache_refs(
            workbook_xml,
            &workbook,
            preserved_xml,
            context,
            container,
            child,
            prefer_after,
            before,
        )?
    };

    if updated == workbook_xml {
        return Ok(updated);
    }

    for prefix in detect_attr_prefixes(preserved_xml, "id") {
        if prefix == "xmlns" {
            continue;
        }
        updated = ensure_workbook_has_namespace_prefix(&updated, part_name, &prefix, REL_NS)?;
    }

    Ok(updated)
}

fn apply_preserved_slicer_caches_to_workbook_xml_with_part(
    workbook_xml: &str,
    part_name: &str,
    preserved_slicer_caches_xml: &str,
) -> Result<String, ChartExtractionError> {
    let prefer_after = ["pivotCaches", "customWorkbookViews"];
    let before = [
        "timelineCaches",
        "smartTagPr",
        "smartTagTypes",
        "webPublishing",
        "fileRecoveryPr",
        "webPublishObjects",
        "extLst",
    ];
    apply_preserved_workbook_cache_refs_to_workbook_xml_with_part(
        workbook_xml,
        part_name,
        preserved_slicer_caches_xml,
        "slicerCaches",
        "slicerCaches",
        "slicerCache",
        &prefer_after,
        &before,
    )
}

fn apply_preserved_timeline_caches_to_workbook_xml_with_part(
    workbook_xml: &str,
    part_name: &str,
    preserved_timeline_caches_xml: &str,
) -> Result<String, ChartExtractionError> {
    let prefer_after = ["slicerCaches", "pivotCaches", "customWorkbookViews"];
    let before = [
        "smartTagPr",
        "smartTagTypes",
        "webPublishing",
        "fileRecoveryPr",
        "webPublishObjects",
        "extLst",
    ];
    apply_preserved_workbook_cache_refs_to_workbook_xml_with_part(
        workbook_xml,
        part_name,
        preserved_timeline_caches_xml,
        "timelineCaches",
        "timelineCaches",
        "timelineCache",
        &prefer_after,
        &before,
    )
}

fn ensure_workbook_has_namespace_prefix(
    workbook_xml: &str,
    part_name: &str,
    prefix: &str,
    uri: &str,
) -> Result<String, ChartExtractionError> {
    if prefix.is_empty() {
        return Ok(workbook_xml.to_string());
    }

    let needle = format!("xmlns:{prefix}=");

    // This helper is commonly called when we've just inserted `*:id="..."` attributes into a
    // workbook that did not previously declare the corresponding `xmlns:*`. At that point the XML
    // isn't namespace-well-formed, so we can't rely on a namespace-aware parser like `roxmltree`.
    //
    // Use a fast streaming parser to locate the `<workbook ...>` start tag and patch it in-place.
    let mut reader = Reader::from_str(workbook_xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    loop {
        let pos_before = reader.buffer_position() as usize;
        let event = reader.read_event_into(&mut buf).map_err(|e| {
            ChartExtractionError::XmlStructure(format!("{part_name}: xml parse error: {e}"))
        })?;
        let pos_after = reader.buffer_position() as usize;

        match event {
            Event::Start(ref e) | Event::Empty(ref e) if e.local_name().as_ref() == b"workbook" => {
                let tag = workbook_xml.get(pos_before..pos_after).ok_or_else(|| {
                    ChartExtractionError::XmlStructure(format!(
                        "{part_name}: invalid <workbook> start tag offsets"
                    ))
                })?;
                if tag.contains(&needle) {
                    return Ok(workbook_xml.to_string());
                }
                let trimmed = tag.trim_end();
                let insert_rel = if trimmed.ends_with("/>") {
                    trimmed.len().saturating_sub(2)
                } else if trimmed.ends_with('>') {
                    trimmed.len().saturating_sub(1)
                } else {
                    return Err(ChartExtractionError::XmlStructure(format!(
                        "{part_name}: invalid <workbook> start tag"
                    )));
                };
                let insert_pos = pos_before + insert_rel;

                let mut out = workbook_xml.to_string();
                out.insert_str(insert_pos, &format!(" xmlns:{prefix}=\"{uri}\""));
                return Ok(out);
            }
            Event::Eof => break,
            _ => {}
        }

        buf.clear();
    }

    Err(ChartExtractionError::XmlStructure(format!(
        "{part_name}: missing <workbook>"
    )))
}

fn is_self_closing_element(xml: &str) -> bool {
    let trimmed = xml.trim_end();
    trimmed.ends_with("/>") && !trimmed.contains("</")
}

fn expand_self_closing_workbook_root_if_needed<'a>(
    workbook_xml: &'a str,
    part_name: &str,
) -> Result<Cow<'a, str>, ChartExtractionError> {
    let mut reader = Reader::from_str(workbook_xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    loop {
        let pos_before = reader.buffer_position() as usize;
        let event = reader.read_event_into(&mut buf).map_err(|e| {
            ChartExtractionError::XmlStructure(format!("{part_name}: xml parse error: {e}"))
        })?;
        let pos_after = reader.buffer_position() as usize;

        match event {
            Event::Start(ref e) if e.local_name().as_ref() == b"workbook" => {
                return Ok(Cow::Borrowed(workbook_xml));
            }
            Event::Empty(ref e) if e.local_name().as_ref() == b"workbook" => {
                let tag = workbook_xml.get(pos_before..pos_after).ok_or_else(|| {
                    ChartExtractionError::XmlStructure(format!(
                        "{part_name}: invalid <workbook/> tag offsets"
                    ))
                })?;
                let qname = extract_qname_from_start_tag(tag, part_name)?;
                let close_tag = format!("</{qname}>");

                let (tag_start, trailing_ws) = split_trailing_whitespace(tag);
                let tag_start = tag_start.trim_end();
                let tag_start = tag_start.strip_suffix("/>").ok_or_else(|| {
                    ChartExtractionError::XmlStructure(format!(
                        "{part_name}: invalid self-closing <workbook/> start tag"
                    ))
                })?;
                let tag_start = tag_start.trim_end();

                let mut out = String::with_capacity(workbook_xml.len() + close_tag.len() + 1);
                out.push_str(&workbook_xml[..pos_before]);
                out.push_str(tag_start);
                out.push('>');
                out.push_str(&close_tag);
                out.push_str(trailing_ws);
                out.push_str(&workbook_xml[pos_after..]);
                return Ok(Cow::Owned(out));
            }
            Event::Eof => break,
            _ => {}
        }

        buf.clear();
    }

    Err(ChartExtractionError::XmlStructure(format!(
        "{part_name}: missing <workbook>"
    )))
}

fn extract_qname_from_start_tag<'a>(
    tag: &'a str,
    part_name: &str,
) -> Result<&'a str, ChartExtractionError> {
    let Some(rest) = tag.strip_prefix('<') else {
        return Err(ChartExtractionError::XmlStructure(format!(
            "{part_name}: invalid <workbook> start tag"
        )));
    };
    let end_rel = rest
        .char_indices()
        .find(|(_, c)| c.is_whitespace() || *c == '>' || *c == '/')
        .map(|(idx, _)| idx)
        .unwrap_or(rest.len());
    let qname = &rest[..end_rel];
    if qname.is_empty() {
        return Err(ChartExtractionError::XmlStructure(format!(
            "{part_name}: invalid <workbook> qualified name"
        )));
    }
    Ok(qname)
}

fn split_trailing_whitespace(s: &str) -> (&str, &str) {
    let trimmed_len = s.trim_end_matches(char::is_whitespace).len();
    (&s[..trimmed_len], &s[trimmed_len..])
}

fn update_pivot_caches_count_attr(pivot_caches_xml: &str, new_count: usize) -> String {
    let Some(tag_end) = pivot_caches_xml.find('>') else {
        return pivot_caches_xml.to_string();
    };
    let start_tag = &pivot_caches_xml[..tag_end];
    let Some(count_idx) = start_tag.find("count=") else {
        return pivot_caches_xml.to_string();
    };

    let value_start = count_idx + "count=".len();
    let quote = pivot_caches_xml.as_bytes().get(value_start).copied();
    let quote = match quote {
        Some(b'"') => '"',
        Some(b'\'') => '\'',
        _ => return pivot_caches_xml.to_string(),
    };

    let value_start = value_start + 1;
    let Some(value_rel_end) = pivot_caches_xml[value_start..].find(quote) else {
        return pivot_caches_xml.to_string();
    };
    let value_end = value_start + value_rel_end;

    let mut out = pivot_caches_xml.to_string();
    out.replace_range(value_start..value_end, &new_count.to_string());
    out
}

fn ensure_workbook_xml_has_pivot_caches(
    workbook_xml: &[u8],
    part_name: &str,
    pivot_caches_xml: &[u8],
) -> Result<Vec<u8>, ChartExtractionError> {
    let workbook_xml = std::str::from_utf8(workbook_xml)
        .map_err(|e| ChartExtractionError::XmlNonUtf8(part_name.to_string(), e))?;
    let pivot_caches_str = std::str::from_utf8(pivot_caches_xml)
        .map_err(|e| ChartExtractionError::XmlNonUtf8("pivotCaches".to_string(), e))?;

    let updated = apply_preserved_pivot_caches_to_workbook_xml_with_part(
        workbook_xml,
        part_name,
        pivot_caches_str,
    )?;
    Ok(updated.into_bytes())
}

fn ensure_workbook_xml_has_slicer_caches(
    workbook_xml: &[u8],
    part_name: &str,
    slicer_caches_xml: &[u8],
) -> Result<Vec<u8>, ChartExtractionError> {
    let workbook_xml = std::str::from_utf8(workbook_xml)
        .map_err(|e| ChartExtractionError::XmlNonUtf8(part_name.to_string(), e))?;
    let slicer_caches_str = std::str::from_utf8(slicer_caches_xml)
        .map_err(|e| ChartExtractionError::XmlNonUtf8("slicerCaches".to_string(), e))?;

    let updated = apply_preserved_slicer_caches_to_workbook_xml_with_part(
        workbook_xml,
        part_name,
        slicer_caches_str,
    )?;
    Ok(updated.into_bytes())
}

fn ensure_workbook_xml_has_timeline_caches(
    workbook_xml: &[u8],
    part_name: &str,
    timeline_caches_xml: &[u8],
) -> Result<Vec<u8>, ChartExtractionError> {
    let workbook_xml = std::str::from_utf8(workbook_xml)
        .map_err(|e| ChartExtractionError::XmlNonUtf8(part_name.to_string(), e))?;
    let timeline_caches_str = std::str::from_utf8(timeline_caches_xml)
        .map_err(|e| ChartExtractionError::XmlNonUtf8("timelineCaches".to_string(), e))?;

    let updated = apply_preserved_timeline_caches_to_workbook_xml_with_part(
        workbook_xml,
        part_name,
        timeline_caches_str,
    )?;
    Ok(updated.into_bytes())
}

/// Merge (or insert) a `<pivotTables>` block into a worksheet XML string.
///
/// Ordering rules (best-effort):
/// - If inserting a new `<pivotTables>`, place it before `<extLst>` when present.
/// - Otherwise insert before `</worksheet>`.
/// - Never insert inside `<sheetData>`.
///
/// If the worksheet already contains `<pivotTables>`, merge instead of inserting
/// a second `<pivotTables>` section by unioning `<pivotTable r:id="..."/>`
/// children by relationship ID.
pub fn ensure_sheet_xml_has_pivot_tables(
    sheet_xml: &[u8],
    part_name: &str,
    pivot_tables_xml: &[u8],
) -> Result<Vec<u8>, ChartExtractionError> {
    let xml_str = std::str::from_utf8(sheet_xml)
        .map_err(|e| ChartExtractionError::XmlNonUtf8(part_name.to_string(), e))?;

    let pivot_tables_str = std::str::from_utf8(pivot_tables_xml)
        .map_err(|e| ChartExtractionError::XmlNonUtf8("pivotTables".to_string(), e))?;

    let doc = Document::parse(xml_str)
        .map_err(|e| ChartExtractionError::XmlParse(part_name.to_string(), e))?;
    let root = doc.root_element();
    if root.tag_name().name() != "worksheet" {
        return Err(ChartExtractionError::XmlStructure(format!(
            "{part_name}: expected <worksheet>, found <{}>",
            root.tag_name().name()
        )));
    }

    let root_start = root.range().start;
    let worksheet_prefix = element_prefix_at(xml_str, root_start);

    let desired_rids = extract_pivot_table_rids(pivot_tables_str, "pivotTables")?;
    if desired_rids.is_empty() {
        return Ok(sheet_xml.to_vec());
    }

    let mut merged_rids: Vec<String> = Vec::new();
    let mut seen: HashSet<String> = HashSet::new();

    // Merge all existing `<pivotTables>` blocks (in case the worksheet is already malformed)
    // plus the preserved block, de-duping by relationship ID.
    let mut blocks: Vec<(usize, usize)> = doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "pivotTables")
        .map(|n| (n.range().start, n.range().end))
        .collect();
    blocks.sort_by_key(|(start, _)| *start);

    for pivot_tables in doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "pivotTables")
    {
        for rid in pivot_tables
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "pivotTable")
            .filter_map(|n| {
                n.attribute((REL_NS, "id"))
                    .or_else(|| n.attribute("r:id"))
                    .or_else(|| n.attribute("id"))
            })
        {
            if seen.insert(rid.to_string()) {
                merged_rids.push(rid.to_string());
            }
        }
    }

    for rid in desired_rids {
        if seen.insert(rid.clone()) {
            merged_rids.push(rid);
        }
    }

    let merged_block = build_pivot_tables_xml(worksheet_prefix, &merged_rids);
    let mut xml = xml_str.to_string();

    if blocks.is_empty() {
        let sheet_data_end = root
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "sheetData")
            .map(|n| n.range().end)
            .unwrap_or(0);

        let close_tag_len = crate::xml::prefixed_tag(worksheet_prefix, "worksheet").len() + 3;
        let close_idx = root.range().end.checked_sub(close_tag_len).ok_or_else(|| {
            ChartExtractionError::XmlStructure(format!("{part_name}: invalid </worksheet> tag"))
        })?;

        let ext_idx = root
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "extLst")
            .map(|n| n.range().start);

        let insert_idx = ext_idx
            .filter(|idx| *idx >= sheet_data_end && *idx <= close_idx)
            .unwrap_or_else(|| close_idx.max(sheet_data_end).min(close_idx));

        xml.insert_str(insert_idx, &merged_block);
    } else {
        // Remove all but the first `<pivotTables>` block, then replace the first with the merged one.
        let first = blocks[0];
        for (start, end) in blocks.iter().rev() {
            if (*start, *end) == first {
                continue;
            }
            xml.replace_range(*start..*end, "");
        }
        xml.replace_range(first.0..first.1, &merged_block);
    }

    // Ensure the `r` namespace exists when we insert `r:id` attributes.
    if !root_start_has_r_namespace(&xml, root_start, part_name)? {
        xml = add_r_namespace_to_root(&xml, root_start, part_name)?;
    }

    Ok(xml.into_bytes())
}

fn rewrite_relationship_ids(
    xml_bytes: &[u8],
    part_name: &str,
    id_map: &HashMap<String, String>,
) -> Result<Vec<u8>, ChartExtractionError> {
    if id_map.is_empty() {
        return Ok(xml_bytes.to_vec());
    }

    let xml = std::str::from_utf8(xml_bytes)
        .map_err(|e| ChartExtractionError::XmlNonUtf8(part_name.to_string(), e))?;

    // Preserved subtrees often rely on `xmlns:*` declarations on an ancestor element (e.g.
    // `<workbook>`/`<worksheet>`). Seed the namespace map with any `prefix:id="..."` prefixes we see
    // so we can resolve relationship attributes even when the prefix declaration is missing from
    // the fragment itself.
    let mut base_ns_map: HashMap<String, String> = HashMap::new();
    for prefix in detect_attr_prefixes(xml, "id") {
        if prefix == "xmlns" {
            continue;
        }
        base_ns_map.insert(prefix, REL_NS.to_string());
    }

    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Vec::with_capacity(xml_bytes.len()));
    let mut buf = Vec::new();

    // Namespace stack, one entry per open element. The top represents the in-scope mappings.
    let mut ns_stack: Vec<HashMap<String, String>> = Vec::new();

    loop {
        let event = reader.read_event_into(&mut buf).map_err(|e| {
            ChartExtractionError::XmlStructure(format!("{part_name}: xml parse error: {e}"))
        })?;

        match event {
            Event::Eof => break,
            Event::Start(ref e) => {
                let mut current = ns_stack
                    .last()
                    .cloned()
                    .unwrap_or_else(|| base_ns_map.clone());

                // Apply any in-fragment `xmlns:...` declarations so we can resolve attribute
                // namespace URIs correctly.
                for attr in e.attributes().with_checks(false) {
                    let attr = attr.map_err(|e| {
                        ChartExtractionError::XmlStructure(format!(
                            "{part_name}: xml attribute error: {e}"
                        ))
                    })?;
                    let key = attr.key.as_ref();
                    if let Some(prefix) = key.strip_prefix(b"xmlns:") {
                        let prefix = std::str::from_utf8(prefix).map_err(|_| {
                            ChartExtractionError::XmlStructure(format!(
                                "{part_name}: non-utf8 xmlns prefix"
                            ))
                        })?;
                        let uri = attr.unescape_value().map_err(|e| {
                            ChartExtractionError::XmlStructure(format!(
                                "{part_name}: xml attribute error: {e}"
                            ))
                        })?;
                        current.insert(prefix.to_string(), uri.into_owned());
                    }
                }

                let qname = e.name();
                let name = std::str::from_utf8(qname.as_ref()).map_err(|_| {
                    ChartExtractionError::XmlStructure(format!(
                        "{part_name}: non-utf8 element name"
                    ))
                })?;
                let mut out = BytesStart::new(name);
                for attr in e.attributes().with_checks(false) {
                    let attr = attr.map_err(|e| {
                        ChartExtractionError::XmlStructure(format!(
                            "{part_name}: xml attribute error: {e}"
                        ))
                    })?;
                    let key = attr.key.as_ref();

                    let colon = key.iter().position(|b| *b == b':');
                    let Some(colon) = colon else {
                        // Best-effort: some producers emit relationship IDs as an unprefixed `id`
                        // attribute (not namespace-well-formed, but observed in the wild). If the
                        // value matches a remapped relationship ID, rewrite it anyway.
                        if key == b"id" {
                            let value = attr.unescape_value().map_err(|e| {
                                ChartExtractionError::XmlStructure(format!(
                                    "{part_name}: xml attribute error: {e}"
                                ))
                            })?;
                            if let Some(new_id) = id_map.get(value.as_ref()) {
                                out.push_attribute((key, new_id.as_bytes()));
                            } else {
                                out.push_attribute((key, attr.value.as_ref()));
                            }
                        } else {
                            out.push_attribute((key, attr.value.as_ref()));
                        }
                        continue;
                    };
                    let prefix_bytes = &key[..colon];
                    let local = &key[colon + 1..];
                    if local != b"id" {
                        out.push_attribute((key, attr.value.as_ref()));
                        continue;
                    }
                    let prefix = std::str::from_utf8(prefix_bytes).map_err(|_| {
                        ChartExtractionError::XmlStructure(format!(
                            "{part_name}: non-utf8 attribute prefix"
                        ))
                    })?;
                    if current.get(prefix).map(|s| s.as_str()) != Some(REL_NS) {
                        out.push_attribute((key, attr.value.as_ref()));
                        continue;
                    }

                    let value = attr.unescape_value().map_err(|e| {
                        ChartExtractionError::XmlStructure(format!(
                            "{part_name}: xml attribute error: {e}"
                        ))
                    })?;
                    if let Some(new_id) = id_map.get(value.as_ref()) {
                        out.push_attribute((key, new_id.as_bytes()));
                    } else {
                        out.push_attribute((key, attr.value.as_ref()));
                    }
                }

                writer
                    .write_event(Event::Start(out))
                    .map_err(|e| ChartExtractionError::XmlStructure(format!("{part_name}: {e}")))?;

                ns_stack.push(current);
            }
            Event::Empty(ref e) => {
                let mut current = ns_stack
                    .last()
                    .cloned()
                    .unwrap_or_else(|| base_ns_map.clone());
                for attr in e.attributes().with_checks(false) {
                    let attr = attr.map_err(|e| {
                        ChartExtractionError::XmlStructure(format!(
                            "{part_name}: xml attribute error: {e}"
                        ))
                    })?;
                    let key = attr.key.as_ref();
                    if let Some(prefix) = key.strip_prefix(b"xmlns:") {
                        let prefix = std::str::from_utf8(prefix).map_err(|_| {
                            ChartExtractionError::XmlStructure(format!(
                                "{part_name}: non-utf8 xmlns prefix"
                            ))
                        })?;
                        let uri = attr.unescape_value().map_err(|e| {
                            ChartExtractionError::XmlStructure(format!(
                                "{part_name}: xml attribute error: {e}"
                            ))
                        })?;
                        current.insert(prefix.to_string(), uri.into_owned());
                    }
                }

                let qname = e.name();
                let name = std::str::from_utf8(qname.as_ref()).map_err(|_| {
                    ChartExtractionError::XmlStructure(format!(
                        "{part_name}: non-utf8 element name"
                    ))
                })?;
                let mut out = BytesStart::new(name);
                for attr in e.attributes().with_checks(false) {
                    let attr = attr.map_err(|e| {
                        ChartExtractionError::XmlStructure(format!(
                            "{part_name}: xml attribute error: {e}"
                        ))
                    })?;
                    let key = attr.key.as_ref();

                    let colon = key.iter().position(|b| *b == b':');
                    let Some(colon) = colon else {
                        // Best-effort: some producers emit relationship IDs as an unprefixed `id`
                        // attribute (not namespace-well-formed, but observed in the wild). If the
                        // value matches a remapped relationship ID, rewrite it anyway.
                        if key == b"id" {
                            let value = attr.unescape_value().map_err(|e| {
                                ChartExtractionError::XmlStructure(format!(
                                    "{part_name}: xml attribute error: {e}"
                                ))
                            })?;
                            if let Some(new_id) = id_map.get(value.as_ref()) {
                                out.push_attribute((key, new_id.as_bytes()));
                            } else {
                                out.push_attribute((key, attr.value.as_ref()));
                            }
                        } else {
                            out.push_attribute((key, attr.value.as_ref()));
                        }
                        continue;
                    };
                    let prefix_bytes = &key[..colon];
                    let local = &key[colon + 1..];
                    if local != b"id" {
                        out.push_attribute((key, attr.value.as_ref()));
                        continue;
                    }
                    let prefix = std::str::from_utf8(prefix_bytes).map_err(|_| {
                        ChartExtractionError::XmlStructure(format!(
                            "{part_name}: non-utf8 attribute prefix"
                        ))
                    })?;
                    if current.get(prefix).map(|s| s.as_str()) != Some(REL_NS) {
                        out.push_attribute((key, attr.value.as_ref()));
                        continue;
                    }

                    let value = attr.unescape_value().map_err(|e| {
                        ChartExtractionError::XmlStructure(format!(
                            "{part_name}: xml attribute error: {e}"
                        ))
                    })?;
                    if let Some(new_id) = id_map.get(value.as_ref()) {
                        out.push_attribute((key, new_id.as_bytes()));
                    } else {
                        out.push_attribute((key, attr.value.as_ref()));
                    }
                }

                writer
                    .write_event(Event::Empty(out))
                    .map_err(|e| ChartExtractionError::XmlStructure(format!("{part_name}: {e}")))?;
            }
            Event::End(ref e) => {
                writer
                    .write_event(Event::End(e.to_owned()))
                    .map_err(|e| ChartExtractionError::XmlStructure(format!("{part_name}: {e}")))?;
                ns_stack.pop();
            }
            _ => {
                writer
                    .write_event(event.to_owned())
                    .map_err(|e| ChartExtractionError::XmlStructure(format!("{part_name}: {e}")))?;
            }
        }

        buf.clear();
    }

    Ok(writer.into_inner())
}

fn extract_pivot_table_rids(
    fragment: &str,
    context: &str,
) -> Result<Vec<String>, ChartExtractionError> {
    let maybe_prefix = detect_prefix_in_fragment(fragment, "pivotTables")
        .or_else(|| detect_prefix_in_fragment(fragment, "pivotTable"));
    let prefix_decl = maybe_prefix
        .as_deref()
        .map(|p| format!(" xmlns:{p}=\"{SPREADSHEETML_NS}\""))
        .unwrap_or_default();

    let mut rel_decls = String::new();
    for prefix in detect_attr_prefixes(fragment, "id") {
        if prefix == "xmlns" {
            continue;
        }
        if maybe_prefix.as_deref() == Some(prefix.as_str()) {
            continue;
        }
        rel_decls.push_str(&format!(" xmlns:{prefix}=\"{REL_NS}\""));
    }

    let wrapped = format!("<worksheet{rel_decls}{prefix_decl}>{fragment}</worksheet>");
    let doc = Document::parse(&wrapped)
        .map_err(|e| ChartExtractionError::XmlParse(context.to_string(), e))?;

    Ok(doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "pivotTable")
        .filter_map(|n| {
            n.attribute((REL_NS, "id"))
                .or_else(|| n.attribute("r:id"))
                .or_else(|| n.attribute("id"))
        })
        .map(|s| s.to_string())
        .collect())
}

fn build_pivot_tables_xml(prefix: Option<&str>, rids: &[String]) -> String {
    let pivot_tables_tag = crate::xml::prefixed_tag(prefix, "pivotTables");
    let pivot_table_tag = crate::xml::prefixed_tag(prefix, "pivotTable");
    let mut out = String::new();
    out.push_str(&format!(r#"<{pivot_tables_tag} count="{}">"#, rids.len()));
    for rid in rids {
        out.push_str(&format!(r#"<{pivot_table_tag} r:id=""#));
        out.push_str(&xml_escape(rid));
        out.push_str(r#""/>"#);
    }
    out.push_str(&format!("</{pivot_tables_tag}>"));
    out
}

fn detect_prefix_in_fragment(fragment: &str, local: &str) -> Option<String> {
    let needle = format!(":{local}");
    let idx = fragment.find(&needle)?;
    let lt = fragment[..idx].rfind('<')?;
    let prefix = &fragment[lt + 1..idx];
    (!prefix.is_empty()).then(|| prefix.to_string())
}

fn detect_attr_prefixes(fragment: &str, local: &str) -> HashSet<String> {
    let bytes = fragment.as_bytes();
    let mut prefixes = HashSet::new();
    let needle = format!(":{local}");
    let mut i = 0usize;
    while let Some(rel_idx) = fragment[i..].find(&needle) {
        let idx = i + rel_idx;

        // Confirm this looks like an attribute name: `prefix:local ...=...`
        let mut j = idx + needle.len();
        while j < bytes.len() && bytes[j].is_ascii_whitespace() {
            j += 1;
        }
        if j >= bytes.len() || bytes[j] != b'=' {
            i = idx + needle.len();
            continue;
        }

        // Scan backwards for the beginning of the prefix.
        let mut start = idx;
        while start > 0 {
            let c = bytes[start - 1];
            let is_name_char = c.is_ascii_alphanumeric() || c == b'_' || c == b'-' || c == b'.';
            if !is_name_char {
                break;
            }
            start -= 1;
        }
        if start < idx {
            prefixes.insert(fragment[start..idx].to_string());
        }

        i = idx + needle.len();
    }
    prefixes
}

fn rewrite_spreadsheetml_prefix_in_fragment(
    fragment: &str,
    desired_prefix: Option<&str>,
    targets: &[&str],
    part_name: &str,
) -> Result<String, ChartExtractionError> {
    let mut reader = Reader::from_str(fragment);
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Vec::with_capacity(fragment.len()));
    let mut buf = Vec::new();

    loop {
        let event = reader.read_event_into(&mut buf).map_err(|e| {
            ChartExtractionError::XmlStructure(format!("{part_name}: xml parse error: {e}"))
        })?;

        match event {
            Event::Eof => break,
            Event::Start(ref e) => {
                let local_name = e.local_name();
                let local = local_name.as_ref();
                let should_rewrite = targets.iter().any(|t| local == t.as_bytes());
                if should_rewrite {
                    let local_str = std::str::from_utf8(local).map_err(|_| {
                        ChartExtractionError::XmlStructure(format!(
                            "{part_name}: non-utf8 element name"
                        ))
                    })?;
                    let new_name = crate::xml::prefixed_tag(desired_prefix, local_str);
                    let mut out = BytesStart::new(new_name.as_str());
                    for attr in e.attributes().with_checks(false) {
                        let attr = attr.map_err(|e| {
                            ChartExtractionError::XmlStructure(format!(
                                "{part_name}: xml attribute error: {e}"
                            ))
                        })?;
                        out.push_attribute((attr.key.as_ref(), attr.value.as_ref()));
                    }
                    writer.write_event(Event::Start(out)).map_err(|e| {
                        ChartExtractionError::XmlStructure(format!(
                            "{part_name}: xml write error: {e}"
                        ))
                    })?;
                } else {
                    writer
                        .write_event(Event::Start(e.to_owned()))
                        .map_err(|e| {
                            ChartExtractionError::XmlStructure(format!(
                                "{part_name}: xml write error: {e}"
                            ))
                        })?;
                }
            }
            Event::Empty(ref e) => {
                let local_name = e.local_name();
                let local = local_name.as_ref();
                let should_rewrite = targets.iter().any(|t| local == t.as_bytes());
                if should_rewrite {
                    let local_str = std::str::from_utf8(local).map_err(|_| {
                        ChartExtractionError::XmlStructure(format!(
                            "{part_name}: non-utf8 element name"
                        ))
                    })?;
                    let new_name = crate::xml::prefixed_tag(desired_prefix, local_str);
                    let mut out = BytesStart::new(new_name.as_str());
                    for attr in e.attributes().with_checks(false) {
                        let attr = attr.map_err(|e| {
                            ChartExtractionError::XmlStructure(format!(
                                "{part_name}: xml attribute error: {e}"
                            ))
                        })?;
                        out.push_attribute((attr.key.as_ref(), attr.value.as_ref()));
                    }
                    writer.write_event(Event::Empty(out)).map_err(|e| {
                        ChartExtractionError::XmlStructure(format!(
                            "{part_name}: xml write error: {e}"
                        ))
                    })?;
                } else {
                    writer
                        .write_event(Event::Empty(e.to_owned()))
                        .map_err(|e| {
                            ChartExtractionError::XmlStructure(format!(
                                "{part_name}: xml write error: {e}"
                            ))
                        })?;
                }
            }
            Event::End(ref e) => {
                let local_name = e.local_name();
                let local = local_name.as_ref();
                let should_rewrite = targets.iter().any(|t| local == t.as_bytes());
                if should_rewrite {
                    let local_str = std::str::from_utf8(local).map_err(|_| {
                        ChartExtractionError::XmlStructure(format!(
                            "{part_name}: non-utf8 element name"
                        ))
                    })?;
                    let new_name = crate::xml::prefixed_tag(desired_prefix, local_str);
                    writer
                        .write_event(Event::End(BytesEnd::new(new_name.as_str())))
                        .map_err(|e| {
                            ChartExtractionError::XmlStructure(format!(
                                "{part_name}: xml write error: {e}"
                            ))
                        })?;
                } else {
                    writer.write_event(Event::End(e.to_owned())).map_err(|e| {
                        ChartExtractionError::XmlStructure(format!(
                            "{part_name}: xml write error: {e}"
                        ))
                    })?;
                }
            }
            _ => {
                writer.write_event(event.to_owned()).map_err(|e| {
                    ChartExtractionError::XmlStructure(format!("{part_name}: xml write error: {e}"))
                })?;
            }
        }

        buf.clear();
    }

    String::from_utf8(writer.into_inner()).map_err(|_| {
        ChartExtractionError::XmlStructure(format!("{part_name}: xml output was not utf-8"))
    })
}

fn root_start_has_r_namespace(
    xml: &str,
    root_start: usize,
    part_name: &str,
) -> Result<bool, ChartExtractionError> {
    let tag_end_rel = xml[root_start..].find('>').ok_or_else(|| {
        ChartExtractionError::XmlStructure(format!("{part_name}: invalid root start tag"))
    })?;
    let tag_end = root_start + tag_end_rel;
    Ok(xml[root_start..=tag_end].contains("xmlns:r="))
}

fn add_r_namespace_to_root(
    xml: &str,
    root_start: usize,
    part_name: &str,
) -> Result<String, ChartExtractionError> {
    let tag_end_rel = xml[root_start..].find('>').ok_or_else(|| {
        ChartExtractionError::XmlStructure(format!("{part_name}: invalid root start tag"))
    })?;
    let tag_end = root_start + tag_end_rel;
    let start_tag = &xml[root_start..=tag_end];
    let trimmed = start_tag.trim_end();
    let insert_pos = if trimmed.ends_with("/>") {
        root_start + trimmed.len() - 2
    } else {
        tag_end
    };

    let mut out = xml.to_string();
    out.insert_str(insert_pos, &format!(" xmlns:r=\"{REL_NS}\""));
    Ok(out)
}

fn element_prefix_at(xml: &str, element_start: usize) -> Option<&str> {
    let rest = xml.get(element_start + 1..)?;
    let end_rel = rest
        .char_indices()
        .find(|(_, c)| c.is_whitespace() || *c == '>' || *c == '/')
        .map(|(idx, _)| idx)
        .unwrap_or(rest.len());
    let qname = &rest[..end_rel];
    qname.split_once(':').map(|(p, _)| p)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn inserts_pivot_tables_before_ext_lst() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/><extLst><ext/></extLst></worksheet>"#;
        let pivot_tables = br#"<pivotTables xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><pivotTable r:id="rId1"/></pivotTables>"#;
        let updated =
            ensure_sheet_xml_has_pivot_tables(xml, "xl/worksheets/sheet1.xml", pivot_tables)
                .expect("patch sheet xml");
        let updated_str = std::str::from_utf8(&updated).unwrap();

        let pivot_pos = updated_str.find("<pivotTables").unwrap();
        let ext_pos = updated_str.find("<extLst").unwrap();
        assert!(
            pivot_pos < ext_pos,
            "pivotTables should be inserted before extLst"
        );
    }

    #[test]
    fn inserts_pivot_caches_into_prefixed_workbook() {
        let workbook = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheets/></x:workbook>"#;
        let fragment = r#"<x:pivotCaches><x:pivotCache cacheId="1" r:id="rId1"/></x:pivotCaches>"#;

        let updated =
            apply_preserved_pivot_caches_to_workbook_xml(workbook, fragment).expect("patch");

        Document::parse(&updated).expect("output should be parseable XML");
        assert!(
            updated.contains("<x:pivotCaches"),
            "missing inserted block: {updated}"
        );
        assert!(
            !updated.contains("</workbook>"),
            "introduced unprefixed close tag: {updated}"
        );
        assert!(
            !updated.contains("</pivotCaches>"),
            "introduced unprefixed pivotCaches close tag: {updated}"
        );
    }

    #[test]
    fn inserts_pivot_caches_adds_relationship_namespace_for_non_r_prefix() {
        let workbook = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="{SPREADSHEETML_NS}" xmlns:r="{REL_NS}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#
        );
        // Simulate a preserved fragment that used a non-`r` relationships prefix with the
        // declaration living on the original `<workbook>` element (so it's missing here).
        let fragment = r#"<pivotCaches><pivotCache cacheId="1" rel:id="rId2"/></pivotCaches>"#;

        let updated =
            apply_preserved_pivot_caches_to_workbook_xml(&workbook, fragment).expect("patch");

        Document::parse(&updated).expect("output should be parseable XML");
        assert!(
            updated.contains(
                r#"xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#
            ),
            "missing xmlns:rel declaration: {updated}"
        );
        assert!(
            updated.contains(r#"rel:id="rId2""#),
            "missing rel:id attribute: {updated}"
        );
    }

    #[test]
    fn inserts_pivot_caches_into_self_closing_prefixed_workbook_root() {
        let workbook = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><x:workbook xmlns:x="{SPREADSHEETML_NS}"/>"#
        );
        let fragment = r#"<x:pivotCaches><x:pivotCache cacheId="1" r:id="rId1"/></x:pivotCaches>"#;

        let updated =
            apply_preserved_pivot_caches_to_workbook_xml(&workbook, fragment).expect("patch");

        Document::parse(&updated).expect("output should be parseable XML");
        assert!(
            updated.contains("<x:pivotCaches"),
            "missing inserted block: {updated}"
        );
        assert!(
            updated.contains("</x:workbook>"),
            "missing prefixed close tag: {updated}"
        );
        assert!(
            !updated.contains("</workbook>"),
            "introduced unprefixed close tag: {updated}"
        );
    }

    #[test]
    fn inserts_slicer_caches_into_prefixed_workbook() {
        let workbook = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheets/></x:workbook>"#;
        let fragment = r#"<x:slicerCaches><x:slicerCache r:id="rId1"/></x:slicerCaches>"#;

        let updated = apply_preserved_slicer_caches_to_workbook_xml_with_part(
            workbook,
            "xl/workbook.xml",
            fragment,
        )
        .expect("patch");

        Document::parse(&updated).expect("output should be parseable XML");
        assert!(
            updated.contains("<x:slicerCaches"),
            "missing inserted block: {updated}"
        );
        assert!(
            !updated.contains("</workbook>"),
            "introduced unprefixed close tag: {updated}"
        );
        assert!(
            !updated.contains("</slicerCaches>"),
            "introduced unprefixed slicerCaches close tag: {updated}"
        );
    }

    #[test]
    fn inserts_slicer_caches_adds_relationship_namespace_for_non_r_prefix() {
        let workbook = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="{SPREADSHEETML_NS}" xmlns:r="{REL_NS}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#
        );
        // Simulate a preserved fragment that used a non-`r` relationships prefix with the
        // declaration living on the original `<workbook>` element (so it's missing here).
        let fragment = r#"<slicerCaches><slicerCache rel:id="rId2"/></slicerCaches>"#;

        let updated = apply_preserved_slicer_caches_to_workbook_xml_with_part(
            &workbook,
            "xl/workbook.xml",
            fragment,
        )
        .expect("patch");

        Document::parse(&updated).expect("output should be parseable XML");
        assert!(
            updated.contains(
                r#"xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#
            ),
            "missing xmlns:rel declaration: {updated}"
        );
        assert!(
            updated.contains(r#"rel:id="rId2""#),
            "missing rel:id attribute: {updated}"
        );
    }

    #[test]
    fn inserts_slicer_caches_into_self_closing_prefixed_workbook_root() {
        let workbook = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><x:workbook xmlns:x="{SPREADSHEETML_NS}"/>"#
        );
        let fragment = r#"<x:slicerCaches><x:slicerCache r:id="rId1"/></x:slicerCaches>"#;

        let updated = apply_preserved_slicer_caches_to_workbook_xml_with_part(
            &workbook,
            "xl/workbook.xml",
            fragment,
        )
        .expect("patch");

        Document::parse(&updated).expect("output should be parseable XML");
        assert!(
            updated.contains("<x:slicerCaches"),
            "missing inserted block: {updated}"
        );
        assert!(
            updated.contains("</x:workbook>"),
            "missing prefixed close tag: {updated}"
        );
        assert!(
            !updated.contains("</workbook>"),
            "introduced unprefixed close tag: {updated}"
        );
        assert!(
            !updated.contains("</slicerCaches>"),
            "introduced unprefixed slicerCaches close tag: {updated}"
        );
    }

    #[test]
    fn inserts_timeline_caches_into_prefixed_workbook() {
        let workbook = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheets/></x:workbook>"#;
        let fragment = r#"<x:timelineCaches><x:timelineCache r:id="rId1"/></x:timelineCaches>"#;

        let updated = apply_preserved_timeline_caches_to_workbook_xml_with_part(
            workbook,
            "xl/workbook.xml",
            fragment,
        )
        .expect("patch");

        Document::parse(&updated).expect("output should be parseable XML");
        assert!(
            updated.contains("<x:timelineCaches"),
            "missing inserted block: {updated}"
        );
        assert!(
            !updated.contains("</workbook>"),
            "introduced unprefixed close tag: {updated}"
        );
        assert!(
            !updated.contains("</timelineCaches>"),
            "introduced unprefixed timelineCaches close tag: {updated}"
        );
    }

    #[test]
    fn inserts_timeline_caches_adds_relationship_namespace_for_non_r_prefix() {
        let workbook = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="{SPREADSHEETML_NS}" xmlns:r="{REL_NS}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#
        );
        // Simulate a preserved fragment that used a non-`r` relationships prefix with the
        // declaration living on the original `<workbook>` element (so it's missing here).
        let fragment = r#"<timelineCaches><timelineCache rel:id="rId2"/></timelineCaches>"#;

        let updated = apply_preserved_timeline_caches_to_workbook_xml_with_part(
            &workbook,
            "xl/workbook.xml",
            fragment,
        )
        .expect("patch");

        Document::parse(&updated).expect("output should be parseable XML");
        assert!(
            updated.contains(
                r#"xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#
            ),
            "missing xmlns:rel declaration: {updated}"
        );
        assert!(
            updated.contains(r#"rel:id="rId2""#),
            "missing rel:id attribute: {updated}"
        );
    }

    #[test]
    fn inserts_timeline_caches_into_self_closing_prefixed_workbook_root() {
        let workbook = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><x:workbook xmlns:x="{SPREADSHEETML_NS}"/>"#
        );
        let fragment = r#"<x:timelineCaches><x:timelineCache r:id="rId1"/></x:timelineCaches>"#;

        let updated = apply_preserved_timeline_caches_to_workbook_xml_with_part(
            &workbook,
            "xl/workbook.xml",
            fragment,
        )
        .expect("patch");

        Document::parse(&updated).expect("output should be parseable XML");
        assert!(
            updated.contains("<x:timelineCaches"),
            "missing inserted block: {updated}"
        );
        assert!(
            updated.contains("</x:workbook>"),
            "missing prefixed close tag: {updated}"
        );
        assert!(
            !updated.contains("</workbook>"),
            "introduced unprefixed close tag: {updated}"
        );
        assert!(
            !updated.contains("</timelineCaches>"),
            "introduced unprefixed timelineCaches close tag: {updated}"
        );
    }

    #[test]
    fn inserts_pivot_caches_into_self_closing_default_ns_workbook_root_adds_relationship_namespace()
    {
        let workbook = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="{SPREADSHEETML_NS}"/>"#
        );
        let fragment = r#"<pivotCaches><pivotCache cacheId="1" r:id="rId1"/></pivotCaches>"#;

        let updated =
            apply_preserved_pivot_caches_to_workbook_xml(&workbook, fragment).expect("patch");

        Document::parse(&updated).expect("output should be parseable XML");
        assert!(
            updated.contains(&format!(
                r#"<workbook xmlns="{SPREADSHEETML_NS}" xmlns:r="{REL_NS}""#
            )),
            "missing xmlns:r declaration on <workbook>: {updated}"
        );
        assert!(
            updated.contains("<pivotCaches"),
            "missing inserted block: {updated}"
        );
    }

    #[test]
    fn merges_into_self_closing_prefixed_pivot_caches() {
        let workbook = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:pivotCaches/></x:workbook>"#;
        let fragment = r#"<x:pivotCaches><x:pivotCache cacheId="1" r:id="rId1"/></x:pivotCaches>"#;

        let updated =
            apply_preserved_pivot_caches_to_workbook_xml(workbook, fragment).expect("patch");

        Document::parse(&updated).expect("output should be parseable XML");
        assert!(
            updated.contains("</x:pivotCaches>"),
            "missing prefixed close tag: {updated}"
        );
        assert!(
            !updated.contains("</pivotCaches>"),
            "introduced unprefixed pivotCaches close tag: {updated}"
        );
    }

    #[test]
    fn rewrites_relationship_ids_with_non_r_prefix() {
        let fragment = format!(r#"<a xmlns:rel="{REL_NS}" rel:id="rId1"/>"#);
        let mut id_map = HashMap::new();
        id_map.insert("rId1".to_string(), "rId9".to_string());

        let rewritten = rewrite_relationship_ids(fragment.as_bytes(), "test", &id_map)
            .expect("rewrite relationship ids");
        let rewritten = std::str::from_utf8(&rewritten).unwrap();
        assert!(
            rewritten.contains(r#"rel:id="rId9""#),
            "unexpected output: {rewritten}"
        );
        assert!(
            !rewritten.contains(r#"rel:id="rId1""#),
            "unexpected output: {rewritten}"
        );
    }

    #[test]
    fn rewrites_relationship_ids_for_unprefixed_id_attr() {
        let fragment = r#"<a id="rId1"/>"#;
        let mut id_map = HashMap::new();
        id_map.insert("rId1".to_string(), "rId9".to_string());

        let rewritten = rewrite_relationship_ids(fragment.as_bytes(), "test", &id_map)
            .expect("rewrite relationship ids");
        let rewritten = std::str::from_utf8(&rewritten).unwrap();
        assert!(
            rewritten.contains(r#"id="rId9""#),
            "unexpected output: {rewritten}"
        );
        assert!(
            !rewritten.contains(r#"id="rId1""#),
            "unexpected output: {rewritten}"
        );
    }
}
