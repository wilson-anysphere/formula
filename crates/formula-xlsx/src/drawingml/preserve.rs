use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet};
use std::io::{Read, Seek, SeekFrom};

use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader as XmlReader, Writer as XmlWriter};
use roxmltree::{Document, Node};
use zip::ZipArchive;

use crate::path::{rels_for_part, resolve_target};
use crate::preserve::rels_merge::{ensure_rels_has_relationships, RelationshipStub};
use crate::preserve::sheet_match::{
    match_sheet_by_name_or_index, workbook_sheet_parts, workbook_sheet_parts_from_workbook_xml,
};
use crate::relationships::parse_relationships;
use crate::workbook::ChartExtractionError;
use crate::zip_util::{ZipInflateBudget, DEFAULT_MAX_ZIP_PART_BYTES, DEFAULT_MAX_ZIP_TOTAL_BYTES};
use crate::XlsxPackage;

const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const DRAWING_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
const IMAGE_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const OLE_OBJECT_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject";
const CHARTSHEET_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";

/// Minimal metadata needed to re-attach an existing drawing part to a worksheet.
///
/// We preserve the worksheet relationship Id (`rId*`) so Excel continues to
/// resolve `xl/drawings/*.xml` references without regenerating relationship IDs.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SheetDrawingRelationship {
    pub rel_id: String,
    pub target: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PreservedSheetDrawings {
    pub sheet_index: usize,
    pub sheet_id: Option<u32>,
    pub drawings: Vec<SheetDrawingRelationship>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SheetRelationshipStub {
    pub rel_id: String,
    pub target: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PreservedSheetPicture {
    pub sheet_index: usize,
    pub sheet_id: Option<u32>,
    /// The `<picture>` element from the worksheet XML (outer XML).
    pub picture_xml: Vec<u8>,
    /// The relationship used by `<picture r:id="...">`.
    pub picture_rel: SheetRelationshipStub,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PreservedSheetOleObjects {
    pub sheet_index: usize,
    pub sheet_id: Option<u32>,
    /// The `<oleObjects>` subtree from the worksheet XML (outer XML).
    pub ole_objects_xml: Vec<u8>,
    /// Relationships from the worksheet `.rels` required by `<oleObjects>`.
    pub ole_object_rels: Vec<SheetRelationshipStub>,
}

/// Relationship metadata required by a preserved worksheet fragment.
///
/// Unlike [`SheetRelationshipStub`], this includes the relationship `Type` attribute so fragments
/// like `<controls>` / `<drawingHF>` (and any nested `r:*` attributes) can be re-attached without
/// guessing the relationship kind.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SheetRelationshipStubWithType {
    pub rel_id: String,
    pub type_: String,
    pub target: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PreservedSheetControls {
    pub sheet_index: usize,
    pub sheet_id: Option<u32>,
    /// The `<controls>` subtree from the worksheet XML (outer XML).
    pub controls_xml: Vec<u8>,
    /// Relationships from the worksheet `.rels` required by `<controls>`.
    pub control_rels: Vec<SheetRelationshipStubWithType>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PreservedSheetDrawingHF {
    pub sheet_index: usize,
    pub sheet_id: Option<u32>,
    /// The `<drawingHF>` element from the worksheet XML (outer XML).
    pub drawing_hf_xml: Vec<u8>,
    /// Relationships from the worksheet `.rels` required by `<drawingHF>`.
    pub drawing_hf_rels: Vec<SheetRelationshipStubWithType>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PreservedChartSheet {
    pub sheet_index: usize,
    pub sheet_id: Option<u32>,
    /// The `rId*` referenced by the `<sheet r:id="...">` entry in `xl/workbook.xml`.
    pub rel_id: String,
    /// The matching relationship target in `xl/_rels/workbook.xml.rels`.
    pub rel_target: String,
    /// Optional `state=` attribute value from `xl/workbook.xml` (`hidden`/`veryHidden`).
    pub state: Option<String>,
    /// The chartsheet XML part path (e.g. `xl/chartsheets/sheet1.xml`).
    pub part_name: String,
}

/// A slice of an XLSX package that is required to preserve DrawingML objects
/// (including charts) across a "read -> write" pipeline that doesn't otherwise
/// round-trip the original package structure.
#[derive(Debug, Clone)]
pub struct PreservedDrawingParts {
    pub content_types_xml: Vec<u8>,
    pub parts: BTreeMap<String, Vec<u8>>,
    pub sheet_drawings: BTreeMap<String, PreservedSheetDrawings>,
    pub sheet_pictures: BTreeMap<String, PreservedSheetPicture>,
    pub sheet_ole_objects: BTreeMap<String, PreservedSheetOleObjects>,
    pub sheet_controls: BTreeMap<String, PreservedSheetControls>,
    pub sheet_drawing_hfs: BTreeMap<String, PreservedSheetDrawingHF>,
    pub chart_sheets: BTreeMap<String, PreservedChartSheet>,
}

impl PreservedDrawingParts {
    pub fn is_empty(&self) -> bool {
        self.parts.is_empty()
            && self.sheet_drawings.values().all(|v| v.drawings.is_empty())
            && self.sheet_pictures.is_empty()
            && self.sheet_ole_objects.is_empty()
            && self.sheet_controls.is_empty()
            && self.sheet_drawing_hfs.is_empty()
            && self.chart_sheets.is_empty()
    }
}

/// Streaming variant of [`XlsxPackage::preserve_drawing_parts`].
///
/// This reads only the subset of ZIP parts required to capture DrawingML objects (including
/// charts) for a later regeneration-based round-trip.
///
/// Unlike [`XlsxPackage::from_bytes`], this does **not** inflate every ZIP entry into memory.
pub fn preserve_drawing_parts_from_reader<R: Read + Seek>(
    reader: R,
) -> Result<PreservedDrawingParts, ChartExtractionError> {
    preserve_drawing_parts_from_reader_limited(reader, DEFAULT_MAX_ZIP_PART_BYTES, DEFAULT_MAX_ZIP_TOTAL_BYTES)
}

/// Streaming variant of [`XlsxPackage::preserve_drawing_parts`] with configurable ZIP inflation
/// limits.
///
/// This is primarily useful for callers that treat the input as untrusted (e.g. desktop IPC
/// surfaces) and want tighter bounds than the crate defaults.
pub fn preserve_drawing_parts_from_reader_limited<R: Read + Seek>(
    mut reader: R,
    max_part_bytes: u64,
    max_total_bytes: u64,
) -> Result<PreservedDrawingParts, ChartExtractionError> {
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

    let workbook_xml =
        read_zip_part_required(&mut archive, "xl/workbook.xml", max_part_bytes, &mut budget)?;
    let workbook_rels_xml =
        read_zip_part_optional(&mut archive, "xl/_rels/workbook.xml.rels", max_part_bytes, &mut budget)?;

    let chart_sheets = extract_workbook_chart_sheets_from_workbook_parts(
        &workbook_xml,
        workbook_rels_xml.as_deref(),
        &part_names,
    )?;
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

    let mut root_parts: BTreeSet<String> = BTreeSet::new();
    for sheet in chart_sheets.values() {
        root_parts.insert(sheet.part_name.clone());
    }

    let mut sheet_drawings: BTreeMap<String, PreservedSheetDrawings> = BTreeMap::new();
    let mut sheet_pictures: BTreeMap<String, PreservedSheetPicture> = BTreeMap::new();
    let mut sheet_ole_objects: BTreeMap<String, PreservedSheetOleObjects> = BTreeMap::new();
    let mut sheet_controls: BTreeMap<String, PreservedSheetControls> = BTreeMap::new();
    let mut sheet_drawing_hfs: BTreeMap<String, PreservedSheetDrawingHF> = BTreeMap::new();

    for sheet in sheets {
        let sheet_rels_part = rels_for_part(&sheet.part_name);

        // Best-effort: some producers emit malformed `.rels` parts. For preservation we skip
        // relationship discovery for this sheet rather than erroring.
        let rels = match read_zip_part_optional(
            &mut archive,
            &sheet_rels_part,
            max_part_bytes,
            &mut budget,
        )? {
            Some(xml) => match parse_relationships(&xml, &sheet_rels_part) {
                Ok(rels) => rels,
                Err(_) => Vec::new(),
            },
            None => Vec::new(),
        };

        for rel in &rels {
            if rel
                .target_mode
                .as_deref()
                .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
            {
                continue;
            }
            let resolved = resolve_target(&sheet.part_name, &rel.target);
            if is_drawing_adjacent_relationship(rel.type_.as_str(), &resolved) {
                if let Some(resolved) = find_part_name(&part_names, &resolved) {
                    root_parts.insert(resolved);
                }
            }
        }

        if sheet.part_name.starts_with("xl/chartsheets/") {
            if let Some(resolved) = find_part_name(&part_names, &sheet.part_name) {
                root_parts.insert(resolved);
            } else {
                root_parts.insert(sheet.part_name.clone());
            }
        }

        // Only parse the worksheet XML when the sheet has at least one drawing-adjacent relationship.
        // This avoids inflating large sheet XML parts for the common case where a workbook has many
        // data-only sheets.
        let should_parse_sheet_xml = sheet.part_name.starts_with("xl/chartsheets/")
            || rels.iter().any(|rel| {
                if rel
                    .target_mode
                    .as_deref()
                    .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
                {
                    return false;
                }
                let resolved = resolve_target(&sheet.part_name, &rel.target);
                is_drawing_adjacent_relationship(rel.type_.as_str(), &resolved)
            });
        if !should_parse_sheet_xml {
            continue;
        }

        let Some(sheet_xml) =
            read_zip_part_optional(&mut archive, &sheet.part_name, max_part_bytes, &mut budget)?
        else {
            continue;
        };

        let sheet_xml_str = std::str::from_utf8(&sheet_xml)
            .map_err(|e| ChartExtractionError::XmlNonUtf8(sheet.part_name.clone(), e))?;
        let doc = Document::parse(sheet_xml_str)
            .map_err(|e| ChartExtractionError::XmlParse(sheet.part_name.clone(), e))?;

        fn is_drawing_node(node: Node<'_, '_>) -> bool {
            node.is_element() && node.tag_name().name() == "drawing"
        }

        let drawing_rids: Vec<String> = crate::drawingml::anchor::descendants_selecting_alternate_content(
            doc.root_element(),
            is_drawing_node,
            is_drawing_node,
        )
        .into_iter()
        .filter_map(|n| {
            n.attribute((REL_NS, "id"))
                .or_else(|| n.attribute("r:id"))
                .or_else(|| n.attribute("id"))
        })
        .map(|s| s.to_string())
        .collect();

        let picture_node = doc
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "picture");
        let picture = picture_node.and_then(|node| {
            let rid = node
                .attribute((REL_NS, "id"))
                .or_else(|| node.attribute("r:id"))
                .or_else(|| node.attribute("id"))?;
            Some((rid.to_string(), sheet_xml_str.as_bytes()[node.range()].to_vec()))
        });

        let ole_objects_node = doc
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "oleObjects");
        let ole_objects = ole_objects_node.map(|node| {
            let rids: Vec<String> = node
                .descendants()
                .filter(|n| n.is_element() && n.tag_name().name() == "oleObject")
                .filter_map(|n| {
                    n.attribute((REL_NS, "id"))
                        .or_else(|| n.attribute("r:id"))
                        .or_else(|| n.attribute("id"))
                })
                .map(|s| s.to_string())
                .collect();
            (sheet_xml_str.as_bytes()[node.range()].to_vec(), rids)
        });

        let controls_node = doc
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "controls");
        let controls = controls_node.map(|node| {
            (
                sheet_xml_str.as_bytes()[node.range()].to_vec(),
                extract_relationship_ids(node),
            )
        });

        let drawing_hf_node = doc
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "drawingHF");
        let drawing_hf = drawing_hf_node.map(|node| {
            (
                sheet_xml_str.as_bytes()[node.range()].to_vec(),
                extract_relationship_ids(node),
            )
        });

        if drawing_rids.is_empty()
            && picture.is_none()
            && ole_objects.is_none()
            && controls.is_none()
            && drawing_hf.is_none()
        {
            continue;
        }

        let rel_map: HashMap<_, _> = rels.into_iter().map(|r| (r.id.clone(), r)).collect();

        if !drawing_rids.is_empty() {
            let mut drawings = Vec::new();
            for rid in drawing_rids {
                if let Some(rel) = rel_map.get(&rid) {
                    if rel.type_ == DRAWING_REL_TYPE
                        && !rel
                            .target_mode
                            .as_deref()
                            .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
                    {
                        drawings.push(SheetDrawingRelationship {
                            rel_id: rid.clone(),
                            target: rel.target.clone(),
                        });
                    }
                }
            }

            if !drawings.is_empty() {
                sheet_drawings.insert(
                    sheet.name.clone(),
                    PreservedSheetDrawings {
                        sheet_index: sheet.index,
                        sheet_id: sheet.sheet_id,
                        drawings,
                    },
                );
            }
        }

        if let Some((rid, picture_xml)) = picture {
            if let Some(rel) = rel_map.get(&rid) {
                if rel.type_ == IMAGE_REL_TYPE
                    && !rel
                        .target_mode
                        .as_deref()
                        .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
                {
                    sheet_pictures.insert(
                        sheet.name.clone(),
                        PreservedSheetPicture {
                            sheet_index: sheet.index,
                            sheet_id: sheet.sheet_id,
                            picture_xml,
                            picture_rel: SheetRelationshipStub {
                                rel_id: rid.clone(),
                                target: rel.target.clone(),
                            },
                        },
                    );
                }
            }
        }

        if let Some((ole_objects_xml, rids)) = ole_objects {
            if !rids.is_empty() {
                let mut rels = Vec::new();
                let mut missing_rel = false;
                for rid in &rids {
                    match rel_map.get(rid) {
                        Some(rel)
                            if rel.type_ == OLE_OBJECT_REL_TYPE
                                && !rel.target_mode.as_deref().is_some_and(|mode| {
                                    mode.trim().eq_ignore_ascii_case("External")
                                }) =>
                        {
                            rels.push(SheetRelationshipStub {
                                rel_id: rid.clone(),
                                target: rel.target.clone(),
                            })
                        }
                        _ => {
                            missing_rel = true;
                            break;
                        }
                    }
                }

                if !missing_rel {
                    sheet_ole_objects.insert(
                        sheet.name.clone(),
                        PreservedSheetOleObjects {
                            sheet_index: sheet.index,
                            sheet_id: sheet.sheet_id,
                            ole_objects_xml,
                            ole_object_rels: rels,
                        },
                    );
                }
            }
        }

        if let Some((controls_xml, rids)) = controls {
            let mut rels = Vec::new();
            let mut missing_rel = false;
            for rid in &rids {
                match rel_map.get(rid) {
                    Some(rel)
                        if !rel.target_mode.as_deref().is_some_and(|mode| {
                            mode.trim().eq_ignore_ascii_case("External")
                        }) =>
                    {
                        rels.push(SheetRelationshipStubWithType {
                            rel_id: rid.clone(),
                            type_: rel.type_.clone(),
                            target: rel.target.clone(),
                        })
                    }
                    _ => {
                        missing_rel = true;
                        break;
                    }
                }
            }

            if !missing_rel {
                sheet_controls.insert(
                    sheet.name.clone(),
                    PreservedSheetControls {
                        sheet_index: sheet.index,
                        sheet_id: sheet.sheet_id,
                        controls_xml,
                        control_rels: rels,
                    },
                );
            }
        }

        if let Some((drawing_hf_xml, rids)) = drawing_hf {
            let mut rels = Vec::new();
            let mut missing_rel = false;
            for rid in &rids {
                match rel_map.get(rid) {
                    Some(rel)
                        if !rel.target_mode.as_deref().is_some_and(|mode| {
                            mode.trim().eq_ignore_ascii_case("External")
                        }) =>
                    {
                        rels.push(SheetRelationshipStubWithType {
                            rel_id: rid.clone(),
                            type_: rel.type_.clone(),
                            target: rel.target.clone(),
                        })
                    }
                    _ => {
                        missing_rel = true;
                        break;
                    }
                }
            }

            if !missing_rel {
                sheet_drawing_hfs.insert(
                    sheet.name.clone(),
                    PreservedSheetDrawingHF {
                        sheet_index: sheet.index,
                        sheet_id: sheet.sheet_id,
                        drawing_hf_xml,
                        drawing_hf_rels: rels,
                    },
                );
            }
        }
    }

    let parts = collect_transitive_related_parts_from_archive(
        &mut archive,
        &part_names,
        root_parts,
        max_part_bytes,
        &mut budget,
    )?;

    Ok(PreservedDrawingParts {
        content_types_xml,
        parts,
        sheet_drawings,
        sheet_pictures,
        sheet_ole_objects,
        sheet_controls,
        sheet_drawing_hfs,
        chart_sheets,
    })
}

impl XlsxPackage {
    /// Extract the DrawingML/chart-related parts of an XLSX package so they can
    /// be re-applied to another package later (e.g. after regenerating sheet XML).
    pub fn preserve_drawing_parts(&self) -> Result<PreservedDrawingParts, ChartExtractionError> {
        let content_types_xml = self
            .part("[Content_Types].xml")
            .ok_or_else(|| ChartExtractionError::MissingPart("[Content_Types].xml".to_string()))?
            .to_vec();
        let chart_sheets = extract_workbook_chart_sheets(self)?;
        let sheets = workbook_sheet_parts(self)?;
        let mut root_parts: BTreeSet<String> = BTreeSet::new();
        for sheet in chart_sheets.values() {
            root_parts.insert(sheet.part_name.clone());
        }
        let mut sheet_drawings: BTreeMap<String, PreservedSheetDrawings> = BTreeMap::new();
        let mut sheet_pictures: BTreeMap<String, PreservedSheetPicture> = BTreeMap::new();
        let mut sheet_ole_objects: BTreeMap<String, PreservedSheetOleObjects> = BTreeMap::new();
        let mut sheet_controls: BTreeMap<String, PreservedSheetControls> = BTreeMap::new();
        let mut sheet_drawing_hfs: BTreeMap<String, PreservedSheetDrawingHF> = BTreeMap::new();

        for sheet in sheets {
            let sheet_rels_part = rels_for_part(&sheet.part_name);
            // Best-effort: some producers emit malformed `.rels` parts. For preservation we skip
            // relationship discovery for this sheet rather than erroring.
            let rels = match self.part(&sheet_rels_part) {
                Some(xml) => match parse_relationships(xml, &sheet_rels_part) {
                    Ok(rels) => rels,
                    Err(_) => Vec::new(),
                },
                None => Vec::new(),
            };

            for rel in &rels {
                if rel
                    .target_mode
                    .as_deref()
                    .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
                {
                    continue;
                }
                let resolved = resolve_target(&sheet.part_name, &rel.target);
                if is_drawing_adjacent_relationship(rel.type_.as_str(), &resolved)
                    && self.part(&resolved).is_some()
                {
                    root_parts.insert(resolved);
                }
            }

            let Some(sheet_xml) = self.part(&sheet.part_name) else {
                continue;
            };

            if sheet.part_name.starts_with("xl/chartsheets/") {
                root_parts.insert(sheet.part_name.clone());
            }

            let sheet_xml_str = std::str::from_utf8(sheet_xml)
                .map_err(|e| ChartExtractionError::XmlNonUtf8(sheet.part_name.clone(), e))?;
            let doc = Document::parse(sheet_xml_str)
                .map_err(|e| ChartExtractionError::XmlParse(sheet.part_name.clone(), e))?;

            fn is_drawing_node(node: Node<'_, '_>) -> bool {
                node.is_element() && node.tag_name().name() == "drawing"
            }

            let drawing_rids: Vec<String> = crate::drawingml::anchor::descendants_selecting_alternate_content(
                doc.root_element(),
                is_drawing_node,
                is_drawing_node,
            )
            .into_iter()
            .filter_map(|n| {
                n.attribute((REL_NS, "id"))
                    .or_else(|| n.attribute("r:id"))
                    .or_else(|| n.attribute("id"))
            })
            .map(|s| s.to_string())
            .collect();

            let picture_node = doc
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "picture");
            let picture = picture_node.and_then(|node| {
                let rid = node
                    .attribute((REL_NS, "id"))
                    .or_else(|| node.attribute("r:id"))
                    .or_else(|| node.attribute("id"))?;
                Some((
                    rid.to_string(),
                    sheet_xml_str.as_bytes()[node.range()].to_vec(),
                ))
            });

            let ole_objects_node = doc
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "oleObjects");
            let ole_objects = ole_objects_node.map(|node| {
                let rids: Vec<String> = node
                    .descendants()
                    .filter(|n| n.is_element() && n.tag_name().name() == "oleObject")
                    .filter_map(|n| {
                        n.attribute((REL_NS, "id"))
                            .or_else(|| n.attribute("r:id"))
                            .or_else(|| n.attribute("id"))
                    })
                    .map(|s| s.to_string())
                    .collect();
                (sheet_xml_str.as_bytes()[node.range()].to_vec(), rids)
            });

            let controls_node = doc
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "controls");
            let controls = controls_node.map(|node| {
                (
                    sheet_xml_str.as_bytes()[node.range()].to_vec(),
                    extract_relationship_ids(node),
                )
            });

            let drawing_hf_node = doc
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "drawingHF");
            let drawing_hf = drawing_hf_node.map(|node| {
                (
                    sheet_xml_str.as_bytes()[node.range()].to_vec(),
                    extract_relationship_ids(node),
                )
            });

            if drawing_rids.is_empty()
                && picture.is_none()
                && ole_objects.is_none()
                && controls.is_none()
                && drawing_hf.is_none()
            {
                continue;
            }

            let rel_map: HashMap<_, _> = rels.into_iter().map(|r| (r.id.clone(), r)).collect();

            if !drawing_rids.is_empty() {
                let mut drawings = Vec::new();
                for rid in drawing_rids {
                    if let Some(rel) = rel_map.get(&rid) {
                        if rel.type_ == DRAWING_REL_TYPE
                            && !rel
                                .target_mode
                                .as_deref()
                                .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
                        {
                            drawings.push(SheetDrawingRelationship {
                                rel_id: rid.clone(),
                                target: rel.target.clone(),
                            });
                        }
                    }
                }

                if !drawings.is_empty() {
                    sheet_drawings.insert(
                        sheet.name.clone(),
                        PreservedSheetDrawings {
                            sheet_index: sheet.index,
                            sheet_id: sheet.sheet_id,
                            drawings,
                        },
                    );
                }
            }

            if let Some((rid, picture_xml)) = picture {
                if let Some(rel) = rel_map.get(&rid) {
                    if rel.type_ == IMAGE_REL_TYPE
                        && !rel
                            .target_mode
                            .as_deref()
                            .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
                    {
                        sheet_pictures.insert(
                            sheet.name.clone(),
                            PreservedSheetPicture {
                                sheet_index: sheet.index,
                                sheet_id: sheet.sheet_id,
                                picture_xml,
                                picture_rel: SheetRelationshipStub {
                                    rel_id: rid.clone(),
                                    target: rel.target.clone(),
                                },
                            },
                        );
                    }
                }
            }

            if let Some((ole_objects_xml, rids)) = ole_objects {
                if !rids.is_empty() {
                    let mut rels = Vec::new();
                    let mut missing_rel = false;
                    for rid in &rids {
                        match rel_map.get(rid) {
                            Some(rel)
                                if rel.type_ == OLE_OBJECT_REL_TYPE
                                    && !rel.target_mode.as_deref().is_some_and(|mode| {
                                        mode.trim().eq_ignore_ascii_case("External")
                                    }) =>
                            {
                                rels.push(SheetRelationshipStub {
                                    rel_id: rid.clone(),
                                    target: rel.target.clone(),
                                })
                            }
                            _ => {
                                missing_rel = true;
                                break;
                            }
                        }
                    }

                    if !missing_rel {
                        sheet_ole_objects.insert(
                            sheet.name.clone(),
                            PreservedSheetOleObjects {
                                sheet_index: sheet.index,
                                sheet_id: sheet.sheet_id,
                                ole_objects_xml,
                                ole_object_rels: rels,
                            },
                        );
                    }
                }
            }

            if let Some((controls_xml, rids)) = controls {
                let mut rels = Vec::new();
                let mut missing_rel = false;
                for rid in &rids {
                    match rel_map.get(rid) {
                        Some(rel)
                            if !rel.target_mode.as_deref().is_some_and(|mode| {
                                mode.trim().eq_ignore_ascii_case("External")
                            }) =>
                        {
                            rels.push(SheetRelationshipStubWithType {
                                rel_id: rid.clone(),
                                type_: rel.type_.clone(),
                                target: rel.target.clone(),
                            })
                        }
                        _ => {
                            missing_rel = true;
                            break;
                        }
                    }
                }

                if !missing_rel {
                    sheet_controls.insert(
                        sheet.name.clone(),
                        PreservedSheetControls {
                            sheet_index: sheet.index,
                            sheet_id: sheet.sheet_id,
                            controls_xml,
                            control_rels: rels,
                        },
                    );
                }
            }

            if let Some((drawing_hf_xml, rids)) = drawing_hf {
                let mut rels = Vec::new();
                let mut missing_rel = false;
                for rid in &rids {
                    match rel_map.get(rid) {
                        Some(rel)
                            if !rel.target_mode.as_deref().is_some_and(|mode| {
                                mode.trim().eq_ignore_ascii_case("External")
                            }) =>
                        {
                            rels.push(SheetRelationshipStubWithType {
                                rel_id: rid.clone(),
                                type_: rel.type_.clone(),
                                target: rel.target.clone(),
                            })
                        }
                        _ => {
                            missing_rel = true;
                            break;
                        }
                    }
                }

                if !missing_rel {
                    sheet_drawing_hfs.insert(
                        sheet.name.clone(),
                        PreservedSheetDrawingHF {
                            sheet_index: sheet.index,
                            sheet_id: sheet.sheet_id,
                            drawing_hf_xml,
                            drawing_hf_rels: rels,
                        },
                    );
                }
            }
        }

        let parts = crate::preserve::opc_graph::collect_transitive_related_parts(
            self,
            root_parts.into_iter(),
        )?;

        Ok(PreservedDrawingParts {
            content_types_xml,
            parts,
            sheet_drawings,
            sheet_pictures,
            sheet_ole_objects,
            sheet_controls,
            sheet_drawing_hfs,
            chart_sheets,
        })
    }

    /// Apply previously captured drawing/chart parts to this package.
    ///
    /// This function is intentionally conservative:
    /// - It copies the raw parts byte-for-byte.
    /// - It re-attaches drawings to matching sheets (by sheet name) by ensuring
    ///   `<drawing r:id="..."/>` exists in the worksheet XML and the
    ///   corresponding relationship exists in the worksheet `.rels`.
    /// - It merges required `[Content_Types].xml` entries for inserted parts.
    pub fn apply_preserved_drawing_parts(
        &mut self,
        preserved: &PreservedDrawingParts,
    ) -> Result<(), ChartExtractionError> {
        for (name, bytes) in &preserved.parts {
            self.set_part(name.clone(), bytes.clone());
        }

        self.merge_content_types(&preserved.content_types_xml, preserved.parts.keys())?;

        if !preserved.chart_sheets.is_empty() {
            ensure_workbook_has_chartsheets(self, &preserved.chart_sheets)?;
        }

        let sheets = workbook_sheet_parts(self)?;
        for (sheet_name, preserved_sheet) in &preserved.sheet_drawings {
            if preserved_sheet.drawings.is_empty() {
                continue;
            }
            let Some(sheet) =
                match_sheet_by_name_or_index(&sheets, sheet_name, preserved_sheet.sheet_index)
            else {
                continue;
            };

            let sheet_rels_part = rels_for_part(&sheet.part_name);
            let drawing_rels: Vec<RelationshipStub> = preserved_sheet
                .drawings
                .iter()
                .map(|drawing| RelationshipStub {
                    rel_id: drawing.rel_id.clone(),
                    target: drawing.target.clone(),
                })
                .collect();
            let (updated_rels, rid_map) = ensure_rels_has_relationships(
                self.part(&sheet_rels_part),
                &sheet_rels_part,
                &sheet.part_name,
                DRAWING_REL_TYPE,
                &drawing_rels,
            )?;
            self.set_part(sheet_rels_part, updated_rels);

            let Some(sheet_xml) = self.part(&sheet.part_name) else {
                continue;
            };
            let updated_sheet_xml = ensure_sheet_xml_has_drawings(
                sheet_xml,
                &sheet.part_name,
                &preserved_sheet.drawings,
                &rid_map,
            )?;
            self.set_part(sheet.part_name.clone(), updated_sheet_xml);
        }

        for (sheet_name, preserved_sheet) in &preserved.sheet_drawing_hfs {
            let Some(sheet) =
                match_sheet_by_name_or_index(&sheets, sheet_name, preserved_sheet.sheet_index)
            else {
                continue;
            };

            let Some(sheet_xml) = self.part(&sheet.part_name) else {
                continue;
            };
            let sheet_xml = sheet_xml.to_vec();

            let sheet_xml_str = std::str::from_utf8(&sheet_xml)
                .map_err(|e| ChartExtractionError::XmlNonUtf8(sheet.part_name.clone(), e))?;
            if xml_contains_local_element(sheet_xml_str, "drawingHF") {
                continue;
            }

            let mut rels_by_type: BTreeMap<String, Vec<RelationshipStub>> = BTreeMap::new();
            for rel in &preserved_sheet.drawing_hf_rels {
                rels_by_type
                    .entry(rel.type_.clone())
                    .or_default()
                    .push(RelationshipStub {
                        rel_id: rel.rel_id.clone(),
                        target: rel.target.clone(),
                    });
            }

            let sheet_rels_part = rels_for_part(&sheet.part_name);
            let mut rels_xml = self.part(&sheet_rels_part).map(|b| b.to_vec());
            let mut rid_map: HashMap<String, String> = HashMap::new();
            for (rel_type, rels) in rels_by_type {
                let (updated_rels, map) = ensure_rels_has_relationships(
                    rels_xml.as_deref(),
                    &sheet_rels_part,
                    &sheet.part_name,
                    &rel_type,
                    &rels,
                )?;
                rels_xml = Some(updated_rels);
                rid_map.extend(map);
            }
            if let Some(updated_rels) = rels_xml {
                if !preserved_sheet.drawing_hf_rels.is_empty() {
                    self.set_part(sheet_rels_part, updated_rels);
                }
            }

            let updated_sheet_xml = ensure_sheet_xml_has_drawing_hf(
                &sheet_xml,
                &sheet.part_name,
                &preserved_sheet.drawing_hf_xml,
                &rid_map,
            )?;
            self.set_part(sheet.part_name.clone(), updated_sheet_xml);
        }

        for (sheet_name, preserved_sheet) in &preserved.sheet_pictures {
            let Some(sheet) =
                match_sheet_by_name_or_index(&sheets, sheet_name, preserved_sheet.sheet_index)
            else {
                continue;
            };

            let Some(sheet_xml) = self.part(&sheet.part_name) else {
                continue;
            };
            let sheet_xml = sheet_xml.to_vec();

            let sheet_xml_str = std::str::from_utf8(&sheet_xml)
                .map_err(|e| ChartExtractionError::XmlNonUtf8(sheet.part_name.clone(), e))?;
            if xml_contains_local_element(sheet_xml_str, "picture") {
                continue;
            }

            let sheet_rels_part = rels_for_part(&sheet.part_name);
            let picture_rels = [RelationshipStub {
                rel_id: preserved_sheet.picture_rel.rel_id.clone(),
                target: preserved_sheet.picture_rel.target.clone(),
            }];
            let (updated_rels, rid_map) = ensure_rels_has_relationships(
                self.part(&sheet_rels_part),
                &sheet_rels_part,
                &sheet.part_name,
                IMAGE_REL_TYPE,
                &picture_rels,
            )?;
            self.set_part(sheet_rels_part, updated_rels);

            let updated_sheet_xml = ensure_sheet_xml_has_picture(
                &sheet_xml,
                &sheet.part_name,
                &preserved_sheet.picture_xml,
                &rid_map,
            )?;
            self.set_part(sheet.part_name.clone(), updated_sheet_xml);
        }

        for (sheet_name, preserved_sheet) in &preserved.sheet_ole_objects {
            let Some(sheet) =
                match_sheet_by_name_or_index(&sheets, sheet_name, preserved_sheet.sheet_index)
            else {
                continue;
            };

            let Some(sheet_xml) = self.part(&sheet.part_name) else {
                continue;
            };
            let sheet_xml = sheet_xml.to_vec();

            let sheet_xml_str = std::str::from_utf8(&sheet_xml)
                .map_err(|e| ChartExtractionError::XmlNonUtf8(sheet.part_name.clone(), e))?;
            if xml_contains_local_element(sheet_xml_str, "oleObjects") {
                continue;
            }

            let sheet_rels_part = rels_for_part(&sheet.part_name);
            let ole_rels: Vec<RelationshipStub> = preserved_sheet
                .ole_object_rels
                .iter()
                .map(|rel| RelationshipStub {
                    rel_id: rel.rel_id.clone(),
                    target: rel.target.clone(),
                })
                .collect();
            let (updated_rels, rid_map) = ensure_rels_has_relationships(
                self.part(&sheet_rels_part),
                &sheet_rels_part,
                &sheet.part_name,
                OLE_OBJECT_REL_TYPE,
                &ole_rels,
            )?;
            self.set_part(sheet_rels_part, updated_rels);

            let updated_sheet_xml = ensure_sheet_xml_has_ole_objects(
                &sheet_xml,
                &sheet.part_name,
                &preserved_sheet.ole_objects_xml,
                &rid_map,
            )?;
            self.set_part(sheet.part_name.clone(), updated_sheet_xml);
        }

        for (sheet_name, preserved_sheet) in &preserved.sheet_controls {
            let Some(sheet) =
                match_sheet_by_name_or_index(&sheets, sheet_name, preserved_sheet.sheet_index)
            else {
                continue;
            };

            let Some(sheet_xml) = self.part(&sheet.part_name) else {
                continue;
            };
            let sheet_xml = sheet_xml.to_vec();

            let sheet_xml_str = std::str::from_utf8(&sheet_xml)
                .map_err(|e| ChartExtractionError::XmlNonUtf8(sheet.part_name.clone(), e))?;
            if xml_contains_local_element(sheet_xml_str, "controls") {
                continue;
            }

            let mut rels_by_type: BTreeMap<String, Vec<RelationshipStub>> = BTreeMap::new();
            for rel in &preserved_sheet.control_rels {
                rels_by_type
                    .entry(rel.type_.clone())
                    .or_default()
                    .push(RelationshipStub {
                        rel_id: rel.rel_id.clone(),
                        target: rel.target.clone(),
                    });
            }

            let sheet_rels_part = rels_for_part(&sheet.part_name);
            let mut rels_xml = self.part(&sheet_rels_part).map(|b| b.to_vec());
            let mut rid_map: HashMap<String, String> = HashMap::new();
            for (rel_type, rels) in rels_by_type {
                let (updated_rels, map) = ensure_rels_has_relationships(
                    rels_xml.as_deref(),
                    &sheet_rels_part,
                    &sheet.part_name,
                    &rel_type,
                    &rels,
                )?;
                rels_xml = Some(updated_rels);
                rid_map.extend(map);
            }
            if let Some(updated_rels) = rels_xml {
                if !preserved_sheet.control_rels.is_empty() {
                    self.set_part(sheet_rels_part, updated_rels);
                }
            }

            let updated_sheet_xml = ensure_sheet_xml_has_controls(
                &sheet_xml,
                &sheet.part_name,
                &preserved_sheet.controls_xml,
                &rid_map,
            )?;
            self.set_part(sheet.part_name.clone(), updated_sheet_xml);
        }

        Ok(())
    }
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

fn find_part_name(part_names: &HashSet<String>, candidate: &str) -> Option<String> {
    if let Some(found) = part_names.get(candidate) {
        return Some(found.clone());
    }
    part_names
        .iter()
        .find(|name| crate::zip_util::zip_part_names_equivalent(name.as_str(), candidate))
        .cloned()
}

fn collect_transitive_related_parts_from_archive<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    part_names: &HashSet<String>,
    root_parts: impl IntoIterator<Item = String>,
    max_part_bytes: u64,
    budget: &mut ZipInflateBudget,
) -> Result<BTreeMap<String, Vec<u8>>, ChartExtractionError> {
    use std::collections::VecDeque;

    let mut out: BTreeMap<String, Vec<u8>> = BTreeMap::new();
    let mut visited: HashSet<String> = HashSet::new();
    let mut queue: VecDeque<String> = root_parts.into_iter().collect();

    fn strip_fragment(target: &str) -> &str {
        target
            .split_once('#')
            .map(|(base, _)| base)
            .unwrap_or(target)
    }

    while let Some(part_name) = queue.pop_front() {
        if !visited.insert(part_name.clone()) {
            continue;
        }

        if !part_names.contains(&part_name) {
            continue;
        }
        let Some(part_bytes) = read_zip_part_optional(archive, &part_name, max_part_bytes, budget)?
        else {
            continue;
        };
        out.insert(part_name.clone(), part_bytes);

        let rels_part_name = rels_for_part(&part_name);
        let Some(rels_part_name) = find_part_name(part_names, &rels_part_name) else {
            continue;
        };
        let Some(rels_bytes) =
            read_zip_part_optional(archive, &rels_part_name, max_part_bytes, budget)?
        else {
            continue;
        };
        out.insert(rels_part_name.clone(), rels_bytes.clone());

        let relationships = match crate::openxml::parse_relationships(&rels_bytes) {
            Ok(rels) => rels,
            Err(_) => continue,
        };

        for rel in relationships {
            if rel
                .target_mode
                .as_deref()
                .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
            {
                continue;
            }

            // Match the non-streaming traversal behavior: strip URI fragments, tolerate invalid
            // Windows-style path separators, and ignore leading `./`.
            let target = strip_fragment(&rel.target);
            if target.is_empty() {
                continue;
            }
            let target: std::borrow::Cow<'_, str> = if target.contains('\\') {
                std::borrow::Cow::Owned(target.replace('\\', "/"))
            } else {
                std::borrow::Cow::Borrowed(target)
            };
            let target = target.as_ref();
            let target = target.strip_prefix("./").unwrap_or(target);

            let target_part = resolve_target(&part_name, target);
            if let Some(found) = find_part_name(part_names, &target_part) {
                queue.push_back(found);
            }
        }
    }

    Ok(out)
}

fn extract_workbook_chart_sheets_from_workbook_parts(
    workbook_xml: &[u8],
    workbook_rels_xml: Option<&[u8]>,
    part_names: &HashSet<String>,
) -> Result<BTreeMap<String, PreservedChartSheet>, ChartExtractionError> {
    let workbook_part = "xl/workbook.xml";
    let workbook_xml = std::str::from_utf8(workbook_xml)
        .map_err(|e| ChartExtractionError::XmlNonUtf8(workbook_part.to_string(), e))?;
    let workbook_doc = Document::parse(workbook_xml)
        .map_err(|e| ChartExtractionError::XmlParse(workbook_part.to_string(), e))?;

    let workbook_rels_part = "xl/_rels/workbook.xml.rels";
    let rel_map: HashMap<String, crate::relationships::Relationship> = match workbook_rels_xml {
        Some(workbook_rels_xml) => match parse_relationships(workbook_rels_xml, workbook_rels_part)
        {
            Ok(rels) => rels.into_iter().map(|r| (r.id.clone(), r)).collect(),
            Err(_) => HashMap::new(),
        },
        None => HashMap::new(),
    };

    let mut out = BTreeMap::new();
    for (index, sheet_node) in workbook_doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "sheet")
        .enumerate()
    {
        let Some(name) = sheet_node.attribute("name") else {
            continue;
        };
        let sheet_id = sheet_node
            .attribute("sheetId")
            .and_then(|v| v.parse::<u32>().ok());
        let Some(rel_id) = sheet_node
            .attribute((REL_NS, "id"))
            .or_else(|| sheet_node.attribute("r:id"))
            .or_else(|| sheet_node.attribute("id"))
        else {
            continue;
        };
        let state = sheet_node.attribute("state").map(|s| s.to_string());

        let Some(rel) = rel_map.get(rel_id) else {
            continue;
        };
        if rel
            .target_mode
            .as_deref()
            .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
        {
            continue;
        }
        if rel.type_ != CHARTSHEET_REL_TYPE {
            continue;
        }

        let resolved_target = resolve_target(workbook_part, &rel.target);
        if !resolved_target.starts_with("xl/chartsheets/") {
            continue;
        }
        let Some(resolved_target) = find_part_name(part_names, &resolved_target) else {
            continue;
        };

        out.insert(
            name.to_string(),
            PreservedChartSheet {
                sheet_index: index,
                sheet_id,
                rel_id: rel_id.to_string(),
                rel_target: rel.target.clone(),
                state,
                part_name: resolved_target,
            },
        );
    }

    Ok(out)
}

fn extract_workbook_chart_sheets(
    pkg: &XlsxPackage,
) -> Result<BTreeMap<String, PreservedChartSheet>, ChartExtractionError> {
    let workbook_part = "xl/workbook.xml";
    let workbook_xml = pkg
        .part(workbook_part)
        .ok_or_else(|| ChartExtractionError::MissingPart(workbook_part.to_string()))?;
    let workbook_xml = std::str::from_utf8(workbook_xml)
        .map_err(|e| ChartExtractionError::XmlNonUtf8(workbook_part.to_string(), e))?;
    let workbook_doc = Document::parse(workbook_xml)
        .map_err(|e| ChartExtractionError::XmlParse(workbook_part.to_string(), e))?;

    let workbook_rels_part = "xl/_rels/workbook.xml.rels";
    let rel_map: HashMap<String, crate::relationships::Relationship> = match pkg.part(workbook_rels_part) {
        Some(workbook_rels_xml) => match parse_relationships(workbook_rels_xml, workbook_rels_part) {
            Ok(rels) => rels.into_iter().map(|r| (r.id.clone(), r)).collect(),
            Err(_) => HashMap::new(),
        },
        None => HashMap::new(),
    };

    let mut out = BTreeMap::new();
    for (index, sheet_node) in workbook_doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "sheet")
        .enumerate()
    {
        let Some(name) = sheet_node.attribute("name") else {
            continue;
        };
        let sheet_id = sheet_node
            .attribute("sheetId")
            .and_then(|v| v.parse::<u32>().ok());
        let Some(rel_id) = sheet_node
            .attribute((REL_NS, "id"))
            .or_else(|| sheet_node.attribute("r:id"))
            .or_else(|| sheet_node.attribute("id"))
        else {
            continue;
        };
        let state = sheet_node.attribute("state").map(|s| s.to_string());

        let Some(rel) = rel_map.get(rel_id) else {
            continue;
        };
        if rel
            .target_mode
            .as_deref()
            .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
        {
            continue;
        }
        if rel.type_ != CHARTSHEET_REL_TYPE {
            continue;
        }

        let resolved_target = resolve_target(workbook_part, &rel.target);
        if !resolved_target.starts_with("xl/chartsheets/") || pkg.part(&resolved_target).is_none() {
            continue;
        }

        out.insert(
            name.to_string(),
            PreservedChartSheet {
                sheet_index: index,
                sheet_id,
                rel_id: rel_id.to_string(),
                rel_target: rel.target.clone(),
                state,
                part_name: resolved_target,
            },
        );
    }

    Ok(out)
}

fn is_drawing_adjacent_relationship(rel_type: &str, resolved_target: &str) -> bool {
    if rel_type == DRAWING_REL_TYPE {
        return true;
    }

    const PREFIXES: &[&str] = &[
        "xl/drawings/",
        "xl/charts/",
        "xl/media/",
        "xl/embeddings/",
        "xl/ctrlProps/",
        "xl/activeX/",
        "xl/diagrams/",
    ];

    PREFIXES
        .iter()
        .any(|prefix| resolved_target.starts_with(prefix))
}

fn xml_contains_local_element(xml: &str, local: &str) -> bool {
    let bytes = xml.as_bytes();
    let local = local.as_bytes();
    let mut i = 0usize;

    while i < bytes.len() {
        if bytes[i] != b'<' {
            i += 1;
            continue;
        }
        i += 1;
        if i >= bytes.len() {
            break;
        }

        // Skip processing instructions (`<? ... ?>`) and directives/comments (`<! ... >`).
        if bytes[i] == b'?' || bytes[i] == b'!' {
            continue;
        }

        let mut j = i;
        if bytes[j] == b'/' {
            j += 1;
        }
        if j >= bytes.len() {
            break;
        }
        if bytes[j] == b'?' || bytes[j] == b'!' {
            continue;
        }

        let name_start = j;
        while j < bytes.len() {
            match bytes[j] {
                b' ' | b'\n' | b'\t' | b'\r' | b'>' | b'/' => break,
                _ => j += 1,
            }
        }
        if name_start == j {
            continue;
        }

        let name = &bytes[name_start..j];
        let local_name = name.rsplit(|b| *b == b':').next().unwrap_or(name);
        if local_name == local {
            return true;
        }

        i = j;
    }

    false
}

fn extract_relationship_ids(node: roxmltree::Node<'_, '_>) -> Vec<String> {
    let mut out = Vec::new();
    let mut seen = HashSet::new();

    for element in node.descendants().filter(|n| n.is_element()) {
        for attr in element.attributes() {
            // Common case: `r:id`, `r:embed`, etc.
            if attr.namespace() == Some(REL_NS) {
                let value = attr.value().to_string();
                if seen.insert(value.clone()) {
                    out.push(value);
                }
                continue;
            }

            // Fallback for producers that emit unprefixed `id="rId*"` attributes.
            if attr.name() == "id" && attr.value().starts_with("rId") {
                let value = attr.value().to_string();
                if seen.insert(value.clone()) {
                    out.push(value);
                }
            }
        }
    }

    out
}

fn ensure_sheet_xml_has_drawings(
    sheet_xml: &[u8],
    part_name: &str,
    drawings: &[SheetDrawingRelationship],
    rid_map: &HashMap<String, String>,
) -> Result<Vec<u8>, ChartExtractionError> {
    if drawings.is_empty() {
        return Ok(sheet_xml.to_vec());
    }

    let xml_str = std::str::from_utf8(sheet_xml)
        .map_err(|e| ChartExtractionError::XmlNonUtf8(part_name.to_string(), e))?;
    let doc = Document::parse(xml_str)
        .map_err(|e| ChartExtractionError::XmlParse(part_name.to_string(), e))?;
    let root = doc.root_element();
    let root_name = root.tag_name().name();
    let root_start = root.range().start;
    let root_prefix = element_prefix_at(xml_str, root_start);
    if root_name != "worksheet" && root_name != "chartsheet" {
        return Err(ChartExtractionError::XmlStructure(format!(
            "{part_name}: expected <worksheet> or <chartsheet>, found <{root_name}>"
        )));
    }

    let root_qname = crate::xml::prefixed_tag(root_prefix, root_name);
    let close_tag = format!("</{root_qname}>");
    let close_idx = xml_str.rfind(&close_tag).ok_or_else(|| {
        ChartExtractionError::XmlStructure(format!("{part_name}: missing {close_tag}"))
    })?;
    let insert_idx = root
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "extLst")
        .map(|n| n.range().start)
        .unwrap_or(close_idx);

    fn is_drawing_node(node: Node<'_, '_>) -> bool {
        node.is_element() && node.tag_name().name() == "drawing"
    }

    let existing: HashSet<String> = crate::drawingml::anchor::descendants_selecting_alternate_content(
        root,
        is_drawing_node,
        is_drawing_node,
    )
    .into_iter()
    .filter_map(|n| {
        n.attribute((REL_NS, "id"))
            .or_else(|| n.attribute("r:id"))
            .or_else(|| n.attribute("id"))
    })
    .map(|s| s.to_string())
    .collect();

    let drawing_tag = crate::xml::prefixed_tag(root_prefix, "drawing");
    let mut inserted: HashSet<String> = HashSet::new();
    let mut to_insert = String::new();
    for drawing in drawings {
        let desired_id = rid_map
            .get(&drawing.rel_id)
            .map(String::as_str)
            .unwrap_or(drawing.rel_id.as_str());
        if existing.contains(desired_id) || !inserted.insert(desired_id.to_string()) {
            continue;
        }
        to_insert.push_str(&format!("<{drawing_tag} r:id=\"{}\"/>", desired_id));
    }

    if to_insert.is_empty() {
        return Ok(sheet_xml.to_vec());
    }

    let mut xml = xml_str.to_string();
    xml.insert_str(insert_idx, &to_insert);

    if !root_start_has_r_namespace(&xml, root_start, part_name)? {
        xml = add_r_namespace_to_root(&xml, root_start, part_name)?;
    }
    Ok(xml.into_bytes())
}

fn ensure_sheet_xml_has_picture(
    sheet_xml: &[u8],
    part_name: &str,
    picture_xml: &[u8],
    rid_map: &HashMap<String, String>,
) -> Result<Vec<u8>, ChartExtractionError> {
    if picture_xml.is_empty() {
        return Ok(sheet_xml.to_vec());
    }

    let picture_str = std::str::from_utf8(picture_xml)
        .map_err(|e| ChartExtractionError::XmlNonUtf8("picture".to_string(), e))?;
    let mut picture_str = remap_relationship_ids(picture_str, rid_map);

    let xml_str = std::str::from_utf8(sheet_xml)
        .map_err(|e| ChartExtractionError::XmlNonUtf8(part_name.to_string(), e))?;

    let doc = Document::parse(xml_str)
        .map_err(|e| ChartExtractionError::XmlParse(part_name.to_string(), e))?;
    let root = doc.root_element();
    let root_name = root.tag_name().name();
    let root_start = root.range().start;
    let root_prefix = element_prefix_at(xml_str, root_start);
    if root_name != "worksheet" {
        return Err(ChartExtractionError::XmlStructure(format!(
            "{part_name}: expected <worksheet>, found <{root_name}>"
        )));
    }

    // Ensure the inserted `<picture>` tag uses the worksheet's SpreadsheetML prefix style.
    // This matters for prefix-only worksheets (`<x:worksheet ...>`) where unprefixed tags are in
    // *no namespace*, and also for default-namespace worksheets where inserting a prefixed tag
    // (e.g. `<x:picture>`) would introduce an undeclared prefix.
    let picture_prefix = root_element_prefix(&picture_str)?;
    if picture_prefix.as_deref() != root_prefix {
        picture_str =
            rewrite_fragment_prefix(&picture_str, picture_prefix.as_deref(), root_prefix)?;
    }

    if root
        .descendants()
        .any(|n| n.is_element() && n.tag_name().name() == "picture")
    {
        return Ok(sheet_xml.to_vec());
    }

    let root_qname = crate::xml::prefixed_tag(root_prefix, root_name);
    let close_tag = format!("</{root_qname}>");
    let close_idx = xml_str.rfind(&close_tag).ok_or_else(|| {
        ChartExtractionError::XmlStructure(format!("{part_name}: missing {close_tag}"))
    })?;

    let sheet_data_end = root
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "sheetData")
        .map(|n| n.range().end)
        .unwrap_or(0);

    let insert_idx = root
        .children()
        .find(|n| {
            n.is_element()
                && matches!(
                    n.tag_name().name(),
                    "oleObjects" | "controls" | "webPublishItems" | "tableParts" | "extLst"
                )
        })
        .map(|n| n.range().start)
        .unwrap_or(close_idx)
        .max(sheet_data_end)
        .min(close_idx);

    let mut xml = xml_str.to_string();
    xml.insert_str(insert_idx, &picture_str);

    if picture_str.contains("r:") && !root_start_has_r_namespace(&xml, root_start, part_name)? {
        xml = add_r_namespace_to_root(&xml, root_start, part_name)?;
    }

    Ok(xml.into_bytes())
}

fn ensure_sheet_xml_has_ole_objects(
    sheet_xml: &[u8],
    part_name: &str,
    ole_objects_xml: &[u8],
    rid_map: &HashMap<String, String>,
) -> Result<Vec<u8>, ChartExtractionError> {
    if ole_objects_xml.is_empty() {
        return Ok(sheet_xml.to_vec());
    }

    let ole_str = std::str::from_utf8(ole_objects_xml)
        .map_err(|e| ChartExtractionError::XmlNonUtf8("oleObjects".to_string(), e))?;
    let mut ole_str = remap_relationship_ids(ole_str, rid_map);

    let xml_str = std::str::from_utf8(sheet_xml)
        .map_err(|e| ChartExtractionError::XmlNonUtf8(part_name.to_string(), e))?;

    let doc = Document::parse(xml_str)
        .map_err(|e| ChartExtractionError::XmlParse(part_name.to_string(), e))?;
    let root = doc.root_element();
    let root_name = root.tag_name().name();
    let root_start = root.range().start;
    let root_prefix = element_prefix_at(xml_str, root_start);
    if root_name != "worksheet" {
        return Err(ChartExtractionError::XmlStructure(format!(
            "{part_name}: expected <worksheet>, found <{root_name}>"
        )));
    }

    let ole_prefix = root_element_prefix(&ole_str)?;
    if ole_prefix.as_deref() != root_prefix {
        ole_str = rewrite_fragment_prefix(&ole_str, ole_prefix.as_deref(), root_prefix)?;
    }

    if root
        .descendants()
        .any(|n| n.is_element() && n.tag_name().name() == "oleObjects")
    {
        return Ok(sheet_xml.to_vec());
    }

    let root_qname = crate::xml::prefixed_tag(root_prefix, root_name);
    let close_tag = format!("</{root_qname}>");
    let close_idx = xml_str.rfind(&close_tag).ok_or_else(|| {
        ChartExtractionError::XmlStructure(format!("{part_name}: missing {close_tag}"))
    })?;

    let sheet_data_end = root
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "sheetData")
        .map(|n| n.range().end)
        .unwrap_or(0);

    let insert_idx = root
        .children()
        .find(|n| {
            n.is_element()
                && matches!(
                    n.tag_name().name(),
                    "controls" | "webPublishItems" | "tableParts" | "extLst"
                )
        })
        .map(|n| n.range().start)
        .unwrap_or(close_idx)
        .max(sheet_data_end)
        .min(close_idx);

    let mut xml = xml_str.to_string();
    xml.insert_str(insert_idx, &ole_str);

    if ole_str.contains("r:") && !root_start_has_r_namespace(&xml, root_start, part_name)? {
        xml = add_r_namespace_to_root(&xml, root_start, part_name)?;
    }

    Ok(xml.into_bytes())
}

fn ensure_sheet_xml_has_controls(
    sheet_xml: &[u8],
    part_name: &str,
    controls_xml: &[u8],
    rid_map: &HashMap<String, String>,
) -> Result<Vec<u8>, ChartExtractionError> {
    if controls_xml.is_empty() {
        return Ok(sheet_xml.to_vec());
    }

    let controls_str = std::str::from_utf8(controls_xml)
        .map_err(|e| ChartExtractionError::XmlNonUtf8("controls".to_string(), e))?;
    let mut controls_str = remap_relationship_ids(controls_str, rid_map);

    let xml_str = std::str::from_utf8(sheet_xml)
        .map_err(|e| ChartExtractionError::XmlNonUtf8(part_name.to_string(), e))?;

    let doc = Document::parse(xml_str)
        .map_err(|e| ChartExtractionError::XmlParse(part_name.to_string(), e))?;
    let root = doc.root_element();
    let root_name = root.tag_name().name();
    let root_start = root.range().start;
    let root_prefix = element_prefix_at(xml_str, root_start);
    if root_name != "worksheet" {
        return Err(ChartExtractionError::XmlStructure(format!(
            "{part_name}: expected <worksheet>, found <{root_name}>"
        )));
    }

    if root
        .descendants()
        .any(|n| n.is_element() && n.tag_name().name() == "controls")
    {
        return Ok(sheet_xml.to_vec());
    }

    let controls_prefix = root_element_prefix(&controls_str)?;
    if controls_prefix.as_deref() != root_prefix {
        controls_str =
            rewrite_fragment_prefix(&controls_str, controls_prefix.as_deref(), root_prefix)?;
    }

    let root_qname = crate::xml::prefixed_tag(root_prefix, root_name);
    let close_tag = format!("</{root_qname}>");
    let close_idx = xml_str.rfind(&close_tag).ok_or_else(|| {
        ChartExtractionError::XmlStructure(format!("{part_name}: missing {close_tag}"))
    })?;

    let sheet_data_end = root
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "sheetData")
        .map(|n| n.range().end)
        .unwrap_or(0);

    let insert_idx = root
        .children()
        .find(|n| {
            n.is_element()
                && matches!(
                    n.tag_name().name(),
                    "webPublishItems" | "tableParts" | "extLst"
                )
        })
        .map(|n| n.range().start)
        .unwrap_or(close_idx)
        .max(sheet_data_end)
        .min(close_idx);

    let mut xml = xml_str.to_string();
    xml.insert_str(insert_idx, &controls_str);

    if controls_str.contains("r:") && !root_start_has_r_namespace(&xml, root_start, part_name)? {
        xml = add_r_namespace_to_root(&xml, root_start, part_name)?;
    }

    Ok(xml.into_bytes())
}

fn ensure_sheet_xml_has_drawing_hf(
    sheet_xml: &[u8],
    part_name: &str,
    drawing_hf_xml: &[u8],
    rid_map: &HashMap<String, String>,
) -> Result<Vec<u8>, ChartExtractionError> {
    if drawing_hf_xml.is_empty() {
        return Ok(sheet_xml.to_vec());
    }

    let drawing_hf_str = std::str::from_utf8(drawing_hf_xml)
        .map_err(|e| ChartExtractionError::XmlNonUtf8("drawingHF".to_string(), e))?;
    let mut drawing_hf_str = remap_relationship_ids(drawing_hf_str, rid_map);

    let xml_str = std::str::from_utf8(sheet_xml)
        .map_err(|e| ChartExtractionError::XmlNonUtf8(part_name.to_string(), e))?;

    let doc = Document::parse(xml_str)
        .map_err(|e| ChartExtractionError::XmlParse(part_name.to_string(), e))?;
    let root = doc.root_element();
    let root_name = root.tag_name().name();
    let root_start = root.range().start;
    let root_prefix = element_prefix_at(xml_str, root_start);
    if root_name != "worksheet" {
        return Err(ChartExtractionError::XmlStructure(format!(
            "{part_name}: expected <worksheet>, found <{root_name}>"
        )));
    }

    if root
        .descendants()
        .any(|n| n.is_element() && n.tag_name().name() == "drawingHF")
    {
        return Ok(sheet_xml.to_vec());
    }

    let drawing_hf_prefix = root_element_prefix(&drawing_hf_str)?;
    if drawing_hf_prefix.as_deref() != root_prefix {
        drawing_hf_str =
            rewrite_fragment_prefix(&drawing_hf_str, drawing_hf_prefix.as_deref(), root_prefix)?;
    }

    let root_qname = crate::xml::prefixed_tag(root_prefix, root_name);
    let close_tag = format!("</{root_qname}>");
    let close_idx = xml_str.rfind(&close_tag).ok_or_else(|| {
        ChartExtractionError::XmlStructure(format!("{part_name}: missing {close_tag}"))
    })?;

    let sheet_data_end = root
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "sheetData")
        .map(|n| n.range().end)
        .unwrap_or(0);

    let insert_idx = root
        .children()
        .find(|n| {
            n.is_element()
                && matches!(
                    n.tag_name().name(),
                    "picture"
                        | "oleObjects"
                        | "controls"
                        | "webPublishItems"
                        | "tableParts"
                        | "extLst"
                )
        })
        .map(|n| n.range().start)
        .unwrap_or(close_idx)
        .max(sheet_data_end)
        .min(close_idx);

    let mut xml = xml_str.to_string();
    xml.insert_str(insert_idx, &drawing_hf_str);

    if drawing_hf_str.contains("r:") && !root_start_has_r_namespace(&xml, root_start, part_name)? {
        xml = add_r_namespace_to_root(&xml, root_start, part_name)?;
    }

    Ok(xml.into_bytes())
}

fn remap_relationship_ids(fragment: &str, rid_map: &HashMap<String, String>) -> String {
    if rid_map.is_empty() {
        return fragment.to_string();
    }

    let patterns: [(&str, char); 8] = [
        ("r:id=\"", '"'),
        ("r:id='", '\''),
        ("r:embed=\"", '"'),
        ("r:embed='", '\''),
        ("r:link=\"", '"'),
        ("r:link='", '\''),
        ("id=\"", '"'),
        ("id='", '\''),
    ];
    let mut out = String::new();
    let _ = out.try_reserve(fragment.len());
    let mut cursor = 0usize;

    while cursor < fragment.len() {
        let mut next_match: Option<(usize, &str, char)> = None;
        for (pat, quote) in patterns {
            if let Some(rel_pos) = fragment[cursor..].find(pat) {
                let abs_pos = cursor + rel_pos;
                if next_match.map_or(true, |(pos, _, _)| abs_pos < pos) {
                    next_match = Some((abs_pos, pat, quote));
                }
            }
        }

        let Some((pos, pat, quote)) = next_match else {
            out.push_str(&fragment[cursor..]);
            break;
        };

        let pat_end = pos + pat.len();
        out.push_str(&fragment[cursor..pat_end]);

        let value_start = pat_end;
        let Some(value_end_rel) = fragment[value_start..].find(quote) else {
            out.push_str(&fragment[value_start..]);
            break;
        };
        let value_end = value_start + value_end_rel;
        let value = &fragment[value_start..value_end];

        if let Some(mapped) = rid_map.get(value) {
            out.push_str(mapped);
        } else {
            out.push_str(value);
        }
        out.push(quote);

        cursor = value_end + 1;
    }

    out
}

fn root_element_prefix(fragment: &str) -> Result<Option<String>, ChartExtractionError> {
    let mut reader = XmlReader::from_str(fragment);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(|e| {
            ChartExtractionError::XmlStructure(format!("failed to parse XML fragment: {e}"))
        })? {
            Event::Start(e) | Event::Empty(e) => {
                let name = e.name();
                let name = name.as_ref();
                let prefix = name
                    .iter()
                    .rposition(|b| *b == b':')
                    .map(|idx| &name[..idx])
                    .and_then(|p| std::str::from_utf8(p).ok())
                    .map(|s| s.to_string());
                return Ok(prefix);
            }
            Event::Eof => return Ok(None),
            _ => {}
        }
        buf.clear();
    }
}

fn rewrite_fragment_prefix(
    fragment: &str,
    from_prefix: Option<&str>,
    to_prefix: Option<&str>,
) -> Result<String, ChartExtractionError> {
    if from_prefix == to_prefix {
        return Ok(fragment.to_string());
    }

    let mut reader = XmlReader::from_str(fragment);
    reader.config_mut().trim_text(false);
    let mut writer = XmlWriter::new(Vec::new());
    let mut buf = Vec::new();

    loop {
        let event = reader.read_event_into(&mut buf).map_err(|e| {
            ChartExtractionError::XmlStructure(format!("failed to parse XML fragment: {e}"))
        })?;

        match event {
            Event::Eof => break,
            Event::Start(e) => {
                let tag = rewrite_tag_name(e.name().as_ref(), from_prefix, to_prefix);
                let mut out = BytesStart::new(tag.as_str()).into_owned();
                for attr in e.attributes().with_checks(false) {
                    let attr = attr.map_err(|e| {
                        ChartExtractionError::XmlStructure(format!(
                            "failed to parse XML fragment attribute: {e}"
                        ))
                    })?;
                    out.push_attribute((attr.key.as_ref(), attr.value.as_ref()));
                }
                writer.write_event(Event::Start(out)).map_err(|e| {
                    ChartExtractionError::XmlStructure(format!("xml write error: {e}"))
                })?;
            }
            Event::Empty(e) => {
                let tag = rewrite_tag_name(e.name().as_ref(), from_prefix, to_prefix);
                let mut out = BytesStart::new(tag.as_str()).into_owned();
                for attr in e.attributes().with_checks(false) {
                    let attr = attr.map_err(|e| {
                        ChartExtractionError::XmlStructure(format!(
                            "failed to parse XML fragment attribute: {e}"
                        ))
                    })?;
                    out.push_attribute((attr.key.as_ref(), attr.value.as_ref()));
                }
                writer.write_event(Event::Empty(out)).map_err(|e| {
                    ChartExtractionError::XmlStructure(format!("xml write error: {e}"))
                })?;
            }
            Event::End(e) => {
                let tag = rewrite_tag_name(e.name().as_ref(), from_prefix, to_prefix);
                writer
                    .write_event(Event::End(BytesEnd::new(tag.as_str())))
                    .map_err(|e| {
                        ChartExtractionError::XmlStructure(format!("xml write error: {e}"))
                    })?;
            }
            other => {
                writer.write_event(other.to_owned()).map_err(|e| {
                    ChartExtractionError::XmlStructure(format!("xml write error: {e}"))
                })?;
            }
        }

        buf.clear();
    }

    String::from_utf8(writer.into_inner()).map_err(|e| {
        ChartExtractionError::XmlStructure(format!("xml write produced invalid UTF-8: {e}"))
    })
}

fn rewrite_tag_name(name: &[u8], from_prefix: Option<&str>, to_prefix: Option<&str>) -> String {
    let name_str = std::str::from_utf8(name).unwrap_or_default();
    let (prefix, local) = match name_str.split_once(':') {
        Some((p, local)) => (Some(p), local),
        None => (None, name_str),
    };
    let matches_from = match from_prefix {
        Some(from) => prefix == Some(from),
        None => prefix.is_none(),
    };

    if matches_from {
        crate::xml::prefixed_tag(to_prefix, local)
    } else {
        name_str.to_string()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    use std::io::{Cursor, Write};

    use zip::write::FileOptions;
    use zip::ZipWriter;

    fn build_zip_bytes(entries: &[(&str, &[u8])]) -> Vec<u8> {
        let cursor = Cursor::new(Vec::new());
        let mut zip = ZipWriter::new(cursor);
        let options =
            FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

        for (name, bytes) in entries {
            zip.start_file(*name, options).unwrap();
            zip.write_all(bytes).unwrap();
        }

        zip.finish().unwrap().into_inner()
    }

    fn build_package(entries: &[(&str, &[u8])]) -> XlsxPackage {
        let bytes = build_zip_bytes(entries);
        XlsxPackage::from_bytes(&bytes).expect("read test pkg")
    }

    #[test]
    fn collect_transitive_related_parts_strips_relationship_fragments() {
        let rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="chart1.xml#something"/>
</Relationships>"#;

        let bytes = build_zip_bytes(&[
            ("xl/charts/chart0.xml", br#"<c:chartSpace/>"#),
            ("xl/charts/_rels/chart0.xml.rels", rels),
            ("xl/charts/chart1.xml", br#"<c:chartSpace/>"#),
        ]);

        let cursor = Cursor::new(bytes);
        let mut archive = zip::ZipArchive::new(cursor).expect("open zip");
        let mut part_names: HashSet<String> = HashSet::new();
        for i in 0..archive.len() {
            let file = archive.by_index(i).expect("zip entry");
            if file.is_dir() {
                continue;
            }
            let name = file.name();
            part_names.insert(name.strip_prefix('/').unwrap_or(name).to_string());
        }

        let mut budget = ZipInflateBudget::new(u64::MAX);
        let preserved = collect_transitive_related_parts_from_archive(
            &mut archive,
            &part_names,
            ["xl/charts/chart0.xml".to_string()],
            u64::MAX,
            &mut budget,
        )
        .expect("traverse");

        assert!(
            preserved.contains_key("xl/charts/chart1.xml"),
            "expected fragment-stripped relationship to pull in chart1.xml, got keys: {:?}",
            preserved.keys().collect::<Vec<_>>()
        );
    }

    #[test]
    fn collect_transitive_related_parts_normalizes_relationship_backslashes() {
        let rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="..\media\image1.png"/>
</Relationships>"#;

        let bytes = build_zip_bytes(&[
            ("xl/drawings/drawing1.xml", br#"<xdr:wsDr/>"#),
            ("xl/drawings/_rels/drawing1.xml.rels", rels),
            ("xl/media/image1.png", b"png-bytes"),
        ]);

        let cursor = Cursor::new(bytes);
        let mut archive = zip::ZipArchive::new(cursor).expect("open zip");
        let mut part_names: HashSet<String> = HashSet::new();
        for i in 0..archive.len() {
            let file = archive.by_index(i).expect("zip entry");
            if file.is_dir() {
                continue;
            }
            let name = file.name();
            part_names.insert(name.strip_prefix('/').unwrap_or(name).to_string());
        }

        let mut budget = ZipInflateBudget::new(u64::MAX);
        let preserved = collect_transitive_related_parts_from_archive(
            &mut archive,
            &part_names,
            ["xl/drawings/drawing1.xml".to_string()],
            u64::MAX,
            &mut budget,
        )
        .expect("traverse");

        assert!(
            preserved.contains_key("xl/media/image1.png"),
            "expected backslash-normalized relationship to pull in media part, got keys: {:?}",
            preserved.keys().collect::<Vec<_>>()
        );
    }

    #[test]
    fn preserve_drawing_parts_tolerates_malformed_sheet_rels() {
        let content_types = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"#;

        let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

        let sheet_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData/>
  <drawing r:id="rId1"/>
</worksheet>"#;

        let pkg = build_package(&[
            ("[Content_Types].xml", content_types),
            ("xl/workbook.xml", workbook_xml),
            ("xl/worksheets/sheet1.xml", sheet_xml),
            ("xl/worksheets/_rels/sheet1.xml.rels", br#"<Relationships><Relationship"#),
        ]);

        let preserved = pkg.preserve_drawing_parts().expect("best-effort preserve");
        assert!(
            preserved.is_empty(),
            "malformed sheet rels should result in skipping drawing preservation"
        );
    }

    #[test]
    fn preserve_drawing_parts_ignores_external_relationships() {
        let content_types = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"#;

        let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

        let sheet_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData/>
  <drawing r:id="rId1"/>
</worksheet>"#;

        // The target happens to look like a valid package part, but TargetMode=External means it
        // should not be traversed/preserved.
        let sheet_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="/xl/drawings/drawing1.xml" TargetMode="External"/>
</Relationships>"#;

        let pkg = build_package(&[
            ("[Content_Types].xml", content_types),
            ("xl/workbook.xml", workbook_xml),
            ("xl/worksheets/sheet1.xml", sheet_xml),
            ("xl/worksheets/_rels/sheet1.xml.rels", sheet_rels),
            ("xl/drawings/drawing1.xml", br#"<xdr:wsDr/>"#),
        ]);

        let preserved = pkg.preserve_drawing_parts().expect("best-effort preserve");
        assert!(
            preserved.is_empty(),
            "TargetMode=External relationships should not be preserved"
        );
    }

    #[test]
    fn inserts_drawing_before_ext_lst() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/><extLst><ext/></extLst></worksheet>"#;
        let drawings = [SheetDrawingRelationship {
            rel_id: "rId1".to_string(),
            target: "drawings/drawing1.xml".to_string(),
        }];
        let updated = ensure_sheet_xml_has_drawings(
            xml,
            "xl/worksheets/sheet1.xml",
            &drawings,
            &HashMap::new(),
        )
        .expect("patch sheet xml");
        let updated_str = std::str::from_utf8(&updated).unwrap();

        let drawing_pos = updated_str.find("<drawing").unwrap();
        let ext_pos = updated_str.find("<extLst").unwrap();
        assert!(
            drawing_pos < ext_pos,
            "drawing should be inserted before extLst"
        );
    }

    #[test]
    fn inserts_drawing_before_ext_lst_in_chartsheet() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><extLst/></chartsheet>"#;
        let drawings = [SheetDrawingRelationship {
            rel_id: "rId1".to_string(),
            target: "drawings/drawing1.xml".to_string(),
        }];
        let updated = ensure_sheet_xml_has_drawings(
            xml,
            "xl/chartsheets/sheet1.xml",
            &drawings,
            &HashMap::new(),
        )
        .expect("patch chartsheet xml");
        let updated_str = std::str::from_utf8(&updated).unwrap();

        let drawing_pos = updated_str.find("<drawing").unwrap();
        let ext_pos = updated_str.find("<extLst").unwrap();
        assert!(
            drawing_pos < ext_pos,
            "drawing should be inserted before extLst"
        );
    }

    #[test]
    fn inserts_prefixed_drawing_when_worksheet_uses_prefix_only_namespaces() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <x:sheetData/>
  <x:extLst><x:ext/></x:extLst>
</x:worksheet>"#;
        let drawings = [SheetDrawingRelationship {
            rel_id: "rId1".to_string(),
            target: "drawings/drawing1.xml".to_string(),
        }];
        let updated = ensure_sheet_xml_has_drawings(
            xml,
            "xl/worksheets/sheet1.xml",
            &drawings,
            &HashMap::new(),
        )
        .expect("patch sheet xml");
        let updated_str = std::str::from_utf8(&updated).unwrap();

        roxmltree::Document::parse(updated_str).expect("output XML should be well-formed");
        assert!(
            updated_str.contains("<x:drawing r:id=\"rId1\"/>"),
            "expected inserted <x:drawing>, got:\n{updated_str}"
        );
        let drawing_pos = updated_str.find("<x:drawing").unwrap();
        let ext_pos = updated_str.find("<x:extLst").unwrap();
        assert!(
            drawing_pos < ext_pos,
            "drawing should be inserted before extLst, got:\n{updated_str}"
        );
        assert!(
            updated_str.contains("xmlns:r="),
            "expected xmlns:r to be added to root when inserting r:id attributes, got:\n{updated_str}"
        );
        assert!(
            !updated_str.contains("<drawing"),
            "should not introduce unprefixed <drawing> tags in prefix-only worksheets, got:\n{updated_str}"
        );
    }

    #[test]
    fn inserts_controls_after_sheet_data_and_adds_r_namespace() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"/></sheetData><extLst><ext/></extLst></worksheet>"#;
        let controls = br#"<controls><control r:id="rId1"/></controls>"#;
        let updated = ensure_sheet_xml_has_controls(
            xml,
            "xl/worksheets/sheet1.xml",
            controls,
            &HashMap::new(),
        )
        .expect("insert controls");
        let updated_str = std::str::from_utf8(&updated).unwrap();

        let sheet_data_end = updated_str.find("</sheetData>").unwrap() + "</sheetData>".len();
        let controls_pos = updated_str.find("<controls").unwrap();
        let ext_pos = updated_str.find("<extLst").unwrap();
        assert!(
            controls_pos >= sheet_data_end && controls_pos < ext_pos,
            "controls should be inserted after sheetData and before extLst, got:\n{updated_str}"
        );
        assert!(
            updated_str.contains("xmlns:r="),
            "expected xmlns:r to be added when inserting r:* attributes, got:\n{updated_str}"
        );
    }

    #[test]
    fn inserts_controls_with_prefix_when_worksheet_is_prefix_only() {
        let sheet_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <x:sheetData/>
</x:worksheet>"#;
        let controls = br#"<controls><control r:id="rId1"/></controls>"#;

        let updated = ensure_sheet_xml_has_controls(
            sheet_xml,
            "xl/worksheets/sheet1.xml",
            controls,
            &HashMap::new(),
        )
        .expect("insert controls");
        let updated_str = std::str::from_utf8(&updated).unwrap();

        roxmltree::Document::parse(updated_str).expect("output XML should be well-formed");
        assert!(
            updated_str.contains("<x:controls>") && updated_str.contains("</x:controls>"),
            "expected inserted <x:controls> block, got:\n{updated_str}"
        );
        assert!(
            updated_str.contains("<x:control r:id=\"rId1\"/>"),
            "expected inserted <x:control>, got:\n{updated_str}"
        );
        assert!(
            !updated_str.contains("<controls"),
            "should not introduce unprefixed <controls> in prefix-only worksheets, got:\n{updated_str}"
        );
        assert!(
            updated_str.contains("xmlns:r="),
            "expected xmlns:r to be added to root when inserting r:* attributes, got:\n{updated_str}"
        );
    }

    #[test]
    fn inserts_controls_without_prefix_when_worksheet_uses_default_namespace() {
        let sheet_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#;
        let controls = br#"<x:controls><x:control r:id="rId1"/></x:controls>"#;

        let updated = ensure_sheet_xml_has_controls(
            sheet_xml,
            "xl/worksheets/sheet1.xml",
            controls,
            &HashMap::new(),
        )
        .expect("insert controls");
        let updated_str = std::str::from_utf8(&updated).unwrap();

        roxmltree::Document::parse(updated_str).expect("output XML should be well-formed");
        assert!(
            updated_str.contains("<controls>") && updated_str.contains("</controls>"),
            "expected inserted <controls> block, got:\n{updated_str}"
        );
        assert!(
            updated_str.contains("<control r:id=\"rId1\"/>"),
            "expected inserted <control>, got:\n{updated_str}"
        );
        assert!(
            !updated_str.contains("<x:controls"),
            "should not introduce prefixed <x:controls> tags in default-namespace worksheets, got:\n{updated_str}"
        );
        assert!(
            updated_str.contains("xmlns:r="),
            "expected xmlns:r to be added to root when inserting r:* attributes, got:\n{updated_str}"
        );
    }

    #[test]
    fn inserts_drawing_hf_before_picture_and_after_sheet_data() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="1"/></sheetData><picture r:id="rId9"/><extLst><ext/></extLst></worksheet>"#;
        let drawing_hf = br#"<drawingHF r:id="rId7"/>"#;
        let updated = ensure_sheet_xml_has_drawing_hf(
            xml,
            "xl/worksheets/sheet1.xml",
            drawing_hf,
            &HashMap::new(),
        )
        .expect("insert drawingHF");
        let updated_str = std::str::from_utf8(&updated).unwrap();

        let sheet_data_end = updated_str.find("</sheetData>").unwrap() + "</sheetData>".len();
        let drawing_hf_pos = updated_str.find("<drawingHF").unwrap();
        let picture_pos = updated_str.find("<picture").unwrap();
        assert!(
            drawing_hf_pos >= sheet_data_end && drawing_hf_pos < picture_pos,
            "drawingHF should be inserted after sheetData and before picture, got:\n{updated_str}"
        );
        assert!(
            updated_str.contains("xmlns:r="),
            "expected xmlns:r to be added when inserting r:* attributes, got:\n{updated_str}"
        );
    }

    #[test]
    fn inserts_drawing_hf_with_prefix_when_worksheet_is_prefix_only() {
        let sheet_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <x:sheetData/>
  <x:picture/>
</x:worksheet>"#;
        let drawing_hf = br#"<drawingHF r:id="rId7"/>"#;

        let updated = ensure_sheet_xml_has_drawing_hf(
            sheet_xml,
            "xl/worksheets/sheet1.xml",
            drawing_hf,
            &HashMap::new(),
        )
        .expect("insert drawingHF");
        let updated_str = std::str::from_utf8(&updated).unwrap();

        roxmltree::Document::parse(updated_str).expect("output XML should be well-formed");
        assert!(
            updated_str.contains("<x:drawingHF r:id=\"rId7\"/>"),
            "expected inserted <x:drawingHF>, got:\n{updated_str}"
        );
        assert!(
            !updated_str.contains("<drawingHF r:id"),
            "should not introduce unprefixed <drawingHF> in prefix-only worksheets, got:\n{updated_str}"
        );
        let sheet_data_end = updated_str.find("<x:sheetData").unwrap();
        let drawing_hf_pos = updated_str.find("<x:drawingHF").unwrap();
        let picture_pos = updated_str.find("<x:picture").unwrap();
        assert!(
            drawing_hf_pos > sheet_data_end && drawing_hf_pos < picture_pos,
            "drawingHF should be inserted after sheetData and before picture, got:\n{updated_str}"
        );
        assert!(
            updated_str.contains("xmlns:r="),
            "expected xmlns:r to be added to root when inserting r:* attributes, got:\n{updated_str}"
        );
    }

    #[test]
    fn inserts_drawing_hf_without_prefix_when_worksheet_uses_default_namespace() {
        let sheet_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
  <picture/>
</worksheet>"#;
        let drawing_hf = br#"<x:drawingHF r:id="rId7"/>"#;

        let updated = ensure_sheet_xml_has_drawing_hf(
            sheet_xml,
            "xl/worksheets/sheet1.xml",
            drawing_hf,
            &HashMap::new(),
        )
        .expect("insert drawingHF");
        let updated_str = std::str::from_utf8(&updated).unwrap();

        roxmltree::Document::parse(updated_str).expect("output XML should be well-formed");
        assert!(
            updated_str.contains("<drawingHF r:id=\"rId7\"/>"),
            "expected inserted <drawingHF> without a prefix, got:\n{updated_str}"
        );
        assert!(
            !updated_str.contains("<x:drawingHF"),
            "should not introduce prefixed <x:drawingHF> tags in default-namespace worksheets, got:\n{updated_str}"
        );
        let sheet_data_pos = updated_str.find("<sheetData").unwrap();
        let drawing_hf_pos = updated_str.find("<drawingHF").unwrap();
        let picture_pos = updated_str.find("<picture").unwrap();
        assert!(
            drawing_hf_pos > sheet_data_pos && drawing_hf_pos < picture_pos,
            "drawingHF should be inserted after sheetData and before picture, got:\n{updated_str}"
        );
        assert!(
            updated_str.contains("xmlns:r="),
            "expected xmlns:r to be added to root when inserting r:* attributes, got:\n{updated_str}"
        );
    }

    #[test]
    fn inserts_picture_with_prefix_when_worksheet_is_prefix_only() {
        let sheet_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <x:sheetData/>
</x:worksheet>"#;
        let picture_xml = br#"<picture r:id="rId1"/>"#;

        let updated = ensure_sheet_xml_has_picture(
            sheet_xml,
            "xl/worksheets/sheet1.xml",
            picture_xml,
            &HashMap::new(),
        )
        .expect("patch sheet xml");
        let updated_str = std::str::from_utf8(&updated).unwrap();

        roxmltree::Document::parse(updated_str).expect("output XML should be well-formed");
        assert!(
            updated_str.contains("<x:picture r:id=\"rId1\"/>"),
            "expected inserted <x:picture>, got:\n{updated_str}"
        );
        assert!(
            !updated_str.contains("<picture r:id"),
            "should not introduce unprefixed <picture> in prefix-only worksheets, got:\n{updated_str}"
        );
    }

    #[test]
    fn inserts_picture_without_prefix_when_worksheet_uses_default_namespace() {
        let sheet_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#;
        let picture_xml = br#"<x:picture r:id="rId1"/>"#;

        let updated = ensure_sheet_xml_has_picture(
            sheet_xml,
            "xl/worksheets/sheet1.xml",
            picture_xml,
            &HashMap::new(),
        )
        .expect("patch sheet xml");
        let updated_str = std::str::from_utf8(&updated).unwrap();

        roxmltree::Document::parse(updated_str).expect("output XML should be well-formed");
        assert!(
            updated_str.contains("<picture r:id=\"rId1\"/>"),
            "expected inserted <picture> without a prefix, got:\n{updated_str}"
        );
        assert!(
            !updated_str.contains("<x:picture"),
            "should not introduce prefixed <x:picture> tags in default-namespace worksheets, got:\n{updated_str}"
        );
        assert!(
            updated_str.contains("xmlns:r="),
            "expected xmlns:r to be added to root when inserting r:id attributes, got:\n{updated_str}"
        );
    }

    #[test]
    fn inserts_ole_objects_with_prefix_when_worksheet_is_prefix_only() {
        let sheet_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <x:sheetData/>
</x:worksheet>"#;
        let ole_xml = br#"<oleObjects><oleObject r:id="rId1"/></oleObjects>"#;

        let updated = ensure_sheet_xml_has_ole_objects(
            sheet_xml,
            "xl/worksheets/sheet1.xml",
            ole_xml,
            &HashMap::new(),
        )
        .expect("patch sheet xml");
        let updated_str = std::str::from_utf8(&updated).unwrap();

        roxmltree::Document::parse(updated_str).expect("output XML should be well-formed");
        assert!(
            updated_str.contains("<x:oleObjects>") && updated_str.contains("</x:oleObjects>"),
            "expected inserted <x:oleObjects> block, got:\n{updated_str}"
        );
        assert!(
            updated_str.contains("<x:oleObject r:id=\"rId1\"/>"),
            "expected inserted <x:oleObject>, got:\n{updated_str}"
        );
        assert!(
            !updated_str.contains("<oleObjects>"),
            "should not introduce unprefixed <oleObjects> in prefix-only worksheets, got:\n{updated_str}"
        );
    }

    #[test]
    fn inserts_ole_objects_without_prefix_when_worksheet_uses_default_namespace() {
        let sheet_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#;
        let ole_xml = br#"<x:oleObjects><x:oleObject r:id="rId1"/></x:oleObjects>"#;

        let updated = ensure_sheet_xml_has_ole_objects(
            sheet_xml,
            "xl/worksheets/sheet1.xml",
            ole_xml,
            &HashMap::new(),
        )
        .expect("patch sheet xml");
        let updated_str = std::str::from_utf8(&updated).unwrap();

        roxmltree::Document::parse(updated_str).expect("output XML should be well-formed");
        assert!(
            updated_str.contains("<oleObjects>") && updated_str.contains("</oleObjects>"),
            "expected inserted <oleObjects> block, got:\n{updated_str}"
        );
        assert!(
            updated_str.contains("<oleObject r:id=\"rId1\"/>"),
            "expected inserted <oleObject>, got:\n{updated_str}"
        );
        assert!(
            !updated_str.contains("<x:oleObjects"),
            "should not introduce prefixed <x:oleObjects> tags in default-namespace worksheets, got:\n{updated_str}"
        );
        assert!(
            updated_str.contains("xmlns:r="),
            "expected xmlns:r to be added to root when inserting r:id attributes, got:\n{updated_str}"
        );
    }

    #[test]
    fn inserts_chartsheets_preserving_relationship_prefix() {
        let content_types = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
</Types>"#;

        let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <x:sheets>
    <x:sheet name="Sheet1" sheetId="1" rel:id="rId1"/>
  </x:sheets>
</x:workbook>"#;

        let mut pkg = build_package(&[
            ("[Content_Types].xml", content_types),
            ("xl/workbook.xml", workbook_xml),
        ]);

        let mut chart_sheets = BTreeMap::new();
        chart_sheets.insert(
            "Chart1".to_string(),
            PreservedChartSheet {
                sheet_index: 1,
                sheet_id: Some(2),
                rel_id: "rId2".to_string(),
                rel_target: "chartsheets/sheet1.xml".to_string(),
                state: None,
                part_name: "xl/chartsheets/sheet1.xml".to_string(),
            },
        );

        let preserved = PreservedDrawingParts {
            content_types_xml: content_types.to_vec(),
            parts: BTreeMap::new(),
            sheet_drawings: BTreeMap::new(),
            sheet_pictures: BTreeMap::new(),
            sheet_ole_objects: BTreeMap::new(),
            sheet_controls: BTreeMap::new(),
            sheet_drawing_hfs: BTreeMap::new(),
            chart_sheets,
        };

        pkg.apply_preserved_drawing_parts(&preserved)
            .expect("apply preserved parts");

        let updated = std::str::from_utf8(
            pkg.part("xl/workbook.xml")
                .expect("workbook.xml should exist after patch"),
        )
        .expect("workbook.xml should be utf-8");

        roxmltree::Document::parse(updated).expect("output workbook.xml should be well-formed");
        assert!(
            updated.contains("<x:sheet name=\"Chart1\" sheetId=\"2\" rel:id=\""),
            "expected inserted chartsheet entry to use rel:id, got:\n{updated}"
        );
        assert!(
            !updated.contains("r:id="),
            "should not introduce r:id when workbook already uses rel:id, got:\n{updated}"
        );
        assert!(
            !updated.contains("xmlns:r="),
            "should not introduce xmlns:r when workbook already declares rel namespace, got:\n{updated}"
        );
    }
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
    let start = element_start.checked_add(1)?;
    let rest = xml.get(start..)?;
    let end_rel = rest
        .char_indices()
        .find(|(_, c)| c.is_whitespace() || *c == '>' || *c == '/')
        .map(|(idx, _)| idx)
        .unwrap_or(rest.len());
    let qname = &rest[..end_rel];
    qname.split_once(':').map(|(p, _)| p)
}

fn workbook_relationship_id_prefix(
    workbook_xml: &str,
) -> Result<(String, bool), ChartExtractionError> {
    // Prefer detecting the relationships prefix from the workbook root `xmlns:*` declarations so we
    // preserve the workbook's prefix style (e.g. `rel:id` instead of forcing `r:id`).
    let mut reader = XmlReader::from_str(workbook_xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut rel_prefix_from_root: Option<String> = None;
    let mut root_xmlns: HashMap<String, String> = HashMap::new();

    loop {
        match reader.read_event_into(&mut buf).map_err(|e| {
            ChartExtractionError::XmlStructure(format!("failed to parse workbook.xml: {e}"))
        })? {
            Event::Start(e) | Event::Empty(e) => {
                for attr in e.attributes().with_checks(false) {
                    let attr = attr.map_err(|e| {
                        ChartExtractionError::XmlStructure(format!(
                            "failed to parse workbook.xml root attribute: {e}"
                        ))
                    })?;
                    let key = attr.key.as_ref();

                    if key == b"xmlns" {
                        // Default namespace declaration. Keep it for completeness even though it
                        // can't be used for namespaced attributes.
                        let value = attr
                            .unescape_value()
                            .map_err(|e| {
                                ChartExtractionError::XmlStructure(format!(
                                    "failed to parse workbook.xml root attribute value: {e}"
                                ))
                            })?
                            .into_owned();
                        root_xmlns.insert(String::new(), value);
                        continue;
                    }

                    let Some(prefix_bytes) = key.strip_prefix(b"xmlns:") else {
                        continue;
                    };
                    let Ok(prefix) = std::str::from_utf8(prefix_bytes) else {
                        continue;
                    };
                    let value = attr
                        .unescape_value()
                        .map_err(|e| {
                            ChartExtractionError::XmlStructure(format!(
                                "failed to parse workbook.xml root attribute value: {e}"
                            ))
                        })?
                        .into_owned();

                    // Record the first prefix bound to the relationships namespace.
                    if rel_prefix_from_root.is_none() && value == REL_NS {
                        rel_prefix_from_root = Some(prefix.to_string());
                    }
                    root_xmlns.insert(prefix.to_string(), value);
                }
                break;
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    if let Some(prefix) = rel_prefix_from_root {
        return Ok((prefix, true));
    }

    // Fallback: scan existing sheet attributes to infer the relationship prefix (`*:id`) used in
    // the workbook when it isn't declared at the root.
    let mut reader = XmlReader::from_str(workbook_xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut inferred: Option<String> = None;
    loop {
        match reader.read_event_into(&mut buf).map_err(|e| {
            ChartExtractionError::XmlStructure(format!("failed to parse workbook.xml: {e}"))
        })? {
            Event::Start(e) | Event::Empty(e)
                if crate::openxml::local_name(e.name().as_ref()).eq_ignore_ascii_case(b"sheet") =>
            {
                for attr in e.attributes().with_checks(false) {
                    let attr = attr.map_err(|e| {
                        ChartExtractionError::XmlStructure(format!(
                            "failed to parse workbook.xml sheet attribute: {e}"
                        ))
                    })?;
                    let key = attr.key.as_ref();
                    if crate::openxml::local_name(key).eq_ignore_ascii_case(b"id") {
                        if let Some(idx) = key.iter().position(|b| *b == b':') {
                            if idx > 0 {
                                if let Ok(prefix) = std::str::from_utf8(&key[..idx]) {
                                    inferred = Some(prefix.to_string());
                                    break;
                                }
                            }
                        }
                    }
                }
                if inferred.is_some() {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    if let Some(prefix) = inferred {
        // Only use the inferred prefix if it is not already declared on the workbook root. If the
        // root already declares this prefix then it must be bound to a different URI (since we
        // didn't find any root mapping to `REL_NS`), and using it would make the inserted
        // `prefix:id` attribute point at the wrong namespace.
        if !root_xmlns.contains_key(&prefix) {
            return Ok((prefix, false));
        }
    }

    // Default to `r` (Excel's conventional relationship prefix). If `r` is already declared on the
    // root, pick a variant (`r1`, `r2`, ...) that is not declared so we can safely inject it.
    let mut prefix = "r".to_string();
    if root_xmlns.contains_key(&prefix) {
        let mut counter = 1u32;
        loop {
            let candidate = format!("r{counter}");
            if !root_xmlns.contains_key(&candidate) {
                prefix = candidate;
                break;
            }
            counter += 1;
        }
    }
    Ok((prefix, false))
}

fn root_start_has_namespace_prefix(
    xml: &str,
    root_start: usize,
    prefix: &str,
    part_name: &str,
) -> Result<bool, ChartExtractionError> {
    let tag_end_rel = xml[root_start..].find('>').ok_or_else(|| {
        ChartExtractionError::XmlStructure(format!("{part_name}: invalid root start tag"))
    })?;
    let tag_end = root_start + tag_end_rel;
    Ok(xml[root_start..=tag_end].contains(&format!("xmlns:{prefix}=")))
}

fn add_namespace_prefix_to_root(
    xml: &str,
    root_start: usize,
    prefix: &str,
    uri: &str,
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
    out.insert_str(insert_pos, &format!(" xmlns:{prefix}=\"{uri}\""));
    Ok(out)
}

fn ensure_workbook_has_chartsheets(
    pkg: &mut XlsxPackage,
    chart_sheets: &BTreeMap<String, PreservedChartSheet>,
) -> Result<(), ChartExtractionError> {
    if chart_sheets.is_empty() {
        return Ok(());
    }

    let workbook_part = "xl/workbook.xml";
    let workbook_xml = pkg
        .part(workbook_part)
        .ok_or_else(|| ChartExtractionError::MissingPart(workbook_part.to_string()))?
        .to_vec();
    let workbook_xml = std::str::from_utf8(&workbook_xml)
        .map_err(|e| ChartExtractionError::XmlNonUtf8(workbook_part.to_string(), e))?
        .to_string();

    let doc = Document::parse(&workbook_xml)
        .map_err(|e| ChartExtractionError::XmlParse(workbook_part.to_string(), e))?;
    let root_start = doc.root_element().range().start;
    let (rel_id_prefix, rel_prefix_declared_on_root) =
        workbook_relationship_id_prefix(&workbook_xml)?;
    let sheets_node = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "sheets")
        .ok_or_else(|| {
            ChartExtractionError::XmlStructure("workbook.xml missing <sheets>".to_string())
        })?;
    let sheets_prefix = element_prefix_at(&workbook_xml, sheets_node.range().start);
    let sheet_tag = crate::xml::prefixed_tag(sheets_prefix, "sheet");
    let sheets_close_tag = format!("</{}>", crate::xml::prefixed_tag(sheets_prefix, "sheets"));

    let mut existing_sheet_names: HashSet<String> = HashSet::new();
    let mut used_sheet_ids: HashSet<u32> = HashSet::new();
    let mut max_sheet_id = 0u32;
    for sheet in doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "sheet")
    {
        if let Some(name) = sheet.attribute("name") {
            existing_sheet_names.insert(name.to_string());
        }
        if let Some(sheet_id) = sheet
            .attribute("sheetId")
            .and_then(|v| v.parse::<u32>().ok())
        {
            used_sheet_ids.insert(sheet_id);
            max_sheet_id = max_sheet_id.max(sheet_id);
        }
    }

    let mut to_insert: Vec<(&String, &PreservedChartSheet)> = chart_sheets
        .iter()
        .filter(|(name, _)| !existing_sheet_names.contains(*name))
        .collect();
    if to_insert.is_empty() {
        return Ok(());
    }
    to_insert.sort_by_key(|(_, sheet)| sheet.sheet_index);

    let rels_part = "xl/_rels/workbook.xml.rels";
    let relationship_stubs: Vec<RelationshipStub> = to_insert
        .iter()
        .map(|(_, sheet)| RelationshipStub {
            rel_id: sheet.rel_id.clone(),
            target: sheet.rel_target.clone(),
        })
        .collect();
    let (updated_rels, rid_map) = ensure_rels_has_relationships(
        pkg.part(rels_part),
        rels_part,
        workbook_part,
        CHARTSHEET_REL_TYPE,
        &relationship_stubs,
    )?;
    pkg.set_part(rels_part, updated_rels);

    let rel_id_attr = format!("{rel_id_prefix}:id");
    let mut insertion = String::new();
    for (name, sheet) in to_insert {
        let desired_sheet_id = sheet.sheet_id.filter(|id| *id > 0);
        let sheet_id = match desired_sheet_id {
            Some(id) if !used_sheet_ids.contains(&id) => id,
            _ => {
                let mut candidate = max_sheet_id + 1;
                while used_sheet_ids.contains(&candidate) {
                    candidate += 1;
                }
                max_sheet_id = candidate;
                candidate
            }
        };
        used_sheet_ids.insert(sheet_id);

        let rel_id = rid_map
            .get(&sheet.rel_id)
            .cloned()
            .unwrap_or_else(|| sheet.rel_id.clone());

        insertion.push_str(&format!(
            "    <{sheet_tag} name=\"{}\" sheetId=\"{}\" {rel_id_attr}=\"{}\"",
            xml_escape(name),
            sheet_id,
            xml_escape(&rel_id)
        ));
        if let Some(state) = sheet.state.as_deref() {
            insertion.push_str(&format!(" state=\"{}\"", xml_escape(state)));
        }
        insertion.push_str("/>\n");
        existing_sheet_names.insert(name.to_string());
    }

    let sheets_end = workbook_xml.rfind(&sheets_close_tag).ok_or_else(|| {
        ChartExtractionError::XmlStructure(format!("workbook.xml missing {sheets_close_tag}"))
    })?;
    let mut workbook_xml_updated = workbook_xml.clone();
    workbook_xml_updated.insert_str(sheets_end, &insertion);

    if !rel_prefix_declared_on_root
        && !root_start_has_namespace_prefix(
            &workbook_xml_updated,
            root_start,
            &rel_id_prefix,
            workbook_part,
        )?
    {
        workbook_xml_updated = add_namespace_prefix_to_root(
            &workbook_xml_updated,
            root_start,
            &rel_id_prefix,
            REL_NS,
            workbook_part,
        )?;
    }

    pkg.set_part(workbook_part, workbook_xml_updated.into_bytes());
    Ok(())
}

fn xml_escape(input: &str) -> String {
    input
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}
