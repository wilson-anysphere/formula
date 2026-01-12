use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet};

use roxmltree::Document;

use crate::path::{rels_for_part, resolve_target};
use crate::preserve::opc_graph::collect_transitive_related_parts;
use crate::preserve::rels_merge::{ensure_rels_has_relationships, RelationshipStub};
use crate::preserve::sheet_match::{match_sheet_by_name_or_index, workbook_sheet_parts};
use crate::relationships::parse_relationships;
use crate::workbook::ChartExtractionError;
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

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PreservedChartSheet {
    pub sheet_index: usize,
    pub sheet_id: Option<u32>,
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
    pub chart_sheets: BTreeMap<String, PreservedChartSheet>,
}

impl PreservedDrawingParts {
    pub fn is_empty(&self) -> bool {
        self.parts.is_empty()
            && self.sheet_drawings.values().all(|v| v.drawings.is_empty())
            && self.sheet_pictures.is_empty()
            && self.sheet_ole_objects.is_empty()
            && self.chart_sheets.is_empty()
    }
}

impl XlsxPackage {
    /// Extract the DrawingML/chart-related parts of an XLSX package so they can
    /// be re-applied to another package later (e.g. after regenerating sheet XML).
    pub fn preserve_drawing_parts(&self) -> Result<PreservedDrawingParts, ChartExtractionError> {
        let content_types_xml = self
            .part("[Content_Types].xml")
            .ok_or_else(|| ChartExtractionError::MissingPart("[Content_Types].xml".to_string()))?
            .to_vec();
        let sheets = workbook_sheet_parts(self)?;
        let mut root_parts: BTreeSet<String> = BTreeSet::new();
        let mut sheet_drawings: BTreeMap<String, PreservedSheetDrawings> = BTreeMap::new();
        let mut sheet_pictures: BTreeMap<String, PreservedSheetPicture> = BTreeMap::new();
        let mut sheet_ole_objects: BTreeMap<String, PreservedSheetOleObjects> = BTreeMap::new();
        let mut chart_sheets: BTreeMap<String, PreservedChartSheet> = BTreeMap::new();

        for sheet in sheets {
            let sheet_rels_part = rels_for_part(&sheet.part_name);
            let rels = self
                .part(&sheet_rels_part)
                .map(|xml| parse_relationships(xml, &sheet_rels_part))
                .transpose()?
                .unwrap_or_default();

            for rel in &rels {
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
                chart_sheets.insert(
                    sheet.name.clone(),
                    PreservedChartSheet {
                        sheet_index: sheet.index,
                        sheet_id: sheet.sheet_id,
                        part_name: sheet.part_name.clone(),
                    },
                );
            }

            let sheet_xml_str = std::str::from_utf8(sheet_xml)
                .map_err(|e| ChartExtractionError::XmlNonUtf8(sheet.part_name.clone(), e))?;
            let doc = Document::parse(sheet_xml_str)
                .map_err(|e| ChartExtractionError::XmlParse(sheet.part_name.clone(), e))?;

            let drawing_rids: Vec<String> = doc
                .descendants()
                .filter(|n| n.is_element() && n.tag_name().name() == "drawing")
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

            if drawing_rids.is_empty() && picture.is_none() && ole_objects.is_none() {
                continue;
            }

            let rel_map: HashMap<_, _> = rels.into_iter().map(|r| (r.id.clone(), r)).collect();

            if !drawing_rids.is_empty() {
                let mut drawings = Vec::new();
                for rid in drawing_rids {
                    if let Some(rel) = rel_map.get(&rid) {
                        if rel.type_ == DRAWING_REL_TYPE {
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
                    if rel.type_ == IMAGE_REL_TYPE {
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
                            Some(rel) if rel.type_ == OLE_OBJECT_REL_TYPE => {
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
        }

        let parts = collect_transitive_related_parts(self, root_parts.into_iter())?;

        Ok(PreservedDrawingParts {
            content_types_xml,
            parts,
            sheet_drawings,
            sheet_pictures,
            sheet_ole_objects,
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
            let Some(sheet) = match_sheet_by_name_or_index(
                &sheets,
                sheet_name,
                preserved_sheet.sheet_index,
            ) else {
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
            if sheet_xml_str.contains("<picture") {
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
            if sheet_xml_str.contains("<oleObjects") {
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

        Ok(())
    }
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

    PREFIXES.iter().any(|prefix| resolved_target.starts_with(prefix))
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
    let doc =
        Document::parse(xml_str).map_err(|e| ChartExtractionError::XmlParse(part_name.to_string(), e))?;
    let root = doc.root_element();
    let root_name = root.tag_name().name();
    if root_name != "worksheet" && root_name != "chartsheet" {
        return Err(ChartExtractionError::XmlStructure(format!(
            "{part_name}: expected <worksheet> or <chartsheet>, found <{root_name}>"
        )));
    }

    let close_tag = format!("</{root_name}>");
    let close_idx = xml_str.rfind(&close_tag).ok_or_else(|| {
        ChartExtractionError::XmlStructure(format!("{part_name}: missing {close_tag}"))
    })?;
    let insert_idx = root
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "extLst")
        .map(|n| n.range().start)
        .unwrap_or(close_idx);

    let existing: HashSet<String> = root
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "drawing")
        .filter_map(|n| {
            n.attribute((REL_NS, "id"))
                .or_else(|| n.attribute("r:id"))
                .or_else(|| n.attribute("id"))
        })
        .map(|s| s.to_string())
        .collect();

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
        to_insert.push_str(&format!("<drawing r:id=\"{}\"/>", desired_id));
    }

    if to_insert.is_empty() {
        return Ok(sheet_xml.to_vec());
    }

    let mut xml = xml_str.to_string();
    xml.insert_str(insert_idx, &to_insert);

    if !root_start_has_r_namespace(&xml, root_name, part_name)? {
        xml = add_r_namespace_to_root(&xml, root_name, part_name)?;
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
    let picture_str = remap_relationship_ids(picture_str, rid_map);

    let xml_str = std::str::from_utf8(sheet_xml)
        .map_err(|e| ChartExtractionError::XmlNonUtf8(part_name.to_string(), e))?;

    if xml_str.contains("<picture") {
        return Ok(sheet_xml.to_vec());
    }

    let doc =
        Document::parse(xml_str).map_err(|e| ChartExtractionError::XmlParse(part_name.to_string(), e))?;
    let root = doc.root_element();
    let root_name = root.tag_name().name();
    if root_name != "worksheet" {
        return Err(ChartExtractionError::XmlStructure(format!(
            "{part_name}: expected <worksheet>, found <{root_name}>"
        )));
    }

    let close_tag = format!("</{root_name}>");
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

    if picture_str.contains("r:id") && !root_start_has_r_namespace(&xml, root_name, part_name)? {
        xml = add_r_namespace_to_root(&xml, root_name, part_name)?;
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
    let ole_str = remap_relationship_ids(ole_str, rid_map);

    let xml_str = std::str::from_utf8(sheet_xml)
        .map_err(|e| ChartExtractionError::XmlNonUtf8(part_name.to_string(), e))?;

    if xml_str.contains("<oleObjects") {
        return Ok(sheet_xml.to_vec());
    }

    let doc =
        Document::parse(xml_str).map_err(|e| ChartExtractionError::XmlParse(part_name.to_string(), e))?;
    let root = doc.root_element();
    let root_name = root.tag_name().name();
    if root_name != "worksheet" {
        return Err(ChartExtractionError::XmlStructure(format!(
            "{part_name}: expected <worksheet>, found <{root_name}>"
        )));
    }

    let close_tag = format!("</{root_name}>");
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

    if ole_str.contains("r:id") && !root_start_has_r_namespace(&xml, root_name, part_name)? {
        xml = add_r_namespace_to_root(&xml, root_name, part_name)?;
    }

    Ok(xml.into_bytes())
}

fn remap_relationship_ids(fragment: &str, rid_map: &HashMap<String, String>) -> String {
    if rid_map.is_empty() {
        return fragment.to_string();
    }

    let patterns: [(&str, char); 4] = [("r:id=\"", '"'), ("r:id='", '\''), ("id=\"", '"'), ("id='", '\'')];
    let mut out = String::with_capacity(fragment.len());
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

#[cfg(test)]
mod tests {
    use super::*;

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
        assert!(drawing_pos < ext_pos, "drawing should be inserted before extLst");
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
}

fn root_start_has_r_namespace(
    xml: &str,
    root_name: &str,
    part_name: &str,
) -> Result<bool, ChartExtractionError> {
    let root_start = xml.find(&format!("<{root_name}")).ok_or_else(|| {
        ChartExtractionError::XmlStructure(format!("{part_name}: missing <{root_name}>"))
    })?;
    let tag_end_rel = xml[root_start..].find('>').ok_or_else(|| {
        ChartExtractionError::XmlStructure(format!("{part_name}: invalid <{root_name}> start tag"))
    })?;
    let tag_end = root_start + tag_end_rel;
    Ok(xml[root_start..=tag_end].contains("xmlns:r="))
}

fn add_r_namespace_to_root(
    xml: &str,
    root_name: &str,
    part_name: &str,
) -> Result<String, ChartExtractionError> {
    let root_start = xml.find(&format!("<{root_name}")).ok_or_else(|| {
        ChartExtractionError::XmlStructure(format!("{part_name}: missing <{root_name}>"))
    })?;
    let tag_end_rel = xml[root_start..].find('>').ok_or_else(|| {
        ChartExtractionError::XmlStructure(format!("{part_name}: invalid <{root_name}> start tag"))
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
        .ok_or_else(|| ChartExtractionError::MissingPart(workbook_part.to_string()))?;
    let workbook_xml = std::str::from_utf8(workbook_xml)
        .map_err(|e| ChartExtractionError::XmlNonUtf8(workbook_part.to_string(), e))?
        .to_string();

    let doc = Document::parse(&workbook_xml)
        .map_err(|e| ChartExtractionError::XmlParse(workbook_part.to_string(), e))?;

    let mut existing_sheet_names: HashMap<String, ()> = HashMap::new();
    let mut max_sheet_id = 0u32;
    for sheet in doc.descendants().filter(|n| n.is_element() && n.tag_name().name() == "sheet") {
        if let Some(name) = sheet.attribute("name") {
            existing_sheet_names.insert(name.to_string(), ());
        }
        if let Some(sheet_id) = sheet.attribute("sheetId").and_then(|v| v.parse::<u32>().ok()) {
            max_sheet_id = max_sheet_id.max(sheet_id);
        }
    }

    let rels_part = "xl/_rels/workbook.xml.rels";
    let rels_xml_bytes = pkg
        .part(rels_part)
        .ok_or_else(|| ChartExtractionError::MissingPart(rels_part.to_string()))?;
    let mut rels_xml = std::str::from_utf8(rels_xml_bytes)
        .map_err(|e| ChartExtractionError::XmlNonUtf8(rels_part.to_string(), e))?
        .to_string();

    let mut workbook_xml_updated = workbook_xml;

    // Insert in original sheet order for determinism when multiple chart sheets are present.
    let mut chart_sheets: Vec<(&String, &PreservedChartSheet)> = chart_sheets.iter().collect();
    chart_sheets.sort_by_key(|(_, sheet)| sheet.sheet_index);

    for (name, sheet) in chart_sheets {
        if existing_sheet_names.contains_key(name) {
            continue;
        }

        let next_sheet_id = max_sheet_id + 1;
        max_sheet_id = next_sheet_id;

        let next_rid = next_relationship_id(&rels_xml);
        let rel_id = format!("rId{next_rid}");

        let target = sheet
            .part_name
            .strip_prefix("xl/")
            .unwrap_or(sheet.part_name.as_str());

        let rels_insert_idx = rels_xml.rfind("</Relationships>").ok_or_else(|| {
            ChartExtractionError::XmlStructure(format!("{rels_part}: missing </Relationships>"))
        })?;
        rels_xml.insert_str(
            rels_insert_idx,
            &format!(
                "  <Relationship Id=\"{}\" Type=\"{}\" Target=\"{}\"/>\n",
                rel_id,
                CHARTSHEET_REL_TYPE,
                xml_escape(target)
            ),
        );

        let sheets_end = workbook_xml_updated.rfind("</sheets>").ok_or_else(|| {
            ChartExtractionError::XmlStructure("workbook.xml missing </sheets>".to_string())
        })?;
        workbook_xml_updated.insert_str(
            sheets_end,
            &format!(
                "    <sheet name=\"{}\" sheetId=\"{}\" r:id=\"{}\"/>\n",
                xml_escape(name),
                next_sheet_id,
                xml_escape(&rel_id)
            ),
        );

        existing_sheet_names.insert(name.to_string(), ());
    }

    if workbook_xml_updated.contains("r:id") && !workbook_xml_updated.contains("xmlns:r=") {
        workbook_xml_updated = ensure_workbook_has_r_namespace(&workbook_xml_updated, workbook_part)?;
    }

    pkg.set_part(workbook_part, workbook_xml_updated.into_bytes());
    pkg.set_part(rels_part, rels_xml.into_bytes());
    Ok(())
}

fn next_relationship_id(xml: &str) -> u32 {
    let mut max_id = 0u32;
    let mut rest = xml;
    while let Some(idx) = rest.find("Id=\"rId") {
        let after = &rest[idx + "Id=\"rId".len()..];
        let mut digits = String::new();
        for ch in after.chars() {
            if ch.is_ascii_digit() {
                digits.push(ch);
            } else {
                break;
            }
        }
        if let Ok(n) = digits.parse::<u32>() {
            max_id = max_id.max(n);
        }
        rest = &after[digits.len()..];
    }
    max_id + 1
}

fn ensure_workbook_has_r_namespace(
    workbook_xml: &str,
    part_name: &str,
) -> Result<String, ChartExtractionError> {
    if workbook_xml.contains("xmlns:r=") {
        return Ok(workbook_xml.to_string());
    }

    let workbook_start = workbook_xml.find("<workbook").ok_or_else(|| {
        ChartExtractionError::XmlStructure(format!("{part_name}: missing <workbook"))
    })?;
    let tag_end_rel = workbook_xml[workbook_start..].find('>').ok_or_else(|| {
        ChartExtractionError::XmlStructure(format!("{part_name}: invalid <workbook> start tag"))
    })?;
    let insert_pos = workbook_start + tag_end_rel;

    let mut out = workbook_xml.to_string();
    out.insert_str(insert_pos, &format!(" xmlns:r=\"{REL_NS}\""));
    Ok(out)
}

fn xml_escape(input: &str) -> String {
    input
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}
