use std::collections::{BTreeMap, HashMap, HashSet};

use roxmltree::Document;

use crate::path::rels_for_part;
use crate::preserve::rels_merge::{ensure_rels_has_relationships, RelationshipStub};
use crate::preserve::sheet_match::{match_sheet_by_name_or_index, workbook_sheet_parts};
use crate::relationships::parse_relationships;
use crate::workbook::ChartExtractionError;
use crate::XlsxPackage;

const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const DRAWING_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";

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

/// A slice of an XLSX package that is required to preserve DrawingML objects
/// (including charts) across a "read -> write" pipeline that doesn't otherwise
/// round-trip the original package structure.
#[derive(Debug, Clone)]
pub struct PreservedDrawingParts {
    pub content_types_xml: Vec<u8>,
    pub parts: BTreeMap<String, Vec<u8>>,
    pub sheet_drawings: BTreeMap<String, PreservedSheetDrawings>,
}

impl PreservedDrawingParts {
    pub fn is_empty(&self) -> bool {
        self.parts.is_empty() && self.sheet_drawings.values().all(|v| v.drawings.is_empty())
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

        let mut parts = BTreeMap::new();
        for (name, bytes) in self.parts() {
            if name.starts_with("xl/drawings/")
                || name.starts_with("xl/charts/")
                || name.starts_with("xl/media/")
            {
                parts.insert(name.to_string(), bytes.to_vec());
            }
        }

        let sheets = workbook_sheet_parts(self)?;
        let mut sheet_drawings: BTreeMap<String, PreservedSheetDrawings> = BTreeMap::new();

        for sheet in sheets {
            let Some(sheet_xml) = self.part(&sheet.part_name) else {
                continue;
            };
            let drawing_rids = extract_sheet_drawing_rids(sheet_xml, &sheet.part_name)?;
            if drawing_rids.is_empty() {
                continue;
            }

            let sheet_rels_part = rels_for_part(&sheet.part_name);
            let Some(sheet_rels_xml) = self.part(&sheet_rels_part) else {
                continue;
            };
            let rels = parse_relationships(sheet_rels_xml, &sheet_rels_part)?;
            let rel_map: HashMap<_, _> = rels.into_iter().map(|r| (r.id.clone(), r)).collect();

            let mut drawings = Vec::new();
            for rid in drawing_rids {
                if let Some(rel) = rel_map.get(&rid) {
                    drawings.push(SheetDrawingRelationship {
                        rel_id: rid.clone(),
                        target: rel.target.clone(),
                    });
                }
            }

            if drawings.is_empty() {
                continue;
            }

            sheet_drawings.insert(
                sheet.name,
                PreservedSheetDrawings {
                    sheet_index: sheet.index,
                    sheet_id: sheet.sheet_id,
                    drawings,
                },
            );
        }

        Ok(PreservedDrawingParts {
            content_types_xml,
            parts,
            sheet_drawings,
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

        Ok(())
    }
}

fn extract_sheet_drawing_rids(
    sheet_xml: &[u8],
    part_name: &str,
) -> Result<Vec<String>, ChartExtractionError> {
    let xml =
        std::str::from_utf8(sheet_xml).map_err(|e| ChartExtractionError::XmlNonUtf8(part_name.to_string(), e))?;
    let doc =
        Document::parse(xml).map_err(|e| ChartExtractionError::XmlParse(part_name.to_string(), e))?;
    Ok(doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "drawing")
        .filter_map(|n| {
            n.attribute((REL_NS, "id"))
                .or_else(|| n.attribute("r:id"))
                .or_else(|| n.attribute("id"))
        })
        .map(|s| s.to_string())
        .collect())
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
