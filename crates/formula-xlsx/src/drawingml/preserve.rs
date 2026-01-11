use std::collections::{BTreeMap, HashMap};

use roxmltree::Document;

use crate::path::rels_for_part;
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
            let Some(sheet) = match_sheet_by_name_or_index(
                &sheets,
                sheet_name,
                preserved_sheet.sheet_index,
            ) else {
                continue;
            };

            let Some(sheet_xml) = self.part(&sheet.part_name) else {
                continue;
            };
            let updated_sheet_xml =
                ensure_sheet_xml_has_drawings(sheet_xml, &sheet.part_name, &preserved_sheet.drawings)?;
            self.set_part(sheet.part_name.clone(), updated_sheet_xml);

            let sheet_rels_part = rels_for_part(&sheet.part_name);
            let updated_rels = ensure_sheet_rels_has_drawings(
                self.part(&sheet_rels_part),
                &sheet_rels_part,
                &preserved_sheet.drawings,
            )?;
            self.set_part(sheet_rels_part, updated_rels);
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
) -> Result<Vec<u8>, ChartExtractionError> {
    if drawings.is_empty() {
        return Ok(sheet_xml.to_vec());
    }

    let mut xml = std::str::from_utf8(sheet_xml)
        .map_err(|e| ChartExtractionError::XmlNonUtf8(part_name.to_string(), e))?
        .to_string();

    if !xml.contains("xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"")
    {
        let worksheet_start = xml
            .find("<worksheet")
            .ok_or_else(|| ChartExtractionError::XmlStructure(format!("{part_name}: missing <worksheet")))?;
        let tag_end_rel = xml[worksheet_start..]
            .find('>')
            .ok_or_else(|| ChartExtractionError::XmlStructure(format!("{part_name}: invalid <worksheet> start tag")))?;
        let insert_pos = worksheet_start + tag_end_rel;
        xml.insert_str(insert_pos, &format!(" xmlns:r=\"{REL_NS}\""));
    }

    let close_idx = xml.rfind("</worksheet>").ok_or_else(|| {
        ChartExtractionError::XmlStructure(format!("{part_name}: missing </worksheet>"))
    })?;
    let insert_idx = xml
        .rfind("<extLst")
        .filter(|idx| *idx < close_idx)
        .unwrap_or(close_idx);

    let mut to_insert = String::new();
    for drawing in drawings {
        if xml.contains(&format!("r:id=\"{}\"", drawing.rel_id)) {
            continue;
        }
        to_insert.push_str(&format!("<drawing r:id=\"{}\"/>", drawing.rel_id));
    }

    if !to_insert.is_empty() {
        xml.insert_str(insert_idx, &to_insert);
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
        let updated = ensure_sheet_xml_has_drawings(xml, "xl/worksheets/sheet1.xml", &drawings)
            .expect("patch sheet xml");
        let updated_str = std::str::from_utf8(&updated).unwrap();

        let drawing_pos = updated_str.find("<drawing").unwrap();
        let ext_pos = updated_str.find("<extLst").unwrap();
        assert!(drawing_pos < ext_pos, "drawing should be inserted before extLst");
    }
}

fn ensure_sheet_rels_has_drawings(
    rels_xml: Option<&[u8]>,
    part_name: &str,
    drawings: &[SheetDrawingRelationship],
) -> Result<Vec<u8>, ChartExtractionError> {
    if drawings.is_empty() {
        return Ok(rels_xml.unwrap_or_default().to_vec());
    }

    let mut xml = match rels_xml {
        Some(bytes) => std::str::from_utf8(bytes)
            .map_err(|e| ChartExtractionError::XmlNonUtf8(part_name.to_string(), e))?
            .to_string(),
        None => String::from(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n</Relationships>\n",
        ),
    };

    let insert_idx = xml.rfind("</Relationships>").ok_or_else(|| {
        ChartExtractionError::XmlStructure(format!("{part_name}: missing </Relationships>"))
    })?;

    for drawing in drawings {
        if xml.contains(&format!("Id=\"{}\"", drawing.rel_id)) {
            continue;
        }
        xml.insert_str(
            insert_idx,
            &format!(
                "  <Relationship Id=\"{}\" Type=\"{}\" Target=\"{}\"/>\n",
                drawing.rel_id, DRAWING_REL_TYPE, drawing.target
            ),
        );
    }

    Ok(xml.into_bytes())
}
