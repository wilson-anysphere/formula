use std::collections::{BTreeMap, HashMap, HashSet};

use roxmltree::Document;

use crate::path::rels_for_part;
use crate::preserve::sheet_match::{match_sheet_by_name_or_index, workbook_sheet_parts};
use crate::relationships::parse_relationships;
use crate::workbook::ChartExtractionError;
use crate::XlsxPackage;

const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PIVOT_CACHE_DEF_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";
const PIVOT_TABLE_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";

/// Minimal metadata needed to re-attach preserved pivot relationships.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RelationshipStub {
    pub rel_id: String,
    pub target: String,
}

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
    /// The `<pivotCaches>` subtree from `xl/workbook.xml` (outer XML).
    pub workbook_pivot_caches: Option<Vec<u8>>,
    /// Relationships from `xl/_rels/workbook.xml.rels` required by `<pivotCaches>`.
    pub workbook_pivot_cache_rels: Vec<RelationshipStub>,
    /// Preserved `<pivotTables>` subtrees and `.rels` metadata per worksheet.
    pub sheet_pivot_tables: BTreeMap<String, PreservedSheetPivotTables>,
}

impl PreservedPivotParts {
    pub fn is_empty(&self) -> bool {
        self.parts.is_empty()
            && self.workbook_pivot_caches.is_none()
            && self.sheet_pivot_tables.is_empty()
            && self.workbook_pivot_cache_rels.is_empty()
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
        let workbook_pivot_caches =
            pivot_caches_node.map(|n| workbook_xml.as_bytes()[n.range()].to_vec());

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

        let workbook_rels_part = "xl/_rels/workbook.xml.rels";
        let rel_map: HashMap<String, crate::relationships::Relationship> =
            match self.part(workbook_rels_part) {
                Some(workbook_rels_xml) => {
                    parse_relationships(workbook_rels_xml, workbook_rels_part)?
                        .into_iter()
                        .map(|r| (r.id.clone(), r))
                        .collect()
                }
                None => HashMap::new(),
            };

        let mut workbook_pivot_cache_rels = Vec::new();
        for rid in workbook_pivot_cache_rids {
            if let Some(rel) = rel_map.get(&rid) {
                if rel.type_ != PIVOT_CACHE_DEF_REL_TYPE {
                    continue;
                }
                workbook_pivot_cache_rels.push(RelationshipStub {
                    rel_id: rid.clone(),
                    target: rel.target.clone(),
                });
            }
        }

        let sheets = workbook_sheet_parts(self)?;
        let mut sheet_pivot_tables: BTreeMap<String, PreservedSheetPivotTables> = BTreeMap::new();

        for sheet in sheets {
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
                if let Some(sheet_rels_xml) = self.part(&sheet_rels_part) {
                    let rels = parse_relationships(sheet_rels_xml, &sheet_rels_part)?;
                    let rel_map: HashMap<_, _> =
                        rels.into_iter().map(|r| (r.id.clone(), r)).collect();

                    for rid in pivot_table_rids {
                        if let Some(rel) = rel_map.get(&rid) {
                            if rel.type_ != PIVOT_TABLE_REL_TYPE {
                                continue;
                            }
                            pivot_table_rels.push(RelationshipStub {
                                rel_id: rid.clone(),
                                target: rel.target.clone(),
                            });
                        }
                    }
                }
            }

            sheet_pivot_tables.insert(
                sheet.name,
                PreservedSheetPivotTables {
                    sheet_index: sheet.index,
                    sheet_id: sheet.sheet_id,
                    pivot_tables_xml,
                    pivot_table_rels,
                },
            );
        }

        Ok(PreservedPivotParts {
            content_types_xml,
            parts,
            workbook_pivot_caches,
            workbook_pivot_cache_rels,
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
        for (name, bytes) in &preserved.parts {
            self.set_part(name.clone(), bytes.clone());
        }

        self.merge_content_types(&preserved.content_types_xml, preserved.parts.keys())?;

        if let Some(pivot_caches) = preserved.workbook_pivot_caches.as_deref() {
            let workbook_part = "xl/workbook.xml";
            let workbook_xml = self
                .part(workbook_part)
                .ok_or_else(|| ChartExtractionError::MissingPart(workbook_part.to_string()))?;
            let updated =
                ensure_workbook_xml_has_pivot_caches(workbook_xml, workbook_part, pivot_caches)?;
            self.set_part(workbook_part, updated);
        }

        if !preserved.workbook_pivot_cache_rels.is_empty() {
            let workbook_rels_part = "xl/_rels/workbook.xml.rels";
            let updated_workbook_rels = ensure_rels_has_relationships(
                self.part(workbook_rels_part),
                workbook_rels_part,
                PIVOT_CACHE_DEF_REL_TYPE,
                &preserved.workbook_pivot_cache_rels,
            )?;
            self.set_part(workbook_rels_part, updated_workbook_rels);
        }

        let sheets = workbook_sheet_parts(self)?;
        for (sheet_name, preserved_sheet) in &preserved.sheet_pivot_tables {
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

            let updated_sheet_xml = ensure_sheet_xml_has_pivot_tables(
                sheet_xml,
                &sheet.part_name,
                &preserved_sheet.pivot_tables_xml,
            )?;
            self.set_part(sheet.part_name.clone(), updated_sheet_xml);

            if !preserved_sheet.pivot_table_rels.is_empty() {
                let sheet_rels_part = rels_for_part(&sheet.part_name);
                let updated_sheet_rels = ensure_rels_has_relationships(
                    self.part(&sheet_rels_part),
                    &sheet_rels_part,
                    PIVOT_TABLE_REL_TYPE,
                    &preserved_sheet.pivot_table_rels,
                )?;
                self.set_part(sheet_rels_part, updated_sheet_rels);
            }
        }

        Ok(())
    }
}

fn ensure_workbook_xml_has_pivot_caches(
    workbook_xml: &[u8],
    part_name: &str,
    pivot_caches_xml: &[u8],
) -> Result<Vec<u8>, ChartExtractionError> {
    let mut xml = std::str::from_utf8(workbook_xml)
        .map_err(|e| ChartExtractionError::XmlNonUtf8(part_name.to_string(), e))?
        .to_string();

    if xml.contains("<pivotCaches") {
        return Ok(workbook_xml.to_vec());
    }

    let pivot_caches_str = std::str::from_utf8(pivot_caches_xml)
        .map_err(|e| ChartExtractionError::XmlNonUtf8("pivotCaches".to_string(), e))?;

    if pivot_caches_str.contains("r:id")
        && !xml.contains("xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"")
    {
        let workbook_start = xml
            .find("<workbook")
            .ok_or_else(|| ChartExtractionError::XmlStructure(format!("{part_name}: missing <workbook")))?;
        let tag_end = xml[workbook_start..].find('>').ok_or_else(|| {
            ChartExtractionError::XmlStructure(format!("{part_name}: invalid <workbook> start tag"))
        })?;
        let insert_pos = workbook_start + tag_end;
        xml.insert_str(insert_pos, &format!(" xmlns:r=\"{REL_NS}\""));
    }

    let close_idx = xml.rfind("</workbook>").ok_or_else(|| {
        ChartExtractionError::XmlStructure(format!("{part_name}: missing </workbook>"))
    })?;
    xml.insert_str(close_idx, pivot_caches_str);
    Ok(xml.into_bytes())
}

fn ensure_sheet_xml_has_pivot_tables(
    sheet_xml: &[u8],
    part_name: &str,
    pivot_tables_xml: &[u8],
) -> Result<Vec<u8>, ChartExtractionError> {
    let mut xml = std::str::from_utf8(sheet_xml)
        .map_err(|e| ChartExtractionError::XmlNonUtf8(part_name.to_string(), e))?
        .to_string();

    if xml.contains("<pivotTables") {
        return Ok(sheet_xml.to_vec());
    }

    let pivot_tables_str = std::str::from_utf8(pivot_tables_xml)
        .map_err(|e| ChartExtractionError::XmlNonUtf8("pivotTables".to_string(), e))?;

    if pivot_tables_str.contains("r:id")
        && !xml.contains("xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"")
    {
        let worksheet_start = xml
            .find("<worksheet")
            .ok_or_else(|| ChartExtractionError::XmlStructure(format!("{part_name}: missing <worksheet")))?;
        let tag_end = xml[worksheet_start..].find('>').ok_or_else(|| {
            ChartExtractionError::XmlStructure(format!("{part_name}: invalid <worksheet> start tag"))
        })?;
        let insert_pos = worksheet_start + tag_end;
        xml.insert_str(insert_pos, &format!(" xmlns:r=\"{REL_NS}\""));
    }

    let close_idx = xml.rfind("</worksheet>").ok_or_else(|| {
        ChartExtractionError::XmlStructure(format!("{part_name}: missing </worksheet>"))
    })?;
    let insert_idx = xml
        .rfind("<extLst")
        .filter(|idx| *idx < close_idx)
        .unwrap_or(close_idx);
    xml.insert_str(insert_idx, pivot_tables_str);
    Ok(xml.into_bytes())
}

fn ensure_rels_has_relationships(
    rels_xml: Option<&[u8]>,
    part_name: &str,
    rel_type: &str,
    relationships: &[RelationshipStub],
) -> Result<Vec<u8>, ChartExtractionError> {
    if relationships.is_empty() {
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

    for relationship in relationships {
        if xml.contains(&format!("Id=\"{}\"", relationship.rel_id)) {
            continue;
        }
        xml.insert_str(
            insert_idx,
            &format!(
                "  <Relationship Id=\"{}\" Type=\"{}\" Target=\"{}\"/>\n",
                relationship.rel_id, rel_type, relationship.target
            ),
        );
    }

    Ok(xml.into_bytes())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn inserts_pivot_tables_before_ext_lst() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/><extLst><ext/></extLst></worksheet>"#;
        let pivot_tables = br#"<pivotTables xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><pivotTable r:id="rId1"/></pivotTables>"#;
        let updated = ensure_sheet_xml_has_pivot_tables(xml, "xl/worksheets/sheet1.xml", pivot_tables)
            .expect("patch sheet xml");
        let updated_str = std::str::from_utf8(&updated).unwrap();

        let pivot_pos = updated_str.find("<pivotTables").unwrap();
        let ext_pos = updated_str.find("<extLst").unwrap();
        assert!(pivot_pos < ext_pos, "pivotTables should be inserted before extLst");
    }
}
