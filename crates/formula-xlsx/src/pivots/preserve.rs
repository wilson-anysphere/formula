use std::collections::{BTreeMap, HashMap, HashSet};

use roxmltree::Document;

use crate::path::{rels_for_part, resolve_target};
use crate::preserve::sheet_match::{match_sheet_by_name_or_index, workbook_sheet_parts};
use crate::relationships::{parse_relationships, Relationship, Relationships};
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
        if preserved.is_empty() {
            return Ok(());
        }

        for (name, bytes) in &preserved.parts {
            self.set_part(name.clone(), bytes.clone());
        }

        self.merge_content_types(&preserved.content_types_xml, preserved.parts.keys())?;

        let workbook_part = "xl/workbook.xml";
        let workbook_rels_part = "xl/_rels/workbook.xml.rels";

        let workbook_rid_map = if !preserved.workbook_pivot_cache_rels.is_empty() {
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

        if let Some(pivot_caches) = preserved.workbook_pivot_caches.as_deref() {
            if !preserved.workbook_pivot_cache_rels.is_empty() {
                let rewritten =
                    rewrite_relationship_ids(pivot_caches, "pivotCaches", &workbook_rid_map)?;
                let workbook_xml = self.part(workbook_part).ok_or_else(|| {
                    ChartExtractionError::MissingPart(workbook_part.to_string())
                })?;
                let updated = ensure_workbook_xml_has_pivot_caches(
                    workbook_xml,
                    workbook_part,
                    &rewritten,
                )?;
                self.set_part(workbook_part, updated);
            }
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
            let updated_sheet_xml = ensure_sheet_xml_has_pivot_tables(
                sheet_xml,
                &sheet.part_name,
                &rewritten,
            )?;
            self.set_part(sheet.part_name.clone(), updated_sheet_xml);
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
        && !xml.contains(
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"",
        )
    {
        let workbook_start = xml.find("<workbook").ok_or_else(|| {
            ChartExtractionError::XmlStructure(format!("{part_name}: missing <workbook"))
        })?;
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
    let mut xml = std::str::from_utf8(sheet_xml)
        .map_err(|e| ChartExtractionError::XmlNonUtf8(part_name.to_string(), e))?
        .to_string();

    let pivot_tables_str = std::str::from_utf8(pivot_tables_xml)
        .map_err(|e| ChartExtractionError::XmlNonUtf8("pivotTables".to_string(), e))?;

    let desired_rids = extract_pivot_table_rids(pivot_tables_str, "pivotTables")?;
    if desired_rids.is_empty() {
        return Ok(sheet_xml.to_vec());
    }

    // Ensure the `r` namespace exists when we insert `r:id` attributes.
    let needs_r_namespace = pivot_tables_str.contains("r:id") || xml.contains("<pivotTables");
    if needs_r_namespace
        && !xml.contains(
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"",
        )
    {
        let worksheet_start = xml.find("<worksheet").ok_or_else(|| {
            ChartExtractionError::XmlStructure(format!("{part_name}: missing <worksheet"))
        })?;
        let tag_end = xml[worksheet_start..].find('>').ok_or_else(|| {
            ChartExtractionError::XmlStructure(format!("{part_name}: invalid <worksheet> start tag"))
        })?;
        let insert_pos = worksheet_start + tag_end;
        xml.insert_str(insert_pos, &format!(" xmlns:r=\"{REL_NS}\""));
    }

    let blocks = find_pivot_tables_blocks(&xml);
    if blocks.is_empty() {
        let insert_idx = pivot_tables_insertion_index(&xml, part_name)?;
        xml.insert_str(insert_idx, pivot_tables_str);
        return Ok(xml.into_bytes());
    }

    // Merge all existing `<pivotTables>` blocks (in case the worksheet is already malformed)
    // plus the preserved block.
    let mut merged_rids: Vec<String> = Vec::new();
    let mut seen: HashSet<String> = HashSet::new();

    for (start, end) in &blocks {
        let block = &xml[*start..*end];
        for rid in extract_pivot_table_rids(block, part_name)? {
            if seen.insert(rid.clone()) {
                merged_rids.push(rid);
            }
        }
    }

    for rid in desired_rids {
        if seen.insert(rid.clone()) {
            merged_rids.push(rid);
        }
    }

    let merged_block = build_pivot_tables_xml(&merged_rids);

    // Remove all but the first `<pivotTables>` block, then replace the first with the merged one.
    let first = blocks[0];
    for (start, end) in blocks.iter().rev() {
        if (*start, *end) == first {
            continue;
        }
        xml.replace_range(*start..*end, "");
    }
    xml.replace_range(first.0..first.1, &merged_block);

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

    let mut xml = std::str::from_utf8(xml_bytes)
        .map_err(|e| ChartExtractionError::XmlNonUtf8(part_name.to_string(), e))?
        .to_string();

    for (old_id, new_id) in id_map {
        xml = xml.replace(&format!("r:id=\"{old_id}\""), &format!("r:id=\"{new_id}\""));
        xml = xml.replace(&format!("r:id='{old_id}'"), &format!("r:id='{new_id}'"));
    }

    Ok(xml.into_bytes())
}

fn ensure_rels_has_relationships(
    rels_xml: Option<&[u8]>,
    part_name: &str,
    base_part: &str,
    rel_type: &str,
    relationships: &[RelationshipStub],
) -> Result<(Vec<u8>, HashMap<String, String>), ChartExtractionError> {
    if relationships.is_empty() {
        return Ok((rels_xml.unwrap_or_default().to_vec(), HashMap::new()));
    }

    let mut xml = match rels_xml {
        Some(bytes) => std::str::from_utf8(bytes)
            .map_err(|e| ChartExtractionError::XmlNonUtf8(part_name.to_string(), e))?
            .to_string(),
        None => String::from(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n</Relationships>\n",
        ),
    };

    let existing_rels = match rels_xml {
        Some(bytes) => parse_relationships(bytes, part_name)?,
        None => Vec::new(),
    };
    let mut rels = Relationships::new(existing_rels);

    let mut id_map: HashMap<String, String> = HashMap::new();
    let mut to_insert: Vec<Relationship> = Vec::new();

    for relationship in relationships {
        let desired_id = relationship.rel_id.as_str();
        let desired_target = relationship.target.as_str();

        if let Some(mapped) = id_map.get(desired_id) {
            // We've already allocated a stable replacement for this ID in this scope.
            // Ensure the relationship exists in the output but don't allocate again.
            if rels.get(mapped).is_none() {
                let rel = Relationship {
                    id: mapped.clone(),
                    type_: rel_type.to_string(),
                    target: desired_target.to_string(),
                };
                rels.push(rel.clone());
                to_insert.push(rel);
            }
            continue;
        }

        let final_id = match rels.get(desired_id) {
            None => desired_id.to_string(),
            Some(existing)
                if existing.type_ == rel_type
                    && resolve_target(base_part, &existing.target)
                        == resolve_target(base_part, desired_target) =>
            {
                desired_id.to_string()
            }
            Some(_) => {
                let new_id = rels.next_r_id();
                id_map.insert(desired_id.to_string(), new_id.clone());
                new_id
            }
        };

        if rels.get(&final_id).is_some() {
            continue;
        }

        let rel = Relationship {
            id: final_id.clone(),
            type_: rel_type.to_string(),
            target: desired_target.to_string(),
        };
        rels.push(rel.clone());
        to_insert.push(rel);
    }

    if !to_insert.is_empty() {
        let insert_idx = xml.rfind("</Relationships>").ok_or_else(|| {
            ChartExtractionError::XmlStructure(format!("{part_name}: missing </Relationships>"))
        })?;

        let mut insertion = String::new();
        for rel in &to_insert {
            insertion.push_str(&format!(
                "  <Relationship Id=\"{}\" Type=\"{}\" Target=\"{}\"/>\n",
                xml_escape(&rel.id),
                xml_escape(&rel.type_),
                xml_escape(&rel.target)
            ));
        }
        xml.insert_str(insert_idx, &insertion);
    }

    Ok((xml.into_bytes(), id_map))
}

fn xml_escape(input: &str) -> String {
    input
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

fn sheet_data_end_idx(xml: &str) -> Option<usize> {
    if let Some(idx) = xml.rfind("</sheetData>") {
        return Some(idx + "</sheetData>".len());
    }

    let start = xml.find("<sheetData")?;
    let gt_rel = xml[start..].find('>')?;
    let gt = start + gt_rel;
    let tag = &xml[start..=gt];
    if tag.trim_end().ends_with("/>") {
        return Some(gt + 1);
    }

    None
}

fn pivot_tables_insertion_index(xml: &str, part_name: &str) -> Result<usize, ChartExtractionError> {
    let close_idx = xml.rfind("</worksheet>").ok_or_else(|| {
        ChartExtractionError::XmlStructure(format!("{part_name}: missing </worksheet>"))
    })?;

    let sheet_data_end = sheet_data_end_idx(xml).unwrap_or(0);

    let ext_idx = xml.rfind("<extLst").filter(|idx| *idx < close_idx);
    if let Some(ext_idx) = ext_idx {
        if ext_idx >= sheet_data_end {
            return Ok(ext_idx);
        }
    }

    Ok(close_idx.max(sheet_data_end))
}

fn find_pivot_tables_blocks(xml: &str) -> Vec<(usize, usize)> {
    let mut out = Vec::new();
    let mut search = 0usize;

    while let Some(start_rel) = xml[search..].find("<pivotTables") {
        let start = search + start_rel;
        let gt_rel = match xml[start..].find('>') {
            Some(idx) => idx,
            None => break,
        };
        let gt = start + gt_rel;
        let open_tag = &xml[start..=gt];
        let is_self_closing = open_tag.trim_end().ends_with("/>");
        if is_self_closing {
            out.push((start, gt + 1));
            search = gt + 1;
            continue;
        }

        let close_tag = "</pivotTables>";
        let close_rel = match xml[gt + 1..].find(close_tag) {
            Some(idx) => idx,
            None => break,
        };
        let end = gt + 1 + close_rel + close_tag.len();
        out.push((start, end));
        search = end;
    }

    out
}

fn extract_pivot_table_rids(fragment: &str, context: &str) -> Result<Vec<String>, ChartExtractionError> {
    let wrapped = format!("<worksheet xmlns:r=\"{REL_NS}\">{fragment}</worksheet>");
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

fn build_pivot_tables_xml(rids: &[String]) -> String {
    let mut out = String::new();
    out.push_str(&format!(r#"<pivotTables count="{}">"#, rids.len()));
    for rid in rids {
        out.push_str(r#"<pivotTable r:id=""#);
        out.push_str(&xml_escape(rid));
        out.push_str(r#""/>"#);
    }
    out.push_str("</pivotTables>");
    out
}

fn xml_escape(input: &str) -> String {
    input
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
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
        assert!(pivot_pos < ext_pos, "pivotTables should be inserted before extLst");
    }
}
