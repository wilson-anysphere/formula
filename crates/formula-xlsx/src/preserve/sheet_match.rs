use std::collections::HashMap;

use roxmltree::Document;

use crate::path::resolve_target;
use crate::relationships::parse_relationships;
use crate::workbook::ChartExtractionError;
use crate::XlsxPackage;

const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct WorkbookSheetPart {
    pub name: String,
    pub index: usize,
    pub sheet_id: Option<u32>,
    pub part_name: String,
}

pub(crate) fn workbook_sheet_parts(
    pkg: &XlsxPackage,
) -> Result<Vec<WorkbookSheetPart>, ChartExtractionError> {
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
        Some(workbook_rels_xml) => parse_relationships(workbook_rels_xml, workbook_rels_part)?
            .into_iter()
            .map(|r| (r.id.clone(), r))
            .collect(),
        None => HashMap::new(),
    };

    let mut sheets = Vec::new();
    for (index, sheet_node) in workbook_doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "sheet")
        .enumerate()
    {
        let sheet_name = match sheet_node.attribute("name") {
            Some(name) => name.to_string(),
            None => continue,
        };
        let sheet_id = sheet_node.attribute("sheetId").and_then(|v| v.parse::<u32>().ok());
        let sheet_rid = match sheet_node
            .attribute((REL_NS, "id"))
            .or_else(|| sheet_node.attribute("r:id"))
            .or_else(|| sheet_node.attribute("id"))
        {
            Some(id) => id,
            None => continue,
        };

        let target = match rel_map.get(sheet_rid) {
            Some(rel) => Some(resolve_target(workbook_part, &rel.target)),
            None => sheet_id
                .map(|sheet_id| format!("xl/worksheets/sheet{sheet_id}.xml"))
                .filter(|candidate| pkg.part(candidate).is_some()),
        };
        let Some(target) = target else {
            continue;
        };

        sheets.push(WorkbookSheetPart {
            name: sheet_name,
            index,
            sheet_id,
            part_name: target,
        });
    }

    Ok(sheets)
}

pub(crate) fn match_sheet_by_name_or_index<'a>(
    sheets: &'a [WorkbookSheetPart],
    preserved_name: &str,
    preserved_index: usize,
) -> Option<&'a WorkbookSheetPart> {
    sheets
        .iter()
        .find(|sheet| sheet.name == preserved_name)
        .or_else(|| sheets.iter().find(|sheet| sheet.index == preserved_index))
}
