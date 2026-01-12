use std::collections::HashMap;

use roxmltree::Document;

use formula_model::sheet_name_eq_case_insensitive;

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

pub(crate) fn workbook_sheet_parts_from_workbook_xml<F>(
    workbook_xml: &[u8],
    workbook_rels_xml: Option<&[u8]>,
    has_part: F,
) -> Result<Vec<WorkbookSheetPart>, ChartExtractionError>
where
    F: Fn(&str) -> bool,
{
    let workbook_part = "xl/workbook.xml";
    let workbook_xml = std::str::from_utf8(workbook_xml)
        .map_err(|e| ChartExtractionError::XmlNonUtf8(workbook_part.to_string(), e))?;
    let workbook_doc = Document::parse(workbook_xml)
        .map_err(|e| ChartExtractionError::XmlParse(workbook_part.to_string(), e))?;

    let workbook_rels_part = "xl/_rels/workbook.xml.rels";
    // Best-effort: workbook relationships are frequently missing in "regenerated" workbooks and
    // some producers emit malformed XML. For preservation, fall back to common sheet naming
    // conventions instead of erroring.
    let rel_map: HashMap<String, crate::relationships::Relationship> = match workbook_rels_xml {
        Some(workbook_rels_xml) => match parse_relationships(workbook_rels_xml, workbook_rels_part)
        {
            Ok(rels) => rels.into_iter().map(|r| (r.id.clone(), r)).collect(),
            Err(_) => HashMap::new(),
        },
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
            Some(rel)
                if !rel
                    .target_mode
                    .as_deref()
                    .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External")) =>
            {
                Some(resolve_target(workbook_part, &rel.target))
            }
            _ => sheet_id
                .map(|sheet_id| format!("xl/worksheets/sheet{sheet_id}.xml"))
                .filter(|candidate| has_part(candidate)),
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

pub(crate) fn workbook_sheet_parts(
    pkg: &XlsxPackage,
) -> Result<Vec<WorkbookSheetPart>, ChartExtractionError> {
    let workbook_part = "xl/workbook.xml";
    let workbook_xml = pkg
        .part(workbook_part)
        .ok_or_else(|| ChartExtractionError::MissingPart(workbook_part.to_string()))?;
    let workbook_rels_part = "xl/_rels/workbook.xml.rels";
    workbook_sheet_parts_from_workbook_xml(
        workbook_xml,
        pkg.part(workbook_rels_part),
        |candidate| pkg.part(candidate).is_some(),
    )
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
    fn workbook_sheet_parts_tolerates_malformed_workbook_rels() {
        let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

        let pkg = build_package(&[
            ("xl/workbook.xml", workbook_xml),
            ("xl/_rels/workbook.xml.rels", br#"<Relationships><Relationship"#),
            ("xl/worksheets/sheet1.xml", br#"<worksheet/>"#),
        ]);

        let sheets = workbook_sheet_parts(&pkg).expect("should fallback to sheetId");
        assert_eq!(sheets.len(), 1);
        assert_eq!(sheets[0].name, "Sheet1");
        assert_eq!(sheets[0].index, 0);
        assert_eq!(sheets[0].sheet_id, Some(1));
        assert_eq!(sheets[0].part_name, "xl/worksheets/sheet1.xml");
    }

    #[test]
    fn match_sheet_by_name_or_index_is_case_insensitive_like_excel() {
        let sheets = vec![WorkbookSheetPart {
            name: "Résumé".to_string(),
            index: 0,
            sheet_id: Some(1),
            part_name: "xl/worksheets/sheet1.xml".to_string(),
        }];

        let matched = match_sheet_by_name_or_index(&sheets, "résumé", 123).expect("name match");
        assert_eq!(matched.index, 0);
    }
}

pub(crate) fn match_sheet_by_name_or_index<'a>(
    sheets: &'a [WorkbookSheetPart],
    preserved_name: &str,
    preserved_index: usize,
) -> Option<&'a WorkbookSheetPart> {
    sheets
        .iter()
        .find(|sheet| sheet_name_eq_case_insensitive(&sheet.name, preserved_name))
        .or_else(|| sheets.iter().find(|sheet| sheet.index == preserved_index))
}
