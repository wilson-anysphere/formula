use std::collections::HashMap;

use formula_model::charts::{Chart, ChartType};
use roxmltree::Document;

use crate::charts::parse_chart;
use crate::drawingml::extract_chart_refs;
use crate::package::XlsxPackage;
use crate::path::{rels_for_part, resolve_target};
use crate::relationships::parse_relationships;

#[derive(Debug, thiserror::Error)]
pub enum ChartExtractionError {
    #[error("missing part: {0}")]
    MissingPart(String),
    #[error("part is not valid UTF-8: {0}: {1}")]
    XmlNonUtf8(String, #[source] std::str::Utf8Error),
    #[error("failed to parse XML: {0}: {1}")]
    XmlParse(String, #[source] roxmltree::Error),
    #[error("invalid XML structure: {0}")]
    XmlStructure(String),
}

impl XlsxPackage {
    pub fn extract_charts(&self) -> Result<Vec<Chart>, ChartExtractionError> {
        let drawing_to_sheet = self.drawing_to_sheet_map().unwrap_or_default();
        let mut charts = Vec::new();

        for drawing_part in self
            .part_names()
            .filter(|name| name.starts_with("xl/drawings/drawing") && name.ends_with(".xml"))
            .map(str::to_string)
            .collect::<Vec<_>>()
        {
            let sheet_info = drawing_to_sheet.get(&drawing_part);
            let sheet_name = sheet_info.map(|info| info.name.clone());
            let sheet_part = sheet_info.and_then(|info| info.part.clone());

            let drawing_xml = match self.part(&drawing_part) {
                Some(bytes) => bytes,
                None => continue,
            };

            let drawing_refs = extract_chart_refs(drawing_xml, &drawing_part)?;
            if drawing_refs.is_empty() {
                continue;
            }

            let drawing_rels_part = rels_for_part(&drawing_part);
            let drawing_rels = self.part(&drawing_rels_part);
            let drawing_rel_map = match drawing_rels {
                Some(xml) => parse_relationships(xml, &drawing_rels_part)?
                    .into_iter()
                    .map(|r| (r.id.clone(), r))
                    .collect::<HashMap<_, _>>(),
                None => HashMap::new(),
            };

            for drawing_ref in drawing_refs {
                let relationship = drawing_rel_map.get(&drawing_ref.rel_id);
                let chart_part = relationship
                    .map(|r| resolve_target(&drawing_part, &r.target))
                    .filter(|target| self.part(target).is_some());

                let parsed = match chart_part
                    .as_deref()
                    .and_then(|chart_part| self.part(chart_part).map(|bytes| (chart_part, bytes)))
                {
                    Some((chart_part, bytes)) => parse_chart(bytes, chart_part).unwrap_or(None),
                    None => None,
                };

                let (chart_type, title, series) = match parsed {
                    Some(parsed) => (parsed.chart_type, parsed.title, parsed.series),
                    None => (
                        ChartType::Unknown {
                            name: "unparsed".to_string(),
                        },
                        None,
                        Vec::new(),
                    ),
                };

                charts.push(Chart {
                    sheet_name: sheet_name.clone(),
                    sheet_part: sheet_part.clone(),
                    drawing_part: drawing_part.clone(),
                    chart_part: chart_part.map(|s| s.to_string()),
                    rel_id: drawing_ref.rel_id,
                    chart_type,
                    title,
                    series,
                    anchor: drawing_ref.anchor,
                });
            }
        }

        Ok(charts)
    }

    fn drawing_to_sheet_map(&self) -> Result<HashMap<String, SheetInfo>, ChartExtractionError> {
        const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        let workbook_part = "xl/workbook.xml";
        let workbook_xml = match self.part(workbook_part) {
            Some(xml) => xml,
            None => return Ok(HashMap::new()),
        };

        let workbook_xml = std::str::from_utf8(workbook_xml)
            .map_err(|e| ChartExtractionError::XmlNonUtf8(workbook_part.to_string(), e))?;
        let workbook_doc = Document::parse(workbook_xml)
            .map_err(|e| ChartExtractionError::XmlParse(workbook_part.to_string(), e))?;

        let workbook_rels_part = "xl/_rels/workbook.xml.rels";
        let workbook_rels_xml = match self.part(workbook_rels_part) {
            Some(xml) => xml,
            None => return Ok(HashMap::new()),
        };

        let workbook_rels = parse_relationships(workbook_rels_xml, workbook_rels_part)?;
        let mut workbook_rel_map = HashMap::new();
        for rel in workbook_rels {
            workbook_rel_map.insert(rel.id.clone(), rel);
        }

        let mut out = HashMap::new();

        for sheet_node in workbook_doc
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "sheet")
        {
            let sheet_name = match sheet_node.attribute("name") {
                Some(name) => name.to_string(),
                None => continue,
            };
            let sheet_rid = match sheet_node
                .attribute((REL_NS, "id"))
                .or_else(|| sheet_node.attribute("r:id"))
                .or_else(|| sheet_node.attribute("id"))
            {
                Some(id) => id,
                None => continue,
            };

            let sheet_target = match workbook_rel_map.get(sheet_rid) {
                Some(rel) => resolve_target(workbook_part, &rel.target),
                None => continue,
            };

            let sheet_xml = match self.part(&sheet_target) {
                Some(xml) => xml,
                None => continue,
            };
            let sheet_xml = std::str::from_utf8(sheet_xml)
                .map_err(|e| ChartExtractionError::XmlNonUtf8(sheet_target.clone(), e))?;
            let sheet_doc = Document::parse(sheet_xml)
                .map_err(|e| ChartExtractionError::XmlParse(sheet_target.clone(), e))?;

            let drawing_rids: Vec<String> = sheet_doc
                .descendants()
                .filter(|n| n.is_element() && n.tag_name().name() == "drawing")
                .filter_map(|n| {
                    n.attribute((REL_NS, "id"))
                        .or_else(|| n.attribute("r:id"))
                        .or_else(|| n.attribute("id"))
                })
                .map(|s| s.to_string())
                .collect();

            if drawing_rids.is_empty() {
                continue;
            }

            let sheet_rels_part = rels_for_part(&sheet_target);
            let sheet_rels_xml = match self.part(&sheet_rels_part) {
                Some(xml) => xml,
                None => continue,
            };
            let sheet_rels = parse_relationships(sheet_rels_xml, &sheet_rels_part)?;
            let rel_map: HashMap<_, _> = sheet_rels.into_iter().map(|r| (r.id.clone(), r)).collect();

            for drawing_rid in drawing_rids {
                let drawing_target = match rel_map.get(&drawing_rid) {
                    Some(rel) => resolve_target(&sheet_target, &rel.target),
                    None => continue,
                };

                out.insert(
                    drawing_target,
                    SheetInfo {
                        name: sheet_name.clone(),
                        part: Some(sheet_target.clone()),
                    },
                );
            }
        }

        Ok(out)
    }
}

#[derive(Debug, Clone)]
struct SheetInfo {
    name: String,
    part: Option<String>,
}
