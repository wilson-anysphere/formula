use formula_model::charts::ChartAnchor;
use roxmltree::{Document, Node};

use crate::workbook::ChartExtractionError;

pub(crate) mod anchor;
pub mod charts;
mod preserve;
pub mod style;

pub use preserve::{
    preserve_drawing_parts_from_reader, preserve_drawing_parts_from_reader_limited,
    PreservedChartSheet, PreservedDrawingParts,
    PreservedSheetControls, PreservedSheetDrawingHF, PreservedSheetDrawings,
    PreservedSheetOleObjects, PreservedSheetPicture, SheetDrawingRelationship,
    SheetRelationshipStub, SheetRelationshipStubWithType,
};

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DrawingChartRef {
    pub rel_id: String,
    pub anchor: ChartAnchor,
}

pub fn extract_chart_refs(
    drawing_xml: &[u8],
    part_name: &str,
) -> Result<Vec<DrawingChartRef>, ChartExtractionError> {
    fn is_chart_node(node: Node<'_, '_>) -> bool {
        node.is_element() && node.tag_name().name() == "chart"
    }

    let xml = std::str::from_utf8(drawing_xml)
        .map_err(|e| ChartExtractionError::XmlNonUtf8(part_name.to_string(), e))?;
    let doc = Document::parse(xml)
        .map_err(|e| ChartExtractionError::XmlParse(part_name.to_string(), e))?;

    let mut out = Vec::new();
    for anchor in anchor::wsdr_anchor_nodes(doc.root_element()) {

        let Some(anchor_model) = anchor::parse_anchor(&anchor) else {
            continue;
        };
        let anchor_model = anchor::anchor_to_chart_anchor(anchor_model);

        for chart in anchor::descendants_selecting_alternate_content(anchor, is_chart_node, is_chart_node)
        {
            let Some(rel_id) = chart
                .attribute((
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                    "id",
                ))
                .or_else(|| chart.attribute("r:id"))
                .or_else(|| chart.attribute("id"))
            else {
                continue;
            };

            out.push(DrawingChartRef {
                rel_id: rel_id.to_string(),
                anchor: anchor_model.clone(),
            });
        }
    }

    Ok(out)
}
