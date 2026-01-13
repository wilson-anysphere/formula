use formula_model::charts::ChartAnchor;
use roxmltree::Document;

use crate::workbook::ChartExtractionError;

pub mod charts;
mod preserve;
pub mod style;

pub use preserve::{
    PreservedChartSheet, PreservedDrawingParts, PreservedSheetControls, PreservedSheetDrawingHF,
    PreservedSheetDrawings, PreservedSheetOleObjects, PreservedSheetPicture,
    SheetDrawingRelationship, SheetRelationshipStub, SheetRelationshipStubWithType,
    preserve_drawing_parts_from_reader,
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
    let xml = std::str::from_utf8(drawing_xml)
        .map_err(|e| ChartExtractionError::XmlNonUtf8(part_name.to_string(), e))?;
    let doc = Document::parse(xml)
        .map_err(|e| ChartExtractionError::XmlParse(part_name.to_string(), e))?;

    let mut out = Vec::new();
    for anchor in doc.descendants().filter(|n| n.is_element()) {
        let anchor_kind = anchor.tag_name().name();
        if anchor_kind != "twoCellAnchor"
            && anchor_kind != "absoluteAnchor"
            && anchor_kind != "oneCellAnchor"
        {
            continue;
        }

        let Some(anchor_model) = parse_anchor(&anchor) else {
            continue;
        };

        for chart in anchor
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "chart")
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

fn parse_anchor(anchor: &roxmltree::Node<'_, '_>) -> Option<ChartAnchor> {
    match anchor.tag_name().name() {
        "absoluteAnchor" => {
            let pos = anchor
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "pos")?;
            let ext = anchor
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "ext")?;

            Some(ChartAnchor::Absolute {
                x_emu: pos.attribute("x")?.trim().parse().ok()?,
                y_emu: pos.attribute("y")?.trim().parse().ok()?,
                cx_emu: ext.attribute("cx")?.trim().parse().ok()?,
                cy_emu: ext.attribute("cy")?.trim().parse().ok()?,
            })
        }
        "oneCellAnchor" => {
            let from = anchor
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "from")?;
            let ext = anchor
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "ext")?;

            Some(ChartAnchor::OneCell {
                from_col: descendant_text(from, "col")
                    .and_then(|t| t.trim().parse().ok())?,
                from_row: descendant_text(from, "row")
                    .and_then(|t| t.trim().parse().ok())?,
                from_col_off_emu: descendant_text(from, "colOff")
                    .unwrap_or("0")
                    .trim()
                    .parse()
                    .ok()?,
                from_row_off_emu: descendant_text(from, "rowOff")
                    .unwrap_or("0")
                    .trim()
                    .parse()
                    .ok()?,
                cx_emu: ext.attribute("cx")?.trim().parse().ok()?,
                cy_emu: ext.attribute("cy")?.trim().parse().ok()?,
            })
        }
        "twoCellAnchor" => {
            let from = anchor
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "from")?;
            let to = anchor
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "to")?;

            Some(ChartAnchor::TwoCell {
                from_col: descendant_text(from, "col")
                    .and_then(|t| t.trim().parse().ok())?,
                from_row: descendant_text(from, "row")
                    .and_then(|t| t.trim().parse().ok())?,
                from_col_off_emu: descendant_text(from, "colOff")
                    .unwrap_or("0")
                    .trim()
                    .parse()
                    .ok()?,
                from_row_off_emu: descendant_text(from, "rowOff")
                    .unwrap_or("0")
                    .trim()
                    .parse()
                    .ok()?,
                to_col: descendant_text(to, "col").and_then(|t| t.trim().parse().ok())?,
                to_row: descendant_text(to, "row").and_then(|t| t.trim().parse().ok())?,
                to_col_off_emu: descendant_text(to, "colOff")
                    .unwrap_or("0")
                    .trim()
                    .parse()
                    .ok()?,
                to_row_off_emu: descendant_text(to, "rowOff")
                    .unwrap_or("0")
                    .trim()
                    .parse()
                    .ok()?,
            })
        }
        _ => None,
    }
}

fn descendant_text<'a>(node: roxmltree::Node<'a, 'a>, tag: &str) -> Option<&'a str> {
    node.children()
        .find(|n| n.is_element() && n.tag_name().name() == tag)
        .and_then(|n| n.text())
}
