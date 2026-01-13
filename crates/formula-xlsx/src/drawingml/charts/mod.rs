use formula_model::charts::ChartModel;
use formula_model::drawings::Anchor;
use roxmltree::Document;

use crate::workbook::ChartExtractionError;

mod cache;
mod parse_chart_color_style;
mod parse_chart_ex;
mod parse_chart_space;
mod parse_chart_style;

pub use parse_chart_color_style::{parse_chart_color_style, ChartColorStyleParseError};
pub use parse_chart_ex::{parse_chart_ex, ChartExParseError};
pub use parse_chart_space::{parse_chart_space, ChartSpaceParseError};
pub use parse_chart_style::{parse_chart_style, ChartStyleParseError};

const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OpcPart {
    pub path: String,
    pub rels_path: Option<String>,
    pub rels_bytes: Option<Vec<u8>>,
    pub bytes: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartParts {
    pub chart: OpcPart,
    pub chart_ex: Option<OpcPart>,
    pub style: Option<OpcPart>,
    pub colors: Option<OpcPart>,
    /// Optional DrawingML part that stores user-defined shapes for the chart
    /// (callouts, overlays, etc.).
    ///
    /// This is referenced from the chart part's `.rels` via the
    /// `.../relationships/chartUserShapes` relationship type.
    pub user_shapes: Option<OpcPart>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct ChartObject {
    pub sheet_name: Option<String>,
    pub sheet_part: Option<String>,
    pub drawing_part: String,
    /// Relationship ID (`r:id`) used inside the drawing part to reference the chart part
    /// (`xl/charts/chartN.xml`).
    pub drawing_rel_id: String,
    /// DrawingML non-visual object id from `<xdr:cNvPr id="..."/>` (when present).
    pub drawing_object_id: Option<u32>,
    /// DrawingML non-visual object name from `<xdr:cNvPr name="..."/>` (when present).
    pub drawing_object_name: Option<String>,
    pub anchor: Anchor,
    /// Raw XML for the `<xdr:graphicFrame>` subtree inside the drawing part.
    pub drawing_frame_xml: String,
    pub parts: ChartParts,
    /// Parsed chart model (optional; set when a parser is available).
    pub model: Option<ChartModel>,
    pub diagnostics: Vec<ChartDiagnostic>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartDiagnostic {
    pub severity: ChartDiagnosticSeverity,
    pub message: String,
    pub part: Option<String>,
    pub xpath: Option<String>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartDiagnosticSeverity {
    Info,
    Warning,
    Error,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DrawingChartObjectRef {
    pub rel_id: String,
    pub drawing_object_id: Option<u32>,
    pub drawing_object_name: Option<String>,
    pub anchor: Anchor,
    pub drawing_frame_xml: String,
}

pub fn extract_chart_object_refs(
    drawing_xml: &[u8],
    part_name: &str,
) -> Result<Vec<DrawingChartObjectRef>, ChartExtractionError> {
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

        let Some(anchor_model) = super::anchor::parse_anchor(&anchor) else {
            continue;
        };

        for frame in anchor
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "graphicFrame")
        {
            let Some(chart) = frame
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "chart")
            else {
                continue;
            };
            let Some(rel_id) = chart
                .attribute((REL_NS, "id"))
                .or_else(|| chart.attribute("r:id"))
                .or_else(|| chart.attribute("id"))
            else {
                continue;
            };

            let frame_xml = slice_node_xml(&frame, xml).unwrap_or_default();

            let (drawing_object_id, drawing_object_name) = frame
                .descendants()
                .find(|n| n.is_element() && n.tag_name().name() == "cNvPr")
                .map(|c_nv_pr| {
                    let id = c_nv_pr
                        .attribute("id")
                        .and_then(|v| v.trim().parse::<u32>().ok());
                    let name = c_nv_pr.attribute("name").map(|v| v.to_string());
                    (id, name)
                })
                .unwrap_or((None, None));

            out.push(DrawingChartObjectRef {
                rel_id: rel_id.to_string(),
                drawing_object_id,
                drawing_object_name,
                anchor: anchor_model,
                drawing_frame_xml: frame_xml,
            });
        }
    }

    Ok(out)
}

#[cfg(test)]
mod tests {
    use super::extract_chart_object_refs;

    #[test]
    fn extract_chart_object_refs_finds_nested_graphic_frame() {
        let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from>
      <xdr:col>0</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>0</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>5</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>10</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:to>
      <xdr:grpSp>
        <xdr:nvGrpSpPr/>
        <xdr:grpSpPr/>
      <xdr:graphicFrame>
        <xdr:nvGraphicFramePr>
          <xdr:cNvPr id="7" name="Chart 7"/>
          <xdr:cNvGraphicFramePr/>
        </xdr:nvGraphicFramePr>
        <xdr:xfrm/>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart r:id="rId42"/>
          </a:graphicData>
        </a:graphic>
      </xdr:graphicFrame>
    </xdr:grpSp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>
"#;

        let refs = extract_chart_object_refs(drawing_xml.as_bytes(), "xl/drawings/drawing1.xml")
            .expect("chart refs parsed");
        assert_eq!(refs.len(), 1);

        let chart_ref = &refs[0];
        assert_eq!(chart_ref.rel_id, "rId42");
        assert_eq!(chart_ref.drawing_object_id, Some(7));
        assert_eq!(chart_ref.drawing_object_name.as_deref(), Some("Chart 7"));
        assert!(chart_ref.drawing_frame_xml.contains("<xdr:graphicFrame"));
        assert!(chart_ref.drawing_frame_xml.contains("r:id=\"rId42\""));
        assert!(!chart_ref.drawing_frame_xml.contains("<xdr:grpSp"));

        let frame_xml = chart_ref.drawing_frame_xml.trim();
        assert!(frame_xml.starts_with("<xdr:graphicFrame"));
        assert!(frame_xml.ends_with("</xdr:graphicFrame>"));
    }
}
fn slice_node_xml(node: &roxmltree::Node<'_, '_>, doc: &str) -> Option<String> {
    let range = node.range();
    doc.get(range).map(|s| s.to_string())
}
