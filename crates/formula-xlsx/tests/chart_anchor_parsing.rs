use formula_model::charts::ChartAnchor;
use formula_model::drawings::{Anchor, AnchorPoint, CellOffset, EmuSize};
use formula_model::CellRef;

use formula_xlsx::drawingml;

fn wrap_wsdr(body: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
{body}
</xdr:wsDr>"#
    )
}

fn chart_graphic_frame(rel_id: &str) -> String {
    format!(
        r#"<xdr:graphicFrame>
  <a:graphic>
    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
      <c:chart r:id="{rel_id}"/>
    </a:graphicData>
  </a:graphic>
</xdr:graphicFrame>"#
    )
}

#[test]
fn chart_anchor_absolute_anchor_parses_whitespace() {
    let xml = wrap_wsdr(&format!(
        r#"<xdr:absoluteAnchor>
  <xdr:pos x=" 100 " y=" 200 "/>
  <xdr:ext cx=" 300 " cy=" 400 "/>
  {}
  <xdr:clientData/>
</xdr:absoluteAnchor>"#,
        chart_graphic_frame("rId1")
    ));

    let drawing_refs =
        drawingml::extract_chart_refs(xml.as_bytes(), "xl/drawings/drawing1.xml").unwrap();
    assert_eq!(drawing_refs.len(), 1);
    assert_eq!(drawing_refs[0].rel_id, "rId1");
    assert_eq!(
        drawing_refs[0].anchor,
        ChartAnchor::Absolute {
            x_emu: 100,
            y_emu: 200,
            cx_emu: 300,
            cy_emu: 400,
        }
    );

    let object_refs =
        drawingml::charts::extract_chart_object_refs(xml.as_bytes(), "xl/drawings/drawing1.xml")
            .unwrap();
    assert_eq!(object_refs.len(), 1);
    assert_eq!(object_refs[0].rel_id, "rId1");
    assert_eq!(
        object_refs[0].anchor,
        Anchor::Absolute {
            pos: CellOffset::new(100, 200),
            ext: EmuSize::new(300, 400),
        }
    );
}

#[test]
fn chart_anchor_one_cell_anchor_defaults_missing_offsets_to_zero() {
    // Note: xdr:colOff / xdr:rowOff are optional in the wild; default them to 0.
    let xml = wrap_wsdr(&format!(
        r#"<xdr:oneCellAnchor>
  <xdr:from>
    <xdr:col> 1 </xdr:col>
    <xdr:row> 2 </xdr:row>
  </xdr:from>
  <xdr:ext cx=" 300 " cy=" 400 "/>
  {}
  <xdr:clientData/>
</xdr:oneCellAnchor>"#,
        chart_graphic_frame("rId2")
    ));

    let drawing_refs =
        drawingml::extract_chart_refs(xml.as_bytes(), "xl/drawings/drawing1.xml").unwrap();
    assert_eq!(drawing_refs.len(), 1);
    assert_eq!(drawing_refs[0].rel_id, "rId2");
    assert_eq!(
        drawing_refs[0].anchor,
        ChartAnchor::OneCell {
            from_col: 1,
            from_row: 2,
            from_col_off_emu: 0,
            from_row_off_emu: 0,
            cx_emu: 300,
            cy_emu: 400,
        }
    );

    let object_refs =
        drawingml::charts::extract_chart_object_refs(xml.as_bytes(), "xl/drawings/drawing1.xml")
            .unwrap();
    assert_eq!(object_refs.len(), 1);
    assert_eq!(object_refs[0].rel_id, "rId2");
    assert_eq!(
        object_refs[0].anchor,
        Anchor::OneCell {
            from: AnchorPoint::new(CellRef::new(2, 1), CellOffset::new(0, 0)),
            ext: EmuSize::new(300, 400),
        }
    );
}

#[test]
fn chart_anchor_two_cell_anchor_defaults_missing_offsets_to_zero() {
    let xml = wrap_wsdr(&format!(
        r#"<xdr:twoCellAnchor>
  <xdr:from>
    <xdr:col>  5 </xdr:col>
    <xdr:row>  6 </xdr:row>
  </xdr:from>
  <xdr:to>
    <xdr:col> 7 </xdr:col>
    <xdr:row> 8 </xdr:row>
  </xdr:to>
  {}
  <xdr:clientData/>
</xdr:twoCellAnchor>"#,
        chart_graphic_frame("rId3")
    ));

    let drawing_refs =
        drawingml::extract_chart_refs(xml.as_bytes(), "xl/drawings/drawing1.xml").unwrap();
    assert_eq!(drawing_refs.len(), 1);
    assert_eq!(drawing_refs[0].rel_id, "rId3");
    assert_eq!(
        drawing_refs[0].anchor,
        ChartAnchor::TwoCell {
            from_col: 5,
            from_row: 6,
            from_col_off_emu: 0,
            from_row_off_emu: 0,
            to_col: 7,
            to_row: 8,
            to_col_off_emu: 0,
            to_row_off_emu: 0,
        }
    );

    let object_refs =
        drawingml::charts::extract_chart_object_refs(xml.as_bytes(), "xl/drawings/drawing1.xml")
            .unwrap();
    assert_eq!(object_refs.len(), 1);
    assert_eq!(object_refs[0].rel_id, "rId3");
    assert_eq!(
        object_refs[0].anchor,
        Anchor::TwoCell {
            from: AnchorPoint::new(CellRef::new(6, 5), CellOffset::new(0, 0)),
            to: AnchorPoint::new(CellRef::new(8, 7), CellOffset::new(0, 0)),
        }
    );
}

