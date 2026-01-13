use std::collections::BTreeMap;

use formula_model::drawings::{Anchor, CellOffset, EmuSize};
use formula_xlsx::drawings::DrawingPart;

#[test]
fn parse_absolute_anchor_drawing_part() {
    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:absoluteAnchor>
    <xdr:pos x="123" y="456"/>
    <xdr:ext cx="789" cy="1011"/>
    <xdr:graphicFrame>
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="1" name="Chart 1"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm>
        <a:off x="0" y="0"/>
        <a:ext cx="0" cy="0"/>
      </xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart r:id="rId1"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:absoluteAnchor>
</xdr:wsDr>"#;

    // `DrawingPart::parse_from_parts` requires the `.rels` file, even if it is empty.
    let rels_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#;

    let parts: BTreeMap<String, Vec<u8>> = [
        ("xl/drawings/drawing1.xml".to_string(), drawing_xml.as_bytes().to_vec()),
        (
            "xl/drawings/_rels/drawing1.xml.rels".to_string(),
            rels_xml.as_bytes().to_vec(),
        ),
    ]
    .into_iter()
    .collect();

    let mut workbook = formula_model::Workbook::new();
    let drawing = DrawingPart::parse_from_parts(
        0,
        "xl/drawings/drawing1.xml",
        &parts,
        &mut workbook,
    )
    .expect("parse drawing part with absoluteAnchor");

    assert_eq!(drawing.objects.len(), 1);
    assert_eq!(
        drawing.objects[0].anchor,
        Anchor::Absolute {
            pos: CellOffset::new(123, 456),
            ext: EmuSize::new(789, 1011),
        }
    );
}

