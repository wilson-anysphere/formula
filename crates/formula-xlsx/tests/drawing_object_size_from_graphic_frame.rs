use std::collections::BTreeMap;

use formula_model::drawings::{Anchor, DrawingObjectKind, EmuSize};
use formula_xlsx::drawings::DrawingPart;

#[test]
fn two_cell_anchor_graphic_frame_extracts_size_from_transform_ext() {
    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from>
      <xdr:col>1</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>2</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>5</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>10</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="42" name="Chart 42"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm>
        <a:off x="0" y="0"/>
        <a:ext cx="1234567" cy="7654321"/>
      </xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart r:id="rId99"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

    let rels_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"#;

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
    let part = DrawingPart::parse_from_parts(0, "xl/drawings/drawing1.xml", &parts, &mut workbook)
        .expect("parse drawing");

    assert_eq!(part.objects.len(), 1);
    let object = &part.objects[0];

    assert!(matches!(object.anchor, Anchor::TwoCell { .. }));
    assert!(matches!(object.kind, DrawingObjectKind::ChartPlaceholder { .. }));
    assert_eq!(object.size, Some(EmuSize::new(1234567, 7654321)));
}

