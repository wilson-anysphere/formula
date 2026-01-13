use std::collections::BTreeMap;

use formula_model::drawings::{Anchor, DrawingObjectKind, EmuSize};
use formula_xlsx::drawings::DrawingPart;

#[test]
fn two_cell_anchor_shape_extracts_size_from_transform_ext() {
    let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from>
      <xdr:col>0</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>0</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>2</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>3</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:to>
    <xdr:sp>
      <xdr:nvSpPr>
        <xdr:cNvPr id="7" name="Shape 7"/>
        <xdr:cNvSpPr/>
      </xdr:nvSpPr>
      <xdr:spPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="111" cy="222"/>
        </a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </xdr:spPr>
    </xdr:sp>
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
    assert!(matches!(object.kind, DrawingObjectKind::Shape { .. }));
    assert_eq!(object.size, Some(EmuSize::new(111, 222)));
}

