use std::collections::BTreeMap;

use formula_model::drawings::DrawingObjectKind;
use formula_xlsx::drawings::DrawingPart;

fn sample_drawing_xml() -> Vec<u8> {
    br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from>
      <xdr:col>0</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>0</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>1</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>1</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:to>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="1" name="Picture 1"/>
        <xdr:cNvPicPr/>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip r:embed="rId1"/>
      </xdr:blipFill>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
  <xdr:twoCellAnchor>
    <xdr:from>
      <xdr:col>2</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>2</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>3</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>3</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:to>
    <xdr:sp>
      <xdr:nvSpPr>
        <xdr:cNvPr id="2" name="Shape 1"/>
        <xdr:cNvSpPr/>
      </xdr:nvSpPr>
      <xdr:spPr>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </xdr:spPr>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#
        .to_vec()
}

#[test]
fn drawing_part_parse_without_rels_is_best_effort() {
    let parts: BTreeMap<String, Vec<u8>> = [("xl/drawings/drawing1.xml".to_string(), sample_drawing_xml())]
        .into_iter()
        .collect();
    let mut workbook = formula_model::Workbook::new();
    let part = DrawingPart::parse_from_parts(
        0,
        "xl/drawings/drawing1.xml",
        &parts,
        &mut workbook,
    )
    .expect("drawing parse should be best-effort without .rels");

    assert!(
        part.objects
            .iter()
            .any(|obj| matches!(obj.kind, DrawingObjectKind::Shape { .. })),
        "expected a shape object to parse"
    );
    assert!(
        part.objects
            .iter()
            .any(|obj| matches!(obj.kind, DrawingObjectKind::Unknown { .. })),
        "expected an unknown object due to missing relationships"
    );
}

#[test]
fn drawing_part_parse_with_malformed_rels_is_best_effort() {
    let drawing_path = "xl/drawings/drawing1.xml";
    let rels_path = "xl/drawings/_rels/drawing1.xml.rels";
    let parts: BTreeMap<String, Vec<u8>> = [
        (drawing_path.to_string(), sample_drawing_xml()),
        // Not well-formed XML.
        (rels_path.to_string(), br#"<Relationships>"#.to_vec()),
    ]
    .into_iter()
    .collect();

    let mut workbook = formula_model::Workbook::new();
    let part =
        DrawingPart::parse_from_parts(0, drawing_path, &parts, &mut workbook)
            .expect("drawing parse should be best-effort even with malformed rels");

    assert!(
        part.objects
            .iter()
            .any(|obj| matches!(obj.kind, DrawingObjectKind::Shape { .. })),
        "expected a shape object to parse"
    );
    assert!(
        part.objects
            .iter()
            .any(|obj| matches!(obj.kind, DrawingObjectKind::Unknown { .. })),
        "expected an unknown object due to malformed relationships"
    );
}

#[test]
fn drawing_part_missing_image_relationship_becomes_unknown() {
    let drawing_path = "xl/drawings/drawing1.xml";
    let rels_path = "xl/drawings/_rels/drawing1.xml.rels";
    let rels_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.png"/>
</Relationships>"#
        .to_vec();

    let parts: BTreeMap<String, Vec<u8>> = [
        (drawing_path.to_string(), sample_drawing_xml()),
        (rels_path.to_string(), rels_xml),
    ]
    .into_iter()
    .collect();

    let mut workbook = formula_model::Workbook::new();
    let part = DrawingPart::parse_from_parts(0, drawing_path, &parts, &mut workbook)
        .expect("drawing parse should be best-effort even when image rel is missing");

    let mut saw_unknown_pic = false;
    for obj in &part.objects {
        if let DrawingObjectKind::Unknown { raw_xml } = &obj.kind {
            if raw_xml.contains(r#"r:embed="rId1""#) && raw_xml.contains("<xdr:pic>") {
                saw_unknown_pic = true;
            }
        }
    }
    assert!(
        saw_unknown_pic,
        "expected the pic anchor referencing missing relationship to become Unknown"
    );
}

