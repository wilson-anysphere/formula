use std::collections::BTreeMap;

use formula_model::drawings::DrawingObjectKind;
use formula_xlsx::drawings::DrawingPart;

#[test]
fn drawings_roundtrip_preserves_anchor_and_client_data_attributes() {
    let drawing_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
 <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
           xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
           xmlns:dup1="urn:example"
           xmlns:dup2="urn:example"
           dup2:rootAttr="abc">
   <xdr:twoCellAnchor xmlns:xdr14="http://schemas.microsoft.com/office/drawing/2010/spreadsheetDrawing" editAs="oneCell" xdr14:anchorId="123">
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
    <xdr:sp>
      <xdr:nvSpPr>
        <xdr:cNvPr id="2" name="Shape 1"/>
        <xdr:cNvSpPr/>
      </xdr:nvSpPr>
      <xdr:spPr>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </xdr:spPr>
    </xdr:sp>
    <xdr:clientData fLocksWithSheet="0" fPrintsWithSheet="1"/>
  </xdr:twoCellAnchor>
</xdr:wsDr>
"#;

    let rels_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>
"#;

    let mut parts: BTreeMap<String, Vec<u8>> = BTreeMap::new();
    parts.insert("xl/drawings/drawing1.xml".to_string(), drawing_xml.to_vec());
    parts.insert(
        "xl/drawings/_rels/drawing1.xml.rels".to_string(),
        rels_xml.to_vec(),
    );

    let mut workbook = formula_model::Workbook::new();
    let mut part =
        DrawingPart::parse_from_parts(0, "xl/drawings/drawing1.xml", &parts, &mut workbook)
            .expect("parse drawing part");

    let shape = part
        .objects
        .iter()
        .find(|o| matches!(o.kind, DrawingObjectKind::Shape { .. }))
        .expect("shape object present");
    assert_eq!(
        shape
            .preserved
            .get("xlsx.anchor_edit_as")
            .map(|s| s.as_str()),
        Some("oneCell")
    );
    let attrs_json = shape
        .preserved
        .get("xlsx.anchor_attrs")
        .expect("anchor attrs preserved");
    let attrs: BTreeMap<String, String> = serde_json::from_str(attrs_json).expect("parse attrs json");
    assert_eq!(
        attrs.get("xmlns:xdr14").map(|s| s.as_str()),
        Some("http://schemas.microsoft.com/office/drawing/2010/spreadsheetDrawing")
    );
    assert_eq!(attrs.get("xdr14:anchorId").map(|s| s.as_str()), Some("123"));
    assert!(shape
        .preserved
        .get("xlsx.client_data_xml")
        .is_some_and(
            |xml| xml.contains("fLocksWithSheet=\"0\"") && xml.contains("fPrintsWithSheet=\"1\"")
        ));

    part.write_into_parts(&mut parts, &workbook)
        .expect("write drawing part");
    let out_xml = parts
        .get("xl/drawings/drawing1.xml")
        .expect("drawing written");

    // Round-trip should preserve anchor editAs and clientData attributes.
    formula_xlsx::assert_xml_semantic_eq(drawing_xml, out_xml).unwrap();
}
