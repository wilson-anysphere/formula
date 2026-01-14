use std::io::{Cursor, Write};

use formula_model::drawings::DrawingObjectKind;
use zip::write::FileOptions;
use zip::ZipWriter;

fn build_package(entries: &[(&str, &[u8])]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    for (name, bytes) in entries {
        zip.start_file(*name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    zip.finish().unwrap().into_inner()
}

#[test]
fn load_from_bytes_trims_drawing_relationship_id_and_type_whitespace() {
    let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let workbook_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    // Note the whitespace around the relationship id.
    let sheet_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData/>
  <drawing r:id=" rId1 "/>
</worksheet>"#;

    // Note whitespace around `Id` and trailing whitespace in `Type`.
    let sheet_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id=" rId1 " Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing " Target="../drawings/drawing1.xml"/>
</Relationships>"#;

    // Minimal drawing with a single shape.
    let drawing_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
    <xdr:sp>
      <xdr:nvSpPr>
        <xdr:cNvPr id="1" name="Shape 1"/>
        <xdr:cNvSpPr/>
      </xdr:nvSpPr>
      <xdr:spPr>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </xdr:spPr>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

    // Drawing relationships part must exist for `DrawingPart::parse_from_parts` even if unused.
    let drawing_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"#;

    let bytes = build_package(&[
        (
            "[Content_Types].xml",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
</Types>"#,
        ),
        ("xl/workbook.xml", workbook_xml),
        ("xl/_rels/workbook.xml.rels", workbook_rels),
        ("xl/worksheets/sheet1.xml", sheet_xml),
        ("xl/worksheets/_rels/sheet1.xml.rels", sheet_rels),
        ("xl/drawings/drawing1.xml", drawing_xml),
        ("xl/drawings/_rels/drawing1.xml.rels", drawing_rels),
    ]);

    let doc = formula_xlsx::load_from_bytes(&bytes).expect("load workbook");
    assert_eq!(doc.workbook.sheets.len(), 1);
    let ws = &doc.workbook.sheets[0];
    assert!(
        ws.drawings
            .iter()
            .any(|o| matches!(o.kind, DrawingObjectKind::Shape { .. })),
        "expected at least one Shape drawing object, got {:#?}",
        ws.drawings
    );
}

