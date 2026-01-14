use formula_model::drawings::DrawingObjectKind;
use formula_xlsx::XlsxPackage;
use std::io::{Cursor, Write};

#[test]
fn extract_drawing_objects_finds_image() {
    let bytes = std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/basic/image.xlsx"
    ))
    .expect("fixture exists");

    let pkg = XlsxPackage::from_bytes(&bytes).expect("read fixture package");
    let drawings = pkg
        .extract_drawing_objects()
        .expect("extract drawing objects");

    let image_count = drawings
        .iter()
        .flat_map(|entry| entry.objects.iter())
        .filter(|obj| matches!(obj.kind, DrawingObjectKind::Image { .. }))
        .count();

    assert_eq!(image_count, 1, "expected one image object in fixture");
}

fn zip_with_alternate_content_drawing_ref() -> Vec<u8> {
    let content_types_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/drawings/drawing2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
</Types>"#;

    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let workbook_rels_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
                Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
           xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <mc:AlternateContent>
    <mc:Choice Requires="x14ac"><drawing r:id="rId1"/></mc:Choice>
    <mc:Fallback><drawing r:id="rId2"/></mc:Fallback>
  </mc:AlternateContent>
</worksheet>"#;

    let sheet_rels_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
                Target="../drawings/drawing1.xml"/>
  <Relationship Id="rId2"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
                Target="../drawings/drawing2.xml"/>
</Relationships>"#;

    fn drawing_xml(shape_name: &str) -> String {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:row>0</xdr:row></xdr:from>
    <xdr:to><xdr:col>1</xdr:col><xdr:row>1</xdr:row></xdr:to>
    <xdr:sp>
      <xdr:nvSpPr><xdr:cNvPr id="1" name="{shape_name}"/></xdr:nvSpPr>
      <xdr:spPr/>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#
        )
    }

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", options)
        .expect("start [Content_Types].xml");
    zip.write_all(content_types_xml.as_bytes())
        .expect("write [Content_Types].xml");

    zip.start_file("xl/workbook.xml", options)
        .expect("start workbook.xml");
    zip.write_all(workbook_xml.as_bytes())
        .expect("write workbook.xml");

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .expect("start workbook rels");
    zip.write_all(workbook_rels_xml.as_bytes())
        .expect("write workbook rels");

    zip.start_file("xl/worksheets/sheet1.xml", options)
        .expect("start sheet1.xml");
    zip.write_all(sheet_xml.as_bytes()).expect("write sheet1");

    zip.start_file("xl/worksheets/_rels/sheet1.xml.rels", options)
        .expect("start sheet rels");
    zip.write_all(sheet_rels_xml.as_bytes())
        .expect("write sheet rels");

    zip.start_file("xl/drawings/drawing1.xml", options)
        .expect("start drawing1.xml");
    zip.write_all(drawing_xml("ChoiceShape").as_bytes())
        .expect("write drawing1");

    zip.start_file("xl/drawings/drawing2.xml", options)
        .expect("start drawing2.xml");
    zip.write_all(drawing_xml("FallbackShape").as_bytes())
        .expect("write drawing2");

    zip.finish().expect("finish zip").into_inner()
}

#[test]
fn extract_drawing_objects_respects_mc_alternate_content_for_drawing_refs() {
    let bytes = zip_with_alternate_content_drawing_ref();
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read package");
    let drawings = pkg
        .extract_drawing_objects()
        .expect("extract drawing objects");

    assert_eq!(drawings.len(), 1, "expected one drawing part for the sheet");
    assert_eq!(drawings[0].drawing_part, "xl/drawings/drawing1.xml");
    assert_eq!(drawings[0].sheet_name, "Sheet1");

    let shape_xmls: Vec<_> = drawings[0]
        .objects
        .iter()
        .filter_map(|obj| match &obj.kind {
            DrawingObjectKind::Shape { raw_xml } => Some(raw_xml.as_str()),
            _ => None,
        })
        .collect();

    assert_eq!(shape_xmls.len(), 1, "expected one shape object");
    assert!(
        shape_xmls[0].contains("ChoiceShape"),
        "expected Choice branch shape to be selected"
    );
    assert!(
        !shape_xmls[0].contains("FallbackShape"),
        "did not expect Fallback branch shape to be selected"
    );
}

#[test]
fn preserve_drawing_parts_respects_mc_alternate_content_for_drawing_refs() {
    let bytes = zip_with_alternate_content_drawing_ref();
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read package");

    let preserved = pkg
        .preserve_drawing_parts()
        .expect("preserve drawing parts");
    let sheet = preserved
        .sheet_drawings
        .get("Sheet1")
        .expect("preserved sheet drawings");

    assert_eq!(
        sheet.drawings.len(),
        1,
        "expected only the Choice branch drawing relationship to be preserved"
    );
    assert_eq!(sheet.drawings[0].rel_id, "rId1");
    assert_eq!(sheet.drawings[0].target, "../drawings/drawing1.xml");
}

#[test]
fn apply_preserved_drawing_parts_does_not_duplicate_alternate_content_drawings() {
    let bytes = zip_with_alternate_content_drawing_ref();
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read package");
    let preserved = pkg
        .preserve_drawing_parts()
        .expect("preserve drawing parts");

    let mut pkg2 = pkg.clone();
    pkg2.apply_preserved_drawing_parts(&preserved)
        .expect("apply preserved drawing parts");

    let sheet_xml = std::str::from_utf8(pkg2.part("xl/worksheets/sheet1.xml").unwrap())
        .expect("sheet xml utf8");
    assert_eq!(
        sheet_xml.matches("<drawing").count(),
        2,
        "expected apply to not insert a duplicate <drawing> outside mc:AlternateContent"
    );
}
