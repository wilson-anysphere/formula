use std::io::{Cursor, Write};

use formula_xlsx::XlsxPackage;

#[test]
fn macro_strip_preserves_worksheet_ole_object_embeddings() {
    let bytes = build_macro_workbook_with_ole_object();
    let mut pkg = XlsxPackage::from_bytes(&bytes).expect("parse pkg");

    assert!(pkg.part("xl/vbaProject.bin").is_some());
    assert!(pkg.part("xl/embeddings/oleObject1.bin").is_some());

    pkg.remove_vba_project().expect("strip vba project");

    let written = pkg.write_to_bytes().expect("write stripped pkg");
    let pkg2 = XlsxPackage::from_bytes(&written).expect("read stripped pkg");

    assert!(
        pkg2.part("xl/vbaProject.bin").is_none(),
        "expected vbaProject.bin to be removed"
    );

    // Non-control OLE objects are valid in `.xlsx` and should be preserved when stripping VBA.
    assert!(
        pkg2.part("xl/embeddings/oleObject1.bin").is_some(),
        "expected worksheet OLE embedding to remain"
    );

    let sheet_xml = std::str::from_utf8(pkg2.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    assert!(
        sheet_xml.contains("<oleObjects"),
        "expected <oleObjects> to remain in worksheet XML"
    );
    assert!(
        sheet_xml.contains("r:id=\"rIdOle\""),
        "expected worksheet oleObject r:id to remain"
    );

    let sheet_rels =
        std::str::from_utf8(pkg2.part("xl/worksheets/_rels/sheet1.xml.rels").unwrap()).unwrap();
    assert!(
        sheet_rels.contains("Id=\"rIdOle\""),
        "expected worksheet relationship for OLE embedding to remain"
    );
    assert!(
        sheet_rels.contains("Target=\"../embeddings/oleObject1.bin\""),
        "expected worksheet relationship to keep targeting the embedding"
    );
}

fn build_macro_workbook_with_ole_object() -> Vec<u8> {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>
  <Override PartName="/xl/embeddings/oleObject1.bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
</Types>"#;

    let root_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>
</Relationships>"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData/>
  <oleObjects>
    <oleObject progId="Package" dvAspect="DVASPECT_ICON" oleUpdate="OLEUPDATE_ALWAYS" shapeId="1" r:id="rIdOle"/>
  </oleObjects>
</worksheet>"#;

    let sheet_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdOle" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="../embeddings/oleObject1.bin"/>
</Relationships>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types.as_bytes()).unwrap();

    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(root_rels.as_bytes()).unwrap();

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(sheet_xml.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/_rels/sheet1.xml.rels", options)
        .unwrap();
    zip.write_all(sheet_rels.as_bytes()).unwrap();

    zip.start_file("xl/embeddings/oleObject1.bin", options)
        .unwrap();
    zip.write_all(b"OLE DATA").unwrap();

    zip.start_file("xl/vbaProject.bin", options).unwrap();
    zip.write_all(b"VBA DATA").unwrap();

    zip.start_file("xl/_rels/vbaProject.bin.rels", options)
        .unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#,
    )
    .unwrap();

    zip.finish().unwrap().into_inner()
}
