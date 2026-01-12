use std::io::{Cursor, Write};

use formula_xlsx::{load_from_bytes, validate_opc_relationships, XlsxPackage};

fn build_prefixed_xlsx() -> Vec<u8> {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types">
  <ct:Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <ct:Default Extension="xml" ContentType="application/xml"/>
  <ct:Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <ct:Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</ct:Types>"#;

    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:workbook
  xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <x:workbookPr/>
  <x:calcPr/>
  <x:bookViews>
    <x:workbookView activeTab="0" firstSheet="0"/>
  </x:bookViews>
  <x:sheets>
    <x:sheet name="Sheet1" sheetId="1" rel:id="rId1"/>
  </x:sheets>
</x:workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pr:Relationships xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships">
  <pr:Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</pr:Relationships>"#;

    let sheet1_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types.as_bytes()).unwrap();

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(sheet1_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

#[test]
fn writer_pipeline_preserves_prefixes_for_workbook_level_parts() {
    let bytes = build_prefixed_xlsx();
    let mut doc = load_from_bytes(&bytes).expect("load");

    doc.workbook.add_sheet("Sheet2").expect("add sheet");
    let out = doc.save_to_vec().expect("save");

    let pkg = XlsxPackage::from_bytes(&out).expect("read output pkg");

    let workbook_xml = std::str::from_utf8(pkg.part("xl/workbook.xml").unwrap()).unwrap();
    assert!(
        workbook_xml.contains(r#"<x:sheet name="Sheet1""#),
        "expected prefixed Sheet1 entry, got: {workbook_xml}"
    );
    assert!(
        workbook_xml.contains(r#"<x:sheet name="Sheet2""#),
        "expected prefixed Sheet2 entry, got: {workbook_xml}"
    );
    assert!(
        workbook_xml.contains(r#" rel:id="rId1""#),
        "expected relationship id attr to preserve prefix, got: {workbook_xml}"
    );
    assert!(
        workbook_xml.contains(r#" rel:id="rId2""#),
        "expected relationship id attr for new sheet to preserve prefix, got: {workbook_xml}"
    );
    assert!(
        !workbook_xml.contains(r#" r:id=""#),
        "writer introduced unexpected r:id prefix, got: {workbook_xml}"
    );
    assert!(
        !workbook_xml.contains("<sheet "),
        "writer introduced unprefixed <sheet> element, got: {workbook_xml}"
    );

    let workbook_rels_xml =
        std::str::from_utf8(pkg.part("xl/_rels/workbook.xml.rels").unwrap()).unwrap();
    assert!(
        workbook_rels_xml.contains(r#"<pr:Relationship Id="rId2""#),
        "expected inserted relationship element to preserve prefix, got: {workbook_rels_xml}"
    );
    assert!(
        !workbook_rels_xml.contains("<Relationship "),
        "writer introduced unprefixed <Relationship>, got: {workbook_rels_xml}"
    );

    let content_types_xml = std::str::from_utf8(pkg.part("[Content_Types].xml").unwrap()).unwrap();
    assert!(
        content_types_xml.contains(
            r#"<ct:Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#
        ),
        "expected Sheet1 override to preserve prefix, got: {content_types_xml}"
    );
    assert!(
        content_types_xml.contains(
            r#"<ct:Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#
        ),
        "expected Sheet2 override to preserve prefix, got: {content_types_xml}"
    );
    assert!(
        !content_types_xml.contains(r#"<Override PartName="/xl/worksheets/sheet2.xml""#),
        "writer introduced unprefixed <Override>, got: {content_types_xml}"
    );

    validate_opc_relationships(pkg.parts_map()).expect("validate relationships");
}

