use std::io::{Cursor, Write};

use formula_xlsx::XlsxPackage;

fn build_macro_control_fixture(leading_slash_entries: bool) -> Vec<u8> {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/controls/control1.xml" ContentType="application/vnd.ms-excel.control+xml"/>
  <Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>
</Types>"#;

    let root_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="xl/workbook.xml"/>
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
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject"
    Target="vbaProject.bin"/>
</Relationships>"#;

    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;

    let sheet_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/control"
    Target="../controls/control1.xml"/>
</Relationships>"#;

    let control_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<control xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" name="Button1"/>"#;

    let control_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    Target="../media/image1.png"/>
</Relationships>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    fn add_file<S: ToString>(
        zip: &mut zip::ZipWriter<Cursor<Vec<u8>>>,
        options: zip::write::FileOptions<()>,
        name: S,
        bytes: &[u8],
    ) {
        zip.start_file(name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    add_file(&mut zip, options, "[Content_Types].xml", content_types.as_bytes());
    add_file(&mut zip, options, "_rels/.rels", root_rels.as_bytes());
    let xl_part_name = |name: &str| {
        if leading_slash_entries {
            format!("/{name}")
        } else {
            name.to_string()
        }
    };

    add_file(&mut zip, options, xl_part_name("xl/workbook.xml"), workbook_xml.as_bytes());
    add_file(
        &mut zip,
        options,
        xl_part_name("xl/_rels/workbook.xml.rels"),
        workbook_rels.as_bytes(),
    );
    add_file(
        &mut zip,
        options,
        xl_part_name("xl/worksheets/sheet1.xml"),
        worksheet_xml.as_bytes(),
    );
    add_file(
        &mut zip,
        options,
        xl_part_name("xl/worksheets/_rels/sheet1.xml.rels"),
        sheet_rels.as_bytes(),
    );

    add_file(
        &mut zip,
        options,
        xl_part_name("xl/vbaProject.bin"),
        b"dummy-vba",
    );
    add_file(
        &mut zip,
        options,
        xl_part_name("xl/_rels/vbaProject.bin.rels"),
        br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#,
    );

    add_file(
        &mut zip,
        options,
        xl_part_name("xl/controls/control1.xml"),
        control_xml.as_bytes(),
    );
    add_file(
        &mut zip,
        options,
        xl_part_name("xl/controls/_rels/control1.xml.rels"),
        control_rels.as_bytes(),
    );
    add_file(
        &mut zip,
        options,
        xl_part_name("xl/media/image1.png"),
        b"not-a-real-png",
    );

    zip.finish().unwrap().into_inner()
}

#[test]
fn macro_stripping_removes_controls_parts_and_relationships() {
    let fixture = build_macro_control_fixture(false);
    let mut pkg = XlsxPackage::from_bytes(&fixture).expect("read fixture");

    pkg.remove_vba_project().expect("strip macros/controls");

    let written = pkg.write_to_bytes().expect("write stripped package");
    let pkg2 = XlsxPackage::from_bytes(&written).expect("read stripped package");

    assert!(pkg2.part("xl/vbaProject.bin").is_none());
    assert!(pkg2.part("xl/controls/control1.xml").is_none());
    assert!(pkg2.part("xl/controls/_rels/control1.xml.rels").is_none());

    // Relationship traversal should remove child parts that were only reachable from the deleted
    // control part.
    assert!(pkg2.part("xl/media/image1.png").is_none());

    let sheet_rels = std::str::from_utf8(pkg2.part("xl/worksheets/_rels/sheet1.xml.rels").unwrap())
        .expect("sheet rels utf-8");
    assert!(
        !sheet_rels.contains("controls/control1.xml"),
        "expected sheet rels to stop referencing deleted controls (got {sheet_rels:?})"
    );

    for (name, bytes) in pkg2.parts() {
        if !name.ends_with(".rels") {
            continue;
        }
        let xml = std::str::from_utf8(bytes).expect("rels utf-8");
        assert!(
            !xml.contains("controls/"),
            "{name} still references `xl/controls/**` parts: {xml:?}"
        );
    }

    let ct = std::str::from_utf8(pkg2.part("[Content_Types].xml").unwrap()).expect("ct utf-8");
    assert!(!ct.contains("vbaProject.bin"));
    assert!(!ct.contains("xl/controls/control1.xml"));
    assert!(
        !ct.contains("macroEnabled.main+xml"),
        "expected workbook content type to be downgraded to .xlsx (got {ct:?})"
    );
}

#[test]
fn macro_stripping_removes_controls_parts_and_relationships_with_leading_slash_entries() {
    let fixture = build_macro_control_fixture(true);
    let mut pkg = XlsxPackage::from_bytes(&fixture).expect("read fixture");

    pkg.remove_vba_project().expect("strip macros/controls");

    let written = pkg.write_to_bytes().expect("write stripped package");
    let pkg2 = XlsxPackage::from_bytes(&written).expect("read stripped package");

    assert!(pkg2.part("xl/vbaProject.bin").is_none());
    assert!(pkg2.part("xl/controls/control1.xml").is_none());
    assert!(pkg2.part("xl/controls/_rels/control1.xml.rels").is_none());

    // Relationship traversal should remove child parts that were only reachable from the deleted
    // control part.
    assert!(pkg2.part("xl/media/image1.png").is_none());

    let sheet_rels = std::str::from_utf8(pkg2.part("xl/worksheets/_rels/sheet1.xml.rels").unwrap())
        .expect("sheet rels utf-8");
    assert!(
        !sheet_rels.contains("controls/control1.xml"),
        "expected sheet rels to stop referencing deleted controls (got {sheet_rels:?})"
    );

    for (name, bytes) in pkg2.parts() {
        if !name.ends_with(".rels") {
            continue;
        }
        let xml = std::str::from_utf8(bytes).expect("rels utf-8");
        assert!(
            !xml.contains("controls/"),
            "{name} still references `xl/controls/**` parts: {xml:?}"
        );
    }

    let ct = std::str::from_utf8(pkg2.part("[Content_Types].xml").unwrap()).expect("ct utf-8");
    assert!(!ct.contains("vbaProject.bin"));
    assert!(!ct.contains("xl/controls/control1.xml"));
    assert!(
        !ct.contains("macroEnabled.main+xml"),
        "expected workbook content type to be downgraded to .xlsx (got {ct:?})"
    );
}
