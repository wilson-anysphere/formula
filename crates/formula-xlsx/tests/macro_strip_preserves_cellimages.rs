use std::io::{Cursor, Write};

use formula_xlsx::XlsxPackage;

#[test]
fn macro_strip_preserves_cellimages_part_and_shared_image_bytes() {
    let fixture = build_macro_control_cellimages_fixture();
    let mut pkg = XlsxPackage::from_bytes(&fixture).expect("read fixture");

    let original_cellimages_xml = pkg
        .part("xl/cellimages.xml")
        .expect("fixture has xl/cellimages.xml")
        .to_vec();
    let original_cellimages_rels = pkg
        .part("xl/_rels/cellimages.xml.rels")
        .expect("fixture has xl/_rels/cellimages.xml.rels")
        .to_vec();
    let original_image_bytes = pkg
        .part("xl/media/image1.png")
        .expect("fixture has xl/media/image1.png")
        .to_vec();

    pkg.remove_vba_project().expect("strip macros/controls");

    let written = pkg.write_to_bytes().expect("write stripped package");
    let pkg2 = XlsxPackage::from_bytes(&written).expect("read stripped package");

    assert!(
        pkg2.part("xl/vbaProject.bin").is_none(),
        "expected vbaProject.bin to be removed"
    );
    assert!(
        pkg2.part("xl/controls/control1.xml").is_none(),
        "expected controls/control1.xml to be removed"
    );
    assert!(
        pkg2.part("xl/controls/_rels/control1.xml.rels").is_none(),
        "expected controls/_rels/control1.xml.rels to be removed"
    );

    assert_eq!(
        pkg2.part("xl/cellimages.xml")
            .expect("cellimages.xml preserved"),
        original_cellimages_xml.as_slice(),
        "expected cellimages.xml to be preserved byte-for-byte"
    );
    assert_eq!(
        pkg2.part("xl/_rels/cellimages.xml.rels")
            .expect("cellimages.xml.rels preserved"),
        original_cellimages_rels.as_slice(),
        "expected cellimages.xml.rels to be preserved byte-for-byte"
    );

    assert_eq!(
        pkg2.part("xl/media/image1.png")
            .expect("shared image preserved"),
        original_image_bytes.as_slice(),
        "expected shared image payload to be preserved byte-for-byte"
    );

    let sheet_rels = std::str::from_utf8(pkg2.part("xl/worksheets/_rels/sheet1.xml.rels").unwrap())
        .expect("sheet rels utf-8");
    assert!(
        !sheet_rels.contains("controls/control1.xml"),
        "expected sheet relationships to stop referencing deleted controls (got {sheet_rels:?})"
    );
}

fn build_macro_control_cellimages_fixture() -> Vec<u8> {
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
  <Override PartName="/xl/cellimages.xml" ContentType="application/vnd.ms-excel.cellimages+xml"/>
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
  <Relationship Id="rId3"
    Type="http://schemas.microsoft.com/office/2017/06/relationships/cellimages"
    Target="cellimages.xml"/>
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

    let cellimages_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cellImages xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/06/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <cellImage r:id="rId1"/>
</cellImages>"#;

    let cellimages_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    Target="media/image1.png"/>
</Relationships>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    fn add_file(
        zip: &mut zip::ZipWriter<Cursor<Vec<u8>>>,
        options: zip::write::FileOptions<()>,
        name: &str,
        bytes: &[u8],
    ) {
        zip.start_file(name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    add_file(
        &mut zip,
        options,
        "[Content_Types].xml",
        content_types.as_bytes(),
    );
    add_file(&mut zip, options, "_rels/.rels", root_rels.as_bytes());
    add_file(
        &mut zip,
        options,
        "xl/workbook.xml",
        workbook_xml.as_bytes(),
    );
    add_file(
        &mut zip,
        options,
        "xl/_rels/workbook.xml.rels",
        workbook_rels.as_bytes(),
    );
    add_file(
        &mut zip,
        options,
        "xl/worksheets/sheet1.xml",
        worksheet_xml.as_bytes(),
    );
    add_file(
        &mut zip,
        options,
        "xl/worksheets/_rels/sheet1.xml.rels",
        sheet_rels.as_bytes(),
    );

    add_file(&mut zip, options, "xl/vbaProject.bin", b"dummy-vba");
    add_file(
        &mut zip,
        options,
        "xl/_rels/vbaProject.bin.rels",
        br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#,
    );

    add_file(
        &mut zip,
        options,
        "xl/controls/control1.xml",
        control_xml.as_bytes(),
    );
    add_file(
        &mut zip,
        options,
        "xl/controls/_rels/control1.xml.rels",
        control_rels.as_bytes(),
    );

    add_file(
        &mut zip,
        options,
        "xl/cellimages.xml",
        cellimages_xml.as_bytes(),
    );
    add_file(
        &mut zip,
        options,
        "xl/_rels/cellimages.xml.rels",
        cellimages_rels.as_bytes(),
    );

    add_file(&mut zip, options, "xl/media/image1.png", b"not-a-real-png");

    zip.finish().unwrap().into_inner()
}
