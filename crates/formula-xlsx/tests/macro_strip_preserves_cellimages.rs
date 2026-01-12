use std::io::{Cursor, Write};

use base64::Engine;
use formula_xlsx::XlsxPackage;
use roxmltree::Document;

fn build_fixture() -> (Vec<u8>, Vec<u8>, Vec<u8>, Vec<u8>) {
    // 1x1 transparent PNG.
    let png_bytes = base64::engine::general_purpose::STANDARD
        .decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/58HAQUBAO3+2NoAAAAASUVORK5CYII=")
        .expect("valid base64 png");

    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/cellimages.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.cellimages+xml"/>
  <Override PartName="/xl/controls/control1.xml" ContentType="application/vnd.ms-excel.control+xml"/>
  <Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>
</Types>"#;

    let root_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2020/relationships/cellImage" Target="cellimages.xml"/>
  <Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>
</Relationships>"#;

    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;

    let worksheet_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/control" Target="../controls/control1.xml"/>
</Relationships>"#;

    let cellimages_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<etc:cellImages xmlns:etc="http://schemas.microsoft.com/office/spreadsheetml/2022/cellimages"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <etc:cellImage r:id="rId1"/>
</etc:cellImages>"#
        .to_vec();

    // Intentionally use `../media/*` which some producers emit for workbook-level parts; macro
    // stripping must still understand that this points at `xl/media/*` and must not delete the
    // shared image when the macro/control part is removed.
    let cellimages_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>"#
        .to_vec();

    let control_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<control xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" name="Button1"/>"#;

    let control_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>"#;

    let empty_rels =
        br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options =
        zip::write::FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    fn add_file(
        zip: &mut zip::ZipWriter<Cursor<Vec<u8>>>,
        options: zip::write::FileOptions<()>,
        name: &str,
        bytes: &[u8],
    ) {
        zip.start_file(name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    add_file(&mut zip, options, "[Content_Types].xml", content_types.as_bytes());
    add_file(&mut zip, options, "_rels/.rels", root_rels.as_bytes());
    add_file(&mut zip, options, "xl/workbook.xml", workbook_xml.as_bytes());
    add_file(
        &mut zip,
        options,
        "xl/_rels/workbook.xml.rels",
        workbook_rels.as_bytes(),
    );
    add_file(&mut zip, options, "xl/worksheets/sheet1.xml", worksheet_xml.as_bytes());
    add_file(
        &mut zip,
        options,
        "xl/worksheets/_rels/sheet1.xml.rels",
        worksheet_rels.as_bytes(),
    );

    add_file(&mut zip, options, "xl/cellimages.xml", &cellimages_xml);
    add_file(
        &mut zip,
        options,
        "xl/_rels/cellimages.xml.rels",
        &cellimages_rels,
    );

    add_file(&mut zip, options, "xl/vbaProject.bin", b"dummy-vba");
    add_file(&mut zip, options, "xl/_rels/vbaProject.bin.rels", empty_rels);

    add_file(&mut zip, options, "xl/controls/control1.xml", control_xml.as_bytes());
    add_file(
        &mut zip,
        options,
        "xl/controls/_rels/control1.xml.rels",
        control_rels.as_bytes(),
    );

    add_file(&mut zip, options, "xl/media/image1.png", &png_bytes);

    (
        zip.finish().unwrap().into_inner(),
        cellimages_xml,
        cellimages_rels,
        png_bytes,
    )
}

#[test]
fn macro_stripping_preserves_cellimages_and_shared_media() {
    let (fixture, cellimages_xml, cellimages_rels, png_bytes) = build_fixture();

    let mut pkg = XlsxPackage::from_bytes(&fixture).expect("read fixture");
    pkg.remove_vba_project().expect("strip macros");

    let written = pkg.write_to_bytes().expect("write stripped package");
    let pkg2 = XlsxPackage::from_bytes(&written).expect("read stripped package");

    // Macro parts removed.
    assert!(pkg2.part("xl/vbaProject.bin").is_none());
    assert!(pkg2.part("xl/controls/control1.xml").is_none());
    assert!(pkg2.part("xl/controls/_rels/control1.xml.rels").is_none());

    // Cellimages preserved byte-for-byte (plus referenced media).
    assert_eq!(pkg2.part("xl/cellimages.xml").unwrap(), cellimages_xml.as_slice());
    assert_eq!(
        pkg2.part("xl/_rels/cellimages.xml.rels").unwrap(),
        cellimages_rels.as_slice()
    );
    assert_eq!(pkg2.part("xl/media/image1.png").unwrap(), png_bytes.as_slice());

    // Workbook/package metadata cleaned but not destructive.
    let ct_xml =
        std::str::from_utf8(pkg2.part("[Content_Types].xml").unwrap()).expect("content types utf8");
    let ct_doc = Document::parse(ct_xml).expect("parse [Content_Types].xml");
    let overrides: Vec<_> = ct_doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "Override")
        .collect();
    assert!(
        overrides
            .iter()
            .all(|n| n.attribute("PartName") != Some("/xl/vbaProject.bin")),
        "expected [Content_Types].xml to drop /xl/vbaProject.bin override (got {ct_xml:?})"
    );
    assert!(
        overrides
            .iter()
            .any(|n| n.attribute("PartName") == Some("/xl/cellimages.xml")),
        "expected [Content_Types].xml to keep /xl/cellimages.xml override (got {ct_xml:?})"
    );

    let wb_rels_xml = std::str::from_utf8(pkg2.part("xl/_rels/workbook.xml.rels").unwrap())
        .expect("workbook rels utf8");
    let rels_doc = Document::parse(wb_rels_xml).expect("parse workbook.xml.rels");
    let rels: Vec<_> = rels_doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "Relationship")
        .collect();
    assert!(
        rels.iter().any(|n| n.attribute("Target") == Some("cellimages.xml")),
        "expected workbook.xml.rels to keep cellimages relationship (got {wb_rels_xml:?})"
    );
    assert!(
        rels.iter().all(|n| n.attribute("Target") != Some("vbaProject.bin")),
        "expected workbook.xml.rels to drop vbaProject relationship (got {wb_rels_xml:?})"
    );

    let sheet_rels_xml =
        std::str::from_utf8(pkg2.part("xl/worksheets/_rels/sheet1.xml.rels").unwrap())
            .expect("sheet rels utf8");
    assert!(
        !sheet_rels_xml.contains("controls/control1.xml"),
        "expected sheet relationships to stop referencing deleted controls (got {sheet_rels_xml:?})"
    );
}

