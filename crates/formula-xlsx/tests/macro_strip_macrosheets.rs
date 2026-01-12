use std::io::{Cursor, Write};

use formula_xlsx::{validate_opc_relationships, XlsxPackage};

fn build_macrosheet_fixture() -> Vec<u8> {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/macrosheets/sheet2.xml" ContentType="application/vnd.ms-excel.macrosheet+xml"/>
  <Override PartName="/xl/dialogsheets/sheet3.xml" ContentType="application/vnd.ms-excel.dialogsheet+xml"/>
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
    <sheet name="MacroSheet" sheetId="2" r:id="rId2"/>
    <sheet name="DialogSheet" sheetId="3" r:id="rId3"/>
  </sheets>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet"
    Target="macrosheets/sheet2.xml"/>
  <Relationship Id="rId3"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet"
    Target="dialogsheets/sheet3.xml"/>
  <Relationship Id="rId4"
    Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject"
    Target="vbaProject.bin"/>
</Relationships>"#;

    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;

    let macro_sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<macroSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;

    let dialog_sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dialogsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;

    let empty_rels = br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#;

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
    add_file(&mut zip, options, "xl/macrosheets/sheet2.xml", macro_sheet_xml.as_bytes());
    add_file(&mut zip, options, "xl/dialogsheets/sheet3.xml", dialog_sheet_xml.as_bytes());

    // Include nested relationship parts so macro stripping needs to delete them as well.
    add_file(
        &mut zip,
        options,
        "xl/macrosheets/_rels/sheet2.xml.rels",
        empty_rels,
    );
    add_file(
        &mut zip,
        options,
        "xl/dialogsheets/_rels/sheet3.xml.rels",
        empty_rels,
    );
    add_file(&mut zip, options, "xl/vbaProject.bin", b"dummy-vba");
    add_file(&mut zip, options, "xl/_rels/vbaProject.bin.rels", empty_rels);

    zip.finish().unwrap().into_inner()
}

#[test]
fn macro_stripping_removes_macrosheets_and_dialogsheets() {
    let fixture = build_macrosheet_fixture();
    let mut pkg = XlsxPackage::from_bytes(&fixture).expect("read fixture");

    pkg.remove_vba_project().expect("strip macros");

    let written = pkg.write_to_bytes().expect("write stripped package");
    let pkg2 = XlsxPackage::from_bytes(&written).expect("read stripped package");

    assert!(pkg2.part("xl/vbaProject.bin").is_none());
    assert!(pkg2.part("xl/macrosheets/sheet2.xml").is_none());
    assert!(pkg2.part("xl/dialogsheets/sheet3.xml").is_none());
    assert!(pkg2.part("xl/macrosheets/_rels/sheet2.xml.rels").is_none());
    assert!(pkg2.part("xl/dialogsheets/_rels/sheet3.xml.rels").is_none());

    let workbook_xml = std::str::from_utf8(pkg2.part("xl/workbook.xml").unwrap())
        .expect("workbook xml utf-8");
    assert!(
        !workbook_xml.contains(r#"name="MacroSheet""#),
        "expected workbook.xml to drop macro sheet entry (got {workbook_xml:?})"
    );
    assert!(
        !workbook_xml.contains(r#"name="DialogSheet""#),
        "expected workbook.xml to drop dialog sheet entry (got {workbook_xml:?})"
    );
    assert!(
        !workbook_xml.contains(r#"r:id="rId2""#) && !workbook_xml.contains(r#"r:id="rId3""#),
        "expected workbook.xml to drop dangling r:ids (got {workbook_xml:?})"
    );

    let workbook_rels = std::str::from_utf8(pkg2.part("xl/_rels/workbook.xml.rels").unwrap())
        .expect("workbook rels utf-8");
    assert!(
        !workbook_rels.contains("macrosheets/"),
        "expected workbook rels to stop referencing macrosheets (got {workbook_rels:?})"
    );
    assert!(
        !workbook_rels.contains("dialogsheets/"),
        "expected workbook rels to stop referencing dialogsheets (got {workbook_rels:?})"
    );

    let content_types = std::str::from_utf8(pkg2.part("[Content_Types].xml").unwrap())
        .expect("content types utf-8");
    assert!(
        !content_types.contains("macroEnabled.main+xml"),
        "expected workbook content type to be downgraded to .xlsx (got {content_types:?})"
    );
    assert!(!content_types.contains("/xl/vbaProject.bin"));
    assert!(!content_types.contains("/xl/macrosheets/sheet2.xml"));
    assert!(!content_types.contains("/xl/dialogsheets/sheet3.xml"));

    // Ensure we didn't leave any dangling relationship targets or relationship id references.
    validate_opc_relationships(pkg2.parts_map()).expect("stripped package relationships are consistent");
}

