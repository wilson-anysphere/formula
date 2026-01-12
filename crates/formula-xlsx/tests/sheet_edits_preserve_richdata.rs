use std::io::{Cursor, Read, Write};

use formula_xlsx::load_from_bytes;
use pretty_assertions::assert_eq;
use zip::write::FileOptions;
use zip::{ZipArchive, ZipWriter};

fn build_fixture_xlsx() -> Vec<u8> {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Sheet2" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>
"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata" Target="metadata.xml"/>
</Relationships>
"#;

    let root_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"#;

    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/metadata.xml" ContentType="application/vnd.ms-excel.metadata+xml"/>
  <Override PartName="/xl/richData/richValueTypes.xml" ContentType="application/vnd.ms-excel.richValueTypes+xml"/>
  <Override PartName="/xl/richData/richValues.xml" ContentType="application/vnd.ms-excel.richValues+xml"/>
</Types>
"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>
"#;

    let metadata_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <metadataTypes count="1">
    <metadataType name="XLD" minSupportedVersion="0"/>
  </metadataTypes>
</metadata>
"#;

    let metadata_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2020/relationships/richValueTypes" Target="../richData/richValueTypes.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2020/relationships/richValues" Target="../richData/richValues.xml"/>
</Relationships>
"#;

    let rich_value_types_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvTypes xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <rvType name="ExampleType"/>
</rvTypes>
"#;

    let rich_values_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <rv value="ExampleValue"/>
</rvData>
"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", options)
        .expect("zip file");
    zip.write_all(content_types.as_bytes())
        .expect("zip write");

    zip.start_file("_rels/.rels", options).expect("zip file");
    zip.write_all(root_rels.as_bytes()).expect("zip write");

    zip.start_file("xl/workbook.xml", options)
        .expect("zip file");
    zip.write_all(workbook_xml.as_bytes()).expect("zip write");

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .expect("zip file");
    zip.write_all(workbook_rels.as_bytes()).expect("zip write");

    zip.start_file("xl/worksheets/sheet1.xml", options)
        .expect("zip file");
    zip.write_all(sheet_xml.as_bytes()).expect("zip write");

    zip.start_file("xl/worksheets/sheet2.xml", options)
        .expect("zip file");
    zip.write_all(sheet_xml.as_bytes()).expect("zip write");

    zip.start_file("xl/metadata.xml", options).expect("zip file");
    zip.write_all(metadata_xml.as_bytes()).expect("zip write");

    zip.start_file("xl/_rels/metadata.xml.rels", options)
        .expect("zip file");
    zip.write_all(metadata_rels.as_bytes()).expect("zip write");

    zip.start_file("xl/richData/richValueTypes.xml", options)
        .expect("zip file");
    zip.write_all(rich_value_types_xml.as_bytes())
        .expect("zip write");

    zip.start_file("xl/richData/richValues.xml", options)
        .expect("zip file");
    zip.write_all(rich_values_xml.as_bytes())
        .expect("zip write");

    zip.finish().expect("finish zip").into_inner()
}

fn zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).expect("read part");
    buf
}

#[test]
fn sheet_edits_preserve_richdata_parts_and_relationships() {
    let fixture = build_fixture_xlsx();

    let original_metadata = zip_part(&fixture, "xl/metadata.xml");
    let original_metadata_rels = zip_part(&fixture, "xl/_rels/metadata.xml.rels");
    let original_rich_value_types = zip_part(&fixture, "xl/richData/richValueTypes.xml");
    let original_rich_values = zip_part(&fixture, "xl/richData/richValues.xml");

    let mut doc = load_from_bytes(&fixture).expect("load fixture");
    assert_eq!(doc.workbook.sheets.len(), 2);

    let sheet2_id = doc.workbook.sheets[1].id;
    doc.workbook.delete_sheet(sheet2_id).expect("delete sheet2");
    doc.workbook.add_sheet("Added").expect("add sheet");

    let saved = doc.save_to_vec().expect("save");

    // Sanity: ensure we actually exercised the sheet-structure rewrite path.
    let workbook_xml = String::from_utf8(zip_part(&saved, "xl/workbook.xml")).expect("utf8");
    assert!(
        !workbook_xml.contains("Sheet2"),
        "expected deleted sheet to be removed from xl/workbook.xml"
    );
    assert!(
        workbook_xml.contains("Added"),
        "expected newly-added sheet to be present in xl/workbook.xml"
    );
    let cursor = Cursor::new(&saved);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    assert!(
        archive.by_name("xl/worksheets/sheet2.xml").is_err(),
        "expected deleted sheet2 part to be removed from output package"
    );
    assert!(
        archive.by_name("xl/worksheets/sheet3.xml").is_ok(),
        "expected newly-added sheet to be materialized as sheet3.xml"
    );

    assert_eq!(
        zip_part(&saved, "xl/metadata.xml"),
        original_metadata,
        "xl/metadata.xml must be preserved byte-for-byte across sheet edits"
    );
    assert_eq!(
        zip_part(&saved, "xl/_rels/metadata.xml.rels"),
        original_metadata_rels,
        "xl/_rels/metadata.xml.rels must be preserved byte-for-byte across sheet edits"
    );
    assert_eq!(
        zip_part(&saved, "xl/richData/richValueTypes.xml"),
        original_rich_value_types,
        "xl/richData/richValueTypes.xml must be preserved byte-for-byte across sheet edits"
    );
    assert_eq!(
        zip_part(&saved, "xl/richData/richValues.xml"),
        original_rich_values,
        "xl/richData/richValues.xml must be preserved byte-for-byte across sheet edits"
    );

    let workbook_rels = String::from_utf8(zip_part(&saved, "xl/_rels/workbook.xml.rels"))
        .expect("workbook.xml.rels utf8");
    assert!(
        workbook_rels.contains(r#"Id="rId9""#) && workbook_rels.contains(r#"Target="metadata.xml""#),
        "workbook.xml.rels must retain the metadata relationship (rId9 -> metadata.xml)"
    );

    let content_types =
        String::from_utf8(zip_part(&saved, "[Content_Types].xml")).expect("content types utf8");
    assert!(
        content_types.contains(r#"/xl/metadata.xml"#),
        "[Content_Types].xml must retain override for /xl/metadata.xml"
    );
    assert!(
        content_types.contains(r#"/xl/richData/richValueTypes.xml"#),
        "[Content_Types].xml must retain override for /xl/richData/richValueTypes.xml"
    );
    assert!(
        content_types.contains(r#"/xl/richData/richValues.xml"#),
        "[Content_Types].xml must retain override for /xl/richData/richValues.xml"
    );
}
