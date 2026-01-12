use std::io::{Cursor, Read, Write};

use formula_model::{CellRef, CellValue};
use formula_xlsx::load_from_bytes;
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

struct RichDataFixture {
    bytes: Vec<u8>,
    metadata_xml: Vec<u8>,
    metadata_rels_xml: Vec<u8>,
    rich_value_types_xml: Vec<u8>,
    rich_values_xml: Vec<u8>,
}

fn build_richdata_fixture_xlsx() -> RichDataFixture {
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
  <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata" Target="metadata.xml"/>
</Relationships>"#;

    let root_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.metadata+xml"/>
  <Override PartName="/xl/richData/richValueTypes.xml" ContentType="application/vnd.ms-excel.richvaluetypes+xml"/>
  <Override PartName="/xl/richData/richValues.xml" ContentType="application/vnd.ms-excel.richvalues+xml"/>
</Types>"#;

    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Hello</t></is></c>
    </row>
  </sheetData>
</worksheet>"#;

    let metadata_xml = br#"<metadata stable="true">rich-data-root</metadata>"#.to_vec();
    let rich_value_types_xml = br#"<richValueTypes>types</richValueTypes>"#.to_vec();
    let rich_values_xml = br#"<richValues>values</richValues>"#.to_vec();
    let metadata_rels_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2020/02/relationships/richValueTypes" Target="../richData/richValueTypes.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2020/02/relationships/richValues" Target="../richData/richValues.xml"/>
</Relationships>"#
        .to_vec();

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(root_rels.as_bytes()).unwrap();

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types.as_bytes()).unwrap();

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(worksheet_xml.as_bytes()).unwrap();

    zip.start_file("xl/metadata.xml", options).unwrap();
    zip.write_all(&metadata_xml).unwrap();

    zip.start_file("xl/_rels/metadata.xml.rels", options)
        .unwrap();
    zip.write_all(&metadata_rels_xml).unwrap();

    zip.start_file("xl/richData/richValueTypes.xml", options)
        .unwrap();
    zip.write_all(&rich_value_types_xml).unwrap();

    zip.start_file("xl/richData/richValues.xml", options)
        .unwrap();
    zip.write_all(&rich_values_xml).unwrap();

    let bytes = zip.finish().unwrap().into_inner();

    RichDataFixture {
        bytes,
        metadata_xml,
        metadata_rels_xml,
        rich_value_types_xml,
        rich_values_xml,
    }
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
fn richdata_parts_are_preserved_when_writer_adds_shared_strings() -> Result<(), Box<dyn std::error::Error>>
{
    let fixture = build_richdata_fixture_xlsx();

    let mut doc = load_from_bytes(&fixture.bytes)?;

    // Force the writer to synthesize a sharedStrings part + metadata by introducing a new string
    // value that lacks preserved cell metadata.
    let sheet_id = doc.workbook.sheets[0].id;
    let sheet = doc.workbook.sheet_mut(sheet_id).expect("sheet exists");
    sheet.set_value(
        CellRef::from_a1("B1")?,
        CellValue::String("World".to_string()),
    );

    let saved = doc.save_to_vec()?;

    // Rich data parts should be preserved byte-for-byte.
    assert_eq!(zip_part(&saved, "xl/metadata.xml"), fixture.metadata_xml);
    assert_eq!(
        zip_part(&saved, "xl/_rels/metadata.xml.rels"),
        fixture.metadata_rels_xml
    );
    assert_eq!(
        zip_part(&saved, "xl/richData/richValueTypes.xml"),
        fixture.rich_value_types_xml
    );
    assert_eq!(
        zip_part(&saved, "xl/richData/richValues.xml"),
        fixture.rich_values_xml
    );

    // Workbook relationship list should retain the original metadata relationship.
    let workbook_rels = String::from_utf8(zip_part(&saved, "xl/_rels/workbook.xml.rels"))?;
    assert!(workbook_rels.contains(r#"Id="rId9""#));
    assert!(workbook_rels.contains(r#"Target="metadata.xml""#));

    // Content types should retain rich-data overrides and gain a sharedStrings override.
    let content_types = String::from_utf8(zip_part(&saved, "[Content_Types].xml"))?;
    assert!(content_types.contains(r#"/xl/metadata.xml"#));
    assert!(content_types.contains(r#"/xl/richData/richValueTypes.xml"#));
    assert!(content_types.contains(r#"/xl/richData/richValues.xml"#));
    assert!(
        content_types.contains(r#"/xl/sharedStrings.xml"#),
        "writer should have added a sharedStrings content-types override"
    );

    Ok(())
}

