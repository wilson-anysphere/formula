use std::io::{Cursor, Read, Write};

use zip::ZipArchive;

fn read_zip_part(bytes: &[u8], name: &str) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let cursor = Cursor::new(bytes);
    let mut zip = ZipArchive::new(cursor)?;
    let mut file = zip.by_name(name)?;
    let mut out = Vec::new();
    file.read_to_end(&mut out)?;
    Ok(out)
}

fn build_synthetic_xlsx_with_rich_data() -> (Vec<u8>, Vec<u8>, Vec<u8>) {
    // Intentionally include a few newer/unrecognized parts that Formula does not currently parse.
    // The key contract we want to enforce is that `XlsxDocument` preserves unknown parts
    // byte-for-byte during a no-op load/save.

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
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata"
    Target="metadata.xml"/>
</Relationships>"#;

    // `vm="1"` is used by Excel to reference value metadata. It's optional for this test, but it
    // mirrors real rich-data workbooks.
    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1">
      <c r="A1" vm="1"><v>1</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    let metadata_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <metadataTypes count="0"/>
</metadata>
"#;

    let rich_data_part_name = "xl/richData/richValueRel.xml";
    let rich_data_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvRel xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <info>test</info>
</rvRel>
"#;

    // Include content-type overrides for both the metadata and richData parts. The exact values
    // aren't important for this regression test; the goal is to ensure the writer does not drop
    // these parts when rewriting core workbook/package files.
    let content_types_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.metadata+xml"/>
  <Override PartName="/{rich_data_part_name}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.richData+xml"/>
</Types>"#
    );

    let root_rels_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="xl/workbook.xml"/>
</Relationships>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types_xml.as_bytes()).unwrap();

    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(root_rels_xml.as_bytes()).unwrap();

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels_xml.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(worksheet_xml.as_bytes()).unwrap();

    zip.start_file("xl/metadata.xml", options).unwrap();
    zip.write_all(metadata_xml.as_bytes()).unwrap();

    zip.start_file(rich_data_part_name, options).unwrap();
    zip.write_all(rich_data_xml.as_bytes()).unwrap();

    let bytes = zip.finish().unwrap().into_inner();

    (
        bytes,
        metadata_xml.as_bytes().to_vec(),
        rich_data_xml.as_bytes().to_vec(),
    )
}

#[test]
fn noop_roundtrip_preserves_metadata_and_richdata_parts_and_workbook_relationship(
) -> Result<(), Box<dyn std::error::Error>> {
    let (input_bytes, original_metadata, original_rich_data) =
        build_synthetic_xlsx_with_rich_data();

    let doc = formula_xlsx::load_from_bytes(&input_bytes)?;
    let output_bytes = doc.save_to_vec()?;

    let out_metadata = read_zip_part(&output_bytes, "xl/metadata.xml")?;
    assert_eq!(
        out_metadata, original_metadata,
        "expected xl/metadata.xml to be preserved byte-for-byte"
    );

    let out_rich_data = read_zip_part(&output_bytes, "xl/richData/richValueRel.xml")?;
    assert_eq!(
        out_rich_data, original_rich_data,
        "expected xl/richData/* part to be preserved byte-for-byte"
    );

    let out_rels = String::from_utf8(read_zip_part(&output_bytes, "xl/_rels/workbook.xml.rels")?)?;
    let rels_doc = roxmltree::Document::parse(&out_rels)?;
    let has_metadata_rel = rels_doc.descendants().any(|n| {
        n.is_element()
            && n.tag_name().name() == "Relationship"
            && n.attribute("Target") == Some("metadata.xml")
    });
    assert!(
        has_metadata_rel,
        "expected xl/_rels/workbook.xml.rels to contain a relationship targeting metadata.xml, got:\n{out_rels}"
    );

    Ok(())
}

