use std::io::{Cursor, Write};

use formula_model::CellRef;
use zip::write::FileOptions;
use zip::ZipWriter;

fn build_fixture_xlsx() -> Vec<u8> {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
"#;

    // Include both relationship types so we can assert the loader prefers the Excel-standard
    // `sheetMetadata` relationship when both are present.
    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata" Target="metadata.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="custom-metadata.xml"/>
</Relationships>
"#;

    // Minimal metadata part that maps `vm` -> rich value index directly via `rc/@v`.
    let metadata_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <valueMetadata count="1">
    <bk>
      <rc t="0" v="7"/>
    </bk>
  </valueMetadata>
</metadata>
"#;

    // `vm="1"` here is intentionally 1-based to ensure the parser handles index base ambiguity.
    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" vm="1"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>
"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/workbook.xml", options)
        .expect("zip file");
    zip.write_all(workbook_xml.as_bytes()).expect("zip write");

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .expect("zip file");
    zip.write_all(workbook_rels.as_bytes()).expect("zip write");

    zip.start_file("xl/custom-metadata.xml", options)
        .expect("zip file");
    zip.write_all(metadata_xml.as_bytes()).expect("zip write");

    zip.start_file("xl/worksheets/sheet1.xml", options)
        .expect("zip file");
    zip.write_all(sheet_xml.as_bytes()).expect("zip write");

    zip.finish().expect("finish zip").into_inner()
}

#[test]
fn loads_metadata_part_from_sheetmetadata_relationship_with_custom_target(
) -> Result<(), Box<dyn std::error::Error>> {
    let fixture = build_fixture_xlsx();
    let doc = formula_xlsx::load_from_bytes(&fixture)?;

    let sheet_id = doc.workbook.sheets[0].id;
    assert_eq!(
        doc.rich_value_index(sheet_id, CellRef::from_a1("A1")?),
        Some(7)
    );

    Ok(())
}

