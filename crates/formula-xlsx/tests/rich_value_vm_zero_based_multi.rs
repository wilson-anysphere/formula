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

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata" Target="metadata.xml"/>
</Relationships>
"#;

    // Two rich values, referenced by valueMetadata via `rc/@v` indices into futureMetadata.
    let metadata_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <metadataTypes count="1">
    <metadataType name="XLRICHVALUE"/>
  </metadataTypes>
  <futureMetadata name="XLRICHVALUE" count="2">
    <bk>
      <extLst>
        <ext uri="{00000000-0000-0000-0000-000000000000}">
          <xlrd:rvb i="5"/>
        </ext>
      </extLst>
    </bk>
    <bk>
      <extLst>
        <ext uri="{00000000-0000-0000-0000-000000000001}">
          <xlrd:rvb i="42"/>
        </ext>
      </extLst>
    </bk>
  </futureMetadata>
  <valueMetadata count="2">
    <bk><rc t="1" v="0"/></bk>
    <bk><rc t="1" v="1"/></bk>
  </valueMetadata>
</metadata>
"#;

    // `vm` is intentionally 0-based (vm="0" selects the first <valueMetadata><bk>).
    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" vm="0"><v>0</v></c>
    </row>
    <row r="2">
      <c r="A2" vm="1"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>
"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/metadata.xml", options).unwrap();
    zip.write_all(metadata_xml.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options)
        .unwrap();
    zip.write_all(sheet_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

#[test]
fn loads_rich_value_indices_from_zero_based_vm_with_multiple_entries() -> Result<(), Box<dyn std::error::Error>>
{
    let fixture = build_fixture_xlsx();
    let doc = formula_xlsx::load_from_bytes(&fixture)?;

    let sheet_id = doc.workbook.sheets[0].id;
    assert_eq!(
        doc.rich_value_index(sheet_id, CellRef::from_a1("A1")?),
        Some(5)
    );
    assert_eq!(
        doc.rich_value_index(sheet_id, CellRef::from_a1("A2")?),
        Some(42)
    );

    Ok(())
}

