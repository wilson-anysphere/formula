use std::io::{Cursor, Write};

use formula_model::CellRef;
use formula_xlsx::load_from_bytes;
use zip::write::FileOptions;
use zip::ZipWriter;

fn build_rich_cell_image_fixture() -> Vec<u8> {
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
</Relationships>
"#;

    let sheet1_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" vm="0"><v>ignored</v></c>
      <c r="B1"><v>2</v></c>
    </row>
  </sheetData>
</worksheet>
"#;

    // Minimal metadata that maps vm=0 -> rich value index 0 (via "v" fallback logic).
    let metadata_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <valueMetadata count="1">
    <bk>
      <rc t="0" v="0"/>
    </bk>
  </valueMetadata>
</metadata>
"#;

    // Rich value index 0 references relationship index 0.
    let rich_value_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValue xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <rv>
    <v t="rel">0</v>
  </rv>
</richValue>
"#;

    let rich_value_rel_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueRel xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rel r:id="rId1"/>
</richValueRel>
"#;

    let rich_value_rel_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png#fragment"/>
</Relationships>
"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options).unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(sheet1_xml.as_bytes()).unwrap();

    zip.start_file("xl/metadata.xml", options).unwrap();
    zip.write_all(metadata_xml.as_bytes()).unwrap();

    zip.start_file("xl/richData/richValue.xml", options).unwrap();
    zip.write_all(rich_value_xml.as_bytes()).unwrap();

    zip.start_file("xl/richData/richValueRel.xml", options).unwrap();
    zip.write_all(rich_value_rel_xml.as_bytes()).unwrap();

    zip.start_file("xl/richData/_rels/richValueRel.xml.rels", options)
        .unwrap();
    zip.write_all(rich_value_rel_rels.as_bytes()).unwrap();

    zip.start_file("xl/media/image1.png", options).unwrap();
    zip.write_all(b"fakepng").unwrap();

    zip.finish().unwrap().into_inner()
}

#[test]
fn resolves_image_target_for_rich_value_cell() -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_rich_cell_image_fixture();
    let doc = load_from_bytes(&bytes)?;
    let sheet_id = doc.workbook.sheets[0].id;

    assert_eq!(
        doc.image_target_for_cell(sheet_id, CellRef::from_a1("A1")?)?,
        Some("xl/media/image1.png".to_string())
    );
    assert_eq!(
        doc.image_target_for_cell(sheet_id, CellRef::from_a1("B1")?)?,
        None
    );

    Ok(())
}
