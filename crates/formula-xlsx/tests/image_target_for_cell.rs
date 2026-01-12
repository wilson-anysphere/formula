use std::io::{Cursor, Write};

use formula_model::CellRef;
use formula_xlsx::load_from_bytes;
use zip::write::FileOptions;
use zip::ZipWriter;

fn build_rich_cell_image_fixture_with_rich_value_part_base(rich_value_part_base: &str) -> Vec<u8> {
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
      <c r="C1"><v>3</v></c>
    </row>
  </sheetData>
  <mergeCells count="1">
    <mergeCell ref="A1:B1"/>
  </mergeCells>
</worksheet>
"#;

    // Minimal metadata that maps vm=0 -> rich value index 1 (via "v" fallback logic).
    let metadata_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <valueMetadata count="1">
    <bk>
      <rc t="0" v="1"/>
    </bk>
  </valueMetadata>
</metadata>
"#;

    // Multiple richValue parts ensure we don't rely on lexicographic part ordering
    // (`richValue10.xml` must not be treated as coming before `richValue2.xml`).
    //
    // Rich value index 0 references relationship index 0.
    let rich_value_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValue xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <rv>
    <v t="rel">0</v>
  </rv>
</richValue>
"#;

    // Rich value index 1 references relationship index 1.
    let rich_value2_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValue xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <rv>
    <v t="rel">1</v>
  </rv>
</richValue>
"#;

    // Rich value index 2 references relationship index 2.
    // This part name is intentionally `richValue10.xml` to catch lexicographic ordering bugs.
    let rich_value10_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValue xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <rv>
    <v t="rel">2</v>
  </rv>
</richValue>
"#;

    // Use non-canonical richValueRel part names to ensure we don't rely on lexicographic ordering
    // (`richValueRel10.xml` must not be chosen over `richValueRel2.xml`).
    let rich_value_rel2_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueRel xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rel r:id="rId1"/>
  <rel r:id="rId2"/>
  <rel r:id="rId3"/>
</richValueRel>
"#;

    // Dummy part that should not be selected (wrong suffix ordering).
    let rich_value_rel10_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueRel xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rel r:id="rId999"/>
</richValueRel>
"#;

    let rich_value_rel2_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.png#fragment"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image3.png"/>
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

    zip.start_file(
        format!("xl/richData/{rich_value_part_base}.xml"),
        options,
    )
    .unwrap();
    zip.write_all(rich_value_xml.as_bytes()).unwrap();

    zip.start_file(
        format!("xl/richData/{rich_value_part_base}2.xml"),
        options,
    )
    .unwrap();
    zip.write_all(rich_value2_xml.as_bytes()).unwrap();

    zip.start_file(
        format!("xl/richData/{rich_value_part_base}10.xml"),
        options,
    )
    .unwrap();
    zip.write_all(rich_value10_xml.as_bytes()).unwrap();

    zip.start_file("xl/richData/richValueRel2.xml", options).unwrap();
    zip.write_all(rich_value_rel2_xml.as_bytes()).unwrap();

    zip.start_file("xl/richData/richValueRel10.xml", options).unwrap();
    zip.write_all(rich_value_rel10_xml.as_bytes()).unwrap();

    zip.start_file("xl/richData/_rels/richValueRel2.xml.rels", options)
        .unwrap();
    zip.write_all(rich_value_rel2_rels.as_bytes()).unwrap();

    zip.start_file("xl/media/image1.png", options).unwrap();
    zip.write_all(b"fakepng").unwrap();

    zip.start_file("xl/media/image2.png", options).unwrap();
    zip.write_all(b"fakepng2").unwrap();

    zip.start_file("xl/media/image3.png", options).unwrap();
    zip.write_all(b"fakepng3").unwrap();

    zip.finish().unwrap().into_inner()
}

fn build_rich_cell_image_fixture() -> Vec<u8> {
    build_rich_cell_image_fixture_with_rich_value_part_base("richValue")
}

fn build_rich_cell_image_fixture_plural_richvalues() -> Vec<u8> {
    build_rich_cell_image_fixture_with_rich_value_part_base("richValues")
}

#[test]
fn resolves_image_target_for_rich_value_cell() -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_rich_cell_image_fixture();
    let doc = load_from_bytes(&bytes)?;
    let sheet_id = doc.workbook.sheets[0].id;

    assert_eq!(
        doc.image_target_for_cell(sheet_id, CellRef::from_a1("A1")?)?,
        Some("xl/media/image2.png".to_string())
    );
    assert_eq!(
        doc.image_target_for_cell(sheet_id, CellRef::from_a1("B1")?)?,
        Some("xl/media/image2.png".to_string())
    );
    assert_eq!(
        doc.image_target_for_cell(sheet_id, CellRef::from_a1("C1")?)?,
        None
    );

    Ok(())
}

#[test]
fn resolves_image_target_for_rich_value_cell_with_plural_richvalues_parts(
) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_rich_cell_image_fixture_plural_richvalues();
    let doc = load_from_bytes(&bytes)?;
    let sheet_id = doc.workbook.sheets[0].id;

    assert_eq!(
        doc.image_target_for_cell(sheet_id, CellRef::from_a1("A1")?)?,
        Some("xl/media/image2.png".to_string())
    );
    assert_eq!(
        doc.image_target_for_cell(sheet_id, CellRef::from_a1("B1")?)?,
        Some("xl/media/image2.png".to_string())
    );
    assert_eq!(
        doc.image_target_for_cell(sheet_id, CellRef::from_a1("C1")?)?,
        None
    );

    Ok(())
}
