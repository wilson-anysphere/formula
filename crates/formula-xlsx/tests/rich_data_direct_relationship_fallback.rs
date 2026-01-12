use std::collections::HashMap;
use std::io::{Cursor, Write};

use formula_model::CellRef;
use formula_xlsx::rich_data::extract_rich_cell_images;
use formula_xlsx::XlsxPackage;

#[test]
fn rich_data_falls_back_to_rich_value_part_relationships_when_rich_value_rel_missing() {
    // This package intentionally omits `xl/richData/richValueRel.xml` entirely. The rich value
    // entry references `rId1` directly (`<v>rId1</v>`), and the relationship is stored on the
    // rich value part itself (`xl/richData/_rels/richValue.xml.rels`).
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
</Relationships>"#;

    let sheet1_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" vm="0"><v>ignored</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    // Minimal metadata mapping: vm 0 -> rich value index 0.
    let metadata_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <valueMetadata count="1">
    <bk>
      <rc t="0" v="0"/>
    </bk>
  </valueMetadata>
  <extLst>
    <ext uri="{D06F3F9D-0A6B-4D0A-80D3-712A9E1D37F4}">
      <xlrd:rvb i="0"/>
    </ext>
  </extLst>
</metadata>"#;

    let rich_value_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValue xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <rv>
    <v>rId1</v>
  </rv>
</richValue>"#;

    let rich_value_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options =
        zip::write::FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(sheet1_xml.as_bytes()).unwrap();

    zip.start_file("xl/metadata.xml", options).unwrap();
    zip.write_all(metadata_xml.as_bytes()).unwrap();

    zip.start_file("xl/richData/richValue.xml", options).unwrap();
    zip.write_all(rich_value_xml.as_bytes()).unwrap();

    zip.start_file("xl/richData/_rels/richValue.xml.rels", options)
        .unwrap();
    zip.write_all(rich_value_rels.as_bytes()).unwrap();

    zip.start_file("xl/media/image1.png", options).unwrap();
    zip.write_all(b"fakepng").unwrap();

    let bytes = zip.finish().unwrap().into_inner();

    let pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
    let images = extract_rich_cell_images(&pkg).expect("extract images");

    let mut expected: HashMap<(String, CellRef), Vec<u8>> = HashMap::new();
    expected.insert(
        ("Sheet1".to_string(), CellRef::from_a1("A1").unwrap()),
        b"fakepng".to_vec(),
    );

    assert_eq!(images, expected);
}
