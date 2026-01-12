use std::io::{Cursor, Write};

use formula_model::CellRef;
use formula_xlsx::XlsxPackage;

fn build_rich_image_xlsx_with_media_relative_target(include_rich_value_part: bool) -> Vec<u8> {
    let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let workbook_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let sheet1_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" vm="0"><v>ignored</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    // Simplest supported vm mapping: vm idx (0-based) -> rc/@v (index into `<rvb i="..."/>` list)
    // without any futureMetadata indirection.
    let metadata_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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

    // Rich value table with one record that points at relationship index 0.
    let rich_value_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValue xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <rv>
    <v t="rel">0</v>
  </rv>
</richValue>"#;

    let rich_value_rel_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueRel xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rel r:id="rId1"/>
</richValueRel>"#;

    // NOTE: Some producers emit `Target="media/image1.png"` (relative to `xl/`) rather than the
    // more common `Target="../media/image1.png"` (relative to `xl/richData/`).
    let rich_value_rel_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png#fragment"/>
</Relationships>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(sheet1_xml).unwrap();

    zip.start_file("xl/metadata.xml", options).unwrap();
    zip.write_all(metadata_xml).unwrap();

    if include_rich_value_part {
        zip.start_file("xl/richData/richValue.xml", options).unwrap();
        zip.write_all(rich_value_xml).unwrap();
    }

    zip.start_file("xl/richData/richValueRel.xml", options)
        .unwrap();
    zip.write_all(rich_value_rel_xml).unwrap();

    zip.start_file("xl/richData/_rels/richValueRel.xml.rels", options)
        .unwrap();
    zip.write_all(rich_value_rel_rels).unwrap();

    zip.start_file("xl/media/image1.png", options).unwrap();
    zip.write_all(b"fakepng").unwrap();

    zip.finish().unwrap().into_inner()
}

#[test]
fn rich_data_cell_images_resolves_media_relative_targets_with_rich_value_parts() {
    let bytes = build_rich_image_xlsx_with_media_relative_target(true);
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
    let images = pkg.extract_rich_cell_images_by_cell().expect("extract images");

    let key = ("Sheet1".to_string(), CellRef::from_a1("A1").unwrap());
    assert_eq!(
        images.get(&key).map(|v| v.as_slice()),
        Some(b"fakepng".as_slice())
    );
    assert_eq!(
        images.len(),
        1,
        "unexpected extra images extracted: keys={:?}",
        images.keys().collect::<Vec<_>>()
    );
}

#[test]
fn rich_data_cell_images_resolves_media_relative_targets_without_rich_value_parts() {
    let bytes = build_rich_image_xlsx_with_media_relative_target(false);
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
    let images = pkg.extract_rich_cell_images_by_cell().expect("extract images");

    let key = ("Sheet1".to_string(), CellRef::from_a1("A1").unwrap());
    assert_eq!(
        images.get(&key).map(|v| v.as_slice()),
        Some(b"fakepng".as_slice())
    );
    assert_eq!(
        images.len(),
        1,
        "unexpected extra images extracted: keys={:?}",
        images.keys().collect::<Vec<_>>()
    );
}

