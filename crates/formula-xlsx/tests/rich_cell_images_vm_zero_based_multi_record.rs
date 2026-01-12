use std::collections::HashMap;
use std::io::{Cursor, Write};

use formula_model::CellRef;
use formula_xlsx::XlsxPackage;

fn build_package(entries: &[(&str, &[u8])]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    for (name, bytes) in entries {
        zip.start_file(*name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    zip.finish().unwrap().into_inner()
}

fn build_rich_cell_image_xlsx(include_rich_value_part: bool) -> Vec<u8> {
    let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let workbook_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata"
    Target="metadata.xml"/>
</Relationships>"#;

    // Two cells with 0-based `vm` values.
    let sheet1_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" vm="0"><v>0</v></c>
    </row>
    <row r="2">
      <c r="A2" vm="1"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    // metadata.xml maps valueMetadata bk[0] -> richValue index 0 and bk[1] -> richValue index 1.
    // Note that `vm` is *canonically* 1-based here, while the worksheet uses 0-based indices.
    let metadata_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <metadataTypes count="1">
    <metadataType name="XLRICHVALUE"/>
  </metadataTypes>
  <futureMetadata name="XLRICHVALUE" count="2">
    <bk><extLst><ext uri="{00000000-0000-0000-0000-000000000000}"><xlrd:rvb i="0"/></ext></extLst></bk>
    <bk><extLst><ext uri="{00000000-0000-0000-0000-000000000000}"><xlrd:rvb i="1"/></ext></extLst></bk>
  </futureMetadata>
  <valueMetadata count="2">
    <bk><rc t="1" v="0"/></bk>
    <bk><rc t="1" v="1"/></bk>
  </valueMetadata>
</metadata>"#;

    // Two rich values, each mapping to relationship slots 0 and 1.
    let rich_value_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValue xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <rv><v t="rel">0</v></rv>
  <rv><v t="rel">1</v></rv>
</richValue>"#;

    let rich_value_rel_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueRel xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rel r:id="rId1"/>
  <rel r:id="rId2"/>
</richValueRel>"#;

    let rich_value_rel_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.png"/>
</Relationships>"#;

    let mut entries: Vec<(&str, &[u8])> = vec![
        ("xl/workbook.xml", workbook_xml),
        ("xl/_rels/workbook.xml.rels", workbook_rels),
        ("xl/worksheets/sheet1.xml", sheet1_xml),
        ("xl/metadata.xml", metadata_xml),
        ("xl/richData/richValueRel.xml", rich_value_rel_xml),
        (
            "xl/richData/_rels/richValueRel.xml.rels",
            rich_value_rel_rels,
        ),
        ("xl/media/image1.png", b"img1"),
        ("xl/media/image2.png", b"img2"),
    ];
    if include_rich_value_part {
        entries.push(("xl/richData/richValue.xml", rich_value_xml));
    }

    build_package(&entries)
}

#[test]
fn rich_cell_images_supports_zero_based_vm_with_multiple_value_metadata_records() {
    // Regression test: storing both 0-based and 1-based `vm` keys in a single map causes
    // collisions for multi-record workbooks. (0-based vm=1 is ambiguous between metadata bk[0]
    // (1-based vm=1) and bk[1] (0-based vm=1)).
    let bytes = build_rich_cell_image_xlsx(true);
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read package");

    let images = pkg
        .extract_rich_cell_images_by_cell()
        .expect("extract rich cell images");

    let mut expected: HashMap<(String, CellRef), Vec<u8>> = HashMap::new();
    expected.insert(
        ("Sheet1".to_string(), CellRef::from_a1("A1").unwrap()),
        b"img1".to_vec(),
    );
    expected.insert(
        ("Sheet1".to_string(), CellRef::from_a1("A2").unwrap()),
        b"img2".to_vec(),
    );

    assert_eq!(images, expected);
}

#[test]
fn rich_cell_images_supports_zero_based_vm_multi_record_without_rich_value_parts() {
    // Same scenario as above, but omitting `xl/richData/richValue*.xml` so
    // `extract_rich_cell_images_by_cell` falls back to indexing directly into
    // `xl/richData/richValueRel.xml`.
    let bytes = build_rich_cell_image_xlsx(false);
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read package");

    let images = pkg
        .extract_rich_cell_images_by_cell()
        .expect("extract rich cell images");

    let mut expected: HashMap<(String, CellRef), Vec<u8>> = HashMap::new();
    expected.insert(
        ("Sheet1".to_string(), CellRef::from_a1("A1").unwrap()),
        b"img1".to_vec(),
    );
    expected.insert(
        ("Sheet1".to_string(), CellRef::from_a1("A2").unwrap()),
        b"img2".to_vec(),
    );

    assert_eq!(images, expected);
}

