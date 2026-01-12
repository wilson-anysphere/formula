use std::collections::HashMap;
use std::io::{Cursor, Write};

use formula_model::CellRef;
use formula_xlsx::extract_rich_cell_images;
use formula_xlsx::XlsxPackage;

fn build_rich_image_xlsx(
    include_metadata: bool,
    include_rich_value_rels: bool,
    fragmented_target: bool,
) -> Vec<u8> {
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

    let rich_value_types_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueTypes xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <richValueType id="0">
    <property name="dummy" t="i"/>
    <property name="image" t="rel"/>
  </richValueType>
</richValueTypes>"#;

    let rich_value_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValue xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <rv s="0">
    <v>123</v>
    <v>0</v>
  </rv>
</richValue>"#;

    let rich_value_rel_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueRel xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rel r:id="rId1"/>
</richValueRel>"#;

    let fragment = if fragmented_target { "#fragment" } else { "" };
    let rich_value_rel_rels = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png{}"/>
</Relationships>"#,
        fragment
    );

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(sheet1_xml.as_bytes()).unwrap();

    if include_metadata {
        zip.start_file("xl/metadata.xml", options).unwrap();
        zip.write_all(metadata_xml.as_bytes()).unwrap();
    }

    zip.start_file("xl/richData/richValueTypes.xml", options)
        .unwrap();
    zip.write_all(rich_value_types_xml.as_bytes()).unwrap();

    zip.start_file("xl/richData/richValue.xml", options).unwrap();
    zip.write_all(rich_value_xml.as_bytes()).unwrap();

    zip.start_file("xl/richData/richValueRel.xml", options)
        .unwrap();
    zip.write_all(rich_value_rel_xml.as_bytes()).unwrap();

    if include_rich_value_rels {
        zip.start_file("xl/richData/_rels/richValueRel.xml.rels", options)
            .unwrap();
        zip.write_all(rich_value_rel_rels.as_bytes()).unwrap();
    }

    zip.start_file("xl/media/image1.png", options).unwrap();
    zip.write_all(b"fakepng").unwrap();

    zip.finish().unwrap().into_inner()
}

fn build_rich_image_xlsx_with_structure_schema() -> Vec<u8> {
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

    // Typical Excel shape: richValueTypes maps type id -> structure id.
    let rich_value_types_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvTypes xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <types>
    <type id="0" structure="s_image"/>
  </types>
</rvTypes>"#;

    let rich_value_structure_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvStruct xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <structures>
    <structure id="s_image">
      <member name="dummy" kind="i"/>
      <member name="imageRel" kind="rel"/>
    </structure>
  </structures>
</rvStruct>"#;

    // Rich value record has two numeric `<v>` fields; only the second is the relationship index.
    let rich_value_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValue xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <rv s="0">
    <v>123</v>
    <v>0</v>
  </rv>
</richValue>"#;

    let rich_value_rel_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueRel xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rel r:id="rId1"/>
</richValueRel>"#;

    let rich_value_rel_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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

    zip.start_file("xl/richData/richValueTypes.xml", options)
        .unwrap();
    zip.write_all(rich_value_types_xml.as_bytes()).unwrap();

    zip.start_file("xl/richData/richValueStructure.xml", options)
        .unwrap();
    zip.write_all(rich_value_structure_xml.as_bytes()).unwrap();

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

fn build_rich_image_xlsx_with_custom_metadata_part() -> Vec<u8> {
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
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="custom-metadata.xml"/>
</Relationships>"#;

    let sheet1_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" vm="0"><v>ignored</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    // Simplest supported vm mapping: vm idx (0-based) -> rc/@v (index into `<rvb i="..."/>` list)
    // without any futureMetadata indirection.
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
    <v t="rel">0</v>
  </rv>
</richValue>"#;

    let rich_value_rel_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueRel xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rel r:id="rId1"/>
</richValueRel>"#;

    let rich_value_rel_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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

    zip.start_file("xl/custom-metadata.xml", options).unwrap();
    zip.write_all(metadata_xml.as_bytes()).unwrap();

    zip.start_file("xl/richData/richValue.xml", options).unwrap();
    zip.write_all(rich_value_xml.as_bytes()).unwrap();

    zip.start_file("xl/richData/richValueRel.xml", options).unwrap();
    zip.write_all(rich_value_rel_xml.as_bytes()).unwrap();

    zip.start_file("xl/richData/_rels/richValueRel.xml.rels", options)
        .unwrap();
    zip.write_all(rich_value_rel_rels.as_bytes()).unwrap();

    zip.start_file("xl/media/image1.png", options).unwrap();
    zip.write_all(b"custompng").unwrap();

    zip.finish().unwrap().into_inner()
}

#[test]
fn extracts_rich_cell_image_bytes_from_vm_chain() {
    let bytes = build_rich_image_xlsx(true, true, false);
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
    let images = pkg.extract_rich_cell_images_by_cell().expect("extract images");

    let mut expected: HashMap<(String, CellRef), Vec<u8>> = HashMap::new();
    expected.insert(
        ("Sheet1".to_string(), CellRef::from_a1("A1").unwrap()),
        b"fakepng".to_vec(),
    );

    assert_eq!(images, expected);
}

#[test]
fn extracts_rich_cell_image_bytes_from_structure_schema() {
    let bytes = build_rich_image_xlsx_with_structure_schema();
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
    let images = extract_rich_cell_images(&pkg).expect("extract images");

    let mut expected: HashMap<(String, CellRef), Vec<u8>> = HashMap::new();
    expected.insert(
        ("Sheet1".to_string(), CellRef::from_a1("A1").unwrap()),
        b"fakepng".to_vec(),
    );

    assert_eq!(images, expected);
}

#[test]
fn missing_metadata_xml_returns_empty_map() {
    let bytes = build_rich_image_xlsx(false, true, false);
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
    let images = pkg.extract_rich_cell_images_by_cell().expect("extract images");
    assert!(images.is_empty());
}

#[test]
fn missing_rich_value_rels_returns_empty_map() {
    let bytes = build_rich_image_xlsx(true, false, false);
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
    let images = pkg.extract_rich_cell_images_by_cell().expect("extract images");
    assert!(images.is_empty());
}

#[test]
fn extracts_rich_cell_image_bytes_when_relationship_target_has_fragment() {
    let bytes = build_rich_image_xlsx(true, true, true);
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
    let images = pkg.extract_rich_cell_images_by_cell().expect("extract images");

    let mut expected: HashMap<(String, CellRef), Vec<u8>> = HashMap::new();
    expected.insert(
        ("Sheet1".to_string(), CellRef::from_a1("A1").unwrap()),
        b"fakepng".to_vec(),
    );

    assert_eq!(images, expected);
}

#[test]
fn extracts_rich_cell_image_bytes_from_workbook_sheet_metadata_relationship() {
    let bytes = build_rich_image_xlsx_with_custom_metadata_part();
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
    let images = pkg.extract_rich_cell_images_by_cell().expect("extract images");

    let mut expected: HashMap<(String, CellRef), Vec<u8>> = HashMap::new();
    expected.insert(
        ("Sheet1".to_string(), CellRef::from_a1("A1").unwrap()),
        b"custompng".to_vec(),
    );

    assert_eq!(images, expected);
}

#[test]
fn rich_data_cell_images_extracts_image_in_cell_richdata_fixture() {
    let fixture_path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/xlsx/basic/image-in-cell-richdata.xlsx");
    let bytes = std::fs::read(&fixture_path)
        .unwrap_or_else(|e| panic!("read fixture {}: {e}", fixture_path.display()));
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");

    let images = pkg.extract_rich_cell_images_by_cell().expect("extract images");

    let key = ("Sheet1".to_string(), CellRef::from_a1("A1").unwrap());
    let expected = pkg
        .part("xl/media/image1.png")
        .expect("xl/media/image1.png present")
        .to_vec();

    assert_eq!(images.get(&key), Some(&expected));
    assert_eq!(
        images.len(),
        1,
        "unexpected extra images extracted: keys={:?}",
        images.keys().collect::<Vec<_>>()
    );
}

#[test]
fn rich_data_cell_images_extracts_image_in_cell_fixture() {
    let fixture_path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/xlsx/basic/image-in-cell.xlsx");
    let bytes = std::fs::read(&fixture_path)
        .unwrap_or_else(|e| panic!("read fixture {}: {e}", fixture_path.display()));
    let pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");

    let images = pkg.extract_rich_cell_images_by_cell().expect("extract images");

    let img1 = pkg
        .part("xl/media/image1.png")
        .expect("xl/media/image1.png present")
        .to_vec();
    let img2 = pkg
        .part("xl/media/image2.png")
        .expect("xl/media/image2.png present")
        .to_vec();

    let b2 = ("Sheet1".to_string(), CellRef::from_a1("B2").unwrap());
    let b3 = ("Sheet1".to_string(), CellRef::from_a1("B3").unwrap());
    let b4 = ("Sheet1".to_string(), CellRef::from_a1("B4").unwrap());

    assert_eq!(images.get(&b2), Some(&img1));
    assert_eq!(images.get(&b3), Some(&img1));
    assert_eq!(images.get(&b4), Some(&img2));

    assert_eq!(
        images.len(),
        3,
        "unexpected extra images extracted: keys={:?}",
        images.keys().collect::<Vec<_>>()
    );
}
