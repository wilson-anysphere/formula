use std::io::{Cursor, Write};

use formula_model::CellRef;
use formula_xlsx::{extract_embedded_images, XlsxPackage};
use zip::write::FileOptions;
use zip::ZipWriter;

fn build_package(entries: &[(&str, &[u8])]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    for (name, bytes) in entries {
        zip.start_file(*name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    zip.finish().unwrap().into_inner()
}

#[test]
fn extract_embedded_images_supports_zero_based_vm_with_multiple_value_metadata_records() {
    // Like `embedded_cell_images_vm_zero_based_multi_record`, but exercises the legacy
    // `extract_embedded_images()` API.

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

    // Two cells with 0-based vm values.
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

    // metadata.xml maps valueMetadata bk[0] -> richValue index 0, and bk[1] -> richValue index 1.
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
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <values>
    <rv><v kind="rel">0</v></rv>
    <rv><v kind="rel">1</v></rv>
  </values>
</rvData>"#;

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

    let bytes = build_package(&[
        ("xl/workbook.xml", workbook_xml),
        ("xl/_rels/workbook.xml.rels", workbook_rels),
        ("xl/worksheets/sheet1.xml", sheet1_xml),
        ("xl/metadata.xml", metadata_xml),
        ("xl/richData/richValue.xml", rich_value_xml),
        ("xl/richData/richValueRel.xml", rich_value_rel_xml),
        ("xl/richData/_rels/richValueRel.xml.rels", rich_value_rel_rels),
        ("xl/media/image1.png", b"img1"),
        ("xl/media/image2.png", b"img2"),
    ]);

    let pkg = XlsxPackage::from_bytes(&bytes).expect("read package");
    let images = extract_embedded_images(&pkg).expect("extract embedded images");
    assert_eq!(images.len(), 2);

    let mut a1 = None;
    let mut a2 = None;
    for entry in &images {
        match entry.cell {
            c if c == CellRef::from_a1("A1").unwrap() => a1 = Some(entry),
            c if c == CellRef::from_a1("A2").unwrap() => a2 = Some(entry),
            _ => {}
        }
    }

    let a1 = a1.expect("expected Sheet1!A1 image");
    let a2 = a2.expect("expected Sheet1!A2 image");

    assert_eq!(a1.image_target, "xl/media/image1.png");
    assert_eq!(a1.bytes, b"img1");
    assert_eq!(a2.image_target, "xl/media/image2.png");
    assert_eq!(a2.bytes, b"img2");
}

#[test]
fn extract_embedded_images_supports_zero_based_vm_even_when_no_cell_uses_vm_zero() {
    // Some producers appear to use 0-based worksheet `c/@vm` values but never reference the first
    // `<valueMetadata>` record (`vm="0"`). In that case we can't infer the base from the presence
    // of `vm="0"` alone and must choose the offset based on which one yields valid image
    // resolutions.

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

    // A single cell with a 0-based vm value, but we never reference vm=0.
    let sheet1_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" vm="1"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    // metadata.xml provides two valueMetadata records; the second maps to rich value index 1.
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

    // Two rich values: the first does not reference a relationship; the second references rel idx 0.
    let rich_value_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <values>
    <rv><v kind="string">No image</v></rv>
    <rv><v kind="rel">0</v></rv>
  </values>
</rvData>"#;

    let rich_value_rel_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueRel xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rel r:id="rId1"/>
</richValueRel>"#;

    let rich_value_rel_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>"#;

    let bytes = build_package(&[
        ("xl/workbook.xml", workbook_xml),
        ("xl/_rels/workbook.xml.rels", workbook_rels),
        ("xl/worksheets/sheet1.xml", sheet1_xml),
        ("xl/metadata.xml", metadata_xml),
        ("xl/richData/richValue.xml", rich_value_xml),
        ("xl/richData/richValueRel.xml", rich_value_rel_xml),
        ("xl/richData/_rels/richValueRel.xml.rels", rich_value_rel_rels),
        ("xl/media/image1.png", b"img1"),
    ]);

    let pkg = XlsxPackage::from_bytes(&bytes).expect("read package");
    let images = extract_embedded_images(&pkg).expect("extract embedded images");
    assert_eq!(images.len(), 1);
    assert_eq!(images[0].cell, CellRef::from_a1("A1").unwrap());
    assert_eq!(images[0].image_target, "xl/media/image1.png");
    assert_eq!(images[0].bytes, b"img1");
}
