use std::io::{Cursor, Write};

use formula_model::CellRef;
use formula_xlsx::XlsxPackage;
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
fn embedded_cell_images_supports_zero_based_vm_with_rd_rich_value_schema() {
    // This is a synthetic workbook that exercises the `rdrichvalue.xml` local-image schema while
    // keeping worksheet `c/@vm` values 0-based (vm="0" selects the first `<valueMetadata><bk>`).
    //
    // This specifically guards against a regression where `extract_embedded_cell_images` used a
    // 1-based lookup (`vm_to_rich_value.get(&vm)`) in its rdRichValue fast path, causing vm="0"
    // cells to be ignored.

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
      <c r="A1" vm="0"><v>0</v></c>
    </row>
    <row r="2">
      <c r="A2" vm="1"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

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

    let rd_rich_value_structure_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rdRichValueStructure xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <s t="_localImage">
    <k n="_rvRel:LocalImageIdentifier"/>
    <k n="CalcOrigin"/>
    <k n="Text"/>
  </s>
</rdRichValueStructure>"#;

    let rd_rich_value_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rdRichValue xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <rv s="0"><v>0</v><v>5</v><v>Alt1</v></rv>
  <rv s="0"><v>1</v><v>6</v><v>Alt2</v></rv>
</rdRichValue>"#;

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
        ("xl/richData/rdrichvalue.xml", rd_rich_value_xml),
        ("xl/richData/rdrichvaluestructure.xml", rd_rich_value_structure_xml),
        ("xl/richData/richValueRel.xml", rich_value_rel_xml),
        ("xl/richData/_rels/richValueRel.xml.rels", rich_value_rel_rels),
        ("xl/media/image1.png", b"img1"),
        ("xl/media/image2.png", b"img2"),
    ]);

    let pkg = XlsxPackage::from_bytes(&bytes).expect("read package");
    let images = pkg
        .extract_embedded_cell_images()
        .expect("extract embedded cell images");

    let a1 = (
        "xl/worksheets/sheet1.xml".to_string(),
        CellRef::from_a1("A1").unwrap(),
    );
    let a2 = (
        "xl/worksheets/sheet1.xml".to_string(),
        CellRef::from_a1("A2").unwrap(),
    );

    let e1 = images.get(&a1).expect("expected A1 image");
    assert_eq!(e1.image_part, "xl/media/image1.png");
    assert_eq!(e1.image_bytes, b"img1");
    assert_eq!(e1.calc_origin, 5);
    assert_eq!(e1.alt_text.as_deref(), Some("Alt1"));

    let e2 = images.get(&a2).expect("expected A2 image");
    assert_eq!(e2.image_part, "xl/media/image2.png");
    assert_eq!(e2.image_bytes, b"img2");
    assert_eq!(e2.calc_origin, 6);
    assert_eq!(e2.alt_text.as_deref(), Some("Alt2"));
}

