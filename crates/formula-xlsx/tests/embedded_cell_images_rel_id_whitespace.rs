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
fn embedded_cell_images_trims_rel_id_attribute_value() {
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

    // Use vm="0" and omit metadata/richValue tables to exercise the fallback path that treats `vm`
    // as a direct index into `richValueRel.xml` relationship slots.
    let sheet_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" vm="0"/>
    </row>
  </sheetData>
</worksheet>"#;

    // Note the whitespace around the relationship id value.
    let rich_value_rel_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueRel xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rel r:id=" rId1 "/>
</richValueRel>"#;

    let rich_value_rel_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>"#;

    let bytes = build_package(&[
        ("xl/workbook.xml", workbook_xml),
        ("xl/_rels/workbook.xml.rels", workbook_rels),
        ("xl/worksheets/sheet1.xml", sheet_xml),
        ("xl/richData/richValueRel.xml", rich_value_rel_xml),
        ("xl/richData/_rels/richValueRel.xml.rels", rich_value_rel_rels),
        ("xl/media/image1.png", b"img1"),
    ]);

    let pkg = XlsxPackage::from_bytes(&bytes).expect("read package");
    let images = pkg
        .extract_embedded_cell_images()
        .expect("extract embedded cell images");

    let addr = (
        "xl/worksheets/sheet1.xml".to_string(),
        CellRef::from_a1("A1").unwrap(),
    );
    let img = images.get(&addr).expect("expected embedded cell image");

    assert_eq!(img.image_part, "xl/media/image1.png");
    assert_eq!(img.image_bytes, b"img1");
}

#[test]
fn embedded_cell_images_trims_relationship_id_attribute_value_in_rels() {
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

    // Use vm="0" and omit metadata/richValue tables to exercise the fallback path that treats `vm`
    // as a direct index into `richValueRel.xml` relationship slots.
    let sheet_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" vm="0"/>
    </row>
  </sheetData>
</worksheet>"#;

    let rich_value_rel_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueRel xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rel r:id="rId1"/>
</richValueRel>"#;

    // Note the whitespace around the relationship `Id` and the trailing whitespace in `Type`.
    let rich_value_rel_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id=" rId1 " Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image " Target="../media/image1.png"/>
</Relationships>"#;

    let bytes = build_package(&[
        ("xl/workbook.xml", workbook_xml),
        ("xl/_rels/workbook.xml.rels", workbook_rels),
        ("xl/worksheets/sheet1.xml", sheet_xml),
        ("xl/richData/richValueRel.xml", rich_value_rel_xml),
        ("xl/richData/_rels/richValueRel.xml.rels", rich_value_rel_rels),
        ("xl/media/image1.png", b"img1"),
    ]);

    let pkg = XlsxPackage::from_bytes(&bytes).expect("read package");
    let images = pkg
        .extract_embedded_cell_images()
        .expect("extract embedded cell images");

    let addr = (
        "xl/worksheets/sheet1.xml".to_string(),
        CellRef::from_a1("A1").unwrap(),
    );
    let img = images.get(&addr).expect("expected embedded cell image");

    assert_eq!(img.image_part, "xl/media/image1.png");
    assert_eq!(img.image_bytes, b"img1");
}

#[test]
fn embedded_cell_images_resolves_media_targets_relative_to_xl() {
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

    // Two cells with vm indices that will be interpreted directly as richValueRel slots.
    let sheet_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" vm="0"/>
    </row>
    <row r="2">
      <c r="A2" vm="1"/>
    </row>
  </sheetData>
</worksheet>"#;

    // Two richValueRel slots.
    let rich_value_rel_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueRel xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rel r:id="rId1"/>
  <rel r:id="rId2"/>
</richValueRel>"#;

    // Some producers emit targets relative to `xl/` (`media/...`) or prefix them with `xl/`
    // without a leading `/`. The embedded-cell-image extractor should tolerate both.
    let rich_value_rel_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="xl/media/image2.png"/>
</Relationships>"#;

    let bytes = build_package(&[
        ("xl/workbook.xml", workbook_xml),
        ("xl/_rels/workbook.xml.rels", workbook_rels),
        ("xl/worksheets/sheet1.xml", sheet_xml),
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

    assert_eq!(images.get(&a1).unwrap().image_part, "xl/media/image1.png");
    assert_eq!(images.get(&a1).unwrap().image_bytes, b"img1");

    assert_eq!(images.get(&a2).unwrap().image_part, "xl/media/image2.png");
    assert_eq!(images.get(&a2).unwrap().image_bytes, b"img2");
}

#[test]
fn embedded_cell_images_resolves_image_targets_with_query_fragments_and_backslashes() {
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

    // Two cells with vm indices that will be interpreted directly as richValueRel slots.
    let sheet_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" vm="0"/>
    </row>
    <row r="2">
      <c r="A2" vm="1"/>
    </row>
  </sheetData>
</worksheet>"#;

    // Two richValueRel slots.
    let rich_value_rel_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueRel xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rel r:id="rId1"/>
  <rel r:id="rId2"/>
</richValueRel>"#;

    // Some producers emit:
    // - URI fragments/query strings on in-package targets, and/or
    // - Windows-style path separators.
    //
    // The embedded-cell-image extractor should still resolve these to valid package part names.
    let rich_value_rel_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target=" media\image1.png?foo=bar#frag "/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="xl\media\image2.png?foo=bar#frag2"/>
</Relationships>"#;

    let bytes = build_package(&[
        ("xl/workbook.xml", workbook_xml),
        ("xl/_rels/workbook.xml.rels", workbook_rels),
        ("xl/worksheets/sheet1.xml", sheet_xml),
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

    assert_eq!(images.get(&a1).unwrap().image_part, "xl/media/image1.png");
    assert_eq!(images.get(&a1).unwrap().image_bytes, b"img1");

    assert_eq!(images.get(&a2).unwrap().image_part, "xl/media/image2.png");
    assert_eq!(images.get(&a2).unwrap().image_bytes, b"img2");
}
