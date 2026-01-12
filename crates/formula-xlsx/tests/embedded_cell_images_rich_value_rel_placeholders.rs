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
fn embedded_cell_images_preserves_rich_value_rel_placeholders() {
    // `richValueRel.xml` is a dense table (relationship-index -> rId). Some workbooks may include
    // placeholder `<rel/>` entries; those must be preserved during parsing to avoid shifting
    // indices.
    //
    // This exercises the fallback mode of `extract_embedded_cell_images` (no metadata/richValue
    // parts): it treats worksheet `vm` as a 1-based index into the richValueRel `<rel>` list.

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

    // vm="2" -> local_image_identifier = 1 (vm-1), which should resolve to the second `<rel>`
    // entry below.
    let sheet1_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" vm="2"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    let rich_value_rel_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueRel xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rel/>
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
        ("xl/richData/richValueRel.xml", rich_value_rel_xml),
        ("xl/richData/_rels/richValueRel.xml.rels", rich_value_rel_rels),
        ("xl/media/image1.png", b"img1"),
    ]);

    let pkg = XlsxPackage::from_bytes(&bytes).expect("read package");
    let images = pkg
        .extract_embedded_cell_images()
        .expect("extract embedded cell images");

    let key = (
        "xl/worksheets/sheet1.xml".to_string(),
        CellRef::from_a1("A1").unwrap(),
    );
    let img = images.get(&key).expect("expected A1 image");
    assert_eq!(img.image_part, "xl/media/image1.png");
    assert_eq!(img.image_bytes, b"img1");
}

