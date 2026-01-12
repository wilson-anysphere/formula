use std::io::{Cursor, Write};

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
fn discovers_cell_images_part_and_resolves_image_targets() {
    let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"></workbook>"#;

    // Intentionally use an unknown relationship type. Discovery must rely on the Target heuristic.
    let workbook_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdCellImages" Type="http://example.com/relationships/unknown" Target="cellimages.xml"/>
</Relationships>"#;

    let cellimages_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cellImages xmlns="http://schemas.microsoft.com/office/spreadsheetml/2019/cellimages"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <cellImage>
    <a:blip r:embed="rId1"/>
  </cellImage>
</cellImages>"#;

    let cellimages_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>"#;

    let png_bytes: &[u8] = b"not-a-real-png";

    let bytes = build_package(&[
        ("xl/workbook.xml", workbook_xml),
        ("xl/_rels/workbook.xml.rels", workbook_rels),
        ("xl/cellimages.xml", cellimages_xml),
        ("xl/_rels/cellimages.xml.rels", cellimages_rels),
        ("xl/media/image1.png", png_bytes),
    ]);

    let pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
    let info = pkg
        .cell_images_part_info()
        .expect("parse cell images")
        .expect("expected cellimages part");

    assert_eq!(info.part_path, "xl/cellimages.xml");
    assert_eq!(info.rels_path, "xl/_rels/cellimages.xml.rels");
    assert_eq!(info.embeds.len(), 1);
    assert_eq!(info.embeds[0].embed_rid, "rId1");
    assert_eq!(info.embeds[0].target_part, "xl/media/image1.png");
    assert_eq!(info.embeds[0].target_bytes.as_slice(), png_bytes);
}
