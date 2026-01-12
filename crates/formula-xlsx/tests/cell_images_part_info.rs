use std::io::{Cursor, Write};

use formula_model::Workbook;
use formula_xlsx::cell_images::CellImages;
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
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <cellImage r:embed="rId1"/>
</cellImages>"#;

    let cellimages_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png#frag"/>
  <Relationship Id="rIdIgnore" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
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

    let part_info = pkg
        .cell_images_part_info()
        .expect("cell_images_part_info")
        .expect("expected cell images part info");
    assert_eq!(part_info.part_path, "xl/cellimages.xml");
    assert_eq!(part_info.rels_path, "xl/_rels/cellimages.xml.rels");
    assert_eq!(part_info.embeds.len(), 1);
    assert_eq!(part_info.embeds[0].embed_rid, "rId1");
    assert_eq!(part_info.embeds[0].target_part, "xl/media/image1.png");
    assert_eq!(part_info.embeds[0].target_bytes.as_slice(), png_bytes);

    let mut workbook = Workbook::default();
    let images = CellImages::parse_from_parts(pkg.parts_map(), &mut workbook)
        .expect("parse cell images parts");

    assert_eq!(images.parts.len(), 1, "expected cellimages part");
    let part = &images.parts[0];

    assert_eq!(part.path, "xl/cellimages.xml");
    assert_eq!(part.rels_path, "xl/_rels/cellimages.xml.rels");
    assert_eq!(part.images.len(), 1);
    assert_eq!(part.images[0].embed_rel_id, "rId1");
    assert_eq!(
        part.images[0].target.as_deref(),
        Some("xl/media/image1.png")
    );

    let image_data = workbook
        .images
        .get(&formula_model::drawings::ImageId::new("image1.png"))
        .expect("expected image bytes to be loaded into workbook image store");
    assert_eq!(image_data.bytes.as_slice(), png_bytes);
}
