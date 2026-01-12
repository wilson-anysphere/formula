use std::io::Write as _;

use formula_model::CellRef;
use formula_xlsx::{extract_embedded_images, XlsxPackage};
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipWriter};

fn build_zip(parts: &[(&str, &[u8])]) -> Vec<u8> {
    let cursor = std::io::Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

    for (name, bytes) in parts {
        zip.start_file(*name, options).expect("start zip file");
        zip.write_all(bytes).expect("write zip bytes");
    }

    zip.finish().expect("finish zip").into_inner()
}

#[test]
fn extract_embedded_images_accepts_plural_richvalues_part_name() {
    // This is the legacy `extract_embedded_images` API. Ensure it also discovers plural
    // `xl/richData/richValues*.xml` parts.

    let worksheet_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="e" vm="1"><v>#VALUE!</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    // Direct metadata mapping (no futureMetadata indirection): vm=1 -> rich value index 1.
    let metadata_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <metadataTypes count="1">
    <metadataType name="XLRICHVALUE"/>
  </metadataTypes>
  <valueMetadata count="1">
    <bk><rc t="1" v="1"/></bk>
  </valueMetadata>
</metadata>"#;

    // Plural rich values part name. Rich value index 1 references relationship slot 0.
    let rich_values_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValue xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <rv><v t="string">not an image</v></rv>
  <rv><v t="rel">0</v></rv>
</richValue>"#;

    let rich_value_rel_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueRel xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rel r:id="rId1"/>
</richValueRel>"#;

    let rich_value_rel_rels_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    Target="../media/image1.png"/>
</Relationships>"#;

    let image_bytes = b"png-bytes";

    let bytes = build_zip(&[
        ("xl/worksheets/sheet1.xml", worksheet_xml),
        ("xl/metadata.xml", metadata_xml),
        ("xl/richData/richValues.xml", rich_values_xml),
        ("xl/richData/richValueRel.xml", rich_value_rel_xml),
        (
            "xl/richData/_rels/richValueRel.xml.rels",
            rich_value_rel_rels_xml,
        ),
        ("xl/media/image1.png", image_bytes),
    ]);

    let pkg = XlsxPackage::from_bytes(&bytes).expect("parse xlsx");
    let images = extract_embedded_images(&pkg).expect("extract embedded images");
    assert_eq!(images.len(), 1);

    let image = &images[0];
    assert_eq!(image.sheet_part, "xl/worksheets/sheet1.xml");
    assert_eq!(image.cell, CellRef::from_a1("A1").unwrap());
    assert_eq!(image.image_target, "xl/media/image1.png");
    assert_eq!(image.bytes, image_bytes);
    assert_eq!(image.alt_text, None);
}

