use std::io::{Cursor, Write};

use formula_xlsx::{ExtractedRichValueImages, RichValueWarning, XlsxPackage};

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

#[test]
fn rich_value_indexing_prefers_explicit_ids_across_multiple_parts() {
    let rich_value1 = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValue xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rv>
    <a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" r:embed="rIdUnused"/>
  </rv>
</richValue>"#;

    let rich_value2 = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValue xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rv id="10">
    <a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" r:embed="rIdImg"/>
  </rv>
</richValue>"#;

    let rich_value2_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>"#;

    let metadata = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <rvb i="10"/>
</metadata>"#;

    let image_bytes = b"fake-png";

    let bytes = build_package(&[
        ("xl/metadata.xml", metadata),
        ("xl/richData/richValue1.xml", rich_value1),
        ("xl/richData/richValue2.xml", rich_value2),
        ("xl/richData/_rels/richValue2.xml.rels", rich_value2_rels),
        ("xl/media/image1.png", image_bytes),
    ]);

    let pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
    let ExtractedRichValueImages { images, warnings } =
        pkg.extract_rich_value_images().expect("extract");

    assert_eq!(
        images.get(&10).map(Vec::as_slice),
        Some(image_bytes.as_slice()),
        "expected rich value index 10 to resolve to image1.png"
    );
    assert!(
        warnings.is_empty(),
        "did not expect warnings, got: {warnings:?}"
    );
}

#[test]
fn rich_value_indexing_collision_is_deterministic_and_warns() {
    let rich_value1 = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValue xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rv id="0">
    <a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" r:embed="rIdImg1"/>
  </rv>
</richValue>"#;

    let rich_value2 = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValue xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rv id="0">
    <a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" r:embed="rIdImg2"/>
  </rv>
</richValue>"#;

    let rich_value1_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>"#;

    let rich_value2_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImg2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.png"/>
</Relationships>"#;

    let metadata = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <rvb i="0"/>
</metadata>"#;

    let image1_bytes = b"first-image";
    let image2_bytes = b"second-image";

    let bytes = build_package(&[
        ("xl/metadata.xml", metadata),
        ("xl/richData/richValue1.xml", rich_value1),
        ("xl/richData/_rels/richValue1.xml.rels", rich_value1_rels),
        ("xl/richData/richValue2.xml", rich_value2),
        ("xl/richData/_rels/richValue2.xml.rels", rich_value2_rels),
        ("xl/media/image1.png", image1_bytes),
        ("xl/media/image2.png", image2_bytes),
    ]);

    let pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");
    let ExtractedRichValueImages { images, warnings } =
        pkg.extract_rich_value_images().expect("extract");

    // Deterministic collision resolution: first wins (richValue1.xml comes before richValue2.xml).
    assert_eq!(images.get(&0).map(Vec::as_slice), Some(image1_bytes.as_slice()));

    assert!(
        warnings
            .iter()
            .any(|w| matches!(w, RichValueWarning::DuplicateIndex { index: 0, .. })),
        "expected a DuplicateIndex warning, got: {warnings:?}"
    );
}
