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
fn rich_data_relationship_targets_strip_uri_fragments() {
    let rich_value_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png#fragment"/>
</Relationships>"#;

    let bytes = build_package(&[
        ("xl/richData/richValueRel1.xml", br#"<richValueRel/>"#),
        (
            "xl/richData/_rels/richValueRel1.xml.rels",
            rich_value_rels,
        ),
        ("xl/media/image1.png", b"png-bytes"),
    ]);

    let pkg = XlsxPackage::from_bytes(&bytes).expect("read test package");
    let media = pkg.rich_data_media_parts();
    let media_via_legacy = pkg
        .extract_rich_data_images()
        .expect("extract richData images via legacy helper");

    assert_eq!(
        media.get("xl/media/image1.png").map(|v| v.as_slice()),
        Some(b"png-bytes".as_slice())
    );

    assert_eq!(media_via_legacy, media);
}

#[test]
fn rich_data_relationship_targets_accept_media_targets_relative_to_xl() {
    // Some producers emit `Target="media/image1.png"` (relative to `xl/`) from a richData part's
    // `.rels` (which lives under `xl/richData/_rels/`). This is technically an invalid relative
    // target (it should usually be `../media/image1.png`), but we should still discover the
    // referenced `xl/media/*` part.
    let rich_value_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>"#;

    let bytes = build_package(&[
        ("xl/richData/richValueRel1.xml", br#"<richValueRel/>"#),
        (
            "xl/richData/_rels/richValueRel1.xml.rels",
            rich_value_rels,
        ),
        ("xl/media/image1.png", b"png-bytes"),
    ]);

    let pkg = XlsxPackage::from_bytes(&bytes).expect("read test package");
    let media = pkg.rich_data_media_parts();
    let media_via_legacy = pkg
        .extract_rich_data_images()
        .expect("extract richData images via legacy helper");

    assert_eq!(
        media.get("xl/media/image1.png").map(|v| v.as_slice()),
        Some(b"png-bytes".as_slice())
    );

    assert_eq!(media_via_legacy, media);
}
