use std::io::{Cursor, Write};

use formula_xlsx::rich_data::RichValueRels;
use formula_xlsx::XlsxPackage;
use zip::write::FileOptions;
use zip::ZipWriter;

fn build_package(entries: &[(&str, &[u8])]) -> XlsxPackage {
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    for (name, bytes) in entries {
        zip.start_file(*name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    let bytes = zip.finish().unwrap().into_inner();
    XlsxPackage::from_bytes(&bytes).expect("read test pkg")
}

#[test]
fn rich_value_rel_extracts_r_ids_and_resolves_targets() {
    let rich_value_rel_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueRel xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rel r:id="rId2"/>
  <rel r:id="rId5"/>
</richValueRel>"#;

    let rels_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    Target="../media/image2.png"/>
  <Relationship Id="rId5"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"
    Target="/xl/custom.xml"/>
</Relationships>"#;

    let pkg = build_package(&[
        ("xl/richData/richValueRel.xml", rich_value_rel_xml),
        ("xl/richData/_rels/richValueRel.xml.rels", rels_xml),
        ("xl/media/image2.png", b"png-bytes"),
        ("xl/custom.xml", br#"<a/>"#),
    ]);

    let rels = RichValueRels::from_package(&pkg)
        .expect("parse richValueRel.xml")
        .expect("richValueRel.xml present");

    assert_eq!(rels.r_ids, vec!["rId2".to_string(), "rId5".to_string()]);
    assert_eq!(
        rels.resolve_target(&pkg, 0).as_deref(),
        Some("xl/media/image2.png")
    );
    assert_eq!(
        rels.resolve_target(&pkg, 1).as_deref(),
        Some("xl/custom.xml")
    );
    assert_eq!(rels.resolve_target(&pkg, 2), None);
}

#[test]
fn rich_value_rel_tolerates_prefixes_and_wrappers() {
    // `bar:rel` has the correct local-name, but the XML uses arbitrary prefixes
    // and wrapper nodes. The parser should still capture the `r:id` values in
    // document order.
    let rich_value_rel_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<foo:root xmlns:foo="urn:foo"
  xmlns:bar="urn:bar"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <foo:wrapper>
    <bar:rel r:id="rId7" foo:unknown="1"/>
    <foo:more>
      <bar:rel r:id="rId8"/>
    </foo:more>
  </foo:wrapper>
</foo:root>"#;

    let rels_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rel:Relationships xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships">
  <rel:Relationship rel:Id="rId7"
    rel:Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    rel:Target="../media/image7.png"/>
  <rel:Relationship rel:Id="rId8"
    rel:Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    rel:Target="../media/image8.png"/>
</rel:Relationships>"#;

    let pkg = build_package(&[
        ("xl/richData/richValueRel.xml", rich_value_rel_xml),
        ("xl/richData/_rels/richValueRel.xml.rels", rels_xml),
        ("xl/media/image7.png", b"png-bytes"),
        ("xl/media/image8.png", b"png-bytes"),
    ]);

    let rels = RichValueRels::from_package(&pkg)
        .expect("parse richValueRel.xml")
        .expect("richValueRel.xml present");

    assert_eq!(rels.r_ids, vec!["rId7".to_string(), "rId8".to_string()]);
    assert_eq!(
        rels.resolve_target(&pkg, 0).as_deref(),
        Some("xl/media/image7.png")
    );
    assert_eq!(
        rels.resolve_target(&pkg, 1).as_deref(),
        Some("xl/media/image8.png")
    );
}
