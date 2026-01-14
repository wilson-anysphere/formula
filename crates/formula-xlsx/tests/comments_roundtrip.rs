use std::fs;
use std::io::{Cursor, Write};

use formula_model::CommentKind;
use formula_xlsx::comments::extract_comment_parts;
use formula_xlsx::XlsxPackage;
use zip::write::FileOptions;
use zip::CompressionMethod;

#[test]
fn preserves_comment_related_parts_on_round_trip() {
    let fixture_path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/comments.xlsx");
    let bytes = fs::read(fixture_path).expect("fixture workbook should be readable");
    let pkg = XlsxPackage::from_bytes(&bytes).expect("fixture should parse as xlsx package");

    let parts = extract_comment_parts(&pkg);
    assert!(parts
        .comments
        .iter()
        .any(|comment| comment.kind == CommentKind::Note));
    assert!(parts
        .comments
        .iter()
        .any(|comment| comment.kind == CommentKind::Threaded));

    assert!(parts.preserved.contains_key("xl/comments1.xml"));
    assert!(parts.preserved.contains_key("xl/drawings/vmlDrawing1.vml"));
    assert!(parts.preserved.contains_key("xl/threadedComments/threadedComments1.xml"));
    assert!(parts.preserved.contains_key("xl/commentsExt1.xml"));
    assert!(parts.preserved.contains_key("xl/persons/persons1.xml"));

    let threaded = parts
        .comments
        .iter()
        .find(|comment| comment.kind == CommentKind::Threaded)
        .expect("fixture should contain threaded comment");
    assert_eq!(threaded.author.name, "Alex");
    assert_eq!(
        threaded
            .replies
            .first()
            .map(|reply| reply.author.name.as_str()),
        Some("Sam")
    );

    let written = pkg.write_to_bytes().expect("write package");
    let pkg2 = XlsxPackage::from_bytes(&written).expect("read package");
    for (path, original_bytes) in parts.preserved.iter() {
        let roundtrip = pkg2
            .part(path)
            .unwrap_or_else(|| panic!("missing roundtrip part {path}"));
        assert_eq!(
            roundtrip,
            original_bytes.as_slice(),
            "part should be preserved byte-for-byte: {path}"
        );
    }
}

fn make_package_with_non_canonical_comment_part_names() -> Vec<u8> {
    // This is intentionally *not* a full XLSX package; `XlsxPackage` only requires a ZIP with the
    // relevant parts present.
    const LEGACY_COMMENTS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors>
    <author>Alex</author>
  </authors>
  <commentList>
    <comment ref="A1" authorId="0">
      <text><r><t xml:space="preserve">Legacy note</t></r></text>
    </comment>
  </commentList>
</comments>
"#;

    const THREADED_COMMENTS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<threadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <threadedComment id="t1" ref="B2" personId="p1" done="1">
    <text><r><t xml:space="preserve">Thread root</t></r></text>
  </threadedComment>
  <threadedComment id="t2" parentId="t1" ref="B2" personId="p2">
    <text><r><t xml:space="preserve">First reply</t></r></text>
  </threadedComment>
</threadedComments>
"#;

    const PERSONS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<persons xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <person id="p1" displayName="Alex"/>
  <person id="p2" displayName="Sam"/>
</persons>
"#;

    const COMMENTS_EXT_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<commentsExt xmlns="http://schemas.microsoft.com/office/spreadsheetml/2014/main">
  <extLst/>
</commentsExt>
"#;

    let cursor = Cursor::new(Vec::<u8>::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options: FileOptions<'_, ()> =
        FileOptions::default().compression_method(CompressionMethod::Stored);

    // Non-canonical casing on part names.
    zip.start_file("XL/COMMENTS1.XML", options)
        .expect("zip entry creation should succeed");
    zip.write_all(LEGACY_COMMENTS_XML.as_bytes())
        .expect("legacy comments xml write should succeed");

    zip.start_file("XL/THREADEDCOMMENTS/THREADEDCOMMENTS1.XML", options)
        .expect("zip entry creation should succeed");
    zip.write_all(THREADED_COMMENTS_XML.as_bytes())
        .expect("threaded comments xml write should succeed");

    zip.start_file("XL/PERSONS/PERSONS1.XML", options)
        .expect("zip entry creation should succeed");
    zip.write_all(PERSONS_XML.as_bytes())
        .expect("persons xml write should succeed");

    zip.start_file("XL/COMMENTSEXT1.XML", options)
        .expect("zip entry creation should succeed");
    zip.write_all(COMMENTS_EXT_XML.as_bytes())
        .expect("commentsExt xml write should succeed");

    zip.finish()
        .expect("zip finalization should succeed")
        .into_inner()
}

#[test]
fn comment_parts_tolerate_non_canonical_part_name_casing() {
    let bytes = make_package_with_non_canonical_comment_part_names();
    let pkg = XlsxPackage::from_bytes(&bytes).expect("should parse as xlsx package");

    let parts = extract_comment_parts(&pkg);
    assert!(
        parts
            .comments
            .iter()
            .any(|comment| comment.kind == CommentKind::Note),
        "expected legacy note comment to be parsed"
    );
    assert!(
        parts
            .comments
            .iter()
            .any(|comment| comment.kind == CommentKind::Threaded),
        "expected threaded comment to be parsed"
    );

    let threaded = parts
        .comments
        .iter()
        .find(|comment| comment.kind == CommentKind::Threaded)
        .expect("threaded comment should exist");
    assert_eq!(threaded.author.name, "Alex");
    assert_eq!(
        threaded
            .replies
            .first()
            .map(|reply| reply.author.name.as_str()),
        Some("Sam")
    );

    // All comment-related parts should be preserved even with non-canonical casing.
    assert!(parts.preserved.contains_key("XL/COMMENTS1.XML"));
    assert!(parts
        .preserved
        .contains_key("XL/THREADEDCOMMENTS/THREADEDCOMMENTS1.XML"));
    assert!(parts.preserved.contains_key("XL/PERSONS/PERSONS1.XML"));
    assert!(parts.preserved.contains_key("XL/COMMENTSEXT1.XML"));
}
