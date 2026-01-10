use std::fs;

use formula_model::CommentKind;
use formula_xlsx::comments::extract_comment_parts;
use formula_xlsx::XlsxPackage;

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
