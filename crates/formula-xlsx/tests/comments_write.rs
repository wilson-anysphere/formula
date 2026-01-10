use std::fs;

use formula_model::CommentKind;
use formula_xlsx::comments::{extract_comment_parts, render_comment_parts};
use formula_xlsx::XlsxPackage;

#[test]
fn writes_updated_comment_xml_and_preserves_unknown_parts() {
    let fixture_path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/comments.xlsx");
    let bytes = fs::read(fixture_path).expect("fixture workbook should be readable");
    let pkg = XlsxPackage::from_bytes(&bytes).expect("fixture should parse as xlsx package");
    let mut parts = extract_comment_parts(&pkg);

    let note = parts
        .comments
        .iter_mut()
        .find(|comment| comment.kind == CommentKind::Note)
        .expect("fixture should contain legacy note");
    note.content = "Updated note".to_string();

    let rendered = render_comment_parts(&parts);
    let updated_xml = rendered
        .get("xl/comments1.xml")
        .expect("legacy comments part should exist");
    let updated_xml = std::str::from_utf8(updated_xml).expect("comments xml should be utf-8");
    assert!(updated_xml.contains("Updated note"));

    assert_eq!(
        rendered.get("xl/commentsExt1.xml"),
        parts.preserved.get("xl/commentsExt1.xml"),
        "unknown comment-related parts should remain byte-for-byte"
    );
}
