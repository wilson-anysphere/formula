use std::fs;

use formula_model::{CommentKind, CommentPatch};
use formula_xlsx::{load_from_bytes, XlsxPackage};

#[test]
fn comments_writeback_noop_preserves_comment_parts() {
    let fixture_path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/comments.xlsx");
    let bytes = fs::read(fixture_path).expect("fixture workbook should be readable");

    let doc = load_from_bytes(&bytes).expect("load_from_bytes");
    let saved = doc.save_to_vec().expect("save_to_vec");

    let orig_pkg = XlsxPackage::from_bytes(&bytes).expect("fixture should parse as xlsx package");
    let pkg = XlsxPackage::from_bytes(&saved).expect("roundtrip should parse as xlsx package");

    for path in [
        "xl/comments1.xml",
        "xl/threadedComments/threadedComments1.xml",
        "xl/commentsExt1.xml",
        "xl/drawings/vmlDrawing1.vml",
        "xl/persons/persons1.xml",
        "xl/worksheets/_rels/sheet1.xml.rels",
    ] {
        let original = orig_pkg.part(path).unwrap_or_else(|| panic!("missing fixture part {path}"));
        let roundtrip = pkg.part(path).unwrap_or_else(|| panic!("missing roundtrip part {path}"));
        assert_eq!(
            roundtrip, original,
            "expected no-op roundtrip to preserve part byte-for-byte: {path}"
        );
    }
}

#[test]
fn comments_writeback_updates_comment_xml_parts_and_preserves_unknown_parts() {
    let fixture_path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/comments.xlsx");
    let bytes = fs::read(fixture_path).expect("fixture workbook should be readable");

    let mut doc = load_from_bytes(&bytes).expect("load_from_bytes");
    let sheet_id = doc.workbook.sheets[0].id;
    let sheet = doc
        .workbook
        .sheet_mut(sheet_id)
        .expect("fixture sheet should exist");

    let note_id = sheet
        .iter_comments()
        .find(|(_, c)| c.kind == CommentKind::Note)
        .map(|(_, c)| c.id.clone())
        .expect("fixture should contain legacy note");
    sheet
        .update_comment(
            &note_id,
            CommentPatch {
                content: Some("Updated note".to_string()),
                ..Default::default()
            },
        )
        .expect("update legacy note");

    let threaded_id = sheet
        .iter_comments()
        .find(|(_, c)| c.kind == CommentKind::Threaded)
        .map(|(_, c)| c.id.clone())
        .expect("fixture should contain threaded comment");
    sheet
        .update_comment(
            &threaded_id,
            CommentPatch {
                content: Some("Updated thread root".to_string()),
                ..Default::default()
            },
        )
        .expect("update threaded comment");

    let saved = doc.save_to_vec().expect("save_to_vec");

    let pkg = XlsxPackage::from_bytes(&saved).expect("roundtrip should parse as xlsx package");
    let legacy = pkg
        .part("xl/comments1.xml")
        .expect("legacy comments part should exist");
    let legacy = std::str::from_utf8(legacy).expect("comments xml should be utf-8");
    assert!(
        legacy.contains("Updated note"),
        "expected updated legacy note content, got:\n{legacy}"
    );

    let threaded = pkg
        .part("xl/threadedComments/threadedComments1.xml")
        .expect("threaded comments part should exist");
    let threaded = std::str::from_utf8(threaded).expect("threaded comments xml should be utf-8");
    assert!(
        threaded.contains("Updated thread root"),
        "expected updated threaded root content, got:\n{threaded}"
    );

    // Unknown comment-related parts must remain byte-for-byte identical.
    let orig_pkg = XlsxPackage::from_bytes(&bytes).expect("fixture should parse as xlsx package");
    for path in [
        "xl/commentsExt1.xml",
        "xl/drawings/vmlDrawing1.vml",
        "xl/persons/persons1.xml",
        "xl/worksheets/_rels/sheet1.xml.rels",
    ] {
        let original = orig_pkg.part(path).unwrap_or_else(|| panic!("missing fixture part {path}"));
        let roundtrip = pkg.part(path).unwrap_or_else(|| panic!("missing roundtrip part {path}"));
        assert_eq!(
            roundtrip, original,
            "expected part to be preserved byte-for-byte: {path}"
        );
    }
}

