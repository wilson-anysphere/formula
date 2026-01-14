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

#[test]
fn comments_writeback_updates_legacy_only_preserves_threaded_part() {
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
                content: Some("Updated note only".to_string()),
                ..Default::default()
            },
        )
        .expect("update legacy note");

    let saved = doc.save_to_vec().expect("save_to_vec");

    let orig_pkg = XlsxPackage::from_bytes(&bytes).expect("fixture should parse as xlsx package");
    let pkg = XlsxPackage::from_bytes(&saved).expect("roundtrip should parse as xlsx package");

    let legacy = pkg
        .part("xl/comments1.xml")
        .expect("legacy comments part should exist");
    let legacy = std::str::from_utf8(legacy).expect("comments xml should be utf-8");
    assert!(
        legacy.contains("Updated note only"),
        "expected updated legacy note content, got:\n{legacy}"
    );

    // Threaded comment part should be preserved byte-for-byte when it is not edited.
    let original_threaded = orig_pkg
        .part("xl/threadedComments/threadedComments1.xml")
        .expect("fixture threaded comments part should exist");
    let roundtrip_threaded = pkg
        .part("xl/threadedComments/threadedComments1.xml")
        .expect("roundtrip threaded comments part should exist");
    assert_eq!(
        roundtrip_threaded, original_threaded,
        "expected threaded comment part to be preserved byte-for-byte when only legacy notes changed"
    );
}

#[test]
fn comments_writeback_updates_threaded_only_preserves_legacy_part() {
    let fixture_path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/comments.xlsx");
    let bytes = fs::read(fixture_path).expect("fixture workbook should be readable");

    let mut doc = load_from_bytes(&bytes).expect("load_from_bytes");
    let sheet_id = doc.workbook.sheets[0].id;
    let sheet = doc
        .workbook
        .sheet_mut(sheet_id)
        .expect("fixture sheet should exist");

    let threaded_id = sheet
        .iter_comments()
        .find(|(_, c)| c.kind == CommentKind::Threaded)
        .map(|(_, c)| c.id.clone())
        .expect("fixture should contain threaded comment");
    sheet
        .update_comment(
            &threaded_id,
            CommentPatch {
                content: Some("Updated thread only".to_string()),
                ..Default::default()
            },
        )
        .expect("update threaded comment");

    let saved = doc.save_to_vec().expect("save_to_vec");

    let orig_pkg = XlsxPackage::from_bytes(&bytes).expect("fixture should parse as xlsx package");
    let pkg = XlsxPackage::from_bytes(&saved).expect("roundtrip should parse as xlsx package");

    let threaded = pkg
        .part("xl/threadedComments/threadedComments1.xml")
        .expect("threaded comments part should exist");
    let threaded = std::str::from_utf8(threaded).expect("threaded comments xml should be utf-8");
    assert!(
        threaded.contains("Updated thread only"),
        "expected updated threaded comment content, got:\n{threaded}"
    );

    // Legacy comment part should be preserved byte-for-byte when it is not edited.
    let original_legacy = orig_pkg
        .part("xl/comments1.xml")
        .expect("fixture legacy comments part should exist");
    let roundtrip_legacy = pkg
        .part("xl/comments1.xml")
        .expect("roundtrip legacy comments part should exist");
    assert_eq!(
        roundtrip_legacy, original_legacy,
        "expected legacy comment part to be preserved byte-for-byte when only threaded comments changed"
    );
}

#[test]
fn comments_writeback_updates_threaded_comment_when_rels_type_is_noncanonical() {
    let fixture_path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/comments.xlsx");
    let bytes = fs::read(fixture_path).expect("fixture workbook should be readable");

    // Simulate a workbook producer that uses a non-canonical relationship Type URI for threaded
    // comments. The importer is tolerant (it matches any URI containing "threadedComment"), and the
    // write-back path should be equally tolerant so we can preserve and rewrite the existing part.
    let mut pkg = XlsxPackage::from_bytes(&bytes).expect("fixture should parse as xlsx package");
    let rels_path = "xl/worksheets/_rels/sheet1.xml.rels";
    let rels = pkg
        .part(rels_path)
        .unwrap_or_else(|| panic!("missing fixture part {rels_path}"));
    let rels = std::str::from_utf8(rels).expect("sheet rels should be utf-8");

    let canonical = "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment";
    assert!(
        rels.contains(canonical),
        "fixture sheet rels should contain threaded comment relationship"
    );
    let modified_rels = rels.replace(
        canonical,
        "http://schemas.microsoft.com/office/2019/10/relationships/threadedComment",
    );
    let modified_rels_bytes = modified_rels.clone().into_bytes();
    pkg.set_part(rels_path, modified_rels_bytes.clone());
    let modified_bytes = pkg.write_to_bytes().expect("write modified fixture package");

    // Baseline package after our `.rels` mutation (used to verify preservation).
    let modified_pkg =
        XlsxPackage::from_bytes(&modified_bytes).expect("modified fixture should parse");

    let mut doc = load_from_bytes(&modified_bytes).expect("load_from_bytes");
    let sheet_id = doc.workbook.sheets[0].id;
    let sheet = doc
        .workbook
        .sheet_mut(sheet_id)
        .expect("fixture sheet should exist");

    let threaded_id = sheet
        .iter_comments()
        .find(|(_, c)| c.kind == CommentKind::Threaded)
        .map(|(_, c)| c.id.clone())
        .expect("fixture should contain threaded comment");
    sheet
        .update_comment(
            &threaded_id,
            CommentPatch {
                content: Some("Updated threaded w/ noncanonical rel".to_string()),
                ..Default::default()
            },
        )
        .expect("update threaded comment");

    let saved = doc.save_to_vec().expect("save_to_vec");
    let roundtrip_pkg = XlsxPackage::from_bytes(&saved).expect("roundtrip should parse as xlsx package");

    let threaded = roundtrip_pkg
        .part("xl/threadedComments/threadedComments1.xml")
        .expect("threaded comments part should exist");
    let threaded = std::str::from_utf8(threaded).expect("threaded comments xml should be utf-8");
    assert!(
        threaded.contains("Updated threaded w/ noncanonical rel"),
        "expected updated threaded comment content, got:\n{threaded}"
    );

    // Ensure we did not touch the worksheet relationships while rewriting comment parts.
    let roundtrip_rels = roundtrip_pkg
        .part(rels_path)
        .expect("roundtrip sheet rels should exist");
    assert_eq!(
        roundtrip_rels,
        modified_rels_bytes.as_slice(),
        "expected worksheet .rels to be preserved byte-for-byte"
    );

    // Unknown comment-related parts must remain byte-for-byte identical.
    for path in [
        "xl/commentsExt1.xml",
        "xl/drawings/vmlDrawing1.vml",
        "xl/persons/persons1.xml",
    ] {
        let original = modified_pkg.part(path).unwrap_or_else(|| panic!("missing part {path}"));
        let roundtrip = roundtrip_pkg.part(path).unwrap_or_else(|| panic!("missing part {path}"));
        assert_eq!(
            roundtrip, original,
            "expected part to be preserved byte-for-byte: {path}"
        );
    }
}
