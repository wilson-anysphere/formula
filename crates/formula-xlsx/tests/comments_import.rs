use std::fs;

use formula_model::{CellRef, CommentKind, Workbook};

fn assert_fixture_comments_present(workbook: &Workbook) {
    let sheet = workbook
        .sheets
        .iter()
        .find(|s| s.name == "Sheet1")
        .expect("fixture should contain Sheet1");

    let mut note = None;
    let mut threaded = None;
    for (_anchor, comment) in sheet.iter_comments() {
        match comment.kind {
            CommentKind::Note => note = Some(comment),
            CommentKind::Threaded => threaded = Some(comment),
        }
    }

    let note = note.expect("fixture should contain a legacy note comment");
    assert_eq!(note.cell_ref, CellRef::from_a1("A1").unwrap());
    assert_eq!(note.author.name, "Alex");
    assert_eq!(note.content, "Legacy note");

    let threaded = threaded.expect("fixture should contain a threaded comment");
    assert_eq!(threaded.cell_ref, CellRef::from_a1("B2").unwrap());
    assert_eq!(threaded.author.name, "Alex");
    assert_eq!(threaded.content, "Thread root");
    assert!(threaded.resolved, "fixture threaded comment should be resolved");

    let reply = threaded
        .replies
        .first()
        .expect("fixture threaded comment should have a reply");
    assert_eq!(reply.author.name, "Sam");
    assert_eq!(reply.content, "First reply");
}

#[test]
fn load_from_bytes_imports_sheet_comments() {
    let fixture_path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/comments.xlsx");
    let bytes = fs::read(fixture_path).expect("fixture workbook should be readable");

    let doc = formula_xlsx::load_from_bytes(&bytes).expect("load_from_bytes should succeed");
    assert_fixture_comments_present(&doc.workbook);
}

#[test]
fn fast_reader_imports_sheet_comments() {
    let fixture_path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/comments.xlsx");
    let bytes = fs::read(fixture_path).expect("fixture workbook should be readable");

    let workbook =
        formula_xlsx::read_workbook_model_from_bytes(&bytes).expect("fast reader should succeed");
    assert_fixture_comments_present(&workbook);
}

