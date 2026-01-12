use std::io::Write;

use formula_model::{CellRef, CommentKind, Range};

mod common;

use common::xls_fixture_builder;

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

#[test]
fn imports_note_comment_records() {
    let bytes = xls_fixture_builder::build_note_comment_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("Notes")
        .expect("Notes missing");

    let a1 = CellRef::from_a1("A1").unwrap();
    let comments = sheet.comments_for_cell(a1);
    assert_eq!(comments.len(), 1, "expected 1 comment on A1");

    let comment = &comments[0];
    assert_eq!(comment.kind, CommentKind::Note);
    assert_eq!(comment.content, "Hello from note");
    assert_eq!(comment.author.name, "Alice");
}

#[test]
fn anchors_note_comments_to_merged_region_top_left_cell() {
    let bytes = xls_fixture_builder::build_note_comment_in_merged_region_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("MergedNotes")
        .expect("MergedNotes missing");

    let merge_range = Range::from_a1("A1:B1").unwrap();
    assert!(
        sheet.merged_regions.iter().any(|region| region.range == merge_range),
        "missing expected merged range A1:B1"
    );

    // The NOTE record in this fixture targets B1, but the model anchors comments
    // to the merged region's top-left cell (A1).
    let a1 = CellRef::from_a1("A1").unwrap();
    let comments = sheet.comments_for_cell(a1);
    assert_eq!(comments.len(), 1, "expected 1 comment on A1");
    assert_eq!(comments[0].content, "Hello from note");
}

