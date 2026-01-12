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
    let result2 = import_fixture(&bytes);

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
    assert_eq!(comment.id, "xls-note:A1:1");

    let sheet2 = result2
        .workbook
        .sheet_by_name("Notes")
        .expect("Notes missing (2)");
    let comments2 = sheet2.comments_for_cell(a1);
    assert_eq!(comments2.len(), 1, "expected 1 comment on A1 (2)");
    assert_eq!(comments2[0].id, comment.id, "comment ids should be stable across repeated imports");
}

#[test]
fn imports_note_comment_records_biff5() {
    let bytes = xls_fixture_builder::build_note_comment_biff5_fixture_xls();
    let result = import_fixture(&bytes);
    let result2 = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("NotesBiff5")
        .expect("NotesBiff5 missing");

    let a1 = CellRef::from_a1("A1").unwrap();
    let comments = sheet.comments_for_cell(a1);
    assert_eq!(comments.len(), 1, "expected 1 comment on A1");

    let comment = &comments[0];
    assert_eq!(comment.kind, CommentKind::Note);
    assert_eq!(comment.content, "Hi \u{0410}");
    assert_eq!(comment.author.name, "\u{0410}");
    assert_eq!(comment.id, "xls-note:A1:1");

    let sheet2 = result2
        .workbook
        .sheet_by_name("NotesBiff5")
        .expect("NotesBiff5 missing (2)");
    let comments2 = sheet2.comments_for_cell(a1);
    assert_eq!(comments2.len(), 1, "expected 1 comment on A1 (2)");
    assert_eq!(comments2[0].id, comment.id, "comment ids should be stable across repeated imports");
}

#[test]
fn imports_note_comment_text_split_across_multiple_continue_records_biff5() {
    let bytes = xls_fixture_builder::build_note_comment_biff5_split_across_continues_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("NotesBiff5Split")
        .expect("NotesBiff5Split missing");

    let a1 = CellRef::from_a1("A1").unwrap();
    let comments = sheet.comments_for_cell(a1);
    assert_eq!(comments.len(), 1, "expected 1 comment on A1");
    assert_eq!(comments[0].content, "Hi \u{0410}");
    assert_eq!(comments[0].author.name, "\u{0410}");
    assert_eq!(comments[0].id, "xls-note:A1:1");
}

#[test]
fn imports_note_comment_text_split_across_multiple_continue_records_with_flags_biff5() {
    let bytes =
        xls_fixture_builder::build_note_comment_biff5_split_across_continues_with_flags_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("NotesBiff5SplitFlags")
        .expect("NotesBiff5SplitFlags missing");

    let a1 = CellRef::from_a1("A1").unwrap();
    let comments = sheet.comments_for_cell(a1);
    assert_eq!(comments.len(), 1, "expected 1 comment on A1");
    assert_eq!(comments[0].content, "Hi \u{0410}");
    assert_eq!(comments[0].author.name, "\u{0410}");
    assert_eq!(comments[0].id, "xls-note:A1:1");
}

#[test]
fn imports_note_comment_text_split_across_continue_records_using_multibyte_codepage_biff5() {
    let bytes =
        xls_fixture_builder::build_note_comment_biff5_split_across_continues_codepage_932_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("NotesBiff5SplitCp932")
        .expect("NotesBiff5SplitCp932 missing");

    let a1 = CellRef::from_a1("A1").unwrap();
    let comments = sheet.comments_for_cell(a1);
    assert_eq!(comments.len(), 1, "expected 1 comment on A1");
    assert_eq!(comments[0].content, "あ");
    assert_eq!(comments[0].author.name, "あ");
    assert_eq!(comments[0].id, "xls-note:A1:1");
}

#[test]
fn imports_note_comment_author_stored_as_biff8_short_string_biff5() {
    let bytes = xls_fixture_builder::build_note_comment_biff5_author_biff8_short_string_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("NotesBiff5AuthorBiff8")
        .expect("NotesBiff5AuthorBiff8 missing");

    let a1 = CellRef::from_a1("A1").unwrap();
    let comments = sheet.comments_for_cell(a1);
    assert_eq!(comments.len(), 1, "expected 1 comment on A1");
    assert_eq!(comments[0].content, "Hi");
    assert_eq!(comments[0].author.name, "\u{0410}");
    assert_eq!(comments[0].id, "xls-note:A1:1");
}

#[test]
fn imports_note_comment_author_stored_as_biff8_unicode_string_biff5() {
    let bytes =
        xls_fixture_builder::build_note_comment_biff5_author_biff8_unicode_string_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("NotesBiff5AuthorBiff8Unicode")
        .expect("NotesBiff5AuthorBiff8Unicode missing");

    let a1 = CellRef::from_a1("A1").unwrap();
    let comments = sheet.comments_for_cell(a1);
    assert_eq!(comments.len(), 1, "expected 1 comment on A1");
    assert_eq!(comments[0].content, "Hi");
    assert_eq!(comments[0].author.name, "\u{0410}");
    assert_eq!(comments[0].id, "xls-note:A1:1");
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
    assert_eq!(comments[0].id, "xls-note:A1:1");
}

#[test]
fn anchors_note_comments_to_merged_region_top_left_cell_biff5() {
    let bytes = xls_fixture_builder::build_note_comment_biff5_in_merged_region_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("MergedNotesBiff5")
        .expect("MergedNotesBiff5 missing");

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
    assert_eq!(comments[0].content, "Hello");
    assert_eq!(comments[0].id, "xls-note:A1:1");
}

#[test]
fn imports_note_comment_text_using_workbook_codepage() {
    let bytes = xls_fixture_builder::build_note_comment_codepage_1251_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("NotesCp1251")
        .expect("NotesCp1251 missing");

    let a1 = CellRef::from_a1("A1").unwrap();
    let comments = sheet.comments_for_cell(a1);
    assert_eq!(comments.len(), 1, "expected 1 comment on A1");
    assert_eq!(comments[0].content, "\u{0410}");
    assert_eq!(comments[0].id, "xls-note:A1:1");
}

#[test]
fn imports_note_comment_author_using_workbook_codepage() {
    let bytes = xls_fixture_builder::build_note_comment_author_codepage_1251_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("NotesAuthorCp1251")
        .expect("NotesAuthorCp1251 missing");

    let a1 = CellRef::from_a1("A1").unwrap();
    let comments = sheet.comments_for_cell(a1);
    assert_eq!(comments.len(), 1, "expected 1 comment on A1");
    assert_eq!(comments[0].author.name, "\u{0410}");
    assert_eq!(comments[0].content, "Hello");
    assert_eq!(comments[0].id, "xls-note:A1:1");
}

#[test]
fn imports_note_comment_author_stored_as_xl_unicode_string() {
    let bytes = xls_fixture_builder::build_note_comment_author_xl_unicode_string_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("NotesAuthorXlUnicode")
        .expect("NotesAuthorXlUnicode missing");

    let a1 = CellRef::from_a1("A1").unwrap();
    let comments = sheet.comments_for_cell(a1);
    assert_eq!(comments.len(), 1, "expected 1 comment on A1");
    assert_eq!(comments[0].author.name, "Alice");
    assert_eq!(comments[0].content, "Hello");
    assert_eq!(comments[0].id, "xls-note:A1:1");
}

#[test]
fn imports_note_comment_author_missing_biff8_flags_byte() {
    let bytes = xls_fixture_builder::build_note_comment_author_missing_flags_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("NotesAuthorNoFlags")
        .expect("NotesAuthorNoFlags missing");

    let a1 = CellRef::from_a1("A1").unwrap();
    let comments = sheet.comments_for_cell(a1);
    assert_eq!(comments.len(), 1, "expected 1 comment on A1");
    assert_eq!(comments[0].author.name, "Alice");
    assert_eq!(comments[0].content, "Hello");
    assert_eq!(comments[0].id, "xls-note:A1:1");
}

#[test]
fn imports_note_comment_text_missing_biff8_flags_byte() {
    let bytes = xls_fixture_builder::build_note_comment_txo_text_missing_flags_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("NotesTxoTextNoFlags")
        .expect("NotesTxoTextNoFlags missing");

    let a1 = CellRef::from_a1("A1").unwrap();
    let comments = sheet.comments_for_cell(a1);
    assert_eq!(comments.len(), 1, "expected 1 comment on A1");
    assert_eq!(comments[0].author.name, "Alice");
    assert_eq!(comments[0].content, "Hello");
    assert_eq!(comments[0].id, "xls-note:A1:1");
}

#[test]
fn imports_note_comment_text_missing_biff8_flags_byte_in_second_fragment() {
    let bytes =
        xls_fixture_builder::build_note_comment_txo_text_missing_flags_in_second_fragment_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("NotesTxoTextNoFlagsMid")
        .expect("NotesTxoTextNoFlagsMid missing");

    let a1 = CellRef::from_a1("A1").unwrap();
    let comments = sheet.comments_for_cell(a1);
    assert_eq!(comments.len(), 1, "expected 1 comment on A1");
    assert_eq!(comments[0].author.name, "Alice");
    assert_eq!(comments[0].content, "Hello");
    assert_eq!(comments[0].id, "xls-note:A1:1");
}

#[test]
fn imports_note_comment_when_txo_cch_text_is_zero() {
    let bytes = xls_fixture_builder::build_note_comment_txo_cch_text_zero_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("NotesTxoCchZero")
        .expect("NotesTxoCchZero missing");

    let a1 = CellRef::from_a1("A1").unwrap();
    let comments = sheet.comments_for_cell(a1);
    assert_eq!(comments.len(), 1, "expected 1 comment on A1");
    assert_eq!(comments[0].content, "Hello");
    assert_eq!(comments[0].author.name, "Alice");
    assert_eq!(comments[0].id, "xls-note:A1:1");
}

#[test]
fn imports_note_comment_text_split_across_multiple_continue_records() {
    let bytes = xls_fixture_builder::build_note_comment_split_across_continues_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("NotesSplit")
        .expect("NotesSplit missing");

    let a1 = CellRef::from_a1("A1").unwrap();
    let comments = sheet.comments_for_cell(a1);
    assert_eq!(comments.len(), 1, "expected 1 comment on A1");
    assert_eq!(comments[0].content, "ABCDE");
    assert_eq!(comments[0].id, "xls-note:A1:1");
}

#[test]
fn skips_note_comment_when_txo_payload_is_missing() {
    let bytes = xls_fixture_builder::build_note_comment_missing_txo_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("NotesMissingTxo")
        .expect("NotesMissingTxo missing");

    let a1 = CellRef::from_a1("A1").unwrap();
    assert!(
        sheet.comments_for_cell(a1).is_empty(),
        "expected no comments when TXO payload is missing"
    );

    assert!(
        result
            .warnings
            .iter()
            .any(|w| w.message.contains("missing TXO payload")),
        "expected missing TXO warning; warnings={:?}",
        result
            .warnings
            .iter()
            .map(|w| w.message.as_str())
            .collect::<Vec<_>>()
    );
}

#[test]
fn skips_note_comment_when_txo_payload_is_missing_biff5() {
    let bytes = xls_fixture_builder::build_note_comment_biff5_missing_txo_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("NotesBiff5MissingTxo")
        .expect("NotesBiff5MissingTxo missing");

    let a1 = CellRef::from_a1("A1").unwrap();
    assert!(
        sheet.comments_for_cell(a1).is_empty(),
        "expected no comments when TXO payload is missing"
    );

    assert!(
        result
            .warnings
            .iter()
            .any(|w| w.message.contains("missing TXO payload")),
        "expected missing TXO warning; warnings={:?}",
        result
            .warnings
            .iter()
            .map(|w| w.message.as_str())
            .collect::<Vec<_>>()
    );
}

#[test]
fn imports_note_comment_text_split_across_continue_records_with_mixed_encoding_flags() {
    let bytes =
        xls_fixture_builder::build_note_comment_split_across_continues_mixed_encoding_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("NotesSplitMixed")
        .expect("NotesSplitMixed missing");

    let a1 = CellRef::from_a1("A1").unwrap();
    let comments = sheet.comments_for_cell(a1);
    assert_eq!(comments.len(), 1, "expected 1 comment on A1");
    assert_eq!(comments[0].content, "ABCDE");
    assert_eq!(comments[0].id, "xls-note:A1:1");
}

#[test]
fn imports_note_comment_text_split_across_continue_records_using_multibyte_codepage() {
    let bytes =
        xls_fixture_builder::build_note_comment_split_across_continues_codepage_932_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("NotesSplitCp932")
        .expect("NotesSplitCp932 missing");

    let a1 = CellRef::from_a1("A1").unwrap();
    let comments = sheet.comments_for_cell(a1);
    assert_eq!(comments.len(), 1, "expected 1 comment on A1");
    assert_eq!(comments[0].content, "あ");
    assert_eq!(comments[0].id, "xls-note:A1:1");
}

#[test]
fn imports_note_comment_text_split_across_continue_records_using_utf8_codepage() {
    let bytes =
        xls_fixture_builder::build_note_comment_split_across_continues_codepage_65001_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("NotesSplitUtf8")
        .expect("NotesSplitUtf8 missing");

    let a1 = CellRef::from_a1("A1").unwrap();
    let comments = sheet.comments_for_cell(a1);
    assert_eq!(comments.len(), 1, "expected 1 comment on A1");
    assert_eq!(comments[0].content, "€");
    assert_eq!(comments[0].id, "xls-note:A1:1");
}

#[test]
fn imports_note_comment_unicode_text_split_mid_code_unit_across_continue_records() {
    let bytes =
        xls_fixture_builder::build_note_comment_split_utf16_code_unit_across_continues_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("NotesSplitUtf16Odd")
        .expect("NotesSplitUtf16Odd missing");

    let a1 = CellRef::from_a1("A1").unwrap();
    let comments = sheet.comments_for_cell(a1);
    assert_eq!(comments.len(), 1, "expected 1 comment on A1");
    assert_eq!(comments[0].content, "€");
    assert_eq!(comments[0].id, "xls-note:A1:1");
}

#[test]
fn imports_note_comment_when_txo_header_is_missing() {
    let bytes = xls_fixture_builder::build_note_comment_missing_txo_header_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("NotesMissingTxoHeader")
        .expect("NotesMissingTxoHeader missing");

    let a1 = CellRef::from_a1("A1").unwrap();
    let comments = sheet.comments_for_cell(a1);
    assert_eq!(comments.len(), 1, "expected 1 comment on A1");
    assert_eq!(comments[0].content, "Hello");
    assert_eq!(comments[0].author.name, "Alice");
    assert_eq!(comments[0].id, "xls-note:A1:1");

    assert!(
        result
            .warnings
            .iter()
            .any(|w| w.message.contains("malformed header") || w.message.contains("falling back")),
        "expected malformed-header warning; warnings={:?}",
        result
            .warnings
            .iter()
            .map(|w| w.message.as_str())
            .collect::<Vec<_>>()
    );
}

#[test]
fn imports_note_comment_when_txo_header_is_truncated_and_missing_cb_runs() {
    let bytes =
        xls_fixture_builder::build_note_comment_truncated_txo_header_missing_cb_runs_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("NotesTxoHeaderNoCbRuns")
        .expect("NotesTxoHeaderNoCbRuns missing");

    let a1 = CellRef::from_a1("A1").unwrap();
    let comments = sheet.comments_for_cell(a1);
    assert_eq!(comments.len(), 1, "expected 1 comment on A1");
    assert_eq!(comments[0].content, "Hello");
    assert_eq!(comments[0].author.name, "Alice");
    assert_eq!(comments[0].id, "xls-note:A1:1");

    assert!(
        result.warnings.iter().any(|w| w.message.contains("truncated text")),
        "expected truncation warning; warnings={:?}",
        result
            .warnings
            .iter()
            .map(|w| w.message.as_str())
            .collect::<Vec<_>>()
    );
}

#[test]
fn imports_note_comment_when_txo_cch_text_is_at_alternate_offset() {
    let bytes = xls_fixture_builder::build_note_comment_txo_cch_text_offset_4_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("NotesTxoCchOffset4")
        .expect("NotesTxoCchOffset4 missing");

    let a1 = CellRef::from_a1("A1").unwrap();
    let comments = sheet.comments_for_cell(a1);
    assert_eq!(comments.len(), 1, "expected 1 comment on A1");
    assert_eq!(comments[0].content, "Hi");
    assert_eq!(comments[0].author.name, "Alice");
    assert_eq!(comments[0].id, "xls-note:A1:1");
}

#[test]
fn imports_note_comment_when_note_obj_id_fields_are_swapped() {
    let bytes = xls_fixture_builder::build_note_comment_note_obj_id_swapped_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("NotesObjIdSwapped")
        .expect("NotesObjIdSwapped missing");

    let a1 = CellRef::from_a1("A1").unwrap();
    let comments = sheet.comments_for_cell(a1);
    assert_eq!(comments.len(), 1, "expected 1 comment on A1");
    assert_eq!(comments[0].content, "Hello from swapped obj id");
    assert_eq!(comments[0].author.name, "Alice");
    assert_eq!(comments[0].id, "xls-note:A1:2");
}
