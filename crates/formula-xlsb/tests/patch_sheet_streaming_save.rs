use std::path::{Path, PathBuf};

use formula_xlsb::{CellEdit, CellValue, XlsbWorkbook};

fn fixture_path() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/simple.xlsb")
}

fn format_report(report: &xlsx_diff::DiffReport) -> String {
    report
        .differences
        .iter()
        .map(|d| d.to_string())
        .collect::<Vec<_>>()
        .join("\n")
}

#[test]
fn save_with_cell_edits_streaming_matches_in_memory_patch_path() {
    let fixture_path = fixture_path();
    let wb = XlsbWorkbook::open(&fixture_path).expect("open xlsb fixture");

    let edits = [CellEdit {
        row: 0,
        col: 1,
        new_value: CellValue::Number(123.0),
        new_formula: None,
        shared_string_index: None,
    }];

    let tmpdir = tempfile::tempdir().expect("create temp dir");
    let in_memory_path = tmpdir.path().join("patched_in_memory.xlsb");
    let streaming_path = tmpdir.path().join("patched_streaming.xlsb");

    wb.save_with_cell_edits(&in_memory_path, 0, &edits)
        .expect("save_with_cell_edits");
    wb.save_with_cell_edits_streaming(&streaming_path, 0, &edits)
        .expect("save_with_cell_edits_streaming");

    let report =
        xlsx_diff::diff_workbooks(&in_memory_path, &streaming_path).expect("diff workbooks");
    assert!(
        report.is_empty(),
        "expected no OPC part diffs between in-memory and streaming paths, got:\n{}",
        format_report(&report)
    );
}

#[test]
fn save_with_cell_edits_streaming_is_lossless_for_noop_edit() {
    let fixture_path = fixture_path();
    let wb = XlsbWorkbook::open(&fixture_path).expect("open xlsb fixture");

    // In `tests/fixtures/simple.xlsb`, B1 is 42.5.
    let edits = [CellEdit {
        row: 0,
        col: 1,
        new_value: CellValue::Number(42.5),
        new_formula: None,
        shared_string_index: None,
    }];

    let tmpdir = tempfile::tempdir().expect("create temp dir");
    let out_path = tmpdir.path().join("noop_streaming.xlsb");

    wb.save_with_cell_edits_streaming(&out_path, 0, &edits)
        .expect("save_with_cell_edits_streaming");

    let report = xlsx_diff::diff_workbooks(&fixture_path, &out_path).expect("diff workbooks");
    assert!(
        report.is_empty(),
        "expected no OPC part diffs for no-op streaming edit, got:\n{}",
        format_report(&report)
    );
}
