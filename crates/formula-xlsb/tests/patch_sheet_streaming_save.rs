use std::path::{Path, PathBuf};

use formula_xlsb::rgce::{encode_rgce_with_context, CellCoord};
use formula_xlsb::{CellEdit, CellValue, XlsbWorkbook};

mod fixture_builder;
use fixture_builder::XlsbFixtureBuilder;

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
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
        new_style: None,
        clear_formula: false,
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
fn save_with_cell_edits_streaming_can_insert_missing_cells() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number(0, 0, 1.0);

    let tmpdir = tempfile::tempdir().expect("create temp dir");
    let fixture_path = tmpdir.path().join("input.xlsb");
    std::fs::write(&fixture_path, builder.build_bytes()).expect("write xlsb fixture");

    let wb = XlsbWorkbook::open(&fixture_path).expect("open xlsb fixture");

    let edits = [CellEdit {
        row: 5,
        col: 3,
        new_value: CellValue::Number(123.0),
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
        new_style: None,
        clear_formula: false,
    }];

    let in_memory_path = tmpdir.path().join("patched_in_memory_insert.xlsb");
    let streaming_path = tmpdir.path().join("patched_streaming_insert.xlsb");

    wb.save_with_cell_edits(&in_memory_path, 0, &edits)
        .expect("save_with_cell_edits");
    wb.save_with_cell_edits_streaming(&streaming_path, 0, &edits)
        .expect("save_with_cell_edits_streaming");

    let report =
        xlsx_diff::diff_workbooks(&in_memory_path, &streaming_path).expect("diff workbooks");
    assert!(
        report.is_empty(),
        "expected no OPC part diffs between in-memory and streaming insertion paths, got:\n{}",
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
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
        new_style: None,
        clear_formula: false,
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

#[test]
fn save_with_cell_edits_streaming_can_patch_formula_rgcb_bytes() {
    let ctx = formula_xlsb::workbook_context::WorkbookContext::default();

    let encoded_123 =
        encode_rgce_with_context("=SUM({1,2,3})", &ctx, CellCoord::new(0, 0)).expect("encode rgce");
    assert!(
        !encoded_123.rgcb.is_empty(),
        "expected array formula encoding to produce rgcb bytes"
    );

    let mut builder = XlsbFixtureBuilder::new();
    builder.set_sheet_name("ArrayRgcb");
    builder.set_cell_formula_num(
        0,
        0,
        6.0,
        encoded_123.rgce.clone(),
        encoded_123.rgcb.clone(),
    );

    let tmpdir = tempfile::tempdir().expect("create temp dir");
    let fixture_path = tmpdir.path().join("input.xlsb");
    std::fs::write(&fixture_path, builder.build_bytes()).expect("write xlsb fixture");

    let wb = XlsbWorkbook::open(&fixture_path).expect("open xlsb fixture");

    let encoded_45 =
        encode_rgce_with_context("=SUM({4,5})", &ctx, CellCoord::new(0, 0)).expect("encode rgce");
    assert_eq!(
        encoded_45.rgce, encoded_123.rgce,
        "expected SUM(array) formulas to share rgce so only rgcb changes"
    );

    let edits = [CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Number(9.0),
        new_formula: None,
        new_rgcb: Some(encoded_45.rgcb.clone()),
        new_formula_flags: None,
        shared_string_index: None,
        new_style: None,
        clear_formula: false,
    }];

    let in_memory_path = tmpdir.path().join("patched_in_memory_rgcb.xlsb");
    let streaming_path = tmpdir.path().join("patched_streaming_rgcb.xlsb");

    wb.save_with_cell_edits(&in_memory_path, 0, &edits)
        .expect("save_with_cell_edits");
    wb.save_with_cell_edits_streaming(&streaming_path, 0, &edits)
        .expect("save_with_cell_edits_streaming");

    let report =
        xlsx_diff::diff_workbooks(&in_memory_path, &streaming_path).expect("diff workbooks");
    assert!(
        report.is_empty(),
        "expected no OPC part diffs between in-memory and streaming rgcb edits, got:\n{}",
        format_report(&report)
    );

    let wb2 = XlsbWorkbook::open(&streaming_path).expect("open patched xlsb");
    let sheet = wb2.read_sheet(0).expect("read sheet");
    let cell = sheet
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 0)
        .expect("A1 exists");
    assert_eq!(cell.value, CellValue::Number(9.0));
    let formula = cell.formula.as_ref().expect("formula metadata");
    assert_eq!(formula.extra, encoded_45.rgcb);
    assert_eq!(formula.text.as_deref(), Some("SUM({4,5})"));
}

#[test]
fn save_with_cell_edits_streaming_can_insert_formula_rgcb_bytes() {
    let ctx = formula_xlsb::workbook_context::WorkbookContext::default();

    let encoded =
        encode_rgce_with_context("=SUM({4,5})", &ctx, CellCoord::new(5, 3)).expect("encode rgce");
    assert!(
        !encoded.rgcb.is_empty(),
        "expected array formula encoding to produce rgcb bytes"
    );

    let mut builder = XlsbFixtureBuilder::new();
    builder.set_sheet_name("ArrayRgcbInsert");
    builder.set_cell_number(0, 0, 1.0);

    let tmpdir = tempfile::tempdir().expect("create temp dir");
    let fixture_path = tmpdir.path().join("input.xlsb");
    std::fs::write(&fixture_path, builder.build_bytes()).expect("write xlsb fixture");

    let wb = XlsbWorkbook::open(&fixture_path).expect("open xlsb fixture");

    let edits = [CellEdit {
        row: 5,
        col: 3,
        new_value: CellValue::Number(9.0),
        new_formula: Some(encoded.rgce.clone()),
        new_rgcb: Some(encoded.rgcb.clone()),
        new_formula_flags: None,
        shared_string_index: None,
        new_style: None,
        clear_formula: false,
    }];

    let in_memory_path = tmpdir.path().join("patched_in_memory_rgcb_insert.xlsb");
    let streaming_path = tmpdir.path().join("patched_streaming_rgcb_insert.xlsb");

    wb.save_with_cell_edits(&in_memory_path, 0, &edits)
        .expect("save_with_cell_edits");
    wb.save_with_cell_edits_streaming(&streaming_path, 0, &edits)
        .expect("save_with_cell_edits_streaming");

    let report =
        xlsx_diff::diff_workbooks(&in_memory_path, &streaming_path).expect("diff workbooks");
    assert!(
        report.is_empty(),
        "expected no OPC part diffs between in-memory and streaming rgcb insertion edits, got:\n{}",
        format_report(&report)
    );

    let wb2 = XlsbWorkbook::open(&streaming_path).expect("open patched xlsb");
    let sheet = wb2.read_sheet(0).expect("read sheet");
    let cell = sheet
        .cells
        .iter()
        .find(|c| c.row == 5 && c.col == 3)
        .expect("D6 exists");
    assert_eq!(cell.value, CellValue::Number(9.0));
    let formula = cell.formula.as_ref().expect("formula metadata");
    assert_eq!(formula.extra, encoded.rgcb);
    assert_eq!(formula.text.as_deref(), Some("SUM({4,5})"));
}
