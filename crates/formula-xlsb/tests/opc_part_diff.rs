use std::collections::BTreeSet;
use std::path::{Path, PathBuf};

use formula_xlsb::{CellValue, XlsbWorkbook};
use pretty_assertions::assert_eq;

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
fn save_as_is_lossless_at_opc_part_level() {
    let fixture_path = fixture_path();
    let wb = XlsbWorkbook::open(&fixture_path).expect("open xlsb fixture");

    let tmpdir = tempfile::tempdir().expect("create temp dir");
    let out_path = tmpdir.path().join("roundtrip.xlsb");
    wb.save_as(&out_path).expect("save_as");

    let report = xlsx_diff::diff_workbooks(&fixture_path, &out_path).expect("diff workbooks");
    assert!(
        report.is_empty(),
        "expected no OPC part diffs, got:\n{}",
        format_report(&report)
    );
}

#[test]
fn patch_writer_changes_only_target_sheet_part() {
    let fixture_path = fixture_path();
    let wb = XlsbWorkbook::open(&fixture_path).expect("open xlsb fixture");

    let tmpdir = tempfile::tempdir().expect("create temp dir");
    let out_path = tmpdir.path().join("patched.xlsb");
    wb.save_with_edits(&out_path, 0, 0, 1, 123.0)
        .expect("save_with_edits");

    let patched = XlsbWorkbook::open(&out_path).expect("re-open patched workbook");
    let sheet = patched.read_sheet(0).expect("read patched sheet");
    let b1 = sheet
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 1)
        .expect("B1 exists");
    assert_eq!(b1.value, CellValue::Number(123.0));

    let report = xlsx_diff::diff_workbooks(&fixture_path, &out_path).expect("diff workbooks");
    assert!(
        report
            .differences
            .iter()
            .any(|d| d.part == "xl/worksheets/sheet1.bin"),
        "expected worksheet part to change, got:\n{}",
        format_report(&report)
    );

    let unexpected_missing: Vec<_> = report
        .differences
        .iter()
        .filter(|d| d.kind == "missing_part" && !is_calc_chain_part(&d.part))
        .map(|d| d.part.clone())
        .collect();
    assert!(
        unexpected_missing.is_empty(),
        "unexpected missing parts: {unexpected_missing:?}\n{}",
        format_report(&report)
    );

    let parts: BTreeSet<String> = report.differences.iter().map(|d| d.part.clone()).collect();
    let unexpected_parts: Vec<_> = parts
        .iter()
        .filter(|part| !is_allowed_patch_diff_part(part))
        .cloned()
        .collect();

    assert!(
        unexpected_parts.is_empty(),
        "unexpected diff parts: {unexpected_parts:?}\n{}",
        format_report(&report)
    );
}

fn is_allowed_patch_diff_part(part: &str) -> bool {
    part == "xl/worksheets/sheet1.bin" || is_calc_chain_part(part)
}

fn is_calc_chain_part(part: &str) -> bool {
    part.starts_with("xl/calcChain.")
}
