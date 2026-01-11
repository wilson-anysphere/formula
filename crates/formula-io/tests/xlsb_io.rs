use std::path::PathBuf;

use formula_io::{open_workbook, save_workbook};

fn xlsb_fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../formula-xlsb/tests/fixtures")
        .join(rel)
}

#[test]
fn roundtrips_xlsb_without_part_diffs() {
    let path = xlsb_fixture_path("simple.xlsb");
    let wb = open_workbook(&path).expect("open workbook");

    let dir = tempfile::tempdir().expect("temp dir");
    let out_path = dir.path().join("roundtrip.xlsb");
    save_workbook(&wb, &out_path).expect("save workbook");

    let report = xlsx_diff::diff_workbooks(&path, &out_path).expect("diff workbooks");
    assert!(
        report.is_empty(),
        "expected no part diffs, got {}\n{}",
        report.differences.len(),
        report
            .differences
            .iter()
            .map(|d| d.to_string())
            .collect::<String>()
    );
}

