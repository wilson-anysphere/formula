use std::path::PathBuf;

use formula_io::{open_workbook, save_workbook, Workbook};
use xlsx_diff::Severity;

fn fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../fixtures").join(rel)
}

#[test]
fn opens_basic_xlsx_fixture() {
    let path = fixture_path("xlsx/basic/basic.xlsx");
    let wb = open_workbook(&path).expect("open workbook");

    match wb {
        Workbook::Xlsx(pkg) => {
            assert!(pkg.part("xl/workbook.xml").is_some());
        }
        other => panic!("expected Workbook::Xlsx, got {other:?}"),
    }
}

#[test]
fn opens_xlsm_fixture_and_preserves_vba_project() {
    let path = fixture_path("xlsx/macros/basic.xlsm");
    let wb = open_workbook(&path).expect("open workbook");

    match wb {
        Workbook::Xlsx(pkg) => {
            assert!(
                pkg.vba_project_bin().is_some(),
                "expected xl/vbaProject.bin to be present"
            );
        }
        other => panic!("expected Workbook::Xlsx, got {other:?}"),
    }
}

#[test]
fn roundtrips_xlsx_package_without_critical_diffs() {
    let path = fixture_path("xlsx/basic/basic.xlsx");
    let wb = open_workbook(&path).expect("open workbook");

    let dir = tempfile::tempdir().expect("temp dir");
    let out_path = dir.path().join("roundtrip.xlsx");
    save_workbook(&wb, &out_path).expect("save workbook");

    let report = xlsx_diff::diff_workbooks(&path, &out_path).expect("diff workbooks");
    let critical = report.count(Severity::Critical);
    assert_eq!(
        critical, 0,
        "expected no critical diffs, got {critical}\n{}",
        report
            .differences
            .iter()
            .map(|d| d.to_string())
            .collect::<String>()
    );
}

