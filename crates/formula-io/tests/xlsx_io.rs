use std::path::PathBuf;

use formula_io::{open_workbook, save_workbook, Workbook};
use formula_xlsx::XlsxPackage;
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

#[test]
fn saving_xlsm_as_xlsx_strips_vba_project() {
    let path = fixture_path("xlsx/macros/basic.xlsm");
    let wb = open_workbook(&path).expect("open workbook");

    let dir = tempfile::tempdir().expect("temp dir");
    let out_path = dir.path().join("out.xlsx");
    save_workbook(&wb, &out_path).expect("save workbook");

    let bytes = std::fs::read(&out_path).expect("read saved workbook");
    let pkg = XlsxPackage::from_bytes(&bytes).expect("re-open saved package");

    assert!(
        pkg.vba_project_bin().is_none(),
        "expected VBA project to be stripped when saving `.xlsm` as `.xlsx`"
    );
    assert!(
        pkg.part("xl/vbaProject.bin").is_none(),
        "expected xl/vbaProject.bin to be removed"
    );
    assert!(
        pkg.part("xl/vbaProjectSignature.bin").is_none(),
        "expected xl/vbaProjectSignature.bin to be removed"
    );

    let content_types =
        std::str::from_utf8(pkg.part("[Content_Types].xml").expect("content types")).unwrap();
    assert!(
        !content_types.contains("macroEnabled"),
        "expected `[Content_Types].xml` to no longer advertise a macro-enabled workbook"
    );
    assert!(
        !content_types.contains("vbaProject.bin"),
        "expected `[Content_Types].xml` to no longer reference vbaProject.bin"
    );

    let wb_rels =
        std::str::from_utf8(pkg.part("xl/_rels/workbook.xml.rels").expect("workbook rels"))
            .unwrap();
    assert!(
        !wb_rels.contains("vbaProject"),
        "expected workbook relationships to no longer reference vbaProject"
    );
}

#[test]
fn saving_xlsm_as_xlsm_preserves_vba_project() {
    let path = fixture_path("xlsx/macros/basic.xlsm");
    let wb = open_workbook(&path).expect("open workbook");

    let dir = tempfile::tempdir().expect("temp dir");
    let out_path = dir.path().join("out.xlsm");
    save_workbook(&wb, &out_path).expect("save workbook");

    let bytes = std::fs::read(&out_path).expect("read saved workbook");
    let pkg = XlsxPackage::from_bytes(&bytes).expect("re-open saved package");

    assert!(
        pkg.vba_project_bin().is_some(),
        "expected VBA project to be preserved when saving `.xlsm` as `.xlsm`"
    );
}
