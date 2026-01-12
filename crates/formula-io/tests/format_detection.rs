use std::path::PathBuf;

use formula_io::{open_workbook, Workbook};

use formula_io::{detect_workbook_format, WorkbookFormat};

fn root_fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../fixtures").join(rel)
}

fn xlsb_fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../formula-xlsb/tests/fixtures")
        .join(rel)
}

fn xls_fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../formula-xls/tests/fixtures")
        .join(rel)
}

#[test]
fn opens_xlsx_even_with_unknown_extension() {
    let src = root_fixture_path("xlsx/basic/basic.xlsx");
    let dir = tempfile::tempdir().expect("temp dir");
    let dst = dir.path().join("basic.bin");
    std::fs::copy(&src, &dst).expect("copy fixture");

    let wb = open_workbook(&dst).expect("open workbook");
    match wb {
        Workbook::Xlsx(pkg) => {
            assert!(pkg.part("xl/workbook.xml").is_some());
        }
        other => panic!("expected Workbook::Xlsx, got {other:?}"),
    }
}

#[test]
fn opens_xlsb_even_with_unknown_extension() {
    let src = xlsb_fixture_path("simple.xlsb");
    let dir = tempfile::tempdir().expect("temp dir");
    let dst = dir.path().join("simple.bin");
    std::fs::copy(&src, &dst).expect("copy fixture");

    let wb = open_workbook(&dst).expect("open workbook");
    match wb {
        Workbook::Xlsb(_) => {}
        other => panic!("expected Workbook::Xlsb, got {other:?}"),
    }
}

#[test]
fn opens_xls_even_with_unknown_extension() {
    let src = xls_fixture_path("basic.xls");
    let dir = tempfile::tempdir().expect("temp dir");
    let dst = dir.path().join("basic.bin");
    std::fs::copy(&src, &dst).expect("copy fixture");

    let wb = open_workbook(&dst).expect("open workbook");
    match wb {
        Workbook::Xls(_) => {}
        other => panic!("expected Workbook::Xls, got {other:?}"),
    }
}

#[test]
fn detect_workbook_format_sniffs_csv_with_wrong_extension() {
    let dir = tempfile::tempdir().expect("temp dir");
    let path = dir.path().join("data.xlsx");
    std::fs::write(&path, "col1,col2\n1,hello\n2,world\n").expect("write csv");

    let fmt = detect_workbook_format(&path).expect("detect format");
    assert_eq!(fmt, WorkbookFormat::Csv);
}

#[test]
fn detect_workbook_format_sniffs_extensionless_csv() {
    let dir = tempfile::tempdir().expect("temp dir");
    let path = dir.path().join("data");
    std::fs::write(&path, "col1,col2\n1,hello\n2,world\n").expect("write csv");

    let fmt = detect_workbook_format(&path).expect("detect format");
    assert_eq!(fmt, WorkbookFormat::Csv);
}
