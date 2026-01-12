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
fn detect_workbook_format_sniffs_csv_with_xls_extension() {
    let dir = tempfile::tempdir().expect("temp dir");
    let path = dir.path().join("data.xls");
    std::fs::write(&path, "col1,col2\n1,hello\n2,world\n").expect("write csv");

    let fmt = detect_workbook_format(&path).expect("detect format");
    assert_eq!(fmt, WorkbookFormat::Csv);
}

#[test]
fn detect_workbook_format_sniffs_csv_with_xlsb_extension() {
    let dir = tempfile::tempdir().expect("temp dir");
    let path = dir.path().join("data.xlsb");
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

#[test]
fn detect_workbook_format_sniffs_single_line_csv_with_wrong_extension() {
    let dir = tempfile::tempdir().expect("temp dir");
    let path = dir.path().join("data.xlsx");
    std::fs::write(&path, "a,b").expect("write csv");

    let fmt = detect_workbook_format(&path).expect("detect format");
    assert_eq!(fmt, WorkbookFormat::Csv);
}

#[test]
fn detect_workbook_format_does_not_misclassify_single_line_prose_as_csv() {
    let dir = tempfile::tempdir().expect("temp dir");
    let path = dir.path().join("note.txt");
    std::fs::write(&path, "Hello, world").expect("write text");

    let fmt = detect_workbook_format(&path).expect("detect format");
    assert_eq!(fmt, WorkbookFormat::Unknown);
}

#[test]
fn detect_workbook_format_sniffs_utf16le_tab_delimited_text() {
    let dir = tempfile::tempdir().expect("temp dir");
    let path = dir.path().join("data.xlsx");

    let tsv = "col1\tcol2\r\n1\thello\r\n2\tworld\r\n";
    let mut bytes = vec![0xFF, 0xFE];
    for unit in tsv.encode_utf16() {
        bytes.extend_from_slice(&unit.to_le_bytes());
    }
    std::fs::write(&path, &bytes).expect("write utf16 tsv");

    let fmt = detect_workbook_format(&path).expect("detect format");
    assert_eq!(fmt, WorkbookFormat::Csv);
}

#[test]
fn detect_workbook_format_sniffs_utf16be_tab_delimited_text() {
    let dir = tempfile::tempdir().expect("temp dir");
    let path = dir.path().join("data.xlsx");

    let tsv = "col1\tcol2\r\n1\thello\r\n2\tworld\r\n";
    let mut bytes = vec![0xFE, 0xFF];
    for unit in tsv.encode_utf16() {
        bytes.extend_from_slice(&unit.to_be_bytes());
    }
    std::fs::write(&path, &bytes).expect("write utf16be tsv");

    let fmt = detect_workbook_format(&path).expect("detect format");
    assert_eq!(fmt, WorkbookFormat::Csv);
}

#[test]
fn detect_workbook_format_sniffs_utf16le_tab_delimited_text_without_bom() {
    let dir = tempfile::tempdir().expect("temp dir");
    let path = dir.path().join("data.xlsx");

    let tsv = "col1\tcol2\r\n1\thello\r\n2\tworld\r\n";
    let mut bytes = Vec::new();
    for unit in tsv.encode_utf16() {
        bytes.extend_from_slice(&unit.to_le_bytes());
    }
    std::fs::write(&path, &bytes).expect("write utf16 tsv");

    let fmt = detect_workbook_format(&path).expect("detect format");
    assert_eq!(fmt, WorkbookFormat::Csv);
}

#[test]
fn detect_workbook_format_sniffs_utf16be_tab_delimited_text_without_bom() {
    let dir = tempfile::tempdir().expect("temp dir");
    let path = dir.path().join("data.xlsx");

    let tsv = "col1\tcol2\r\n1\thello\r\n2\tworld\r\n";
    let mut bytes = Vec::new();
    for unit in tsv.encode_utf16() {
        bytes.extend_from_slice(&unit.to_be_bytes());
    }
    std::fs::write(&path, &bytes).expect("write utf16be tsv");

    let fmt = detect_workbook_format(&path).expect("detect format");
    assert_eq!(fmt, WorkbookFormat::Csv);
}

#[test]
fn detect_workbook_format_sniffs_utf16le_tab_delimited_text_without_bom_mostly_non_ascii() {
    let dir = tempfile::tempdir().expect("temp dir");
    let path = dir.path().join("data.xlsx");

    let left = "あ".repeat(200);
    let right = "い".repeat(200);
    let tsv = format!("{left}\t{right}\r\n{left}\t{right}\r\n");
    let mut bytes = Vec::new();
    for unit in tsv.encode_utf16() {
        bytes.extend_from_slice(&unit.to_le_bytes());
    }
    std::fs::write(&path, &bytes).expect("write utf16 tsv");

    let fmt = detect_workbook_format(&path).expect("detect format");
    assert_eq!(fmt, WorkbookFormat::Csv);
}

#[test]
fn detect_workbook_format_sniffs_utf16be_tab_delimited_text_without_bom_mostly_non_ascii() {
    let dir = tempfile::tempdir().expect("temp dir");
    let path = dir.path().join("data.xlsx");

    let left = "あ".repeat(200);
    let right = "い".repeat(200);
    let tsv = format!("{left}\t{right}\r\n{left}\t{right}\r\n");
    let mut bytes = Vec::new();
    for unit in tsv.encode_utf16() {
        bytes.extend_from_slice(&unit.to_be_bytes());
    }
    std::fs::write(&path, &bytes).expect("write utf16be tsv");

    let fmt = detect_workbook_format(&path).expect("detect format");
    assert_eq!(fmt, WorkbookFormat::Csv);
}

#[test]
fn detect_workbook_format_does_not_misclassify_binary_as_csv() {
    let dir = tempfile::tempdir().expect("temp dir");
    let path = dir.path().join("blob");

    // Include NUL/control bytes so the text heuristic should reject it.
    std::fs::write(&path, b"\x00\x01\x02\x03not csv").expect("write binary");

    let fmt = detect_workbook_format(&path).expect("detect format");
    assert_eq!(fmt, WorkbookFormat::Unknown);
}
