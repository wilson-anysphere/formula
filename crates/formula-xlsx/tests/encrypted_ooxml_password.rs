use std::path::{Path, PathBuf};

use formula_model::{CellRef, CellValue};
use formula_xlsx::{
    load_from_bytes_with_password, read_workbook_model_from_bytes_with_password, ReadError,
    XlsxError, XlsxPackage,
};

fn fixture_path(rel: &str) -> PathBuf {
    Path::new(concat!(env!("CARGO_MANIFEST_DIR"), "/../../fixtures/encrypted/ooxml/")).join(rel)
}

fn read_fixture(name: &str) -> Vec<u8> {
    let path = fixture_path(name);
    std::fs::read(&path).unwrap_or_else(|err| panic!("read fixture {path:?}: {err}"))
}

fn assert_plaintext_contents(workbook: &formula_model::Workbook) {
    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "Sheet1");

    let sheet = &workbook.sheets[0];
    assert_eq!(
        sheet.value(CellRef::from_a1("A1").unwrap()),
        CellValue::Number(1.0)
    );
    assert_eq!(
        sheet.value(CellRef::from_a1("B1").unwrap()),
        CellValue::String("Hello".to_string())
    );
}

#[test]
fn workbook_and_document_loaders_decrypt_agile_standard_and_empty_password_fixtures() {
    let cases = [
        ("plaintext.xlsx", "ignored"),
        ("agile.xlsx", "password"),
        ("standard.xlsx", "password"),
        ("agile-empty-password.xlsx", ""),
    ];

    for (name, password) in cases {
        let bytes = read_fixture(name);

        let model = read_workbook_model_from_bytes_with_password(&bytes, password)
            .unwrap_or_else(|err| panic!("read_workbook_model_from_bytes_with_password({name}): {err:?}"));
        assert_plaintext_contents(&model);

        let doc = load_from_bytes_with_password(&bytes, password)
            .unwrap_or_else(|err| panic!("load_from_bytes_with_password({name}): {err:?}"));
        assert_plaintext_contents(&doc.workbook);
    }
}

#[test]
fn xlsx_package_loader_decrypts_standard_fixture_and_exposes_password_errors() {
    let plaintext = read_fixture("plaintext.xlsx");
    let plaintext_pkg = XlsxPackage::from_bytes(&plaintext).expect("open plaintext package");

    for encrypted_name in ["agile.xlsx", "standard.xlsx"] {
        let encrypted = read_fixture(encrypted_name);
        let pkg = XlsxPackage::from_bytes_with_password(&encrypted, "password")
            .unwrap_or_else(|err| panic!("from_bytes_with_password({encrypted_name}): {err:?}"));

        assert_eq!(
            pkg.part("xl/workbook.xml"),
            plaintext_pkg.part("xl/workbook.xml"),
            "xl/workbook.xml should match plaintext for {encrypted_name}"
        );
        assert_eq!(
            pkg.part("xl/worksheets/sheet1.xml"),
            plaintext_pkg.part("xl/worksheets/sheet1.xml"),
            "xl/worksheets/sheet1.xml should match plaintext for {encrypted_name}"
        );
    }

    let standard = read_fixture("standard.xlsx");
    let err = load_from_bytes_with_password(&standard, "wrong").expect_err("expected failure");
    assert!(
        matches!(err, ReadError::InvalidPassword),
        "expected ReadError::InvalidPassword, got {err:?}"
    );

    let err = XlsxPackage::from_bytes_with_password(&standard, "wrong").expect_err("bad password");
    assert!(
        matches!(err, XlsxError::InvalidPassword),
        "expected XlsxError::InvalidPassword, got {err:?}"
    );
}
