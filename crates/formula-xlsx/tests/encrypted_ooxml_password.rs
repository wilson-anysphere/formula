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

fn assert_excel_plaintext_contents(workbook: &formula_model::Workbook) {
    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "Sheet1");

    let sheet = &workbook.sheets[0];
    assert_eq!(
        sheet.value(CellRef::from_a1("A1").unwrap()),
        CellValue::String("lorem".to_string())
    );
    assert_eq!(
        sheet.value(CellRef::from_a1("B1").unwrap()),
        CellValue::String("ipsum".to_string())
    );
}

#[test]
fn workbook_and_document_loaders_decrypt_agile_standard_and_empty_password_fixtures() {
    let cases = [
        ("plaintext.xlsx", "ignored"),
        ("agile.xlsx", "password"),
        ("agile-unicode.xlsx", "pässwörd"),
        ("standard.xlsx", "password"),
        ("standard-rc4.xlsx", "password"),
        ("standard-unicode.xlsx", "pässwörd🔒"),
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
fn workbook_and_document_loaders_decrypt_agile_unicode_excel_fixture() {
    let plaintext_name = "plaintext-excel.xlsx";
    let encrypted_name = "agile-unicode-excel.xlsx";
    let password = "pässwörd🔒";

    let plaintext_bytes = read_fixture(plaintext_name);
    let plaintext_model = read_workbook_model_from_bytes_with_password(&plaintext_bytes, "ignored")
        .unwrap_or_else(|err| {
            panic!("read_workbook_model_from_bytes_with_password({plaintext_name}): {err:?}")
        });
    assert_excel_plaintext_contents(&plaintext_model);

    let encrypted_bytes = read_fixture(encrypted_name);
    let model = read_workbook_model_from_bytes_with_password(&encrypted_bytes, password)
        .unwrap_or_else(|err| {
            panic!("read_workbook_model_from_bytes_with_password({encrypted_name}): {err:?}")
        });
    assert_excel_plaintext_contents(&model);

    let doc = load_from_bytes_with_password(&encrypted_bytes, password)
        .unwrap_or_else(|err| panic!("load_from_bytes_with_password({encrypted_name}): {err:?}"));
    assert_excel_plaintext_contents(&doc.workbook);
}

#[test]
fn xlsx_package_loader_decrypts_standard_fixture_and_exposes_password_errors() {
    let plaintext = read_fixture("plaintext.xlsx");
    let plaintext_pkg = XlsxPackage::from_bytes(&plaintext).expect("open plaintext package");

    for (encrypted_name, password) in [
        ("agile.xlsx", "password"),
        ("standard.xlsx", "password"),
        ("standard-rc4.xlsx", "password"),
        ("standard-unicode.xlsx", "pässwörd🔒"),
    ] {
        let encrypted = read_fixture(encrypted_name);
        let pkg = XlsxPackage::from_bytes_with_password(&encrypted, password)
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

    let standard_rc4 = read_fixture("standard-rc4.xlsx");
    let err = load_from_bytes_with_password(&standard_rc4, "wrong").expect_err("expected failure");
    assert!(
        matches!(err, ReadError::InvalidPassword),
        "expected ReadError::InvalidPassword, got {err:?}"
    );

    let err = XlsxPackage::from_bytes_with_password(&standard_rc4, "wrong").expect_err("bad password");
    assert!(
        matches!(err, XlsxError::InvalidPassword),
        "expected XlsxError::InvalidPassword, got {err:?}"
    );
}

#[test]
fn unicode_password_normalization_mismatch_fails() {
    // NFC password is "pässwörd" (U+00E4, U+00F6). NFD decomposes those into combining marks.
    let nfd = "pa\u{0308}sswo\u{0308}rd";
    assert_ne!(nfd, "pässwörd");

    let bytes = read_fixture("agile-unicode.xlsx");
    let err = read_workbook_model_from_bytes_with_password(&bytes, nfd).expect_err("expected error");
    assert!(
        matches!(err, ReadError::InvalidPassword),
        "expected ReadError::InvalidPassword, got {err:?}"
    );
}

#[test]
fn unicode_emoji_password_normalization_mismatch_fails() {
    // NFC password is "pässwörd🔒" (U+00E4, U+00F6). NFD decomposes those into combining marks, but
    // leaves the non-BMP emoji alone.
    let nfd = "pa\u{0308}sswo\u{0308}rd🔒";
    assert_ne!(nfd, "pässwörd🔒");

    let bytes = read_fixture("agile-unicode-excel.xlsx");
    let err = read_workbook_model_from_bytes_with_password(&bytes, nfd).expect_err("expected error");
    assert!(
        matches!(err, ReadError::InvalidPassword),
        "expected ReadError::InvalidPassword, got {err:?}"
    );
}

#[test]
fn standard_unicode_emoji_password_normalization_mismatch_fails() {
    // NFC password is "pässwörd🔒" (U+00E4, U+00F6). NFD decomposes those into combining marks, but
    // leaves the non-BMP emoji alone.
    let nfd = "pa\u{0308}sswo\u{0308}rd🔒";
    assert_ne!(nfd, "pässwörd🔒");

    let bytes = read_fixture("standard-unicode.xlsx");
    let err =
        read_workbook_model_from_bytes_with_password(&bytes, nfd).expect_err("expected error");
    assert!(
        matches!(err, ReadError::InvalidPassword),
        "expected ReadError::InvalidPassword, got {err:?}"
    );
}
