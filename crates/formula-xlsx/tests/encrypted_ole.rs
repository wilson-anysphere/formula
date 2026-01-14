use std::io::Cursor;
use std::path::Path;

use formula_xlsx::XlsxError;

fn fixture_path(rel: &str) -> std::path::PathBuf {
    Path::new(concat!(env!("CARGO_MANIFEST_DIR"), "/../../fixtures/encrypted/ooxml/")).join(rel)
}

#[test]
fn encrypted_ole_correct_password_opens_workbook() {
    let bytes = std::fs::read(fixture_path("agile.xlsx")).expect("read encrypted fixture");

    let pkg = formula_xlsx::load_from_encrypted_ole_bytes(&bytes, "password")
        .expect("decrypt + open as xlsx package");
    assert!(
        pkg.part("xl/workbook.xml").is_some(),
        "expected decrypted package to contain xl/workbook.xml"
    );

    let workbook =
        formula_xlsx::read_workbook_from_encrypted_reader(Cursor::new(bytes), "password")
            .expect("decrypt + parse workbook");

    assert!(
        workbook.sheets.iter().any(|sheet| sheet.name == "Sheet1"),
        "expected decrypted workbook to contain Sheet1"
    );
}

#[test]
fn encrypted_ole_wrong_password_returns_invalid_password() {
    let bytes = std::fs::read(fixture_path("agile.xlsx")).expect("read encrypted fixture");

    let err = formula_xlsx::read_workbook_from_encrypted_reader(Cursor::new(bytes), "wrong")
        .expect_err("expected invalid password to error");
    assert!(
        matches!(err, XlsxError::InvalidPassword),
        "expected InvalidPassword, got {err:?}"
    );
}
