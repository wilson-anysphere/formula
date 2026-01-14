#![cfg(feature = "encrypted-workbooks")]

use std::path::PathBuf;

use formula_io::{open_workbook_model_with_password, open_workbook_with_password, Error, Workbook};
use formula_model::CellValue;

fn fixture_path(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/encrypted")
        .join(name)
}

fn fixture_password(name: &str) -> &'static str {
    match name {
        "encrypted_agile.xlsx" => "Password1234_",
        // Standard/CryptoAPI encrypted XLSX (Apache POI output).
        "encrypted_standard.xlsx" => "password",
        // Sourced from Apache POI test corpus (`protected_passtika.xlsb`).
        "encrypted.xlsb" => "tika",
        // Sourced from `crates/formula-xls/tests/fixtures/encrypted/biff8_rc4_cryptoapi_pw_open.xls`.
        "encrypted.xls" => "correct horse battery staple",
        _ => panic!("unknown fixture {name}"),
    }
}

#[test]
fn opens_encrypted_agile_xlsx_fixture() {
    let name = "encrypted_agile.xlsx";
    let path = fixture_path(name);
    let pw = fixture_password(name);

    let wb = open_workbook_with_password(&path, Some(pw)).expect("decrypt + open");
    assert!(
        matches!(wb, Workbook::Xlsx(_)),
        "expected Workbook::Xlsx, got {wb:?}"
    );

    let workbook = open_workbook_model_with_password(&path, Some(pw)).expect("decrypt + open");

    let sheet = workbook.sheet_by_name("Sheet1").expect("Sheet1 missing");
    assert_eq!(
        sheet.value_a1("A1").unwrap(),
        CellValue::String("lorem".to_string())
    );
    assert_eq!(
        sheet.value_a1("B1").unwrap(),
        CellValue::String("ipsum".to_string())
    );
}

#[test]
fn opens_encrypted_standard_xlsx_fixture() {
    let name = "encrypted_standard.xlsx";
    let path = fixture_path(name);
    let pw = fixture_password(name);

    let wb = open_workbook_with_password(&path, Some(pw)).expect("decrypt + open");
    assert!(
        matches!(wb, Workbook::Xlsx(_)),
        "expected Workbook::Xlsx, got {wb:?}"
    );

    let workbook = open_workbook_model_with_password(&path, Some(pw)).expect("decrypt + open");
    let sheet = workbook.sheet_by_name("Sheet1").expect("Sheet1 missing");
    assert_eq!(sheet.value_a1("A1").unwrap(), CellValue::Number(1.0));
    assert_eq!(
        sheet.value_a1("B1").unwrap(),
        CellValue::String("Hello".to_string())
    );
}

#[test]
fn opens_encrypted_xlsb_fixture() {
    let name = "encrypted.xlsb";
    let path = fixture_path(name);
    let pw = fixture_password(name);

    let wb = open_workbook_with_password(&path, Some(pw)).expect("decrypt + open");
    assert!(
        matches!(wb, Workbook::Xlsb(_)),
        "expected Workbook::Xlsb, got {wb:?}"
    );

    let workbook = open_workbook_model_with_password(&path, Some(pw)).expect("decrypt + open");

    let sheet = workbook.sheet_by_name("Sheet1").expect("Sheet1 missing");
    assert!(
        matches!(sheet.value_a1("A1").unwrap(), CellValue::String(_)),
        "expected string in A1, got {:?}",
        sheet.value_a1("A1").unwrap()
    );

    assert_eq!(
        sheet.value_a1("A1").unwrap(),
        CellValue::String("You can't see me".to_string())
    );
    assert_eq!(sheet.value_a1("B1").unwrap(), CellValue::Empty);
}

#[test]
fn opens_encrypted_legacy_xls_fixture() {
    let name = "encrypted.xls";
    let path = fixture_path(name);
    let pw = fixture_password(name);

    let wb = open_workbook_with_password(&path, Some(pw)).expect("decrypt + open");
    assert!(
        matches!(wb, Workbook::Xls(_)),
        "expected Workbook::Xls, got {wb:?}"
    );

    let workbook = open_workbook_model_with_password(&path, Some(pw)).expect("decrypt + open");

    let sheet = workbook.sheet_by_name("Sheet1").expect("Sheet1 missing");
    assert_eq!(sheet.value_a1("A1").unwrap(), CellValue::Number(42.0));
}

#[test]
fn errors_with_wrong_password_for_all_fixtures() {
    for name in [
        "encrypted_agile.xlsx",
        "encrypted_standard.xlsx",
        "encrypted.xlsb",
        "encrypted.xls",
    ] {
        let path = fixture_path(name);
        let err = open_workbook_model_with_password(&path, Some("wrong-password"))
            .expect_err("expected invalid password error");
        assert!(
            matches!(err, Error::InvalidPassword { .. }),
            "expected Error::InvalidPassword for {name}, got {err:?}"
        );
    }
}
