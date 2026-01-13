//! End-to-end decryption tests for Office-encrypted OOXML workbooks (Agile + Standard).
//!
//! These are gated behind the `encrypted-workbooks` feature because decryption support is still
//! landing. The companion fixture validation test (`encrypted_ooxml_fixture_validation.rs`) is
//! always enabled and provides immediate sanity checks once fixtures are present.
#![cfg(feature = "encrypted-workbooks")]

use std::io::Write as _;
use std::path::{Path, PathBuf};

use formula_io::{open_workbook_model_with_password, Error};
use formula_model::{CellRef, CellValue};

fn fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/encrypted/ooxml")
        .join(rel)
}

fn assert_expected_contents(workbook: &formula_model::Workbook) {
    assert_eq!(workbook.sheets.len(), 1, "expected exactly one sheet");
    assert_eq!(workbook.sheets[0].name, "Sheet1");

    let sheet = workbook.sheet_by_name("Sheet1").expect("Sheet1 missing");
    assert_eq!(
        sheet.value(CellRef::from_a1("A1").unwrap()),
        CellValue::Number(1.0)
    );
    assert_eq!(
        sheet.value(CellRef::from_a1("B1").unwrap()),
        CellValue::String("Hello".to_string())
    );
}

fn open_model_with_password(path: &Path, password: &str) -> formula_model::Workbook {
    open_workbook_model_with_password(path, Some(password))
        .unwrap_or_else(|err| panic!("open encrypted workbook {path:?} failed: {err:?}"))
}

#[test]
fn decrypts_agile_and_standard_with_correct_password() {
    let plaintext_path = fixture_path("plaintext.xlsx");
    let agile_path = fixture_path("agile.xlsx");
    let standard_path = fixture_path("standard.xlsx");

    // Baseline sanity: plaintext fixture matches expected contents.
    let plaintext = formula_io::open_workbook_model(&plaintext_path).expect("open plaintext.xlsx");
    assert_expected_contents(&plaintext);

    let agile = open_model_with_password(&agile_path, "password");
    assert_expected_contents(&agile);

    let standard = open_model_with_password(&standard_path, "password");
    assert_expected_contents(&standard);

    // Nice-to-have: content sniffing should work regardless of extension.
    let tmp = tempfile::tempdir().expect("temp dir");
    for (src, name) in [
        (&agile_path, "agile.xls"),
        (&agile_path, "agile.xlsb"),
        (&standard_path, "standard.xls"),
        (&standard_path, "standard.xlsb"),
    ] {
        let bytes = std::fs::read(src).expect("read encrypted fixture");

        let mut file = tempfile::Builder::new()
            .prefix(name)
            .tempfile_in(tmp.path())
            .expect("tempfile");
        file.write_all(&bytes).expect("write tempfile bytes");

        let wb = open_model_with_password(file.path(), "password");
        assert_expected_contents(&wb);
    }
}

#[test]
fn errors_on_wrong_password() {
    let agile_path = fixture_path("agile.xlsx");
    let standard_path = fixture_path("standard.xlsx");

    for path in [&agile_path, &standard_path] {
        assert!(
            matches!(
                open_workbook_model_with_password(path, Some("wrong-password")),
                Err(Error::InvalidPassword { .. })
            ),
            "expected InvalidPassword error for {path:?}"
        );
    }
}
