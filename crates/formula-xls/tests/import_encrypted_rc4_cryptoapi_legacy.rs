use std::path::PathBuf;

use formula_model::CellValue;

const UNICODE_EMOJI_PASSWORD: &str = "pässwörd🔒";

fn legacy_fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("encrypted")
        .join("biff8_rc4_cryptoapi_legacy_unicode_emoji_pw_open.xls")
}

#[test]
fn decrypts_rc4_cryptoapi_legacy_biff8_xls_with_unicode_emoji_password() {
    let result = formula_xls::import_xls_path_with_password(
        legacy_fixture_path(),
        Some(UNICODE_EMOJI_PASSWORD),
    )
    .expect("expected decrypt + import to succeed");
    let sheet = result.workbook.sheet_by_name("Sheet1").expect("Sheet1");
    assert_eq!(sheet.value_a1("A1").unwrap(), CellValue::Number(42.0));
}

#[test]
fn rc4_cryptoapi_legacy_unicode_emoji_wrong_password_errors() {
    let err = formula_xls::import_xls_path_with_password(legacy_fixture_path(), Some("wrong password"))
        .expect_err("expected wrong password error");
    assert!(matches!(err, formula_xls::ImportError::InvalidPassword));
}

#[test]
fn rc4_cryptoapi_legacy_unicode_emoji_password_different_normalization_fails() {
    // NFC password is "pässwörd🔒" (U+00E4, U+00F6). NFD decomposes those into combining marks.
    let nfd = "pa\u{0308}sswo\u{0308}rd🔒";
    assert_ne!(
        nfd, UNICODE_EMOJI_PASSWORD,
        "strings should differ before UTF-16 encoding"
    );

    let err = formula_xls::import_xls_path_with_password(legacy_fixture_path(), Some(nfd))
        .expect_err("expected wrong password error");
    assert!(matches!(err, formula_xls::ImportError::InvalidPassword));
}

