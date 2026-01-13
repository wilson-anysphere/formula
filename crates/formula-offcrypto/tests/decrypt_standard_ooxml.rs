// Fixtures in `tests/fixtures/` are copied from the MIT-licensed `nolze/msoffcrypto-tool` repo:
// https://github.com/nolze/msoffcrypto-tool
//
// The upstream project is MIT licensed; see their repository for the full license text.

use std::path::PathBuf;

use formula_offcrypto::decrypt_standard_ooxml_from_bytes;
use formula_offcrypto::{EncryptionType, OffcryptoError};

fn fixture(path: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("tests").join("fixtures").join(path)
}

#[test]
fn decrypts_standard_fixture_docx() {
    let encrypted = std::fs::read(fixture("inputs/ecma376standard_password.docx"))
        .expect("read encrypted fixture");
    let expected = std::fs::read(fixture("outputs/ecma376standard_password_plain.docx"))
        .expect("read expected decrypted fixture");

    let decrypted =
        decrypt_standard_ooxml_from_bytes(encrypted, "Password1234_").expect("decrypt fixture");
    assert!(decrypted.starts_with(b"PK"));
    assert_eq!(decrypted, expected);
}

#[test]
fn wrong_password_returns_error() {
    let encrypted = std::fs::read(fixture("inputs/ecma376standard_password.docx"))
        .expect("read encrypted fixture");

    let err = decrypt_standard_ooxml_from_bytes(encrypted, "not-the-password")
        .expect_err("expected wrong password to error");
    assert!(matches!(err, formula_offcrypto::OffcryptoError::InvalidPassword));
}

#[test]
fn rejects_agile_fixture() {
    // `example_password.xlsx` is an Agile-encrypted OOXML package (EncryptionInfo v4.4).
    let encrypted =
        std::fs::read(fixture("inputs/example_password.xlsx")).expect("read encrypted fixture");

    let err = decrypt_standard_ooxml_from_bytes(encrypted, "any-password")
        .expect_err("expected Agile encryption to be rejected");
    assert!(
        matches!(
            err,
            OffcryptoError::UnsupportedEncryption {
                encryption_type: EncryptionType::Agile
            }
        ),
        "expected UnsupportedEncryption(Agile), got {err:?}"
    );
}
