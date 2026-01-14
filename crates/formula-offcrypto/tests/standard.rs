#![cfg(not(target_arch = "wasm32"))]

use formula_offcrypto::{decrypt_from_bytes, OffcryptoError};

// Fixtures vendored from https://github.com/nolze/msoffcrypto-tool (MIT license).
const ENCRYPTED_DOCX: &[u8] = include_bytes!("inputs/ecma376standard_password.docx");
const PLAIN_DOCX: &[u8] = include_bytes!("outputs/ecma376standard_password_plain.docx");

#[test]
fn decrypt_standard_docx_roundtrip() {
    let decrypted = decrypt_from_bytes(ENCRYPTED_DOCX, "Password1234_").expect("decrypt");
    assert_eq!(decrypted, PLAIN_DOCX);
    assert!(decrypted.starts_with(b"PK"));
}

#[test]
fn decrypt_standard_wrong_password_errors() {
    let err = decrypt_from_bytes(ENCRYPTED_DOCX, "wrong-password").expect_err("expected error");
    assert!(
        matches!(&err, OffcryptoError::InvalidPassword),
        "expected InvalidPassword, got {err:?}"
    );
}
