use formula_offcrypto::{
    decrypt_encrypted_package, decrypt_ooxml_from_ole_bytes, DecryptOptions, OffcryptoError,
};

mod support;

#[test]
fn decrypt_agile_roundtrip_with_derived_password_key_iv() {
    let password = "Password";
    let plaintext = b"PK\0\0formula-offcrypto-derived-iv-test".to_vec();

    let (encryption_info, encrypted_package) =
        support::encrypt_agile_password_key_derived_iv(&plaintext, password);

    let decrypted = decrypt_encrypted_package(
        &encryption_info,
        &encrypted_package,
        password,
        DecryptOptions::default(),
    )
    .expect("decrypt derived-IV Agile package");

    assert_eq!(decrypted, plaintext);
}

#[test]
fn decrypt_agile_wrong_password_with_derived_password_key_iv_is_invalid_password() {
    let plaintext = b"PK\0\0formula-offcrypto-derived-iv-test".to_vec();
    let (encryption_info, encrypted_package) =
        support::encrypt_agile_password_key_derived_iv(&plaintext, "password-1");

    let err = decrypt_encrypted_package(
        &encryption_info,
        &encrypted_package,
        "password-2",
        DecryptOptions::default(),
    )
    .expect_err("wrong password should fail");

    assert_eq!(err, OffcryptoError::InvalidPassword);
}

#[test]
fn decrypt_agile_ole_roundtrip_with_derived_password_key_iv() {
    let password = "Password";
    let plaintext = b"PK\0\0formula-offcrypto-derived-iv-test".to_vec();

    let (encryption_info, encrypted_package) =
        support::encrypt_agile_password_key_derived_iv(&plaintext, password);
    let ole_bytes = support::wrap_in_ole_cfb(&encryption_info, &encrypted_package);

    let decrypted =
        decrypt_ooxml_from_ole_bytes(ole_bytes, password).expect("decrypt derived-IV Agile OLE");
    assert_eq!(decrypted, plaintext);
}

#[test]
fn decrypt_agile_ole_wrong_password_with_derived_password_key_iv_is_invalid_password() {
    let plaintext = b"PK\0\0formula-offcrypto-derived-iv-test".to_vec();
    let (encryption_info, encrypted_package) =
        support::encrypt_agile_password_key_derived_iv(&plaintext, "password-1");
    let ole_bytes = support::wrap_in_ole_cfb(&encryption_info, &encrypted_package);

    let err = decrypt_ooxml_from_ole_bytes(ole_bytes, "password-2")
        .expect_err("wrong password should fail");
    assert_eq!(err, OffcryptoError::InvalidPassword);
}
