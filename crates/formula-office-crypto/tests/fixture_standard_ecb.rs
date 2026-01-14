use formula_office_crypto::decrypt_encrypted_package_ole;

#[test]
fn decrypts_repo_standard_fixture() {
    // This is the canonical "Standard" (CryptoAPI / ECMA-376) fixture in this repo.
    //
    // It uses AES-ECB (no IV) for both the password verifier fields and the `EncryptedPackage`
    // payload. This test ensures our Standard decryptor stays compatible with the fixture (and
    // with `msoffcrypto-tool`).
    let encrypted = include_bytes!("../../../fixtures/encrypted/ooxml/standard.xlsx");
    let expected = include_bytes!("../../../fixtures/encrypted/ooxml/plaintext.xlsx");

    let decrypted = decrypt_encrypted_package_ole(encrypted, "password").expect("decrypt fixture");
    assert_eq!(decrypted.as_slice(), expected.as_slice());
}

