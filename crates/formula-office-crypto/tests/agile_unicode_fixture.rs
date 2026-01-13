use std::io::Read as _;

#[test]
fn decrypts_agile_unicode_fixture() {
    let path = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/encrypted/ooxml/agile-unicode.xlsx");
    let ole_bytes = std::fs::read(&path).expect("read agile-unicode.xlsx");

    let decrypted =
        formula_office_crypto::decrypt_encrypted_package_ole(&ole_bytes, "pässwörd")
            .expect("decrypt agile-unicode.xlsx with unicode password");
    assert!(
        decrypted.starts_with(b"PK"),
        "expected decrypted bytes to start with PK"
    );

    let mut reader =
        formula_office_crypto::decrypt_encrypted_package_ole_to_reader(&ole_bytes, "pässwörd")
            .expect("decrypt agile-unicode.xlsx to reader");
    let mut prefix = [0u8; 2];
    reader.read_exact(&mut prefix).expect("read prefix");
    assert_eq!(&prefix, b"PK");
}

