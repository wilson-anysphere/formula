use std::path::PathBuf;

use formula_offcrypto::decrypt_standard_ooxml_from_bytes;

fn fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures")
        .join(rel)
}

#[test]
fn decrypts_standard_rc4_fixture_to_plaintext_zip() {
    let encrypted_path = fixture_path("encrypted/ooxml/standard-rc4.xlsx");
    let plaintext_path = fixture_path("encrypted/ooxml/plaintext.xlsx");

    let encrypted = std::fs::read(&encrypted_path).expect("read encrypted fixture");
    let decrypted =
        decrypt_standard_ooxml_from_bytes(encrypted, "password").expect("decrypt standard-rc4.xlsx");

    let plaintext = std::fs::read(&plaintext_path).expect("read plaintext fixture");
    assert_eq!(decrypted, plaintext);
}
