use std::path::Path;

use formula_offcrypto::decrypt_standard_ooxml_from_bytes;

fn fixture_path(rel: &str) -> std::path::PathBuf {
    Path::new(concat!(env!("CARGO_MANIFEST_DIR"), "/../../fixtures/encrypted/ooxml/")).join(rel)
}

#[test]
fn decrypts_repo_standard_xlsx_fixture() {
    let encrypted =
        std::fs::read(fixture_path("standard.xlsx")).expect("read standard.xlsx fixture");
    let expected =
        std::fs::read(fixture_path("plaintext.xlsx")).expect("read plaintext.xlsx fixture");

    let decrypted =
        decrypt_standard_ooxml_from_bytes(encrypted, "password").expect("decrypt standard.xlsx");
    assert!(decrypted.starts_with(b"PK"));
    assert_eq!(decrypted, expected);
}
