use std::io::Cursor;
use std::path::{Path, PathBuf};

use formula_xlsx::{decrypt_ooxml_from_cfb, XlsxPackage};

const PASSWORD: &str = "password";

fn fixture_path_buf(rel: &str) -> PathBuf {
    Path::new(concat!(env!("CARGO_MANIFEST_DIR"), "/../../fixtures/encrypted/ooxml/")).join(rel)
}

fn decrypt_fixture(encrypted_name: &str) -> Vec<u8> {
    let path = fixture_path_buf(encrypted_name);
    let bytes =
        std::fs::read(&path).unwrap_or_else(|err| panic!("read fixture {path:?}: {err}"));

    let cursor = Cursor::new(bytes);
    let mut ole = cfb::CompoundFile::open(cursor).expect("open OLE container");
    decrypt_ooxml_from_cfb(&mut ole, PASSWORD).expect("decrypt encrypted package")
}

#[test]
fn decrypts_agile_and_standard_large_fixtures() {
    let plaintext_path = fixture_path_buf("plaintext-large.xlsx");
    let plaintext =
        std::fs::read(plaintext_path).expect("read plaintext-large.xlsx fixture bytes");

    // Sanity: ensure we actually exercise multi-segment (4096-byte) Agile decryption.
    assert!(
        plaintext.len() > 4096,
        "expected plaintext-large.xlsx to be > 4096 bytes, got {}",
        plaintext.len()
    );

    for encrypted in ["agile-large.xlsx", "standard-large.xlsx"] {
        let decrypted = decrypt_fixture(encrypted);
        assert_eq!(
            decrypted, plaintext,
            "decrypted bytes must match plaintext-large.xlsx for {encrypted}"
        );

        // Additional sanity: the decrypted bytes should be a valid OPC/ZIP workbook package.
        let pkg = XlsxPackage::from_bytes(&decrypted).expect("open decrypted package as XLSX");
        assert!(
            pkg.part_names().any(|n| n.eq_ignore_ascii_case("xl/workbook.xml")),
            "decrypted package missing xl/workbook.xml"
        );
    }
}
