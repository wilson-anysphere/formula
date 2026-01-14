use std::io::Cursor;
use std::path::PathBuf;

use formula_office_crypto::{decrypt_encrypted_package, is_encrypted_ooxml_ole};

fn fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures")
        .join(rel)
}

fn assert_valid_zip(bytes: &[u8]) {
    assert!(
        bytes.starts_with(b"PK"),
        "expected decrypted bytes to be a ZIP"
    );
    let mut zip = zip::ZipArchive::new(Cursor::new(bytes)).expect("open decrypted zip");
    zip.by_name("xl/workbook.xml")
        .expect("expected xl/workbook.xml in decrypted zip");
}

#[test]
fn decrypts_ooxml_encrypted_fixtures_to_valid_zip() {
    let fixtures = [
        ("encryption/encrypted_agile.xlsx", "password", false),
        ("encryption/encrypted_standard.xlsx", "password", false),
        ("encryption/encrypted_agile_unicode.xlsx", "pässwörd", false),
        ("encryption/encrypted_agile.xlsm", "password", true),
    ];

    for (rel, password, expect_vba) in fixtures {
        let bytes =
            std::fs::read(fixture_path(rel)).unwrap_or_else(|_| panic!("read {rel} fixture"));
        assert!(
            is_encrypted_ooxml_ole(&bytes),
            "expected {rel} to be detected as an encrypted OOXML package"
        );

        let decrypted =
            decrypt_encrypted_package(&bytes, password).unwrap_or_else(|_| panic!("decrypt {rel}"));
        assert_valid_zip(&decrypted);

        if expect_vba {
            let mut zip = zip::ZipArchive::new(Cursor::new(&decrypted)).expect("open zip");
            zip.by_name("xl/vbaProject.bin")
                .expect("expected xl/vbaProject.bin in decrypted xlsm");
        }
    }
}
