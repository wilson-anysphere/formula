use std::io::Cursor;
use std::path::{Path, PathBuf};

use formula_xlsx::{decrypt_ooxml_from_cfb, XlsxPackage};

const PASSWORD: &str = "password";

fn fixture_path_buf(rel: &str) -> PathBuf {
    Path::new(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/encrypted/ooxml/"
    ))
    .join(rel)
}

fn decrypt_fixture(encrypted_name: &str) -> Vec<u8> {
    let path = fixture_path_buf(encrypted_name);
    let bytes = std::fs::read(&path).unwrap_or_else(|err| panic!("read fixture {path:?}: {err}"));

    let cursor = Cursor::new(bytes);
    let mut ole = cfb::CompoundFile::open(cursor).expect("open OLE container");
    decrypt_ooxml_from_cfb(&mut ole, PASSWORD)
        .unwrap_or_else(|err| panic!("decrypt {encrypted_name} encrypted package: {err}"))
}

#[test]
fn decrypts_agile_and_standard_large_fixtures() {
    let plaintext_path = fixture_path_buf("plaintext-large.xlsx");
    let plaintext = std::fs::read(plaintext_path).expect("read plaintext-large.xlsx fixture bytes");

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
            pkg.part_names()
                .any(|n| n.eq_ignore_ascii_case("xl/workbook.xml")),
            "decrypted package missing xl/workbook.xml"
        );
    }
}

#[test]
fn decrypts_agile_and_standard_small_fixtures() {
    let plaintext_path = fixture_path_buf("plaintext.xlsx");
    let plaintext = std::fs::read(plaintext_path).expect("read plaintext.xlsx fixture bytes");

    // Sanity: ensure we actually exercise the <=4096-byte edge case (padding/truncation).
    assert!(
        plaintext.len() < 4096,
        "expected plaintext.xlsx to be < 4096 bytes, got {}",
        plaintext.len()
    );

    for encrypted in ["agile.xlsx", "standard.xlsx", "standard-rc4.xlsx"] {
        let decrypted = decrypt_fixture(encrypted);
        assert_eq!(
            decrypted, plaintext,
            "decrypted bytes must match plaintext.xlsx for {encrypted}"
        );

        // Additional sanity: the decrypted bytes should be a valid OPC/ZIP workbook package.
        let pkg = XlsxPackage::from_bytes(&decrypted).expect("open decrypted package as XLSX");
        assert!(
            pkg.part_names()
                .any(|n| n.eq_ignore_ascii_case("xl/workbook.xml")),
            "decrypted package missing xl/workbook.xml"
        );
    }
}

#[test]
fn decrypts_standard_4_2_fixture() {
    let plaintext_path = fixture_path_buf("plaintext.xlsx");
    let plaintext = std::fs::read(plaintext_path).expect("read plaintext.xlsx fixture bytes");

    let decrypted = decrypt_fixture("standard-4.2.xlsx");
    assert_eq!(
        decrypted, plaintext,
        "decrypted bytes must match plaintext.xlsx for standard-4.2.xlsx"
    );

    let pkg = XlsxPackage::from_bytes(&decrypted).expect("open decrypted package as XLSX");
    assert!(
        pkg.part_names()
            .any(|n| n.eq_ignore_ascii_case("xl/workbook.xml")),
        "decrypted package missing xl/workbook.xml"
    );
}

#[test]
fn decrypts_standard_small_fixtures_match_plaintext_bytes() {
    let plaintext_xlsx = std::fs::read(fixture_path_buf("plaintext.xlsx"))
        .expect("read plaintext.xlsx fixture bytes");
    let plaintext_xlsm = std::fs::read(fixture_path_buf("plaintext-basic.xlsm"))
        .expect("read plaintext-basic.xlsm fixture bytes");

    let decrypted_xlsx = decrypt_fixture("standard.xlsx");
    assert_eq!(
        decrypted_xlsx, plaintext_xlsx,
        "decrypted bytes must match plaintext.xlsx for standard.xlsx"
    );
    assert!(decrypted_xlsx.starts_with(b"PK"));

    let decrypted_rc4 = decrypt_fixture("standard-rc4.xlsx");
    assert_eq!(
        decrypted_rc4, plaintext_xlsx,
        "decrypted bytes must match plaintext.xlsx for standard-rc4.xlsx"
    );
    assert!(decrypted_rc4.starts_with(b"PK"));

    let decrypted_xlsm = decrypt_fixture("standard-basic.xlsm");
    assert_eq!(
        decrypted_xlsm, plaintext_xlsm,
        "decrypted bytes must match plaintext-basic.xlsm for standard-basic.xlsm"
    );
    assert!(decrypted_xlsm.starts_with(b"PK"));
}

#[test]
fn decrypts_agile_basic_xlsm_fixture_matches_plaintext_bytes() {
    let plaintext_xlsm = std::fs::read(fixture_path_buf("plaintext-basic.xlsm"))
        .expect("read plaintext-basic.xlsm fixture bytes");
    let decrypted_xlsm = decrypt_fixture("agile-basic.xlsm");
    assert_eq!(
        decrypted_xlsm, plaintext_xlsm,
        "decrypted bytes must match plaintext-basic.xlsm for agile-basic.xlsm"
    );
    assert!(decrypted_xlsm.starts_with(b"PK"));
}

#[test]
fn xlsxpackage_from_bytes_with_password_supports_agile_and_standard() {
    for encrypted in [
        "agile.xlsx",
        "standard.xlsx",
        "standard-rc4.xlsx",
        "agile-large.xlsx",
        "standard-large.xlsx",
    ] {
        let path = fixture_path_buf(encrypted);
        let bytes =
            std::fs::read(&path).unwrap_or_else(|err| panic!("read fixture {path:?}: {err}"));

        let pkg = XlsxPackage::from_bytes_with_password(&bytes, PASSWORD)
            .unwrap_or_else(|err| panic!("from_bytes_with_password {encrypted}: {err}"));
        assert!(
            pkg.part_names()
                .any(|n| n.eq_ignore_ascii_case("xl/workbook.xml")),
            "{encrypted}: decrypted package missing xl/workbook.xml"
        );
    }
}

#[test]
fn xlsxpackage_from_bytes_with_password_decrypts_agile_and_standard_xlsm() {
    for encrypted in ["agile-basic.xlsm", "standard-basic.xlsm"] {
        let path = fixture_path_buf(encrypted);
        let bytes =
            std::fs::read(&path).unwrap_or_else(|err| panic!("read fixture {path:?}: {err}"));

        let pkg = XlsxPackage::from_bytes_with_password(&bytes, PASSWORD)
            .unwrap_or_else(|err| panic!("from_bytes_with_password {encrypted}: {err}"));
        assert!(
            pkg.part_names()
                .any(|n| n.eq_ignore_ascii_case("xl/workbook.xml")),
            "{encrypted}: decrypted package missing xl/workbook.xml"
        );
        let vba = pkg
            .vba_project_bin()
            .expect("expected decrypted xlsm to contain xl/vbaProject.bin");
        assert!(!vba.is_empty(), "expected vbaProject.bin to be non-empty");
    }
}
