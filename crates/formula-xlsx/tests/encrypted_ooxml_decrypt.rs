use std::io::{Cursor, Read, Seek};
use std::path::{Path, PathBuf};

use formula_xlsx::offcrypto::decrypt_ooxml_encrypted_package;
use formula_xlsx::XlsxPackage;

const PASSWORD: &str = "password";

fn fixture_path_buf(rel: &str) -> PathBuf {
    Path::new(concat!(env!("CARGO_MANIFEST_DIR"), "/../../fixtures/encrypted/ooxml/")).join(rel)
}

fn read_stream<R: Read + Seek + std::io::Write>(
    ole: &mut cfb::CompoundFile<R>,
    name: &str,
) -> Vec<u8> {
    // Some producers include a leading slash in the stream name. Be tolerant.
    let mut stream = ole
        .open_stream(name)
        .or_else(|_| ole.open_stream(&format!("/{name}")))
        .unwrap_or_else(|err| panic!("open {name} stream: {err}"));

    let mut buf = Vec::new();
    stream
        .read_to_end(&mut buf)
        .unwrap_or_else(|err| panic!("read {name} stream: {err}"));
    buf
}

fn decrypt_fixture(encrypted_name: &str) -> Vec<u8> {
    let path = fixture_path_buf(encrypted_name);
    let bytes =
        std::fs::read(&path).unwrap_or_else(|err| panic!("read fixture {path:?}: {err}"));

    let cursor = Cursor::new(bytes);
    let mut ole = cfb::CompoundFile::open(cursor).expect("open OLE container");

    let encryption_info = read_stream(&mut ole, "EncryptionInfo");
    let encrypted_package = read_stream(&mut ole, "EncryptedPackage");

    // Header version sanity check (fixtures should be deterministic).
    assert!(
        encryption_info.len() >= 4,
        "EncryptionInfo stream too short ({} bytes)",
        encryption_info.len()
    );
    let major = u16::from_le_bytes([encryption_info[0], encryption_info[1]]);
    let minor = u16::from_le_bytes([encryption_info[2], encryption_info[3]]);
    match encrypted_name {
        name if name.starts_with("agile") => assert_eq!((major, minor), (4, 4)),
        name if name.starts_with("standard") => assert_eq!((major, minor), (3, 2)),
        _ => {}
    }

    decrypt_ooxml_encrypted_package(&encryption_info, &encrypted_package, PASSWORD)
        .expect("decrypt encrypted package")
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
