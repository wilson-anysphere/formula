use std::io::{Cursor, Read, Seek};
use std::path::PathBuf;

use formula_xlsx::offcrypto::decrypt_ooxml_encrypted_package;

const PASSWORD: &str = "password";

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../formula-io/tests/fixtures/offcrypto_standard_cryptoapi_password.xlsx")
}

fn read_stream<R: Read + Seek + std::io::Write>(
    ole: &mut cfb::CompoundFile<R>,
    name: &str,
) -> Vec<u8> {
    // Some producers include a leading slash in stream names. Be tolerant.
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

#[test]
fn decrypts_real_standard_cryptoapi_fixture() {
    let path = fixture_path();
    let bytes = std::fs::read(&path).unwrap_or_else(|err| panic!("read fixture {path:?}: {err}"));

    let cursor = Cursor::new(bytes);
    let mut ole = cfb::CompoundFile::open(cursor).expect("open OLE container");

    let encryption_info = read_stream(&mut ole, "EncryptionInfo");
    let encrypted_package = read_stream(&mut ole, "EncryptedPackage");

    let decrypted = decrypt_ooxml_encrypted_package(&encryption_info, &encrypted_package, PASSWORD)
        .expect("decrypt EncryptedPackage");

    assert!(
        decrypted.starts_with(b"PK"),
        "expected decrypted OOXML package to start with ZIP signature (PK)"
    );
}
