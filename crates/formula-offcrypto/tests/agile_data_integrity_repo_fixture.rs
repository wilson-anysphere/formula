use std::io::{Cursor, Read};
use std::path::PathBuf;

use formula_offcrypto::{decrypt_encrypted_package, DecryptOptions};

fn fixture(path: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("..")
        .join("..")
        .join("fixtures")
        .join("encrypted")
        .join("ooxml")
        .join(path)
}

fn read_ole_stream(bytes: &[u8], name: &str) -> Vec<u8> {
    let mut ole = cfb::CompoundFile::open(Cursor::new(bytes)).expect("open fixture cfb");
    let mut stream = ole.open_stream(name).expect("open stream");
    let mut out = Vec::new();
    stream.read_to_end(&mut out).expect("read stream");
    out
}

#[test]
fn decrypt_encrypted_package_verify_integrity_accepts_plaintext_hmac_variant() {
    let encrypted =
        std::fs::read(fixture("agile-basic.xlsm")).expect("read encrypted fixture");
    let expected =
        std::fs::read(fixture("plaintext-basic.xlsm")).expect("read expected decrypted bytes");

    let encryption_info = read_ole_stream(&encrypted, "EncryptionInfo");
    let encrypted_package = read_ole_stream(&encrypted, "EncryptedPackage");

    let decrypted = decrypt_encrypted_package(
        &encryption_info,
        &encrypted_package,
        "password",
        DecryptOptions {
            verify_integrity: true,
            ..Default::default()
        },
    )
    .expect("decrypt with integrity verification");

    assert_eq!(decrypted, expected);
}
