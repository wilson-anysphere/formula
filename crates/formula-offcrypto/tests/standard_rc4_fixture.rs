use std::fs::File;
use std::io::Read;
use std::path::PathBuf;

use formula_offcrypto::{parse_encryption_info, standard_rc4, EncryptionInfo, StandardEncryptionInfo};

fn fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures")
        .join(rel)
}

#[test]
fn decrypts_standard_rc4_fixture_to_plaintext_zip() {
    let encrypted_path = fixture_path("encrypted/ooxml/standard-rc4.xlsx");
    let plaintext_path = fixture_path("encrypted/ooxml/plaintext.xlsx");

    let file = File::open(&encrypted_path).expect("open encrypted fixture");
    let mut ole = cfb::CompoundFile::open(file).expect("open OLE container");

    let mut info_bytes = Vec::new();
    ole.open_stream("EncryptionInfo")
        .expect("EncryptionInfo stream")
        .read_to_end(&mut info_bytes)
        .expect("read EncryptionInfo");

    let info = parse_encryption_info(&info_bytes).expect("parse EncryptionInfo");
    let (header, verifier) = match info {
        EncryptionInfo::Standard { header, verifier, .. } => (header, verifier),
        other => panic!("expected Standard EncryptionInfo, got {other:?}"),
    };
    let standard = StandardEncryptionInfo { header, verifier };

    let mut encrypted_package = Vec::new();
    ole.open_stream("EncryptedPackage")
        .expect("EncryptedPackage stream")
        .read_to_end(&mut encrypted_package)
        .expect("read EncryptedPackage");

    let decrypted =
        standard_rc4::decrypt_encrypted_package(&standard, &encrypted_package, "password")
            .expect("decrypt package");

    let plaintext = std::fs::read(&plaintext_path).expect("read plaintext fixture");
    assert_eq!(decrypted, plaintext);
}

