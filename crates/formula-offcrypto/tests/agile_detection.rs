//! Agile-encrypted XLSX negative fixture.
//!
//! Fixture provenance:
//! - Source repository: https://github.com/nolze/msoffcrypto-tool (MIT)
//! - Files copied from:
//!   - `tests/inputs/example_password.xlsx` (password `Password1234_`)
//!   - `tests/outputs/example.xlsx` (expected plaintext)
//!
//! We only use this fixture for **detection** and **negative tests**. This crate's
//! "standard-only" entrypoint must reject Agile encryption with
//! `UnsupportedEncryption { encryption_type: EncryptionType::Agile }`, not misreport it as
//! `InvalidPassword`.

use std::io::{Cursor, Read};
use std::path::PathBuf;

use formula_offcrypto::{
    decrypt_standard_ooxml_from_bytes, parse_encryption_info, EncryptionInfo, EncryptionType,
    OffcryptoError,
};

fn fixture(path: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join(path)
}

#[test]
fn detects_agile_and_rejects_in_standard_only_mode() {
    let encrypted = std::fs::read(fixture("inputs/example_password.xlsx"))
        .expect("read agile encrypted fixture");

    let mut ole =
        cfb::CompoundFile::open(Cursor::new(encrypted.as_slice())).expect("open cfb");
    let mut encryption_info = Vec::new();
    ole.open_stream("EncryptionInfo")
        .expect("open EncryptionInfo stream")
        .read_to_end(&mut encryption_info)
        .expect("read EncryptionInfo stream");

    let parsed = parse_encryption_info(&encryption_info).expect("parse EncryptionInfo");
    let EncryptionInfo::Agile { version, .. } = parsed else {
        panic!("expected Agile EncryptionInfo, got {parsed:?}");
    };
    assert_eq!(version.major, 4);
    assert_eq!(version.minor, 4);

    let err = decrypt_standard_ooxml_from_bytes(encrypted, "Password1234_")
        .expect_err("expected standard-only decrypt to reject Agile encryption");
    assert!(
        matches!(
            &err,
            OffcryptoError::UnsupportedEncryption {
                encryption_type: EncryptionType::Agile
            }
        ),
        "expected UnsupportedEncryption(Agile), got {err:?}"
    );
}
