//! Regression test for the MS-OFFCRYPTO *Standard* (CryptoAPI) `EncryptedPackage` mode used by our
//! committed OOXML fixtures.
//!
//! There is long-running ambiguity across implementations/docs about whether Standard
//! `EncryptedPackage` uses AES-ECB or AES-CBC-with-segmentation. This test locks down the on-disk
//! mode for our committed Standard **AES** fixtures under `fixtures/encrypted/ooxml/standard*`.
//!
//! This test is a canary to ensure fixture regeneration does not silently drift to a different
//! mode.

use std::io::{Cursor, Read};
use std::path::PathBuf;

use sha2::Digest as _;

use formula_offcrypto::{
    decrypt_from_bytes, decrypt_standard_encrypted_package, parse_encryption_info,
    standard_derive_key_zeroizing, standard_verify_key, EncryptionInfo, StandardEncryptionInfo,
};

fn fixture(path: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("..")
        .join("..")
        .join("fixtures")
        .join("encrypted")
        .join("ooxml")
        .join(path)
}

fn read_ole_stream(raw_ole: &[u8], name: &str) -> Vec<u8> {
    let cursor = Cursor::new(raw_ole.to_vec());
    let mut ole = cfb::CompoundFile::open(cursor).expect("open OLE container");
    let mut stream = ole
        .open_stream(name)
        .or_else(|_| ole.open_stream(&format!("/{name}")))
        .unwrap_or_else(|err| panic!("open OLE stream {name:?}: {err}"));
    let mut buf = Vec::new();
    stream
        .read_to_end(&mut buf)
        .unwrap_or_else(|err| panic!("read OLE stream {name:?}: {err}"));
    buf
}

/// Attempt a CBC+segmented `EncryptedPackage` decrypt using the *same* Standard password verifier
/// implementation in this crate.
///
/// This is only used as a diagnostic if ECB decryption does not match the plaintext fixture.
fn assert_fixture_is_ecb(encrypted_name: &str, plaintext_name: &str, password: &str) {
    let encrypted =
        std::fs::read(fixture(encrypted_name)).unwrap_or_else(|_| panic!("read {encrypted_name}"));
    let expected =
        std::fs::read(fixture(plaintext_name)).unwrap_or_else(|_| panic!("read {plaintext_name}"));

    // 1) Sanity check: the high-level Standard decrypt entrypoint can open the fixture.
    let decrypted = decrypt_from_bytes(&encrypted, password)
        .unwrap_or_else(|err| panic!("decrypt_from_bytes({encrypted_name}) failed: {err:?}"));
    assert_eq!(
        decrypted, expected,
        "decrypt_from_bytes output must match {plaintext_name} for {encrypted_name}"
    );

    // 2) Regression: confirm the fixture `EncryptedPackage` specifically uses ECB mode.
    let encryption_info_bytes = read_ole_stream(&encrypted, "EncryptionInfo");
    let encrypted_package_bytes = read_ole_stream(&encrypted, "EncryptedPackage");
    let info =
        match parse_encryption_info(&encryption_info_bytes).expect("parse EncryptionInfo") {
            EncryptionInfo::Standard { header, verifier, .. } => StandardEncryptionInfo { header, verifier },
            other => panic!("expected Standard EncryptionInfo for {encrypted_name}, got {other:?}"),
        };

    // Derive/verify the ECB key (50k spinCount + CryptDeriveKey expansion).
    let key = standard_derive_key_zeroizing(&info, password)
        .unwrap_or_else(|err| panic!("standard_derive_key({encrypted_name}) failed: {err:?}"));
    standard_verify_key(&info, &key)
        .unwrap_or_else(|err| panic!("standard_verify_key({encrypted_name}) failed: {err:?}"));

    let ecb = decrypt_standard_encrypted_package(&key, &encrypted_package_bytes)
        .unwrap_or_else(|err| panic!("ECB decrypt_standard_encrypted_package({encrypted_name}) failed: {err:?}"));
    if ecb != expected {
        let ecb_sha = hex::encode(sha2::Sha256::digest(&ecb));
        let expected_sha = hex::encode(sha2::Sha256::digest(&expected));

        let cbc = formula_offcrypto::encrypted_package::decrypt_standard_encrypted_package_cbc(
            &key,
            &info.verifier.salt,
            &encrypted_package_bytes,
        );
        match cbc {
            Ok(cbc) if cbc == expected => {
                panic!(
                    "{encrypted_name} is NOT ECB-compatible.\n\
                     ECB sha256={ecb_sha}\n\
                     expected sha256={expected_sha}\n\
                     CBC+segmented decrypt matches {plaintext_name}, indicating the fixture EncryptedPackage \
                     uses CBC segmentation, not ECB."
                );
            }
            Ok(cbc) => {
                let cbc_sha = hex::encode(sha2::Sha256::digest(&cbc));
                panic!(
                    "{encrypted_name} ECB decrypt did not match {plaintext_name}.\n\
                     ECB sha256={ecb_sha}\n\
                     CBC sha256={cbc_sha}\n\
                     expected sha256={expected_sha}"
                );
            }
            Err(err) => {
                panic!(
                    "{encrypted_name} ECB decrypt did not match {plaintext_name}.\n\
                     ECB sha256={ecb_sha}\n\
                     expected sha256={expected_sha}\n\
                     CBC+segmented diagnostic attempt errored: {err:?}"
                );
            }
        }
    }

    assert!(ecb.starts_with(b"PK"), "ECB decrypted output should be a ZIP");
    assert_eq!(
        ecb, expected,
        "ECB decrypted bytes must match {plaintext_name} for {encrypted_name}"
    );
}

#[test]
fn standard_fixtures_encryptedpackage_is_ecb() {
    // These are explicitly meant to be “baseline” Standard (CryptoAPI) **AES** fixtures. (The RC4
    // fixture is excluded because ECB/CBC is not applicable to RC4.)
    for (encrypted, plaintext, password) in [
        ("standard.xlsx", "plaintext.xlsx", "password"),
        ("standard-large.xlsx", "plaintext-large.xlsx", "password"),
        ("standard-4.2.xlsx", "plaintext.xlsx", "password"),
        ("standard-unicode.xlsx", "plaintext.xlsx", "pässwörd🔒"),
        ("standard-basic.xlsm", "plaintext-basic.xlsm", "password"),
    ] {
        assert_fixture_is_ecb(encrypted, plaintext, password);
    }
}
