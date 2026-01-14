//! Regression test for the MS-OFFCRYPTO *Standard* (CryptoAPI) `EncryptedPackage` mode used by our
//! committed OOXML fixtures.
//!
//! There is long-running ambiguity across implementations/docs about whether Standard
//! `EncryptedPackage` uses AES-ECB or AES-CBC-with-segmentation. Our `decrypt_from_bytes` helper is
//! the Standard-only **ECB** implementation.
//!
//! This test is a canary to ensure fixture regeneration does not silently drift to a different
//! mode.

use std::io::{Cursor, Read};
use std::path::PathBuf;

use sha2::Digest as _;

use formula_offcrypto::{
    decrypt_from_bytes, parse_encryption_info, standard_derive_key_zeroizing, standard_verify_key,
    EncryptionInfo, OffcryptoError, StandardEncryptionInfo,
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
fn decrypt_standard_cbc_segmented_from_ole(
    raw_ole: &[u8],
    password: &str,
) -> Result<Vec<u8>, OffcryptoError> {
    let encryption_info = read_ole_stream(raw_ole, "EncryptionInfo");
    let encrypted_package = read_ole_stream(raw_ole, "EncryptedPackage");

    let info = match parse_encryption_info(&encryption_info)? {
        EncryptionInfo::Standard { header, verifier, .. } => StandardEncryptionInfo { header, verifier },
        other => {
            return Err(OffcryptoError::InvalidStructure(format!(
                "expected Standard EncryptionInfo for CBC diagnostic, got {other:?}"
            )))
        }
    };

    let key = standard_derive_key_zeroizing(&info, password)?;
    standard_verify_key(&info, &key)?;

    formula_offcrypto::encrypted_package::decrypt_standard_encrypted_package_cbc(
        &key,
        &info.verifier.salt,
        &encrypted_package,
    )
}

#[test]
fn standard_fixture_encryptedpackage_is_ecb() {
    let encrypted = std::fs::read(fixture("standard.xlsx")).expect("read standard.xlsx fixture");
    let expected = std::fs::read(fixture("plaintext.xlsx")).expect("read plaintext.xlsx fixture");

    // Primary assertion: fixtures must be decryptable by the ECB implementation.
    let ecb = match decrypt_from_bytes(&encrypted, "password") {
        Ok(bytes) => bytes,
        Err(err) => {
            // Best-effort diagnostic: try CBC+segmented decryption to see if the fixture drifted.
            let cbc = decrypt_standard_cbc_segmented_from_ole(&encrypted, "password");
            panic!(
                "ECB decrypt_from_bytes failed for fixtures/encrypted/ooxml/standard.xlsx: {err:?}\n\
                 CBC-segmented diagnostic result: {cbc:?}\n\
                 This likely indicates fixture regeneration drift (ECB vs CBC ambiguity)."
            );
        }
    };

    if ecb != expected {
        let ecb_sha = hex::encode(sha2::Sha256::digest(&ecb));
        let expected_sha = hex::encode(sha2::Sha256::digest(&expected));

        // If ECB output doesn't match, try CBC+segmented decryption and include the result in the
        // failure message so it's obvious which mode the fixture uses.
        match decrypt_standard_cbc_segmented_from_ole(&encrypted, "password") {
            Ok(cbc) if cbc == expected => {
                panic!(
                    "fixtures/encrypted/ooxml/standard.xlsx is NOT ECB-compatible: ECB output does \
                     not match plaintext.xlsx.\n\
                     ECB sha256={ecb_sha}\n\
                     expected sha256={expected_sha}\n\
                     CBC+segmented decrypt (diagnostic) *does* match plaintext, indicating the \
                     fixture EncryptedPackage uses CBC segmentation, not ECB."
                );
            }
            Ok(cbc) => {
                let cbc_sha = hex::encode(sha2::Sha256::digest(&cbc));
                panic!(
                    "fixtures/encrypted/ooxml/standard.xlsx ECB decrypt did not match plaintext \
                     (possible fixture drift or implementation bug).\n\
                     ECB sha256={ecb_sha}\n\
                     CBC sha256={cbc_sha}\n\
                     expected sha256={expected_sha}"
                );
            }
            Err(err) => {
                panic!(
                    "fixtures/encrypted/ooxml/standard.xlsx ECB decrypt did not match plaintext \
                     (possible fixture drift or implementation bug).\n\
                     ECB sha256={ecb_sha}\n\
                     expected sha256={expected_sha}\n\
                     CBC+segmented diagnostic attempt errored: {err:?}"
                );
            }
        }
    }

    assert!(
        ecb.starts_with(b"PK"),
        "decrypted output should be a ZIP (PK signature)"
    );
    assert_eq!(ecb, expected, "ECB decrypted bytes must match plaintext.xlsx");
}
