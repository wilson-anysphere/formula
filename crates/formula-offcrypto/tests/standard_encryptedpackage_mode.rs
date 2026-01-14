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

use sha1::Digest as _;

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
    // Avoid copying the whole OLE container for stream extraction.
    let cursor = Cursor::new(raw_ole);
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
    let key = match standard_derive_key_zeroizing(&info, password) {
        Ok(key) => key,
        Err(err) => {
            report_non_ecb_key_derivation(encrypted_name, plaintext_name, password, &info, &encrypted_package_bytes, &expected, &err);
        }
    };
    if let Err(err) = standard_verify_key(&info, &key) {
        report_non_ecb_key_derivation(encrypted_name, plaintext_name, password, &info, &encrypted_package_bytes, &expected, &err);
    }

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

#[cold]
#[track_caller]
fn report_non_ecb_key_derivation(
    encrypted_name: &str,
    plaintext_name: &str,
    password: &str,
    info: &StandardEncryptionInfo,
    encrypted_package: &[u8],
    expected: &[u8],
    ecb_err: &dyn std::fmt::Debug,
) -> ! {
    // If the baseline Standard key derivation fails, try the “CBC variant” key derivation observed
    // in some third-party producers (spinCount=1000 + AES-CBC verifier). We treat this as fixture
    // drift, but include a clear diagnosis in the failure message.
    let key_len = match usize::try_from(info.header.key_size_bits / 8) {
        Ok(v) if v > 0 => v,
        _ => {
            panic!(
                "{encrypted_name}: Standard ECB key derivation/verifier failed ({ecb_err:?}), and keySize={} is not valid",
                info.header.key_size_bits
            );
        }
    };

    let cbc_variant_key = derive_standard_cbc_variant_key(password, &info.verifier.salt, key_len);
    let ecb_out = decrypt_standard_encrypted_package(&cbc_variant_key, encrypted_package);
    let cbc_out = formula_offcrypto::encrypted_package::decrypt_standard_encrypted_package_cbc(
        &cbc_variant_key,
        &info.verifier.salt,
        encrypted_package,
    );

    let ecb_ok = ecb_out.as_ref().is_ok_and(|b| b == expected);
    let cbc_ok = cbc_out.as_ref().is_ok_and(|b| b == expected);

    panic!(
        "{encrypted_name}: fixture is not compatible with baseline Standard ECB key derivation/verifier ({ecb_err:?}).\n\
         CBC-variant key derivation results:\n\
         - ECB EncryptedPackage decrypt matches {plaintext_name}: {ecb_ok}\n\
         - CBC-segmented EncryptedPackage decrypt matches {plaintext_name}: {cbc_ok}\n\
         This likely indicates fixture regeneration drift (Standard key derivation/mode ambiguity)."
    );
}

fn derive_standard_cbc_variant_key(password: &str, salt: &[u8], key_len: usize) -> Vec<u8> {
    // Matches `standard_derive_key_cbc_variant` in `crates/formula-offcrypto/src/lib.rs`:
    // - spinCount = 1000
    // - pwHash = SHA1(salt || UTF16LE(password)), then 1000 rounds of SHA1(LE32(i) || pwHash)
    // - key = SHA1(pwHash || LE32(0)), then truncate/pad to `key_len` (padding byte 0x36)
    const SPIN_COUNT: u32 = 1_000;

    let pw_utf16 = password_utf16le_bytes(password);

    let mut hasher = sha1::Sha1::new();
    hasher.update(salt);
    hasher.update(&pw_utf16);
    let mut h: [u8; 20] = hasher.finalize().into();

    for i in 0..SPIN_COUNT {
        let mut d = sha1::Sha1::new();
        d.update(i.to_le_bytes());
        d.update(h);
        h = d.finalize().into();
    }

    let mut buf = Vec::with_capacity(20 + 4);
    buf.extend_from_slice(&h);
    buf.extend_from_slice(&0u32.to_le_bytes());
    let digest = sha1::Sha1::digest(&buf);

    if key_len <= digest.len() {
        digest[..key_len].to_vec()
    } else {
        let mut out = vec![0x36u8; key_len];
        out[..digest.len()].copy_from_slice(&digest);
        out
    }
}

fn password_utf16le_bytes(password: &str) -> Vec<u8> {
    let mut out = Vec::with_capacity(password.len().saturating_mul(2));
    for unit in password.encode_utf16() {
        out.extend_from_slice(&unit.to_le_bytes());
    }
    out
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
