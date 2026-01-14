#![cfg(not(target_arch = "wasm32"))]

use std::io::{Cursor, Read, Write};
use std::path::PathBuf;

use aes::Aes128;
use base64::engine::general_purpose::STANDARD as BASE64;
use base64::Engine as _;
use cbc::cipher::block_padding::NoPadding;
use cbc::cipher::{BlockEncryptMut, KeyIvInit};
use cfb::CompoundFile;
use hmac::{Hmac, Mac as _};
use ms_offcrypto_writer::Ecma376AgileWriter;
use rand::{rngs::StdRng, SeedableRng};
use sha1::Digest as _;
use zip::write::FileOptions;

use formula_xlsx::{
    decrypt_agile_encrypted_package, decrypt_agile_encrypted_package_with_warnings, OffCryptoError,
    OffCryptoWarning,
};
use formula_xlsx::offcrypto::{
    decrypt_agile_encrypted_package_bytes, decrypt_agile_encrypted_package_with_options, derive_iv,
    derive_key, hash_password, DecryptOptions, HashAlgorithm, DEFAULT_MAX_SPIN_COUNT, HMAC_KEY_BLOCK,
    HMAC_VALUE_BLOCK, KEY_VALUE_BLOCK, VERIFIER_HASH_INPUT_BLOCK, VERIFIER_HASH_VALUE_BLOCK,
};

fn fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/encrypted/ooxml")
        .join(rel)
}

fn build_tiny_zip() -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut writer = zip::ZipWriter::new(cursor);
    writer
        .start_file("hello.txt", FileOptions::<()>::default())
        .expect("start zip file");
    writer.write_all(b"hello").expect("write zip contents");
    writer.finish().expect("finish zip").into_inner()
}

fn build_zip_with_padding() -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut writer = zip::ZipWriter::new(cursor);
    let stored = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Stored);
    writer
        .start_file("hello.txt", stored)
        .expect("start zip file");
    writer.write_all(b"hello").expect("write zip contents");

    // `office-crypto`'s Agile decrypt implementation currently assumes the plaintext payload is at
    // least one full 4096-byte segment. Keep this fixture larger than that so we can compare our
    // implementation against it.
    writer
        .start_file("padding.bin", stored)
        .expect("start padding file");
    writer
        .write_all(&vec![0xA5; 8 * 1024])
        .expect("write padding");

    writer.finish().expect("finish zip").into_inner()
}

fn encrypt_zip_with_password(plain_zip: &[u8], password: &str) -> Vec<u8> {
    let mut cursor = Cursor::new(Vec::new());
    // Use a deterministic RNG seed so these tests don't depend on OS entropy and remain
    // reproducible across CI runs.
    let mut rng = StdRng::from_seed([0u8; 32]);
    let mut agile = Ecma376AgileWriter::create(&mut rng, password, &mut cursor).expect("create agile");
    agile
        .write_all(plain_zip)
        .expect("write plaintext zip to agile writer");
    agile.finalize().expect("finalize agile writer");
    cursor.into_inner()
}

fn extract_stream_bytes(cfb_bytes: &[u8], stream_name: &str) -> Vec<u8> {
    let mut ole = CompoundFile::open(Cursor::new(cfb_bytes)).expect("open cfb");
    // `cfb` stream names can be addressed with or without a leading `/` depending on context, and
    // some producers vary casing. Mirror Formula's best-effort stream resolution so these tests are
    // robust across fixture sources.
    let trimmed = stream_name.trim_start_matches('/');
    let mut stream = ole
        .open_stream(stream_name)
        .or_else(|_| ole.open_stream(trimmed))
        .or_else(|_| ole.open_stream(&format!("/{trimmed}")))
        .or_else(|_| {
            let mut found_path: Option<String> = None;
            for entry in ole.walk() {
                if !entry.is_stream() {
                    continue;
                }
                let path = entry.path().to_string_lossy();
                let normalized = path.as_ref().strip_prefix('/').unwrap_or(path.as_ref());
                if normalized.eq_ignore_ascii_case(trimmed) {
                    found_path = Some(path.into_owned());
                    break;
                }
            }
            let Some(found_path) = found_path else {
                return ole.open_stream(stream_name);
            };
            ole.open_stream(&found_path)
                .or_else(|_| ole.open_stream(found_path.trim_start_matches('/')))
                .or_else(|_| ole.open_stream(&format!("/{trimmed}")))
        })
        .expect("open stream");
    let mut buf = Vec::new();
    stream.read_to_end(&mut buf).expect("read stream");
    buf
}

fn encrypt_aes128_cbc_no_padding(key: &[u8], iv: &[u8], plaintext: &[u8]) -> Vec<u8> {
    assert_eq!(key.len(), 16);
    assert_eq!(iv.len(), 16);
    assert!(
        plaintext.len() % 16 == 0,
        "plaintext must be AES-block aligned"
    );

    let mut buf = plaintext.to_vec();
    let len = buf.len();
    cbc::Encryptor::<Aes128>::new_from_slices(key, iv)
        .expect("valid key/iv")
        .encrypt_padded_mut::<NoPadding>(&mut buf, len)
        .expect("encrypt");
    buf
}

fn compute_hmac_sha1(key: &[u8], data: &[u8]) -> Vec<u8> {
    let mut mac: Hmac<sha1::Sha1> = Hmac::new_from_slice(key).expect("hmac key");
    mac.update(data);
    mac.finalize().into_bytes().to_vec()
}

#[test]
fn agile_decrypt_roundtrip() {
    let password = "correct horse battery staple";
    let plain_zip = build_tiny_zip();

    let encrypted_cfb = encrypt_zip_with_password(&plain_zip, password);
    let encryption_info = extract_stream_bytes(&encrypted_cfb, "/EncryptionInfo");
    let encrypted_package = extract_stream_bytes(&encrypted_cfb, "/EncryptedPackage");

    let decrypted =
        decrypt_agile_encrypted_package(&encryption_info, &encrypted_package, password).unwrap();
    assert_eq!(decrypted, plain_zip);
}

#[test]
fn agile_decrypt_roundtrip_empty_password() {
    let plain_zip = build_tiny_zip();
    let encrypted_cfb = encrypt_zip_with_password(&plain_zip, "");
    let encryption_info = extract_stream_bytes(&encrypted_cfb, "/EncryptionInfo");
    let encrypted_package = extract_stream_bytes(&encrypted_cfb, "/EncryptedPackage");

    let decrypted = decrypt_agile_encrypted_package(&encryption_info, &encrypted_package, "")
        .expect("empty password should decrypt");
    assert_eq!(decrypted, plain_zip);
}

#[test]
fn agile_decrypt_tampered_size_header_fails_integrity() {
    let password = "correct horse battery staple";
    let plain_zip = build_tiny_zip();

    let encrypted_cfb = encrypt_zip_with_password(&plain_zip, password);
    let encryption_info = extract_stream_bytes(&encrypted_cfb, "/EncryptionInfo");
    let mut encrypted_package = extract_stream_bytes(&encrypted_cfb, "/EncryptedPackage");

    // Tamper the 8-byte plaintext size prefix. Excel computes `dataIntegrity` over the full
    // `EncryptedPackage` stream bytes (including this header), so this must fail.
    let original_size = u64::from_le_bytes(
        encrypted_package[..8]
            .try_into()
            .expect("EncryptedPackage header is 8 bytes"),
    );
    assert!(original_size > 0, "unexpected empty EncryptedPackage payload");
    let tampered_size = original_size - 1;
    encrypted_package[..8].copy_from_slice(&tampered_size.to_le_bytes());

    let err = decrypt_agile_encrypted_package(&encryption_info, &encrypted_package, password)
        .expect_err("expected integrity failure");
    assert!(
        matches!(err, OffCryptoError::IntegrityMismatch),
        "expected IntegrityMismatch, got {err:?}"
    );
}

#[test]
fn agile_decrypt_tampered_size_header_high_dword_fails_integrity() {
    let password = "correct horse battery staple";
    let plain_zip = build_tiny_zip();

    let encrypted_cfb = encrypt_zip_with_password(&plain_zip, password);
    let encryption_info = extract_stream_bytes(&encrypted_cfb, "/EncryptionInfo");
    let mut encrypted_package = extract_stream_bytes(&encrypted_cfb, "/EncryptedPackage");

    // Tamper only the high DWORD of the 8-byte size prefix. Some producers treat it as reserved for
    // length semantics, but MS-OFFCRYPTO defines `dataIntegrity` over the full EncryptedPackage
    // stream bytes (including all 8 header bytes), so this must fail.
    assert!(
        encrypted_package.len() >= 8,
        "EncryptedPackage stream is unexpectedly small"
    );
    encrypted_package[4..8].copy_from_slice(&1u32.to_le_bytes());

    let err = decrypt_agile_encrypted_package(&encryption_info, &encrypted_package, password)
        .expect_err("expected integrity failure");
    assert!(
        matches!(err, OffCryptoError::IntegrityMismatch),
        "expected IntegrityMismatch, got {err:?}"
    );
}

#[test]
fn agile_decrypt_appended_ciphertext_fails_integrity() {
    let password = "correct horse battery staple";
    let plain_zip = build_tiny_zip();

    let encrypted_cfb = encrypt_zip_with_password(&plain_zip, password);
    let encryption_info = extract_stream_bytes(&encrypted_cfb, "/EncryptionInfo");
    let mut encrypted_package = extract_stream_bytes(&encrypted_cfb, "/EncryptedPackage");

    // Append an extra AES block to the ciphertext. `dataIntegrity` authenticates the *entire*
    // `EncryptedPackage` stream bytes, so this should be detected.
    encrypted_package.extend_from_slice(&[0xA5u8; 16]);

    let err = decrypt_agile_encrypted_package(&encryption_info, &encrypted_package, password)
        .expect_err("expected integrity failure");
    assert!(
        matches!(err, OffCryptoError::IntegrityMismatch),
        "expected IntegrityMismatch, got {err:?}"
    );
}

#[test]
fn agile_decrypt_wrong_password_fails() {
    let plain_zip = build_tiny_zip();
    let encrypted_cfb = encrypt_zip_with_password(&plain_zip, "password-1");

    let encryption_info = extract_stream_bytes(&encrypted_cfb, "/EncryptionInfo");
    let encrypted_package = extract_stream_bytes(&encrypted_cfb, "/EncryptedPackage");

    let err = decrypt_agile_encrypted_package(&encryption_info, &encrypted_package, "password-2")
        .expect_err("wrong password should fail");
    match err {
        OffCryptoError::WrongPassword | OffCryptoError::IntegrityMismatch => {}
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn agile_decrypt_decrypts_empty_password_fixture() {
    let encrypted_cfb = std::fs::read(fixture_path("agile-empty-password.xlsx"))
        .expect("read agile-empty-password.xlsx");
    let expected = std::fs::read(fixture_path("plaintext.xlsx")).expect("read plaintext.xlsx");

    let encryption_info = extract_stream_bytes(&encrypted_cfb, "/EncryptionInfo");
    let encrypted_package = extract_stream_bytes(&encrypted_cfb, "/EncryptedPackage");

    let decrypted = decrypt_agile_encrypted_package(&encryption_info, &encrypted_package, "")
        .expect("decrypt empty-password fixture");
    assert_eq!(decrypted, expected);
}

#[test]
fn agile_decrypt_decrypts_password_fixture() {
    let encrypted_cfb = std::fs::read(fixture_path("agile.xlsx")).expect("read agile.xlsx");
    let expected = std::fs::read(fixture_path("plaintext.xlsx")).expect("read plaintext.xlsx");

    let encryption_info = extract_stream_bytes(&encrypted_cfb, "/EncryptionInfo");
    let encrypted_package = extract_stream_bytes(&encrypted_cfb, "/EncryptedPackage");

    let decrypted = decrypt_agile_encrypted_package(&encryption_info, &encrypted_package, "password")
        .expect("decrypt agile fixture");
    assert_eq!(decrypted, expected);
}

#[test]
fn agile_decrypt_errors_on_non_block_aligned_verifier_ciphertext() {
    // Build a minimal Agile `EncryptionInfo` stream where one ciphertext blob decodes to a
    // non-multiple-of-16 length. The decrypt implementation should fail early with a structured
    // error pointing at the offending field.
    let invalid = BASE64.encode([0u8; 15]); // 15 % 16 != 0
    let valid = BASE64.encode([0u8; 16]);

  let xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
            xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
  <keyData saltValue="{valid}" hashAlgorithm="SHA1" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
           blockSize="16" keyBits="128" hashSize="20" saltSize="16" />
  <dataIntegrity encryptedHmacKey="{valid}" encryptedHmacValue="{valid}" />
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey saltValue="{valid}" hashAlgorithm="SHA1" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                      spinCount="0" blockSize="16" keyBits="128" hashSize="20" saltSize="16"
                      encryptedVerifierHashInput="{invalid}"
                      encryptedVerifierHashValue="{valid}"
                      encryptedKeyValue="{valid}" />
    </keyEncryptor>
  </keyEncryptors>
</encryption>"#
    );

    let mut encryption_info = Vec::new();
    encryption_info.extend_from_slice(&4u16.to_le_bytes());
    encryption_info.extend_from_slice(&4u16.to_le_bytes());
    encryption_info.extend_from_slice(&0u32.to_le_bytes());
    encryption_info.extend_from_slice(xml.as_bytes());

    let err = decrypt_agile_encrypted_package(&encryption_info, &[], "pw")
        .expect_err("expected CiphertextNotBlockAligned");
    assert!(
        matches!(
            err,
            OffCryptoError::CiphertextNotBlockAligned {
                field: "encryptedVerifierHashInput",
                len: 15
            }
        ),
        "unexpected error: {err:?}"
    );
}

#[test]
fn agile_decrypt_rejects_spin_count_above_default_max() {
    let invalid = BASE64.encode([0u8; 15]); // 15 % 16 != 0
    let valid = BASE64.encode([0u8; 16]);
    let spin_count = DEFAULT_MAX_SPIN_COUNT.saturating_add(1);

    let xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
            xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
  <keyData saltValue="{valid}" hashAlgorithm="SHA1" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
           blockSize="16" keyBits="128" hashSize="20" saltSize="16" />
  <dataIntegrity encryptedHmacKey="{valid}" encryptedHmacValue="{valid}" />
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey saltValue="{valid}" hashAlgorithm="SHA1" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                      spinCount="{spin_count}" blockSize="16" keyBits="128" hashSize="20" saltSize="16"
                      encryptedVerifierHashInput="{invalid}"
                      encryptedVerifierHashValue="{valid}"
                      encryptedKeyValue="{valid}" />
    </keyEncryptor>
  </keyEncryptors>
</encryption>"#
    );

    let mut encryption_info = Vec::new();
    encryption_info.extend_from_slice(&4u16.to_le_bytes());
    encryption_info.extend_from_slice(&4u16.to_le_bytes());
    encryption_info.extend_from_slice(&0u32.to_le_bytes());
    encryption_info.extend_from_slice(xml.as_bytes());

    let err = decrypt_agile_encrypted_package(&encryption_info, &[], "pw")
        .expect_err("expected SpinCountTooLarge");
    assert!(
        matches!(
            err,
            OffCryptoError::SpinCountTooLarge {
                spin_count: s,
                max
            } if s == spin_count && max == DEFAULT_MAX_SPIN_COUNT
        ),
        "unexpected error: {err:?}"
    );
}

#[test]
fn agile_decrypt_checks_spin_count_before_decoding_password_salt() {
    // If `spinCount` is above the configured max, we should fail fast without attempting to decode
    // other (possibly malformed) base64 fields.
    //
    // This guards against inputs that combine an oversized `spinCount` (CPU DoS) with malformed
    // blobs intended to trigger additional work during parse.
    let valid = BASE64.encode([0u8; 16]);
    let spin_count = DEFAULT_MAX_SPIN_COUNT.saturating_add(1);

    let xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
            xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
  <keyData saltValue="{valid}" hashAlgorithm="SHA1" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
           blockSize="16" keyBits="128" hashSize="20" saltSize="16" />
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey saltValue="!!!!" hashAlgorithm="SHA1" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                      spinCount="{spin_count}" blockSize="16" keyBits="128" hashSize="20" saltSize="16"
                      encryptedVerifierHashInput="{valid}"
                      encryptedVerifierHashValue="{valid}"
                      encryptedKeyValue="{valid}" />
    </keyEncryptor>
  </keyEncryptors>
</encryption>"#
    );

    let mut encryption_info = Vec::new();
    encryption_info.extend_from_slice(&4u16.to_le_bytes());
    encryption_info.extend_from_slice(&4u16.to_le_bytes());
    encryption_info.extend_from_slice(&0u32.to_le_bytes());
    encryption_info.extend_from_slice(xml.as_bytes());

    let err = decrypt_agile_encrypted_package(&encryption_info, &[], "pw")
        .expect_err("expected SpinCountTooLarge");
    assert!(
        matches!(
            err,
            OffCryptoError::SpinCountTooLarge {
                spin_count: s,
                max
            } if s == spin_count && max == DEFAULT_MAX_SPIN_COUNT
        ),
        "unexpected error: {err:?}"
    );
}

#[test]
fn agile_decrypt_allows_overriding_max_spin_count() {
    let invalid = BASE64.encode([0u8; 15]); // 15 % 16 != 0
    let valid = BASE64.encode([0u8; 16]);
    let spin_count = DEFAULT_MAX_SPIN_COUNT.saturating_add(1);

    let xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
            xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
  <keyData saltValue="{valid}" hashAlgorithm="SHA1" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
           blockSize="16" keyBits="128" hashSize="20" saltSize="16" />
  <dataIntegrity encryptedHmacKey="{valid}" encryptedHmacValue="{valid}" />
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey saltValue="{valid}" hashAlgorithm="SHA1" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                      spinCount="{spin_count}" blockSize="16" keyBits="128" hashSize="20" saltSize="16"
                      encryptedVerifierHashInput="{invalid}"
                      encryptedVerifierHashValue="{valid}"
                      encryptedKeyValue="{valid}" />
    </keyEncryptor>
  </keyEncryptors>
</encryption>"#
    );

    let mut encryption_info = Vec::new();
    encryption_info.extend_from_slice(&4u16.to_le_bytes());
    encryption_info.extend_from_slice(&4u16.to_le_bytes());
    encryption_info.extend_from_slice(&0u32.to_le_bytes());
    encryption_info.extend_from_slice(xml.as_bytes());

    let opts = DecryptOptions {
        max_spin_count: spin_count,
    };
    let err = decrypt_agile_encrypted_package_with_options(&encryption_info, &[], "pw", &opts)
        .expect_err("expected CiphertextNotBlockAligned");
    assert!(
        matches!(
            err,
            OffCryptoError::CiphertextNotBlockAligned {
                field: "encryptedVerifierHashInput",
                len: 15
            }
        ),
        "unexpected error: {err:?}"
    );
}

#[test]
fn agile_decrypt_ignores_trailing_padding_in_verifier_and_hmac_values() {
    // Synthetic Agile EncryptionInfo where the decrypted verifierHashValue and encryptedHmacValue
    // contain a correct digest prefix followed by non-zero garbage (e.g. producer-specific padding).
    //
    // Our decryptor/validator must compare only the first `hashSize` bytes.
    let password = "pw";
    let plain_zip = build_tiny_zip();

    // Use SHA1 so `hashSize=20` is not AES-block aligned (16), forcing padding/extra bytes.
    let hash_alg = HashAlgorithm::Sha1;
    let hash_size = 20usize;
    let block_size = 16usize;
    let key_encrypt_key_len = 16usize;

    let key_data_salt: Vec<u8> = (0u8..=15).collect();
    let password_salt: Vec<u8> = (16u8..=31).collect();
    let spin_count = 10u32;

    let package_key: Vec<u8> = (32u8..=47).collect(); // AES-128 keyValue

    // --- Build EncryptedPackage stream ---------------------------------------------------------
    let iv0 = derive_iv(&key_data_salt, &0u32.to_le_bytes(), block_size, hash_alg).unwrap();
    let padded_plain_zip = {
        let mut out = plain_zip.clone();
        out.extend(std::iter::repeat(0u8).take((16 - (out.len() % 16)) % 16));
        out
    };
    let ciphertext = encrypt_aes128_cbc_no_padding(&package_key, &iv0, &padded_plain_zip);
    let mut encrypted_package = Vec::new();
    encrypted_package.extend_from_slice(&(plain_zip.len() as u64).to_le_bytes());
    encrypted_package.extend_from_slice(&ciphertext);

    // --- Build password key encryptor fields ---------------------------------------------------
    let password_hash = hash_password(password, &password_salt, spin_count, hash_alg).unwrap();
    let verifier_iv = &password_salt[..block_size];

    let verifier_input: Vec<u8> = b"abcdefghijklmnop".to_vec();
    let verifier_hash: Vec<u8> = sha1::Sha1::digest(&verifier_input).to_vec();

    // Make verifierHashValue plaintext block-aligned by appending non-zero garbage after the digest.
    let mut verifier_hash_value_plain = verifier_hash.clone();
    verifier_hash_value_plain.extend_from_slice(&[0xA5u8; 12]); // garbage beyond hashSize
    assert_eq!(verifier_hash_value_plain.len(), 32);

    let encrypt_pw_blob = |block_key: &[u8], plaintext: &[u8]| -> Vec<u8> {
        let k = derive_key(&password_hash, block_key, key_encrypt_key_len, hash_alg).unwrap();
        encrypt_aes128_cbc_no_padding(&k, verifier_iv, plaintext)
    };

    let encrypted_verifier_hash_input = encrypt_pw_blob(&VERIFIER_HASH_INPUT_BLOCK, &verifier_input);
    let encrypted_verifier_hash_value =
        encrypt_pw_blob(&VERIFIER_HASH_VALUE_BLOCK, &verifier_hash_value_plain);
    let encrypted_key_value = encrypt_pw_blob(&KEY_VALUE_BLOCK, &package_key);

    // --- Build dataIntegrity fields ------------------------------------------------------------
    let hmac_key_plain: Vec<u8> = vec![0x11u8; hash_size];
    let actual_hmac = compute_hmac_sha1(&hmac_key_plain, &encrypted_package);
    assert_eq!(actual_hmac.len(), hash_size);

    // Non-zero garbage padding after hashSize.
    let mut hmac_key_blob = hmac_key_plain.clone();
    hmac_key_blob.extend_from_slice(&[0x5Au8; 12]);
    let mut hmac_value_blob = actual_hmac.clone();
    hmac_value_blob.extend_from_slice(&[0xC3u8; 12]);

    let hmac_key_iv = derive_iv(&key_data_salt, &HMAC_KEY_BLOCK, block_size, hash_alg).unwrap();
    let encrypted_hmac_key = encrypt_aes128_cbc_no_padding(&package_key, &hmac_key_iv, &hmac_key_blob);
    let hmac_val_iv = derive_iv(&key_data_salt, &HMAC_VALUE_BLOCK, block_size, hash_alg).unwrap();
    let encrypted_hmac_value =
        encrypt_aes128_cbc_no_padding(&package_key, &hmac_val_iv, &hmac_value_blob);

    // --- Build EncryptionInfo stream -----------------------------------------------------------
    let xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
            xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
  <keyData saltSize="16" blockSize="{block_size}" keyBits="128" hashSize="{hash_size}"
           cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA1"
           saltValue="{key_data_salt_b64}"/>
  <dataIntegrity encryptedHmacKey="{ehk_b64}" encryptedHmacValue="{ehv_b64}"/>
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey saltSize="16" blockSize="{block_size}" keyBits="128" hashSize="{hash_size}"
                      spinCount="{spin_count}" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA1"
                      saltValue="{password_salt_b64}"
                      encryptedVerifierHashInput="{evhi_b64}"
                      encryptedVerifierHashValue="{evhv_b64}"
                      encryptedKeyValue="{ekv_b64}"/>
    </keyEncryptor>
  </keyEncryptors>
</encryption>"#,
        key_data_salt_b64 = BASE64.encode(&key_data_salt),
        password_salt_b64 = BASE64.encode(&password_salt),
        ehk_b64 = BASE64.encode(&encrypted_hmac_key),
        ehv_b64 = BASE64.encode(&encrypted_hmac_value),
        evhi_b64 = BASE64.encode(&encrypted_verifier_hash_input),
        evhv_b64 = BASE64.encode(&encrypted_verifier_hash_value),
        ekv_b64 = BASE64.encode(&encrypted_key_value),
    );

    let mut encryption_info = Vec::new();
    encryption_info.extend_from_slice(&4u16.to_le_bytes()); // major
    encryption_info.extend_from_slice(&4u16.to_le_bytes()); // minor
    encryption_info.extend_from_slice(&0u32.to_le_bytes()); // flags
    encryption_info.extend_from_slice(xml.as_bytes());

    let decrypted =
        decrypt_agile_encrypted_package(&encryption_info, &encrypted_package, password).unwrap();
    assert_eq!(decrypted, plain_zip);
}

#[test]
fn agile_decrypt_accepts_short_hmac_key() {
    // Some producers emit a decrypted `encryptedHmacKey` whose length is shorter than `hashSize`.
    // HMAC accepts any key length, so we should accept such files as long as the computed digest
    // matches `encryptedHmacValue` (first `hashSize` bytes).
    let password = "pw";
    let plain_zip = build_tiny_zip();

    // Use SHA1 so `hashSize=20` is not AES-block aligned (16), forcing padding in `encryptedHmacValue`.
    let hash_alg = HashAlgorithm::Sha1;
    let hash_size = 20usize;
    let block_size = 16usize;
    let key_encrypt_key_len = 16usize;

    let key_data_salt: Vec<u8> = (0u8..=15).collect();
    let password_salt: Vec<u8> = (16u8..=31).collect();
    let spin_count = 10u32;

    let package_key: Vec<u8> = (32u8..=47).collect(); // AES-128 keyValue

    // --- Build EncryptedPackage stream ---------------------------------------------------------
    let iv0 = derive_iv(&key_data_salt, &0u32.to_le_bytes(), block_size, hash_alg).unwrap();
    let padded_plain_zip = {
        let mut out = plain_zip.clone();
        out.extend(std::iter::repeat(0u8).take((16 - (out.len() % 16)) % 16));
        out
    };
    let ciphertext = encrypt_aes128_cbc_no_padding(&package_key, &iv0, &padded_plain_zip);
    let mut encrypted_package = Vec::new();
    encrypted_package.extend_from_slice(&(plain_zip.len() as u64).to_le_bytes());
    encrypted_package.extend_from_slice(&ciphertext);

    // --- Build password key encryptor fields ---------------------------------------------------
    let password_hash = hash_password(password, &password_salt, spin_count, hash_alg).unwrap();
    let verifier_iv = &password_salt[..block_size];

    let verifier_input: Vec<u8> = b"abcdefghijklmnop".to_vec();
    let verifier_hash: Vec<u8> = sha1::Sha1::digest(&verifier_input).to_vec();

    let mut verifier_hash_value_plain = verifier_hash.clone();
    verifier_hash_value_plain.extend_from_slice(&[0xA5u8; 12]); // garbage beyond hashSize
    assert_eq!(verifier_hash_value_plain.len(), 32);

    let encrypt_pw_blob = |block_key: &[u8], plaintext: &[u8]| -> Vec<u8> {
        let k = derive_key(&password_hash, block_key, key_encrypt_key_len, hash_alg).unwrap();
        encrypt_aes128_cbc_no_padding(&k, verifier_iv, plaintext)
    };

    let encrypted_verifier_hash_input = encrypt_pw_blob(&VERIFIER_HASH_INPUT_BLOCK, &verifier_input);
    let encrypted_verifier_hash_value =
        encrypt_pw_blob(&VERIFIER_HASH_VALUE_BLOCK, &verifier_hash_value_plain);
    let encrypted_key_value = encrypt_pw_blob(&KEY_VALUE_BLOCK, &package_key);

    // --- Build dataIntegrity fields ------------------------------------------------------------
    let hmac_key_plain: Vec<u8> = vec![0x11u8; 16]; // shorter than hashSize=20
    let actual_hmac = compute_hmac_sha1(&hmac_key_plain, &encrypted_package);
    assert_eq!(actual_hmac.len(), hash_size);

    // `encryptedHmacKey` decrypts to 16 bytes (one AES block), but we still compare the full
    // `hashSize` bytes of the HMAC output against `encryptedHmacValue`.
    let hmac_key_blob = hmac_key_plain;

    // Non-zero garbage padding after hashSize.
    let mut hmac_value_blob = actual_hmac.clone();
    hmac_value_blob.extend_from_slice(&[0xC3u8; 12]);

    let hmac_key_iv = derive_iv(&key_data_salt, &HMAC_KEY_BLOCK, block_size, hash_alg).unwrap();
    let encrypted_hmac_key = encrypt_aes128_cbc_no_padding(&package_key, &hmac_key_iv, &hmac_key_blob);
    let hmac_val_iv = derive_iv(&key_data_salt, &HMAC_VALUE_BLOCK, block_size, hash_alg).unwrap();
    let encrypted_hmac_value =
        encrypt_aes128_cbc_no_padding(&package_key, &hmac_val_iv, &hmac_value_blob);

    // --- Build EncryptionInfo stream -----------------------------------------------------------
    let xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
            xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
  <keyData saltSize="16" blockSize="{block_size}" keyBits="128" hashSize="{hash_size}"
           cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA1"
           saltValue="{key_data_salt_b64}"/>
  <dataIntegrity encryptedHmacKey="{ehk_b64}" encryptedHmacValue="{ehv_b64}"/>
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey saltSize="16" blockSize="{block_size}" keyBits="128" hashSize="{hash_size}"
                      spinCount="{spin_count}" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA1"
                      saltValue="{password_salt_b64}"
                      encryptedVerifierHashInput="{evhi_b64}"
                      encryptedVerifierHashValue="{evhv_b64}"
                      encryptedKeyValue="{ekv_b64}"/>
    </keyEncryptor>
  </keyEncryptors>
</encryption>"#,
        key_data_salt_b64 = BASE64.encode(&key_data_salt),
        password_salt_b64 = BASE64.encode(&password_salt),
        ehk_b64 = BASE64.encode(&encrypted_hmac_key),
        ehv_b64 = BASE64.encode(&encrypted_hmac_value),
        evhi_b64 = BASE64.encode(&encrypted_verifier_hash_input),
        evhv_b64 = BASE64.encode(&encrypted_verifier_hash_value),
        ekv_b64 = BASE64.encode(&encrypted_key_value),
    );

    let mut encryption_info = Vec::new();
    encryption_info.extend_from_slice(&4u16.to_le_bytes()); // major
    encryption_info.extend_from_slice(&4u16.to_le_bytes()); // minor
    encryption_info.extend_from_slice(&0u32.to_le_bytes()); // flags
    encryption_info.extend_from_slice(xml.as_bytes());

    let decrypted =
        decrypt_agile_encrypted_package(&encryption_info, &encrypted_package, password).unwrap();
    assert_eq!(decrypted, plain_zip);
}

#[test]
fn agile_decrypt_accepts_key_encryptor_blobs_as_child_elements() {
    // Some producers encode the three password key-encryptor ciphertext blobs as child elements
    // instead of attributes:
    //   <p:encryptedKey ...>
    //     <p:encryptedVerifierHashInput>...</p:encryptedVerifierHashInput>
    //     <p:encryptedVerifierHashValue>...</p:encryptedVerifierHashValue>
    //     <p:encryptedKeyValue>...</p:encryptedKeyValue>
    //   </p:encryptedKey>
    //
    // Formula should accept either representation.
    let password = "pw";
    let plain_zip = build_tiny_zip();

    // Use SHA1 so `hashSize=20` is not AES-block aligned (16), forcing padding in verifier/HMAC values.
    let hash_alg = HashAlgorithm::Sha1;
    let hash_size = 20usize;
    let block_size = 16usize;
    let key_encrypt_key_len = 16usize;

    let key_data_salt: Vec<u8> = (0u8..=15).collect();
    let password_salt: Vec<u8> = (16u8..=31).collect();
    let spin_count = 10u32;

    let package_key: Vec<u8> = (32u8..=47).collect(); // AES-128 keyValue

    // --- Build EncryptedPackage stream ---------------------------------------------------------
    let iv0 = derive_iv(&key_data_salt, &0u32.to_le_bytes(), block_size, hash_alg).unwrap();
    let padded_plain_zip = {
        let mut out = plain_zip.clone();
        out.extend(std::iter::repeat(0u8).take((16 - (out.len() % 16)) % 16));
        out
    };
    let ciphertext = encrypt_aes128_cbc_no_padding(&package_key, &iv0, &padded_plain_zip);
    let mut encrypted_package = Vec::new();
    encrypted_package.extend_from_slice(&(plain_zip.len() as u64).to_le_bytes());
    encrypted_package.extend_from_slice(&ciphertext);

    // --- Build password key encryptor fields ---------------------------------------------------
    let password_hash = hash_password(password, &password_salt, spin_count, hash_alg).unwrap();
    let verifier_iv = &password_salt[..block_size];

    let verifier_input: Vec<u8> = b"abcdefghijklmnop".to_vec();
    let verifier_hash: Vec<u8> = sha1::Sha1::digest(&verifier_input).to_vec();

    let mut verifier_hash_value_plain = verifier_hash.clone();
    verifier_hash_value_plain.extend_from_slice(&[0xA5u8; 12]); // garbage beyond hashSize
    assert_eq!(verifier_hash_value_plain.len(), 32);

    let encrypt_pw_blob = |block_key: &[u8], plaintext: &[u8]| -> Vec<u8> {
        let k = derive_key(&password_hash, block_key, key_encrypt_key_len, hash_alg).unwrap();
        encrypt_aes128_cbc_no_padding(&k, verifier_iv, plaintext)
    };

    let encrypted_verifier_hash_input = encrypt_pw_blob(&VERIFIER_HASH_INPUT_BLOCK, &verifier_input);
    let encrypted_verifier_hash_value =
        encrypt_pw_blob(&VERIFIER_HASH_VALUE_BLOCK, &verifier_hash_value_plain);
    let encrypted_key_value = encrypt_pw_blob(&KEY_VALUE_BLOCK, &package_key);

    // --- Build dataIntegrity fields ------------------------------------------------------------
    let hmac_key_plain: Vec<u8> = vec![0x11u8; hash_size];
    let actual_hmac = compute_hmac_sha1(&hmac_key_plain, &encrypted_package);
    assert_eq!(actual_hmac.len(), hash_size);

    let mut hmac_key_blob = hmac_key_plain.clone();
    hmac_key_blob.extend_from_slice(&[0x5Au8; 12]);
    let mut hmac_value_blob = actual_hmac.clone();
    hmac_value_blob.extend_from_slice(&[0xC3u8; 12]);

    let hmac_key_iv = derive_iv(&key_data_salt, &HMAC_KEY_BLOCK, block_size, hash_alg).unwrap();
    let encrypted_hmac_key = encrypt_aes128_cbc_no_padding(&package_key, &hmac_key_iv, &hmac_key_blob);
    let hmac_val_iv = derive_iv(&key_data_salt, &HMAC_VALUE_BLOCK, block_size, hash_alg).unwrap();
    let encrypted_hmac_value =
        encrypt_aes128_cbc_no_padding(&package_key, &hmac_val_iv, &hmac_value_blob);

    // --- Build EncryptionInfo stream -----------------------------------------------------------
    let xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
            xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
  <keyData saltSize="16" blockSize="{block_size}" keyBits="128" hashSize="{hash_size}"
           cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA1"
           saltValue="{key_data_salt_b64}"/>
  <dataIntegrity encryptedHmacKey="{ehk_b64}" encryptedHmacValue="{ehv_b64}"/>
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey saltSize="16" blockSize="{block_size}" keyBits="128" hashSize="{hash_size}"
                      spinCount="{spin_count}" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA1"
                      saltValue="{password_salt_b64}">
        <p:encryptedVerifierHashInput>{evhi_b64}</p:encryptedVerifierHashInput>
        <p:encryptedVerifierHashValue>{evhv_b64}</p:encryptedVerifierHashValue>
        <p:encryptedKeyValue>{ekv_b64}</p:encryptedKeyValue>
      </p:encryptedKey>
    </keyEncryptor>
  </keyEncryptors>
</encryption>"#,
        key_data_salt_b64 = BASE64.encode(&key_data_salt),
        password_salt_b64 = BASE64.encode(&password_salt),
        ehk_b64 = BASE64.encode(&encrypted_hmac_key),
        ehv_b64 = BASE64.encode(&encrypted_hmac_value),
        evhi_b64 = BASE64.encode(&encrypted_verifier_hash_input),
        evhv_b64 = BASE64.encode(&encrypted_verifier_hash_value),
        ekv_b64 = BASE64.encode(&encrypted_key_value),
    );

    let mut encryption_info = Vec::new();
    encryption_info.extend_from_slice(&4u16.to_le_bytes()); // major
    encryption_info.extend_from_slice(&4u16.to_le_bytes()); // minor
    encryption_info.extend_from_slice(&0u32.to_le_bytes()); // flags
    encryption_info.extend_from_slice(xml.as_bytes());

    let decrypted =
        decrypt_agile_encrypted_package(&encryption_info, &encrypted_package, password).unwrap();
    assert_eq!(decrypted, plain_zip);

    // Also ensure the alternate Agile parser/decryptor path (offcrypto::agile) accepts the same XML.
    let decrypted_alt =
        decrypt_agile_encrypted_package_bytes(&encryption_info, &encrypted_package, password).unwrap();
    assert_eq!(decrypted_alt, plain_zip);
}

#[test]
fn agile_decrypt_falls_back_to_derived_iv_for_password_key_encryptor_blobs() {
    // Synthetic Agile EncryptionInfo where the password key-encryptor blobs are encrypted with a
    // derived IV (Hash(saltValue || blockKey)) instead of Excel's typical `IV = saltValue[..blockSize]`.
    //
    // Formula should try the Excel IV strategy first, then fall back to the derived IV strategy on
    // verifier mismatch.
    let password = "pw";
    let plain_zip = build_tiny_zip();

    // Use SHA1 to keep the fixture compact and to exercise padding in verifier/HMAC values.
    let hash_alg = HashAlgorithm::Sha1;
    let hash_size = 20usize;
    let block_size = 16usize;
    let key_encrypt_key_len = 16usize;

    let key_data_salt: Vec<u8> = (0u8..=15).collect();
    let password_salt: Vec<u8> = (16u8..=31).collect();
    let spin_count = 10u32;

    let package_key: Vec<u8> = (32u8..=47).collect(); // AES-128 keyValue

    // --- Build EncryptedPackage stream ---------------------------------------------------------
    let iv0 = derive_iv(&key_data_salt, &0u32.to_le_bytes(), block_size, hash_alg).unwrap();
    let padded_plain_zip = {
        let mut out = plain_zip.clone();
        out.extend(std::iter::repeat(0u8).take((16 - (out.len() % 16)) % 16));
        out
    };
    let ciphertext = encrypt_aes128_cbc_no_padding(&package_key, &iv0, &padded_plain_zip);
    let mut encrypted_package = Vec::new();
    encrypted_package.extend_from_slice(&(plain_zip.len() as u64).to_le_bytes());
    encrypted_package.extend_from_slice(&ciphertext);

    // --- Build password key encryptor fields ---------------------------------------------------
    let password_hash = hash_password(password, &password_salt, spin_count, hash_alg).unwrap();
    let salt_iv = &password_salt[..block_size];
    let derived_iv =
        derive_iv(&password_salt, &VERIFIER_HASH_INPUT_BLOCK, block_size, hash_alg).unwrap();
    assert_ne!(
        derived_iv.as_slice(),
        salt_iv,
        "derived-IV scheme should not accidentally match Excel's saltValue IV"
    );

    let verifier_input: Vec<u8> = b"abcdefghijklmnop".to_vec();
    let verifier_hash: Vec<u8> = sha1::Sha1::digest(&verifier_input).to_vec();

    // Make verifierHashValue plaintext block-aligned by appending non-zero garbage after the digest.
    let mut verifier_hash_value_plain = verifier_hash.clone();
    verifier_hash_value_plain.extend_from_slice(&[0xA5u8; 12]); // garbage beyond hashSize
    assert_eq!(verifier_hash_value_plain.len(), 32);

    let encrypt_pw_blob = |block_key: &[u8], plaintext: &[u8]| -> Vec<u8> {
        let k = derive_key(&password_hash, block_key, key_encrypt_key_len, hash_alg).unwrap();
        let iv = derive_iv(&password_salt, block_key, block_size, hash_alg).unwrap();
        encrypt_aes128_cbc_no_padding(&k, &iv, plaintext)
    };

    let encrypted_verifier_hash_input = encrypt_pw_blob(&VERIFIER_HASH_INPUT_BLOCK, &verifier_input);
    let encrypted_verifier_hash_value =
        encrypt_pw_blob(&VERIFIER_HASH_VALUE_BLOCK, &verifier_hash_value_plain);
    let encrypted_key_value = encrypt_pw_blob(&KEY_VALUE_BLOCK, &package_key);

    // --- Build dataIntegrity fields ------------------------------------------------------------
    let hmac_key_plain: Vec<u8> = vec![0x11u8; hash_size];
    let actual_hmac = compute_hmac_sha1(&hmac_key_plain, &encrypted_package);
    assert_eq!(actual_hmac.len(), hash_size);

    // Non-zero garbage padding after hashSize.
    let mut hmac_key_blob = hmac_key_plain.clone();
    hmac_key_blob.extend_from_slice(&[0x5Au8; 12]);
    let mut hmac_value_blob = actual_hmac.clone();
    hmac_value_blob.extend_from_slice(&[0xC3u8; 12]);

    let hmac_key_iv = derive_iv(&key_data_salt, &HMAC_KEY_BLOCK, block_size, hash_alg).unwrap();
    let encrypted_hmac_key =
        encrypt_aes128_cbc_no_padding(&package_key, &hmac_key_iv, &hmac_key_blob);
    let hmac_val_iv = derive_iv(&key_data_salt, &HMAC_VALUE_BLOCK, block_size, hash_alg).unwrap();
    let encrypted_hmac_value =
        encrypt_aes128_cbc_no_padding(&package_key, &hmac_val_iv, &hmac_value_blob);

    // --- Build EncryptionInfo stream -----------------------------------------------------------
    let xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
            xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
  <keyData saltSize="16" blockSize="{block_size}" keyBits="128" hashSize="{hash_size}"
           cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA1"
           saltValue="{key_data_salt_b64}"/>
  <dataIntegrity encryptedHmacKey="{ehk_b64}" encryptedHmacValue="{ehv_b64}"/>
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey saltSize="16" blockSize="{block_size}" keyBits="128" hashSize="{hash_size}"
                      spinCount="{spin_count}" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA1"
                      saltValue="{password_salt_b64}"
                      encryptedVerifierHashInput="{evhi_b64}"
                      encryptedVerifierHashValue="{evhv_b64}"
                      encryptedKeyValue="{ekv_b64}"/>
    </keyEncryptor>
  </keyEncryptors>
</encryption>"#,
        key_data_salt_b64 = BASE64.encode(&key_data_salt),
        password_salt_b64 = BASE64.encode(&password_salt),
        ehk_b64 = BASE64.encode(&encrypted_hmac_key),
        ehv_b64 = BASE64.encode(&encrypted_hmac_value),
        evhi_b64 = BASE64.encode(&encrypted_verifier_hash_input),
        evhv_b64 = BASE64.encode(&encrypted_verifier_hash_value),
        ekv_b64 = BASE64.encode(&encrypted_key_value),
    );

    let mut encryption_info = Vec::new();
    encryption_info.extend_from_slice(&4u16.to_le_bytes()); // major
    encryption_info.extend_from_slice(&4u16.to_le_bytes()); // minor
    encryption_info.extend_from_slice(&0u32.to_le_bytes()); // flags
    encryption_info.extend_from_slice(xml.as_bytes());

    let decrypted =
        decrypt_agile_encrypted_package(&encryption_info, &encrypted_package, password).unwrap();
    assert_eq!(decrypted, plain_zip);

    let err = decrypt_agile_encrypted_package(&encryption_info, &encrypted_package, "wrong-password")
        .expect_err("wrong password should fail");
    assert!(
        matches!(err, OffCryptoError::WrongPassword),
        "expected WrongPassword, got: {err:?}"
    );
}

#[test]
fn agile_decrypt_warns_on_multiple_password_key_encryptors() {
    let password = "correct horse battery staple";
    let plain_zip = build_tiny_zip();
    let encrypted_cfb = encrypt_zip_with_password(&plain_zip, password);

    let mut encryption_info = extract_stream_bytes(&encrypted_cfb, "/EncryptionInfo");
    let encrypted_package = extract_stream_bytes(&encrypted_cfb, "/EncryptedPackage");

    // Duplicate the existing password `<keyEncryptor>` block in the XML.
    let xml_start = encryption_info
        .iter()
        .position(|b| *b == b'<')
        .expect("XML must be present");
    let header = encryption_info[..xml_start].to_vec();
    let xml = std::str::from_utf8(&encryption_info[xml_start..]).expect("EncryptionInfo XML is UTF-8");

    let marker = "<keyEncryptor uri=\"http://schemas.microsoft.com/office/2006/keyEncryptor/password\"";
    let start = xml
        .find(marker)
        .expect("expected password keyEncryptor");
    let end_rel = xml[start..]
        .find("</keyEncryptor>")
        .expect("expected closing keyEncryptor tag");
    let end = start + end_rel + "</keyEncryptor>".len();
    let key_encryptor_block = &xml[start..end];

    let insert_pos = xml
        .rfind("</keyEncryptors>")
        .expect("expected </keyEncryptors>");
    let mut patched_xml = String::new();
    patched_xml.push_str(&xml[..insert_pos]);
    patched_xml.push_str(key_encryptor_block);
    patched_xml.push_str(&xml[insert_pos..]);

    encryption_info = header.into_iter().chain(patched_xml.into_bytes()).collect();

    let (decrypted, warnings) =
        decrypt_agile_encrypted_package_with_warnings(&encryption_info, &encrypted_package, password)
            .expect("decrypt should succeed");
    assert_eq!(decrypted, plain_zip);
    assert!(
        warnings.contains(&OffCryptoWarning::MultiplePasswordKeyEncryptors { count: 2 }),
        "expected MultiplePasswordKeyEncryptors warning, got: {warnings:?}"
    );
}

#[test]
fn agile_decrypt_succeeds_without_data_integrity_and_warns() {
    // Some non-Excel producers omit the `<dataIntegrity>` element. Formula should still be able to
    // decrypt, but must skip integrity verification and surface a warning through the warnings API.
    let encrypted_cfb = std::fs::read(fixture_path("agile.xlsx")).expect("read agile.xlsx");
    let expected = std::fs::read(fixture_path("plaintext.xlsx")).expect("read plaintext.xlsx");

    let mut encryption_info = extract_stream_bytes(&encrypted_cfb, "/EncryptionInfo");
    let encrypted_package = extract_stream_bytes(&encrypted_cfb, "/EncryptedPackage");

    // Remove the `<dataIntegrity .../>` element from the XML.
    let xml_start = encryption_info
        .iter()
        .position(|b| *b == b'<')
        .expect("EncryptionInfo must contain XML");
    let header = encryption_info[..xml_start].to_vec();
    let xml = std::str::from_utf8(&encryption_info[xml_start..]).expect("EncryptionInfo XML is UTF-8");

    let start = xml
        .find("<dataIntegrity")
        .expect("expected <dataIntegrity> element");
    let end = if let Some(end_rel) = xml[start..].find("/>") {
        start + end_rel + 2
    } else if let Some(end_rel) = xml[start..].find("</dataIntegrity>") {
        start + end_rel + "</dataIntegrity>".len()
    } else {
        panic!("expected </dataIntegrity> or />");
    };

    let mut patched_xml = String::new();
    patched_xml.push_str(&xml[..start]);
    patched_xml.push_str(&xml[end..]);

    encryption_info = header.into_iter().chain(patched_xml.into_bytes()).collect();

    let decrypted = decrypt_agile_encrypted_package(&encryption_info, &encrypted_package, "password")
        .expect("decrypt without dataIntegrity should succeed");
    assert_eq!(decrypted, expected);

    let (decrypted, warnings) = decrypt_agile_encrypted_package_with_warnings(
        &encryption_info,
        &encrypted_package,
        "password",
    )
    .expect("decrypt with warnings should succeed");
    assert_eq!(decrypted, expected);
    assert!(
        warnings.contains(&OffCryptoWarning::MissingDataIntegrity),
        "expected MissingDataIntegrity warning, got: {warnings:?}"
    );
}

#[test]
fn agile_decrypt_matches_office_crypto_reference() {
    let password = "pass";
    let plain_zip = build_zip_with_padding();

    let encrypted_cfb = encrypt_zip_with_password(&plain_zip, password);
    let encryption_info = extract_stream_bytes(&encrypted_cfb, "/EncryptionInfo");
    let encrypted_package = extract_stream_bytes(&encrypted_cfb, "/EncryptedPackage");

    let decrypted =
        decrypt_agile_encrypted_package(&encryption_info, &encrypted_package, password).unwrap();
    assert_eq!(decrypted, plain_zip);

    // Sanity-check the higher-level wrapper that detects encryption and extracts streams from the
    // OLE container.
    let mut ole = CompoundFile::open(Cursor::new(encrypted_cfb.as_slice())).expect("open cfb");
    let decrypted_from_cfb =
        formula_xlsx::decrypt_ooxml_from_cfb(&mut ole, password).expect("decrypt from cfb");
    assert_eq!(decrypted_from_cfb, plain_zip);
    assert_eq!(decrypted_from_cfb, decrypted);

    let decrypted_from_ole_bytes =
        formula_xlsx::decrypt_ooxml_from_ole_bytes(encrypted_cfb.as_slice(), password)
            .expect("decrypt from ole bytes");
    assert_eq!(decrypted_from_ole_bytes, plain_zip);
    assert_eq!(decrypted_from_ole_bytes, decrypted);

    let decrypted_from_ole_reader =
        formula_xlsx::decrypt_ooxml_from_ole_reader(Cursor::new(encrypted_cfb.as_slice()), password)
            .expect("decrypt from ole reader");
    assert_eq!(decrypted_from_ole_reader, plain_zip);
    assert_eq!(decrypted_from_ole_reader, decrypted);

    let office_crypto_decrypted =
        office_crypto::decrypt_from_bytes(encrypted_cfb, password).unwrap();
    assert_eq!(office_crypto_decrypted, plain_zip);
    assert_eq!(office_crypto_decrypted, decrypted);
}

#[test]
fn agile_decrypt_empty_password_matches_office_crypto_reference() {
    let password = "";
    let plain_zip = build_zip_with_padding();

    let encrypted_cfb = encrypt_zip_with_password(&plain_zip, password);
    let encryption_info = extract_stream_bytes(&encrypted_cfb, "/EncryptionInfo");
    let encrypted_package = extract_stream_bytes(&encrypted_cfb, "/EncryptedPackage");

    let decrypted =
        decrypt_agile_encrypted_package(&encryption_info, &encrypted_package, password).unwrap();
    assert_eq!(decrypted, plain_zip);

    let mut ole = CompoundFile::open(Cursor::new(encrypted_cfb.as_slice())).expect("open cfb");
    let decrypted_from_cfb =
        formula_xlsx::decrypt_ooxml_from_cfb(&mut ole, password).expect("decrypt from cfb");
    assert_eq!(decrypted_from_cfb, plain_zip);
    assert_eq!(decrypted_from_cfb, decrypted);

    let decrypted_from_ole_bytes =
        formula_xlsx::decrypt_ooxml_from_ole_bytes(encrypted_cfb.as_slice(), password)
            .expect("decrypt from ole bytes");
    assert_eq!(decrypted_from_ole_bytes, plain_zip);
    assert_eq!(decrypted_from_ole_bytes, decrypted);

    let decrypted_from_ole_reader =
        formula_xlsx::decrypt_ooxml_from_ole_reader(Cursor::new(encrypted_cfb.as_slice()), password)
            .expect("decrypt from ole reader");
    assert_eq!(decrypted_from_ole_reader, plain_zip);
    assert_eq!(decrypted_from_ole_reader, decrypted);

    let office_crypto_decrypted = office_crypto::decrypt_from_bytes(encrypted_cfb, password).unwrap();
    assert_eq!(office_crypto_decrypted, plain_zip);
    assert_eq!(office_crypto_decrypted, decrypted);
}

#[test]
fn agile_decrypt_large_fixture_matches_office_crypto_reference() {
    // Cross-check our Agile decrypt against the independent `office-crypto` implementation on a
    // real (pre-generated) encrypted workbook fixture.
    let password = "password";
    let encrypted_cfb =
        std::fs::read(fixture_path("agile-large.xlsx")).expect("read agile-large.xlsx");
    let expected =
        std::fs::read(fixture_path("plaintext-large.xlsx")).expect("read plaintext-large.xlsx");

    let encryption_info = extract_stream_bytes(&encrypted_cfb, "/EncryptionInfo");
    let encrypted_package = extract_stream_bytes(&encrypted_cfb, "/EncryptedPackage");

    let decrypted =
        decrypt_agile_encrypted_package(&encryption_info, &encrypted_package, password).unwrap();
    assert_eq!(decrypted, expected);

    let office_crypto_decrypted = office_crypto::decrypt_from_bytes(encrypted_cfb, password).unwrap();
    assert_eq!(office_crypto_decrypted, expected);
    assert_eq!(office_crypto_decrypted, decrypted);
}

#[test]
fn agile_decrypt_unicode_excel_fixture_matches_office_crypto_reference() {
    // Cross-check against a real Excel-produced Agile-encrypted workbook with a non-BMP (emoji)
    // password, to validate password encoding + multi-segment decryption.
    let password = "pässwörd🔒";
    let encrypted_cfb = std::fs::read(fixture_path("agile-unicode-excel.xlsx"))
        .expect("read agile-unicode-excel.xlsx");
    let expected =
        std::fs::read(fixture_path("plaintext-excel.xlsx")).expect("read plaintext-excel.xlsx");

    // Sanity: ensure we cover multi-segment (4096-byte) Agile decryption for this fixture.
    assert!(
        expected.len() > 4096,
        "expected plaintext-excel.xlsx to be > 4096 bytes, got {}",
        expected.len()
    );

    let encryption_info = extract_stream_bytes(&encrypted_cfb, "/EncryptionInfo");
    let encrypted_package = extract_stream_bytes(&encrypted_cfb, "/EncryptedPackage");

    let decrypted =
        decrypt_agile_encrypted_package(&encryption_info, &encrypted_package, password).unwrap();
    assert_eq!(decrypted, expected);

    let office_crypto_decrypted = office_crypto::decrypt_from_bytes(encrypted_cfb, password).unwrap();
    assert_eq!(office_crypto_decrypted, expected);
    assert_eq!(office_crypto_decrypted, decrypted);
}
