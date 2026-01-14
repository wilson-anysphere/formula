#![cfg(not(target_arch = "wasm32"))]

use std::io::{Cursor, Read, Write};
use std::path::PathBuf;

use aes::cipher::{generic_array::GenericArray, BlockEncrypt, KeyInit};
use aes::Aes128;
use base64::engine::general_purpose::STANDARD as BASE64;
use base64::Engine as _;
use cbc::cipher::block_padding::NoPadding;
use cbc::cipher::{BlockEncryptMut, KeyIvInit};
use cfb::CompoundFile;
use hmac::{Hmac, Mac as _};
use ms_offcrypto_writer::Ecma376AgileWriter;
use rand::{rngs::StdRng, SeedableRng as _};
use sha1::{Digest as _, Sha1};
use zip::write::FileOptions;

use formula_offcrypto::{
    decrypt_encrypted_package, standard_derive_key, DecryptOptions, OffcryptoError,
    StandardEncryptionHeader, StandardEncryptionHeaderFlags, StandardEncryptionInfo,
    StandardEncryptionVerifier,
};

const CALG_AES_128: u32 = 0x0000_660E;
const CALG_SHA1: u32 = 0x0000_8004;

fn fixture(path: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join(path)
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

fn encrypt_zip_with_password_agile(plain_zip: &[u8], password: &str) -> Vec<u8> {
    let mut cursor = Cursor::new(Vec::new());
    let mut rng = StdRng::from_seed([0u8; 32]);
    let mut agile =
        Ecma376AgileWriter::create(&mut rng, password, &mut cursor).expect("create agile");
    agile
        .write_all(plain_zip)
        .expect("write plaintext zip to agile writer");
    agile.finalize().expect("finalize agile writer");
    cursor.into_inner()
}

fn extract_stream_bytes(cfb_bytes: &[u8], stream_name: &str) -> Vec<u8> {
    let mut ole = CompoundFile::open(Cursor::new(cfb_bytes)).expect("open cfb");
    let mut stream = ole.open_stream(stream_name).expect("open stream");
    let mut buf = Vec::new();
    stream.read_to_end(&mut buf).expect("read stream");
    buf
}

#[test]
fn decrypt_agile_roundtrip_matches_plain_zip() {
    let password = "correct horse battery staple";
    let plain_zip = build_tiny_zip();

    let encrypted_cfb = encrypt_zip_with_password_agile(&plain_zip, password);
    let encryption_info = extract_stream_bytes(&encrypted_cfb, "/EncryptionInfo");
    let encrypted_package = extract_stream_bytes(&encrypted_cfb, "/EncryptedPackage");

    // Ensure we validate `dataIntegrity` (HMAC over EncryptedPackage stream bytes).
    let options = DecryptOptions {
        verify_integrity: true,
        ..DecryptOptions::default()
    };
    let decrypted = decrypt_encrypted_package(
        &encryption_info,
        &encrypted_package,
        password,
        options,
    )
    .expect("decrypt agile package");
    assert_eq!(decrypted, plain_zip);
}

#[test]
fn decrypt_agile_wrong_password_is_invalid_password() {
    let plain_zip = build_tiny_zip();
    let encrypted_cfb = encrypt_zip_with_password_agile(&plain_zip, "password-1");

    let encryption_info = extract_stream_bytes(&encrypted_cfb, "/EncryptionInfo");
    let encrypted_package = extract_stream_bytes(&encrypted_cfb, "/EncryptedPackage");

    let err = decrypt_encrypted_package(
        &encryption_info,
        &encrypted_package,
        "password-2",
        DecryptOptions::default(),
    )
    .expect_err("wrong password should fail");
    assert_eq!(err, OffcryptoError::InvalidPassword);
}

#[test]
fn decrypt_agile_tampered_ciphertext_fails_integrity() {
    let password = "correct horse battery staple";
    let plain_zip = build_tiny_zip();
    let encrypted_cfb = encrypt_zip_with_password_agile(&plain_zip, password);
    let encryption_info = extract_stream_bytes(&encrypted_cfb, "/EncryptionInfo");
    let mut encrypted_package = extract_stream_bytes(&encrypted_cfb, "/EncryptedPackage");

    // Flip a byte in the ciphertext (after the 8-byte length header). Integrity verification
    // should fail before decryption is attempted.
    assert!(
        encrypted_package.len() > 8,
        "EncryptedPackage stream is unexpectedly small"
    );
    encrypted_package[8] ^= 0x55;

    let options = DecryptOptions {
        verify_integrity: true,
        ..DecryptOptions::default()
    };
    let err = decrypt_encrypted_package(
        &encryption_info,
        &encrypted_package,
        password,
        options,
    )
    .expect_err("tampered EncryptedPackage should fail integrity");
    assert_eq!(err, OffcryptoError::IntegrityCheckFailed);
}

#[test]
fn decrypt_agile_tampered_size_header_fails_integrity() {
    let password = "correct horse battery staple";
    let plain_zip = build_tiny_zip();
    let encrypted_cfb = encrypt_zip_with_password_agile(&plain_zip, password);
    let encryption_info = extract_stream_bytes(&encrypted_cfb, "/EncryptionInfo");
    let mut encrypted_package = extract_stream_bytes(&encrypted_cfb, "/EncryptedPackage");

    // Tamper the 8-byte plaintext size prefix. Integrity verification should cover these bytes as
    // part of the full `EncryptedPackage` stream.
    let original_size = u64::from_le_bytes(
        encrypted_package[..8]
            .try_into()
            .expect("EncryptedPackage header is 8 bytes"),
    );
    assert!(original_size > 0, "unexpected empty EncryptedPackage payload");
    let tampered_size = original_size - 1;
    encrypted_package[..8].copy_from_slice(&tampered_size.to_le_bytes());

    let options = DecryptOptions {
        verify_integrity: true,
        ..DecryptOptions::default()
    };
    let err = decrypt_encrypted_package(
        &encryption_info,
        &encrypted_package,
        password,
        options,
    )
    .expect_err("tampered EncryptedPackage header should fail integrity");
    assert_eq!(err, OffcryptoError::IntegrityCheckFailed);
}

#[test]
fn decrypt_agile_appended_ciphertext_fails_integrity() {
    let password = "correct horse battery staple";
    let plain_zip = build_tiny_zip();
    let encrypted_cfb = encrypt_zip_with_password_agile(&plain_zip, password);
    let encryption_info = extract_stream_bytes(&encrypted_cfb, "/EncryptionInfo");
    let mut encrypted_package = extract_stream_bytes(&encrypted_cfb, "/EncryptedPackage");

    // Append an extra AES block to simulate trailing bytes stored in the stream (e.g. sector slack
    // or producer quirks). MS-OFFCRYPTO defines `dataIntegrity` as an HMAC over the *entire*
    // `EncryptedPackage` stream bytes, so this should fail integrity verification.
    encrypted_package.extend_from_slice(&[0xA5u8; 16]);

    let options = DecryptOptions {
        verify_integrity: true,
        ..DecryptOptions::default()
    };
    let err = decrypt_encrypted_package(
        &encryption_info,
        &encrypted_package,
        password,
        options,
    )
    .expect_err("tampered EncryptedPackage should fail integrity");
    assert_eq!(err, OffcryptoError::IntegrityCheckFailed);
}

fn password_to_utf16le_bytes(password: &str) -> Vec<u8> {
    let mut out = Vec::with_capacity(password.len() * 2);
    for unit in password.encode_utf16() {
        out.extend_from_slice(&unit.to_le_bytes());
    }
    out
}

fn derive_iterated_hash_sha1(password: &str, salt_value: &[u8], spin_count: u32) -> Vec<u8> {
    let password_utf16 = password_to_utf16le_bytes(password);
    let mut buf = Vec::with_capacity(salt_value.len() + password_utf16.len());
    buf.extend_from_slice(salt_value);
    buf.extend_from_slice(&password_utf16);

    let mut h = Sha1::digest(&buf).to_vec();
    for i in 0..spin_count {
        let mut round = Vec::with_capacity(4 + h.len());
        round.extend_from_slice(&i.to_le_bytes());
        round.extend_from_slice(&h);
        h = Sha1::digest(&round).to_vec();
    }
    h
}

fn derive_encryption_key_sha1(h: &[u8], block_key: &[u8; 8], key_bits: usize) -> Vec<u8> {
    assert_eq!(
        key_bits % 8,
        0,
        "keyBits must be divisible by 8 for AES"
    );
    let key_len = key_bits / 8;
    let mut buf = Vec::with_capacity(h.len() + block_key.len());
    buf.extend_from_slice(h);
    buf.extend_from_slice(block_key);
    let digest = Sha1::digest(&buf);
    digest[..key_len].to_vec()
}

fn derive_iv_sha1(salt: &[u8], suffix: &[u8]) -> [u8; 16] {
    let mut h = Sha1::new();
    h.update(salt);
    h.update(suffix);
    let digest = h.finalize();
    let mut iv = [0u8; 16];
    iv.copy_from_slice(&digest[..16]);
    iv
}

fn aes128_cbc_encrypt_no_padding(key: &[u8], iv: &[u8], plaintext: &[u8]) -> Vec<u8> {
    assert_eq!(key.len(), 16);
    assert_eq!(iv.len(), 16);
    assert_eq!(
        plaintext.len() % 16,
        0,
        "AES-CBC NoPadding requires block-aligned plaintext"
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
    let mut mac: Hmac<Sha1> =
        <Hmac<Sha1> as hmac::Mac>::new_from_slice(key).expect("HMAC key length is unrestricted");
    mac.update(data);
    mac.finalize().into_bytes().to_vec()
}

#[test]
fn decrypt_agile_accepts_short_hmac_key() {
    // Some producers emit a decrypted `encryptedHmacKey` whose length is shorter than the hash
    // output size (`SHA1`=20). HMAC accepts any key length, so we should accept such files as long
    // as the computed digest matches `encryptedHmacValue` (first `hash_len` bytes).
    let password = "pw";
    let plain_zip = build_tiny_zip();

    // Use SHA1 so `hash_len=20` is not AES-block aligned (16), forcing padding/extra bytes in
    // `encryptedHmacValue`.
    let hash_len = 20usize;
    let block_size = 16usize;
    let key_bits = 128usize;
    let spin_count = 10u32;

    let key_data_salt: Vec<u8> = (0u8..=15).collect();
    let password_salt: Vec<u8> = (16u8..=31).collect();
    let package_key: Vec<u8> = (32u8..=47).collect(); // AES-128 package key

    // Agile block keys (MS-OFFCRYPTO).
    const BLK_KEY_VERIFIER_HASH_INPUT: [u8; 8] =
        [0xFE, 0xA7, 0xD2, 0x76, 0x3B, 0x4B, 0x9E, 0x79];
    const BLK_KEY_VERIFIER_HASH_VALUE: [u8; 8] =
        [0xD7, 0xAA, 0x0F, 0x6D, 0x30, 0x61, 0x34, 0x4E];
    const BLK_KEY_ENCRYPTED_KEY_VALUE: [u8; 8] =
        [0x14, 0x6E, 0x0B, 0xE7, 0xAB, 0xAC, 0xD0, 0xD6];
    const BLK_KEY_HMAC_KEY: [u8; 8] = [0x5F, 0xB2, 0xAD, 0x01, 0x0C, 0xB9, 0xE1, 0xF6];
    const BLK_KEY_HMAC_VALUE: [u8; 8] = [0xA0, 0x67, 0x7F, 0x02, 0xB2, 0x2C, 0x84, 0x33];

    // --- Build EncryptedPackage stream ---------------------------------------------------------
    let iv0 = derive_iv_sha1(&key_data_salt, &0u32.to_le_bytes());
    let padded_plain_zip = {
        let mut out = plain_zip.clone();
        out.extend(std::iter::repeat(0u8).take((16 - (out.len() % 16)) % 16));
        out
    };
    let ciphertext = aes128_cbc_encrypt_no_padding(&package_key, &iv0, &padded_plain_zip);
    let mut encrypted_package = Vec::new();
    encrypted_package.extend_from_slice(&(plain_zip.len() as u64).to_le_bytes());
    encrypted_package.extend_from_slice(&ciphertext);

    // --- Build password key-encryptor fields ---------------------------------------------------
    let password_hash = derive_iterated_hash_sha1(password, &password_salt, spin_count);
    let verifier_iv = &password_salt[..block_size];

    let verifier_input: Vec<u8> = b"abcdefghijklmnop".to_vec();
    let verifier_hash: Vec<u8> = Sha1::digest(&verifier_input).to_vec();

    let mut verifier_hash_value_plain = verifier_hash.clone();
    verifier_hash_value_plain.extend_from_slice(&[0xA5u8; 12]); // garbage beyond hash_len
    assert_eq!(verifier_hash_value_plain.len(), 32);

    let encrypt_pw_blob = |block_key: &[u8; 8], plaintext: &[u8]| -> Vec<u8> {
        let k = derive_encryption_key_sha1(&password_hash, block_key, key_bits);
        aes128_cbc_encrypt_no_padding(&k, verifier_iv, plaintext)
    };

    let encrypted_verifier_hash_input =
        encrypt_pw_blob(&BLK_KEY_VERIFIER_HASH_INPUT, &verifier_input);
    let encrypted_verifier_hash_value =
        encrypt_pw_blob(&BLK_KEY_VERIFIER_HASH_VALUE, &verifier_hash_value_plain);
    let encrypted_key_value = encrypt_pw_blob(&BLK_KEY_ENCRYPTED_KEY_VALUE, &package_key);

    // --- Build dataIntegrity fields ------------------------------------------------------------
    let hmac_key_plain: Vec<u8> = vec![0x11u8; 16]; // shorter than hash_len=20
    let actual_hmac = compute_hmac_sha1(&hmac_key_plain, &encrypted_package);
    assert_eq!(actual_hmac.len(), hash_len);

    // Non-zero garbage padding after hash_len.
    let mut hmac_value_blob = actual_hmac.clone();
    hmac_value_blob.extend_from_slice(&[0xC3u8; 12]);
    assert_eq!(hmac_value_blob.len(), 32);

    let iv_hmac_key = derive_iv_sha1(&key_data_salt, &BLK_KEY_HMAC_KEY);
    let encrypted_hmac_key = aes128_cbc_encrypt_no_padding(&package_key, &iv_hmac_key, &hmac_key_plain);
    let iv_hmac_value = derive_iv_sha1(&key_data_salt, &BLK_KEY_HMAC_VALUE);
    let encrypted_hmac_value =
        aes128_cbc_encrypt_no_padding(&package_key, &iv_hmac_value, &hmac_value_blob);

    // --- Build EncryptionInfo stream -----------------------------------------------------------
    let xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
            xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
  <keyData saltSize="16" blockSize="{block_size}" keyBits="128" hashSize="{hash_len}"
           cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA1"
           saltValue="{key_data_salt_b64}"/>
  <dataIntegrity encryptedHmacKey="{ehk_b64}" encryptedHmacValue="{ehv_b64}"/>
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey saltSize="16" blockSize="{block_size}" keyBits="128" hashSize="{hash_len}"
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

    let options = DecryptOptions {
        verify_integrity: true,
        ..DecryptOptions::default()
    };
    let decrypted = decrypt_encrypted_package(&encryption_info, &encrypted_package, password, options)
        .expect("decrypt agile package with short HMAC key");
    assert_eq!(decrypted, plain_zip);
}

fn aes128_ecb_encrypt_in_place(key: &[u8], buf: &mut [u8]) {
    assert_eq!(key.len(), 16, "expected AES-128 key");
    assert_eq!(buf.len() % 16, 0, "ECB input must be block-aligned");
    let cipher = Aes128::new_from_slice(key).expect("valid AES-128 key");
    for block in buf.chunks_mut(16) {
        cipher.encrypt_block(GenericArray::from_mut_slice(block));
    }
}

fn build_standard_encryption_info_bytes(
    salt: &[u8; 16],
    encrypted_verifier: [u8; 16],
    encrypted_verifier_hash: [u8; 32],
) -> Vec<u8> {
    let mut bytes = Vec::new();
    // EncryptionVersionInfo (major=3, minor=2).
    bytes.extend_from_slice(&3u16.to_le_bytes());
    bytes.extend_from_slice(&2u16.to_le_bytes());
    bytes.extend_from_slice(&0u32.to_le_bytes()); // flags

    // EncryptionHeader (8 DWORDs, no CSPName string).
    let mut header = Vec::new();
    let header_flags =
        StandardEncryptionHeaderFlags::F_CRYPTOAPI | StandardEncryptionHeaderFlags::F_AES;
    header.extend_from_slice(&header_flags.to_le_bytes()); // header.flags
    header.extend_from_slice(&0u32.to_le_bytes()); // header.sizeExtra
    header.extend_from_slice(&CALG_AES_128.to_le_bytes()); // header.algId
    header.extend_from_slice(&CALG_SHA1.to_le_bytes()); // header.algIdHash
    header.extend_from_slice(&128u32.to_le_bytes()); // header.keySize
    header.extend_from_slice(&0u32.to_le_bytes()); // header.providerType
    header.extend_from_slice(&0u32.to_le_bytes()); // header.reserved1
    header.extend_from_slice(&0u32.to_le_bytes()); // header.reserved2

    bytes.extend_from_slice(&(header.len() as u32).to_le_bytes());
    bytes.extend_from_slice(&header);

    // EncryptionVerifier.
    bytes.extend_from_slice(&16u32.to_le_bytes()); // saltSize
    bytes.extend_from_slice(salt); // salt
    bytes.extend_from_slice(&encrypted_verifier);
    bytes.extend_from_slice(&20u32.to_le_bytes()); // verifierHashSize (SHA1)
    bytes.extend_from_slice(&encrypted_verifier_hash);

    bytes
}

fn encrypt_standard_encrypted_package_ecb(key: &[u8], plaintext: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    out.extend_from_slice(&(plaintext.len() as u64).to_le_bytes());

    let mut buf = plaintext.to_vec();
    let pad_len = (16 - (buf.len() % 16)) % 16;
    buf.extend(std::iter::repeat(0u8).take(pad_len));
    aes128_ecb_encrypt_in_place(key, &mut buf);
    out.extend_from_slice(&buf);

    out
}

#[test]
fn decrypt_standard_roundtrip_matches_plain_zip() {
    let password = "Password1234_";
    let salt: [u8; 16] = [
        0xe8, 0x82, 0x66, 0x49, 0x0c, 0x5b, 0xd1, 0xee, 0xbd, 0x2b, 0x43, 0x94, 0xe3, 0xf8,
        0x30, 0xef,
    ];

    let key = {
        let info = StandardEncryptionInfo {
            header: StandardEncryptionHeader {
                flags: StandardEncryptionHeaderFlags::from_raw(
                    StandardEncryptionHeaderFlags::F_CRYPTOAPI
                        | StandardEncryptionHeaderFlags::F_AES,
                ),
                size_extra: 0,
                alg_id: CALG_AES_128,
                alg_id_hash: CALG_SHA1,
                key_size_bits: 128,
                provider_type: 0,
                reserved1: 0,
                reserved2: 0,
                csp_name: String::new(),
            },
            verifier: StandardEncryptionVerifier {
                salt: Vec::from(salt),
                encrypted_verifier: [0u8; 16],
                verifier_hash_size: 20,
                encrypted_verifier_hash: vec![0u8; 32],
            },
        };
        standard_derive_key(&info, password).expect("derive key")
    };

    let verifier_plain: [u8; 16] = [
        0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0a, 0x0b, 0x0c, 0x0d,
        0x0e, 0x0f,
    ];
    let verifier_hash: [u8; 20] = Sha1::digest(&verifier_plain).into();
    let mut verifier_hash_padded = [0u8; 32];
    verifier_hash_padded[..20].copy_from_slice(&verifier_hash);
    verifier_hash_padded[20..].fill(0xa5);

    let mut encrypted_verifier = verifier_plain;
    aes128_ecb_encrypt_in_place(&key, &mut encrypted_verifier);
    let mut encrypted_verifier_hash = verifier_hash_padded;
    aes128_ecb_encrypt_in_place(&key, &mut encrypted_verifier_hash);

    let encryption_info = build_standard_encryption_info_bytes(&salt, encrypted_verifier, encrypted_verifier_hash);

    let plain_zip = build_tiny_zip();
    let encrypted_package = encrypt_standard_encrypted_package_ecb(&key, &plain_zip);

    let decrypted = decrypt_encrypted_package(
        &encryption_info,
        &encrypted_package,
        password,
        DecryptOptions::default(),
    )
    .expect("decrypt standard package");
    assert_eq!(decrypted, plain_zip);
}

#[test]
fn decrypt_standard_rejects_non_zip_plaintext_via_zip_check() {
    let password = "Password1234_";
    let salt: [u8; 16] = [
        0xe8, 0x82, 0x66, 0x49, 0x0c, 0x5b, 0xd1, 0xee, 0xbd, 0x2b, 0x43, 0x94, 0xe3, 0xf8,
        0x30, 0xef,
    ];

    let key = {
        let info = StandardEncryptionInfo {
            header: StandardEncryptionHeader {
                flags: StandardEncryptionHeaderFlags::from_raw(
                    StandardEncryptionHeaderFlags::F_CRYPTOAPI
                        | StandardEncryptionHeaderFlags::F_AES,
                ),
                size_extra: 0,
                alg_id: CALG_AES_128,
                alg_id_hash: CALG_SHA1,
                key_size_bits: 128,
                provider_type: 0,
                reserved1: 0,
                reserved2: 0,
                csp_name: String::new(),
            },
            verifier: StandardEncryptionVerifier {
                salt: Vec::from(salt),
                encrypted_verifier: [0u8; 16],
                verifier_hash_size: 20,
                encrypted_verifier_hash: vec![0u8; 32],
            },
        };
        standard_derive_key(&info, password).expect("derive key")
    };

    let verifier_plain: [u8; 16] = [
        0x10, 0x11, 0x12, 0x13, 0x14, 0x15, 0x16, 0x17, 0x18, 0x19, 0x1a, 0x1b, 0x1c, 0x1d,
        0x1e, 0x1f,
    ];
    let verifier_hash: [u8; 20] = Sha1::digest(&verifier_plain).into();
    let mut verifier_hash_padded = [0u8; 32];
    verifier_hash_padded[..20].copy_from_slice(&verifier_hash);

    let mut encrypted_verifier = verifier_plain;
    aes128_ecb_encrypt_in_place(&key, &mut encrypted_verifier);
    let mut encrypted_verifier_hash = verifier_hash_padded;
    aes128_ecb_encrypt_in_place(&key, &mut encrypted_verifier_hash);

    let encryption_info = build_standard_encryption_info_bytes(&salt, encrypted_verifier, encrypted_verifier_hash);

    let plaintext = b"this is not a zip file";
    let encrypted_package = encrypt_standard_encrypted_package_ecb(&key, plaintext);

    let err = decrypt_encrypted_package(
        &encryption_info,
        &encrypted_package,
        password,
        DecryptOptions::default(),
    )
    .expect_err("expected non-zip plaintext to be rejected");

    assert_eq!(err, OffcryptoError::InvalidPassword);
}

#[test]
fn decrypt_standard_fixture_encrypted_package_stream_matches_expected_plaintext() {
    let encrypted = std::fs::read(fixture("inputs/ecma376standard_password.docx"))
        .expect("read encrypted fixture");
    let expected = std::fs::read(fixture("outputs/ecma376standard_password_plain.docx"))
        .expect("read expected decrypted fixture");

    let encryption_info = extract_stream_bytes(&encrypted, "EncryptionInfo");
    let encrypted_package = extract_stream_bytes(&encrypted, "EncryptedPackage");

    let decrypted = decrypt_encrypted_package(
        &encryption_info,
        &encrypted_package,
        "Password1234_",
        DecryptOptions::default(),
    )
    .expect("decrypt standard EncryptedPackage stream");
    assert_eq!(decrypted, expected);
}

#[test]
fn decrypt_agile_fixture_encrypted_package_stream_matches_expected_plaintext() {
    let encrypted = std::fs::read(fixture("inputs/example_password.xlsx"))
        .expect("read encrypted fixture");
    let expected = std::fs::read(fixture("outputs/example.xlsx")).expect("read expected fixture");

    let encryption_info = extract_stream_bytes(&encrypted, "EncryptionInfo");
    let encrypted_package = extract_stream_bytes(&encrypted, "EncryptedPackage");

    let decrypted = decrypt_encrypted_package(
        &encryption_info,
        &encrypted_package,
        "Password1234_",
        DecryptOptions::default(),
    )
    .expect("decrypt agile EncryptedPackage stream");
    assert_eq!(decrypted, expected);
}
