use std::io::{Cursor, Write as _};

use aes::Aes128;
use base64::engine::general_purpose::STANDARD as BASE64;
use base64::Engine as _;
use cbc::cipher::{block_padding::NoPadding, BlockEncryptMut, KeyIvInit};
use sha1::{Digest as _, Sha1};

const ENCRYPTED_PACKAGE_SEGMENT_LEN: usize = 0x1000;
const AES_BLOCK_LEN: usize = 16;

const PASSWORD_KEY_ENCRYPTOR_NS: &str =
    "http://schemas.microsoft.com/office/2006/keyEncryptor/password";

// MS-OFFCRYPTO Agile block keys.
const BLK_KEY_VERIFIER_HASH_INPUT: [u8; 8] = [0xFE, 0xA7, 0xD2, 0x76, 0x3B, 0x4B, 0x9E, 0x79];
const BLK_KEY_VERIFIER_HASH_VALUE: [u8; 8] = [0xD7, 0xAA, 0x0F, 0x6D, 0x30, 0x61, 0x34, 0x4E];
const BLK_KEY_ENCRYPTED_KEY_VALUE: [u8; 8] = [0x14, 0x6E, 0x0B, 0xE7, 0xAB, 0xAC, 0xD0, 0xD6];

fn sha1(data: &[u8]) -> [u8; 20] {
    Sha1::digest(data).into()
}

fn derive_segment_iv_sha1(salt: &[u8], segment_index: u32) -> [u8; AES_BLOCK_LEN] {
    let mut hasher = Sha1::new();
    hasher.update(salt);
    hasher.update(segment_index.to_le_bytes());
    let digest = hasher.finalize();

    let mut iv = [0u8; AES_BLOCK_LEN];
    iv.copy_from_slice(&digest[..AES_BLOCK_LEN]);
    iv
}

fn pad16_zero(bytes: &[u8]) -> Vec<u8> {
    let padded_len = (bytes.len() + (AES_BLOCK_LEN - 1)) / AES_BLOCK_LEN * AES_BLOCK_LEN;
    let mut out = bytes.to_vec();
    out.resize(padded_len, 0);
    out
}

fn password_to_utf16le_bytes(password: &str) -> Vec<u8> {
    let mut out = Vec::with_capacity(password.len() * 2);
    for unit in password.encode_utf16() {
        out.extend_from_slice(&unit.to_le_bytes());
    }
    out
}

fn derive_iterated_hash_sha1(password: &str, salt_value: &[u8], spin_count: u32) -> Vec<u8> {
    // H = sha1(salt + password_utf16le)
    let password_utf16le = password_to_utf16le_bytes(password);
    let mut buf = Vec::with_capacity(salt_value.len() + password_utf16le.len());
    buf.extend_from_slice(salt_value);
    buf.extend_from_slice(&password_utf16le);
    let mut h = Sha1::digest(&buf).to_vec();

    // For i in 0..spinCount: H = sha1(u32le(i) + H)
    for i in 0..spin_count {
        let mut round = Vec::with_capacity(4 + h.len());
        round.extend_from_slice(&i.to_le_bytes());
        round.extend_from_slice(&h);
        h = Sha1::digest(&round).to_vec();
    }

    h
}

fn derive_agile_encryption_key_sha1(h: &[u8], block_key: &[u8; 8], key_bits: usize) -> Vec<u8> {
    assert_eq!(key_bits % 8, 0);
    let key_len = key_bits / 8;
    let mut buf = Vec::with_capacity(h.len() + block_key.len());
    buf.extend_from_slice(h);
    buf.extend_from_slice(block_key);
    let digest = Sha1::digest(&buf);
    digest[..key_len].to_vec()
}

fn aes128_cbc_encrypt_no_padding(key: &[u8], iv: &[u8], plaintext: &[u8]) -> Vec<u8> {
    assert_eq!(key.len(), 16);
    assert_eq!(iv.len(), 16);
    assert_eq!(
        plaintext.len() % AES_BLOCK_LEN,
        0,
        "AES-CBC NoPadding requires block-aligned plaintext"
    );

    let mut buf = plaintext.to_vec();
    cbc::Encryptor::<Aes128>::new_from_slices(key, iv)
        .expect("key/iv")
        .encrypt_padded_mut::<NoPadding>(&mut buf, plaintext.len())
        .expect("encrypt");
    buf
}

/// Create bytes for an Agile (ECMA-376) MS-OFFCRYPTO encrypted OOXML OLE container that wraps the
/// provided plaintext package (ZIP) bytes.
///
/// This is intended for CLI integration tests, so it uses deterministic crypto parameters (not
/// intended to be secure).
pub fn build_agile_encrypted_ooxml_ole_bytes(package_bytes: &[u8], password: &str) -> Vec<u8> {
    let spin_count: u32 = 1000;
    let key_bits: usize = 128;

    let password_salt: [u8; 16] = [
        0xA0, 0xA1, 0xA2, 0xA3, 0xA4, 0xA5, 0xA6, 0xA7, 0xA8, 0xA9, 0xAA, 0xAB, 0xAC, 0xAD, 0xAE,
        0xAF,
    ];
    let key_data_salt: [u8; 16] = [
        0xB0, 0xB1, 0xB2, 0xB3, 0xB4, 0xB5, 0xB6, 0xB7, 0xB8, 0xB9, 0xBA, 0xBB, 0xBC, 0xBD, 0xBE,
        0xBF,
    ];

    let secret_key_plain: [u8; 16] = [0x11u8; 16];
    let verifier_hash_input_plain: [u8; 16] = *b"formula-agl-test";

    // Password verifier fields + secret key unwrap use keys derived from the password.
    let h = derive_iterated_hash_sha1(password, &password_salt, spin_count);
    let key_kv = derive_agile_encryption_key_sha1(&h, &BLK_KEY_ENCRYPTED_KEY_VALUE, key_bits);
    let key_vhi = derive_agile_encryption_key_sha1(&h, &BLK_KEY_VERIFIER_HASH_INPUT, key_bits);
    let key_vhv = derive_agile_encryption_key_sha1(&h, &BLK_KEY_VERIFIER_HASH_VALUE, key_bits);

    let encrypted_key_value =
        aes128_cbc_encrypt_no_padding(&key_kv, &password_salt, &secret_key_plain);
    let encrypted_verifier_hash_input =
        aes128_cbc_encrypt_no_padding(&key_vhi, &password_salt, &verifier_hash_input_plain);

    let verifier_hash_value = sha1(&verifier_hash_input_plain);
    let verifier_hash_value_padded = pad16_zero(&verifier_hash_value);
    let encrypted_verifier_hash_value =
        aes128_cbc_encrypt_no_padding(&key_vhv, &password_salt, &verifier_hash_value_padded);

    // Build `EncryptionInfo` XML descriptor (version 4.4).
    let password_salt_b64 = BASE64.encode(password_salt);
    let key_data_salt_b64 = BASE64.encode(key_data_salt);
    let encrypted_key_value_b64 = BASE64.encode(encrypted_key_value);
    let encrypted_vhi_b64 = BASE64.encode(encrypted_verifier_hash_input);
    let encrypted_vhv_b64 = BASE64.encode(encrypted_verifier_hash_value);

    // Dummy integrity fields (not validated by `formula-offcrypto`).
    let dummy_integrity = BASE64.encode([0u8; 16]);

    let xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
    xmlns:p="{PASSWORD_KEY_ENCRYPTOR_NS}">
  <keyData saltValue="{key_data_salt_b64}" hashAlgorithm="SHA1" blockSize="16"/>
  <dataIntegrity encryptedHmacKey="{dummy_integrity}" encryptedHmacValue="{dummy_integrity}"/>
  <keyEncryptors>
    <keyEncryptor uri="{PASSWORD_KEY_ENCRYPTOR_NS}">
      <p:encryptedKey spinCount="{spin_count}" saltValue="{password_salt_b64}" hashAlgorithm="SHA1" keyBits="{key_bits}"
        encryptedKeyValue="{encrypted_key_value_b64}"
        encryptedVerifierHashInput="{encrypted_vhi_b64}"
        encryptedVerifierHashValue="{encrypted_vhv_b64}"/>
    </keyEncryptor>
  </keyEncryptors>
</encryption>
"#
    );

    let mut encryption_info_bytes = Vec::new();
    encryption_info_bytes.extend_from_slice(&4u16.to_le_bytes()); // major
    encryption_info_bytes.extend_from_slice(&4u16.to_le_bytes()); // minor
    encryption_info_bytes.extend_from_slice(&0u32.to_le_bytes()); // flags
    encryption_info_bytes.extend_from_slice(xml.as_bytes());

    // Encrypt the package data in 4096-byte segments.
    let mut encrypted_package_bytes = Vec::new();
    encrypted_package_bytes.extend_from_slice(&(package_bytes.len() as u64).to_le_bytes());

    for (idx, chunk) in package_bytes
        .chunks(ENCRYPTED_PACKAGE_SEGMENT_LEN)
        .enumerate()
    {
        let padded = pad16_zero(chunk);
        let iv = derive_segment_iv_sha1(&key_data_salt, idx as u32);
        let ciphertext = aes128_cbc_encrypt_no_padding(&secret_key_plain, &iv, &padded);
        encrypted_package_bytes.extend_from_slice(&ciphertext);
    }

    // Wrap the `EncryptionInfo` and `EncryptedPackage` streams in an OLE/CFB container.
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");

    {
        let mut s = ole
            .create_stream("EncryptionInfo")
            .expect("create EncryptionInfo stream");
        s.write_all(&encryption_info_bytes)
            .expect("write EncryptionInfo");
    }
    {
        let mut s = ole
            .create_stream("EncryptedPackage")
            .expect("create EncryptedPackage stream");
        s.write_all(&encrypted_package_bytes)
            .expect("write EncryptedPackage");
    }

    ole.into_inner().into_inner()
}
