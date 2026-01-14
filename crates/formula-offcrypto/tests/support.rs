#![allow(dead_code)]

use std::io::{Cursor, Write};

use aes::cipher::{generic_array::GenericArray, BlockEncrypt, KeyInit};
use aes::{Aes128, Aes192, Aes256};
use base64::engine::general_purpose::STANDARD as BASE64;
use base64::Engine;
use cbc::Encryptor;
use cipher::block_padding::NoPadding;
use cipher::{BlockEncryptMut, KeyIvInit};
use sha1::{Digest as _, Sha1};

use formula_offcrypto::{
    StandardEncryptionHeader, StandardEncryptionHeaderFlags, StandardEncryptionInfo,
    StandardEncryptionVerifier,
};

const ENCRYPTED_PACKAGE_SEGMENT_LEN: usize = 4096;

// CryptoAPI algorithm identifiers used by Standard encryption.
const CALG_AES_128: u32 = 0x0000_660E;
const CALG_SHA1: u32 = 0x0000_8004;

const PASSWORD_KEY_ENCRYPTOR_NS: &str =
    "http://schemas.microsoft.com/office/2006/keyEncryptor/password";

// Agile block keys (MS-OFFCRYPTO).
const BLK_KEY_VERIFIER_HASH_INPUT: [u8; 8] = [0xFE, 0xA7, 0xD2, 0x76, 0x3B, 0x4B, 0x9E, 0x79];
const BLK_KEY_VERIFIER_HASH_VALUE: [u8; 8] = [0xD7, 0xAA, 0x0F, 0x6D, 0x30, 0x61, 0x34, 0x4E];
const BLK_KEY_ENCRYPTED_KEY_VALUE: [u8; 8] = [0x14, 0x6E, 0x0B, 0xE7, 0xAB, 0xAC, 0xD0, 0xD6];

fn sha1(data: &[u8]) -> [u8; 20] {
    Sha1::digest(data).into()
}

fn aes_ecb_encrypt_in_place(key: &[u8], buf: &mut [u8]) {
    assert_eq!(buf.len() % 16, 0, "AES-ECB requires full blocks");

    fn encrypt_with<C>(key: &[u8], buf: &mut [u8])
    where
        C: BlockEncrypt + KeyInit,
    {
        let cipher = C::new_from_slice(key).expect("valid AES key length");
        for block in buf.chunks_mut(16) {
            cipher.encrypt_block(GenericArray::from_mut_slice(block));
        }
    }

    match key.len() {
        16 => encrypt_with::<Aes128>(key, buf),
        24 => encrypt_with::<Aes192>(key, buf),
        32 => encrypt_with::<Aes256>(key, buf),
        other => panic!("invalid AES key length {other}"),
    }
}

fn aes128_cbc_encrypt(key: &[u8], iv: &[u8; 16], plaintext: &[u8]) -> Vec<u8> {
    assert_eq!(
        plaintext.len() % 16,
        0,
        "AES-CBC NoPadding requires block-aligned plaintext"
    );
    let mut buf = plaintext.to_vec();
    let enc = Encryptor::<Aes128>::new_from_slices(key, iv).expect("key/iv");
    let len = buf.len();
    enc.encrypt_padded_mut::<NoPadding>(&mut buf, len)
        .expect("encrypt");
    buf
}

fn derive_iv_sha1(salt: &[u8], block_index: u32) -> [u8; 16] {
    let mut h = Sha1::new();
    h.update(salt);
    h.update(&block_index.to_le_bytes());
    let digest = h.finalize();
    let mut iv = [0u8; 16];
    iv.copy_from_slice(&digest[..16]);
    iv
}

fn pad16_zero(bytes: &[u8]) -> Vec<u8> {
    let padded_len = (bytes.len() + 15) / 16 * 16;
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

/// Deterministically encrypt a plaintext OOXML ZIP using Standard (CryptoAPI) encryption.
///
/// Returns `(EncryptionInfo, EncryptedPackage)` stream contents.
pub fn encrypt_standard(
    plaintext_zip: &[u8],
    password: &str,
) -> (Vec<u8>, Vec<u8>) {
    // Deterministic parameters (not intended to be secure).
    let salt: [u8; 16] = [
        0x10, 0x11, 0x12, 0x13, 0x14, 0x15, 0x16, 0x17, 0x18, 0x19, 0x1A, 0x1B, 0x1C, 0x1D,
        0x1E, 0x1F,
    ];
    let verifier_plain: [u8; 16] = *b"formula-std-test";

    let header = StandardEncryptionHeader {
        flags: StandardEncryptionHeaderFlags::from_raw(
            StandardEncryptionHeaderFlags::F_CRYPTOAPI | StandardEncryptionHeaderFlags::F_AES,
        ),
        size_extra: 0,
        alg_id: CALG_AES_128,
        alg_id_hash: CALG_SHA1,
        key_size_bits: 128,
        provider_type: 0,
        reserved1: 0,
        reserved2: 0,
        csp_name: String::new(),
    };

    // Derive the key using production code (depends only on salt + password + key_size_bits).
    let placeholder = StandardEncryptionInfo {
        header: header.clone(),
        verifier: StandardEncryptionVerifier {
            salt: salt.to_vec(),
            encrypted_verifier: [0u8; 16],
            verifier_hash_size: 20,
            encrypted_verifier_hash: vec![0u8; 32],
        },
    };
    let key = formula_offcrypto::standard_derive_key(&placeholder, password).expect("derive key");

    // Encrypt verifier fields using AES-ECB (what `standard_verify_key` expects).
    let mut encrypted_verifier = verifier_plain;
    aes_ecb_encrypt_in_place(&key, &mut encrypted_verifier);

    let verifier_hash = sha1(&verifier_plain);
    let mut verifier_hash_padded = verifier_hash.to_vec();
    verifier_hash_padded.resize(32, 0);
    aes_ecb_encrypt_in_place(&key, &mut verifier_hash_padded);

    // Build `EncryptionInfo` binary payload (version 3.2).
    let mut header_bytes = Vec::new();
    header_bytes.extend_from_slice(
        &(StandardEncryptionHeaderFlags::F_CRYPTOAPI | StandardEncryptionHeaderFlags::F_AES)
            .to_le_bytes(),
    ); // flags
    header_bytes.extend_from_slice(&0u32.to_le_bytes()); // sizeExtra
    header_bytes.extend_from_slice(&CALG_AES_128.to_le_bytes());
    header_bytes.extend_from_slice(&CALG_SHA1.to_le_bytes());
    header_bytes.extend_from_slice(&128u32.to_le_bytes()); // keySize
    header_bytes.extend_from_slice(&0u32.to_le_bytes()); // providerType
    header_bytes.extend_from_slice(&0u32.to_le_bytes()); // reserved1
    header_bytes.extend_from_slice(&0u32.to_le_bytes()); // reserved2
    header_bytes.extend_from_slice(&[0u8, 0u8]); // empty UTF-16LE NUL-terminated CSPName

    let header_size = header_bytes.len() as u32;

    let mut encryption_info = Vec::new();
    encryption_info.extend_from_slice(&3u16.to_le_bytes()); // major
    encryption_info.extend_from_slice(&2u16.to_le_bytes()); // minor
    encryption_info.extend_from_slice(&0u32.to_le_bytes()); // flags
    encryption_info.extend_from_slice(&header_size.to_le_bytes());
    encryption_info.extend_from_slice(&header_bytes);

    encryption_info.extend_from_slice(&16u32.to_le_bytes()); // saltSize
    encryption_info.extend_from_slice(&salt);
    encryption_info.extend_from_slice(&encrypted_verifier);
    encryption_info.extend_from_slice(&20u32.to_le_bytes()); // verifierHashSize
    encryption_info.extend_from_slice(&verifier_hash_padded);

    // Build `EncryptedPackage`: u64 original size + CBC-encrypted 4096-byte segments.
    let mut encrypted_package = Vec::new();
    encrypted_package.extend_from_slice(&(plaintext_zip.len() as u64).to_le_bytes());

    for (idx, chunk) in plaintext_zip.chunks(ENCRYPTED_PACKAGE_SEGMENT_LEN).enumerate() {
        let block_index = idx as u32;
        let padded = pad16_zero(chunk);
        let iv = derive_iv_sha1(&salt, block_index);
        let ct = aes128_cbc_encrypt(&key, &iv, &padded);
        encrypted_package.extend_from_slice(&ct);
    }

    (encryption_info, encrypted_package)
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
        plaintext.len() % 16,
        0,
        "AES-CBC NoPadding requires block-aligned plaintext"
    );

    let mut buf = plaintext.to_vec();
    let enc = Encryptor::<Aes128>::new_from_slices(key, iv).expect("key/iv");
    let len = buf.len();
    enc.encrypt_padded_mut::<NoPadding>(&mut buf, len)
        .expect("encrypt");
    buf
}

/// Deterministically encrypt a plaintext OOXML ZIP using Agile encryption.
///
/// Returns `(EncryptionInfo, EncryptedPackage)` stream contents.
pub fn encrypt_agile(
    plaintext_zip: &[u8],
    password: &str,
) -> (Vec<u8>, Vec<u8>) {
    // Deterministic parameters (not intended to be secure).
    let spin_count: u32 = 1000;
    let key_bits: usize = 128;

    let password_salt: [u8; 16] = [
        0xA0, 0xA1, 0xA2, 0xA3, 0xA4, 0xA5, 0xA6, 0xA7, 0xA8, 0xA9, 0xAA, 0xAB, 0xAC, 0xAD,
        0xAE, 0xAF,
    ];
    let key_data_salt: [u8; 16] = [
        0xB0, 0xB1, 0xB2, 0xB3, 0xB4, 0xB5, 0xB6, 0xB7, 0xB8, 0xB9, 0xBA, 0xBB, 0xBC, 0xBD,
        0xBE, 0xBF,
    ];

    let secret_key_plain: [u8; 16] = [
        0x11, 0x11, 0x11, 0x11, 0x11, 0x11, 0x11, 0x11, 0x11, 0x11, 0x11, 0x11, 0x11, 0x11,
        0x11, 0x11,
    ];
    let verifier_hash_input_plain: [u8; 16] = *b"formula-agl-test";

    // Encrypt password verifier fields and the package key.
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

    // Build XML descriptor (version 4.4). Use attributes to match the crate's parser.
    let password_salt_b64 = BASE64.encode(password_salt);
    let key_data_salt_b64 = BASE64.encode(key_data_salt);
    let encrypted_key_value_b64 = BASE64.encode(encrypted_key_value);
    let encrypted_vhi_b64 = BASE64.encode(encrypted_verifier_hash_input);
    let encrypted_vhv_b64 = BASE64.encode(encrypted_verifier_hash_value);

    // Dummy integrity fields (not validated by this crate; must be valid base64).
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

    let mut encryption_info = Vec::new();
    encryption_info.extend_from_slice(&4u16.to_le_bytes()); // major
    encryption_info.extend_from_slice(&4u16.to_le_bytes()); // minor
    encryption_info.extend_from_slice(&0u32.to_le_bytes()); // flags
    encryption_info.extend_from_slice(xml.as_bytes());

    // Encrypt the package data in 4096-byte segments using the "secret key" and per-block IVs.
    let mut encrypted_package = Vec::new();
    encrypted_package.extend_from_slice(&(plaintext_zip.len() as u64).to_le_bytes());

    for (idx, chunk) in plaintext_zip.chunks(ENCRYPTED_PACKAGE_SEGMENT_LEN).enumerate() {
        let block_index = idx as u32;
        let padded = pad16_zero(chunk);
        let iv = derive_iv_sha1(&key_data_salt, block_index);
        let ct = aes128_cbc_encrypt(&secret_key_plain, &iv, &padded);
        encrypted_package.extend_from_slice(&ct);
    }

    (encryption_info, encrypted_package)
}

/// Wrap `EncryptionInfo` and `EncryptedPackage` streams in an OLE/CFB container.
pub fn wrap_in_ole_cfb(encryption_info: &[u8], encrypted_package: &[u8]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");

    ole.create_stream("EncryptionInfo")
        .expect("create EncryptionInfo stream")
        .write_all(encryption_info)
        .expect("write EncryptionInfo stream");

    ole.create_stream("EncryptedPackage")
        .expect("create EncryptedPackage stream")
        .write_all(encrypted_package)
        .expect("write EncryptedPackage stream");

    ole.into_inner().into_inner()
}
