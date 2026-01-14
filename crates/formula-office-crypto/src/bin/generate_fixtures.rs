//! Deterministic generator for MS-OFFCRYPTO `EncryptedPackage` fixtures used by tests.
//!
//! This binary is intentionally *not* part of the library API surface. Writing/encrypting Office
//! documents is out of scope for `formula-office-crypto`, but we still want reproducible encrypted
//! fixtures for regression tests.

use std::io::{Cursor, Write};

use aes::cipher::generic_array::GenericArray;
use aes::cipher::{BlockEncrypt, KeyInit};
use base64::Engine as _;
use hmac::{Hmac, Mac};
use sha2::{Sha512, Digest};

const FIXTURE_PASSWORD: &str = "password";
const PLAINTEXT_XLSX: &str = "fixtures/xlsx/basic/basic.xlsx";

const OUT_DIR: &str = "fixtures/encryption";
const OUT_AGILE: &str = "fixtures/encryption/encrypted_agile.xlsx";
const OUT_STANDARD: &str = "fixtures/encryption/encrypted_standard.xlsx";

const SEGMENT_LENGTH: usize = 4096;

// Agile block keys.
const BLOCK_VERIFIER_HASH_INPUT: [u8; 8] = [0xFE, 0xA7, 0xD2, 0x76, 0x3B, 0x4B, 0x9E, 0x79];
const BLOCK_VERIFIER_HASH_VALUE: [u8; 8] = [0xD7, 0xAA, 0x0F, 0x6D, 0x30, 0x61, 0x34, 0x4E];
const BLOCK_KEY_VALUE: [u8; 8] = [0x14, 0x6E, 0x0B, 0xE7, 0xAB, 0xAC, 0xD0, 0xD6];
const BLOCK_HMAC_KEY: [u8; 8] = [0x5F, 0xB2, 0xAD, 0x01, 0x0C, 0xB9, 0xE1, 0xF6];
const BLOCK_HMAC_VALUE: [u8; 8] = [0xA0, 0x67, 0x7F, 0x02, 0xB2, 0x2C, 0x84, 0x33];

fn encode_password_utf16le(password: &str) -> Vec<u8> {
    let mut out = Vec::with_capacity(password.len() * 2);
    for ch in password.encode_utf16() {
        out.extend_from_slice(&ch.to_le_bytes());
    }
    out
}

fn sha512(data: &[u8]) -> Vec<u8> {
    Sha512::digest(data).to_vec()
}

fn sha1(data: &[u8]) -> [u8; 20] {
    let digest = sha1::Sha1::digest(data);
    digest.into()
}

fn aes256_encrypt_block(key: &[u8; 32], block: &mut [u8; 16]) {
    let cipher = aes::Aes256::new(GenericArray::from_slice(key));
    cipher.encrypt_block(GenericArray::from_mut_slice(block));
}

fn aes128_encrypt_block(key: &[u8; 16], block: &mut [u8; 16]) {
    let cipher = aes::Aes128::new(GenericArray::from_slice(key));
    cipher.encrypt_block(GenericArray::from_mut_slice(block));
}

fn aes128_ecb_encrypt(key: &[u8; 16], data: &[u8]) -> Vec<u8> {
    assert!(data.len() % 16 == 0);
    let mut out = Vec::with_capacity(data.len());
    for chunk in data.chunks_exact(16) {
        let mut block = [0u8; 16];
        block.copy_from_slice(chunk);
        aes128_encrypt_block(key, &mut block);
        out.extend_from_slice(&block);
    }
    out
}

fn aes256_cbc_encrypt(key: &[u8; 32], iv: &[u8; 16], data: &[u8]) -> Vec<u8> {
    assert!(data.len() % 16 == 0);
    let mut prev = *iv;
    let mut out = Vec::with_capacity(data.len());
    for chunk in data.chunks_exact(16) {
        let mut block = [0u8; 16];
        block.copy_from_slice(chunk);
        for i in 0..16 {
            block[i] ^= prev[i];
        }
        aes256_encrypt_block(key, &mut block);
        out.extend_from_slice(&block);
        prev = block;
    }
    out
}

fn derive_iterated_hash_from_password(password: &str, salt: &[u8; 16], spin_count: u32) -> Vec<u8> {
    let pw = encode_password_utf16le(password);
    let mut buf = Vec::with_capacity(salt.len() + pw.len());
    buf.extend_from_slice(salt);
    buf.extend_from_slice(&pw);
    let mut h = sha512(&buf);
    for i in 0..spin_count {
        let mut b = Vec::with_capacity(4 + h.len());
        b.extend_from_slice(&i.to_le_bytes());
        b.extend_from_slice(&h);
        h = sha512(&b);
    }
    h
}

fn derive_agile_key(iter_hash: &[u8], block_key: &[u8; 8], key_bits: u32) -> [u8; 32] {
    assert_eq!(key_bits, 256);
    let mut buf = Vec::with_capacity(iter_hash.len() + block_key.len());
    buf.extend_from_slice(iter_hash);
    buf.extend_from_slice(block_key);
    let digest = sha512(&buf);
    let mut out = [0u8; 32];
    out.copy_from_slice(&digest[..32]);
    out
}

fn generate_agile_fixture(plaintext: &[u8]) -> Vec<u8> {
    // Deterministic parameters (SHA-512 + AES-256).
    let password_salt: [u8; 16] = *b"passwordsalt-123"; // 16 bytes
    let key_data_salt: [u8; 16] = *b"keydatasalt--456"; // 16 bytes
    let spin_count: u32 = 100_000;
    let secret_key: [u8; 32] = [
        0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D,
        0x0E, 0x0F, 0x10, 0x11, 0x12, 0x13, 0x14, 0x15, 0x16, 0x17, 0x18, 0x19, 0x1A, 0x1B,
        0x1C, 0x1D, 0x1E, 0x1F,
    ];

    // Password verifier.
    let verifier_input: [u8; 16] = *b"verifier-input!!"; // 16 bytes
    let verifier_hash = sha512(&verifier_input); // 64 bytes

    let iter_hash =
        derive_iterated_hash_from_password(FIXTURE_PASSWORD, &password_salt, spin_count);
    let key1 = derive_agile_key(&iter_hash, &BLOCK_VERIFIER_HASH_INPUT, 256);
    let key2 = derive_agile_key(&iter_hash, &BLOCK_VERIFIER_HASH_VALUE, 256);
    let key3 = derive_agile_key(&iter_hash, &BLOCK_KEY_VALUE, 256);

    let enc_verifier_input = aes256_cbc_encrypt(&key1, &password_salt, &verifier_input);
    let enc_verifier_hash = aes256_cbc_encrypt(
        &key2,
        &password_salt,
        &verifier_hash, // 64 bytes
    );
    let enc_key_value = aes256_cbc_encrypt(&key3, &password_salt, &secret_key);

    // Encrypt the payload in 4096-byte segments.
    let total_size = plaintext.len() as u64;
    let mut encrypted_package = Vec::new();
    encrypted_package.extend_from_slice(&total_size.to_le_bytes());

    let mut padded = plaintext.to_vec();
    // Pad to a full segment (Excel does this; makes fixtures deterministic).
    let remainder = padded.len() % SEGMENT_LENGTH;
    if remainder != 0 {
        padded.extend(std::iter::repeat(0u8).take(SEGMENT_LENGTH - remainder));
    }
    assert!(padded.len() % SEGMENT_LENGTH == 0);

    for (i, seg) in padded.chunks_exact(SEGMENT_LENGTH).enumerate() {
        let mut iv_buf = Vec::with_capacity(key_data_salt.len() + 4);
        iv_buf.extend_from_slice(&key_data_salt);
        iv_buf.extend_from_slice(&(i as u32).to_le_bytes());
        let iv_full = sha512(&iv_buf);
        let mut iv = [0u8; 16];
        iv.copy_from_slice(&iv_full[..16]);

        // Segment is already multiple of 16.
        let enc = aes256_cbc_encrypt(&secret_key, &iv, seg);
        encrypted_package.extend_from_slice(&enc);
    }

    // Data integrity (HMAC over the full EncryptedPackage stream).
    let hmac_key_plain = vec![0xA5u8; 64];
    let mut mac = <Hmac<Sha512> as Mac>::new_from_slice(&hmac_key_plain).unwrap();
    mac.update(&encrypted_package);
    let hmac_value_plain = mac.finalize().into_bytes().to_vec();
    assert_eq!(hmac_value_plain.len(), 64);

    let mut iv1_buf = Vec::with_capacity(key_data_salt.len() + BLOCK_HMAC_KEY.len());
    iv1_buf.extend_from_slice(&key_data_salt);
    iv1_buf.extend_from_slice(&BLOCK_HMAC_KEY);
    let iv1_full = sha512(&iv1_buf);
    let mut iv1 = [0u8; 16];
    iv1.copy_from_slice(&iv1_full[..16]);

    let mut iv2_buf = Vec::with_capacity(key_data_salt.len() + BLOCK_HMAC_VALUE.len());
    iv2_buf.extend_from_slice(&key_data_salt);
    iv2_buf.extend_from_slice(&BLOCK_HMAC_VALUE);
    let iv2_full = sha512(&iv2_buf);
    let mut iv2 = [0u8; 16];
    iv2.copy_from_slice(&iv2_full[..16]);

    let enc_hmac_key = aes256_cbc_encrypt(&secret_key, &iv1, &hmac_key_plain);
    let enc_hmac_value = aes256_cbc_encrypt(&secret_key, &iv2, &hmac_value_plain);

    // Build EncryptionInfo XML.
    let b64 = base64::engine::general_purpose::STANDARD;
    let xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password" xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">
  <keyData saltValue="{key_data_salt}" hashAlgorithm="SHA512" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" saltSize="16" blockSize="16" keyBits="256" hashSize="64"/>
  <dataIntegrity encryptedHmacKey="{enc_hmac_key}" encryptedHmacValue="{enc_hmac_value}"/>
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey spinCount="{spin_count}" saltValue="{password_salt}" hashAlgorithm="SHA512" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" saltSize="16" blockSize="16" keyBits="256" hashSize="64" encryptedVerifierHashInput="{enc_verifier_input}" encryptedVerifierHashValue="{enc_verifier_hash}" encryptedKeyValue="{enc_key_value}"/>
    </keyEncryptor>
  </keyEncryptors>
</encryption>
"#,
        key_data_salt = b64.encode(key_data_salt),
        enc_hmac_key = b64.encode(enc_hmac_key),
        enc_hmac_value = b64.encode(enc_hmac_value),
        spin_count = spin_count,
        password_salt = b64.encode(password_salt),
        enc_verifier_input = b64.encode(enc_verifier_input),
        enc_verifier_hash = b64.encode(enc_verifier_hash),
        enc_key_value = b64.encode(enc_key_value),
    );

    let mut encryption_info = Vec::new();
    encryption_info.extend_from_slice(&4u16.to_le_bytes()); // major
    encryption_info.extend_from_slice(&4u16.to_le_bytes()); // minor
    encryption_info.extend_from_slice(&0u32.to_le_bytes()); // flags (unused by our parser/tests)
    encryption_info.extend_from_slice(xml.as_bytes());

    // Build OLE container.
    build_ole(encryption_info, encrypted_package)
}

fn make_standard_key_from_password(password: &str, salt: &[u8; 16], key_size_bits: u32) -> [u8; 16] {
    assert_eq!(key_size_bits, 128);
    const ITER_COUNT: u32 = 50_000;

    let pw_utf16 = encode_password_utf16le(password);
    let mut buf = Vec::with_capacity(salt.len() + pw_utf16.len());
    buf.extend_from_slice(salt);
    buf.extend_from_slice(&pw_utf16);
    let mut h = sha1(&buf).to_vec();
    for i in 0..ITER_COUNT {
        let mut b = Vec::with_capacity(4 + h.len());
        b.extend_from_slice(&i.to_le_bytes());
        b.extend_from_slice(&h);
        h = sha1(&b).to_vec();
    }
    let mut bfinal = Vec::with_capacity(h.len() + 4);
    bfinal.extend_from_slice(&h);
    bfinal.extend_from_slice(&0u32.to_le_bytes());
    let hfinal = sha1(&bfinal);

    let mut buf1 = vec![0x36u8; 64];
    for i in 0..20 {
        buf1[i] ^= hfinal[i];
    }
    let x1 = sha1(&buf1);

    let mut buf2 = vec![0x5cu8; 64];
    for i in 0..20 {
        buf2[i] ^= hfinal[i];
    }
    let x2 = sha1(&buf2);

    let mut x3 = Vec::with_capacity(40);
    x3.extend_from_slice(&x1);
    x3.extend_from_slice(&x2);

    let mut key = [0u8; 16];
    key.copy_from_slice(&x3[..16]);
    key
}

fn generate_standard_fixture(plaintext: &[u8]) -> Vec<u8> {
    let salt: [u8; 16] = *b"standard-salt-16"; // 16 bytes
    let key = make_standard_key_from_password(FIXTURE_PASSWORD, &salt, 128);

    // Password verifier.
    let verifier: [u8; 16] = *b"std-verifier--!!"; // 16 bytes
    let verifier_hash = sha1(&verifier);
    let mut verifier_hash_padded = [0u8; 32];
    verifier_hash_padded[..20].copy_from_slice(&verifier_hash);

    let encrypted_verifier = aes128_ecb_encrypt(&key, &verifier);
    let encrypted_verifier_hash = aes128_ecb_encrypt(&key, &verifier_hash_padded);

    // Encrypted package: AES-ECB over the entire payload, padded to a multiple of 16.
    let total_size = plaintext.len() as u64;
    let mut padded = plaintext.to_vec();
    let rem = padded.len() % 16;
    if rem != 0 {
        padded.extend(std::iter::repeat(0u8).take(16 - rem));
    }
    let encrypted_payload = aes128_ecb_encrypt(&key, &padded);

    let mut encrypted_package = Vec::new();
    encrypted_package.extend_from_slice(&total_size.to_le_bytes());
    encrypted_package.extend_from_slice(&encrypted_payload);

    // EncryptionHeader (variable-length).
    let csp_name = "Microsoft Enhanced RSA and AES Cryptographic Provider\0";
    let mut csp_utf16 = Vec::new();
    for ch in csp_name.encode_utf16() {
        csp_utf16.extend_from_slice(&ch.to_le_bytes());
    }

    let mut header = Vec::new();
    header.extend_from_slice(&0u32.to_le_bytes()); // flags
    header.extend_from_slice(&0u32.to_le_bytes()); // sizeExtra
    header.extend_from_slice(&0x0000_660Eu32.to_le_bytes()); // algId AES-128
    header.extend_from_slice(&0x0000_8004u32.to_le_bytes()); // algIdHash SHA1
    header.extend_from_slice(&128u32.to_le_bytes()); // keySize bits
    header.extend_from_slice(&0x18u32.to_le_bytes()); // providerType PROV_RSA_AES
    header.extend_from_slice(&0u32.to_le_bytes()); // reserved1
    header.extend_from_slice(&0u32.to_le_bytes()); // reserved2
    header.extend_from_slice(&csp_utf16);

    let header_size = header.len() as u32;

    let mut verifier_struct = Vec::new();
    verifier_struct.extend_from_slice(&16u32.to_le_bytes()); // saltSize
    verifier_struct.extend_from_slice(&salt);
    verifier_struct.extend_from_slice(&encrypted_verifier);
    verifier_struct.extend_from_slice(&20u32.to_le_bytes()); // verifierHashSize
    verifier_struct.extend_from_slice(&encrypted_verifier_hash);

    let mut encryption_info = Vec::new();
    encryption_info.extend_from_slice(&3u16.to_le_bytes()); // major
    encryption_info.extend_from_slice(&2u16.to_le_bytes()); // minor
    encryption_info.extend_from_slice(&0u32.to_le_bytes()); // headerFlags
    encryption_info.extend_from_slice(&header_size.to_le_bytes());
    encryption_info.extend_from_slice(&header);
    encryption_info.extend_from_slice(&verifier_struct);

    build_ole(encryption_info, encrypted_package)
}

fn build_ole(encryption_info: Vec<u8>, encrypted_package: Vec<u8>) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create OLE");

    {
        let mut s = ole
            .create_stream("EncryptionInfo")
            .expect("create stream EncryptionInfo");
        s.write_all(&encryption_info)
            .expect("write EncryptionInfo");
    }
    {
        let mut s = ole
            .create_stream("EncryptedPackage")
            .expect("create stream EncryptedPackage");
        s.write_all(&encrypted_package)
            .expect("write EncryptedPackage");
    }

    ole.into_inner().into_inner()
}

fn main() {
    let plaintext = std::fs::read(PLAINTEXT_XLSX).expect("read plaintext xlsx fixture");

    std::fs::create_dir_all(OUT_DIR).expect("create output dir");

    let agile = generate_agile_fixture(&plaintext);
    std::fs::write(OUT_AGILE, &agile).expect("write agile fixture");

    let standard = generate_standard_fixture(&plaintext);
    std::fs::write(OUT_STANDARD, &standard).expect("write standard fixture");

    eprintln!("Wrote {OUT_AGILE} and {OUT_STANDARD}");

    // Basic self-check: decrypt with our own library to ensure we're not committing broken fixtures.
    let decrypted_agile =
        formula_office_crypto::decrypt_encrypted_package(&agile, FIXTURE_PASSWORD).expect("decrypt agile");
    assert_eq!(decrypted_agile, plaintext);
    let decrypted_std =
        formula_office_crypto::decrypt_encrypted_package(&standard, FIXTURE_PASSWORD).expect("decrypt standard");
    assert_eq!(decrypted_std, plaintext);
}
