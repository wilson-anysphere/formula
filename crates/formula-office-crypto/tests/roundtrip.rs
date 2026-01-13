use std::io::{Cursor, Write};

use base64::Engine;
use cipher::block_padding::NoPadding;
use cipher::{BlockEncryptMut, KeyIvInit};
use hmac::{Hmac, Mac};
use sha2::Digest;

use formula_office_crypto::{
    decrypt_encrypted_package_ole, is_encrypted_ooxml_ole, OfficeCryptoError,
};

const BLOCK_KEY_VERIFIER_HASH_INPUT: &[u8; 8] = b"\xFE\xA7\xD2\x76\x3B\x4B\x9E\x79";
const BLOCK_KEY_VERIFIER_HASH_VALUE: &[u8; 8] = b"\xD7\xAA\x0F\x6D\x30\x61\x34\x4E";
const BLOCK_KEY_ENCRYPTED_KEY_VALUE: &[u8; 8] = b"\x14\x6E\x0B\xE7\xAB\xAC\xD0\xD6";
const BLOCK_KEY_INTEGRITY_HMAC_KEY: &[u8; 8] = b"\x5F\xB2\xAD\x01\x0C\xB9\xE1\xF6";
const BLOCK_KEY_INTEGRITY_HMAC_VALUE: &[u8; 8] = b"\xA0\x67\x7F\x02\xB2\x2C\x84\x33";

#[test]
fn roundtrip_standard_encryption() {
    let password = "Password";
    let plaintext = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/basic/basic.xlsx"
    ));

    let ole_bytes = encrypt_standard_ooxml_ole(plaintext, password);
    assert!(is_encrypted_ooxml_ole(&ole_bytes));

    let decrypted = decrypt_encrypted_package_ole(&ole_bytes, password).expect("decrypt");
    assert_eq!(decrypted, plaintext);
    assert_zip_contains_workbook_xml(&decrypted);

    let err = decrypt_encrypted_package_ole(&ole_bytes, "wrong-password").expect_err("wrong pw");
    assert!(matches!(err, OfficeCryptoError::InvalidPassword));
}

#[test]
fn roundtrip_agile_encryption() {
    let password = "Password";
    let plaintext = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/basic/basic.xlsx"
    ));

    let ole_bytes = encrypt_agile_ooxml_ole(plaintext, password);
    assert!(is_encrypted_ooxml_ole(&ole_bytes));

    let decrypted = decrypt_encrypted_package_ole(&ole_bytes, password).expect("decrypt");
    assert_eq!(decrypted, plaintext);
    assert_zip_contains_workbook_xml(&decrypted);

    let err = decrypt_encrypted_package_ole(&ole_bytes, "wrong-password").expect_err("wrong pw");
    assert!(matches!(err, OfficeCryptoError::InvalidPassword));
}

fn assert_zip_contains_workbook_xml(bytes: &[u8]) {
    let cursor = Cursor::new(bytes);
    let zip = zip::ZipArchive::new(cursor).expect("zip archive");
    let mut found = false;
    for name in zip.file_names() {
        if name.eq_ignore_ascii_case("xl/workbook.xml") {
            found = true;
            break;
        }
    }
    assert!(found, "zip should contain xl/workbook.xml");
}

fn encrypt_standard_ooxml_ole(plaintext: &[u8], password: &str) -> Vec<u8> {
    // Deterministic parameters (not intended to be secure).
    let salt: [u8; 16] = [
        0x10, 0x11, 0x12, 0x13, 0x14, 0x15, 0x16, 0x17, 0x18, 0x19, 0x1A, 0x1B, 0x1C, 0x1D, 0x1E,
        0x1F,
    ];
    let key_bits = 128u32;

    // EncryptionInfo header.
    let version_major = 4u16;
    let version_minor = 2u16;
    let flags = 0x0000_0040u32;

    let header_flags = 0u32;
    let size_extra = 0u32;
    let alg_id = 0x0000_660Eu32; // CALG_AES_128
    let alg_id_hash = 0x0000_8004u32; // CALG_SHA1
    let provider_type = 0x0000_0018u32; // PROV_RSA_AES
    let reserved1 = 0u32;
    let reserved2 = 0u32;
    let csp_name_utf16_nul = [0u8, 0u8];

    let mut header_bytes = Vec::new();
    header_bytes.extend_from_slice(&header_flags.to_le_bytes());
    header_bytes.extend_from_slice(&size_extra.to_le_bytes());
    header_bytes.extend_from_slice(&alg_id.to_le_bytes());
    header_bytes.extend_from_slice(&alg_id_hash.to_le_bytes());
    header_bytes.extend_from_slice(&key_bits.to_le_bytes());
    header_bytes.extend_from_slice(&provider_type.to_le_bytes());
    header_bytes.extend_from_slice(&reserved1.to_le_bytes());
    header_bytes.extend_from_slice(&reserved2.to_le_bytes());
    header_bytes.extend_from_slice(&csp_name_utf16_nul);

    let header_size = header_bytes.len() as u32;

    // Build EncryptionVerifier.
    let verifier_plain: [u8; 16] = *b"formula-std-test";
    let verifier_hash = sha1_digest(&verifier_plain);
    let mut verifier_hash_padded = verifier_hash.to_vec();
    verifier_hash_padded.resize(32, 0);

    let key0 = standard_derive_key_sha1(password, &salt, key_bits, 0);
    let iv0 = [0u8; 16];
    let encrypted_verifier = aes128_cbc_encrypt(&key0, &iv0, &verifier_plain);
    let encrypted_verifier_hash = aes128_cbc_encrypt(&key0, &iv0, &verifier_hash_padded);

    let salt_size = salt.len() as u32;
    let verifier_hash_size = verifier_hash.len() as u32;
    let mut verifier_bytes = Vec::new();
    verifier_bytes.extend_from_slice(&salt_size.to_le_bytes());
    verifier_bytes.extend_from_slice(&salt);
    verifier_bytes.extend_from_slice(&encrypted_verifier);
    verifier_bytes.extend_from_slice(&verifier_hash_size.to_le_bytes());
    verifier_bytes.extend_from_slice(&encrypted_verifier_hash);

    let mut encryption_info = Vec::new();
    encryption_info.extend_from_slice(&version_major.to_le_bytes());
    encryption_info.extend_from_slice(&version_minor.to_le_bytes());
    encryption_info.extend_from_slice(&flags.to_le_bytes());
    encryption_info.extend_from_slice(&header_size.to_le_bytes());
    encryption_info.extend_from_slice(&header_bytes);
    encryption_info.extend_from_slice(&verifier_bytes);

    // Encrypt the package in 4096-byte segments using per-block keys (blockIndex=N, IV=0).
    let mut encrypted_package = Vec::new();
    encrypted_package.extend_from_slice(&(plaintext.len() as u64).to_le_bytes());

    const SEGMENT_LEN: usize = 4096;
    let mut offset = 0usize;
    let mut block = 0u32;
    while offset < plaintext.len() {
        let seg_len = (plaintext.len() - offset).min(SEGMENT_LEN);
        let seg = &plaintext[offset..offset + seg_len];
        let mut padded = seg.to_vec();
        let padded_len = (padded.len() + 15) / 16 * 16;
        padded.resize(padded_len, 0);

        let key = standard_derive_key_sha1(password, &salt, key_bits, block);
        let cipher = aes128_cbc_encrypt(&key, &iv0, &padded);
        encrypted_package.extend_from_slice(&cipher);

        offset += seg_len;
        block += 1;
    }

    // Write the OLE/CFB wrapper.
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    ole.create_stream("EncryptionInfo")
        .expect("create stream")
        .write_all(&encryption_info)
        .expect("write EncryptionInfo");
    ole.create_stream("EncryptedPackage")
        .expect("create stream")
        .write_all(&encrypted_package)
        .expect("write EncryptedPackage");
    ole.into_inner().into_inner()
}

fn encrypt_agile_ooxml_ole(plaintext: &[u8], password: &str) -> Vec<u8> {
    // Deterministic parameters (not intended to be secure).
    let spin_count = 10_000u32; // keep test runtime reasonable
    let block_size = 16usize;
    let key_bits = 256usize;

    let salt_key_encryptor: [u8; 16] = [
        0xA0, 0xA1, 0xA2, 0xA3, 0xA4, 0xA5, 0xA6, 0xA7, 0xA8, 0xA9, 0xAA, 0xAB, 0xAC, 0xAD, 0xAE,
        0xAF,
    ];
    let salt_key_data: [u8; 16] = [
        0xB0, 0xB1, 0xB2, 0xB3, 0xB4, 0xB5, 0xB6, 0xB7, 0xB8, 0xB9, 0xBA, 0xBB, 0xBC, 0xBD, 0xBE,
        0xBF,
    ];

    let pw_utf16 = password_to_utf16le(password);

    let verifier_hash_input_plain: [u8; 16] = *b"formula-agl-test";
    let verifier_hash_value_plain = sha512_digest(&verifier_hash_input_plain);
    let package_key_plain: [u8; 32] = [0x11; 32];

    // Encrypt password verifier fields and package key.
    let enc_vhi = agile_encrypt_with_block_key(
        &salt_key_encryptor,
        &pw_utf16,
        spin_count,
        key_bits,
        block_size,
        BLOCK_KEY_VERIFIER_HASH_INPUT,
        &verifier_hash_input_plain,
    );
    let enc_vhv = agile_encrypt_with_block_key(
        &salt_key_encryptor,
        &pw_utf16,
        spin_count,
        key_bits,
        block_size,
        BLOCK_KEY_VERIFIER_HASH_VALUE,
        &verifier_hash_value_plain,
    );
    let enc_kv = agile_encrypt_with_block_key(
        &salt_key_encryptor,
        &pw_utf16,
        spin_count,
        key_bits,
        block_size,
        BLOCK_KEY_ENCRYPTED_KEY_VALUE,
        &package_key_plain,
    );

    // Build XML descriptor.
    let b64 = base64::engine::general_purpose::STANDARD;
    let salt_key_encryptor_b64 = b64.encode(salt_key_encryptor);
    let salt_key_data_b64 = b64.encode(salt_key_data);
    let enc_vhi_b64 = b64.encode(enc_vhi);
    let enc_vhv_b64 = b64.encode(enc_vhv);
    let enc_kv_b64 = b64.encode(enc_kv);

    // Encrypt the package data in 4096-byte segments using a single package key and per-block IVs
    // derived from the keyData salt + block index.
    let mut encrypted_package = Vec::new();
    encrypted_package.extend_from_slice(&(plaintext.len() as u64).to_le_bytes());

    const SEGMENT_LEN: usize = 4096;
    let mut offset = 0usize;
    let mut block = 0u32;
    while offset < plaintext.len() {
        let seg_len = (plaintext.len() - offset).min(SEGMENT_LEN);
        let seg = &plaintext[offset..offset + seg_len];
        let mut padded = seg.to_vec();
        let padded_len = (padded.len() + 15) / 16 * 16;
        padded.resize(padded_len, 0);

        let iv = sha512_digest(&[&salt_key_data[..], &block.to_le_bytes()[..]].concat());
        let iv = &iv[..block_size];

        let cipher = aes256_cbc_encrypt(&package_key_plain, iv, &padded);
        encrypted_package.extend_from_slice(&cipher);

        offset += seg_len;
        block += 1;
    }

    // Integrity (HMAC over the EncryptedPackage stream).
    //
    // Match the crate's Agile writer: encryptedHmacKey/value are AES-CBC encrypted using the
    // package key and IVs derived from the keyData salt + fixed block keys.
    let hmac_key_plain = [0x22u8; 64];
    let hmac_value_plain = hmac_sha512(&hmac_key_plain, &encrypted_package);

    let iv_hmac_key =
        sha512_digest(&[&salt_key_data[..], &BLOCK_KEY_INTEGRITY_HMAC_KEY[..]].concat());
    let iv_hmac_key = &iv_hmac_key[..block_size];
    let encrypted_hmac_key = aes256_cbc_encrypt(&package_key_plain, iv_hmac_key, &hmac_key_plain);

    let iv_hmac_val =
        sha512_digest(&[&salt_key_data[..], &BLOCK_KEY_INTEGRITY_HMAC_VALUE[..]].concat());
    let iv_hmac_val = &iv_hmac_val[..block_size];
    let encrypted_hmac_value =
        aes256_cbc_encrypt(&package_key_plain, iv_hmac_val, &hmac_value_plain);

    let encrypted_hmac_key_b64 = b64.encode(encrypted_hmac_key);
    let encrypted_hmac_value_b64 = b64.encode(encrypted_hmac_value);

    let xml = format!(
        r#"<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption">
  <keyData saltSize="16" blockSize="16" keyBits="256" hashAlgorithm="SHA512" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" saltValue="{salt_key_data_b64}"/>
  <dataIntegrity encryptedHmacKey="{encrypted_hmac_key_b64}" encryptedHmacValue="{encrypted_hmac_value_b64}"/>
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password"
        saltSize="16" blockSize="16" keyBits="256" spinCount="{spin_count}" hashAlgorithm="SHA512" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" saltValue="{salt_key_encryptor_b64}">
        <p:encryptedVerifierHashInput>{enc_vhi_b64}</p:encryptedVerifierHashInput>
        <p:encryptedVerifierHashValue>{enc_vhv_b64}</p:encryptedVerifierHashValue>
        <p:encryptedKeyValue>{enc_kv_b64}</p:encryptedKeyValue>
      </p:encryptedKey>
    </keyEncryptor>
  </keyEncryptors>
</encryption>"#
    );

    // Build the EncryptionInfo stream.
    let version_major = 4u16;
    let version_minor = 4u16;
    let flags = 0x0000_0040u32;

    let mut encryption_info = Vec::new();
    encryption_info.extend_from_slice(&version_major.to_le_bytes());
    encryption_info.extend_from_slice(&version_minor.to_le_bytes());
    encryption_info.extend_from_slice(&flags.to_le_bytes());
    encryption_info.extend_from_slice(xml.as_bytes());

    // Write the OLE/CFB wrapper.
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    ole.create_stream("EncryptionInfo")
        .expect("create stream")
        .write_all(&encryption_info)
        .expect("write EncryptionInfo");
    ole.create_stream("EncryptedPackage")
        .expect("create stream")
        .write_all(&encrypted_package)
        .expect("write EncryptedPackage");
    ole.into_inner().into_inner()
}

fn password_to_utf16le(password: &str) -> Vec<u8> {
    let mut out = Vec::with_capacity(password.len() * 2);
    for cu in password.encode_utf16() {
        out.extend_from_slice(&cu.to_le_bytes());
    }
    out
}

fn sha1_digest(data: &[u8]) -> [u8; 20] {
    let mut hasher = sha1::Sha1::new();
    hasher.update(data);
    let out = hasher.finalize();
    out.into()
}

fn sha512_digest(data: &[u8]) -> Vec<u8> {
    let mut hasher = sha2::Sha512::new();
    hasher.update(data);
    hasher.finalize().to_vec()
}

fn hmac_sha512(key: &[u8], data: &[u8]) -> [u8; 64] {
    let mut mac: Hmac<sha2::Sha512> = Hmac::new_from_slice(key).expect("HMAC accepts any key size");
    mac.update(data);
    mac.finalize().into_bytes().into()
}

fn standard_derive_key_sha1(password: &str, salt: &[u8], key_bits: u32, block: u32) -> [u8; 16] {
    let pw = password_to_utf16le(password);
    let mut buf = Vec::with_capacity(salt.len() + pw.len());
    buf.extend_from_slice(salt);
    buf.extend_from_slice(&pw);
    let mut h = sha1_digest(&buf).to_vec();
    for i in 0..50_000u32 {
        let mut tmp = Vec::with_capacity(4 + h.len());
        tmp.extend_from_slice(&i.to_le_bytes());
        tmp.extend_from_slice(&h);
        h = sha1_digest(&tmp).to_vec();
    }
    let mut tmp = Vec::with_capacity(h.len() + 4);
    tmp.extend_from_slice(&h);
    tmp.extend_from_slice(&block.to_le_bytes());
    let h2 = sha1_digest(&tmp);
    let key_len = (key_bits as usize) / 8;
    let mut out = [0u8; 16];
    out.copy_from_slice(&h2[..key_len]);
    out
}

fn aes128_cbc_encrypt(key: &[u8; 16], iv: &[u8; 16], plaintext: &[u8]) -> Vec<u8> {
    use aes::Aes128;
    use cbc::Encryptor;
    let mut buf = plaintext.to_vec();
    let enc = Encryptor::<Aes128>::new_from_slices(key, iv).expect("key/iv");
    enc.encrypt_padded_mut::<NoPadding>(&mut buf, plaintext.len())
        .expect("encrypt");
    buf
}

fn aes256_cbc_encrypt(key: &[u8; 32], iv: &[u8], plaintext: &[u8]) -> Vec<u8> {
    use aes::Aes256;
    use cbc::Encryptor;
    let mut buf = plaintext.to_vec();
    let enc = Encryptor::<Aes256>::new_from_slices(key, iv).expect("key/iv");
    enc.encrypt_padded_mut::<NoPadding>(&mut buf, plaintext.len())
        .expect("encrypt");
    buf
}

fn agile_encrypt_with_block_key(
    salt: &[u8],
    password_utf16le: &[u8],
    spin_count: u32,
    key_bits: usize,
    block_size: usize,
    block_key: &[u8; 8],
    plaintext: &[u8],
) -> Vec<u8> {
    let key = agile_derive_key_sha512(salt, password_utf16le, spin_count, key_bits, block_key);
    let iv = &salt[..block_size];

    // plaintext must be multiple of 16 for NoPadding.
    assert!(plaintext.len() % 16 == 0);

    use aes::Aes256;
    use cbc::Encryptor;
    let mut buf = plaintext.to_vec();
    let enc = Encryptor::<Aes256>::new_from_slices(&key, iv).expect("key/iv");
    enc.encrypt_padded_mut::<NoPadding>(&mut buf, plaintext.len())
        .expect("encrypt");
    buf
}

fn agile_derive_key_sha512(
    salt: &[u8],
    password_utf16le: &[u8],
    spin_count: u32,
    key_bits: usize,
    block_key: &[u8; 8],
) -> [u8; 32] {
    let mut initial = Vec::with_capacity(salt.len() + password_utf16le.len());
    initial.extend_from_slice(salt);
    initial.extend_from_slice(password_utf16le);
    let mut h = sha512_digest(&initial);
    for i in 0..spin_count {
        let mut tmp = Vec::with_capacity(4 + h.len());
        tmp.extend_from_slice(&i.to_le_bytes());
        tmp.extend_from_slice(&h);
        h = sha512_digest(&tmp);
    }
    let mut tmp = Vec::with_capacity(h.len() + block_key.len());
    tmp.extend_from_slice(&h);
    tmp.extend_from_slice(block_key);
    let h2 = sha512_digest(&tmp);

    let key_len = key_bits / 8;
    assert_eq!(key_len, 32);
    let mut out = [0u8; 32];
    out.copy_from_slice(&h2[..32]);
    out
}
