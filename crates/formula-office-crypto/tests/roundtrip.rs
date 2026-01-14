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
fn roundtrip_standard_rc4_sha1_encryption() {
    let password = "password";
    let plaintext = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/basic/basic.xlsx"
    ));
    let ole_bytes =
        encrypt_standard_rc4_ooxml_ole(plaintext, password, Rc4HashAlgorithm::Sha1, 0);
    assert!(is_encrypted_ooxml_ole(&ole_bytes));

    let decrypted = decrypt_encrypted_package_ole(&ole_bytes, password).expect("decrypt");
    assert_eq!(decrypted, plaintext);
    assert_zip_contains_workbook_xml(&decrypted);
}

#[test]
fn roundtrip_standard_rc4_md5_encryption() {
    let password = "password";
    let plaintext = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/basic/basic.xlsx"
    ));
    // Add some trailing bytes after the verifier hash to ensure the parser does not over-read.
    let ole_bytes =
        encrypt_standard_rc4_ooxml_ole(plaintext, password, Rc4HashAlgorithm::Md5, 12);
    assert!(is_encrypted_ooxml_ole(&ole_bytes));

    let decrypted = decrypt_encrypted_package_ole(&ole_bytes, password).expect("decrypt");
    assert_eq!(decrypted, plaintext);
    assert_zip_contains_workbook_xml(&decrypted);
}

#[test]
fn standard_rc4_bogus_verifier_hash_size_errors() {
    let password = "password";
    let plaintext = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/basic/basic.xlsx"
    ));

    // Build an RC4 + MD5 EncryptionInfo, but lie about verifierHashSize (declare 20 bytes, which is
    // SHA-1's digest length). This should be rejected as invalid format (not InvalidPassword).
    let ole_bytes = encrypt_standard_rc4_ooxml_ole_with_overridden_verifier_hash_size(
        plaintext,
        password,
        Rc4HashAlgorithm::Md5,
        20,
    );

    let err = decrypt_encrypted_package_ole(&ole_bytes, password).expect_err("expected error");
    match err {
        OfficeCryptoError::InvalidFormat(msg) => {
            assert!(
                msg.contains("verifierHashSize"),
                "expected error message to mention verifierHashSize, got: {msg}"
            );
        }
        other => panic!("expected InvalidFormat, got {other:?}"),
    }
}

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

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Rc4HashAlgorithm {
    Sha1,
    Md5,
}

fn rc4_hash_alg_id(hash_alg: Rc4HashAlgorithm) -> u32 {
    match hash_alg {
        Rc4HashAlgorithm::Sha1 => 0x0000_8004, // CALG_SHA1
        Rc4HashAlgorithm::Md5 => 0x0000_8003,  // CALG_MD5
    }
}

fn rc4_hash_digest(hash_alg: Rc4HashAlgorithm, data: &[u8]) -> Vec<u8> {
    match hash_alg {
        Rc4HashAlgorithm::Sha1 => sha1_digest(data).to_vec(),
        Rc4HashAlgorithm::Md5 => {
            let mut hasher = md5::Md5::new();
            hasher.update(data);
            hasher.finalize().to_vec()
        }
    }
}

fn standard_rc4_spun_password_hash(
    hash_alg: Rc4HashAlgorithm,
    password: &str,
    salt: &[u8],
) -> Vec<u8> {
    let pw = password_to_utf16le(password);

    // h = Hash(salt || pw)
    let mut h = match hash_alg {
        Rc4HashAlgorithm::Sha1 => {
            let mut hasher = sha1::Sha1::new();
            hasher.update(salt);
            hasher.update(&pw);
            hasher.finalize().to_vec()
        }
        Rc4HashAlgorithm::Md5 => {
            let mut hasher = md5::Md5::new();
            hasher.update(salt);
            hasher.update(&pw);
            hasher.finalize().to_vec()
        }
    };

    // spin: h = Hash(LE32(i) || h)
    for i in 0..50_000u32 {
        h = match hash_alg {
            Rc4HashAlgorithm::Sha1 => {
                let mut hasher = sha1::Sha1::new();
                hasher.update(i.to_le_bytes());
                hasher.update(&h);
                hasher.finalize().to_vec()
            }
            Rc4HashAlgorithm::Md5 => {
                let mut hasher = md5::Md5::new();
                hasher.update(i.to_le_bytes());
                hasher.update(&h);
                hasher.finalize().to_vec()
            }
        };
    }

    h
}

fn standard_rc4_derive_key(
    hash_alg: Rc4HashAlgorithm,
    spun_hash: &[u8],
    key_len: usize,
    block_index: u32,
) -> Vec<u8> {
    let digest = match hash_alg {
        Rc4HashAlgorithm::Sha1 => {
            let mut hasher = sha1::Sha1::new();
            hasher.update(spun_hash);
            hasher.update(block_index.to_le_bytes());
            hasher.finalize().to_vec()
        }
        Rc4HashAlgorithm::Md5 => {
            let mut hasher = md5::Md5::new();
            hasher.update(spun_hash);
            hasher.update(block_index.to_le_bytes());
            hasher.finalize().to_vec()
        }
    };
    digest[..key_len].to_vec()
}

fn rc4_apply(key: &[u8], data: &mut [u8]) {
    assert!(!key.is_empty(), "RC4 key must be non-empty");
    let mut s = [0u8; 256];
    for (i, v) in s.iter_mut().enumerate() {
        *v = i as u8;
    }

    // KSA
    let mut j: u8 = 0;
    for i in 0..256u16 {
        let idx = i as usize;
        j = j
            .wrapping_add(s[idx])
            .wrapping_add(key[idx % key.len()]);
        s.swap(idx, j as usize);
    }

    // PRGA
    let mut i: u8 = 0;
    j = 0;
    for b in data {
        i = i.wrapping_add(1);
        j = j.wrapping_add(s[i as usize]);
        s.swap(i as usize, j as usize);
        let idx = s[i as usize].wrapping_add(s[j as usize]);
        let k = s[idx as usize];
        *b ^= k;
    }
}

fn encrypt_standard_rc4_ooxml_ole(
    plaintext: &[u8],
    password: &str,
    hash_alg: Rc4HashAlgorithm,
    verifier_trailing_len: usize,
) -> Vec<u8> {
    encrypt_standard_rc4_ooxml_ole_inner(plaintext, password, hash_alg, None, verifier_trailing_len)
}

fn encrypt_standard_rc4_ooxml_ole_with_overridden_verifier_hash_size(
    plaintext: &[u8],
    password: &str,
    hash_alg: Rc4HashAlgorithm,
    verifier_hash_size_override: u32,
) -> Vec<u8> {
    encrypt_standard_rc4_ooxml_ole_inner(
        plaintext,
        password,
        hash_alg,
        Some(verifier_hash_size_override),
        0,
    )
}

fn encrypt_standard_rc4_ooxml_ole_inner(
    plaintext: &[u8],
    password: &str,
    hash_alg: Rc4HashAlgorithm,
    verifier_hash_size_override: Option<u32>,
    verifier_trailing_len: usize,
) -> Vec<u8> {
    // Deterministic parameters (not intended to be secure).
    let salt: Vec<u8> = (0u8..=0x0F).collect();
    let key_bits = 128u32;
    let key_len = (key_bits / 8) as usize;

    // Derive spun password hash + block 0 key.
    let spun = standard_rc4_spun_password_hash(hash_alg, password, &salt);

    // Test vectors (lock down derivation for both SHA1 and MD5).
    if hash_alg == Rc4HashAlgorithm::Sha1 {
        let key0 = standard_rc4_derive_key(hash_alg, &spun, key_len, 0);
        assert_eq!(
            key0,
            hex_decode("6ad7dedf2da3514b1d85eabee069d47d"),
            "SHA1 block0 key vector mismatch"
        );
    }
    if hash_alg == Rc4HashAlgorithm::Md5 {
        let key0 = standard_rc4_derive_key(hash_alg, &spun, key_len, 0);
        assert_eq!(
            key0,
            hex_decode("69badcae244868e209d4e053ccd2a3bc"),
            "MD5 block0 key vector mismatch"
        );
    }

    let key0 = standard_rc4_derive_key(hash_alg, &spun, key_len, 0);

    // Build EncryptionVerifier.
    let verifier_plain: [u8; 16] = [
        0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D, 0x0E,
        0x0F,
    ];
    let verifier_hash = rc4_hash_digest(hash_alg, &verifier_plain);
    let verifier_hash_size = verifier_hash_size_override.unwrap_or(verifier_hash.len() as u32);

    let mut verifier_buf = Vec::new();
    verifier_buf.extend_from_slice(&verifier_plain);
    // Truncate/pad hash to match the declared size (for bogus size tests).
    if verifier_hash_size as usize <= verifier_hash.len() {
        verifier_buf.extend_from_slice(&verifier_hash[..verifier_hash_size as usize]);
    } else {
        verifier_buf.extend_from_slice(&verifier_hash);
        verifier_buf.resize(16 + verifier_hash_size as usize, 0);
    }
    rc4_apply(&key0, &mut verifier_buf);
    let encrypted_verifier = &verifier_buf[..16];
    let encrypted_verifier_hash = &verifier_buf[16..];

    // EncryptionInfo header.
    let version_major = 4u16;
    let version_minor = 2u16;
    let flags = 0x0000_0000u32;
    let header_flags = 0u32;
    let size_extra = 0u32;
    let alg_id = 0x0000_6801u32; // CALG_RC4
    let alg_id_hash = rc4_hash_alg_id(hash_alg);
    let provider_type = 0u32;
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

    let mut verifier_bytes = Vec::new();
    verifier_bytes.extend_from_slice(&(salt.len() as u32).to_le_bytes());
    verifier_bytes.extend_from_slice(&salt);
    verifier_bytes.extend_from_slice(encrypted_verifier);
    verifier_bytes.extend_from_slice(&verifier_hash_size.to_le_bytes());
    verifier_bytes.extend_from_slice(encrypted_verifier_hash);
    verifier_bytes.extend(std::iter::repeat(0xCCu8).take(verifier_trailing_len));

    let mut encryption_info = Vec::new();
    encryption_info.extend_from_slice(&version_major.to_le_bytes());
    encryption_info.extend_from_slice(&version_minor.to_le_bytes());
    encryption_info.extend_from_slice(&flags.to_le_bytes());
    encryption_info.extend_from_slice(&header_size.to_le_bytes());
    encryption_info.extend_from_slice(&header_bytes);
    encryption_info.extend_from_slice(&verifier_bytes);

    // Encrypt the package in 0x200-byte blocks using per-block keys.
    let mut ciphertext = plaintext.to_vec();
    let mut block_index: u32 = 0;
    for chunk in ciphertext.chunks_mut(0x200) {
        let key = standard_rc4_derive_key(hash_alg, &spun, key_len, block_index);
        rc4_apply(&key, chunk);
        block_index = block_index.checked_add(1).expect("block counter overflow");
    }

    let mut encrypted_package = Vec::new();
    encrypted_package.extend_from_slice(&(plaintext.len() as u64).to_le_bytes());
    encrypted_package.extend_from_slice(&ciphertext);

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

fn hex_decode(mut s: &str) -> Vec<u8> {
    s = s.trim();
    let mut compact = String::with_capacity(s.len());
    for ch in s.chars() {
        if ch.is_ascii_hexdigit() {
            compact.push(ch);
        }
    }
    assert_eq!(compact.len() % 2, 0, "hex string must have even length");
    let bytes = compact.as_bytes();
    let mut out = Vec::with_capacity(bytes.len() / 2);
    for i in (0..bytes.len()).step_by(2) {
        let hi = (bytes[i] as char).to_digit(16).unwrap();
        let lo = (bytes[i + 1] as char).to_digit(16).unwrap();
        out.push(((hi << 4) | lo) as u8);
    }
    out
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
    // The password key-encryptor fields use `saltValue` as the IV (truncated to blockSize),
    // unlike `EncryptedPackage` segment IVs which are derived from `keyData/@saltValue`.
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
