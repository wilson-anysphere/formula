//! Regression test: MS-OFFCRYPTO Standard / CryptoAPI / RC4 where `EncryptionHeader.keySize == 0`.
//!
//! MS-OFFCRYPTO specifies that a `keySize` of 0 MUST be interpreted as 40-bit RC4.
//! `formula-io` should accept such files and decrypt them correctly.
#![cfg(all(feature = "encrypted-workbooks", not(target_arch = "wasm32")))]

use std::io::{Cursor, Write as _};

use formula_io::{open_workbook_model_with_password, Error};
use formula_model::{CellRef, CellValue};
use sha1::{Digest as _, Sha1};

const CALG_RC4: u32 = 0x0000_6801;
const CALG_SHA1: u32 = 0x0000_8004;
const PROV_RSA_FULL: u32 = 1;
const F_CRYPTOAPI: u32 = 0x0000_0004;
const F_AES: u32 = 0x0000_0020;
const SPIN_COUNT: u32 = 50_000;
const RC4_BLOCK_SIZE: usize = 0x200;

fn build_tiny_xlsx() -> Vec<u8> {
    let mut workbook = formula_model::Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1").expect("add sheet");
    let sheet = workbook.sheet_mut(sheet_id).expect("sheet exists");
    sheet.set_value(CellRef::from_a1("A1").unwrap(), CellValue::Number(1.0));
    sheet.set_value(
        CellRef::from_a1("B1").unwrap(),
        CellValue::String("Hello".to_string()),
    );

    let mut cursor = Cursor::new(Vec::new());
    formula_io::xlsx::write_workbook_to_writer(&workbook, &mut cursor).expect("write xlsx bytes");
    cursor.into_inner()
}

fn password_utf16le_bytes(password: &str) -> Vec<u8> {
    let mut out = Vec::with_capacity(password.len().saturating_mul(2));
    for cu in password.encode_utf16() {
        out.extend_from_slice(&cu.to_le_bytes());
    }
    out
}

fn spun_password_hash_sha1(password: &str, salt: &[u8; 16]) -> [u8; 20] {
    // h = SHA1(salt || UTF-16LE(password))
    let pw = password_utf16le_bytes(password);
    let mut hasher = Sha1::new();
    hasher.update(salt);
    hasher.update(&pw);
    let mut h: [u8; 20] = hasher.finalize().into();

    // for i in 0..SPIN_COUNT: h = SHA1(LE32(i) || h)
    for i in 0..SPIN_COUNT {
        let mut hasher = Sha1::new();
        hasher.update(i.to_le_bytes());
        hasher.update(h);
        h = hasher.finalize().into();
    }

    h
}

fn derive_block_digest_sha1(h: &[u8; 20], block_index: u32) -> [u8; 20] {
    let mut hasher = Sha1::new();
    hasher.update(h);
    hasher.update(block_index.to_le_bytes());
    hasher.finalize().into()
}

/// Minimal RC4 implementation (KSA + PRGA).
fn rc4_apply(key: &[u8], data: &[u8]) -> Vec<u8> {
    assert!(!key.is_empty(), "RC4 key must be non-empty");

    let mut s = [0u8; 256];
    for (i, b) in s.iter_mut().enumerate() {
        *b = i as u8;
    }

    let mut j: u8 = 0;
    for i in 0..256u16 {
        j = j
            .wrapping_add(s[i as usize])
            .wrapping_add(key[i as usize % key.len()]);
        s.swap(i as usize, j as usize);
    }

    let mut i: u8 = 0;
    j = 0;
    let mut out = Vec::with_capacity(data.len());
    for &b in data {
        i = i.wrapping_add(1);
        j = j.wrapping_add(s[i as usize]);
        s.swap(i as usize, j as usize);
        let k = s[(s[i as usize].wrapping_add(s[j as usize])) as usize];
        out.push(b ^ k);
    }
    out
}

fn rc4_key_40bit_from_digest(digest: &[u8; 20]) -> [u8; 16] {
    // CryptoAPI/Office represent a "40-bit" RC4 key as a 128-bit key with the high 88 bits zero.
    let mut key = [0u8; 16];
    key[..5].copy_from_slice(&digest[..5]);
    key
}

fn encrypt_rc4_cryptoapi_per_block_40bit(plaintext: &[u8], h: &[u8; 20]) -> Vec<u8> {
    let mut out = Vec::with_capacity(plaintext.len());
    let mut offset = 0usize;
    let mut block_index = 0u32;
    while offset < plaintext.len() {
        let block_len = (plaintext.len() - offset).min(RC4_BLOCK_SIZE);
        let digest = derive_block_digest_sha1(h, block_index);
        let key = rc4_key_40bit_from_digest(&digest);
        let chunk = rc4_apply(&key, &plaintext[offset..offset + block_len]);
        out.extend_from_slice(&chunk);
        offset += block_len;
        block_index += 1;
    }
    out
}

fn build_encryption_info_standard_rc4_keysize_zero(
    salt: &[u8; 16],
    encrypted_verifier: &[u8; 16],
    encrypted_verifier_hash: &[u8; 20],
) -> Vec<u8> {
    let mut out = Vec::new();

    // EncryptionVersionInfo (Standard): 3.2
    out.extend_from_slice(&3u16.to_le_bytes()); // major
    out.extend_from_slice(&2u16.to_le_bytes()); // minor
    // Standard/CryptoAPI EncryptionInfo commonly uses 0x0000_0040 for this outer flags field.
    // The critical bits for decryptors are in the inner `EncryptionHeader.flags`.
    out.extend_from_slice(&0x0000_0040u32.to_le_bytes()); // flags

    // EncryptionHeader
    let csp_name = "Microsoft Enhanced Cryptographic Provider v1.0";
    let mut csp_utf16le = Vec::new();
    for cu in csp_name.encode_utf16() {
        csp_utf16le.extend_from_slice(&cu.to_le_bytes());
    }
    csp_utf16le.extend_from_slice(&0u16.to_le_bytes()); // NUL terminator

    let mut header = Vec::new();
    // MS-OFFCRYPTO Standard `EncryptionHeader.flags`:
    // - fCryptoAPI must be set for Standard/CryptoAPI encryption.
    // - fAES must be unset for RC4.
    header.extend_from_slice(&(F_CRYPTOAPI & !F_AES).to_le_bytes()); // Flags
    header.extend_from_slice(&0u32.to_le_bytes()); // SizeExtra
    header.extend_from_slice(&CALG_RC4.to_le_bytes()); // AlgID
    header.extend_from_slice(&CALG_SHA1.to_le_bytes()); // AlgIDHash
    header.extend_from_slice(&0u32.to_le_bytes()); // KeySize (0 => 40-bit)
    header.extend_from_slice(&PROV_RSA_FULL.to_le_bytes()); // ProviderType
    header.extend_from_slice(&0u32.to_le_bytes()); // Reserved1
    header.extend_from_slice(&0u32.to_le_bytes()); // Reserved2
    header.extend_from_slice(&csp_utf16le);

    out.extend_from_slice(&(header.len() as u32).to_le_bytes()); // HeaderSize
    out.extend_from_slice(&header);

    // EncryptionVerifier
    out.extend_from_slice(&16u32.to_le_bytes()); // saltSize
    out.extend_from_slice(salt);
    out.extend_from_slice(encrypted_verifier);
    out.extend_from_slice(&20u32.to_le_bytes()); // verifierHashSize (SHA1)
    out.extend_from_slice(encrypted_verifier_hash);

    out
}

#[test]
fn decrypts_standard_cryptoapi_rc4_with_keysize_zero() {
    let password = "password";
    let plaintext = build_tiny_xlsx();

    // Deterministic parameters for a stable fixture.
    let salt: [u8; 16] = [
        0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D,
        0x0E, 0x0F,
    ];
    let verifier_plain: [u8; 16] = *b"0123456789ABCDEF";
    let verifier_hash: [u8; 20] = Sha1::digest(verifier_plain).into();

    // --- Encrypt -------------------------------------------------------------------------------
    let h = spun_password_hash_sha1(password, &salt);
    let digest0 = derive_block_digest_sha1(&h, 0);
    let key0 = rc4_key_40bit_from_digest(&digest0);

    // Encrypt verifier + verifier hash as a single RC4 stream.
    let mut verifier_concat = Vec::new();
    verifier_concat.extend_from_slice(&verifier_plain);
    verifier_concat.extend_from_slice(&verifier_hash);
    let verifier_cipher = rc4_apply(&key0, &verifier_concat);
    let encrypted_verifier: [u8; 16] = verifier_cipher[..16].try_into().unwrap();
    let encrypted_verifier_hash: [u8; 20] = verifier_cipher[16..].try_into().unwrap();

    let encryption_info = build_encryption_info_standard_rc4_keysize_zero(
        &salt,
        &encrypted_verifier,
        &encrypted_verifier_hash,
    );

    // Encrypt EncryptedPackage payload (RC4 in 0x200-byte blocks) with the same `keySize=0` => 40-bit semantics.
    let ciphertext = encrypt_rc4_cryptoapi_per_block_40bit(&plaintext, &h);
    let mut encrypted_package = Vec::new();
    encrypted_package.extend_from_slice(&(plaintext.len() as u64).to_le_bytes());
    encrypted_package.extend_from_slice(&ciphertext);

    // Wrap in an OLE/CFB container with the required streams.
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create OLE container");
    ole.create_stream("EncryptionInfo")
        .expect("create EncryptionInfo")
        .write_all(&encryption_info)
        .expect("write EncryptionInfo");
    ole.create_stream("EncryptedPackage")
        .expect("create EncryptedPackage")
        .write_all(&encrypted_package)
        .expect("write EncryptedPackage");
    let bytes = ole.into_inner().into_inner();

    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("standard-rc4-keysize0.xlsx");
    std::fs::write(&path, bytes).expect("write encrypted file");

    // --- Decrypt + open ------------------------------------------------------------------------
    let wrong = open_workbook_model_with_password(&path, Some("wrong-password"));
    assert!(
        matches!(wrong, Err(Error::InvalidPassword { .. })),
        "wrong password should return InvalidPassword, got {wrong:?}"
    );

    let workbook =
        open_workbook_model_with_password(&path, Some(password)).expect("decrypt + open workbook");
    let sheet = workbook.sheet_by_name("Sheet1").expect("Sheet1 missing");
    assert_eq!(
        sheet.value(CellRef::from_a1("A1").unwrap()),
        CellValue::Number(1.0)
    );
    assert_eq!(
        sheet.value(CellRef::from_a1("B1").unwrap()),
        CellValue::String("Hello".to_string())
    );
}
