//! Regression test: MS-OFFCRYPTO Standard / CryptoAPI / RC4 with `AlgIDHash == CALG_MD5`.
//!
//! Most Office-produced Standard RC4 workbooks use SHA-1 (CALG_SHA1) for password hashing, but the
//! spec allows MD5 (CALG_MD5). Ensure `formula-xlsx` can decrypt RC4+MD5 EncryptionInfo payloads.

use std::io::{Cursor, Write as _};

use formula_offcrypto::{decrypt_encrypted_package, DecryptOptions, OffcryptoError};
use formula_xlsx::offcrypto::{decrypt_ooxml_encrypted_package, OffCryptoError};
use md5::{Digest as _, Md5};
use zip::write::FileOptions;
use zip::ZipWriter;

const PASSWORD: &str = "password";

const CALG_RC4: u32 = 0x0000_6801;
const CALG_MD5: u32 = 0x0000_8003;
const PROV_RSA_FULL: u32 = 1;
const SPIN_COUNT: u32 = 50_000;
const RC4_BLOCK_SIZE: usize = 0x200;

const F_CRYPTOAPI: u32 = 0x0000_0004;

fn build_minimal_zip() -> Vec<u8> {
    // Small but valid ZIP container (enough for `formula-xlsx`'s "looks like ZIP" validation).
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Stored);

    zip.start_file("[Content_Types].xml", options)
        .expect("start file");
    zip.write_all(
        b"<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"/>",
    )
    .expect("write file");

    zip.finish().expect("finish zip").into_inner()
}

fn password_utf16le_bytes(password: &str) -> Vec<u8> {
    let mut out = Vec::with_capacity(password.len().saturating_mul(2));
    for cu in password.encode_utf16() {
        out.extend_from_slice(&cu.to_le_bytes());
    }
    out
}

fn spun_password_hash_md5(password: &str, salt: &[u8; 16]) -> [u8; 16] {
    // h = MD5(salt || UTF-16LE(password))
    let pw = password_utf16le_bytes(password);
    let mut hasher = Md5::new();
    hasher.update(salt);
    hasher.update(&pw);
    let mut h: [u8; 16] = hasher.finalize().into();

    // for i in 0..SPIN_COUNT: h = MD5(LE32(i) || h)
    for i in 0..SPIN_COUNT {
        let mut hasher = Md5::new();
        hasher.update(i.to_le_bytes());
        hasher.update(h);
        h = hasher.finalize().into();
    }

    h
}

fn derive_block_digest_md5(h: &[u8; 16], block_index: u32) -> [u8; 16] {
    let mut hasher = Md5::new();
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

fn encrypt_rc4_cryptoapi_per_block_128bit(plaintext: &[u8], h: &[u8; 16]) -> Vec<u8> {
    let mut out = Vec::with_capacity(plaintext.len());
    let mut offset = 0usize;
    let mut block_index = 0u32;
    while offset < plaintext.len() {
        let block_len = (plaintext.len() - offset).min(RC4_BLOCK_SIZE);
        let digest = derive_block_digest_md5(h, block_index);
        let chunk = rc4_apply(&digest, &plaintext[offset..offset + block_len]);
        out.extend_from_slice(&chunk);
        offset += block_len;
        block_index += 1;
    }
    out
}

fn build_encryption_info_standard_rc4_md5_128bit(
    salt: &[u8; 16],
    encrypted_verifier: &[u8; 16],
    encrypted_verifier_hash: &[u8; 16],
) -> Vec<u8> {
    let mut out = Vec::new();

    // EncryptionVersionInfo (Standard): 3.2
    out.extend_from_slice(&3u16.to_le_bytes()); // major
    out.extend_from_slice(&2u16.to_le_bytes()); // minor
    out.extend_from_slice(&0u32.to_le_bytes()); // flags

    // EncryptionHeader
    let csp_name = "Microsoft Enhanced Cryptographic Provider v1.0";
    let mut csp_utf16le = Vec::new();
    for cu in csp_name.encode_utf16() {
        csp_utf16le.extend_from_slice(&cu.to_le_bytes());
    }
    csp_utf16le.extend_from_slice(&0u16.to_le_bytes()); // NUL terminator

    let mut header = Vec::new();
    header.extend_from_slice(&F_CRYPTOAPI.to_le_bytes()); // Flags
    header.extend_from_slice(&0u32.to_le_bytes()); // SizeExtra
    header.extend_from_slice(&CALG_RC4.to_le_bytes()); // AlgID
    header.extend_from_slice(&CALG_MD5.to_le_bytes()); // AlgIDHash
    header.extend_from_slice(&128u32.to_le_bytes()); // KeySize (bits)
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
    out.extend_from_slice(&16u32.to_le_bytes()); // verifierHashSize (MD5)
    out.extend_from_slice(encrypted_verifier_hash);

    out
}

#[test]
fn decrypts_standard_cryptoapi_rc4_md5() {
    let plaintext = build_minimal_zip();

    // Deterministic parameters for a stable test.
    let salt: [u8; 16] = [
        0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D, 0x0E,
        0x0F,
    ];
    let verifier_plain: [u8; 16] = *b"0123456789ABCDEF";
    let verifier_hash: [u8; 16] = Md5::digest(verifier_plain).into();

    // --- Encrypt -------------------------------------------------------------------------------
    let h = spun_password_hash_md5(PASSWORD, &salt);
    let key0 = derive_block_digest_md5(&h, 0);

    // Encrypt verifier + verifier hash as a single RC4 stream.
    let mut verifier_concat = Vec::new();
    verifier_concat.extend_from_slice(&verifier_plain);
    verifier_concat.extend_from_slice(&verifier_hash);
    let verifier_cipher = rc4_apply(&key0, &verifier_concat);
    let encrypted_verifier: [u8; 16] = verifier_cipher[..16].try_into().unwrap();
    let encrypted_verifier_hash: [u8; 16] = verifier_cipher[16..].try_into().unwrap();

    let encryption_info =
        build_encryption_info_standard_rc4_md5_128bit(&salt, &encrypted_verifier, &encrypted_verifier_hash);

    let ciphertext = encrypt_rc4_cryptoapi_per_block_128bit(&plaintext, &h);
    let mut encrypted_package = Vec::new();
    encrypted_package.extend_from_slice(&(plaintext.len() as u64).to_le_bytes());
    encrypted_package.extend_from_slice(&ciphertext);

    // --- Decrypt -------------------------------------------------------------------------------
    let decrypted_offcrypto = decrypt_encrypted_package(
        &encryption_info,
        &encrypted_package,
        PASSWORD,
        DecryptOptions::default(),
    )
    .expect("decrypt EncryptedPackage (formula-offcrypto)");
    assert_eq!(decrypted_offcrypto, plaintext);

    let err = decrypt_encrypted_package(
        &encryption_info,
        &encrypted_package,
        "wrong-password",
        DecryptOptions::default(),
    )
    .expect_err("expected invalid password (formula-offcrypto)");
    assert!(matches!(err, OffcryptoError::InvalidPassword));

    let decrypted = decrypt_ooxml_encrypted_package(&encryption_info, &encrypted_package, PASSWORD)
        .expect("decrypt EncryptedPackage");
    assert_eq!(decrypted, plaintext);

    let err = decrypt_ooxml_encrypted_package(&encryption_info, &encrypted_package, "wrong-password")
        .expect_err("expected invalid password");
    assert!(matches!(err, OffCryptoError::WrongPassword));
}
