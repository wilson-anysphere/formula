//! Regenerate Standard (CryptoAPI) encrypted OOXML fixtures under `fixtures/encrypted/ooxml/`.
//!
//! This is a developer utility and is not invoked by tests/CI. It produces deterministic output so
//! the committed fixtures are stable across runs.
//!
//! Usage (from repo root):
//! ```bash
//! bash scripts/cargo_agent.sh run -p formula-xlsx --example regen_standard_fixtures
//! ```

use std::io::{Cursor, Write as _};
use std::path::PathBuf;

use aes::cipher::{generic_array::GenericArray, BlockEncrypt, KeyInit};
use aes::{Aes128, Aes192, Aes256};
use cbc::cipher::{block_padding::NoPadding, BlockEncryptMut, KeyIvInit};
use sha1::{Digest as _, Sha1};

use formula_xlsx::offcrypto::{hash_password, HashAlgorithm};

const PASSWORD: &str = "password";
const STANDARD_SPIN_COUNT: u32 = 50_000;

// CryptoAPI algorithm identifiers used by Standard encryption.
const CALG_AES_128: u32 = 0x0000_660E;
const CALG_SHA1: u32 = 0x0000_8004;

const ENCRYPTED_PACKAGE_SEGMENT_LEN: usize = 0x1000;
const AES_BLOCK_LEN: usize = 16;

fn derive_segment_iv(salt: &[u8], segment_index: u32) -> [u8; AES_BLOCK_LEN] {
    let mut hasher = Sha1::new();
    hasher.update(salt);
    hasher.update(segment_index.to_le_bytes());
    let digest = hasher.finalize();

    let mut iv = [0u8; AES_BLOCK_LEN];
    iv.copy_from_slice(&digest[..AES_BLOCK_LEN]);
    iv
}

fn pkcs7_pad(plaintext: &[u8]) -> Vec<u8> {
    if plaintext.is_empty() {
        return Vec::new();
    }
    let mut out = plaintext.to_vec();
    let mut pad_len = AES_BLOCK_LEN - (out.len() % AES_BLOCK_LEN);
    if pad_len == 0 {
        pad_len = AES_BLOCK_LEN;
    }
    out.extend(std::iter::repeat(pad_len as u8).take(pad_len));
    out
}

fn aes_ecb_encrypt_in_place(key: &[u8], buf: &mut [u8]) {
    assert_eq!(buf.len() % AES_BLOCK_LEN, 0);
    match key.len() {
        16 => {
            let cipher = Aes128::new_from_slice(key).expect("valid AES-128 key");
            for block in buf.chunks_mut(AES_BLOCK_LEN) {
                cipher.encrypt_block(GenericArray::from_mut_slice(block));
            }
        }
        24 => {
            let cipher = Aes192::new_from_slice(key).expect("valid AES-192 key");
            for block in buf.chunks_mut(AES_BLOCK_LEN) {
                cipher.encrypt_block(GenericArray::from_mut_slice(block));
            }
        }
        32 => {
            let cipher = Aes256::new_from_slice(key).expect("valid AES-256 key");
            for block in buf.chunks_mut(AES_BLOCK_LEN) {
                cipher.encrypt_block(GenericArray::from_mut_slice(block));
            }
        }
        other => panic!("unexpected AES key length {other}"),
    }
}

fn encrypt_segment_aes_cbc_no_padding(key: &[u8], iv: &[u8; AES_BLOCK_LEN], plaintext: &[u8]) -> Vec<u8> {
    assert!(plaintext.len() % AES_BLOCK_LEN == 0);

    let mut buf = plaintext.to_vec();
    match key.len() {
        16 => {
            cbc::Encryptor::<Aes128>::new_from_slices(key, iv)
                .unwrap()
                .encrypt_padded_mut::<NoPadding>(&mut buf, plaintext.len())
                .unwrap();
        }
        24 => {
            cbc::Encryptor::<Aes192>::new_from_slices(key, iv)
                .unwrap()
                .encrypt_padded_mut::<NoPadding>(&mut buf, plaintext.len())
                .unwrap();
        }
        32 => {
            cbc::Encryptor::<Aes256>::new_from_slices(key, iv)
                .unwrap()
                .encrypt_padded_mut::<NoPadding>(&mut buf, plaintext.len())
                .unwrap();
        }
        other => panic!("unexpected AES key length {other}"),
    }
    buf
}

fn derive_cryptoapi_aes_key_sha1(password: &str, salt: &[u8], key_len: usize) -> Vec<u8> {
    let pw_hash = hash_password(password, salt, STANDARD_SPIN_COUNT, HashAlgorithm::Sha1)
        .expect("password hash");

    assert_eq!(pw_hash.len(), 20, "SHA1 hash length");

    // H_block0 = SHA1(pw_hash || LE32(0))
    let mut buf0 = [0u8; 20 + 4];
    buf0[..20].copy_from_slice(&pw_hash);
    buf0[20..].copy_from_slice(&0u32.to_le_bytes());
    let h_block0 = Sha1::digest(&buf0);

    // CryptoAPI `CryptDeriveKey` ipad/opad expansion for SHA1 (64-byte block size).
    let mut ipad = [0x36u8; 64];
    let mut opad = [0x5cu8; 64];
    for i in 0..20 {
        ipad[i] ^= h_block0[i];
        opad[i] ^= h_block0[i];
    }
    let x1 = Sha1::digest(&ipad);
    let x2 = Sha1::digest(&opad);

    let mut key_material = [0u8; 40];
    key_material[..20].copy_from_slice(&x1);
    key_material[20..].copy_from_slice(&x2);

    key_material[..key_len].to_vec()
}

fn encrypt_encrypted_package_stream_standard_cryptoapi(key: &[u8], salt: &[u8], plaintext: &[u8]) -> Vec<u8> {
    let orig_size = plaintext.len() as u64;

    let mut out = Vec::new();
    out.extend_from_slice(&orig_size.to_le_bytes());

    if plaintext.is_empty() {
        return out;
    }

    let padded = pkcs7_pad(plaintext);
    for (i, chunk) in padded.chunks(ENCRYPTED_PACKAGE_SEGMENT_LEN).enumerate() {
        let iv = derive_segment_iv(salt, i as u32);
        let ciphertext = encrypt_segment_aes_cbc_no_padding(key, &iv, chunk);
        out.extend_from_slice(&ciphertext);
    }

    out
}

fn build_standard_encrypted_ooxml_ole_bytes(package_bytes: &[u8], password: &str) -> Vec<u8> {
    // Deterministic fixture params.
    let salt: [u8; 16] = [
        0xe8, 0x82, 0x66, 0x49, 0x0c, 0x5b, 0xd1, 0xee, 0xbd, 0x2b, 0x43, 0x94, 0xe3, 0xf8, 0x30,
        0xef,
    ];
    let key_size_bits: u32 = 128;
    let key_len = (key_size_bits / 8) as usize;

    let key = derive_cryptoapi_aes_key_sha1(password, &salt, key_len);

    let verifier_plain: [u8; 16] = [
        0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0a, 0x0b, 0x0c, 0x0d, 0x0e,
        0x0f,
    ];
    let verifier_hash: [u8; 20] = Sha1::digest(&verifier_plain).into();
    let mut verifier_hash_padded = [0u8; 32];
    verifier_hash_padded[..20].copy_from_slice(&verifier_hash);
    verifier_hash_padded[20..].fill(0xa5);

    let mut encrypted_verifier = verifier_plain;
    aes_ecb_encrypt_in_place(&key, &mut encrypted_verifier);

    let mut encrypted_verifier_hash = verifier_hash_padded;
    aes_ecb_encrypt_in_place(&key, &mut encrypted_verifier_hash);

    // --- Build EncryptionInfo stream bytes (Standard 3.2) ---
    let mut encryption_info_bytes = Vec::new();
    encryption_info_bytes.extend_from_slice(&3u16.to_le_bytes()); // major
    encryption_info_bytes.extend_from_slice(&2u16.to_le_bytes()); // minor
    encryption_info_bytes.extend_from_slice(&0u32.to_le_bytes()); // flags

    // EncryptionHeader: 8 fixed u32s + optional UTF-16LE cspName bytes.
    let header_bytes_len = 8 * 4;
    encryption_info_bytes.extend_from_slice(&(header_bytes_len as u32).to_le_bytes());

    encryption_info_bytes.extend_from_slice(&0u32.to_le_bytes()); // flags
    encryption_info_bytes.extend_from_slice(&0u32.to_le_bytes()); // sizeExtra
    encryption_info_bytes.extend_from_slice(&CALG_AES_128.to_le_bytes()); // algId
    encryption_info_bytes.extend_from_slice(&CALG_SHA1.to_le_bytes()); // algIdHash
    encryption_info_bytes.extend_from_slice(&key_size_bits.to_le_bytes()); // keySize (bits)
    encryption_info_bytes.extend_from_slice(&0u32.to_le_bytes()); // providerType
    encryption_info_bytes.extend_from_slice(&0u32.to_le_bytes()); // reserved1
    encryption_info_bytes.extend_from_slice(&0u32.to_le_bytes()); // reserved2

    // EncryptionVerifier.
    encryption_info_bytes.extend_from_slice(&(salt.len() as u32).to_le_bytes());
    encryption_info_bytes.extend_from_slice(&salt);
    encryption_info_bytes.extend_from_slice(&encrypted_verifier);
    encryption_info_bytes.extend_from_slice(&(20u32).to_le_bytes());
    encryption_info_bytes.extend_from_slice(&encrypted_verifier_hash);

    // --- Build EncryptedPackage stream bytes ---
    let encrypted_package_bytes = encrypt_encrypted_package_stream_standard_cryptoapi(&key, &salt, package_bytes);

    // --- Wrap in an OLE/CFB container ---
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

fn main() {
    let fixture_dir: PathBuf = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/encrypted/ooxml");

    let plaintext_small = std::fs::read(fixture_dir.join("plaintext.xlsx")).expect("read plaintext.xlsx");
    let plaintext_large = std::fs::read(fixture_dir.join("plaintext-large.xlsx")).expect("read plaintext-large.xlsx");

    let standard_small = build_standard_encrypted_ooxml_ole_bytes(&plaintext_small, PASSWORD);
    let standard_large = build_standard_encrypted_ooxml_ole_bytes(&plaintext_large, PASSWORD);

    std::fs::write(fixture_dir.join("standard.xlsx"), standard_small).expect("write standard.xlsx");
    std::fs::write(fixture_dir.join("standard-large.xlsx"), standard_large)
        .expect("write standard-large.xlsx");

    eprintln!("regenerated Standard/CryptoAPI fixtures in {}", fixture_dir.display());
}

