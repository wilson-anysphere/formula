#![cfg(not(target_arch = "wasm32"))]

//! End-to-end Standard (CryptoAPI) RC4 tests for non-128-bit key sizes.
//!
//! Motivation: non-128-bit Standard/CryptoAPI RC4 key sizes are easy to implement incorrectly.
//! MS-OFFCRYPTO specifies:
//! - `keyLen = keySize/8` bytes (40→5 bytes, 56→7 bytes)
//! - `keySize == 0` MUST be interpreted as 40-bit RC4
//!
//! This test builds a small synthetic Standard RC4 encrypted OLE container and ensures decryption
//! works for:
//! - keySize = 56 (7-byte RC4 key, **no** zero padding)
//! - keySize = 0 (treated as 40-bit)
//!
//! The encryption logic in this file is intentionally self-contained (does not call
//! `formula_offcrypto::cryptoapi::*`) so it can catch regressions in key derivation semantics.

use std::io::{Cursor, Write};

use md5::Md5;
use sha1::{Digest as _, Sha1};
use zip::write::FileOptions;

// CryptoAPI constants used by MS-OFFCRYPTO Standard encryption.
const CALG_RC4: u32 = 0x0000_6801;
const CALG_MD5: u32 = 0x0000_8003;
const CALG_SHA1: u32 = 0x0000_8004;

const SPIN_COUNT: u32 = 50_000;
const RC4_BLOCK_LEN: usize = 0x200;

#[derive(Debug, Clone, Copy)]
enum HashAlg {
    Sha1,
    Md5,
}

impl HashAlg {
    fn calg_id(self) -> u32 {
        match self {
            HashAlg::Sha1 => CALG_SHA1,
            HashAlg::Md5 => CALG_MD5,
        }
    }

    fn digest_len(self) -> usize {
        match self {
            HashAlg::Sha1 => 20,
            HashAlg::Md5 => 16,
        }
    }
}

#[derive(Clone)]
struct Rc4 {
    s: [u8; 256],
    i: u8,
    j: u8,
}

impl Rc4 {
    fn new(key: &[u8]) -> Self {
        assert!(!key.is_empty(), "RC4 key must be non-empty");
        let mut s = [0u8; 256];
        for (i, b) in s.iter_mut().enumerate() {
            *b = i as u8;
        }
        let mut j = 0u8;
        for i in 0..256usize {
            j = j.wrapping_add(s[i]).wrapping_add(key[i % key.len()]);
            s.swap(i, j as usize);
        }
        Self { s, i: 0, j: 0 }
    }

    fn apply_keystream(&mut self, buf: &mut [u8]) {
        for b in buf {
            self.i = self.i.wrapping_add(1);
            self.j = self.j.wrapping_add(self.s[self.i as usize]);
            self.s.swap(self.i as usize, self.j as usize);
            let t = self.s[self.i as usize].wrapping_add(self.s[self.j as usize]);
            let k = self.s[t as usize];
            *b ^= k;
        }
    }
}

fn utf16le_bytes(s: &str) -> Vec<u8> {
    let mut out = Vec::with_capacity(s.len().saturating_mul(2));
    for cu in s.encode_utf16() {
        out.extend_from_slice(&cu.to_le_bytes());
    }
    out
}

fn iterated_hash(password: &str, salt: &[u8; 16], hash_alg: HashAlg) -> Vec<u8> {
    let pw = utf16le_bytes(password);
    match hash_alg {
        HashAlg::Sha1 => {
            // H0 = SHA1(salt || password_utf16le)
            let mut hasher = Sha1::new();
            hasher.update(salt);
            hasher.update(&pw);
            let mut h: [u8; 20] = hasher.finalize().into();

            // Hi = SHA1(LE32(i) || H_{i-1}) for i in 0..49999
            let mut buf = [0u8; 4 + 20];
            for i in 0..SPIN_COUNT {
                buf[..4].copy_from_slice(&i.to_le_bytes());
                buf[4..].copy_from_slice(&h);
                h = Sha1::digest(&buf).into();
            }
            h.to_vec()
        }
        HashAlg::Md5 => {
            // H0 = MD5(salt || password_utf16le)
            let mut hasher = Md5::new();
            hasher.update(salt);
            hasher.update(&pw);
            let mut h: [u8; 16] = hasher.finalize().into();

            // Hi = MD5(LE32(i) || H_{i-1}) for i in 0..49999
            let mut buf = [0u8; 4 + 16];
            for i in 0..SPIN_COUNT {
                buf[..4].copy_from_slice(&i.to_le_bytes());
                buf[4..].copy_from_slice(&h);
                h = Md5::digest(&buf).into();
            }
            h.to_vec()
        }
    }
}

fn block_hash(h: &[u8], block: u32, hash_alg: HashAlg) -> Vec<u8> {
    match hash_alg {
        HashAlg::Sha1 => {
            let mut hasher = Sha1::new();
            hasher.update(h);
            hasher.update(block.to_le_bytes());
            let digest: [u8; 20] = hasher.finalize().into();
            digest.to_vec()
        }
        HashAlg::Md5 => {
            let mut hasher = Md5::new();
            hasher.update(h);
            hasher.update(block.to_le_bytes());
            let digest: [u8; 16] = hasher.finalize().into();
            digest.to_vec()
        }
    }
}

fn rc4_key_for_block(h: &[u8], block: u32, key_size_bits: u32, hash_alg: HashAlg) -> Vec<u8> {
    // MS-OFFCRYPTO: keySize==0 => 40-bit.
    let key_size_bits = if key_size_bits == 0 { 40 } else { key_size_bits };
    assert!(key_size_bits % 8 == 0);
    let key_len = (key_size_bits / 8) as usize;
    let digest = block_hash(h, block, hash_alg);

    digest[..key_len].to_vec()
}

fn build_tiny_zip() -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut writer = zip::ZipWriter::new(cursor);
    writer
        .start_file("hello.txt", FileOptions::<()>::default())
        .expect("start file");
    writer.write_all(b"hello").expect("write file");
    writer.finish().expect("finish zip").into_inner()
}

fn build_standard_rc4_encryption_info(
    password: &str,
    salt: [u8; 16],
    key_size_bits: u32,
    hash_alg: HashAlg,
) -> Vec<u8> {
    // Deterministic verifier plaintext.
    let verifier_plain: [u8; 16] = [
        0x10, 0x11, 0x12, 0x13, 0x14, 0x15, 0x16, 0x17, 0x18, 0x19, 0x1A, 0x1B, 0x1C, 0x1D,
        0x1E, 0x1F,
    ];
    let verifier_hash = match hash_alg {
        HashAlg::Sha1 => {
            let d: [u8; 20] = Sha1::digest(&verifier_plain).into();
            d.to_vec()
        }
        HashAlg::Md5 => {
            let d: [u8; 16] = Md5::digest(&verifier_plain).into();
            d.to_vec()
        }
    };

    let h = iterated_hash(password, &salt, hash_alg);
    let key0 = rc4_key_for_block(&h, 0, key_size_bits, hash_alg);

    // Encrypt verifier + verifier hash with a single RC4 stream (no reset between fields).
    let mut rc4 = Rc4::new(&key0);
    let mut encrypted_verifier = verifier_plain;
    rc4.apply_keystream(&mut encrypted_verifier);
    let mut encrypted_verifier_hash = verifier_hash;
    rc4.apply_keystream(&mut encrypted_verifier_hash);

    // EncryptionInfo header (Standard 3.2).
    let mut out = Vec::new();
    out.extend_from_slice(&3u16.to_le_bytes()); // versionMajor
    out.extend_from_slice(&2u16.to_le_bytes()); // versionMinor
    out.extend_from_slice(&0u32.to_le_bytes()); // EncryptionInfo.flags

    // EncryptionHeader (8 DWORDs + UTF-16LE CSPName).
    let mut header_bytes = Vec::new();
    let header_flags = 0x0000_0004u32; // fCryptoAPI
    header_bytes.extend_from_slice(&header_flags.to_le_bytes());
    header_bytes.extend_from_slice(&0u32.to_le_bytes()); // sizeExtra
    header_bytes.extend_from_slice(&CALG_RC4.to_le_bytes()); // algId
    header_bytes.extend_from_slice(&hash_alg.calg_id().to_le_bytes()); // algIdHash
    header_bytes.extend_from_slice(&key_size_bits.to_le_bytes()); // keySize (bits)
    header_bytes.extend_from_slice(&0u32.to_le_bytes()); // providerType
    header_bytes.extend_from_slice(&0u32.to_le_bytes()); // reserved1
    header_bytes.extend_from_slice(&0u32.to_le_bytes()); // reserved2
    // Empty CSPName but include a UTF-16 NUL terminator to keep the header well-formed.
    header_bytes.extend_from_slice(&0u16.to_le_bytes());

    out.extend_from_slice(&(header_bytes.len() as u32).to_le_bytes());
    out.extend_from_slice(&header_bytes);

    // EncryptionVerifier
    out.extend_from_slice(&16u32.to_le_bytes()); // saltSize
    out.extend_from_slice(&salt);
    out.extend_from_slice(&encrypted_verifier);
    out.extend_from_slice(&(hash_alg.digest_len() as u32).to_le_bytes()); // verifierHashSize
    out.extend_from_slice(&encrypted_verifier_hash);

    out
}

fn build_standard_rc4_encrypted_package(
    password: &str,
    salt: [u8; 16],
    key_size_bits: u32,
    hash_alg: HashAlg,
    plaintext_zip: &[u8],
) -> Vec<u8> {
    let h = iterated_hash(password, &salt, hash_alg);

    let mut ciphertext = plaintext_zip.to_vec();
    for (block, chunk) in ciphertext.chunks_mut(RC4_BLOCK_LEN).enumerate() {
        let key = rc4_key_for_block(&h, block as u32, key_size_bits, hash_alg);
        let mut rc4 = Rc4::new(&key);
        rc4.apply_keystream(chunk);
    }

    let mut out = Vec::new();
    out.extend_from_slice(&(plaintext_zip.len() as u64).to_le_bytes());
    out.extend_from_slice(&ciphertext);
    out
}

fn build_ole_container(encryption_info: &[u8], encrypted_package: &[u8]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");

    {
        let mut s = ole
            .create_stream("/EncryptionInfo")
            .expect("create EncryptionInfo");
        s.write_all(encryption_info).expect("write EncryptionInfo");
    }
    {
        let mut s = ole
            .create_stream("/EncryptedPackage")
            .expect("create EncryptedPackage");
        s.write_all(encrypted_package)
            .expect("write EncryptedPackage");
    }

    ole.into_inner().into_inner()
}

#[test]
fn decrypts_standard_rc4_keysize_56_roundtrip() {
    for hash_alg in [HashAlg::Sha1, HashAlg::Md5] {
        let password = "password";
        let salt: [u8; 16] = [
            0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C,
            0x0D, 0x0E, 0x0F,
        ];
        let key_size_bits = 56;

        let plaintext = build_tiny_zip();
        assert!(plaintext.starts_with(b"PK"));

        let encryption_info =
            build_standard_rc4_encryption_info(password, salt, key_size_bits, hash_alg);
        let encrypted_package =
            build_standard_rc4_encrypted_package(password, salt, key_size_bits, hash_alg, &plaintext);
        let ole_bytes = build_ole_container(&encryption_info, &encrypted_package);

        let decrypted =
            formula_offcrypto::decrypt_standard_ooxml_from_bytes(ole_bytes, password).unwrap_or_else(
                |err| {
                    panic!("decrypt failed for {hash_alg:?}: {err}");
                },
            );
        assert_eq!(decrypted, plaintext, "hash_alg={hash_alg:?}");
    }
}

#[test]
fn decrypts_standard_rc4_keysize_zero_means_40_roundtrip() {
    for hash_alg in [HashAlg::Sha1, HashAlg::Md5] {
        let password = "password";
        let salt: [u8; 16] = [
            0x10, 0x11, 0x12, 0x13, 0x14, 0x15, 0x16, 0x17, 0x18, 0x19, 0x1A, 0x1B, 0x1C,
            0x1D, 0x1E, 0x1F,
        ];
        let key_size_bits = 0; // MUST be treated as 40-bit per MS-OFFCRYPTO.

        let plaintext = build_tiny_zip();
        assert!(plaintext.starts_with(b"PK"));

        let encryption_info =
            build_standard_rc4_encryption_info(password, salt, key_size_bits, hash_alg);
        let encrypted_package =
            build_standard_rc4_encrypted_package(password, salt, key_size_bits, hash_alg, &plaintext);
        let ole_bytes = build_ole_container(&encryption_info, &encrypted_package);

        let decrypted =
            formula_offcrypto::decrypt_standard_ooxml_from_bytes(ole_bytes, password).unwrap_or_else(
                |err| {
                    panic!("decrypt failed for {hash_alg:?}: {err}");
                },
            );
        assert_eq!(decrypted, plaintext, "hash_alg={hash_alg:?}");
    }
}
