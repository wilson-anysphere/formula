//! Generate a Standard/CryptoAPI RC4 encrypted OOXML fixture (e.g. `.xlsx`/`.xlsm`/`.xlsb`).
//!
//! This example is **not** part of the normal test suite; it exists to document how the
//! `fixtures/encrypted/ooxml/standard-rc4.xlsx` blob was generated.
//!
//! Usage (from repo root):
//!
//! ```bash
//! # Regenerate the committed fixture (deterministic output).
//! cargo run -p formula-io --example generate_standard_rc4_ooxml_fixture -- \
//!   fixtures/encrypted/ooxml/plaintext.xlsx \
//!   fixtures/encrypted/ooxml/standard-rc4.xlsx
//! ```
//!
//! Notes:
//! - This produces an OLE/CFB container (it is **not** a ZIP on disk).
//! - The password is hard-coded to `password` to match other encrypted OOXML fixtures.
//! - The salt and verifier plaintext are deterministic for stable bytes.

use std::fs;
use std::io::{Cursor, Write as _};
use std::path::PathBuf;

use sha1::{Digest as _, Sha1};

const PASSWORD: &str = "password";

// CryptoAPI algorithm identifiers (WinCrypt).
const CALG_RC4: u32 = 0x0000_6801;
const CALG_SHA1: u32 = 0x0000_8004;
const PROV_RSA_FULL: u32 = 1;
const ENCRYPTION_HEADER_F_CRYPTOAPI: u32 = 0x0000_0004;

const SPIN_COUNT: u32 = 50_000;
const RC4_BLOCK_SIZE: usize = 0x200;

fn password_utf16le_bytes(password: &str) -> Vec<u8> {
    let mut out = Vec::with_capacity(password.len().saturating_mul(2));
    for cu in password.encode_utf16() {
        out.extend_from_slice(&cu.to_le_bytes());
    }
    out
}

fn sha1(data: &[u8]) -> [u8; 20] {
    Sha1::digest(data).into()
}

fn spun_password_hash_sha1(password: &str, salt: &[u8]) -> [u8; 20] {
    let pw = password_utf16le_bytes(password);

    // h = sha1(salt || pw)
    let mut hasher = Sha1::new();
    hasher.update(salt);
    hasher.update(&pw);
    let mut h: [u8; 20] = hasher.finalize().into();

    // for i in 0..SPIN_COUNT-1: h = sha1(LE32(i) || h)
    let mut buf = [0u8; 4 + 20];
    for i in 0..SPIN_COUNT {
        buf[..4].copy_from_slice(&i.to_le_bytes());
        buf[4..].copy_from_slice(&h);
        h = sha1(&buf);
    }

    h
}

fn derive_rc4_key(h: &[u8; 20], block_index: u32, key_len: usize) -> [u8; 20] {
    assert!(key_len <= 20);
    let mut hasher = Sha1::new();
    hasher.update(h);
    hasher.update(block_index.to_le_bytes());
    hasher.finalize().into()
}

#[derive(Clone)]
struct Rc4 {
    s: [u8; 256],
    i: u8,
    j: u8,
}

impl Rc4 {
    fn new(key: &[u8]) -> Self {
        assert!(!key.is_empty());
        let mut s = [0u8; 256];
        for (i, b) in s.iter_mut().enumerate() {
            *b = i as u8;
        }
        let mut j = 0u8;
        for i in 0..256u16 {
            let si = s[i as usize];
            j = j.wrapping_add(si).wrapping_add(key[i as usize % key.len()]);
            s.swap(i as usize, j as usize);
        }
        Self { s, i: 0, j: 0 }
    }

    fn next_byte(&mut self) -> u8 {
        self.i = self.i.wrapping_add(1);
        self.j = self.j.wrapping_add(self.s[self.i as usize]);
        self.s.swap(self.i as usize, self.j as usize);
        let idx = self.s[self.i as usize].wrapping_add(self.s[self.j as usize]);
        self.s[idx as usize]
    }

    fn apply_keystream(&mut self, data: &mut [u8]) {
        for b in data {
            *b ^= self.next_byte();
        }
    }
}

fn encrypt_rc4_stream(key: &[u8], plaintext: &[u8]) -> Vec<u8> {
    let mut rc4 = Rc4::new(key);
    let mut out = plaintext.to_vec();
    rc4.apply_keystream(&mut out);
    out
}

fn encrypt_rc4_cryptoapi_per_block(h: &[u8; 20], key_len: usize, plaintext: &[u8]) -> Vec<u8> {
    let mut out = Vec::with_capacity(plaintext.len());
    let mut offset = 0usize;
    let mut block_index = 0u32;
    while offset < plaintext.len() {
        let block_len = (plaintext.len() - offset).min(RC4_BLOCK_SIZE);
        let digest = derive_rc4_key(h, block_index, key_len);
        let key = &digest[..key_len];
        let mut chunk = plaintext[offset..offset + block_len].to_vec();
        let mut rc4 = Rc4::new(key);
        rc4.apply_keystream(&mut chunk);
        out.extend_from_slice(&chunk);
        offset += block_len;
        block_index += 1;
    }
    out
}

fn build_encryption_info_rc4(
    salt: &[u8; 16],
    verifier_plain: &[u8; 16],
    encrypted_verifier: &[u8; 16],
    encrypted_verifier_hash: &[u8; 20],
    key_size_bits: u32,
) -> Vec<u8> {
    let mut out = Vec::new();

    // EncryptionVersionInfo (Standard): 3.2
    out.extend_from_slice(&3u16.to_le_bytes()); // major
    out.extend_from_slice(&2u16.to_le_bytes()); // minor
    out.extend_from_slice(&0u32.to_le_bytes()); // flags (not required for our decryptor)

    // EncryptionHeader
    let csp_name = "Microsoft Enhanced Cryptographic Provider v1.0";
    let mut csp_utf16le = Vec::new();
    for cu in csp_name.encode_utf16() {
        csp_utf16le.extend_from_slice(&cu.to_le_bytes());
    }
    csp_utf16le.extend_from_slice(&0u16.to_le_bytes()); // NUL terminator

    let mut header = Vec::new();
    // `EncryptionHeader.Flags` must include fCryptoAPI for Standard/CryptoAPI encryption.
    header.extend_from_slice(&ENCRYPTION_HEADER_F_CRYPTOAPI.to_le_bytes());
    header.extend_from_slice(&0u32.to_le_bytes()); // SizeExtra
    header.extend_from_slice(&CALG_RC4.to_le_bytes()); // AlgID
    header.extend_from_slice(&CALG_SHA1.to_le_bytes()); // AlgIDHash
    header.extend_from_slice(&key_size_bits.to_le_bytes()); // KeySize (bits)
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

    // Sanity: verifier_plain is unused here; keep to avoid accidental mismatches when editing.
    let _ = verifier_plain;

    out
}

fn main() {
    let mut args = std::env::args_os();
    let exe = args
        .next()
        .unwrap_or_else(|| std::ffi::OsString::from("generate_standard_rc4_ooxml_fixture"));
    let Some(in_path) = args.next().map(PathBuf::from) else {
        eprintln!(
            "usage: {} <plaintext_ooxml_zip> <out_encrypted_ooxml>",
            exe.to_string_lossy()
        );
        std::process::exit(2);
    };
    let Some(out_path) = args.next().map(PathBuf::from) else {
        eprintln!(
            "usage: {} <plaintext_ooxml_zip> <out_encrypted_ooxml>",
            exe.to_string_lossy()
        );
        std::process::exit(2);
    };
    if args.next().is_some() {
        eprintln!(
            "usage: {} <plaintext_ooxml_zip> <out_encrypted_ooxml>",
            exe.to_string_lossy()
        );
        std::process::exit(2);
    }

    let plaintext = fs::read(&in_path).expect("read plaintext xlsx");

    // Deterministic parameters.
    let salt: [u8; 16] = [
        0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D,
        0x0E, 0x0F,
    ];
    let verifier_plain: [u8; 16] = *b"0123456789ABCDEF";

    let verifier_hash: [u8; 20] = Sha1::digest(verifier_plain).into();

    let h = spun_password_hash_sha1(PASSWORD, &salt);
    let key_len = 16usize;
    let key_size_bits = (key_len * 8) as u32;

    let block0_digest = derive_rc4_key(&h, 0, key_len);
    let rc4_key0 = &block0_digest[..key_len];

    // Encrypt verifier + verifier hash as a single RC4 stream.
    let mut verifier_concat = Vec::new();
    verifier_concat.extend_from_slice(&verifier_plain);
    verifier_concat.extend_from_slice(&verifier_hash);
    let verifier_cipher = encrypt_rc4_stream(rc4_key0, &verifier_concat);
    let encrypted_verifier: [u8; 16] = verifier_cipher[..16].try_into().unwrap();
    let encrypted_verifier_hash: [u8; 20] = verifier_cipher[16..].try_into().unwrap();

    let encryption_info = build_encryption_info_rc4(
        &salt,
        &verifier_plain,
        &encrypted_verifier,
        &encrypted_verifier_hash,
        key_size_bits,
    );

    let ciphertext = encrypt_rc4_cryptoapi_per_block(&h, key_len, &plaintext);
    let mut encrypted_package = Vec::new();
    encrypted_package.extend_from_slice(&(plaintext.len() as u64).to_le_bytes());
    encrypted_package.extend_from_slice(&ciphertext);

    // Write OLE container.
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

    if let Some(parent) = out_path.parent() {
        fs::create_dir_all(parent).expect("create output dir");
    }
    fs::write(&out_path, bytes).expect("write encrypted fixture");
}
