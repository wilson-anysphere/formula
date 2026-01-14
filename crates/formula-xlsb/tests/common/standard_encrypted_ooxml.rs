use std::io::{Cursor, Write as _};

use aes::cipher::{generic_array::GenericArray, BlockEncrypt, KeyInit};
use aes::{Aes128, Aes192, Aes256};
use cbc::cipher::{block_padding::NoPadding, BlockEncryptMut, KeyIvInit};
use formula_offcrypto::{
    standard_derive_key, StandardEncryptionHeader, StandardEncryptionHeaderFlags,
    StandardEncryptionInfo, StandardEncryptionVerifier,
};
use sha1::{Digest as _, Sha1};

const ENCRYPTED_PACKAGE_SEGMENT_LEN: usize = 0x1000;
const AES_BLOCK_LEN: usize = 16;

fn sha1(data: &[u8]) -> [u8; 20] {
    Sha1::digest(data).into()
}

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
    assert_eq!(buf.len() % 16, 0);
    match key.len() {
        16 => {
            let cipher = Aes128::new_from_slice(key).expect("valid AES-128 key");
            for block in buf.chunks_mut(16) {
                cipher.encrypt_block(GenericArray::from_mut_slice(block));
            }
        }
        24 => {
            let cipher = Aes192::new_from_slice(key).expect("valid AES-192 key");
            for block in buf.chunks_mut(16) {
                cipher.encrypt_block(GenericArray::from_mut_slice(block));
            }
        }
        32 => {
            let cipher = Aes256::new_from_slice(key).expect("valid AES-256 key");
            for block in buf.chunks_mut(16) {
                cipher.encrypt_block(GenericArray::from_mut_slice(block));
            }
        }
        _ => panic!("unexpected AES key length"),
    }
}

fn encrypt_segment_aes_cbc_no_padding(
    key: &[u8],
    iv: &[u8; AES_BLOCK_LEN],
    plaintext: &[u8],
) -> Vec<u8> {
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
        _ => panic!("unsupported key length"),
    }
    buf
}

fn encrypt_encrypted_package_stream_standard_cryptoapi(
    key: &[u8],
    salt: &[u8],
    plaintext: &[u8],
) -> Vec<u8> {
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

/// Create bytes for a Standard (CryptoAPI) MS-OFFCRYPTO encrypted OOXML OLE container that wraps the
/// provided plaintext package (ZIP) bytes.
///
/// This is intended for CLI integration tests, so it keeps the EncryptionInfo minimal but
/// well-formed enough for `formula-offcrypto` parsing.
pub fn build_standard_encrypted_ooxml_ole_bytes(package_bytes: &[u8], password: &str) -> Vec<u8> {
    // Deterministic test vectors.
    let salt: [u8; 16] = [
        0xe8, 0x82, 0x66, 0x49, 0x0c, 0x5b, 0xd1, 0xee, 0xbd, 0x2b, 0x43, 0x94, 0xe3, 0xf8, 0x30,
        0xef,
    ];

    let header = StandardEncryptionHeader {
        flags: StandardEncryptionHeaderFlags::from_raw(
            StandardEncryptionHeaderFlags::F_CRYPTOAPI | StandardEncryptionHeaderFlags::F_AES,
        ),
        size_extra: 0,
        alg_id: 0x0000_660E,      // AES-128
        alg_id_hash: 0x0000_8004, // CALG_SHA1
        key_size_bits: 128,
        provider_type: 0,
        reserved1: 0,
        reserved2: 0,
        csp_name: String::new(),
    };

    let verifier_plain: [u8; 16] = [
        0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0a, 0x0b, 0x0c, 0x0d, 0x0e,
        0x0f,
    ];
    let verifier_hash = sha1(&verifier_plain);
    let mut verifier_hash_padded = [0u8; 32];
    verifier_hash_padded[..20].copy_from_slice(&verifier_hash);
    verifier_hash_padded[20..].fill(0xa5);

    let base_info = StandardEncryptionInfo {
        header: header.clone(),
        verifier: StandardEncryptionVerifier {
            salt: Vec::from(salt),
            encrypted_verifier: [0u8; 16],
            verifier_hash_size: 20,
            encrypted_verifier_hash: vec![0u8; 32],
        },
    };
    let key = standard_derive_key(&base_info, password).expect("derive key");

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

    encryption_info_bytes.extend_from_slice(&header.flags.raw.to_le_bytes());
    encryption_info_bytes.extend_from_slice(&header.size_extra.to_le_bytes());
    encryption_info_bytes.extend_from_slice(&header.alg_id.to_le_bytes());
    encryption_info_bytes.extend_from_slice(&header.alg_id_hash.to_le_bytes());
    encryption_info_bytes.extend_from_slice(&header.key_size_bits.to_le_bytes());
    encryption_info_bytes.extend_from_slice(&header.provider_type.to_le_bytes());
    encryption_info_bytes.extend_from_slice(&header.reserved1.to_le_bytes());
    encryption_info_bytes.extend_from_slice(&header.reserved2.to_le_bytes());

    // EncryptionVerifier.
    encryption_info_bytes.extend_from_slice(&(salt.len() as u32).to_le_bytes());
    encryption_info_bytes.extend_from_slice(&salt);
    encryption_info_bytes.extend_from_slice(&encrypted_verifier);
    encryption_info_bytes.extend_from_slice(&(20u32).to_le_bytes());
    encryption_info_bytes.extend_from_slice(&encrypted_verifier_hash);

    // --- Build EncryptedPackage stream bytes ---
    let encrypted_package_bytes =
        encrypt_encrypted_package_stream_standard_cryptoapi(&key, &salt, package_bytes);

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
