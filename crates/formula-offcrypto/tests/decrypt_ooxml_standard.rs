#![cfg(not(target_arch = "wasm32"))]

use aes::cipher::{generic_array::GenericArray, BlockEncrypt, KeyInit};
use aes::{Aes128, Aes192, Aes256};
use cbc::Encryptor;
use cipher::{block_padding::NoPadding, BlockEncryptMut, KeyIvInit};
use formula_offcrypto::{
    decrypt_ooxml_standard, standard_derive_key, OffcryptoError, StandardEncryptionHeader,
    StandardEncryptionHeaderFlags, StandardEncryptionInfo, StandardEncryptionVerifier,
};
use sha1::{Digest as _, Sha1};
use std::io::{Cursor, Read, Write};

const CALG_AES_128: u32 = 0x0000_660E;
const CALG_SHA1: u32 = 0x0000_8004;
const STANDARD_SALT: [u8; 16] = [
    0x10, 0x11, 0x12, 0x13, 0x14, 0x15, 0x16, 0x17, 0x18, 0x19, 0x1a, 0x1b, 0x1c, 0x1d, 0x1e,
    0x1f,
];

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

fn build_standard_encryption_info_and_key(password: &str) -> (Vec<u8>, Vec<u8>) {
    let mut info = StandardEncryptionInfo {
        header: StandardEncryptionHeader {
            // MS-OFFCRYPTO Standard encryption must set `fCryptoAPI`, and because we declare an AES
            // `algId` we must also set `fAES` to satisfy strict header validation.
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
        },
        verifier: StandardEncryptionVerifier {
            salt: STANDARD_SALT.to_vec(),
            encrypted_verifier: [0u8; 16],
            verifier_hash_size: 20,
            encrypted_verifier_hash: vec![0u8; 32],
        },
    };

    let key = standard_derive_key(&info, password).expect("derive key");

    // Build a verifier that will validate for this password/key.
    let verifier_plain: [u8; 16] = *b"formula-std-test";
    let verifier_hash: [u8; 20] = Sha1::digest(&verifier_plain).into();

    let mut verifier_hash_padded = [0u8; 32];
    verifier_hash_padded[..20].copy_from_slice(&verifier_hash);

    let mut encrypted_verifier = verifier_plain;
    aes_ecb_encrypt_in_place(&key, &mut encrypted_verifier);

    let mut encrypted_verifier_hash = verifier_hash_padded;
    aes_ecb_encrypt_in_place(&key, &mut encrypted_verifier_hash);

    info.verifier.encrypted_verifier = encrypted_verifier;
    info.verifier.encrypted_verifier_hash = encrypted_verifier_hash.to_vec();

    // Serialize Standard (3.2) EncryptionInfo.
    let mut out = Vec::new();
    out.extend_from_slice(&3u16.to_le_bytes()); // major
    out.extend_from_slice(&2u16.to_le_bytes()); // minor
    out.extend_from_slice(&0u32.to_le_bytes()); // flags

    let mut header_bytes = Vec::new();
    header_bytes.extend_from_slice(&info.header.flags.raw.to_le_bytes());
    header_bytes.extend_from_slice(&info.header.size_extra.to_le_bytes());
    header_bytes.extend_from_slice(&info.header.alg_id.to_le_bytes());
    header_bytes.extend_from_slice(&info.header.alg_id_hash.to_le_bytes());
    header_bytes.extend_from_slice(&info.header.key_size_bits.to_le_bytes());
    header_bytes.extend_from_slice(&info.header.provider_type.to_le_bytes());
    header_bytes.extend_from_slice(&info.header.reserved1.to_le_bytes());
    header_bytes.extend_from_slice(&info.header.reserved2.to_le_bytes());
    // `csp_name` is empty; omit UTF-16 bytes entirely.

    out.extend_from_slice(&(header_bytes.len() as u32).to_le_bytes());
    out.extend_from_slice(&header_bytes);

    // EncryptionVerifier
    out.extend_from_slice(&16u32.to_le_bytes()); // saltSize
    out.extend_from_slice(&info.verifier.salt);
    out.extend_from_slice(&info.verifier.encrypted_verifier);
    out.extend_from_slice(&info.verifier.verifier_hash_size.to_le_bytes()); // verifierHashSize
    out.extend_from_slice(&info.verifier.encrypted_verifier_hash);

    (out, key)
}

fn encrypt_encrypted_package_ecb(key: &[u8], plaintext: &[u8]) -> Vec<u8> {
    let total_size = plaintext.len() as u64;
    let mut out = Vec::new();
    out.extend_from_slice(&total_size.to_le_bytes());

    let mut padded = plaintext.to_vec();
    let rem = padded.len() % 16;
    if rem != 0 {
        padded.resize(padded.len() + (16 - rem), 0);
    }
    aes_ecb_encrypt_in_place(key, &mut padded);
    out.extend_from_slice(&padded);

    out
}

fn derive_standard_segment_iv(salt: &[u8], segment_index: u32) -> [u8; 16] {
    let mut h = Sha1::new();
    h.update(salt);
    h.update(&segment_index.to_le_bytes());
    let digest = h.finalize();
    let mut iv = [0u8; 16];
    iv.copy_from_slice(&digest[..16]);
    iv
}

fn aes_cbc_encrypt_in_place(key: &[u8], iv: &[u8; 16], buf: &mut [u8]) {
    assert_eq!(buf.len() % 16, 0);

    let len = buf.len();
    match key.len() {
        16 => {
            let enc = Encryptor::<Aes128>::new_from_slices(key, iv).expect("key/iv");
            enc.encrypt_padded_mut::<NoPadding>(buf, len)
                .expect("AES-CBC encrypt");
        }
        24 => {
            let enc = Encryptor::<Aes192>::new_from_slices(key, iv).expect("key/iv");
            enc.encrypt_padded_mut::<NoPadding>(buf, len)
                .expect("AES-CBC encrypt");
        }
        32 => {
            let enc = Encryptor::<Aes256>::new_from_slices(key, iv).expect("key/iv");
            enc.encrypt_padded_mut::<NoPadding>(buf, len)
                .expect("AES-CBC encrypt");
        }
        _ => panic!("unexpected AES key length"),
    }
}

fn encrypt_encrypted_package_cbc_segmented(key: &[u8], salt: &[u8], plaintext: &[u8]) -> Vec<u8> {
    const SEGMENT_LEN: usize = 0x1000;
    let total_size = plaintext.len() as u64;

    let mut out = Vec::new();
    out.extend_from_slice(&total_size.to_le_bytes());

    for (idx, chunk) in plaintext.chunks(SEGMENT_LEN).enumerate() {
        let mut padded = chunk.to_vec();
        let rem = padded.len() % 16;
        if rem != 0 {
            padded.resize(padded.len() + (16 - rem), 0);
        }
        let iv = derive_standard_segment_iv(salt, idx as u32);
        aes_cbc_encrypt_in_place(key, &iv, &mut padded);
        out.extend_from_slice(&padded);
    }

    out
}

fn build_tiny_zip() -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Stored);
    zip.start_file("hello.txt", options)
        .expect("start zip entry");
    zip.write_all(b"hello").expect("write zip entry");
    let cursor = zip.finish().expect("finish zip");
    cursor.into_inner()
}

#[test]
fn decrypt_ooxml_standard_roundtrips_zip() {
    let password = "Password1234_";
    let (encryption_info, key) = build_standard_encryption_info_and_key(password);

    let zip_bytes = build_tiny_zip();
    assert_eq!(&zip_bytes[..2], b"PK");

    let encrypted_package = encrypt_encrypted_package_ecb(&key, &zip_bytes);
    let decrypted =
        decrypt_ooxml_standard(&encryption_info, &encrypted_package, password).expect("decrypt");

    assert_eq!(decrypted, zip_bytes);

    // Additional sanity: ensure the decrypted bytes form a valid ZIP we can open.
    let mut archive = zip::ZipArchive::new(Cursor::new(&decrypted)).expect("open zip");
    let mut file = archive.by_name("hello.txt").expect("open file");
    let mut contents = String::new();
    file.read_to_string(&mut contents).expect("read file");
    assert_eq!(contents, "hello");
}

#[test]
fn decrypt_ooxml_standard_supports_cbc_segmented_encryptedpackage_variant() {
    // Some producers encrypt Standard/CryptoAPI `EncryptedPackage` using per-segment AES-CBC with an
    // IV derived from the verifier salt and segment index:
    // `iv_i = SHA1(salt || LE32(i))[0..16]`.
    //
    // `formula-offcrypto` is expected to auto-detect ECB vs CBC-segmented by probing the first
    // decrypted segment for a ZIP signature.
    let password = "Password1234_";
    let (encryption_info, key) = build_standard_encryption_info_and_key(password);

    let zip_bytes = build_tiny_zip();
    assert_eq!(&zip_bytes[..2], b"PK");

    let encrypted_package =
        encrypt_encrypted_package_cbc_segmented(&key, &STANDARD_SALT, &zip_bytes);
    let decrypted = decrypt_ooxml_standard(&encryption_info, &encrypted_package, password)
        .expect("decrypt CBC-segmented Standard EncryptedPackage");

    assert_eq!(decrypted, zip_bytes);
}

#[test]
fn decrypt_ooxml_standard_invalid_pk_returns_invalid_password() {
    let password = "Password1234_";
    let (encryption_info, key) = build_standard_encryption_info_and_key(password);

    let plaintext = b"NOTAZIP".to_vec();
    let encrypted_package = encrypt_encrypted_package_ecb(&key, &plaintext);

    let err = decrypt_ooxml_standard(&encryption_info, &encrypted_package, password)
        .expect_err("expected PK sanity check to fail");
    assert_eq!(err, OffcryptoError::InvalidPassword);
}
