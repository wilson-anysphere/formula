use aes::{Aes128, Aes192, Aes256};
use cbc::cipher::{block_padding::NoPadding, BlockEncryptMut, KeyIvInit};
use formula_offcrypto::{agile_decrypt_package, AgileEncryptionInfo, HashAlgorithm, OffcryptoError};
use sha1::Sha1;
use sha2::{Digest, Sha256, Sha384, Sha512};

const SEGMENT_LENGTH: usize = 4096;

fn fill_deterministic(bytes: &mut [u8]) {
    // Xorshift32.
    let mut x = 0x5EED_1234u32;
    for b in bytes {
        x ^= x << 13;
        x ^= x >> 17;
        x ^= x << 5;
        *b = x as u8;
    }
}

fn segment_iv(hash: HashAlgorithm, salt: &[u8], segment_index: u32) -> [u8; 16] {
    let index_bytes = segment_index.to_le_bytes();
    let mut iv = [0u8; 16];

    match hash {
        HashAlgorithm::Sha1 => {
            let mut hasher = Sha1::new();
            hasher.update(salt);
            hasher.update(index_bytes);
            let digest = hasher.finalize();
            iv.copy_from_slice(&digest[..16]);
        }
        HashAlgorithm::Sha256 => {
            let mut hasher = Sha256::new();
            hasher.update(salt);
            hasher.update(index_bytes);
            let digest = hasher.finalize();
            iv.copy_from_slice(&digest[..16]);
        }
        HashAlgorithm::Sha384 => {
            let mut hasher = Sha384::new();
            hasher.update(salt);
            hasher.update(index_bytes);
            let digest = hasher.finalize();
            iv.copy_from_slice(&digest[..16]);
        }
        HashAlgorithm::Sha512 => {
            let mut hasher = Sha512::new();
            hasher.update(salt);
            hasher.update(index_bytes);
            let digest = hasher.finalize();
            iv.copy_from_slice(&digest[..16]);
        }
    }

    iv
}

fn aes_cbc_encrypt(secret_key: &[u8], iv: &[u8; 16], plaintext: &[u8]) -> Vec<u8> {
    assert!(plaintext.len().is_multiple_of(16));
    let mut buf = plaintext.to_vec();
    let len = buf.len();

    match secret_key.len() {
        16 => {
            let encryptor = cbc::Encryptor::<Aes128>::new_from_slices(secret_key, iv).unwrap();
            encryptor
                .encrypt_padded_mut::<NoPadding>(&mut buf, len)
                .unwrap();
        }
        24 => {
            let encryptor = cbc::Encryptor::<Aes192>::new_from_slices(secret_key, iv).unwrap();
            encryptor
                .encrypt_padded_mut::<NoPadding>(&mut buf, len)
                .unwrap();
        }
        32 => {
            let encryptor = cbc::Encryptor::<Aes256>::new_from_slices(secret_key, iv).unwrap();
            encryptor
                .encrypt_padded_mut::<NoPadding>(&mut buf, len)
                .unwrap();
        }
        _ => panic!("unsupported key length"),
    }

    buf
}

fn minimal_agile_info(hash: HashAlgorithm) -> AgileEncryptionInfo {
    AgileEncryptionInfo {
        key_data_salt: b"unit test salt value".to_vec(),
        key_data_hash_algorithm: hash,
        key_data_block_size: 16,

        encrypted_hmac_key: Vec::new(),
        encrypted_hmac_value: Vec::new(),

        spin_count: 1,
        password_salt: Vec::new(),
        password_hash_algorithm: HashAlgorithm::Sha256,
        password_key_bits: 256,
        encrypted_key_value: Vec::new(),
        encrypted_verifier_hash_input: Vec::new(),
        encrypted_verifier_hash_value: Vec::new(),
    }
}

fn build_encrypted_package(info: &AgileEncryptionInfo, secret_key: &[u8], plaintext: &[u8]) -> Vec<u8> {
    let total_size = plaintext.len() as u64;
    let mut out = Vec::new();
    out.extend_from_slice(&total_size.to_le_bytes());

    for (segment_index, segment) in plaintext.chunks(SEGMENT_LENGTH).enumerate() {
        let iv = segment_iv(info.key_data_hash_algorithm, &info.key_data_salt, segment_index as u32);
        let mut padded = segment.to_vec();
        let pad_len = (16 - (padded.len() % 16)) % 16;
        padded.extend(std::iter::repeat(0u8).take(pad_len));
        let cipher_segment = aes_cbc_encrypt(secret_key, &iv, &padded);
        out.extend_from_slice(&cipher_segment);
    }

    out
}

#[test]
fn agile_decrypt_package_roundtrips_for_all_hashes() {
    let secret_key = [0x42u8; 32]; // AES-256

    let mut plaintext = b"FORMULA_OFFCRYPTO_TEST\0".to_vec();
    let mut tail = vec![0u8; 32 * 1024];
    fill_deterministic(&mut tail);
    plaintext.extend_from_slice(&tail);

    for alg in [
        HashAlgorithm::Sha1,
        HashAlgorithm::Sha256,
        HashAlgorithm::Sha384,
        HashAlgorithm::Sha512,
    ] {
        let info = minimal_agile_info(alg);
        let encrypted = build_encrypted_package(&info, &secret_key, &plaintext);
        let decrypted = agile_decrypt_package(&info, &secret_key, &encrypted).expect("decrypt");
        assert_eq!(decrypted, plaintext, "hash algorithm {alg:?}");
    }
}

#[test]
fn agile_decrypt_package_rejects_short_header() {
    let info = minimal_agile_info(HashAlgorithm::Sha1);
    let secret_key = [0u8; 16];
    let err = agile_decrypt_package(&info, &secret_key, &[0u8; 7]).unwrap_err();
    assert!(matches!(err, OffcryptoError::Truncated { .. }));
}

#[test]
fn agile_decrypt_package_rejects_implausible_size() {
    let info = minimal_agile_info(HashAlgorithm::Sha1);
    let secret_key = [0u8; 16];
    let mut encrypted = vec![0u8; 8 + 16];
    encrypted[0..8].copy_from_slice(&u64::MAX.to_le_bytes());
    let err = agile_decrypt_package(&info, &secret_key, &encrypted).unwrap_err();
    assert!(
        matches!(
            &err,
            OffcryptoError::EncryptedPackageSizeOverflow { total_size }
                if *total_size == u64::MAX
        ),
        "expected EncryptedPackageSizeOverflow(u64::MAX), got {err:?}"
    );
}
