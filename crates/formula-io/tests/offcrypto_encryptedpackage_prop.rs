use aes::{Aes128, Aes192, Aes256};
use cbc::cipher::block_padding::NoPadding;
use cbc::cipher::{BlockEncryptMut, KeyIvInit};
use formula_io::offcrypto::decrypt_standard_encrypted_package_stream;
use proptest::prelude::*;
use sha1::{Digest, Sha1};

const ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN: usize = 8;
const ENCRYPTED_PACKAGE_SEGMENT_LEN: usize = 0x1000;
const AES_BLOCK_LEN: usize = 16;

const KEY: [u8; 16] = [0x42; 16];
const SALT: [u8; 16] = [0x24; 16];

fn derive_iv(salt: &[u8], segment_index: u32) -> [u8; AES_BLOCK_LEN] {
    let mut hasher = Sha1::new();
    hasher.update(salt);
    hasher.update(segment_index.to_le_bytes());
    let digest = hasher.finalize();
    let mut iv = [0u8; AES_BLOCK_LEN];
    iv.copy_from_slice(&digest[..AES_BLOCK_LEN]);
    iv
}

fn pkcs7_pad(plaintext: &[u8]) -> Vec<u8> {
    let mut out = plaintext.to_vec();
    // PKCS#7 always pads, even if already aligned.
    let pad_len = AES_BLOCK_LEN - (out.len() % AES_BLOCK_LEN);
    out.extend(std::iter::repeat(pad_len as u8).take(pad_len));
    out
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
        other => panic!("unsupported AES key length {other}"),
    }

    buf
}

fn encrypt_standard_encrypted_package_stream(plaintext: &[u8], key: &[u8], salt: &[u8]) -> Vec<u8> {
    let orig_size = plaintext.len() as u64;
    let mut out = Vec::new();
    out.extend_from_slice(&orig_size.to_le_bytes());

    // The EncryptedPackage stream can represent an empty package as just the size prefix.
    if plaintext.is_empty() {
        return out;
    }

    // Pad the *overall* plaintext, then split into 0x1000-byte segments.
    let padded = pkcs7_pad(plaintext);
    for (i, chunk) in padded.chunks(ENCRYPTED_PACKAGE_SEGMENT_LEN).enumerate() {
        let iv = derive_iv(salt, i as u32);
        out.extend_from_slice(&encrypt_segment_aes_cbc_no_padding(key, &iv, chunk));
    }

    out
}

proptest! {
    #![proptest_config(ProptestConfig {
        cases: 64,
        max_shrink_iters: 0,
        .. ProptestConfig::default()
    })]

    #[test]
    fn prop_encryptedpackage_roundtrip(plaintext in proptest::collection::vec(any::<u8>(), 0..=20_000)) {
        let ciphertext = encrypt_standard_encrypted_package_stream(&plaintext, &KEY, &SALT);
        let decrypted = decrypt_standard_encrypted_package_stream(&ciphertext, &KEY, &SALT)
            .expect("decrypt(encrypt(pt)) should succeed");
        prop_assert_eq!(decrypted, plaintext);
    }
}

#[derive(Debug, Clone, Copy)]
enum Corruption {
    FlipHeaderLenHuge,
    TruncateCiphertext,
    MakeCiphertextNotBlockAligned,
    RemoveBytesFromNonFinalSegment,
}

proptest! {
    #![proptest_config(ProptestConfig {
        cases: 128,
        max_shrink_iters: 0,
        .. ProptestConfig::default()
    })]

    #[test]
    fn prop_encryptedpackage_rejects_corruption(
        (plaintext, corruption) in prop_oneof![
            (proptest::collection::vec(any::<u8>(), 0..=20_000), Just(Corruption::FlipHeaderLenHuge)),
            (proptest::collection::vec(any::<u8>(), 1..=20_000), Just(Corruption::TruncateCiphertext)),
            (proptest::collection::vec(any::<u8>(), 0..=20_000), Just(Corruption::MakeCiphertextNotBlockAligned)),
            (proptest::collection::vec(any::<u8>(), ENCRYPTED_PACKAGE_SEGMENT_LEN..=20_000), Just(Corruption::RemoveBytesFromNonFinalSegment)),
        ]
    ) {
        let mut ciphertext = encrypt_standard_encrypted_package_stream(&plaintext, &KEY, &SALT);

        match corruption {
            Corruption::FlipHeaderLenHuge => {
                ciphertext[..ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN].copy_from_slice(&u64::MAX.to_le_bytes());
            }
            Corruption::TruncateCiphertext => {
                // Remove all ciphertext bytes. This should fail for non-empty orig_size.
                ciphertext.truncate(ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN);
            }
            Corruption::MakeCiphertextNotBlockAligned => {
                ciphertext.pop();
            }
            Corruption::RemoveBytesFromNonFinalSegment => {
                // Remove >16 bytes (more than the maximum PKCS#7 padding) so the ciphertext is
                // guaranteed to be too short to reproduce `orig_size`.
                //
                // Remove from the first segment (non-final) to also exercise segment boundary logic.
                let start = ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN + (ENCRYPTED_PACKAGE_SEGMENT_LEN / 2);
                ciphertext.drain(start..start + (2 * AES_BLOCK_LEN));
            }
        }

        let res = decrypt_standard_encrypted_package_stream(&ciphertext, &KEY, &SALT);
        prop_assert!(res.is_err(), "expected decrypt to fail for corruption={corruption:?}");
    }
}
