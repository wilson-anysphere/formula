use aes::cipher::generic_array::GenericArray;
use aes::cipher::{BlockEncrypt, KeyInit};
use aes::{Aes128, Aes192, Aes256};
use formula_io::offcrypto::decrypt_standard_encrypted_package_stream;
use proptest::prelude::*;
use sha1::{Digest as _, Sha1};

const ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN: usize = 8;
const AES_BLOCK_LEN: usize = 16;
const ENCRYPTED_PACKAGE_SEGMENT_LEN: usize = 0x1000;

const KEY: [u8; 16] = [0x42; 16];
const SALT: [u8; 16] = [0x24; 16];

fn aes_ecb_encrypt_in_place(key: &[u8], buf: &mut [u8]) {
    fn encrypt_with<C>(key: &[u8], buf: &mut [u8])
    where
        C: BlockEncrypt + KeyInit,
    {
        let cipher = C::new_from_slice(key).expect("key length validated by caller");
        for block in buf.chunks_mut(AES_BLOCK_LEN) {
            cipher.encrypt_block(GenericArray::from_mut_slice(block));
        }
    }

    assert!(
        buf.len() % AES_BLOCK_LEN == 0,
        "plaintext must be block-aligned for AES-ECB"
    );
    match key.len() {
        16 => encrypt_with::<Aes128>(key, buf),
        24 => encrypt_with::<Aes192>(key, buf),
        32 => encrypt_with::<Aes256>(key, buf),
        other => panic!("unsupported AES key length {other}"),
    }
}

fn encrypt_standard_encrypted_package_stream(plaintext: &[u8], key: &[u8]) -> Vec<u8> {
    let orig_size = plaintext.len() as u64;
    let mut out = Vec::new();
    out.extend_from_slice(&orig_size.to_le_bytes());

    // The EncryptedPackage stream can represent an empty package as just the size prefix.
    if plaintext.is_empty() {
        return out;
    }

    // Standard/CryptoAPI AES `EncryptedPackage` uses AES-ECB (no IV). Ciphertext is padded to a
    // whole number of AES blocks, and consumers truncate the decrypted plaintext to `orig_size`.
    let mut buf = plaintext.to_vec();
    let rem = buf.len() % AES_BLOCK_LEN;
    if rem != 0 {
        buf.resize(buf.len() + (AES_BLOCK_LEN - rem), 0);
    }
    aes_ecb_encrypt_in_place(key, &mut buf);
    out.extend_from_slice(&buf);
    out
}

fn derive_standard_cryptoapi_iv_sha1(salt: &[u8], segment_index: u32) -> [u8; AES_BLOCK_LEN] {
    let mut hasher = Sha1::new();
    hasher.update(salt);
    hasher.update(segment_index.to_le_bytes());
    let digest = hasher.finalize();
    let mut iv = [0u8; AES_BLOCK_LEN];
    iv.copy_from_slice(&digest[..AES_BLOCK_LEN]);
    iv
}

fn aes_cbc_encrypt_in_place(key: &[u8], iv: &[u8; AES_BLOCK_LEN], buf: &mut [u8]) {
    assert!(
        buf.len() % AES_BLOCK_LEN == 0,
        "plaintext must be block-aligned for AES-CBC"
    );

    fn encrypt_with<C>(key: &[u8], iv: &[u8; AES_BLOCK_LEN], buf: &mut [u8])
    where
        C: BlockEncrypt + KeyInit,
    {
        let cipher = C::new_from_slice(key).expect("valid AES key length");
        let mut prev = *iv;

        for block in buf.chunks_exact_mut(AES_BLOCK_LEN) {
            for (b, p) in block.iter_mut().zip(prev.iter()) {
                *b ^= p;
            }
            cipher.encrypt_block(GenericArray::from_mut_slice(block));
            prev.copy_from_slice(block);
        }
    }

    match key.len() {
        16 => encrypt_with::<Aes128>(key, iv, buf),
        24 => encrypt_with::<Aes192>(key, iv, buf),
        32 => encrypt_with::<Aes256>(key, iv, buf),
        other => panic!("unsupported AES key length {other} (expected 16/24/32)"),
    }
}

fn pkcs7_pad(mut buf: Vec<u8>) -> Vec<u8> {
    let pad_len = AES_BLOCK_LEN - (buf.len() % AES_BLOCK_LEN);
    buf.extend(std::iter::repeat(pad_len as u8).take(pad_len));
    buf
}

fn encrypt_standard_encrypted_package_stream_segmented_cbc_pkcs7(
    plaintext: &[u8],
    key: &[u8],
    salt: &[u8],
) -> Vec<u8> {
    let orig_size = plaintext.len() as u64;
    let mut out = Vec::new();
    out.extend_from_slice(&orig_size.to_le_bytes());

    if plaintext.is_empty() {
        return out;
    }

    // A best-effort compatibility scheme observed in some non-Excel producers:
    // - Ciphertext is AES-CBC, but CBC chaining resets every 0x1000-byte segment.
    // - For segment i, IV = Truncate16(SHA1(salt || LE32(i))).
    // - PKCS#7 padding is applied to the overall plaintext before segmentation so the ciphertext
    //   is always AES-block-aligned.
    let padded = pkcs7_pad(plaintext.to_vec());
    for (segment_index, segment) in padded.chunks(ENCRYPTED_PACKAGE_SEGMENT_LEN).enumerate() {
        let iv = derive_standard_cryptoapi_iv_sha1(salt, segment_index as u32);
        let mut buf = segment.to_vec();
        aes_cbc_encrypt_in_place(key, &iv, &mut buf);
        out.extend_from_slice(&buf);
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
        let ciphertext = encrypt_standard_encrypted_package_stream(&plaintext, &KEY);
        let decrypted = decrypt_standard_encrypted_package_stream(&ciphertext, &KEY, &[])
            .expect("decrypt(encrypt(pt)) should succeed");
        prop_assert_eq!(decrypted, plaintext);
    }
}

proptest! {
    #![proptest_config(ProptestConfig {
        cases: 32,
        max_shrink_iters: 0,
        .. ProptestConfig::default()
    })]

    #[test]
    fn prop_encryptedpackage_roundtrip_ecb_with_salt_prefers_zip_prefix(
        rest in proptest::collection::vec(any::<u8>(), 0..=19_996),
        extra_blocks in 0usize..=4,
    ) {
        // When `salt` is present, the decryptor may try multiple candidate modes and select the
        // most plausible output (preferring a ZIP prefix). Real-world EncryptedPackage payloads
        // are always OOXML ZIP bytes, so ensure we exercise that mode-selection behavior.
        let mut plaintext = Vec::with_capacity(4 + rest.len());
        plaintext.extend_from_slice(b"PK\x03\x04");
        plaintext.extend_from_slice(&rest);

        let mut ciphertext = encrypt_standard_encrypted_package_stream(&plaintext, &KEY);
        ciphertext.extend(std::iter::repeat(0u8).take(extra_blocks * AES_BLOCK_LEN));

        let decrypted = decrypt_standard_encrypted_package_stream(&ciphertext, &KEY, &SALT)
            .expect("decrypt(encrypt(pt)) should succeed");
        prop_assert_eq!(decrypted, plaintext);
    }
}

proptest! {
    #![proptest_config(ProptestConfig {
        cases: 64,
        max_shrink_iters: 0,
        .. ProptestConfig::default()
    })]

    #[test]
    fn prop_encryptedpackage_roundtrip_segmented_cbc_pkcs7(
        plaintext in proptest::collection::vec(any::<u8>(), 0..=20_000),
    ) {
        let ciphertext = encrypt_standard_encrypted_package_stream_segmented_cbc_pkcs7(
            &plaintext,
            &KEY,
            &SALT,
        );
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
            // Keep payload large enough that we can safely remove bytes from the middle of the ciphertext.
            (proptest::collection::vec(any::<u8>(), 4_096..=20_000), Just(Corruption::RemoveBytesFromNonFinalSegment)),
        ]
    ) {
        let mut ciphertext = encrypt_standard_encrypted_package_stream(&plaintext, &KEY);

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
                // Remove two AES blocks so ciphertext remains block-aligned but is guaranteed to be
                // too short to reproduce `orig_size` (accounting for AES block padding).
                //
                // Remove from the first segment (non-final) to also exercise segment boundary logic.
                let start = ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN + (ENCRYPTED_PACKAGE_SEGMENT_LEN / 2);
                ciphertext.drain(start..start + (2 * AES_BLOCK_LEN));
            }
        }

        let res = decrypt_standard_encrypted_package_stream(&ciphertext, &KEY, &[]);
        prop_assert!(res.is_err(), "expected decrypt to fail for corruption={corruption:?}");
    }
}

proptest! {
    #![proptest_config(ProptestConfig {
        cases: 128,
        max_shrink_iters: 0,
        .. ProptestConfig::default()
    })]

    #[test]
    fn prop_encryptedpackage_rejects_corruption_segmented_cbc_pkcs7(
        (plaintext, corruption) in prop_oneof![
            (proptest::collection::vec(any::<u8>(), 0..=20_000), Just(Corruption::FlipHeaderLenHuge)),
            (proptest::collection::vec(any::<u8>(), 1..=20_000), Just(Corruption::TruncateCiphertext)),
            (proptest::collection::vec(any::<u8>(), 0..=20_000), Just(Corruption::MakeCiphertextNotBlockAligned)),
            (proptest::collection::vec(any::<u8>(), ENCRYPTED_PACKAGE_SEGMENT_LEN..=20_000), Just(Corruption::RemoveBytesFromNonFinalSegment)),
        ]
    ) {
        let mut ciphertext = encrypt_standard_encrypted_package_stream_segmented_cbc_pkcs7(
            &plaintext,
            &KEY,
            &SALT,
        );

        match corruption {
            Corruption::FlipHeaderLenHuge => {
                ciphertext[..ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN].copy_from_slice(&u64::MAX.to_le_bytes());
            }
            Corruption::TruncateCiphertext => {
                ciphertext.truncate(ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN);
            }
            Corruption::MakeCiphertextNotBlockAligned => {
                ciphertext.pop();
            }
            Corruption::RemoveBytesFromNonFinalSegment => {
                let start = ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN + (ENCRYPTED_PACKAGE_SEGMENT_LEN / 2);
                ciphertext.drain(start..start + (2 * AES_BLOCK_LEN));
            }
        }

        let res = decrypt_standard_encrypted_package_stream(&ciphertext, &KEY, &SALT);
        prop_assert!(res.is_err(), "expected decrypt to fail for corruption={corruption:?}");
    }
}
