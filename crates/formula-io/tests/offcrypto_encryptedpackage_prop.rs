use aes::cipher::generic_array::GenericArray;
use aes::cipher::{BlockEncrypt, KeyInit};
use aes::{Aes128, Aes192, Aes256};
use formula_io::offcrypto::decrypt_standard_encrypted_package_stream;
use proptest::prelude::*;

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

proptest! {
    #![proptest_config(ProptestConfig {
        cases: 64,
        max_shrink_iters: 0,
        .. ProptestConfig::default()
    })]

    #[test]
    fn prop_encryptedpackage_roundtrip(plaintext in proptest::collection::vec(any::<u8>(), 0..=20_000)) {
        let ciphertext = encrypt_standard_encrypted_package_stream(&plaintext, &KEY);
        let decrypted_empty_salt = decrypt_standard_encrypted_package_stream(&ciphertext, &KEY, &[])
            .expect("decrypt(encrypt(pt)) should succeed");
        prop_assert_eq!(decrypted_empty_salt.as_slice(), plaintext.as_slice());

        // The Standard/CryptoAPI AES `EncryptedPackage` ciphertext is AES-ECB (no IV). The verifier
        // salt is only used for key derivation / password verification and must not affect package
        // decryption.
        let decrypted_nonempty_salt =
            decrypt_standard_encrypted_package_stream(&ciphertext, &KEY, &SALT)
                .expect("decrypt should ignore salt for Standard AES-ECB");
        prop_assert_eq!(decrypted_nonempty_salt.as_slice(), plaintext.as_slice());
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
