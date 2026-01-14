#![cfg(not(target_arch = "wasm32"))]

use std::io::Cursor;

use aes::cipher::generic_array::GenericArray;
use aes::cipher::{BlockEncrypt, KeyInit};
use aes::{Aes128, Aes192, Aes256};

use formula_io::offcrypto::decrypt_encrypted_package_standard_aes_to_writer;
use proptest::prelude::*;

const ENCRYPTED_PACKAGE_SEGMENT_LEN: usize = 4096;
const AES_BLOCK_LEN: usize = 16;

fn encrypt_encrypted_package_standard_aes_ecb(plaintext: &[u8], key: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    out.extend_from_slice(&(plaintext.len() as u64).to_le_bytes());

    if plaintext.is_empty() {
        return out;
    }

    let mut buf = plaintext.to_vec();
    let rem = buf.len() % AES_BLOCK_LEN;
    if rem != 0 {
        buf.resize(buf.len() + (AES_BLOCK_LEN - rem), 0);
    }

    fn encrypt_with<C: BlockEncrypt + KeyInit>(key: &[u8], buf: &mut [u8]) {
        let cipher = C::new_from_slice(key).expect("valid AES key length");
        for block in buf.chunks_exact_mut(AES_BLOCK_LEN) {
            cipher.encrypt_block(GenericArray::from_mut_slice(block));
        }
    }

    match key.len() {
        16 => encrypt_with::<Aes128>(key, &mut buf),
        24 => encrypt_with::<Aes192>(key, &mut buf),
        32 => encrypt_with::<Aes256>(key, &mut buf),
        other => panic!("unsupported key length for test fixture: {other}"),
    }

    out.extend_from_slice(&buf);
    out
}

fn make_plaintext(len: usize) -> Vec<u8> {
    (0..len).map(|i| (i as u8).wrapping_mul(31).wrapping_add(7)).collect()
}

fn key_strategy() -> impl Strategy<Value = Vec<u8>> {
    prop_oneof![
        Just(vec![0x11u8; 16]),
        Just(vec![0x22u8; 24]),
        Just(vec![0x33u8; 32]),
    ]
}

proptest! {
    #![proptest_config(ProptestConfig {
        cases: 64,
        max_shrink_iters: 0,
        .. ProptestConfig::default()
    })]

    #[test]
    fn prop_streaming_decrypt_roundtrip_with_trailing_bytes(
        plaintext in proptest::collection::vec(any::<u8>(), 0..=20_000),
        key in key_strategy(),
        trailing_len in 0usize..=64,
        trailing_byte in any::<u8>(),
    ) {
        let mut encrypted = encrypt_encrypted_package_standard_aes_ecb(&plaintext, &key);
        encrypted.extend(std::iter::repeat(trailing_byte).take(trailing_len));

        let mut out = Vec::new();
        let bytes_written = decrypt_encrypted_package_standard_aes_to_writer(
            Cursor::new(encrypted),
            &key,
            &[],
            &mut out,
        )
        .expect("decrypt");

        prop_assert_eq!(bytes_written, plaintext.len() as u64);
        prop_assert_eq!(out, plaintext);
    }
}

proptest! {
    #![proptest_config(ProptestConfig {
        cases: 64,
        max_shrink_iters: 0,
        .. ProptestConfig::default()
    })]

    #[test]
    fn prop_streaming_decrypt_rejects_truncation(
        plaintext in proptest::collection::vec(any::<u8>(), 1..=20_000),
        key in key_strategy(),
        remove_bytes in 1usize..=64,
    ) {
        let mut encrypted = encrypt_encrypted_package_standard_aes_ecb(&plaintext, &key);
        let new_len = encrypted.len().saturating_sub(remove_bytes);
        encrypted.truncate(new_len);

        let mut out = Vec::new();
        let res = decrypt_encrypted_package_standard_aes_to_writer(
            Cursor::new(encrypted),
            &key,
            &[],
            &mut out,
        );

        prop_assert!(res.is_err());
    }
}

#[test]
fn decrypts_standard_aes_encrypted_package_across_boundaries() {
    let key = [0x42u8; 16];

    for plain_len in [
        1usize,
        AES_BLOCK_LEN,
        ENCRYPTED_PACKAGE_SEGMENT_LEN,
        ENCRYPTED_PACKAGE_SEGMENT_LEN + 1,
    ] {
        let plaintext = make_plaintext(plain_len);
        let encrypted = encrypt_encrypted_package_standard_aes_ecb(&plaintext, &key);

        let mut out = Vec::new();
        let bytes_written = decrypt_encrypted_package_standard_aes_to_writer(
            Cursor::new(encrypted),
            &key,
            &[],
            &mut out,
        )
        .expect("decrypt");

        assert_eq!(bytes_written, plaintext.len() as u64);
        assert_eq!(out, plaintext);
    }
}

#[test]
fn stops_at_orig_size_and_ignores_trailing_ciphertext() {
    let key = [0x11u8; 16];

    let plaintext = make_plaintext(ENCRYPTED_PACKAGE_SEGMENT_LEN);
    let mut encrypted = encrypt_encrypted_package_standard_aes_ecb(&plaintext, &key);

    // Add trailing garbage that is *not* block aligned. The decryptor must not read it.
    encrypted.extend_from_slice(&[0xDE, 0xAD, 0xBE]);

    let mut out = Vec::new();
    let bytes_written = decrypt_encrypted_package_standard_aes_to_writer(
        Cursor::new(encrypted),
        &key,
        &[],
        &mut out,
    )
    .expect("decrypt");

    assert_eq!(bytes_written, plaintext.len() as u64);
    assert_eq!(out, plaintext);
}

#[test]
fn decrypts_standard_aes_encrypted_package_when_size_prefix_high_dword_is_reserved() {
    let key = [0x42u8; 16];
    let plaintext = make_plaintext(1234);
    let mut encrypted = encrypt_encrypted_package_standard_aes_ecb(&plaintext, &key);

    // Mutate the size prefix to mimic producers that store it as (u32 size, u32 reserved).
    encrypted[4..8].copy_from_slice(&1u32.to_le_bytes());

    let mut out = Vec::new();
    let bytes_written = decrypt_encrypted_package_standard_aes_to_writer(
        Cursor::new(encrypted),
        &key,
        &[],
        &mut out,
    )
    .expect("decrypt");
    assert_eq!(bytes_written, plaintext.len() as u64);
    assert_eq!(out, plaintext);
}
