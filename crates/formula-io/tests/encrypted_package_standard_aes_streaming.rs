use std::io::Cursor;

use aes::cipher::generic_array::GenericArray;
use aes::cipher::{BlockEncrypt, KeyInit};
use aes::{Aes128, Aes192, Aes256};

use formula_io::offcrypto::decrypt_encrypted_package_standard_aes_to_writer;

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
