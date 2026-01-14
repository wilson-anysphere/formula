use std::io::Cursor;

use aes::cipher::generic_array::GenericArray;
use aes::cipher::{BlockEncrypt, KeyInit};
use sha1::{Digest, Sha1};

use formula_io::offcrypto::decrypt_encrypted_package_standard_aes_to_writer;

const ENCRYPTED_PACKAGE_SEGMENT_LEN: usize = 4096;
const AES_BLOCK_LEN: usize = 16;

fn standard_aes_segment_iv(salt: &[u8], segment_index: u32) -> [u8; 16] {
    let mut hasher = Sha1::new();
    hasher.update(salt);
    hasher.update(segment_index.to_le_bytes());
    let digest = hasher.finalize();

    let mut iv = [0u8; 16];
    iv.copy_from_slice(&digest[..16]);
    iv
}

fn round_up_to_multiple(n: usize, multiple: usize) -> usize {
    (n + multiple - 1) / multiple * multiple
}

fn encrypt_cbc_in_place<C: BlockEncrypt>(cipher: &C, iv: &[u8; 16], buf: &mut [u8]) {
    assert_eq!(buf.len() % AES_BLOCK_LEN, 0);

    let mut prev = *iv;
    for block in buf.chunks_exact_mut(AES_BLOCK_LEN) {
        for (b, p) in block.iter_mut().zip(prev.iter()) {
            *b ^= *p;
        }
        cipher.encrypt_block(GenericArray::from_mut_slice(block));

        prev.copy_from_slice(block);
    }
}

fn encrypt_encrypted_package_standard_aes(plaintext: &[u8], key: &[u8], salt: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    out.extend_from_slice(&(plaintext.len() as u64).to_le_bytes());

    let mut segment_index: u32 = 0;
    let mut offset = 0usize;

    match key.len() {
        16 => {
            let cipher = aes::Aes128::new_from_slice(key).expect("valid AES-128 key length");

            while offset < plaintext.len() {
                let end = (offset + ENCRYPTED_PACKAGE_SEGMENT_LEN).min(plaintext.len());
                let chunk = &plaintext[offset..end];

                let cipher_len = round_up_to_multiple(chunk.len(), AES_BLOCK_LEN);
                let mut buf = vec![0u8; cipher_len];
                buf[..chunk.len()].copy_from_slice(chunk);

                let iv = standard_aes_segment_iv(salt, segment_index);
                encrypt_cbc_in_place(&cipher, &iv, &mut buf);

                out.extend_from_slice(&buf);

                segment_index += 1;
                offset = end;
            }
        }
        other => panic!("unsupported key length for test fixture: {other}"),
    }

    out
}

fn make_plaintext(len: usize) -> Vec<u8> {
    (0..len).map(|i| (i as u8).wrapping_mul(31).wrapping_add(7)).collect()
}

#[test]
fn decrypts_standard_aes_encrypted_package_across_boundaries() {
    let key = [0x42u8; 16];
    let salt = [0xA5u8; 16];

    for plain_len in [1usize, 16, 4096, 4097] {
        let plaintext = make_plaintext(plain_len);
        let encrypted = encrypt_encrypted_package_standard_aes(&plaintext, &key, &salt);

        let mut out = Vec::new();
        let bytes_written = decrypt_encrypted_package_standard_aes_to_writer(
            Cursor::new(encrypted),
            &key,
            &salt,
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
    let salt = [0x22u8; 16];

    let plaintext = make_plaintext(4096);
    let mut encrypted = encrypt_encrypted_package_standard_aes(&plaintext, &key, &salt);

    // Add trailing garbage that is *not* block aligned. The decryptor must not read it.
    encrypted.extend_from_slice(&[0xDE, 0xAD, 0xBE]);

    let mut out = Vec::new();
    let bytes_written = decrypt_encrypted_package_standard_aes_to_writer(
        Cursor::new(encrypted),
        &key,
        &salt,
        &mut out,
    )
    .expect("decrypt");

    assert_eq!(bytes_written, plaintext.len() as u64);
    assert_eq!(out, plaintext);
}
