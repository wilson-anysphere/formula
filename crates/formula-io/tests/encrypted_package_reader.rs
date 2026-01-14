use aes::{Aes128, Aes192, Aes256};
use cbc::Encryptor;
use cbc::cipher::{
    block_padding::{NoPadding, Pkcs7},
    BlockCipher, BlockEncryptMut, KeyIvInit,
};
use sha1::{Digest, Sha1};
use std::io::{Cursor, Read, Seek, SeekFrom};
use std::io::ErrorKind;

use formula_io::StandardAesEncryptedPackageReader;

const SEGMENT_LEN: usize = 0x1000;

fn derive_iv(salt: &[u8], segment_index: u32) -> [u8; 16] {
    let mut hasher = Sha1::new();
    hasher.update(salt);
    hasher.update(&segment_index.to_le_bytes());
    let digest = hasher.finalize();
    let mut iv = [0u8; 16];
    iv.copy_from_slice(&digest[..16]);
    iv
}

fn encrypt_segment(key: &[u8], iv: &[u8; 16], plaintext: &[u8], final_segment: bool) -> Vec<u8> {
    let msg_len = plaintext.len();
    match key.len() {
        16 => {
            let cipher = Encryptor::<Aes128>::new_from_slices(key, iv).expect("valid aes-128 key");
            encrypt_with_cipher(cipher, plaintext, msg_len, final_segment)
        }
        24 => {
            let cipher = Encryptor::<Aes192>::new_from_slices(key, iv).expect("valid aes-192 key");
            encrypt_with_cipher(cipher, plaintext, msg_len, final_segment)
        }
        32 => {
            let cipher = Encryptor::<Aes256>::new_from_slices(key, iv).expect("valid aes-256 key");
            encrypt_with_cipher(cipher, plaintext, msg_len, final_segment)
        }
        other => panic!("unexpected key length {other}"),
    }
}

fn encrypt_with_cipher<C>(
    cipher: Encryptor<C>,
    plaintext: &[u8],
    msg_len: usize,
    final_segment: bool,
) -> Vec<u8>
where
    C: BlockCipher + BlockEncryptMut,
{
    let mut buf = plaintext.to_vec();
    if final_segment {
        // Ensure enough space for PKCS7 padding (up to one full block).
        buf.resize(msg_len + 16, 0);
        let ct_len = cipher
            .encrypt_padded_mut::<Pkcs7>(&mut buf, msg_len)
            .expect("encrypt pkcs7")
            .len();
        buf.truncate(ct_len);
        buf
    } else {
        let ct_len = cipher
            .encrypt_padded_mut::<NoPadding>(&mut buf, msg_len)
            .expect("encrypt nopad")
            .len();
        buf.truncate(ct_len);
        buf
    }
}

fn make_encrypted_package(plaintext: &[u8], key: &[u8], salt: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    out.extend_from_slice(&(plaintext.len() as u64).to_le_bytes());

    let seg_count = if plaintext.is_empty() {
        0usize
    } else {
        (plaintext.len() + SEGMENT_LEN - 1) / SEGMENT_LEN
    };

    for seg_idx in 0..seg_count {
        let start = seg_idx * SEGMENT_LEN;
        let end = ((seg_idx + 1) * SEGMENT_LEN).min(plaintext.len());
        let final_segment = seg_idx + 1 == seg_count;
        let iv = derive_iv(salt, seg_idx as u32);
        let ct = encrypt_segment(key, &iv, &plaintext[start..end], final_segment);
        out.extend_from_slice(&ct);
    }

    out
}

fn make_plaintext(len: usize) -> Vec<u8> {
    // Deterministic, non-compressible-ish bytes.
    (0..len).map(|i| ((i * 31) % 251) as u8).collect()
}

#[test]
fn read_sequential_yields_exact_plaintext() {
    let key = [0x42u8; 32];
    let salt = [0xA5u8; 16];
    let plaintext = make_plaintext(10_000);
    let encrypted = make_encrypted_package(&plaintext, &key, &salt);

    let cursor = Cursor::new(encrypted);
    let mut reader =
        StandardAesEncryptedPackageReader::new(cursor, key.to_vec(), salt.to_vec()).expect("new reader");

    let mut out = Vec::new();
    reader.read_to_end(&mut out).expect("read_to_end");
    assert_eq!(out, plaintext);
    assert_eq!(reader.orig_size(), plaintext.len() as u64);
}

#[test]
fn seek_and_read_across_segment_boundaries() {
    let key = [0x11u8; 16];
    let salt = [0x22u8; 16];
    let plaintext = make_plaintext(9_000);
    let encrypted = make_encrypted_package(&plaintext, &key, &salt);
    let cursor = Cursor::new(encrypted);
    let mut reader =
        StandardAesEncryptedPackageReader::new(cursor, key.to_vec(), salt.to_vec()).expect("new reader");

    // Cross segment boundary: 4090..4110.
    reader.seek(SeekFrom::Start(4090)).expect("seek");
    let mut buf = vec![0u8; 20];
    reader.read_exact(&mut buf).expect("read_exact");
    assert_eq!(&buf, &plaintext[4090..4110]);

    // Middle of stream.
    reader.seek(SeekFrom::Start(5000)).expect("seek");
    let mut buf = vec![0u8; 123];
    reader.read_exact(&mut buf).expect("read_exact");
    assert_eq!(&buf, &plaintext[5000..5123]);

    // Tail via SeekFrom::End.
    reader.seek(SeekFrom::End(-10)).expect("seek end");
    let mut buf = vec![0u8; 10];
    reader.read_exact(&mut buf).expect("read_exact");
    assert_eq!(&buf, &plaintext[plaintext.len() - 10..]);
}

#[test]
fn seeking_beyond_eof_returns_eof_on_read() {
    let key = [0x33u8; 24];
    let salt = [0x44u8; 16];
    let plaintext = make_plaintext(1234);
    let encrypted = make_encrypted_package(&plaintext, &key, &salt);
    let cursor = Cursor::new(encrypted);
    let mut reader =
        StandardAesEncryptedPackageReader::new(cursor, key.to_vec(), salt.to_vec()).expect("new reader");

    reader
        .seek(SeekFrom::Start(plaintext.len() as u64 + 999))
        .expect("seek");
    let mut buf = [0u8; 16];
    let n = reader.read(&mut buf).expect("read");
    assert_eq!(n, 0);

    // SeekFrom::End(0) is EOF.
    reader.seek(SeekFrom::End(0)).expect("seek end");
    let n = reader.read(&mut buf).expect("read");
    assert_eq!(n, 0);
}

#[test]
fn final_segment_can_be_larger_than_4096_due_to_padding() {
    let key = [0x55u8; 32];
    let salt = [0x66u8; 16];
    let plaintext = make_plaintext(SEGMENT_LEN * 2); // exact multiple of 4096
    let encrypted = make_encrypted_package(&plaintext, &key, &salt);

    // Ensure our synthetic fixture actually exercises the "final segment > 4096" case.
    assert!(
        encrypted.len() > 8 + SEGMENT_LEN * 2,
        "expected final ciphertext segment to include PKCS7 padding"
    );

    let cursor = Cursor::new(encrypted);
    let mut reader =
        StandardAesEncryptedPackageReader::new(cursor, key.to_vec(), salt.to_vec()).expect("new reader");

    let mut out = Vec::new();
    reader.read_to_end(&mut out).expect("read_to_end");
    assert_eq!(out.len(), plaintext.len());
    assert_eq!(out, plaintext);
}

#[test]
fn errors_on_truncated_size_prefix() {
    let key = [0x42u8; 16];
    let salt = [0xA5u8; 16];

    let encrypted = vec![0u8; 7]; // < 8-byte u64le prefix
    let cursor = Cursor::new(encrypted);
    let err = StandardAesEncryptedPackageReader::new(cursor, key.to_vec(), salt.to_vec())
        .expect_err("expected truncated size prefix to error");
    assert_eq!(err.kind(), ErrorKind::InvalidData);
}

#[test]
fn errors_on_truncated_non_final_ciphertext_segment() {
    let key = [0x11u8; 16];
    let salt = [0x22u8; 16];
    let plaintext = make_plaintext(SEGMENT_LEN + 1); // 2 segments
    let mut encrypted = make_encrypted_package(&plaintext, &key, &salt);

    // Truncate inside the first (non-final) ciphertext segment so it is < 4096 bytes.
    encrypted.truncate(8 + SEGMENT_LEN - 1);

    let cursor = Cursor::new(encrypted);
    let mut reader =
        StandardAesEncryptedPackageReader::new(cursor, key.to_vec(), salt.to_vec()).expect("new reader");

    let mut out = Vec::new();
    let err = reader
        .read_to_end(&mut out)
        .expect_err("expected truncated ciphertext segment to error");
    assert_eq!(err.kind(), ErrorKind::InvalidData);
}

#[test]
fn errors_on_final_segment_ciphertext_not_block_aligned() {
    let key = [0x99u8; 16];
    let salt = [0x77u8; 16];

    let mut encrypted = Vec::new();
    encrypted.extend_from_slice(&1u64.to_le_bytes()); // orig_size = 1
    encrypted.extend_from_slice(&[0u8; 17]); // not a multiple of 16

    let cursor = Cursor::new(encrypted);
    let mut reader =
        StandardAesEncryptedPackageReader::new(cursor, key.to_vec(), salt.to_vec()).expect("new reader");

    let mut buf = [0u8; 8];
    let err = reader
        .read(&mut buf)
        .expect_err("expected non-block-aligned ciphertext to error");
    assert_eq!(err.kind(), ErrorKind::InvalidData);
}

#[test]
fn errors_when_final_segment_ciphertext_is_block_aligned_but_too_short() {
    let key = [0x55u8; 16];
    let salt = [0x66u8; 16];

    // orig_size = 33, but ciphertext only contains 32 bytes (block-aligned).
    let mut encrypted = Vec::new();
    encrypted.extend_from_slice(&33u64.to_le_bytes());
    encrypted.extend_from_slice(&[0u8; 32]);

    let cursor = Cursor::new(encrypted);
    let mut reader =
        StandardAesEncryptedPackageReader::new(cursor, key.to_vec(), salt.to_vec()).expect("new reader");

    let mut out = Vec::new();
    let err = reader
        .read_to_end(&mut out)
        .expect_err("expected too-short final ciphertext segment to error");
    assert_eq!(err.kind(), ErrorKind::InvalidData);
}

#[test]
fn errors_on_u64_max_orig_size_without_panicking() {
    let key = [0xAAu8; 16];
    let salt = [0xBBu8; 16];

    // u64::MAX should not overflow segment count math.
    let mut encrypted = Vec::new();
    encrypted.extend_from_slice(&u64::MAX.to_le_bytes());
    // No ciphertext bytes: should be treated as truncated/corrupt, not as "0 segments".

    let cursor = Cursor::new(encrypted);
    let mut reader =
        StandardAesEncryptedPackageReader::new(cursor, key.to_vec(), salt.to_vec()).expect("new reader");

    let mut buf = [0u8; 1];
    let err = reader
        .read(&mut buf)
        .expect_err("expected truncated ciphertext to error");
    assert_eq!(err.kind(), ErrorKind::InvalidData);
}
