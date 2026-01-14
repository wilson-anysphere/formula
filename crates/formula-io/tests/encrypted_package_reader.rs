#![cfg(not(target_arch = "wasm32"))]

use aes::{Aes128, Aes192, Aes256};
use cbc::Encryptor;
use cbc::cipher::{
    block_padding::{NoPadding, Pkcs7},
    BlockCipher, BlockEncryptMut, KeyIvInit,
};
use proptest::prelude::*;
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

#[derive(Debug, Clone)]
enum ReadSeekOp {
    SeekStart(u64),
    SeekCurrent(i64),
    SeekEnd(i64),
    Read(usize),
}

fn read_seek_op_strategy() -> impl Strategy<Value = ReadSeekOp> {
    // Keep the ranges conservative so the test is fast and avoids exploring pathological i64
    // corner cases. The goal is to stress the segmented-ciphertext + caching logic across a wide
    // variety of random access patterns.
    let seek_abs_max = 25_000u64;
    let seek_rel_max = 25_000i64;
    let read_len_max = SEGMENT_LEN * 2;

    prop_oneof![
        (0u64..=seek_abs_max).prop_map(ReadSeekOp::SeekStart),
        (-seek_rel_max..=seek_rel_max).prop_map(ReadSeekOp::SeekCurrent),
        (-seek_rel_max..=seek_rel_max).prop_map(ReadSeekOp::SeekEnd),
        (0usize..=read_len_max).prop_map(ReadSeekOp::Read),
    ]
}

proptest! {
    #![proptest_config(ProptestConfig {
        cases: 32,
        max_shrink_iters: 0,
        failure_persistence: None,
        .. ProptestConfig::default()
    })]

    #[test]
    fn prop_reader_seek_read_matches_plaintext(
        plaintext in proptest::collection::vec(any::<u8>(), 0..=20_000),
        ops in proptest::collection::vec(read_seek_op_strategy(), 0..=64),
    ) {
        let key = [0x11u8; 32];
        let salt = [0x22u8; 16];

        let encrypted = make_encrypted_package(&plaintext, &key, &salt);
        let cursor = Cursor::new(encrypted);
        let mut reader = StandardAesEncryptedPackageReader::new(cursor, key.to_vec(), salt.to_vec())
            .expect("new reader");

        let mut expected_pos: u64 = 0;
        let plaintext_len = plaintext.len() as u64;

        for op in ops {
            match op {
                ReadSeekOp::SeekStart(pos) => {
                    let res = reader.seek(SeekFrom::Start(pos));
                    prop_assert!(res.is_ok());
                    expected_pos = pos;
                }
                ReadSeekOp::SeekCurrent(off) => {
                    let new_pos = expected_pos as i128 + off as i128;
                    let res = reader.seek(SeekFrom::Current(off));
                    if new_pos < 0 {
                        prop_assert!(res.is_err());
                    } else {
                        prop_assert_eq!(res.unwrap(), new_pos as u64);
                        expected_pos = new_pos as u64;
                    }
                }
                ReadSeekOp::SeekEnd(off) => {
                    let new_pos = plaintext_len as i128 + off as i128;
                    let res = reader.seek(SeekFrom::End(off));
                    if new_pos < 0 {
                        prop_assert!(res.is_err());
                    } else {
                        prop_assert_eq!(res.unwrap(), new_pos as u64);
                        expected_pos = new_pos as u64;
                    }
                }
                ReadSeekOp::Read(len) => {
                    let mut buf = vec![0u8; len];
                    let n = reader.read(&mut buf).expect("read should not error");

                    let expected_n = if expected_pos >= plaintext_len {
                        0usize
                    } else {
                        let remaining = (plaintext_len - expected_pos) as usize;
                        remaining.min(len)
                    };
                    prop_assert_eq!(n, expected_n);

                    if n > 0 {
                        prop_assert_eq!(
                            &buf[..n],
                            &plaintext[expected_pos as usize..expected_pos as usize + n]
                        );
                        expected_pos += n as u64;
                    }
                }
            }
        }
    }
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
fn reader_accepts_reserved_high_dword_in_size_prefix() {
    let key = [0x42u8; 32];
    let salt = [0xA5u8; 16];
    let plaintext = make_plaintext(10_000);
    let mut encrypted = make_encrypted_package(&plaintext, &key, &salt);
    encrypted[4..8].copy_from_slice(&1u32.to_le_bytes());

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
    let err = StandardAesEncryptedPackageReader::new(cursor, key.to_vec(), salt.to_vec())
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
    let err = StandardAesEncryptedPackageReader::new(cursor, key.to_vec(), salt.to_vec())
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
    let err = StandardAesEncryptedPackageReader::new(cursor, key.to_vec(), salt.to_vec())
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
    let err = StandardAesEncryptedPackageReader::new(cursor, key.to_vec(), salt.to_vec())
        .expect_err("expected u64::MAX orig_size to error");
    assert_eq!(err.kind(), ErrorKind::InvalidData);
    assert!(
        err.to_string().contains("orig_size"),
        "unexpected error message: {err}"
    );
}

#[test]
fn errors_when_orig_size_requires_a_missing_final_ciphertext_segment() {
    let key = [0x11u8; 16];
    let salt = [0x22u8; 16];
    let plaintext = make_plaintext(SEGMENT_LEN + 1); // 2 segments
    let mut encrypted = make_encrypted_package(&plaintext, &key, &salt);

    // Drop the entire final ciphertext segment.
    //
    // The `orig_size` prefix is attacker-controlled; the reader should reject inputs where the
    // declared plaintext length is implausible for the available ciphertext bytes.
    encrypted.truncate(8 + SEGMENT_LEN);

    let cursor = Cursor::new(encrypted);
    let err = StandardAesEncryptedPackageReader::new(cursor, key.to_vec(), salt.to_vec())
        .expect_err("expected missing final ciphertext segment to error");
    assert_eq!(err.kind(), ErrorKind::InvalidData);

    // Simulate a stream that reports the full ciphertext length via `Seek` but returns EOF when
    // reading the second segment. This exercises the reader's "return partial bytes, surface the
    // error on the next call" behavior (common for `Read` adapters).
    #[derive(Debug)]
    struct TruncatedRead<R> {
        inner: R,
        truncate_at: u64,
    }

    impl<R: Read + Seek> Read for TruncatedRead<R> {
        fn read(&mut self, buf: &mut [u8]) -> std::io::Result<usize> {
            let pos = self.inner.seek(SeekFrom::Current(0))?;
            if pos >= self.truncate_at || buf.is_empty() {
                return Ok(0);
            }
            let remaining = (self.truncate_at - pos) as usize;
            let n = remaining.min(buf.len());
            self.inner.read(&mut buf[..n])
        }
    }

    impl<R: Read + Seek> Seek for TruncatedRead<R> {
        fn seek(&mut self, pos: SeekFrom) -> std::io::Result<u64> {
            self.inner.seek(pos)
        }
    }

    let encrypted = make_encrypted_package(&plaintext, &key, &salt);
    let cursor = Cursor::new(encrypted);
    let cursor = TruncatedRead {
        inner: cursor,
        truncate_at: (8 + SEGMENT_LEN) as u64,
    };
    let mut reader = StandardAesEncryptedPackageReader::new(cursor, key.to_vec(), salt.to_vec())
        .expect("new reader");

    let mut buf = vec![0u8; SEGMENT_LEN + 100];
    let n = reader.read(&mut buf).expect("read");
    assert_eq!(n, SEGMENT_LEN);
    assert_eq!(&buf[..n], &plaintext[..SEGMENT_LEN]);

    let err = reader.read(&mut buf).expect_err("expected error on follow-up read");
    assert_eq!(err.kind(), ErrorKind::InvalidData);
}
