use std::io::{Read, Write};

use super::cryptoapi::{
    final_hash,
    hash_password_fixed_spin,
    password_to_utf16le,
    HashAlg as CryptoApiHashAlg,
};

use aes::cipher::{generic_array::GenericArray, BlockDecrypt, KeyInit};
use aes::{Aes128, Aes192, Aes256};

use crate::rc4_encrypted_package::{
    parse_rc4_encrypted_package_stream, Rc4EncryptedPackageParseError,
    Rc4EncryptedPackageParseOptions,
};

const ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN: usize = 8;
const AES_BLOCK_LEN: usize = 16;
const ENCRYPTED_PACKAGE_SEGMENT_LEN: usize = 0x1000;
const ENCRYPTED_PACKAGE_RC4_BLOCK_LEN: usize = 0x200;

#[derive(Debug, Clone, PartialEq, Eq, thiserror::Error)]
pub enum InvalidCiphertextLenReason {
    #[error("segment index/offset overflow")]
    SegmentIndexOverflow,
}

/// Errors returned by [`decrypt_standard_encrypted_package_stream`].
///
/// These errors aim to provide enough context (segment index + offset/length) to debug corrupt
/// `EncryptedPackage` streams without needing to re-run with a debugger.
#[derive(Debug, thiserror::Error, PartialEq, Eq)]
pub enum EncryptedPackageDecryptError {
    #[error("truncated `EncryptedPackage` size prefix: expected {ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN} bytes, got {len}")]
    TruncatedPrefix { len: usize },

    #[error("invalid AES key length {len} (expected 16, 24, or 32 bytes)")]
    InvalidKeyLength { len: usize },

    #[error("invalid ciphertext length for segment {segment} at offset {offset}: len={len} ({reason})")]
    InvalidCiphertextLen {
        segment: u32,
        offset: usize,
        len: usize,
        reason: InvalidCiphertextLenReason,
    },

    #[error("truncated ciphertext segment {segment} at offset {offset}: expected {expected} bytes, got {got}")]
    TruncatedSegment {
        segment: u32,
        offset: usize,
        expected: usize,
        got: usize,
    },

    #[error("ciphertext segment {segment} at offset {offset} is not AES-block aligned: len={len}")]
    CiphertextNotBlockAligned { segment: u32, offset: usize, len: usize },

    #[error("`EncryptedPackage` orig_size {orig_size} is too large for ciphertext length {ciphertext_len}")]
    OrigSizeTooLarge { orig_size: u64, ciphertext_len: usize },

    #[error("crypto error while decrypting segment {segment} at offset {offset}")]
    CryptoError { segment: u32, offset: usize },
}

/// Errors returned by the legacy RC4 `EncryptedPackage` helpers in this module.
#[derive(Debug, thiserror::Error, PartialEq, Eq)]
pub enum EncryptedPackageError {
    #[error(
        "`EncryptedPackage` stream is too short: expected at least {ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN} bytes, got {len}"
    )]
    StreamTooShort { len: usize },

    #[error(
        "`EncryptedPackage` orig_size {orig_size} does not fit into the current platform `usize`"
    )]
    OrigSizeTooLargeForPlatform { orig_size: u64 },

    #[error(
        "`EncryptedPackage` orig_size {orig_size} is implausibly large for ciphertext length {ciphertext_len}"
    )]
    ImplausibleOrigSize {
        orig_size: u64,
        ciphertext_len: usize,
    },

    #[error("`EncryptedPackage` orig_size {orig_size} exceeds configured maximum {max}")]
    OrigSizeExceedsMax { orig_size: u64, max: u64 },

    #[error("`EncryptedPackage` ciphertext length {ciphertext_len} is not a multiple of AES block size ({AES_BLOCK_LEN})")]
    CiphertextLenNotBlockAligned { ciphertext_len: usize },

    #[error(
        "`EncryptedPackage` AES key length must be 16, 24, or 32 bytes (AES-128/192/256), got {key_len}"
    )]
    InvalidAesKeyLength { key_len: usize },

    #[error(
        "decrypted plaintext is shorter than expected: got {decrypted_len} bytes, expected at least {orig_size}"
    )]
    DecryptedTooShort {
        decrypted_len: usize,
        orig_size: u64,
    },

    #[error(
        "`EncryptedPackage` RC4 key length must be non-zero and must not exceed the hash digest length, got {key_len}"
    )]
    Rc4InvalidKeyLength { key_len: usize },

    #[error(
        "unsupported `EncryptionHeader.algIdHash` {alg_id_hash:#010x} for RC4 (supported: CALG_SHA1=0x00008004, CALG_MD5=0x00008003)"
    )]
    Rc4UnsupportedHashAlgorithm { alg_id_hash: u32 },
}

/// Errors returned by [`decrypt_encrypted_package_standard_aes_to_writer`].
#[derive(Debug, thiserror::Error)]
pub enum EncryptedPackageToWriterError {
    #[error(transparent)]
    Crypto(#[from] EncryptedPackageError),

    #[error("failed to read `EncryptedPackage` original size prefix: {source}")]
    ReadOrigSize {
        #[source]
        source: std::io::Error,
    },

    #[error(
        "failed to read `EncryptedPackage` ciphertext for segment {segment_index} (needed {needed} bytes): {source}"
    )]
    ReadCiphertext {
        segment_index: u32,
        needed: usize,
        #[source]
        source: std::io::Error,
    },

    #[error("failed to write decrypted `EncryptedPackage` bytes: {source}")]
    WritePlaintext {
        #[source]
        source: std::io::Error,
    },
}

fn padded_aes_len(len: usize) -> usize {
    // `len` is at most 4096 bytes, so this cannot overflow.
    let rem = len % AES_BLOCK_LEN;
    if rem == 0 {
        len
    } else {
        len + (AES_BLOCK_LEN - rem)
    }
}

/// Decrypt a Standard (CryptoAPI) AES `EncryptedPackage` stream to an arbitrary writer.
///
/// This implements the baseline MS-OFFCRYPTO/ECMA-376 behavior for Standard AES `EncryptedPackage`:
/// **AES-ECB** (no IV).
///
/// For baseline AES-ECB framing + truncation rules, see `docs/offcrypto-standard-encryptedpackage.md`.
///
/// This API is **sequential**:
/// - It does **not** require `Seek`.
/// - It never allocates a buffer proportional to the package size.
/// - It stops after writing exactly `orig_size` bytes (truncating any padding).
///
/// The caller must provide:
/// - `key`: the AES key bytes (16/24/32 bytes for AES-128/192/256).
/// - `salt`: unused for Standard AES-ECB package decryption (accepted for API compatibility).
///
/// Returns the number of plaintext bytes written (always `orig_size` on success).
pub fn decrypt_encrypted_package_standard_aes_to_writer<R: Read, W: Write>(
    mut reader: R,
    key: &[u8],
    _salt: &[u8],
    mut out: W,
) -> Result<u64, EncryptedPackageToWriterError> {
    let mut size_bytes = [0u8; ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN];
    reader
        .read_exact(&mut size_bytes)
        .map_err(|source| EncryptedPackageToWriterError::ReadOrigSize { source })?;
    let orig_size = crate::parse_encrypted_package_size_prefix_bytes(size_bytes, None);

    if orig_size == 0 {
        return Ok(0);
    }

    // Validate key length up front so we fail fast even for empty ciphertext.
    if !matches!(key.len(), 16 | 24 | 32) {
        return Err(EncryptedPackageError::InvalidAesKeyLength { key_len: key.len() }.into());
    }

    let mut remaining = orig_size;
    let mut segment_index: u32 = 0;
    let mut scratch = [0u8; ENCRYPTED_PACKAGE_SEGMENT_LEN];

    while remaining > 0 {
        let plain_len = remaining.min(ENCRYPTED_PACKAGE_SEGMENT_LEN as u64) as usize;
        let cipher_len = padded_aes_len(plain_len);

        reader
            .read_exact(&mut scratch[..cipher_len])
            .map_err(|source| EncryptedPackageToWriterError::ReadCiphertext {
                segment_index,
                needed: cipher_len,
                source,
            })?;

        aes_ecb_decrypt_in_place(key, &mut scratch[..cipher_len])?;

        out.write_all(&scratch[..plain_len])
            .map_err(|source| EncryptedPackageToWriterError::WritePlaintext { source })?;

        remaining -= plain_len as u64;
        segment_index = segment_index.wrapping_add(1);
    }

    Ok(orig_size)
}

/// Decrypt an MS-OFFCRYPTO "Standard" (CryptoAPI) `EncryptedPackage` stream.
///
/// The caller must provide:
/// - `key`: the file encryption key (AES-128/192/256), derived from the password and
///   `EncryptionInfo`.
///
/// Algorithm summary:
/// - First 8 bytes are the original plaintext size (`orig_size`) as a little-endian `u64`.
/// - Remaining bytes are AES ciphertext (AES-ECB, no IV).
/// - The concatenated plaintext is truncated to `orig_size` (do **not** rely on PKCS#7 unpadding).
/// - `salt` is unused for Standard AES `EncryptedPackage` decryption (accepted for API compatibility).
pub fn decrypt_standard_encrypted_package_stream(
    encrypted_package_stream: &[u8],
    key: &[u8],
    _salt: &[u8],
) -> Result<Vec<u8>, EncryptedPackageDecryptError> {
    if encrypted_package_stream.len() < ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN {
        return Err(EncryptedPackageDecryptError::TruncatedPrefix {
            len: encrypted_package_stream.len(),
        });
    }

    if !matches!(key.len(), 16 | 24 | 32) {
        return Err(EncryptedPackageDecryptError::InvalidKeyLength { len: key.len() });
    }

    let mut size_bytes = [0u8; ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN];
    size_bytes.copy_from_slice(&encrypted_package_stream[..ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN]);
    let ciphertext = &encrypted_package_stream[ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN..];
    let ciphertext_len = ciphertext.len();
    let orig_size = crate::parse_encrypted_package_size_prefix_bytes(size_bytes, Some(ciphertext_len as u64));
    if ciphertext_len == 0 && orig_size == 0 {
        return Ok(Vec::new());
    }

    if orig_size > ciphertext_len as u64 {
        return Err(EncryptedPackageDecryptError::OrigSizeTooLarge {
            orig_size,
            ciphertext_len,
        });
    }

    let orig_size_usize = usize::try_from(orig_size).map_err(|_| {
        EncryptedPackageDecryptError::OrigSizeTooLarge {
            orig_size,
            ciphertext_len,
        }
    })?;
    if orig_size_usize == 0 {
        // Be permissive: treat non-empty ciphertext as trailing padding bytes.
        return Ok(Vec::new());
    }

    let segment_count = orig_size_usize.div_ceil(ENCRYPTED_PACKAGE_SEGMENT_LEN);
    let full_segments = segment_count.saturating_sub(1);
    let full_cipher_len = full_segments.checked_mul(ENCRYPTED_PACKAGE_SEGMENT_LEN).ok_or(
        EncryptedPackageDecryptError::InvalidCiphertextLen {
            segment: 0,
            offset: ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN,
            len: ciphertext_len,
            reason: InvalidCiphertextLenReason::SegmentIndexOverflow,
        },
    )?;

    // --- Segment-level truncation checks (before block-alignment checks) ---
    //
    // Prefer `TruncatedSegment` when the ciphertext is clearly missing bytes, even if the remaining
    // bytes are also not block-aligned. This yields more actionable diagnostics (expected/got).
    if ciphertext_len < full_cipher_len {
        let seg = ciphertext_len / ENCRYPTED_PACKAGE_SEGMENT_LEN;
        let seg_u32 =
            u32::try_from(seg).map_err(|_| EncryptedPackageDecryptError::InvalidCiphertextLen {
                segment: 0,
                offset: ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN,
                len: ciphertext_len,
                reason: InvalidCiphertextLenReason::SegmentIndexOverflow,
            })?;
        let seg_offset = seg * ENCRYPTED_PACKAGE_SEGMENT_LEN;
        let got = ciphertext_len - seg_offset;
        return Err(EncryptedPackageDecryptError::TruncatedSegment {
            segment: seg_u32,
            offset: ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN + seg_offset,
            expected: ENCRYPTED_PACKAGE_SEGMENT_LEN,
            got,
        });
    }

    let last_plain_len = orig_size_usize - full_cipher_len;
    let expected_min_last_cipher_len = padded_aes_len(last_plain_len);
    let last_cipher_available = ciphertext_len - full_cipher_len;
    if last_cipher_available < expected_min_last_cipher_len {
        let seg_u32 = u32::try_from(full_segments).map_err(|_| {
            EncryptedPackageDecryptError::InvalidCiphertextLen {
                segment: 0,
                offset: ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN,
                len: ciphertext_len,
                reason: InvalidCiphertextLenReason::SegmentIndexOverflow,
            }
        })?;
        return Err(EncryptedPackageDecryptError::TruncatedSegment {
            segment: seg_u32,
            offset: ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN + full_cipher_len,
            expected: expected_min_last_cipher_len,
            got: last_cipher_available,
        });
    }

    if ciphertext_len % AES_BLOCK_LEN != 0 {
        let seg = ciphertext_len / ENCRYPTED_PACKAGE_SEGMENT_LEN;
        let seg_u32 =
            u32::try_from(seg).map_err(|_| EncryptedPackageDecryptError::InvalidCiphertextLen {
                segment: 0,
                offset: ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN,
                len: ciphertext_len,
                reason: InvalidCiphertextLenReason::SegmentIndexOverflow,
            })?;
        let seg_offset = seg * ENCRYPTED_PACKAGE_SEGMENT_LEN;
        let seg_len = ciphertext_len - seg_offset;
        return Err(EncryptedPackageDecryptError::CiphertextNotBlockAligned {
            segment: seg_u32,
            offset: ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN + seg_offset,
            len: seg_len,
        });
    }

    fn decrypt_with<C: BlockDecrypt + KeyInit>(
        key: &[u8],
        ciphertext: &[u8],
        orig_size_usize: usize,
        ciphertext_len: usize,
    ) -> Result<Vec<u8>, EncryptedPackageDecryptError> {
        let cipher = C::new_from_slice(key).map_err(|_| EncryptedPackageDecryptError::CryptoError {
            segment: 0,
            offset: ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN,
        })?;
        // Standard/CryptoAPI AES `EncryptedPackage` uses AES-ECB (no IV).

        let mut out = Vec::with_capacity(orig_size_usize);
        let mut segment_index: u32 = 0;
        while out.len() < orig_size_usize {
            let seg_offset = (segment_index as usize)
                .checked_mul(ENCRYPTED_PACKAGE_SEGMENT_LEN)
                .ok_or(EncryptedPackageDecryptError::InvalidCiphertextLen {
                    segment: segment_index,
                    offset: ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN,
                    len: ciphertext_len,
                    reason: InvalidCiphertextLenReason::SegmentIndexOverflow,
                })?;

            let plain_len = (orig_size_usize - out.len()).min(ENCRYPTED_PACKAGE_SEGMENT_LEN);
            let cipher_len = padded_aes_len(plain_len);

            let seg_end = seg_offset.checked_add(cipher_len).ok_or(
                EncryptedPackageDecryptError::InvalidCiphertextLen {
                    segment: segment_index,
                    offset: ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN + seg_offset,
                    len: ciphertext_len,
                    reason: InvalidCiphertextLenReason::SegmentIndexOverflow,
                },
            )?;
            let Some(seg_cipher) = ciphertext.get(seg_offset..seg_end) else {
                let got = ciphertext_len.saturating_sub(seg_offset);
                return Err(EncryptedPackageDecryptError::TruncatedSegment {
                    segment: segment_index,
                    offset: ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN + seg_offset,
                    expected: cipher_len,
                    got,
                });
            };

            let mut decrypted = seg_cipher.to_vec();
            for block in decrypted.chunks_exact_mut(AES_BLOCK_LEN) {
                cipher.decrypt_block(GenericArray::from_mut_slice(block));
            }

            out.extend_from_slice(&decrypted[..plain_len]);

            segment_index = segment_index.checked_add(1).ok_or(
                EncryptedPackageDecryptError::InvalidCiphertextLen {
                    segment: segment_index,
                    offset: ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN + seg_offset,
                    len: ciphertext_len,
                    reason: InvalidCiphertextLenReason::SegmentIndexOverflow,
                },
            )?;
        }

        out.truncate(orig_size_usize);
        Ok(out)
    }

    match key.len() {
        16 => decrypt_with::<Aes128>(key, ciphertext, orig_size_usize, ciphertext_len),
        24 => decrypt_with::<Aes192>(key, ciphertext, orig_size_usize, ciphertext_len),
        32 => decrypt_with::<Aes256>(key, ciphertext, orig_size_usize, ciphertext_len),
        other => Err(EncryptedPackageDecryptError::InvalidKeyLength { len: other }),
    }
}

fn aes_ecb_decrypt_in_place(key: &[u8], buf: &mut [u8]) -> Result<(), EncryptedPackageError> {
    if buf.len() % AES_BLOCK_LEN != 0 {
        return Err(EncryptedPackageError::CiphertextLenNotBlockAligned {
            ciphertext_len: buf.len(),
        });
    }

    fn decrypt_with<C>(key: &[u8], buf: &mut [u8]) -> Result<(), EncryptedPackageError>
    where
        C: BlockDecrypt + KeyInit,
    {
        let cipher = C::new_from_slice(key)
            .map_err(|_| EncryptedPackageError::InvalidAesKeyLength { key_len: key.len() })?;
        for block in buf.chunks_mut(AES_BLOCK_LEN) {
            cipher.decrypt_block(GenericArray::from_mut_slice(block));
        }
        Ok(())
    }

    match key.len() {
        16 => decrypt_with::<Aes128>(key, buf),
        24 => decrypt_with::<Aes192>(key, buf),
        32 => decrypt_with::<Aes256>(key, buf),
        other => Err(EncryptedPackageError::InvalidAesKeyLength { key_len: other }),
    }
}

struct Rc4 {
    s: [u8; 256],
    i: u8,
    j: u8,
}

impl Rc4 {
    fn new(key: &[u8]) -> Self {
        debug_assert!(!key.is_empty(), "RC4 key must be non-empty");
        let mut s = [0u8; 256];
        for (i, v) in s.iter_mut().enumerate() {
            *v = i as u8;
        }
        let mut j: u8 = 0;
        for i in 0..256usize {
            j = j.wrapping_add(s[i]).wrapping_add(key[i % key.len()]);
            s.swap(i, j as usize);
        }
        Self { s, i: 0, j: 0 }
    }

    fn apply_keystream(&mut self, data: &mut [u8]) {
        for b in data {
            self.i = self.i.wrapping_add(1);
            self.j = self.j.wrapping_add(self.s[self.i as usize]);
            self.s.swap(self.i as usize, self.j as usize);
            let t = self.s[self.i as usize].wrapping_add(self.s[self.j as usize]);
            let k = self.s[t as usize];
            *b ^= k;
        }
    }
}

struct CryptoapiRc4EncryptedPackageDecryptor {
    password_hash: Vec<u8>,
    key_len: usize,
    hash_alg: CryptoApiHashAlg,
}

impl CryptoapiRc4EncryptedPackageDecryptor {
    fn new(password: &str, salt: &[u8], key_len: usize) -> Result<Self, EncryptedPackageError> {
        Self::new_with_hash_alg(password, salt, key_len, CryptoApiHashAlg::Sha1)
    }

    fn new_with_hash_alg(
        password: &str,
        salt: &[u8],
        key_len: usize,
        hash_alg: CryptoApiHashAlg,
    ) -> Result<Self, EncryptedPackageError> {
        if key_len == 0 || key_len > hash_alg.hash_len() {
            return Err(EncryptedPackageError::Rc4InvalidKeyLength { key_len });
        }

        let pw_utf16le = password_to_utf16le(password);
        let password_hash = hash_password_fixed_spin(&pw_utf16le, salt, hash_alg);
        Ok(Self {
            password_hash,
            key_len,
            hash_alg,
        })
    }

    fn decrypt_encrypted_package_stream(
        &self,
        encrypted_package_stream: &[u8],
    ) -> Result<Vec<u8>, EncryptedPackageError> {
        let parsed = parse_rc4_encrypted_package_stream(
            encrypted_package_stream,
            &Rc4EncryptedPackageParseOptions::default(),
        )
        .map_err(|err| match err {
            Rc4EncryptedPackageParseError::TruncatedHeader => {
                EncryptedPackageError::StreamTooShort {
                    len: encrypted_package_stream.len(),
                }
            }
            Rc4EncryptedPackageParseError::DeclaredSizeExceedsPayload {
                declared,
                available,
            } => EncryptedPackageError::ImplausibleOrigSize {
                orig_size: declared,
                ciphertext_len: available as usize,
            },
            Rc4EncryptedPackageParseError::DeclaredSizeExceedsMax { declared, max } => {
                EncryptedPackageError::OrigSizeExceedsMax {
                    orig_size: declared,
                    max,
                }
            }
        })?;
        let orig_size = parsed.header.package_size;
        let ciphertext = parsed.encrypted_payload;

        if ciphertext.is_empty() && orig_size == 0 {
            return Ok(Vec::new());
        }

        let orig_size_usize = usize::try_from(orig_size)
            .map_err(|_| EncryptedPackageError::OrigSizeTooLargeForPlatform { orig_size })?;

        let mut out = ciphertext[..orig_size_usize].to_vec();
        for (block_index, chunk) in out.chunks_mut(ENCRYPTED_PACKAGE_RC4_BLOCK_LEN).enumerate() {
            let digest = final_hash(&self.password_hash, block_index as u32, self.hash_alg);
            let mut rc4 = Rc4::new(&digest[..self.key_len]);
            rc4.apply_keystream(chunk);
        }

        Ok(out)
    }
}

/// Decrypt an MS-OFFCRYPTO "Standard" (CryptoAPI) RC4 `EncryptedPackage` stream.
///
/// Stream framing:
/// - first 8 bytes: `orig_size` (`u64le`, plaintext size)
/// - remaining bytes: ciphertext (RC4-encrypted package bytes)
///
/// Unlike the Standard/CryptoAPI AES `EncryptedPackage` variant above (AES-ECB), the RC4 variant uses
/// **0x200-byte blocks** (note: this differs from BIFF8 `FILEPASS` RC4, which re-keys every 0x400 bytes)
/// and derives a fresh RC4 key for each block:
/// - `password_hash = Hash(salt || UTF16LE(password))`
/// - for `i in 0..50000`: `password_hash = Hash(LE32(i) || password_hash)`
/// - `h_i = Hash(password_hash || LE32(i))`
/// - `rc4_key_i = h_i[0..key_len]` where `key_len = keySize/8` (40→5 bytes, 56→7 bytes, 128→16 bytes)
/// - RC4 is **reset** per block (do not carry keystream state across blocks).
///
/// See `docs/offcrypto-standard-cryptoapi-rc4.md` for additional notes and test vectors.
///
/// This helper assumes `EncryptionHeader.algIdHash == CALG_SHA1`. For non-SHA1 Standard RC4 (e.g.
/// `CALG_MD5`), use [`decrypt_standard_cryptoapi_rc4_encrypted_package_stream_with_hash`].
pub fn decrypt_standard_cryptoapi_rc4_encrypted_package_stream(
    encrypted_package_stream: &[u8],
    password: &str,
    salt: &[u8],
    key_len: usize,
) -> Result<Vec<u8>, EncryptedPackageError> {
    CryptoapiRc4EncryptedPackageDecryptor::new(password, salt, key_len)?
        .decrypt_encrypted_package_stream(encrypted_package_stream)
}

/// Decrypt an MS-OFFCRYPTO "Standard" (CryptoAPI) RC4 `EncryptedPackage` stream, using the hash
/// algorithm specified by `EncryptionHeader.algIdHash`.
///
/// `alg_id_hash` must be `CALG_SHA1` or `CALG_MD5`.
pub fn decrypt_standard_cryptoapi_rc4_encrypted_package_stream_with_hash(
    encrypted_package_stream: &[u8],
    password: &str,
    salt: &[u8],
    key_len: usize,
    alg_id_hash: u32,
) -> Result<Vec<u8>, EncryptedPackageError> {
    let hash_alg = CryptoApiHashAlg::from_calg_id(alg_id_hash)
        .map_err(|_| EncryptedPackageError::Rc4UnsupportedHashAlgorithm { alg_id_hash })?;
    CryptoapiRc4EncryptedPackageDecryptor::new_with_hash_alg(password, salt, key_len, hash_alg)?
        .decrypt_encrypted_package_stream(encrypted_package_stream)
}

#[cfg(test)]
mod tests {
    use super::*;
    use aes::cipher::{generic_array::GenericArray, BlockEncrypt, KeyInit};
    use aes::{Aes128, Aes192, Aes256};
    use md5::Md5;
    use sha1::{Digest, Sha1};
    use std::io::Cursor;

    fn fixed_key(len: usize) -> Vec<u8> {
        (0..len).map(|i| i as u8).collect()
    }

    fn fixed_key_16() -> Vec<u8> {
        fixed_key(16)
    }

    fn fixed_key_24() -> Vec<u8> {
        (0u8..=0x17).collect()
    }

    fn fixed_key_32() -> Vec<u8> {
        (0u8..=0x1F).collect()
    }

    fn aes_ecb_encrypt_in_place(key: &[u8], buf: &mut [u8]) {
        assert!(
            buf.len() % AES_BLOCK_LEN == 0,
            "plaintext must be block-aligned for AES-ECB"
        );

        fn encrypt_with<C>(key: &[u8], buf: &mut [u8])
        where
            C: BlockEncrypt + KeyInit,
        {
            let cipher = C::new_from_slice(key).expect("valid key length for AES");
            for block in buf.chunks_mut(AES_BLOCK_LEN) {
                cipher.encrypt_block(GenericArray::from_mut_slice(block));
            }
        }

        match key.len() {
            16 => encrypt_with::<Aes128>(key, buf),
            24 => encrypt_with::<Aes192>(key, buf),
            32 => encrypt_with::<Aes256>(key, buf),
            other => panic!("unsupported AES key length {other}"),
        }
    }

    fn encrypt_encrypted_package_stream_standard_cryptoapi_ecb(
        key: &[u8],
        plaintext: &[u8],
    ) -> Vec<u8> {
        let orig_size = plaintext.len() as u64;

        let mut out = Vec::with_capacity(ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN + plaintext.len());
        out.extend_from_slice(&orig_size.to_le_bytes());

        if plaintext.is_empty() {
            return out;
        }

        // Standard/CryptoAPI AES `EncryptedPackage` uses AES-ECB (no IV). Ciphertext is padded to
        // a whole number of AES blocks, and consumers truncate the decrypted plaintext to the
        // declared `orig_size`.
        let mut buf = plaintext.to_vec();
        let rem = buf.len() % AES_BLOCK_LEN;
        if rem != 0 {
            buf.extend(std::iter::repeat(0u8).take(AES_BLOCK_LEN - rem));
        }
        aes_ecb_encrypt_in_place(key, &mut buf);
        out.extend_from_slice(&buf);

        out
    }

    fn make_plaintext(len: usize) -> Vec<u8> {
        (0..len).map(|i| (i % 251) as u8).collect()
    }

    #[test]
    fn round_trip_decrypts_various_sizes_and_key_lengths() {
        let keys = [fixed_key_16(), fixed_key_24(), fixed_key_32()];

        for key in keys {
            for size in [0usize, 1, 15, 16, 17, 4095, 4096, 4097, 8192 + 123] {
                let plaintext = make_plaintext(size);
                let encrypted =
                    encrypt_encrypted_package_stream_standard_cryptoapi_ecb(&key, &plaintext);
                let decrypted =
                    decrypt_standard_encrypted_package_stream(&encrypted, &key, &[]).unwrap();
                assert_eq!(
                    decrypted,
                    plaintext,
                    "failed for size={size} key_len={}",
                    key.len()
                );
            }
        }
    }

    #[test]
    fn decrypt_truncates_to_orig_size_even_with_trailing_bytes() {
        let key = fixed_key_16();

        let plaintext = make_plaintext(4096);
        let mut encrypted =
            encrypt_encrypted_package_stream_standard_cryptoapi_ecb(&key, &plaintext);

        // Append extra ciphertext blocks beyond what `orig_size` requires.
        encrypted.extend_from_slice(&[0u8; AES_BLOCK_LEN * 3]);

        let decrypted = decrypt_standard_encrypted_package_stream(&encrypted, &key, &[]).unwrap();
        assert_eq!(decrypted, plaintext);
    }

    #[test]
    fn decrypt_empty_ciphertext_and_zero_size_is_ok() {
        let key = fixed_key_16();

        let encrypted = 0u64.to_le_bytes().to_vec();
        let decrypted = decrypt_standard_encrypted_package_stream(&encrypted, &key, &[]).unwrap();
        assert!(decrypted.is_empty());
    }

    #[test]
    fn errors_on_short_stream() {
        let key = fixed_key_16();

        let err = decrypt_standard_encrypted_package_stream(&[0u8; 7], &key, &[]).unwrap_err();
        assert_eq!(err, EncryptedPackageDecryptError::TruncatedPrefix { len: 7 });
    }

    #[test]
    fn errors_on_truncated_ciphertext() {
        let key = fixed_key_16();

        let mut encrypted = Vec::new();
        encrypted.extend_from_slice(&(1u64).to_le_bytes());
        encrypted.extend_from_slice(&[0u8; 15]); // not multiple of 16

        let err = decrypt_standard_encrypted_package_stream(&encrypted, &key, &[]).unwrap_err();
        assert_eq!(
            err,
            EncryptedPackageDecryptError::TruncatedSegment {
                segment: 0,
                offset: 8,
                expected: 16,
                got: 15
            }
        );
    }

    #[test]
    fn errors_when_length_header_exceeds_ciphertext() {
        let key = fixed_key_16();

        // orig_size claims 32 bytes, but we only have 16 bytes of ciphertext.
        let mut encrypted = Vec::new();
        encrypted.extend_from_slice(&(32u64).to_le_bytes());
        encrypted.extend_from_slice(&[0u8; 16]);

        let err = decrypt_standard_encrypted_package_stream(&encrypted, &key, &[]).unwrap_err();
        assert_eq!(
            err,
            EncryptedPackageDecryptError::OrigSizeTooLarge {
                orig_size: 32,
                ciphertext_len: 16,
            }
        );
    }

    #[test]
    fn errors_on_invalid_key_length() {
        let key = vec![0u8; 15];
        let encrypted = 0u64.to_le_bytes().to_vec();

        let err = decrypt_standard_encrypted_package_stream(&encrypted, &key, &[]).unwrap_err();
        assert_eq!(err, EncryptedPackageDecryptError::InvalidKeyLength { len: 15 });
    }

    #[test]
    fn errors_truncated_segment_reports_segment_index() {
        let key = fixed_key_16();

        // Two segments required, but the final segment is 1 byte short of the minimum 16-byte block.
        let mut encrypted = Vec::new();
        encrypted.extend_from_slice(&(4096u64 + 1).to_le_bytes());
        encrypted.extend_from_slice(&vec![0u8; 4096 + 15]);

        let err = decrypt_standard_encrypted_package_stream(&encrypted, &key, &[]).unwrap_err();
        assert!(
            matches!(
                err,
                EncryptedPackageDecryptError::TruncatedSegment {
                    segment: 1,
                    expected: 16,
                    got: 15,
                    ..
                }
            ),
            "unexpected error: {err:?}"
        );
    }

    #[test]
    fn decrypts_real_standard_cryptoapi_fixture_to_valid_zip() {
        let path = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("tests/fixtures/offcrypto_standard_cryptoapi_password.xlsx");
        let file = std::fs::File::open(&path).expect("open fixture");
        let mut ole = cfb::CompoundFile::open(file).expect("parse OLE container");

        let mut encryption_info = Vec::new();
        ole.open_stream("EncryptionInfo")
            .expect("open EncryptionInfo stream")
            .read_to_end(&mut encryption_info)
            .expect("read EncryptionInfo");

        let mut encrypted_package = Vec::new();
        ole.open_stream("EncryptedPackage")
            .expect("open EncryptedPackage stream")
            .read_to_end(&mut encrypted_package)
            .expect("read EncryptedPackage");

        let info = match formula_offcrypto::parse_encryption_info(&encryption_info)
            .expect("parse EncryptionInfo")
        {
            formula_offcrypto::EncryptionInfo::Standard {
                header, verifier, ..
            } => formula_offcrypto::StandardEncryptionInfo { header, verifier },
            other => panic!("expected Standard EncryptionInfo, got {other:?}"),
        };

        let password = "password";
        let key = formula_offcrypto::standard_derive_key(&info, password).expect("derive key");
        formula_offcrypto::standard_verify_key(&info, &key).expect("verify key");

        let decrypted = decrypt_standard_encrypted_package_stream(
            &encrypted_package,
            &key,
            &info.verifier.salt,
        )
        .expect("decrypt EncryptedPackage");
        assert!(
            decrypted.starts_with(b"PK"),
            "expected decrypted package to start with PK, got {:02x?}",
            decrypted.get(..2)
        );

        let mut zip = zip::ZipArchive::new(Cursor::new(decrypted.as_slice()))
            .expect("open ZIP from decrypted bytes");
        zip.by_name("xl/workbook.xml")
            .expect("expected xl/workbook.xml in decrypted ZIP");
    }

    #[test]
    fn errors_ciphertext_not_block_aligned_reports_segment_index() {
        let key = fixed_key_16();

        // Two segments required, but the final segment has 17 bytes (not block aligned).
        let mut encrypted = Vec::new();
        encrypted.extend_from_slice(&(4096u64 + 1).to_le_bytes());
        encrypted.extend_from_slice(&vec![0u8; 4096 + 17]);

        let err = decrypt_standard_encrypted_package_stream(&encrypted, &key, &[]).unwrap_err();
        assert!(
            matches!(
                err,
                EncryptedPackageDecryptError::CiphertextNotBlockAligned {
                    segment: 1,
                    len: 17,
                    ..
                }
            ),
            "unexpected error: {err:?}"
        );
    }

    struct TestRc4 {
        s: [u8; 256],
        i: u8,
        j: u8,
    }

    impl TestRc4 {
        fn new(key: &[u8]) -> Self {
            let mut s = [0u8; 256];
            for (i, v) in s.iter_mut().enumerate() {
                *v = i as u8;
            }
            let mut j: u8 = 0;
            for i in 0..256usize {
                j = j.wrapping_add(s[i]).wrapping_add(key[i % key.len()]);
                s.swap(i, j as usize);
            }
            Self { s, i: 0, j: 0 }
        }

        fn apply_keystream(&mut self, data: &mut [u8]) {
            for b in data {
                self.i = self.i.wrapping_add(1);
                self.j = self.j.wrapping_add(self.s[self.i as usize]);
                self.s.swap(self.i as usize, self.j as usize);
                let t = self.s[self.i as usize].wrapping_add(self.s[self.j as usize]);
                let k = self.s[t as usize];
                *b ^= k;
            }
        }
    }

    fn test_password_utf16le(password: &str) -> Vec<u8> {
        password
            .encode_utf16()
            .flat_map(|cu| cu.to_le_bytes())
            .collect()
    }

    fn test_cryptoapi_password_hash_sha1(password: &str, salt: &[u8]) -> [u8; 20] {
        const SPIN_COUNT: u32 = 50_000;
        let pw = test_password_utf16le(password);

        let mut hasher = Sha1::new();
        hasher.update(salt);
        hasher.update(&pw);
        let mut h: [u8; 20] = hasher.finalize().into();

        let mut buf = [0u8; 4 + 20];
        for i in 0..SPIN_COUNT {
            buf[..4].copy_from_slice(&i.to_le_bytes());
            buf[4..].copy_from_slice(&h);
            h = Sha1::digest(&buf).into();
        }
        h
    }

    fn test_cryptoapi_block_key_sha1(
        password_hash: &[u8; 20],
        block_index: u32,
        key_len: usize,
    ) -> Vec<u8> {
        let mut hasher = Sha1::new();
        hasher.update(password_hash);
        hasher.update(block_index.to_le_bytes());
        let digest: [u8; 20] = hasher.finalize().into();
        digest[..key_len].to_vec()
    }

    fn test_cryptoapi_password_hash_md5(password: &str, salt: &[u8]) -> [u8; 16] {
        use md5::Digest as _;
        const SPIN_COUNT: u32 = 50_000;
        let pw = test_password_utf16le(password);

        let mut hasher = Md5::new();
        hasher.update(salt);
        hasher.update(&pw);
        let mut h: [u8; 16] = hasher.finalize().into();

        let mut buf = [0u8; 4 + 16];
        for i in 0..SPIN_COUNT {
            buf[..4].copy_from_slice(&i.to_le_bytes());
            buf[4..].copy_from_slice(&h);
            h = Md5::digest(&buf).into();
        }
        h
    }

    fn test_cryptoapi_block_key_md5(
        password_hash: &[u8; 16],
        block_index: u32,
        key_len: usize,
    ) -> Vec<u8> {
        use md5::Digest as _;
        let mut hasher = Md5::new();
        hasher.update(password_hash);
        hasher.update(block_index.to_le_bytes());
        let digest: [u8; 16] = hasher.finalize().into();
        digest[..key_len].to_vec()
    }

    struct TestCryptoapiRc4Encryptor {
        password_hash: [u8; 20],
        key_len: usize,
    }

    impl TestCryptoapiRc4Encryptor {
        fn new(password: &str, salt: &[u8], key_len: usize) -> Self {
            assert!(key_len <= 20);
            let password_hash = test_cryptoapi_password_hash_sha1(password, salt);
            Self {
                password_hash,
                key_len,
            }
        }

        fn encrypt_encrypted_package_stream(&self, plaintext: &[u8]) -> Vec<u8> {
            const BLOCK_LEN: usize = 0x200;

            let mut ciphertext = plaintext.to_vec();
            for (block_index, chunk) in ciphertext.chunks_mut(BLOCK_LEN).enumerate() {
                let key = test_cryptoapi_block_key_sha1(
                    &self.password_hash,
                    block_index as u32,
                    self.key_len,
                );
                let mut rc4 = TestRc4::new(&key);
                rc4.apply_keystream(chunk);
            }

            let mut out = Vec::with_capacity(8 + ciphertext.len());
            out.extend_from_slice(&(plaintext.len() as u64).to_le_bytes());
            out.extend_from_slice(&ciphertext);
            out
        }
    }

    struct TestCryptoapiRc4EncryptorMd5 {
        password_hash: [u8; 16],
        key_len: usize,
    }

    impl TestCryptoapiRc4EncryptorMd5 {
        fn new(password: &str, salt: &[u8], key_len: usize) -> Self {
            assert!(key_len <= 16);
            let password_hash = test_cryptoapi_password_hash_md5(password, salt);
            Self {
                password_hash,
                key_len,
            }
        }

        fn encrypt_encrypted_package_stream(&self, plaintext: &[u8]) -> Vec<u8> {
            const BLOCK_LEN: usize = 0x200;

            let mut ciphertext = plaintext.to_vec();
            for (block_index, chunk) in ciphertext.chunks_mut(BLOCK_LEN).enumerate() {
                let key = test_cryptoapi_block_key_md5(
                    &self.password_hash,
                    block_index as u32,
                    self.key_len,
                );
                let mut rc4 = TestRc4::new(&key);
                rc4.apply_keystream(chunk);
            }

            let mut out = Vec::with_capacity(8 + ciphertext.len());
            out.extend_from_slice(&(plaintext.len() as u64).to_le_bytes());
            out.extend_from_slice(&ciphertext);
            out
        }
    }

    fn make_plaintext_pattern(len: usize) -> Vec<u8> {
        (0..len)
            .map(|i| {
                let i = i as u32;
                let x = i
                    .wrapping_mul(31)
                    .wrapping_add(i.rotate_left(13))
                    .wrapping_add(0x9E37_79B9);
                (x ^ (x >> 8) ^ (x >> 16) ^ (x >> 24)) as u8
            })
            .collect()
    }

    #[test]
    fn rc4_cryptoapi_encryptedpackage_block_boundary_regression() {
        let password = "correct horse battery staple";
        let salt: [u8; 16] = [
            0x01, 0x23, 0x45, 0x67, 0x89, 0xAB, 0xCD, 0xEF, 0x10, 0x32, 0x54, 0x76, 0x98, 0xBA,
            0xDC, 0xFE,
        ];
        for key_len in [5usize, 7, 16] {
            let encryptor = TestCryptoapiRc4Encryptor::new(password, &salt, key_len);
            let decryptor = CryptoapiRc4EncryptedPackageDecryptor::new(password, &salt, key_len)
                .expect("decryptor");

            for len in [0usize, 1, 511, 512, 513, 1023, 1024, 1025, 10_000] {
                let plaintext = make_plaintext_pattern(len);
                let encrypted = encryptor.encrypt_encrypted_package_stream(&plaintext);
                let decrypted = decryptor
                    .decrypt_encrypted_package_stream(&encrypted)
                    .expect("decrypt");
                assert_eq!(decrypted, plaintext, "failed for len={len} (key_len={key_len})");
            }
        }
    }

    #[test]
    fn rc4_cryptoapi_encryptedpackage_roundtrip_with_40_bit_key() {
        // Regression: 40-bit CryptoAPI RC4 uses a 5-byte key (no padding).
        let password = "correct horse battery staple";
        let salt: [u8; 16] = [
            0x01, 0x23, 0x45, 0x67, 0x89, 0xAB, 0xCD, 0xEF, 0x10, 0x32, 0x54, 0x76, 0x98, 0xBA,
            0xDC, 0xFE,
        ];
        let key_len = 5;

        let encryptor = TestCryptoapiRc4Encryptor::new(password, &salt, key_len);
        let decryptor = CryptoapiRc4EncryptedPackageDecryptor::new(password, &salt, key_len)
            .expect("decryptor");

        // Ensure we cross at least one 0x200-byte boundary so per-block rekeying is exercised.
        let plaintext = make_plaintext_pattern(10_000);
        let encrypted = encryptor.encrypt_encrypted_package_stream(&plaintext);
        let decrypted = decryptor
            .decrypt_encrypted_package_stream(&encrypted)
            .expect("decrypt");
        assert_eq!(decrypted, plaintext);
    }

    #[test]
    fn rc4_cryptoapi_encryptedpackage_md5_block_boundary_regression() {
        let password = "correct horse battery staple";
        let salt: [u8; 16] = [
            0x01, 0x23, 0x45, 0x67, 0x89, 0xAB, 0xCD, 0xEF, 0x10, 0x32, 0x54, 0x76, 0x98, 0xBA,
            0xDC, 0xFE,
        ];
        let key_len = 16;

        let encryptor = TestCryptoapiRc4EncryptorMd5::new(password, &salt, key_len);
        let decryptor = CryptoapiRc4EncryptedPackageDecryptor::new_with_hash_alg(
            password,
            &salt,
            key_len,
            CryptoApiHashAlg::Md5,
        )
        .expect("decryptor");

        for len in [0usize, 1, 511, 512, 513, 1023, 1024, 1025, 10_000] {
            let plaintext = make_plaintext_pattern(len);
            let encrypted = encryptor.encrypt_encrypted_package_stream(&plaintext);
            let decrypted = decryptor
                .decrypt_encrypted_package_stream(&encrypted)
                .expect("decrypt");
            assert_eq!(decrypted, plaintext, "failed for len={len}");
        }
    }

    #[test]
    fn rc4_cryptoapi_encryptedpackage_md5_roundtrip_with_40_bit_key() {
        // Regression: 40-bit CryptoAPI RC4 uses a padded 16-byte RC4 key, not a raw 5-byte key.
        let password = "correct horse battery staple";
        let salt: [u8; 16] = [
            0x01, 0x23, 0x45, 0x67, 0x89, 0xAB, 0xCD, 0xEF, 0x10, 0x32, 0x54, 0x76, 0x98, 0xBA,
            0xDC, 0xFE,
        ];
        let key_len = 5;

        let encryptor = TestCryptoapiRc4EncryptorMd5::new(password, &salt, key_len);
        let decryptor = CryptoapiRc4EncryptedPackageDecryptor::new_with_hash_alg(
            password,
            &salt,
            key_len,
            CryptoApiHashAlg::Md5,
        )
        .expect("decryptor");

        // Ensure we cross at least one 0x200-byte boundary so per-block rekeying is exercised.
        let plaintext = make_plaintext_pattern(10_000);
        let encrypted = encryptor.encrypt_encrypted_package_stream(&plaintext);
        let decrypted = decryptor
            .decrypt_encrypted_package_stream(&encrypted)
            .expect("decrypt");
        assert_eq!(decrypted, plaintext);
    }

    #[test]
    fn rejects_u64_max_orig_size_without_panicking() {
        let key = fixed_key_16();

        // Tiny `EncryptedPackage`: just the size prefix, no ciphertext.
        let encrypted = u64::MAX.to_le_bytes().to_vec();
        let res = std::panic::catch_unwind(|| {
            decrypt_standard_encrypted_package_stream(&encrypted, &key, &[])
        });
        assert!(
            res.is_ok(),
            "decryptor should not panic on u64::MAX orig_size"
        );
        assert!(
            res.unwrap().is_err(),
            "expected error on u64::MAX orig_size"
        );
    }

    #[test]
    fn rejects_orig_size_larger_than_usize_max_when_possible() {
        // On 64-bit, `usize::MAX == u64::MAX`, so the "+1" test case cannot be represented.
        if usize::BITS >= 64 {
            return;
        }

        let key = fixed_key_16();

        let orig_size = (usize::MAX as u64) + 1;
        let encrypted = orig_size.to_le_bytes().to_vec();
        let res = std::panic::catch_unwind(|| {
            decrypt_standard_encrypted_package_stream(&encrypted, &key, &[])
        });
        assert!(
            res.is_ok(),
            "decryptor should not panic on orig_size > usize::MAX"
        );
        let err = res
            .unwrap()
            .expect_err("expected error on orig_size > usize::MAX");
        assert!(matches!(
            err,
            EncryptedPackageDecryptError::OrigSizeTooLarge { orig_size: _, .. }
        ));
    }

    #[test]
    fn rejects_orig_size_that_would_overflow_naive_segment_math() {
        let key = fixed_key_16();

        // This value would overflow naive `(orig_size + 15)` padding math.
        let orig_size = u64::MAX - 4094;
        let encrypted = orig_size.to_le_bytes().to_vec();

        let res = std::panic::catch_unwind(|| {
            decrypt_standard_encrypted_package_stream(&encrypted, &key, &[])
        });
        assert!(
            res.is_ok(),
            "decryptor should not panic when computing padded ciphertext lengths"
        );
        assert!(res.unwrap().is_err(), "expected error on huge orig_size");
    }
}
