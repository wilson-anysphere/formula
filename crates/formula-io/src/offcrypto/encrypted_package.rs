use std::io::{Read, Write};

use sha1::{Digest, Sha1};

use aes::cipher::{generic_array::GenericArray, BlockDecrypt, KeyInit};
use aes::{Aes128, Aes192, Aes256};

use crate::rc4_encrypted_package::{
    parse_rc4_encrypted_package_stream, Rc4EncryptedPackageParseError, Rc4EncryptedPackageParseOptions,
};

const ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN: usize = 8;
const AES_BLOCK_LEN: usize = 16;
const ENCRYPTED_PACKAGE_SEGMENT_LEN: usize = 0x1000;
const ENCRYPTED_PACKAGE_RC4_BLOCK_LEN: usize = 0x200;
const CRYPTOAPI_SPIN_COUNT: u32 = 50_000;

/// Errors returned by [`decrypt_standard_encrypted_package_stream`].
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
        "`EncryptedPackage` RC4 key length must be between 1 and 20 bytes for SHA-1, got {key_len}"
    )]
    Rc4InvalidKeyLength { key_len: usize },
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

fn derive_standard_cryptoapi_iv_sha1(salt: &[u8], segment_index: u32) -> [u8; AES_BLOCK_LEN] {
    // See docs/offcrypto-standard-encryptedpackage.md ("Variant B").
    let mut hasher = Sha1::new();
    hasher.update(salt);
    hasher.update(segment_index.to_le_bytes());
    let digest = hasher.finalize();
    let mut iv = [0u8; AES_BLOCK_LEN];
    iv.copy_from_slice(&digest[..AES_BLOCK_LEN]);
    iv
}

fn aes_cbc_decrypt_in_place(
    key: &[u8],
    iv: &[u8; AES_BLOCK_LEN],
    buf: &mut [u8],
) -> Result<(), EncryptedPackageError> {
    if buf.len() % AES_BLOCK_LEN != 0 {
        return Err(EncryptedPackageError::CiphertextLenNotBlockAligned {
            ciphertext_len: buf.len(),
        });
    }

    fn decrypt_with<C: BlockDecrypt + KeyInit>(
        key: &[u8],
        iv: &[u8; AES_BLOCK_LEN],
        buf: &mut [u8],
    ) -> Result<(), EncryptedPackageError> {
        let cipher = C::new_from_slice(key)
            .map_err(|_| EncryptedPackageError::InvalidAesKeyLength { key_len: key.len() })?;
        let mut prev = *iv;
        for block in buf.chunks_exact_mut(AES_BLOCK_LEN) {
            let mut cur = [0u8; AES_BLOCK_LEN];
            cur.copy_from_slice(block);

            cipher.decrypt_block(GenericArray::from_mut_slice(block));
            for (b, p) in block.iter_mut().zip(prev.iter()) {
                *b ^= p;
            }
            prev = cur;
        }
        Ok(())
    }

    match key.len() {
        16 => decrypt_with::<Aes128>(key, iv, buf),
        24 => decrypt_with::<Aes192>(key, iv, buf),
        32 => decrypt_with::<Aes256>(key, iv, buf),
        other => Err(EncryptedPackageError::InvalidAesKeyLength { key_len: other }),
    }
}

fn looks_like_zip_prefix(buf: &[u8]) -> bool {
    // Local file header / empty archive / spanning signature.
    buf.starts_with(b"PK\x03\x04")
        || buf.starts_with(b"PK\x05\x06")
        || buf.starts_with(b"PK\x07\x08")
}

fn pkcs7_padding_matches(decrypted_ciphertext: &[u8], orig_size: usize) -> bool {
    if orig_size == 0 {
        return false;
    }
    let rem = orig_size % AES_BLOCK_LEN;
    if rem != 0 {
        let padded_len = orig_size + (AES_BLOCK_LEN - rem);
        if decrypted_ciphertext.len() < padded_len {
            return false;
        }
        let pad_len = padded_len - orig_size;
        decrypted_ciphertext[orig_size..padded_len]
            .iter()
            .all(|b| *b == pad_len as u8)
    } else {
        let padded_len = orig_size + AES_BLOCK_LEN;
        if decrypted_ciphertext.len() < padded_len {
            return false;
        }
        decrypted_ciphertext[orig_size..padded_len]
            .iter()
            .all(|b| *b == AES_BLOCK_LEN as u8)
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
    let orig_size = u64::from_le_bytes(size_bytes);

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
/// - Remaining bytes are AES ciphertext (baseline MS-OFFCRYPTO/ECMA-376: AES-ECB, no IV).
/// - The concatenated plaintext is truncated to `orig_size` (do **not** rely on PKCS#7 unpadding).
///
/// When `salt` is present, this function will try both modes and pick the most plausible output
/// (preferring a ZIP prefix, then falling back to PKCS#7 padding shape).
pub fn decrypt_standard_encrypted_package_stream(
    encrypted_package_stream: &[u8],
    key: &[u8],
    salt: &[u8],
) -> Result<Vec<u8>, EncryptedPackageError> {
    if encrypted_package_stream.len() < ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN {
        return Err(EncryptedPackageError::StreamTooShort {
            len: encrypted_package_stream.len(),
        });
    }

    let mut size_bytes = [0u8; ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN];
    size_bytes.copy_from_slice(&encrypted_package_stream[..ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN]);
    let orig_size = u64::from_le_bytes(size_bytes);
    let ciphertext = &encrypted_package_stream[ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN..];

    if ciphertext.len() % AES_BLOCK_LEN != 0 {
        return Err(EncryptedPackageError::CiphertextLenNotBlockAligned {
            ciphertext_len: ciphertext.len(),
        });
    }

    let orig_size_usize = usize::try_from(orig_size)
        .map_err(|_| EncryptedPackageError::OrigSizeTooLargeForPlatform { orig_size })?;

    // Guardrail: `EncryptedPackage` carries the original plaintext size separately. Treat inputs as
    // malformed when the ciphertext is too short to possibly contain `orig_size` bytes (accounting
    // for AES block padding).
    //
    // This prevents:
    // - panics from integer overflow when computing padded sizes
    // - OOM from allocating based on attacker-controlled `orig_size`
    let expected_min_ciphertext_len = orig_size
        .checked_add((AES_BLOCK_LEN - 1) as u64)
        .and_then(|v| v.checked_div(AES_BLOCK_LEN as u64))
        .and_then(|blocks| blocks.checked_mul(AES_BLOCK_LEN as u64))
        .ok_or(EncryptedPackageError::ImplausibleOrigSize {
            orig_size,
            ciphertext_len: ciphertext.len(),
        })?;
    if (ciphertext.len() as u64) < expected_min_ciphertext_len {
        return Err(EncryptedPackageError::ImplausibleOrigSize {
            orig_size,
            ciphertext_len: ciphertext.len(),
        });
    }

    let mut plaintext = if !salt.is_empty() {
        // Some producers encrypt Standard/CryptoAPI `EncryptedPackage` using a non-standard
        // segmented mode.
        //
        // Unfortunately, some real-world files still use AES-ECB while also carrying a salt, so
        // we cannot select the mode purely from `salt.is_empty()`.
        //
        // Approach:
        // - Try both candidate decryptions (ECB and the segmented fallback).
        // - Prefer whichever candidate looks like a valid ZIP (OOXML payload).
        // - If neither looks like a ZIP, fall back to PKCS#7 padding shape as a weak heuristic,
        //   then default to ECB for compatibility.

        let mut cbc = ciphertext.to_vec();
        for (segment_index, segment) in cbc.chunks_mut(ENCRYPTED_PACKAGE_SEGMENT_LEN).enumerate() {
            let iv = derive_standard_cryptoapi_iv_sha1(salt, segment_index as u32);
            aes_cbc_decrypt_in_place(key, &iv, segment)?;
        }

        let mut ecb = ciphertext.to_vec();
        aes_ecb_decrypt_in_place(key, &mut ecb)?;

        let cbc_prefix = cbc.get(..orig_size_usize.min(cbc.len())).unwrap_or(&[]);
        if looks_like_zip_prefix(cbc_prefix) {
            cbc
        } else {
            let ecb_prefix = ecb.get(..orig_size_usize.min(ecb.len())).unwrap_or(&[]);
            if looks_like_zip_prefix(ecb_prefix) {
                ecb
            } else if pkcs7_padding_matches(&cbc, orig_size_usize)
                && !pkcs7_padding_matches(&ecb, orig_size_usize)
            {
                cbc
            } else if !pkcs7_padding_matches(&cbc, orig_size_usize)
                && pkcs7_padding_matches(&ecb, orig_size_usize)
            {
                ecb
            } else if pkcs7_padding_matches(&cbc, orig_size_usize) {
                cbc
            } else {
                ecb
            }
        }
    } else {
        // AES-ECB variant (no IV).
        let mut out = ciphertext.to_vec();
        aes_ecb_decrypt_in_place(key, &mut out)?;
        out
    };

    if plaintext.len() < orig_size_usize {
        return Err(EncryptedPackageError::DecryptedTooShort {
            decrypted_len: plaintext.len(),
            orig_size,
        });
    }

    plaintext.truncate(orig_size_usize);
    Ok(plaintext)
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

fn password_utf16le_bytes(password: &str) -> Vec<u8> {
    let mut out = Vec::with_capacity(password.len().saturating_mul(2));
    for cu in password.encode_utf16() {
        out.extend_from_slice(&cu.to_le_bytes());
    }
    out
}

fn cryptoapi_password_hash_sha1(password: &str, salt: &[u8]) -> [u8; 20] {
    // H = SHA1(salt || password_utf16le)
    let pw = password_utf16le_bytes(password);
    let mut hasher = Sha1::new();
    hasher.update(salt);
    hasher.update(&pw);
    let mut h: [u8; 20] = hasher.finalize().into();

    // for i in 0..CRYPTOAPI_SPIN_COUNT: H = SHA1(u32le(i) || H)
    let mut buf = [0u8; 4 + 20];
    for i in 0..CRYPTOAPI_SPIN_COUNT {
        buf[..4].copy_from_slice(&i.to_le_bytes());
        buf[4..].copy_from_slice(&h);
        h = Sha1::digest(&buf).into();
    }
    h
}

fn cryptoapi_block_key_sha1(password_hash: &[u8; 20], block_index: u32) -> [u8; 20] {
    let mut hasher = Sha1::new();
    hasher.update(password_hash);
    hasher.update(block_index.to_le_bytes());
    hasher.finalize().into()
}

struct CryptoapiRc4EncryptedPackageDecryptor {
    password_hash: [u8; 20],
    key_len: usize,
}

impl CryptoapiRc4EncryptedPackageDecryptor {
    fn new(password: &str, salt: &[u8], key_len: usize) -> Result<Self, EncryptedPackageError> {
        if !(1..=20).contains(&key_len) {
            return Err(EncryptedPackageError::Rc4InvalidKeyLength { key_len });
        }
        Ok(Self {
            password_hash: cryptoapi_password_hash_sha1(password, salt),
            key_len,
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
            Rc4EncryptedPackageParseError::TruncatedHeader => EncryptedPackageError::StreamTooShort {
                len: encrypted_package_stream.len(),
            },
            Rc4EncryptedPackageParseError::DeclaredSizeExceedsPayload { declared, available } => {
                EncryptedPackageError::ImplausibleOrigSize {
                    orig_size: declared,
                    ciphertext_len: available as usize,
                }
            }
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
            let digest = cryptoapi_block_key_sha1(&self.password_hash, block_index as u32);
            let mut rc4 = if self.key_len == 5 {
                // CryptoAPI/Office represent a "40-bit" RC4 key as a 128-bit RC4 key with the high
                // 88 bits zero. Using a raw 5-byte key changes RC4 KSA and yields the wrong
                // keystream.
                let mut padded = [0u8; 16];
                padded[..5].copy_from_slice(&digest[..5]);
                Rc4::new(&padded)
            } else {
                Rc4::new(&digest[..self.key_len])
            };
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
/// - `password_hash = SHA1(salt || UTF16LE(password))`
/// - for `i in 0..50000`: `password_hash = SHA1(LE32(i) || password_hash)`
/// - `h_i = SHA1(password_hash || LE32(i))`
/// - If `key_len == 5` (40-bit): `rc4_key_i = h_i[0..5] || 0x00 * 11` (16 bytes total)
/// - Otherwise: `rc4_key_i = h_i[0..key_len]`
/// - RC4 is **reset** per block (do not carry keystream state across blocks).
///
/// See `docs/offcrypto-standard-cryptoapi-rc4.md` for additional notes and test vectors.
pub fn decrypt_standard_cryptoapi_rc4_encrypted_package_stream(
    encrypted_package_stream: &[u8],
    password: &str,
    salt: &[u8],
    key_len: usize,
) -> Result<Vec<u8>, EncryptedPackageError> {
    CryptoapiRc4EncryptedPackageDecryptor::new(password, salt, key_len)?
        .decrypt_encrypted_package_stream(encrypted_package_stream)
}

#[cfg(test)]
mod tests {
    use super::*;
    use aes::cipher::{generic_array::GenericArray, BlockEncrypt, KeyInit};
    use aes::{Aes128, Aes192, Aes256};
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
        assert_eq!(err, EncryptedPackageError::StreamTooShort { len: 7 });
    }

    #[test]
    fn errors_on_non_block_aligned_ciphertext() {
        let key = fixed_key_16();

        let mut encrypted = Vec::new();
        encrypted.extend_from_slice(&(1u64).to_le_bytes());
        encrypted.extend_from_slice(&[0u8; 15]); // not multiple of 16

        let err = decrypt_standard_encrypted_package_stream(&encrypted, &key, &[]).unwrap_err();
        assert_eq!(
            err,
            EncryptedPackageError::CiphertextLenNotBlockAligned { ciphertext_len: 15 }
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
            EncryptedPackageError::ImplausibleOrigSize {
                orig_size: 32,
                ciphertext_len: 16
            }
        );
    }

    #[test]
    fn errors_on_invalid_aes_key_length() {
        let key = [0u8; 17];
        let encrypted = 0u64.to_le_bytes().to_vec();
        let err = decrypt_standard_encrypted_package_stream(&encrypted, &key, &[]).unwrap_err();
        assert_eq!(
            err,
            EncryptedPackageError::InvalidAesKeyLength { key_len: 17 }
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
        if key_len == 5 {
            let mut key = Vec::with_capacity(16);
            key.extend_from_slice(&digest[..5]);
            key.resize(16, 0);
            key
        } else {
            digest[..key_len].to_vec()
        }
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
        let key_len = 16;

        let encryptor = TestCryptoapiRc4Encryptor::new(password, &salt, key_len);
        let decryptor = CryptoapiRc4EncryptedPackageDecryptor::new(password, &salt, key_len)
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
    fn rc4_cryptoapi_encryptedpackage_roundtrip_with_40_bit_key() {
        // Regression: 40-bit CryptoAPI RC4 uses a padded 16-byte RC4 key, not a raw 5-byte key.
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
        assert_eq!(
            err,
            EncryptedPackageError::OrigSizeTooLargeForPlatform { orig_size }
        );
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
