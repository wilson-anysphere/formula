use sha1::{Digest, Sha1};

use formula_xlsx::offcrypto::{decrypt_aes_cbc_no_padding_in_place, AesCbcDecryptError};

const ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN: usize = 8;
const ENCRYPTED_PACKAGE_SEGMENT_LEN: usize = 0x1000;
const AES_BLOCK_LEN: usize = 16;
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

    #[error("`EncryptedPackage` ciphertext length {ciphertext_len} is not a multiple of AES block size ({AES_BLOCK_LEN})")]
    CiphertextLenNotBlockAligned { ciphertext_len: usize },

    #[error(
        "`EncryptedPackage` AES key length must be 16, 24, or 32 bytes (AES-128/192/256), got {key_len}"
    )]
    InvalidAesKeyLength { key_len: usize },

    #[error("`EncryptedPackage` segment index overflow (file too large)")]
    SegmentIndexOverflow,

    #[error("AES-CBC decryption failed for segment {segment_index}")]
    SegmentDecryptFailed { segment_index: u32 },

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

fn derive_segment_iv(salt: &[u8], segment_index: u32) -> [u8; AES_BLOCK_LEN] {
    let mut hasher = Sha1::new();
    hasher.update(salt);
    hasher.update(segment_index.to_le_bytes());
    let digest = hasher.finalize();

    let mut iv = [0u8; AES_BLOCK_LEN];
    iv.copy_from_slice(&digest[..AES_BLOCK_LEN]);
    iv
}

/// Decrypt an MS-OFFCRYPTO "Standard" (CryptoAPI) `EncryptedPackage` stream.
///
/// The caller must provide:
/// - `key`: the file encryption key (AES-128/192/256).
/// - `salt`: the `EncryptionVerifier` salt used to derive per-segment IVs.
///
/// Algorithm summary (MS-OFFCRYPTO Standard/CryptoAPI):
/// - First 8 bytes are the original plaintext size (`orig_size`) as a little-endian `u64`.
/// - Remaining bytes are AES-CBC ciphertext split into 0x1000-byte segments (except the last).
/// - Each segment `i` uses IV = `SHA1(salt || LE32(i))[0..16]` and is decrypted independently.
/// - The concatenated plaintext is truncated to `orig_size` (do **not** rely on PKCS#7 unpadding).
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

    if ciphertext.is_empty() && orig_size == 0 {
        return Ok(Vec::new());
    }

    if ciphertext.len() % AES_BLOCK_LEN != 0 {
        return Err(EncryptedPackageError::CiphertextLenNotBlockAligned {
            ciphertext_len: ciphertext.len(),
        });
    }

    let orig_size_usize = usize::try_from(orig_size)
        .map_err(|_| EncryptedPackageError::OrigSizeTooLargeForPlatform { orig_size })?;

    // --- Guardrails for malicious `orig_size` ---
    //
    // `EncryptedPackage` stores the unencrypted package size (`orig_size`) separately from the
    // ciphertext bytes. A corrupt/malicious size can otherwise:
    // - overflow segment math like `ceil(orig_size / 4096)` or `8 + i*4096`
    // - induce large allocations (OOM) in naive implementations
    //
    // We keep these checks conservative to avoid rejecting valid-but-unusual files.

    // Compute number of 4096-byte segments implied by `orig_size`, but do it in a way that cannot
    // overflow (avoid `(orig_size + 4095) / 4096`).
    let n_segments: u64 = if orig_size == 0 {
        0
    } else {
        let seg = ENCRYPTED_PACKAGE_SEGMENT_LEN as u64;
        (orig_size / seg) + u64::from(orig_size % seg != 0)
    };

    // Prevent overflow in ciphertext offset calculations like `8 + i*4096` by validating that the
    // final segment start is representable (and plausible for the provided ciphertext length).
    //
    // We only need this for malformed inputs; for well-formed packages, `orig_size` and the
    // ciphertext length agree.
    if n_segments > 0 {
        let seg_len_u64 = ENCRYPTED_PACKAGE_SEGMENT_LEN as u64;
        let last_seg_start = n_segments
            .saturating_sub(1)
            .checked_mul(seg_len_u64)
            .ok_or(EncryptedPackageError::ImplausibleOrigSize {
                orig_size,
                ciphertext_len: ciphertext.len(),
            })?;

        // Require at least one AES block of ciphertext for the final segment.
        let min_ciphertext_needed = last_seg_start.checked_add(AES_BLOCK_LEN as u64).ok_or(
            EncryptedPackageError::ImplausibleOrigSize {
                orig_size,
                ciphertext_len: ciphertext.len(),
            },
        )?;
        if (ciphertext.len() as u64) < min_ciphertext_needed {
            return Err(EncryptedPackageError::ImplausibleOrigSize {
                orig_size,
                ciphertext_len: ciphertext.len(),
            });
        }
    }

    // If ciphertext length is known (buffer-based decrypt), reject clearly implausible `orig_size`
    // values. Allow up to one extra segment of slop to account for producer differences (e.g.
    // padding to a 4096-byte boundary, OLE sector slack, etc).
    let plausible_max =
        (ciphertext.len() as u64).saturating_add(ENCRYPTED_PACKAGE_SEGMENT_LEN as u64);
    if orig_size > plausible_max {
        return Err(EncryptedPackageError::ImplausibleOrigSize {
            orig_size,
            ciphertext_len: ciphertext.len(),
        });
    }

    // Guardrail: `EncryptedPackage` carries the original plaintext size separately. Treat inputs as
    // malformed when the ciphertext is too short to possibly contain `orig_size` bytes (accounting
    // for AES block padding).
    //
    // This prevents:
    // - panics from segment math/slicing when the length header is corrupt
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

    // Decrypt segment-by-segment until we have produced `orig_size` bytes (or run out of input).
    //
    // Allocate at most `orig_size` bytes.
    //
    // We validated above that:
    // - `orig_size` fits in `usize`
    // - the ciphertext is long enough to plausibly contain `orig_size` bytes
    //
    // Using `orig_size` avoids allocating based on trailing ciphertext/padding bytes when the
    // `EncryptedPackage` stream is larger than the declared original size.
    let mut out = Vec::with_capacity(orig_size_usize);
    let mut segment_index: u32 = 0;
    while out.len() < orig_size_usize {
        // Compute segment start as `8 + i*4096` using checked arithmetic to avoid integer overflow.
        let seg_start = (segment_index as usize)
            .checked_mul(ENCRYPTED_PACKAGE_SEGMENT_LEN)
            .and_then(|v| v.checked_add(ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN))
            .ok_or(EncryptedPackageError::SegmentIndexOverflow)?;

        if seg_start >= encrypted_package_stream.len() {
            break;
        }

        let remaining = encrypted_package_stream.len() - seg_start;
        let seg_len = remaining.min(ENCRYPTED_PACKAGE_SEGMENT_LEN);

        // `seg_len` is either 0x1000 (full segment) or the final remainder. Since 0x1000 is a
        // multiple of 16, validating the entire ciphertext length above is sufficient to guarantee
        // that `seg_len` is also block-aligned. Keep the check anyway for defense-in-depth.
        if seg_len % AES_BLOCK_LEN != 0 {
            return Err(EncryptedPackageError::CiphertextLenNotBlockAligned {
                ciphertext_len: seg_len,
            });
        }

        let iv = derive_segment_iv(salt, segment_index);
        let mut decrypted = encrypted_package_stream[seg_start..seg_start + seg_len].to_vec();
        decrypt_aes_cbc_no_padding_in_place(key, &iv, &mut decrypted).map_err(|err| match err {
            AesCbcDecryptError::UnsupportedKeyLength(key_len) => {
                EncryptedPackageError::InvalidAesKeyLength { key_len }
            }
            AesCbcDecryptError::InvalidIvLength(_) => {
                EncryptedPackageError::SegmentDecryptFailed { segment_index }
            }
            AesCbcDecryptError::InvalidCiphertextLength(ciphertext_len) => {
                EncryptedPackageError::CiphertextLenNotBlockAligned { ciphertext_len }
            }
        })?;

        let remaining_needed = orig_size_usize - out.len();
        if decrypted.len() > remaining_needed {
            out.extend_from_slice(&decrypted[..remaining_needed]);
            break;
        }

        out.extend_from_slice(&decrypted);
        segment_index = segment_index
            .checked_add(1)
            .ok_or(EncryptedPackageError::SegmentIndexOverflow)?;
    }

    if out.len() < orig_size_usize {
        return Err(EncryptedPackageError::DecryptedTooShort {
            decrypted_len: out.len(),
            orig_size,
        });
    }

    // The loop truncates at `orig_size` already, but keep this for clarity.
    out.truncate(orig_size_usize);
    Ok(out)
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
        if encrypted_package_stream.len() < ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN {
            return Err(EncryptedPackageError::StreamTooShort {
                len: encrypted_package_stream.len(),
            });
        }

        let mut size_bytes = [0u8; ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN];
        size_bytes.copy_from_slice(&encrypted_package_stream[..ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN]);
        let orig_size = u64::from_le_bytes(size_bytes);
        let ciphertext = &encrypted_package_stream[ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN..];

        if ciphertext.is_empty() && orig_size == 0 {
            return Ok(Vec::new());
        }

        let orig_size_usize = usize::try_from(orig_size)
            .map_err(|_| EncryptedPackageError::OrigSizeTooLargeForPlatform { orig_size })?;

        // RC4 has no block padding requirements, so ciphertext must contain at least `orig_size` bytes.
        if orig_size > ciphertext.len() as u64 {
            return Err(EncryptedPackageError::ImplausibleOrigSize {
                orig_size,
                ciphertext_len: ciphertext.len(),
            });
        }

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
/// Unlike AES-based `EncryptedPackage` encryption (0x1000-byte segments), the RC4 variant uses
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
    use aes::{Aes128, Aes192, Aes256};
    use cbc::cipher::block_padding::NoPadding;
    use cbc::cipher::{BlockEncryptMut, KeyIvInit};

    fn derive_segment_iv_reference(salt: &[u8], segment_index: u32) -> [u8; AES_BLOCK_LEN] {
        // Spec reference: SHA1(salt || LE32(i))[:16]
        let mut hasher = Sha1::new();
        hasher.update(salt);
        hasher.update(segment_index.to_le_bytes());
        let digest = hasher.finalize();
        let mut iv = [0u8; AES_BLOCK_LEN];
        iv.copy_from_slice(&digest[..AES_BLOCK_LEN]);
        iv
    }

    fn fixed_key(len: usize) -> Vec<u8> {
        (0..len).map(|i| i as u8).collect()
    }

    fn fixed_key_16() -> Vec<u8> {
        fixed_key(16)
    }

    fn fixed_salt_16() -> Vec<u8> {
        (0x10u8..=0x1F).collect()
    }

    fn pkcs7_pad(plaintext: &[u8]) -> Vec<u8> {
        if plaintext.is_empty() {
            return Vec::new();
        }
        let mut out = plaintext.to_vec();
        let mut pad_len = AES_BLOCK_LEN - (out.len() % AES_BLOCK_LEN);
        if pad_len == 0 {
            pad_len = AES_BLOCK_LEN;
        }
        out.extend(std::iter::repeat(pad_len as u8).take(pad_len));
        out
    }

    fn encrypt_segment_aes_cbc_no_padding(
        key: &[u8],
        iv: &[u8; AES_BLOCK_LEN],
        plaintext: &[u8],
    ) -> Vec<u8> {
        assert!(plaintext.len() % AES_BLOCK_LEN == 0);

        // Encrypt in-place to avoid needing the higher-level padding helpers.
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
            _ => panic!("unsupported key length"),
        }
        buf
    }

    fn encrypt_encrypted_package_stream_standard_cryptoapi(
        key: &[u8],
        salt: &[u8],
        plaintext: &[u8],
    ) -> Vec<u8> {
        let orig_size = plaintext.len() as u64;

        let mut out = Vec::new();
        out.extend_from_slice(&orig_size.to_le_bytes());

        if plaintext.is_empty() {
            return out;
        }

        let padded = pkcs7_pad(plaintext);
        for (i, chunk) in padded.chunks(ENCRYPTED_PACKAGE_SEGMENT_LEN).enumerate() {
            let iv = derive_segment_iv(salt, i as u32);
            let ciphertext = encrypt_segment_aes_cbc_no_padding(key, &iv, chunk);
            out.extend_from_slice(&ciphertext);
        }

        out
    }

    fn encrypt_encrypted_package_stream_standard_cryptoapi_reference(
        key: &[u8],
        salt: &[u8],
        plaintext: &[u8],
    ) -> Vec<u8> {
        let orig_size = plaintext.len() as u64;

        let mut out = Vec::new();
        out.extend_from_slice(&orig_size.to_le_bytes());

        if plaintext.is_empty() {
            return out;
        }

        let padded = pkcs7_pad(plaintext);
        for (i, chunk) in padded.chunks(ENCRYPTED_PACKAGE_SEGMENT_LEN).enumerate() {
            // Reference IV derivation (do not use the decryptor's helper) so this test fails if the
            // decryptor ever truncates/pads the salt or uses the wrong segment index endianness.
            let iv = derive_segment_iv_reference(salt, i as u32);
            let ciphertext = encrypt_segment_aes_cbc_no_padding(key, &iv, chunk);
            out.extend_from_slice(&ciphertext);
        }

        out
    }

    fn make_plaintext(len: usize) -> Vec<u8> {
        (0..len).map(|i| (i % 251) as u8).collect()
    }

    #[test]
    fn round_trip_decrypts_aes_128_192_256() {
        let salt = fixed_salt_16();

        for key_len in [16usize, 24, 32] {
            let key = fixed_key(key_len);
            for size in [0usize, 1, 15, 16, 17, 4095, 4096, 4097, 8192 + 123] {
                let plaintext = make_plaintext(size);
                let encrypted =
                    encrypt_encrypted_package_stream_standard_cryptoapi(&key, &salt, &plaintext);
                let decrypted =
                    decrypt_standard_encrypted_package_stream(&encrypted, &key, &salt).unwrap();
                assert_eq!(decrypted, plaintext, "failed for key_len={key_len} size={size}");
            }
        }
    }

    #[test]
    fn decrypt_truncates_to_orig_size_even_with_trailing_bytes() {
        let key = fixed_key(16);
        let salt = fixed_salt_16();

        // Use a size that ends on a segment boundary (4096) so callers can stop after decrypting
        // only the first 0x1000-byte segment.
        let plaintext = make_plaintext(4096);
        let mut encrypted =
            encrypt_encrypted_package_stream_standard_cryptoapi(&key, &salt, &plaintext);

        // Append extra bytes (simulating OLE sector padding). Ensure the overall ciphertext is still
        // block-aligned but the trailing bytes do not represent valid PKCS#7 padding.
        encrypted.extend_from_slice(&[0u8; 16]);

        let decrypted = decrypt_standard_encrypted_package_stream(&encrypted, &key, &salt).unwrap();
        assert_eq!(decrypted, plaintext);
    }

    #[test]
    fn decrypt_empty_ciphertext_and_zero_size_is_ok() {
        let key = fixed_key(16);
        let salt = fixed_salt_16();

        let encrypted = 0u64.to_le_bytes().to_vec();
        let decrypted = decrypt_standard_encrypted_package_stream(&encrypted, &key, &salt).unwrap();
        assert!(decrypted.is_empty());
    }

    #[test]
    fn errors_on_short_stream() {
        let key = fixed_key(16);
        let salt = fixed_salt_16();

        let err = decrypt_standard_encrypted_package_stream(&[0u8; 7], &key, &salt).unwrap_err();
        assert_eq!(err, EncryptedPackageError::StreamTooShort { len: 7 });
    }

    #[test]
    fn errors_on_non_block_aligned_ciphertext() {
        let key = fixed_key(16);
        let salt = fixed_salt_16();

        let mut encrypted = Vec::new();
        encrypted.extend_from_slice(&(1u64).to_le_bytes());
        encrypted.extend_from_slice(&[0u8; 15]); // not multiple of 16

        let err = decrypt_standard_encrypted_package_stream(&encrypted, &key, &salt).unwrap_err();
        assert_eq!(
            err,
            EncryptedPackageError::CiphertextLenNotBlockAligned { ciphertext_len: 15 }
        );
    }

    #[test]
    fn errors_when_length_header_exceeds_ciphertext() {
        let key = fixed_key(16);
        let salt = fixed_salt_16();

        // orig_size claims 32 bytes, but we only have 16 bytes of ciphertext.
        let mut encrypted = Vec::new();
        encrypted.extend_from_slice(&(32u64).to_le_bytes());
        encrypted.extend_from_slice(&[0u8; 16]);

        let err = decrypt_standard_encrypted_package_stream(&encrypted, &key, &salt).unwrap_err();
        assert_eq!(
            err,
            EncryptedPackageError::ImplausibleOrigSize {
                orig_size: 32,
                ciphertext_len: 16
            }
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
        let decryptor =
            CryptoapiRc4EncryptedPackageDecryptor::new(password, &salt, key_len).expect("decryptor");

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
        let salt = fixed_salt_16();

        // Tiny `EncryptedPackage`: just the size prefix, no ciphertext.
        let encrypted = u64::MAX.to_le_bytes().to_vec();
        let res = std::panic::catch_unwind(|| {
            decrypt_standard_encrypted_package_stream(&encrypted, &key, &salt)
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
        let salt = fixed_salt_16();

        let orig_size = (usize::MAX as u64) + 1;
        let encrypted = orig_size.to_le_bytes().to_vec();
        let res = std::panic::catch_unwind(|| {
            decrypt_standard_encrypted_package_stream(&encrypted, &key, &salt)
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
        let salt = fixed_salt_16();

        // This value would overflow `(orig_size + 4095)` in naive `ceil(orig_size/4096)` math.
        let orig_size = u64::MAX - 4094;
        let encrypted = orig_size.to_le_bytes().to_vec();

        let res = std::panic::catch_unwind(|| {
            decrypt_standard_encrypted_package_stream(&encrypted, &key, &salt)
        });
        assert!(
            res.is_ok(),
            "decryptor should not panic when computing segment counts/offsets"
        );
        assert!(res.unwrap().is_err(), "expected error on huge orig_size");
    }

    #[test]
    fn derive_iv_known_answer_uses_full_salt_and_le_u32_segment_index() {
        // salt = 0x00..0x07 (8 bytes, intentionally not 16)
        let salt: Vec<u8> = (0u8..8).collect();

        // Expected SHA1(salt || LE32(i))[:16]
        // i = 0: SHA1(0001020304050607 00000000) = 5eb83f233710c17da63ba93c6e11a0a928db13b5
        // i = 1: SHA1(0001020304050607 01000000) = e53ad9c78a02910481b7ead1e296876ebb94c934
        let expected_i0: [u8; 16] = [
            0x5e, 0xb8, 0x3f, 0x23, 0x37, 0x10, 0xc1, 0x7d, 0xa6, 0x3b, 0xa9, 0x3c, 0x6e, 0x11,
            0xa0, 0xa9,
        ];
        let expected_i1: [u8; 16] = [
            0xe5, 0x3a, 0xd9, 0xc7, 0x8a, 0x02, 0x91, 0x04, 0x81, 0xb7, 0xea, 0xd1, 0xe2, 0x96,
            0x87, 0x6e,
        ];

        assert_eq!(derive_segment_iv(&salt, 0), expected_i0);
        assert_eq!(derive_segment_iv(&salt, 1), expected_i1);
    }

    #[test]
    fn decrypt_round_trips_non_16_byte_salts_across_multiple_segments() {
        let key = fixed_key_16();

        // >4096 bytes so we exercise segmentIndex 0 and 1.
        let plaintext = make_plaintext(5000);

        for salt in [
            (0u8..8).collect::<Vec<u8>>(),
            (0u8..32).collect::<Vec<u8>>(),
        ] {
            let encrypted = encrypt_encrypted_package_stream_standard_cryptoapi_reference(
                &key, &salt, &plaintext,
            );
            let decrypted =
                decrypt_standard_encrypted_package_stream(&encrypted, &key, &salt).unwrap();
            assert_eq!(
                decrypted,
                plaintext,
                "round-trip failed for salt len={}",
                salt.len()
            );
        }
    }

    #[test]
    fn errors_on_invalid_key_lengths() {
        let salt = fixed_salt_16();
        let plaintext = make_plaintext(ENCRYPTED_PACKAGE_SEGMENT_LEN + 123);
        let encrypted =
            encrypt_encrypted_package_stream_standard_cryptoapi(&fixed_key(16), &salt, &plaintext);

        for bad_len in [0usize, 1, 15, 17, 23, 25, 31, 33] {
            let bad_key = vec![0u8; bad_len];
            let err =
                decrypt_standard_encrypted_package_stream(&encrypted, &bad_key, &salt).unwrap_err();
            assert_eq!(
                err,
                EncryptedPackageError::InvalidAesKeyLength { key_len: bad_len },
                "bad_len={bad_len}"
            );
        }
    }
}
