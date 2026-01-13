use crate::{AgileEncryptionInfo, DecryptOptions, OffcryptoError, Reader};

use aes::{Aes128, Aes192, Aes256};
use cbc::Decryptor;
use cipher::{block_padding::NoPadding, BlockDecryptMut, KeyIvInit};
use sha1::{Digest as _, Sha1};

/// Office encrypted packages are segmented into 4096-byte blocks.
pub const ENCRYPTED_PACKAGE_SEGMENT_LEN: usize = 4096;

/// The AES block size used by both Agile and Standard encryption.
pub const AES_BLOCK_LEN: usize = 16;

#[inline]
fn padded_aes_len(len: usize) -> usize {
    // `len` is at most `ENCRYPTED_PACKAGE_SEGMENT_LEN` (4096), so this cannot overflow.
    let rem = len % AES_BLOCK_LEN;
    if rem == 0 {
        len
    } else {
        len + (AES_BLOCK_LEN - rem)
    }
}

fn checked_output_len(total_size: u64) -> Result<usize, OffcryptoError> {
    let len = usize::try_from(total_size)
        .map_err(|_| OffcryptoError::EncryptedPackageSizeOverflow { total_size })?;

    // `Vec<u8>` cannot exceed `isize::MAX` due to `Layout::array`/pointer offset invariants.
    isize::try_from(len)
        .map_err(|_| OffcryptoError::EncryptedPackageSizeOverflow { total_size })?;

    Ok(len)
}

/// Decrypt an `EncryptedPackage` stream using the provided per-block decryptor.
///
/// The decryptor is invoked with a 0-based segment index, the ciphertext bytes for the segment
/// (padded to AES block size), and a plaintext output buffer (same length as ciphertext).
///
/// Callers are expected to copy only the first `min(remaining, 4096)` bytes of each decrypted
/// segment; the remainder contains padding.
pub fn decrypt_encrypted_package<F>(
    encrypted_package: &[u8],
    decrypt_block: F,
) -> Result<Vec<u8>, OffcryptoError>
where
    F: FnMut(u32, &[u8], &mut [u8]) -> Result<(), OffcryptoError>,
{
    decrypt_encrypted_package_with_options(encrypted_package, &DecryptOptions::default(), decrypt_block)
}

/// Decrypt an `EncryptedPackage` stream using the provided per-block decryptor, with configurable
/// limits.
///
/// See [`decrypt_encrypted_package`] for details on the `decrypt_block` callback.
pub fn decrypt_encrypted_package_with_options<F>(
    encrypted_package: &[u8],
    options: &DecryptOptions,
    mut decrypt_block: F,
) -> Result<Vec<u8>, OffcryptoError>
where
    F: FnMut(u32, &[u8], &mut [u8]) -> Result<(), OffcryptoError>,
{
    let mut reader = Reader::new(encrypted_package);
    let size_bytes = reader.take(8, "EncryptedPackageHeader.original_size")?;
    let len_lo = u32::from_le_bytes([size_bytes[0], size_bytes[1], size_bytes[2], size_bytes[3]])
        as u64;
    let len_hi = u32::from_le_bytes([size_bytes[4], size_bytes[5], size_bytes[6], size_bytes[7]])
        as u64;
    let total_size_u64 = len_lo | (len_hi << 32);
    let ciphertext_len_u64 = reader.remaining().len() as u64;
    // MS-OFFCRYPTO describes `original_size` as a `u64le`, but some producers/libraries treat it as
    // `u32 totalSize` + `u32 reserved` (often 0). When the high DWORD is non-zero but the combined
    // 64-bit value is not plausible for the available ciphertext, fall back to the low DWORD
    // *only when it is non-zero* (so we don't misinterpret true 64-bit sizes that are exact
    // multiples of 2^32).
    let total_size =
        if len_lo != 0
            && len_hi != 0
            && total_size_u64 > ciphertext_len_u64
            && len_lo <= ciphertext_len_u64
        {
            len_lo
        } else {
            total_size_u64
        };
    if let Some(max) = options.limits.max_output_size {
        if total_size > max {
            return Err(OffcryptoError::OutputTooLarge { total_size, max });
        }
    }
    let output_len = checked_output_len(total_size)?;

    // Validate ciphertext framing before allocating based on attacker-controlled `total_size`.
    //
    // Each ciphertext segment is padded up to the AES block size (16). The total ciphertext length
    // should therefore also be block-aligned, and must be large enough to cover the padded segment
    // lengths implied by `total_size`.
    let ciphertext_len = reader.remaining().len();
    if ciphertext_len % AES_BLOCK_LEN != 0 {
        return Err(OffcryptoError::InvalidCiphertextLength {
            len: ciphertext_len,
        });
    }

    // Minimum ciphertext length implied by `total_size`.
    let last_len = (total_size % ENCRYPTED_PACKAGE_SEGMENT_LEN as u64) as usize;
    let full_bytes = total_size - (last_len as u64);
    let required_ciphertext_len = if last_len == 0 {
        full_bytes
    } else {
        full_bytes
            .checked_add(padded_aes_len(last_len) as u64)
            .ok_or(OffcryptoError::EncryptedPackageSizeOverflow { total_size })?
    };
    if required_ciphertext_len > ciphertext_len as u64 {
        return Err(OffcryptoError::EncryptedPackageSizeMismatch {
            total_size,
            ciphertext_len,
        });
    }

    let mut out = Vec::new();
    out.try_reserve_exact(output_len)
        .map_err(|_| OffcryptoError::EncryptedPackageAllocationFailed { total_size })?;
    out.resize(output_len, 0);

    let mut remaining: u64 = total_size;
    let mut out_offset: usize = 0;
    let mut block_index: u32 = 0;

    let mut plaintext_buf = [0u8; ENCRYPTED_PACKAGE_SEGMENT_LEN];

    while remaining > 0 {
        let plaintext_len = std::cmp::min(remaining, ENCRYPTED_PACKAGE_SEGMENT_LEN as u64) as usize;
        let ciphertext_len = padded_aes_len(plaintext_len);
        let ciphertext = reader.take(ciphertext_len, "EncryptedPackage.ciphertext_segment")?;

        decrypt_block(
            block_index,
            ciphertext,
            &mut plaintext_buf[..ciphertext_len],
        )?;

        out[out_offset..out_offset + plaintext_len]
            .copy_from_slice(&plaintext_buf[..plaintext_len]);

        out_offset += plaintext_len;
        remaining -= plaintext_len as u64;
        block_index = block_index
            .checked_add(1)
            .ok_or(OffcryptoError::EncryptedPackageSizeOverflow { total_size })?;
    }

    Ok(out)
}

/// Decrypt an ECMA-376 Agile `EncryptedPackage` stream.
///
/// Algorithm notes (matching `msoffcrypto`):
/// - The first 8 bytes are the *unencrypted* payload size (little-endian).
/// - Ciphertext begins at offset 8 and is decrypted in 4096-byte segments.
/// - Each segment uses an IV derived from `HASH(keyDataSalt || u32le(i))[:16]`.
/// - The final output is truncated to exactly the declared plaintext size.
pub fn agile_decrypt_package(
    info: &AgileEncryptionInfo,
    secret_key: &[u8],
    encrypted_package: &[u8],
) -> Result<Vec<u8>, OffcryptoError> {
    // Fast plausibility check to avoid attempting huge allocations for corrupt inputs.
    if encrypted_package.len() < 8 {
        return Err(OffcryptoError::Truncated {
            context: "EncryptedPackageHeader.original_size",
        });
    }
    let size_bytes: [u8; 8] =
        encrypted_package[..8]
            .try_into()
            .map_err(|_| OffcryptoError::Truncated {
                context: "EncryptedPackageHeader.original_size",
            })?;
    let len_lo = u32::from_le_bytes([size_bytes[0], size_bytes[1], size_bytes[2], size_bytes[3]])
        as u64;
    let len_hi = u32::from_le_bytes([size_bytes[4], size_bytes[5], size_bytes[6], size_bytes[7]])
        as u64;
    let total_size_u64 = len_lo | (len_hi << 32);
    let ciphertext_len = encrypted_package.len().saturating_sub(8) as u64;
    let total_size =
        if len_lo != 0 && len_hi != 0 && total_size_u64 > ciphertext_len && len_lo <= ciphertext_len
        {
            len_lo
        } else {
            total_size_u64
        };
    let plausible_max = (encrypted_package.len() as u64).saturating_mul(2);
    if total_size > plausible_max {
        return Err(OffcryptoError::EncryptedPackageSizeOverflow { total_size });
    }

    let digest_len = info.key_data_hash_algorithm.digest_len();
    if digest_len < AES_BLOCK_LEN {
        return Err(OffcryptoError::InvalidEncryptionInfo {
            context: "hash output too short for AES IV",
        });
    }
    let mut digest_buf = [0u8; crate::MAX_DIGEST_LEN];

    decrypt_encrypted_package(encrypted_package, |segment_index, ciphertext, plaintext| {
        let mut iv = [0u8; 16];
        info.key_data_hash_algorithm.digest_two_into(
            &info.key_data_salt,
            &segment_index.to_le_bytes(),
            &mut digest_buf[..digest_len],
        );
        iv.copy_from_slice(&digest_buf[..AES_BLOCK_LEN]);

        plaintext.copy_from_slice(ciphertext);
        let pt_len = plaintext.len();

        match secret_key.len() {
            16 => {
                let decryptor =
                    Decryptor::<Aes128>::new_from_slices(secret_key, &iv).map_err(|_| {
                        OffcryptoError::InvalidKeyLength {
                            len: secret_key.len(),
                        }
                    })?;
                decryptor
                    .decrypt_padded_mut::<NoPadding>(plaintext)
                    .map_err(|_| OffcryptoError::InvalidCiphertextLength { len: pt_len })?;
            }
            24 => {
                let decryptor =
                    Decryptor::<Aes192>::new_from_slices(secret_key, &iv).map_err(|_| {
                        OffcryptoError::InvalidKeyLength {
                            len: secret_key.len(),
                        }
                    })?;
                decryptor
                    .decrypt_padded_mut::<NoPadding>(plaintext)
                    .map_err(|_| OffcryptoError::InvalidCiphertextLength { len: pt_len })?;
            }
            32 => {
                let decryptor =
                    Decryptor::<Aes256>::new_from_slices(secret_key, &iv).map_err(|_| {
                        OffcryptoError::InvalidKeyLength {
                            len: secret_key.len(),
                        }
                    })?;
                decryptor
                    .decrypt_padded_mut::<NoPadding>(plaintext)
                    .map_err(|_| OffcryptoError::InvalidCiphertextLength { len: pt_len })?;
            }
            _ => {
                return Err(OffcryptoError::InvalidKeyLength {
                    len: secret_key.len(),
                })
            }
        }

        Ok(())
    })
}

/// Decrypt a Standard (CryptoAPI) `EncryptedPackage` stream using AES-ECB.
///
/// The key is the output of `standard_derive_key`.
pub fn decrypt_standard_encrypted_package(
    key: &[u8],
    encrypted_package: &[u8],
) -> Result<Vec<u8>, OffcryptoError> {
    crate::validate_standard_encrypted_package_stream(encrypted_package)?;
    decrypt_encrypted_package(encrypted_package, |_idx, ct, pt| {
        pt.copy_from_slice(ct);
        crate::aes_ecb_decrypt_in_place(key, pt)
    })
}

fn looks_like_zip_prefix(bytes: &[u8]) -> bool {
    if bytes.len() >= 4 {
        matches!(&bytes[..4], b"PK\x03\x04" | b"PK\x05\x06" | b"PK\x07\x08")
    } else {
        bytes.len() >= 2 && &bytes[..2] == b"PK"
    }
}

fn derive_standard_iv(salt: &[u8], segment_index: u32) -> [u8; 16] {
    let mut hasher = Sha1::new();
    hasher.update(salt);
    hasher.update(&segment_index.to_le_bytes());
    let digest = hasher.finalize();
    let mut iv = [0u8; 16];
    iv.copy_from_slice(&digest[..AES_BLOCK_LEN]);
    iv
}

fn aes_cbc_decrypt_in_place(
    key: &[u8],
    iv: &[u8; 16],
    buf: &mut [u8],
) -> Result<(), OffcryptoError> {
    if buf.len() % AES_BLOCK_LEN != 0 {
        return Err(OffcryptoError::InvalidCiphertextLength { len: buf.len() });
    }
    let len = buf.len();

    match key.len() {
        16 => {
            let decryptor = Decryptor::<Aes128>::new_from_slices(key, iv)
                .map_err(|_| OffcryptoError::InvalidKeyLength { len: key.len() })?;
            decryptor
                .decrypt_padded_mut::<NoPadding>(buf)
                .map_err(|_| OffcryptoError::InvalidCiphertextLength { len })?;
        }
        24 => {
            let decryptor = Decryptor::<Aes192>::new_from_slices(key, iv)
                .map_err(|_| OffcryptoError::InvalidKeyLength { len: key.len() })?;
            decryptor
                .decrypt_padded_mut::<NoPadding>(buf)
                .map_err(|_| OffcryptoError::InvalidCiphertextLength { len })?;
        }
        32 => {
            let decryptor = Decryptor::<Aes256>::new_from_slices(key, iv)
                .map_err(|_| OffcryptoError::InvalidKeyLength { len: key.len() })?;
            decryptor
                .decrypt_padded_mut::<NoPadding>(buf)
                .map_err(|_| OffcryptoError::InvalidCiphertextLength { len })?;
        }
        _ => return Err(OffcryptoError::InvalidKeyLength { len: key.len() }),
    }

    Ok(())
}

/// Decrypt a Standard (CryptoAPI) `EncryptedPackage` stream using segmented AES-CBC.
///
/// This scheme is observed in the wild for Standard-encrypted OOXML and matches the per-segment IV
/// derivation used by Agile encryption:
/// `IV = SHA1(salt || u32le(segment_index))[:16]`.
///
/// The key is the output of `standard_derive_key`; the salt comes from `EncryptionVerifier.salt`.
pub fn decrypt_standard_encrypted_package_cbc(
    key: &[u8],
    salt: &[u8],
    encrypted_package: &[u8],
) -> Result<Vec<u8>, OffcryptoError> {
    crate::validate_standard_encrypted_package_stream(encrypted_package)?;
    decrypt_encrypted_package(encrypted_package, |segment_index, ciphertext, plaintext| {
        let iv = derive_standard_iv(salt, segment_index);
        plaintext.copy_from_slice(ciphertext);
        aes_cbc_decrypt_in_place(key, &iv, plaintext)
    })
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum StandardEncryptedPackageScheme {
    Ecb,
    CbcSegmented,
}

fn detect_standard_scheme(
    key: &[u8],
    salt: &[u8],
    encrypted_package: &[u8],
) -> Result<Option<StandardEncryptedPackageScheme>, OffcryptoError> {
    // Only decrypt the first segment for detection (avoids allocating the full output twice).
    let mut reader = Reader::new(encrypted_package);
    let size_bytes = reader.take(8, "EncryptedPackageHeader.original_size")?;
    let len_lo = u32::from_le_bytes([size_bytes[0], size_bytes[1], size_bytes[2], size_bytes[3]])
        as u64;
    let len_hi = u32::from_le_bytes([size_bytes[4], size_bytes[5], size_bytes[6], size_bytes[7]])
        as u64;
    let total_size_u64 = len_lo | (len_hi << 32);

    // Enforce basic framing invariants up front.
    let ciphertext_total = reader.remaining();
    let ciphertext_total_len = ciphertext_total.len() as u64;
    let total_size = if len_hi != 0
        && total_size_u64 > ciphertext_total_len
        && len_lo != 0
        && len_lo <= ciphertext_total_len
    {
        len_lo
    } else {
        total_size_u64
    };
    let plaintext_len = std::cmp::min(total_size, ENCRYPTED_PACKAGE_SEGMENT_LEN as u64) as usize;
    let ciphertext_len = padded_aes_len(plaintext_len);
    if ciphertext_total.len() % AES_BLOCK_LEN != 0 {
        return Err(OffcryptoError::InvalidCiphertextLength {
            len: ciphertext_total.len(),
        });
    }
    let ciphertext = reader.take(ciphertext_len, "EncryptedPackage.ciphertext_segment")?;

    // `ciphertext_len <= 4096 + 15`; use a fixed stack buffer.
    let mut buf = [0u8; ENCRYPTED_PACKAGE_SEGMENT_LEN + AES_BLOCK_LEN];

    buf[..ciphertext_len].copy_from_slice(ciphertext);
    crate::aes_ecb_decrypt_in_place(key, &mut buf[..ciphertext_len])?;
    let ecb_ok = looks_like_zip_prefix(&buf[..plaintext_len]);

    buf[..ciphertext_len].copy_from_slice(ciphertext);
    let iv = derive_standard_iv(salt, 0);
    aes_cbc_decrypt_in_place(key, &iv, &mut buf[..ciphertext_len])?;
    let cbc_ok = looks_like_zip_prefix(&buf[..plaintext_len]);

    Ok(match (ecb_ok, cbc_ok) {
        (true, false) => Some(StandardEncryptedPackageScheme::Ecb),
        (false, true) => Some(StandardEncryptedPackageScheme::CbcSegmented),
        (true, true) => Some(StandardEncryptedPackageScheme::Ecb),
        (false, false) => None,
    })
}

/// Decrypt a Standard (CryptoAPI) `EncryptedPackage` stream using scheme auto-detection.
///
/// Standard-encrypted OOXML packages are ZIP archives; we therefore detect the correct decryption
/// scheme by attempting both ECB and segmented CBC on the first package segment and selecting the
/// scheme that yields a `PK..` ZIP signature.
pub fn decrypt_standard_encrypted_package_auto(
    key: &[u8],
    salt: &[u8],
    encrypted_package: &[u8],
) -> Result<Vec<u8>, OffcryptoError> {
    crate::validate_standard_encrypted_package_stream(encrypted_package)?;

    let scheme = detect_standard_scheme(key, salt, encrypted_package)?;
    match scheme {
        Some(StandardEncryptedPackageScheme::Ecb) => {
            decrypt_standard_encrypted_package(key, encrypted_package)
        }
        Some(StandardEncryptedPackageScheme::CbcSegmented) => {
            decrypt_standard_encrypted_package_cbc(key, salt, encrypted_package)
        }
        None => Err(OffcryptoError::InvalidPassword),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::test_alloc::MAX_ALLOC;
    use crate::OffcryptoError;
    use aes::cipher::{generic_array::GenericArray, BlockEncrypt, KeyInit};
    use aes::Aes128;
    use std::sync::atomic::Ordering;

    fn aes128_ecb_encrypt_in_place(key: &[u8; 16], buf: &mut [u8]) {
        assert_eq!(buf.len() % 16, 0);
        let cipher = Aes128::new_from_slice(key).expect("valid AES-128 key");
        for block in buf.chunks_mut(16) {
            cipher.encrypt_block(GenericArray::from_mut_slice(block));
        }
    }

    #[test]
    fn encrypted_package_shorter_than_header_returns_truncated() {
        let bytes = [0u8; 7];
        let err = decrypt_encrypted_package(&bytes, |_idx, _ct, _pt| {
            panic!("decryptor should not be invoked for truncated header");
        })
        .expect_err("expected truncated header");
        assert_eq!(
            err,
            OffcryptoError::Truncated {
                context: "EncryptedPackageHeader.original_size"
            }
        );
    }

    #[test]
    fn ciphertext_length_not_multiple_of_16_errors_before_invoking_decryptor() {
        // total_size=0, but provide a non-block-aligned ciphertext tail (15 bytes).
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&0u64.to_le_bytes());
        bytes.extend_from_slice(&[0u8; 15]);

        let err = decrypt_encrypted_package(bytes.as_slice(), |_idx, _ct, _pt| {
            panic!("decryptor should not be invoked for invalid ciphertext framing");
        })
        .expect_err("expected invalid ciphertext length");

        assert_eq!(err, OffcryptoError::InvalidCiphertextLength { len: 15 });
    }

    #[test]
    fn truncated_ciphertext_for_large_total_size_errors_without_large_allocation() {
        // `original_size` is attacker-controlled. Ensure we reject obviously truncated ciphertext
        // *before* attempting to allocate the full output buffer.
        let total_size: u64 = 100 * 1024 * 1024; // 100MiB

        let mut bytes = Vec::new();
        bytes.extend_from_slice(&total_size.to_le_bytes());
        bytes.extend_from_slice(&[0u8; 16]); // 1 AES block of ciphertext (far too short)

        MAX_ALLOC.store(0, Ordering::Relaxed);

        let err = decrypt_encrypted_package(bytes.as_slice(), |_idx, _ct, _pt| Ok(()))
            .expect_err("expected truncated ciphertext");
        assert!(
            matches!(
                err,
                OffcryptoError::EncryptedPackageSizeMismatch {
                    total_size: got_size,
                    ciphertext_len: got_ct
                } if got_size == total_size && got_ct == 16
            ),
            "expected EncryptedPackageSizeMismatch({total_size}, 16), got {err:?}"
        );

        let max_alloc = MAX_ALLOC.load(Ordering::Relaxed);
        assert!(
            max_alloc < 10 * 1024 * 1024,
            "expected no large allocations, observed max allocation request: {max_alloc} bytes"
        );
    }

    #[test]
    fn decrypt_encrypted_package_identity_reads_size_and_payload() {
        let total_size: u64 = 10;
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&total_size.to_le_bytes());
        bytes.extend_from_slice(b"0123456789");
        // Pad to AES block size to match `EncryptedPackage` padding rules.
        bytes.extend_from_slice(&[0u8; 6]);

        let out = decrypt_encrypted_package(bytes.as_slice(), |_idx, ct, pt| {
            pt.copy_from_slice(ct);
            Ok(())
        })
        .expect("decrypt");

        assert_eq!(out, b"0123456789");
    }

    #[test]
    fn decrypt_encrypted_package_falls_back_to_low_dword_when_high_dword_is_reserved() {
        // Some producers treat the 8-byte size prefix as (u32 totalSize, u32 reserved). Ensure we
        // tolerate a non-zero "reserved" high DWORD.
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&10u32.to_le_bytes()); // size (low DWORD)
        bytes.extend_from_slice(&1u32.to_le_bytes()); // reserved (high DWORD)
        bytes.extend_from_slice(b"0123456789");
        // Pad to AES block size to match `EncryptedPackage` padding rules.
        bytes.extend_from_slice(&[0u8; 6]);

        let out = decrypt_encrypted_package(bytes.as_slice(), |_idx, ct, pt| {
            pt.copy_from_slice(ct);
            Ok(())
        })
        .expect("decrypt");

        assert_eq!(out, b"0123456789");
    }

    #[test]
    fn decrypt_standard_encrypted_package_round_trip_truncates_to_total_size() {
        let key = [0x42u8; 16];
        let plaintext: Vec<u8> = (0u8..37).collect();
        let total_size = plaintext.len() as u64;

        let padded_len = ((plaintext.len() + 15) / 16) * 16;
        let mut ciphertext = plaintext.clone();
        ciphertext.resize(padded_len, 0);
        aes128_ecb_encrypt_in_place(&key, &mut ciphertext);

        let mut encrypted_package = Vec::new();
        encrypted_package.extend_from_slice(&total_size.to_le_bytes());
        encrypted_package.extend_from_slice(&ciphertext);

        let out = decrypt_standard_encrypted_package(&key, &encrypted_package).expect("decrypt");
        assert_eq!(out, plaintext);
    }

    #[test]
    fn decrypt_standard_encrypted_package_rejects_non_block_aligned_ciphertext() {
        let key = [0u8; 16];
        let mut encrypted_package = 0u64.to_le_bytes().to_vec();
        encrypted_package.extend_from_slice(&[0u8; 15]);

        let err =
            decrypt_standard_encrypted_package(&key, &encrypted_package).expect_err("expected error");
        assert_eq!(err, OffcryptoError::InvalidCiphertextLength { len: 15 });
    }

    #[test]
    fn oversized_total_size_errors_without_large_allocation() {
        let total_size: u64 = if usize::BITS < 64 {
            (usize::MAX as u64) + 1
        } else {
            u64::MAX
        };

        let mut bytes = Vec::new();
        bytes.extend_from_slice(&total_size.to_le_bytes());

        MAX_ALLOC.store(0, Ordering::Relaxed);

        let err = decrypt_encrypted_package(bytes.as_slice(), |_idx, _ct, _pt| Ok(()))
            .expect_err("expected output too large");

        assert!(
            matches!(
                err,
                OffcryptoError::OutputTooLarge { total_size: got, max }
                    if got == total_size && max == crate::MAX_ENCRYPTED_PACKAGE_ORIGINAL_SIZE
            ),
            "expected OutputTooLarge({total_size}, {}), got {err:?}",
            crate::MAX_ENCRYPTED_PACKAGE_ORIGINAL_SIZE
        );

        let max_alloc = MAX_ALLOC.load(Ordering::Relaxed);
        assert!(
            max_alloc < 1024 * 1024,
            "expected no large allocations, observed max allocation request: {max_alloc} bytes"
        );
    }

    #[test]
    fn decrypt_encrypted_package_ecb_rejects_oversized_total_size_without_large_allocation() {
        // Ensure the Standard AES-ECB helper rejects header sizes that cannot fit into a `Vec<u8>`
        // even if the ciphertext is empty (avoid depending on ciphertext mismatch errors).
        let total_size: u64 = if usize::BITS < 64 {
            (usize::MAX as u64) + 1
        } else {
            u64::MAX
        };

        let mut bytes = Vec::new();
        bytes.extend_from_slice(&total_size.to_le_bytes());

        MAX_ALLOC.store(0, Ordering::Relaxed);

        let key = [0u8; 16];
        let err = crate::decrypt_encrypted_package_ecb(&key, bytes.as_slice())
            .expect_err("expected size overflow");

        assert!(
            matches!(err, OffcryptoError::EncryptedPackageSizeOverflow { total_size: got } if got == total_size),
            "expected EncryptedPackageSizeOverflow({total_size}), got {err:?}"
        );

        let max_alloc = MAX_ALLOC.load(Ordering::Relaxed);
        assert!(
            max_alloc < 1024 * 1024,
            "expected no large allocations, observed max allocation request: {max_alloc} bytes"
        );
    }

    #[test]
    fn output_too_large_errors_without_large_allocation() {
        // Choose a very large size that would be a dangerous allocation on 64-bit platforms.
        let total_size: u64 = 4 * 1024 * 1024 * 1024; // 4GiB
        let max: u64 = 1024 * 1024; // 1MiB

        let mut bytes = Vec::new();
        bytes.extend_from_slice(&total_size.to_le_bytes());

        MAX_ALLOC.store(0, Ordering::Relaxed);

        let opts = crate::DecryptOptions {
            verify_integrity: true,
            limits: crate::DecryptLimits {
                max_output_size: Some(max),
                ..Default::default()
            },
        };
        let err =
            decrypt_encrypted_package_with_options(bytes.as_slice(), &opts, |_idx, _ct, _pt| Ok(()))
                .expect_err("expected output too large");

        assert!(
            matches!(err, OffcryptoError::OutputTooLarge { total_size: got, max: m } if got == total_size && m == max),
            "expected OutputTooLarge({total_size}, {max}), got {err:?}"
        );

        let max_alloc = MAX_ALLOC.load(Ordering::Relaxed);
        assert!(
            max_alloc < 1024 * 1024,
            "expected no large allocations, observed max allocation request: {max_alloc} bytes"
        );
    }
}
