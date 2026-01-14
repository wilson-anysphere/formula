use crate::{AgileEncryptionInfo, OffcryptoError, Reader};

use aes::{Aes128, Aes192, Aes256};
use cbc::Decryptor;
use cipher::{block_padding::NoPadding, BlockDecryptMut, KeyIvInit};

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
    let len =
        usize::try_from(total_size).map_err(|_| OffcryptoError::EncryptedPackageSizeOverflow {
            total_size,
        })?;

    // `Vec<u8>` cannot exceed `isize::MAX` due to `Layout::array`/pointer offset invariants.
    isize::try_from(len).map_err(|_| OffcryptoError::EncryptedPackageSizeOverflow { total_size })?;

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
    mut decrypt_block: F,
) -> Result<Vec<u8>, OffcryptoError>
where
    F: FnMut(u32, &[u8], &mut [u8]) -> Result<(), OffcryptoError>,
{
    let mut reader = Reader::new(encrypted_package);
    let total_size = reader.read_u64_le("EncryptedPackageHeader.original_size")?;
    let output_len = checked_output_len(total_size)?;

    // Validate ciphertext framing before allocating based on attacker-controlled `total_size`.
    //
    // Each ciphertext segment is padded up to the AES block size (16). The total ciphertext length
    // should therefore also be block-aligned, and must be large enough to cover the padded segment
    // lengths implied by `total_size`.
    let ciphertext_len = reader.remaining().len();
    if ciphertext_len % AES_BLOCK_LEN != 0 {
        return Err(OffcryptoError::InvalidCiphertextLength { len: ciphertext_len });
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
        return Err(OffcryptoError::Truncated {
            context: "EncryptedPackage.ciphertext_segment",
        });
    }

    let mut out = Vec::new();
    out.try_reserve_exact(output_len).map_err(|_| {
        OffcryptoError::EncryptedPackageAllocationFailed {
            total_size,
        }
    })?;
    out.resize(output_len, 0);

    let mut remaining: u64 = total_size;
    let mut out_offset: usize = 0;
    let mut block_index: u32 = 0;

    let mut plaintext_buf = [0u8; ENCRYPTED_PACKAGE_SEGMENT_LEN];

    while remaining > 0 {
        let plaintext_len = std::cmp::min(remaining, ENCRYPTED_PACKAGE_SEGMENT_LEN as u64) as usize;
        let ciphertext_len = padded_aes_len(plaintext_len);
        let ciphertext =
            reader.take(ciphertext_len, "EncryptedPackage.ciphertext_segment")?;

        decrypt_block(block_index, ciphertext, &mut plaintext_buf[..ciphertext_len])?;

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
    let size_bytes: [u8; 8] = encrypted_package[..8]
        .try_into()
        .map_err(|_| OffcryptoError::Truncated {
            context: "EncryptedPackageHeader.original_size",
        })?;
    let total_size = u64::from_le_bytes(size_bytes);
    let plausible_max = (encrypted_package.len() as u64).saturating_mul(2);
    if total_size > plausible_max {
        return Err(OffcryptoError::EncryptedPackageSizeOverflow { total_size });
    }

    let mut iv_seed = Vec::with_capacity(info.key_data_salt.len() + 4);

    decrypt_encrypted_package(encrypted_package, |segment_index, ciphertext, plaintext| {
        iv_seed.clear();
        iv_seed.extend_from_slice(&info.key_data_salt);
        iv_seed.extend_from_slice(&segment_index.to_le_bytes());

        let digest = info.key_data_hash_algorithm.digest(&iv_seed);
        let mut iv = [0u8; 16];
        iv.copy_from_slice(&digest[..16]);

        plaintext.copy_from_slice(ciphertext);
        let pt_len = plaintext.len();

        match secret_key.len() {
            16 => {
                let decryptor = Decryptor::<Aes128>::new_from_slices(secret_key, &iv)
                    .map_err(|_| OffcryptoError::InvalidKeyLength { len: secret_key.len() })?;
                decryptor
                    .decrypt_padded_mut::<NoPadding>(plaintext)
                    .map_err(|_| OffcryptoError::InvalidCiphertextLength { len: pt_len })?;
            }
            24 => {
                let decryptor = Decryptor::<Aes192>::new_from_slices(secret_key, &iv)
                    .map_err(|_| OffcryptoError::InvalidKeyLength { len: secret_key.len() })?;
                decryptor
                    .decrypt_padded_mut::<NoPadding>(plaintext)
                    .map_err(|_| OffcryptoError::InvalidCiphertextLength { len: pt_len })?;
            }
            32 => {
                let decryptor = Decryptor::<Aes256>::new_from_slices(secret_key, &iv)
                    .map_err(|_| OffcryptoError::InvalidKeyLength { len: secret_key.len() })?;
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::test_alloc::MAX_ALLOC;
    use crate::OffcryptoError;
    use std::sync::atomic::Ordering;

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
        assert!(matches!(err, OffcryptoError::Truncated { .. }));

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
}
