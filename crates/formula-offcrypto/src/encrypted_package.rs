use crate::{OffcryptoError, Reader};

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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::test_alloc::MAX_ALLOC;
    use crate::OffcryptoError;
    use std::sync::atomic::Ordering;

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
