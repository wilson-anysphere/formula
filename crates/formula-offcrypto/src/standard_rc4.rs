//! Standard (CryptoAPI) RC4 decryption helpers.
//!
//! Standard encryption supports an RC4-CryptoAPI variant (ALG_ID = CALG_RC4). Unlike the AES
//! variants (which use AES-ECB for verifier fields and the package stream), RC4 uses:
//! - RC4 for verifier fields, encrypted sequentially with the block-0 RC4 key
//! - RC4 for the package stream, re-keyed per 512-byte block
//!
//! Key derivation for the RC4 per-block key uses **raw hash truncation**:
//! `Hfinal = Hash(H || LE32(b))`, `rc4_key_b = Hfinal[0..keySize/8]`.

use crate::cryptoapi;
use crate::rc4::Rc4;
use crate::standard::verify_verifier;
use crate::{
    parse_encrypted_package_header, HashAlgorithm, OffcryptoError, StandardEncryptionInfo,
};
use zeroize::Zeroizing;

// CryptoAPI alg ids.
const CALG_RC4: u32 = 0x0000_6801;
const CALG_MD5: u32 = 0x0000_8003;
const CALG_SHA1: u32 = 0x0000_8004;

fn hash_algorithm_from_alg_id_hash(alg_id_hash: u32) -> Result<HashAlgorithm, OffcryptoError> {
    match alg_id_hash {
        CALG_MD5 => Ok(HashAlgorithm::Md5),
        CALG_SHA1 => Ok(HashAlgorithm::Sha1),
        other => Err(OffcryptoError::UnsupportedAlgorithm(format!(
            "algIdHash=0x{other:08x}"
        ))),
    }
}

/// Verify a password against the Standard RC4 verifier fields.
///
/// Returns the CryptoAPI base hash `H` when the verifier check succeeds.
pub fn verify_password(
    info: &StandardEncryptionInfo,
    password: &str,
) -> Result<Zeroizing<Vec<u8>>, OffcryptoError> {
    if info.header.alg_id != CALG_RC4 {
        return Err(OffcryptoError::UnsupportedAlgorithm(format!(
            "algId=0x{:08x}",
            info.header.alg_id
        )));
    }
    let hash_alg = hash_algorithm_from_alg_id_hash(info.header.alg_id_hash)?;

    let h = cryptoapi::iterated_hash_from_password(
        password,
        &info.verifier.salt,
        cryptoapi::STANDARD_SPIN_COUNT,
        hash_alg,
    )?;

    let key0 = cryptoapi::rc4_key_for_block(&h, 0, info.header.key_size_bits, hash_alg)?;
    let mut rc4 = Rc4::new(key0.as_slice());

    let mut verifier = info.verifier.encrypted_verifier;
    rc4.apply_keystream(&mut verifier);

    let mut verifier_hash = Zeroizing::new(info.verifier.encrypted_verifier_hash.clone());
    rc4.apply_keystream(&mut verifier_hash);

    let hash_len = info.verifier.verifier_hash_size as usize;
    if verifier_hash.len() < hash_len {
        return Err(OffcryptoError::InvalidEncryptionInfo {
            context: "decrypted verifierHash shorter than verifierHashSize",
        });
    }

    verify_verifier(&verifier, &verifier_hash[..hash_len], hash_alg)?;
    Ok(h)
}

/// Decrypt an `EncryptedPackage` stream for Standard RC4 encryption using a precomputed base hash.
///
/// `h` must be the CryptoAPI base hash returned by [`verify_password`].
pub(crate) fn decrypt_encrypted_package_with_h(
    info: &StandardEncryptionInfo,
    encrypted_package_stream: &[u8],
    h: &[u8],
) -> Result<Vec<u8>, OffcryptoError> {
    if info.header.alg_id != CALG_RC4 {
        return Err(OffcryptoError::UnsupportedAlgorithm(format!(
            "algId=0x{:08x}",
            info.header.alg_id
        )));
    }
    let hash_alg = hash_algorithm_from_alg_id_hash(info.header.alg_id_hash)?;

    let header = parse_encrypted_package_header(encrypted_package_stream)?;
    let total_size = header.original_size;

    let output_len = usize::try_from(total_size)
        .map_err(|_| OffcryptoError::EncryptedPackageSizeOverflow { total_size })?;
    isize::try_from(output_len)
        .map_err(|_| OffcryptoError::EncryptedPackageSizeOverflow { total_size })?;

    // `original_size` is attacker-controlled. Ensure we reject obviously truncated ciphertext
    // before attempting large allocations.
    let ciphertext_len = encrypted_package_stream.len().saturating_sub(8);
    if ciphertext_len < output_len {
        return Err(OffcryptoError::EncryptedPackageSizeMismatch {
            total_size,
            ciphertext_len,
        });
    }

    let mut out = Vec::new();
    out.try_reserve_exact(output_len)
        .map_err(|_| OffcryptoError::EncryptedPackageAllocationFailed { total_size })?;
    out.resize(output_len, 0);
    let mut remaining = output_len;
    let mut in_offset: usize = 8;
    let mut out_offset: usize = 0;
    let mut block_index: u32 = 0;

    while remaining > 0 {
        let chunk_len = std::cmp::min(remaining, cryptoapi::RC4_BLOCK_LEN);
        let end = in_offset
            .checked_add(chunk_len)
            .ok_or(OffcryptoError::Truncated {
                context: "EncryptedPackage.ciphertext",
            })?;
        if end > encrypted_package_stream.len() {
            return Err(OffcryptoError::Truncated {
                context: "EncryptedPackage.ciphertext",
            });
        }

        out[out_offset..out_offset + chunk_len]
            .copy_from_slice(&encrypted_package_stream[in_offset..end]);

        let key =
            cryptoapi::rc4_key_for_block(h, block_index, info.header.key_size_bits, hash_alg)?;
        let mut rc4 = Rc4::new(key.as_slice());
        rc4.apply_keystream(&mut out[out_offset..out_offset + chunk_len]);

        in_offset = end;
        out_offset += chunk_len;
        remaining -= chunk_len;
        block_index = block_index
            .checked_add(1)
            .ok_or(OffcryptoError::EncryptedPackageSizeOverflow { total_size })?;
    }

    Ok(out)
}

/// Decrypt an `EncryptedPackage` stream for Standard RC4 encryption.
pub fn decrypt_encrypted_package(
    info: &StandardEncryptionInfo,
    encrypted_package_stream: &[u8],
    password: &str,
) -> Result<Vec<u8>, OffcryptoError> {
    let h = verify_password(info, password)?;
    decrypt_encrypted_package_with_h(info, encrypted_package_stream, h.as_slice())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::test_alloc::MAX_ALLOC;
    use crate::{
        StandardEncryptionHeader, StandardEncryptionHeaderFlags, StandardEncryptionVerifier,
    };
    use std::sync::atomic::Ordering;

    #[test]
    fn rc4_decrypt_rejects_truncated_ciphertext_without_large_allocation() {
        // `original_size` is attacker-controlled. Ensure we reject obviously truncated ciphertext
        // *before* attempting to allocate the full output buffer.
        let total_size: u64 = 100 * 1024 * 1024; // 100MiB

        let mut encrypted_package = Vec::new();
        encrypted_package.extend_from_slice(&total_size.to_le_bytes());
        encrypted_package.extend_from_slice(&[0u8; 16]); // far too short

        let info = StandardEncryptionInfo {
            header: StandardEncryptionHeader {
                flags: StandardEncryptionHeaderFlags::from_raw(0),
                size_extra: 0,
                alg_id: CALG_RC4,
                alg_id_hash: CALG_SHA1,
                key_size_bits: 128,
                provider_type: 0,
                reserved1: 0,
                reserved2: 0,
                csp_name: String::new(),
            },
            verifier: StandardEncryptionVerifier {
                salt: vec![0u8; 16],
                encrypted_verifier: [0u8; 16],
                verifier_hash_size: 20,
                encrypted_verifier_hash: vec![0u8; 20],
            },
        };

        // SHA1 digest length for the CryptoAPI base hash `H`.
        let h = vec![0u8; 20];

        MAX_ALLOC.store(0, Ordering::Relaxed);

        let err = decrypt_encrypted_package_with_h(&info, &encrypted_package, &h)
            .expect_err("expected ciphertext size mismatch");
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
}
