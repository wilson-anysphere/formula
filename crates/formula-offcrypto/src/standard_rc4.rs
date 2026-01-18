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

    // Some producers incorrectly treat export-grade 40-bit keys as 128-bit RC4 keys with trailing
    // zeros (changing RC4 KSA). For compatibility, attempt both verifier decryption schemes when the
    // declared key size implies a 40-bit key.
    let key_bits = if info.header.key_size_bits == 0 {
        40
    } else {
        info.header.key_size_bits
    };
    let pad_candidates: &[bool] = if key_bits == 40 { &[false, true] } else { &[false] };

    for &pad_40_bit_to_128 in pad_candidates {
        match verify_password_with_h(info, h.as_slice(), hash_alg, pad_40_bit_to_128) {
            Ok(()) => return Ok(h),
            Err(OffcryptoError::InvalidPassword) => continue,
            Err(other) => return Err(other),
        }
    }

    Err(OffcryptoError::InvalidPassword)
}

fn rc4_key_for_block_with_optional_padding(
    h: &[u8],
    block: u32,
    key_size_bits: u32,
    hash_alg: HashAlgorithm,
    pad_40_bit_to_128: bool,
) -> Result<Zeroizing<Vec<u8>>, OffcryptoError> {
    let key = cryptoapi::rc4_key_for_block(h, block, key_size_bits, hash_alg)?;
    if !pad_40_bit_to_128 {
        return Ok(key);
    }

    let effective_bits = if key_size_bits == 0 { 40 } else { key_size_bits };
    if effective_bits != 40 {
        return Ok(key);
    }

    // Compatibility: Some producers (notably some CryptoAPI implementations) treat a 40-bit RC4
    // key as a 128-bit key with the high 88 bits set to zero.
    //
    // This changes the RC4 key scheduling because RC4 includes the key length; it is *not*
    // equivalent to using the 5-byte key directly.
    let mut padded = vec![0u8; 16];
    padded[..key.len()].copy_from_slice(&key);
    Ok(Zeroizing::new(padded))
}

fn verify_password_with_h(
    info: &StandardEncryptionInfo,
    h: &[u8],
    hash_alg: HashAlgorithm,
    pad_40_bit_to_128: bool,
) -> Result<(), OffcryptoError> {
    let key0 = rc4_key_for_block_with_optional_padding(
        h,
        0,
        info.header.key_size_bits,
        hash_alg,
        pad_40_bit_to_128,
    )?;
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

    verify_verifier(&verifier, &verifier_hash[..hash_len], hash_alg)
}

/// Decrypt an `EncryptedPackage` stream for Standard RC4 encryption using a precomputed base hash.
///
/// `h` must be the CryptoAPI base hash returned by [`verify_password`].
pub(crate) fn decrypt_encrypted_package_with_h(
    info: &StandardEncryptionInfo,
    encrypted_package_stream: &[u8],
    h: &[u8],
) -> Result<Vec<u8>, OffcryptoError> {
    decrypt_encrypted_package_with_h_and_padding(info, encrypted_package_stream, h, false)
}

fn decrypt_encrypted_package_with_h_and_padding(
    info: &StandardEncryptionInfo,
    encrypted_package_stream: &[u8],
    h: &[u8],
    pad_40_bit_to_128: bool,
) -> Result<Vec<u8>, OffcryptoError> {
    if info.header.alg_id != CALG_RC4 {
        return Err(OffcryptoError::UnsupportedAlgorithm(format!(
            "algId=0x{:08x}",
            info.header.alg_id
        )));
    }
    let hash_alg = hash_algorithm_from_alg_id_hash(info.header.alg_id_hash)?;

    // MS-OFFCRYPTO describes `EncryptedPackage.original_size` as a `u64le`, but some
    // producers/libraries treat it as `(u32 totalSize, u32 reserved)` and may leave the high DWORD
    // non-zero. Mirror the heuristic used by the AES decryptors: when the high DWORD is non-zero
    // and the combined 64-bit size is not plausible for the available ciphertext, fall back to the
    // low DWORD (only when it is non-zero, so true 64-bit sizes that are exact multiples of 2^32
    // are not misinterpreted).
    parse_encrypted_package_header(encrypted_package_stream)?;
    let header_bytes: [u8; 8] = encrypted_package_stream
        .get(..8)
        .ok_or(OffcryptoError::Truncated {
            context: "EncryptedPackageHeader.original_size",
        })?
        .try_into()
        .map_err(|_| OffcryptoError::Truncated {
            context: "EncryptedPackageHeader.original_size",
        })?;
    let len_lo =
        u32::from_le_bytes([header_bytes[0], header_bytes[1], header_bytes[2], header_bytes[3]])
            as u64;
    let len_hi =
        u32::from_le_bytes([header_bytes[4], header_bytes[5], header_bytes[6], header_bytes[7]])
            as u64;
    let total_size_u64 = len_lo | (len_hi << 32);
    let ciphertext_len_u64 = encrypted_package_stream.len().saturating_sub(8) as u64;
    let total_size =
        if len_lo != 0 && len_hi != 0 && total_size_u64 > ciphertext_len_u64 && len_lo <= ciphertext_len_u64 {
            len_lo
        } else {
            total_size_u64
        };

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

        let out_end = out_offset
            .checked_add(chunk_len)
            .ok_or(OffcryptoError::EncryptedPackageSizeOverflow { total_size })?;
        let dst = out
            .get_mut(out_offset..out_end)
            .ok_or(OffcryptoError::EncryptedPackageSizeOverflow { total_size })?;
        let src = encrypted_package_stream.get(in_offset..end).ok_or(OffcryptoError::Truncated {
            context: "EncryptedPackage.ciphertext",
        })?;
        dst.copy_from_slice(src);

        let key = rc4_key_for_block_with_optional_padding(
            h,
            block_index,
            info.header.key_size_bits,
            hash_alg,
            pad_40_bit_to_128,
        )?;
        let mut rc4 = Rc4::new(key.as_slice());
        rc4.apply_keystream(dst);

        in_offset = end;
        out_offset = out_end;
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

    // Some 40-bit RC4 producers pad the derived key to 128-bit by appending zeros. For maximum
    // compatibility, attempt both derivations when keySize indicates 40-bit RC4.
    let key_bits = if info.header.key_size_bits == 0 {
        40
    } else {
        info.header.key_size_bits
    };
    let pad_candidates: &[bool] = if key_bits == 40 { &[false, true] } else { &[false] };

    for &pad_40_bit_to_128 in pad_candidates {
        match verify_password_with_h(info, h.as_slice(), hash_alg, pad_40_bit_to_128) {
            Ok(()) => {
                if !pad_40_bit_to_128 {
                    // Use the shared implementation when we're not applying any key-length
                    // compatibility hacks.
                    return decrypt_encrypted_package_with_h(info, encrypted_package_stream, h.as_slice());
                }
                return decrypt_encrypted_package_with_h_and_padding(
                    info,
                    encrypted_package_stream,
                    h.as_slice(),
                    pad_40_bit_to_128,
                );
            }
            Err(OffcryptoError::InvalidPassword) => continue,
            Err(other) => return Err(other),
        }
    }

    Err(OffcryptoError::InvalidPassword)
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

    #[test]
    fn rc4_decrypt_falls_back_to_low_dword_when_high_dword_is_reserved() {
        // Some producers treat the 8-byte size prefix as `(u32 totalSize, u32 reserved)`. When the
        // high DWORD is non-zero but implausible for the ciphertext length, fall back to the low
        // DWORD (compatibility).

        let plaintext_len: usize = 1000;
        let plaintext: Vec<u8> = (0..plaintext_len).map(|i| (i % 256) as u8).collect();

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

        // Encrypt the plaintext using the same per-block RC4 scheme (RC4 is symmetric).
        let mut ciphertext = plaintext.clone();
        for (block, chunk) in ciphertext.chunks_mut(cryptoapi::RC4_BLOCK_LEN).enumerate() {
            let key = cryptoapi::rc4_key_for_block(&h, block as u32, info.header.key_size_bits, HashAlgorithm::Sha1)
                .expect("rc4 key");
            let mut rc4 = Rc4::new(key.as_slice());
            rc4.apply_keystream(chunk);
        }

        // Build `EncryptedPackage` with a non-zero high DWORD.
        let mut encrypted_package = Vec::new();
        encrypted_package.extend_from_slice(&(plaintext_len as u32).to_le_bytes());
        encrypted_package.extend_from_slice(&1u32.to_le_bytes()); // reserved/high DWORD
        encrypted_package.extend_from_slice(&ciphertext);

        MAX_ALLOC.store(0, Ordering::Relaxed);

        let decrypted = decrypt_encrypted_package_with_h(&info, &encrypted_package, &h)
            .expect("decrypt with reserved high DWORD");
        assert_eq!(decrypted, plaintext);

        let max_alloc = MAX_ALLOC.load(Ordering::Relaxed);
        assert!(
            max_alloc < 10 * 1024 * 1024,
            "expected no large allocations, observed max allocation request: {max_alloc} bytes"
        );
    }

    #[test]
    fn verify_password_accepts_zero_padded_40_bit_rc4_keys_for_compatibility() {
        // MS-OFFCRYPTO Standard RC4 specifies that 40-bit keys are 5-byte RC4 keys derived from the
        // per-block digest. However, some producers incorrectly treat the 40-bit key material as a
        // 16-byte RC4 key with trailing zeros. Ensure we accept that variant when validating the
        // verifier.

        let password = "password";
        let wrong_password = "wrong password";

        // Use the deterministic MD5 vector from `docs/offcrypto-standard-cryptoapi-rc4.md`.
        let salt: [u8; 16] = [
            0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D,
            0x0E, 0x0F,
        ];
        let h: [u8; 16] = [
            0x20, 0x79, 0x47, 0x60, 0x89, 0xFD, 0xA7, 0x84, 0xC3, 0xA3, 0xCF, 0xEB, 0x98, 0x10,
            0x2C, 0x7E,
        ];

        let key_size_bits: u32 = 0; // interpreted as 40-bit for RC4
        let hash_alg = HashAlgorithm::Md5;

        // Encrypt the verifier fields using the padded 16-byte key scheme.
        let key_material =
            cryptoapi::rc4_key_for_block(&h, 0, key_size_bits, hash_alg).expect("rc4 key");
        assert_eq!(key_material.len(), 5, "expected 40-bit key material length");
        let mut padded_key = vec![0u8; 16];
        padded_key[..key_material.len()].copy_from_slice(key_material.as_slice());

        let verifier_plain: [u8; 16] = [
            0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D,
            0x0E, 0x0F,
        ];
        let verifier_hash_plain = hash_alg.digest(&verifier_plain);
        assert_eq!(
            verifier_hash_plain.len(),
            16,
            "expected MD5 verifier digest length"
        );

        let mut rc4 = Rc4::new(&padded_key);
        let mut encrypted_verifier = verifier_plain;
        rc4.apply_keystream(&mut encrypted_verifier);
        let mut encrypted_verifier_hash = verifier_hash_plain.clone();
        rc4.apply_keystream(&mut encrypted_verifier_hash);

        let info = StandardEncryptionInfo {
            header: StandardEncryptionHeader {
                flags: StandardEncryptionHeaderFlags::from_raw(0),
                size_extra: 0,
                alg_id: CALG_RC4,
                alg_id_hash: CALG_MD5,
                key_size_bits,
                provider_type: 0,
                reserved1: 0,
                reserved2: 0,
                csp_name: String::new(),
            },
            verifier: StandardEncryptionVerifier {
                salt: salt.to_vec(),
                encrypted_verifier,
                verifier_hash_size: 16,
                encrypted_verifier_hash,
            },
        };

        // Sanity check: without padding, verifier decryption should fail for this construction.
        let err = verify_password_with_h(&info, &h, hash_alg, false)
            .expect_err("expected raw 5-byte key verifier to fail");
        assert!(matches!(err, OffcryptoError::InvalidPassword));
        verify_password_with_h(&info, &h, hash_alg, true).expect("padded verifier should succeed");

        let derived = verify_password(&info, password).expect("password should verify");
        assert_eq!(derived.as_slice(), &h);

        let err = verify_password(&info, wrong_password).expect_err("expected wrong password error");
        assert!(matches!(err, OffcryptoError::InvalidPassword));
    }
}
