//! Agile encryption password helpers.
//!
//! In the Agile encryption scheme (OOXML), password verification is performed by decrypting the
//! `encryptedVerifierHashInput` and `encryptedVerifierHashValue` fields using keys derived from the
//! provided password, then checking:
//!
//! `Hash(verifierHashInput) == verifierHashValue`.
//!
//! The verifier hash value (and other Agile digests like `encryptedHmacValue`) are stored as
//! AES-CBC ciphertext, and decrypt to a buffer that may be padded to a 16-byte boundary when the
//! digest size is not a multiple of 16 (e.g. SHA1=20 bytes). Callers should therefore compare only
//! the digest prefix.
//!
//! The derived password hash uses `spinCount` iterations (commonly 100,000). To avoid recomputing
//! this expensive iterated hash multiple times during a single decryption attempt, this module
//! exposes [`agile_iterated_hash`] and reuses its output when deriving the block keys for:
//! - block 1: `encryptedVerifierHashInput`
//! - block 2: `encryptedVerifierHashValue`
//! - block 3: `encryptedKeyValue` (the secret/package key)

use crate::util::ct_eq;
use crate::{AgileEncryptionInfo, DecryptOptions, HashAlgorithm, OffcryptoError};
use sha1::Digest as _;
use zeroize::Zeroizing;

#[cfg(test)]
use std::cell::Cell;

/// MS-OFFCRYPTO Agile: block key used for deriving the "verifierHashInput" key.
const VERIFIER_HASH_INPUT_BLOCK: [u8; 8] = [0xFE, 0xA7, 0xD2, 0x76, 0x3B, 0x4B, 0x9E, 0x79];
/// MS-OFFCRYPTO Agile: block key used for deriving the "verifierHashValue" key.
const VERIFIER_HASH_VALUE_BLOCK: [u8; 8] = [0xD7, 0xAA, 0x0F, 0x6D, 0x30, 0x61, 0x34, 0x4E];
/// MS-OFFCRYPTO Agile: block key used for deriving the "keyValue" key.
const KEY_VALUE_BLOCK: [u8; 8] = [0x14, 0x6E, 0x0B, 0xE7, 0xAB, 0xAC, 0xD0, 0xD6];

const VERIFIER_HASH_INPUT_LEN: usize = 16;

#[cfg(test)]
thread_local! {
    static ITERATED_HASH_CALLS: Cell<usize> = Cell::new(0);
}

#[cfg(test)]
fn inc_iterated_hash_calls() {
    ITERATED_HASH_CALLS.with(|calls| calls.set(calls.get().saturating_add(1)));
}

#[cfg(test)]
fn reset_iterated_hash_calls() {
    ITERATED_HASH_CALLS.with(|calls| calls.set(0));
}

#[cfg(test)]
fn iterated_hash_call_count() -> usize {
    ITERATED_HASH_CALLS.with(|calls| calls.get())
}

/// Verify the decrypted Agile verifier fields for a candidate password.
///
/// Callers are expected to pass the *decrypted* `verifierHashInput` and `verifierHashValue` fields.
/// `verifierHashValue` may include AES-CBC padding; this function compares only the digest prefix.
pub fn verify_password(
    verifier_hash_input: &[u8],
    verifier_hash_value: &[u8],
    hash_alg: HashAlgorithm,
) -> Result<(), OffcryptoError> {
    let digest_len = hash_alg.digest_len();
    let mut digest_buf = [0u8; 64];
    hash_alg.digest_into(verifier_hash_input, &mut digest_buf[..digest_len]);
    let expected = verifier_hash_value
        .get(..digest_len)
        .ok_or(OffcryptoError::InvalidPassword)?;
    if !ct_eq(&digest_buf[..digest_len], expected) {
        return Err(OffcryptoError::InvalidPassword);
    }
    Ok(())
}

/// Verify an Agile integrity HMAC value.
///
/// `computed_hmac` should be the computed HMAC digest, and `decrypted_hmac_value` should be the
/// decrypted `encryptedHmacValue` field from the Agile `dataIntegrity` element.
///
/// Note: `decrypted_hmac_value` may include AES-CBC padding (e.g. SHA1=20 bytes padded to 32).
/// This function compares only the digest prefix.
pub fn verify_hmac(computed_hmac: &[u8], decrypted_hmac_value: &[u8]) -> Result<(), OffcryptoError> {
    let expected = decrypted_hmac_value
        .get(..computed_hmac.len())
        .ok_or(OffcryptoError::InvalidPassword)?;
    if !ct_eq(computed_hmac, expected) {
        return Err(OffcryptoError::InvalidPassword);
    }
    Ok(())
}

/// Compute the Agile password *iterated hash*.
///
/// Algorithm:
/// 1. `H = Hash(salt || password_utf16le)`
/// 2. For `i in 0..spinCount`: `H = Hash(LE32(i) || H)`
pub fn agile_iterated_hash(
    password_utf16le: &[u8],
    salt: &[u8],
    hash_alg: HashAlgorithm,
    spin_count: u32,
) -> Zeroizing<Vec<u8>> {
    #[cfg(test)]
    inc_iterated_hash_calls();

    let digest_len = hash_alg.digest_len();
    debug_assert!(digest_len <= crate::MAX_DIGEST_LEN);

    // Avoid per-iteration allocations (spinCount is often 100k): keep the current digest in a fixed
    // buffer and overwrite it each round.
    let mut h_buf: Zeroizing<[u8; crate::MAX_DIGEST_LEN]> =
        Zeroizing::new([0u8; crate::MAX_DIGEST_LEN]);
    hash_alg.digest_two_into(salt, password_utf16le, &mut h_buf[..digest_len]);

    match hash_alg {
        HashAlgorithm::Md5 => {
            for i in 0..spin_count {
                let mut hasher = md5::Md5::new();
                hasher.update(i.to_le_bytes());
                hasher.update(&h_buf[..16]);
                h_buf[..16].copy_from_slice(&hasher.finalize());
            }
        }
        HashAlgorithm::Sha1 => {
            for i in 0..spin_count {
                let mut hasher = sha1::Sha1::new();
                hasher.update(i.to_le_bytes());
                hasher.update(&h_buf[..20]);
                h_buf[..20].copy_from_slice(&hasher.finalize());
            }
        }
        HashAlgorithm::Sha256 => {
            for i in 0..spin_count {
                let mut hasher = sha2::Sha256::new();
                hasher.update(i.to_le_bytes());
                hasher.update(&h_buf[..32]);
                h_buf[..32].copy_from_slice(&hasher.finalize());
            }
        }
        HashAlgorithm::Sha384 => {
            for i in 0..spin_count {
                let mut hasher = sha2::Sha384::new();
                hasher.update(i.to_le_bytes());
                hasher.update(&h_buf[..48]);
                h_buf[..48].copy_from_slice(&hasher.finalize());
            }
        }
        HashAlgorithm::Sha512 => {
            for i in 0..spin_count {
                let mut hasher = sha2::Sha512::new();
                hasher.update(i.to_le_bytes());
                hasher.update(&h_buf[..64]);
                h_buf[..64].copy_from_slice(&hasher.finalize());
            }
        }
    }

    Zeroizing::new(h_buf[..digest_len].to_vec())
}

/// Derive and decrypt the Agile secret key (encryptedKeyValue) *with password verification*.
///
/// This decrypt path needs three derived block keys. The expensive iterated password hash is
/// computed once and reused for all key derivations.
pub fn agile_secret_key_from_password(
    info: &AgileEncryptionInfo,
    password: &str,
) -> Result<Zeroizing<Vec<u8>>, OffcryptoError> {
    let options = DecryptOptions::default();
    agile_secret_key_from_password_with_options(info, password, &options)
}

/// Like [`agile_secret_key_from_password`], but allows overriding resource limits.
pub fn agile_secret_key_from_password_with_options(
    info: &AgileEncryptionInfo,
    password: &str,
    options: &DecryptOptions,
) -> Result<Zeroizing<Vec<u8>>, OffcryptoError> {
    if info.password_salt.len() != 16 {
        return Err(OffcryptoError::InvalidEncryptionInfo {
            context: "encryptedKey.saltValue must be 16 bytes",
        });
    }
    if info.password_key_bits == 0 || info.password_key_bits % 8 != 0 {
        return Err(OffcryptoError::InvalidEncryptionInfo {
            context: "encryptedKey.keyBits is not divisible by 8",
        });
    }
    let key_len = info.password_key_bits / 8;

    if info.encrypted_verifier_hash_input.is_empty() || info.encrypted_verifier_hash_value.is_empty()
    {
        return Err(OffcryptoError::InvalidEncryptionInfo {
            context: "missing encryptedVerifierHashInput/encryptedVerifierHashValue",
        });
    }

    // `spinCount` is attacker-controlled; enforce limits up front to avoid CPU DoS.
    crate::check_spin_count(info.spin_count, &options.limits)?;

    let password_utf16le = Zeroizing::new(crate::password_to_utf16le_bytes(password));
    let h = agile_iterated_hash(
        &password_utf16le,
        &info.password_salt,
        info.password_hash_algorithm,
        info.spin_count,
    );

    // Derive keys once; only IV handling changes between modes.
    let key1 = crate::derive_encryption_key(
        &h,
        &VERIFIER_HASH_INPUT_BLOCK,
        info.password_hash_algorithm,
        info.password_key_bits,
    )?;
    let key2 = crate::derive_encryption_key(
        &h,
        &VERIFIER_HASH_VALUE_BLOCK,
        info.password_hash_algorithm,
        info.password_key_bits,
    )?;
    let key3 = crate::derive_encryption_key(
        &h,
        &KEY_VALUE_BLOCK,
        info.password_hash_algorithm,
        info.password_key_bits,
    )?;

    #[derive(Debug, Clone, Copy, PartialEq, Eq)]
    enum PasswordKeyIvMode {
        /// MS-OFFCRYPTO spec behavior: use the raw password `saltValue` as the AES-CBC IV.
        Salt,
        /// Compatibility behavior observed in some producers: use per-blob derived IVs
        /// (`Hash(saltValue || blockKey)[:16]`).
        Derived,
    }

    let try_mode = |mode: PasswordKeyIvMode| -> Result<Zeroizing<Vec<u8>>, OffcryptoError> {
        let decrypt_with_ivs =
            |iv_vhi: &[u8], iv_vhv: &[u8], iv_kv: &[u8]| -> Result<Zeroizing<Vec<u8>>, OffcryptoError> {
                // Decrypt verifierHashInput.
                let verifier_hash_input = crate::aes_cbc_decrypt(
                    &info.encrypted_verifier_hash_input,
                    &key1,
                    iv_vhi,
                )?;
                if verifier_hash_input.len() < VERIFIER_HASH_INPUT_LEN {
                    return Err(OffcryptoError::InvalidEncryptionInfo {
                        context: "decrypted verifierHashInput is truncated",
                    });
                }

                // Decrypt verifierHashValue and verify.
                let verifier_hash_value = crate::aes_cbc_decrypt(
                    &info.encrypted_verifier_hash_value,
                    &key2,
                    iv_vhv,
                )?;

                let digest_len = info.password_hash_algorithm.digest_len();
                if verifier_hash_value.len() < digest_len {
                    return Err(OffcryptoError::InvalidEncryptionInfo {
                        context: "decrypted verifierHashValue is truncated",
                    });
                }
                verify_password(
                    &verifier_hash_input[..VERIFIER_HASH_INPUT_LEN],
                    &verifier_hash_value,
                    info.password_hash_algorithm,
                )?;

                // Decrypt encryptedKeyValue (secret key).
                let key_value = crate::aes_cbc_decrypt(&info.encrypted_key_value, &key3, iv_kv)?;
                if key_value.len() < key_len {
                    return Err(OffcryptoError::InvalidEncryptionInfo {
                        context: "decrypted keyValue is truncated",
                    });
                }
                Ok(Zeroizing::new(key_value[..key_len].to_vec()))
            };

        match mode {
            PasswordKeyIvMode::Salt => decrypt_with_ivs(
                &info.password_salt,
                &info.password_salt,
                &info.password_salt,
            ),
            PasswordKeyIvMode::Derived => {
                let iv1 = crate::derive_iv_from_salt(
                    &info.password_salt,
                    &VERIFIER_HASH_INPUT_BLOCK,
                    info.password_hash_algorithm,
                )?;
                let iv2 = crate::derive_iv_from_salt(
                    &info.password_salt,
                    &VERIFIER_HASH_VALUE_BLOCK,
                    info.password_hash_algorithm,
                )?;
                let iv3 = crate::derive_iv_from_salt(
                    &info.password_salt,
                    &KEY_VALUE_BLOCK,
                    info.password_hash_algorithm,
                )?;
                decrypt_with_ivs(&iv1, &iv2, &iv3)
            }
        }
    };

    match try_mode(PasswordKeyIvMode::Salt) {
        Ok(key) => Ok(key),
        Err(OffcryptoError::InvalidPassword) => try_mode(PasswordKeyIvMode::Derived),
        Err(other) => Err(other),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::util::{ct_eq_call_count, reset_ct_eq_calls};
    use aes::Aes128;
    use cbc::Encryptor;
    use cipher::{block_padding::NoPadding, BlockEncryptMut, KeyIvInit};

    fn zero_pad_to_aes_block(mut bytes: Vec<u8>) -> Vec<u8> {
        let rem = bytes.len() % 16;
        if rem != 0 {
            bytes.resize(bytes.len() + (16 - rem), 0);
        }
        bytes
    }

    fn encrypt_aes128_cbc_no_padding(key: &[u8], iv: &[u8], plaintext: &[u8]) -> Vec<u8> {
        assert_eq!(key.len(), 16);
        assert_eq!(iv.len(), 16);
        assert_eq!(plaintext.len() % 16, 0);
        let mut buf = plaintext.to_vec();
        Encryptor::<Aes128>::new_from_slices(key, iv)
            .unwrap()
            .encrypt_padded_mut::<NoPadding>(&mut buf, plaintext.len())
            .unwrap();
        buf
    }

    #[test]
    fn agile_verify_password_mismatch_returns_invalid_password_and_uses_ct_eq() {
        let input = b"verifier-hash-input";
        for alg in [HashAlgorithm::Sha1, HashAlgorithm::Md5] {
            reset_ct_eq_calls();

            let mut expected = alg.digest(input);
            // Flip a bit to force a mismatch.
            expected[0] ^= 0x01;

            let err = verify_password(input, &expected, alg)
                .expect_err("expected verifier mismatch to return an error");
            assert!(matches!(err, OffcryptoError::InvalidPassword));

            assert!(
                ct_eq_call_count() >= 1,
                "expected constant-time compare helper to be invoked (alg={alg})"
            );
        }
    }

    #[test]
    fn sha1_verifier_hash_value_padding_is_ignored() {
        reset_ct_eq_calls();

        let verifier_input = b"0123456789abcdef"; // 16 bytes
        let digest = HashAlgorithm::Sha1.digest(verifier_input); // 20 bytes

        let mut padded = digest.clone();
        padded.extend([0xA5u8; 12]); // pad to 32 bytes
        assert_eq!(padded.len(), 32);

        verify_password(verifier_input, &padded, HashAlgorithm::Sha1).expect("verify");
        assert!(ct_eq_call_count() >= 1, "expected ct_eq to be used");
    }

    #[test]
    fn sha1_hmac_value_padding_is_ignored() {
        reset_ct_eq_calls();

        let computed = (0u8..20).collect::<Vec<_>>();
        let mut padded = computed.clone();
        padded.extend([0xA5u8; 12]);

        verify_hmac(&computed, &padded).expect("verify");
        assert!(ct_eq_call_count() >= 1, "expected ct_eq to be used");
    }

    #[test]
    fn agile_secret_key_from_password_computes_iterated_hash_once() {
        let password = "password";
        let salt = vec![0x11u8; 16];
        let spin_count = 1000;
        let hash_algorithm = HashAlgorithm::Sha1;
        let key_bits = 128usize;

        // Build an AgileEncryptionInfo with synthetic encrypted verifier + keyValue fields.
        let password_utf16le = crate::password_to_utf16le_bytes(password);
        let h = agile_iterated_hash(&password_utf16le, &salt, hash_algorithm, spin_count);

        let key1 = crate::derive_encryption_key(&h, &VERIFIER_HASH_INPUT_BLOCK, hash_algorithm, key_bits)
            .unwrap();
        let verifier_hash_input_plain = vec![0x22u8; VERIFIER_HASH_INPUT_LEN];
        let encrypted_verifier_hash_input =
            encrypt_aes128_cbc_no_padding(&key1, &salt, &verifier_hash_input_plain);

        let key2 =
            crate::derive_encryption_key(&h, &VERIFIER_HASH_VALUE_BLOCK, hash_algorithm, key_bits)
                .unwrap();
        let digest = hash_algorithm.digest(&verifier_hash_input_plain);
        let verifier_hash_value_plain = zero_pad_to_aes_block(digest.clone());
        let encrypted_verifier_hash_value =
            encrypt_aes128_cbc_no_padding(&key2, &salt, &verifier_hash_value_plain);

        let key3 = crate::derive_encryption_key(&h, &KEY_VALUE_BLOCK, hash_algorithm, key_bits).unwrap();
        let secret_key_plain = vec![0x33u8; key_bits / 8];
        let encrypted_key_value =
            encrypt_aes128_cbc_no_padding(&key3, &salt, &secret_key_plain);

        let info = AgileEncryptionInfo {
            key_data_salt: Vec::new(),
            key_data_hash_algorithm: hash_algorithm,
            key_data_block_size: 16,
            data_integrity: None,
            spin_count,
            password_salt: salt.clone(),
            password_hash_algorithm: hash_algorithm,
            password_key_bits: key_bits,
            encrypted_key_value,
            encrypted_verifier_hash_input,
            encrypted_verifier_hash_value,
        };

        // Setup above called `agile_iterated_hash`.
        reset_iterated_hash_calls();

        let out = agile_secret_key_from_password(&info, password).unwrap();
        assert_eq!(out.as_slice(), secret_key_plain.as_slice());
        assert_eq!(iterated_hash_call_count(), 1);
    }

    #[test]
    fn agile_secret_key_from_password_falls_back_to_derived_iv() {
        let password = "password";
        let salt = vec![0x11u8; 16];
        let spin_count = 1000;
        let hash_algorithm = HashAlgorithm::Sha1;
        let key_bits = 128usize;

        // Build an AgileEncryptionInfo with verifier/keyValue blobs encrypted using derived IVs
        // (Hash(saltValue || blockKey)[:16]) instead of the raw salt bytes.
        let password_utf16le = crate::password_to_utf16le_bytes(password);
        let h = agile_iterated_hash(&password_utf16le, &salt, hash_algorithm, spin_count);

        let key1 = crate::derive_encryption_key(&h, &VERIFIER_HASH_INPUT_BLOCK, hash_algorithm, key_bits)
            .unwrap();
        let key2 =
            crate::derive_encryption_key(&h, &VERIFIER_HASH_VALUE_BLOCK, hash_algorithm, key_bits)
                .unwrap();
        let key3 =
            crate::derive_encryption_key(&h, &KEY_VALUE_BLOCK, hash_algorithm, key_bits).unwrap();

        let iv1 = crate::derive_iv_from_salt(&salt, &VERIFIER_HASH_INPUT_BLOCK, hash_algorithm).unwrap();
        let iv2 = crate::derive_iv_from_salt(&salt, &VERIFIER_HASH_VALUE_BLOCK, hash_algorithm).unwrap();
        let iv3 = crate::derive_iv_from_salt(&salt, &KEY_VALUE_BLOCK, hash_algorithm).unwrap();
        assert_ne!(
            &iv1[..],
            &salt[..],
            "derived-IV scheme should not accidentally match the raw salt IV"
        );

        let verifier_hash_input_plain = vec![0x22u8; VERIFIER_HASH_INPUT_LEN];
        let encrypted_verifier_hash_input =
            encrypt_aes128_cbc_no_padding(&key1, &iv1, &verifier_hash_input_plain);

        let digest = hash_algorithm.digest(&verifier_hash_input_plain);
        let verifier_hash_value_plain = zero_pad_to_aes_block(digest.clone());
        let encrypted_verifier_hash_value =
            encrypt_aes128_cbc_no_padding(&key2, &iv2, &verifier_hash_value_plain);

        let secret_key_plain = vec![0x33u8; key_bits / 8];
        let encrypted_key_value = encrypt_aes128_cbc_no_padding(&key3, &iv3, &secret_key_plain);

        let info = AgileEncryptionInfo {
            key_data_salt: Vec::new(),
            key_data_hash_algorithm: hash_algorithm,
            key_data_block_size: 16,
            data_integrity: None,
            spin_count,
            password_salt: salt,
            password_hash_algorithm: hash_algorithm,
            password_key_bits: key_bits,
            encrypted_key_value,
            encrypted_verifier_hash_input,
            encrypted_verifier_hash_value,
        };

        // Setup above called `agile_iterated_hash`.
        reset_iterated_hash_calls();

        let out = agile_secret_key_from_password(&info, password).unwrap();
        assert_eq!(out.as_slice(), secret_key_plain.as_slice());
        assert_eq!(
            iterated_hash_call_count(),
            1,
            "expected iterated hash to be computed once even when trying both IV schemes"
        );
    }

    #[test]
    fn agile_secret_key_from_password_rejects_spin_count_over_limit_without_hashing() {
        let password = "password";
        let salt = vec![0x11u8; 16];
        let spin_count = u32::MAX;
        let hash_algorithm = HashAlgorithm::Sha1;
        let key_bits = 128usize;

        let info = AgileEncryptionInfo {
            key_data_salt: Vec::new(),
            key_data_hash_algorithm: hash_algorithm,
            key_data_block_size: 16,
            data_integrity: None,
            spin_count,
            password_salt: salt,
            password_hash_algorithm: hash_algorithm,
            password_key_bits: key_bits,
            encrypted_key_value: vec![0u8; 16],
            encrypted_verifier_hash_input: vec![0u8; 16],
            encrypted_verifier_hash_value: vec![0u8; 16],
        };

        let options = DecryptOptions {
            verify_integrity: true,
            limits: crate::DecryptLimits {
                max_spin_count: Some(10),
                ..Default::default()
            },
        };

        reset_iterated_hash_calls();

        let err = agile_secret_key_from_password_with_options(&info, password, &options)
            .expect_err("expected spinCount over limit to error");
        assert_eq!(
            err,
            OffcryptoError::SpinCountTooLarge {
                spin_count,
                max: 10
            }
        );
        assert_eq!(
            iterated_hash_call_count(),
            0,
            "expected iterated hash to not run when spinCount is over the limit"
        );
    }
}
