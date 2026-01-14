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
use crate::{AgileEncryptionInfo, HashAlgorithm, OffcryptoError};
use zeroize::Zeroizing;

#[cfg(test)]
use std::sync::atomic::{AtomicUsize, Ordering};

/// MS-OFFCRYPTO Agile: block key used for deriving the "verifierHashInput" key.
const VERIFIER_HASH_INPUT_BLOCK: [u8; 8] = [0xFE, 0xA7, 0xD2, 0x76, 0x3B, 0x4B, 0x9E, 0x79];
/// MS-OFFCRYPTO Agile: block key used for deriving the "verifierHashValue" key.
const VERIFIER_HASH_VALUE_BLOCK: [u8; 8] = [0xD7, 0xAA, 0x0F, 0x6D, 0x30, 0x61, 0x34, 0x4E];
/// MS-OFFCRYPTO Agile: block key used for deriving the "keyValue" key.
const KEY_VALUE_BLOCK: [u8; 8] = [0x14, 0x6E, 0x0B, 0xE7, 0xAB, 0xAC, 0xD0, 0xD6];

const VERIFIER_HASH_INPUT_LEN: usize = 16;

#[cfg(test)]
static ITERATED_HASH_CALLS: AtomicUsize = AtomicUsize::new(0);

fn hash_output_len(hash_alg: HashAlgorithm) -> usize {
    match hash_alg {
        HashAlgorithm::Sha1 => 20,
        HashAlgorithm::Sha256 => 32,
        HashAlgorithm::Sha384 => 48,
        HashAlgorithm::Sha512 => 64,
    }
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
    let digest = hash_alg.digest(verifier_hash_input);
    let expected = verifier_hash_value
        .get(..digest.len())
        .ok_or(OffcryptoError::InvalidPassword)?;
    if !ct_eq(&digest, expected) {
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
    ITERATED_HASH_CALLS.fetch_add(1, Ordering::Relaxed);

    fn hash_into(hash_alg: HashAlgorithm, data: &[u8], out: &mut [u8]) {
        match hash_alg {
            HashAlgorithm::Sha1 => {
                use sha1::Digest as _;
                let digest = sha1::Sha1::digest(data);
                out.copy_from_slice(&digest);
            }
            HashAlgorithm::Sha256 => {
                use sha2::Digest as _;
                let digest = sha2::Sha256::digest(data);
                out.copy_from_slice(&digest);
            }
            HashAlgorithm::Sha384 => {
                use sha2::Digest as _;
                let digest = sha2::Sha384::digest(data);
                out.copy_from_slice(&digest);
            }
            HashAlgorithm::Sha512 => {
                use sha2::Digest as _;
                let digest = sha2::Sha512::digest(data);
                out.copy_from_slice(&digest);
            }
        }
    }

    let digest_len = hash_output_len(hash_alg);
    let mut h = Zeroizing::new(vec![0u8; digest_len]);

    // Initial round: Hash(salt || password_utf16le)
    let mut buf = Zeroizing::new(Vec::with_capacity(salt.len() + password_utf16le.len()));
    buf.extend_from_slice(salt);
    buf.extend_from_slice(password_utf16le);
    hash_into(hash_alg, &buf[..], &mut h[..]);

    // Iteration 0..spinCount-1: Hash(LE32(i) || H)
    //
    // Avoid allocating in the loop: reuse a fixed-size buffer and overwrite the hash output.
    let mut round = Zeroizing::new(vec![0u8; 4 + digest_len]);
    for i in 0..spin_count {
        round[..4].copy_from_slice(&i.to_le_bytes());
        round[4..].copy_from_slice(&h);
        hash_into(hash_alg, &round[..], &mut h[..]);
    }

    h
}

/// Derive and decrypt the Agile secret key (encryptedKeyValue) *with password verification*.
///
/// This decrypt path needs three derived block keys. The expensive iterated password hash is
/// computed once and reused for all key derivations.
pub fn agile_secret_key_from_password(
    info: &AgileEncryptionInfo,
    password: &str,
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

    let password_utf16le = Zeroizing::new(crate::password_to_utf16le_bytes(password));
    let h = agile_iterated_hash(
        &password_utf16le,
        &info.password_salt,
        info.password_hash_algorithm,
        info.spin_count,
    );

    // Block 1: decrypt verifierHashInput.
    let key1 = crate::derive_encryption_key(
        &h,
        &VERIFIER_HASH_INPUT_BLOCK,
        info.password_hash_algorithm,
        info.password_key_bits,
    )?;
    let verifier_hash_input =
        crate::aes_cbc_decrypt(&info.encrypted_verifier_hash_input, &key1, &info.password_salt)?;
    if verifier_hash_input.len() < VERIFIER_HASH_INPUT_LEN {
        return Err(OffcryptoError::InvalidEncryptionInfo {
            context: "decrypted verifierHashInput is truncated",
        });
    }

    // Block 2: decrypt verifierHashValue and verify.
    let key2 = crate::derive_encryption_key(
        &h,
        &VERIFIER_HASH_VALUE_BLOCK,
        info.password_hash_algorithm,
        info.password_key_bits,
    )?;
    let verifier_hash_value =
        crate::aes_cbc_decrypt(&info.encrypted_verifier_hash_value, &key2, &info.password_salt)?;

    let digest_len = hash_output_len(info.password_hash_algorithm);
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

    // Block 3: decrypt encryptedKeyValue (secret key).
    let key3 = crate::derive_encryption_key(
        &h,
        &KEY_VALUE_BLOCK,
        info.password_hash_algorithm,
        info.password_key_bits,
    )?;
    let key_value =
        crate::aes_cbc_decrypt(&info.encrypted_key_value, &key3, &info.password_salt)?;
    if key_value.len() < key_len {
        return Err(OffcryptoError::InvalidEncryptionInfo {
            context: "decrypted keyValue is truncated",
        });
    }
    Ok(Zeroizing::new(key_value[..key_len].to_vec()))
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
        reset_ct_eq_calls();

        let input = b"verifier-hash-input";
        let mut expected = HashAlgorithm::Sha1.digest(input);
        // Flip a bit to force a mismatch.
        expected[0] ^= 0x01;

        let err = verify_password(input, &expected, HashAlgorithm::Sha1)
            .expect_err("expected verifier mismatch to return an error");
        assert!(matches!(err, OffcryptoError::InvalidPassword));

        assert!(
            ct_eq_call_count() >= 1,
            "expected constant-time compare helper to be invoked"
        );
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
            encrypted_hmac_key: Vec::new(),
            encrypted_hmac_value: Vec::new(),
            spin_count,
            password_salt: salt.clone(),
            password_hash_algorithm: hash_algorithm,
            password_key_bits: key_bits,
            encrypted_key_value,
            encrypted_verifier_hash_input,
            encrypted_verifier_hash_value,
        };

        // Setup above called `agile_iterated_hash`.
        ITERATED_HASH_CALLS.store(0, Ordering::Relaxed);

        let out = agile_secret_key_from_password(&info, password).unwrap();
        assert_eq!(out.as_slice(), secret_key_plain.as_slice());
        assert_eq!(ITERATED_HASH_CALLS.load(Ordering::Relaxed), 1);
    }
}
