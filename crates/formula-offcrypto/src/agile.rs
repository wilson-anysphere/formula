//! Agile encryption password verification helpers.
//!
//! In the Agile encryption scheme (OOXML), password verification is performed
//! by decrypting the `encryptedVerifierHashInput` and `encryptedVerifierHashValue`
//! using a key derived from the provided password, then checking that:
//!
//! `Hash(verifierHashInput) == verifierHashValue`.
//!
//! This module provides the digest verification step, using constant-time
//! comparisons.

use crate::util::ct_eq;
use crate::{HashAlgorithm, OffcryptoError};

/// Verify the decrypted Agile verifier fields for a candidate password.
///
/// Callers are expected to pass the *decrypted* `verifierHashInput` and
/// `verifierHashValue` fields (trimmed to the hash output size).
pub fn verify_password(
    verifier_hash_input: &[u8],
    verifier_hash_value: &[u8],
    hash_alg: HashAlgorithm,
) -> Result<(), OffcryptoError> {
    let digest = hash_alg.digest(verifier_hash_input);
    if !ct_eq(&digest, verifier_hash_value) {
        return Err(OffcryptoError::InvalidPassword);
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::util::{ct_eq_call_count, reset_ct_eq_calls};

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
}
