//! Standard encryption password verification helpers.
//!
//! In the Standard encryption scheme (legacy Office formats), the verifier
//! check is:
//!
//! `Hash(verifier) == verifierHash`.
//!
//! This module provides the digest verification step, using constant-time
//! comparisons.

use crate::util::ct_eq;
use crate::{HashAlgorithm, OffcryptoError};
use zeroize::Zeroizing;

/// Verify the decrypted Standard verifier fields for a candidate password.
///
/// Callers are expected to pass the *decrypted* `verifier` and `verifierHash`
/// fields (trimmed to the hash output size).
pub fn verify_verifier(
    verifier: &[u8],
    verifier_hash: &[u8],
    hash_alg: HashAlgorithm,
) -> Result<(), OffcryptoError> {
    let digest = Zeroizing::new(hash_alg.digest(verifier));
    if !ct_eq(&digest[..], verifier_hash) {
        return Err(OffcryptoError::InvalidPassword);
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::util::{ct_eq_call_count, reset_ct_eq_calls};

    #[test]
    fn standard_verify_verifier_mismatch_returns_invalid_password_and_uses_ct_eq() {
        let verifier = b"standard-verifier";
        for alg in [HashAlgorithm::Sha1, HashAlgorithm::Md5] {
            reset_ct_eq_calls();

            let mut expected = alg.digest(verifier);
            expected[0] ^= 0x80;

            let err = verify_verifier(verifier, &expected, alg)
                .expect_err("expected verifier mismatch to return an error");
            assert!(matches!(err, OffcryptoError::InvalidPassword));

            assert!(
                ct_eq_call_count() >= 1,
                "expected constant-time compare helper to be invoked (alg={alg})"
            );
        }
    }
}
