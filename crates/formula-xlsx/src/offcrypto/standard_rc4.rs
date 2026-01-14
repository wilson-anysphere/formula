//! MS-OFFCRYPTO "Standard Encryption" / CryptoAPI RC4 password verifier.
//!
//! This module implements the **password verifier** check for the legacy
//! MS-OFFCRYPTO *Standard* encryption scheme (CryptoAPI RC4).
//!
//! The verifier is the first-line check for wrong passwords. A common bug is
//! to reinitialize RC4 when decrypting `EncryptedVerifierHash`; per the spec,
//! `EncryptedVerifier` and `EncryptedVerifierHash` are encrypted with **one
//! continuous RC4 keystream** (no reset between the two fields).

use digest::Digest as _;

#[cfg(test)]
use std::cell::Cell;

// Unit tests run in parallel by default. Use a thread-local counter so tests that reset/inspect
// the counter don't race each other.
#[cfg(test)]
thread_local! {
    static CT_EQ_CALLS: Cell<usize> = Cell::new(0);
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum CryptoApiHashAlg {
    Sha1,
    Md5,
}

impl CryptoApiHashAlg {
    fn hash_parts(self, parts: &[&[u8]]) -> Vec<u8> {
        match self {
            CryptoApiHashAlg::Sha1 => {
                let mut h = sha1::Sha1::new();
                for part in parts {
                    h.update(part);
                }
                h.finalize().to_vec()
            }
            CryptoApiHashAlg::Md5 => {
                let mut h = md5::Md5::new();
                for part in parts {
                    h.update(part);
                }
                h.finalize().to_vec()
            }
        }
    }

    fn hash(self, data: &[u8]) -> Vec<u8> {
        self.hash_parts(&[data])
    }
}

/// MS-OFFCRYPTO Standard (CryptoAPI) uses a fixed 50,000 iteration count for password hashing.
const STANDARD_SPIN_COUNT: u32 = 50_000;

fn password_utf16le_bytes(password: &str) -> Vec<u8> {
    // UTF-16LE with no BOM and no terminator.
    let mut out = Vec::with_capacity(password.len().saturating_mul(2));
    for unit in password.encode_utf16() {
        out.extend_from_slice(&unit.to_le_bytes());
    }
    out
}

fn ct_eq(a: &[u8], b: &[u8]) -> bool {
    #[cfg(test)]
    CT_EQ_CALLS.with(|calls| calls.set(calls.get().saturating_add(1)));

    let mut diff = 0u8;
    let max_len = a.len().max(b.len());
    for idx in 0..max_len {
        let av = a.get(idx).copied().unwrap_or(0);
        let bv = b.get(idx).copied().unwrap_or(0);
        diff |= av ^ bv;
    }
    diff == 0 && a.len() == b.len()
}

#[cfg(test)]
fn reset_ct_eq_calls() {
    CT_EQ_CALLS.with(|calls| calls.set(0));
}

#[cfg(test)]
fn ct_eq_call_count() -> usize {
    CT_EQ_CALLS.with(|calls| calls.get())
}

/// RC4 stream cipher.
#[derive(Clone)]
struct Rc4 {
    s: [u8; 256],
    i: u8,
    j: u8,
}

impl Rc4 {
    fn new(key: &[u8]) -> Self {
        assert!(!key.is_empty(), "RC4 key must be non-empty");
        assert!(
            key.len() <= 256,
            "RC4 key length must be <= 256 bytes (got {})",
            key.len()
        );

        let mut s = [0u8; 256];
        for (idx, b) in s.iter_mut().enumerate() {
            *b = idx as u8;
        }

        // Key scheduling algorithm (KSA).
        let mut j: u8 = 0;
        for i in 0u16..256 {
            j = j
                .wrapping_add(s[i as usize])
                .wrapping_add(key[i as usize % key.len()]);
            s.swap(i as usize, j as usize);
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

/// Derive the RC4 key for a given `EncryptedPackage` block number (or 0 for the verifier).
///
/// MS-OFFCRYPTO Standard/CryptoAPI RC4 key derivation:
/// 1. `H = Hash(salt || UTF16LE(password))`
/// 2. For `i` in `0..50000`:
///    `H = Hash(LE32(i) || H)`
/// 3. `Hfinal = Hash(H || LE32(block_index))`
/// 4. `rc4_key = Hfinal[0..key_size_bytes]`
fn derive_rc4_key_for_block(
    password: &str,
    salt: &[u8],
    hash_alg: CryptoApiHashAlg,
    key_size_bits: u32,
    block_index: u32,
) -> Vec<u8> {
    // MS-OFFCRYPTO: for Standard/CryptoAPI RC4, `keySize == 0` MUST be interpreted as 40-bit.
    let key_size_bits = if key_size_bits == 0 { 40 } else { key_size_bits };
    assert!(key_size_bits.is_multiple_of(8), "key size must be byte-aligned");
    let key_size_bytes = (key_size_bits / 8) as usize;
    assert!(key_size_bytes > 0, "key size must be non-zero");

    let pw = password_utf16le_bytes(password);
    let mut h = hash_alg.hash_parts(&[salt, &pw]);
    for i in 0u32..STANDARD_SPIN_COUNT {
        let i_le = i.to_le_bytes();
        h = hash_alg.hash_parts(&[&i_le, &h]);
    }

    let block = block_index.to_le_bytes();
    let h_final = hash_alg.hash_parts(&[&h, &block]);
    assert!(
        key_size_bytes <= h_final.len(),
        "key size must be <= digest length (key={} bytes digest={} bytes)",
        key_size_bytes,
        h_final.len()
    );

    h_final[..key_size_bytes].to_vec()
}

/// Standard / CryptoAPI RC4 verifier bundle (MS-OFFCRYPTO `EncryptionVerifier`).
#[derive(Debug, Clone)]
pub(crate) struct StandardRc4CryptoApiVerifier {
    pub(crate) salt: [u8; 16],
    pub(crate) encrypted_verifier: [u8; 16],
    pub(crate) verifier_hash_size: u32,
    pub(crate) encrypted_verifier_hash: Vec<u8>,
    pub(crate) key_size_bits: u32,
    pub(crate) hash_alg: CryptoApiHashAlg,
}

impl StandardRc4CryptoApiVerifier {
    /// Verify `password` using the MS-OFFCRYPTO Standard/CryptoAPI RC4 verifier check.
    pub(crate) fn verify_password(&self, password: &str) -> bool {
        // The verifier uses the same key derivation as `EncryptedPackage` block 0.
        let rc4_key_block0 =
            derive_rc4_key_for_block(password, &self.salt, self.hash_alg, self.key_size_bits, 0);

        // Critical: one RC4 state across verifier + hash (single keystream).
        let mut rc4 = Rc4::new(&rc4_key_block0);

        let mut verifier = self.encrypted_verifier;
        rc4.apply_keystream(&mut verifier);

        let mut verifier_hash = self.encrypted_verifier_hash.clone();
        rc4.apply_keystream(&mut verifier_hash);

        let computed_hash = self.hash_alg.hash(&verifier);
        let expected_len = self.verifier_hash_size as usize;
        if expected_len == 0 {
            return false;
        }
        if expected_len > computed_hash.len() || expected_len > verifier_hash.len() {
            return false;
        }
        ct_eq(&computed_hash[..expected_len], &verifier_hash[..expected_len])
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn build_fixture(
        password: &str,
        salt: [u8; 16],
        verifier_plaintext: [u8; 16],
        hash_alg: CryptoApiHashAlg,
        key_size_bits: u32,
    ) -> StandardRc4CryptoApiVerifier {
        let verifier_hash_plaintext = hash_alg.hash(&verifier_plaintext);
        let verifier_hash_size = verifier_hash_plaintext.len() as u32;

        let rc4_key_block0 = derive_rc4_key_for_block(password, &salt, hash_alg, key_size_bits, 0);

        // Encrypt verifier + verifierHash using *one* RC4 stream (no reset).
        let mut rc4 = Rc4::new(&rc4_key_block0);
        let mut encrypted_verifier = verifier_plaintext;
        rc4.apply_keystream(&mut encrypted_verifier);
        let mut encrypted_verifier_hash = verifier_hash_plaintext.clone();
        rc4.apply_keystream(&mut encrypted_verifier_hash);

        StandardRc4CryptoApiVerifier {
            salt,
            encrypted_verifier,
            verifier_hash_size,
            encrypted_verifier_hash,
            key_size_bits,
            hash_alg,
        }
    }

    /// Simulate the classic bug: reinitialize RC4 before decrypting `EncryptedVerifierHash`.
    fn verify_password_with_keystream_reset_bug(
        verifier: &StandardRc4CryptoApiVerifier,
        password: &str,
    ) -> bool {
        let key = derive_rc4_key_for_block(
            password,
            &verifier.salt,
            verifier.hash_alg,
            verifier.key_size_bits,
            0,
        );

        let mut rc4_a = Rc4::new(&key);
        let mut decrypted_verifier = verifier.encrypted_verifier;
        rc4_a.apply_keystream(&mut decrypted_verifier);

        // BUG: RC4 is reset for the hash instead of continuing the stream.
        let mut rc4_b = Rc4::new(&key);
        let mut decrypted_hash = verifier.encrypted_verifier_hash.clone();
        rc4_b.apply_keystream(&mut decrypted_hash);

        let computed = verifier.hash_alg.hash(&decrypted_verifier);
        let n = verifier.verifier_hash_size as usize;
        if n == 0 || n > computed.len() || n > decrypted_hash.len() {
            return false;
        }
        ct_eq(&computed[..n], &decrypted_hash[..n])
    }

    #[test]
    fn standard_cryptoapi_rc4_verifier_uses_constant_time_compare() {
        let password = "correct horse battery staple";
        let salt = [
            0xA9, 0x38, 0x13, 0x7C, 0x20, 0x54, 0xD2, 0xB6, 0x9D, 0x01, 0xFA, 0xEE, 0x8B, 0x3C, 0x47, 0x11,
        ];
        let verifier_plaintext = [
            0x10, 0x11, 0x12, 0x13, 0x14, 0x15, 0x16, 0x17, 0x18, 0x19, 0x1A, 0x1B, 0x1C, 0x1D, 0x1E, 0x1F,
        ];

        let verifier = build_fixture(
            password,
            salt,
            verifier_plaintext,
            CryptoApiHashAlg::Sha1,
            128,
        );

        reset_ct_eq_calls();
        assert!(!verifier.verify_password("wrong password"));
        assert!(
            ct_eq_call_count() >= 1,
            "expected ct_eq to be used for verifier digest comparison"
        );
    }

    #[test]
    fn standard_cryptoapi_rc4_verifier_uses_single_keystream_sha1() {
        let password = "correct horse battery staple";
        let salt = [
            0xA9, 0x38, 0x13, 0x7C, 0x20, 0x54, 0xD2, 0xB6, 0x9D, 0x01, 0xFA, 0xEE, 0x8B, 0x3C,
            0x47, 0x11,
        ];
        let verifier_plaintext = [
            0x10, 0x11, 0x12, 0x13, 0x14, 0x15, 0x16, 0x17, 0x18, 0x19, 0x1A, 0x1B, 0x1C, 0x1D,
            0x1E, 0x1F,
        ];

        let verifier = build_fixture(
            password,
            salt,
            verifier_plaintext,
            CryptoApiHashAlg::Sha1,
            128,
        );

        assert!(verifier.verify_password(password));
        assert!(!verifier.verify_password("wrong password"));

        // Ensure the test detects the keystream-reset bug.
        assert!(!verify_password_with_keystream_reset_bug(
            &verifier, password
        ));
    }

    #[test]
    fn standard_cryptoapi_rc4_verifier_uses_single_keystream_md5() {
        let password = "P@ssw0rd!";
        let salt = [
            0x1B, 0x2A, 0x3C, 0x4D, 0x5E, 0x6F, 0x70, 0x81, 0x92, 0xA3, 0xB4, 0xC5, 0xD6, 0xE7,
            0xF8, 0x09,
        ];
        let verifier_plaintext = [
            0xFE, 0xED, 0xFA, 0xCE, 0xBA, 0xAD, 0xF0, 0x0D, 0x12, 0x34, 0x56, 0x78, 0x9A, 0xBC,
            0xDE, 0xF0,
        ];

        // Use 40-bit key size to exercise 5-byte RC4 keys.
        let verifier = build_fixture(
            password,
            salt,
            verifier_plaintext,
            CryptoApiHashAlg::Md5,
            40,
        );

        assert!(verifier.verify_password(password));
        assert!(!verifier.verify_password("definitely not the password"));
        assert!(!verify_password_with_keystream_reset_bug(
            &verifier, password
        ));
    }

    #[test]
    fn standard_cryptoapi_rc4_keysize_zero_is_interpreted_as_40_bit() {
        // MS-OFFCRYPTO: for RC4, `keySize == 0` MUST be interpreted as 40-bit RC4.
        //
        // Reuse the deterministic vectors from `docs/offcrypto-standard-cryptoapi-rc4.md`.
        let password = "password";
        let salt: [u8; 16] = [
            0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D,
            0x0E, 0x0F,
        ];

        let expected_key: [u8; 5] = [0x6A, 0xD7, 0xDE, 0xDF, 0x2D];

        let key_size_bits = 0;
        let key0 = derive_rc4_key_for_block(password, &salt, CryptoApiHashAlg::Sha1, key_size_bits, 0);
        assert_eq!(key0.as_slice(), expected_key.as_slice());

        // Ensure `keySize=0` matches the `keySize=40` behavior.
        let key0_40 = derive_rc4_key_for_block(password, &salt, CryptoApiHashAlg::Sha1, 40, 0);
        assert_eq!(key0, key0_40);
    }

    #[test]
    fn standard_cryptoapi_rc4_verifier_hash_size_truncates_computed_hash() {
        // Some producers set VerifierHashSize smaller than the digest output size.
        // We should compare only the first VerifierHashSize bytes.
        let password = "truncate-test";
        let salt = [0x55u8; 16];
        let verifier_plaintext = [0xA5u8; 16];
        let hash_alg = CryptoApiHashAlg::Sha1;
        let full_hash = hash_alg.hash(&verifier_plaintext);
        assert_eq!(full_hash.len(), 20);

        let verifier_hash_size = 16u32;
        let key_size_bits = 128;
        let rc4_key_block0 = derive_rc4_key_for_block(password, &salt, hash_alg, key_size_bits, 0);

        let mut rc4 = Rc4::new(&rc4_key_block0);
        let mut encrypted_verifier = verifier_plaintext;
        rc4.apply_keystream(&mut encrypted_verifier);

        // Only encrypt the truncated hash bytes (single RC4 stream).
        let mut truncated_hash = full_hash[..verifier_hash_size as usize].to_vec();
        rc4.apply_keystream(&mut truncated_hash);

        let verifier = StandardRc4CryptoApiVerifier {
            salt,
            encrypted_verifier,
            verifier_hash_size,
            encrypted_verifier_hash: truncated_hash,
            key_size_bits,
            hash_alg,
        };

        assert!(verifier.verify_password(password));
    }

    #[test]
    fn standard_cryptoapi_rc4_keysize_zero_is_interpreted_as_40bit() {
        // MS-OFFCRYPTO specifies `keySize == 0` MUST be interpreted as 40-bit for Standard/CryptoAPI
        // RC4.
        let password = "password";
        let salt = [0x42u8; 16];

        let key0 = derive_rc4_key_for_block(password, &salt, CryptoApiHashAlg::Sha1, 0, 0);
        let key40 = derive_rc4_key_for_block(password, &salt, CryptoApiHashAlg::Sha1, 40, 0);

        assert_eq!(key0, key40);
        // Office represents the 40-bit RC4 key as a 16-byte key padded with zeros.
        assert_eq!(key0.len(), 16);
        assert!(key0[5..].iter().all(|b| *b == 0));
    }
}
