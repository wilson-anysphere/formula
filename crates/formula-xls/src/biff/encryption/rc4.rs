use md5::{Digest as _, Md5};
use zeroize::{Zeroize, Zeroizing};

use crate::ct::ct_eq;

/// Minimal RC4 stream cipher implementation (KSA + PRGA).
///
/// BIFF8 legacy encryption uses RC4 with per-block keys derived from password material.
#[derive(Clone)]
pub(crate) struct Rc4 {
    s: [u8; 256],
    i: u8,
    j: u8,
}

impl Rc4 {
    pub(crate) fn new(key: &[u8]) -> Self {
        // A well-formed BIFF `FILEPASS` record never results in an empty RC4 key, but this module
        // is used on untrusted input. Avoid panicking in case of malformed headers.
        let key = if key.is_empty() { &[0u8][..] } else { key };

        let mut s = [0u8; 256];
        for (i, v) in s.iter_mut().enumerate() {
            *v = i as u8;
        }

        let mut j: u8 = 0;
        for i in 0..256u16 {
            let idx = i as usize;
            j = j.wrapping_add(s[idx]).wrapping_add(key[idx % key.len()]);
            s.swap(idx, j as usize);
        }

        Self { s, i: 0, j: 0 }
    }

    pub(crate) fn apply_keystream(&mut self, data: &mut [u8]) {
        for b in data {
            self.i = self.i.wrapping_add(1);
            self.j = self.j.wrapping_add(self.s[self.i as usize]);
            self.s.swap(self.i as usize, self.j as usize);
            let idx = self.s[self.i as usize].wrapping_add(self.s[self.j as usize]);
            let k = self.s[idx as usize];
            *b ^= k;
        }
    }
}

impl Zeroize for Rc4 {
    fn zeroize(&mut self) {
        self.s.zeroize();
        self.i = 0;
        self.j = 0;
    }
}

impl Drop for Rc4 {
    fn drop(&mut self) {
        self.zeroize();
    }
}

/// Derive the BIFF8 *RC4* (non-CryptoAPI) per-block key.
///
/// This corresponds to the legacy Office 97-2000 RC4 scheme (MD5-based).
///
/// `key_len` is the RC4 key length in bytes. Excel commonly uses 5 bytes (40-bit) for this scheme.
pub(crate) fn derive_biff8_rc4_key(
    password: &str,
    salt: &[u8; 16],
    block_index: u32,
    key_len: usize,
) -> Zeroizing<Vec<u8>> {
    if key_len == 0 {
        return Zeroizing::new(Vec::new());
    }

    let intermediate_key = derive_biff8_rc4_intermediate_key(password, salt);
    let block_key = derive_biff8_rc4_block_key(&*intermediate_key, block_index);
    drop(intermediate_key);

    let key = block_key[..key_len.min(block_key.len())].to_vec();
    drop(block_key);
    Zeroizing::new(key)
}

/// Derive the BIFF8 RC4 "intermediate key" (MD5-based) from `password` and `salt`.
///
/// This corresponds to `H1 = MD5(MD5(password_utf16le) || salt)` in [MS-OFFCRYPTO].
pub(crate) fn derive_biff8_rc4_intermediate_key(
    password: &str,
    salt: &[u8; 16],
) -> Zeroizing<[u8; 16]> {
    // Excel 97-2003 Standard Encryption only considers the first 15 UTF-16 code units of the
    // password.
    let pw_bytes = super::password_to_utf16le(password);

    // H0 = MD5(password_utf16le)
    let mut md5 = Md5::new();
    md5.update(&*pw_bytes);
    let mut h0 = md5.finalize();
    drop(pw_bytes);

    // H1 = MD5(H0 || salt)
    let mut md5 = Md5::new();
    md5.update(&h0);
    md5.update(salt);
    let mut digest = md5.finalize();
    h0.as_mut_slice().zeroize();

    let mut out = [0u8; 16];
    out.copy_from_slice(&digest);
    digest.as_mut_slice().zeroize();
    Zeroizing::new(out)
}

/// Derive the BIFF8 RC4 per-block key (16 bytes).
///
/// This corresponds to `H2 = MD5(intermediate_key || block_index_le32)`.
pub(crate) fn derive_biff8_rc4_block_key(
    intermediate_key: &[u8; 16],
    block_index: u32,
) -> Zeroizing<[u8; 16]> {
    let mut md5 = Md5::new();
    md5.update(intermediate_key);
    md5.update(&block_index.to_le_bytes());
    let mut digest = md5.finalize();

    let mut out = [0u8; 16];
    out.copy_from_slice(&digest);
    digest.as_mut_slice().zeroize();
    Zeroizing::new(out)
}

/// Decrypt the legacy RC4 verifier and verifier hash.
///
/// Returns `(verifier, verifier_hash)` in plaintext.
pub(crate) fn decrypt_biff8_rc4_verifier(
    password: &str,
    salt: &[u8; 16],
    encrypted_verifier: &[u8; 16],
    encrypted_verifier_hash: &[u8; 16],
    key_len: usize,
) -> (Zeroizing<[u8; 16]>, Zeroizing<[u8; 16]>) {
    let key = derive_biff8_rc4_key(password, salt, 0, key_len);
    let mut rc4 = Rc4::new(&key[..]);
    drop(key);

    let mut buf = Zeroizing::new([0u8; 32]);
    buf[..16].copy_from_slice(encrypted_verifier);
    buf[16..].copy_from_slice(encrypted_verifier_hash);
    rc4.apply_keystream(&mut buf[..]);

    let mut verifier = Zeroizing::new([0u8; 16]);
    verifier.copy_from_slice(&buf[..16]);
    let mut verifier_hash = Zeroizing::new([0u8; 16]);
    verifier_hash.copy_from_slice(&buf[16..]);
    (verifier, verifier_hash)
}

/// Validate a password against the legacy RC4 verifier.
pub(crate) fn validate_biff8_rc4_password(
    password: &str,
    salt: &[u8; 16],
    encrypted_verifier: &[u8; 16],
    encrypted_verifier_hash: &[u8; 16],
    key_len: usize,
) -> bool {
    let (verifier, verifier_hash) = decrypt_biff8_rc4_verifier(
        password,
        salt,
        encrypted_verifier,
        encrypted_verifier_hash,
        key_len,
    );

    let mut md5 = Md5::new();
    md5.update(&verifier[..]);
    let mut expected = md5.finalize();
    let ok = ct_eq(expected.as_slice(), &verifier_hash[..]);
    expected.as_mut_slice().zeroize();
    ok
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn rc4_cipher_implements_zeroize() {
        fn assert_zeroize<T: Zeroize>() {}
        assert_zeroize::<Rc4>();
    }

    // Deterministic vector generated from Excel's documented algorithms (MD5 + RC4), using:
    // - password = "SecretPassword"
    // - salt = 16 fixed bytes
    // - verifier = 16 fixed bytes
    //
    // The encryptedVerifier/encryptedVerifierHash values are RC4-encrypted using the derived
    // block-0 key, and are intentionally embedded so that key derivation changes will be caught
    // by unit tests.
    #[test]
    fn rc4_key_derivation_and_verifier_decrypt_matches_vector() {
        let password = "SecretPassword";
        let salt: [u8; 16] = [
            0x01, 0x23, 0x45, 0x67, 0x89, 0xAB, 0xCD, 0xEF, 0x10, 0x32, 0x54, 0x76, 0x98, 0xBA,
            0xDC, 0xFE,
        ];

        // 40-bit key (5 bytes) is typical for BIFF8 RC4.
        let key_len = 5usize;

        let expected_key: [u8; 5] = [0x50, 0x0A, 0xE2, 0x80, 0xEE];
        let expected_key_block1: [u8; 5] = [0x8E, 0xE6, 0xCF, 0x4E, 0x3C];

        let encrypted_verifier: [u8; 16] = [
            0x3A, 0xAE, 0x87, 0x68, 0x86, 0xB8, 0x19, 0xF4, 0x34, 0x28, 0x11, 0x4A, 0x4F, 0x62,
            0xA8, 0x70,
        ];
        let encrypted_verifier_hash: [u8; 16] = [
            0xBD, 0x24, 0x21, 0xEE, 0xEB, 0x88, 0x08, 0x35, 0x01, 0xEC, 0x4F, 0xA5, 0x26, 0xF3,
            0xFD, 0x9A,
        ];

        let expected_verifier: [u8; 16] = [
            0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D,
            0x0E, 0x0F,
        ];
        let expected_verifier_hash: [u8; 16] = [
            0x1A, 0xC1, 0xEF, 0x01, 0xE9, 0x6C, 0xAF, 0x1B, 0xE0, 0xD3, 0x29, 0x33, 0x1A, 0x4F,
            0xC2, 0xA8,
        ];

        let derived_key = derive_biff8_rc4_key(password, &salt, 0, key_len);
        assert_eq!(&derived_key[..], &expected_key, "derived_key mismatch");
        let derived_key_block1 = derive_biff8_rc4_key(password, &salt, 1, key_len);
        assert_eq!(
            &derived_key_block1[..],
            &expected_key_block1,
            "derived_key(block=1) mismatch"
        );

        let (verifier, verifier_hash) = decrypt_biff8_rc4_verifier(
            password,
            &salt,
            &encrypted_verifier,
            &encrypted_verifier_hash,
            key_len,
        );
        assert_eq!(&verifier[..], &expected_verifier, "verifier mismatch");
        assert_eq!(
            &verifier_hash[..],
            &expected_verifier_hash,
            "verifier_hash mismatch"
        );

        assert!(validate_biff8_rc4_password(
            password,
            &salt,
            &encrypted_verifier,
            &encrypted_verifier_hash,
            key_len
        ));
        assert!(!validate_biff8_rc4_password(
            "wrong",
            &salt,
            &encrypted_verifier,
            &encrypted_verifier_hash,
            key_len
        ));
    }
}
