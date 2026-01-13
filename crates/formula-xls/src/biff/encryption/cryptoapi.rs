use sha1::{Digest as _, Sha1};

use super::rc4::Rc4;

/// Derive the BIFF8 RC4 CryptoAPI key for a given `block_index`.
///
/// This corresponds to the "RC4 CryptoAPI" / "CryptoAPI" encryption scheme used by newer BIFF8
/// workbooks (Excel 2002/2003), which uses SHA-1 + a spin count to harden password derivation.
///
/// `key_len` is the length of the RC4 key in bytes (commonly 16 for 128-bit RC4).
pub(crate) fn derive_biff8_cryptoapi_key(
    password: &str,
    salt: &[u8; 16],
    spin_count: u32,
    block_index: u32,
    key_len: usize,
) -> Vec<u8> {
    assert!(key_len > 0, "key_len must be > 0");

    let pw_bytes: Vec<u8> = password
        .encode_utf16()
        .flat_map(|c| c.to_le_bytes())
        .collect();

    // Initial hash: H0 = SHA1(salt || password_utf16le)
    let mut sha1 = Sha1::new();
    sha1.update(salt);
    sha1.update(&pw_bytes);
    let mut hash = sha1.finalize().to_vec();

    // Spin: Hi+1 = SHA1(i_le || Hi)
    for i in 0..spin_count {
        let mut sha1 = Sha1::new();
        sha1.update(&i.to_le_bytes());
        sha1.update(&hash);
        hash = sha1.finalize().to_vec();
    }

    // Final block key hash: H = SHA1(Hspin || block_index_le)
    let mut sha1 = Sha1::new();
    sha1.update(&hash);
    sha1.update(&block_index.to_le_bytes());
    let block_hash = sha1.finalize();

    block_hash[..key_len.min(block_hash.len())].to_vec()
}

/// Decrypt the CryptoAPI verifier and verifier hash.
///
/// Returns `(verifier, verifier_hash)` in plaintext.
pub(crate) fn decrypt_biff8_cryptoapi_verifier(
    password: &str,
    salt: &[u8; 16],
    spin_count: u32,
    encrypted_verifier: &[u8; 16],
    encrypted_verifier_hash: &[u8; 20],
    key_len: usize,
) -> ([u8; 16], [u8; 20]) {
    let key = derive_biff8_cryptoapi_key(password, salt, spin_count, 0, key_len);
    let mut rc4 = Rc4::new(&key);

    let mut buf = [0u8; 36];
    buf[..16].copy_from_slice(encrypted_verifier);
    buf[16..].copy_from_slice(encrypted_verifier_hash);
    rc4.apply_keystream(&mut buf);

    let mut verifier = [0u8; 16];
    verifier.copy_from_slice(&buf[..16]);
    let mut verifier_hash = [0u8; 20];
    verifier_hash.copy_from_slice(&buf[16..]);
    (verifier, verifier_hash)
}

/// Validate a password against a CryptoAPI verifier.
pub(crate) fn validate_biff8_cryptoapi_password(
    password: &str,
    salt: &[u8; 16],
    spin_count: u32,
    encrypted_verifier: &[u8; 16],
    encrypted_verifier_hash: &[u8; 20],
    key_len: usize,
) -> bool {
    let (verifier, verifier_hash) = decrypt_biff8_cryptoapi_verifier(
        password,
        salt,
        spin_count,
        encrypted_verifier,
        encrypted_verifier_hash,
        key_len,
    );
    let mut sha1 = Sha1::new();
    sha1.update(&verifier);
    let expected = sha1.finalize();
    expected.as_slice() == verifier_hash
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn cryptoapi_key_derivation_and_verifier_decrypt_matches_vector() {
        let password = "SecretPassword";
        let salt: [u8; 16] = [
            0xA0, 0xA1, 0xA2, 0xA3, 0xA4, 0xA5, 0xA6, 0xA7, 0xA8, 0xA9, 0xAA, 0xAB, 0xAC,
            0xAD, 0xAE, 0xAF,
        ];
        let spin_count: u32 = 50_000;
        let key_len: usize = 16; // 128-bit

        let expected_key: [u8; 16] = [
            0x3D, 0x7D, 0x0B, 0x04, 0x2E, 0xCF, 0x02, 0xA7, 0xBC, 0xE1, 0x26, 0xA1, 0xE2,
            0x20, 0xE2, 0xC8,
        ];

        let encrypted_verifier: [u8; 16] = [
            0xBB, 0xFF, 0x8B, 0x22, 0x0E, 0x9A, 0x35, 0x3E, 0x6E, 0xC5, 0xE1, 0x4A, 0x40,
            0x98, 0x63, 0xA2,
        ];
        let encrypted_verifier_hash: [u8; 20] = [
            0xF5, 0xDB, 0x86, 0xB1, 0x65, 0x02, 0xB7, 0xED, 0xFE, 0x95, 0x97, 0x6F, 0x97,
            0xD0, 0x27, 0x35, 0xC2, 0x63, 0x26, 0xA0,
        ];

        let expected_verifier: [u8; 16] = [
            0xF0, 0xE1, 0xD2, 0xC3, 0xB4, 0xA5, 0x96, 0x87, 0x78, 0x69, 0x5A, 0x4B, 0x3C,
            0x2D, 0x1E, 0x0F,
        ];
        let expected_verifier_hash: [u8; 20] = [
            0x93, 0xEC, 0x7C, 0x96, 0x8F, 0x9A, 0x40, 0xFE, 0xDA, 0x5C, 0x38, 0x55, 0xF1,
            0x37, 0x82, 0x29, 0xD7, 0xE0, 0x0C, 0x53,
        ];

        let derived_key = derive_biff8_cryptoapi_key(password, &salt, spin_count, 0, key_len);
        assert_eq!(derived_key, expected_key, "derived_key mismatch");

        let (verifier, verifier_hash) = decrypt_biff8_cryptoapi_verifier(
            password,
            &salt,
            spin_count,
            &encrypted_verifier,
            &encrypted_verifier_hash,
            key_len,
        );
        assert_eq!(verifier, expected_verifier, "verifier mismatch");
        assert_eq!(verifier_hash, expected_verifier_hash, "verifier_hash mismatch");

        assert!(validate_biff8_cryptoapi_password(
            password,
            &salt,
            spin_count,
            &encrypted_verifier,
            &encrypted_verifier_hash,
            key_len
        ));
        assert!(!validate_biff8_cryptoapi_password(
            "wrong",
            &salt,
            spin_count,
            &encrypted_verifier,
            &encrypted_verifier_hash,
            key_len
        ));
    }
}
