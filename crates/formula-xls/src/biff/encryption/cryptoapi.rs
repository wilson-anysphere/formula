use md5::Md5;
use sha1::{Digest as _, Sha1};
use zeroize::{Zeroize, Zeroizing};

use crate::ct::ct_eq;

use super::rc4::Rc4;
use super::DecryptError;

// CryptoAPI ALG_ID values for hash functions.
pub(crate) const CALG_MD5: u32 = 0x0000_8003;
pub(crate) const CALG_SHA1: u32 = 0x0000_8004;

fn cryptoapi_hash2(
    alg_id_hash: u32,
    a: &[u8],
    b: &[u8],
) -> Result<Zeroizing<Vec<u8>>, DecryptError> {
    match alg_id_hash {
        CALG_MD5 => {
            let mut h = Md5::new();
            h.update(a);
            h.update(b);
            let mut digest = h.finalize();
            let out = digest.to_vec();
            digest.as_mut_slice().zeroize();
            Ok(Zeroizing::new(out))
        }
        CALG_SHA1 => {
            let mut h = Sha1::new();
            h.update(a);
            h.update(b);
            let mut digest = h.finalize();
            let out = digest.to_vec();
            digest.as_mut_slice().zeroize();
            Ok(Zeroizing::new(out))
        }
        other => Err(DecryptError::UnsupportedEncryption(format!(
            "unsupported CryptoAPI hash alg_id_hash=0x{other:08X}"
        ))),
    }
}

/// Derive the BIFF8 RC4 CryptoAPI key for a given `block_index`.
///
/// This corresponds to the "RC4 CryptoAPI" / "CryptoAPI" encryption scheme used by newer BIFF8
/// workbooks (Excel 2002/2003), which uses SHA-1 + a spin count to harden password derivation.
///
/// `key_len` is `KeySize / 8` from the CryptoAPI header (e.g. 16 for 128-bit RC4, 5 for 40-bit).
pub(crate) fn derive_biff8_cryptoapi_key(
    alg_id_hash: u32,
    password: &str,
    salt: &[u8; 16],
    spin_count: u32,
    block_index: u32,
    key_len: usize,
) -> Result<Zeroizing<Vec<u8>>, DecryptError> {
    if key_len == 0 {
        return Err(DecryptError::InvalidFilePass(
            "CryptoAPI key length must be > 0".to_string(),
        ));
    }

    let pw_bytes: Zeroizing<Vec<u8>> = Zeroizing::new(
        password
            .encode_utf16()
            .flat_map(|c| c.to_le_bytes())
            .collect(),
    );

    // Initial hash: H0 = Hash(salt || password_utf16le)
    let mut hash = cryptoapi_hash2(alg_id_hash, salt, &pw_bytes[..])?;
    drop(pw_bytes);

    // Spin: Hi+1 = Hash(i_le || Hi)
    for i in 0..spin_count {
        let i_bytes = i.to_le_bytes();
        hash = cryptoapi_hash2(alg_id_hash, &i_bytes, &hash[..])?;
    }

    // Final block key hash: H = Hash(Hspin || block_index_le)
    let block_bytes = block_index.to_le_bytes();
    let block_hash = cryptoapi_hash2(alg_id_hash, &hash[..], &block_bytes)?;
    drop(hash);

    let key = block_hash[..key_len.min(block_hash.len())].to_vec();
    drop(block_hash);
    Ok(Zeroizing::new(key))
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
) -> Result<(Zeroizing<[u8; 16]>, Zeroizing<[u8; 20]>), DecryptError> {
    let key = derive_biff8_cryptoapi_key(CALG_SHA1, password, salt, spin_count, 0, key_len)?;
    let mut rc4 = Rc4::new(&key[..]);
    drop(key);

    let mut buf = Zeroizing::new([0u8; 36]);
    buf[..16].copy_from_slice(encrypted_verifier);
    buf[16..].copy_from_slice(encrypted_verifier_hash);
    rc4.apply_keystream(&mut buf[..]);

    let mut verifier = Zeroizing::new([0u8; 16]);
    verifier.copy_from_slice(&buf[..16]);
    let mut verifier_hash = Zeroizing::new([0u8; 20]);
    verifier_hash.copy_from_slice(&buf[16..]);
    Ok((verifier, verifier_hash))
}

/// Validate a password against a CryptoAPI verifier.
pub(crate) fn validate_biff8_cryptoapi_password(
    password: &str,
    salt: &[u8; 16],
    spin_count: u32,
    encrypted_verifier: &[u8; 16],
    encrypted_verifier_hash: &[u8; 20],
    key_len: usize,
) -> Result<bool, DecryptError> {
    let (verifier, verifier_hash) = decrypt_biff8_cryptoapi_verifier(
        password,
        salt,
        spin_count,
        encrypted_verifier,
        encrypted_verifier_hash,
        key_len,
    )?;
    let mut sha1 = Sha1::new();
    sha1.update(&verifier[..]);
    let mut expected = sha1.finalize();
    let ok = ct_eq(expected.as_slice(), &verifier_hash[..]);
    expected.as_mut_slice().zeroize();
    Ok(ok)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn cryptoapi_key_derivation_and_verifier_decrypt_matches_vector() {
        let password = "SecretPassword";
        let salt: [u8; 16] = [
            0xA0, 0xA1, 0xA2, 0xA3, 0xA4, 0xA5, 0xA6, 0xA7, 0xA8, 0xA9, 0xAA, 0xAB, 0xAC, 0xAD,
            0xAE, 0xAF,
        ];
        let spin_count: u32 = 50_000;
        let key_len: usize = 16; // 128-bit

        let expected_key: [u8; 16] = [
            0x3D, 0x7D, 0x0B, 0x04, 0x2E, 0xCF, 0x02, 0xA7, 0xBC, 0xE1, 0x26, 0xA1, 0xE2, 0x20,
            0xE2, 0xC8,
        ];
        let expected_key_block1: [u8; 16] = [
            0xF8, 0x06, 0xD7, 0x4E, 0x99, 0x94, 0x8E, 0xE8, 0xD3, 0x68, 0xD6, 0x1C, 0xEA,
            0xAA, 0x7C, 0x36,
        ];

        let encrypted_verifier: [u8; 16] = [
            0xBB, 0xFF, 0x8B, 0x22, 0x0E, 0x9A, 0x35, 0x3E, 0x6E, 0xC5, 0xE1, 0x4A, 0x40, 0x98,
            0x63, 0xA2,
        ];
        let encrypted_verifier_hash: [u8; 20] = [
            0xF5, 0xDB, 0x86, 0xB1, 0x65, 0x02, 0xB7, 0xED, 0xFE, 0x95, 0x97, 0x6F, 0x97, 0xD0,
            0x27, 0x35, 0xC2, 0x63, 0x26, 0xA0,
        ];

        let expected_verifier: [u8; 16] = [
            0xF0, 0xE1, 0xD2, 0xC3, 0xB4, 0xA5, 0x96, 0x87, 0x78, 0x69, 0x5A, 0x4B, 0x3C, 0x2D,
            0x1E, 0x0F,
        ];
        let expected_verifier_hash: [u8; 20] = [
            0x93, 0xEC, 0x7C, 0x96, 0x8F, 0x9A, 0x40, 0xFE, 0xDA, 0x5C, 0x38, 0x55, 0xF1, 0x37,
            0x82, 0x29, 0xD7, 0xE0, 0x0C, 0x53,
        ];

        let derived_key = derive_biff8_cryptoapi_key(CALG_SHA1, password, &salt, spin_count, 0, key_len)
            .expect("derive key");
        assert_eq!(&derived_key[..], &expected_key, "derived_key mismatch");
        let derived_key_block1 =
            derive_biff8_cryptoapi_key(CALG_SHA1, password, &salt, spin_count, 1, key_len)
                .expect("derive key block1");
        assert_eq!(
            &derived_key_block1[..],
            &expected_key_block1,
            "derived_key(block=1) mismatch"
        );

        let (verifier, verifier_hash) = decrypt_biff8_cryptoapi_verifier(
            password,
            &salt,
            spin_count,
            &encrypted_verifier,
            &encrypted_verifier_hash,
            key_len,
        )
        .expect("decrypt verifier");
        assert_eq!(&verifier[..], &expected_verifier, "verifier mismatch");
        assert_eq!(
            &verifier_hash[..],
            &expected_verifier_hash,
            "verifier_hash mismatch"
        );

        assert!(validate_biff8_cryptoapi_password(
            password,
            &salt,
            spin_count,
            &encrypted_verifier,
            &encrypted_verifier_hash,
            key_len
        )
        .expect("validate"));
        assert!(!validate_biff8_cryptoapi_password(
            "wrong",
            &salt,
            spin_count,
            &encrypted_verifier,
            &encrypted_verifier_hash,
            key_len
        )
        .expect("validate wrong"));
    }

    #[test]
    fn cryptoapi_key_derivation_md5_vectors() {
        let password = "password";
        let salt: [u8; 16] = [
            0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C,
            0x0D, 0x0E, 0x0F,
        ];
        let spin_count: u32 = 50_000;

        let expected: &[(u32, [u8; 16])] = &[
            (
                0,
                [
                    0x69, 0xBA, 0xDC, 0xAE, 0x24, 0x48, 0x68, 0xE2, 0x09, 0xD4, 0xE0, 0x53,
                    0xCC, 0xD2, 0xA3, 0xBC,
                ],
            ),
            (
                1,
                [
                    0x6F, 0x4D, 0x50, 0x2A, 0xB3, 0x77, 0x00, 0xFF, 0xDA, 0xB5, 0x70, 0x41,
                    0x60, 0x45, 0x5B, 0x47,
                ],
            ),
            (
                2,
                [
                    0xAC, 0x69, 0x02, 0x2E, 0x39, 0x6C, 0x77, 0x50, 0x87, 0x21, 0x33, 0xF3,
                    0x7E, 0x2C, 0x7A, 0xFC,
                ],
            ),
            (
                3,
                [
                    0x1B, 0x05, 0x6E, 0x71, 0x18, 0xAB, 0x8D, 0x35, 0xE9, 0xD6, 0x7A, 0xDE,
                    0xE8, 0xB1, 0x11, 0x04,
                ],
            ),
        ];

        for (block, expected_key) in expected {
            let key =
                derive_biff8_cryptoapi_key(CALG_MD5, password, &salt, spin_count, *block, 16)
                    .expect("derive");
            assert_eq!(key.as_slice(), expected_key, "block={block}");
        }

        // 40-bit CryptoAPI RC4 keys are 5 bytes (`keySize/8`), not padded to 16 bytes.
        let key_40 = derive_biff8_cryptoapi_key(CALG_MD5, password, &salt, spin_count, 0, 5)
            .expect("derive 40-bit");
        assert_eq!(key_40.as_slice(), vec![0x69, 0xBA, 0xDC, 0xAE, 0x24].as_slice());
        assert_eq!(key_40.len(), 5);
    }
}
