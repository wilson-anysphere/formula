//! MS-OFFCRYPTO RC4 CryptoAPI ("Standard") key derivation helpers.
//!
//! This module implements the per-block RC4 key derivation used by the "Standard" (CryptoAPI)
//! encryption scheme. It is intentionally small and self-contained so unit tests can lock in the
//! subtle byte ordering and iteration behavior.
#![allow(dead_code)]

use md5::Md5;
use sha1::Sha1;
use sha2::{Digest as _, Sha256};

/// Spin count used by MS-OFFCRYPTO "Standard" (CryptoAPI) encryption.
///
/// This value is a fixed constant in the "Standard" encryption scheme (unlike Agile encryption,
/// where it is an explicit parameter).
pub(crate) const STANDARD_SPIN_COUNT: u32 = 50_000;

/// Hash algorithm used by MS-OFFCRYPTO key derivation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum HashAlg {
    Md5,
    Sha1,
    Sha256,
}

fn hash(alg: HashAlg, bytes: &[u8]) -> Vec<u8> {
    match alg {
        HashAlg::Md5 => Md5::digest(bytes).to_vec(),
        HashAlg::Sha1 => Sha1::digest(bytes).to_vec(),
        HashAlg::Sha256 => Sha256::digest(bytes).to_vec(),
    }
}

fn password_utf16le_bytes(password: &str) -> Vec<u8> {
    // MS-OFFCRYPTO uses the UTF-16LE code units of the password string (without a NUL terminator).
    let mut out = Vec::with_capacity(password.len() * 2);
    out.extend(password.encode_utf16().flat_map(|u| u.to_le_bytes()));
    out
}

/// Derive `rc4_key_b` for MS-OFFCRYPTO RC4 CryptoAPI ("Standard") encryption.
///
/// For SHA-1, this corresponds to the spec pseudocode:
///
/// ```text
/// H = Hash(salt || password_utf16le)
/// for i in 0..50000:
///   H = Hash(LE32(i) || H)
/// h_block = Hash(H || LE32(block_index))
/// key_b = h_block[0 .. keySize/8]
/// ```
pub(crate) fn derive_rc4_key_b(
    password: &str,
    salt: &[u8],
    key_size_bits: usize,
    block_index: u32,
    hash_alg: HashAlg,
) -> Vec<u8> {
    // MS-OFFCRYPTO specifies that for Standard/CryptoAPI RC4, `EncryptionHeader.keySize == 0` MUST
    // be interpreted as 40-bit.
    let key_size_bits = if key_size_bits == 0 { 40 } else { key_size_bits };
    debug_assert!(
        key_size_bits % 8 == 0,
        "key_size_bits must be a whole number of bytes"
    );
    let key_len = key_size_bits / 8;

    let mut h_in = Vec::with_capacity(salt.len() + password.len() * 2);
    h_in.extend_from_slice(salt);
    h_in.extend_from_slice(&password_utf16le_bytes(password));

    let mut h = hash(hash_alg, &h_in);

    // Standard/CryptoAPI uses a fixed 50,000-iteration spin loop, hashing the 32-bit little-endian
    // counter *prepended* to the previous hash output each time.
    for i in 0..STANDARD_SPIN_COUNT {
        let mut buf = Vec::with_capacity(4 + h.len());
        buf.extend_from_slice(&i.to_le_bytes());
        buf.extend_from_slice(&h);
        h = hash(hash_alg, &buf);
    }

    // Derive the per-block key by appending the 32-bit little-endian block index to H.
    let mut buf = Vec::with_capacity(h.len() + 4);
    buf.extend_from_slice(&h);
    buf.extend_from_slice(&block_index.to_le_bytes());
    let block_hash = hash(hash_alg, &buf);
    block_hash.into_iter().take(key_len).collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    fn decode_hex(s: &str) -> Vec<u8> {
        let s = s.trim();
        assert!(
            s.len() % 2 == 0,
            "hex string must have an even number of characters"
        );
        let mut out = Vec::with_capacity(s.len() / 2);
        for pair in s.as_bytes().chunks_exact(2) {
            let hi = (pair[0] as char)
                .to_digit(16)
                .unwrap_or_else(|| panic!("invalid hex digit {:?}", pair[0] as char));
            let lo = (pair[1] as char)
                .to_digit(16)
                .unwrap_or_else(|| panic!("invalid hex digit {:?}", pair[1] as char));
            out.push(((hi << 4) | lo) as u8);
        }
        out
    }

    #[test]
    fn rc4_cryptoapi_standard_derive_rc4_key_b_sha1_keysize_128_vectors() {
        // Deterministic test vectors to catch subtle mistakes:
        // - password encoding must be UTF-16LE (not UTF-8)
        // - input concatenation order must be salt || password
        // - spin loop must be 50,000 iterations of Hash(LE32(i) || H)
        // - per-block key must be Hash(H || LE32(block_index))
        let password = "password";
        let salt: [u8; 16] = [
            0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D,
            0x0E, 0x0F,
        ];

        let expected = [
            (0u32, "6ad7dedf2da3514b1d85eabee069d47d"),
            (1u32, "2ed4e8825cd48aa4a47994cda7415b4a"),
            (2u32, "9ce57d0699be3938951f47fa949361db"),
            (3u32, "e65b2643eaba3815a37a61159f137840"),
        ];

        for (block, expected_hex) in expected {
            let derived = derive_rc4_key_b(password, &salt, 128, block, HashAlg::Sha1);
            assert_eq!(derived, decode_hex(expected_hex), "block index {block}");
        }
    }

    #[test]
    fn rc4_cryptoapi_standard_derive_rc4_key_b_sha1_keysize_40_truncates_to_5_bytes() {
        let password = "password";
        let salt: [u8; 16] = [
            0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D,
            0x0E, 0x0F,
        ];

        let derived = derive_rc4_key_b(password, &salt, 40, 0, HashAlg::Sha1);
        assert_eq!(derived, decode_hex("6ad7dedf2d"));
        assert_eq!(derived.len(), 5);
    }

    #[test]
    fn rc4_cryptoapi_standard_derive_rc4_key_b_sha1_keysize_0_is_interpreted_as_40bit() {
        let password = "password";
        let salt: [u8; 16] = [
            0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D,
            0x0E, 0x0F,
        ];

        let derived_0 = derive_rc4_key_b(password, &salt, 0, 0, HashAlg::Sha1);
        let derived_40 = derive_rc4_key_b(password, &salt, 40, 0, HashAlg::Sha1);
        assert_eq!(derived_0, derived_40);
        assert_eq!(derived_0, decode_hex("6ad7dedf2d"));
        assert_eq!(derived_0.len(), 5);
    }

    #[test]
    fn rc4_cryptoapi_standard_derive_rc4_key_b_sha1_keysize_56_truncates() {
        let password = "password";
        let salt: [u8; 16] = [
            0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D,
            0x0E, 0x0F,
        ];

        let derived = derive_rc4_key_b(password, &salt, 56, 0, HashAlg::Sha1);
        assert_eq!(derived, decode_hex("6ad7dedf2da351"));
        assert_eq!(derived.len(), 7);
    }
}
