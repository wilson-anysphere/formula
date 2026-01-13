//! MS-OFFCRYPTO "Standard" (CryptoAPI) primitives shared by RC4 and AES variants.
//!
//! This module intentionally focuses on *deterministic* key derivation helpers so callers can
//! implement the required stream decryption logic for Office's `EncryptedPackage` and verifier
//! fields.

use crate::rc4::Rc4;
use crate::{HashAlgorithm, OffcryptoError};

/// The fixed "spin count" used by Office for CryptoAPI standard encryption.
///
/// MS-OFFCRYPTO standard encryption does not carry this value in the `EncryptionInfo` stream; Office
/// uses a fixed 50k iteration count for the password hash transform.
pub const STANDARD_SPIN_COUNT: u32 = 50_000;

/// Per-block re-keying length for CryptoAPI RC4 stream encryption.
///
/// MS-OFFCRYPTO applies RC4 in 512-byte blocks, re-deriving the RC4 key for each block index `b`
/// using `Hfinal = Hash(H || LE32(b))`.
pub const RC4_BLOCK_LEN: usize = 0x200;

fn digest_len(hash_alg: HashAlgorithm) -> Result<usize, OffcryptoError> {
    match hash_alg {
        HashAlgorithm::Md5 => Ok(16),
        HashAlgorithm::Sha1 => Ok(20),
        // Standard/CryptoAPI encryption only uses SHA1/MD5.
        other => Err(OffcryptoError::InvalidEncryptionInfo {
            context: match other {
                HashAlgorithm::Sha256
                | HashAlgorithm::Sha384
                | HashAlgorithm::Sha512 => "unsupported CryptoAPI hash algorithm",
                _ => "unsupported CryptoAPI hash algorithm",
            },
        }),
    }
}

fn password_to_utf16le_bytes(password: &str) -> Vec<u8> {
    let mut out = Vec::with_capacity(password.len() * 2);
    for unit in password.encode_utf16() {
        out.extend_from_slice(&unit.to_le_bytes());
    }
    out
}

/// Compute the CryptoAPI "iterated hash" `H` from `(password, salt, spin_count, hash_alg)`.
///
/// Pseudocode:
/// ```text
/// H = Hash(salt || password_utf16le)
/// for i in 0..spinCount-1:
///     H = Hash(LE32(i) || H)
/// ```
pub fn iterated_hash_from_password(
    password: &str,
    salt: &[u8],
    spin_count: u32,
    hash_alg: HashAlgorithm,
) -> Result<Vec<u8>, OffcryptoError> {
    let _ = digest_len(hash_alg)?;

    let password_utf16 = password_to_utf16le_bytes(password);
    let mut buf = Vec::with_capacity(salt.len() + password_utf16.len());
    buf.extend_from_slice(salt);
    buf.extend_from_slice(&password_utf16);

    let mut h = hash_alg.digest(&buf);
    for i in 0..spin_count {
        let mut round = Vec::with_capacity(4 + h.len());
        round.extend_from_slice(&i.to_le_bytes());
        round.extend_from_slice(&h);
        h = hash_alg.digest(&round);
    }

    Ok(h)
}

/// Compute `Hfinal = Hash(H || LE32(block))`.
pub fn block_hash(h: &[u8], block: u32, hash_alg: HashAlgorithm) -> Result<Vec<u8>, OffcryptoError> {
    let _ = digest_len(hash_alg)?;
    let mut buf = Vec::with_capacity(h.len() + 4);
    buf.extend_from_slice(h);
    buf.extend_from_slice(&block.to_le_bytes());
    Ok(hash_alg.digest(&buf))
}

/// CryptoAPI `CryptDeriveKey` expansion from a hash value.
///
/// This is the derivation used by CryptoAPI for standard AES encryption.
///
/// The output length is `key_size_bits / 8` (must be divisible by 8).
pub fn crypt_derive_key(
    hash: &[u8],
    key_size_bits: u32,
    hash_alg: HashAlgorithm,
) -> Result<Vec<u8>, OffcryptoError> {
    let hash_len = digest_len(hash_alg)?;
    if hash.len() != hash_len {
        return Err(OffcryptoError::InvalidEncryptionInfo {
            context: "hash length mismatch for CryptoAPI derivation",
        });
    }
    if key_size_bits % 8 != 0 {
        return Err(OffcryptoError::InvalidKeySizeBits { key_size_bits });
    }
    let key_len = (key_size_bits / 8) as usize;

    // buf1 = (hash XOR 0x36) || 0x36*(64-hashLen)
    let mut buf1 = vec![0x36u8; 64];
    let mut buf2 = vec![0x5cu8; 64];
    for i in 0..hash_len {
        buf1[i] ^= hash[i];
        buf2[i] ^= hash[i];
    }
    let x1 = hash_alg.digest(&buf1);
    let x2 = hash_alg.digest(&buf2);

    let mut out = Vec::with_capacity(x1.len() + x2.len());
    out.extend_from_slice(&x1);
    out.extend_from_slice(&x2);
    out.truncate(key_len);
    Ok(out)
}

/// Derive the CryptoAPI RC4 key bytes for a given `block` index.
///
/// **Important:** Standard/CryptoAPI RC4 uses *raw `Hfinal` truncation* (plus the 40-bit padding
/// rule), not `CryptDeriveKey` ipad/opad expansion.
pub fn rc4_key_for_block(
    h: &[u8],
    block: u32,
    key_size_bits: u32,
    hash_alg: HashAlgorithm,
) -> Result<Vec<u8>, OffcryptoError> {
    if key_size_bits % 8 != 0 {
        return Err(OffcryptoError::InvalidKeySizeBits { key_size_bits });
    }
    let key_len = (key_size_bits / 8) as usize;

    let hfinal = block_hash(h, block, hash_alg)?;
    if hfinal.len() < key_len {
        return Err(OffcryptoError::InvalidEncryptionInfo {
            context: "derived RC4 key is shorter than key_size_bits",
        });
    }

    let mut key = hfinal[..key_len].to_vec();
    if key_size_bits == 40 {
        // CryptoAPI "40-bit RC4" keys are represented as a 16-byte key blob where the first 5 bytes
        // contain key material and the remaining 11 bytes are zeros.
        key.truncate(5);
        key.resize(16, 0);
    }
    Ok(key)
}

/// Decrypt a CryptoAPI RC4 stream that is re-keyed per 512-byte block.
pub fn rc4_decrypt_stream(
    ciphertext: &[u8],
    h: &[u8],
    key_size_bits: u32,
    hash_alg: HashAlgorithm,
) -> Result<Vec<u8>, OffcryptoError> {
    let mut out = ciphertext.to_vec();
    for (block, chunk) in out.chunks_exact_mut(RC4_BLOCK_LEN).enumerate() {
        let key = rc4_key_for_block(h, block as u32, key_size_bits, hash_alg)?;
        let mut rc4 = Rc4::new(&key);
        rc4.apply_keystream(chunk);
    }
    let rem = out.len() % RC4_BLOCK_LEN;
    if rem != 0 {
        let block = (out.len() / RC4_BLOCK_LEN) as u32;
        let key = rc4_key_for_block(h, block, key_size_bits, hash_alg)?;
        let mut rc4 = Rc4::new(&key);
        let start = out.len() - rem;
        rc4.apply_keystream(&mut out[start..]);
    }
    Ok(out)
}

