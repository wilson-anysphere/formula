//! MS-OFFCRYPTO "Standard" (CryptoAPI) primitives shared by RC4 and AES variants.
//!
//! This module intentionally focuses on *deterministic* key derivation helpers so callers can
//! implement the required stream decryption logic for Office's `EncryptedPackage` and verifier
//! fields.

use crate::rc4::Rc4;
use crate::{HashAlgorithm, OffcryptoError};
use zeroize::Zeroizing;

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
                HashAlgorithm::Sha256 | HashAlgorithm::Sha384 | HashAlgorithm::Sha512 => {
                    "unsupported CryptoAPI hash algorithm"
                }
                _ => "unsupported CryptoAPI hash algorithm",
            },
        }),
    }
}

fn password_to_utf16le_bytes(password: &str) -> Zeroizing<Vec<u8>> {
    let mut out = Vec::new();
    let _ = out.try_reserve_exact(password.len() * 2);
    for unit in password.encode_utf16() {
        out.extend_from_slice(&unit.to_le_bytes());
    }
    Zeroizing::new(out)
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
) -> Result<Zeroizing<Vec<u8>>, OffcryptoError> {
    let digest_len = digest_len(hash_alg)?;
    debug_assert!(digest_len <= crate::MAX_DIGEST_LEN);

    let password_utf16 = password_to_utf16le_bytes(password);

    // Avoid per-iteration allocations: keep the current digest in a fixed buffer and overwrite it
    // each round.
    let mut h_buf: Zeroizing<[u8; crate::MAX_DIGEST_LEN]> =
        Zeroizing::new([0u8; crate::MAX_DIGEST_LEN]);
    hash_alg.digest_two_into(salt, &password_utf16, &mut h_buf[..digest_len]);

    let mut round_buf: Zeroizing<[u8; crate::MAX_DIGEST_LEN]> =
        Zeroizing::new([0u8; crate::MAX_DIGEST_LEN]);
    for i in 0..spin_count {
        hash_alg.digest_two_into(
            &i.to_le_bytes(),
            &h_buf[..digest_len],
            &mut round_buf[..digest_len],
        );
        h_buf[..digest_len].copy_from_slice(&round_buf[..digest_len]);
    }

    Ok(Zeroizing::new(h_buf[..digest_len].to_vec()))
}

/// Compute `Hfinal = Hash(H || LE32(block))`.
pub fn block_hash(
    h: &[u8],
    block: u32,
    hash_alg: HashAlgorithm,
) -> Result<Zeroizing<Vec<u8>>, OffcryptoError> {
    let digest_len = digest_len(hash_alg)?;
    let mut out = Zeroizing::new(vec![0u8; digest_len]);
    hash_alg.digest_two_into(h, &block.to_le_bytes(), &mut out);
    Ok(out)
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
) -> Result<Zeroizing<Vec<u8>>, OffcryptoError> {
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
    let mut buf1 = Zeroizing::new(vec![0x36u8; 64]);
    let mut buf2 = Zeroizing::new(vec![0x5cu8; 64]);
    for i in 0..hash_len {
        buf1[i] ^= hash[i];
        buf2[i] ^= hash[i];
    }
    let x1 = Zeroizing::new(hash_alg.digest(&buf1));
    let x2 = Zeroizing::new(hash_alg.digest(&buf2));

    if key_len > x1.len() + x2.len() {
        return Err(OffcryptoError::InvalidEncryptionInfo {
            context: "keySize too large for CryptoAPI derivation",
        });
    }

    // Avoid `truncate()` (which can leave the unused bytes in the `Vec` spare capacity) by copying
    // only the prefix we need.
    let mut out = Zeroizing::new(vec![0u8; key_len]);
    let x1_take = std::cmp::min(key_len, x1.len());
    out[..x1_take].copy_from_slice(&x1[..x1_take]);
    if x1_take < key_len {
        let rem = key_len - x1_take;
        out[x1_take..].copy_from_slice(&x2[..rem]);
    }
    Ok(out)
}

/// Derive the CryptoAPI RC4 key bytes for a given `block` index.
///
/// **Important:** Standard/CryptoAPI RC4 uses *raw `Hfinal` truncation*, not `CryptDeriveKey`
/// ipad/opad expansion.
pub fn rc4_key_for_block(
    h: &[u8],
    block: u32,
    key_size_bits: u32,
    hash_alg: HashAlgorithm,
) -> Result<Zeroizing<Vec<u8>>, OffcryptoError> {
    // Standard/CryptoAPI RC4 key derivation uses *raw hash truncation*, not `CryptDeriveKey`:
    // `Hfinal = Hash(H || LE32(block))`, `rc4_key_b = Hfinal[0..keySize/8]`.
    //
    // MS-OFFCRYPTO specifies that for RC4, `keySize=0` MUST be interpreted as 40-bit (5 bytes).
    //
    // Note: some producers treat 40-bit keys as 128-bit keys padded with zeros; that
    // compatibility behavior is handled by the Standard RC4 decryptor, not here.
    let key_size_bits = if key_size_bits == 0 { 40 } else { key_size_bits };
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

    let mut key = Zeroizing::new(vec![0u8; key_len]);
    key.copy_from_slice(&hfinal[..key_len]);
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
        let mut rc4 = Rc4::new(key.as_slice());
        rc4.apply_keystream(chunk);
    }
    let rem = out.len() % RC4_BLOCK_LEN;
    if rem != 0 {
        let block = (out.len() / RC4_BLOCK_LEN) as u32;
        let key = rc4_key_for_block(h, block, key_size_bits, hash_alg)?;
        let mut rc4 = Rc4::new(key.as_slice());
        let start = out.len() - rem;
        rc4.apply_keystream(&mut out[start..]);
    }
    Ok(out)
}
