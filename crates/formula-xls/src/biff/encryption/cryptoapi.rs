use md5::{Digest as _, Md5};
use sha1::Sha1;
use zeroize::{Zeroize, Zeroizing};

use crate::ct::ct_eq;

use super::{rc4::Rc4, records, DecryptError};

// CryptoAPI ALG_ID values for hash functions.
pub(crate) const CALG_MD5: u32 = 0x0000_8003;
pub(crate) const CALG_SHA1: u32 = 0x0000_8004;
// CryptoAPI ALG_ID for RC4.
pub(crate) const CALG_RC4: u32 = 0x0000_6801;

// BIFF8 RC4 CryptoAPI password hardening spin count.
pub(crate) const BIFF8_CRYPTOAPI_SPIN_COUNT: u32 = 50_000;

// CryptoAPI EncryptionHeader is 32 bytes of fixed fields plus an optional CSP name.
// Cap this defensively so malformed files cannot request unbounded allocations.
const MAX_ENCRYPTION_HEADER_SIZE: usize = 4096;

/// Normalize CryptoAPI RC4 `EncryptionHeader.keySize` (bits).
///
/// In MS-OFFCRYPTO Standard/CryptoAPI encryption, `keySize == 0` is defined to mean **40-bit RC4**
/// (legacy export restrictions). Some BIFF8 `FILEPASS` CryptoAPI workbooks follow the same
/// convention, so we treat `0` as `40` for compatibility.
fn normalize_cryptoapi_rc4_key_size_bits(key_size_bits: u32) -> u32 {
    if key_size_bits == 0 {
        40
    } else {
        key_size_bits
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum CryptoApiHashAlg {
    Sha1,
    Md5,
}

impl CryptoApiHashAlg {
    fn from_alg_id_hash(alg_id_hash: u32) -> Option<Self> {
        match alg_id_hash {
            CALG_SHA1 => Some(Self::Sha1),
            CALG_MD5 => Some(Self::Md5),
            _ => None,
        }
    }

    fn digest_len(self) -> usize {
        match self {
            Self::Sha1 => 20,
            Self::Md5 => 16,
        }
    }
}

fn read_u16_le(bytes: &[u8], offset: usize) -> Result<u16, DecryptError> {
    let end = offset.checked_add(2).ok_or_else(|| {
        DecryptError::InvalidFilePass(format!(
            "FILEPASS offset overflow (need u16 at offset {offset}, len={})",
            bytes.len()
        ))
    })?;
    let b = bytes.get(offset..end).ok_or_else(|| {
        DecryptError::InvalidFilePass(format!(
            "truncated FILEPASS payload (need u16 at offset {offset}, len={})",
            bytes.len()
        ))
    })?;
    Ok(u16::from_le_bytes([b[0], b[1]]))
}

fn read_u32_le(bytes: &[u8], offset: usize) -> Result<u32, DecryptError> {
    let end = offset.checked_add(4).ok_or_else(|| {
        DecryptError::InvalidFilePass(format!(
            "FILEPASS offset overflow (need u32 at offset {offset}, len={})",
            bytes.len()
        ))
    })?;
    let b = bytes.get(offset..end).ok_or_else(|| {
        DecryptError::InvalidFilePass(format!(
            "truncated FILEPASS payload (need u32 at offset {offset}, len={})",
            bytes.len()
        ))
    })?;
    Ok(u32::from_le_bytes([b[0], b[1], b[2], b[3]]))
}

fn utf16le_bytes(password: &str) -> Zeroizing<Vec<u8>> {
    let mut out = Zeroizing::new(Vec::new());
    if let Some(cap) = password.len().checked_mul(2) {
        let _ = out.try_reserve(cap);
    }
    for unit in password.encode_utf16() {
        out.extend_from_slice(&unit.to_le_bytes());
    }
    out
}

fn sha1_bytes(chunks: &[&[u8]]) -> [u8; 20] {
    let mut hasher = Sha1::new();
    for chunk in chunks {
        hasher.update(chunk);
    }
    let mut digest = hasher.finalize();
    let mut out = [0u8; 20];
    out.copy_from_slice(&digest);
    digest.as_mut_slice().zeroize();
    out
}

fn md5_bytes(chunks: &[&[u8]]) -> [u8; 16] {
    let mut hasher = Md5::new();
    for chunk in chunks {
        hasher.update(chunk);
    }
    let mut digest = hasher.finalize();
    let mut out = [0u8; 16];
    out.copy_from_slice(&digest);
    digest.as_mut_slice().zeroize();
    out
}

fn derive_key_material(
    hash_alg: CryptoApiHashAlg,
    password: &str,
    salt: &[u8],
    spin_count: u32,
) -> Zeroizing<Vec<u8>> {
    // CryptoAPI password hashing [MS-OFFCRYPTO]:
    //   H0 = Hash(salt + UTF16LE(password))
    //   for i in 0..spin_count-1: H0 = Hash(i_le32 + H0)
    let pw_bytes = utf16le_bytes(password);
    match hash_alg {
        CryptoApiHashAlg::Sha1 => {
            let mut hash = Zeroizing::new(sha1_bytes(&[salt, &pw_bytes]));
            drop(pw_bytes);

            for i in 0..spin_count {
                let iter = i.to_le_bytes();
                let mut next = sha1_bytes(&[&iter, &hash[..]]);
                hash[..].copy_from_slice(&next);
                next.zeroize();
            }

            Zeroizing::new(hash.to_vec())
        }
        CryptoApiHashAlg::Md5 => {
            let mut hash = Zeroizing::new(md5_bytes(&[salt, &pw_bytes]));
            drop(pw_bytes);

            for i in 0..spin_count {
                let iter = i.to_le_bytes();
                let mut next = md5_bytes(&[&iter, &hash[..]]);
                hash[..].copy_from_slice(&next);
                next.zeroize();
            }

            Zeroizing::new(hash.to_vec())
        }
    }
}

fn derive_key_material_legacy(
    hash_alg: CryptoApiHashAlg,
    password: &str,
    salt: &[u8],
) -> Result<Zeroizing<Vec<u8>>, DecryptError> {
    if salt.len() != 16 {
        return Err(DecryptError::InvalidFilePass(format!(
            "CryptoAPI legacy salt length {} (expected 16)",
            salt.len()
        )));
    }
    Ok(derive_key_material(hash_alg, password, salt, 0))
}

fn derive_block_key(
    hash_alg: CryptoApiHashAlg,
    key_material: &[u8],
    block: u32,
    key_len: usize,
    pad_40_bit_to_16: bool,
) -> Zeroizing<Vec<u8>> {
    let block_bytes = block.to_le_bytes();

    // Compatibility note:
    //
    // Some CryptoAPI/WinCrypt implementations treat 40-bit RC4 keys as a 128-bit (16-byte) key
    // where the low 40 bits are set and the remaining 88 bits are zero ("effective key length").
    //
    // This changes the RC4 key scheduling (RC4 includes the key length), so it is *not* equivalent
    // to using the 5-byte key directly. Some producers appear to use the raw 5-byte key; we
    // support both and select the correct variant via verifier validation.
    let key = match hash_alg {
        CryptoApiHashAlg::Sha1 => {
            let mut digest = sha1_bytes(&[key_material, &block_bytes]);
            let key = if key_len == 5 && pad_40_bit_to_16 {
                let mut key = Vec::new();
                key.extend_from_slice(&digest[..5]);
                key.resize(16, 0);
                key
            } else {
                digest[..key_len].to_vec()
            };
            digest.zeroize();
            key
        }
        CryptoApiHashAlg::Md5 => {
            let mut digest = md5_bytes(&[key_material, &block_bytes]);
            let key = if key_len == 5 && pad_40_bit_to_16 {
                let mut key = Vec::new();
                key.extend_from_slice(&digest[..5]);
                key.resize(16, 0);
                key
            } else {
                digest[..key_len].to_vec()
            };
            digest.zeroize();
            key
        }
    };

    Zeroizing::new(key)
}

fn rc4_discard(rc4: &mut Rc4, mut n: usize) {
    // Advance the internal RC4 state without caring about the output bytes.
    let mut scratch = [0u8; 64];
    while n > 0 {
        let take = n.min(scratch.len());
        rc4.apply_keystream(&mut scratch[..take]);
        n -= take;
    }
}

fn decrypt_range_by_offset(
    bytes: &mut [u8],
    start_offset: usize,
    hash_alg: CryptoApiHashAlg,
    key_material: &[u8],
    key_len: usize,
    pad_40_bit_to_16: bool,
) {
    // Decrypt `bytes` assuming the CryptoAPI RC4 keystream position corresponds to
    // `start_offset` in the workbook stream.
    let mut stream_pos = start_offset;
    let mut remaining = bytes.len();
    let mut pos = 0usize;
    while remaining > 0 {
        let block = (stream_pos / super::RC4_BLOCK_SIZE) as u32;
        let in_block = stream_pos % super::RC4_BLOCK_SIZE;
        let take = remaining.min(super::RC4_BLOCK_SIZE - in_block);

        let key = derive_block_key(hash_alg, key_material, block, key_len, pad_40_bit_to_16);
        let mut rc4 = Rc4::new(&key[..]);
        drop(key);
        rc4_discard(&mut rc4, in_block);
        let Some(end) = pos.checked_add(take) else {
            return;
        };
        let Some(chunk) = bytes.get_mut(pos..end) else {
            return;
        };
        rc4.apply_keystream(chunk);

        let Some(next_stream_pos) = stream_pos.checked_add(take) else {
            return;
        };
        stream_pos = next_stream_pos;
        let Some(next_pos) = pos.checked_add(take) else {
            return;
        };
        pos = next_pos;
        remaining -= take;
    }
}

/// Derive the BIFF8 RC4 CryptoAPI key for a given `block_index`.
///
/// This corresponds to the "RC4 CryptoAPI" / "CryptoAPI" encryption scheme used by newer BIFF8
/// workbooks (Excel 2002/2003), which uses SHA-1 + a spin count to harden password derivation.
///
/// `key_len` is `KeySize / 8` from the CryptoAPI header (e.g. 16 for 128-bit RC4, 5 for 40-bit).
///
/// Note: this helper returns the **WinCrypt/Excel-style** key shape for 40-bit RC4: the derived
/// 5 bytes padded with 11 zero bytes (so the effective RC4 key length is 16 bytes). The main BIFF
/// CryptoAPI decryptor also accepts the raw 5-byte key variant by verifier validation.
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

    let Some(hash_alg) = CryptoApiHashAlg::from_alg_id_hash(alg_id_hash) else {
        return Err(DecryptError::UnsupportedEncryption(format!(
            "unsupported CryptoAPI hash alg_id_hash=0x{alg_id_hash:08X}"
        )));
    };

    // Derive base key material (Hspin) and then the per-block key.
    let key_material = derive_key_material(hash_alg, password, salt, spin_count);

    // Preserve historical behavior: if the caller asks for more bytes than the hash digest
    // provides, clamp to the digest length rather than panicking.
    let effective_key_len = if key_len == 5 {
        5
    } else {
        key_len.min(hash_alg.digest_len())
    };

    Ok(derive_block_key(
        hash_alg,
        key_material.as_slice(),
        block_index,
        effective_key_len,
        true,
    ))
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

#[derive(Debug, Clone)]
struct EncryptionHeader {
    alg_id: u32,
    alg_id_hash: u32,
    key_size_bits: u32,
}

#[derive(Debug, Clone)]
struct EncryptionVerifier {
    salt: Vec<u8>,
    encrypted_verifier: [u8; 16],
    verifier_hash_size: u32,
    encrypted_verifier_hash: Vec<u8>,
}

#[derive(Debug, Clone)]
struct CryptoApiEncryptionInfo {
    header: EncryptionHeader,
    verifier: EncryptionVerifier,
}

fn parse_encryption_header(bytes: &[u8]) -> Result<EncryptionHeader, DecryptError> {
    // Fixed-length header fields are 8 DWORDs (32 bytes).
    if bytes.len() < 32 {
        return Err(DecryptError::InvalidFilePass(format!(
            "EncryptionHeader truncated (len={})",
            bytes.len()
        )));
    }

    // EncryptionHeader layout [MS-OFFCRYPTO] 2.3.1:
    //   DWORD Flags;
    //   DWORD SizeExtra;
    //   DWORD AlgID;
    //   DWORD AlgIDHash;
    //   DWORD KeySize;
    //   DWORD ProviderType;
    //   DWORD Reserved1;
    //   DWORD Reserved2;
    //   WCHAR CSPName[];
    let alg_id = read_u32_le(bytes, 8)?;
    let alg_id_hash = read_u32_le(bytes, 12)?;
    let key_size_bits = read_u32_le(bytes, 16)?;
    Ok(EncryptionHeader {
        alg_id,
        alg_id_hash,
        key_size_bits,
    })
}

fn parse_encryption_verifier(bytes: &[u8]) -> Result<EncryptionVerifier, DecryptError> {
    // EncryptionVerifier layout [MS-OFFCRYPTO] 2.3.2:
    //   DWORD SaltSize;
    //   BYTE  Salt[SaltSize];
    //   BYTE  EncryptedVerifier[16];
    //   DWORD VerifierHashSize;
    //   BYTE  EncryptedVerifierHash[VerifierHashSize];

    if bytes.len() < 4 {
        return Err(DecryptError::InvalidFilePass(
            "EncryptionVerifier truncated".to_string(),
        ));
    }

    let salt_size = read_u32_le(bytes, 0)? as usize;
    const MAX_SALT_SIZE: usize = 64;
    if salt_size > MAX_SALT_SIZE {
        return Err(DecryptError::InvalidFilePass(format!(
            "EncryptionVerifier SaltSize {salt_size} exceeds max {MAX_SALT_SIZE}",
        )));
    }

    let salt_start = 4usize;
    let salt_end = salt_start
        .checked_add(salt_size)
        .ok_or_else(|| DecryptError::InvalidFilePass("SaltSize overflow".to_string()))?;
    let verifier_start = salt_end;
    let verifier_end = verifier_start.checked_add(16).ok_or_else(|| {
        DecryptError::InvalidFilePass("EncryptedVerifier offset overflow".to_string())
    })?;
    let hash_size_start = verifier_end;
    let hash_size_end = hash_size_start.checked_add(4).ok_or_else(|| {
        DecryptError::InvalidFilePass("VerifierHashSize offset overflow".to_string())
    })?;

    if hash_size_end > bytes.len() {
        return Err(DecryptError::InvalidFilePass(format!(
            "EncryptionVerifier truncated (len={}, need={hash_size_end})",
            bytes.len()
        )));
    }

    let salt = bytes
        .get(salt_start..salt_end)
        .ok_or_else(|| {
            DecryptError::InvalidFilePass("EncryptionVerifier salt out of bounds".to_string())
        })?
        .to_vec();
    let mut encrypted_verifier = [0u8; 16];
    encrypted_verifier.copy_from_slice(
        bytes
            .get(verifier_start..verifier_end)
            .ok_or_else(|| {
                DecryptError::InvalidFilePass(
                    "EncryptionVerifier encryptedVerifier out of bounds".to_string(),
                )
            })?,
    );
    let verifier_hash_size = read_u32_le(bytes, hash_size_start)?;
    let verifier_hash_size_usize = verifier_hash_size as usize;
    const MAX_VERIFIER_HASH_SIZE: usize = 64;
    if verifier_hash_size_usize > MAX_VERIFIER_HASH_SIZE {
        return Err(DecryptError::InvalidFilePass(format!(
            "EncryptionVerifierHash VerifierHashSize {verifier_hash_size_usize} exceeds max {MAX_VERIFIER_HASH_SIZE}",
        )));
    }
    let encrypted_hash_start = hash_size_end;
    let encrypted_hash_end = encrypted_hash_start
        .checked_add(verifier_hash_size_usize)
        .ok_or_else(|| DecryptError::InvalidFilePass("VerifierHashSize overflow".to_string()))?;
    if encrypted_hash_end > bytes.len() {
        return Err(DecryptError::InvalidFilePass(format!(
            "EncryptionVerifierHash truncated (len={}, need={encrypted_hash_end})",
            bytes.len()
        )));
    }

    let encrypted_verifier_hash = bytes
        .get(encrypted_hash_start..encrypted_hash_end)
        .ok_or_else(|| {
            DecryptError::InvalidFilePass(
                "EncryptionVerifier encryptedVerifierHash out of bounds".to_string(),
            )
        })?
        .to_vec();

    Ok(EncryptionVerifier {
        salt,
        encrypted_verifier,
        verifier_hash_size,
        encrypted_verifier_hash,
    })
}

fn parse_cryptoapi_encryption_info(bytes: &[u8]) -> Result<CryptoApiEncryptionInfo, DecryptError> {
    // EncryptionInfo [MS-OFFCRYPTO] 2.3.1:
    //   EncryptionVersionInfo (Major, Minor) 4 bytes
    //   DWORD Flags;
    //   DWORD HeaderSize;
    //   EncryptionHeader (HeaderSize bytes)
    //   EncryptionVerifier (remaining bytes)
    if bytes.len() < 12 {
        return Err(DecryptError::InvalidFilePass(format!(
            "EncryptionInfo truncated (len={})",
            bytes.len()
        )));
    }

    let header_size = read_u32_le(bytes, 8)? as usize;
    if header_size > MAX_ENCRYPTION_HEADER_SIZE {
        return Err(DecryptError::InvalidFilePass(format!(
            "EncryptionInfo HeaderSize {header_size} exceeds max {MAX_ENCRYPTION_HEADER_SIZE}",
        )));
    }
    let header_start = 12usize;
    let header_end = header_start
        .checked_add(header_size)
        .ok_or_else(|| DecryptError::InvalidFilePass("HeaderSize overflow".to_string()))?;
    if header_end > bytes.len() {
        return Err(DecryptError::InvalidFilePass(format!(
            "EncryptionInfo header out of bounds (len={}, header_end={header_end})",
            bytes.len()
        )));
    }

    let header_bytes = bytes.get(header_start..header_end).ok_or_else(|| {
        DecryptError::InvalidFilePass("EncryptionInfo header out of bounds".to_string())
    })?;
    let verifier_bytes =
        bytes.get(header_end..).ok_or_else(|| {
            DecryptError::InvalidFilePass("EncryptionInfo verifier out of bounds".to_string())
        })?;
    let header = parse_encryption_header(header_bytes)?;
    let verifier = parse_encryption_verifier(verifier_bytes)?;
    Ok(CryptoApiEncryptionInfo { header, verifier })
}

fn parse_cryptoapi_encryption_info_legacy_filepass(
    payload: &[u8],
) -> Result<CryptoApiEncryptionInfo, DecryptError> {
    // Legacy BIFF8 FILEPASS layout for RC4 CryptoAPI:
    // - u16 wEncryptionType (0x0001)
    // - u16 wEncryptionInfo (0x0004)
    // - u16 vMajor
    // - u16 vMinor
    // - u16 reserved (0)
    // - u32 headerSize
    // - EncryptionHeader (headerSize bytes)
    // - EncryptionVerifier (remaining bytes)
    if payload.len() < 14 {
        return Err(DecryptError::InvalidFilePass(format!(
            "FILEPASS payload truncated (len={})",
            payload.len()
        )));
    }

    let encryption_type = read_u16_le(payload, 0)?;
    if encryption_type != super::BIFF8_ENCRYPTION_TYPE_RC4 {
        return Err(DecryptError::InvalidFilePass(format!(
            "FILEPASS wEncryptionType=0x{encryption_type:04X} (expected 0x0001 for RC4)"
        )));
    }

    let encryption_info = read_u16_le(payload, 2)?;
    if encryption_info != super::BIFF8_RC4_ENCRYPTION_INFO_CRYPTOAPI_LEGACY {
        return Err(DecryptError::InvalidFilePass(format!(
            "FILEPASS RC4 wEncryptionInfo=0x{encryption_info:04X} (expected 0x0004)"
        )));
    }

    let header_size = read_u32_le(payload, 10)? as usize;
    if header_size > MAX_ENCRYPTION_HEADER_SIZE {
        return Err(DecryptError::InvalidFilePass(format!(
            "FILEPASS headerSize {header_size} exceeds max {MAX_ENCRYPTION_HEADER_SIZE}",
        )));
    }
    let header_start = 14usize;
    let header_end = header_start
        .checked_add(header_size)
        .ok_or_else(|| DecryptError::InvalidFilePass("headerSize overflow".to_string()))?;
    if header_end > payload.len() {
        return Err(DecryptError::InvalidFilePass(format!(
            "FILEPASS header out of bounds (payload_len={}, header_end={header_end})",
            payload.len()
        )));
    }

    let header_bytes = payload.get(header_start..header_end).ok_or_else(|| {
        DecryptError::InvalidFilePass(format!(
            "FILEPASS header slice out of bounds (payload_len={}, header_start={header_start}, header_end={header_end})",
            payload.len()
        ))
    })?;
    let verifier_bytes = payload.get(header_end..).ok_or_else(|| {
        DecryptError::InvalidFilePass(format!(
            "FILEPASS verifier slice out of bounds (payload_len={}, header_end={header_end})",
            payload.len()
        ))
    })?;

    let header = parse_encryption_header(header_bytes)?;
    let verifier = parse_encryption_verifier(verifier_bytes)?;
    Ok(CryptoApiEncryptionInfo { header, verifier })
}

fn verify_password(
    info: &CryptoApiEncryptionInfo,
    password: &str,
) -> Result<(CryptoApiHashAlg, Zeroizing<Vec<u8>>, usize, bool), DecryptError> {
    let hash_alg =
        CryptoApiHashAlg::from_alg_id_hash(info.header.alg_id_hash).ok_or_else(|| {
            DecryptError::UnsupportedEncryption(format!(
                "CryptoAPI AlgID=0x{:08X} AlgIDHash=0x{:08X}",
                info.header.alg_id, info.header.alg_id_hash
            ))
        })?;
    if info.header.alg_id != CALG_RC4 {
        return Err(DecryptError::UnsupportedEncryption(format!(
            "CryptoAPI AlgID=0x{:08X} AlgIDHash=0x{:08X}",
            info.header.alg_id, info.header.alg_id_hash
        )));
    }

    // RC4 CryptoAPI uses a fixed 16-byte salt.
    if info.verifier.salt.len() != 16 {
        return Err(DecryptError::UnsupportedEncryption(format!(
            "CryptoAPI verifier salt length {} (expected 16)",
            info.verifier.salt.len()
        )));
    }

    let key_size_bits = normalize_cryptoapi_rc4_key_size_bits(info.header.key_size_bits);
    if key_size_bits % 8 != 0 {
        return Err(DecryptError::UnsupportedEncryption(format!(
            "CryptoAPI keySizeBits={key_size_bits} (not byte-aligned)"
        )));
    }
    let key_len = (key_size_bits / 8) as usize;
    if !matches!(key_len, 5 | 7 | 16) {
        return Err(DecryptError::UnsupportedEncryption(format!(
            "CryptoAPI keySizeBits={key_size_bits}"
        )));
    }

    let verifier_hash_size = info.verifier.verifier_hash_size as usize;
    let expected_verifier_hash_size = hash_alg.digest_len();
    if verifier_hash_size != expected_verifier_hash_size {
        return Err(DecryptError::UnsupportedEncryption(format!(
            "CryptoAPI verifierHashSize={verifier_hash_size}"
        )));
    }

    // Derive base key material and decrypt verifier using block 0.
    let key_material = derive_key_material(
        hash_alg,
        password,
        info.verifier.salt.as_slice(),
        BIFF8_CRYPTOAPI_SPIN_COUNT,
    );

    let pad_candidates: &[bool] = if key_len == 5 {
        // Excel/WinCrypt uses the padded 16-byte form for 40-bit RC4 keys. Some producers appear to
        // use the raw 5-byte key, so try both and select via verifier validation.
        &[true, false]
    } else {
        &[false]
    };

    for &pad_40_bit_to_16 in pad_candidates {
        let key0 = derive_block_key(
            hash_alg,
            key_material.as_slice(),
            0,
            key_len,
            pad_40_bit_to_16,
        );
        let mut rc4 = Rc4::new(&key0[..]);
        drop(key0);

        let mut verifier = Zeroizing::new(info.verifier.encrypted_verifier);
        rc4.apply_keystream(&mut verifier[..]);

        let mut verifier_hash = Zeroizing::new(info.verifier.encrypted_verifier_hash.clone());
        rc4.apply_keystream(&mut verifier_hash[..]);

        let expected = Zeroizing::new(match hash_alg {
            CryptoApiHashAlg::Sha1 => sha1_bytes(&[&verifier[..]]).to_vec(),
            CryptoApiHashAlg::Md5 => md5_bytes(&[&verifier[..]]).to_vec(),
        });
        if verifier_hash.len() < verifier_hash_size {
            return Err(DecryptError::InvalidFilePass(format!(
                "EncryptedVerifierHash length {} shorter than VerifierHashSize {verifier_hash_size}",
                verifier_hash.len()
            )));
        }
        if ct_eq(
            &verifier_hash[..verifier_hash_size],
            &expected[..verifier_hash_size],
        ) {
            return Ok((hash_alg, key_material, key_len, pad_40_bit_to_16));
        }
    }

    Err(DecryptError::WrongPassword)
}

fn verify_password_legacy(
    info: &CryptoApiEncryptionInfo,
    password: &str,
) -> Result<(CryptoApiHashAlg, Zeroizing<Vec<u8>>, usize, bool), DecryptError> {
    // Some legacy CryptoAPI FILEPASS headers omit/zero AlgIDHash; Excel still behaves like SHA1.
    // Default to SHA1 for AlgIDHash=0 to preserve historical behaviour, but honor MD5 when set.
    let hash_alg = match CryptoApiHashAlg::from_alg_id_hash(info.header.alg_id_hash) {
        Some(hash_alg) => hash_alg,
        None if info.header.alg_id_hash == 0 => CryptoApiHashAlg::Sha1,
        None => {
            return Err(DecryptError::UnsupportedEncryption(format!(
                "CryptoAPI AlgID=0x{:08X} AlgIDHash=0x{:08X}",
                info.header.alg_id, info.header.alg_id_hash
            )))
        }
    };
    if info.header.alg_id != CALG_RC4 {
        return Err(DecryptError::UnsupportedEncryption(format!(
            "CryptoAPI AlgID=0x{:08X} AlgIDHash=0x{:08X}",
            info.header.alg_id, info.header.alg_id_hash
        )));
    }

    let key_size_bits = normalize_cryptoapi_rc4_key_size_bits(info.header.key_size_bits);
    if key_size_bits % 8 != 0 {
        return Err(DecryptError::UnsupportedEncryption(format!(
            "CryptoAPI keySizeBits={key_size_bits} (not byte-aligned)"
        )));
    }
    let key_len = (key_size_bits / 8) as usize;
    if !matches!(key_len, 5 | 7 | 16) {
        return Err(DecryptError::UnsupportedEncryption(format!(
            "CryptoAPI keySizeBits={key_size_bits}"
        )));
    }

    let verifier_hash_size = info.verifier.verifier_hash_size as usize;
    let expected_verifier_hash_size = hash_alg.digest_len();
    if verifier_hash_size != expected_verifier_hash_size {
        return Err(DecryptError::UnsupportedEncryption(format!(
            "CryptoAPI verifierHashSize={verifier_hash_size}"
        )));
    }

    let key_material =
        derive_key_material_legacy(hash_alg, password, info.verifier.salt.as_slice())?;
    let pad_candidates: &[bool] = if key_len == 5 {
        &[true, false]
    } else {
        &[false]
    };
    for &pad_40_bit_to_16 in pad_candidates {
        let key0 = derive_block_key(
            hash_alg,
            key_material.as_slice(),
            0,
            key_len,
            pad_40_bit_to_16,
        );
        let mut rc4 = Rc4::new(&key0[..]);
        drop(key0);

        let mut verifier = Zeroizing::new(info.verifier.encrypted_verifier);
        rc4.apply_keystream(&mut verifier[..]);

        let mut verifier_hash = Zeroizing::new(info.verifier.encrypted_verifier_hash.clone());
        rc4.apply_keystream(&mut verifier_hash[..]);

        let expected = Zeroizing::new(match hash_alg {
            CryptoApiHashAlg::Sha1 => sha1_bytes(&[&verifier[..]]).to_vec(),
            CryptoApiHashAlg::Md5 => md5_bytes(&[&verifier[..]]).to_vec(),
        });
        if verifier_hash.len() < verifier_hash_size {
            return Err(DecryptError::InvalidFilePass(format!(
                "EncryptedVerifierHash length {} shorter than VerifierHashSize {verifier_hash_size}",
                verifier_hash.len()
            )));
        }
        if ct_eq(
            &verifier_hash[..verifier_hash_size],
            &expected[..verifier_hash_size],
        ) {
            return Ok((hash_alg, key_material, key_len, pad_40_bit_to_16));
        }
    }

    Err(DecryptError::WrongPassword)
}

struct PayloadRc4 {
    hash_alg: CryptoApiHashAlg,
    key_material: Zeroizing<Vec<u8>>,
    key_len: usize,
    pad_40_bit_to_16: bool,
    block: u32,
    pos_in_block: usize,
    rc4: Rc4,
}

impl PayloadRc4 {
    fn new(
        hash_alg: CryptoApiHashAlg,
        key_material: Zeroizing<Vec<u8>>,
        key_len: usize,
        pad_40_bit_to_16: bool,
    ) -> Self {
        let key0 = derive_block_key(
            hash_alg,
            key_material.as_slice(),
            0,
            key_len,
            pad_40_bit_to_16,
        );
        let rc4 = Rc4::new(&key0);
        drop(key0);
        Self {
            hash_alg,
            key_material,
            key_len,
            pad_40_bit_to_16,
            block: 0,
            pos_in_block: 0,
            rc4,
        }
    }

    fn rekey(&mut self) {
        self.block = self.block.wrapping_add(1);
        let key = derive_block_key(
            self.hash_alg,
            self.key_material.as_slice(),
            self.block,
            self.key_len,
            self.pad_40_bit_to_16,
        );
        self.rc4 = Rc4::new(&key);
        self.pos_in_block = 0;
    }

    fn apply_keystream(&mut self, mut data: &mut [u8]) {
        while !data.is_empty() {
            if self.pos_in_block == super::RC4_BLOCK_SIZE {
                self.rekey();
            }

            let remaining_in_block = super::RC4_BLOCK_SIZE - self.pos_in_block;
            let chunk_len = data.len().min(remaining_in_block);
            let (chunk, rest) = data.split_at_mut(chunk_len);
            self.rc4.apply_keystream(chunk);
            self.pos_in_block += chunk_len;
            data = rest;
        }
    }
}

fn is_never_encrypted_record(record_id: u16) -> bool {
    // Mirror Apache POI's BIFF8 RC4 CryptoAPI legacy behaviour:
    // - BOF
    // - FILEPASS
    // - INTERFACEHDR (0x00E1)
    records::is_bof_record(record_id)
        || record_id == records::RECORD_FILEPASS
        || record_id == super::RECORD_INTERFACEHDR
}

fn decrypt_cryptoapi_standard(
    workbook_stream: &mut [u8],
    encrypted_start: usize,
    password: &str,
    filepass_payload: &[u8],
) -> Result<(), DecryptError> {
    // FILEPASS payload:
    //   u16 wEncryptionType
    //   u16 wEncryptionSubType (=0x0002)
    //   u32 dwEncryptionInfoLen
    //   EncryptionInfo bytes...
    if filepass_payload.len() < 8 {
        return Err(DecryptError::InvalidFilePass(format!(
            "FILEPASS payload truncated (len={})",
            filepass_payload.len()
        )));
    }

    let enc_info_len = read_u32_le(filepass_payload, 4)? as usize;
    let enc_info_start = 8usize;
    let enc_info_end = enc_info_start
        .checked_add(enc_info_len)
        .ok_or_else(|| DecryptError::InvalidFilePass("dwEncryptionInfoLen overflow".to_string()))?;
    if enc_info_end > filepass_payload.len() {
        return Err(DecryptError::InvalidFilePass(format!(
            "FILEPASS dwEncryptionInfoLen out of bounds (len={}, need={enc_info_end})",
            filepass_payload.len()
        )));
    }

    let enc_info_bytes = filepass_payload.get(enc_info_start..enc_info_end).ok_or_else(|| {
        DecryptError::InvalidFilePass("FILEPASS EncryptionInfo out of bounds".to_string())
    })?;
    let info = parse_cryptoapi_encryption_info(enc_info_bytes)?;
    let (hash_alg, key_material, key_len, pad_40_bit_to_16) = verify_password(&info, password)?;

    let mut cipher = PayloadRc4::new(hash_alg, key_material, key_len, pad_40_bit_to_16);

    // Decrypt record payload bytes after FILEPASS using the record-payload-only stream model.
    let mut offset = encrypted_start;
    while offset < workbook_stream.len() {
        let remaining = workbook_stream.len().checked_sub(offset).unwrap_or(0);
        if remaining < 4 {
            // Some writers may include trailing padding bytes after the final EOF record.
            break;
        }

        let Some(header) = workbook_stream.get(offset..).and_then(|rest| rest.get(..4)) else {
            return Err(DecryptError::InvalidFilePass(
                "truncated BIFF record header while decrypting CryptoAPI".to_string(),
            ));
        };
        let record_id = u16::from_le_bytes([header[0], header[1]]);
        let len = u16::from_le_bytes([header[2], header[3]]) as usize;
        let data_start = offset.checked_add(4).ok_or_else(|| {
            DecryptError::InvalidFilePass("BIFF record offset overflow".to_string())
        })?;
        let data_end = data_start.checked_add(len).ok_or_else(|| {
            DecryptError::InvalidFilePass("BIFF record length overflow".to_string())
        })?;
        if data_end > workbook_stream.len() {
            return Err(DecryptError::InvalidFilePass(format!(
                "BIFF record 0x{record_id:04X} at offset {offset} extends past end of stream (len={}, end={data_end})",
                workbook_stream.len()
            )));
        }

        let Some(payload) = workbook_stream.get_mut(data_start..data_end) else {
            return Err(DecryptError::InvalidFilePass(
                "BIFF record payload out of bounds while decrypting CryptoAPI".to_string(),
            ));
        };
        cipher.apply_keystream(payload);
        offset = data_end;
    }

    Ok(())
}

fn decrypt_cryptoapi_legacy(
    workbook_stream: &mut [u8],
    encrypted_start: usize,
    password: &str,
    filepass_payload: &[u8],
) -> Result<(), DecryptError> {
    let info = parse_cryptoapi_encryption_info_legacy_filepass(filepass_payload)?;
    let (hash_alg, key_material, key_len, pad_40_bit_to_16) =
        verify_password_legacy(&info, password)?;

    let mut offset = encrypted_start;
    // "Encryption stream position" is keyed by the absolute offset within the workbook stream.
    let mut stream_pos: usize = encrypted_start;

    while offset < workbook_stream.len() {
        let remaining = workbook_stream.len().checked_sub(offset).unwrap_or(0);
        if remaining < 4 {
            break;
        }

        let Some(header) = workbook_stream.get(offset..).and_then(|rest| rest.get(..4)) else {
            return Err(DecryptError::InvalidFilePass(
                "truncated BIFF record header while decrypting CryptoAPI legacy".to_string(),
            ));
        };
        let record_id = u16::from_le_bytes([header[0], header[1]]);
        let len = u16::from_le_bytes([header[2], header[3]]) as usize;
        let data_start = offset.checked_add(4).ok_or_else(|| {
            DecryptError::InvalidFilePass("BIFF record offset overflow".to_string())
        })?;
        let data_end = data_start.checked_add(len).ok_or_else(|| {
            DecryptError::InvalidFilePass("BIFF record length overflow".to_string())
        })?;
        if data_end > workbook_stream.len() {
            return Err(DecryptError::InvalidFilePass(format!(
                "BIFF record 0x{record_id:04X} at offset {offset} extends past end of stream (len={}, end={data_end})",
                workbook_stream.len()
            )));
        }

        // Record headers are not encrypted but still advance the CryptoAPI RC4 stream position.
        stream_pos = stream_pos
            .checked_add(4)
            .ok_or_else(|| DecryptError::InvalidFilePass("stream position overflow".to_string()))?;

        if !is_never_encrypted_record(record_id) {
            match record_id {
                // BoundSheet8.lbPlyPos (first 4 bytes) must remain plaintext so sheet offsets remain
                // valid after masking FILEPASS. The remaining fields are encrypted.
                super::RECORD_BOUNDSHEET => {
                    if len > 4 {
                        let decrypt_start = stream_pos.checked_add(4).ok_or_else(|| {
                            DecryptError::InvalidFilePass("stream position overflow".to_string())
                        })?;
                        let range_start = data_start.checked_add(4).ok_or_else(|| {
                            DecryptError::InvalidFilePass("BIFF record offset overflow".to_string())
                        })?;
                        let Some(payload) = workbook_stream.get_mut(range_start..data_end) else {
                            return Err(DecryptError::InvalidFilePass(
                                "BIFF record payload out of bounds while decrypting BoundSheet (CryptoAPI legacy)"
                                    .to_string(),
                            ));
                        };
                        decrypt_range_by_offset(
                            payload,
                            decrypt_start,
                            hash_alg,
                            key_material.as_slice(),
                            key_len,
                            pad_40_bit_to_16,
                        );
                    }
                }
                _ => {
                    let Some(payload) = workbook_stream.get_mut(data_start..data_end) else {
                        return Err(DecryptError::InvalidFilePass(
                            "BIFF record payload out of bounds while decrypting CryptoAPI legacy"
                                .to_string(),
                        ));
                    };
                    decrypt_range_by_offset(
                        payload,
                        stream_pos,
                        hash_alg,
                        key_material.as_slice(),
                        key_len,
                        pad_40_bit_to_16,
                    )
                }
            }
        }

        // Advance past the record payload, regardless of whether we decrypted it.
        stream_pos = stream_pos
            .checked_add(len)
            .ok_or_else(|| DecryptError::InvalidFilePass("stream position overflow".to_string()))?;
        offset = data_end;
    }

    Ok(())
}

/// Decrypt a BIFF8 workbook stream protected with RC4 CryptoAPI.
///
/// This supports both:
/// - Standard FILEPASS CryptoAPI layout (`wEncryptionSubType == 0x0002`)
/// - Legacy FILEPASS CryptoAPI layout (`wEncryptionInfo == 0x0004`)
pub(crate) fn decrypt_workbook_stream_rc4_cryptoapi(
    workbook_stream: &mut [u8],
    encrypted_start: usize,
    password: &str,
    filepass_payload: &[u8],
) -> Result<(), DecryptError> {
    let encryption_type = read_u16_le(filepass_payload, 0)?;
    if encryption_type != super::BIFF8_ENCRYPTION_TYPE_RC4 {
        return Err(DecryptError::UnsupportedEncryption(format!(
            "FILEPASS wEncryptionType=0x{encryption_type:04X} (expected RC4)"
        )));
    }

    let second = read_u16_le(filepass_payload, 2)?;
    match second {
        super::BIFF8_RC4_SUBTYPE_CRYPTOAPI => {
            decrypt_cryptoapi_standard(workbook_stream, encrypted_start, password, filepass_payload)
        }
        super::BIFF8_RC4_ENCRYPTION_INFO_CRYPTOAPI_LEGACY => {
            decrypt_cryptoapi_legacy(workbook_stream, encrypted_start, password, filepass_payload)
        }
        other => Err(DecryptError::UnsupportedEncryption(format!(
            "FILEPASS RC4 wEncryptionSubType/wEncryptionInfo=0x{other:04X}"
        ))),
    }
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
            0xF8, 0x06, 0xD7, 0x4E, 0x99, 0x94, 0x8E, 0xE8, 0xD3, 0x68, 0xD6, 0x1C, 0xEA, 0xAA,
            0x7C, 0x36,
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

        let derived_key =
            derive_biff8_cryptoapi_key(CALG_SHA1, password, &salt, spin_count, 0, key_len)
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
            0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D,
            0x0E, 0x0F,
        ];
        let spin_count: u32 = 50_000;

        let expected: &[(u32, [u8; 16])] = &[
            (
                0,
                [
                    0x69, 0xBA, 0xDC, 0xAE, 0x24, 0x48, 0x68, 0xE2, 0x09, 0xD4, 0xE0, 0x53, 0xCC,
                    0xD2, 0xA3, 0xBC,
                ],
            ),
            (
                1,
                [
                    0x6F, 0x4D, 0x50, 0x2A, 0xB3, 0x77, 0x00, 0xFF, 0xDA, 0xB5, 0x70, 0x41, 0x60,
                    0x45, 0x5B, 0x47,
                ],
            ),
            (
                2,
                [
                    0xAC, 0x69, 0x02, 0x2E, 0x39, 0x6C, 0x77, 0x50, 0x87, 0x21, 0x33, 0xF3, 0x7E,
                    0x2C, 0x7A, 0xFC,
                ],
            ),
            (
                3,
                [
                    0x1B, 0x05, 0x6E, 0x71, 0x18, 0xAB, 0x8D, 0x35, 0xE9, 0xD6, 0x7A, 0xDE, 0xE8,
                    0xB1, 0x11, 0x04,
                ],
            ),
        ];

        for (block, expected_key) in expected {
            let key = derive_biff8_cryptoapi_key(CALG_MD5, password, &salt, spin_count, *block, 16)
                .expect("derive");
            assert_eq!(key.as_slice(), expected_key, "block={block}");
        }

        // 40-bit CryptoAPI RC4 keys are represented as a 16-byte RC4 key with the high 88 bits
        // zero.
        let key_40 = derive_biff8_cryptoapi_key(CALG_MD5, password, &salt, spin_count, 0, 5)
            .expect("derive 40-bit");
        let mut expected_40 = vec![0x69, 0xBA, 0xDC, 0xAE, 0x24];
        expected_40.resize(16, 0);
        assert_eq!(key_40.as_slice(), expected_40.as_slice());
        assert_eq!(key_40.len(), 16);
    }

    #[test]
    fn decrypt_cryptoapi_standard_accepts_raw_40_bit_rc4_keys() {
        // Compatibility regression test: some producers emit 40-bit CryptoAPI RC4 workbooks where
        // the RC4 key is the raw 5-byte truncation (`keyLen = keySize/8`) rather than the WinCrypt
        // 16-byte "effective key length" form. `formula-xls` should accept both by validating the
        // verifier with both key shapes.
        let password = "pw";
        let bof_payload = [0x00, 0x06, 0x05, 0x00];

        fn record(record_id: u16, payload: &[u8]) -> Vec<u8> {
            let mut out = Vec::new();
            out.extend_from_slice(&record_id.to_le_bytes());
            out.extend_from_slice(&(payload.len() as u16).to_le_bytes());
            out.extend_from_slice(payload);
            out
        }

        fn dummy_payload(len: usize, seed: u8) -> Vec<u8> {
            (0..len)
                .map(|i| {
                    seed.wrapping_add((i as u8).wrapping_mul(31))
                        .wrapping_add((i >> 8) as u8)
                })
                .collect()
        }

        // Build a minimal CryptoAPI `EncryptionInfo` with `keySizeBits=0` (=> 40-bit), SHA-1, and a
        // deterministic salt/verifier.
        let salt: [u8; 16] = core::array::from_fn(|i| 0x10u8.wrapping_add(i as u8));
        let verifier_plain: [u8; 16] = core::array::from_fn(|i| 0xA0u8.wrapping_add(i as u8));
        let verifier_hash_plain = sha1_bytes(&[&verifier_plain]);

        let key_material = derive_key_material(
            CryptoApiHashAlg::Sha1,
            password,
            &salt,
            BIFF8_CRYPTOAPI_SPIN_COUNT,
        );
        let key0_raw =
            derive_block_key(CryptoApiHashAlg::Sha1, key_material.as_slice(), 0, 5, false);
        let mut rc4 = Rc4::new(&key0_raw[..]);
        drop(key0_raw);
        let mut verifier_buf = [0u8; 36];
        verifier_buf[..16].copy_from_slice(&verifier_plain);
        verifier_buf[16..].copy_from_slice(&verifier_hash_plain);
        rc4.apply_keystream(&mut verifier_buf);
        let encrypted_verifier: [u8; 16] = verifier_buf[..16].try_into().unwrap();
        let encrypted_verifier_hash: [u8; 20] = verifier_buf[16..].try_into().unwrap();

        // EncryptionHeader (32 bytes, no CSP name).
        let mut enc_header = Vec::<u8>::new();
        enc_header.extend_from_slice(&0u32.to_le_bytes()); // Flags
        enc_header.extend_from_slice(&0u32.to_le_bytes()); // SizeExtra
        enc_header.extend_from_slice(&CALG_RC4.to_le_bytes()); // AlgID
        enc_header.extend_from_slice(&CALG_SHA1.to_le_bytes()); // AlgIDHash
        enc_header.extend_from_slice(&0u32.to_le_bytes()); // KeySize bits (0 => 40-bit)
        enc_header.extend_from_slice(&0u32.to_le_bytes()); // ProviderType
        enc_header.extend_from_slice(&0u32.to_le_bytes()); // Reserved1
        enc_header.extend_from_slice(&0u32.to_le_bytes()); // Reserved2

        // EncryptionVerifier.
        let mut enc_verifier = Vec::<u8>::new();
        enc_verifier.extend_from_slice(&(salt.len() as u32).to_le_bytes());
        enc_verifier.extend_from_slice(&salt);
        enc_verifier.extend_from_slice(&encrypted_verifier);
        enc_verifier.extend_from_slice(&(encrypted_verifier_hash.len() as u32).to_le_bytes());
        enc_verifier.extend_from_slice(&encrypted_verifier_hash);

        // EncryptionInfo (version fields are ignored by parser; use 4.2).
        let mut enc_info = Vec::<u8>::new();
        enc_info.extend_from_slice(&4u16.to_le_bytes()); // Major
        enc_info.extend_from_slice(&2u16.to_le_bytes()); // Minor
        enc_info.extend_from_slice(&0u32.to_le_bytes()); // Flags
        enc_info.extend_from_slice(&(enc_header.len() as u32).to_le_bytes()); // HeaderSize
        enc_info.extend_from_slice(&enc_header);
        enc_info.extend_from_slice(&enc_verifier);

        // FILEPASS payload (layout A).
        let mut filepass_payload = Vec::<u8>::new();
        filepass_payload.extend_from_slice(&super::super::BIFF8_ENCRYPTION_TYPE_RC4.to_le_bytes());
        filepass_payload
            .extend_from_slice(&super::super::BIFF8_RC4_SUBTYPE_CRYPTOAPI.to_le_bytes());
        filepass_payload.extend_from_slice(&(enc_info.len() as u32).to_le_bytes());
        filepass_payload.extend_from_slice(&enc_info);

        // Build plaintext workbook stream.
        let r1 = record(0x00FC, &dummy_payload(1000, 0x11));
        let r2 = record(0x00FD, &dummy_payload(80, 0x22));
        let bof = record(records::RECORD_BOF_BIFF8, &bof_payload);
        let filepass = record(records::RECORD_FILEPASS, &filepass_payload);

        let plain = [bof, filepass, r1, r2, record(records::RECORD_EOF, &[])].concat();

        // Encrypt payload bytes after FILEPASS using the **raw 5-byte** RC4 key variant.
        let encrypted_start = filepass_payload
            .len()
            .checked_add(4)
            .and_then(|l| l.checked_add(bof_payload.len() + 4))
            .unwrap();
        let mut encrypted = plain.clone();
        let mut cipher = PayloadRc4::new(CryptoApiHashAlg::Sha1, key_material, 5, false);
        let mut offset = encrypted_start;
        while offset < encrypted.len() {
            if encrypted.len() - offset < 4 {
                break;
            }
            let Some(header) = encrypted.get(offset..).and_then(|rest| rest.get(..4)) else {
                break;
            };
            let len = u16::from_le_bytes([header[2], header[3]]) as usize;
            let data_start = match offset.checked_add(4) {
                Some(v) => v,
                None => break,
            };
            let data_end = match data_start.checked_add(len) {
                Some(v) => v,
                None => break,
            };
            let Some(payload) = encrypted.get_mut(data_start..data_end) else {
                break;
            };
            cipher.apply_keystream(payload);
            offset = data_end;
        }

        // Decrypt must recover the original plaintext stream.
        crate::biff::encryption::decrypt_workbook_stream(&mut encrypted, password)
            .expect("decrypt");
        assert_eq!(encrypted, plain);
    }

    #[test]
    fn decrypt_cryptoapi_legacy_accepts_raw_40_bit_rc4_keys() {
        // Like `decrypt_cryptoapi_standard_accepts_raw_40_bit_rc4_keys`, but for the legacy BIFF8
        // FILEPASS CryptoAPI layout (`wEncryptionInfo == 0x0004`) and the legacy absolute-position
        // RC4 stream model.
        let password = "pw";
        let bof_payload = [0x00, 0x06, 0x05, 0x00];

        fn record(record_id: u16, payload: &[u8]) -> Vec<u8> {
            let mut out = Vec::new();
            out.extend_from_slice(&record_id.to_le_bytes());
            out.extend_from_slice(&(payload.len() as u16).to_le_bytes());
            out.extend_from_slice(payload);
            out
        }

        fn dummy_payload(len: usize, seed: u8) -> Vec<u8> {
            (0..len)
                .map(|i| {
                    seed.wrapping_add((i as u8).wrapping_mul(31))
                        .wrapping_add((i >> 8) as u8)
                })
                .collect()
        }

        // Deterministic salt/verifier for reproducibility.
        let salt: [u8; 16] = core::array::from_fn(|i| 0x10u8.wrapping_add(i as u8));
        let verifier_plain: [u8; 16] = core::array::from_fn(|i| 0xA0u8.wrapping_add(i as u8));
        let verifier_hash_plain = sha1_bytes(&[&verifier_plain]);

        // Encrypt verifier using the **raw 5-byte** RC4 key variant (no WinCrypt padding).
        let key_material =
            derive_key_material_legacy(CryptoApiHashAlg::Sha1, password, &salt).expect("kdf");
        let key0_raw =
            derive_block_key(CryptoApiHashAlg::Sha1, key_material.as_slice(), 0, 5, false);
        let mut rc4 = Rc4::new(&key0_raw[..]);
        drop(key0_raw);
        let mut verifier_buf = [0u8; 36];
        verifier_buf[..16].copy_from_slice(&verifier_plain);
        verifier_buf[16..].copy_from_slice(&verifier_hash_plain);
        rc4.apply_keystream(&mut verifier_buf);
        let encrypted_verifier: [u8; 16] = verifier_buf[..16].try_into().unwrap();
        let encrypted_verifier_hash: [u8; 20] = verifier_buf[16..].try_into().unwrap();

        // EncryptionHeader (32 bytes, no CSP name).
        let mut enc_header = Vec::<u8>::new();
        enc_header.extend_from_slice(&0u32.to_le_bytes()); // Flags
        enc_header.extend_from_slice(&0u32.to_le_bytes()); // SizeExtra
        enc_header.extend_from_slice(&CALG_RC4.to_le_bytes()); // AlgID
        enc_header.extend_from_slice(&CALG_SHA1.to_le_bytes()); // AlgIDHash
        enc_header.extend_from_slice(&0u32.to_le_bytes()); // KeySize bits (0 => 40-bit)
        enc_header.extend_from_slice(&0u32.to_le_bytes()); // ProviderType
        enc_header.extend_from_slice(&0u32.to_le_bytes()); // Reserved1
        enc_header.extend_from_slice(&0u32.to_le_bytes()); // Reserved2

        // EncryptionVerifier.
        let mut enc_verifier = Vec::<u8>::new();
        enc_verifier.extend_from_slice(&(salt.len() as u32).to_le_bytes());
        enc_verifier.extend_from_slice(&salt);
        enc_verifier.extend_from_slice(&encrypted_verifier);
        enc_verifier.extend_from_slice(&(encrypted_verifier_hash.len() as u32).to_le_bytes());
        enc_verifier.extend_from_slice(&encrypted_verifier_hash);

        // FILEPASS payload (legacy layout B).
        let mut filepass_payload = Vec::<u8>::new();
        filepass_payload.extend_from_slice(&super::super::BIFF8_ENCRYPTION_TYPE_RC4.to_le_bytes()); // wEncryptionType
        filepass_payload
            .extend_from_slice(&super::super::BIFF8_RC4_ENCRYPTION_INFO_CRYPTOAPI_LEGACY.to_le_bytes()); // wEncryptionInfo (0x0004)
        filepass_payload.extend_from_slice(&4u16.to_le_bytes()); // vMajor
        filepass_payload.extend_from_slice(&2u16.to_le_bytes()); // vMinor
        filepass_payload.extend_from_slice(&0u16.to_le_bytes()); // reserved
        filepass_payload.extend_from_slice(&(enc_header.len() as u32).to_le_bytes()); // headerSize
        filepass_payload.extend_from_slice(&enc_header);
        filepass_payload.extend_from_slice(&enc_verifier);

        // Build plaintext workbook stream.
        let r1 = record(0x00FC, &dummy_payload(1000, 0x11));
        let r2 = record(0x00FD, &dummy_payload(80, 0x22));
        let bof = record(records::RECORD_BOF_BIFF8, &bof_payload);
        let filepass = record(records::RECORD_FILEPASS, &filepass_payload);
        let plain = [bof, filepass, r1, r2, record(records::RECORD_EOF, &[])].concat();

        // Encrypt record payload bytes after FILEPASS using the legacy absolute-position stream
        // mapping and the raw 5-byte key variant.
        let encrypted_start = filepass_payload
            .len()
            .checked_add(4)
            .and_then(|l| l.checked_add(bof_payload.len() + 4))
            .unwrap();
        let mut encrypted = plain.clone();

        let mut offset = encrypted_start;
        let mut stream_pos = encrypted_start;
        while offset < encrypted.len() {
            let remaining = encrypted.len().checked_sub(offset).unwrap_or(0);
            if remaining < 4 {
                break;
            }

            let Some(header) = encrypted.get(offset..).and_then(|rest| rest.get(..4)) else {
                break;
            };
            let record_id = u16::from_le_bytes([header[0], header[1]]);
            let len = u16::from_le_bytes([header[2], header[3]]) as usize;
            let data_start = match offset.checked_add(4) {
                Some(v) => v,
                None => break,
            };
            let data_end = match data_start.checked_add(len) {
                Some(v) => v,
                None => break,
            };

            // Record headers are not encrypted but still advance the CryptoAPI RC4 stream position.
            stream_pos += 4;

            // Avoid record types with special-case decryption behavior (we only emit dummy records
            // here), but mirror the "never encrypted" list for completeness.
            if !is_never_encrypted_record(record_id) && len > 0 {
                if let Some(payload) = encrypted.get_mut(data_start..data_end) {
                    decrypt_range_by_offset(
                        payload,
                        stream_pos,
                        CryptoApiHashAlg::Sha1,
                        key_material.as_slice(),
                        5,
                        false,
                    );
                } else {
                    break;
                }
            }

            stream_pos += len;
            offset = data_end;
        }

        // Decrypt must recover the original plaintext stream.
        crate::biff::encryption::decrypt_workbook_stream(&mut encrypted, password)
            .expect("decrypt");
        assert_eq!(encrypted, plain);
    }

    #[test]
    fn decrypt_cryptoapi_legacy_respects_boundsheet_and_never_encrypted_records() {
        // Legacy CryptoAPI RC4 in BIFF8 has a few record-specific rules:
        // - INTERFACEHDR is never encrypted.
        // - BoundSheet8.lbPlyPos (first 4 bytes) is never encrypted.
        //
        // This test ensures those exceptions are honored and that "40-bit" keys padded to 16 bytes
        // (the Excel/WinCrypt effective-key-length behavior) decrypt correctly.
        let password = "pw";
        let bof_payload = [0x00, 0x06, 0x05, 0x00];

        fn record(record_id: u16, payload: &[u8]) -> Vec<u8> {
            let mut out = Vec::new();
            out.extend_from_slice(&record_id.to_le_bytes());
            out.extend_from_slice(&(payload.len() as u16).to_le_bytes());
            out.extend_from_slice(payload);
            out
        }

        fn dummy_payload(len: usize, seed: u8) -> Vec<u8> {
            (0..len)
                .map(|i| {
                    seed.wrapping_add((i as u8).wrapping_mul(31))
                        .wrapping_add((i >> 8) as u8)
                })
                .collect()
        }

        // Deterministic salt/verifier for reproducibility.
        let salt: [u8; 16] = core::array::from_fn(|i| 0x22u8.wrapping_add(i as u8));
        let verifier_plain: [u8; 16] = core::array::from_fn(|i| 0xB0u8.wrapping_add(i as u8));
        let verifier_hash_plain = sha1_bytes(&[&verifier_plain]);

        // Encrypt verifier using the WinCrypt padded-16-byte representation for 40-bit keys.
        let key_material =
            derive_key_material_legacy(CryptoApiHashAlg::Sha1, password, &salt).expect("kdf");
        let key0_padded =
            derive_block_key(CryptoApiHashAlg::Sha1, key_material.as_slice(), 0, 5, true);
        let mut rc4 = Rc4::new(&key0_padded[..]);
        drop(key0_padded);
        let mut verifier_buf = [0u8; 36];
        verifier_buf[..16].copy_from_slice(&verifier_plain);
        verifier_buf[16..].copy_from_slice(&verifier_hash_plain);
        rc4.apply_keystream(&mut verifier_buf);
        let encrypted_verifier: [u8; 16] = verifier_buf[..16].try_into().unwrap();
        let encrypted_verifier_hash: [u8; 20] = verifier_buf[16..].try_into().unwrap();

        // EncryptionHeader (32 bytes, no CSP name).
        let mut enc_header = Vec::<u8>::new();
        enc_header.extend_from_slice(&0u32.to_le_bytes()); // Flags
        enc_header.extend_from_slice(&0u32.to_le_bytes()); // SizeExtra
        enc_header.extend_from_slice(&CALG_RC4.to_le_bytes()); // AlgID
        enc_header.extend_from_slice(&CALG_SHA1.to_le_bytes()); // AlgIDHash
        enc_header.extend_from_slice(&0u32.to_le_bytes()); // KeySize bits (0 => 40-bit)
        enc_header.extend_from_slice(&0u32.to_le_bytes()); // ProviderType
        enc_header.extend_from_slice(&0u32.to_le_bytes()); // Reserved1
        enc_header.extend_from_slice(&0u32.to_le_bytes()); // Reserved2

        // EncryptionVerifier.
        let mut enc_verifier = Vec::<u8>::new();
        enc_verifier.extend_from_slice(&(salt.len() as u32).to_le_bytes());
        enc_verifier.extend_from_slice(&salt);
        enc_verifier.extend_from_slice(&encrypted_verifier);
        enc_verifier.extend_from_slice(&(encrypted_verifier_hash.len() as u32).to_le_bytes());
        enc_verifier.extend_from_slice(&encrypted_verifier_hash);

        // FILEPASS payload (legacy layout B).
        let mut filepass_payload = Vec::<u8>::new();
        filepass_payload.extend_from_slice(&super::super::BIFF8_ENCRYPTION_TYPE_RC4.to_le_bytes()); // wEncryptionType
        filepass_payload
            .extend_from_slice(&super::super::BIFF8_RC4_ENCRYPTION_INFO_CRYPTOAPI_LEGACY.to_le_bytes()); // wEncryptionInfo (0x0004)
        filepass_payload.extend_from_slice(&4u16.to_le_bytes()); // vMajor
        filepass_payload.extend_from_slice(&2u16.to_le_bytes()); // vMinor
        filepass_payload.extend_from_slice(&0u16.to_le_bytes()); // reserved
        filepass_payload.extend_from_slice(&(enc_header.len() as u32).to_le_bytes()); // headerSize
        filepass_payload.extend_from_slice(&enc_header);
        filepass_payload.extend_from_slice(&enc_verifier);

        // Build plaintext workbook stream. Include:
        // - INTERFACEHDR (never encrypted)
        // - BOUNDSHEET (lbPlyPos plaintext)
        // - Dummy record that crosses the 1024-byte rekey boundary.
        let interface_hdr = record(
            super::super::RECORD_INTERFACEHDR,
            &dummy_payload(32, 0x33),
        );
        let boundsheet_payload = {
            let mut payload = Vec::<u8>::new();
            payload.extend_from_slice(&0x11223344u32.to_le_bytes()); // lbPlyPos (plaintext)
            payload.extend_from_slice(&dummy_payload(20, 0x44)); // encrypted remainder
            payload
        };
        let boundsheet = record(super::super::RECORD_BOUNDSHEET, &boundsheet_payload);
        let r1 = record(0x00FC, &dummy_payload(1500, 0x55));

        let bof = record(records::RECORD_BOF_BIFF8, &bof_payload);
        let filepass = record(records::RECORD_FILEPASS, &filepass_payload);
        let plain = [
            bof,
            filepass,
            interface_hdr,
            boundsheet,
            r1,
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        // Encrypt record payload bytes after FILEPASS using the legacy absolute-position stream
        // mapping.
        let encrypted_start = filepass_payload
            .len()
            .checked_add(4)
            .and_then(|l| l.checked_add(bof_payload.len() + 4))
            .unwrap();
        let mut encrypted = plain.clone();

        let mut offset = encrypted_start;
        let mut stream_pos = encrypted_start;
        while offset < encrypted.len() {
            let remaining = encrypted.len().checked_sub(offset).unwrap_or(0);
            if remaining < 4 {
                break;
            }

            let Some(header) = encrypted.get(offset..).and_then(|rest| rest.get(..4)) else {
                break;
            };
            let record_id = u16::from_le_bytes([header[0], header[1]]);
            let len = u16::from_le_bytes([header[2], header[3]]) as usize;
            let data_start = match offset.checked_add(4) {
                Some(v) => v,
                None => break,
            };
            let data_end = match data_start.checked_add(len) {
                Some(v) => v,
                None => break,
            };

            // Record headers are not encrypted but still advance the CryptoAPI RC4 stream position.
            stream_pos += 4;

            if !is_never_encrypted_record(record_id) && len > 0 {
                if record_id == super::super::RECORD_BOUNDSHEET {
                    if len > 4 {
                        if let Some(range_start) = data_start.checked_add(4) {
                            if let Some(payload) = encrypted.get_mut(range_start..data_end) {
                                decrypt_range_by_offset(
                                    payload,
                                    stream_pos + 4,
                                    CryptoApiHashAlg::Sha1,
                                    key_material.as_slice(),
                                    5,
                                    true,
                                );
                            } else {
                                break;
                            }
                        } else {
                            break;
                        }
                    }
                } else {
                    if let Some(payload) = encrypted.get_mut(data_start..data_end) {
                        decrypt_range_by_offset(
                            payload,
                            stream_pos,
                            CryptoApiHashAlg::Sha1,
                            key_material.as_slice(),
                            5,
                            true,
                        );
                    } else {
                        break;
                    }
                }
            }

            stream_pos += len;
            offset = data_end;
        }

        crate::biff::encryption::decrypt_workbook_stream(&mut encrypted, password)
            .expect("decrypt");
        assert_eq!(encrypted, plain);
    }
}
