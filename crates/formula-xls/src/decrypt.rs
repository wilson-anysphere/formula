use sha1::{Digest as _, Sha1};
use thiserror::Error;
use zeroize::Zeroizing;

use crate::biff::encryption::rc4::Rc4;
use crate::ct::ct_eq;

/// Errors returned while decrypting password-protected `.xls` BIFF8 workbooks.
#[derive(Debug, Error, Clone, PartialEq, Eq)]
pub(crate) enum DecryptError {
    #[error("unsupported encryption scheme: {0}")]
    UnsupportedEncryption(String),
    #[error("wrong password")]
    WrongPassword,
    #[error("invalid encryption info: {0}")]
    InvalidFormat(String),
}

const RECORD_FILEPASS: u16 = 0x002F;
// BIFF record id reserved for "unknown" sanitization. Any value that calamine doesn't treat as a
// special record is fine; we use 0xFFFF which is not a defined BIFF record id.
const RECORD_MASKED: u16 = 0xFFFF;

// FILEPASS.wEncryptionType values [MS-XLS 2.4.105].
const ENCRYPTION_TYPE_XOR: u16 = 0x0000;
const ENCRYPTION_TYPE_RC4: u16 = 0x0001;
// FILEPASS.wEncryptionSubType values for `wEncryptionType == 0x0001`.
const ENCRYPTION_SUBTYPE_STANDARD: u16 = 0x0001;
const ENCRYPTION_SUBTYPE_CRYPTOAPI: u16 = 0x0002;
// Some BIFF8 RC4 CryptoAPI workbooks use an older FILEPASS layout where the second field is
// `wEncryptionInfo == 0x0004` and the CryptoAPI EncryptionHeader/EncryptionVerifier structures are
// embedded directly in the FILEPASS payload (rather than using a length-prefixed EncryptionInfo
// blob).
const ENCRYPTION_INFO_CRYPTOAPI_LEGACY: u16 = 0x0004;

// CryptoAPI algorithm identifiers [MS-OFFCRYPTO] / WinCrypt.h.
const CALG_RC4: u32 = 0x0000_6801;
const CALG_SHA1: u32 = 0x0000_8004;

const PAYLOAD_BLOCK_SIZE: usize = 1024;
const PASSWORD_HASH_ITERATIONS: u32 = 50_000;

fn map_biff_decrypt_error(err: crate::biff::encryption::DecryptError) -> DecryptError {
    match err {
        crate::biff::encryption::DecryptError::WrongPassword => DecryptError::WrongPassword,
        crate::biff::encryption::DecryptError::UnsupportedEncryption(scheme) => {
            DecryptError::UnsupportedEncryption(scheme)
        }
        crate::biff::encryption::DecryptError::InvalidFilePass(message) => {
            DecryptError::InvalidFormat(message)
        }
        crate::biff::encryption::DecryptError::NoFilePass => {
            DecryptError::InvalidFormat("missing FILEPASS record".to_string())
        }
        crate::biff::encryption::DecryptError::PasswordRequired => DecryptError::WrongPassword,
    }
}

/// Decrypt an in-memory BIFF workbook stream for any supported `FILEPASS` scheme.
///
/// The returned workbook stream has the `FILEPASS` record id *masked* (replaced with `0xFFFF`) so
/// downstream BIFF parsers (and `calamine`) treat the stream as plaintext without shifting any
/// record offsets (e.g. `BoundSheet8.lbPlyPos`).
pub(crate) fn decrypt_biff_workbook_stream(
    workbook_stream: &[u8],
    password: &str,
) -> Result<Vec<u8>, DecryptError> {
    let biff_version = crate::biff::detect_biff_version(workbook_stream);
    if biff_version != crate::biff::BiffVersion::Biff8 {
        // Legacy BIFF5 XOR (and any future non-BIFF8 schemes) are handled by the BIFF decryptor.
        let mut out = workbook_stream.to_vec();
        crate::biff::encryption::decrypt_workbook_stream(&mut out, password)
            .map_err(map_biff_decrypt_error)?;
        crate::biff::records::mask_workbook_globals_filepass_record_id_in_place(&mut out);
        return Ok(out);
    }

    // For BIFF8, dispatch CryptoAPI RC4 decryption to the existing implementation (it returns a
    // new Vec and masks FILEPASS), but use the BIFF decryptor for XOR + legacy RC4.
    let (filepass_offset, filepass_len) = find_filepass_record_offset(workbook_stream)?;
    let filepass_data_start = filepass_offset + 4;
    let filepass_data_end = filepass_data_start
        .checked_add(filepass_len)
        .ok_or_else(|| DecryptError::InvalidFormat("FILEPASS length overflow".to_string()))?;
    let filepass_payload = workbook_stream
        .get(filepass_data_start..filepass_data_end)
        .ok_or_else(|| DecryptError::InvalidFormat("FILEPASS payload out of bounds".to_string()))?;

    // FILEPASS payload begins with `wEncryptionType` (u16). For BIFF8 RC4, the next u16 is
    // `wEncryptionSubType` (0x0001=Standard, 0x0002=CryptoAPI).
    let encryption_type = read_u16_le(filepass_payload, 0).ok_or_else(|| {
        DecryptError::InvalidFormat("FILEPASS missing wEncryptionType".to_string())
    })?;

    if encryption_type == ENCRYPTION_TYPE_RC4 {
        // For BIFF8 RC4, the second u16 is either:
        // - `wEncryptionSubType` (0x0001=Standard, 0x0002=CryptoAPI), or
        // - legacy `wEncryptionInfo` (0x0004) for older CryptoAPI RC4 layouts.
        let second_field = read_u16_le(filepass_payload, 2).ok_or_else(|| {
            DecryptError::InvalidFormat("FILEPASS missing subtype/encryption info".to_string())
        })?;
        if matches!(second_field, ENCRYPTION_SUBTYPE_CRYPTOAPI | ENCRYPTION_INFO_CRYPTOAPI_LEGACY) {
            return decrypt_biff8_workbook_stream_rc4_cryptoapi(workbook_stream, password);
        }
        if second_field != ENCRYPTION_SUBTYPE_STANDARD {
            return Err(DecryptError::UnsupportedEncryption(format!(
                "FILEPASS RC4 wEncryptionSubType=0x{second_field:04X}"
            )));
        }
    } else if encryption_type != ENCRYPTION_TYPE_XOR {
        return Err(DecryptError::UnsupportedEncryption(format!(
            "FILEPASS wEncryptionType=0x{encryption_type:04X}"
        )));
    }

    let mut out = workbook_stream.to_vec();
    crate::biff::encryption::decrypt_workbook_stream(&mut out, password)
        .map_err(map_biff_decrypt_error)?;
    crate::biff::records::mask_workbook_globals_filepass_record_id_in_place(&mut out);
    Ok(out)
}

#[derive(Debug, Clone)]
struct EncryptionHeader {
    alg_id: u32,
    alg_id_hash: u32,
    key_size_bits: u32,
    #[allow(dead_code)]
    provider_type: u32,
    #[allow(dead_code)]
    csp_name: Option<String>,
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

fn read_u16_le(bytes: &[u8], offset: usize) -> Option<u16> {
    let b = bytes.get(offset..offset + 2)?;
    Some(u16::from_le_bytes([b[0], b[1]]))
}

fn read_u32_le(bytes: &[u8], offset: usize) -> Option<u32> {
    let b = bytes.get(offset..offset + 4)?;
    Some(u32::from_le_bytes([b[0], b[1], b[2], b[3]]))
}

fn utf16le_bytes(s: &str) -> Zeroizing<Vec<u8>> {
    let mut out = Vec::with_capacity(s.len().saturating_mul(2));
    for unit in s.encode_utf16() {
        out.extend_from_slice(&unit.to_le_bytes());
    }
    Zeroizing::new(out)
}

fn sha1_bytes(chunks: &[&[u8]]) -> [u8; 20] {
    let mut hasher = Sha1::new();
    for chunk in chunks {
        hasher.update(chunk);
    }
    let digest = hasher.finalize();
    let mut out = [0u8; 20];
    out.copy_from_slice(&digest);
    out
}

fn derive_key_material_legacy(password: &str, salt: &[u8]) -> Result<[u8; 20], DecryptError> {
    // Legacy BIFF8 RC4 CryptoAPI key derivation (FILEPASS wEncryptionInfo=0x0004):
    //
    //   key_material = SHA1(salt + password_utf16le)
    //
    // This intentionally does *not* apply the 50,000-iteration hashing step used by other
    // CryptoAPI encodings.
    if salt.len() != 16 {
        return Err(DecryptError::InvalidFormat(format!(
            "CryptoAPI legacy salt length {} (expected 16)",
            salt.len()
        )));
    }
    let pw_bytes = utf16le_bytes(password);
    Ok(sha1_bytes(&[salt, &pw_bytes]))
}

fn parse_encryption_header(bytes: &[u8]) -> Result<EncryptionHeader, DecryptError> {
    // Fixed-length header fields are 8 DWORDs.
    if bytes.len() < 32 {
        return Err(DecryptError::InvalidFormat(format!(
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
    let alg_id = read_u32_le(bytes, 8).ok_or_else(|| {
        DecryptError::InvalidFormat("EncryptionHeader missing AlgID".to_string())
    })?;
    let alg_id_hash = read_u32_le(bytes, 12).ok_or_else(|| {
        DecryptError::InvalidFormat("EncryptionHeader missing AlgIDHash".to_string())
    })?;
    let key_size_bits = read_u32_le(bytes, 16).ok_or_else(|| {
        DecryptError::InvalidFormat("EncryptionHeader missing KeySize".to_string())
    })?;
    let provider_type = read_u32_le(bytes, 20).ok_or_else(|| {
        DecryptError::InvalidFormat("EncryptionHeader missing ProviderType".to_string())
    })?;

    let csp_bytes = &bytes[32..];
    let csp_name = if csp_bytes.is_empty() {
        None
    } else {
        // CSPName is a null-terminated UTF-16LE string.
        let even_len = csp_bytes.len().saturating_sub(csp_bytes.len() % 2);
        let mut units: Vec<u16> = Vec::with_capacity(even_len / 2);
        for chunk in csp_bytes[..even_len].chunks_exact(2) {
            let unit = u16::from_le_bytes([chunk[0], chunk[1]]);
            if unit == 0 {
                break;
            }
            units.push(unit);
        }
        Some(String::from_utf16_lossy(&units))
    };

    Ok(EncryptionHeader {
        alg_id,
        alg_id_hash,
        key_size_bits,
        provider_type,
        csp_name,
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
        return Err(DecryptError::InvalidFormat(
            "EncryptionVerifier truncated".to_string(),
        ));
    }

    let salt_size = read_u32_le(bytes, 0).ok_or_else(|| {
        DecryptError::InvalidFormat("EncryptionVerifier missing SaltSize".to_string())
    })? as usize;

    let salt_start = 4usize;
    let salt_end = salt_start
        .checked_add(salt_size)
        .ok_or_else(|| DecryptError::InvalidFormat("SaltSize overflow".to_string()))?;
    let verifier_start = salt_end;
    let verifier_end = verifier_start.checked_add(16).ok_or_else(|| {
        DecryptError::InvalidFormat("EncryptedVerifier offset overflow".to_string())
    })?;
    let hash_size_start = verifier_end;
    let hash_size_end = hash_size_start.checked_add(4).ok_or_else(|| {
        DecryptError::InvalidFormat("VerifierHashSize offset overflow".to_string())
    })?;

    if hash_size_end > bytes.len() {
        return Err(DecryptError::InvalidFormat(format!(
            "EncryptionVerifier truncated (len={}, need={hash_size_end})",
            bytes.len()
        )));
    }

    let salt = bytes[salt_start..salt_end].to_vec();
    let mut encrypted_verifier = [0u8; 16];
    encrypted_verifier.copy_from_slice(&bytes[verifier_start..verifier_end]);
    let verifier_hash_size =
        read_u32_le(bytes, hash_size_start).ok_or_else(|| {
            DecryptError::InvalidFormat("EncryptionVerifier missing VerifierHashSize".to_string())
        })?;
    let encrypted_hash_start = hash_size_end;
    let encrypted_hash_end = encrypted_hash_start
        .checked_add(verifier_hash_size as usize)
        .ok_or_else(|| DecryptError::InvalidFormat("VerifierHashSize overflow".to_string()))?;
    if encrypted_hash_end > bytes.len() {
        return Err(DecryptError::InvalidFormat(format!(
            "EncryptionVerifierHash truncated (len={}, need={encrypted_hash_end})",
            bytes.len()
        )));
    }

    let encrypted_verifier_hash = bytes[encrypted_hash_start..encrypted_hash_end].to_vec();

    Ok(EncryptionVerifier {
        salt,
        encrypted_verifier,
        verifier_hash_size,
        encrypted_verifier_hash,
    })
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
    // - EncryptionHeader (headerSize bytes, includes CSPName)
    // - EncryptionVerifier (remaining bytes)
    if payload.len() < 14 {
        return Err(DecryptError::InvalidFormat(format!(
            "FILEPASS payload truncated (len={})",
            payload.len()
        )));
    }

    let encryption_type =
        read_u16_le(payload, 0).ok_or_else(|| {
            DecryptError::InvalidFormat("FILEPASS missing wEncryptionType".to_string())
        })?;
    if encryption_type != ENCRYPTION_TYPE_RC4 {
        return Err(DecryptError::InvalidFormat(format!(
            "FILEPASS wEncryptionType=0x{encryption_type:04X} (expected 0x0001 for RC4)"
        )));
    }

    let encryption_info =
        read_u16_le(payload, 2).ok_or_else(|| {
            DecryptError::InvalidFormat("FILEPASS missing wEncryptionInfo".to_string())
        })?;
    if encryption_info != ENCRYPTION_INFO_CRYPTOAPI_LEGACY {
        return Err(DecryptError::InvalidFormat(format!(
            "FILEPASS RC4 wEncryptionInfo=0x{encryption_info:04X} (expected 0x0004)"
        )));
    }

    let header_size =
        read_u32_le(payload, 10).ok_or_else(|| {
            DecryptError::InvalidFormat("FILEPASS missing headerSize".to_string())
        })? as usize;

    let header_start = 14usize;
    let header_end = header_start
        .checked_add(header_size)
        .ok_or_else(|| DecryptError::InvalidFormat("headerSize overflow".to_string()))?;
    if header_end > payload.len() {
        return Err(DecryptError::InvalidFormat(format!(
            "FILEPASS header out of bounds (payload_len={}, header_end={header_end})",
            payload.len()
        )));
    }

    let header = parse_encryption_header(&payload[header_start..header_end])?;
    let verifier = parse_encryption_verifier(&payload[header_end..])?;
    Ok(CryptoApiEncryptionInfo { header, verifier })
}

fn parse_cryptoapi_encryption_info(bytes: &[u8]) -> Result<CryptoApiEncryptionInfo, DecryptError> {
    // EncryptionInfo [MS-OFFCRYPTO] 2.3.1:
    //   EncryptionVersionInfo (Major, Minor) 4 bytes
    //   DWORD Flags;
    //   DWORD HeaderSize;
    //   EncryptionHeader (HeaderSize bytes)
    //   EncryptionVerifier (remaining bytes)
    if bytes.len() < 12 {
        return Err(DecryptError::InvalidFormat(format!(
            "EncryptionInfo truncated (len={})",
            bytes.len()
        )));
    }

    let header_size = read_u32_le(bytes, 8).ok_or_else(|| {
        DecryptError::InvalidFormat("EncryptionInfo missing HeaderSize".to_string())
    })? as usize;

    let header_start = 12usize;
    let header_end = header_start
        .checked_add(header_size)
        .ok_or_else(|| DecryptError::InvalidFormat("HeaderSize overflow".to_string()))?;
    if header_end > bytes.len() {
        return Err(DecryptError::InvalidFormat(format!(
            "EncryptionInfo header out of bounds (len={}, header_end={header_end})",
            bytes.len()
        )));
    }

    let header = parse_encryption_header(&bytes[header_start..header_end])?;
    let verifier = parse_encryption_verifier(&bytes[header_end..])?;

    Ok(CryptoApiEncryptionInfo { header, verifier })
}

fn derive_key_material(password: &str, salt: &[u8]) -> [u8; 20] {
    // CryptoAPI password hashing [MS-OFFCRYPTO]:
    //   H0 = SHA1(salt + UTF16LE(password))
    //   for i in 0..49999: H0 = SHA1(i_le32 + H0)
    let pw_bytes = utf16le_bytes(password);
    let mut hash: Zeroizing<[u8; 20]> = Zeroizing::new(sha1_bytes(&[salt, pw_bytes.as_slice()]));

    for i in 0..PASSWORD_HASH_ITERATIONS {
        let iter = i.to_le_bytes();
        *hash = sha1_bytes(&[&iter, &hash[..]]);
    }

    *hash
}

fn derive_block_key(key_material: &[u8; 20], block: u32, key_len: usize) -> Vec<u8> {
    let block_bytes = block.to_le_bytes();
    let digest = sha1_bytes(&[key_material, &block_bytes]);

    // Office/WinCrypt quirk: 40-bit RC4 keys are expressed as a 128-bit (16-byte) key where the
    // low 40 bits are set and the remaining 88 bits are zero. Using the raw 5-byte key changes the
    // RC4 key-scheduling algorithm (KSA) and yields the wrong keystream.
    //
    // [MS-OFFCRYPTO] calls this out for CryptoAPI RC4; Excel uses the same convention for BIFF8
    // `FILEPASS` CryptoAPI encryption.
    if key_len == 5 {
        let mut key = Vec::with_capacity(16);
        key.extend_from_slice(&digest[..5]);
        key.resize(16, 0);
        key
    } else {
        digest[..key_len].to_vec()
    }
}

fn rc4_discard(rc4: &mut Rc4, mut n: usize) {
    // Advance the internal RC4 state without caring about the output bytes. This is used by the
    // absolute-offset BIFF8 CryptoAPI RC4 variant to jump to `pos_in_block` within a 1024-byte
    // rekey segment.
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
    key_material: &[u8; 20],
    key_len: usize,
) {
    let mut stream_pos = start_offset;
    let mut remaining = bytes.len();
    let mut pos = 0usize;
    while remaining > 0 {
        let block = (stream_pos / PAYLOAD_BLOCK_SIZE) as u32;
        let in_block = stream_pos % PAYLOAD_BLOCK_SIZE;
        let take = remaining.min(PAYLOAD_BLOCK_SIZE - in_block);

        let key = derive_block_key(key_material, block, key_len);
        let mut rc4 = Rc4::new(&key);
        rc4_discard(&mut rc4, in_block);
        rc4.apply_keystream(&mut bytes[pos..pos + take]);

        stream_pos += take;
        pos += take;
        remaining -= take;
    }
}

fn is_never_encrypted_record(record_id: u16) -> bool {
    // Mirror Apache POI's BIFF8 RC4 CryptoAPI behaviour:
    // - BOF
    // - FILEPASS
    // - INTERFACEHDR (0x00E1)
    matches!(record_id, 0x0809 | 0x0009 | RECORD_FILEPASS | 0x00E1)
}

#[cfg(test)]
mod filepass_tests {
    use super::*;
    use crate::ct::{ct_eq_call_count, reset_ct_eq_calls};
    use formula_model::{CellRef, VerticalAlignment};
    use std::path::PathBuf;

    fn record(record_id: u16, payload: &[u8]) -> Vec<u8> {
        let mut out = Vec::with_capacity(4 + payload.len());
        out.extend_from_slice(&record_id.to_le_bytes());
        out.extend_from_slice(&(payload.len() as u16).to_le_bytes());
        out.extend_from_slice(payload);
        out
    }

    /// Spec-correct per-block RC4 key derivation for CryptoAPI RC4 (BIFF8 FILEPASS subtype 0x0002).
    ///
    /// This intentionally does **not** call `derive_block_key` so tests catch regressions in the
    /// 40-bit padding behaviour.
    fn derive_block_key_spec(key_material: &[u8; 20], block: u32, key_size_bits: u32) -> Vec<u8> {
        let block_bytes = block.to_le_bytes();
        let digest = sha1_bytes(&[key_material, &block_bytes]);
        let key_len = (key_size_bits / 8) as usize;
        if key_size_bits == 40 {
            let mut key = Vec::with_capacity(16);
            key.extend_from_slice(&digest[..5]);
            key.resize(16, 0);
            return key;
        }
        digest[..key_len].to_vec()
    }

    struct PayloadRc4Spec {
        key_material: [u8; 20],
        key_size_bits: u32,
        block: u32,
        pos_in_block: usize,
        rc4: Rc4,
    }

    impl PayloadRc4Spec {
        fn new(key_material: [u8; 20], key_size_bits: u32) -> Self {
            let key = derive_block_key_spec(&key_material, 0, key_size_bits);
            let rc4 = Rc4::new(&key);
            Self {
                key_material,
                key_size_bits,
                block: 0,
                pos_in_block: 0,
                rc4,
            }
        }

        fn rekey(&mut self) {
            self.block = self.block.wrapping_add(1);
            let key = derive_block_key_spec(&self.key_material, self.block, self.key_size_bits);
            self.rc4 = Rc4::new(&key);
            self.pos_in_block = 0;
        }

        fn apply_keystream(&mut self, mut data: &mut [u8]) {
            while !data.is_empty() {
                if self.pos_in_block == PAYLOAD_BLOCK_SIZE {
                    self.rekey();
                }
                let remaining_in_block = PAYLOAD_BLOCK_SIZE.saturating_sub(self.pos_in_block);
                let chunk_len = data.len().min(remaining_in_block);
                let (chunk, rest) = data.split_at_mut(chunk_len);
                self.rc4.apply_keystream(chunk);
                self.pos_in_block += chunk_len;
                data = rest;
            }
        }
    }

    #[test]
    fn derive_block_key_pads_40_bit_rc4_to_16_bytes() {
        let key_material = [0x11u8; 20];
        let block = 0u32;

        let block_bytes = block.to_le_bytes();
        let digest = sha1_bytes(&[&key_material, &block_bytes]);
        let mut expected = Vec::from(&digest[..5]);
        expected.resize(16, 0);

        let got = derive_block_key(&key_material, block, 5);
        assert_eq!(got, expected);
        assert_eq!(got.len(), 16);
        assert!(got[5..].iter().all(|b| *b == 0));
    }

    #[test]
    fn cryptoapi_kdf_matches_known_vectors() {
        // Test vectors mirrored from `biff::encryption::cryptoapi` (Apache POI-compatible).
        let password = "SecretPassword";
        let salt: [u8; 16] = [
            0xA0, 0xA1, 0xA2, 0xA3, 0xA4, 0xA5, 0xA6, 0xA7, 0xA8, 0xA9, 0xAA, 0xAB, 0xAC,
            0xAD, 0xAE, 0xAF,
        ];
        let key_material = derive_key_material(password, &salt);

        let expected_key_block0: [u8; 16] = [
            0x3D, 0x7D, 0x0B, 0x04, 0x2E, 0xCF, 0x02, 0xA7, 0xBC, 0xE1, 0x26, 0xA1, 0xE2,
            0x20, 0xE2, 0xC8,
        ];
        let expected_key_block1: [u8; 16] = [
            0xF8, 0x06, 0xD7, 0x4E, 0x99, 0x94, 0x8E, 0xE8, 0xD3, 0x68, 0xD6, 0x1C, 0xEA,
            0xAA, 0x7C, 0x36,
        ];

        assert_eq!(
            derive_block_key(&key_material, 0, 16),
            expected_key_block0
        );
        assert_eq!(
            derive_block_key(&key_material, 1, 16),
            expected_key_block1
        );
    }

    #[test]
    fn decrypts_rc4_cryptoapi_40_bit_by_using_padded_rc4_key() {
        // Build a minimal BIFF8 workbook stream:
        // BOF (plaintext) + FILEPASS (plaintext) + one record with encrypted payload + EOF.
        const RECORD_BOF: u16 = 0x0809;
        const RECORD_EOF: u16 = 0x000A;

        let password = "password";
        let key_size_bits: u32 = 40;
        let salt: [u8; 16] = [
            0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D,
            0x0E, 0x0F, 0x10,
        ];

        // Derive key material per MS-OFFCRYPTO (SHA1).
        let key_material = derive_key_material(password, &salt);

        // Build the verifier fields (encrypted with block 0 key).
        let verifier_plain: [u8; 16] = [
            0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C,
            0x0D, 0x0E, 0x0F,
        ];
        let verifier_hash_plain: [u8; 20] = sha1_bytes(&[&verifier_plain]);

        let key0 = derive_block_key_spec(&key_material, 0, key_size_bits);
        assert_eq!(key0.len(), 16, "40-bit RC4 key must be padded to 16 bytes");

        let mut rc4 = Rc4::new(&key0);
        let mut encrypted_verifier = verifier_plain;
        rc4.apply_keystream(&mut encrypted_verifier);
        let mut encrypted_verifier_hash = verifier_hash_plain.to_vec();
        rc4.apply_keystream(&mut encrypted_verifier_hash);

        // Build CryptoAPI EncryptionInfo (minimal, SHA1 + RC4).
        let mut enc_header = Vec::new();
        enc_header.extend_from_slice(&0u32.to_le_bytes()); // flags
        enc_header.extend_from_slice(&0u32.to_le_bytes()); // sizeExtra
        enc_header.extend_from_slice(&CALG_RC4.to_le_bytes()); // algId
        enc_header.extend_from_slice(&CALG_SHA1.to_le_bytes()); // algIdHash
        enc_header.extend_from_slice(&key_size_bits.to_le_bytes()); // keySize (bits)
        enc_header.extend_from_slice(&0u32.to_le_bytes()); // providerType
        enc_header.extend_from_slice(&0u32.to_le_bytes()); // reserved1
        enc_header.extend_from_slice(&0u32.to_le_bytes()); // reserved2

        let mut enc_info = Vec::new();
        enc_info.extend_from_slice(&4u16.to_le_bytes()); // majorVersion (ignored by parser)
        enc_info.extend_from_slice(&2u16.to_le_bytes()); // minorVersion (ignored by parser)
        enc_info.extend_from_slice(&0u32.to_le_bytes()); // flags
        enc_info.extend_from_slice(&(enc_header.len() as u32).to_le_bytes()); // headerSize
        enc_info.extend_from_slice(&enc_header);
        // EncryptionVerifier
        enc_info.extend_from_slice(&(salt.len() as u32).to_le_bytes());
        enc_info.extend_from_slice(&salt);
        enc_info.extend_from_slice(&encrypted_verifier);
        enc_info.extend_from_slice(&20u32.to_le_bytes()); // verifierHashSize
        enc_info.extend_from_slice(&encrypted_verifier_hash);

        let mut filepass_payload = Vec::new();
        filepass_payload.extend_from_slice(&ENCRYPTION_TYPE_RC4.to_le_bytes());
        filepass_payload.extend_from_slice(&ENCRYPTION_SUBTYPE_CRYPTOAPI.to_le_bytes());
        filepass_payload.extend_from_slice(&(enc_info.len() as u32).to_le_bytes());
        filepass_payload.extend_from_slice(&enc_info);

        // Plaintext record payload after FILEPASS. Make it >1024 bytes to ensure the decryptor
        // rekeys (block 1 derivation must also follow the 40-bit padding rule).
        let mut plaintext_payload = vec![0u8; 2048];
        for (i, b) in plaintext_payload.iter_mut().enumerate() {
            *b = (i % 251) as u8;
        }
        const RECORD_DUMMY: u16 = 0x1234;

        let bof_payload = [0u8; 16];
        let plaintext_stream = [
            record(RECORD_BOF, &bof_payload),
            record(RECORD_FILEPASS, &filepass_payload),
            record(RECORD_DUMMY, &plaintext_payload),
            record(RECORD_EOF, &[]),
        ]
        .concat();

        // Encrypt record payloads after FILEPASS using the spec-correct RC4 key derivation.
        let (filepass_offset, filepass_len) =
            find_filepass_record_offset(&plaintext_stream).expect("FILEPASS offset");
        let filepass_data_end = filepass_offset + 4 + filepass_len;

        let mut encrypted_stream = plaintext_stream.clone();
        let mut payload_cipher = PayloadRc4Spec::new(key_material, key_size_bits);

        let mut offset = filepass_data_end;
        while offset < encrypted_stream.len() {
            let len = u16::from_le_bytes([encrypted_stream[offset + 2], encrypted_stream[offset + 3]])
                as usize;
            let data_start = offset + 4;
            let data_end = data_start + len;
            payload_cipher.apply_keystream(&mut encrypted_stream[data_start..data_end]);
            offset = data_end;
        }

        // Decrypt using the implementation under test.
        let decrypted =
            decrypt_biff8_workbook_stream_rc4_cryptoapi(&encrypted_stream, password).expect("decrypt");

        // The decryptor masks the FILEPASS record id but otherwise yields the original plaintext.
        let mut expected = plaintext_stream;
        expected[filepass_offset..filepass_offset + 2]
            .copy_from_slice(&RECORD_MASKED.to_le_bytes());
        assert_eq!(decrypted, expected);
    }

    #[test]
    fn decrypts_real_cryptoapi_fixture_and_preserves_workbook_globals_structure() {
        // Regression guard for `.xls` files where workbook-global records after FILEPASS
        // (XF/FONT/etc) must be decrypted so downstream BIFF parsers can import styles and other
        // metadata.
        //
        // The fixture's Sheet1!A1 uses a non-default XF (vertical alignment = Top) so we can assert
        // that:
        // - workbook-globals (XF table) survived decryption,
        // - style resolution can still see the non-default alignment, and
        // - the XF remains "interesting" so the importer retains it.
        const PASSWORD: &str = "correct horse battery staple";

        let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("tests")
            .join("fixtures")
            .join("encrypted")
            .join("biff8_rc4_cryptoapi_pw_open.xls");

        let workbook_stream =
            crate::biff::read_workbook_stream_from_xls(&path).expect("read Workbook stream");
        let decrypted = decrypt_biff8_workbook_stream_rc4_cryptoapi(&workbook_stream, PASSWORD)
            .expect("decrypt");

        // The decryptor masks FILEPASS so parsers do not treat it as an encryption terminator.
        assert!(
            !crate::biff::records::workbook_globals_has_filepass_record(&decrypted),
            "expected FILEPASS record to be masked after decryption"
        );

        // Ensure the decrypted workbook globals still contain the expected record types.
        let mut xf_count = 0usize;
        let mut font_count = 0usize;
        let mut boundsheet_count = 0usize;

        let mut iter = crate::biff::records::BiffRecordIter::from_offset(&decrypted, 0)
            .expect("record iter");
        while let Some(next) = iter.next() {
            let record = next.expect("record");
            match record.record_id {
                0x00E0 => xf_count += 1,         // XF
                0x0031 => font_count += 1,       // FONT
                0x0085 => boundsheet_count += 1, // BOUNDSHEET
                crate::biff::records::RECORD_EOF => break,
                _ => {}
            }
        }

        assert!(xf_count > 0, "expected at least one XF record after decryption");
        assert!(font_count > 0, "expected at least one FONT record after decryption");
        assert!(
            boundsheet_count > 0,
            "expected at least one BOUNDSHEET record after decryption"
        );

        // Sanity-check that BIFF global/style parsing sees at least one non-default cell style on
        // the fixture (Sheet1!A1 uses a non-default vertical alignment in the source workbook).
        let biff_version = crate::biff::detect_biff_version(&decrypted);
        let codepage = crate::biff::parse_biff_codepage(&decrypted);

        let globals = crate::biff::parse_biff_workbook_globals(&decrypted, biff_version, codepage)
            .expect("parse workbook globals");
        assert!(
            globals.xf_count() != 0,
            "expected XF records to be parsed from decrypted workbook globals"
        );

        let bound_sheets = crate::biff::parse_biff_bound_sheets(&decrypted, biff_version, codepage)
            .expect("parse bound sheets");
        assert!(!bound_sheets.is_empty(), "expected at least one bound sheet");
        let sheet0_offset = bound_sheets[0].offset;

        let cell_xfs =
            crate::biff::parse_biff_sheet_cell_xf_indices_filtered(&decrypted, sheet0_offset, None)
                .expect("parse cell xfs");
        let xf_idx = *cell_xfs
            .get(&CellRef::new(0, 0))
            .expect("expected A1 xf index in sheet stream");

        // Sanity check: verify the decrypted sheet payload contains the expected A1 value (42.0).
        {
            let mut found = None;
            for record in crate::biff::records::BestEffortSubstreamIter::from_offset(
                &decrypted,
                sheet0_offset,
            )
            .expect("sheet record iter")
            {
                // Stop before consuming the next substream.
                if record.offset != sheet0_offset
                    && crate::biff::records::is_bof_record(record.record_id)
                {
                    break;
                }
                // NUMBER [MS-XLS 2.4.164]
                if record.record_id == 0x0203 && record.data.len() >= 14 {
                    let row = u16::from_le_bytes([record.data[0], record.data[1]]);
                    let col = u16::from_le_bytes([record.data[2], record.data[3]]);
                    if row == 0 && col == 0 {
                        let mut buf = [0u8; 8];
                        buf.copy_from_slice(&record.data[6..14]);
                        found = Some(f64::from_le_bytes(buf));
                        break;
                    }
                }
            }
            assert_eq!(found, Some(42.0), "expected decrypted A1 NUMBER record");
        }

        // Style resolution should preserve non-default vertical alignment.
        let style = globals.resolve_style(xf_idx as u32);
        assert_ne!(style, formula_model::Style::default());
        assert_eq!(
            style
                .alignment
                .as_ref()
                .and_then(|alignment| alignment.vertical),
            Some(VerticalAlignment::Top),
            "expected A1 style to preserve vertical alignment from decrypted XF records"
        );

        // Inspect the raw decrypted XF record bytes backing the A1 style.
        {
            let mut seen = 0u16;
            let mut alignment_byte = None;
            let mut used_attr = None;
            for record in crate::biff::records::BestEffortSubstreamIter::from_offset(&decrypted, 0)
                .expect("globals record iter")
            {
                if record.record_id == 0x00E0 {
                    if seen == xf_idx {
                        alignment_byte = record.data.get(6).copied();
                        used_attr = record.data.get(9).copied();
                        break;
                    }
                    seen = seen.saturating_add(1);
                }
            }
            assert_eq!(
                alignment_byte,
                Some(0),
                "expected A1 XF alignment byte to encode vertical=Top"
            );
            assert_eq!(
                used_attr,
                Some(0x3F),
                "expected A1 XF to have apply-all attribute flags"
            );
        }

        let mask = globals.xf_is_interesting_mask();
        assert!(
            mask.get(xf_idx as usize).copied().unwrap_or(false),
            "expected A1 XF to be marked as interesting so style import retains it"
        );
    }

    #[test]
    fn verify_password_uses_constant_time_compare() {
        // Ensure the password verifier hash check uses our constant-time equality helper.
        let password = "password";
        let wrong_password = "wrong password";
        let key_size_bits: u32 = 40;
        let salt: [u8; 16] = [
            0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D,
            0x0E, 0x0F, 0x10,
        ];

        // Derive key material per MS-OFFCRYPTO (SHA1).
        let key_material = derive_key_material(password, &salt);

        // Build the verifier fields (encrypted with block 0 key).
        let verifier_plain: [u8; 16] = [
            0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C,
            0x0D, 0x0E, 0x0F,
        ];
        let verifier_hash_plain: [u8; 20] = sha1_bytes(&[&verifier_plain]);

        let key0 = derive_block_key_spec(&key_material, 0, key_size_bits);
        let mut rc4 = Rc4::new(&key0);
        let mut encrypted_verifier = verifier_plain;
        rc4.apply_keystream(&mut encrypted_verifier);
        let mut encrypted_verifier_hash = verifier_hash_plain.to_vec();
        rc4.apply_keystream(&mut encrypted_verifier_hash);

        // Build CryptoAPI EncryptionInfo (minimal, SHA1 + RC4).
        let mut enc_header = Vec::new();
        enc_header.extend_from_slice(&0u32.to_le_bytes()); // flags
        enc_header.extend_from_slice(&0u32.to_le_bytes()); // sizeExtra
        enc_header.extend_from_slice(&CALG_RC4.to_le_bytes()); // algId
        enc_header.extend_from_slice(&CALG_SHA1.to_le_bytes()); // algIdHash
        enc_header.extend_from_slice(&key_size_bits.to_le_bytes()); // keySize (bits)
        enc_header.extend_from_slice(&0u32.to_le_bytes()); // providerType
        enc_header.extend_from_slice(&0u32.to_le_bytes()); // reserved1
        enc_header.extend_from_slice(&0u32.to_le_bytes()); // reserved2

        let mut enc_info = Vec::new();
        enc_info.extend_from_slice(&4u16.to_le_bytes()); // majorVersion (ignored by parser)
        enc_info.extend_from_slice(&2u16.to_le_bytes()); // minorVersion (ignored by parser)
        enc_info.extend_from_slice(&0u32.to_le_bytes()); // flags
        enc_info.extend_from_slice(&(enc_header.len() as u32).to_le_bytes()); // headerSize
        enc_info.extend_from_slice(&enc_header);
        // EncryptionVerifier
        enc_info.extend_from_slice(&(salt.len() as u32).to_le_bytes());
        enc_info.extend_from_slice(&salt);
        enc_info.extend_from_slice(&encrypted_verifier);
        enc_info.extend_from_slice(&20u32.to_le_bytes()); // verifierHashSize
        enc_info.extend_from_slice(&encrypted_verifier_hash);

        let mut filepass_payload = Vec::new();
        filepass_payload.extend_from_slice(&ENCRYPTION_TYPE_RC4.to_le_bytes());
        filepass_payload.extend_from_slice(&ENCRYPTION_SUBTYPE_CRYPTOAPI.to_le_bytes());
        filepass_payload.extend_from_slice(&(enc_info.len() as u32).to_le_bytes());
        filepass_payload.extend_from_slice(&enc_info);

        let info =
            parse_filepass_record_payload(&filepass_payload).expect("parse FILEPASS payload");

        reset_ct_eq_calls();
        let err = verify_password(&info, wrong_password).expect_err("expected wrong password error");
        assert_eq!(err, DecryptError::WrongPassword);
        assert!(
            ct_eq_call_count() > 0,
            "expected constant-time compare helper to be invoked"
        );
    }
}

struct PayloadRc4 {
    key_material: Zeroizing<[u8; 20]>,
    key_len: usize,
    block: u32,
    pos_in_block: usize,
    rc4: Rc4,
}

impl PayloadRc4 {
    fn new(key_material: [u8; 20], key_len: usize) -> Self {
        let key_material: Zeroizing<[u8; 20]> = Zeroizing::new(key_material);
        let key: Zeroizing<Vec<u8>> = Zeroizing::new(derive_block_key(&key_material, 0, key_len));
        let rc4 = Rc4::new(key.as_slice());
        Self {
            key_material,
            key_len,
            block: 0,
            pos_in_block: 0,
            rc4,
        }
    }

    fn rekey(&mut self) {
        self.block = self.block.wrapping_add(1);
        let key: Zeroizing<Vec<u8>> =
            Zeroizing::new(derive_block_key(&self.key_material, self.block, self.key_len));
        self.rc4 = Rc4::new(key.as_slice());
        self.pos_in_block = 0;
    }

    fn apply_keystream(&mut self, mut data: &mut [u8]) {
        while !data.is_empty() {
            if self.pos_in_block == PAYLOAD_BLOCK_SIZE {
                self.rekey();
            }

            let remaining_in_block = PAYLOAD_BLOCK_SIZE.saturating_sub(self.pos_in_block);
            let chunk_len = data.len().min(remaining_in_block);
            let (chunk, rest) = data.split_at_mut(chunk_len);
            self.rc4.apply_keystream(chunk);
            self.pos_in_block += chunk_len;
            data = rest;
        }
    }
}

fn verify_password(info: &CryptoApiEncryptionInfo, password: &str) -> Result<[u8; 20], DecryptError> {
    if info.header.alg_id != CALG_RC4 || info.header.alg_id_hash != CALG_SHA1 {
        return Err(DecryptError::UnsupportedEncryption(format!(
            "CryptoAPI AlgID=0x{:08X} AlgIDHash=0x{:08X}",
            info.header.alg_id, info.header.alg_id_hash
        )));
    }

    // RC4 CryptoAPI uses a fixed 16-byte salt.
    if info.verifier.salt.len() != 16 {
        return Err(DecryptError::UnsupportedEncryption);
    }

    let key_size_bits = info.header.key_size_bits;
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
    if verifier_hash_size != 20 {
        // Office 97-2003 CryptoAPI RC4 uses SHA1 verifier hashes.
        return Err(DecryptError::UnsupportedEncryption(format!(
            "CryptoAPI verifierHashSize={verifier_hash_size}"
        )));
    }

    // Derive the base key material and decrypt the verifier using block 0.
    let key_material: Zeroizing<[u8; 20]> =
        Zeroizing::new(derive_key_material(password, &info.verifier.salt));
    let key0: Zeroizing<Vec<u8>> = Zeroizing::new(derive_block_key(&key_material, 0, key_len));
    let mut rc4 = Rc4::new(key0.as_slice());

    let mut verifier: Zeroizing<[u8; 16]> = Zeroizing::new(info.verifier.encrypted_verifier);
    rc4.apply_keystream(&mut verifier[..]);

    let mut verifier_hash: Zeroizing<Vec<u8>> =
        Zeroizing::new(info.verifier.encrypted_verifier_hash.clone());
    rc4.apply_keystream(&mut verifier_hash[..]);

    let expected: Zeroizing<[u8; 20]> = Zeroizing::new(sha1_bytes(&[&verifier[..]]));
    if verifier_hash.len() < verifier_hash_size {
        return Err(DecryptError::InvalidFormat(format!(
            "EncryptedVerifierHash length {} shorter than VerifierHashSize {verifier_hash_size}",
            verifier_hash.len()
        )));
    }
    if !ct_eq(&verifier_hash[..verifier_hash_size], &expected[..verifier_hash_size]) {
        return Err(DecryptError::WrongPassword);
    }

    Ok(*key_material)
}

fn verify_password_legacy(
    info: &CryptoApiEncryptionInfo,
    password: &str,
) -> Result<[u8; 20], DecryptError> {
    if info.header.alg_id != CALG_RC4 || info.header.alg_id_hash != CALG_SHA1 {
        return Err(DecryptError::UnsupportedEncryption(format!(
            "CryptoAPI AlgID=0x{:08X} AlgIDHash=0x{:08X}",
            info.header.alg_id, info.header.alg_id_hash
        )));
    }

    let key_size_bits = info.header.key_size_bits;
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
    if verifier_hash_size != 20 {
        return Err(DecryptError::UnsupportedEncryption(format!(
            "CryptoAPI verifierHashSize={verifier_hash_size}"
        )));
    }

    let key_material = derive_key_material_legacy(password, &info.verifier.salt)?;
    let key0 = derive_block_key(&key_material, 0, key_len);
    let mut rc4 = Rc4::new(&key0);

    let mut verifier = info.verifier.encrypted_verifier;
    rc4.apply_keystream(&mut verifier);

    let mut verifier_hash = info.verifier.encrypted_verifier_hash.clone();
    rc4.apply_keystream(&mut verifier_hash);

    let expected = sha1_bytes(&[&verifier]);
    if verifier_hash.len() < verifier_hash_size {
        return Err(DecryptError::InvalidFormat(format!(
            "EncryptedVerifierHash length {} shorter than VerifierHashSize {verifier_hash_size}",
            verifier_hash.len()
        )));
    }
    if !ct_eq(&verifier_hash[..verifier_hash_size], &expected[..verifier_hash_size]) {
        return Err(DecryptError::WrongPassword);
    }

    Ok(key_material)
}

fn parse_filepass_record_payload(payload: &[u8]) -> Result<CryptoApiEncryptionInfo, DecryptError> {
    // FILEPASS payload layout for BIFF8 [MS-XLS] 2.4.117 (FilePass):
    // - u16 wEncryptionType
    // - (if wEncryptionType == 0x0000) XOR obfuscation: u16 key, u16 verifier
    // - (if wEncryptionType == 0x0001) RC4 encryption:
    //     u16 wEncryptionSubType
    //     (subType-specific payload)
    //
    // This decryptor only supports RC4 CryptoAPI (`wEncryptionType=0x0001`, `wEncryptionSubType=0x0002`).
    // Be careful to return `UnsupportedEncryption` (not `InvalidFormat`) when the payload is well-formed
    // enough to identify an unsupported scheme.

    let encryption_type = read_u16_le(payload, 0).ok_or_else(|| {
        DecryptError::InvalidFormat(format!(
            "FILEPASS payload truncated while reading wEncryptionType (len={})",
            payload.len()
        ))
    })?;

    if encryption_type != ENCRYPTION_TYPE_RC4 {
        return Err(DecryptError::UnsupportedEncryption(format!(
            "FILEPASS wEncryptionType=0x{encryption_type:04X}"
        )));
    }

    let encryption_subtype = read_u16_le(payload, 2).ok_or_else(|| {
        DecryptError::InvalidFormat(format!(
            "FILEPASS payload truncated while reading wEncryptionSubType (len={})",
            payload.len()
        ))
    })?;

    if encryption_subtype != ENCRYPTION_SUBTYPE_CRYPTOAPI {
        return Err(DecryptError::UnsupportedEncryption(format!(
            "FILEPASS RC4 wEncryptionSubType=0x{encryption_subtype:04X}"
        )));
    }

    if payload.len() < 8 {
        return Err(DecryptError::InvalidFormat(format!(
            "FILEPASS payload truncated (len={})",
            payload.len()
        )));
    }
    let enc_info_len = read_u32_le(payload, 4).ok_or_else(|| {
        DecryptError::InvalidFormat("FILEPASS missing dwEncryptionInfoLen".to_string())
    })? as usize;

    let enc_info_start = 8usize;
    let enc_info_end = enc_info_start.checked_add(enc_info_len).ok_or_else(|| {
        DecryptError::InvalidFormat("dwEncryptionInfoLen overflow".to_string())
    })?;
    let enc_info = payload.get(enc_info_start..enc_info_end).ok_or_else(|| {
        DecryptError::InvalidFormat(format!(
            "FILEPASS dwEncryptionInfoLen out of bounds (payload_len={}, enc_info_end={enc_info_end})",
            payload.len()
        ))
    })?;

    parse_cryptoapi_encryption_info(enc_info)
}

fn find_filepass_record_offset(workbook_stream: &[u8]) -> Result<(usize, usize), DecryptError> {
    let mut offset: usize = 0;
    while offset < workbook_stream.len() {
        if offset + 4 > workbook_stream.len() {
            return Err(DecryptError::InvalidFormat(
                "truncated BIFF record header".to_string(),
            ));
        }

        let record_id = u16::from_le_bytes([workbook_stream[offset], workbook_stream[offset + 1]]);
        let len = u16::from_le_bytes([workbook_stream[offset + 2], workbook_stream[offset + 3]])
            as usize;
        let data_start = offset + 4;
        let data_end = data_start.checked_add(len).ok_or_else(|| {
            DecryptError::InvalidFormat("BIFF record length overflow".to_string())
        })?;
        if data_end > workbook_stream.len() {
            return Err(DecryptError::InvalidFormat(format!(
                "BIFF record 0x{record_id:04X} at offset {offset} extends past end of stream (len={}, end={data_end})",
                workbook_stream.len()
            )));
        }

        if record_id == RECORD_FILEPASS {
            return Ok((offset, len));
        }

        offset = data_end;
    }

    Err(DecryptError::InvalidFormat(
        "missing FILEPASS record".to_string(),
    ))
}

/// Decrypt an in-memory BIFF8 workbook stream encrypted with RC4 CryptoAPI.
///
/// Supports:
/// - `FILEPASS` `wEncryptionSubType = 0x0002` (length-prefixed CryptoAPI `EncryptionInfo`)
/// - legacy `FILEPASS` `wEncryptionInfo = 0x0004` (embedded `EncryptionHeader` + `EncryptionVerifier`)
///
/// The returned workbook stream has the `FILEPASS` record *masked* (record id replaced with
/// `0xFFFF`) so downstream parsers that do not implement BIFF encryption treat the stream as
/// plaintext without shifting any record offsets (e.g. `BoundSheet8.lbPlyPos`).
pub(crate) fn decrypt_biff8_workbook_stream_rc4_cryptoapi(
    workbook_stream: &[u8],
    password: &str,
) -> Result<Vec<u8>, DecryptError> {
    let (filepass_offset, filepass_len) = find_filepass_record_offset(workbook_stream)?;
    let filepass_data_start = filepass_offset + 4;
    let filepass_data_end = filepass_data_start
        .checked_add(filepass_len)
        .ok_or_else(|| DecryptError::InvalidFormat("FILEPASS length overflow".to_string()))?;
    let filepass_payload = workbook_stream
        .get(filepass_data_start..filepass_data_end)
        .ok_or_else(|| {
            DecryptError::InvalidFormat("FILEPASS payload out of bounds".to_string())
        })?;

    let encryption_type = read_u16_le(filepass_payload, 0).ok_or_else(|| {
        DecryptError::InvalidFormat("FILEPASS missing wEncryptionType".to_string())
    })?;
    if encryption_type != ENCRYPTION_TYPE_RC4 {
        return Err(DecryptError::UnsupportedEncryption(format!(
            "FILEPASS wEncryptionType=0x{encryption_type:04X}"
        )));
    }
    let second_field = read_u16_le(filepass_payload, 2).ok_or_else(|| {
        DecryptError::InvalidFormat("FILEPASS missing subtype/encryption info".to_string())
    })?;

    match second_field {
        // Modern BIFF8 RC4 CryptoAPI FILEPASS layout: `wEncryptionSubType == 0x0002` followed by a
        // length-prefixed EncryptionInfo blob.
        ENCRYPTION_SUBTYPE_CRYPTOAPI => {
            let info = parse_filepass_record_payload(filepass_payload)?;
            let key_material = verify_password(&info, password)?;

            let key_size_bits = info.header.key_size_bits;
            let key_len = (key_size_bits / 8) as usize;

            let mut out = workbook_stream.to_vec();
            out[filepass_offset..filepass_offset + 2].copy_from_slice(&RECORD_MASKED.to_le_bytes());

            // Decrypt all subsequent record payloads in-place using the record-payload-only stream
            // model.
            let mut cipher = PayloadRc4::new(key_material, key_len);

            let mut offset = filepass_data_end;
            while offset < out.len() {
                let remaining = out.len().saturating_sub(offset);
                if remaining < 4 {
                    // Some writers may include trailing padding bytes after the final EOF record.
                    // Those bytes are not part of any record payload and should be ignored rather
                    // than treated as a truncated record header.
                    break;
                }

                if offset + 4 > out.len() {
                    return Err(DecryptError::InvalidFormat(
                        "truncated BIFF record header".to_string(),
                    ));
                }

                let len = u16::from_le_bytes([out[offset + 2], out[offset + 3]]) as usize;
                let data_start = offset + 4;
                let data_end = data_start.checked_add(len).ok_or_else(|| {
                    DecryptError::InvalidFormat("BIFF record length overflow".to_string())
                })?;
                if data_end > out.len() {
                    return Err(DecryptError::InvalidFormat(format!(
                        "BIFF record at offset {offset} extends past end of stream (len={}, end={data_end})",
                        out.len()
                    )));
                }

                cipher.apply_keystream(&mut out[data_start..data_end]);
                offset = data_end;
            }

            Ok(out)
        }
        // Legacy BIFF8 RC4 CryptoAPI FILEPASS layout (`wEncryptionInfo == 0x0004`): no
        // length-prefixed EncryptionInfo blob; the EncryptionHeader + Verifier are embedded
        // directly in the FILEPASS payload.
        //
        // Excel derives per-1024-byte block keys based on an absolute stream position that includes
        // record headers and unencrypted bytes.
        ENCRYPTION_INFO_CRYPTOAPI_LEGACY => {
            const RECORD_BOUNDSHEET8: u16 = 0x0085;

            let info = parse_cryptoapi_encryption_info_legacy_filepass(filepass_payload)?;
            let key_material = verify_password_legacy(&info, password)?;

            let key_size_bits = info.header.key_size_bits;
            let key_len = (key_size_bits / 8) as usize;

            let mut out = workbook_stream.to_vec();
            out[filepass_offset..filepass_offset + 2].copy_from_slice(&RECORD_MASKED.to_le_bytes());

            let mut offset = filepass_data_end;
            // "Encryption stream position" is keyed by the *absolute* offset within the workbook
            // stream. Encryption begins after FILEPASS, but the block counter (and keystream
            // discard) still incorporate the preceding plaintext bytes.
            let mut stream_pos: usize = filepass_data_end;

            while offset < out.len() {
                let remaining = out.len().saturating_sub(offset);
                if remaining < 4 {
                    break;
                }

                let record_id = u16::from_le_bytes([out[offset], out[offset + 1]]);
                let len = u16::from_le_bytes([out[offset + 2], out[offset + 3]]) as usize;
                let data_start = offset + 4;
                let data_end = data_start.checked_add(len).ok_or_else(|| {
                    DecryptError::InvalidFormat("BIFF record length overflow".to_string())
                })?;
                if data_end > out.len() {
                    return Err(DecryptError::InvalidFormat(format!(
                        "BIFF record 0x{record_id:04X} at offset {offset} extends past end of stream (len={}, end={data_end})",
                        out.len()
                    )));
                }

                // Record headers are not encrypted but still advance the CryptoAPI RC4 stream
                // position (Apache POI / Excel behaviour).
                stream_pos = stream_pos.checked_add(4).ok_or_else(|| {
                    DecryptError::InvalidFormat("stream position overflow".to_string())
                })?;

                if !is_never_encrypted_record(record_id) {
                    match record_id {
                        // BoundSheet8.lbPlyPos (first 4 bytes) must remain plaintext so sheet
                        // offsets remain valid after we mask FILEPASS. The remaining fields are
                        // encrypted.
                        RECORD_BOUNDSHEET8 => {
                            if len > 4 {
                                let decrypt_start = stream_pos.checked_add(4).ok_or_else(|| {
                                    DecryptError::InvalidFormat("stream position overflow".to_string())
                                })?;
                                decrypt_range_by_offset(
                                    &mut out[data_start + 4..data_end],
                                    decrypt_start,
                                    &key_material,
                                    key_len,
                                );
                            }
                        }
                        _ => decrypt_range_by_offset(
                            &mut out[data_start..data_end],
                            stream_pos,
                            &key_material,
                            key_len,
                        ),
                    }
                }

                // Advance past the record payload, regardless of whether we decrypted it.
                stream_pos = stream_pos.checked_add(len).ok_or_else(|| {
                    DecryptError::InvalidFormat("stream position overflow".to_string())
                })?;
                offset = data_end;
            }

            Ok(out)
        }
        _ => Err(DecryptError::UnsupportedEncryption(format!(
            "FILEPASS RC4 wEncryptionSubType/wEncryptionInfo=0x{second_field:04X}"
        ))),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use proptest::prelude::*;

    const RECORD_BOF_BIFF8: u16 = 0x0809;
    const RECORD_EOF: u16 = 0x000A;

    fn record(record_id: u16, payload: &[u8]) -> Vec<u8> {
        let mut out = Vec::with_capacity(4 + payload.len());
        out.extend_from_slice(&record_id.to_le_bytes());
        out.extend_from_slice(&(payload.len() as u16).to_le_bytes());
        out.extend_from_slice(payload);
        out
    }

    // Minimal BIFF8 workbook-globals BOF payload: BIFF8 version + workbook globals subtype.
    const BOF_GLOBALS: [u8; 4] = [0x00, 0x06, 0x05, 0x00];

    #[test]
    fn filepass_xor_is_reported_as_unsupported_encryption() {
        // BIFF8 XOR obfuscation FILEPASS payload is:
        //   u16 wEncryptionType (0x0000)
        //   u16 key
        //   u16 verifier
        let payload = [0x00, 0x00, 0x34, 0x12, 0x78, 0x56];
        let err = parse_filepass_record_payload(&payload).expect_err("expected error");
        assert!(matches!(err, DecryptError::UnsupportedEncryption(_)));
    }

    #[test]
    fn filepass_rc4_non_cryptoapi_is_reported_as_unsupported_encryption() {
        // RC4 "standard" (non-CryptoAPI) has wEncryptionType==0x0001 but a different subtype.
        let payload = [0x01, 0x00, 0x01, 0x00];
        let err = parse_filepass_record_payload(&payload).expect_err("expected error");
        assert!(matches!(err, DecryptError::UnsupportedEncryption(_)));
    }

    #[test]
    fn filepass_missing_encryption_type_is_invalid_format() {
        let payload = [0x00];
        let err = parse_filepass_record_payload(&payload).expect_err("expected error");
        assert!(matches!(err, DecryptError::InvalidFormat(_)));
    }

    #[test]
    fn filepass_truncated_missing_subtype_is_invalid_format() {
        let payload = [0x01, 0x00, 0x02];
        let err = parse_filepass_record_payload(&payload).expect_err("expected error");
        assert!(matches!(err, DecryptError::InvalidFormat(_)));
    }

    #[test]
    fn filepass_truncated_missing_encryption_info_len_is_invalid_format() {
        // wEncryptionType=RC4, wEncryptionSubType=CryptoAPI, but missing dwEncryptionInfoLen.
        let payload = [0x01, 0x00, 0x02, 0x00, 0x00, 0x00, 0x00];
        let err = parse_filepass_record_payload(&payload).expect_err("expected error");
        assert!(matches!(err, DecryptError::InvalidFormat(_)));
    }

    proptest! {
        #![proptest_config(ProptestConfig {
            cases: 64,
            rng_seed: proptest::test_runner::RngSeed::Fixed(0),
            failure_persistence: None,
            .. ProptestConfig::default()
        })]

        #[test]
        fn decryptor_handles_truncated_stream(prefix in proptest::collection::vec(any::<u8>(), 0..4)) {
            let res = decrypt_biff8_workbook_stream_rc4_cryptoapi(&prefix, "pw");
            prop_assert!(matches!(res, Err(DecryptError::InvalidFormat(_))));
        }

        #[test]
        fn decryptor_handles_truncated_record_header_after_bof(trailing in proptest::collection::vec(any::<u8>(), 0..4)) {
            let mut stream = record(RECORD_BOF_BIFF8, &BOF_GLOBALS);
            stream.extend_from_slice(&trailing);

            let res = decrypt_biff8_workbook_stream_rc4_cryptoapi(&stream, "pw");
            prop_assert!(matches!(res, Err(DecryptError::InvalidFormat(_))));
        }

        #[test]
        fn decryptor_rejects_too_short_cryptoapi_filepass_payload(len in 0u16..8u16) {
            // CryptoAPI FILEPASS requires at least:
            //   wEncryptionType (2) + wEncryptionSubType (2) + dwEncryptionInfoLen (4) = 8 bytes.
            let mut payload = vec![0u8; len as usize];
            if payload.len() >= 2 {
                payload[0..2].copy_from_slice(&ENCRYPTION_TYPE_RC4.to_le_bytes());
            }
            if payload.len() >= 4 {
                payload[2..4].copy_from_slice(&ENCRYPTION_SUBTYPE_CRYPTOAPI.to_le_bytes());
            }

            let stream = [
                record(RECORD_BOF_BIFF8, &BOF_GLOBALS),
                record(RECORD_FILEPASS, &payload),
                record(RECORD_EOF, &[]),
            ].concat();

            let res = decrypt_biff_workbook_stream(&stream, "pw");
            prop_assert!(matches!(res, Err(DecryptError::InvalidFormat(_))));
        }

        #[test]
        fn decryptor_rejects_declared_encryptioninfo_len_that_exceeds_payload(enc_info_len in 1u16..512u16) {
            // FILEPASS declares a dwEncryptionInfoLen but does not provide enough bytes.
            let mut payload = Vec::new();
            payload.extend_from_slice(&ENCRYPTION_TYPE_RC4.to_le_bytes());
            payload.extend_from_slice(&ENCRYPTION_SUBTYPE_CRYPTOAPI.to_le_bytes());
            payload.extend_from_slice(&(enc_info_len as u32).to_le_bytes());
            // Intentionally omit EncryptionInfo bytes.

            let stream = [
                record(RECORD_BOF_BIFF8, &BOF_GLOBALS),
                record(RECORD_FILEPASS, &payload),
                record(RECORD_EOF, &[]),
            ].concat();

            let res = decrypt_biff_workbook_stream(&stream, "pw");
            prop_assert!(matches!(res, Err(DecryptError::InvalidFormat(_))));
        }
    }

    #[test]
    fn decryptor_rejects_huge_record_len_that_exceeds_stream_end() {
        // BIFF record declaring a huge payload but missing bytes should not panic or read OOB.
        let mut bogus = Vec::new();
        bogus.extend_from_slice(&RECORD_FILEPASS.to_le_bytes());
        bogus.extend_from_slice(&u16::MAX.to_le_bytes());
        bogus.extend_from_slice(&[0xAA, 0xBB, 0xCC]); // truncated payload

        let stream = [record(RECORD_BOF_BIFF8, &BOF_GLOBALS), bogus].concat();
        let res = decrypt_biff8_workbook_stream_rc4_cryptoapi(&stream, "pw");
        assert!(matches!(res, Err(DecryptError::InvalidFormat(_))), "res={res:?}");
    }

    #[test]
    fn decryptor_reports_unsupported_encryption_type() {
        // FILEPASS with an unknown encryption type should surface UnsupportedEncryption.
        let mut payload = Vec::new();
        payload.extend_from_slice(&0x0002u16.to_le_bytes()); // unknown wEncryptionType
        payload.extend_from_slice(&0u16.to_le_bytes()); // padding

        let stream = [
            record(RECORD_BOF_BIFF8, &BOF_GLOBALS),
            record(RECORD_FILEPASS, &payload),
            record(RECORD_EOF, &[]),
        ]
        .concat();

        let res = decrypt_biff_workbook_stream(&stream, "pw");
        assert!(matches!(res, Err(DecryptError::UnsupportedEncryption)), "res={res:?}");
    }
}
