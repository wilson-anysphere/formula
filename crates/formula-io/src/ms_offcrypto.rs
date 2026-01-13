//! Minimal parsers for MS-OFFCRYPTO structures used by encrypted OOXML (OLE) containers.
//!
//! This module intentionally implements **only** what we need for conservative detection and
//! future decryption support. The goal is to reject random/garbage inputs deterministically.
//!
//! References:
//! - MS-OFFCRYPTO: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/
//! - `EncryptionInfo` stream (Standard): `EncryptionInfo` → `EncryptionHeader` → `EncryptionVerifier`

use thiserror::Error;

/// MS-OFFCRYPTO "Standard" encryption version (CryptoAPI).
///
/// This is used by the `EncryptionInfo` stream in an encrypted OOXML OLE container.
const STANDARD_VERSION_MAJOR: u16 = 3;
const STANDARD_VERSION_MINOR: u16 = 2;

/// CryptoAPI `ALG_ID` values for AES (Windows).
///
/// These are commonly used in Office Standard encryption headers.
const CALG_AES_128: u32 = 0x0000_660E;
const CALG_AES_192: u32 = 0x0000_660F;
const CALG_AES_256: u32 = 0x0000_6610;
// Some references mention `CALG_AES` without key size; keep it in the "AES" family to avoid
// misclassifying it as non-AES.
const CALG_AES: u32 = 0x0000_6611;

/// CryptoAPI `ALG_ID` value for RC4 (Windows).
#[allow(dead_code)]
const CALG_RC4: u32 = 0x0000_6801;

#[derive(Debug, Error)]
pub(crate) enum EncryptionInfoError {
    #[error("EncryptionInfo is truncated")]
    Truncated,
    #[error("unsupported EncryptionInfo version: {major}.{minor}")]
    UnsupportedVersion { major: u16, minor: u16 },
    #[error("invalid EncryptionHeaderSize: {0}")]
    InvalidEncryptionHeaderSize(u32),
    #[error("invalid EncryptionHeader CSPName (missing NUL terminator)")]
    InvalidCspName,
    #[error("unsupported external Standard encryption (fExternal flag set)")]
    UnsupportedExternalEncryption,
    #[error("unsupported Standard encryption: fCryptoAPI flag not set")]
    UnsupportedNonCryptoApiStandardEncryption,
    #[error("invalid Standard EncryptionHeader flags for algId: flags={flags:#010x}, alg_id={alg_id:#010x}")]
    InvalidFlags { flags: u32, alg_id: u32 },
}

fn read_u16_le(buf: &[u8], offset: &mut usize) -> Result<u16, EncryptionInfoError> {
    let b = buf
        .get(*offset..*offset + 2)
        .ok_or(EncryptionInfoError::Truncated)?;
    *offset += 2;
    Ok(u16::from_le_bytes([b[0], b[1]]))
}

fn read_u32_le(buf: &[u8], offset: &mut usize) -> Result<u32, EncryptionInfoError> {
    let b = buf
        .get(*offset..*offset + 4)
        .ok_or(EncryptionInfoError::Truncated)?;
    *offset += 4;
    Ok(u32::from_le_bytes([b[0], b[1], b[2], b[3]]))
}

fn is_aes_alg_id(alg_id: u32) -> bool {
    matches!(
        alg_id,
        CALG_AES_128 | CALG_AES_192 | CALG_AES_256 | CALG_AES
    )
}

/// Parsed MS-OFFCRYPTO `EncryptionHeader.Flags` bits for Standard encryption.
///
/// The raw value contains additional bits not currently modeled here.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct EncryptionHeaderFlags {
    pub(crate) raw: u32,
    pub(crate) f_cryptoapi: bool,
    pub(crate) f_doc_props: bool,
    pub(crate) f_external: bool,
    pub(crate) f_aes: bool,
}

impl EncryptionHeaderFlags {
    pub(crate) const F_CRYPTOAPI: u32 = 0x0000_0004;
    pub(crate) const F_DOCPROPS: u32 = 0x0000_0008;
    pub(crate) const F_EXTERNAL: u32 = 0x0000_0010;
    pub(crate) const F_AES: u32 = 0x0000_0020;

    pub(crate) fn from_raw(raw: u32) -> Self {
        Self {
            raw,
            f_cryptoapi: raw & Self::F_CRYPTOAPI != 0,
            f_doc_props: raw & Self::F_DOCPROPS != 0,
            f_external: raw & Self::F_EXTERNAL != 0,
            f_aes: raw & Self::F_AES != 0,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct StandardEncryptionHeader {
    pub(crate) flags: EncryptionHeaderFlags,
    pub(crate) alg_id: u32,
}

/// Parse a Standard `EncryptionHeader` and validate key flags.
fn parse_standard_encryption_header(
    buf: &[u8],
) -> Result<StandardEncryptionHeader, EncryptionInfoError> {
    // `EncryptionHeader` fixed-size fields are 8 DWORDs, then a variable-length UTF-16LE CSPName.
    const FIXED_LEN: usize = 8 * 4;
    if buf.len() < FIXED_LEN + 2 {
        return Err(EncryptionInfoError::Truncated);
    }

    let mut offset = 0usize;
    let flags_raw = read_u32_le(buf, &mut offset)?;
    let flags = EncryptionHeaderFlags::from_raw(flags_raw);
    let _size_extra = read_u32_le(buf, &mut offset)?;
    let alg_id = read_u32_le(buf, &mut offset)?;
    let _alg_id_hash = read_u32_le(buf, &mut offset)?;
    let _key_size = read_u32_le(buf, &mut offset)?;
    let _provider_type = read_u32_le(buf, &mut offset)?;
    let _reserved1 = read_u32_le(buf, &mut offset)?;
    let _reserved2 = read_u32_le(buf, &mut offset)?;

    // Validate flag semantics.
    if flags.f_external {
        return Err(EncryptionInfoError::UnsupportedExternalEncryption);
    }
    if !flags.f_cryptoapi {
        return Err(EncryptionInfoError::UnsupportedNonCryptoApiStandardEncryption);
    }

    // Policy: treat AES flag/algId mismatches as invalid. This is conservative and avoids false
    // positives when parsing arbitrary OLE streams.
    let alg_is_aes = is_aes_alg_id(alg_id);
    if flags.f_aes != alg_is_aes {
        return Err(EncryptionInfoError::InvalidFlags {
            flags: flags_raw,
            alg_id,
        });
    }

    // Validate CSPName (must be UTF-16LE NUL-terminated within the declared header size).
    let csp_bytes = buf.get(offset..).ok_or(EncryptionInfoError::Truncated)?;
    if csp_bytes.len() % 2 != 0 {
        return Err(EncryptionInfoError::InvalidCspName);
    }
    let mut has_nul = false;
    for pair in csp_bytes.chunks_exact(2) {
        if pair == [0, 0] {
            has_nul = true;
            break;
        }
    }
    if !has_nul {
        return Err(EncryptionInfoError::InvalidCspName);
    }

    Ok(StandardEncryptionHeader { flags, alg_id })
}

/// Parse an MS-OFFCRYPTO `EncryptionInfo` stream as Standard encryption.
pub(crate) fn parse_standard_encryption_info(
    bytes: &[u8],
) -> Result<StandardEncryptionHeader, EncryptionInfoError> {
    let mut offset = 0usize;
    let major = read_u16_le(bytes, &mut offset)?;
    let minor = read_u16_le(bytes, &mut offset)?;

    if major != STANDARD_VERSION_MAJOR || minor != STANDARD_VERSION_MINOR {
        return Err(EncryptionInfoError::UnsupportedVersion { major, minor });
    }

    // `EncryptionInfo.Flags` (not to be confused with `EncryptionHeader.Flags`).
    let _info_flags = read_u32_le(bytes, &mut offset)?;

    // Standard encryption includes a DWORD size prefix for the subsequent `EncryptionHeader`.
    let header_size = read_u32_le(bytes, &mut offset)?;
    let header_size_usize = usize::try_from(header_size).map_err(|_| {
        // Header size is untrusted input; treat overflow as invalid.
        EncryptionInfoError::InvalidEncryptionHeaderSize(header_size)
    })?;

    // The header contains at least the fixed fields plus a UTF-16 NUL terminator.
    if header_size_usize < 8 * 4 + 2 {
        return Err(EncryptionInfoError::InvalidEncryptionHeaderSize(
            header_size,
        ));
    }
    if header_size_usize % 2 != 0 {
        return Err(EncryptionInfoError::InvalidEncryptionHeaderSize(
            header_size,
        ));
    }

    let header_end = offset
        .checked_add(header_size_usize)
        .ok_or(EncryptionInfoError::Truncated)?;
    let header = bytes
        .get(offset..header_end)
        .ok_or(EncryptionInfoError::Truncated)?;
    offset += header_size_usize;

    let parsed_header = parse_standard_encryption_header(header)?;

    // Best-effort parse of the following `EncryptionVerifier` header fields to ensure the stream is
    // not trivially truncated. We don't currently surface these values, but checking lengths makes
    // the overall parser more conservative.
    //
    // `EncryptionVerifier` begins with:
    //   DWORD SaltSize;
    //   BYTE  Salt[SaltSize];
    //   BYTE  EncryptedVerifier[16];
    //   DWORD VerifierHashSize;
    //   BYTE  EncryptedVerifierHash[...]
    //
    // The `EncryptedVerifierHash` length depends on the encryption algorithm: for AES it's padded
    // to a 16-byte block.
    let verifier_salt_size = read_u32_le(bytes, &mut offset)? as usize;
    let verifier_salt_end = offset
        .checked_add(verifier_salt_size)
        .ok_or(EncryptionInfoError::Truncated)?;
    let verifier_salt = bytes
        .get(offset..verifier_salt_end)
        .ok_or(EncryptionInfoError::Truncated)?;
    let _ = verifier_salt;
    offset += verifier_salt_size;

    let encrypted_verifier_end = offset
        .checked_add(16)
        .ok_or(EncryptionInfoError::Truncated)?;
    let _encrypted_verifier = bytes
        .get(offset..encrypted_verifier_end)
        .ok_or(EncryptionInfoError::Truncated)?;
    offset += 16;

    let verifier_hash_size = read_u32_le(bytes, &mut offset)? as usize;
    let expected_hash_len = if is_aes_alg_id(parsed_header.alg_id) {
        // AES is a block cipher (16-byte blocks).
        verifier_hash_size
            .checked_add(15)
            .map(|v| (v / 16) * 16)
            .ok_or(EncryptionInfoError::Truncated)?
    } else {
        // RC4 is stream-based (no padding).
        verifier_hash_size
    };
    let verifier_hash_end = offset
        .checked_add(expected_hash_len)
        .ok_or(EncryptionInfoError::Truncated)?;
    let _encrypted_verifier_hash = bytes
        .get(offset..verifier_hash_end)
        .ok_or(EncryptionInfoError::Truncated)?;
    // We intentionally do not require that the verifier hash consumes the rest of the stream;
    // producers may append additional data.

    Ok(parsed_header)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn build_standard_encryption_info(encryption_header_flags: u32, alg_id: u32) -> Vec<u8> {
        // Build a minimal Standard EncryptionInfo buffer:
        // - Version (3.2)
        // - Info flags (0)
        // - Header size
        // - EncryptionHeader (fixed fields + empty CSPName)
        // - EncryptionVerifier (minimal lengths)

        let mut out = Vec::new();
        out.extend_from_slice(&STANDARD_VERSION_MAJOR.to_le_bytes());
        out.extend_from_slice(&STANDARD_VERSION_MINOR.to_le_bytes());
        out.extend_from_slice(&0u32.to_le_bytes()); // EncryptionInfo.Flags (ignored)

        // Header: 8 DWORDs + UTF-16LE NUL terminator.
        let header_size = (8 * 4 + 2) as u32;
        out.extend_from_slice(&header_size.to_le_bytes());

        // EncryptionHeader.
        out.extend_from_slice(&encryption_header_flags.to_le_bytes()); // Flags
        out.extend_from_slice(&0u32.to_le_bytes()); // SizeExtra
        out.extend_from_slice(&alg_id.to_le_bytes()); // AlgId
        out.extend_from_slice(&0u32.to_le_bytes()); // AlgIdHash
        out.extend_from_slice(&128u32.to_le_bytes()); // KeySize (bits)
        out.extend_from_slice(&0u32.to_le_bytes()); // ProviderType
        out.extend_from_slice(&0u32.to_le_bytes()); // Reserved1
        out.extend_from_slice(&0u32.to_le_bytes()); // Reserved2
        out.extend_from_slice(&0u16.to_le_bytes()); // CSPName (empty string terminator)

        // EncryptionVerifier.
        out.extend_from_slice(&16u32.to_le_bytes()); // SaltSize
        out.extend_from_slice(&[0u8; 16]); // Salt
        out.extend_from_slice(&[0u8; 16]); // EncryptedVerifier
        out.extend_from_slice(&20u32.to_le_bytes()); // VerifierHashSize (SHA-1)
        let hash_len = if is_aes_alg_id(alg_id) { 32 } else { 20 };
        out.extend_from_slice(&vec![0u8; hash_len]); // EncryptedVerifierHash (padded for AES)

        out
    }

    #[test]
    fn standard_rejects_external_encryption_flag() {
        let flags = EncryptionHeaderFlags::F_CRYPTOAPI | EncryptionHeaderFlags::F_EXTERNAL;
        let bytes = build_standard_encryption_info(flags, CALG_RC4);
        let err = parse_standard_encryption_info(&bytes).expect_err("expected error");
        assert!(matches!(
            err,
            EncryptionInfoError::UnsupportedExternalEncryption
        ));
    }

    #[test]
    fn standard_rejects_non_cryptoapi_encryption() {
        let flags = 0u32;
        let bytes = build_standard_encryption_info(flags, CALG_RC4);
        let err = parse_standard_encryption_info(&bytes).expect_err("expected error");
        assert!(matches!(
            err,
            EncryptionInfoError::UnsupportedNonCryptoApiStandardEncryption
        ));
    }

    #[test]
    fn standard_rejects_aes_alg_id_without_faes_flag() {
        let flags = EncryptionHeaderFlags::F_CRYPTOAPI;
        let bytes = build_standard_encryption_info(flags, CALG_AES_128);
        let err = parse_standard_encryption_info(&bytes).expect_err("expected error");
        assert!(matches!(err, EncryptionInfoError::InvalidFlags { .. }));
    }

    #[test]
    fn standard_rejects_faes_flag_without_aes_alg_id() {
        let flags = EncryptionHeaderFlags::F_CRYPTOAPI | EncryptionHeaderFlags::F_AES;
        let bytes = build_standard_encryption_info(flags, CALG_RC4);
        let err = parse_standard_encryption_info(&bytes).expect_err("expected error");
        assert!(matches!(err, EncryptionInfoError::InvalidFlags { .. }));
    }

    #[test]
    fn standard_accepts_cryptoapi_rc4_flags() {
        let flags = EncryptionHeaderFlags::F_CRYPTOAPI | EncryptionHeaderFlags::F_DOCPROPS;
        let bytes = build_standard_encryption_info(flags, CALG_RC4);
        let header = parse_standard_encryption_info(&bytes).expect("should parse");
        assert_eq!(header.flags.f_cryptoapi, true);
        assert_eq!(header.flags.f_doc_props, true);
        assert_eq!(header.flags.f_external, false);
        assert_eq!(header.flags.f_aes, false);
        assert_eq!(header.alg_id, CALG_RC4);
    }

    #[test]
    fn standard_accepts_cryptoapi_aes_flags() {
        let flags = EncryptionHeaderFlags::F_CRYPTOAPI | EncryptionHeaderFlags::F_AES;
        let bytes = build_standard_encryption_info(flags, CALG_AES_128);
        let header = parse_standard_encryption_info(&bytes).expect("should parse");
        assert_eq!(header.flags.f_cryptoapi, true);
        assert_eq!(header.flags.f_aes, true);
        assert_eq!(header.alg_id, CALG_AES_128);
    }
}
