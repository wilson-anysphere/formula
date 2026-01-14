//! Minimal parsers for MS-OFFCRYPTO structures used by encrypted OOXML (OLE) containers.
//!
//! This module intentionally implements **only** what we need for conservative detection and
//! future decryption support. The goal is to reject random/garbage inputs deterministically.
//!
//! References:
//! - MS-OFFCRYPTO: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/
//! - `EncryptionInfo` stream (Standard): `EncryptionInfo` → `EncryptionHeader` → `EncryptionVerifier`
//!
//! ## `EncryptionHeader.sizeExtra`
//!
//! For Standard encryption (`EncryptionVersionInfo` major=3, minor=2), the `EncryptionInfo` stream
//! contains a length-prefixed `EncryptionHeader` blob. The header begins with 8 DWORD fields
//! (32 bytes) and ends with:
//! - a UTF-16LE `CSPName` field (variable length), followed by
//! - `sizeExtra` algorithm-specific bytes (also variable length).
//!
//! `headerSize` is the total byte length of the `EncryptionHeader` blob, so the number of bytes
//! that belong to `CSPName` must be computed as:
//!
//! `csp_name_bytes_len = headerSize - 32 - sizeExtra`
//!
//! The trailing `sizeExtra` bytes must be skipped (or stored) before parsing the subsequent
//! `EncryptionVerifier` fields.

// This module is intentionally not wired into the main `formula-io` open path yet; keep it
// compiling for future work without spamming downstream builds with dead-code warnings.
#![allow(dead_code)]

use encoding_rs::UTF_16LE;
use thiserror::Error;

/// MS-OFFCRYPTO "Standard" encryption version (CryptoAPI).
///
/// MS-OFFCRYPTO identifies Standard encryption via `versionMinor == 2`, but real-world files vary
/// `versionMajor` across Office generations (2/3/4). Keep the canonical major version here for
/// tests that construct synthetic `EncryptionInfo` streams.
const STANDARD_VERSION_MAJOR: u16 = 3;
const STANDARD_VERSION_MINOR: u16 = 2;

/// `EncryptionHeader` fixed-size fields are 8 DWORDs (32 bytes).
const ENCRYPTION_HEADER_FIXED_LEN: usize = 8 * 4;

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
    #[error("invalid CSPName byte length: {len} (must be even for UTF-16LE)")]
    InvalidCspNameLength { len: usize },
    #[error("invalid CSPName: missing UTF-16LE NUL terminator")]
    InvalidCspNameMissingTerminator,
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
    pub(crate) size_extra: u32,
    pub(crate) alg_id: u32,
    pub(crate) alg_id_hash: u32,
    pub(crate) key_size: u32,
    pub(crate) provider_type: u32,
    pub(crate) reserved1: u32,
    pub(crate) reserved2: u32,
    pub(crate) csp_name: String,
    pub(crate) header_extra: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct EncryptionVerifier {
    pub(crate) salt_size: u32,
    pub(crate) salt: Vec<u8>,
    pub(crate) encrypted_verifier: Vec<u8>,
    pub(crate) verifier_hash_size: u32,
    pub(crate) encrypted_verifier_hash: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct StandardEncryptionInfo {
    pub(crate) header_size: u32,
    pub(crate) header: StandardEncryptionHeader,
    pub(crate) verifier: EncryptionVerifier,
}

/// Parse a Standard `EncryptionHeader` and validate key flags.
fn parse_standard_encryption_header(
    buf: &[u8],
) -> Result<StandardEncryptionHeader, EncryptionInfoError> {
    // `EncryptionHeader` fixed-size fields are 8 DWORDs, followed by variable-length
    // CSPName bytes and trailing sizeExtra bytes.
    if buf.len() < ENCRYPTION_HEADER_FIXED_LEN {
        return Err(EncryptionInfoError::Truncated);
    }

    let mut offset = 0usize;
    let flags_raw = read_u32_le(buf, &mut offset)?;
    let flags = EncryptionHeaderFlags::from_raw(flags_raw);
    let size_extra = read_u32_le(buf, &mut offset)?;
    let alg_id = read_u32_le(buf, &mut offset)?;
    let alg_id_hash = read_u32_le(buf, &mut offset)?;
    let key_size = read_u32_le(buf, &mut offset)?;
    let provider_type = read_u32_le(buf, &mut offset)?;
    let reserved1 = read_u32_le(buf, &mut offset)?;
    let reserved2 = read_u32_le(buf, &mut offset)?;

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

    let tail = &buf[ENCRYPTION_HEADER_FIXED_LEN..];
    let size_extra_usize = size_extra as usize;
    if size_extra_usize > tail.len() {
        return Err(EncryptionInfoError::InvalidEncryptionHeaderSize(
            buf.len() as u32,
        ));
    }

    let csp_bytes = &tail[..tail.len() - size_extra_usize];
    if csp_bytes.len() % 2 != 0 {
        return Err(EncryptionInfoError::InvalidCspNameLength {
            len: csp_bytes.len(),
        });
    }
    // `CSPName` must be NUL-terminated (`0x0000`) in UTF-16LE.
    if !csp_bytes.chunks_exact(2).any(|pair| pair == [0x00, 0x00]) {
        return Err(EncryptionInfoError::InvalidCspNameMissingTerminator);
    }

    let (cow, _) = UTF_16LE.decode_without_bom_handling(csp_bytes);
    let csp_name = cow.into_owned();

    let header_extra = tail[tail.len() - size_extra_usize..].to_vec();
    debug_assert_eq!(header_extra.len(), size_extra_usize);

    Ok(StandardEncryptionHeader {
        flags,
        size_extra,
        alg_id,
        alg_id_hash,
        key_size,
        provider_type,
        reserved1,
        reserved2,
        csp_name,
        header_extra,
    })
}

/// Parse an MS-OFFCRYPTO `EncryptionInfo` stream as Standard encryption.
pub(crate) fn parse_standard_encryption_info(
    bytes: &[u8],
) -> Result<StandardEncryptionInfo, EncryptionInfoError> {
    let mut offset = 0usize;
    let major = read_u16_le(bytes, &mut offset)?;
    let minor = read_u16_le(bytes, &mut offset)?;

    if minor != STANDARD_VERSION_MINOR || !(2..=4).contains(&major) {
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

    // The header must contain at least the fixed 8 DWORD fields plus a UTF-16LE NUL terminator for
    // the (possibly empty) CSPName string.
    if header_size_usize < ENCRYPTION_HEADER_FIXED_LEN + 2 {
        return Err(EncryptionInfoError::InvalidEncryptionHeaderSize(header_size));
    }

    let header_end = offset
        .checked_add(header_size_usize)
        .ok_or(EncryptionInfoError::Truncated)?;
    let header = bytes
        .get(offset..header_end)
        .ok_or(EncryptionInfoError::Truncated)?;
    offset += header_size_usize;

    let parsed_header = parse_standard_encryption_header(header)?;

    // `EncryptionVerifier` begins with:
    //   DWORD SaltSize;
    //   BYTE  Salt[SaltSize];
    //   BYTE  EncryptedVerifier[16];
    //   DWORD VerifierHashSize;
    //   BYTE  EncryptedVerifierHash[...]
    let salt_size = read_u32_le(bytes, &mut offset)?;
    let salt_size_usize = salt_size as usize;
    let salt_end = offset
        .checked_add(salt_size_usize)
        .ok_or(EncryptionInfoError::Truncated)?;
    let salt = bytes
        .get(offset..salt_end)
        .ok_or(EncryptionInfoError::Truncated)?
        .to_vec();
    offset += salt_size_usize;

    let encrypted_verifier_end = offset
        .checked_add(16)
        .ok_or(EncryptionInfoError::Truncated)?;
    let encrypted_verifier = bytes
        .get(offset..encrypted_verifier_end)
        .ok_or(EncryptionInfoError::Truncated)?
        .to_vec();
    offset += 16;

    let verifier_hash_size = read_u32_le(bytes, &mut offset)?;
    let verifier_hash_size_usize = verifier_hash_size as usize;
    let expected_hash_len = if is_aes_alg_id(parsed_header.alg_id) {
        // AES is a block cipher (16-byte blocks).
        verifier_hash_size_usize
            .checked_add(15)
            .map(|v| (v / 16) * 16)
            .ok_or(EncryptionInfoError::Truncated)?
    } else {
        // RC4 is stream-based (no padding).
        verifier_hash_size_usize
    };
    let verifier_hash_end = offset
        .checked_add(expected_hash_len)
        .ok_or(EncryptionInfoError::Truncated)?;
    let encrypted_verifier_hash = bytes
        .get(offset..verifier_hash_end)
        .ok_or(EncryptionInfoError::Truncated)?
        .to_vec();

    Ok(StandardEncryptionInfo {
        header_size,
        header: parsed_header,
        verifier: EncryptionVerifier {
            salt_size,
            salt,
            encrypted_verifier,
            verifier_hash_size,
            encrypted_verifier_hash,
        },
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn build_standard_encryption_info(encryption_header_flags: u32, alg_id: u32) -> Vec<u8> {
        // Build a minimal Standard EncryptionInfo buffer:
        // - Version (3.2)
        // - Info flags (0)
        // - Header size
        // - EncryptionHeader (fixed fields + empty CSPName bytes)
        // - EncryptionVerifier (minimal lengths)

        let mut out = Vec::new();
        out.extend_from_slice(&STANDARD_VERSION_MAJOR.to_le_bytes());
        out.extend_from_slice(&STANDARD_VERSION_MINOR.to_le_bytes());
        out.extend_from_slice(&0u32.to_le_bytes()); // EncryptionInfo.Flags (ignored)

        // Header: 8 DWORDs + 2 bytes of empty UTF-16LE (one NUL code unit).
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
        out.extend_from_slice(&0u16.to_le_bytes()); // CSPName (empty string / terminator)

        // EncryptionVerifier.
        out.extend_from_slice(&16u32.to_le_bytes()); // SaltSize
        out.extend_from_slice(&[0u8; 16]); // Salt
        out.extend_from_slice(&[0u8; 16]); // EncryptedVerifier
        out.extend_from_slice(&20u32.to_le_bytes()); // VerifierHashSize (SHA-1)
        let hash_len = if is_aes_alg_id(alg_id) { 32 } else { 20 };
        out.extend_from_slice(&vec![0u8; hash_len]); // EncryptedVerifierHash (padded for AES)

        out
    }

    fn utf16le_bytes(s: &str) -> Vec<u8> {
        s.encode_utf16().flat_map(|cu| cu.to_le_bytes()).collect()
    }

    #[test]
    fn standard_parses_csp_name_and_skips_size_extra_bytes() {
        // Build a synthetic Standard EncryptionInfo buffer with a non-zero sizeExtra and confirm:
        // - CSPName is derived from headerSize - 32 - sizeExtra bytes
        // - verifier fields are parsed starting *after* the extra bytes
        let flags = EncryptionHeaderFlags::F_CRYPTOAPI;
        let alg_id = CALG_RC4;

        let csp_name = "Test CSP\0";
        let csp_name_bytes = utf16le_bytes(csp_name);
        let header_extra = vec![0xAA, 0xBB, 0xCC, 0xDD];
        let size_extra = header_extra.len() as u32;

        let header_size = (8 * 4 + csp_name_bytes.len() + header_extra.len()) as u32;

        let salt_size = 16u32;
        let salt: Vec<u8> = (0u8..16).collect();
        let encrypted_verifier: Vec<u8> = (0x10u8..0x20).collect();
        let verifier_hash_size = 20u32;
        let encrypted_verifier_hash = vec![0xEE; 20];

        let mut out = Vec::new();
        out.extend_from_slice(&STANDARD_VERSION_MAJOR.to_le_bytes());
        out.extend_from_slice(&STANDARD_VERSION_MINOR.to_le_bytes());
        out.extend_from_slice(&0u32.to_le_bytes()); // EncryptionInfo.Flags
        out.extend_from_slice(&header_size.to_le_bytes());

        // EncryptionHeader fixed fields.
        out.extend_from_slice(&flags.to_le_bytes()); // Flags
        out.extend_from_slice(&size_extra.to_le_bytes()); // SizeExtra
        out.extend_from_slice(&alg_id.to_le_bytes()); // AlgId
        out.extend_from_slice(&0u32.to_le_bytes()); // AlgIdHash
        out.extend_from_slice(&128u32.to_le_bytes()); // KeySize
        out.extend_from_slice(&0u32.to_le_bytes()); // ProviderType
        out.extend_from_slice(&0u32.to_le_bytes()); // Reserved1
        out.extend_from_slice(&0u32.to_le_bytes()); // Reserved2
        out.extend_from_slice(&csp_name_bytes);
        out.extend_from_slice(&header_extra);

        // EncryptionVerifier.
        out.extend_from_slice(&salt_size.to_le_bytes());
        out.extend_from_slice(&salt);
        out.extend_from_slice(&encrypted_verifier);
        out.extend_from_slice(&verifier_hash_size.to_le_bytes());
        out.extend_from_slice(&encrypted_verifier_hash);

        let parsed = parse_standard_encryption_info(&out).expect("should parse");
        assert_eq!(parsed.header_size, header_size);
        assert_eq!(parsed.header.size_extra, size_extra);
        assert_eq!(parsed.header.csp_name, csp_name);
        assert_eq!(parsed.header.header_extra, header_extra);

        assert_eq!(parsed.verifier.salt_size, salt_size);
        assert_eq!(parsed.verifier.salt, salt);
        assert_eq!(parsed.verifier.encrypted_verifier, encrypted_verifier);
        assert_eq!(parsed.verifier.verifier_hash_size, verifier_hash_size);
        assert_eq!(parsed.verifier.encrypted_verifier_hash, encrypted_verifier_hash);
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
        let info = parse_standard_encryption_info(&bytes).expect("should parse");
        assert_eq!(info.header.flags.f_cryptoapi, true);
        assert_eq!(info.header.flags.f_doc_props, true);
        assert_eq!(info.header.flags.f_external, false);
        assert_eq!(info.header.flags.f_aes, false);
        assert_eq!(info.header.alg_id, CALG_RC4);
    }

    #[test]
    fn standard_accepts_cryptoapi_aes_flags() {
        let flags = EncryptionHeaderFlags::F_CRYPTOAPI | EncryptionHeaderFlags::F_AES;
        let bytes = build_standard_encryption_info(flags, CALG_AES_128);
        let info = parse_standard_encryption_info(&bytes).expect("should parse");
        assert_eq!(info.header.flags.f_cryptoapi, true);
        assert_eq!(info.header.flags.f_aes, true);
        assert_eq!(info.header.alg_id, CALG_AES_128);
    }

    #[test]
    fn standard_rejects_odd_csp_name_byte_length() {
        // headerSize=35 => CSPName byte length = 3 (odd) when sizeExtra=0.
        let header_size = 35u32;
        let mut out = Vec::new();
        out.extend_from_slice(&STANDARD_VERSION_MAJOR.to_le_bytes());
        out.extend_from_slice(&STANDARD_VERSION_MINOR.to_le_bytes());
        out.extend_from_slice(&0u32.to_le_bytes()); // EncryptionInfo.Flags
        out.extend_from_slice(&header_size.to_le_bytes());

        // EncryptionHeader fixed fields.
        out.extend_from_slice(&EncryptionHeaderFlags::F_CRYPTOAPI.to_le_bytes()); // Flags
        out.extend_from_slice(&0u32.to_le_bytes()); // SizeExtra
        out.extend_from_slice(&CALG_RC4.to_le_bytes()); // AlgId
        out.extend_from_slice(&0u32.to_le_bytes()); // AlgIdHash
        out.extend_from_slice(&128u32.to_le_bytes()); // KeySize
        out.extend_from_slice(&0u32.to_le_bytes()); // ProviderType
        out.extend_from_slice(&0u32.to_le_bytes()); // Reserved1
        out.extend_from_slice(&0u32.to_le_bytes()); // Reserved2

        // 3 bytes of CSPName (invalid UTF-16LE length).
        out.extend_from_slice(&[0x41, 0x00, 0x42]);

        let err = parse_standard_encryption_info(&out).expect_err("expected error");
        assert!(
            matches!(err, EncryptionInfoError::InvalidCspNameLength { len: 3 }),
            "unexpected error: {err:?}"
        );
    }

    #[test]
    fn standard_accepts_version_major_2_and_4() {
        let flags = EncryptionHeaderFlags::F_CRYPTOAPI | EncryptionHeaderFlags::F_AES;
        let bytes = build_standard_encryption_info(flags, CALG_AES_128);

        for major in [2u16, 4u16] {
            let mut patched = bytes.clone();
            patched[..2].copy_from_slice(&major.to_le_bytes());
            let info = parse_standard_encryption_info(&patched).expect("should parse");
            assert_eq!(info.header.flags.f_cryptoapi, true);
            assert_eq!(info.header.flags.f_aes, true);
            assert_eq!(info.header.alg_id, CALG_AES_128);
        }
    }

    #[test]
    fn standard_accepts_odd_total_header_size_when_size_extra_is_odd() {
        // Some real-world Standard headers include `sizeExtra` bytes that are not UTF-16LE and can
        // make the overall `headerSize` odd. The parser must validate CSPName only over
        // `headerSize - 32 - sizeExtra` bytes.
        let flags = EncryptionHeaderFlags::F_CRYPTOAPI;
        let alg_id = CALG_RC4;

        let csp_name = "Odd sizeExtra CSP\0";
        let csp_name_bytes = utf16le_bytes(csp_name);
        assert_eq!(csp_name_bytes.len() % 2, 0);

        let header_extra = vec![0xAB]; // sizeExtra = 1 (odd)
        let header_size = (ENCRYPTION_HEADER_FIXED_LEN + csp_name_bytes.len() + header_extra.len()) as u32;
        assert_eq!(header_size % 2, 1, "expected odd headerSize");

        let mut out = Vec::new();
        out.extend_from_slice(&STANDARD_VERSION_MAJOR.to_le_bytes());
        out.extend_from_slice(&STANDARD_VERSION_MINOR.to_le_bytes());
        out.extend_from_slice(&0u32.to_le_bytes()); // EncryptionInfo.Flags
        out.extend_from_slice(&header_size.to_le_bytes());

        // EncryptionHeader fixed fields.
        out.extend_from_slice(&flags.to_le_bytes()); // Flags
        out.extend_from_slice(&(header_extra.len() as u32).to_le_bytes()); // SizeExtra
        out.extend_from_slice(&alg_id.to_le_bytes()); // AlgId
        out.extend_from_slice(&0u32.to_le_bytes()); // AlgIdHash
        out.extend_from_slice(&128u32.to_le_bytes()); // KeySize
        out.extend_from_slice(&0u32.to_le_bytes()); // ProviderType
        out.extend_from_slice(&0u32.to_le_bytes()); // Reserved1
        out.extend_from_slice(&0u32.to_le_bytes()); // Reserved2
        out.extend_from_slice(&csp_name_bytes);
        out.extend_from_slice(&header_extra);

        // Minimal EncryptionVerifier.
        out.extend_from_slice(&16u32.to_le_bytes()); // SaltSize
        out.extend_from_slice(&[0u8; 16]); // Salt
        out.extend_from_slice(&[0u8; 16]); // EncryptedVerifier
        out.extend_from_slice(&20u32.to_le_bytes()); // VerifierHashSize
        out.extend_from_slice(&[0u8; 20]); // EncryptedVerifierHash (RC4 exact length)

        let parsed = parse_standard_encryption_info(&out).expect("should parse");
        assert_eq!(parsed.header_size, header_size);
        assert_eq!(parsed.header.size_extra, 1);
        assert_eq!(parsed.header.csp_name, csp_name);
        assert_eq!(parsed.header.header_extra, header_extra);
    }

    #[test]
    fn standard_rejects_size_extra_larger_than_header_tail() {
        // headerSize includes CSPName bytes + sizeExtra. If sizeExtra exceeds the remaining tail
        // bytes after the fixed 32-byte header fields, the header is internally inconsistent and
        // must be rejected without panicking.
        let header_size = (ENCRYPTION_HEADER_FIXED_LEN + 2) as u32; // fixed fields + UTF-16LE NUL

        let mut out = Vec::new();
        out.extend_from_slice(&STANDARD_VERSION_MAJOR.to_le_bytes());
        out.extend_from_slice(&STANDARD_VERSION_MINOR.to_le_bytes());
        out.extend_from_slice(&0u32.to_le_bytes()); // EncryptionInfo.Flags
        out.extend_from_slice(&header_size.to_le_bytes());

        // EncryptionHeader fixed fields.
        out.extend_from_slice(&EncryptionHeaderFlags::F_CRYPTOAPI.to_le_bytes()); // Flags
        out.extend_from_slice(&3u32.to_le_bytes()); // SizeExtra (tail len is only 2)
        out.extend_from_slice(&CALG_RC4.to_le_bytes()); // AlgId
        out.extend_from_slice(&0u32.to_le_bytes()); // AlgIdHash
        out.extend_from_slice(&128u32.to_le_bytes()); // KeySize
        out.extend_from_slice(&0u32.to_le_bytes()); // ProviderType
        out.extend_from_slice(&0u32.to_le_bytes()); // Reserved1
        out.extend_from_slice(&0u32.to_le_bytes()); // Reserved2
        out.extend_from_slice(&0u16.to_le_bytes()); // CSPName terminator (2 bytes)

        let err = parse_standard_encryption_info(&out).expect_err("expected error");
        assert!(
            matches!(
                err,
                EncryptionInfoError::Truncated | EncryptionInfoError::InvalidEncryptionHeaderSize(_)
            ),
            "unexpected error: {err:?}"
        );
    }
}
