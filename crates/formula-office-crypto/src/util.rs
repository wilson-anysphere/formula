use crate::error::OfficeCryptoError;

use subtle::ConstantTimeEq as _;

#[cfg(test)]
use std::sync::atomic::{AtomicUsize, Ordering};

#[cfg(test)]
static CT_EQ_CALLS: AtomicUsize = AtomicUsize::new(0);

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum EncryptionInfoKind {
    Agile,
    Standard,
}

#[derive(Debug, Clone)]
pub(crate) struct EncryptionInfoHeader {
    pub(crate) version_major: u16,
    pub(crate) version_minor: u16,
    pub(crate) flags: u32,
    pub(crate) header_size: u32,
    pub(crate) kind: EncryptionInfoKind,
    pub(crate) header_offset: usize,
}

pub(crate) fn parse_encryption_info_header(
    bytes: &[u8],
) -> Result<EncryptionInfoHeader, OfficeCryptoError> {
    if bytes.len() < 8 {
        return Err(OfficeCryptoError::InvalidFormat(
            "EncryptionInfo stream too short".to_string(),
        ));
    }
    let version_major = read_u16_le(bytes, 0)?;
    let version_minor = read_u16_le(bytes, 2)?;
    let flags = read_u32_le(bytes, 4)?;

    // MS-OFFCRYPTO / ECMA-376 identifies "Standard" encryption via `versionMinor == 2`, but
    // real-world files vary `versionMajor` across Office generations (2/3/4).
    //
    // "Extensible" encryption uses `versionMinor == 3` with `versionMajor` 3 or 4.
    match (version_major, version_minor) {
        (4, 4) => {
            // Agile (XML) EncryptionInfo.
            //
            // Some producers include a 4-byte XML length prefix after the 8-byte
            // `EncryptionVersionInfo` header. Others store the XML document directly starting at
            // byte offset 8. Accept both forms.
            let (header_offset, header_size) = if bytes.len() >= 12 {
                let candidate = read_u32_le(bytes, 8)?;
                let available = bytes.len().saturating_sub(12);
                if (candidate as usize) <= available {
                    (12usize, candidate)
                } else {
                    (8usize, bytes.len().saturating_sub(8) as u32)
                }
            } else {
                (8usize, bytes.len().saturating_sub(8) as u32)
            };

            Ok(EncryptionInfoHeader {
                version_major,
                version_minor,
                flags,
                header_size,
                kind: EncryptionInfoKind::Agile,
                header_offset,
            })
        }
        (major, 2) if (2..=4).contains(&major) => {
            if bytes.len() < 12 {
                return Err(OfficeCryptoError::InvalidFormat(
                    "EncryptionInfo stream too short".to_string(),
                ));
            }
            let header_size = read_u32_le(bytes, 8)?;
            Ok(EncryptionInfoHeader {
                version_major,
                version_minor,
                flags,
                header_size,
                kind: EncryptionInfoKind::Standard,
                header_offset: 12,
            })
        }
        _ => Err(OfficeCryptoError::UnsupportedEncryption(format!(
            "unsupported EncryptionInfo version {version_major}.{version_minor} (flags={flags:#x})"
        ))),
    }
}

pub(crate) fn read_u16_le(bytes: &[u8], offset: usize) -> Result<u16, OfficeCryptoError> {
    let b = bytes
        .get(offset..offset + 2)
        .ok_or_else(|| OfficeCryptoError::InvalidFormat("unexpected EOF".to_string()))?;
    Ok(u16::from_le_bytes([b[0], b[1]]))
}

pub(crate) fn read_u32_le(bytes: &[u8], offset: usize) -> Result<u32, OfficeCryptoError> {
    let b = bytes
        .get(offset..offset + 4)
        .ok_or_else(|| OfficeCryptoError::InvalidFormat("unexpected EOF".to_string()))?;
    Ok(u32::from_le_bytes([b[0], b[1], b[2], b[3]]))
}

pub(crate) fn read_u64_le(bytes: &[u8], offset: usize) -> Result<u64, OfficeCryptoError> {
    let b = bytes
        .get(offset..offset + 8)
        .ok_or_else(|| OfficeCryptoError::InvalidFormat("unexpected EOF".to_string()))?;
    Ok(u64::from_le_bytes([
        b[0], b[1], b[2], b[3], b[4], b[5], b[6], b[7],
    ]))
}

pub(crate) fn decode_utf16le_nul_terminated(bytes: &[u8]) -> Result<String, OfficeCryptoError> {
    if bytes.len() % 2 != 0 {
        return Err(OfficeCryptoError::InvalidFormat(
            "UTF-16LE string has odd length".to_string(),
        ));
    }

    let mut code_units: Vec<u16> = Vec::new();
    for pair in bytes.chunks_exact(2) {
        let cu = u16::from_le_bytes([pair[0], pair[1]]);
        if cu == 0 {
            break;
        }
        code_units.push(cu);
    }
    String::from_utf16(&code_units).map_err(|_| {
        OfficeCryptoError::InvalidFormat("invalid UTF-16LE in EncryptionInfo".to_string())
    })
}

pub(crate) fn checked_vec_len(total_size: u64) -> Result<usize, OfficeCryptoError> {
    let len = usize::try_from(total_size)
        .map_err(|_| OfficeCryptoError::EncryptedPackageSizeOverflow { total_size })?;

    // `Vec<u8>` cannot exceed `isize::MAX` due to `Layout::array`/pointer offset invariants.
    isize::try_from(len)
        .map_err(|_| OfficeCryptoError::EncryptedPackageSizeOverflow { total_size })?;

    Ok(len)
}

/// Compare two byte slices in constant time.
///
/// This should be used for comparing password verifier digests (both Standard and Agile) to avoid
/// obvious timing side channels from Rust's early-exit `==` / `!=` implementations.
///
/// Note: This does not aim to make the overall decryption flow perfectly side-channel resistant; it
/// only hardens the digest comparison step.
pub(crate) fn ct_eq(a: &[u8], b: &[u8]) -> bool {
    #[cfg(test)]
    CT_EQ_CALLS.fetch_add(1, Ordering::Relaxed);
    bool::from(a.ct_eq(b))
}

#[cfg(test)]
pub(crate) fn reset_ct_eq_calls() {
    CT_EQ_CALLS.store(0, Ordering::Relaxed);
}

#[cfg(test)]
pub(crate) fn ct_eq_call_count() -> usize {
    CT_EQ_CALLS.load(Ordering::Relaxed)
}
