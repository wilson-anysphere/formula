use crate::error::OfficeCryptoError;
use subtle::{Choice, ConstantTimeEq};

#[cfg(test)]
use std::cell::Cell;

// Unit tests run in parallel by default. Use a thread-local counter so tests that reset/inspect
// the counter don't race each other.
#[cfg(test)]
thread_local! {
    static CT_EQ_CALLS: Cell<usize> = Cell::new(0);
}

/// Compare two byte slices in (mostly) constant time.
///
/// This is intended for comparing password verifier digests to reduce timing side channels. Length
/// mismatches return `false` and are handled without panicking or early-exiting on the first
/// mismatched byte.
pub(crate) fn ct_eq(a: &[u8], b: &[u8]) -> bool {
    #[cfg(test)]
    CT_EQ_CALLS.with(|calls| calls.set(calls.get().saturating_add(1)));

    // We treat lengths as non-secret metadata, but still avoid early returns so callers don't
    // accidentally reintroduce short-circuit timing differences.
    let max_len = a.len().max(b.len());
    let mut ok = Choice::from(1u8);
    for i in 0..max_len {
        let av = a.get(i).copied().unwrap_or(0);
        let bv = b.get(i).copied().unwrap_or(0);
        ok &= av.ct_eq(&bv);
    }
    ok &= Choice::from((a.len() == b.len()) as u8);

    bool::from(ok)
}

#[cfg(test)]
pub(crate) fn reset_ct_eq_calls() {
    CT_EQ_CALLS.with(|calls| calls.set(0));
}

#[cfg(test)]
pub(crate) fn ct_eq_call_count() -> usize {
    CT_EQ_CALLS.with(|calls| calls.get())
}

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
    let (kind, header_offset, header_size) = match (version_major, version_minor) {
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
            (EncryptionInfoKind::Agile, header_offset, header_size)
        }
        (major, 2) if (2..=4).contains(&major) => {
            if bytes.len() < 12 {
                return Err(OfficeCryptoError::InvalidFormat(
                    "EncryptionInfo stream too short".to_string(),
                ));
            }
            let header_size = read_u32_le(bytes, 8)?;
            (EncryptionInfoKind::Standard, 12usize, header_size)
        }
        _ => {
            return Err(OfficeCryptoError::UnsupportedEncryption(format!(
                "unsupported EncryptionInfo version {version_major}.{version_minor} (flags={flags:#x})"
            )));
        }
    };

    let header_size_usize = header_size as usize;
    match kind {
        EncryptionInfoKind::Agile => {
            if header_size_usize > crate::MAX_AGILE_ENCRYPTION_INFO_XML_BYTES {
                return Err(OfficeCryptoError::SizeLimitExceeded {
                    context: "EncryptionInfo XML",
                    limit: crate::MAX_AGILE_ENCRYPTION_INFO_XML_BYTES,
                });
            }
        }
        EncryptionInfoKind::Standard => {
            if header_size_usize > crate::MAX_STANDARD_ENCRYPTION_HEADER_BYTES {
                return Err(OfficeCryptoError::SizeLimitExceeded {
                    context: "EncryptionInfo.headerSize",
                    limit: crate::MAX_STANDARD_ENCRYPTION_HEADER_BYTES,
                });
            }
        }
    }

    let end = header_offset.checked_add(header_size_usize).ok_or_else(|| {
        OfficeCryptoError::InvalidFormat("EncryptionInfo header size overflow".to_string())
    })?;
    if end > bytes.len() {
        return Err(OfficeCryptoError::InvalidFormat(
            "EncryptionInfo header size out of range".to_string(),
        ));
    }

    Ok(EncryptionInfoHeader {
        version_major,
        version_minor,
        flags,
        header_size,
        kind,
        header_offset,
    })
}

pub(crate) fn read_u16_le(bytes: &[u8], offset: usize) -> Result<u16, OfficeCryptoError> {
    let end = offset.checked_add(2).ok_or_else(|| {
        OfficeCryptoError::InvalidFormat("offset overflow".to_string())
    })?;
    let b = bytes.get(offset..end).ok_or_else(|| {
        OfficeCryptoError::InvalidFormat("unexpected EOF".to_string())
    })?;
    Ok(u16::from_le_bytes([b[0], b[1]]))
}

pub(crate) fn read_u32_le(bytes: &[u8], offset: usize) -> Result<u32, OfficeCryptoError> {
    let end = offset.checked_add(4).ok_or_else(|| {
        OfficeCryptoError::InvalidFormat("offset overflow".to_string())
    })?;
    let b = bytes.get(offset..end).ok_or_else(|| {
        OfficeCryptoError::InvalidFormat("unexpected EOF".to_string())
    })?;
    Ok(u32::from_le_bytes([b[0], b[1], b[2], b[3]]))
}

/// Parse the 8-byte plaintext size prefix at the start of an `EncryptedPackage` stream.
///
/// MS-OFFCRYPTO describes this field as a `u64le`, but some producers/libraries treat it as
/// `(u32 totalSize, u32 reserved)` (often with `reserved = 0`).
///
/// For compatibility, when the high DWORD is non-zero *and* the combined 64-bit value is not
/// plausible for the available ciphertext, we fall back to the low DWORD **only when it is
/// non-zero** (so we don't misinterpret true 64-bit sizes that are exact multiples of `2^32`).
pub(crate) fn parse_encrypted_package_original_size(
    encrypted_package: &[u8],
) -> Result<u64, OfficeCryptoError> {
    if encrypted_package.len() < 8 {
        return Err(OfficeCryptoError::InvalidFormat(
            "EncryptedPackage stream too short".to_string(),
        ));
    }

    let len_lo = read_u32_le(encrypted_package, 0)? as u64;
    let len_hi = read_u32_le(encrypted_package, 4)? as u64;
    let size_u64 = len_lo | (len_hi << 32);

    let ciphertext_len = encrypted_package.len().saturating_sub(8) as u64;
    Ok(if len_lo != 0 && len_hi != 0 && size_u64 > ciphertext_len && len_lo <= ciphertext_len {
        len_lo
    } else {
        size_u64
    })
}

pub(crate) fn decode_utf16le_nul_terminated(bytes: &[u8]) -> Result<String, OfficeCryptoError> {
    if bytes.len() > crate::MAX_STANDARD_CSPNAME_BYTES {
        return Err(OfficeCryptoError::SizeLimitExceeded {
            context: "EncryptionHeader.cspName",
            limit: crate::MAX_STANDARD_CSPNAME_BYTES,
        });
    }
    if bytes.len() % 2 != 0 {
        return Err(OfficeCryptoError::InvalidFormat(
            "UTF-16LE string has odd length".to_string(),
        ));
    }

    let mut code_units: Vec<u16> = Vec::with_capacity(bytes.len() / 2);
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{MAX_AGILE_ENCRYPTION_INFO_XML_BYTES, MAX_STANDARD_ENCRYPTION_HEADER_BYTES};

    #[test]
    fn encrypted_package_size_header_does_not_fall_back_when_low_dword_is_zero() {
        // The header is specified as a `u64le`. Some producers treat it as `(u32 size, u32 reserved)`,
        // but that heuristic must not misread sizes that are exact multiples of 2^32 (low DWORD = 0).
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&0u32.to_le_bytes()); // low DWORD
        bytes.extend_from_slice(&1u32.to_le_bytes()); // high DWORD
        let size = parse_encrypted_package_original_size(&bytes).expect("parse size");
        assert_eq!(size, 1u64 << 32);
    }

    #[test]
    fn encryption_info_header_rejects_truncated() {
        // Standard encryption requires a 4-byte headerSize field at offset 8.
        // Provide a valid version header but truncate before that field.
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&3u16.to_le_bytes());
        bytes.extend_from_slice(&2u16.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(&[0u8; 3]); // 11 bytes total (< 12)
        let err = parse_encryption_info_header(&bytes).unwrap_err();
        assert!(matches!(err, OfficeCryptoError::InvalidFormat(_)));
    }

    #[test]
    fn encryption_info_header_rejects_unsupported_version() {
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&5u16.to_le_bytes());
        bytes.extend_from_slice(&5u16.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        let err = parse_encryption_info_header(&bytes).unwrap_err();
        assert!(matches!(err, OfficeCryptoError::UnsupportedEncryption(_)));
    }

    #[test]
    fn encryption_info_header_rejects_agile_xml_too_large() {
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&4u16.to_le_bytes());
        bytes.extend_from_slice(&4u16.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        // No-length-prefix form: XML begins directly after the 8-byte version header.
        // Fill with a plausible XML start byte (`<`) to ensure the parser falls back to the
        // no-prefix interpretation.
        bytes.extend(std::iter::repeat(b'<').take(MAX_AGILE_ENCRYPTION_INFO_XML_BYTES + 1));
        let err = parse_encryption_info_header(&bytes).unwrap_err();
        assert!(
            matches!(err, OfficeCryptoError::SizeLimitExceeded { .. }),
            "err={err:?}"
        );
    }

    #[test]
    fn encryption_info_header_rejects_standard_header_too_large() {
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&3u16.to_le_bytes());
        bytes.extend_from_slice(&2u16.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(&((MAX_STANDARD_ENCRYPTION_HEADER_BYTES + 1) as u32).to_le_bytes());
        let err = parse_encryption_info_header(&bytes).unwrap_err();
        assert!(
            matches!(err, OfficeCryptoError::SizeLimitExceeded { .. }),
            "err={err:?}"
        );
    }
}
/// Very lightweight ZIP validator for decrypted OOXML packages.
///
/// This is intentionally stricter than just checking the `PK` prefix: when decrypting with the
/// wrong scheme/key it's not impossible to produce bytes starting with `PK` by chance.
///
/// We look for the End of Central Directory (EOCD) record and validate that it is consistent.
pub(crate) fn looks_like_zip(bytes: &[u8]) -> bool {
    // ZIP local file header signature `PK\x03\x04`.
    if bytes.len() < 4 || &bytes[..2] != b"PK" {
        return false;
    }

    // EOCD signature `PK\x05\x06` is at least 22 bytes from the end, and the comment length is
    // stored in the final 2 bytes of the fixed-size structure.
    const EOCD_LEN: usize = 22;
    const EOCD_SIG: [u8; 4] = [0x50, 0x4B, 0x05, 0x06];

    if bytes.len() < EOCD_LEN {
        return false;
    }

    // Per ZIP spec, comment length is a u16 so EOCD should appear within the last 65535+22 bytes.
    let search_window = EOCD_LEN + 65_535;
    let start = bytes.len().saturating_sub(search_window);

    for i in (start..=bytes.len() - 4).rev() {
        if bytes[i..i + 4] != EOCD_SIG {
            continue;
        }
        // Fixed-size EOCD must fit.
        if i + EOCD_LEN > bytes.len() {
            continue;
        }

        let disk_no = u16::from_le_bytes([bytes[i + 4], bytes[i + 5]]);
        let cd_disk_no = u16::from_le_bytes([bytes[i + 6], bytes[i + 7]]);
        let entries_disk = u16::from_le_bytes([bytes[i + 8], bytes[i + 9]]);
        let entries_total = u16::from_le_bytes([bytes[i + 10], bytes[i + 11]]);
        let cd_size =
            u32::from_le_bytes([bytes[i + 12], bytes[i + 13], bytes[i + 14], bytes[i + 15]])
                as usize;
        let cd_offset =
            u32::from_le_bytes([bytes[i + 16], bytes[i + 17], bytes[i + 18], bytes[i + 19]])
                as usize;
        let comment_len = u16::from_le_bytes([bytes[i + 20], bytes[i + 21]]) as usize;

        // EOCD record should end at EOF (including comment).
        if i + EOCD_LEN + comment_len != bytes.len() {
            continue;
        }

        // Reject multi-disk archives (OOXML packages are not spanned).
        if disk_no != 0 || cd_disk_no != 0 || entries_disk != entries_total {
            continue;
        }

        // Basic bounds checks.
        if cd_offset >= bytes.len() || cd_offset.checked_add(cd_size).is_none() {
            continue;
        }
        if cd_offset + cd_size > i {
            continue;
        }
        // Central directory file header signature `PK\x01\x02`.
        if bytes.get(cd_offset..cd_offset + 4) != Some(b"PK\x01\x02") {
            continue;
        }

        return true;
    }

    false
}
