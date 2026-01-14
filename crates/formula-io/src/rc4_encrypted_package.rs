//! Helpers for parsing the `EncryptedPackage` stream header used by Office "Standard" / CryptoAPI
//! RC4 encryption (MS-OFFCRYPTO).
//!
//! The `EncryptedPackage` stream begins with an **8-byte plaintext** size header, followed by the
//! encrypted OOXML package bytes.
//!
//! Real-world producers disagree on how to interpret the 8-byte prefix:
//! - `u32 totalSize` + `u32 reserved` (where `reserved` is often `0`), or
//! - a single `u64` little-endian size.
//!
//! To remain compatible, we parse the prefix as a 64-bit size split across two little-endian DWORDs:
//!
//! ```text
//! lo = u32le(bytes[0..4])
//! hi = u32le(bytes[4..8])
//! package_size = lo as u64 | ((hi as u64) << 32)
//! ```
//!
//! MS-OFFCRYPTO describes this field as a `u64le`, but some producers/libraries treat it as
//! `(u32 totalSize, u32 reserved)`. When the high DWORD is non-zero and the combined 64-bit value
//! is not plausible for the available ciphertext, we fall back to the low DWORD **only when it is
//! non-zero** (so we don't misinterpret true 64-bit sizes that are exact multiples of `2^32`).
//!
//! This module is intentionally self-contained and does *not* implement RC4 decryption; it only
//! parses and validates the size header. The decryption logic should treat the first 8 bytes as
//! plaintext and begin RC4 block processing at offset 8.
#![allow(dead_code)]

/// Plaintext header length (in bytes) at the start of an RC4 `EncryptedPackage` stream.
pub(crate) const ENCRYPTED_PACKAGE_SIZE_HEADER_LEN: usize = 8;

/// Conservative default max size (bytes) for the decrypted OOXML package.
///
/// This is a safety belt against malformed headers that could otherwise cause OOM when callers
/// allocate based on the declared package size.
const fn min_u64(a: u64, b: u64) -> u64 {
    if a < b { a } else { b }
}

pub(crate) const DEFAULT_MAX_DECRYPTED_PACKAGE_SIZE: u64 = min_u64(
    2 * 1024 * 1024 * 1024, // 2GiB
    isize::MAX as u64,      // avoid `Vec`/allocation invariants on 32-bit targets
);

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Rc4EncryptedPackageParseOptions {
    /// When enabled, return a warning when the high DWORD is non-zero (some specs treat it as a
    /// reserved field that should be 0).
    pub(crate) strict: bool,
    /// Hard cap on the declared decrypted package size.
    pub(crate) max_decrypted_package_size: u64,
}

impl Default for Rc4EncryptedPackageParseOptions {
    fn default() -> Self {
        Self {
            strict: false,
            max_decrypted_package_size: DEFAULT_MAX_DECRYPTED_PACKAGE_SIZE,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct Rc4EncryptedPackageSizeHeader {
    pub(crate) package_size: u64,
    pub(crate) lo: u32,
    pub(crate) hi: u32,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct ParsedRc4EncryptedPackage<'a> {
    pub(crate) header: Rc4EncryptedPackageSizeHeader,
    /// Encrypted OOXML package bytes (does **not** include the plaintext 8-byte size header).
    pub(crate) encrypted_payload: &'a [u8],
    /// Best-effort parse warnings (only emitted in `strict` mode).
    pub(crate) warnings: Vec<String>,
}

#[derive(Debug, thiserror::Error, Clone, PartialEq, Eq)]
pub(crate) enum Rc4EncryptedPackageParseError {
    #[error("EncryptedPackage stream is truncated (missing 8-byte size header)")]
    TruncatedHeader,
    #[error(
        "EncryptedPackage declared size {declared} bytes exceeds encrypted payload length {available} bytes"
    )]
    DeclaredSizeExceedsPayload { declared: u64, available: u64 },
    #[error("EncryptedPackage declared size {declared} bytes exceeds configured maximum {max} bytes")]
    DeclaredSizeExceedsMax { declared: u64, max: u64 },
}

fn parse_u32_le(bytes: &[u8]) -> u32 {
    u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]])
}

/// Parse the 8-byte plaintext size header for an RC4 `EncryptedPackage` stream.
///
/// `available_payload_len` is the number of bytes following the 8-byte header (i.e. the encrypted
/// package byte length).
pub(crate) fn parse_rc4_encrypted_package_size_header(
    header: &[u8; ENCRYPTED_PACKAGE_SIZE_HEADER_LEN],
    available_payload_len: u64,
    opts: &Rc4EncryptedPackageParseOptions,
) -> Result<(Rc4EncryptedPackageSizeHeader, Vec<String>), Rc4EncryptedPackageParseError> {
    let lo = parse_u32_le(&header[0..4]);
    let hi = parse_u32_le(&header[4..8]);
    let package_size_raw = (lo as u64) | ((hi as u64) << 32);
    let package_size = crate::parse_encrypted_package_size_prefix_bytes(*header, Some(available_payload_len));

    // Warnings are best-effort diagnostics; they don't influence parsing unless
    // strict callers choose to treat them as fatal.
    let mut warnings = Vec::new();
    if opts.strict && hi != 0 {
        // Many specs/libraries treat the high DWORD as a reserved field that MUST be 0. In the
        // wild, some producers appear to store a true u64 size instead. Warn so strict callers can
        // decide whether to accept it.
        if package_size != package_size_raw {
            warnings.push(format!(
                "EncryptedPackage size header high DWORD is non-zero ({hi}); treating high DWORD as reserved and using low DWORD size {lo}"
            ));
        } else {
            warnings.push(format!(
                "EncryptedPackage size header high DWORD is non-zero ({hi}); treating header as 64-bit size"
            ));
        }
    }

    // Safety checks: avoid pathological allocations and truncated reads.
    if package_size > opts.max_decrypted_package_size {
        return Err(Rc4EncryptedPackageParseError::DeclaredSizeExceedsMax {
            declared: package_size,
            max: opts.max_decrypted_package_size,
        });
    }
    if package_size > available_payload_len {
        return Err(Rc4EncryptedPackageParseError::DeclaredSizeExceedsPayload {
            declared: package_size,
            available: available_payload_len,
        });
    }

    Ok((
        Rc4EncryptedPackageSizeHeader {
            package_size,
            lo,
            hi,
        },
        warnings,
    ))
}

/// Parse an in-memory RC4 `EncryptedPackage` stream as `(header, encrypted_payload)`.
///
/// This function guarantees that the plaintext header is **not** treated as part of the encrypted
/// payload.
pub(crate) fn parse_rc4_encrypted_package_stream<'a>(
    encrypted_package_stream: &'a [u8],
    opts: &Rc4EncryptedPackageParseOptions,
) -> Result<ParsedRc4EncryptedPackage<'a>, Rc4EncryptedPackageParseError> {
    let header_bytes = encrypted_package_stream
        .get(..ENCRYPTED_PACKAGE_SIZE_HEADER_LEN)
        .ok_or(Rc4EncryptedPackageParseError::TruncatedHeader)?;
    let header_bytes: &[u8; ENCRYPTED_PACKAGE_SIZE_HEADER_LEN] = header_bytes
        .try_into()
        .map_err(|_| Rc4EncryptedPackageParseError::TruncatedHeader)?;

    let encrypted_payload = &encrypted_package_stream[ENCRYPTED_PACKAGE_SIZE_HEADER_LEN..];
    let available_payload_len = encrypted_payload.len() as u64;

    let (header, warnings) =
        parse_rc4_encrypted_package_size_header(header_bytes, available_payload_len, opts)?;

    Ok(ParsedRc4EncryptedPackage {
        header,
        encrypted_payload,
        warnings,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn default_max_decrypted_package_size_respects_isize_max() {
        let two_gib = 2u64 * 1024 * 1024 * 1024;
        let expected = if (isize::MAX as u64) < two_gib {
            isize::MAX as u64
        } else {
            two_gib
        };
        assert_eq!(DEFAULT_MAX_DECRYPTED_PACKAGE_SIZE, expected);
    }

    #[test]
    fn rc4_encrypted_package_size_header_hi_zero() {
        let opts = Rc4EncryptedPackageParseOptions {
            strict: true,
            max_decrypted_package_size: 10_000,
        };

        let lo = 1234u32;
        let hi = 0u32;
        let mut header = [0u8; ENCRYPTED_PACKAGE_SIZE_HEADER_LEN];
        header[0..4].copy_from_slice(&lo.to_le_bytes());
        header[4..8].copy_from_slice(&hi.to_le_bytes());

        let (parsed, warnings) =
            parse_rc4_encrypted_package_size_header(&header, 5000, &opts).expect("parse header");
        assert_eq!(parsed.package_size, 1234);
        assert_eq!(parsed.lo, lo);
        assert_eq!(parsed.hi, hi);
        assert!(warnings.is_empty());
    }

    #[test]
    fn rc4_encrypted_package_size_header_hi_nonzero_bounds_check_and_arithmetic() {
        let opts = Rc4EncryptedPackageParseOptions {
            strict: true,
            // Intentionally tiny to trigger the max-size sanity bound.
            max_decrypted_package_size: 1024,
        };

        let lo = 0u32;
        let hi = 1u32;
        let expected = (lo as u64) | ((hi as u64) << 32);
        assert_eq!(expected, 1u64 << 32);

        let mut header = [0u8; ENCRYPTED_PACKAGE_SIZE_HEADER_LEN];
        header[0..4].copy_from_slice(&lo.to_le_bytes());
        header[4..8].copy_from_slice(&hi.to_le_bytes());

        let err =
            parse_rc4_encrypted_package_size_header(&header, u64::MAX, &opts).expect_err("too big");
        match err {
            Rc4EncryptedPackageParseError::DeclaredSizeExceedsMax { declared, max } => {
                assert_eq!(declared, expected, "u64 arithmetic must be correct");
                assert_eq!(max, opts.max_decrypted_package_size);
            }
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn rc4_encrypted_package_size_header_hi_nonzero_emits_warning_in_strict_mode() {
        let lo = 0u32;
        let hi = 1u32;
        let expected = (lo as u64) | ((hi as u64) << 32);

        // Allow the large declared size so we can observe the warning.
        let opts = Rc4EncryptedPackageParseOptions {
            strict: true,
            max_decrypted_package_size: expected,
        };

        let mut header = [0u8; ENCRYPTED_PACKAGE_SIZE_HEADER_LEN];
        header[0..4].copy_from_slice(&lo.to_le_bytes());
        header[4..8].copy_from_slice(&hi.to_le_bytes());

        let (_parsed, warnings) =
            parse_rc4_encrypted_package_size_header(&header, expected, &opts).expect("parse header");
        assert_eq!(warnings.len(), 1);
        assert!(warnings[0].contains("high DWORD is non-zero"));
    }

    #[test]
    fn rc4_encrypted_package_stream_excludes_plaintext_header_from_payload() {
        let opts = Rc4EncryptedPackageParseOptions {
            strict: false,
            max_decrypted_package_size: 16,
        };

        let mut stream = Vec::new();
        stream.extend_from_slice(&4u32.to_le_bytes()); // lo
        stream.extend_from_slice(&0u32.to_le_bytes()); // hi
        stream.extend_from_slice(&[1u8, 2, 3, 4]); // encrypted bytes (dummy)

        let parsed = parse_rc4_encrypted_package_stream(&stream, &opts).expect("parse stream");
        assert_eq!(parsed.header.package_size, 4);
        assert_eq!(parsed.encrypted_payload, &[1, 2, 3, 4]);
    }

    #[test]
    fn rc4_encrypted_package_size_header_falls_back_to_low_dword_when_high_dword_is_reserved() {
        let opts = Rc4EncryptedPackageParseOptions {
            strict: true,
            max_decrypted_package_size: 10_000,
        };

        // Producer writes (lo=size, hi=reserved) and sets reserved to a non-zero value.
        let lo = 1234u32;
        let hi = 1u32;
        let mut header = [0u8; ENCRYPTED_PACKAGE_SIZE_HEADER_LEN];
        header[0..4].copy_from_slice(&lo.to_le_bytes());
        header[4..8].copy_from_slice(&hi.to_le_bytes());

        // Available payload is too small for the combined u64, but large enough for the low DWORD.
        let (parsed, warnings) =
            parse_rc4_encrypted_package_size_header(&header, 5000, &opts).expect("parse header");
        assert_eq!(parsed.package_size, lo as u64);
        assert_eq!(parsed.lo, lo);
        assert_eq!(parsed.hi, hi);
        assert_eq!(warnings.len(), 1);
        assert!(warnings[0].contains("reserved"));
    }
}
