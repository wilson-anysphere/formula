//! Legacy `.xls` BIFF encryption (FILEPASS) parsing and workbook-stream decryption.
//!
//! This module is currently an internal skeleton: it can classify BIFF `FILEPASS`
//! records and locate the start of encrypted bytes in a workbook stream, but does
//! not yet implement the underlying crypto algorithms.
//!
//! Supported FILEPASS variants (classification only):
//! - BIFF5 XOR obfuscation
//! - BIFF8 XOR obfuscation
//! - BIFF8 RC4
//! - BIFF8 RC4 CryptoAPI

#![allow(dead_code)]

use thiserror::Error;

use super::{records, BiffVersion};

// BIFF8 FILEPASS.wEncryptionType values.
// [MS-XLS] 2.4.117 (FilePass).
const BIFF8_ENCRYPTION_TYPE_XOR: u16 = 0x0000;
const BIFF8_ENCRYPTION_TYPE_RC4: u16 = 0x0001;

// BIFF8 RC4 FILEPASS "subType" values (wEncryptionType==0x0001).
// In BIFF8 these correspond to two different RC4-based layouts.
const BIFF8_RC4_SUBTYPE_RC4: u16 = 0x0001;
const BIFF8_RC4_SUBTYPE_CRYPTOAPI: u16 = 0x0002;

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) enum BiffEncryption {
    /// BIFF5 XOR obfuscation. FILEPASS payload is `key` + `verifier`.
    Biff5Xor { key: u16, verifier: u16 },
    /// BIFF8 XOR obfuscation. FILEPASS payload is `wEncryptionType` + `key` + `verifier`.
    Biff8Xor { key: u16, verifier: u16 },
    /// BIFF8 RC4 encryption (legacy non-CryptoAPI).
    ///
    /// The full FILEPASS payload is preserved for future decryption work.
    Biff8Rc4 { filepass_payload: Vec<u8> },
    /// BIFF8 RC4 encryption using CryptoAPI.
    ///
    /// The full FILEPASS payload is preserved for future decryption work.
    Biff8Rc4CryptoApi { filepass_payload: Vec<u8> },
}

#[derive(Debug, Error, Clone, PartialEq, Eq)]
pub(crate) enum DecryptError {
    #[error("workbook is not encrypted (missing FILEPASS record)")]
    NoFilePass,
    #[error("invalid FILEPASS record: {0}")]
    InvalidFilePass(String),
    #[error("unsupported encryption scheme: {0}")]
    UnsupportedEncryption(String),
    #[error("password required")]
    PasswordRequired,
    #[error("wrong password")]
    WrongPassword,
}

fn read_u16(data: &[u8], offset: usize, ctx: &str) -> Result<u16, DecryptError> {
    let bytes = data.get(offset..offset + 2).ok_or_else(|| {
        DecryptError::InvalidFilePass(format!(
            "truncated FILEPASS payload while reading {ctx} at offset {offset} (len={})",
            data.len()
        ))
    })?;
    Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
}

/// Parse the payload of a `FILEPASS` record (record id `0x002F`).
///
/// The record header is *not* included in `data`.
pub(crate) fn parse_filepass_record(
    biff_version: BiffVersion,
    data: &[u8],
) -> Result<BiffEncryption, DecryptError> {
    match biff_version {
        BiffVersion::Biff5 => {
            // BIFF5 FILEPASS is XOR obfuscation: 4 bytes (key + verifier).
            if data.len() < 4 {
                return Err(DecryptError::InvalidFilePass(format!(
                    "BIFF5 FILEPASS too short: expected 4 bytes, got {}",
                    data.len()
                )));
            }
            let key = read_u16(data, 0, "key")?;
            let verifier = read_u16(data, 2, "verifier")?;
            Ok(BiffEncryption::Biff5Xor { key, verifier })
        }
        BiffVersion::Biff8 => {
            let encryption_type = read_u16(data, 0, "wEncryptionType")?;
            match encryption_type {
                BIFF8_ENCRYPTION_TYPE_XOR => {
                    // BIFF8 XOR obfuscation: 6 bytes (type + key + verifier).
                    if data.len() < 6 {
                        return Err(DecryptError::InvalidFilePass(format!(
                            "BIFF8 XOR FILEPASS too short: expected at least 6 bytes, got {}",
                            data.len()
                        )));
                    }
                    let key = read_u16(data, 2, "key")?;
                    let verifier = read_u16(data, 4, "verifier")?;
                    Ok(BiffEncryption::Biff8Xor { key, verifier })
                }
                BIFF8_ENCRYPTION_TYPE_RC4 => {
                    // BIFF8 RC4 encryption: type + subType + algorithm-specific payload.
                    let sub_type = read_u16(data, 2, "subType")?;
                    match sub_type {
                        BIFF8_RC4_SUBTYPE_RC4 => Ok(BiffEncryption::Biff8Rc4 {
                            filepass_payload: data.to_vec(),
                        }),
                        BIFF8_RC4_SUBTYPE_CRYPTOAPI => Ok(BiffEncryption::Biff8Rc4CryptoApi {
                            filepass_payload: data.to_vec(),
                        }),
                        _ => Err(DecryptError::UnsupportedEncryption(format!(
                            "BIFF8 RC4 FILEPASS has unsupported subType=0x{sub_type:04X}"
                        ))),
                    }
                }
                _ => Err(DecryptError::UnsupportedEncryption(format!(
                    "BIFF8 FILEPASS has unsupported wEncryptionType=0x{encryption_type:04X}"
                ))),
            }
        }
    }
}

/// Decrypt an in-memory BIFF workbook stream using the provided password.
///
/// This function:
/// 1. Iterates BIFF records from offset 0 (record headers are plaintext).
/// 2. Locates the workbook-global `FILEPASS` record.
/// 3. Parses the FILEPASS payload to determine encryption scheme.
/// 4. Dispatches to an algorithm-specific decryptor to decrypt record payloads *after*
///    FILEPASS.
///
/// Note: Bytes before FILEPASS are always plaintext; encryption begins immediately
/// after the FILEPASS record.
pub(crate) fn decrypt_workbook_stream(
    workbook_stream: &mut [u8],
    password: &str,
) -> Result<(), DecryptError> {
    let biff_version = super::detect_biff_version(workbook_stream);

    let mut iter = records::BiffRecordIter::from_offset(workbook_stream, 0)
        .map_err(DecryptError::InvalidFilePass)?;

    while let Some(next) = iter.next() {
        let record = next.map_err(DecryptError::InvalidFilePass)?;

        // The FILEPASS record only appears in the workbook globals substream. Stop scanning
        // once we hit the next substream (BOF) or EOF.
        if record.offset != 0 && records::is_bof_record(record.record_id) {
            break;
        }
        if record.record_id == records::RECORD_EOF {
            break;
        }

        if record.record_id != records::RECORD_FILEPASS {
            continue;
        }

        if password.is_empty() {
            return Err(DecryptError::PasswordRequired);
        }

        let encryption = parse_filepass_record(biff_version, record.data)?;
        let encrypted_start = record
            .offset
            .checked_add(4)
            .and_then(|v| v.checked_add(record.data.len()))
            .ok_or_else(|| {
                DecryptError::InvalidFilePass("FILEPASS offset overflow".to_string())
            })?;

        return decrypt_after_filepass(&encryption, workbook_stream, encrypted_start, password);
    }

    Err(DecryptError::NoFilePass)
}

fn decrypt_after_filepass(
    encryption: &BiffEncryption,
    _workbook_stream: &mut [u8],
    _encrypted_start: usize,
    _password: &str,
) -> Result<(), DecryptError> {
    // We intentionally do *not* implement any crypto algorithms yet. This module's
    // immediate goal is to establish module boundaries and correct FILEPASS parsing.
    match encryption {
        BiffEncryption::Biff5Xor { .. } => Err(DecryptError::UnsupportedEncryption(
            "BIFF5 XOR decryption not implemented".to_string(),
        )),
        BiffEncryption::Biff8Xor { .. } => Err(DecryptError::UnsupportedEncryption(
            "BIFF8 XOR decryption not implemented".to_string(),
        )),
        BiffEncryption::Biff8Rc4 { .. } => Err(DecryptError::UnsupportedEncryption(
            "BIFF8 RC4 decryption not implemented".to_string(),
        )),
        BiffEncryption::Biff8Rc4CryptoApi { .. } => Err(DecryptError::UnsupportedEncryption(
            "BIFF8 RC4 CryptoAPI decryption not implemented".to_string(),
        )),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_biff5_xor_filepass() {
        let payload = [0x34, 0x12, 0x78, 0x56];
        let parsed = parse_filepass_record(BiffVersion::Biff5, &payload).expect("parse");
        assert_eq!(
            parsed,
            BiffEncryption::Biff5Xor {
                key: 0x1234,
                verifier: 0x5678
            }
        );
    }

    #[test]
    fn parses_biff8_xor_filepass() {
        let payload = [
            0x00, 0x00, // wEncryptionType
            0x34, 0x12, // key
            0x78, 0x56, // verifier
        ];
        let parsed = parse_filepass_record(BiffVersion::Biff8, &payload).expect("parse");
        assert_eq!(
            parsed,
            BiffEncryption::Biff8Xor {
                key: 0x1234,
                verifier: 0x5678
            }
        );
    }

    #[test]
    fn parses_biff8_rc4_filepass() {
        let payload = [
            0x01, 0x00, // wEncryptionType
            0x01, 0x00, // subType (RC4)
            0xAA, 0xBB, 0xCC,
        ];
        let parsed = parse_filepass_record(BiffVersion::Biff8, &payload).expect("parse");
        assert_eq!(
            parsed,
            BiffEncryption::Biff8Rc4 {
                filepass_payload: payload.to_vec()
            }
        );
    }

    #[test]
    fn parses_biff8_rc4_cryptoapi_filepass() {
        let payload = [
            0x01, 0x00, // wEncryptionType
            0x02, 0x00, // subType (CryptoAPI)
            0xDE, 0xAD, 0xBE, 0xEF,
        ];
        let parsed = parse_filepass_record(BiffVersion::Biff8, &payload).expect("parse");
        assert_eq!(
            parsed,
            BiffEncryption::Biff8Rc4CryptoApi {
                filepass_payload: payload.to_vec()
            }
        );
    }

    #[test]
    fn errors_on_unsupported_biff8_encryption_type() {
        let payload = [0x02, 0x00];
        let err = parse_filepass_record(BiffVersion::Biff8, &payload).expect_err("expected err");
        assert!(matches!(err, DecryptError::UnsupportedEncryption(_)));
    }

    #[test]
    fn errors_on_unsupported_biff8_rc4_subtype() {
        let payload = [0x01, 0x00, 0x03, 0x00];
        let err = parse_filepass_record(BiffVersion::Biff8, &payload).expect_err("expected err");
        assert!(matches!(err, DecryptError::UnsupportedEncryption(_)));
    }

    #[test]
    fn errors_on_truncated_biff5_payload() {
        let payload = [0x34, 0x12, 0x78];
        let err = parse_filepass_record(BiffVersion::Biff5, &payload).expect_err("expected err");
        assert!(matches!(err, DecryptError::InvalidFilePass(_)));
    }

    #[test]
    fn errors_on_truncated_biff8_payload_type() {
        let payload = [0x00];
        let err = parse_filepass_record(BiffVersion::Biff8, &payload).expect_err("expected err");
        assert!(matches!(err, DecryptError::InvalidFilePass(_)));
    }

    #[test]
    fn errors_on_truncated_biff8_xor_payload() {
        // wEncryptionType + key but missing verifier.
        let payload = [0x00, 0x00, 0x34, 0x12];
        let err = parse_filepass_record(BiffVersion::Biff8, &payload).expect_err("expected err");
        assert!(matches!(err, DecryptError::InvalidFilePass(_)));
    }

    #[test]
    fn errors_on_truncated_biff8_rc4_payload() {
        // wEncryptionType without subType.
        let payload = [0x01, 0x00, 0x01];
        let err = parse_filepass_record(BiffVersion::Biff8, &payload).expect_err("expected err");
        assert!(matches!(err, DecryptError::InvalidFilePass(_)));
    }
}

