//! Legacy `.xls` BIFF encryption (FILEPASS) parsing and workbook-stream decryption.
//!
//! This module provides:
//! - BIFF `FILEPASS` parsing / encryption scheme classification
//! - Crypto primitives (key derivation + verifier decryption) for BIFF8 RC4 and RC4 CryptoAPI
//!
//! BIFF workbook streams can indicate encryption/password protection with a `FILEPASS` record.
//! BIFF record headers (record id + length) remain plaintext so parsers can iterate records, but
//! record *payload* bytes after `FILEPASS` are encrypted.
//!
//! This module provides:
//! - Best-effort parsing/classification of `FILEPASS` payloads.
//! - In-place decryption of workbook stream bytes for supported schemes.

#![allow(dead_code)]

#[cfg(test)]
use md5::{Digest as _, Md5};
use thiserror::Error;
use zeroize::{Zeroize, Zeroizing};

use crate::ct::ct_eq;

use super::{records, BiffVersion};

// CryptoAPI RC4 parsing + key derivation helpers.
pub(crate) mod cryptoapi;
pub(crate) mod rc4;
pub(crate) mod xor;

use rc4::Rc4;
/// BIFF8 RC4 encryption uses 1024-byte blocks for the record-data byte stream.
///
/// Record headers (record id + length) are not encrypted and must not contribute to the block
/// position.
const RC4_BLOCK_SIZE: usize = 1024;

// BIFF8 FILEPASS.wEncryptionType values.
// [MS-XLS] 2.4.105 (FILEPASS).
const BIFF8_ENCRYPTION_TYPE_XOR: u16 = 0x0000;
const BIFF8_ENCRYPTION_TYPE_RC4: u16 = 0x0001;

// BIFF8 RC4 FILEPASS "subType" values (wEncryptionType==0x0001).
//
// In practice this corresponds to the `EncryptionInfo` major version:
// - 0x0001 => legacy RC4 "Standard Encryption"
// - 0x0002 => RC4 CryptoAPI
const BIFF8_RC4_SUBTYPE_RC4: u16 = 0x0001;
const BIFF8_RC4_SUBTYPE_CRYPTOAPI: u16 = 0x0002;
// Some BIFF8 RC4 CryptoAPI workbooks use an older FILEPASS layout where the second field is
// `wEncryptionInfo == 0x0004` (rather than `wEncryptionSubType == 0x0002`). In that layout the
// CryptoAPI `EncryptionHeader`/`EncryptionVerifier` structures are embedded directly in the
// FILEPASS payload (rather than using a length-prefixed `EncryptionInfo` blob).
const BIFF8_RC4_ENCRYPTION_INFO_CRYPTOAPI_LEGACY: u16 = 0x0004;

// BIFF record ids used by the legacy XOR obfuscation scheme that are either not encrypted or
// partially encrypted even when they appear after `FILEPASS`.
//
// See [MS-XLS] 2.2.10 "Encryption (Password to Open)".
const RECORD_BOUNDSHEET: u16 = 0x0085;
const RECORD_INTERFACEHDR: u16 = 0x00E1;
const RECORD_RRDINFO: u16 = 0x0138;
const RECORD_RRDHEAD: u16 = 0x0139;
const RECORD_USREXCL: u16 = 0x0194;
const RECORD_FILELOCK: u16 = 0x0195;

/// Maximum allowed FILEPASS record payload size.
///
/// The BIFF record header stores the size in a `u16`, but the payload is typically well under a
/// few hundred bytes. Capping the payload prevents unnecessary allocations when parsing untrusted
/// workbook streams.
pub(crate) const MAX_FILEPASS_PAYLOAD_BYTES: usize = 4 * 1024;

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) enum BiffEncryption {
    /// BIFF5 XOR obfuscation. FILEPASS payload is `key` + `verifier`.
    Biff5Xor { key: u16, verifier: u16 },
    /// BIFF8 XOR obfuscation. FILEPASS payload is `wEncryptionType` + `key` + `verifier`.
    Biff8Xor { key: u16, verifier: u16 },
    /// BIFF8 RC4 encryption (legacy non-CryptoAPI).
    ///
    /// The full FILEPASS payload is preserved so decryptors can parse algorithm details.
    Biff8Rc4 {
        filepass_payload: Zeroizing<Vec<u8>>,
    },
    /// BIFF8 RC4 encryption using CryptoAPI.
    ///
    /// The full FILEPASS payload is preserved so decryptors can parse algorithm details.
    Biff8Rc4CryptoApi {
        filepass_payload: Zeroizing<Vec<u8>>,
    },
    /// BIFF8 RC4 encryption using CryptoAPI with a legacy FILEPASS layout (`wEncryptionInfo=0x0004`).
    ///
    /// This variant uses different RC4 stream-position semantics than the standard CryptoAPI
    /// encoding.
    Biff8Rc4CryptoApiLegacy {
        filepass_payload: Zeroizing<Vec<u8>>,
    },
}

#[derive(Debug, Error, Clone, PartialEq, Eq)]
pub(crate) enum DecryptError {
    #[error("workbook is not encrypted (missing FILEPASS record)")]
    NoFilePass,
    #[error("invalid FILEPASS record: {0}")]
    InvalidFilePass(String),
    #[error("unsupported encryption scheme: {0}")]
    UnsupportedEncryption(String),
    #[error("{context} exceeds maximum allowed size ({limit} bytes)")]
    SizeLimitExceeded { context: &'static str, limit: usize },
    #[error("password required")]
    PasswordRequired,
    #[error("wrong password")]
    WrongPassword,
}

fn read_u16(data: &[u8], offset: usize, ctx: &str) -> Result<u16, DecryptError> {
    let end = offset.checked_add(2).ok_or_else(|| {
        DecryptError::InvalidFilePass(format!(
            "offset overflow while reading {ctx} at offset {offset}"
        ))
    })?;
    let bytes = data.get(offset..end).ok_or_else(|| {
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
    if data.len() > MAX_FILEPASS_PAYLOAD_BYTES {
        return Err(DecryptError::SizeLimitExceeded {
            context: "FILEPASS payload",
            limit: MAX_FILEPASS_PAYLOAD_BYTES,
        });
    }

    match biff_version {
        BiffVersion::Biff5 => {
            // BIFF5 FILEPASS is typically XOR obfuscation: 4 bytes (key + verifier).
            //
            // Some writers (notably LibreOffice when saving "Excel 5.0/95 Workbook") emit a BIFF8-
            // style FILEPASS payload that starts with `wEncryptionType` followed by the XOR
            // key/verifier (6 bytes total). Accept both forms.
            match data.len() {
                n if n < 4 => Err(DecryptError::InvalidFilePass(format!(
                    "BIFF5 FILEPASS too short: expected at least 4 bytes, got {n}",
                ))),
                4 => {
                    let key = read_u16(data, 0, "key")?;
                    let verifier = read_u16(data, 2, "verifier")?;
                    Ok(BiffEncryption::Biff5Xor { key, verifier })
                }
                n => {
                    // BIFF8-style header: wEncryptionType + key + verifier.
                    let encryption_type = read_u16(data, 0, "wEncryptionType")?;
                    match encryption_type {
                        BIFF8_ENCRYPTION_TYPE_XOR => {
                            if n < 6 {
                                return Err(DecryptError::InvalidFilePass(format!(
                                    "BIFF5 XOR FILEPASS too short: expected at least 6 bytes, got {n}"
                                )));
                            }
                            let key = read_u16(data, 2, "key")?;
                            let verifier = read_u16(data, 4, "verifier")?;
                            Ok(BiffEncryption::Biff5Xor { key, verifier })
                        }
                        BIFF8_ENCRYPTION_TYPE_RC4 => Err(DecryptError::UnsupportedEncryption(
                            "BIFF5 RC4 encryption is not supported".to_string(),
                        )),
                        _ => Err(DecryptError::UnsupportedEncryption(format!(
                            "BIFF5 FILEPASS has unsupported wEncryptionType=0x{encryption_type:04X}"
                        ))),
                    }
                }
            }
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
                            filepass_payload: Zeroizing::new(data.to_vec()),
                        }),
                        BIFF8_RC4_SUBTYPE_CRYPTOAPI => Ok(BiffEncryption::Biff8Rc4CryptoApi {
                            filepass_payload: Zeroizing::new(data.to_vec()),
                        }),
                        BIFF8_RC4_ENCRYPTION_INFO_CRYPTOAPI_LEGACY => {
                            Ok(BiffEncryption::Biff8Rc4CryptoApiLegacy {
                                filepass_payload: Zeroizing::new(data.to_vec()),
                            })
                        }
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

fn find_filepass_record(workbook_stream: &[u8]) -> Result<Option<(usize, usize)>, DecryptError> {
    let mut iter = records::BiffRecordIter::from_offset(workbook_stream, 0)
        .map_err(DecryptError::InvalidFilePass)?;

    // Require a valid BIFF workbook stream (must start with BOF) to avoid accidentally treating
    // arbitrary byte buffers as encrypted just because they contain the FILEPASS record id.
    let first = match iter.next() {
        Some(Ok(record)) => record,
        Some(Err(err)) => return Err(DecryptError::InvalidFilePass(err)),
        None => {
            return Err(DecryptError::InvalidFilePass(
                "empty workbook stream".to_string(),
            ))
        }
    };
    if !records::is_bof_record(first.record_id) {
        return Err(DecryptError::InvalidFilePass(format!(
            "workbook stream does not start with BOF (found 0x{:04X})",
            first.record_id
        )));
    }

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

        return Ok(Some((record.offset, record.data.len())));
    }

    Ok(None)
}

#[cfg(test)]
fn apply_cipher_after_offset<F>(
    workbook_stream: &mut [u8],
    start_offset: usize,
    mut cipher: F,
) -> Result<(), DecryptError>
where
    F: FnMut(u32, usize, &mut [u8]) -> Result<(), DecryptError>,
{
    if start_offset > workbook_stream.len() {
        return Err(DecryptError::InvalidFilePass(format!(
            "encrypted start offset {start_offset} out of bounds (len={})",
            workbook_stream.len()
        )));
    }

    // Collect record boundaries using the physical iterator (borrowed immutably) so we can
    // decrypt record payloads in-place without fighting Rust's borrow checker.
    let records: Vec<(usize, usize)> = {
        let mut iter = records::BiffRecordIter::from_offset(&*workbook_stream, start_offset)
            .map_err(DecryptError::InvalidFilePass)?;
        let mut out = Vec::new();
        while let Some(next) = iter.next() {
            let record = next.map_err(DecryptError::InvalidFilePass)?;
            out.push((record.offset, record.data.len()));
        }
        out
    };

    // Position within the encrypted byte stream, counting only record payload bytes (not headers).
    let mut encrypted_pos: usize = 0;

    for (record_offset, payload_len) in records {
        let data_start = record_offset
            .checked_add(4)
            .ok_or_else(|| DecryptError::InvalidFilePass("record offset overflow".to_string()))?;

        let mut local = 0usize;
        while local < payload_len {
            let block_index = (encrypted_pos / RC4_BLOCK_SIZE) as u32;
            let block_offset = encrypted_pos % RC4_BLOCK_SIZE;
            let remaining_in_block = RC4_BLOCK_SIZE - block_offset;
            let remaining_in_record = payload_len - local;
            let chunk_len = remaining_in_block.min(remaining_in_record);

            let start = data_start
                .checked_add(local)
                .ok_or_else(|| DecryptError::InvalidFilePass("record offset overflow".to_string()))?;
            let end = start.checked_add(chunk_len).ok_or_else(|| {
                DecryptError::InvalidFilePass("record length overflow".to_string())
            })?;
            let chunk = workbook_stream.get_mut(start..end).ok_or_else(|| {
                DecryptError::InvalidFilePass("record payload out of bounds".to_string())
            })?;

            cipher(block_index, block_offset, chunk)?;

            encrypted_pos = encrypted_pos.saturating_add(chunk_len);
            local = local.saturating_add(chunk_len);
        }

        // `local` must advance exactly to the payload end.
        debug_assert_eq!(local, payload_len);
    }

    Ok(())
}

/// Test-only helper: decrypt a workbook stream using a caller-provided "cipher" implementation.
///
/// This exists to fuzz/property-test the decryptor's record-walking / block-counter logic without
/// needing a real BIFF crypto implementation.
#[cfg(test)]
pub(crate) fn decrypt_workbook_stream_with_cipher<F>(
    workbook_stream: &mut [u8],
    password: &str,
    cipher: F,
) -> Result<(), DecryptError>
where
    F: FnMut(u32, usize, &mut [u8]) -> Result<(), DecryptError>,
{
    let Some((filepass_offset, filepass_len)) = find_filepass_record(&*workbook_stream)? else {
        return Err(DecryptError::NoFilePass);
    };

    if password.is_empty() {
        return Err(DecryptError::PasswordRequired);
    }

    let encrypted_start = filepass_offset
        .checked_add(4)
        .and_then(|v| v.checked_add(filepass_len))
        .ok_or_else(|| DecryptError::InvalidFilePass("FILEPASS offset overflow".to_string()))?;

    apply_cipher_after_offset(workbook_stream, encrypted_start, cipher)
}

/// Decrypt an in-memory BIFF workbook stream using the provided password.
///
/// This function:
/// 1. Iterates BIFF records from offset 0 (record headers are plaintext).
/// 2. Locates the workbook-global `FILEPASS` record.
/// 3. Parses the FILEPASS payload to determine encryption scheme.
/// 4. Dispatches to an algorithm-specific decryptor to decrypt record payloads *after* FILEPASS.
///
/// Note: Bytes before FILEPASS are always plaintext; encryption begins immediately after the
/// FILEPASS record.
pub(crate) fn decrypt_workbook_stream(
    workbook_stream: &mut [u8],
    password: &str,
) -> Result<(), DecryptError> {
    let biff_version = super::detect_biff_version(workbook_stream);

    let mut iter = records::BiffRecordIter::from_offset(workbook_stream, 0)
        .map_err(DecryptError::InvalidFilePass)?;

    // Require a valid BIFF workbook stream (must start with BOF) to avoid accidentally treating
    // arbitrary byte buffers as encrypted just because they contain the FILEPASS record id.
    let first = match iter.next() {
        Some(Ok(record)) => record,
        Some(Err(err)) => return Err(DecryptError::InvalidFilePass(err)),
        None => {
            return Err(DecryptError::InvalidFilePass(
                "empty workbook stream".to_string(),
            ))
        }
    };
    if !records::is_bof_record(first.record_id) {
        return Err(DecryptError::InvalidFilePass(format!(
            "workbook stream does not start with BOF (found 0x{:04X})",
            first.record_id
        )));
    }

    while let Some(next) = iter.next() {
        let record = next.map_err(DecryptError::InvalidFilePass)?;

        // The FILEPASS record only appears in the workbook globals substream. Stop scanning once we
        // hit the next substream (BOF) or EOF.
        if record.offset != 0 && records::is_bof_record(record.record_id) {
            break;
        }
        if record.record_id == records::RECORD_EOF {
            break;
        }

        if record.record_id != records::RECORD_FILEPASS {
            continue;
        }

        let encryption = parse_filepass_record(biff_version, record.data)?;
        let encrypted_start = record
            .offset
            .checked_add(4)
            .and_then(|v| v.checked_add(record.data.len()))
            .ok_or_else(|| DecryptError::InvalidFilePass("FILEPASS offset overflow".to_string()))?;

        return decrypt_after_filepass(&encryption, workbook_stream, encrypted_start, password);
    }

    Err(DecryptError::NoFilePass)
}

fn decrypt_after_filepass(
    encryption: &BiffEncryption,
    workbook_stream: &mut [u8],
    encrypted_start: usize,
    password: &str,
) -> Result<(), DecryptError> {
    match encryption {
        BiffEncryption::Biff5Xor { key, verifier } | BiffEncryption::Biff8Xor { key, verifier } => {
            decrypt_biff_xor_obfuscation(workbook_stream, encrypted_start, password, *key, *verifier)
        }
        BiffEncryption::Biff8Rc4 { filepass_payload } => {
            decrypt_biff8_rc4_standard(workbook_stream, encrypted_start, password, filepass_payload)
        }
        BiffEncryption::Biff8Rc4CryptoApi { filepass_payload }
        | BiffEncryption::Biff8Rc4CryptoApiLegacy { filepass_payload } => cryptoapi::decrypt_workbook_stream_rc4_cryptoapi(
            workbook_stream,
            encrypted_start,
            password,
            filepass_payload,
        ),
    }
}

fn derive_xor_array(password: &str) -> Zeroizing<[u8; 16]> {
    // Best-effort: derive a deterministic 16-byte XOR array from the password's low UTF-16 bytes.
    //
    // BIFF XOR obfuscation is legacy and not cryptographically secure; this helper exists to make
    // record-boundary handling testable without relying on large binary fixtures.
    const PAD: [u8; 16] = [
        0xBB, 0xFF, 0xFF, 0xBA, 0xFF, 0xFF, 0xB9, 0xFF, 0xFF, 0xB8, 0xFF, 0xFF, 0xB7, 0xFF,
        0xFF, 0xB6,
    ];

    let mut out = PAD;
    for (i, ch) in password.encode_utf16().take(out.len()).enumerate() {
        out[i] ^= (ch & 0xFF) as u8;
    }
    Zeroizing::new(out)
}

fn apply_xor_obfuscation_in_place(
    workbook_stream: &mut [u8],
    encrypted_start: usize,
    key: u16,
    xor_array: &[u8; 16],
) -> Result<(), DecryptError> {
    // Ensure key bytes don't linger on the stack on early-return error paths.
    let key_bytes = Zeroizing::new(key.to_le_bytes());
    let mut pos = 0usize;

    let mut offset = encrypted_start;
    while offset < workbook_stream.len() {
        let remaining = workbook_stream.len().saturating_sub(offset);
        if remaining < 4 {
            // Some writers may include trailing padding bytes after the final EOF record. Those
            // bytes are not part of any record payload and should be ignored rather than treated
            // as a truncated record header.
            break;
        }

        let record_id = u16::from_le_bytes([workbook_stream[offset], workbook_stream[offset + 1]]);
        let len =
            u16::from_le_bytes([workbook_stream[offset + 2], workbook_stream[offset + 3]]) as usize;

        let data_start = offset
            .checked_add(4)
            .ok_or_else(|| DecryptError::InvalidFilePass("BIFF record offset overflow".to_string()))?;
        let data_end = data_start
            .checked_add(len)
            .ok_or_else(|| DecryptError::InvalidFilePass("BIFF record length overflow".to_string()))?;
        if data_end > workbook_stream.len() {
            return Err(DecryptError::InvalidFilePass(format!(
                "BIFF record 0x{record_id:04X} at offset {offset} extends past end of stream while decrypting XOR (len={}, end={data_end})",
                workbook_stream.len()
            )));
        }

        let data = &mut workbook_stream[data_start..data_end];
        for b in data.iter_mut() {
            let ks = xor_array[pos % xor_array.len()] ^ key_bytes[pos % 2];
            *b ^= ks;
            pos = pos.saturating_add(1);
        }

        offset = data_end;
    }
    Ok(())
}

fn decrypt_biff_xor_obfuscation(
    workbook_stream: &mut [u8],
    encrypted_start: usize,
    password: &str,
    key: u16,
    verifier: u16,
) -> Result<(), DecryptError> {
    // First, try to validate/decrypt using the real Excel XOR obfuscation scheme as described in
    // MS-OFFCRYPTO/MS-XLS:
    // - Verify the password by recomputing the FILEPASS `key` + `verifier` fields.
    // - If those fields match, decrypt record payload bytes using the derived 16-byte XOR array.
    //
    // Some of our existing tiny BIFF8 XOR fixtures are generated deterministically for tests using
    // a simplified XOR scheme (to avoid depending on large binary blobs). That legacy format is
    // still supported as a fallback when the FILEPASS fields do not match the spec algorithm.
    if let Some(xor_array) = xor_array_method1_for_password(password, key, verifier) {
        return decrypt_payloads_after_filepass_xor_method1(workbook_stream, encrypted_start, &xor_array);
    }

    // Fallback: simplified XOR scheme used by deterministic in-repo fixtures.
    // This uses the legacy worksheet/workbook protection password hash as a verifier and applies
    // a repeating XOR keystream across record payload bytes.
    let mut expected = xor::xor_password_verifier(password);
    let mut expected_bytes = expected.to_le_bytes();
    let mut verifier_bytes = verifier.to_le_bytes();
    let ok = ct_eq(&expected_bytes, &verifier_bytes);
    expected_bytes.zeroize();
    verifier_bytes.zeroize();
    if !ok {
        expected.zeroize();
        return Err(DecryptError::WrongPassword);
    }
    // Guard against accidentally treating a real Method-1 XOR-encrypted workbook as a simplified
    // test fixture. The deterministic test encryptor uses `key = verifier ^ 0xFFFF`; real Excel
    // workbooks use `CreateXorKey_Method1` for the key.
    if key != (expected ^ 0xFFFF) {
        expected.zeroize();
        return Err(DecryptError::WrongPassword);
    }
    expected.zeroize();

    let xor_array = derive_xor_array(password);
    apply_xor_obfuscation_in_place(workbook_stream, encrypted_start, key, &xor_array)
}

// -------------------------------------------------------------------------------------------------
// XOR obfuscation (MS-OFFCRYPTO/MS-XLS) implementation ("Method 1")
// -------------------------------------------------------------------------------------------------

fn xor_password_byte_candidates(password: &str) -> Vec<Zeroizing<Vec<u8>>> {
    use encoding_rs::WINDOWS_1252;

    let mut out = Vec::new();

    // 1) Windows-1252 encoding (common Excel default on Windows).
    {
        let (cow, _, _) = WINDOWS_1252.encode(password);
        let mut bytes = Zeroizing::new(cow.into_owned());
        if bytes.len() > 15 {
            bytes[15..].zeroize();
            bytes.truncate(15);
        }
        out.push(bytes);
    }

    // 2) MS-OFFCRYPTO 2.3.7.4 "method 2": copy low byte unless zero, else high byte.
    {
        let mut bytes = Zeroizing::new(Vec::with_capacity(15));
        for ch in password.encode_utf16() {
            if bytes.len() >= 15 {
                break;
            }
            let lo = (ch & 0x00FF) as u8;
            let hi = (ch >> 8) as u8;
            bytes.push(if lo != 0 { lo } else { hi });
        }
        out.push(bytes);
    }

    out
}

// [MS-OFFCRYPTO] 2.3.7.2 (CreateXorArray_Method1) constants.
const XOR_PAD_ARRAY: [u8; 15] = [
    0xBB, 0xFF, 0xFF, 0xBA, 0xFF, 0xFF, 0xB9, 0x80, 0x00, 0xBE, 0x0F, 0x00, 0xBF, 0x0F, 0x00,
];

const XOR_INITIAL_CODE: [u16; 15] = [
    0xE1F0, 0x1D0F, 0xCC9C, 0x84C0, 0x110C, 0x0E10, 0xF1CE, 0x313E, 0x1872, 0xE139, 0xD40F,
    0x84F9, 0x280C, 0xA96A, 0x4EC3,
];

const XOR_MATRIX: [u16; 105] = [
    0xAEFC, 0x4DD9, 0x9BB2, 0x2745, 0x4E8A, 0x9D14, 0x2A09, 0x7B61, 0xF6C2, 0xFDA5, 0xEB6B,
    0xC6F7, 0x9DCF, 0x2BBF, 0x4563, 0x8AC6, 0x05AD, 0x0B5A, 0x16B4, 0x2D68, 0x5AD0, 0x0375,
    0x06EA, 0x0DD4, 0x1BA8, 0x3750, 0x6EA0, 0xDD40, 0xD849, 0xA0B3, 0x5147, 0xA28E, 0x553D,
    0xAA7A, 0x44D5, 0x6F45, 0xDE8A, 0xAD35, 0x4A4B, 0x9496, 0x390D, 0x721A, 0xEB23, 0xC667,
    0x9CEF, 0x29FF, 0x53FE, 0xA7FC, 0x5FD9, 0x47D3, 0x8FA6, 0x0F6D, 0x1EDA, 0x3DB4, 0x7B68,
    0xF6D0, 0xB861, 0x60E3, 0xC1C6, 0x93AD, 0x377B, 0x6EF6, 0xDDEC, 0x45A0, 0x8B40, 0x06A1,
    0x0D42, 0x1A84, 0x3508, 0x6A10, 0xAA51, 0x4483, 0x8906, 0x022D, 0x045A, 0x08B4, 0x1168,
    0x76B4, 0xED68, 0xCAF1, 0x85C3, 0x1BA7, 0x374E, 0x6E9C, 0x3730, 0x6E60, 0xDCC0, 0xA9A1,
    0x4363, 0x86C6, 0x1DAD, 0x3331, 0x6662, 0xCCC4, 0x89A9, 0x0373, 0x06E6, 0x0DCC, 0x1021,
    0x2042, 0x4084, 0x8108, 0x1231, 0x2462, 0x48C4,
];

fn xor_ror(byte1: u8, byte2: u8) -> u8 {
    (byte1 ^ byte2).rotate_right(1)
}

fn create_password_verifier_method1(password: &[u8]) -> u16 {
    let mut verifier: u16 = 0;
    let mut password_array =
        Zeroizing::new(Vec::<u8>::with_capacity(password.len().saturating_add(1)));
    password_array.push(password.len() as u8);
    password_array.extend_from_slice(password);

    for &b in password_array.iter().rev() {
        let intermediate1 = if (verifier & 0x4000) == 0 { 0u16 } else { 1u16 };
        let intermediate2 = verifier.wrapping_mul(2) & 0x7FFF;
        let intermediate3 = intermediate1 | intermediate2;
        verifier = intermediate3 ^ (b as u16);
    }

    verifier ^ 0xCE4B
}

fn create_xor_key_method1(password: &[u8]) -> u16 {
    if password.is_empty() || password.len() > 15 {
        return 0;
    }

    let mut xor_key = XOR_INITIAL_CODE[password.len() - 1];
    let mut current_element: i32 = 0x68;

    for &byte in password.iter().rev() {
        let mut ch = byte;
        for _ in 0..7 {
            if (ch & 0x40) != 0 {
                if current_element < 0 || current_element as usize >= XOR_MATRIX.len() {
                    return xor_key;
                }
                xor_key ^= XOR_MATRIX[current_element as usize];
            }
            ch = ch.wrapping_mul(2);
            current_element -= 1;
        }
    }

    xor_key
}

fn create_xor_array_method1(password: &[u8], xor_key: u16) -> [u8; 16] {
    let mut out = [0u8; 16];
    let mut index = password.len();

    let key_high = (xor_key >> 8) as u8;
    let key_low = (xor_key & 0x00FF) as u8;

    if index % 2 == 1 {
        if index < out.len() {
            out[index] = xor_ror(XOR_PAD_ARRAY[0], key_high);
        }

        index = index.saturating_sub(1);

        if !password.is_empty() && index < out.len() {
            let password_last = password[password.len() - 1];
            out[index] = xor_ror(password_last, key_low);
        }
    }

    while index > 0 {
        index = index.saturating_sub(1);
        if index < password.len() {
            out[index] = xor_ror(password[index], key_high);
        }

        index = index.saturating_sub(1);
        if index < password.len() {
            out[index] = xor_ror(password[index], key_low);
        }
    }

    let mut out_index: i32 = 15;
    let mut pad_index: i32 = 15i32 - (password.len() as i32);
    while pad_index > 0 {
        if out_index < 0 {
            break;
        }

        let pi = pad_index as usize;
        if pi < XOR_PAD_ARRAY.len() {
            out[out_index as usize] = xor_ror(XOR_PAD_ARRAY[pi], key_high);
        }
        out_index -= 1;
        pad_index -= 1;

        if out_index < 0 {
            break;
        }

        let pi = pad_index.max(0) as usize;
        if pi < XOR_PAD_ARRAY.len() {
            out[out_index as usize] = xor_ror(XOR_PAD_ARRAY[pi], key_low);
        }
        out_index -= 1;
        pad_index -= 1;
    }

    out
}

fn xor_array_method1_for_password(
    password: &str,
    stored_key: u16,
    stored_verifier: u16,
) -> Option<Zeroizing<[u8; 16]>> {
    for candidate in xor_password_byte_candidates(password) {
        // Passwords are limited to 15 bytes, but some writers can emit an empty password (length 0).
        if candidate.len() > 15 {
            continue;
        }
        // These values are derived from the password and should not linger longer than needed.
        let mut key = create_xor_key_method1(&candidate);
        let mut verifier = create_password_verifier_method1(&candidate);

        let mut key_bytes = key.to_le_bytes();
        let mut stored_key_bytes = stored_key.to_le_bytes();
        let mut verifier_bytes = verifier.to_le_bytes();
        let mut stored_verifier_bytes = stored_verifier.to_le_bytes();

        let ok = ct_eq(&key_bytes, &stored_key_bytes) & ct_eq(&verifier_bytes, &stored_verifier_bytes);
        key_bytes.zeroize();
        stored_key_bytes.zeroize();
        verifier_bytes.zeroize();
        stored_verifier_bytes.zeroize();

        if ok {
            let out = Zeroizing::new(create_xor_array_method1(&candidate, key));
            key.zeroize();
            verifier.zeroize();
            return Some(out);
        }

        key.zeroize();
        verifier.zeroize();
    }
    None
}

fn decrypt_payloads_after_filepass_xor_method1(
    workbook_stream: &mut [u8],
    start_offset: usize,
    xor_array: &[u8; 16],
) -> Result<(), DecryptError> {
    let mut offset = start_offset;
    while offset < workbook_stream.len() {
        let remaining = workbook_stream.len().saturating_sub(offset);
        if remaining < 4 {
            // Some writers include trailing padding bytes after the final EOF record. Those bytes
            // are not part of any record header/payload and should be ignored.
            break;
        }

        let record_id = u16::from_le_bytes([workbook_stream[offset], workbook_stream[offset + 1]]);
        let len =
            u16::from_le_bytes([workbook_stream[offset + 2], workbook_stream[offset + 3]]) as usize;

        let data_start = offset
            .checked_add(4)
            .ok_or_else(|| DecryptError::InvalidFilePass("BIFF record offset overflow".to_string()))?;
        let data_end = data_start
            .checked_add(len)
            .ok_or_else(|| DecryptError::InvalidFilePass("BIFF record length overflow".to_string()))?;
        if data_end > workbook_stream.len() {
            return Err(DecryptError::InvalidFilePass(format!(
                "BIFF record 0x{record_id:04X} at offset {offset} extends past end of stream while decrypting XOR (len={}, end={data_end})",
                workbook_stream.len()
            )));
        }

        // Per [MS-XLS] 2.2.10, some record payloads are not encrypted even in an encrypted BIFF
        // record stream.
        let mut decrypt_from = 0usize;
        let skip_entire_payload = matches!(
            record_id,
            // BOF (both BIFF5 0x0009 and BIFF8 0x0809 ids) + FILEPASS
            records::RECORD_BOF_BIFF8
                | records::RECORD_BOF_BIFF5
                | records::RECORD_FILEPASS
                | RECORD_INTERFACEHDR
                | RECORD_FILELOCK
                | RECORD_USREXCL
                | RECORD_RRDINFO
                | RECORD_RRDHEAD
        );

        if !skip_entire_payload {
            if record_id == RECORD_BOUNDSHEET {
                // BoundSheet.lbPlyPos MUST NOT be encrypted.
                decrypt_from = 4.min(len);
            }

            let payload = &mut workbook_stream[data_start..data_end];
            for i in decrypt_from..payload.len() {
                let abs_pos = data_start + i;
                let mut value = payload[i];
                value ^= xor_array[abs_pos % 16];
                value = value.rotate_right(5);
                payload[i] = value;
            }
        }

        offset = data_end;
    }

    Ok(())
}

/// Parsed BIFF8 RC4 FILEPASS payload for legacy "Standard Encryption".
#[derive(Debug, Clone, PartialEq, Eq)]
struct FilePassRc4 {
    /// RC4 key length in bytes (either 5 bytes / 40-bit, or 16 bytes / 128-bit).
    key_len: usize,
    /// Random salt / "DocId" (16 bytes).
    salt: [u8; 16],
    /// Encrypted verifier (16 bytes).
    encrypted_verifier: [u8; 16],
    /// Encrypted verifier hash (16 bytes, MD5).
    encrypted_verifier_hash: [u8; 16],
}

impl Drop for FilePassRc4 {
    fn drop(&mut self) {
        self.key_len = 0;
        self.salt.zeroize();
        self.encrypted_verifier.zeroize();
        self.encrypted_verifier_hash.zeroize();
    }
}
fn parse_filepass_rc4(payload: &[u8]) -> Result<FilePassRc4, DecryptError> {
    // FILEPASS payload begins with wEncryptionType (u16).
    if payload.len() < 2 {
        return Err(DecryptError::InvalidFilePass(
            "truncated FILEPASS record".to_string(),
        ));
    }
    let encryption_type = u16::from_le_bytes([payload[0], payload[1]]);
    if encryption_type != BIFF8_ENCRYPTION_TYPE_RC4 {
        return Err(DecryptError::UnsupportedEncryption(format!(
            "FILEPASS wEncryptionType=0x{encryption_type:04X} (expected RC4)"
        )));
    }

    // EncryptionInfo: major/minor version.
    if payload.len() < 6 {
        return Err(DecryptError::InvalidFilePass(
            "truncated FILEPASS RC4 header".to_string(),
        ));
    }

    let major = u16::from_le_bytes([payload[2], payload[3]]);
    let minor = u16::from_le_bytes([payload[4], payload[5]]);

    // Excel 97-2003 Standard Encryption uses major version 1.
    if major != 1 {
        return Err(DecryptError::UnsupportedEncryption(format!(
            "unsupported FILEPASS RC4 major version {major} (expected 1 for Standard Encryption)"
        )));
    }

    // Minor version determines key length:
    // - 1 => 40-bit (5 bytes)
    // - 2 => 128-bit (16 bytes)
    let key_len = match minor {
        1 => 5,
        2 => 16,
        _ => {
            return Err(DecryptError::UnsupportedEncryption(format!(
                "unsupported FILEPASS RC4 minor version {minor} (expected 1 or 2)"
            )));
        }
    };

    // Standard encryption stores: salt (16), encrypted verifier (16), encrypted verifier hash (16)
    const EXPECTED_LEN: usize = 6 + 16 + 16 + 16;
    if payload.len() < EXPECTED_LEN {
        return Err(DecryptError::InvalidFilePass(format!(
            "truncated FILEPASS RC4 payload (len={}, need at least {EXPECTED_LEN})",
            payload.len()
        )));
    }

    let salt: [u8; 16] = payload
        .get(6..22)
        .and_then(|bytes| bytes.try_into().ok())
        .ok_or_else(|| {
            DecryptError::InvalidFilePass(
                "truncated FILEPASS RC4 salt (expected 16 bytes)".to_string(),
            )
        })?;
    let encrypted_verifier: [u8; 16] = payload
        .get(22..38)
        .and_then(|bytes| bytes.try_into().ok())
        .ok_or_else(|| {
            DecryptError::InvalidFilePass(
                "truncated FILEPASS RC4 encrypted verifier (expected 16 bytes)".to_string(),
            )
        })?;
    let encrypted_verifier_hash: [u8; 16] = payload
        .get(38..54)
        .and_then(|bytes| bytes.try_into().ok())
        .ok_or_else(|| {
            DecryptError::InvalidFilePass(
                "truncated FILEPASS RC4 encrypted verifier hash (expected 16 bytes)".to_string(),
            )
        })?;

    Ok(FilePassRc4 {
        key_len,
        salt,
        encrypted_verifier,
        encrypted_verifier_hash,
    })
}

fn password_to_utf16le(password: &str) -> Zeroizing<Vec<u8>> {
    // Excel 97-2003 passwords are limited to 15 characters for legacy RC4 encryption.
    //
    // Use UTF-16LE and truncate to 15 UTF-16 code units.
    let mut out = Zeroizing::new(Vec::with_capacity(password.len().min(15) * 2));
    for u in password.encode_utf16().take(15) {
        out.extend_from_slice(&u.to_le_bytes());
    }
    out
}

/// Applies BIFF8 RC4 encryption/decryption to a byte stream representing *record data* (not record
/// headers).
///
/// This is symmetric: applying it twice with the same key yields the original bytes.
struct Rc4BiffStream {
    intermediate_key: Zeroizing<[u8; 16]>,
    key_len: usize,
    block: u32,
    pos_in_block: usize,
    cipher: Rc4,
}

impl Rc4BiffStream {
    fn new(intermediate_key: Zeroizing<[u8; 16]>, key_len: usize) -> Self {
        let block_key = rc4::derive_biff8_rc4_block_key(&*intermediate_key, 0);
        let cipher = Rc4::new(&block_key[..key_len]);
        drop(block_key);
        Self {
            intermediate_key,
            key_len,
            block: 0,
            pos_in_block: 0,
            cipher,
        }
    }

    fn rekey(&mut self) {
        let block_key = rc4::derive_biff8_rc4_block_key(&*self.intermediate_key, self.block);
        self.cipher = Rc4::new(&block_key[..self.key_len]);
        drop(block_key);
    }

    fn apply(&mut self, mut data: &mut [u8]) {
        while !data.is_empty() {
            if self.pos_in_block == RC4_BLOCK_SIZE {
                self.block = self.block.wrapping_add(1);
                self.pos_in_block = 0;
                self.rekey();
            }
            let remaining_in_block = RC4_BLOCK_SIZE - self.pos_in_block;
            let n = remaining_in_block.min(data.len());
            let (chunk, rest) = data.split_at_mut(n);
            self.cipher.apply_keystream(chunk);
            self.pos_in_block += n;
            data = rest;
        }
    }
}

impl Drop for Rc4BiffStream {
    fn drop(&mut self) {
        // Ensure the expanded key schedule doesn't linger beyond the decryptor's lifetime.
        self.cipher.zeroize();
        self.block = 0;
        self.pos_in_block = 0;
    }
}

fn verify_rc4_password(
    filepass: &FilePassRc4,
    password: &str,
) -> Result<Zeroizing<[u8; 16]>, DecryptError> {
    if !rc4::validate_biff8_rc4_password(
        password,
        &filepass.salt,
        &filepass.encrypted_verifier,
        &filepass.encrypted_verifier_hash,
        filepass.key_len,
    ) {
        return Err(DecryptError::WrongPassword);
    }

    Ok(rc4::derive_biff8_rc4_intermediate_key(password, &filepass.salt))
}

fn decrypt_biff8_rc4_standard(
    workbook_stream: &mut [u8],
    encrypted_start: usize,
    password: &str,
    filepass_payload: &[u8],
) -> Result<(), DecryptError> {
    let filepass = parse_filepass_rc4(filepass_payload)?;
    let key_len = filepass.key_len;
    let intermediate_key = verify_rc4_password(&filepass, password)?;
    drop(filepass);

    let mut rc4_stream = Rc4BiffStream::new(intermediate_key, key_len);

    // Decrypt record payloads after FILEPASS.
    let mut offset = encrypted_start;
    while offset < workbook_stream.len() {
        let remaining = workbook_stream.len().saturating_sub(offset);
        if remaining < 4 {
            // Some writers include trailing padding bytes after the final EOF record. Those bytes
            // are not part of any record header/payload and should be ignored.
            break;
        }

        let record_id = u16::from_le_bytes([workbook_stream[offset], workbook_stream[offset + 1]]);
        let len =
            u16::from_le_bytes([workbook_stream[offset + 2], workbook_stream[offset + 3]]) as usize;

        let data_start = offset.checked_add(4).ok_or_else(|| {
            DecryptError::InvalidFilePass("BIFF record offset overflow".to_string())
        })?;
        let data_end = data_start.checked_add(len).ok_or_else(|| {
            DecryptError::InvalidFilePass("BIFF record length overflow".to_string())
        })?;
        if data_end > workbook_stream.len() {
            return Err(DecryptError::InvalidFilePass(format!(
                "BIFF record 0x{record_id:04X} at offset {offset} extends past end of stream while decrypting (len={}, end={data_end})",
                workbook_stream.len()
            )));
        }

        rc4_stream.apply(&mut workbook_stream[data_start..data_end]);
        offset = data_end;
    }

    Ok(())
}

// -------------------------------------------------------------------------------------------------
// Test-only encryption helpers
// -------------------------------------------------------------------------------------------------

#[cfg(test)]
fn collect_payload_ranges_after_offset(
    workbook_stream: &[u8],
    start_offset: usize,
) -> Result<(Vec<(std::ops::Range<usize>, usize)>, usize), DecryptError> {
    let mut ranges = Vec::<(std::ops::Range<usize>, usize)>::new();
    let mut offset = start_offset;
    let mut pos = 0usize;
    while offset < workbook_stream.len() {
        let remaining = workbook_stream.len().saturating_sub(offset);
        if remaining < 4 {
            return Err(DecryptError::InvalidFilePass(
                "truncated BIFF record header while scanning payload ranges".to_string(),
            ));
        }
        let len = u16::from_le_bytes([workbook_stream[offset + 2], workbook_stream[offset + 3]]) as usize;
        let data_start = offset + 4;
        let data_end = data_start.checked_add(len).ok_or_else(|| {
            DecryptError::InvalidFilePass("record length overflow while scanning payload ranges".to_string())
        })?;
        if data_end > workbook_stream.len() {
            return Err(DecryptError::InvalidFilePass(
                "record extends past end of stream while scanning payload ranges".to_string(),
            ));
        }
        ranges.push((data_start..data_end, pos));
        pos = pos.saturating_add(len);
        offset = data_end;
    }
    Ok((ranges, pos))
}

#[cfg(test)]
pub(crate) fn encrypt_workbook_stream_for_test(
    workbook_stream: &mut [u8],
    password: &str,
) -> Result<(), DecryptError> {
    let biff_version = super::detect_biff_version(workbook_stream);
    let mut iter =
        records::BiffRecordIter::from_offset(workbook_stream, 0).map_err(DecryptError::InvalidFilePass)?;

    let first = match iter.next() {
        Some(Ok(record)) => record,
        Some(Err(err)) => return Err(DecryptError::InvalidFilePass(err)),
        None => return Err(DecryptError::InvalidFilePass("empty workbook stream".to_string())),
    };
    if !records::is_bof_record(first.record_id) {
        return Err(DecryptError::InvalidFilePass(
            "workbook stream does not start with BOF".to_string(),
        ));
    }

    while let Some(next) = iter.next() {
        let record = next.map_err(DecryptError::InvalidFilePass)?;

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
        let payload_start = record.offset + 4;
        let payload_end = payload_start + record.data.len();
        let encrypted_start = payload_end;

        match encryption {
            BiffEncryption::Biff5Xor { .. } => {
                if record.data.len() < 4 {
                    return Err(DecryptError::InvalidFilePass(
                        "BIFF5 XOR FILEPASS payload too short".to_string(),
                    ));
                }
                let verifier = xor::xor_password_verifier(password);
                let key = verifier ^ 0xFFFF;
                workbook_stream[payload_start..payload_start + 2].copy_from_slice(&key.to_le_bytes());
                workbook_stream[payload_start + 2..payload_start + 4]
                    .copy_from_slice(&verifier.to_le_bytes());

                let (ranges, _total) = collect_payload_ranges_after_offset(workbook_stream, encrypted_start)?;
                let xor_array = derive_xor_array(password);
                let key_bytes = key.to_le_bytes();
                for (range, start_pos) in ranges {
                    for (i, b) in workbook_stream[range.clone()].iter_mut().enumerate() {
                        let pos = start_pos + i;
                        let ks = xor_array[pos % xor_array.len()] ^ key_bytes[pos % 2];
                        *b ^= ks;
                    }
                }
                return Ok(());
            }
            BiffEncryption::Biff8Xor { .. } => {
                if record.data.len() < 6 {
                    return Err(DecryptError::InvalidFilePass(
                        "BIFF8 XOR FILEPASS payload too short".to_string(),
                    ));
                }
                let verifier = xor::xor_password_verifier(password);
                let key = verifier ^ 0xFFFF;
                workbook_stream[payload_start..payload_start + 2]
                    .copy_from_slice(&BIFF8_ENCRYPTION_TYPE_XOR.to_le_bytes());
                workbook_stream[payload_start + 2..payload_start + 4].copy_from_slice(&key.to_le_bytes());
                workbook_stream[payload_start + 4..payload_start + 6]
                    .copy_from_slice(&verifier.to_le_bytes());

                let (ranges, _total) = collect_payload_ranges_after_offset(workbook_stream, encrypted_start)?;
                let xor_array = derive_xor_array(password);
                let key_bytes = key.to_le_bytes();
                for (range, start_pos) in ranges {
                    for (i, b) in workbook_stream[range.clone()].iter_mut().enumerate() {
                        let pos = start_pos + i;
                        let ks = xor_array[pos % xor_array.len()] ^ key_bytes[pos % 2];
                        *b ^= ks;
                    }
                }
                return Ok(());
            }
            BiffEncryption::Biff8Rc4 { .. } => {
                // Standard RC4 FILEPASS layout:
                // - wEncryptionType (2)
                // - major (2) == 1
                // - minor (2) == 1 (40-bit) or 2 (128-bit)
                // - salt (16)
                // - encrypted verifier (16)
                // - encrypted verifier hash (16)
                const EXPECTED_LEN: usize = 6 + 16 + 16 + 16;
                if record.data.len() < EXPECTED_LEN {
                    return Err(DecryptError::InvalidFilePass(format!(
                        "truncated FILEPASS RC4 payload (len={}, need at least {EXPECTED_LEN})",
                        record.data.len()
                    )));
                }

                // Respect the placeholder's minor version (key length) when possible.
                let minor =
                    u16::from_le_bytes([workbook_stream[payload_start + 4], workbook_stream[payload_start + 5]]);
                let key_len = match minor {
                    1 => 5usize,
                    2 => 16usize,
                    _ => {
                        return Err(DecryptError::UnsupportedEncryption(format!(
                            "unsupported FILEPASS RC4 minor version {minor} (expected 1 or 2)"
                        )))
                    }
                };

                // Deterministic salt/verifier so tests are reproducible.
                let salt: [u8; 16] = core::array::from_fn(|i| i as u8);
                let verifier: [u8; 16] = core::array::from_fn(|i| 0xA0u8.wrapping_add(i as u8));
                let verifier_hash: [u8; 16] = Md5::digest(verifier).into();

                let intermediate_key = rc4::derive_biff8_rc4_intermediate_key(password, &salt);
                let block_key = rc4::derive_biff8_rc4_block_key(&*intermediate_key, 0);
                let mut rc4 = Rc4::new(&block_key[..key_len]);
                drop(block_key);
                let mut buf = [0u8; 32];
                buf[..16].copy_from_slice(&verifier);
                buf[16..].copy_from_slice(&verifier_hash);
                rc4.apply_keystream(&mut buf);

                // Patch FILEPASS payload.
                workbook_stream[payload_start..payload_start + 2]
                    .copy_from_slice(&BIFF8_ENCRYPTION_TYPE_RC4.to_le_bytes());
                workbook_stream[payload_start + 2..payload_start + 4].copy_from_slice(&1u16.to_le_bytes()); // major
                workbook_stream[payload_start + 4..payload_start + 6]
                    .copy_from_slice(&minor.to_le_bytes());
                workbook_stream[payload_start + 6..payload_start + 22].copy_from_slice(&salt);
                workbook_stream[payload_start + 22..payload_start + 38].copy_from_slice(&buf[..16]);
                workbook_stream[payload_start + 38..payload_start + 54].copy_from_slice(&buf[16..]);

                // Encrypt record payload bytes after FILEPASS using an absolute-position mapping so
                // boundary bugs in the production decryptor are caught by roundtrip tests.
                let (ranges, total) = collect_payload_ranges_after_offset(workbook_stream, encrypted_start)?;
                let blocks = total.div_ceil(RC4_BLOCK_SIZE).max(1);
                let mut keystreams = Vec::<[u8; RC4_BLOCK_SIZE]>::with_capacity(blocks);
                for b in 0..blocks {
                    let block_key = rc4::derive_biff8_rc4_block_key(&*intermediate_key, b as u32);
                    let mut rc4 = Rc4::new(&block_key[..key_len]);
                    drop(block_key);
                    let mut ks = [0u8; RC4_BLOCK_SIZE];
                    rc4.apply_keystream(&mut ks);
                    keystreams.push(ks);
                }

                for (range, start_pos) in ranges {
                    for (i, b) in workbook_stream[range.clone()].iter_mut().enumerate() {
                        let abs = start_pos + i;
                        let block = abs / RC4_BLOCK_SIZE;
                        let off = abs % RC4_BLOCK_SIZE;
                        *b ^= keystreams[block][off];
                    }
                }

                return Ok(());
            }
            BiffEncryption::Biff8Rc4CryptoApi { .. } => {
                // RC4 CryptoAPI FILEPASS layout:
                // - wEncryptionType (2) == 0x0001
                // - wEncryptionSubType (2) == 0x0002
                // - dwEncryptionInfoLen (4)
                // - EncryptionInfo bytes
                //
                // For the self-consistency tests we generate deterministic EncryptionInfo
                // structures (salt/verifier), then encrypt record payload bytes after FILEPASS
                // using the CryptoAPI per-block RC4 keystream.
                //
                // Note: This is test-only; production `.xls` decryption is implemented by
                // `crate::biff::encryption` (see `cryptoapi` submodule).
                const CALG_RC4: u32 = 0x0000_6801;
                const SPIN_COUNT: u32 = 50_000;

                // Minimal EncryptionInfo payload sizes:
                // - Header: 32 bytes (no CSP name)
                // - Verifier: saltSize(4)+salt(16)+encVerifier(16)+hashSize(4)+encHash(20)=60
                // - EncryptionInfo: version(4)+flags(4)+headerSize(4)+header(32)+verifier(60)=104
                const ENC_HEADER_SIZE: usize = 32;
                const ENC_INFO_LEN: usize = 12 + ENC_HEADER_SIZE + 60;
                const FILEPASS_PAYLOAD_LEN: usize = 8 + ENC_INFO_LEN;

                if record.data.len() < FILEPASS_PAYLOAD_LEN {
                    return Err(DecryptError::InvalidFilePass(format!(
                        "BIFF8 RC4 CryptoAPI FILEPASS payload too short: expected at least {FILEPASS_PAYLOAD_LEN} bytes, got {}",
                        record.data.len()
                    )));
                }

                // Always use 128-bit RC4 for test streams.
                let key_len: usize = 16;
                let key_size_bits: u32 = (key_len as u32) * 8;

                // Deterministic salt/verifier so tests are reproducible.
                let salt: [u8; 16] = core::array::from_fn(|i| 0x10u8.wrapping_add(i as u8));
                let verifier_plain: [u8; 16] =
                    core::array::from_fn(|i| 0xF0u8.wrapping_sub(i as u8));

                use sha1::{Digest as _, Sha1};
                let digest = Sha1::digest(verifier_plain);
                let mut verifier_hash_plain = [0u8; 20];
                verifier_hash_plain.copy_from_slice(&digest);

                // Encrypt verifier + hash using the block-0 key.
                let key0 = cryptoapi::derive_biff8_cryptoapi_key(
                    cryptoapi::CALG_SHA1,
                    password,
                    &salt,
                    SPIN_COUNT,
                    0,
                    key_len,
                )?;
                let mut rc4 = Rc4::new(&key0);
                let mut verifier_buf = [0u8; 36];
                verifier_buf[..16].copy_from_slice(&verifier_plain);
                verifier_buf[16..].copy_from_slice(&verifier_hash_plain);
                rc4.apply_keystream(&mut verifier_buf);

                let mut encrypted_verifier = [0u8; 16];
                encrypted_verifier.copy_from_slice(&verifier_buf[..16]);
                let mut encrypted_verifier_hash = [0u8; 20];
                encrypted_verifier_hash.copy_from_slice(&verifier_buf[16..]);

                // EncryptionHeader (32 bytes) [MS-OFFCRYPTO].
                let mut enc_header = Vec::<u8>::new();
                enc_header.extend_from_slice(&0u32.to_le_bytes()); // Flags
                enc_header.extend_from_slice(&0u32.to_le_bytes()); // SizeExtra
                enc_header.extend_from_slice(&CALG_RC4.to_le_bytes()); // AlgID
                enc_header.extend_from_slice(&cryptoapi::CALG_SHA1.to_le_bytes()); // AlgIDHash
                enc_header.extend_from_slice(&key_size_bits.to_le_bytes()); // KeySize bits
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

                // EncryptionInfo:
                //   u16 MajorVersion, u16 MinorVersion, u32 Flags, u32 HeaderSize, EncryptionHeader, EncryptionVerifier.
                let mut enc_info = Vec::<u8>::new();
                enc_info.extend_from_slice(&4u16.to_le_bytes()); // Major
                enc_info.extend_from_slice(&2u16.to_le_bytes()); // Minor
                enc_info.extend_from_slice(&0u32.to_le_bytes()); // Flags
                enc_info.extend_from_slice(&(enc_header.len() as u32).to_le_bytes()); // HeaderSize
                enc_info.extend_from_slice(&enc_header);
                enc_info.extend_from_slice(&enc_verifier);

                // FILEPASS payload.
                let mut filepass_payload = Vec::<u8>::new();
                filepass_payload.extend_from_slice(&BIFF8_ENCRYPTION_TYPE_RC4.to_le_bytes());
                filepass_payload.extend_from_slice(&BIFF8_RC4_SUBTYPE_CRYPTOAPI.to_le_bytes());
                filepass_payload.extend_from_slice(&(enc_info.len() as u32).to_le_bytes());
                filepass_payload.extend_from_slice(&enc_info);

                debug_assert_eq!(filepass_payload.len(), FILEPASS_PAYLOAD_LEN);
                workbook_stream[payload_start..payload_start + filepass_payload.len()]
                    .copy_from_slice(&filepass_payload);
                // Zero any remaining bytes in the placeholder so expected streams are stable.
                for b in workbook_stream[payload_start + filepass_payload.len()..payload_end].iter_mut() {
                    *b = 0;
                }

                // Encrypt record payload bytes after FILEPASS using an absolute-position mapping.
                let (ranges, total) = collect_payload_ranges_after_offset(workbook_stream, encrypted_start)?;
                let blocks = total.div_ceil(RC4_BLOCK_SIZE).max(1);
                let mut keystreams = Vec::<[u8; RC4_BLOCK_SIZE]>::with_capacity(blocks);
                for b in 0..blocks {
                    let key = cryptoapi::derive_biff8_cryptoapi_key(
                        cryptoapi::CALG_SHA1,
                        password,
                        &salt,
                        SPIN_COUNT,
                        b as u32,
                        key_len,
                    )?;
                    let mut rc4 = Rc4::new(&key);
                    let mut ks = [0u8; RC4_BLOCK_SIZE];
                    rc4.apply_keystream(&mut ks);
                    keystreams.push(ks);
                }

                for (range, start_pos) in ranges {
                    for (i, b) in workbook_stream[range.clone()].iter_mut().enumerate() {
                        let abs = start_pos + i;
                        let block = abs / RC4_BLOCK_SIZE;
                        let off = abs % RC4_BLOCK_SIZE;
                        *b ^= keystreams[block][off];
                    }
                }

                return Ok(());
            }
            BiffEncryption::Biff8Rc4CryptoApiLegacy { .. } => {
                return Err(DecryptError::UnsupportedEncryption(
                    "BIFF8 RC4 CryptoAPI legacy encryption is not supported by the test encryptor"
                        .to_string(),
                ));
            }
        }
    }

    Err(DecryptError::NoFilePass)
}

#[cfg(test)]
mod tests {
    use super::*;

    const RECORD_BOF: u16 = 0x0809;
    const RECORD_EOF: u16 = 0x000A;
    const RECORD_DUMMY: u16 = 0x00FC;
    fn record(record_id: u16, payload: &[u8]) -> Vec<u8> {
        let mut out = Vec::with_capacity(4 + payload.len());
        out.extend_from_slice(&record_id.to_le_bytes());
        out.extend_from_slice(&(payload.len() as u16).to_le_bytes());
        out.extend_from_slice(payload);
        out
    }

    fn make_filepass_rc4_record(password: &str, salt: [u8; 16], key_len: usize) -> Vec<u8> {
        let (major, minor) = match key_len {
            5 => (1u16, 1u16),
            16 => (1u16, 2u16),
            _ => panic!("unsupported key_len"),
        };

        let verifier: [u8; 16] = (0..16u8)
            .map(|b| b.wrapping_mul(31))
            .collect::<Vec<_>>()
            .as_slice()
            .try_into()
            .unwrap();
        let verifier_hash: [u8; 16] = Md5::digest(verifier).into();

        let intermediate_key = rc4::derive_biff8_rc4_intermediate_key(password, &salt);
        let block_key = rc4::derive_biff8_rc4_block_key(&*intermediate_key, 0);
        let mut rc4 = Rc4::new(&block_key[..key_len]);
        drop(block_key);

        let mut buf = [0u8; 32];
        buf[0..16].copy_from_slice(&verifier);
        buf[16..32].copy_from_slice(&verifier_hash);
        rc4.apply_keystream(&mut buf);

        let encrypted_verifier = &buf[0..16];
        let encrypted_verifier_hash = &buf[16..32];

        let mut payload = Vec::with_capacity(54);
        payload.extend_from_slice(&0x0001u16.to_le_bytes()); // wEncryptionType = RC4
        payload.extend_from_slice(&major.to_le_bytes());
        payload.extend_from_slice(&minor.to_le_bytes());
        payload.extend_from_slice(&salt);
        payload.extend_from_slice(encrypted_verifier);
        payload.extend_from_slice(encrypted_verifier_hash);

        record(records::RECORD_FILEPASS, &payload)
    }

    fn encrypt_record_payloads_in_place(
        workbook_stream: &mut [u8],
        encrypted_start: usize,
        intermediate_key: &[u8; 16],
        key_len: usize,
    ) -> Result<(), String> {
        // Manual encryption reference: apply RC4 stream to record payload bytes only.
        let mut block: u32 = 0;
        let mut pos_in_block: usize = 0;
        let mut block_key = rc4::derive_biff8_rc4_block_key(intermediate_key, block);
        let mut cipher = Rc4::new(&block_key[..key_len]);

        let mut offset = encrypted_start;
        while offset < workbook_stream.len() {
            if workbook_stream.len() - offset < 4 {
                return Err("truncated record header".to_string());
            }
            let len = u16::from_le_bytes([workbook_stream[offset + 2], workbook_stream[offset + 3]])
                as usize;
            let data_start = offset + 4;
            let data_end = data_start + len;
            if data_end > workbook_stream.len() {
                return Err("record extends past end".to_string());
            }

            let mut data = &mut workbook_stream[data_start..data_end];
            while !data.is_empty() {
                if pos_in_block == RC4_BLOCK_SIZE {
                    block = block.wrapping_add(1);
                    pos_in_block = 0;
                    block_key = rc4::derive_biff8_rc4_block_key(intermediate_key, block);
                    cipher = Rc4::new(&block_key[..key_len]);
                }
                let remaining_in_block = RC4_BLOCK_SIZE - pos_in_block;
                let n = remaining_in_block.min(data.len());
                let (chunk, rest) = data.split_at_mut(n);
                cipher.apply_keystream(chunk);
                pos_in_block += n;
                data = rest;
            }

            offset = data_end;
        }
        Ok(())
    }

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
    fn parses_biff5_xor_filepass_with_biff8_style_header() {
        // Some BIFF5 writers emit a BIFF8-style FILEPASS payload:
        //   wEncryptionType (u16) + key (u16) + verifier (u16)
        let payload = [
            0x00, 0x00, // wEncryptionType (XOR)
            0x34, 0x12, // key
            0x78, 0x56, // verifier
        ];
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
                filepass_payload: Zeroizing::new(payload.to_vec())
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
                filepass_payload: Zeroizing::new(payload.to_vec())
            }
        );
    }

    #[test]
    fn parses_biff8_rc4_cryptoapi_legacy_filepass() {
        // Some BIFF8 RC4 CryptoAPI workbooks use an older FILEPASS layout where the second
        // field is `wEncryptionInfo == 0x0004` (rather than `wEncryptionSubType == 0x0002`).
        let payload = [
            0x01, 0x00, // wEncryptionType
            0x04, 0x00, // wEncryptionInfo (legacy CryptoAPI)
            0xDE, 0xAD, 0xBE, 0xEF,
        ];
        let parsed = parse_filepass_record(BiffVersion::Biff8, &payload).expect("parse");
        assert_eq!(
            parsed,
            BiffEncryption::Biff8Rc4CryptoApiLegacy {
                filepass_payload: Zeroizing::new(payload.to_vec()),
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

    #[test]
    fn errors_on_oversized_filepass_payload() {
        let payload = vec![0u8; MAX_FILEPASS_PAYLOAD_BYTES + 1];
        let err = parse_filepass_record(BiffVersion::Biff8, &payload).expect_err("expected err");
        assert!(matches!(err, DecryptError::SizeLimitExceeded { .. }));
    }

    #[test]
    fn rc4_decrypt_allows_empty_password_when_file_was_encrypted_with_empty_password() {
        let password = "";
        let salt: [u8; 16] = (0..16u8).collect::<Vec<_>>()[..].try_into().unwrap();
        let key_len = 5;

        let mut plain = Vec::new();
        plain.extend_from_slice(&record(RECORD_BOF, &[0u8; 16]));
        let filepass_record = make_filepass_rc4_record(password, salt, key_len);
        let filepass_offset = plain.len();
        plain.extend_from_slice(&filepass_record);

        let payload = vec![0x42u8; 64];
        plain.extend_from_slice(&record(RECORD_DUMMY, &payload));
        plain.extend_from_slice(&record(RECORD_EOF, &[]));

        let mut encrypted = plain.clone();
        let filepass_len = u16::from_le_bytes([
            encrypted[filepass_offset + 2],
            encrypted[filepass_offset + 3],
        ]) as usize;
        let encrypted_start = filepass_offset + 4 + filepass_len;

        let intermediate_key = rc4::derive_biff8_rc4_intermediate_key(password, &salt);
        encrypt_record_payloads_in_place(&mut encrypted, encrypted_start, &*intermediate_key, key_len)
            .expect("encrypt");

        decrypt_workbook_stream(&mut encrypted, password).expect("decrypt");
        assert_eq!(encrypted, plain);
    }

    #[test]
    fn xor_method1_decrypt_allows_empty_password_when_file_was_encrypted_with_empty_password() {
        let password = "";
        let bof_payload = [0x00, 0x06, 0x05, 0x00];

        let key = create_xor_key_method1(&[]);
        let verifier = create_password_verifier_method1(&[]);
        let filepass_payload = [
            0x00, 0x00, // wEncryptionType (XOR)
            key.to_le_bytes()[0],
            key.to_le_bytes()[1],
            verifier.to_le_bytes()[0],
            verifier.to_le_bytes()[1],
        ];

        let plain = [
            record(records::RECORD_BOF_BIFF8, &bof_payload),
            record(records::RECORD_FILEPASS, &filepass_payload),
            record(RECORD_DUMMY, &dummy_payload(64, 0x42)),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let mut encrypted = plain.clone();
        let encrypted_start = filepass_payload_range(&encrypted).end;
        let xor_array = create_xor_array_method1(&[], key);
        encrypt_payloads_after_filepass_xor_method1(&mut encrypted, encrypted_start, &xor_array)
            .expect("encrypt");

        decrypt_workbook_stream(&mut encrypted, password).expect("decrypt");
        assert_eq!(encrypted, plain);
    }

    #[test]
    fn xor_method1_decrypt_uses_method2_password_byte_encoding() {
        // Some BIFF XOR "Method 1" implementations derive password bytes from UTF-16 code units
        // using MS-OFFCRYPTO 2.3.7.4 "method 2": use low byte unless it is zero, else use high
        // byte.
        //
        // This test uses U+0100 ("Ā") which is not representable in Windows-1252 (common Excel ANSI
        // encoding), and whose UTF-16 low byte is 0. This ensures the decryptor's candidate
        // generation must consider method-2 bytes to successfully match the stored key/verifier.
        let password = "Ā";
        let bof_payload = [0x00, 0x06, 0x05, 0x00];

        // Method 2 byte encoding for U+0100 => [0x01] (low byte is 0 so we use high byte).
        let password_bytes = [0x01u8];

        let key = create_xor_key_method1(&password_bytes);
        let verifier = create_password_verifier_method1(&password_bytes);
        let filepass_payload = [
            0x00, 0x00, // wEncryptionType (XOR)
            key.to_le_bytes()[0],
            key.to_le_bytes()[1],
            verifier.to_le_bytes()[0],
            verifier.to_le_bytes()[1],
        ];

        let plain = [
            record(records::RECORD_BOF_BIFF8, &bof_payload),
            record(records::RECORD_FILEPASS, &filepass_payload),
            record(RECORD_DUMMY, &dummy_payload(64, 0x99)),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let mut encrypted = plain.clone();
        let encrypted_start = filepass_payload_range(&encrypted).end;
        let xor_array = create_xor_array_method1(&password_bytes, key);
        encrypt_payloads_after_filepass_xor_method1(&mut encrypted, encrypted_start, &xor_array)
            .expect("encrypt");

        decrypt_workbook_stream(&mut encrypted, password).expect("decrypt");
        assert_eq!(encrypted, plain);
    }

    #[test]
    fn decrypt_workbook_stream_returns_no_filepass_when_missing() {
        let bof_payload = [0x00, 0x06, 0x05, 0x00];
        let mut stream = [
            record(records::RECORD_BOF_BIFF8, &bof_payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let err = decrypt_workbook_stream(&mut stream, "pw").expect_err("expected error");
        assert_eq!(err, DecryptError::NoFilePass);
    }

    #[test]
    fn decrypt_workbook_stream_ignores_filepass_after_next_bof() {
        let bof_payload = [0x00, 0x06, 0x05, 0x00];
        let mut stream = [
            record(records::RECORD_BOF_BIFF8, &bof_payload),
            // Next BOF indicates the start of another substream; FILEPASS must not be scanned there.
            record(records::RECORD_BOF_BIFF8, &bof_payload),
            record(records::RECORD_FILEPASS, &[]),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let err = decrypt_workbook_stream(&mut stream, "pw").expect_err("expected error");
        assert_eq!(err, DecryptError::NoFilePass);
    }

    #[test]
    fn decrypt_workbook_stream_rejects_stream_without_bof() {
        let mut stream = [
            record(0x0001, &[0xAA]),
            record(records::RECORD_FILEPASS, &[]),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let err = decrypt_workbook_stream(&mut stream, "pw").expect_err("expected error");
        assert!(matches!(err, DecryptError::InvalidFilePass(_)));
    }

    #[test]
    fn decrypt_workbook_stream_parses_filepass_and_dispatches_biff8_xor() {
        let bof_payload = [0x00, 0x06, 0x05, 0x00];
        let wrong_verifier = xor::xor_password_verifier("not-pw");
        let filepass_payload = [
            0x00, 0x00, // wEncryptionType (XOR)
            0x34, 0x12, // key
            wrong_verifier.to_le_bytes()[0],
            wrong_verifier.to_le_bytes()[1],
        ];
        let mut stream = [
            record(records::RECORD_BOF_BIFF8, &bof_payload),
            record(records::RECORD_FILEPASS, &filepass_payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let err = decrypt_workbook_stream(&mut stream, "pw").expect_err("expected error");
        assert_eq!(err, DecryptError::WrongPassword);
    }

    #[test]
    fn decrypt_workbook_stream_parses_filepass_and_dispatches_biff5_xor() {
        let bof_payload = [0x00, 0x05, 0x05, 0x00];
        let wrong_verifier = xor::xor_password_verifier("not-pw");
        let filepass_payload = [
            0x34, 0x12, // key
            wrong_verifier.to_le_bytes()[0],
            wrong_verifier.to_le_bytes()[1],
        ];
        let mut stream = [
            record(records::RECORD_BOF_BIFF5, &bof_payload),
            record(records::RECORD_FILEPASS, &filepass_payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let err = decrypt_workbook_stream(&mut stream, "pw").expect_err("expected error");
        assert_eq!(err, DecryptError::WrongPassword);
    }

    #[test]
    fn decrypt_workbook_stream_returns_invalid_filepass_when_payload_truncated() {
        let bof_payload = [0x00, 0x06, 0x05, 0x00];
        let filepass_payload = [0x00]; // truncated wEncryptionType
        let mut stream = [
            record(records::RECORD_BOF_BIFF8, &bof_payload),
            record(records::RECORD_FILEPASS, &filepass_payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let err = decrypt_workbook_stream(&mut stream, "pw").expect_err("expected error");
        assert!(matches!(err, DecryptError::InvalidFilePass(_)));
    }

    #[test]
    fn rc4_decrypt_respects_record_headers_and_block_boundaries() {
        let password = "secret";
        let salt: [u8; 16] = (0..16u8).collect::<Vec<_>>()[..].try_into().unwrap();
        let key_len = 16;

        // Build a stream where the first encrypted record is exactly one block (1024 bytes), so the
        // next record's first payload byte must use block 1 key. If the decryptor incorrectly counts
        // record headers as encrypted data, it will misalign and fail this test.
        let mut plain = Vec::new();
        plain.extend_from_slice(&record(RECORD_BOF, &[0u8; 16]));
        let filepass_record = make_filepass_rc4_record(password, salt, key_len);
        let filepass_offset = plain.len();
        plain.extend_from_slice(&filepass_record);

        let record_a_payload = vec![0x42u8; 1024];
        plain.extend_from_slice(&record(RECORD_DUMMY, &record_a_payload));
        let record_b_payload = vec![0x99u8];
        plain.extend_from_slice(&record(RECORD_DUMMY, &record_b_payload));
        plain.extend_from_slice(&record(RECORD_EOF, &[]));

        // Encrypt in place using a reference implementation.
        let mut encrypted = plain.clone();
        let filepass_len = u16::from_le_bytes([
            encrypted[filepass_offset + 2],
            encrypted[filepass_offset + 3],
        ]) as usize;
        let encrypted_start = filepass_offset + 4 + filepass_len;

        let intermediate_key = rc4::derive_biff8_rc4_intermediate_key(password, &salt);
        encrypt_record_payloads_in_place(
            &mut encrypted,
            encrypted_start,
            &*intermediate_key,
            key_len,
        )
        .expect("encrypt");

        // Now decrypt and ensure we get the original plaintext.
        decrypt_workbook_stream(&mut encrypted, password).expect("decrypt");
        assert_eq!(encrypted, plain);
    }

    #[test]
    fn rc4_decrypt_large_synthetic_stream_is_deterministic() {
        let password = "benchmark-password";
        let salt: [u8; 16] = [
            0x10, 0x20, 0x30, 0x40, 0x50, 0x60, 0x70, 0x80, 0x90, 0xA0, 0xB0, 0xC0, 0xD0, 0xE0,
            0xF0, 0x00,
        ];
        let key_len = 16;

        // Build a synthetic BIFF8-like workbook stream with many small records to exercise the
        // per-record header skipping and 1024-byte RC4 block re-keying logic.
        let mut plain = Vec::new();
        plain.extend_from_slice(&record(RECORD_BOF, &[0u8; 16]));
        let filepass_record = make_filepass_rc4_record(password, salt, key_len);
        let filepass_offset = plain.len();
        plain.extend_from_slice(&filepass_record);

        let mut seed: u32 = 0x1234_5678;
        for i in 0..4000u32 {
            // Simple LCG for deterministic payload bytes.
            seed = seed.wrapping_mul(1664525).wrapping_add(1013904223);
            let len = 1 + (seed as usize % 64);
            let mut payload = Vec::with_capacity(len);
            let mut x = seed ^ i;
            for _ in 0..len {
                x = x.wrapping_mul(1103515245).wrapping_add(12345);
                payload.push((x >> 16) as u8);
            }
            plain.extend_from_slice(&record(RECORD_DUMMY, &payload));
        }
        plain.extend_from_slice(&record(RECORD_EOF, &[]));

        // Encrypt the record payloads after FILEPASS.
        let mut encrypted = plain.clone();
        let filepass_len = u16::from_le_bytes([
            encrypted[filepass_offset + 2],
            encrypted[filepass_offset + 3],
        ]) as usize;
        let encrypted_start = filepass_offset + 4 + filepass_len;
        let intermediate_key = rc4::derive_biff8_rc4_intermediate_key(password, &salt);
        encrypt_record_payloads_in_place(
            &mut encrypted,
            encrypted_start,
            &*intermediate_key,
            key_len,
        )
        .expect("encrypt");

        // Decrypt and assert we recover the original bytes.
        decrypt_workbook_stream(&mut encrypted, password).expect("decrypt");
        assert_eq!(encrypted, plain);
    }

    #[test]
    fn rc4_decrypt_ignores_trailing_padding_bytes_after_last_record() {
        let password = "secret";
        let salt: [u8; 16] = (0..16u8).collect::<Vec<_>>()[..].try_into().unwrap();
        let key_len = 16;

        let mut plain = Vec::new();
        plain.extend_from_slice(&record(RECORD_BOF, &[0u8; 16]));
        let filepass_record = make_filepass_rc4_record(password, salt, key_len);
        let filepass_offset = plain.len();
        plain.extend_from_slice(&filepass_record);
        plain.extend_from_slice(&record(RECORD_DUMMY, &dummy_payload(64, 0xAA)));
        plain.extend_from_slice(&record(RECORD_EOF, &[]));

        // Encrypt record payloads after FILEPASS.
        let mut encrypted = plain.clone();
        let filepass_len = u16::from_le_bytes([
            encrypted[filepass_offset + 2],
            encrypted[filepass_offset + 3],
        ]) as usize;
        let encrypted_start = filepass_offset + 4 + filepass_len;
        let intermediate_key = rc4::derive_biff8_rc4_intermediate_key(password, &salt);
        encrypt_record_payloads_in_place(
            &mut encrypted,
            encrypted_start,
            &*intermediate_key,
            key_len,
        )
        .expect("encrypt");

        // Append trailing bytes that do not form a full BIFF record header. Some writers include
        // such padding after the final EOF record.
        let padding = [0xDEu8, 0xADu8, 0xBEu8];
        encrypted.extend_from_slice(&padding);
        let mut expected = plain.clone();
        expected.extend_from_slice(&padding);

        decrypt_workbook_stream(&mut encrypted, password).expect("decrypt");
        assert_eq!(encrypted, expected);
    }

    fn filepass_payload_range(stream: &[u8]) -> std::ops::Range<usize> {
        let mut offset = 0usize;
        while offset + 4 <= stream.len() {
            let record_id = u16::from_le_bytes([stream[offset], stream[offset + 1]]);
            let len = u16::from_le_bytes([stream[offset + 2], stream[offset + 3]]) as usize;
            let data_start = offset + 4;
            let data_end = data_start + len;
            if record_id == records::RECORD_FILEPASS {
                return data_start..data_end;
            }
            offset = data_end;
        }
        panic!("FILEPASS record not found")
    }

    fn dummy_payload(len: usize, seed: u8) -> Vec<u8> {
        // Deterministic, "busy" payload so cipher alignment bugs are visible.
        (0..len)
            .map(|i| seed.wrapping_add((i as u8).wrapping_mul(31)).wrapping_add((i >> 8) as u8))
            .collect()
    }

    fn encrypt_payloads_after_filepass_xor_method1(
        workbook_stream: &mut [u8],
        start_offset: usize,
        xor_array: &[u8; 16],
    ) -> Result<(), DecryptError> {
        let mut offset = start_offset;
        while offset < workbook_stream.len() {
            let remaining = workbook_stream.len().saturating_sub(offset);
            if remaining < 4 {
                // Some writers include trailing padding bytes after the final EOF record. Those bytes
                // are not part of any record header/payload and should be ignored.
                break;
            }

            let record_id = u16::from_le_bytes([workbook_stream[offset], workbook_stream[offset + 1]]);
            let len =
                u16::from_le_bytes([workbook_stream[offset + 2], workbook_stream[offset + 3]]) as usize;

            let data_start = offset
                .checked_add(4)
                .ok_or_else(|| DecryptError::InvalidFilePass("BIFF record offset overflow".to_string()))?;
            let data_end = data_start
                .checked_add(len)
                .ok_or_else(|| DecryptError::InvalidFilePass("BIFF record length overflow".to_string()))?;
            if data_end > workbook_stream.len() {
                return Err(DecryptError::InvalidFilePass(format!(
                    "BIFF record 0x{record_id:04X} at offset {offset} extends past end of stream while encrypting XOR (len={}, end={data_end})",
                    workbook_stream.len()
                )));
            }

            // Mirror the decryptor's per-record skip rules for XOR obfuscation.
            let mut encrypt_from = 0usize;
            let skip_entire_payload = matches!(
                record_id,
                records::RECORD_BOF_BIFF8
                    | records::RECORD_BOF_BIFF5
                    | records::RECORD_FILEPASS
                    | RECORD_INTERFACEHDR
                    | RECORD_FILELOCK
                    | RECORD_USREXCL
                    | RECORD_RRDINFO
                    | RECORD_RRDHEAD
            );

            if !skip_entire_payload {
                if record_id == RECORD_BOUNDSHEET {
                    // BoundSheet.lbPlyPos MUST NOT be encrypted.
                    encrypt_from = 4.min(len);
                }

                let payload = &mut workbook_stream[data_start..data_end];
                for i in encrypt_from..payload.len() {
                    let abs_pos = data_start + i;
                    let mut value = payload[i];
                    value = value.rotate_left(5);
                    value ^= xor_array[abs_pos % 16];
                    payload[i] = value;
                }
            }

            offset = data_end;
        }

        Ok(())
    }

    #[test]
    fn xor_method1_password_is_truncated_to_15_bytes() {
        // Excel legacy XOR "method1" passwords are effectively limited to 15 bytes. Extra characters
        // are ignored (i.e. only the first 15 bytes are significant).
        let base = "0123456789ABCDE"; // 15 ASCII chars
        let password = format!("{base}X"); // 16th char should be ignored
        let password_same_prefix = format!("{base}Y");
        let wrong_password = "0123456789ABCDZ"; // differs within the first 15 chars

        let candidates = xor_password_byte_candidates(&password);
        assert_eq!(candidates[0].as_slice(), base.as_bytes());
        assert_eq!(candidates[1].as_slice(), base.as_bytes());

        let password_bytes = base.as_bytes();
        let key = create_xor_key_method1(password_bytes);
        let verifier = create_password_verifier_method1(password_bytes);
        let xor_array = create_xor_array_method1(password_bytes, key);

        // BIFF8 BOF payload (4 bytes) is sufficient for stream detection in this test.
        let bof_payload = [0x00, 0x06, 0x05, 0x00];
        let filepass_payload = [
            0x00, 0x00, // wEncryptionType (XOR)
            key.to_le_bytes()[0],
            key.to_le_bytes()[1],
            verifier.to_le_bytes()[0],
            verifier.to_le_bytes()[1],
        ];

        let mut plain = Vec::new();
        plain.extend_from_slice(&record(records::RECORD_BOF_BIFF8, &bof_payload));
        let filepass_offset = plain.len();
        plain.extend_from_slice(&record(records::RECORD_FILEPASS, &filepass_payload));
        plain.extend_from_slice(&record(RECORD_DUMMY, &dummy_payload(64, 0x10)));
        plain.extend_from_slice(&record(RECORD_EOF, &[]));

        let encrypted_start = filepass_offset + 4 + filepass_payload.len();
        let mut encrypted = plain.clone();
        encrypt_payloads_after_filepass_xor_method1(&mut encrypted, encrypted_start, &xor_array)
            .expect("encrypt");

        for candidate in [password.as_str(), password_same_prefix.as_str(), base] {
            let mut decrypted = encrypted.clone();
            decrypt_workbook_stream(&mut decrypted, candidate).expect("decrypt");
            assert_eq!(decrypted, plain);
        }

        let mut buf = encrypted.clone();
        let err = decrypt_workbook_stream(&mut buf, wrong_password).expect_err("wrong password");
        assert_eq!(err, DecryptError::WrongPassword);
    }

    #[test]
    fn xor_method1_allows_empty_password_when_file_was_encrypted_with_empty_password() {
        // The spec-defined XOR algorithm can be applied to an empty password (some third-party
        // writers may emit such files). Ensure our method1 path handles it to avoid falling back to
        // the legacy deterministic fixture scheme.
        let password = "";
        let password_bytes: &[u8] = &[];
        let key = create_xor_key_method1(password_bytes);
        let verifier = create_password_verifier_method1(password_bytes);
        let xor_array = create_xor_array_method1(password_bytes, key);

        assert_eq!(key, 0);
        assert_eq!(verifier, 0xCE4B);

        let bof_payload = [0x00, 0x06, 0x05, 0x00];
        let filepass_payload = [
            0x00, 0x00, // wEncryptionType (XOR)
            key.to_le_bytes()[0],
            key.to_le_bytes()[1],
            verifier.to_le_bytes()[0],
            verifier.to_le_bytes()[1],
        ];

        let mut plain = Vec::new();
        plain.extend_from_slice(&record(records::RECORD_BOF_BIFF8, &bof_payload));
        let filepass_offset = plain.len();
        plain.extend_from_slice(&record(records::RECORD_FILEPASS, &filepass_payload));
        plain.extend_from_slice(&record(RECORD_DUMMY, &dummy_payload(64, 0x11)));
        plain.extend_from_slice(&record(RECORD_EOF, &[]));

        let encrypted_start = filepass_offset + 4 + filepass_payload.len();
        let mut encrypted = plain.clone();
        encrypt_payloads_after_filepass_xor_method1(&mut encrypted, encrypted_start, &xor_array)
            .expect("encrypt");

        let mut decrypted = encrypted.clone();
        decrypt_workbook_stream(&mut decrypted, password).expect("decrypt");
        assert_eq!(decrypted, plain);

        let mut buf = encrypted.clone();
        let err = decrypt_workbook_stream(&mut buf, "not-empty").expect_err("wrong password");
        assert_eq!(err, DecryptError::WrongPassword);
    }

    #[test]
    fn xor_encrypt_decrypt_roundtrip_across_records() {
        let password = "pw";
        let bof_payload = [0x00, 0x06, 0x05, 0x00];

        // BIFF8 XOR FILEPASS placeholder: type + key + verifier.
        let filepass_placeholder = [0u8; 6];

        // Include enough payload bytes to exercise record-boundary logic.
        let r1 = record(0x00FC, &dummy_payload(1000, 0x10));
        let r2 = record(0x00FD, &dummy_payload(80, 0x20));

        let mut plain = [
            record(records::RECORD_BOF_BIFF8, &bof_payload),
            record(records::RECORD_FILEPASS, &filepass_placeholder),
            r1,
            r2,
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        // Encrypt using the test-only helper (fills FILEPASS, encrypts payload bytes after it).
        let mut encrypted = plain.clone();
        encrypt_workbook_stream_for_test(&mut encrypted, password).expect("encrypt");

        // Patch expected plaintext FILEPASS payload from the encrypted stream (FILEPASS is plaintext).
        let range = filepass_payload_range(&plain);
        plain[range.clone()].copy_from_slice(&encrypted[range]);

        decrypt_workbook_stream(&mut encrypted, password).expect("decrypt");
        assert_eq!(encrypted, plain);
    }

    #[test]
    fn xor_encrypt_decrypt_roundtrip_biff5_across_records() {
        let password = "pw";
        // BIFF5 BOF payload: version=0x0500, dt arbitrary.
        let bof_payload = [0x00, 0x05, 0x05, 0x00];

        // BIFF5 XOR FILEPASS placeholder: key + verifier.
        let filepass_placeholder = [0u8; 4];

        let r1 = record(0x00FC, &dummy_payload(1000, 0x11));
        let r2 = record(0x00FD, &dummy_payload(80, 0x22));

        let mut plain = [
            record(records::RECORD_BOF_BIFF5, &bof_payload),
            record(records::RECORD_FILEPASS, &filepass_placeholder),
            r1,
            r2,
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let mut encrypted = plain.clone();
        encrypt_workbook_stream_for_test(&mut encrypted, password).expect("encrypt");

        let range = filepass_payload_range(&plain);
        plain[range.clone()].copy_from_slice(&encrypted[range]);

        decrypt_workbook_stream(&mut encrypted, password).expect("decrypt");
        assert_eq!(encrypted, plain);
    }

    #[test]
    fn rc4_encrypt_decrypt_roundtrip_crosses_1024_boundary_mid_record() {
        let password = "pw";
        let bof_payload = [0x00, 0x06, 0x05, 0x00];

        // BIFF8 RC4 FILEPASS placeholder (Standard Encryption / major=1, minor=2 => 128-bit key).
        let mut filepass_placeholder = vec![0u8; 6 + 16 + 16 + 16];
        filepass_placeholder[0..2].copy_from_slice(&BIFF8_ENCRYPTION_TYPE_RC4.to_le_bytes());
        filepass_placeholder[2..4].copy_from_slice(&1u16.to_le_bytes()); // major
        filepass_placeholder[4..6].copy_from_slice(&2u16.to_le_bytes()); // minor (128-bit)

        // Chosen so record2 crosses the 1024-byte block boundary: 1000 + 80.
        let r1 = record(0x00FC, &dummy_payload(1000, 0x30));
        let r2 = record(0x00FD, &dummy_payload(80, 0x40));

        let mut plain = [
            record(records::RECORD_BOF_BIFF8, &bof_payload),
            record(records::RECORD_FILEPASS, &filepass_placeholder),
            r1,
            r2,
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let mut encrypted = plain.clone();
        encrypt_workbook_stream_for_test(&mut encrypted, password).expect("encrypt");

        let range = filepass_payload_range(&plain);
        plain[range.clone()].copy_from_slice(&encrypted[range]);

        decrypt_workbook_stream(&mut encrypted, password).expect("decrypt");
        assert_eq!(encrypted, plain);
    }

    #[test]
    fn rc4_encrypt_decrypt_roundtrip_40_bit_key_crosses_1024_boundary_mid_record() {
        let password = "pw";
        let bof_payload = [0x00, 0x06, 0x05, 0x00];

        // BIFF8 RC4 FILEPASS placeholder (Standard Encryption / major=1, minor=1 => 40-bit key).
        let mut filepass_placeholder = vec![0u8; 6 + 16 + 16 + 16];
        filepass_placeholder[0..2].copy_from_slice(&BIFF8_ENCRYPTION_TYPE_RC4.to_le_bytes());
        filepass_placeholder[2..4].copy_from_slice(&1u16.to_le_bytes()); // major
        filepass_placeholder[4..6].copy_from_slice(&1u16.to_le_bytes()); // minor (40-bit)

        // Chosen so record2 crosses the 1024-byte block boundary: 1000 + 80.
        let r1 = record(0x00FC, &dummy_payload(1000, 0x31));
        let r2 = record(0x00FD, &dummy_payload(80, 0x41));

        let mut plain = [
            record(records::RECORD_BOF_BIFF8, &bof_payload),
            record(records::RECORD_FILEPASS, &filepass_placeholder),
            r1,
            r2,
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let mut encrypted = plain.clone();
        encrypt_workbook_stream_for_test(&mut encrypted, password).expect("encrypt");

        let range = filepass_payload_range(&plain);
        plain[range.clone()].copy_from_slice(&encrypted[range]);

        decrypt_workbook_stream(&mut encrypted, password).expect("decrypt");
        assert_eq!(encrypted, plain);
    }

    #[test]
    fn rc4_cryptoapi_encrypt_decrypt_roundtrip_crosses_1024_boundary_mid_record() {
        let password = "pw";
        let bof_payload = [0x00, 0x06, 0x05, 0x00];

        // BIFF8 RC4 CryptoAPI FILEPASS placeholder. The test helper overwrites the full payload
        // deterministically, but it needs the correct type/subtype so `parse_filepass_record`
        // selects CryptoAPI.
        let mut filepass_placeholder = vec![0u8; 112];
        filepass_placeholder[0..2].copy_from_slice(&BIFF8_ENCRYPTION_TYPE_RC4.to_le_bytes());
        filepass_placeholder[2..4].copy_from_slice(&BIFF8_RC4_SUBTYPE_CRYPTOAPI.to_le_bytes());

        // Chosen so record2 crosses the 1024-byte block boundary: 1000 + 80.
        let r1 = record(0x00FC, &dummy_payload(1000, 0x50));
        let r2 = record(0x00FD, &dummy_payload(80, 0x60));

        let mut plain = [
            record(records::RECORD_BOF_BIFF8, &bof_payload),
            record(records::RECORD_FILEPASS, &filepass_placeholder),
            r1,
            r2,
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let mut encrypted = plain.clone();
        encrypt_workbook_stream_for_test(&mut encrypted, password).expect("encrypt");

        let range = filepass_payload_range(&plain);
        plain[range.clone()].copy_from_slice(&encrypted[range.clone()]);
        decrypt_workbook_stream(&mut encrypted, password).expect("decrypt");
        assert_eq!(encrypted, plain);
    }

    #[test]
    fn rc4_standard_password_is_truncated_to_15_utf16_code_units() {
        // Excel 97-2003 Standard Encryption only considers the first 15 UTF-16 code units of the
        // password.
        let base = "0123456789ABCDE"; // 15 ASCII chars = 15 UTF-16 code units
        let password = format!("{base}X"); // 16th char should be ignored by the cipher
        let password_same_prefix = format!("{base}Y");
        let wrong_password = "0123456789ABCDZ"; // differs in the first 15 code units

        let bof_payload = [0x00, 0x06, 0x05, 0x00];

        // BIFF8 RC4 FILEPASS placeholder (Standard Encryption / major=1, minor=2 => 128-bit key).
        let mut filepass_placeholder = vec![0u8; 6 + 16 + 16 + 16];
        filepass_placeholder[0..2].copy_from_slice(&BIFF8_ENCRYPTION_TYPE_RC4.to_le_bytes());
        filepass_placeholder[2..4].copy_from_slice(&1u16.to_le_bytes()); // major
        filepass_placeholder[4..6].copy_from_slice(&2u16.to_le_bytes()); // minor (128-bit)

        // Chosen so record2 crosses the 1024-byte block boundary: 1000 + 80.
        let r1 = record(0x00FC, &dummy_payload(1000, 0x70));
        let r2 = record(0x00FD, &dummy_payload(80, 0x80));

        let mut plain = [
            record(records::RECORD_BOF_BIFF8, &bof_payload),
            record(records::RECORD_FILEPASS, &filepass_placeholder),
            r1,
            r2,
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let mut encrypted = plain.clone();
        encrypt_workbook_stream_for_test(&mut encrypted, &password).expect("encrypt");

        // FILEPASS is plaintext and populated by the encryption helper; patch the expected plaintext
        // stream so we can compare full bytes after decryption.
        let range = filepass_payload_range(&plain);
        plain[range.clone()].copy_from_slice(&encrypted[range.clone()]);

        for candidate in [password.as_str(), password_same_prefix.as_str(), base] {
            let mut decrypted = encrypted.clone();
            decrypt_workbook_stream(&mut decrypted, candidate).expect("decrypt");
            assert_eq!(decrypted, plain);
        }

        let mut buf = encrypted.clone();
        let err = decrypt_workbook_stream(&mut buf, wrong_password).expect_err("wrong password");
        assert_eq!(err, DecryptError::WrongPassword);
    }

    #[test]
    fn rc4_standard_password_truncation_can_split_surrogate_pairs() {
        // Truncation is defined in terms of UTF-16 code units, not Unicode scalar values.
        // This means it can split a non-BMP character's surrogate pair when the 15-code-unit limit
        // falls between the high and low surrogate.
        //
        // These two passwords differ only in the *low surrogate* of their final non-BMP emoji, so
        // they should be treated as equivalent by the legacy RC4 Standard key derivation.
        let prefix = "0123456789ABCD"; // 14 ASCII chars = 14 UTF-16 code units
        let password = format!("{prefix}🔒"); // U+1F512 => surrogate pair D83D DD12
        let password_same_truncation = format!("{prefix}😀"); // U+1F600 => surrogate pair D83D DE00
        let wrong_password = format!("1123456789ABCD🔒"); // differs within the first 15 code units

        // Confirm the full UTF-16 representations differ.
        assert_ne!(
            password.encode_utf16().collect::<Vec<_>>(),
            password_same_truncation.encode_utf16().collect::<Vec<_>>(),
            "sanity: passwords should differ before truncation"
        );

        // Confirm our UTF-16LE truncation drops the low surrogate (16th code unit).
        let pw_bytes = password_to_utf16le(&password);
        let pw_bytes_same = password_to_utf16le(&password_same_truncation);
        assert_eq!(
            &pw_bytes[..],
            &pw_bytes_same[..],
            "expected passwords to match after truncation to 15 UTF-16 code units"
        );
        assert_eq!(pw_bytes.len(), 15 * 2);
        assert_eq!(
            &pw_bytes[28..30],
            &[0x3D, 0xD8], // 0xD83D little-endian (high surrogate)
            "expected last retained code unit to be the emoji high surrogate"
        );

        let bof_payload = [0x00, 0x06, 0x05, 0x00];

        // BIFF8 RC4 FILEPASS placeholder (Standard Encryption / major=1, minor=2 => 128-bit key).
        let mut filepass_placeholder = vec![0u8; 6 + 16 + 16 + 16];
        filepass_placeholder[0..2].copy_from_slice(&BIFF8_ENCRYPTION_TYPE_RC4.to_le_bytes());
        filepass_placeholder[2..4].copy_from_slice(&1u16.to_le_bytes()); // major
        filepass_placeholder[4..6].copy_from_slice(&2u16.to_le_bytes()); // minor (128-bit)

        // Cross the 1024-byte rekey boundary to ensure we exercise block stepping.
        let r1 = record(0x00FC, &dummy_payload(1000, 0x90));
        let r2 = record(0x00FD, &dummy_payload(80, 0xA0));

        let mut plain = [
            record(records::RECORD_BOF_BIFF8, &bof_payload),
            record(records::RECORD_FILEPASS, &filepass_placeholder),
            r1,
            r2,
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let mut encrypted = plain.clone();
        encrypt_workbook_stream_for_test(&mut encrypted, &password).expect("encrypt");

        // FILEPASS is plaintext and populated by the encryption helper; patch the expected plaintext
        // stream so we can compare full bytes after decryption.
        let range = filepass_payload_range(&plain);
        plain[range.clone()].copy_from_slice(&encrypted[range.clone()]);

        for candidate in [password.as_str(), password_same_truncation.as_str()] {
            let mut decrypted = encrypted.clone();
            decrypt_workbook_stream(&mut decrypted, candidate).expect("decrypt");
            assert_eq!(decrypted, plain);
        }

        let mut buf = encrypted.clone();
        let err = decrypt_workbook_stream(&mut buf, &wrong_password).expect_err("wrong password");
        assert_eq!(err, DecryptError::WrongPassword);
    }

    #[test]
    fn xor_decrypt_tolerates_trailing_padding_bytes_after_eof() {
        let password = "pw";
        let bof_payload = [0x00, 0x06, 0x05, 0x00];

        // BIFF8 XOR FILEPASS placeholder: type + key + verifier.
        let filepass_placeholder = [0u8; 6];

        let mut plain = [
            record(records::RECORD_BOF_BIFF8, &bof_payload),
            record(records::RECORD_FILEPASS, &filepass_placeholder),
            record(RECORD_DUMMY, &dummy_payload(32, 0x10)),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let mut encrypted = plain.clone();
        encrypt_workbook_stream_for_test(&mut encrypted, password).expect("encrypt");

        // FILEPASS is plaintext and populated by the encryption helper; patch the expected plaintext.
        let range = filepass_payload_range(&plain);
        plain[range.clone()].copy_from_slice(&encrypted[range]);

        let padding = [0xAAu8, 0xBB, 0xCC];
        for len in 1..=3 {
            let mut plain_padded = plain.clone();
            plain_padded.extend_from_slice(&padding[..len]);

            let mut encrypted_padded = encrypted.clone();
            encrypted_padded.extend_from_slice(&padding[..len]);

            decrypt_workbook_stream(&mut encrypted_padded, password).expect("decrypt");
            assert_eq!(encrypted_padded, plain_padded);
        }
    }

    #[test]
    fn rc4_standard_decrypt_tolerates_trailing_padding_bytes_after_eof() {
        let password = "pw";
        let bof_payload = [0x00, 0x06, 0x05, 0x00];

        // BIFF8 RC4 FILEPASS placeholder (Standard Encryption / major=1, minor=2 => 128-bit key).
        let mut filepass_placeholder = vec![0u8; 6 + 16 + 16 + 16];
        filepass_placeholder[0..2].copy_from_slice(&BIFF8_ENCRYPTION_TYPE_RC4.to_le_bytes());
        filepass_placeholder[2..4].copy_from_slice(&1u16.to_le_bytes()); // major
        filepass_placeholder[4..6].copy_from_slice(&2u16.to_le_bytes()); // minor (128-bit)

        let mut plain = [
            record(records::RECORD_BOF_BIFF8, &bof_payload),
            record(records::RECORD_FILEPASS, &filepass_placeholder),
            record(RECORD_DUMMY, &dummy_payload(32, 0x20)),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let mut encrypted = plain.clone();
        encrypt_workbook_stream_for_test(&mut encrypted, password).expect("encrypt");

        // FILEPASS is plaintext and populated by the encryption helper; patch the expected plaintext.
        let range = filepass_payload_range(&plain);
        plain[range.clone()].copy_from_slice(&encrypted[range]);

        let padding = [0x01u8, 0x02, 0x03];
        for len in 1..=3 {
            let mut plain_padded = plain.clone();
            plain_padded.extend_from_slice(&padding[..len]);

            let mut encrypted_padded = encrypted.clone();
            encrypted_padded.extend_from_slice(&padding[..len]);

            decrypt_workbook_stream(&mut encrypted_padded, password).expect("decrypt");
            assert_eq!(encrypted_padded, plain_padded);
        }
    }
}

#[cfg(test)]
mod tests_proptest;
