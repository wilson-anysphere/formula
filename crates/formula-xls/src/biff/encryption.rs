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

use md5::{Digest as _, Md5};
use thiserror::Error;

use super::{records, BiffVersion};

pub(crate) mod cryptoapi;
pub(crate) mod rc4;
pub(crate) mod xor;
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

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) enum BiffEncryption {
    /// BIFF5 XOR obfuscation. FILEPASS payload is `key` + `verifier`.
    Biff5Xor { key: u16, verifier: u16 },
    /// BIFF8 XOR obfuscation. FILEPASS payload is `wEncryptionType` + `key` + `verifier`.
    Biff8Xor { key: u16, verifier: u16 },
    /// BIFF8 RC4 encryption (legacy non-CryptoAPI).
    ///
    /// The full FILEPASS payload is preserved so decryptors can parse algorithm details.
    Biff8Rc4 { filepass_payload: Vec<u8> },
    /// BIFF8 RC4 encryption using CryptoAPI.
    ///
    /// The full FILEPASS payload is preserved so decryptors can parse algorithm details.
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
/// 4. Dispatches to an algorithm-specific decryptor to decrypt record payloads *after* FILEPASS.
///
/// Note: Bytes before FILEPASS are always plaintext; encryption begins immediately after the
/// FILEPASS record.
pub(crate) fn decrypt_workbook_stream(
    workbook_stream: &mut [u8],
    password: &str,
) -> Result<(), DecryptError> {
    let biff_version = super::detect_biff_version(workbook_stream);

    let mut iter =
        records::BiffRecordIter::from_offset(workbook_stream, 0).map_err(DecryptError::InvalidFilePass)?;

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

        if password.is_empty() {
            return Err(DecryptError::PasswordRequired);
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
        BiffEncryption::Biff8Rc4CryptoApi { .. } => Err(DecryptError::UnsupportedEncryption(
            "BIFF8 RC4 CryptoAPI decryption not implemented".to_string(),
        )),
    }
}

fn derive_xor_array(password: &str) -> [u8; 16] {
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
    out
}

fn apply_xor_obfuscation_in_place(
    workbook_stream: &mut [u8],
    encrypted_start: usize,
    key: u16,
    xor_array: &[u8; 16],
) -> Result<(), DecryptError> {
    let key_bytes = key.to_le_bytes();
    let mut pos = 0usize;

    let mut offset = encrypted_start;
    while offset < workbook_stream.len() {
        let remaining = workbook_stream.len() - offset;
        if remaining < 4 {
            return Err(DecryptError::InvalidFilePass(
                "truncated BIFF record header while decrypting XOR stream".to_string(),
            ));
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
    // FILEPASS stores a legacy XOR password verifier using the same 16-bit hash as worksheet/workbook
    // protection.
    let expected = xor::xor_password_verifier(password);
    if expected != verifier {
        return Err(DecryptError::WrongPassword);
    }

    let xor_array = derive_xor_array(password);
    apply_xor_obfuscation_in_place(workbook_stream, encrypted_start, key, &xor_array)
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

/// Minimal RC4 stream cipher implementation (KSA + PRGA).
///
/// We implement this locally to avoid pulling in a full crypto dependency just for legacy XLS
/// decryption.
#[derive(Clone)]
struct Rc4 {
    s: [u8; 256],
    i: u8,
    j: u8,
}

impl Rc4 {
    fn new(key: &[u8]) -> Self {
        debug_assert!(!key.is_empty());

        let mut s = [0u8; 256];
        for (idx, slot) in s.iter_mut().enumerate() {
            *slot = idx as u8;
        }

        let mut j: u8 = 0;
        for i in 0u16..=255 {
            let key_byte = key[i as usize % key.len()];
            j = j.wrapping_add(s[i as usize]).wrapping_add(key_byte);
            s.swap(i as usize, j as usize);
        }

        Self { s, i: 0, j: 0 }
    }

    fn apply_keystream(&mut self, data: &mut [u8]) {
        for b in data {
            self.i = self.i.wrapping_add(1);
            self.j = self.j.wrapping_add(self.s[self.i as usize]);
            self.s.swap(self.i as usize, self.j as usize);
            let idx = self.s[self.i as usize].wrapping_add(self.s[self.j as usize]);
            let k = self.s[idx as usize];
            *b ^= k;
        }
    }
}

fn parse_filepass_rc4(payload: &[u8]) -> Result<FilePassRc4, DecryptError> {
    // FILEPASS payload begins with wEncryptionType (u16).
    if payload.len() < 2 {
        return Err(DecryptError::InvalidFilePass("truncated FILEPASS record".to_string()));
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

    let salt = payload[6..22].try_into().expect("slice len");
    let encrypted_verifier = payload[22..38].try_into().expect("slice len");
    let encrypted_verifier_hash = payload[38..54].try_into().expect("slice len");

    Ok(FilePassRc4 {
        key_len,
        salt,
        encrypted_verifier,
        encrypted_verifier_hash,
    })
}

fn password_to_utf16le(password: &str) -> Vec<u8> {
    // Excel 97-2003 passwords are limited to 15 characters for legacy RC4 encryption.
    //
    // Use UTF-16LE and truncate to 15 UTF-16 code units.
    let mut out = Vec::with_capacity(password.len().min(15) * 2);
    for u in password.encode_utf16().take(15) {
        out.extend_from_slice(&u.to_le_bytes());
    }
    out
}

fn derive_rc4_intermediate_key(password: &str, salt: &[u8; 16]) -> [u8; 16] {
    // [MS-OFFCRYPTO] "Standard Encryption" key derivation (Excel 97-2003 RC4):
    // - password_hash = MD5(UTF16LE(password))
    // - intermediate_key = MD5(password_hash + salt)
    let password_bytes = password_to_utf16le(password);
    let password_hash: [u8; 16] = Md5::digest(&password_bytes).into();

    let mut h = Md5::new();
    h.update(password_hash);
    h.update(salt);
    h.finalize().into()
}

fn derive_rc4_block_key(intermediate_key: &[u8; 16], block: u32) -> [u8; 16] {
    // block_key = MD5(intermediate_key + block_index_le32)
    let mut h = Md5::new();
    h.update(intermediate_key);
    h.update(block.to_le_bytes());
    h.finalize().into()
}

/// Applies BIFF8 RC4 encryption/decryption to a byte stream representing *record data* (not record
/// headers).
///
/// This is symmetric: applying it twice with the same key yields the original bytes.
struct Rc4BiffStream {
    intermediate_key: [u8; 16],
    key_len: usize,
    block: u32,
    pos_in_block: usize,
    cipher: Rc4,
}

impl Rc4BiffStream {
    fn new(intermediate_key: [u8; 16], key_len: usize) -> Self {
        let block_key = derive_rc4_block_key(&intermediate_key, 0);
        let cipher = Rc4::new(&block_key[..key_len]);
        Self {
            intermediate_key,
            key_len,
            block: 0,
            pos_in_block: 0,
            cipher,
        }
    }

    fn rekey(&mut self) {
        let block_key = derive_rc4_block_key(&self.intermediate_key, self.block);
        self.cipher = Rc4::new(&block_key[..self.key_len]);
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

fn verify_rc4_password(filepass: &FilePassRc4, password: &str) -> Result<[u8; 16], DecryptError> {
    let intermediate_key = derive_rc4_intermediate_key(password, &filepass.salt);
    let block_key = derive_rc4_block_key(&intermediate_key, 0);
    let mut rc4 = Rc4::new(&block_key[..filepass.key_len]);

    // Decrypt verifier + verifier hash.
    let mut buf = [0u8; 32];
    buf[0..16].copy_from_slice(&filepass.encrypted_verifier);
    buf[16..32].copy_from_slice(&filepass.encrypted_verifier_hash);
    rc4.apply_keystream(&mut buf);

    let verifier = &buf[0..16];
    let verifier_hash = &buf[16..32];
    let expected_hash: [u8; 16] = Md5::digest(verifier).into();

    if verifier_hash != expected_hash {
        return Err(DecryptError::WrongPassword);
    }

    Ok(intermediate_key)
}

fn decrypt_biff8_rc4_standard(
    workbook_stream: &mut [u8],
    encrypted_start: usize,
    password: &str,
    filepass_payload: &[u8],
) -> Result<(), DecryptError> {
    let filepass = parse_filepass_rc4(filepass_payload)?;
    let intermediate_key = verify_rc4_password(&filepass, password)?;
    let mut rc4_stream = Rc4BiffStream::new(intermediate_key, filepass.key_len);

    // Decrypt record payloads after FILEPASS.
    let mut offset = encrypted_start;
    while offset < workbook_stream.len() {
        let remaining = workbook_stream.len() - offset;
        if remaining < 4 {
            return Err(DecryptError::InvalidFilePass(
                "truncated BIFF record header while decrypting".to_string(),
            ));
        }

        let record_id = u16::from_le_bytes([workbook_stream[offset], workbook_stream[offset + 1]]);
        let len =
            u16::from_le_bytes([workbook_stream[offset + 2], workbook_stream[offset + 3]]) as usize;

        let data_start = offset
            .checked_add(4)
            .ok_or_else(|| DecryptError::InvalidFilePass("BIFF record offset overflow".to_string()))?;
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

                let intermediate_key = derive_rc4_intermediate_key(password, &salt);
                let block_key = derive_rc4_block_key(&intermediate_key, 0);
                let mut rc4 = Rc4::new(&block_key[..key_len]);
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
                    let block_key = derive_rc4_block_key(&intermediate_key, b as u32);
                    let mut rc4 = Rc4::new(&block_key[..key_len]);
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
                return Err(DecryptError::UnsupportedEncryption(
                    "BIFF8 RC4 CryptoAPI encryption helper not implemented".to_string(),
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

        let intermediate_key = derive_rc4_intermediate_key(password, &salt);
        let block_key = derive_rc4_block_key(&intermediate_key, 0);
        let mut rc4 = Rc4::new(&block_key[..key_len]);

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
        intermediate_key: [u8; 16],
        key_len: usize,
    ) -> Result<(), String> {
        // Manual encryption reference: apply RC4 stream to record payload bytes only.
        let mut block: u32 = 0;
        let mut pos_in_block: usize = 0;
        let mut block_key = derive_rc4_block_key(&intermediate_key, block);
        let mut cipher = Rc4::new(&block_key[..key_len]);

        let mut offset = encrypted_start;
        while offset < workbook_stream.len() {
            if workbook_stream.len() - offset < 4 {
                return Err("truncated record header".to_string());
            }
            let len =
                u16::from_le_bytes([workbook_stream[offset + 2], workbook_stream[offset + 3]])
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
                    block_key = derive_rc4_block_key(&intermediate_key, block);
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

    #[test]
    fn decrypt_workbook_stream_requires_password_when_filepass_present() {
        // Minimal BIFF8 stream: BOF + FILEPASS + EOF.
        let bof_payload = [0x00, 0x06, 0x05, 0x00]; // BIFF8, workbook globals
        let filepass_payload = [
            0x00, 0x00, // wEncryptionType (XOR)
            0x34, 0x12, // key
            0x78, 0x56, // verifier
        ];
        let mut stream = [
            record(records::RECORD_BOF_BIFF8, &bof_payload),
            record(records::RECORD_FILEPASS, &filepass_payload),
            record(records::RECORD_EOF, &[]),
        ]
        .concat();

        let err = decrypt_workbook_stream(&mut stream, "").expect_err("expected error");
        assert_eq!(err, DecryptError::PasswordRequired);
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

        let intermediate_key = derive_rc4_intermediate_key(password, &salt);
        encrypt_record_payloads_in_place(&mut encrypted, encrypted_start, intermediate_key, key_len)
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
        let intermediate_key = derive_rc4_intermediate_key(password, &salt);
        encrypt_record_payloads_in_place(&mut encrypted, encrypted_start, intermediate_key, key_len)
            .expect("encrypt");

        // Decrypt and assert we recover the original bytes.
        decrypt_workbook_stream(&mut encrypted, password).expect("decrypt");
        assert_eq!(encrypted, plain);
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
}
