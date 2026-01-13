use base64::engine::general_purpose::{
    STANDARD as BASE64_STANDARD, STANDARD_NO_PAD as BASE64_STANDARD_NO_PAD,
};
use base64::Engine;
use hmac::{Hmac, Mac};
use md5::Md5;
use quick_xml::events::Event;
use quick_xml::Reader;
use rand::rngs::OsRng;
use rand::RngCore;
use sha1::Sha1;
use sha2::{Sha256, Sha384, Sha512};

use crate::crypto::{
    aes_cbc_decrypt, aes_cbc_encrypt, derive_agile_key, derive_iv, password_to_utf16le,
    HashAlgorithm,
};
use crate::error::OfficeCryptoError;
use crate::util::{
    checked_vec_len, ct_eq, parse_encrypted_package_original_size, EncryptionInfoHeader,
};
use zeroize::Zeroizing;
use zeroize::Zeroize;

const BLOCK_KEY_VERIFIER_HASH_INPUT: &[u8; 8] = b"\xFE\xA7\xD2\x76\x3B\x4B\x9E\x79";
const BLOCK_KEY_VERIFIER_HASH_VALUE: &[u8; 8] = b"\xD7\xAA\x0F\x6D\x30\x61\x34\x4E";
const BLOCK_KEY_ENCRYPTED_KEY_VALUE: &[u8; 8] = b"\x14\x6E\x0B\xE7\xAB\xAC\xD0\xD6";
const BLOCK_KEY_INTEGRITY_HMAC_KEY: &[u8; 8] = b"\x5F\xB2\xAD\x01\x0C\xB9\xE1\xF6";
const BLOCK_KEY_INTEGRITY_HMAC_VALUE: &[u8; 8] = b"\xA0\x67\x7F\x02\xB2\x2C\x84\x33";

#[derive(Debug, Clone, Copy)]
enum PasswordKeyEncryptorIvScheme {
    /// Use `saltValue[..blockSize]` as the AES-CBC IV for verifier/key blobs.
    SaltValue,
    /// Derive the IV as `TruncateHash(Hash(saltValue || blockKey), blockSize)`.
    DerivedFromBlockKey,
}

impl PasswordKeyEncryptorIvScheme {
    fn iv_for_block_key(
        self,
        hash_alg: HashAlgorithm,
        salt: &[u8],
        block_key: &[u8],
        block_size: usize,
    ) -> Result<Vec<u8>, OfficeCryptoError> {
        match self {
            PasswordKeyEncryptorIvScheme::SaltValue => salt
                .get(..block_size)
                .ok_or_else(|| {
                    OfficeCryptoError::InvalidFormat(format!(
                        "password saltValue shorter than blockSize ({block_size})"
                    ))
                })
                .map(<[u8]>::to_vec),
            PasswordKeyEncryptorIvScheme::DerivedFromBlockKey => {
                Ok(derive_iv(hash_alg, salt, block_key, block_size))
            }
        }
    }
}

#[derive(Debug, Clone)]
pub(crate) struct AgileEncryptionInfo {
    #[allow(dead_code)]
    pub(crate) version_major: u16,
    #[allow(dead_code)]
    pub(crate) version_minor: u16,
    #[allow(dead_code)]
    pub(crate) flags: u32,
    pub(crate) key_data: AgileKeyData,
    pub(crate) data_integrity: Option<AgileDataIntegrity>,
    pub(crate) password_key_encryptor: AgilePasswordKeyEncryptor,
}

#[derive(Debug, Clone)]
pub(crate) struct AgileKeyData {
    pub(crate) salt: Vec<u8>,
    pub(crate) block_size: usize,
    pub(crate) key_bits: usize,
    pub(crate) hash_algorithm: HashAlgorithm,
    /// Hash output size in bytes (`hashSize` attribute).
    ///
    /// For standard producers this matches `hash_algorithm.digest_len()`, but we keep it explicit
    /// to correctly ignore AES block padding when comparing verifier/HMAC digests.
    pub(crate) hash_size: usize,
    pub(crate) cipher_algorithm: String,
    pub(crate) cipher_chaining: String,
}

#[derive(Debug, Clone)]
pub(crate) struct AgileDataIntegrity {
    #[allow(dead_code)]
    pub(crate) encrypted_hmac_key: Vec<u8>,
    #[allow(dead_code)]
    pub(crate) encrypted_hmac_value: Vec<u8>,
}

#[derive(Debug, Clone)]
pub(crate) struct AgilePasswordKeyEncryptor {
    pub(crate) salt: Vec<u8>,
    pub(crate) block_size: usize,
    pub(crate) key_bits: usize,
    pub(crate) spin_count: u32,
    pub(crate) hash_algorithm: HashAlgorithm,
    /// Hash output size in bytes (`hashSize` attribute).
    pub(crate) hash_size: usize,
    pub(crate) cipher_algorithm: String,
    pub(crate) cipher_chaining: String,
    pub(crate) encrypted_verifier_hash_input: Vec<u8>,
    pub(crate) encrypted_verifier_hash_value: Vec<u8>,
    pub(crate) encrypted_key_value: Vec<u8>,
}

pub(crate) fn parse_agile_encryption_info(
    bytes: &[u8],
    header: &EncryptionInfoHeader,
) -> Result<AgileEncryptionInfo, OfficeCryptoError> {
    // Real-world Agile `EncryptionInfo` streams vary in how they wrap/encode the XML descriptor:
    // - Some include a 4-byte XML length prefix after the 8-byte version header.
    // - Others start the XML directly after the version header (no length prefix).
    // - The XML may be UTF-8 (optionally with BOM) or UTF-16LE, and can be padded with trailing
    //   NUL bytes.
    //
    // We mirror the robustness of `formula_io::extract_agile_encryption_info_xml` by extracting
    // candidates from the post-version payload and attempting to parse the descriptor across those
    // variants.
    // Use the already-validated `EncryptionInfoHeader` size/offset to bound how many bytes we
    // consider part of the XML descriptor. This prevents crafted inputs from appending arbitrarily
    // large trailing data to the stream and forcing the XML parser to process unbounded input.
    let end = header
        .header_offset
        .checked_add(header.header_size as usize)
        .ok_or_else(|| OfficeCryptoError::InvalidFormat("EncryptionInfo XML size overflow".to_string()))?;
    let bounded = bytes.get(..end).ok_or_else(|| {
        OfficeCryptoError::InvalidFormat("EncryptionInfo XML size out of range".to_string())
    })?;

    let descriptor = parse_agile_descriptor_from_stream(bounded)?;

    Ok(AgileEncryptionInfo {
        version_major: header.version_major,
        version_minor: header.version_minor,
        flags: header.flags,
        key_data: descriptor.key_data,
        data_integrity: descriptor.data_integrity,
        password_key_encryptor: descriptor.password_key_encryptor,
    })
}

fn parse_agile_descriptor_from_stream(bytes: &[u8]) -> Result<AgileDescriptor, OfficeCryptoError> {
    if bytes.len() < 8 {
        return Err(OfficeCryptoError::InvalidFormat(
            "EncryptionInfo stream too short".to_string(),
        ));
    }

    let payload = &bytes[8..];
    let mut errors: Vec<String> = Vec::new();

    // --- Primary: UTF-8 payload (trim UTF-8 BOM, trim trailing NULs). ---
    match parse_agile_descriptor_utf8(payload) {
        Ok(d) => return Ok(d),
        Err(err) => errors.push(format!("utf-8: {err}")),
    }

    // --- Fallback: UTF-16LE when there are many NUL bytes (ASCII UTF-16 pattern). ---
    if is_nul_heavy(payload) {
        match parse_agile_descriptor_utf16le(payload) {
            Ok(d) => return Ok(d),
            Err(err) => errors.push(format!("utf-16le: {err}")),
        }
    }

    // --- Fallback: length-prefix heuristic (u32 LE) ---
    if let Some(len_slice) = length_prefixed_slice(payload) {
        match parse_agile_descriptor_utf8(len_slice) {
            Ok(d) => return Ok(d),
            Err(err) => errors.push(format!("len+utf-8: {err}")),
        }
        if is_nul_heavy(len_slice) {
            match parse_agile_descriptor_utf16le(len_slice) {
                Ok(d) => return Ok(d),
                Err(err) => errors.push(format!("len+utf-16le: {err}")),
            }
        }
    }

    // --- Fallback: scan forward to the first `<` when the payload contains `<encryption` later. ---
    if let Some(scanned) = scan_to_first_xml_tag(payload) {
        match parse_agile_descriptor_utf8(scanned) {
            Ok(d) => return Ok(d),
            Err(err) => errors.push(format!("scan+utf-8: {err}")),
        }
        if is_nul_heavy(scanned) {
            match parse_agile_descriptor_utf16le(scanned) {
                Ok(d) => return Ok(d),
                Err(err) => errors.push(format!("scan+utf-16le: {err}")),
            }
        }
    }

    if errors.is_empty() {
        errors.push("no candidates".to_string());
    }
    Err(OfficeCryptoError::InvalidFormat(format!(
        "failed to extract Agile EncryptionInfo XML: {}",
        errors.join("; ")
    )))
}

fn trim_trailing_nul_bytes(mut bytes: &[u8]) -> &[u8] {
    while let Some((&last, rest)) = bytes.split_last() {
        if last == 0 {
            bytes = rest;
        } else {
            break;
        }
    }
    bytes
}

fn trim_trailing_utf16le_nul_units(mut bytes: &[u8]) -> &[u8] {
    while bytes.len() >= 2 {
        let n = bytes.len();
        if bytes[n - 2] == 0 && bytes[n - 1] == 0 {
            bytes = &bytes[..n - 2];
        } else {
            break;
        }
    }
    bytes
}

fn trim_utf8_bom(bytes: &[u8]) -> &[u8] {
    bytes.strip_prefix(&[0xEF, 0xBB, 0xBF]).unwrap_or(bytes)
}

fn trim_start_ascii_whitespace(bytes: &[u8]) -> &[u8] {
    let mut idx = 0usize;
    while idx < bytes.len() {
        if matches!(bytes[idx], b' ' | b'\t' | b'\r' | b'\n') {
            idx += 1;
        } else {
            break;
        }
    }
    &bytes[idx..]
}

fn is_nul_heavy(bytes: &[u8]) -> bool {
    if bytes.is_empty() {
        return false;
    }
    let zeros = bytes.iter().filter(|&&b| b == 0).count();
    zeros > bytes.len() / 8
}

fn parse_agile_descriptor_utf8(bytes: &[u8]) -> Result<AgileDescriptor, String> {
    let bytes = trim_trailing_nul_bytes(bytes);
    let bytes = trim_utf8_bom(bytes);
    let xml = std::str::from_utf8(bytes).map_err(|e| e.to_string())?;
    // In case the stream was decoded through a path that preserved U+FEFF.
    let xml = xml.strip_prefix('\u{FEFF}').unwrap_or(xml);
    parse_agile_descriptor(xml).map_err(|e| e.to_string())
}

fn parse_agile_descriptor_utf16le(bytes: &[u8]) -> Result<AgileDescriptor, String> {
    let mut bytes = trim_trailing_utf16le_nul_units(bytes);
    if bytes.starts_with(&[0xFF, 0xFE]) {
        bytes = &bytes[2..];
    }

    // UTF-16 requires an even number of bytes; ignore a trailing odd byte.
    bytes = &bytes[..bytes.len().saturating_sub(bytes.len() % 2)];

    let mut code_units: Vec<u16> = Vec::with_capacity(bytes.len() / 2);
    for pair in bytes.chunks_exact(2) {
        code_units.push(u16::from_le_bytes([pair[0], pair[1]]));
    }
    let mut xml = String::from_utf16(&code_units).map_err(|_| "invalid UTF-16LE".to_string())?;
    if let Some(stripped) = xml.strip_prefix('\u{FEFF}') {
        xml = stripped.to_string();
    }
    while xml.ends_with('\0') {
        xml.pop();
    }
    parse_agile_descriptor(&xml).map_err(|e| e.to_string())
}

fn length_prefixed_slice(payload: &[u8]) -> Option<&[u8]> {
    let len_bytes: [u8; 4] = payload.get(0..4)?.try_into().ok()?;
    let len = u32::from_le_bytes(len_bytes) as usize;
    if len == 0 || len > payload.len().saturating_sub(4) {
        return None;
    }
    let candidate = payload.get(4..4 + len)?;

    // Ensure the candidate *looks* like XML to avoid false positives on arbitrary data.
    let candidate_trimmed = trim_start_ascii_whitespace(candidate);
    let candidate_trimmed = trim_utf8_bom(candidate_trimmed);

    if candidate_trimmed.first() == Some(&b'<') {
        return Some(candidate);
    }
    // UTF-16LE BOM.
    if candidate_trimmed.starts_with(&[0xFF, 0xFE]) {
        return Some(candidate);
    }

    None
}

fn scan_to_first_xml_tag(payload: &[u8]) -> Option<&[u8]> {
    // Be conservative: only scan if we see the expected root tag bytes somewhere later.
    const NEEDLE: &[u8] = b"<encryption";
    if !payload
        .windows(NEEDLE.len())
        .any(|w| w.eq_ignore_ascii_case(NEEDLE))
    {
        return None;
    }

    let payload = trim_utf8_bom(payload);
    let trimmed = trim_start_ascii_whitespace(payload);
    if trimmed.first() == Some(&b'<') {
        return None;
    }

    let idx = payload.iter().position(|&b| b == b'<')?;
    Some(&payload[idx..])
}

pub(crate) fn decrypt_agile_encrypted_package(
    info: &AgileEncryptionInfo,
    encrypted_package: &[u8],
    password: &str,
    opts: &crate::DecryptOptions,
) -> Result<Vec<u8>, OfficeCryptoError> {
    let total_size = parse_encrypted_package_original_size(encrypted_package)?;
    if total_size > crate::MAX_ENCRYPTED_PACKAGE_ORIGINAL_SIZE {
        return Err(OfficeCryptoError::SizeLimitExceededU64 {
            context: "EncryptedPackage.originalSize",
            limit: crate::MAX_ENCRYPTED_PACKAGE_ORIGINAL_SIZE,
        });
    }
    let expected_len = checked_vec_len(total_size)?;
    let ciphertext = &encrypted_package[8..];

    // `EncryptedPackage` begins with an 8-byte (little-endian) decrypted length. This value is
    // untrusted and must not be used to drive large allocations without plausibility checks.
    //
    // Special-case: allow an empty package only when both the declared length and ciphertext are
    // empty.
    if expected_len == 0 {
        if !ciphertext.is_empty() {
            return Err(OfficeCryptoError::InvalidFormat(
                "EncryptedPackage size is zero but ciphertext is non-empty".to_string(),
            ));
        }
    } else if ciphertext.is_empty() {
        return Err(OfficeCryptoError::InvalidFormat(
            "EncryptedPackage ciphertext missing".to_string(),
        ));
    }
    // Conservative bound: AES-CBC encryption cannot produce fewer bytes than the original
    // plaintext length (padding can only increase the size).
    if expected_len > ciphertext.len() {
        return Err(OfficeCryptoError::InvalidFormat(format!(
            "EncryptedPackage size {} larger than ciphertext length {}",
            expected_len,
            ciphertext.len()
        )));
    }

    if info.key_data.cipher_algorithm != "AES" {
        return Err(OfficeCryptoError::UnsupportedEncryption(format!(
            "unsupported cipherAlgorithm {}",
            info.key_data.cipher_algorithm
        )));
    }
    if info.key_data.cipher_chaining != "ChainingModeCBC" {
        return Err(OfficeCryptoError::UnsupportedEncryption(format!(
            "unsupported cipherChaining {}",
            info.key_data.cipher_chaining
        )));
    }

    if info.password_key_encryptor.cipher_algorithm != "AES" {
        return Err(OfficeCryptoError::UnsupportedEncryption(format!(
            "unsupported password cipherAlgorithm {}",
            info.password_key_encryptor.cipher_algorithm
        )));
    }
    if info.password_key_encryptor.cipher_chaining != "ChainingModeCBC" {
        return Err(OfficeCryptoError::UnsupportedEncryption(format!(
            "unsupported password cipherChaining {}",
            info.password_key_encryptor.cipher_chaining
        )));
    }

    let pw_utf16 = password_to_utf16le(password);

    // Agile password key-encryptor IV handling:
    //
    // Excel-compatible implementations use `saltValue[..blockSize]` as the AES-CBC IV for the
    // password-key-encryptor blobs (`encryptedVerifierHashInput`, `encryptedVerifierHashValue`,
    // `encryptedKeyValue`). Some non-Excel producers appear to derive per-blob IVs as
    // `TruncateHash(Hash(saltValue || blockKey), blockSize)` (similar to other Agile IV derivations).
    //
    // Real-world fixture sets contain both variants, so we try both schemes.
    let password_block_size = info.password_key_encryptor.block_size;
    if password_block_size != 16 {
        return Err(OfficeCryptoError::UnsupportedEncryption(format!(
            "unsupported password blockSize {password_block_size}"
        )));
    }
    if info.key_data.block_size != 16 {
        return Err(OfficeCryptoError::UnsupportedEncryption(format!(
            "unsupported keyData blockSize {}",
            info.key_data.block_size
        )));
    }

    // `spinCount` is attacker-controlled; enforce limits up front to avoid CPU DoS.
    if info.password_key_encryptor.spin_count > opts.max_spin_count {
        return Err(OfficeCryptoError::SpinCountTooLarge {
            spin_count: info.password_key_encryptor.spin_count,
            max: opts.max_spin_count,
        });
    }

    // Ciphertext is stored in 4096-byte segments, with each segment padded to the AES block size.
    // Some producers may include trailing bytes in the OLE stream beyond the padded plaintext
    // length; ignore them by decrypting only what we need.
    const SEGMENT_LEN: usize = 4096;
    let required_ciphertext_len = if expected_len == 0 {
        0usize
    } else {
        let full_segments_len = (expected_len / SEGMENT_LEN) * SEGMENT_LEN;
        let rem = expected_len % SEGMENT_LEN;
        let last_padded = if rem == 0 {
            0usize
        } else {
            rem.checked_add(15)
                .ok_or_else(|| OfficeCryptoError::InvalidFormat("EncryptedPackage expected length overflow".to_string()))?
                / 16
                * 16
        };
        full_segments_len
            .checked_add(last_padded)
            .ok_or_else(|| OfficeCryptoError::InvalidFormat("EncryptedPackage expected length overflow".to_string()))?
    };
    if ciphertext.len() < required_ciphertext_len {
        return Err(OfficeCryptoError::InvalidFormat(format!(
            "EncryptedPackage ciphertext truncated (len {}, expected at least {})",
            ciphertext.len(),
            required_ciphertext_len
        )));
    }
    let ciphertext_decrypt = &ciphertext[..required_ciphertext_len];

    let schemes = [
        PasswordKeyEncryptorIvScheme::SaltValue,
        PasswordKeyEncryptorIvScheme::DerivedFromBlockKey,
    ];

    let mut last_err: Option<OfficeCryptoError> = None;
    for scheme in schemes {
        // Password verification.
        let verifier_input_key = derive_agile_key(
            info.password_key_encryptor.hash_algorithm,
            &info.password_key_encryptor.salt,
            &pw_utf16,
            info.password_key_encryptor.spin_count,
            info.password_key_encryptor.key_bits / 8,
            BLOCK_KEY_VERIFIER_HASH_INPUT,
        );
        let iv_vhi = scheme.iv_for_block_key(
            info.password_key_encryptor.hash_algorithm,
            &info.password_key_encryptor.salt,
            BLOCK_KEY_VERIFIER_HASH_INPUT,
            password_block_size,
        )?;
        let verifier_hash_input_plain: Zeroizing<Vec<u8>> = Zeroizing::new(aes_cbc_decrypt(
            &verifier_input_key,
            &iv_vhi,
            &info.password_key_encryptor.encrypted_verifier_hash_input,
        )?);
        let verifier_hash_input_slice = verifier_hash_input_plain
            .get(..password_block_size)
            .ok_or_else(|| {
                OfficeCryptoError::InvalidFormat(
                    "decrypted verifierHashInput shorter than 16 bytes".to_string(),
                )
            })?;

        let verifier_hash_full: Zeroizing<Vec<u8>> = Zeroizing::new(
            info.password_key_encryptor
                .hash_algorithm
                .digest(verifier_hash_input_slice),
        );
        let verifier_hash = verifier_hash_full
            .get(..info.password_key_encryptor.hash_size)
            .ok_or_else(|| {
                OfficeCryptoError::InvalidFormat(
                    "hash output shorter than encryptedKey hashSize".to_string(),
                )
            })?;

        let verifier_value_key = derive_agile_key(
            info.password_key_encryptor.hash_algorithm,
            &info.password_key_encryptor.salt,
            &pw_utf16,
            info.password_key_encryptor.spin_count,
            info.password_key_encryptor.key_bits / 8,
            BLOCK_KEY_VERIFIER_HASH_VALUE,
        );
        let iv_vhv = scheme.iv_for_block_key(
            info.password_key_encryptor.hash_algorithm,
            &info.password_key_encryptor.salt,
            BLOCK_KEY_VERIFIER_HASH_VALUE,
            password_block_size,
        )?;
        let verifier_hash_value_plain: Zeroizing<Vec<u8>> = Zeroizing::new(aes_cbc_decrypt(
            &verifier_value_key,
            &iv_vhv,
            &info.password_key_encryptor.encrypted_verifier_hash_value,
        )?);
        let expected_hash_slice = verifier_hash_value_plain
            .get(..info.password_key_encryptor.hash_size)
            .ok_or_else(|| {
                OfficeCryptoError::InvalidFormat(
                    "decrypted verifierHashValue shorter than encryptedKey hashSize".to_string(),
                )
            })?;

        if !ct_eq(expected_hash_slice, verifier_hash) {
            // Don't overwrite a prior integrity failure with InvalidPassword; if we successfully
            // validated the password for one IV scheme but can't decrypt/verify, prefer returning
            // an integrity/scheme error.
            if last_err.is_none() {
                last_err = Some(OfficeCryptoError::InvalidPassword);
            }
            continue;
        }

        // Decrypt the package key.
        let key_value_key = derive_agile_key(
            info.password_key_encryptor.hash_algorithm,
            &info.password_key_encryptor.salt,
            &pw_utf16,
            info.password_key_encryptor.spin_count,
            info.password_key_encryptor.key_bits / 8,
            BLOCK_KEY_ENCRYPTED_KEY_VALUE,
        );
        let iv_kv = scheme.iv_for_block_key(
            info.password_key_encryptor.hash_algorithm,
            &info.password_key_encryptor.salt,
            BLOCK_KEY_ENCRYPTED_KEY_VALUE,
            password_block_size,
        )?;
        let key_value_plain: Zeroizing<Vec<u8>> = Zeroizing::new(aes_cbc_decrypt(
            &key_value_key,
            &iv_kv,
            &info.password_key_encryptor.encrypted_key_value,
        )?);
        let key_len = info.key_data.key_bits / 8;
        let package_key_bytes = key_value_plain.get(..key_len).ok_or_else(|| {
            OfficeCryptoError::InvalidFormat("decrypted keyValue shorter than keyBytes".to_string())
        })?;
        let package_key: Zeroizing<Vec<u8>> = Zeroizing::new(package_key_bytes.to_vec());

        // Decrypt the HMAC key/value (dataIntegrity), if present.
        //
        // The HMAC key/value are encrypted using the package key, with IVs derived from the keyData
        // salt and fixed block keys.
        //
        // Some real-world producers omit the `<dataIntegrity>` element. In that case we can still
        // decrypt the package but cannot validate integrity.
        let mut hash_size: usize = 0;
        let mut hmac_key_len: usize = 0;
        let (hmac_key_plain, expected_hmac) = if let Some(data_integrity) = &info.data_integrity {
            hash_size = info.key_data.hash_size;
            if hash_size == 0 {
                return Err(OfficeCryptoError::InvalidFormat(
                    "keyData hashSize must be non-zero".to_string(),
                ));
            }
            let digest_len = info.key_data.hash_algorithm.digest_len();
            if hash_size > digest_len {
                return Err(OfficeCryptoError::InvalidFormat(format!(
                    "keyData hashSize {hash_size} exceeds {} digest length {digest_len}",
                    info.key_data.hash_algorithm.as_ooxml_name()
                )));
            }

            let iv_hmac_key = derive_iv(
                info.key_data.hash_algorithm,
                &info.key_data.salt,
                BLOCK_KEY_INTEGRITY_HMAC_KEY,
                info.key_data.block_size,
            );
            let hmac_key_plain: Zeroizing<Vec<u8>> = Zeroizing::new(aes_cbc_decrypt(
                &package_key,
                &iv_hmac_key,
                &data_integrity.encrypted_hmac_key,
            )?);
            hmac_key_len = std::cmp::min(hash_size, hmac_key_plain.len());
            if hmac_key_len == 0 {
                return Err(OfficeCryptoError::InvalidFormat(
                    "decrypted encryptedHmacKey is empty".to_string(),
                ));
            }

            let iv_hmac_val = derive_iv(
                info.key_data.hash_algorithm,
                &info.key_data.salt,
                BLOCK_KEY_INTEGRITY_HMAC_VALUE,
                info.key_data.block_size,
            );
            let hmac_value_plain: Zeroizing<Vec<u8>> = Zeroizing::new(aes_cbc_decrypt(
                &package_key,
                &iv_hmac_val,
                &data_integrity.encrypted_hmac_value,
            )?);
            let expected_hmac = hmac_value_plain.get(..hash_size).ok_or_else(|| {
                OfficeCryptoError::InvalidFormat(
                    "decrypted encryptedHmacValue shorter than hash output".to_string(),
                )
            })?;
            (Some(hmac_key_plain), Some(expected_hmac.to_vec()))
        } else {
            (None, None)
        };

        // Decrypt the package data in 4096-byte segments. We only decrypt the ciphertext needed for
        // the declared `originalSize` to avoid attacker-controlled ciphertext lengths forcing huge
        // allocations.
        let mut out = Vec::new();
        out.try_reserve_exact(expected_len).map_err(|source| {
            OfficeCryptoError::EncryptedPackageAllocationFailed { total_size, source }
        })?;

        let mut remaining = expected_len;
        let mut offset = 0usize;
        let mut block_index = 0u32;
        while remaining > 0 {
            let plain_len = remaining.min(SEGMENT_LEN);
            let cipher_len = padded_aes_len(plain_len)?;
            let end = offset.checked_add(cipher_len).ok_or_else(|| {
                OfficeCryptoError::InvalidFormat(
                    "EncryptedPackage ciphertext offset overflow".to_string(),
                )
            })?;
            let seg = ciphertext_decrypt.get(offset..end).ok_or_else(|| {
                OfficeCryptoError::InvalidFormat(
                    "EncryptedPackage ciphertext shorter than declared originalSize".to_string(),
                )
            })?;
            let iv = derive_iv(
                info.key_data.hash_algorithm,
                &info.key_data.salt,
                &block_index.to_le_bytes(),
                info.key_data.block_size,
            );
            let mut plain = aes_cbc_decrypt(&package_key, &iv, seg)?;
            if plain_len > plain.len() {
                return Err(OfficeCryptoError::InvalidFormat(
                    "decrypted segment shorter than expected".to_string(),
                ));
            }
            out.extend_from_slice(&plain[..plain_len]);
            plain.zeroize();
            offset = end;
            remaining -= plain_len;
            block_index = block_index.checked_add(1).ok_or_else(|| {
                OfficeCryptoError::InvalidFormat("segment counter overflow".to_string())
            })?;
        }

        // Validate data integrity (HMAC) when `<dataIntegrity>` is present.
        //
        // MS-OFFCRYPTO describes `dataIntegrity` as an HMAC over the **EncryptedPackage stream bytes**
        // (length prefix + ciphertext). This matches Excel and the `ms-offcrypto-writer` crate.
        //
        // However, in the wild there are at least three compatibility variants:
        // - HMAC over the ciphertext bytes only (excluding the 8-byte size prefix)
        // - HMAC over the decrypted package bytes (plaintext ZIP)
        // - HMAC over header + plaintext (8-byte size prefix + plaintext ZIP bytes)
        //
        // For robustness, accept any of these targets, preferring the spec'd EncryptedPackage stream.
        if let (Some(hmac_key_plain), Some(expected_hmac)) = (&hmac_key_plain, &expected_hmac) {
            let hmac_key_plain = &hmac_key_plain[..hmac_key_len];

            let computed_hmac_stream_full =
                compute_hmac(info.key_data.hash_algorithm, hmac_key_plain, encrypted_package)?;
            let computed_hmac_stream =
                computed_hmac_stream_full.get(..hash_size).ok_or_else(|| {
                    OfficeCryptoError::InvalidFormat(
                        "HMAC output shorter than hashSize".to_string(),
                    )
                })?;

            let computed_hmac_ciphertext_full =
                compute_hmac(info.key_data.hash_algorithm, hmac_key_plain, ciphertext)?;
            let computed_hmac_ciphertext =
                computed_hmac_ciphertext_full.get(..hash_size).ok_or_else(|| {
                    OfficeCryptoError::InvalidFormat(
                        "HMAC output shorter than hashSize".to_string(),
                    )
                })?;

            let mut integrity_ok = ct_eq(expected_hmac, computed_hmac_stream)
                || ct_eq(expected_hmac, computed_hmac_ciphertext);
            if !integrity_ok {
                let computed_hmac_plaintext_full =
                    compute_hmac(info.key_data.hash_algorithm, hmac_key_plain, &out)?;
                let computed_hmac_plaintext =
                    computed_hmac_plaintext_full.get(..hash_size).ok_or_else(|| {
                        OfficeCryptoError::InvalidFormat(
                            "HMAC output shorter than hashSize".to_string(),
                        )
                    })?;
                let computed_hmac_plaintext_with_size_full = compute_hmac_two(
                    info.key_data.hash_algorithm,
                    hmac_key_plain,
                    encrypted_package.get(..8).ok_or_else(|| {
                        OfficeCryptoError::InvalidFormat(
                            "EncryptedPackage stream too short for size prefix".to_string(),
                        )
                    })?,
                    &out,
                )?;
                let computed_hmac_plaintext_with_size =
                    computed_hmac_plaintext_with_size_full.get(..hash_size).ok_or_else(|| {
                        OfficeCryptoError::InvalidFormat(
                            "HMAC output shorter than hashSize".to_string(),
                        )
                    })?;
                integrity_ok = ct_eq(expected_hmac, computed_hmac_plaintext)
                    || ct_eq(expected_hmac, computed_hmac_plaintext_with_size);
            }
            if !integrity_ok {
                last_err = Some(OfficeCryptoError::IntegrityCheckFailed);
                continue;
            }
        }

        return Ok(out);
    }

    Err(last_err.unwrap_or(OfficeCryptoError::InvalidPassword))
}

pub(crate) fn encrypt_agile_encrypted_package(
    zip_bytes: &[u8],
    password: &str,
    opts: &crate::EncryptOptions,
) -> Result<(Vec<u8>, Vec<u8>), OfficeCryptoError> {
    if opts.key_bits % 8 != 0 {
        return Err(OfficeCryptoError::InvalidOptions(
            "key_bits must be divisible by 8".to_string(),
        ));
    }
    if opts.key_bits != 128 && opts.key_bits != 256 {
        return Err(OfficeCryptoError::InvalidOptions(format!(
            "unsupported key_bits {} (expected 128 or 256)",
            opts.key_bits
        )));
    }

    let key_bytes = opts.key_bits / 8;
    let block_size = 16usize;
    let hash_alg = opts.hash_algorithm;

    let pw_utf16 = password_to_utf16le(password);

    // Random salts and keys.
    let mut salt_key_encryptor = vec![0u8; 16];
    let mut salt_key_data = vec![0u8; 16];
    OsRng.fill_bytes(&mut salt_key_encryptor);
    OsRng.fill_bytes(&mut salt_key_data);

    let mut package_key_plain = vec![0u8; key_bytes];
    OsRng.fill_bytes(&mut package_key_plain);
    let package_key_plain: Zeroizing<Vec<u8>> = Zeroizing::new(package_key_plain);

    let mut verifier_hash_input_plain = [0u8; 16];
    OsRng.fill_bytes(&mut verifier_hash_input_plain);
    let verifier_hash_value_plain = hash_alg.digest(&verifier_hash_input_plain);
    let verifier_hash_value_plain = pad_zero(&verifier_hash_value_plain, block_size);

    // See `decrypt_agile_encrypted_package`: password-key-encryptor fields use `saltValue`
    // as the IV (truncated to blockSize).
    let verifier_iv = salt_key_encryptor.get(..block_size).ok_or_else(|| {
        OfficeCryptoError::InvalidFormat("saltValue shorter than blockSize".to_string())
    })?;

    // Encrypt verifierHashInput.
    let key_vhi = derive_agile_key(
        hash_alg,
        &salt_key_encryptor,
        &pw_utf16,
        opts.spin_count,
        key_bytes,
        BLOCK_KEY_VERIFIER_HASH_INPUT,
    );
    let enc_vhi = aes_cbc_encrypt(&key_vhi, verifier_iv, &verifier_hash_input_plain)?;

    // Encrypt verifierHashValue.
    let key_vhv = derive_agile_key(
        hash_alg,
        &salt_key_encryptor,
        &pw_utf16,
        opts.spin_count,
        key_bytes,
        BLOCK_KEY_VERIFIER_HASH_VALUE,
    );
    let enc_vhv = aes_cbc_encrypt(&key_vhv, verifier_iv, &verifier_hash_value_plain)?;

    // Encrypt package key (encryptedKeyValue).
    let key_kv = derive_agile_key(
        hash_alg,
        &salt_key_encryptor,
        &pw_utf16,
        opts.spin_count,
        key_bytes,
        BLOCK_KEY_ENCRYPTED_KEY_VALUE,
    );
    let enc_kv = aes_cbc_encrypt(&key_kv, verifier_iv, &package_key_plain)?;

    // Encrypt package bytes.
    let encrypted_package = encrypt_encrypted_package_stream(
        zip_bytes,
        &package_key_plain,
        hash_alg,
        &salt_key_data,
        block_size,
    )?;

    // Integrity (HMAC over the EncryptedPackage stream).
    let mut hmac_key_plain = vec![0u8; hash_alg.digest_len()];
    OsRng.fill_bytes(&mut hmac_key_plain);
    let hmac_key_plain: Zeroizing<Vec<u8>> = Zeroizing::new(hmac_key_plain);
    let hmac_value_plain = compute_hmac(hash_alg, &hmac_key_plain, &encrypted_package)?;
    let hmac_value_plain = pad_zero(&hmac_value_plain, block_size);

    let iv_hmac_key = derive_iv(
        hash_alg,
        &salt_key_data,
        BLOCK_KEY_INTEGRITY_HMAC_KEY,
        block_size,
    );
    let encrypted_hmac_key = aes_cbc_encrypt(
        &package_key_plain,
        &iv_hmac_key,
        &pad_zero(&hmac_key_plain, block_size),
    )?;
    let iv_hmac_val = derive_iv(
        hash_alg,
        &salt_key_data,
        BLOCK_KEY_INTEGRITY_HMAC_VALUE,
        block_size,
    );
    let encrypted_hmac_value =
        aes_cbc_encrypt(&package_key_plain, &iv_hmac_val, &hmac_value_plain)?;

    // Build EncryptionInfo XML.
    let b64 = base64::engine::general_purpose::STANDARD;
    let salt_key_encryptor_b64 = b64.encode(&salt_key_encryptor);
    let salt_key_data_b64 = b64.encode(&salt_key_data);
    let enc_vhi_b64 = b64.encode(enc_vhi);
    let enc_vhv_b64 = b64.encode(enc_vhv);
    let enc_kv_b64 = b64.encode(enc_kv);
    let enc_hmac_key_b64 = b64.encode(encrypted_hmac_key);
    let enc_hmac_value_b64 = b64.encode(encrypted_hmac_value);
    let hash_size = hash_alg.digest_len();

    let xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
            xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
  <keyData saltSize="16" blockSize="16" keyBits="{key_bits}" hashSize="{hash_size}" hashAlgorithm="{hash_alg_name}"
           cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" saltValue="{salt_key_data_b64}"/>
  <dataIntegrity encryptedHmacKey="{enc_hmac_key_b64}" encryptedHmacValue="{enc_hmac_value_b64}"/>
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey saltSize="16" blockSize="16" keyBits="{key_bits}" hashSize="{hash_size}" spinCount="{spin_count}"
                      hashAlgorithm="{hash_alg_name}" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                      saltValue="{salt_key_encryptor_b64}"
                      encryptedVerifierHashInput="{enc_vhi_b64}"
                      encryptedVerifierHashValue="{enc_vhv_b64}"
                      encryptedKeyValue="{enc_kv_b64}"/>
    </keyEncryptor>
  </keyEncryptors>
</encryption>"#,
        key_bits = opts.key_bits,
        hash_size = hash_size,
        spin_count = opts.spin_count,
        hash_alg_name = hash_alg.as_ooxml_name(),
        salt_key_data_b64 = salt_key_data_b64,
        salt_key_encryptor_b64 = salt_key_encryptor_b64,
        enc_vhi_b64 = enc_vhi_b64,
        enc_vhv_b64 = enc_vhv_b64,
        enc_kv_b64 = enc_kv_b64,
        enc_hmac_key_b64 = enc_hmac_key_b64,
        enc_hmac_value_b64 = enc_hmac_value_b64,
    );

    // Build EncryptionInfo stream: version header + xml bytes.
    let flags: u32 = 0x0000_0040;
    let mut encryption_info = Vec::new();
    encryption_info.extend_from_slice(&4u16.to_le_bytes());
    encryption_info.extend_from_slice(&4u16.to_le_bytes());
    encryption_info.extend_from_slice(&flags.to_le_bytes());
    encryption_info.extend_from_slice(xml.as_bytes());

    Ok((encryption_info, encrypted_package))
}

fn encrypt_encrypted_package_stream(
    zip_bytes: &[u8],
    package_key: &[u8],
    hash_alg: HashAlgorithm,
    salt: &[u8],
    block_size: usize,
) -> Result<Vec<u8>, OfficeCryptoError> {
    const SEGMENT_LEN: usize = 4096;
    let original_size = zip_bytes.len() as u64;
    let mut out = Vec::with_capacity(8 + zip_bytes.len());
    out.extend_from_slice(&original_size.to_le_bytes());

    let mut block_index = 0u32;
    for chunk in zip_bytes.chunks(SEGMENT_LEN) {
        let iv = derive_iv(hash_alg, salt, &block_index.to_le_bytes(), block_size);
        let plain = pad_zero(chunk, block_size);
        let enc = aes_cbc_encrypt(package_key, &iv, &plain)?;
        out.extend_from_slice(&enc);
        block_index = block_index.checked_add(1).ok_or_else(|| {
            OfficeCryptoError::InvalidFormat("segment counter overflow".to_string())
        })?;
    }

    Ok(out)
}

fn pad_zero(data: &[u8], block_size: usize) -> Vec<u8> {
    if data.len() % block_size == 0 {
        return data.to_vec();
    }
    let mut out = data.to_vec();
    let pad = block_size - (out.len() % block_size);
    out.extend(std::iter::repeat(0u8).take(pad));
    out
}

fn padded_aes_len(len: usize) -> Result<usize, OfficeCryptoError> {
    let rem = len % 16;
    if rem == 0 {
        Ok(len)
    } else {
        len.checked_add(16 - rem).ok_or_else(|| {
            OfficeCryptoError::InvalidFormat(
                "length overflow while padding to AES block".to_string(),
            )
        })
    }
}

fn compute_hmac(
    hash_alg: HashAlgorithm,
    key: &[u8],
    data: &[u8],
) -> Result<Vec<u8>, OfficeCryptoError> {
    match hash_alg {
        HashAlgorithm::Md5 => {
            let mut mac: Hmac<Md5> = Hmac::new_from_slice(key).map_err(|_| {
                OfficeCryptoError::InvalidFormat("invalid HMAC key".to_string())
            })?;
            mac.update(data);
            Ok(mac.finalize().into_bytes().to_vec())
        }
        HashAlgorithm::Sha1 => {
            let mut mac: Hmac<Sha1> = Hmac::new_from_slice(key).map_err(|_| {
                OfficeCryptoError::InvalidFormat("invalid HMAC key".to_string())
            })?;
            mac.update(data);
            Ok(mac.finalize().into_bytes().to_vec())
        }
        HashAlgorithm::Sha256 => {
            let mut mac: Hmac<Sha256> = Hmac::new_from_slice(key).map_err(|_| {
                OfficeCryptoError::InvalidFormat("invalid HMAC key".to_string())
            })?;
            mac.update(data);
            Ok(mac.finalize().into_bytes().to_vec())
        }
        HashAlgorithm::Sha384 => {
            let mut mac: Hmac<Sha384> = Hmac::new_from_slice(key).map_err(|_| {
                OfficeCryptoError::InvalidFormat("invalid HMAC key".to_string())
            })?;
            mac.update(data);
            Ok(mac.finalize().into_bytes().to_vec())
        }
        HashAlgorithm::Sha512 => {
            let mut mac: Hmac<Sha512> = Hmac::new_from_slice(key).map_err(|_| {
                OfficeCryptoError::InvalidFormat("invalid HMAC key".to_string())
            })?;
            mac.update(data);
            Ok(mac.finalize().into_bytes().to_vec())
        }
    }
}

fn compute_hmac_two(
    hash_alg: HashAlgorithm,
    key: &[u8],
    part1: &[u8],
    part2: &[u8],
) -> Result<Vec<u8>, OfficeCryptoError> {
    match hash_alg {
        HashAlgorithm::Md5 => {
            let mut mac: Hmac<Md5> = Hmac::new_from_slice(key).map_err(|_| {
                OfficeCryptoError::InvalidFormat("invalid HMAC key".to_string())
            })?;
            mac.update(part1);
            mac.update(part2);
            Ok(mac.finalize().into_bytes().to_vec())
        }
        HashAlgorithm::Sha1 => {
            let mut mac: Hmac<Sha1> = Hmac::new_from_slice(key).map_err(|_| {
                OfficeCryptoError::InvalidFormat("invalid HMAC key".to_string())
            })?;
            mac.update(part1);
            mac.update(part2);
            Ok(mac.finalize().into_bytes().to_vec())
        }
        HashAlgorithm::Sha256 => {
            let mut mac: Hmac<Sha256> = Hmac::new_from_slice(key).map_err(|_| {
                OfficeCryptoError::InvalidFormat("invalid HMAC key".to_string())
            })?;
            mac.update(part1);
            mac.update(part2);
            Ok(mac.finalize().into_bytes().to_vec())
        }
        HashAlgorithm::Sha384 => {
            let mut mac: Hmac<Sha384> = Hmac::new_from_slice(key).map_err(|_| {
                OfficeCryptoError::InvalidFormat("invalid HMAC key".to_string())
            })?;
            mac.update(part1);
            mac.update(part2);
            Ok(mac.finalize().into_bytes().to_vec())
        }
        HashAlgorithm::Sha512 => {
            let mut mac: Hmac<Sha512> = Hmac::new_from_slice(key).map_err(|_| {
                OfficeCryptoError::InvalidFormat("invalid HMAC key".to_string())
            })?;
            mac.update(part1);
            mac.update(part2);
            Ok(mac.finalize().into_bytes().to_vec())
        }
    }
}
struct AgileDescriptor {
    key_data: AgileKeyData,
    data_integrity: Option<AgileDataIntegrity>,
    password_key_encryptor: AgilePasswordKeyEncryptor,
}

fn parse_agile_descriptor(xml: &str) -> Result<AgileDescriptor, OfficeCryptoError> {
    let mut reader = Reader::from_str(xml);
    reader.trim_text(true);
    let mut buf = Vec::new();

    let mut key_data: Option<AgileKeyData> = None;
    let mut data_integrity: Option<AgileDataIntegrity> = None;
    let mut password_key_encryptor: Option<AgilePasswordKeyEncryptor> = None;

    let mut in_password_key_encryptor = false;
    let mut in_encrypted_key = false;
    let mut capture: Option<CaptureKind> = None;

    let mut tmp_encrypted_verifier_hash_input: Option<Vec<u8>> = None;
    let mut tmp_encrypted_verifier_hash_value: Option<Vec<u8>> = None;
    let mut tmp_encrypted_key_value: Option<Vec<u8>> = None;

    let mut tmp_password_attrs: Option<AgilePasswordAttrs> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let qname = e.name();
                let name = local_name(qname.as_ref());
                match name {
                    b"keyEncryptor" => {
                        // The `<keyEncryptor>` element indicates how the package key is protected
                        // (password vs certificate). We only support password-based decryption.
                        in_password_key_encryptor = is_password_key_encryptor(&e, &reader)?;
                    }
                    b"keyData" => {
                        let kd = parse_key_data_attrs(&e, &reader)?;
                        key_data = Some(kd);
                    }
                    b"dataIntegrity" => {
                        let di = parse_data_integrity_attrs(&e, &reader)?;
                        data_integrity = Some(di);
                    }
                    b"encryptedKey" if in_password_key_encryptor => {
                        in_encrypted_key = true;
                        tmp_password_attrs = Some(parse_password_key_encryptor_attrs(&e, &reader)?);

                        // Some producers (e.g. `ms_offcrypto_writer`) encode the verifier/key
                        // blobs as base64 attributes on the `<encryptedKey/>` element instead of
                        // child elements. Accept either form.
                        let (vhi, vhv, kv) = parse_encrypted_key_value_attrs(&e, &reader)?;
                        if vhi.is_some() {
                            tmp_encrypted_verifier_hash_input = vhi;
                        }
                        if vhv.is_some() {
                            tmp_encrypted_verifier_hash_value = vhv;
                        }
                        if kv.is_some() {
                            tmp_encrypted_key_value = kv;
                        }
                    }
                    b"encryptedVerifierHashInput" if in_encrypted_key => {
                        capture = Some(CaptureKind::VerifierHashInput);
                    }
                    b"encryptedVerifierHashValue" if in_encrypted_key => {
                        capture = Some(CaptureKind::VerifierHashValue);
                    }
                    b"encryptedKeyValue" if in_encrypted_key => {
                        capture = Some(CaptureKind::KeyValue);
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(e)) => {
                let qname = e.name();
                let name = local_name(qname.as_ref());
                match name {
                    b"keyData" => {
                        let kd = parse_key_data_attrs(&e, &reader)?;
                        key_data = Some(kd);
                    }
                    b"dataIntegrity" => {
                        let di = parse_data_integrity_attrs(&e, &reader)?;
                        data_integrity = Some(di);
                    }
                    b"encryptedKey" if in_password_key_encryptor => {
                        let attrs = parse_password_key_encryptor_attrs(&e, &reader)?;
                        let (vhi, vhv, kv) = parse_encrypted_key_value_attrs(&e, &reader)?;
                        password_key_encryptor = Some(AgilePasswordKeyEncryptor {
                            salt: attrs.salt,
                            block_size: attrs.block_size,
                            key_bits: attrs.key_bits,
                            spin_count: attrs.spin_count,
                            hash_algorithm: attrs.hash_algorithm,
                            hash_size: attrs.hash_size,
                            cipher_algorithm: attrs.cipher_algorithm,
                            cipher_chaining: attrs.cipher_chaining,
                            encrypted_verifier_hash_input: vhi.ok_or_else(|| {
                                OfficeCryptoError::InvalidFormat(
                                    "missing encryptedVerifierHashInput".to_string(),
                                )
                            })?,
                            encrypted_verifier_hash_value: vhv.ok_or_else(|| {
                                OfficeCryptoError::InvalidFormat(
                                    "missing encryptedVerifierHashValue".to_string(),
                                )
                            })?,
                            encrypted_key_value: kv.ok_or_else(|| {
                                OfficeCryptoError::InvalidFormat(
                                    "missing encryptedKeyValue".to_string(),
                                )
                            })?,
                        });
                    }
                    _ => {}
                }
            }
            Ok(Event::End(e)) => {
                let qname = e.name();
                let name = local_name(qname.as_ref());
                if name == b"keyEncryptor" {
                    in_password_key_encryptor = false;
                }
                if name == b"encryptedKey" && in_encrypted_key {
                    in_encrypted_key = false;
                    capture = None;
                    let attrs = tmp_password_attrs.take().ok_or_else(|| {
                        OfficeCryptoError::InvalidFormat(
                            "encryptedKey missing required attributes".to_string(),
                        )
                    })?;
                    let encrypted_verifier_hash_input =
                        tmp_encrypted_verifier_hash_input.take().ok_or_else(|| {
                            OfficeCryptoError::InvalidFormat(
                                "missing encryptedVerifierHashInput".to_string(),
                            )
                        })?;
                    let encrypted_verifier_hash_value =
                        tmp_encrypted_verifier_hash_value.take().ok_or_else(|| {
                            OfficeCryptoError::InvalidFormat(
                                "missing encryptedVerifierHashValue".to_string(),
                            )
                        })?;
                    let encrypted_key_value = tmp_encrypted_key_value.take().ok_or_else(|| {
                        OfficeCryptoError::InvalidFormat("missing encryptedKeyValue".to_string())
                    })?;
                    password_key_encryptor = Some(AgilePasswordKeyEncryptor {
                        salt: attrs.salt,
                        block_size: attrs.block_size,
                        key_bits: attrs.key_bits,
                        spin_count: attrs.spin_count,
                        hash_algorithm: attrs.hash_algorithm,
                        hash_size: attrs.hash_size,
                        cipher_algorithm: attrs.cipher_algorithm,
                        cipher_chaining: attrs.cipher_chaining,
                        encrypted_verifier_hash_input,
                        encrypted_verifier_hash_value,
                        encrypted_key_value,
                    });
                }
                if matches!(
                    name,
                    b"encryptedVerifierHashInput"
                        | b"encryptedVerifierHashValue"
                        | b"encryptedKeyValue"
                ) {
                    capture = None;
                }
            }
            Ok(Event::Text(t)) => {
                if let Some(kind) = capture {
                    let text = t
                        .unescape()
                        .map_err(|_| {
                            OfficeCryptoError::InvalidFormat(
                                "invalid XML escape in base64 text".to_string(),
                            )
                        })?
                        .to_string();
                    let decoded = decode_b64_attr(&text).map_err(|_| {
                        OfficeCryptoError::InvalidFormat(
                            "invalid base64 in EncryptionInfo".to_string(),
                        )
                    })?;
                    match kind {
                        CaptureKind::VerifierHashInput => {
                            tmp_encrypted_verifier_hash_input = Some(decoded)
                        }
                        CaptureKind::VerifierHashValue => {
                            tmp_encrypted_verifier_hash_value = Some(decoded)
                        }
                        CaptureKind::KeyValue => tmp_encrypted_key_value = Some(decoded),
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(OfficeCryptoError::InvalidFormat(format!(
                    "failed to parse EncryptionInfo XML: {e}"
                )))
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(AgileDescriptor {
        key_data: key_data.ok_or_else(|| {
            OfficeCryptoError::InvalidFormat("missing keyData element".to_string())
        })?,
        // Some third-party producers omit `<dataIntegrity>` entirely. Excel treats the HMAC as
        // optional (integrity check is skipped when missing), so be permissive and allow it.
        data_integrity,
        password_key_encryptor: password_key_encryptor.ok_or_else(|| {
            OfficeCryptoError::InvalidFormat("missing password keyEncryptor".to_string())
        })?,
    })
}

#[derive(Clone, Copy)]
enum CaptureKind {
    VerifierHashInput,
    VerifierHashValue,
    KeyValue,
}

fn local_name(name: &[u8]) -> &[u8] {
    match name.iter().position(|&b| b == b':') {
        Some(idx) => &name[idx + 1..],
        None => name,
    }
}

fn is_password_key_encryptor<B: std::io::BufRead>(
    e: &quick_xml::events::BytesStart<'_>,
    reader: &Reader<B>,
) -> Result<bool, OfficeCryptoError> {
    const PASSWORD_URI: &str = "http://schemas.microsoft.com/office/2006/keyEncryptor/password";
    for attr in e.attributes() {
        let attr = attr
            .map_err(|_| OfficeCryptoError::InvalidFormat("invalid XML attribute".to_string()))?;
        let key = local_name(attr.key.as_ref());
        if key != b"uri" {
            continue;
        }
        let value = attr.decode_and_unescape_value(reader).map_err(|_| {
            OfficeCryptoError::InvalidFormat("invalid XML attribute encoding".to_string())
        })?;
        return Ok(value.as_ref() == PASSWORD_URI);
    }
    Ok(false)
}

fn parse_key_data_attrs<B: std::io::BufRead>(
    e: &quick_xml::events::BytesStart<'_>,
    reader: &Reader<B>,
) -> Result<AgileKeyData, OfficeCryptoError> {
    let mut salt_value: Option<Vec<u8>> = None;
    let mut block_size: Option<usize> = None;
    let mut key_bits: Option<usize> = None;
    let mut hash_algorithm: Option<HashAlgorithm> = None;
    let mut hash_size: Option<usize> = None;
    let mut cipher_algorithm: Option<String> = None;
    let mut cipher_chaining: Option<String> = None;

    for attr in e.attributes() {
        let attr = attr
            .map_err(|_| OfficeCryptoError::InvalidFormat("invalid XML attribute".to_string()))?;
        let key = local_name(attr.key.as_ref());
        let value = attr.decode_and_unescape_value(reader).map_err(|_| {
            OfficeCryptoError::InvalidFormat("invalid XML attribute encoding".to_string())
        })?;
        match key {
            b"saltValue" => {
                salt_value = Some(decode_b64_attr(value.as_ref()).map_err(|_| {
                    OfficeCryptoError::InvalidFormat("invalid base64 saltValue".to_string())
                })?);
            }
            b"blockSize" => {
                block_size = Some(value.parse::<usize>().map_err(|_| {
                    OfficeCryptoError::InvalidFormat("invalid blockSize".to_string())
                })?);
            }
            b"keyBits" => {
                key_bits = Some(value.parse::<usize>().map_err(|_| {
                    OfficeCryptoError::InvalidFormat("invalid keyBits".to_string())
                })?);
            }
            b"hashAlgorithm" => {
                hash_algorithm = Some(HashAlgorithm::from_name(value.as_ref())?);
            }
            b"hashSize" => {
                hash_size = Some(value.parse::<usize>().map_err(|_| {
                    OfficeCryptoError::InvalidFormat("invalid hashSize".to_string())
                })?);
            }
            b"cipherAlgorithm" => {
                cipher_algorithm = Some(value.as_ref().to_string());
            }
            b"cipherChaining" => {
                cipher_chaining = Some(value.as_ref().to_string());
            }
            _ => {}
        }
    }

    if let Some(bits) = key_bits {
        if bits > 512 {
            return Err(OfficeCryptoError::SizeLimitExceeded {
                context: "keyData.keyBits",
                limit: 512,
            });
        }
        if bits % 8 != 0 {
            return Err(OfficeCryptoError::InvalidFormat(
                "keyData.keyBits must be divisible by 8".to_string(),
            ));
        }
    }
    if let Some(bs) = block_size {
        if bs > 1024 {
            return Err(OfficeCryptoError::SizeLimitExceeded {
                context: "keyData.blockSize",
                limit: 1024,
            });
        }
    }

    let hash_algorithm = hash_algorithm.ok_or_else(|| {
        OfficeCryptoError::InvalidFormat("keyData missing hashAlgorithm".to_string())
    })?;
    let hash_size = hash_size.unwrap_or_else(|| hash_algorithm.digest_len());
    if hash_size == 0 {
        return Err(OfficeCryptoError::InvalidFormat(
            "keyData hashSize must be non-zero".to_string(),
        ));
    }
    let digest_len = hash_algorithm.digest_len();
    if hash_size > digest_len {
        return Err(OfficeCryptoError::InvalidFormat(format!(
            "keyData hashSize {hash_size} exceeds {} digest length {digest_len}",
            hash_algorithm.as_ooxml_name()
        )));
    }
    Ok(AgileKeyData {
        salt: salt_value.ok_or_else(|| {
            OfficeCryptoError::InvalidFormat("keyData missing saltValue".to_string())
        })?,
        block_size: block_size.ok_or_else(|| {
            OfficeCryptoError::InvalidFormat("keyData missing blockSize".to_string())
        })?,
        key_bits: key_bits.ok_or_else(|| {
            OfficeCryptoError::InvalidFormat("keyData missing keyBits".to_string())
        })?,
        hash_algorithm,
        hash_size,
        cipher_algorithm: cipher_algorithm.ok_or_else(|| {
            OfficeCryptoError::InvalidFormat("keyData missing cipherAlgorithm".to_string())
        })?,
        cipher_chaining: cipher_chaining.ok_or_else(|| {
            OfficeCryptoError::InvalidFormat("keyData missing cipherChaining".to_string())
        })?,
    })
}

fn parse_data_integrity_attrs<B: std::io::BufRead>(
    e: &quick_xml::events::BytesStart<'_>,
    reader: &Reader<B>,
) -> Result<AgileDataIntegrity, OfficeCryptoError> {
    let mut encrypted_hmac_key: Option<Vec<u8>> = None;
    let mut encrypted_hmac_value: Option<Vec<u8>> = None;
    for attr in e.attributes() {
        let attr = attr
            .map_err(|_| OfficeCryptoError::InvalidFormat("invalid XML attribute".to_string()))?;
        let key = local_name(attr.key.as_ref());
        let value = attr.decode_and_unescape_value(reader).map_err(|_| {
            OfficeCryptoError::InvalidFormat("invalid XML attribute encoding".to_string())
        })?;
        match key {
            b"encryptedHmacKey" => {
                encrypted_hmac_key = Some(decode_b64_attr(value.as_ref()).map_err(|_| {
                    OfficeCryptoError::InvalidFormat("invalid base64 encryptedHmacKey".to_string())
                })?);
            }
            b"encryptedHmacValue" => {
                encrypted_hmac_value = Some(decode_b64_attr(value.as_ref()).map_err(|_| {
                    OfficeCryptoError::InvalidFormat(
                        "invalid base64 encryptedHmacValue".to_string(),
                    )
                })?);
            }
            _ => {}
        }
    }
    Ok(AgileDataIntegrity {
        encrypted_hmac_key: encrypted_hmac_key.ok_or_else(|| {
            OfficeCryptoError::InvalidFormat("dataIntegrity missing encryptedHmacKey".to_string())
        })?,
        encrypted_hmac_value: encrypted_hmac_value.ok_or_else(|| {
            OfficeCryptoError::InvalidFormat("dataIntegrity missing encryptedHmacValue".to_string())
        })?,
    })
}

fn parse_encrypted_key_value_attrs<B: std::io::BufRead>(
    e: &quick_xml::events::BytesStart<'_>,
    reader: &Reader<B>,
) -> Result<(Option<Vec<u8>>, Option<Vec<u8>>, Option<Vec<u8>>), OfficeCryptoError> {
    let mut encrypted_verifier_hash_input: Option<Vec<u8>> = None;
    let mut encrypted_verifier_hash_value: Option<Vec<u8>> = None;
    let mut encrypted_key_value: Option<Vec<u8>> = None;

    for attr in e.attributes() {
        let attr = attr
            .map_err(|_| OfficeCryptoError::InvalidFormat("invalid XML attribute".to_string()))?;
        let key = local_name(attr.key.as_ref());
        let value = attr.decode_and_unescape_value(reader).map_err(|_| {
            OfficeCryptoError::InvalidFormat("invalid XML attribute encoding".to_string())
        })?;
        match key {
            b"encryptedVerifierHashInput" => {
                encrypted_verifier_hash_input =
                    Some(decode_b64_attr(value.as_ref()).map_err(|_| {
                        OfficeCryptoError::InvalidFormat(
                            "invalid base64 encryptedVerifierHashInput".to_string(),
                        )
                    })?);
            }
            b"encryptedVerifierHashValue" => {
                encrypted_verifier_hash_value =
                    Some(decode_b64_attr(value.as_ref()).map_err(|_| {
                        OfficeCryptoError::InvalidFormat(
                            "invalid base64 encryptedVerifierHashValue".to_string(),
                        )
                    })?);
            }
            b"encryptedKeyValue" => {
                encrypted_key_value = Some(decode_b64_attr(value.as_ref()).map_err(|_| {
                    OfficeCryptoError::InvalidFormat("invalid base64 encryptedKeyValue".to_string())
                })?);
            }
            _ => {}
        }
    }

    Ok((
        encrypted_verifier_hash_input,
        encrypted_verifier_hash_value,
        encrypted_key_value,
    ))
}

#[derive(Debug)]
struct AgilePasswordAttrs {
    salt: Vec<u8>,
    block_size: usize,
    key_bits: usize,
    spin_count: u32,
    hash_algorithm: HashAlgorithm,
    hash_size: usize,
    cipher_algorithm: String,
    cipher_chaining: String,
}

fn parse_password_key_encryptor_attrs<B: std::io::BufRead>(
    e: &quick_xml::events::BytesStart<'_>,
    reader: &Reader<B>,
) -> Result<AgilePasswordAttrs, OfficeCryptoError> {
    let mut salt_value: Option<Vec<u8>> = None;
    let mut block_size: Option<usize> = None;
    let mut key_bits: Option<usize> = None;
    let mut spin_count: Option<u32> = None;
    let mut hash_algorithm: Option<HashAlgorithm> = None;
    let mut hash_size: Option<usize> = None;
    let mut cipher_algorithm: Option<String> = None;
    let mut cipher_chaining: Option<String> = None;

    for attr in e.attributes() {
        let attr = attr
            .map_err(|_| OfficeCryptoError::InvalidFormat("invalid XML attribute".to_string()))?;
        let key = local_name(attr.key.as_ref());
        let value = attr.decode_and_unescape_value(reader).map_err(|_| {
            OfficeCryptoError::InvalidFormat("invalid XML attribute encoding".to_string())
        })?;
        match key {
            b"saltValue" => {
                salt_value = Some(decode_b64_attr(value.as_ref()).map_err(|_| {
                    OfficeCryptoError::InvalidFormat("invalid base64 saltValue".to_string())
                })?);
            }
            b"blockSize" => {
                block_size = Some(value.parse::<usize>().map_err(|_| {
                    OfficeCryptoError::InvalidFormat("invalid blockSize".to_string())
                })?);
            }
            b"keyBits" => {
                key_bits = Some(value.parse::<usize>().map_err(|_| {
                    OfficeCryptoError::InvalidFormat("invalid keyBits".to_string())
                })?);
            }
            b"spinCount" => {
                spin_count = Some(value.parse::<u32>().map_err(|_| {
                    OfficeCryptoError::InvalidFormat("invalid spinCount".to_string())
                })?);
            }
            b"hashAlgorithm" => {
                hash_algorithm = Some(HashAlgorithm::from_name(value.as_ref())?);
            }
            b"hashSize" => {
                hash_size = Some(value.parse::<usize>().map_err(|_| {
                    OfficeCryptoError::InvalidFormat("invalid hashSize".to_string())
                })?);
            }
            b"cipherAlgorithm" => {
                cipher_algorithm = Some(value.as_ref().to_string());
            }
            b"cipherChaining" => {
                cipher_chaining = Some(value.as_ref().to_string());
            }
            _ => {}
        }
    }

    if let Some(bits) = key_bits {
        if bits > 512 {
            return Err(OfficeCryptoError::SizeLimitExceeded {
                context: "encryptedKey.keyBits",
                limit: 512,
            });
        }
        if bits % 8 != 0 {
            return Err(OfficeCryptoError::InvalidFormat(
                "encryptedKey.keyBits must be divisible by 8".to_string(),
            ));
        }
    }
    if let Some(bs) = block_size {
        if bs > 1024 {
            return Err(OfficeCryptoError::SizeLimitExceeded {
                context: "encryptedKey.blockSize",
                limit: 1024,
            });
        }
    }

    let hash_algorithm = hash_algorithm.ok_or_else(|| {
        OfficeCryptoError::InvalidFormat("encryptedKey missing hashAlgorithm".to_string())
    })?;
    let hash_size = hash_size.unwrap_or_else(|| hash_algorithm.digest_len());
    if hash_size == 0 {
        return Err(OfficeCryptoError::InvalidFormat(
            "encryptedKey hashSize must be non-zero".to_string(),
        ));
    }
    let digest_len = hash_algorithm.digest_len();
    if hash_size > digest_len {
        return Err(OfficeCryptoError::InvalidFormat(format!(
            "encryptedKey hashSize {hash_size} exceeds {} digest length {digest_len}",
            hash_algorithm.as_ooxml_name()
        )));
    }
    Ok(AgilePasswordAttrs {
        salt: salt_value.ok_or_else(|| {
            OfficeCryptoError::InvalidFormat("encryptedKey missing saltValue".to_string())
        })?,
        block_size: block_size.ok_or_else(|| {
            OfficeCryptoError::InvalidFormat("encryptedKey missing blockSize".to_string())
        })?,
        key_bits: key_bits.ok_or_else(|| {
            OfficeCryptoError::InvalidFormat("encryptedKey missing keyBits".to_string())
        })?,
        spin_count: spin_count.ok_or_else(|| {
            OfficeCryptoError::InvalidFormat("encryptedKey missing spinCount".to_string())
        })?,
        hash_algorithm,
        hash_size,
        cipher_algorithm: cipher_algorithm.ok_or_else(|| {
            OfficeCryptoError::InvalidFormat("encryptedKey missing cipherAlgorithm".to_string())
        })?,
        cipher_chaining: cipher_chaining.ok_or_else(|| {
            OfficeCryptoError::InvalidFormat("encryptedKey missing cipherChaining".to_string())
        })?,
    })
}

fn decode_b64_attr(value: &str) -> Result<Vec<u8>, base64::DecodeError> {
    let bytes = value.as_bytes();

    // Avoid allocating for the common case where no whitespace is present.
    let mut cleaned: Option<Vec<u8>> = None;
    for (idx, &b) in bytes.iter().enumerate() {
        if matches!(b, b'\r' | b'\n' | b'\t' | b' ') {
            let mut out = Vec::with_capacity(bytes.len());
            out.extend_from_slice(&bytes[..idx]);
            for &b2 in &bytes[idx..] {
                if !matches!(b2, b'\r' | b'\n' | b'\t' | b' ') {
                    out.push(b2);
                }
            }
            cleaned = Some(out);
            break;
        }
    }

    let input = cleaned.as_deref().unwrap_or(bytes);
    BASE64_STANDARD
        .decode(input)
        .or_else(|_| BASE64_STANDARD_NO_PAD.decode(input))
}

#[cfg(test)]
pub(crate) mod tests {
    use super::*;
    use crate::crypto::{
        aes_cbc_encrypt, derive_agile_key, derive_iv, password_to_utf16le, HashAlgorithm,
    };
    use crate::util::parse_encryption_info_header;
    use crate::OfficeCryptoError;

    #[test]
    fn default_max_spin_count_is_one_million() {
        assert_eq!(crate::DEFAULT_MAX_SPIN_COUNT, 1_000_000);
        assert_eq!(
            crate::DecryptOptions::default().max_spin_count,
            crate::DEFAULT_MAX_SPIN_COUNT
        );
    }

    #[test]
    fn compute_hmac_two_matches_compute_hmac_over_concatenated_bytes() {
        // Regression test for the `header + plaintext` HMAC compatibility target.
        // `compute_hmac_two` is used to avoid allocating a temporary `header||plaintext` buffer.
        let key = b"formula-hmac-key";
        let a = b"header";
        let b = b"plaintext";
        for hash_alg in [
            HashAlgorithm::Md5,
            HashAlgorithm::Sha1,
            HashAlgorithm::Sha256,
            HashAlgorithm::Sha384,
            HashAlgorithm::Sha512,
        ] {
            let mut combined = Vec::new();
            combined.extend_from_slice(a);
            combined.extend_from_slice(b);
            let one = compute_hmac(hash_alg, key, &combined).expect("compute_hmac");
            let two = compute_hmac_two(hash_alg, key, a, b).expect("compute_hmac_two");
            assert_eq!(one, two, "hash_alg={hash_alg:?}");

            // Also exercise empty suffix/prefix combinations.
            let one = compute_hmac(hash_alg, key, a).expect("compute_hmac");
            let two = compute_hmac_two(hash_alg, key, a, b"").expect("compute_hmac_two");
            assert_eq!(one, two, "hash_alg={hash_alg:?} (empty b)");

            let one = compute_hmac(hash_alg, key, b).expect("compute_hmac");
            let two = compute_hmac_two(hash_alg, key, b"", b).expect("compute_hmac_two");
            assert_eq!(one, two, "hash_alg={hash_alg:?} (empty a)");
        }
    }

    pub(crate) fn agile_encryption_info_fixture() -> Vec<u8> {
        // A small, deterministic Agile EncryptionInfo fixture for parsing tests.
        let xml = agile_descriptor_fixture_xml();
        let version_major = 4u16;
        let version_minor = 4u16;
        let flags = 0x0000_0040u32;
        let xml_len = xml.as_bytes().len() as u32;

        let mut out = Vec::new();
        out.extend_from_slice(&version_major.to_le_bytes());
        out.extend_from_slice(&version_minor.to_le_bytes());
        out.extend_from_slice(&flags.to_le_bytes());
        out.extend_from_slice(&xml_len.to_le_bytes());
        out.extend_from_slice(xml.as_bytes());

        let hdr = parse_encryption_info_header(&out).expect("header");
        assert_eq!(hdr.kind, crate::util::EncryptionInfoKind::Agile);
        out
    }

    pub(crate) fn agile_descriptor_fixture_xml() -> String {
        // Build a minimal-but-valid agile descriptor (values not meant to be secure).
        let password = "Password";
        let pw_utf16 = password_to_utf16le(password);
        let hash_alg = HashAlgorithm::Sha512;
        let spin_count = 100_000u32;
        let key_bits = 256usize;

        let salt_key_encryptor: [u8; 16] = [
            0xA0, 0xA1, 0xA2, 0xA3, 0xA4, 0xA5, 0xA6, 0xA7, 0xA8, 0xA9, 0xAA, 0xAB, 0xAC, 0xAD,
            0xAE, 0xAF,
        ];
        let salt_key_data: [u8; 16] = [
            0xB0, 0xB1, 0xB2, 0xB3, 0xB4, 0xB5, 0xB6, 0xB7, 0xB8, 0xB9, 0xBA, 0xBB, 0xBC, 0xBD,
            0xBE, 0xBF,
        ];

        let verifier_hash_input_plain: [u8; 16] = *b"formula-agl-test";
        let verifier_hash_value_plain = hash_alg.digest(&verifier_hash_input_plain);
        let package_key_plain: [u8; 32] = [0x11; 32];

        let key_vhi = derive_agile_key(
            hash_alg,
            &salt_key_encryptor,
            &pw_utf16,
            spin_count,
            key_bits / 8,
            BLOCK_KEY_VERIFIER_HASH_INPUT,
        );
        let iv_vhi = salt_key_encryptor.to_vec();
        let enc_vhi =
            aes_cbc_encrypt(&key_vhi, &iv_vhi, &verifier_hash_input_plain).expect("enc vhi");

        let key_vhv = derive_agile_key(
            hash_alg,
            &salt_key_encryptor,
            &pw_utf16,
            spin_count,
            key_bits / 8,
            BLOCK_KEY_VERIFIER_HASH_VALUE,
        );
        let iv_vhv = salt_key_encryptor.to_vec();
        let enc_vhv =
            aes_cbc_encrypt(&key_vhv, &iv_vhv, &verifier_hash_value_plain).expect("enc vhv");

        let key_kv = derive_agile_key(
            hash_alg,
            &salt_key_encryptor,
            &pw_utf16,
            spin_count,
            key_bits / 8,
            BLOCK_KEY_ENCRYPTED_KEY_VALUE,
        );
        let iv_kv = salt_key_encryptor.to_vec();
        let enc_kv = aes_cbc_encrypt(&key_kv, &iv_kv, &package_key_plain).expect("enc key");

        let b64 = base64::engine::general_purpose::STANDARD;
        let salt_key_encryptor_b64 = b64.encode(salt_key_encryptor);
        let salt_key_data_b64 = b64.encode(salt_key_data);
        let enc_vhi_b64 = b64.encode(enc_vhi);
        let enc_vhv_b64 = b64.encode(enc_vhv);
        let enc_kv_b64 = b64.encode(enc_kv);

        // Dummy integrity fields.
        let dummy = b64.encode([0u8; 32]);

        format!(
            r#"<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption">
  <keyData saltSize="16" blockSize="16" keyBits="256" hashAlgorithm="SHA512" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" saltValue="{salt_key_data_b64}"/>
  <dataIntegrity encryptedHmacKey="{dummy}" encryptedHmacValue="{dummy}"/>
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password"
        saltSize="16" blockSize="16" keyBits="256" spinCount="100000" hashAlgorithm="SHA512" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" saltValue="{salt_key_encryptor_b64}">
        <p:encryptedVerifierHashInput>{enc_vhi_b64}</p:encryptedVerifierHashInput>
        <p:encryptedVerifierHashValue>{enc_vhv_b64}</p:encryptedVerifierHashValue>
        <p:encryptedKeyValue>{enc_kv_b64}</p:encryptedKeyValue>
      </p:encryptedKey>
    </keyEncryptor>
  </keyEncryptors>
</encryption>"#
        )
    }

    #[test]
    fn password_encrypted_key_uses_saltvalue_as_cbc_iv() {
        // Regression test: for `p:encryptedKey` fields, the AES-CBC IV is the saltValue itself
        // (truncated to blockSize), not Hash(saltValue || blockKey).

        let password = "Password";
        let pw_utf16 = password_to_utf16le(password);
        let hash_alg = HashAlgorithm::Sha512;
        let spin_count = 1_000u32;
        let key_bits = 256usize;
        let block_size = 16usize;

        let salt_key_encryptor: [u8; 16] = [
            0xA0, 0xA1, 0xA2, 0xA3, 0xA4, 0xA5, 0xA6, 0xA7, 0xA8, 0xA9, 0xAA, 0xAB, 0xAC, 0xAD,
            0xAE, 0xAF,
        ];
        let salt_key_data: [u8; 16] = [
            0xB0, 0xB1, 0xB2, 0xB3, 0xB4, 0xB5, 0xB6, 0xB7, 0xB8, 0xB9, 0xBA, 0xBB, 0xBC, 0xBD,
            0xBE, 0xBF,
        ];

        // Empty ciphertext: the test focuses on password verifier + package key unwrap.
        let encrypted_package = 0u64.to_le_bytes().to_vec();

        let verifier_hash_input_plain: [u8; 16] = *b"formula-agl-test";
        let verifier_hash_value_plain = hash_alg.digest(&verifier_hash_input_plain);
        let package_key_plain: [u8; 32] = [0x11; 32];

        // Derive keys (block keys are used only for key derivation).
        let key_vhi = derive_agile_key(
            hash_alg,
            &salt_key_encryptor,
            &pw_utf16,
            spin_count,
            key_bits / 8,
            BLOCK_KEY_VERIFIER_HASH_INPUT,
        );
        let key_vhv = derive_agile_key(
            hash_alg,
            &salt_key_encryptor,
            &pw_utf16,
            spin_count,
            key_bits / 8,
            BLOCK_KEY_VERIFIER_HASH_VALUE,
        );
        let key_kv = derive_agile_key(
            hash_alg,
            &salt_key_encryptor,
            &pw_utf16,
            spin_count,
            key_bits / 8,
            BLOCK_KEY_ENCRYPTED_KEY_VALUE,
        );

        // Correct IV: saltValue itself.
        let verifier_iv = &salt_key_encryptor[..block_size];
        let enc_vhi =
            aes_cbc_encrypt(&key_vhi, verifier_iv, &verifier_hash_input_plain).expect("enc vhi");
        let enc_vhv =
            aes_cbc_encrypt(&key_vhv, verifier_iv, &verifier_hash_value_plain).expect("enc vhv");
        let enc_kv = aes_cbc_encrypt(&key_kv, verifier_iv, &package_key_plain).expect("enc key");

        // Data integrity: create a valid HMAC for this (empty) EncryptedPackage stream.
        let digest_len = hash_alg.digest_len();
        let hmac_key_plain = vec![0x22u8; digest_len];
        let computed_hmac =
            compute_hmac(hash_alg, &hmac_key_plain, &encrypted_package).expect("hmac");

        let iv_hmac_key = derive_iv(
            hash_alg,
            &salt_key_data,
            BLOCK_KEY_INTEGRITY_HMAC_KEY,
            block_size,
        );
        let encrypted_hmac_key = aes_cbc_encrypt(
            &package_key_plain,
            &iv_hmac_key,
            &pad_zero(&hmac_key_plain, block_size),
        )
        .expect("enc hmac key");

        let iv_hmac_val = derive_iv(
            hash_alg,
            &salt_key_data,
            BLOCK_KEY_INTEGRITY_HMAC_VALUE,
            block_size,
        );
        let encrypted_hmac_value = aes_cbc_encrypt(
            &package_key_plain,
            &iv_hmac_val,
            &pad_zero(&computed_hmac, block_size),
        )
        .expect("enc hmac value");

        let info = AgileEncryptionInfo {
            version_major: 4,
            version_minor: 4,
            flags: 0,
            key_data: AgileKeyData {
                salt: salt_key_data.to_vec(),
                block_size,
                key_bits,
                hash_algorithm: hash_alg,
                hash_size: digest_len,
                cipher_algorithm: "AES".to_string(),
                cipher_chaining: "ChainingModeCBC".to_string(),
            },
            data_integrity: Some(AgileDataIntegrity {
                encrypted_hmac_key,
                encrypted_hmac_value,
            }),
            password_key_encryptor: AgilePasswordKeyEncryptor {
                salt: salt_key_encryptor.to_vec(),
                block_size,
                key_bits,
                spin_count,
                hash_algorithm: hash_alg,
                hash_size: digest_len,
                cipher_algorithm: "AES".to_string(),
                cipher_chaining: "ChainingModeCBC".to_string(),
                encrypted_verifier_hash_input: enc_vhi,
                encrypted_verifier_hash_value: enc_vhv,
                encrypted_key_value: enc_kv,
            },
        };

        let out = decrypt_agile_encrypted_package(
            &info,
            &encrypted_package,
            password,
            &crate::DecryptOptions::default(),
        )
        .expect("decrypt");
        assert!(out.is_empty());

        // Alternative IV mode (seen in other producers): Hash(saltValue || blockKey).
        let wrong_iv_vhi = derive_iv(
            hash_alg,
            &salt_key_encryptor,
            BLOCK_KEY_VERIFIER_HASH_INPUT,
            block_size,
        );
        assert_ne!(wrong_iv_vhi.as_slice(), verifier_iv);
        let wrong_iv_vhv = derive_iv(
            hash_alg,
            &salt_key_encryptor,
            BLOCK_KEY_VERIFIER_HASH_VALUE,
            block_size,
        );
        let wrong_iv_kv = derive_iv(
            hash_alg,
            &salt_key_encryptor,
            BLOCK_KEY_ENCRYPTED_KEY_VALUE,
            block_size,
        );

        let wrong_enc_vhi =
            aes_cbc_encrypt(&key_vhi, &wrong_iv_vhi, &verifier_hash_input_plain).expect("enc vhi");
        let wrong_enc_vhv =
            aes_cbc_encrypt(&key_vhv, &wrong_iv_vhv, &verifier_hash_value_plain).expect("enc vhv");
        let wrong_enc_kv =
            aes_cbc_encrypt(&key_kv, &wrong_iv_kv, &package_key_plain).expect("enc key");

        let mut wrong_info = info.clone();
        wrong_info
            .password_key_encryptor
            .encrypted_verifier_hash_input = wrong_enc_vhi;
        wrong_info
            .password_key_encryptor
            .encrypted_verifier_hash_value = wrong_enc_vhv;
        wrong_info.password_key_encryptor.encrypted_key_value = wrong_enc_kv;

        let out = decrypt_agile_encrypted_package(
            &wrong_info,
            &encrypted_package,
            password,
            &crate::DecryptOptions::default(),
        )
        .expect("decrypt derived-iv variant");
        assert!(out.is_empty());
    }

    #[test]
    fn agile_decrypt_ignores_trailing_padding_in_verifier_and_hmac_values() {
        // `encryptedVerifierHashValue` and `encryptedHmacValue` decrypt to AES-block-aligned buffers.
        // When `hashSize` is not a multiple of 16 (e.g. SHA1=20 bytes), producers may include
        // trailing padding bytes. Validation should compare only the digest prefix.
        let password = "pw";
        let pw_utf16 = password_to_utf16le(password);
        let hash_alg = HashAlgorithm::Sha1;
        let spin_count = 10u32;
        let key_bits = 128usize;
        let key_bytes = key_bits / 8;
        let block_size = 16usize;
        let digest_len = hash_alg.digest_len();
        assert_eq!(digest_len, 20);

        let key_data_salt: Vec<u8> = (0u8..=15).collect();
        let password_salt: Vec<u8> = (16u8..=31).collect();

        // Empty ciphertext: the test focuses on verifier + HMAC validation logic.
        let encrypted_package = 0u64.to_le_bytes().to_vec();

        let verifier_hash_input_plain: [u8; 16] = *b"abcdefghijklmnop";
        let mut verifier_hash_value_plain = hash_alg.digest(&verifier_hash_input_plain);
        verifier_hash_value_plain.extend_from_slice(&[0xA5u8; 12]);
        assert_eq!(verifier_hash_value_plain.len(), 32);

        let package_key_plain: [u8; 16] = [0x11; 16];

        // Derive keys (block keys are used only for key derivation).
        let key_vhi = derive_agile_key(
            hash_alg,
            &password_salt,
            &pw_utf16,
            spin_count,
            key_bytes,
            BLOCK_KEY_VERIFIER_HASH_INPUT,
        );
        let key_vhv = derive_agile_key(
            hash_alg,
            &password_salt,
            &pw_utf16,
            spin_count,
            key_bytes,
            BLOCK_KEY_VERIFIER_HASH_VALUE,
        );
        let key_kv = derive_agile_key(
            hash_alg,
            &password_salt,
            &pw_utf16,
            spin_count,
            key_bytes,
            BLOCK_KEY_ENCRYPTED_KEY_VALUE,
        );

        // Password key encryptor IV scheme: saltValue.
        let verifier_iv = password_salt
            .get(..block_size)
            .expect("saltValue shorter than blockSize");
        let enc_vhi =
            aes_cbc_encrypt(&key_vhi, verifier_iv, &verifier_hash_input_plain).expect("enc vhi");
        let enc_vhv =
            aes_cbc_encrypt(&key_vhv, verifier_iv, &verifier_hash_value_plain).expect("enc vhv");
        let enc_kv = aes_cbc_encrypt(&key_kv, verifier_iv, &package_key_plain).expect("enc key");

        // HMAC key/value fields with non-zero trailing bytes.
        let hmac_key_plain = vec![0x22u8; digest_len];
        let computed_hmac = compute_hmac(hash_alg, &hmac_key_plain, &encrypted_package)
            .expect("compute hmac");
        assert_eq!(computed_hmac.len(), digest_len);

        let mut hmac_key_plain_padded = hmac_key_plain.clone();
        hmac_key_plain_padded.extend_from_slice(&[0x5Au8; 12]);
        let mut hmac_value_plain_padded = computed_hmac.clone();
        hmac_value_plain_padded.extend_from_slice(&[0xC3u8; 12]);

        let iv_hmac_key = derive_iv(
            hash_alg,
            &key_data_salt,
            BLOCK_KEY_INTEGRITY_HMAC_KEY,
            block_size,
        );
        let encrypted_hmac_key =
            aes_cbc_encrypt(&package_key_plain, &iv_hmac_key, &hmac_key_plain_padded)
                .expect("enc hmac key");

        let iv_hmac_val = derive_iv(
            hash_alg,
            &key_data_salt,
            BLOCK_KEY_INTEGRITY_HMAC_VALUE,
            block_size,
        );
        let encrypted_hmac_value =
            aes_cbc_encrypt(&package_key_plain, &iv_hmac_val, &hmac_value_plain_padded)
                .expect("enc hmac value");

        let info = AgileEncryptionInfo {
            version_major: 4,
            version_minor: 4,
            flags: 0,
            key_data: AgileKeyData {
                salt: key_data_salt,
                block_size,
                key_bits,
                hash_algorithm: hash_alg,
                hash_size: digest_len,
                cipher_algorithm: "AES".to_string(),
                cipher_chaining: "ChainingModeCBC".to_string(),
            },
            data_integrity: Some(AgileDataIntegrity {
                encrypted_hmac_key,
                encrypted_hmac_value,
            }),
            password_key_encryptor: AgilePasswordKeyEncryptor {
                salt: password_salt,
                block_size,
                key_bits,
                spin_count,
                hash_algorithm: hash_alg,
                hash_size: digest_len,
                cipher_algorithm: "AES".to_string(),
                cipher_chaining: "ChainingModeCBC".to_string(),
                encrypted_verifier_hash_input: enc_vhi,
                encrypted_verifier_hash_value: enc_vhv,
                encrypted_key_value: enc_kv,
            },
        };

        let out = decrypt_agile_encrypted_package(
            &info,
            &encrypted_package,
            password,
            &crate::DecryptOptions::default(),
        )
        .expect("decrypt");
        assert!(out.is_empty());
    }

    #[test]
    fn agile_verifier_and_hmac_use_constant_time_compare() {
        // Ensure Agile password verifier + dataIntegrity HMAC comparisons route through the shared
        // `ct_eq` helper (constant-time).
        crate::util::reset_ct_eq_calls();

        let password = "Password";
        let pw_utf16 = password_to_utf16le(password);
        let hash_alg = HashAlgorithm::Sha512;
        let spin_count = 1_000u32;
        let key_bits = 256usize;
        let key_bytes = key_bits / 8;
        let block_size = 16usize;

        let salt_key_encryptor: [u8; 16] = [
            0xA0, 0xA1, 0xA2, 0xA3, 0xA4, 0xA5, 0xA6, 0xA7, 0xA8, 0xA9, 0xAA, 0xAB, 0xAC, 0xAD,
            0xAE, 0xAF,
        ];
        let salt_key_data: [u8; 16] = [
            0xB0, 0xB1, 0xB2, 0xB3, 0xB4, 0xB5, 0xB6, 0xB7, 0xB8, 0xB9, 0xBA, 0xBB, 0xBC, 0xBD,
            0xBE, 0xBF,
        ];

        // Empty ciphertext: we only need the size prefix for this test.
        let encrypted_package = 0u64.to_le_bytes().to_vec();

        let verifier_hash_input_plain: [u8; 16] = *b"formula-agl-test";
        let verifier_hash_value_plain = hash_alg.digest(&verifier_hash_input_plain);
        let package_key_plain: [u8; 32] = [0x11; 32];

        // Derive keys (block keys are used only for key derivation).
        let key_vhi = derive_agile_key(
            hash_alg,
            &salt_key_encryptor,
            &pw_utf16,
            spin_count,
            key_bytes,
            BLOCK_KEY_VERIFIER_HASH_INPUT,
        );
        let key_vhv = derive_agile_key(
            hash_alg,
            &salt_key_encryptor,
            &pw_utf16,
            spin_count,
            key_bytes,
            BLOCK_KEY_VERIFIER_HASH_VALUE,
        );
        let key_kv = derive_agile_key(
            hash_alg,
            &salt_key_encryptor,
            &pw_utf16,
            spin_count,
            key_bytes,
            BLOCK_KEY_ENCRYPTED_KEY_VALUE,
        );

        // Password verifier fields use saltValue as the IV.
        let verifier_iv = &salt_key_encryptor[..block_size];
        let enc_vhi =
            aes_cbc_encrypt(&key_vhi, verifier_iv, &verifier_hash_input_plain).expect("enc vhi");
        let enc_vhv =
            aes_cbc_encrypt(&key_vhv, verifier_iv, &verifier_hash_value_plain).expect("enc vhv");
        let enc_kv = aes_cbc_encrypt(&key_kv, verifier_iv, &package_key_plain).expect("enc key");

        // Data integrity: create a valid HMAC for this EncryptedPackage stream.
        let digest_len = hash_alg.digest_len();
        let hmac_key_plain = vec![0x22u8; digest_len];
        let computed_hmac =
            compute_hmac(hash_alg, &hmac_key_plain, &encrypted_package).expect("compute hmac");

        let iv_hmac_key = derive_iv(
            hash_alg,
            &salt_key_data,
            BLOCK_KEY_INTEGRITY_HMAC_KEY,
            block_size,
        );
        let encrypted_hmac_key = aes_cbc_encrypt(
            &package_key_plain,
            &iv_hmac_key,
            &pad_zero(&hmac_key_plain, block_size),
        )
        .expect("enc hmac key");

        let iv_hmac_val = derive_iv(
            hash_alg,
            &salt_key_data,
            BLOCK_KEY_INTEGRITY_HMAC_VALUE,
            block_size,
        );
        let encrypted_hmac_value = aes_cbc_encrypt(
            &package_key_plain,
            &iv_hmac_val,
            &pad_zero(&computed_hmac, block_size),
        )
        .expect("enc hmac value");

        let info = AgileEncryptionInfo {
            version_major: 4,
            version_minor: 4,
            flags: 0,
            key_data: AgileKeyData {
                salt: salt_key_data.to_vec(),
                block_size,
                key_bits,
                hash_algorithm: hash_alg,
                hash_size: digest_len,
                cipher_algorithm: "AES".to_string(),
                cipher_chaining: "ChainingModeCBC".to_string(),
            },
            data_integrity: Some(AgileDataIntegrity {
                encrypted_hmac_key,
                encrypted_hmac_value,
            }),
            password_key_encryptor: AgilePasswordKeyEncryptor {
                salt: salt_key_encryptor.to_vec(),
                block_size,
                key_bits,
                spin_count,
                hash_algorithm: hash_alg,
                hash_size: digest_len,
                cipher_algorithm: "AES".to_string(),
                cipher_chaining: "ChainingModeCBC".to_string(),
                encrypted_verifier_hash_input: enc_vhi,
                encrypted_verifier_hash_value: enc_vhv,
                encrypted_key_value: enc_kv,
            },
        };

        let out = decrypt_agile_encrypted_package(
            &info,
            &encrypted_package,
            password,
            &crate::DecryptOptions::default(),
        )
        .expect("decrypt");
        assert!(out.is_empty());

        let calls = crate::util::ct_eq_call_count();
        assert!(
            calls >= 2,
            "expected ct_eq to be used for both verifier and HMAC comparisons (calls={calls})"
        );
    }

    #[test]
    fn decrypt_agile_rejects_spin_count_too_large() {
        // Keep this test cheap: we only want to validate the early guard (and error surfacing),
        // not actually run an expensive password KDF in CI.
        let opts = crate::DecryptOptions { max_spin_count: 1 };
        let spin_count = 2;

        let info = AgileEncryptionInfo {
            version_major: 4,
            version_minor: 4,
            flags: 0,
            key_data: AgileKeyData {
                salt: vec![0u8; 16],
                block_size: 16,
                key_bits: 256,
                hash_algorithm: HashAlgorithm::Sha512,
                hash_size: 64,
                cipher_algorithm: "AES".to_string(),
                cipher_chaining: "ChainingModeCBC".to_string(),
            },
            data_integrity: Some(AgileDataIntegrity {
                // Not used: decryption fails before integrity checks.
                encrypted_hmac_key: Vec::new(),
                encrypted_hmac_value: Vec::new(),
            }),
            password_key_encryptor: AgilePasswordKeyEncryptor {
                salt: vec![0u8; 16],
                block_size: 16,
                key_bits: 256,
                spin_count,
                hash_algorithm: HashAlgorithm::Sha512,
                hash_size: 64,
                cipher_algorithm: "AES".to_string(),
                cipher_chaining: "ChainingModeCBC".to_string(),
                encrypted_verifier_hash_input: Vec::new(),
                encrypted_verifier_hash_value: Vec::new(),
                encrypted_key_value: Vec::new(),
            },
        };

        // Empty `EncryptedPackage` payload: 8-byte length prefix only.
        let encrypted_package = 0u64.to_le_bytes().to_vec();
        let err = decrypt_agile_encrypted_package(&info, &encrypted_package, "pw", &opts)
            .expect_err("expected spinCount limit error");

        assert!(
            matches!(
                err,
                OfficeCryptoError::SpinCountTooLarge { spin_count: got, max }
                    if got == spin_count && max == opts.max_spin_count
            ),
            "unexpected error: {err:?}"
        );
    }

    fn parsed_info() -> super::AgileEncryptionInfo {
        let info_bytes = agile_encryption_info_fixture();
        let header = parse_encryption_info_header(&info_bytes).expect("parse header");
        super::parse_agile_encryption_info(&info_bytes, &header).expect("parse agile")
    }

    #[test]
    fn decrypt_agile_rejects_u64_max_encrypted_package_size() {
        let info = parsed_info();

        // `u64::MAX` should be rejected as an absurd EncryptedPackage size before any allocation.
        let encrypted_package = u64::MAX.to_le_bytes().to_vec();

        let err = decrypt_agile_encrypted_package(
            &info,
            &encrypted_package,
            "Password",
            &crate::DecryptOptions::default(),
        )
        .expect_err("expected size limit error");
        assert!(
            matches!(
                err,
                OfficeCryptoError::SizeLimitExceededU64 {
                    context: "EncryptedPackage.originalSize",
                    limit
                } if limit == crate::MAX_ENCRYPTED_PACKAGE_ORIGINAL_SIZE
            ),
            "err={err:?}"
        );
    }

    #[test]
    fn decrypt_agile_rejects_size_larger_than_ciphertext() {
        let info = parsed_info();

        let mut encrypted_package = Vec::new();
        encrypted_package.extend_from_slice(&100u64.to_le_bytes());
        encrypted_package.extend_from_slice(&[0u8; 16]);

        let err = decrypt_agile_encrypted_package(
            &info,
            &encrypted_package,
            "Password",
            &crate::DecryptOptions::default(),
        )
        .unwrap_err();
        assert!(matches!(err, OfficeCryptoError::InvalidFormat(_)));
    }

    #[test]
    fn decrypt_rejects_oversized_encrypted_package_original_size() {
        let info = parsed_info();

        let mut encrypted_package = Vec::new();
        encrypted_package.extend_from_slice(
            &(crate::MAX_ENCRYPTED_PACKAGE_ORIGINAL_SIZE + 1).to_le_bytes(),
        );
        let err = decrypt_agile_encrypted_package(
            &info,
            &encrypted_package,
            "Password",
            &crate::DecryptOptions::default(),
        )
        .expect_err("expected size limit error");

        assert!(
            matches!(
                err,
                OfficeCryptoError::SizeLimitExceededU64 {
                    context: "EncryptedPackage.originalSize",
                    limit
                } if limit == crate::MAX_ENCRYPTED_PACKAGE_ORIGINAL_SIZE
            ),
            "err={err:?}"
        );
    }
}
