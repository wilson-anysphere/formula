//! Decryption helpers for Office-encrypted OOXML workbooks (OLE `EncryptionInfo` + `EncryptedPackage`).
//!
//! This module is behind the `encrypted-workbooks` feature because password-based decryption is
//! still landing.
// Some helpers in this module are used only by fixtures/tests while encrypted workbook support is
// being hardened. Keep `cargo check` warning-free in default builds.
#![allow(dead_code)]

use std::io;
use std::io::{Read, Seek};

use base64::engine::general_purpose::STANDARD as BASE64;
use base64::Engine as _;
use roxmltree::Document;
use zeroize::Zeroizing;

use crate::encrypted_package_reader::{DecryptedPackageReader, EncryptionMethod};

use formula_xlsx::offcrypto::{
    decrypt_aes_cbc_no_padding_in_place, derive_iv, derive_key, hash_password, CryptoError,
    HashAlgorithm, DEFAULT_MAX_SPIN_COUNT, KEY_VALUE_BLOCK, VERIFIER_HASH_INPUT_BLOCK,
    VERIFIER_HASH_VALUE_BLOCK,
};

#[derive(Debug, thiserror::Error)]
pub(crate) enum DecryptError {
    #[error("unsupported EncryptionInfo version {major}.{minor}")]
    UnsupportedVersion { major: u16, minor: u16 },
    #[error("invalid EncryptionInfo: {0}")]
    InvalidInfo(String),
    #[error("invalid password")]
    InvalidPassword,
    #[error(transparent)]
    Io(#[from] io::Error),
}

pub(crate) fn decrypted_package_reader<R: Read + Seek>(
    ciphertext_reader: R,
    plaintext_len: u64,
    encryption_info: &[u8],
    password: &str,
) -> Result<DecryptedPackageReader<R>, DecryptError> {
    if encryption_info.len() < 4 {
        return Err(DecryptError::InvalidInfo(
            "EncryptionInfo truncated (missing version header)".to_string(),
        ));
    }

    let major = u16::from_le_bytes([encryption_info[0], encryption_info[1]]);
    let minor = u16::from_le_bytes([encryption_info[2], encryption_info[3]]);

    // MS-OFFCRYPTO identifies "Standard" encryption by `versionMinor == 2`, but real-world files
    // vary the major version across Office generations (2/3/4).
    match (major, minor) {
        (4, 4) => decrypted_package_reader_agile(
            ciphertext_reader,
            plaintext_len,
            encryption_info,
            password,
        ),
        // MS-OFFCRYPTO identifies Standard (CryptoAPI) encryption via `versionMinor == 2`, but
        // real-world files vary `versionMajor` across 2/3/4 (commonly 3.2 or 4.2).
        (2 | 3 | 4, 2) => decrypted_package_reader_standard(
            ciphertext_reader,
            plaintext_len,
            encryption_info,
            password,
        ),
        _ => Err(DecryptError::UnsupportedVersion { major, minor }),
    }
}

/// Convenience helper that decrypts an `EncryptedPackage` stream fully into memory.
///
/// This is primarily used by path-based open APIs that need the decrypted ZIP bytes.
pub(crate) fn decrypt_encrypted_package(
    encryption_info: &[u8],
    encrypted_package: &[u8],
    password: &str,
) -> Result<Vec<u8>, DecryptError> {
    if encryption_info.len() < 4 {
        return Err(DecryptError::InvalidInfo(
            "EncryptionInfo truncated (missing version header)".to_string(),
        ));
    }

    let major = u16::from_le_bytes([encryption_info[0], encryption_info[1]]);
    let minor = u16::from_le_bytes([encryption_info[2], encryption_info[3]]);

    if encrypted_package.len() < 8 {
        return Err(DecryptError::InvalidInfo(
            "EncryptedPackage truncated (missing size prefix)".to_string(),
        ));
    }

    let plaintext_len = crate::parse_encrypted_package_original_size(encrypted_package).ok_or_else(|| {
        DecryptError::InvalidInfo("EncryptedPackage truncated (missing size prefix)".to_string())
    })?;

    match (major, minor) {
        (4, 4) => {
            // The ciphertext begins immediately after the 8-byte plaintext length header.
            let ciphertext = &encrypted_package[8..];
            let cursor = std::io::Cursor::new(ciphertext);
            let mut reader =
                decrypted_package_reader(cursor, plaintext_len, encryption_info, password)?;

            let mut out = Vec::new();
            if let Ok(cap) = usize::try_from(plaintext_len) {
                out.reserve(cap);
            }
            reader.read_to_end(&mut out)?;
            Ok(out)
        }
        (major, 2) if (2..=4).contains(&major) => {
            decrypt_encrypted_package_standard(encryption_info, encrypted_package, password)
        }
        _ => Err(DecryptError::UnsupportedVersion { major, minor }),
    }
}

#[derive(Debug, Clone, Copy)]
/// Candidate `EncryptedPackage` cipher/layout schemes we try for Standard/CryptoAPI AES.
///
/// Baseline MS-OFFCRYPTO/ECMA-376 Standard AES uses **AES-ECB** (no IV). The AES-CBC variants below
/// are compatibility fallbacks for some non-Excel producers.
enum StandardAesScheme {
    /// Baseline MS-OFFCRYPTO Standard AES: decrypt with AES-ECB using the block-0 key.
    Ecb,
    /// Decrypt the ciphertext as a single AES-CBC stream using:
    /// - key = block-0 key
    /// - iv  = 0
    CbcIvZeroStream,
    /// Decrypt the ciphertext as a single AES-CBC stream using:
    /// - key = block-0 key
    /// - iv  = Hash(salt || LE32(0))[:16]
    CbcIvHash0Stream,
    /// Segment the ciphertext in 4096-byte chunks; decrypt each chunk with AES-CBC using:
    /// - key = block-0 key
    /// - iv  = Hash(salt || LE32(segmentIndex))[:16]
    CbcConstKeyPerSegmentIvHash,
    /// Segment the ciphertext in 4096-byte chunks; decrypt each chunk with AES-CBC using:
    /// - key = block-0 key
    /// - iv  = 0
    CbcConstKeyPerSegmentIvZero,
    /// Segment the ciphertext in 4096-byte chunks; decrypt each chunk with AES-CBC using:
    /// - key = block-0 key
    /// - iv  = salt[:16]
    CbcConstKeyPerSegmentIvSalt,
    /// Segment the ciphertext in 4096-byte chunks; decrypt each chunk with AES-CBC using:
    /// - key = key derived for blockIndex=segmentIndex
    /// - iv  = 0
    CbcPerSegmentKeyIvZero,
    /// Decrypt the ciphertext as a single AES-CBC stream using:
    /// - key = block-0 key
    /// - iv  = salt[:16]
    CbcIvSaltStream,
}

fn decrypt_encrypted_package_standard(
    encryption_info: &[u8],
    encrypted_package: &[u8],
    password: &str,
) -> Result<Vec<u8>, DecryptError> {
    use crate::offcrypto;
    use crate::offcrypto::standard::{derive_file_key_standard, verify_password_standard_with_key};

    let info = offcrypto::parse_encryption_info_standard(encryption_info).map_err(|err| match err {
        offcrypto::OffcryptoError::UnsupportedEncryptionInfoVersion { major, minor, .. } => {
            DecryptError::UnsupportedVersion { major, minor }
        }
        other => DecryptError::InvalidInfo(format!(
            "failed to parse Standard EncryptionInfo: {other}"
        )),
    })?;

    // Validate orig_size (and avoid allocating attacker-controlled buffers).
    let plaintext_len = crate::parse_encrypted_package_original_size(encrypted_package).ok_or_else(|| {
        DecryptError::InvalidInfo("EncryptedPackage truncated (missing size prefix)".to_string())
    })?;
    let plaintext_len_usize = usize::try_from(plaintext_len).map_err(|_| {
        DecryptError::InvalidInfo(format!(
            "EncryptedPackage orig_size {plaintext_len} does not fit into usize"
        ))
    })?;

    match info.header.alg_id {
        offcrypto::CALG_RC4 => {
            let key0 = derive_file_key_standard(&info, password).map_err(|err| {
                DecryptError::InvalidInfo(format!("failed to derive Standard RC4 key: {err}"))
            })?;

            // Verify password using the EncryptionVerifier before decrypting the package.
            let ok = verify_password_standard_with_key(&info, &key0).map_err(|err| {
                DecryptError::InvalidInfo(format!("failed to verify Standard password: {err}"))
            })?;
            if !ok {
                return Err(DecryptError::InvalidPassword);
            }

            let key_len = (info.header.key_size / 8) as usize;
            let out = offcrypto::decrypt_standard_cryptoapi_rc4_encrypted_package_stream(
                encrypted_package,
                password,
                &info.verifier.salt,
                key_len,
            )
            .map_err(|err| {
                DecryptError::InvalidInfo(format!(
                    "failed to decrypt Standard RC4 EncryptedPackage: {err}"
                ))
            })?;
            Ok(out)
        }
        offcrypto::CALG_AES_128 | offcrypto::CALG_AES_192 | offcrypto::CALG_AES_256 => {
            // Try a small set of schemes seen in the wild.
            //
            // Some producers appear to derive the AES-128 key by truncating the block hash instead
            // of running the CryptoAPI `CryptDeriveKey` ipad/opad expansion. For robustness, we try
            // both key derivation variants when keySize=128.
            //
            // Additionally, some schemes share the same initial IV (so the first decrypted block can
            // look like a ZIP even when the overall scheme is wrong). Therefore, when we get a
            // promising `PK` prefix we validate by parsing the full decrypted bytes as a ZIP archive.
            let schemes = [
                StandardAesScheme::Ecb,
                StandardAesScheme::CbcIvZeroStream,
                StandardAesScheme::CbcIvSaltStream,
                StandardAesScheme::CbcIvHash0Stream,
                StandardAesScheme::CbcConstKeyPerSegmentIvHash,
                StandardAesScheme::CbcConstKeyPerSegmentIvZero,
                StandardAesScheme::CbcConstKeyPerSegmentIvSalt,
                StandardAesScheme::CbcPerSegmentKeyIvZero,
            ];

            let ciphertext = &encrypted_package[8..];

            let try_with_key0 = |key0: &[u8]| -> Result<Option<Vec<u8>>, DecryptError> {
                for scheme in schemes {
                    let prefix_ok = decrypt_standard_aes_prefix_is_zip(
                        &info,
                        plaintext_len_usize,
                        ciphertext,
                        scheme,
                        key0,
                    )?;
                    if !prefix_ok {
                        continue;
                    }

                    let Ok(out) = decrypt_standard_aes_with_scheme(
                        &info,
                        password,
                        plaintext_len_usize,
                        ciphertext,
                        scheme,
                        key0,
                    ) else {
                        continue;
                    };

                    if is_valid_zip(&out) {
                        return Ok(Some(out));
                    }
                }
                Ok(None)
            };

            let key0 = derive_file_key_standard(&info, password).map_err(|err| {
                DecryptError::InvalidInfo(format!("failed to derive Standard key: {err}"))
            })?;
            if let Some(out) = try_with_key0(&key0)? {
                return Ok(out);
            }

            // AES-128 truncation fallback.
            if info.header.key_size == 128 {
                if let Ok(key_trunc) = derive_standard_aes_key_truncate(&info, password) {
                    if key_trunc != key0 {
                        if let Some(out) = try_with_key0(&key_trunc)? {
                            return Ok(out);
                        }
                    }
                }
            }

            Err(DecryptError::InvalidPassword)
        }
        other => Err(DecryptError::InvalidInfo(format!(
            "unsupported Standard CryptoAPI algorithm algId=0x{other:08x}"
        ))),
    }
}

fn decrypt_standard_aes_with_scheme(
    info: &crate::offcrypto::StandardEncryptionInfo,
    password: &str,
    plaintext_len: usize,
    ciphertext: &[u8],
    scheme: StandardAesScheme,
    key0: &[u8],
) -> Result<Vec<u8>, DecryptError> {
    use crate::offcrypto::standard::derive_key_standard_for_block;

    const AES_BLOCK: usize = 16;
    const SEGMENT_LEN: usize = 0x1000;

    let needed_cipher_len = round_up_to_multiple(plaintext_len, AES_BLOCK);

    match scheme {
        StandardAesScheme::Ecb => {
            if ciphertext.len() < needed_cipher_len {
                return Err(DecryptError::InvalidInfo(format!(
                    "EncryptedPackage ciphertext truncated: have {}, need at least {}",
                    ciphertext.len(),
                    needed_cipher_len
                )));
            }
            let mut buf = ciphertext[..needed_cipher_len].to_vec();
            aes_ecb_decrypt_in_place(key0, &mut buf)
                .map_err(|msg| DecryptError::InvalidInfo(msg.to_string()))?;
            buf.truncate(plaintext_len);
            Ok(buf)
        }
        StandardAesScheme::CbcIvZeroStream => {
            if ciphertext.len() < needed_cipher_len {
                return Err(DecryptError::InvalidInfo(format!(
                    "EncryptedPackage ciphertext truncated: have {}, need at least {}",
                    ciphertext.len(),
                    needed_cipher_len
                )));
            }

            let iv = [0u8; AES_BLOCK];
            let mut buf = ciphertext[..needed_cipher_len].to_vec();
            decrypt_aes_cbc_no_padding_in_place(key0, &iv, &mut buf).map_err(|e| {
                DecryptError::InvalidInfo(format!("AES-CBC decrypt failed: {e}"))
            })?;
            buf.truncate(plaintext_len);
            Ok(buf)
        }
        StandardAesScheme::CbcIvHash0Stream => {
            if ciphertext.len() < needed_cipher_len {
                return Err(DecryptError::InvalidInfo(format!(
                    "EncryptedPackage ciphertext truncated: have {}, need at least {}",
                    ciphertext.len(),
                    needed_cipher_len
                )));
            }

            let iv = derive_standard_segment_iv(info.header.alg_id_hash, &info.verifier.salt, 0)?;
            let mut buf = ciphertext[..needed_cipher_len].to_vec();
            decrypt_aes_cbc_no_padding_in_place(key0, &iv, &mut buf).map_err(|e| {
                DecryptError::InvalidInfo(format!("AES-CBC decrypt failed: {e}"))
            })?;
            buf.truncate(plaintext_len);
            Ok(buf)
        }
        StandardAesScheme::CbcIvSaltStream => {
            if ciphertext.len() < needed_cipher_len {
                return Err(DecryptError::InvalidInfo(format!(
                    "EncryptedPackage ciphertext truncated: have {}, need at least {}",
                    ciphertext.len(),
                    needed_cipher_len
                )));
            }

            let iv = info.verifier.salt.get(..AES_BLOCK).ok_or_else(|| {
                DecryptError::InvalidInfo("EncryptionVerifier.salt shorter than 16 bytes".to_string())
            })?;
            let iv: [u8; AES_BLOCK] = iv
                .try_into()
                .expect("slice length checked to be 16 bytes");

            let mut buf = ciphertext[..needed_cipher_len].to_vec();
            decrypt_aes_cbc_no_padding_in_place(key0, &iv, &mut buf).map_err(|e| {
                DecryptError::InvalidInfo(format!("AES-CBC decrypt failed: {e}"))
            })?;
            buf.truncate(plaintext_len);
            Ok(buf)
        }
        StandardAesScheme::CbcConstKeyPerSegmentIvHash => {
            decrypt_segmented_aes_cbc(
                ciphertext,
                plaintext_len,
                |_segment_index| Ok(key0.to_vec()),
                |segment_index| derive_standard_segment_iv(info.header.alg_id_hash, &info.verifier.salt, segment_index),
                SEGMENT_LEN,
            )
        }
        StandardAesScheme::CbcConstKeyPerSegmentIvZero => decrypt_segmented_aes_cbc(
            ciphertext,
            plaintext_len,
            |_segment_index| Ok(key0.to_vec()),
            |_segment_index| Ok([0u8; AES_BLOCK]),
            SEGMENT_LEN,
        ),
        StandardAesScheme::CbcConstKeyPerSegmentIvSalt => {
            let iv = info.verifier.salt.get(..AES_BLOCK).ok_or_else(|| {
                DecryptError::InvalidInfo("EncryptionVerifier.salt shorter than 16 bytes".to_string())
            })?;
            let iv: [u8; AES_BLOCK] = iv
                .try_into()
                .expect("slice length checked to be 16 bytes");

            decrypt_segmented_aes_cbc(
                ciphertext,
                plaintext_len,
                |_segment_index| Ok(key0.to_vec()),
                |_segment_index| Ok(iv),
                SEGMENT_LEN,
            )
        }
        StandardAesScheme::CbcPerSegmentKeyIvZero => decrypt_segmented_aes_cbc(
            ciphertext,
            plaintext_len,
            |segment_index| {
                if segment_index == 0 {
                    return Ok(key0.to_vec());
                }

                let key = derive_key_standard_for_block(info, password, segment_index).map_err(|err| {
                    DecryptError::InvalidInfo(format!(
                        "failed to derive Standard per-segment key (segment_index={segment_index}): {err}"
                    ))
                })?;
                Ok(key)
            },
            |_segment_index| Ok([0u8; AES_BLOCK]),
            SEGMENT_LEN,
        ),
    }
}

fn decrypt_standard_aes_prefix_is_zip(
    info: &crate::offcrypto::StandardEncryptionInfo,
    plaintext_len: usize,
    ciphertext: &[u8],
    scheme: StandardAesScheme,
    key0: &[u8],
) -> Result<bool, DecryptError> {
    const AES_BLOCK: usize = 16;
    if plaintext_len == 0 {
        return Ok(false);
    }
    let first = ciphertext.get(..AES_BLOCK).ok_or_else(|| {
        DecryptError::InvalidInfo("EncryptedPackage ciphertext truncated (missing first AES block)".to_string())
    })?;
    let mut buf = first.to_vec();

    match scheme {
        StandardAesScheme::Ecb => {
            aes_ecb_decrypt_in_place(key0, &mut buf)
                .map_err(|msg| DecryptError::InvalidInfo(msg.to_string()))?;
        }
        StandardAesScheme::CbcIvZeroStream
        | StandardAesScheme::CbcConstKeyPerSegmentIvZero
        | StandardAesScheme::CbcPerSegmentKeyIvZero => {
            let iv = [0u8; AES_BLOCK];
            decrypt_aes_cbc_no_padding_in_place(key0, &iv, &mut buf).map_err(|e| {
                DecryptError::InvalidInfo(format!("AES-CBC decrypt failed: {e}"))
            })?;
        }
        StandardAesScheme::CbcIvHash0Stream | StandardAesScheme::CbcConstKeyPerSegmentIvHash => {
            let iv = derive_standard_segment_iv(info.header.alg_id_hash, &info.verifier.salt, 0)?;
            decrypt_aes_cbc_no_padding_in_place(key0, &iv, &mut buf).map_err(|e| {
                DecryptError::InvalidInfo(format!("AES-CBC decrypt failed: {e}"))
            })?;
        }
        StandardAesScheme::CbcIvSaltStream | StandardAesScheme::CbcConstKeyPerSegmentIvSalt => {
            let iv = info.verifier.salt.get(..AES_BLOCK).ok_or_else(|| {
                DecryptError::InvalidInfo("EncryptionVerifier.salt shorter than 16 bytes".to_string())
            })?;
            let iv: [u8; AES_BLOCK] = iv
                .try_into()
                .expect("slice length checked to be 16 bytes");
            decrypt_aes_cbc_no_padding_in_place(key0, &iv, &mut buf).map_err(|e| {
                DecryptError::InvalidInfo(format!("AES-CBC decrypt failed: {e}"))
            })?;
        }
    }

    Ok(buf.len() >= 2 && &buf[..2] == b"PK")
}

fn derive_standard_aes_key_truncate(
    info: &crate::offcrypto::StandardEncryptionInfo,
    password: &str,
) -> Result<Vec<u8>, DecryptError> {
    use crate::offcrypto::cryptoapi::{
        final_hash, hash_password_fixed_spin, password_to_utf16le, HashAlg,
    };

    let hash_alg = HashAlg::from_calg_id(info.header.alg_id_hash).map_err(|err| {
        DecryptError::InvalidInfo(format!(
            "unsupported Standard algIdHash=0x{:08x}: {err}",
            info.header.alg_id_hash
        ))
    })?;

    let pw_utf16le = password_to_utf16le(password);
    let h_final = hash_password_fixed_spin(&pw_utf16le, &info.verifier.salt, hash_alg);
    let h_block0 = final_hash(&h_final, 0, hash_alg);

    let key_len = (info.header.key_size / 8) as usize;
    if key_len > h_block0.len() {
        return Err(DecryptError::InvalidInfo(format!(
            "invalid keySize {} bits for truncation-based Standard AES derivation: key_len={key_len} > digest_len={}",
            info.header.key_size,
            h_block0.len()
        )));
    }
    Ok(h_block0[..key_len].to_vec())
}

fn is_valid_zip(bytes: &[u8]) -> bool {
    if bytes.len() < 2 || &bytes[..2] != b"PK" {
        return false;
    }
    let cursor = std::io::Cursor::new(bytes);
    zip::ZipArchive::new(cursor).is_ok()
}

fn decrypt_segmented_aes_cbc(
    ciphertext: &[u8],
    plaintext_len: usize,
    mut key_for_segment: impl FnMut(u32) -> Result<Vec<u8>, DecryptError>,
    mut iv_for_segment: impl FnMut(u32) -> Result<[u8; 16], DecryptError>,
    segment_len: usize,
) -> Result<Vec<u8>, DecryptError> {
    const AES_BLOCK: usize = 16;

    let mut out = Vec::new();
    out.reserve(plaintext_len);

    let mut segment_index: u32 = 0;
    while out.len() < plaintext_len {
        let remaining = plaintext_len - out.len();
        let seg_plain_len = remaining.min(segment_len);
        let seg_cipher_len = round_up_to_multiple(seg_plain_len, AES_BLOCK);

        let seg_start = (segment_index as usize)
            .checked_mul(segment_len)
            .ok_or_else(|| DecryptError::InvalidInfo("segment index overflow".to_string()))?;
        let seg_end = seg_start
            .checked_add(seg_cipher_len)
            .ok_or_else(|| DecryptError::InvalidInfo("segment length overflow".to_string()))?;
        let seg_cipher = ciphertext.get(seg_start..seg_end).ok_or_else(|| {
            DecryptError::InvalidInfo(format!(
                "EncryptedPackage ciphertext truncated for segment {segment_index}: need bytes {seg_start}..{seg_end} (len={})",
                ciphertext.len()
            ))
        })?;

        let key = key_for_segment(segment_index)?;
        let iv = iv_for_segment(segment_index)?;

        let mut buf = seg_cipher.to_vec();
        decrypt_aes_cbc_no_padding_in_place(&key, &iv, &mut buf).map_err(|e| {
            DecryptError::InvalidInfo(format!("AES-CBC decrypt failed for segment {segment_index}: {e}"))
        })?;
        out.extend_from_slice(&buf[..seg_plain_len]);

        segment_index = segment_index
            .checked_add(1)
            .ok_or_else(|| DecryptError::InvalidInfo("segment index overflow".to_string()))?;
    }

    Ok(out)
}

fn derive_standard_segment_iv(
    alg_id_hash: u32,
    salt: &[u8],
    segment_index: u32,
) -> Result<[u8; 16], DecryptError> {
    // Segment IV derivation used by some non-standard AES-CBC Standard/CryptoAPI `EncryptedPackage`
    // layouts:
    //
    //   iv_i = Hash(salt || LE32(i))[0..16]
    //
    // Note: this is not used by the baseline Standard AES-ECB scheme.
    let mut iv = [0u8; 16];
    match alg_id_hash {
        crate::offcrypto::CALG_SHA1 => {
            use sha1::Digest as _;
            let mut hasher = sha1::Sha1::new();
            hasher.update(salt);
            hasher.update(segment_index.to_le_bytes());
            let digest = hasher.finalize();
            iv.copy_from_slice(&digest[..16]); // SHA-1 digest is 20 bytes
            Ok(iv)
        }
        crate::offcrypto::CALG_MD5 => {
            use md5::Digest as _;
            let mut hasher = md5::Md5::new();
            hasher.update(salt);
            hasher.update(segment_index.to_le_bytes());
            let digest = hasher.finalize();
            iv.copy_from_slice(&digest[..16]); // MD5 digest is 16 bytes
            Ok(iv)
        }
        other => Err(DecryptError::InvalidInfo(format!(
            "unsupported Standard algIdHash=0x{other:08x} for AES-CBC compatibility IV derivation"
        ))),
    }
}

fn aes_ecb_decrypt_in_place(key: &[u8], buf: &mut [u8]) -> Result<(), &'static str> {
    use aes::cipher::{generic_array::GenericArray, BlockDecrypt, KeyInit};
    use aes::{Aes128, Aes192, Aes256};

    const AES_BLOCK: usize = 16;
    if buf.len() % AES_BLOCK != 0 {
        return Err("AES-ECB ciphertext length is not a multiple of 16 bytes");
    }

    fn decrypt_with<C>(key: &[u8], buf: &mut [u8]) -> Result<(), &'static str>
    where
        C: BlockDecrypt + KeyInit,
    {
        let cipher = C::new_from_slice(key).map_err(|_| "invalid AES key length")?;
        for block in buf.chunks_mut(AES_BLOCK) {
            cipher.decrypt_block(GenericArray::from_mut_slice(block));
        }
        Ok(())
    }

    match key.len() {
        16 => decrypt_with::<Aes128>(key, buf),
        24 => decrypt_with::<Aes192>(key, buf),
        32 => decrypt_with::<Aes256>(key, buf),
        _ => Err("invalid AES key length"),
    }
}

fn round_up_to_multiple(value: usize, multiple: usize) -> usize {
    if multiple == 0 {
        return value;
    }
    let rem = value % multiple;
    if rem == 0 {
        value
    } else {
        value + (multiple - rem)
    }
}

fn decrypted_package_reader_standard<R: Read + Seek>(
    mut ciphertext_reader: R,
    plaintext_len: u64,
    encryption_info: &[u8],
    password: &str,
) -> Result<DecryptedPackageReader<R>, DecryptError> {
    use crate::offcrypto;
    use crate::offcrypto::standard::{derive_file_key_standard, verify_password_standard_with_key};

    let info = offcrypto::parse_encryption_info_standard(encryption_info).map_err(|err| match err {
        offcrypto::OffcryptoError::UnsupportedEncryptionInfoVersion { major, minor, .. } => {
            DecryptError::UnsupportedVersion { major, minor }
        }
        other => DecryptError::InvalidInfo(format!(
            "failed to parse Standard EncryptionInfo: {other}"
        )),
    })?;

    // The streaming decryptor supports Standard/CryptoAPI AES `EncryptedPackage` decryption via:
    // - baseline AES-ECB (no IV)
    // - a non-standard AES-CBC-per-segment mode observed in some producers/fixtures
    //
    // (It does not support the RC4 variant in this path.)
    match info.header.alg_id {
        offcrypto::CALG_AES_128 | offcrypto::CALG_AES_192 | offcrypto::CALG_AES_256 => {}
        other => {
            return Err(DecryptError::InvalidInfo(format!(
                "unsupported Standard CryptoAPI algorithm algId=0x{other:08x}"
            )))
        }
    }

    let key = Zeroizing::new(derive_file_key_standard(&info, password).map_err(|err| {
        DecryptError::InvalidInfo(format!("failed to derive Standard key: {err}"))
    })?);

    let ok = verify_password_standard_with_key(&info, key.as_slice()).map_err(|err| {
        DecryptError::InvalidInfo(format!("failed to verify Standard password: {err}"))
    })?;
    if !ok {
        return Err(DecryptError::InvalidPassword);
    }

    if plaintext_len == 0 {
        return Ok(DecryptedPackageReader::new(
            ciphertext_reader,
            EncryptionMethod::StandardAesEcb { key },
            plaintext_len,
        ));
    }

    fn derive_standard_segment_iv(salt: &[u8], segment_index: u32) -> [u8; 16] {
        use sha1::{Digest as _, Sha1};
        let mut hasher = Sha1::new();
        hasher.update(salt);
        hasher.update(segment_index.to_le_bytes());
        let digest = hasher.finalize();
        let mut iv = [0u8; 16];
        iv.copy_from_slice(&digest[..16]);
        iv
    }

    // Detect whether the Standard `EncryptedPackage` payload uses AES-ECB or a CBC-per-segment
    // framing by decrypting the first ciphertext block and checking for the `PK` ZIP signature.
    let cipher_pos = ciphertext_reader.seek(io::SeekFrom::Current(0))?;
    let mut first_block = [0u8; 16];
    ciphertext_reader.read_exact(&mut first_block)?;
    ciphertext_reader.seek(io::SeekFrom::Start(cipher_pos))?;

    let salt = info.verifier.salt.clone();

    let ecb_ok = {
        let mut block = first_block;
        aes_ecb_decrypt_in_place(&key, &mut block).is_ok() && block.starts_with(b"PK")
    };

    let cbc_ok = {
        let mut block = first_block;
        let iv = derive_standard_segment_iv(&salt, 0);
        decrypt_aes_cbc_no_padding_in_place(&key, &iv, &mut block).is_ok()
            && block.starts_with(b"PK")
    };

    let method = if ecb_ok {
        EncryptionMethod::StandardAesEcb { key }
    } else if cbc_ok {
        EncryptionMethod::StandardCryptoApi { key, salt }
    } else {
        return Err(DecryptError::InvalidInfo(
            "unable to detect Standard EncryptedPackage cipher mode (expected ZIP magic)".into(),
        ));
    };

    Ok(DecryptedPackageReader::new(ciphertext_reader, method, plaintext_len))
}

#[derive(Debug, Clone)]
struct AgileKeyData {
    salt_value: Vec<u8>,
    hash_algorithm: HashAlgorithm,
    block_size: usize,
    key_bits: usize,
}

#[derive(Debug, Clone)]
struct AgilePasswordKeyEncryptor {
    salt_value: Vec<u8>,
    hash_algorithm: HashAlgorithm,
    spin_count: u32,
    block_size: usize,
    key_bits: usize,
    hash_size: usize,
    encrypted_verifier_hash_input: Vec<u8>,
    encrypted_verifier_hash_value: Vec<u8>,
    encrypted_key_value: Vec<u8>,
}

#[derive(Debug, Clone)]
struct AgileEncryptionInfo {
    key_data: AgileKeyData,
    password_key: AgilePasswordKeyEncryptor,
}

fn decrypted_package_reader_agile<R: Read + Seek>(
    ciphertext_reader: R,
    plaintext_len: u64,
    encryption_info: &[u8],
    password: &str,
) -> Result<DecryptedPackageReader<R>, DecryptError> {
    let xml = crate::extract_agile_encryption_info_xml(encryption_info)
        .map_err(|err| DecryptError::InvalidInfo(err.to_string()))?;
    let info = parse_agile_encryption_info(&xml)?;

    let key = agile_decrypt_package_key(password, &info)?;

    Ok(DecryptedPackageReader::new(
        ciphertext_reader,
        EncryptionMethod::Agile {
            key,
            salt: info.key_data.salt_value,
            hash_alg: info.key_data.hash_algorithm,
            block_size: info.key_data.block_size,
        },
        plaintext_len,
    ))
}

fn parse_agile_encryption_info(xml: &str) -> Result<AgileEncryptionInfo, DecryptError> {
    let doc = Document::parse(xml)
        .map_err(|err| DecryptError::InvalidInfo(format!("EncryptionInfo XML parse: {err}")))?;

    let key_data_node = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "keyData")
        .ok_or_else(|| DecryptError::InvalidInfo("missing keyData element".into()))?;

    validate_cipher_settings(key_data_node)?;

    let key_data = AgileKeyData {
        salt_value: parse_base64_attr(key_data_node, "saltValue")?,
        hash_algorithm: parse_hash_algorithm(key_data_node, "hashAlgorithm")?,
        block_size: parse_usize_attr(key_data_node, "blockSize")?,
        key_bits: parse_usize_attr(key_data_node, "keyBits")?,
    };

    let key_encryptor_node = doc
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "keyEncryptor"
                && n.attribute("uri")
                    .is_some_and(|u| u.to_ascii_lowercase().contains("password"))
        })
        .ok_or_else(|| DecryptError::InvalidInfo("missing keyEncryptor (password)".into()))?;

    let encrypted_key_node = key_encryptor_node
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "encryptedKey")
        .ok_or_else(|| DecryptError::InvalidInfo("missing encryptedKey element".into()))?;

    validate_cipher_settings(encrypted_key_node)?;

    let spin_count = parse_u32_attr(encrypted_key_node, "spinCount")?;
    if spin_count > DEFAULT_MAX_SPIN_COUNT {
        return Err(DecryptError::InvalidInfo(format!(
            "spinCount {spin_count} exceeds maximum allowed {DEFAULT_MAX_SPIN_COUNT} (refusing to run expensive password KDF)"
        )));
    }

    let password_key = AgilePasswordKeyEncryptor {
        salt_value: parse_base64_attr(encrypted_key_node, "saltValue")?,
        hash_algorithm: parse_hash_algorithm(encrypted_key_node, "hashAlgorithm")?,
        spin_count,
        block_size: parse_usize_attr(encrypted_key_node, "blockSize")?,
        key_bits: parse_usize_attr(encrypted_key_node, "keyBits")?,
        hash_size: parse_usize_attr(encrypted_key_node, "hashSize")?,
        encrypted_verifier_hash_input: parse_base64_attr(
            encrypted_key_node,
            "encryptedVerifierHashInput",
        )?,
        encrypted_verifier_hash_value: parse_base64_attr(
            encrypted_key_node,
            "encryptedVerifierHashValue",
        )?,
        encrypted_key_value: parse_base64_attr(encrypted_key_node, "encryptedKeyValue")?,
    };

    Ok(AgileEncryptionInfo {
        key_data,
        password_key,
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::offcrypto::standard::EncryptionHeaderFlags;
    use crate::offcrypto::{EncryptionHeader, EncryptionVerifier, StandardEncryptionInfo, CALG_AES_128, CALG_SHA1};

    #[test]
    fn rejects_spin_count_above_default_max() {
        let xml = r#"
            <encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
                        xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
              <keyData saltValue="AA==" hashAlgorithm="SHA1" hashSize="20"
                       cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                       keyBits="128" blockSize="16" />
              <keyEncryptors>
                <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
                  <p:encryptedKey saltValue="AA==" spinCount="4294967295" hashAlgorithm="SHA1" hashSize="20"
                                  cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                                  keyBits="128" blockSize="16"
                                  encryptedVerifierHashInput="AA=="
                                  encryptedVerifierHashValue="AA=="
                                  encryptedKeyValue="AA=="/>
                </keyEncryptor>
              </keyEncryptors>
            </encryption>
        "#;

        let err = parse_agile_encryption_info(xml).expect_err("expected error");
        assert!(
            matches!(err, DecryptError::InvalidInfo(ref msg) if msg.contains("spinCount") && msg.contains("maximum")),
            "unexpected error: {err:?}"
        );
    }

    #[test]
    fn derive_standard_aes_key_truncate_uses_hblock0_prefix_for_known_vector() {
        // Some producers in the wild appear to derive the Standard AES-128 key by truncating the
        // per-block hash value, instead of running the CryptoAPI `CryptDeriveKey` ipad/opad
        // expansion. `decrypt_standard_encrypted_ooxml_package` keeps a compatibility fallback for
        // this behavior.
        //
        // This test uses the vector from `docs/offcrypto-standard-cryptoapi.md` (§8.2) and asserts
        // that the truncation variant yields `H_block0[0:16]`.
        let salt: [u8; 16] = [
            0xE8, 0x82, 0x66, 0x49, 0x0C, 0x5B, 0xD1, 0xEE, 0xBD, 0x2B, 0x43, 0x94, 0xE3, 0xF8,
            0x30, 0xEF,
        ];

        let info = StandardEncryptionInfo {
            header: EncryptionHeader {
                flags: EncryptionHeaderFlags::from_raw(
                    EncryptionHeaderFlags::F_CRYPTOAPI | EncryptionHeaderFlags::F_AES,
                ),
                size_extra: 0,
                alg_id: CALG_AES_128,
                alg_id_hash: CALG_SHA1,
                key_size: 128,
                provider_type: 0,
                reserved1: 0,
                reserved2: 0,
                csp_name: String::new(),
            },
            verifier: EncryptionVerifier {
                salt: salt.to_vec(),
                encrypted_verifier: [0u8; 16],
                verifier_hash_size: 20,
                encrypted_verifier_hash: Vec::new(),
            },
        };

        let key = derive_standard_aes_key_truncate(&info, "Password1234_").unwrap();
        let expected: [u8; 16] = [
            0xE2, 0xF8, 0xCD, 0xE4, 0x57, 0xE5, 0xD4, 0x49, 0xEB, 0x20, 0x50, 0x57, 0xC8,
            0x8D, 0x20, 0x1D,
        ];
        assert_eq!(key, expected);
    }

    #[test]
    #[cfg(not(target_arch = "wasm32"))]
    fn standard_fixtures_decrypt_to_exact_plaintext_via_streaming_reader() {
        use std::io::Read as _;

        let fixture_dir = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../fixtures/encrypted/ooxml");

        let cases = [
            ("standard.xlsx", "plaintext.xlsx", "password"),
            ("standard-4.2.xlsx", "plaintext.xlsx", "password"),
            ("standard-unicode.xlsx", "plaintext.xlsx", "pässwörd🔒"),
            ("standard-large.xlsx", "plaintext-large.xlsx", "password"),
            ("standard-basic.xlsm", "plaintext-basic.xlsm", "password"),
        ];

        for (encrypted_name, plaintext_name, password) in cases {
            let encrypted_path = fixture_dir.join(encrypted_name);
            let plaintext_path = fixture_dir.join(plaintext_name);

            let file =
                std::fs::File::open(&encrypted_path).expect("open encrypted OOXML fixture");
            let mut ole = cfb::CompoundFile::open(file).expect("parse OLE");

            let mut encryption_info = Vec::new();
            ole.open_stream("EncryptionInfo")
                .or_else(|_| ole.open_stream("/EncryptionInfo"))
                .expect("open EncryptionInfo")
                .read_to_end(&mut encryption_info)
                .expect("read EncryptionInfo");

            let mut encrypted_package = Vec::new();
            ole.open_stream("EncryptedPackage")
                .or_else(|_| ole.open_stream("/EncryptedPackage"))
                .expect("open EncryptedPackage")
                .read_to_end(&mut encrypted_package)
                .expect("read EncryptedPackage");

            assert!(
                encrypted_package.len() >= 8,
                "EncryptedPackage too short (missing size prefix) for {encrypted_name}"
            );
            let plaintext_len = u64::from_le_bytes(
                encrypted_package[..8]
                    .try_into()
                    .expect("EncryptedPackage size prefix"),
            );
            let ciphertext = encrypted_package[8..].to_vec();

            let expected = std::fs::read(&plaintext_path).expect("read plaintext fixture");
            assert_eq!(
                plaintext_len,
                expected.len() as u64,
                "{encrypted_name}: EncryptedPackage orig_size does not match {plaintext_name} length"
            );

            // Sequential read path.
            let mut reader = decrypted_package_reader(
                std::io::Cursor::new(ciphertext.clone()),
                plaintext_len,
                &encryption_info,
                password,
            )
            .unwrap_or_else(|err| panic!("create streaming decrypt reader for {encrypted_name}: {err:?}"));
            let mut out = Vec::new();
            reader
                .read_to_end(&mut out)
                .unwrap_or_else(|err| panic!("read decrypted bytes for {encrypted_name}: {err}"));
            assert_eq!(out, expected, "{encrypted_name}: decrypted bytes mismatch");

            // Random-access path: ensure we can read ZIP metadata via Seek (central directory).
            let reader = decrypted_package_reader(
                std::io::Cursor::new(ciphertext),
                plaintext_len,
                &encryption_info,
                password,
            )
            .unwrap_or_else(|err| panic!("create streaming decrypt reader for {encrypted_name}: {err:?}"));
            let mut zip = zip::ZipArchive::new(reader)
                .unwrap_or_else(|err| panic!("open decrypted ZIP for {encrypted_name}: {err}"));
            let mut part = zip
                .by_name("[Content_Types].xml")
                .unwrap_or_else(|err| panic!("read [Content_Types].xml for {encrypted_name}: {err}"));
            let mut xml = String::new();
            part.read_to_string(&mut xml)
                .unwrap_or_else(|err| panic!("read [Content_Types].xml bytes for {encrypted_name}: {err}"));
            assert!(
                xml.contains("<Types"),
                "{encrypted_name}: expected [Content_Types].xml to contain <Types, got: {xml:?}"
            );
        }
    }
}

fn validate_cipher_settings(node: roxmltree::Node<'_, '_>) -> Result<(), DecryptError> {
    let cipher_alg = required_attr(node, "cipherAlgorithm")?;
    if !cipher_alg.eq_ignore_ascii_case("AES") {
        return Err(DecryptError::InvalidInfo(format!(
            "unsupported cipherAlgorithm {cipher_alg}"
        )));
    }
    let chaining = required_attr(node, "cipherChaining")?;
    if !chaining.eq_ignore_ascii_case("ChainingModeCBC") {
        return Err(DecryptError::InvalidInfo(format!(
            "unsupported cipherChaining {chaining}"
        )));
    }
    Ok(())
}

#[derive(Debug, Clone, Copy)]
enum PasswordKeyIvMode {
    Salt,
    Derived,
}

fn password_key_iv(
    key: &AgilePasswordKeyEncryptor,
    mode: PasswordKeyIvMode,
    block_key: &[u8],
) -> Result<Vec<u8>, DecryptError> {
    match mode {
        PasswordKeyIvMode::Salt => key
            .salt_value
            .get(..key.block_size)
            .ok_or_else(|| {
                DecryptError::InvalidInfo("encryptedKey.saltValue shorter than blockSize".into())
            })
            .map(|iv| iv.to_vec()),
        PasswordKeyIvMode::Derived => derive_iv(
            &key.salt_value,
            block_key,
            key.block_size,
            key.hash_algorithm,
        )
        .map_err(map_crypto_err("derive_iv")),
    }
}

fn agile_decrypt_package_key(
    password: &str,
    info: &AgileEncryptionInfo,
) -> Result<Zeroizing<Vec<u8>>, DecryptError> {
    let password_key = &info.password_key;

    let password_hash = Zeroizing::new(
        hash_password(
            password,
            &password_key.salt_value,
            password_key.spin_count,
            password_key.hash_algorithm,
        )
        .map_err(|e| DecryptError::InvalidInfo(format!("hash_password: {e}")))?,
    );

    let key_encrypt_key_len = key_len_bytes(password_key.key_bits, "encryptedKey", "keyBits")?;
    let package_key_len = key_len_bytes(info.key_data.key_bits, "keyData", "keyBits")?;

    if password_key.block_size != formula_xlsx::offcrypto::AES_BLOCK_SIZE {
        return Err(DecryptError::InvalidInfo(format!(
            "unsupported encryptedKey.blockSize {} (expected {})",
            password_key.block_size,
            formula_xlsx::offcrypto::AES_BLOCK_SIZE
        )));
    }

    fn try_unwrap_key(
        info: &AgileEncryptionInfo,
        password_hash: &[u8],
        key_encrypt_key_len: usize,
        package_key_len: usize,
        mode: PasswordKeyIvMode,
    ) -> Result<Zeroizing<Vec<u8>>, DecryptError> {
        let password_key = &info.password_key;

        let verifier_input_iv = password_key_iv(password_key, mode, &VERIFIER_HASH_INPUT_BLOCK)?;
        let verifier_value_iv = password_key_iv(password_key, mode, &VERIFIER_HASH_VALUE_BLOCK)?;
        let key_value_iv = password_key_iv(password_key, mode, &KEY_VALUE_BLOCK)?;

        let verifier_input = {
            let k = Zeroizing::new(
                derive_key(
                    password_hash,
                    &VERIFIER_HASH_INPUT_BLOCK,
                    key_encrypt_key_len,
                    password_key.hash_algorithm,
                )
                .map_err(map_crypto_err("derive_key(verifierHashInput)"))?,
            );
            let mut decrypted = Zeroizing::new(password_key.encrypted_verifier_hash_input.clone());
            decrypt_aes_cbc_no_padding_in_place(k.as_slice(), &verifier_input_iv, &mut decrypted)
                .map_err(|e| DecryptError::InvalidInfo(format!("decrypt verifierHashInput: {e}")))?;
            if decrypted.len() < password_key.block_size {
                return Err(DecryptError::InvalidInfo(
                    "decrypted verifierHashInput shorter than blockSize".into(),
                ));
            }
            decrypted.truncate(password_key.block_size);
            decrypted
        };

        let verifier_hash = {
            let k = Zeroizing::new(
                derive_key(
                    password_hash,
                    &VERIFIER_HASH_VALUE_BLOCK,
                    key_encrypt_key_len,
                    password_key.hash_algorithm,
                )
                .map_err(map_crypto_err("derive_key(verifierHashValue)"))?,
            );
            let mut decrypted = Zeroizing::new(password_key.encrypted_verifier_hash_value.clone());
            decrypt_aes_cbc_no_padding_in_place(k.as_slice(), &verifier_value_iv, &mut decrypted)
                .map_err(|e| DecryptError::InvalidInfo(format!("decrypt verifierHashValue: {e}")))?;
            if decrypted.len() < password_key.hash_size {
                return Err(DecryptError::InvalidInfo(
                    "decrypted verifierHashValue shorter than hashSize".into(),
                ));
            }
            decrypted.truncate(password_key.hash_size);
            decrypted
        };

        let expected_hash_full =
            Zeroizing::new(hash_bytes(password_key.hash_algorithm, verifier_input.as_slice()));
        let expected_hash = expected_hash_full
            .get(..password_key.hash_size)
            .ok_or_else(|| DecryptError::InvalidInfo("hash output shorter than hashSize".into()))?;

        if !ct_eq(expected_hash, verifier_hash.as_slice()) {
            return Err(DecryptError::InvalidPassword);
        }

        let key_value = {
            let k = Zeroizing::new(
                derive_key(
                    password_hash,
                    &KEY_VALUE_BLOCK,
                    key_encrypt_key_len,
                    password_key.hash_algorithm,
                )
                .map_err(map_crypto_err("derive_key(keyValue)"))?,
            );
            let mut decrypted = Zeroizing::new(password_key.encrypted_key_value.clone());
            decrypt_aes_cbc_no_padding_in_place(k.as_slice(), &key_value_iv, &mut decrypted)
                .map_err(|e| DecryptError::InvalidInfo(format!("decrypt encryptedKeyValue: {e}")))?;
            if decrypted.len() < package_key_len {
                return Err(DecryptError::InvalidInfo(
                    "decrypted keyValue shorter than keyData.keyBits".into(),
                ));
            }
            decrypted.truncate(package_key_len);
            decrypted
        };

        Ok(key_value)
    }

    // Some producers (notably `msoffcrypto-tool`) derive the IV for verifier/keyValue fields via
    // `derive_iv(saltValue, blockKey)` instead of using the raw saltValue.
    let key_value = match try_unwrap_key(
        info,
        &password_hash,
        key_encrypt_key_len,
        package_key_len,
        PasswordKeyIvMode::Salt,
    ) {
        Ok(k) => k,
        Err(DecryptError::InvalidPassword) => try_unwrap_key(
            info,
            &password_hash,
            key_encrypt_key_len,
            package_key_len,
            PasswordKeyIvMode::Derived,
        )?,
        Err(other) => return Err(other),
    };

    Ok(key_value)
}

fn map_crypto_err(ctx: &'static str) -> impl FnOnce(CryptoError) -> DecryptError {
    move |e| DecryptError::InvalidInfo(format!("{ctx}: {e}"))
}

fn key_len_bytes(
    key_bits: usize,
    element: &'static str,
    attr: &'static str,
) -> Result<usize, DecryptError> {
    if key_bits % 8 != 0 {
        return Err(DecryptError::InvalidInfo(format!(
            "{element}.{attr} must be divisible by 8"
        )));
    }
    Ok(key_bits / 8)
}

fn hash_bytes(alg: HashAlgorithm, data: &[u8]) -> Vec<u8> {
    use sha2::Digest as _;

    match alg {
        HashAlgorithm::Sha1 => sha1::Sha1::digest(data).to_vec(),
        HashAlgorithm::Sha256 => sha2::Sha256::digest(data).to_vec(),
        HashAlgorithm::Sha384 => sha2::Sha384::digest(data).to_vec(),
        HashAlgorithm::Sha512 => sha2::Sha512::digest(data).to_vec(),
    }
}

/// Compare two byte slices in constant time.
///
/// Use this for password verifier digests to avoid timing side channels.
fn ct_eq(a: &[u8], b: &[u8]) -> bool {
    if a.len() != b.len() {
        return false;
    }
    let mut diff = 0u8;
    for (&x, &y) in a.iter().zip(b.iter()) {
        diff |= x ^ y;
    }
    diff == 0
}

fn required_attr<'a>(node: roxmltree::Node<'a, '_>, attr: &str) -> Result<&'a str, DecryptError> {
    node.attribute(attr).ok_or_else(|| {
        DecryptError::InvalidInfo(format!(
            "missing attribute `{attr}` on element `{}`",
            node.tag_name().name()
        ))
    })
}

fn parse_usize_attr(node: roxmltree::Node<'_, '_>, attr: &str) -> Result<usize, DecryptError> {
    let val = required_attr(node, attr)?;
    val.trim()
        .parse::<usize>()
        .map_err(|err| DecryptError::InvalidInfo(format!("invalid `{attr}` value `{val}`: {err}")))
}

fn parse_u32_attr(node: roxmltree::Node<'_, '_>, attr: &str) -> Result<u32, DecryptError> {
    let val = required_attr(node, attr)?;
    val.trim()
        .parse::<u32>()
        .map_err(|err| DecryptError::InvalidInfo(format!("invalid `{attr}` value `{val}`: {err}")))
}

fn parse_base64_attr(node: roxmltree::Node<'_, '_>, attr: &str) -> Result<Vec<u8>, DecryptError> {
    let val = required_attr(node, attr)?;
    BASE64
        .decode(val.trim())
        .map_err(|err| DecryptError::InvalidInfo(format!("base64 decode `{attr}`: {err}")))
}

fn parse_hash_algorithm(
    node: roxmltree::Node<'_, '_>,
    attr: &str,
) -> Result<HashAlgorithm, DecryptError> {
    let val = required_attr(node, attr)?;
    HashAlgorithm::parse_offcrypto_name(val)
        .map_err(|_| DecryptError::InvalidInfo(format!("unsupported hashAlgorithm `{val}`")))
}
