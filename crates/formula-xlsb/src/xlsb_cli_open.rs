use std::fs::File;
use std::io::{self, Read, Seek};
use std::path::Path;

use aes::{Aes128, Aes192, Aes256};
use cbc::cipher::{block_padding::NoPadding, BlockDecryptMut, KeyIvInit};
use formula_offcrypto::{
    AgileEncryptionInfo, EncryptionInfo, HashAlgorithm, StandardEncryptionInfo,
};
use formula_xlsb::XlsbWorkbook;
use sha1::Sha1;
use sha2::{Digest as _, Sha256, Sha384, Sha512};

/// OLE/CFB file signature.
///
/// See: <https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/>
const OLE_MAGIC: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

const ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN: usize = 8;
const ENCRYPTED_PACKAGE_SEGMENT_LEN: usize = 0x1000;
const AES_BLOCK_LEN: usize = 16;

// MS-OFFCRYPTO Agile: block keys used when decrypting password key-encryptor fields.
const VERIFIER_HASH_INPUT_BLOCK: [u8; 8] = [0xFE, 0xA7, 0xD2, 0x76, 0x3B, 0x4B, 0x9E, 0x79];
const VERIFIER_HASH_VALUE_BLOCK: [u8; 8] = [0xD7, 0xAA, 0x0F, 0x6D, 0x30, 0x61, 0x34, 0x4E];
const KEY_VALUE_BLOCK: [u8; 8] = [0x14, 0x6E, 0x0B, 0xE7, 0xAB, 0xAC, 0xD0, 0xD6];

/// Open an `.xlsb` workbook, optionally using a password for Office-encrypted OLE wrappers.
pub fn open_xlsb_workbook(
    path: &Path,
    password: Option<&str>,
) -> Result<XlsbWorkbook, Box<dyn std::error::Error>> {
    // Fast path: ZIP-based `.xlsb`.
    if !looks_like_ole_compound_file(path)? {
        return Ok(XlsbWorkbook::open(path)?);
    }

    // OLE-based: could be legacy `.xls` or Office-encrypted OOXML.
    let file = File::open(path)?;
    let mut ole = cfb::CompoundFile::open(file)?;

    if stream_exists(&mut ole, "EncryptionInfo") && stream_exists(&mut ole, "EncryptedPackage") {
        let password = password.ok_or_else(|| {
            io::Error::new(
                io::ErrorKind::InvalidInput,
                "encrypted workbook requires a password; pass --password <pw>",
            )
        })?;

        let zip_bytes = decrypt_ooxml_encrypted_package(&mut ole, password)?;
        return Ok(XlsbWorkbook::open_from_bytes(&zip_bytes)?);
    }

    // Fall back to the normal ZIP open path so the caller gets a sensible parse error.
    Ok(XlsbWorkbook::open(path)?)
}

fn looks_like_ole_compound_file(path: &Path) -> Result<bool, io::Error> {
    let mut file = File::open(path)?;
    let mut header = [0u8; OLE_MAGIC.len()];
    let n = file.read(&mut header)?;
    Ok(n == OLE_MAGIC.len() && header == OLE_MAGIC)
}

fn stream_exists<R: Read + Seek + std::io::Write>(
    ole: &mut cfb::CompoundFile<R>,
    name: &str,
) -> bool {
    if ole.open_stream(name).is_ok() {
        return true;
    }
    let with_leading_slash = format!("/{name}");
    ole.open_stream(&with_leading_slash).is_ok()
}

fn open_stream<R: Read + Seek + std::io::Write>(
    ole: &mut cfb::CompoundFile<R>,
    name: &str,
) -> Result<cfb::Stream<R>, Box<dyn std::error::Error>> {
    match ole.open_stream(name) {
        Ok(s) => Ok(s),
        Err(_) => {
            let with_leading_slash = format!("/{name}");
            Ok(ole.open_stream(&with_leading_slash)?)
        }
    }
}

fn read_stream<R: Read + Seek + std::io::Write>(
    ole: &mut cfb::CompoundFile<R>,
    name: &str,
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let mut stream = open_stream(ole, name)?;
    let mut bytes = Vec::new();
    stream.read_to_end(&mut bytes)?;
    Ok(bytes)
}

fn decrypt_ooxml_encrypted_package<R: Read + Seek + std::io::Write>(
    ole: &mut cfb::CompoundFile<R>,
    password: &str,
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let encryption_info_bytes = read_stream(ole, "EncryptionInfo")?;
    let encrypted_package_bytes = read_stream(ole, "EncryptedPackage")?;

    let info = formula_offcrypto::parse_encryption_info(&encryption_info_bytes)?;
    match info {
        EncryptionInfo::Standard { header, verifier, .. } => {
            let info = StandardEncryptionInfo { header, verifier };
            let key = formula_offcrypto::standard_derive_key(&info, password)?;
            formula_offcrypto::standard_verify_key(&info, &key)?;
            decrypt_standard_encrypted_package_stream(&encrypted_package_bytes, &key, &info.verifier.salt)
                .map_err(|err| err.into())
        }
        EncryptionInfo::Agile { info, .. } => decrypt_agile_encrypted_package_stream(
            &encrypted_package_bytes,
            &info,
            password,
        )
        .map_err(|err| err.into()),
        EncryptionInfo::Unsupported { version } => Err(Box::new(io::Error::new(
            io::ErrorKind::InvalidInput,
            format!(
                "unsupported EncryptionInfo version {}.{} (only Standard 3.2 and Agile 4.4 are supported)",
                version.major, version.minor
            ),
        ))),
    }
}

#[derive(Debug, thiserror::Error)]
enum EncryptedPackageError {
    #[error("`EncryptedPackage` stream is too short: expected at least {ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN} bytes, got {len}")]
    StreamTooShort { len: usize },
    #[error(
        "`EncryptedPackage` orig_size {orig_size} does not fit into the current platform `usize`"
    )]
    OrigSizeTooLargeForPlatform { orig_size: u64 },
    #[error("`EncryptedPackage` ciphertext length {ciphertext_len} is not a multiple of AES block size ({AES_BLOCK_LEN})")]
    CiphertextLenNotBlockAligned { ciphertext_len: usize },
    #[error("`EncryptedPackage` ciphertext segment length {ciphertext_len} is not a multiple of AES block size ({AES_BLOCK_LEN})")]
    SegmentLenNotBlockAligned { ciphertext_len: usize },
    #[error("`EncryptedPackage` segment index overflow (file too large)")]
    SegmentIndexOverflow,
    #[error("AES-CBC decryption failed for segment {segment_index}")]
    SegmentDecryptFailed { segment_index: u32 },
    #[error("decrypted plaintext is shorter than expected: got {decrypted_len} bytes, expected at least {orig_size}")]
    DecryptedTooShort {
        decrypted_len: usize,
        orig_size: u64,
    },
    #[error("unsupported hash algorithm for IV derivation")]
    UnsupportedHashAlgorithm,
    #[error("invalid encrypted workbook password (verifier mismatch)")]
    InvalidPassword,
    #[error("unsupported AES key length {key_len} bytes (expected 16, 24, or 32)")]
    InvalidAesKeyLength { key_len: usize },
    #[error("unsupported cipher block size {block_size} (expected 16 for AES-CBC)")]
    UnsupportedBlockSize { block_size: usize },
}

fn decrypt_standard_encrypted_package_stream(
    encrypted_package_stream: &[u8],
    key: &[u8],
    salt: &[u8],
) -> Result<Vec<u8>, EncryptedPackageError> {
    decrypt_encrypted_package_stream_with_iv_hash(
        encrypted_package_stream,
        key,
        salt,
        HashAlgorithm::Sha1,
        AES_BLOCK_LEN,
    )
}

fn decrypt_agile_encrypted_package_stream(
    encrypted_package_stream: &[u8],
    info: &AgileEncryptionInfo,
    password: &str,
) -> Result<Vec<u8>, EncryptedPackageError> {
    if info.key_data_block_size != AES_BLOCK_LEN {
        return Err(EncryptedPackageError::UnsupportedBlockSize {
            block_size: info.key_data_block_size,
        });
    }

    let key_len = info
        .password_key_bits
        .checked_div(8)
        .filter(|_| info.password_key_bits % 8 == 0)
        .ok_or(EncryptedPackageError::InvalidAesKeyLength { key_len: 0 })?;

    if !matches!(key_len, 16 | 24 | 32) {
        return Err(EncryptedPackageError::InvalidAesKeyLength { key_len });
    }

    // Password → hash (`H`) per MS-OFFCRYPTO Agile.
    let h = hash_password_agile(
        password,
        &info.password_salt,
        info.spin_count,
        info.password_hash_algorithm,
    )?;

    // Decrypt verifier fields and validate password.
    let verifier_hash_input = decrypt_agile_value(
        &info.encrypted_verifier_hash_input,
        &h,
        &info.password_salt,
        &VERIFIER_HASH_INPUT_BLOCK,
        key_len,
        info.password_hash_algorithm,
    )?;
    let verifier_hash_value = decrypt_agile_value(
        &info.encrypted_verifier_hash_value,
        &h,
        &info.password_salt,
        &VERIFIER_HASH_VALUE_BLOCK,
        key_len,
        info.password_hash_algorithm,
    )?;

    let expected = hash_bytes(&verifier_hash_input, info.password_hash_algorithm)?;
    let digest_len = hash_digest_len(info.password_hash_algorithm)?;
    if verifier_hash_value.len() < digest_len || expected[..] != verifier_hash_value[..digest_len] {
        return Err(EncryptedPackageError::InvalidPassword);
    }

    // Decrypt the package key (`encryptedKeyValue`).
    let key_value = decrypt_agile_value(
        &info.encrypted_key_value,
        &h,
        &info.password_salt,
        &KEY_VALUE_BLOCK,
        key_len,
        info.password_hash_algorithm,
    )?;
    if key_value.len() < key_len {
        return Err(EncryptedPackageError::InvalidPassword);
    }
    let package_key = &key_value[..key_len];

    decrypt_encrypted_package_stream_with_iv_hash(
        encrypted_package_stream,
        package_key,
        &info.key_data_salt,
        info.key_data_hash_algorithm,
        info.key_data_block_size,
    )
}

fn decrypt_encrypted_package_stream_with_iv_hash(
    encrypted_package_stream: &[u8],
    key: &[u8],
    salt: &[u8],
    iv_hash_alg: HashAlgorithm,
    iv_len: usize,
) -> Result<Vec<u8>, EncryptedPackageError> {
    if encrypted_package_stream.len() < ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN {
        return Err(EncryptedPackageError::StreamTooShort {
            len: encrypted_package_stream.len(),
        });
    }
    if iv_len != AES_BLOCK_LEN {
        return Err(EncryptedPackageError::UnsupportedBlockSize { block_size: iv_len });
    }
    if !matches!(key.len(), 16 | 24 | 32) {
        return Err(EncryptedPackageError::InvalidAesKeyLength { key_len: key.len() });
    }

    let mut size_bytes = [0u8; ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN];
    size_bytes.copy_from_slice(&encrypted_package_stream[..ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN]);
    let orig_size = u64::from_le_bytes(size_bytes);
    let ciphertext = &encrypted_package_stream[ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN..];

    let orig_size_usize = usize::try_from(orig_size)
        .map_err(|_| EncryptedPackageError::OrigSizeTooLargeForPlatform { orig_size })?;

    if ciphertext.len() % AES_BLOCK_LEN != 0 {
        return Err(EncryptedPackageError::CiphertextLenNotBlockAligned {
            ciphertext_len: ciphertext.len(),
        });
    }

    // Decrypt segment-by-segment until we have produced `orig_size` bytes (or run out of input).
    let mut out = Vec::with_capacity(orig_size_usize);
    let mut offset = 0usize;
    let mut segment_index: u32 = 0;
    while offset < ciphertext.len() && out.len() < orig_size_usize {
        let remaining = ciphertext.len() - offset;
        let seg_len = if remaining > ENCRYPTED_PACKAGE_SEGMENT_LEN {
            ENCRYPTED_PACKAGE_SEGMENT_LEN
        } else {
            remaining
        };

        if seg_len % AES_BLOCK_LEN != 0 {
            return Err(EncryptedPackageError::SegmentLenNotBlockAligned {
                ciphertext_len: seg_len,
            });
        }

        let iv = derive_segment_iv(salt, segment_index, iv_hash_alg)?;
        let decrypted = decrypt_aes_cbc_no_padding(
            key,
            &iv,
            &ciphertext[offset..offset + seg_len],
            segment_index,
        )?;

        let remaining_needed = orig_size_usize - out.len();
        if decrypted.len() > remaining_needed {
            out.extend_from_slice(&decrypted[..remaining_needed]);
            break;
        }

        out.extend_from_slice(&decrypted);
        offset += seg_len;
        segment_index = segment_index
            .checked_add(1)
            .ok_or(EncryptedPackageError::SegmentIndexOverflow)?;
    }

    if out.len() < orig_size_usize {
        return Err(EncryptedPackageError::DecryptedTooShort {
            decrypted_len: out.len(),
            orig_size,
        });
    }
    out.truncate(orig_size_usize);
    Ok(out)
}

fn derive_segment_iv(
    salt: &[u8],
    segment_index: u32,
    hash_alg: HashAlgorithm,
) -> Result<[u8; AES_BLOCK_LEN], EncryptedPackageError> {
    let digest = hash_two(salt, &segment_index.to_le_bytes(), hash_alg)?;
    let mut iv = [0u8; AES_BLOCK_LEN];
    iv.copy_from_slice(&digest[..AES_BLOCK_LEN]);
    Ok(iv)
}

fn decrypt_aes_cbc_no_padding(
    key: &[u8],
    iv: &[u8; AES_BLOCK_LEN],
    ciphertext: &[u8],
    segment_index: u32,
) -> Result<Vec<u8>, EncryptedPackageError> {
    let mut buf = ciphertext.to_vec();

    let res = match key.len() {
        16 => cbc::Decryptor::<Aes128>::new_from_slices(key, iv)
            .map_err(|_| EncryptedPackageError::InvalidAesKeyLength { key_len: key.len() })?
            .decrypt_padded_mut::<NoPadding>(&mut buf),
        24 => cbc::Decryptor::<Aes192>::new_from_slices(key, iv)
            .map_err(|_| EncryptedPackageError::InvalidAesKeyLength { key_len: key.len() })?
            .decrypt_padded_mut::<NoPadding>(&mut buf),
        32 => cbc::Decryptor::<Aes256>::new_from_slices(key, iv)
            .map_err(|_| EncryptedPackageError::InvalidAesKeyLength { key_len: key.len() })?
            .decrypt_padded_mut::<NoPadding>(&mut buf),
        _ => return Err(EncryptedPackageError::InvalidAesKeyLength { key_len: key.len() }),
    };

    res.map_err(|_| EncryptedPackageError::SegmentDecryptFailed { segment_index })?;
    Ok(buf)
}

fn hash_digest_len(alg: HashAlgorithm) -> Result<usize, EncryptedPackageError> {
    Ok(match alg {
        HashAlgorithm::Sha1 => 20,
        HashAlgorithm::Sha256 => 32,
        HashAlgorithm::Sha384 => 48,
        HashAlgorithm::Sha512 => 64,
    })
}

fn hash_two(a: &[u8], b: &[u8], alg: HashAlgorithm) -> Result<Vec<u8>, EncryptedPackageError> {
    let out = match alg {
        HashAlgorithm::Sha1 => {
            let mut h = Sha1::new();
            h.update(a);
            h.update(b);
            h.finalize().to_vec()
        }
        HashAlgorithm::Sha256 => {
            let mut h = Sha256::new();
            h.update(a);
            h.update(b);
            h.finalize().to_vec()
        }
        HashAlgorithm::Sha384 => {
            let mut h = Sha384::new();
            h.update(a);
            h.update(b);
            h.finalize().to_vec()
        }
        HashAlgorithm::Sha512 => {
            let mut h = Sha512::new();
            h.update(a);
            h.update(b);
            h.finalize().to_vec()
        }
    };
    // Some callers truncate the result; keep full digest here.
    // Validate that we can safely truncate to AES_BLOCK_LEN for IV derivation.
    if out.len() < AES_BLOCK_LEN {
        return Err(EncryptedPackageError::UnsupportedHashAlgorithm);
    }
    Ok(out)
}

fn hash_bytes(bytes: &[u8], alg: HashAlgorithm) -> Result<Vec<u8>, EncryptedPackageError> {
    match alg {
        HashAlgorithm::Sha1 => {
            let digest: [u8; 20] = Sha1::digest(bytes).into();
            Ok(digest.to_vec())
        }
        HashAlgorithm::Sha256 => Ok(Sha256::digest(bytes).to_vec()),
        HashAlgorithm::Sha384 => Ok(Sha384::digest(bytes).to_vec()),
        HashAlgorithm::Sha512 => Ok(Sha512::digest(bytes).to_vec()),
    }
}

fn password_to_utf16le_bytes(password: &str) -> Vec<u8> {
    let mut out = Vec::with_capacity(password.len().saturating_mul(2));
    for unit in password.encode_utf16() {
        out.extend_from_slice(&unit.to_le_bytes());
    }
    out
}

fn hash_password_agile(
    password: &str,
    salt: &[u8],
    spin: u32,
    hash_alg: HashAlgorithm,
) -> Result<Vec<u8>, EncryptedPackageError> {
    let pw = password_to_utf16le_bytes(password);

    let mut h = hash_two(salt, &pw, hash_alg)?;
    for i in 0..spin {
        h = hash_two(&i.to_le_bytes(), &h, hash_alg)?;
    }
    Ok(h)
}

fn derive_key_agile(
    h: &[u8],
    block_key: &[u8; 8],
    key_len: usize,
    hash_alg: HashAlgorithm,
) -> Result<Vec<u8>, EncryptedPackageError> {
    let digest_len = hash_digest_len(hash_alg)?;
    let mut out = hash_two(h, block_key, hash_alg)?;
    if key_len <= digest_len {
        out.truncate(key_len);
    } else {
        out.resize(key_len, 0u8);
    }
    Ok(out)
}

fn derive_iv_agile(
    salt: &[u8],
    block_key: &[u8; 8],
    hash_alg: HashAlgorithm,
) -> Result<[u8; AES_BLOCK_LEN], EncryptedPackageError> {
    let digest = hash_two(salt, block_key, hash_alg)?;
    let mut iv = [0u8; AES_BLOCK_LEN];
    iv.copy_from_slice(&digest[..AES_BLOCK_LEN]);
    Ok(iv)
}

fn decrypt_agile_value(
    ciphertext: &[u8],
    password_hash: &[u8],
    salt: &[u8],
    block_key: &[u8; 8],
    key_len: usize,
    hash_alg: HashAlgorithm,
) -> Result<Vec<u8>, EncryptedPackageError> {
    let key = derive_key_agile(password_hash, block_key, key_len, hash_alg)?;
    let iv = derive_iv_agile(salt, block_key, hash_alg)?;

    let mut buf = ciphertext.to_vec();
    // For AES-CBC the ciphertext is always a multiple of 16 bytes.
    if buf.len() % AES_BLOCK_LEN != 0 {
        return Err(EncryptedPackageError::CiphertextLenNotBlockAligned {
            ciphertext_len: buf.len(),
        });
    }

    let res = match key.len() {
        16 => cbc::Decryptor::<Aes128>::new_from_slices(&key, &iv)
            .map_err(|_| EncryptedPackageError::InvalidAesKeyLength { key_len: key.len() })?
            .decrypt_padded_mut::<NoPadding>(&mut buf),
        24 => cbc::Decryptor::<Aes192>::new_from_slices(&key, &iv)
            .map_err(|_| EncryptedPackageError::InvalidAesKeyLength { key_len: key.len() })?
            .decrypt_padded_mut::<NoPadding>(&mut buf),
        32 => cbc::Decryptor::<Aes256>::new_from_slices(&key, &iv)
            .map_err(|_| EncryptedPackageError::InvalidAesKeyLength { key_len: key.len() })?
            .decrypt_padded_mut::<NoPadding>(&mut buf),
        _ => return Err(EncryptedPackageError::InvalidAesKeyLength { key_len: key.len() }),
    };

    res.map_err(|_| EncryptedPackageError::SegmentDecryptFailed { segment_index: 0 })?;
    Ok(buf)
}
