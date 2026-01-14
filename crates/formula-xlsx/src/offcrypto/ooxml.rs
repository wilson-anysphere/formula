use super::{
    decrypt_aes_cbc_no_padding, decrypt_agile_encrypted_package, derive_iv, derive_key,
    hash_password, HashAlgorithm, OffCryptoError, Result, AES_BLOCK_SIZE,
};

const ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN: usize = 8;
const ENCRYPTED_PACKAGE_SEGMENT_LEN: usize = 0x1000;

/// Decrypt an Excel "Encrypt with Password" OOXML encrypted container.
///
/// The caller is responsible for extracting the `EncryptionInfo` and `EncryptedPackage` streams
/// from the surrounding OLE/CFB container.
pub fn decrypt_ooxml_encrypted_package(
    encryption_info_stream: &[u8],
    encrypted_package_stream: &[u8],
    password: &str,
) -> Result<Vec<u8>> {
    if encryption_info_stream.len() < 8 {
        return Err(OffCryptoError::EncryptionInfoTooShort {
            len: encryption_info_stream.len(),
        });
    }

    let major = u16::from_le_bytes([encryption_info_stream[0], encryption_info_stream[1]]);
    let minor = u16::from_le_bytes([encryption_info_stream[2], encryption_info_stream[3]]);

    match (major, minor) {
        (4, 4) => decrypt_agile_encrypted_package(
            encryption_info_stream,
            encrypted_package_stream,
            password,
        ),
        (3, 2) => decrypt_standard(encryption_info_stream, encrypted_package_stream, password),
        _ => Err(OffCryptoError::UnsupportedEncryptionVersion { major, minor }),
    }
}

#[derive(Debug)]
struct StandardEncryptionInfo {
    salt: Vec<u8>,
    key_len: usize,
    verifier_hash_size: usize,
    encrypted_verifier: Vec<u8>,
    encrypted_verifier_hash: Vec<u8>,
}

// Standard (CryptoAPI) encryption uses a fixed spin count; keep it low so fixtures decrypt quickly.
const STANDARD_SPIN_COUNT: u32 = 1000;

fn decrypt_standard(
    encryption_info_stream: &[u8],
    encrypted_package_stream: &[u8],
    password: &str,
) -> Result<Vec<u8>> {
    let info = parse_standard_encryption_info(encryption_info_stream)?;

    // Derive the file key (AES key) from the password.
    let pw_hash = hash_password(
        password,
        &info.salt,
        STANDARD_SPIN_COUNT,
        HashAlgorithm::Sha1,
    )
    .map_err(|err| OffCryptoError::StandardEncryptionInfoMalformed {
        reason: err.to_string(),
    })?;

    let key = derive_key(
        &pw_hash,
        &0u32.to_le_bytes(),
        info.key_len,
        HashAlgorithm::Sha1,
    )
    .map_err(|err| OffCryptoError::StandardEncryptionInfoMalformed {
        reason: err.to_string(),
    })?;

    // Verify password by decrypting the verifier + verifier hash.
    let iv_ver = derive_iv(
        &info.salt,
        &0u32.to_le_bytes(),
        AES_BLOCK_SIZE,
        HashAlgorithm::Sha1,
    )
    .map_err(|err| OffCryptoError::StandardEncryptionInfoMalformed {
        reason: err.to_string(),
    })?;
    let verifier = decrypt_aes_cbc_no_padding(&key, &iv_ver, &info.encrypted_verifier)?;

    let iv_hash = derive_iv(
        &info.salt,
        &1u32.to_le_bytes(),
        AES_BLOCK_SIZE,
        HashAlgorithm::Sha1,
    )
    .map_err(|err| OffCryptoError::StandardEncryptionInfoMalformed {
        reason: err.to_string(),
    })?;
    let verifier_hash = decrypt_aes_cbc_no_padding(&key, &iv_hash, &info.encrypted_verifier_hash)?;

    let expected = HashAlgorithm::Sha1.hash(&verifier);
    let expected = expected.get(..info.verifier_hash_size).ok_or_else(|| {
        OffCryptoError::StandardEncryptionInfoMalformed {
            reason: "verifierHashSize larger than SHA1 digest length".to_string(),
        }
    })?;
    let got = verifier_hash
        .get(..info.verifier_hash_size)
        .ok_or_else(|| OffCryptoError::StandardEncryptionInfoMalformed {
            reason: "decrypted verifierHash is truncated".to_string(),
        })?;
    if expected != got {
        return Err(OffCryptoError::WrongPassword);
    }

    decrypt_encrypted_package_stream(
        encrypted_package_stream,
        &key,
        &info.salt,
        HashAlgorithm::Sha1,
        AES_BLOCK_SIZE,
    )
}

fn parse_standard_encryption_info(bytes: &[u8]) -> Result<StandardEncryptionInfo> {
    if bytes.len() < 8 {
        return Err(OffCryptoError::EncryptionInfoTooShort { len: bytes.len() });
    }

    // Bytes[0..8] are EncryptionVersionInfo (major/minor/flags). We already dispatch on major/minor
    // at the entrypoint, so just skip them here.
    let mut offset = 8usize;

    let header_size = read_u32_le(bytes, &mut offset).ok_or_else(|| {
        OffCryptoError::StandardEncryptionInfoMalformed {
            reason: "truncated EncryptionHeader.size".to_string(),
        }
    })? as usize;
    if bytes.len() < offset + header_size {
        return Err(OffCryptoError::StandardEncryptionInfoMalformed {
            reason: "truncated EncryptionHeader".to_string(),
        });
    }
    let header_bytes = &bytes[offset..offset + header_size];
    offset += header_size;

    if header_bytes.len() < 8 * 4 {
        return Err(OffCryptoError::StandardEncryptionInfoMalformed {
            reason: "EncryptionHeader missing fixed fields".to_string(),
        });
    }

    // keySize (bits) is DWORD #5 (0-indexed) in the fixed fields.
    let key_size_bits =
        u32::from_le_bytes(header_bytes[16..20].try_into().expect("slice is 4 bytes")) as usize;
    let key_len = key_size_bits
        .checked_div(8)
        .filter(|n| *n > 0)
        .ok_or_else(|| OffCryptoError::StandardEncryptionInfoMalformed {
            reason: "invalid keySize".to_string(),
        })?;

    let salt_size = read_u32_le(bytes, &mut offset).ok_or_else(|| {
        OffCryptoError::StandardEncryptionInfoMalformed {
            reason: "truncated EncryptionVerifier.saltSize".to_string(),
        }
    })? as usize;
    if bytes.len() < offset + salt_size {
        return Err(OffCryptoError::StandardEncryptionInfoMalformed {
            reason: "truncated EncryptionVerifier.salt".to_string(),
        });
    }
    let salt = bytes[offset..offset + salt_size].to_vec();
    offset += salt_size;

    if bytes.len() < offset + 16 {
        return Err(OffCryptoError::StandardEncryptionInfoMalformed {
            reason: "truncated EncryptionVerifier.encryptedVerifier".to_string(),
        });
    }
    let encrypted_verifier = bytes[offset..offset + 16].to_vec();
    offset += 16;

    let verifier_hash_size = read_u32_le(bytes, &mut offset).ok_or_else(|| {
        OffCryptoError::StandardEncryptionInfoMalformed {
            reason: "truncated EncryptionVerifier.verifierHashSize".to_string(),
        }
    })? as usize;

    let encrypted_verifier_hash = bytes.get(offset..).unwrap_or_default().to_vec();
    if encrypted_verifier_hash.is_empty() {
        return Err(OffCryptoError::StandardEncryptionInfoMalformed {
            reason: "missing EncryptionVerifier.encryptedVerifierHash".to_string(),
        });
    }

    Ok(StandardEncryptionInfo {
        salt,
        key_len,
        verifier_hash_size,
        encrypted_verifier,
        encrypted_verifier_hash,
    })
}

fn decrypt_encrypted_package_stream(
    encrypted_package_stream: &[u8],
    key: &[u8],
    salt: &[u8],
    hash_alg: HashAlgorithm,
    block_size: usize,
) -> Result<Vec<u8>> {
    if encrypted_package_stream.len() < ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN {
        return Err(OffCryptoError::EncryptedPackageTooShort {
            len: encrypted_package_stream.len(),
        });
    }
    let mut size_bytes = [0u8; ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN];
    size_bytes.copy_from_slice(&encrypted_package_stream[..ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN]);
    let orig_size = u64::from_le_bytes(size_bytes);
    let orig_size_usize =
        usize::try_from(orig_size).map_err(|_| OffCryptoError::InvalidAttribute {
            element: "EncryptedPackage".to_string(),
            attr: "origSize".to_string(),
            reason: "origSize does not fit into usize".to_string(),
        })?;

    let ciphertext = &encrypted_package_stream[ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN..];
    if ciphertext.is_empty() && orig_size == 0 {
        return Ok(Vec::new());
    }
    if ciphertext.len() % AES_BLOCK_SIZE != 0 {
        return Err(OffCryptoError::CiphertextNotBlockAligned {
            field: "EncryptedPackage",
            len: ciphertext.len(),
        });
    }

    let mut out = Vec::with_capacity(orig_size_usize);
    let mut offset = 0usize;
    let mut segment_index: u32 = 0;
    while offset < ciphertext.len() && out.len() < orig_size_usize {
        let remaining = ciphertext.len() - offset;
        let seg_len = remaining.min(ENCRYPTED_PACKAGE_SEGMENT_LEN);

        if seg_len % AES_BLOCK_SIZE != 0 {
            return Err(OffCryptoError::CiphertextNotBlockAligned {
                field: "EncryptedPackage segment",
                len: seg_len,
            });
        }

        let iv =
            derive_iv(salt, &segment_index.to_le_bytes(), block_size, hash_alg).map_err(|err| {
                OffCryptoError::InvalidAttribute {
                    element: "EncryptedPackage".to_string(),
                    attr: "iv".to_string(),
                    reason: err.to_string(),
                }
            })?;

        let decrypted =
            decrypt_aes_cbc_no_padding(key, &iv, &ciphertext[offset..offset + seg_len])?;

        let remaining_needed = orig_size_usize - out.len();
        if decrypted.len() > remaining_needed {
            out.extend_from_slice(&decrypted[..remaining_needed]);
            break;
        }
        out.extend_from_slice(&decrypted);

        offset += seg_len;
        segment_index = segment_index.wrapping_add(1);
    }

    if out.len() < orig_size_usize {
        return Err(OffCryptoError::DecryptedLengthShorterThanHeader {
            declared_len: orig_size_usize,
            available_len: out.len(),
        });
    }
    out.truncate(orig_size_usize);
    Ok(out)
}

fn read_u32_le(bytes: &[u8], offset: &mut usize) -> Option<u32> {
    let b = bytes.get(*offset..*offset + 4)?;
    *offset += 4;
    Some(u32::from_le_bytes([b[0], b[1], b[2], b[3]]))
}
