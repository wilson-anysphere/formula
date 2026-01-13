use crate::crypto::{aes_cbc_decrypt, derive_iv, HashAlgorithm, StandardKeyDeriver};
use crate::error::OfficeCryptoError;
use crate::util::{
    checked_vec_len, decode_utf16le_nul_terminated, read_u32_le, read_u64_le, EncryptionInfoHeader,
};

#[derive(Debug, Clone)]
pub(crate) struct StandardEncryptionInfo {
    #[allow(dead_code)]
    pub(crate) version_major: u16,
    #[allow(dead_code)]
    pub(crate) version_minor: u16,
    #[allow(dead_code)]
    pub(crate) flags: u32,
    pub(crate) header: EncryptionHeader,
    pub(crate) verifier: EncryptionVerifier,
}

#[derive(Debug, Clone)]
pub(crate) struct EncryptionHeader {
    pub(crate) alg_id: u32,
    pub(crate) alg_id_hash: u32,
    pub(crate) key_bits: u32,
    #[allow(dead_code)]
    pub(crate) provider_type: u32,
    #[allow(dead_code)]
    pub(crate) csp_name: String,
}

#[derive(Debug, Clone)]
pub(crate) struct EncryptionVerifier {
    pub(crate) salt: Vec<u8>,
    pub(crate) encrypted_verifier: Vec<u8>,
    pub(crate) verifier_hash_size: u32,
    pub(crate) encrypted_verifier_hash: Vec<u8>,
}

pub(crate) fn parse_standard_encryption_info(
    bytes: &[u8],
    header: &EncryptionInfoHeader,
) -> Result<StandardEncryptionInfo, OfficeCryptoError> {
    let start = header.header_offset;
    let header_size = header.header_size as usize;
    let header_bytes = bytes.get(start..start + header_size).ok_or_else(|| {
        OfficeCryptoError::InvalidFormat("EncryptionInfo header size out of range".to_string())
    })?;
    let verifier_bytes = bytes.get(start + header_size..).ok_or_else(|| {
        OfficeCryptoError::InvalidFormat("EncryptionInfo missing verifier".to_string())
    })?;

    let enc_header = parse_encryption_header(header_bytes)?;
    let verifier = parse_encryption_verifier(verifier_bytes, &enc_header)?;

    Ok(StandardEncryptionInfo {
        version_major: header.version_major,
        version_minor: header.version_minor,
        flags: header.flags,
        header: enc_header,
        verifier,
    })
}

fn parse_encryption_header(bytes: &[u8]) -> Result<EncryptionHeader, OfficeCryptoError> {
    if bytes.len() < 8 * 4 {
        return Err(OfficeCryptoError::InvalidFormat(
            "EncryptionHeader too short".to_string(),
        ));
    }
    // DWORD flags, sizeExtra, algId, algIdHash, keySize, providerType, reserved1, reserved2
    let _flags = read_u32_le(bytes, 0)?;
    let _size_extra = read_u32_le(bytes, 4)?;
    let alg_id = read_u32_le(bytes, 8)?;
    let alg_id_hash = read_u32_le(bytes, 12)?;
    let key_bits = read_u32_le(bytes, 16)?;
    let provider_type = read_u32_le(bytes, 20)?;
    let _reserved1 = read_u32_le(bytes, 24)?;
    let _reserved2 = read_u32_le(bytes, 28)?;
    let csp_name = decode_utf16le_nul_terminated(&bytes[32..])?;

    Ok(EncryptionHeader {
        alg_id,
        alg_id_hash,
        key_bits,
        provider_type,
        csp_name,
    })
}

fn parse_encryption_verifier(
    bytes: &[u8],
    header: &EncryptionHeader,
) -> Result<EncryptionVerifier, OfficeCryptoError> {
    // saltSize, salt, encryptedVerifier(16), verifierHashSize, encryptedVerifierHash(variable)
    if bytes.len() < 4 {
        return Err(OfficeCryptoError::InvalidFormat(
            "EncryptionVerifier too short".to_string(),
        ));
    }
    let salt_size = read_u32_le(bytes, 0)? as usize;
    let mut offset = 4usize;
    let salt = bytes.get(offset..offset + salt_size).ok_or_else(|| {
        OfficeCryptoError::InvalidFormat("EncryptionVerifier salt out of range".to_string())
    })?;
    offset += salt_size;
    let encrypted_verifier = bytes.get(offset..offset + 16).ok_or_else(|| {
        OfficeCryptoError::InvalidFormat("EncryptionVerifier missing verifier".to_string())
    })?;
    offset += 16;
    let verifier_hash_size = read_u32_le(bytes, offset)?;
    offset += 4;

    // Ciphertext is padded to the cipher block size. For AES-CBC, that's 16.
    let encrypted_hash_len = ((verifier_hash_size as usize + 15) / 16) * 16;
    let encrypted_verifier_hash =
        bytes
            .get(offset..offset + encrypted_hash_len)
            .ok_or_else(|| {
                OfficeCryptoError::InvalidFormat(
                    "EncryptionVerifier missing verifier hash".to_string(),
                )
            })?;

    // Basic algorithm sanity check: support AES only for now.
    match header.alg_id {
        0x0000_660E | 0x0000_660F | 0x0000_6610 => {} // CALG_AES_128/192/256
        other => {
            return Err(OfficeCryptoError::UnsupportedEncryption(format!(
                "unsupported cipher AlgID {other:#x}"
            )))
        }
    }

    Ok(EncryptionVerifier {
        salt: salt.to_vec(),
        encrypted_verifier: encrypted_verifier.to_vec(),
        verifier_hash_size,
        encrypted_verifier_hash: encrypted_verifier_hash.to_vec(),
    })
}

pub(crate) fn decrypt_standard_encrypted_package(
    info: &StandardEncryptionInfo,
    encrypted_package: &[u8],
    password: &str,
) -> Result<Vec<u8>, OfficeCryptoError> {
    if encrypted_package.len() < 8 {
        return Err(OfficeCryptoError::InvalidFormat(
            "EncryptedPackage stream too short".to_string(),
        ));
    }
    let total_size = read_u64_le(encrypted_package, 0)?;
    let expected_len = checked_vec_len(total_size)?;
    let ciphertext = &encrypted_package[8..];

    let hash_alg = HashAlgorithm::from_cryptoapi_alg_id_hash(info.header.alg_id_hash)?;

    // Try a small set of schemes seen in the wild. We validate via the password verifier and by
    // checking that the decrypted output starts with `PK`.
    let schemes: [StandardScheme; 3] = [
        StandardScheme::PerBlockKeyIvZero,
        StandardScheme::ConstKeyPerBlockIvHash,
        StandardScheme::ConstKeyIvSaltStream,
    ];

    for scheme in schemes {
        let Ok(out) = decrypt_standard_with_scheme(
            info,
            ciphertext,
            total_size,
            expected_len,
            password,
            hash_alg,
            scheme,
        ) else {
            continue;
        };

        if out.len() >= 2 && &out[..2] == b"PK" {
            return Ok(out);
        }
    }

    Err(OfficeCryptoError::InvalidPassword)
}

#[derive(Debug, Clone, Copy)]
enum StandardScheme {
    /// Segment the data in 4096-byte chunks; for chunk N use a key derived with blockIndex=N and
    /// IV=0.
    PerBlockKeyIvZero,
    /// Segment the data in 4096-byte chunks; use a single key derived with blockIndex=0 and IV
    /// derived as hash(salt||blockIndex) for each chunk.
    ConstKeyPerBlockIvHash,
    /// Treat the ciphertext as a single AES-CBC stream using key(block=0) and IV=salt.
    ConstKeyIvSaltStream,
}

fn decrypt_standard_with_scheme(
    info: &StandardEncryptionInfo,
    ciphertext: &[u8],
    total_size: u64,
    expected_len: usize,
    password: &str,
    hash_alg: HashAlgorithm,
    scheme: StandardScheme,
) -> Result<Vec<u8>, OfficeCryptoError> {
    let deriver = StandardKeyDeriver::new(
        hash_alg,
        info.header.key_bits,
        &info.verifier.salt,
        password,
    );
    let key0 = deriver.derive_key_for_block(0)?;

    // Verify password using EncryptionVerifier.
    let (verifier_key, verifier_iv) = match scheme {
        StandardScheme::PerBlockKeyIvZero => (key0.clone(), [0u8; 16].to_vec()),
        StandardScheme::ConstKeyPerBlockIvHash => (
            key0.clone(),
            derive_iv(hash_alg, &info.verifier.salt, &0u32.to_le_bytes(), 16),
        ),
        StandardScheme::ConstKeyIvSaltStream => {
            let iv = info.verifier.salt.get(..16).ok_or_else(|| {
                OfficeCryptoError::InvalidFormat("EncryptionVerifier salt too short".to_string())
            })?;
            (key0.clone(), iv.to_vec())
        }
    };

    let verifier = aes_cbc_decrypt(
        &verifier_key,
        &verifier_iv,
        &info.verifier.encrypted_verifier,
    )?;
    let verifier_hash_plain = aes_cbc_decrypt(
        &verifier_key,
        &verifier_iv,
        &info.verifier.encrypted_verifier_hash,
    )?;
    let verifier_hash_plain = verifier_hash_plain
        .get(..info.verifier.verifier_hash_size as usize)
        .ok_or_else(|| {
            OfficeCryptoError::InvalidFormat(
                "EncryptionVerifier decrypted hash shorter than verifierHashSize".to_string(),
            )
        })?;

    let verifier_hash = hash_alg.digest(&verifier);
    if verifier_hash_plain != verifier_hash.as_slice() {
        return Err(OfficeCryptoError::InvalidPassword);
    }

    // Password is valid; decrypt the package.
    match scheme {
        StandardScheme::PerBlockKeyIvZero => {
            decrypt_segmented(ciphertext, total_size, expected_len, |block| {
                let key = deriver.derive_key_for_block(block)?;
                Ok((key, [0u8; 16].to_vec()))
            })
        }
        StandardScheme::ConstKeyPerBlockIvHash => {
            decrypt_segmented(ciphertext, total_size, expected_len, |block| {
                let iv = derive_iv(hash_alg, &info.verifier.salt, &block.to_le_bytes(), 16);
                Ok((key0.clone(), iv))
            })
        }
        StandardScheme::ConstKeyIvSaltStream => {
            let iv = info.verifier.salt.get(..16).ok_or_else(|| {
                OfficeCryptoError::InvalidFormat("EncryptionVerifier salt too short".to_string())
            })?;
            let mut plain = aes_cbc_decrypt(&key0, iv, ciphertext)?;
            if expected_len > plain.len() {
                return Err(OfficeCryptoError::InvalidFormat(format!(
                    "decrypted package length {} shorter than expected {}",
                    plain.len(),
                    expected_len
                )));
            }
            plain.truncate(expected_len);
            Ok(plain)
        }
    }
}

fn decrypt_segmented<F>(
    ciphertext: &[u8],
    total_size: u64,
    expected_len: usize,
    mut key_iv_for_block: F,
) -> Result<Vec<u8>, OfficeCryptoError>
where
    F: FnMut(u32) -> Result<(zeroize::Zeroizing<Vec<u8>>, Vec<u8>), OfficeCryptoError>,
{
    const SEGMENT_LEN: usize = 4096;
    let mut out = Vec::new();
    out.try_reserve_exact(ciphertext.len()).map_err(|source| {
        OfficeCryptoError::EncryptedPackageAllocationFailed { total_size, source }
    })?;
    let mut offset = 0usize;
    let mut block = 0u32;
    while offset < ciphertext.len() {
        let seg_len = (ciphertext.len() - offset).min(SEGMENT_LEN);
        let seg = &ciphertext[offset..offset + seg_len];
        let (key, iv) = key_iv_for_block(block)?;
        let mut plain = aes_cbc_decrypt(&key, &iv, seg)?;
        out.append(&mut plain);
        offset += seg_len;
        block = block.checked_add(1).ok_or_else(|| {
            OfficeCryptoError::InvalidFormat("segment counter overflow".to_string())
        })?;
    }
    if expected_len > out.len() {
        return Err(OfficeCryptoError::InvalidFormat(format!(
            "decrypted package length {} shorter than expected {}",
            out.len(),
            expected_len
        )));
    }
    out.truncate(expected_len);
    Ok(out)
}

#[cfg(test)]
pub(crate) mod tests {
    use crate::crypto::{aes_cbc_encrypt, HashAlgorithm, StandardKeyDeriver};
    use crate::util::parse_encryption_info_header;

    pub(crate) fn standard_encryption_info_fixture() -> Vec<u8> {
        // Minimal EncryptionInfo (Standard) fixture:
        // - Version 4.2
        // - AES-128 + SHA1
        // - Empty CSP name
        let version_major = 4u16;
        let version_minor = 2u16;
        let flags = 0x0000_0040u32;

        let header_flags = 0u32;
        let size_extra = 0u32;
        let alg_id = 0x0000_660Eu32; // CALG_AES_128
        let alg_id_hash = 0x0000_8004u32; // CALG_SHA1
        let key_bits = 128u32;
        let provider_type = 0x0000_0018u32; // PROV_RSA_AES
        let reserved1 = 0u32;
        let reserved2 = 0u32;
        let csp_name_utf16_nul = [0u8, 0u8];

        let mut header_bytes = Vec::new();
        header_bytes.extend_from_slice(&header_flags.to_le_bytes());
        header_bytes.extend_from_slice(&size_extra.to_le_bytes());
        header_bytes.extend_from_slice(&alg_id.to_le_bytes());
        header_bytes.extend_from_slice(&alg_id_hash.to_le_bytes());
        header_bytes.extend_from_slice(&key_bits.to_le_bytes());
        header_bytes.extend_from_slice(&provider_type.to_le_bytes());
        header_bytes.extend_from_slice(&reserved1.to_le_bytes());
        header_bytes.extend_from_slice(&reserved2.to_le_bytes());
        header_bytes.extend_from_slice(&csp_name_utf16_nul);

        let header_size = header_bytes.len() as u32;

        // Build a minimal verifier that will pass for password="Password".
        let password = "Password";
        let salt: [u8; 16] = [
            0x10, 0x11, 0x12, 0x13, 0x14, 0x15, 0x16, 0x17, 0x18, 0x19, 0x1A, 0x1B, 0x1C, 0x1D,
            0x1E, 0x1F,
        ];
        let verifier_plain: [u8; 16] = *b"formula-std-test";
        let verifier_hash = HashAlgorithm::Sha1.digest(&verifier_plain);

        let deriver = StandardKeyDeriver::new(HashAlgorithm::Sha1, key_bits, &salt, password);
        let key0 = deriver.derive_key_for_block(0).expect("key0");
        let iv = [0u8; 16];

        let encrypted_verifier =
            aes_cbc_encrypt(&key0, &iv, &verifier_plain).expect("encrypt verifier");
        let mut verifier_hash_padded = verifier_hash.clone();
        verifier_hash_padded.resize(32, 0);
        let encrypted_verifier_hash =
            aes_cbc_encrypt(&key0, &iv, &verifier_hash_padded).expect("encrypt verifier hash");

        let salt_size = salt.len() as u32;
        let verifier_hash_size = verifier_hash.len() as u32;

        let mut verifier_bytes = Vec::new();
        verifier_bytes.extend_from_slice(&salt_size.to_le_bytes());
        verifier_bytes.extend_from_slice(&salt);
        verifier_bytes.extend_from_slice(&encrypted_verifier);
        verifier_bytes.extend_from_slice(&verifier_hash_size.to_le_bytes());
        verifier_bytes.extend_from_slice(&encrypted_verifier_hash);

        let mut out = Vec::new();
        out.extend_from_slice(&version_major.to_le_bytes());
        out.extend_from_slice(&version_minor.to_le_bytes());
        out.extend_from_slice(&flags.to_le_bytes());
        out.extend_from_slice(&header_size.to_le_bytes());
        out.extend_from_slice(&header_bytes);
        out.extend_from_slice(&verifier_bytes);

        // Sanity: should parse as standard.
        let hdr = parse_encryption_info_header(&out).expect("header");
        assert_eq!(hdr.kind, crate::util::EncryptionInfoKind::Standard);
        out
    }
}
