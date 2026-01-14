use crate::crypto::{aes_cbc_decrypt, derive_iv, HashAlgorithm, StandardKeyDeriver};
use crate::error::OfficeCryptoError;
use crate::util::{
    checked_vec_len, ct_eq, decode_utf16le_nul_terminated, read_u32_le, read_u64_le,
    EncryptionInfoHeader,
};
use zeroize::Zeroizing;

// CryptoAPI algorithm identifiers (MS-OFFCRYPTO Standard / CryptoAPI encryption).
#[allow(dead_code)]
const CALG_RC4: u32 = 0x0000_6801;
const CALG_AES_128: u32 = 0x0000_660E;
const CALG_AES_192: u32 = 0x0000_660F;
const CALG_AES_256: u32 = 0x0000_6610;

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
        CALG_AES_128 | CALG_AES_192 | CALG_AES_256 => {}
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

/// Verify an MS-OFFCRYPTO Standard password by decrypting the `EncryptionVerifier` fields.
///
/// This is a lightweight check that does **not** require decrypting the full `EncryptedPackage`.
/// It is intended to be used as an early password validation step.
#[allow(dead_code)]
pub(crate) fn verify_password_standard(
    header: &EncryptionHeader,
    verifier: &EncryptionVerifier,
    password: &str,
) -> Result<(), OfficeCryptoError> {
    let hash_alg = HashAlgorithm::from_cryptoapi_alg_id_hash(header.alg_id_hash)?;
    let deriver = StandardKeyDeriver::new(hash_alg, header.key_bits, &verifier.salt, password);
    let key0 = deriver.derive_key_for_block(0)?;

    let expected_hash_len = verifier.verifier_hash_size as usize;
    let iv = [0u8; 16];

    let (verifier_plain, verifier_hash_plain_full) = match header.alg_id {
        CALG_RC4 => {
            if verifier.encrypted_verifier.len() != 16 {
                return Err(OfficeCryptoError::InvalidFormat(format!(
                    "EncryptionVerifier.encryptedVerifier must be 16 bytes for RC4 (got {})",
                    verifier.encrypted_verifier.len()
                )));
            }

            // RC4 is a stream cipher. CryptoAPI encrypts/decrypts the verifier and verifier hash
            // using the **same** RC4 stream (continuing the keystream), so we must apply RC4 to the
            // concatenated bytes rather than resetting the cipher per field.
            //
            // Additionally, CryptoAPI/Office represent "40-bit" RC4 keys as a 128-bit key where the
            // high 88 bits are zero (see `docs/offcrypto-standard-cryptoapi.md`).
            let rc4_key: Zeroizing<Vec<u8>> = if header.key_bits == 40 {
                if key0.len() != 5 {
                    return Err(OfficeCryptoError::InvalidFormat(format!(
                        "derived RC4 key for keySize=40 must be 5 bytes (got {})",
                        key0.len()
                    )));
                }
                let mut padded = vec![0u8; 16];
                padded[..5].copy_from_slice(&key0[..5]);
                Zeroizing::new(padded)
            } else {
                // Other key sizes are passed through directly.
                key0.clone()
            };

            let mut buf = Vec::with_capacity(
                verifier.encrypted_verifier.len() + verifier.encrypted_verifier_hash.len(),
            );
            buf.extend_from_slice(&verifier.encrypted_verifier);
            buf.extend_from_slice(&verifier.encrypted_verifier_hash);
            crate::crypto::rc4_xor_in_place(rc4_key.as_slice(), &mut buf)?;

            let verifier_plain = buf[..16].to_vec();
            let verifier_hash_plain = buf[16..].to_vec();
            (verifier_plain, verifier_hash_plain)
        }
        CALG_AES_128 | CALG_AES_192 | CALG_AES_256 => {
            let verifier_plain = aes_cbc_decrypt(key0.as_slice(), &iv, &verifier.encrypted_verifier)?;
            let verifier_hash_plain =
                aes_cbc_decrypt(key0.as_slice(), &iv, &verifier.encrypted_verifier_hash)?;
            (verifier_plain, verifier_hash_plain)
        }
        other => {
            return Err(OfficeCryptoError::UnsupportedEncryption(format!(
                "unsupported cipher AlgID {other:#x}"
            )))
        }
    };

    let verifier_hash_plain = verifier_hash_plain_full.get(..expected_hash_len).ok_or_else(|| {
        OfficeCryptoError::InvalidFormat(format!(
            "decrypted verifier hash shorter than verifierHashSize (got {}, need {})",
            verifier_hash_plain_full.len(),
            expected_hash_len
        ))
    })?;

    let verifier_hash = hash_alg.digest(verifier_plain.as_slice());
    let verifier_hash = verifier_hash.get(..expected_hash_len).ok_or_else(|| {
        OfficeCryptoError::InvalidFormat(format!(
            "hash output shorter than verifierHashSize (got {}, need {})",
            verifier_hash.len(),
            expected_hash_len
        ))
    })?;

    if !ct_eq(verifier_hash_plain, verifier_hash) {
        return Err(OfficeCryptoError::InvalidPassword);
    }

    Ok(())
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

    let verifier: Zeroizing<Vec<u8>> = Zeroizing::new(aes_cbc_decrypt(
        &verifier_key,
        &verifier_iv,
        &info.verifier.encrypted_verifier,
    )?);
    let verifier_hash_plain: Zeroizing<Vec<u8>> = Zeroizing::new(aes_cbc_decrypt(
        &verifier_key,
        &verifier_iv,
        &info.verifier.encrypted_verifier_hash,
    )?);
    let verifier_hash_plain = verifier_hash_plain
        .get(..info.verifier.verifier_hash_size as usize)
        .ok_or_else(|| {
            OfficeCryptoError::InvalidFormat(
                "EncryptionVerifier decrypted hash shorter than verifierHashSize".to_string(),
            )
        })?;

    let verifier_hash: Zeroizing<Vec<u8>> = Zeroizing::new(hash_alg.digest(verifier.as_slice()));
    if !ct_eq(verifier_hash_plain, verifier_hash.as_slice()) {
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
    use super::*;
    use crate::crypto::{
        aes_cbc_encrypt, hash_password, password_to_utf16le, rc4_xor_in_place, HashAlgorithm,
        StandardKeyDeriver,
    };
    use crate::util::{ct_eq_call_count, parse_encryption_info_header, reset_ct_eq_calls};

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

    #[test]
    fn standard_password_verifier_uses_constant_time_compare() {
        reset_ct_eq_calls();

        let encryption_info = standard_encryption_info_fixture();
        let header = parse_encryption_info_header(&encryption_info).expect("parse header");
        let info = parse_standard_encryption_info(&encryption_info, &header).expect("parse standard");

        // Only the 8-byte length prefix is required for this test because the wrong password fails
        // during verifier validation (before package decryption is attempted).
        let encrypted_package = [0u8; 8];
        let err =
            decrypt_standard_encrypted_package(&info, &encrypted_package, "wrong-password").expect_err("wrong pw");
        assert!(matches!(err, OfficeCryptoError::InvalidPassword));
        assert!(
            ct_eq_call_count() > 0,
            "expected constant-time compare helper to be invoked"
        );
    }

    fn derive_key_ref(
        hash_alg: HashAlgorithm,
        key_bits: u32,
        salt: &[u8],
        password: &str,
        block_index: u32,
    ) -> Vec<u8> {
        // Matches `StandardKeyDeriver` (50k spin) + CryptoAPI CryptDeriveKey behavior.
        let pw = password_to_utf16le(password);
        let pw_hash = hash_password(hash_alg, salt, &pw, 50_000);

        let mut buf = Vec::with_capacity(pw_hash.len() + 4);
        buf.extend_from_slice(&pw_hash);
        buf.extend_from_slice(&block_index.to_le_bytes());
        let h = hash_alg.digest(&buf);

        let key_len = (key_bits as usize) / 8;
        if key_len <= h.len() {
            return h[..key_len].to_vec();
        }

        // CryptoAPI `CryptDeriveKey` expansion used by MS-OFFCRYPTO Standard encryption: pad `h` to
        // 64 bytes with zeros, XOR with 0x36 and 0x5C, hash each, and concatenate.
        let mut buf = h.clone();
        buf.resize(64, 0);
        let mut ipad = vec![0u8; 64];
        let mut opad = vec![0u8; 64];
        for i in 0..64 {
            ipad[i] = buf[i] ^ 0x36;
            opad[i] = buf[i] ^ 0x5c;
        }
        let mut out = hash_alg.digest(&ipad);
        out.extend_from_slice(&hash_alg.digest(&opad));
        out.truncate(key_len);
        out
    }

    #[test]
    fn verify_password_standard_rc4_keysize_hash_matrix() {
        // Synthetic (password, salt, verifier) fixture used across the parameter matrix.
        let password = "correct horse battery staple";
        let wrong_password = "not the password";
        let salt: [u8; 16] = [
            0x10, 0x32, 0x54, 0x76, 0x98, 0xba, 0xdc, 0xfe, 0x01, 0x23, 0x45, 0x67, 0x89, 0xab,
            0xcd, 0xef,
        ];
        let verifier_plain: [u8; 16] = [
            0x00, 0x11, 0x22, 0x33, 0x44, 0x55, 0x66, 0x77, 0x88, 0x99, 0xaa, 0xbb, 0xcc, 0xdd,
            0xee, 0xff,
        ];

        for (hash_alg, alg_id_hash) in [
            (HashAlgorithm::Sha1, 0x0000_8004u32), // CALG_SHA1
            (HashAlgorithm::Md5, 0x0000_8003u32),  // CALG_MD5
        ] {
            for key_bits in [40u32, 56u32, 128u32] {
                // Derive the expected key for block 0 and validate truncation.
                let key_ref = derive_key_ref(hash_alg, key_bits, &salt, password, 0);
                let deriver = StandardKeyDeriver::new(hash_alg, key_bits, &salt, password);
                let key0 = deriver.derive_key_for_block(0).expect("derive key");
                assert_eq!(
                    key0.as_slice(),
                    key_ref.as_slice(),
                    "StandardKeyDeriver should match the reference derivation"
                );
                assert_eq!(
                    key_ref.len(),
                    (key_bits / 8) as usize,
                    "derived key length should be keySize/8"
                );
                assert!(
                    key_ref.len() <= hash_alg.digest_len(),
                    "for RC4, keyLen should be <= hashLen (truncation case)"
                );

                // Encrypt verifier + verifierHash with RC4 using a **single** stream that continues
                // across both fields (CryptoAPI behavior). For 40-bit keys, Office uses a 128-bit
                // key with the high bits zero.
                let verifier_hash = hash_alg.digest(&verifier_plain);

                let rc4_key = if key_bits == 40 {
                    let mut padded = vec![0u8; 16];
                    padded[..5].copy_from_slice(&key_ref[..5]);
                    padded
                } else {
                    key_ref.clone()
                };

                let mut buf = Vec::with_capacity(verifier_plain.len() + verifier_hash.len());
                buf.extend_from_slice(&verifier_plain);
                buf.extend_from_slice(&verifier_hash);
                rc4_xor_in_place(&rc4_key, &mut buf).expect("rc4 encrypt verifier+hash");

                let encrypted_verifier = buf[..16].to_vec();
                let encrypted_verifier_hash = buf[16..].to_vec();

                let header = EncryptionHeader {
                    alg_id: CALG_RC4,
                    alg_id_hash,
                    key_bits,
                    provider_type: 0,
                    csp_name: String::new(),
                };
                let verifier = EncryptionVerifier {
                    salt: salt.to_vec(),
                    encrypted_verifier,
                    verifier_hash_size: verifier_hash.len() as u32,
                    encrypted_verifier_hash,
                };

                verify_password_standard(&header, &verifier, password)
                    .expect("correct password should verify");
                let err = verify_password_standard(&header, &verifier, wrong_password)
                    .expect_err("wrong password should fail");
                assert!(matches!(err, OfficeCryptoError::InvalidPassword));
            }
        }
    }

    #[test]
    fn cryptderivekey_expansion_is_exercised_for_aes256_sha1() {
        let password = "correct horse battery staple";
        let wrong_password = "not the password";
        let salt: [u8; 16] = [0x42u8; 16];
        let verifier_plain: [u8; 16] = *b"formula-std-test";

        let key_bits = 256u32;
        let hash_alg = HashAlgorithm::Sha1;
        let key_ref = derive_key_ref(hash_alg, key_bits, &salt, password, 0);
        let deriver = StandardKeyDeriver::new(hash_alg, key_bits, &salt, password);
        let key0 = deriver.derive_key_for_block(0).expect("derive key");
        assert_eq!(key0.as_slice(), key_ref.as_slice());
        assert_eq!(key_ref.len(), 32);
        assert!(
            key_ref.len() > hash_alg.digest_len(),
            "AES-256+SHA1 should require CryptDeriveKey expansion"
        );

        // Encrypt verifier and verifierHash with AES-CBC, IV=0.
        let iv = [0u8; 16];
        let encrypted_verifier =
            aes_cbc_encrypt(&key_ref, &iv, &verifier_plain).expect("encrypt verifier");

        let verifier_hash = hash_alg.digest(&verifier_plain);
        let mut verifier_hash_padded = verifier_hash.clone();
        verifier_hash_padded.resize(32, 0);
        let encrypted_verifier_hash =
            aes_cbc_encrypt(&key_ref, &iv, &verifier_hash_padded).expect("encrypt verifier hash");

        let header = EncryptionHeader {
            alg_id: CALG_AES_256,
            alg_id_hash: 0x0000_8004u32, // CALG_SHA1
            key_bits,
            provider_type: 0,
            csp_name: String::new(),
        };
        let verifier = EncryptionVerifier {
            salt: salt.to_vec(),
            encrypted_verifier,
            verifier_hash_size: verifier_hash.len() as u32,
            encrypted_verifier_hash,
        };

        verify_password_standard(&header, &verifier, password).expect("correct password should verify");
        let err =
            verify_password_standard(&header, &verifier, wrong_password).expect_err("wrong password");
        assert!(matches!(err, OfficeCryptoError::InvalidPassword));
    }

    #[test]
    fn cryptderivekey_expansion_is_exercised_for_aes256_md5() {
        let password = "correct horse battery staple";
        let wrong_password = "not the password";
        let salt: [u8; 16] = [0x24u8; 16];
        let verifier_plain: [u8; 16] = *b"formula-std-md5!";

        let key_bits = 256u32;
        let hash_alg = HashAlgorithm::Md5;
        let key_ref = derive_key_ref(hash_alg, key_bits, &salt, password, 0);
        let deriver = StandardKeyDeriver::new(hash_alg, key_bits, &salt, password);
        let key0 = deriver.derive_key_for_block(0).expect("derive key");
        assert_eq!(key0.as_slice(), key_ref.as_slice());
        assert_eq!(key_ref.len(), 32);
        assert!(
            key_ref.len() > hash_alg.digest_len(),
            "AES-256+MD5 should require CryptDeriveKey expansion"
        );

        // Encrypt verifier and verifierHash with AES-CBC, IV=0. Pad the verifier hash ciphertext to
        // 2 blocks to ensure we only compare the first `verifierHashSize` bytes after decryption.
        let iv = [0u8; 16];
        let encrypted_verifier =
            aes_cbc_encrypt(&key_ref, &iv, &verifier_plain).expect("encrypt verifier");

        let verifier_hash = hash_alg.digest(&verifier_plain);
        let mut verifier_hash_padded = verifier_hash.clone();
        verifier_hash_padded.resize(32, 0);
        let encrypted_verifier_hash =
            aes_cbc_encrypt(&key_ref, &iv, &verifier_hash_padded).expect("encrypt verifier hash");

        let header = EncryptionHeader {
            alg_id: CALG_AES_256,
            alg_id_hash: 0x0000_8003u32, // CALG_MD5
            key_bits,
            provider_type: 0,
            csp_name: String::new(),
        };
        let verifier = EncryptionVerifier {
            salt: salt.to_vec(),
            encrypted_verifier,
            verifier_hash_size: verifier_hash.len() as u32,
            encrypted_verifier_hash,
        };

        verify_password_standard(&header, &verifier, password)
            .expect("correct password should verify");
        let err =
            verify_password_standard(&header, &verifier, wrong_password).expect_err("wrong password");
        assert!(matches!(err, OfficeCryptoError::InvalidPassword));
    }
} 
