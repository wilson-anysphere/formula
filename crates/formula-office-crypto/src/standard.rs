use rand::rngs::OsRng;
use rand::RngCore;
use zeroize::Zeroizing;

use crate::crypto::{
    aes_cbc_decrypt, aes_cbc_decrypt_in_place, aes_ecb_decrypt, aes_ecb_decrypt_in_place,
    aes_ecb_encrypt, derive_iv, rc4_xor_in_place, HashAlgorithm, StandardKeyDerivation,
    StandardKeyDeriver,
};
use crate::error::OfficeCryptoError;
use crate::util::{
    checked_vec_len, ct_eq, decode_utf16le_nul_terminated, parse_encrypted_package_original_size,
    read_u32_le, EncryptionInfoHeader,
};
// CryptoAPI algorithm identifiers (MS-OFFCRYPTO Standard / CryptoAPI encryption).
#[allow(dead_code)]
const CALG_RC4: u32 = 0x0000_6801;
const CALG_AES_128: u32 = 0x0000_660E;
const CALG_AES_192: u32 = 0x0000_660F;
const CALG_AES_256: u32 = 0x0000_6610;

/// Parsed MS-OFFCRYPTO `EncryptionHeader.Flags` bits for Standard (CryptoAPI) encryption.
///
/// The raw value may contain additional bits not currently modeled here.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct EncryptionHeaderFlags {
    raw: u32,
    f_cryptoapi: bool,
    #[allow(dead_code)]
    f_doc_props: bool,
    f_external: bool,
    f_aes: bool,
}

impl EncryptionHeaderFlags {
    const F_CRYPTOAPI: u32 = 0x0000_0004;
    const F_DOCPROPS: u32 = 0x0000_0008;
    const F_EXTERNAL: u32 = 0x0000_0010;
    const F_AES: u32 = 0x0000_0020;

    fn from_raw(raw: u32) -> Self {
        Self {
            raw,
            f_cryptoapi: raw & Self::F_CRYPTOAPI != 0,
            f_doc_props: raw & Self::F_DOCPROPS != 0,
            f_external: raw & Self::F_EXTERNAL != 0,
            f_aes: raw & Self::F_AES != 0,
        }
    }
}

fn is_aes_alg_id(alg_id: u32) -> bool {
    matches!(alg_id, CALG_AES_128 | CALG_AES_192 | CALG_AES_256)
}

/// Conservative upper bound on `EncryptionVerifier.saltSize` to avoid allocating attacker-controlled
/// buffers.
///
/// Office-produced Standard encryption uses 16-byte salts, but we accept other sizes within this
/// bound for robustness.
const MAX_VERIFIER_SALT_SIZE: usize = 1024;

/// Conservative upper bound on `EncryptionVerifier.verifierHashSize`.
///
/// Standard encryption uses CryptoAPI hash algorithms; the largest supported digest length in this
/// crate is SHA-512 (64 bytes).
const MAX_VERIFIER_HASH_SIZE: u32 = 64;

/// Standard/CryptoAPI RC4 `EncryptedPackage` block size.
const RC4_ENCRYPTED_PACKAGE_BLOCK_SIZE: usize = 0x200;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum StandardAesCipherMode {
    Ecb,
    Cbc { iv: [u8; 16] },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Rc4KeyStyle {
    /// Use the derived RC4 key bytes as-is (key length = `keySize/8`).
    Raw,
    /// For 40-bit RC4 (keyLen=5), use a 16-byte key where the remaining bytes are zero.
    Padded40Bit,
}

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
    if header_size > crate::MAX_STANDARD_ENCRYPTION_HEADER_BYTES {
        return Err(OfficeCryptoError::SizeLimitExceeded {
            context: "EncryptionInfo.headerSize",
            limit: crate::MAX_STANDARD_ENCRYPTION_HEADER_BYTES,
        });
    }
    let header_end = start.checked_add(header_size).ok_or_else(|| {
        OfficeCryptoError::InvalidFormat("EncryptionInfo header size overflow".to_string())
    })?;
    let header_bytes = bytes.get(start..header_end).ok_or_else(|| {
        OfficeCryptoError::InvalidFormat("EncryptionInfo header size out of range".to_string())
    })?;
    let verifier_bytes = bytes.get(header_end..).ok_or_else(|| {
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
    let flags_raw = read_u32_le(bytes, 0)?;
    let size_extra = read_u32_le(bytes, 4)?;
    let alg_id = read_u32_le(bytes, 8)?;
    let alg_id_hash = read_u32_le(bytes, 12)?;
    let mut key_bits = read_u32_le(bytes, 16)?;
    // MS-OFFCRYPTO specifies that for RC4, `keySize=0` MUST be interpreted as 40-bit.
    if alg_id == CALG_RC4 && key_bits == 0 {
        key_bits = 40;
    }
    let provider_type = read_u32_le(bytes, 20)?;
    let _reserved1 = read_u32_le(bytes, 24)?;
    let _reserved2 = read_u32_le(bytes, 28)?;
    // MS-OFFCRYPTO `EncryptionHeader` stores a UTF-16LE CSPName string followed by `sizeExtra`
    // opaque bytes. Some real-world files set `sizeExtra` to a non-zero (and potentially odd)
    // value. We must avoid decoding those trailing bytes as UTF-16LE.
    let tail_len = bytes.len() - 32;
    if (size_extra as usize) > tail_len {
        return Err(OfficeCryptoError::InvalidFormat(format!(
            "EncryptionHeader sizeExtra out of range (sizeExtra={size_extra}, tailLen={tail_len})"
        )));
    }
    let csp_len = tail_len - (size_extra as usize);
    let csp_bytes = bytes.get(32..32 + csp_len).ok_or_else(|| {
        OfficeCryptoError::InvalidFormat("EncryptionHeader CSPName out of range".to_string())
    })?;
    let csp_name = decode_utf16le_nul_terminated(csp_bytes)?;

    // Validate key `EncryptionHeader.Flags` semantics.
    //
    // Note: some real-world Standard-encrypted workbooks omit these flags (e.g. `flags_raw == 0`)
    // even when the cipher parameters clearly indicate CryptoAPI + AES. We treat the flags as
    // advisory and derive behavior from the algorithm identifiers instead of rejecting those
    // files outright.
    let flags = EncryptionHeaderFlags::from_raw(flags_raw);
    if flags.f_external {
        return Err(OfficeCryptoError::UnsupportedEncryption(
            "unsupported external Standard encryption (fExternal flag set)".to_string(),
        ));
    }

    // Be conservative when `fAES` is set: it must be consistent with the algId. When `fAES` is not
    // set we still accept AES algIds to handle producers that omit the flag.
    let alg_is_aes = is_aes_alg_id(alg_id);
    if flags.f_aes && !alg_is_aes {
        return Err(OfficeCryptoError::InvalidFormat(format!(
            "invalid Standard EncryptionHeader flags for algId {alg_id:#x}: flags={flags_raw:#x}"
        )));
    }

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
    let hash_alg = HashAlgorithm::from_cryptoapi_alg_id_hash(header.alg_id_hash)?;
    let salt_size = read_u32_le(bytes, 0)? as usize;
    if salt_size == 0 || salt_size > MAX_VERIFIER_SALT_SIZE {
        return Err(OfficeCryptoError::InvalidFormat(format!(
            "EncryptionVerifier saltSize {salt_size} is out of bounds (expected 1..={MAX_VERIFIER_SALT_SIZE})"
        )));
    }
    let mut offset = 4usize;
    let salt_end = offset.checked_add(salt_size).ok_or_else(|| {
        OfficeCryptoError::InvalidFormat("EncryptionVerifier salt size overflow".to_string())
    })?;
    let salt = bytes.get(offset..salt_end).ok_or_else(|| {
        OfficeCryptoError::InvalidFormat("EncryptionVerifier salt out of range".to_string())
    })?;
    offset = salt_end;

    let verifier_end = offset.checked_add(16).ok_or_else(|| {
        OfficeCryptoError::InvalidFormat("EncryptionVerifier verifier offset overflow".to_string())
    })?;
    let encrypted_verifier = bytes.get(offset..verifier_end).ok_or_else(|| {
        OfficeCryptoError::InvalidFormat("EncryptionVerifier missing verifier".to_string())
    })?;
    offset = verifier_end;

    let verifier_hash_size = read_u32_le(bytes, offset)?;
    offset = offset.checked_add(4).ok_or_else(|| {
        OfficeCryptoError::InvalidFormat(
            "EncryptionVerifier verifierHashSize offset overflow".to_string(),
        )
    })?;
    let hash_size = verifier_hash_size as usize;
    if hash_size > crate::MAX_STANDARD_VERIFIER_HASH_SIZE_BYTES {
        return Err(OfficeCryptoError::SizeLimitExceeded {
            context: "EncryptionVerifier.verifierHashSize",
            limit: crate::MAX_STANDARD_VERIFIER_HASH_SIZE_BYTES,
        });
    }

    if verifier_hash_size == 0 || verifier_hash_size > MAX_VERIFIER_HASH_SIZE {
        return Err(OfficeCryptoError::InvalidFormat(format!(
            "EncryptionVerifier verifierHashSize {verifier_hash_size} is out of bounds (expected 1..={MAX_VERIFIER_HASH_SIZE})"
        )));
    }
    if verifier_hash_size as usize != hash_alg.digest_len() {
        return Err(OfficeCryptoError::InvalidFormat(format!(
            "EncryptionVerifier verifierHashSize {verifier_hash_size} does not match {} digest length {}",
            hash_alg.as_ooxml_name(),
            hash_alg.digest_len()
        )));
    }

    // Ciphertext size depends on the cipher:
    // - RC4 is a stream cipher (no padding).
    // - AES uses 16-byte blocks and pads `encryptedVerifierHash` to a multiple of 16.
    let encrypted_hash_len = match header.alg_id {
        CALG_RC4 => verifier_hash_size as usize,
        CALG_AES_128 | CALG_AES_192 | CALG_AES_256 => padded_aes_len(verifier_hash_size as usize)?,
        other => {
            return Err(OfficeCryptoError::UnsupportedEncryption(format!(
                "unsupported cipher AlgID {other:#x}"
            )))
        }
    };
    let encrypted_verifier_hash =
        bytes
            .get(offset..offset + encrypted_hash_len)
            .ok_or_else(|| {
                OfficeCryptoError::InvalidFormat(
                    "EncryptionVerifier missing verifier hash".to_string(),
                )
            })?;

    Ok(EncryptionVerifier {
        salt: salt.to_vec(),
        encrypted_verifier: encrypted_verifier.to_vec(),
        verifier_hash_size,
        encrypted_verifier_hash: encrypted_verifier_hash.to_vec(),
    })
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

fn derive_standard_aes_key0_and_mode(
    header: &EncryptionHeader,
    verifier: &EncryptionVerifier,
    hash_alg: HashAlgorithm,
    password: &str,
) -> Result<(Zeroizing<Vec<u8>>, StandardAesCipherMode), OfficeCryptoError> {
    if header.key_bits % 8 != 0 {
        return Err(OfficeCryptoError::InvalidFormat(format!(
            "EncryptionHeader keyBits must be divisible by 8 (got {})",
            header.key_bits
        )));
    }
    let key_len = (header.key_bits / 8) as usize;

    // Standard AES key derivation varies in the wild:
    // - CryptoAPI `CryptDeriveKey` semantics (ipad/opad expansion) (common; matches `standard.xlsx`)
    // - MS-OFFCRYPTO `TruncateHash` semantics (some producers; matches `standard-basic.xlsm`)
    //
    // Try both to preserve fixture compatibility.
    let key_derivations = [
        StandardKeyDerivation::Aes,
        StandardKeyDerivation::AesTruncateHash,
    ];
    for derivation in key_derivations {
        let deriver = StandardKeyDeriver::new(
            hash_alg,
            header.key_bits,
            &verifier.salt,
            password,
            derivation,
        );
        let key0 = deriver.derive_key_for_block(0)?;
        match verify_password_standard_with_key_and_mode(
            header,
            verifier,
            hash_alg,
            key0.as_slice(),
        ) {
            Ok(mode) => return Ok((key0, mode)),
            Err(OfficeCryptoError::InvalidPassword) => continue,
            Err(other) => return Err(other),
        }
    }

    // Compatibility fallback: some producers appear to use the RC4-style key truncation derivation
    // even when AlgID indicates AES.
    //
    // This derivation can only produce `digest_len` bytes of key material. If the requested AES key
    // length is larger (e.g. AES-256 with SHA-1), skip the fallback and report an invalid password
    // instead of surfacing a confusing UnsupportedEncryption error.
    if key_len <= hash_alg.digest_len() {
        let deriver_rc4 = StandardKeyDeriver::new(
            hash_alg,
            header.key_bits,
            &verifier.salt,
            password,
            StandardKeyDerivation::Rc4,
        );
        let key0_rc4 = match deriver_rc4.derive_key_for_block(0) {
            Ok(key) => key,
            Err(OfficeCryptoError::UnsupportedEncryption(_)) => {
                return Err(OfficeCryptoError::InvalidPassword);
            }
            Err(e) => return Err(e),
        };
        let mode = verify_password_standard_with_key_and_mode(
            header,
            verifier,
            hash_alg,
            key0_rc4.as_slice(),
        )?;
        return Ok((key0_rc4, mode));
    }

    Err(OfficeCryptoError::InvalidPassword)
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
    match header.alg_id {
        CALG_RC4 => {
            let deriver = StandardKeyDeriver::new(
                hash_alg,
                header.key_bits,
                &verifier.salt,
                password,
                StandardKeyDerivation::Rc4,
            );
            let key0 = deriver.derive_key_for_block(0)?;
            let _ = verify_password_standard_rc4_key_style(
                header,
                verifier,
                hash_alg,
                key0.as_slice(),
            )?;
            Ok(())
        }
        CALG_AES_128 | CALG_AES_192 | CALG_AES_256 => {
            let _ = derive_standard_aes_key0_and_mode(header, verifier, hash_alg, password)?;
            Ok(())
        }
        other => Err(OfficeCryptoError::UnsupportedEncryption(format!(
            "unsupported cipher AlgID {other:#x}"
        ))),
    }
}

fn verify_password_standard_rc4_key_style(
    header: &EncryptionHeader,
    verifier: &EncryptionVerifier,
    hash_alg: HashAlgorithm,
    key0: &[u8],
) -> Result<Rc4KeyStyle, OfficeCryptoError> {
    match verify_password_standard_with_key_and_mode(header, verifier, hash_alg, key0) {
        Ok(_) => Ok(Rc4KeyStyle::Raw),
        Err(OfficeCryptoError::InvalidPassword) => {
            // Some real-world CryptoAPI RC4 producers treat 40-bit keys as a 16-byte key blob where
            // the remaining bytes are zero. Try that as a compatibility fallback.
            if key0.len() == 5 {
                let mut padded_key = [0u8; 16];
                padded_key[..5].copy_from_slice(key0);
                match verify_password_standard_with_key_and_mode(
                    header,
                    verifier,
                    hash_alg,
                    padded_key.as_slice(),
                ) {
                    Ok(_) => Ok(Rc4KeyStyle::Padded40Bit),
                    Err(OfficeCryptoError::InvalidPassword) => {
                        Err(OfficeCryptoError::InvalidPassword)
                    }
                    Err(e) => Err(e),
                }
            } else {
                Err(OfficeCryptoError::InvalidPassword)
            }
        }
        Err(e) => Err(e),
    }
}

fn verify_password_standard_with_key_and_mode(
    header: &EncryptionHeader,
    verifier: &EncryptionVerifier,
    hash_alg: HashAlgorithm,
    key0: &[u8],
) -> Result<StandardAesCipherMode, OfficeCryptoError> {
    let expected_hash_len = verifier.verifier_hash_size as usize;

    match header.alg_id {
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
            let mut buf = Vec::with_capacity(
                verifier.encrypted_verifier.len() + verifier.encrypted_verifier_hash.len(),
            );
            buf.extend_from_slice(&verifier.encrypted_verifier);
            buf.extend_from_slice(&verifier.encrypted_verifier_hash);

            rc4_xor_in_place(key0, &mut buf)?;

            let verifier_plain = buf.get(..16).ok_or_else(|| {
                OfficeCryptoError::InvalidFormat("RC4 verifier out of range".to_string())
            })?;
            let verifier_hash_plain_full = buf.get(16..).ok_or_else(|| {
                OfficeCryptoError::InvalidFormat("RC4 verifier hash out of range".to_string())
            })?;
            let verifier_hash_plain = verifier_hash_plain_full
                .get(..expected_hash_len)
                .ok_or_else(|| {
                    OfficeCryptoError::InvalidFormat(format!(
                        "decrypted verifier hash shorter than verifierHashSize (got {}, need {})",
                        verifier_hash_plain_full.len(),
                        expected_hash_len
                    ))
                })?;

            let verifier_hash = hash_alg.digest(&verifier_plain);
            let verifier_hash = verifier_hash.get(..expected_hash_len).ok_or_else(|| {
                OfficeCryptoError::InvalidFormat(format!(
                    "hash output shorter than verifierHashSize (got {}, need {})",
                    verifier_hash.len(),
                    expected_hash_len
                ))
            })?;

            if ct_eq(verifier_hash_plain, verifier_hash) {
                Ok(StandardAesCipherMode::Ecb)
            } else {
                Err(OfficeCryptoError::InvalidPassword)
            }
        }
        CALG_AES_128 | CALG_AES_192 | CALG_AES_256 => {
            // MS-OFFCRYPTO Standard AES verifier fields are decrypted with AES-ECB (no IV).
            let verifier_plain = aes_ecb_decrypt(key0, &verifier.encrypted_verifier)?;
            let verifier_hash_plain_full =
                aes_ecb_decrypt(key0, &verifier.encrypted_verifier_hash)?;

            let verifier_hash_plain = verifier_hash_plain_full
                .get(..expected_hash_len)
                .ok_or_else(|| {
                    OfficeCryptoError::InvalidFormat(format!(
                        "decrypted verifier hash shorter than verifierHashSize (got {}, need {})",
                        verifier_hash_plain_full.len(),
                        expected_hash_len
                    ))
                })?;

            let verifier_hash = hash_alg.digest(&verifier_plain);
            let verifier_hash = verifier_hash.get(..expected_hash_len).ok_or_else(|| {
                OfficeCryptoError::InvalidFormat(format!(
                    "hash output shorter than verifierHashSize (got {}, need {})",
                    verifier_hash.len(),
                    expected_hash_len
                ))
            })?;

            if ct_eq(verifier_hash_plain, verifier_hash) {
                return Ok(StandardAesCipherMode::Ecb);
            }

            // Compatibility fallback: some producers appear to use AES-CBC for the verifier fields
            // even though the canonical Standard/CryptoAPI fixture in this repo uses AES-ECB.
            //
            // Try a small set of plausible IVs and accept whichever yields a matching verifier
            // hash.
            let salt_iv = {
                let mut iv = [0u8; 16];
                let n = verifier.salt.len().min(16);
                iv[..n].copy_from_slice(&verifier.salt[..n]);
                iv
            };
            let derived_iv_vec =
                crate::crypto::derive_iv(hash_alg, &verifier.salt, &0u32.to_le_bytes(), 16);
            let mut derived_iv = [0u8; 16];
            derived_iv.copy_from_slice(&derived_iv_vec[..16]);
            let mut iv_candidates: Vec<[u8; 16]> = Vec::new();
            iv_candidates.push([0u8; 16]);
            if salt_iv != [0u8; 16] {
                iv_candidates.push(salt_iv);
            }
            if derived_iv != [0u8; 16] && derived_iv != salt_iv {
                iv_candidates.push(derived_iv);
            }
            for iv in iv_candidates.iter() {
                // 1) Attempt decrypting each field independently with the same IV.
                let verifier_plain = aes_cbc_decrypt(key0, iv, &verifier.encrypted_verifier)?;
                let verifier_hash_plain_full =
                    aes_cbc_decrypt(key0, iv, &verifier.encrypted_verifier_hash)?;
                let verifier_hash_plain =
                    verifier_hash_plain_full.get(..expected_hash_len).ok_or_else(|| {
                        OfficeCryptoError::InvalidFormat(format!(
                            "decrypted verifier hash shorter than verifierHashSize (got {}, need {})",
                            verifier_hash_plain_full.len(),
                            expected_hash_len
                        ))
                    })?;
                let verifier_hash = hash_alg.digest(&verifier_plain);
                let verifier_hash = verifier_hash.get(..expected_hash_len).ok_or_else(|| {
                    OfficeCryptoError::InvalidFormat(format!(
                        "hash output shorter than verifierHashSize (got {}, need {})",
                        verifier_hash.len(),
                        expected_hash_len
                    ))
                })?;
                if ct_eq(verifier_hash_plain, verifier_hash) {
                    return Ok(StandardAesCipherMode::Cbc { iv: *iv });
                }

                // 2) Some producers may encrypt the verifier + verifier hash as one CBC stream.
                // Attempt that as well (decrypt concatenation, then split).
                let mut concat = Vec::with_capacity(
                    verifier.encrypted_verifier.len() + verifier.encrypted_verifier_hash.len(),
                );
                concat.extend_from_slice(&verifier.encrypted_verifier);
                concat.extend_from_slice(&verifier.encrypted_verifier_hash);
                let concat_plain = aes_cbc_decrypt(key0, iv, &concat)?;
                let verifier_plain = concat_plain.get(..16).ok_or_else(|| {
                    OfficeCryptoError::InvalidFormat("CBC verifier out of range".to_string())
                })?;
                let verifier_hash_plain_full = concat_plain.get(16..).ok_or_else(|| {
                    OfficeCryptoError::InvalidFormat("CBC verifier hash out of range".to_string())
                })?;
                let verifier_hash_plain =
                    verifier_hash_plain_full.get(..expected_hash_len).ok_or_else(|| {
                        OfficeCryptoError::InvalidFormat(format!(
                            "decrypted verifier hash shorter than verifierHashSize (got {}, need {})",
                            verifier_hash_plain_full.len(),
                            expected_hash_len
                        ))
                    })?;
                let verifier_hash = hash_alg.digest(verifier_plain);
                let verifier_hash = verifier_hash.get(..expected_hash_len).ok_or_else(|| {
                    OfficeCryptoError::InvalidFormat(format!(
                        "hash output shorter than verifierHashSize (got {}, need {})",
                        verifier_hash.len(),
                        expected_hash_len
                    ))
                })?;
                if ct_eq(verifier_hash_plain, verifier_hash) {
                    return Ok(StandardAesCipherMode::Cbc { iv: *iv });
                }
            }

            Err(OfficeCryptoError::InvalidPassword)
        }
        other => Err(OfficeCryptoError::UnsupportedEncryption(format!(
            "unsupported cipher AlgID {other:#x}"
        ))),
    }
}

/// Verify the Standard/CryptoAPI password using the *spec* AES-ECB verifier mode only.
///
/// Some producers encrypt verifier fields with AES-CBC; callers that need that compatibility should
/// use `verify_password_standard_with_key_and_mode` or their own scheme-specific verification.
#[allow(dead_code)]
fn verify_password_standard_with_key(
    header: &EncryptionHeader,
    verifier: &EncryptionVerifier,
    hash_alg: HashAlgorithm,
    key0: &[u8],
) -> Result<(), OfficeCryptoError> {
    match verify_password_standard_with_key_and_mode(header, verifier, hash_alg, key0) {
        Ok(StandardAesCipherMode::Ecb) => Ok(()),
        Ok(StandardAesCipherMode::Cbc { .. }) => Err(OfficeCryptoError::InvalidPassword),
        Err(e) => Err(e),
    }
}

pub(crate) fn decrypt_standard_encrypted_package(
    info: &StandardEncryptionInfo,
    encrypted_package: &[u8],
    password: &str,
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
    // untrusted and must not be used to drive allocations without plausibility checks.
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

    let hash_alg = HashAlgorithm::from_cryptoapi_alg_id_hash(info.header.alg_id_hash)?;

    match info.header.alg_id {
        CALG_RC4 => {
            // Standard RC4 has no padding; ciphertext must contain at least `expected_len` bytes.
            // Check this before running the password KDF to reject obviously truncated inputs.
            if ciphertext.len() < expected_len {
                return Err(OfficeCryptoError::InvalidFormat(format!(
                    "EncryptedPackage ciphertext truncated (len {}, expected at least {})",
                    ciphertext.len(),
                    expected_len
                )));
            }
            decrypt_standard_encrypted_package_rc4(
                info,
                ciphertext,
                total_size,
                expected_len,
                password,
                hash_alg,
            )
        }
        CALG_AES_128 | CALG_AES_192 | CALG_AES_256 => {
            // The encrypted payload is padded to the AES block size (16 bytes). Some producers may
            // include trailing bytes in the OLE stream beyond the padded plaintext length; ignore
            // them by decrypting only what we need.
            //
            // Check ciphertext length *before* running the password KDF to reject obviously
            // truncated inputs cheaply.
            let padded_len = if expected_len == 0 {
                0usize
            } else {
                expected_len.checked_add(15).ok_or_else(|| {
                    OfficeCryptoError::InvalidFormat(
                        "EncryptedPackage expected length overflow".to_string(),
                    )
                })? / 16
                    * 16
            };
            if ciphertext.len() < padded_len {
                return Err(OfficeCryptoError::InvalidFormat(format!(
                    "EncryptedPackage ciphertext truncated (len {}, expected at least {})",
                    ciphertext.len(),
                    padded_len
                )));
            }

            let (key0, mode) = derive_standard_aes_key0_and_mode(
                &info.header,
                &info.verifier,
                hash_alg,
                password,
            )?;

            let to_decrypt = &ciphertext[..padded_len];
            let mut plain = Vec::new();
            plain.try_reserve_exact(padded_len).map_err(|source| {
                OfficeCryptoError::EncryptedPackageAllocationFailed { total_size, source }
            })?;
            plain.extend_from_slice(to_decrypt);
            match mode {
                StandardAesCipherMode::Ecb => {
                    aes_ecb_decrypt_in_place(key0.as_slice(), &mut plain)?;
                }
                StandardAesCipherMode::Cbc { iv } => {
                    aes_cbc_decrypt_in_place(key0.as_slice(), &iv, &mut plain)?;
                }
            }
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
        other => Err(OfficeCryptoError::UnsupportedEncryption(format!(
            "unsupported cipher AlgID {other:#x} for EncryptedPackage"
        ))),
    }
}

fn decrypt_standard_encrypted_package_rc4(
    info: &StandardEncryptionInfo,
    ciphertext: &[u8],
    total_size: u64,
    expected_len: usize,
    password: &str,
    hash_alg: HashAlgorithm,
) -> Result<Vec<u8>, OfficeCryptoError> {
    // Standard/CryptoAPI RC4 uses 0x200-byte blocks with per-block keys derived from the password
    // hash + block index.

    // MS-OFFCRYPTO: a keySize of 0 must be interpreted as 40-bit RC4.
    let effective_key_bits = if info.header.key_bits == 0 {
        40
    } else {
        info.header.key_bits
    };

    if effective_key_bits % 8 != 0 {
        return Err(OfficeCryptoError::InvalidFormat(format!(
            "EncryptionHeader keyBits must be divisible by 8 (got {})",
            info.header.key_bits
        )));
    }
    let key_len = (effective_key_bits / 8) as usize;
    if !matches!(key_len, 5 | 7 | 16) {
        return Err(OfficeCryptoError::UnsupportedEncryption(format!(
            "unsupported RC4 key length {key_len} bytes (keyBits={})",
            info.header.key_bits
        )));
    }

    let deriver = StandardKeyDeriver::new(
        hash_alg,
        info.header.key_bits,
        &info.verifier.salt,
        password,
        StandardKeyDerivation::Rc4,
    );
    let key0 = deriver.derive_key_for_block(0)?;
    let key_style = verify_password_standard_rc4_key_style(
        &info.header,
        &info.verifier,
        hash_alg,
        key0.as_slice(),
    )?;

    // Password is valid; decrypt the package. Standard RC4 has no padding; treat trailing bytes
    // (OLE sector slack) as irrelevant.
    if ciphertext.len() < expected_len {
        return Err(OfficeCryptoError::InvalidFormat(format!(
            "EncryptedPackage ciphertext truncated (len {}, expected at least {})",
            ciphertext.len(),
            expected_len
        )));
    }

    let mut out = Vec::new();
    out.try_reserve_exact(expected_len).map_err(|source| {
        OfficeCryptoError::EncryptedPackageAllocationFailed { total_size, source }
    })?;
    out.extend_from_slice(&ciphertext[..expected_len]);

    let mut block_index: u32 = 0;
    for chunk in out.chunks_mut(RC4_ENCRYPTED_PACKAGE_BLOCK_SIZE) {
        let key = deriver.derive_key_for_block(block_index)?;
        match key_style {
            Rc4KeyStyle::Raw => {
                rc4_xor_in_place(&key, chunk)?;
            }
            Rc4KeyStyle::Padded40Bit => {
                let mut padded_key = [0u8; 16];
                padded_key[..5].copy_from_slice(&key);
                rc4_xor_in_place(padded_key.as_slice(), chunk)?;
            }
        }
        block_index = block_index.checked_add(1).ok_or_else(|| {
            OfficeCryptoError::InvalidFormat("RC4 block index overflow".to_string())
        })?;
    }

    Ok(out)
}

#[allow(dead_code)]
#[derive(Debug, Clone, Copy)]
enum StandardScheme {
    /// Standard / ECMA-376: AES-ECB using the block-0 file key (no IV).
    ///
    /// This matches `msoffcrypto-tool`'s Standard decryptor and decrypts the repo's
    /// `fixtures/encrypted/ooxml/standard.xlsx`.
    Ecb,
    /// Segment the data in 4096-byte chunks; for chunk N use a key derived with blockIndex=N and
    /// IV=0.
    PerBlockKeyIvZero,
    /// Segment the data in 4096-byte chunks; use a single key derived with blockIndex=0 and IV
    /// derived as hash(salt||blockIndex) for each chunk.
    ConstKeyPerBlockIvHash,
    /// Treat the ciphertext as a single AES-CBC stream using key(block=0) and IV=salt.
    ConstKeyIvSaltStream,
}

#[allow(dead_code)]
fn decrypt_standard_with_scheme(
    info: &StandardEncryptionInfo,
    ciphertext: &[u8],
    total_size: u64,
    expected_len: usize,
    password: &str,
    hash_alg: HashAlgorithm,
    derivation: StandardKeyDerivation,
    scheme: StandardScheme,
) -> Result<Vec<u8>, OfficeCryptoError> {
    let deriver = StandardKeyDeriver::new(
        hash_alg,
        info.header.key_bits,
        &info.verifier.salt,
        password,
        derivation,
    );
    let key0 = deriver.derive_key_for_block(0)?;

    // Verify password using `EncryptionVerifier`.
    //
    // MS-OFFCRYPTO Standard AES specifies that verifier fields are encrypted with AES-ECB (no IV),
    // but some producers appear to use AES-CBC for the verifier as well. To maximize
    // compatibility, we try the spec ECB verifier first, then fall back to scheme-specific CBC
    // verification for the non-ECB variants.
    let mut verifier_ok = match verify_password_standard_with_key_and_mode(
        &info.header,
        &info.verifier,
        hash_alg,
        key0.as_slice(),
    ) {
        Ok(_) => true,
        Err(OfficeCryptoError::InvalidPassword) => false,
        Err(other) => return Err(other),
    };

    if !verifier_ok && !matches!(scheme, StandardScheme::Ecb) {
        let expected_hash_len = info.verifier.verifier_hash_size as usize;

        let verifier_iv: Vec<u8> = match scheme {
            StandardScheme::PerBlockKeyIvZero => vec![0u8; 16],
            StandardScheme::ConstKeyPerBlockIvHash => {
                derive_iv(hash_alg, &info.verifier.salt, &0u32.to_le_bytes(), 16)
            }
            StandardScheme::ConstKeyIvSaltStream => {
                let iv = info.verifier.salt.get(..16).ok_or_else(|| {
                    OfficeCryptoError::InvalidFormat(
                        "EncryptionVerifier salt too short".to_string(),
                    )
                })?;
                iv.to_vec()
            }
            StandardScheme::Ecb => unreachable!("guarded by !matches!(scheme, Ecb)"),
        };

        let verifier_plain = aes_cbc_decrypt(
            key0.as_slice(),
            &verifier_iv,
            &info.verifier.encrypted_verifier,
        )?;
        let verifier_hash_plain_full = aes_cbc_decrypt(
            key0.as_slice(),
            &verifier_iv,
            &info.verifier.encrypted_verifier_hash,
        )?;

        let verifier_hash_plain = verifier_hash_plain_full
            .get(..expected_hash_len)
            .ok_or_else(|| {
                OfficeCryptoError::InvalidFormat(format!(
                    "decrypted verifier hash shorter than verifierHashSize (got {}, need {})",
                    verifier_hash_plain_full.len(),
                    expected_hash_len
                ))
            })?;

        let verifier_hash = hash_alg.digest(&verifier_plain);
        let verifier_hash = verifier_hash.get(..expected_hash_len).ok_or_else(|| {
            OfficeCryptoError::InvalidFormat(format!(
                "hash output shorter than verifierHashSize (got {}, need {})",
                verifier_hash.len(),
                expected_hash_len
            ))
        })?;

        verifier_ok = ct_eq(verifier_hash_plain, verifier_hash);
    }

    if !verifier_ok {
        return Err(OfficeCryptoError::InvalidPassword);
    }

    // Password is valid; decrypt the package.
    match scheme {
        StandardScheme::Ecb => {
            // The encrypted payload is padded to the AES block size (16 bytes). Some producers may
            // include trailing bytes in the OLE stream beyond the padded plaintext length; ignore
            // them by decrypting only what we need.
            let padded_len = if expected_len == 0 {
                0usize
            } else {
                expected_len.checked_add(15).ok_or_else(|| {
                    OfficeCryptoError::InvalidFormat(
                        "EncryptedPackage expected length overflow".to_string(),
                    )
                })? / 16
                    * 16
            };
            if ciphertext.len() < padded_len {
                return Err(OfficeCryptoError::InvalidFormat(format!(
                    "EncryptedPackage ciphertext truncated (len {}, expected at least {})",
                    ciphertext.len(),
                    padded_len
                )));
            }
            let to_decrypt = &ciphertext[..padded_len];
            let mut plain = aes_ecb_decrypt(key0.as_slice(), to_decrypt)?;
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
        StandardScheme::PerBlockKeyIvZero => {
            decrypt_segmented(ciphertext, total_size, expected_len, |block| {
                let key = deriver.derive_key_for_block(block)?;
                Ok((key, vec![0u8; 16]))
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
            let mut plain = aes_cbc_decrypt(key0.as_slice(), iv, ciphertext)?;
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

#[allow(dead_code)]
fn decrypt_segmented<F>(
    ciphertext: &[u8],
    total_size: u64,
    expected_len: usize,
    mut key_iv_for_block: F,
) -> Result<Vec<u8>, OfficeCryptoError>
where
    F: FnMut(u32) -> Result<(Zeroizing<Vec<u8>>, Vec<u8>), OfficeCryptoError>,
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
        let mut plain = aes_cbc_decrypt(key.as_slice(), &iv, seg)?;
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

pub(crate) fn encrypt_standard_encrypted_package(
    zip_bytes: &[u8],
    password: &str,
    opts: &crate::EncryptOptions,
) -> Result<(Vec<u8>, Vec<u8>), OfficeCryptoError> {
    if opts.hash_algorithm != HashAlgorithm::Sha1 {
        return Err(OfficeCryptoError::InvalidOptions(
            "Standard encryption requires SHA1 (CryptoAPI)".to_string(),
        ));
    }
    if opts.spin_count != 50_000 {
        return Err(OfficeCryptoError::InvalidOptions(format!(
            "Standard encryption uses a fixed spin_count=50_000 (got {})",
            opts.spin_count
        )));
    }

    let key_bits_u32 = u32::try_from(opts.key_bits).map_err(|_| {
        OfficeCryptoError::InvalidOptions("key_bits does not fit in u32".to_string())
    })?;
    let alg_id: u32 = match key_bits_u32 {
        128 => 0x0000_660E, // CALG_AES_128
        192 => 0x0000_660F, // CALG_AES_192
        256 => 0x0000_6610, // CALG_AES_256
        other => {
            return Err(OfficeCryptoError::InvalidOptions(format!(
                "unsupported key_bits {other} for Standard encryption (expected 128/192/256)"
            )))
        }
    };
    let alg_id_hash = 0x0000_8004u32; // CALG_SHA1

    // CryptoAPI parameters (Excel-compatible).
    let provider_type = 0x0000_0018u32; // PROV_RSA_AES
    let csp_name = "Microsoft Enhanced RSA and AES Cryptographic Provider";

    // Random verifier salt and verifier bytes.
    let mut salt = [0u8; 16];
    OsRng.fill_bytes(&mut salt);

    let deriver = StandardKeyDeriver::new(
        HashAlgorithm::Sha1,
        key_bits_u32,
        &salt,
        password,
        StandardKeyDerivation::Aes,
    );
    let key0 = deriver.derive_key_for_block(0)?;

    let mut verifier_plain = [0u8; 16];
    OsRng.fill_bytes(&mut verifier_plain);
    let verifier_hash = HashAlgorithm::Sha1.digest(&verifier_plain);
    let verifier_hash_size = verifier_hash.len() as u32;
    let verifier_hash_padded = pad_zero(&verifier_hash, 16);

    // Standard encryption verifier fields are encrypted with AES-ECB (no IV).
    let encrypted_verifier = aes_ecb_encrypt(&key0, &verifier_plain)?;
    let encrypted_verifier_hash = aes_ecb_encrypt(&key0, &verifier_hash_padded)?;

    let mut verifier_bytes = Vec::new();
    verifier_bytes.extend_from_slice(&(salt.len() as u32).to_le_bytes());
    verifier_bytes.extend_from_slice(&salt);
    verifier_bytes.extend_from_slice(&encrypted_verifier);
    verifier_bytes.extend_from_slice(&verifier_hash_size.to_le_bytes());
    verifier_bytes.extend_from_slice(&encrypted_verifier_hash);

    // EncryptionHeader (see MS-OFFCRYPTO).
    let header_flags = EncryptionHeaderFlags::F_CRYPTOAPI | EncryptionHeaderFlags::F_AES;
    let size_extra = 0u32;
    let reserved1 = 0u32;
    let reserved2 = 0u32;
    let csp_name_utf16_nul = encode_utf16le_nul_terminated(csp_name);

    let mut header_bytes = Vec::new();
    header_bytes.extend_from_slice(&header_flags.to_le_bytes());
    header_bytes.extend_from_slice(&size_extra.to_le_bytes());
    header_bytes.extend_from_slice(&alg_id.to_le_bytes());
    header_bytes.extend_from_slice(&alg_id_hash.to_le_bytes());
    header_bytes.extend_from_slice(&key_bits_u32.to_le_bytes());
    header_bytes.extend_from_slice(&provider_type.to_le_bytes());
    header_bytes.extend_from_slice(&reserved1.to_le_bytes());
    header_bytes.extend_from_slice(&reserved2.to_le_bytes());
    header_bytes.extend_from_slice(&csp_name_utf16_nul);
    let header_size = header_bytes.len() as u32;

    // EncryptionInfo header. We use version 3.2 since it's commonly accepted by Office tooling.
    let version_major = 3u16;
    let version_minor = 2u16;
    let flags = 0x0000_0040u32;

    let mut encryption_info = Vec::new();
    encryption_info.extend_from_slice(&version_major.to_le_bytes());
    encryption_info.extend_from_slice(&version_minor.to_le_bytes());
    encryption_info.extend_from_slice(&flags.to_le_bytes());
    encryption_info.extend_from_slice(&header_size.to_le_bytes());
    encryption_info.extend_from_slice(&header_bytes);
    encryption_info.extend_from_slice(&verifier_bytes);

    // Standard encryption encrypts the full package with AES-ECB using key(block=0), with
    // zero-padding to a multiple of 16 bytes.
    let mut encrypted_package = Vec::new();
    encrypted_package.extend_from_slice(&(zip_bytes.len() as u64).to_le_bytes());
    let padded = pad_zero(zip_bytes, 16);
    let enc = aes_ecb_encrypt(&key0, &padded)?;
    encrypted_package.extend_from_slice(&enc);

    Ok((encryption_info, encrypted_package))
}

fn encode_utf16le_nul_terminated(s: &str) -> Vec<u8> {
    let mut out = Vec::with_capacity(s.len() * 2 + 2);
    for cu in s.encode_utf16() {
        out.extend_from_slice(&cu.to_le_bytes());
    }
    out.extend_from_slice(&0u16.to_le_bytes());
    out
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

#[cfg(test)]
pub(crate) mod tests {
    use super::*;
    use crate::crypto::{
        aes_ecb_encrypt, hash_password, password_to_utf16le, rc4_xor_in_place, HashAlgorithm,
        StandardKeyDerivation, StandardKeyDeriver,
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

        let header_flags = EncryptionHeaderFlags::F_CRYPTOAPI | EncryptionHeaderFlags::F_AES;
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

        let deriver = StandardKeyDeriver::new(
            HashAlgorithm::Sha1,
            key_bits,
            &salt,
            password,
            StandardKeyDerivation::Aes,
        );
        let key0 = deriver.derive_key_for_block(0).expect("key0");

        let encrypted_verifier = aes_ecb_encrypt(&key0, &verifier_plain).expect("encrypt verifier");
        let mut verifier_hash_padded = verifier_hash.clone();
        verifier_hash_padded.resize(32, 0);
        let encrypted_verifier_hash =
            aes_ecb_encrypt(&key0, &verifier_hash_padded).expect("encrypt verifier hash");

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

    fn minimal_encryption_header_bytes(flags: u32, alg_id: u32) -> Vec<u8> {
        let size_extra = 0u32;
        let alg_id_hash = 0x0000_8004u32; // CALG_SHA1
        let key_bits = 128u32;
        let provider_type = 0u32;
        let reserved1 = 0u32;
        let reserved2 = 0u32;
        let csp_name_utf16_nul = [0u8, 0u8];

        let mut header_bytes = Vec::new();
        header_bytes.extend_from_slice(&flags.to_le_bytes());
        header_bytes.extend_from_slice(&size_extra.to_le_bytes());
        header_bytes.extend_from_slice(&alg_id.to_le_bytes());
        header_bytes.extend_from_slice(&alg_id_hash.to_le_bytes());
        header_bytes.extend_from_slice(&key_bits.to_le_bytes());
        header_bytes.extend_from_slice(&provider_type.to_le_bytes());
        header_bytes.extend_from_slice(&reserved1.to_le_bytes());
        header_bytes.extend_from_slice(&reserved2.to_le_bytes());
        header_bytes.extend_from_slice(&csp_name_utf16_nul);
        header_bytes
    }

    #[test]
    fn standard_encryption_header_ignores_sizeextra_trailing_bytes() {
        let flags = EncryptionHeaderFlags::F_CRYPTOAPI | EncryptionHeaderFlags::F_AES;
        let size_extra = 1u32;
        let alg_id = CALG_AES_128;
        let alg_id_hash = 0x0000_8004u32; // CALG_SHA1
        let key_bits = 128u32;
        let provider_type = 0x0000_0018u32; // PROV_RSA_AES
        let reserved1 = 0u32;
        let reserved2 = 0u32;

        let expected_csp_name = "Test CSP";
        let csp_name_utf16_nul = encode_utf16le_nul_terminated(expected_csp_name);

        let mut header_bytes = Vec::new();
        header_bytes.extend_from_slice(&flags.to_le_bytes());
        header_bytes.extend_from_slice(&size_extra.to_le_bytes());
        header_bytes.extend_from_slice(&alg_id.to_le_bytes());
        header_bytes.extend_from_slice(&alg_id_hash.to_le_bytes());
        header_bytes.extend_from_slice(&key_bits.to_le_bytes());
        header_bytes.extend_from_slice(&provider_type.to_le_bytes());
        header_bytes.extend_from_slice(&reserved1.to_le_bytes());
        header_bytes.extend_from_slice(&reserved2.to_le_bytes());
        header_bytes.extend_from_slice(&csp_name_utf16_nul);
        header_bytes.push(0xAA); // trailing opaque sizeExtra byte (odd length total)

        let header = parse_encryption_header(&header_bytes).expect("parse header");
        assert_eq!(header.csp_name, expected_csp_name);
    }

    #[test]
    fn standard_encryption_header_rejects_external_flag() {
        let flags = EncryptionHeaderFlags::F_CRYPTOAPI | EncryptionHeaderFlags::F_EXTERNAL;
        let bytes = minimal_encryption_header_bytes(flags, CALG_RC4);
        let err = parse_encryption_header(&bytes).expect_err("expected error");
        assert!(
            matches!(
                err,
                OfficeCryptoError::UnsupportedEncryption(ref msg) if msg.contains("external")
            ),
            "unexpected error: {err:?}"
        );
    }

    #[test]
    fn standard_encryption_header_accepts_missing_cryptoapi_flag() {
        let flags = 0u32;
        let bytes = minimal_encryption_header_bytes(flags, CALG_RC4);
        let header = parse_encryption_header(&bytes).expect("parse header");
        assert_eq!(header.alg_id, CALG_RC4);
    }

    #[test]
    fn standard_encryption_header_accepts_aes_algid_without_faes_flag() {
        let flags = 0u32;
        let bytes = minimal_encryption_header_bytes(flags, CALG_AES_128);
        let header = parse_encryption_header(&bytes).expect("parse header");
        assert_eq!(header.alg_id, CALG_AES_128);
    }

    #[test]
    fn standard_encryption_header_rejects_faes_flag_with_rc4_algid() {
        let flags = EncryptionHeaderFlags::F_CRYPTOAPI | EncryptionHeaderFlags::F_AES;
        let bytes = minimal_encryption_header_bytes(flags, CALG_RC4);
        let err = parse_encryption_header(&bytes).expect_err("expected error");
        assert!(matches!(err, OfficeCryptoError::InvalidFormat(_)));
    }

    #[test]
    fn standard_encryption_header_rc4_keysize_zero_is_interpreted_as_40bit() {
        // MS-OFFCRYPTO specifies that for RC4, `keySize=0` MUST be interpreted as 40-bit.
        let flags = EncryptionHeaderFlags::F_CRYPTOAPI;
        let mut bytes = minimal_encryption_header_bytes(flags, CALG_RC4);
        bytes[16..20].copy_from_slice(&0u32.to_le_bytes()); // keyBits offset

        let header = parse_encryption_header(&bytes).expect("parse header");
        assert_eq!(header.alg_id, CALG_RC4);
        assert_eq!(header.key_bits, 40);
    }

    #[test]
    fn standard_password_verifier_uses_constant_time_compare() {
        reset_ct_eq_calls();

        let encryption_info = standard_encryption_info_fixture();
        let header = parse_encryption_info_header(&encryption_info).expect("parse header");
        let info =
            parse_standard_encryption_info(&encryption_info, &header).expect("parse standard");

        // Only the 8-byte length prefix is required for this test because the wrong password fails
        // during verifier validation (before package decryption is attempted).
        let encrypted_package = [0u8; 8];
        let err = decrypt_standard_encrypted_package(&info, &encrypted_package, "wrong-password")
            .expect_err("wrong pw");
        assert!(matches!(err, OfficeCryptoError::InvalidPassword));
        assert!(
            ct_eq_call_count() > 0,
            "expected constant-time compare helper to be invoked"
        );
    }

    #[test]
    fn standard_rc4_password_verifier_uses_constant_time_compare() {
        reset_ct_eq_calls();

        // Build a minimal RC4 verifier and ensure the wrong-password path still uses `ct_eq` for
        // digest comparison.
        let password = "Password";
        let wrong_password = "wrong-password";
        let salt: [u8; 16] = [0x11u8; 16];
        let verifier_plain: [u8; 16] = *b"formula-rc4-test";

        let hash_alg = HashAlgorithm::Sha1;
        let key_bits = 128u32;
        let alg_id_hash = 0x0000_8004u32; // CALG_SHA1

        let key0 = derive_key_ref(
            hash_alg,
            key_bits,
            &salt,
            password,
            0,
            StandardKeyDerivation::Rc4,
        );

        let verifier_hash = hash_alg.digest(&verifier_plain);
        let mut verifier_buf = Vec::with_capacity(verifier_plain.len() + verifier_hash.len());
        verifier_buf.extend_from_slice(&verifier_plain);
        verifier_buf.extend_from_slice(&verifier_hash);
        rc4_xor_in_place(&key0, &mut verifier_buf).expect("rc4 encrypt verifier");

        let header = EncryptionHeader {
            alg_id: CALG_RC4,
            alg_id_hash,
            key_bits,
            provider_type: 0,
            csp_name: String::new(),
        };
        let verifier = EncryptionVerifier {
            salt: salt.to_vec(),
            encrypted_verifier: verifier_buf[..16].to_vec(),
            verifier_hash_size: verifier_hash.len() as u32,
            encrypted_verifier_hash: verifier_buf[16..].to_vec(),
        };

        let err =
            verify_password_standard(&header, &verifier, wrong_password).expect_err("wrong pw");
        assert!(matches!(err, OfficeCryptoError::InvalidPassword));
        assert!(
            ct_eq_call_count() > 0,
            "expected constant-time compare helper to be invoked"
        );
    }

    #[test]
    fn decrypt_rejects_oversized_encrypted_package_original_size() {
        let info_bytes = standard_encryption_info_fixture();
        let hdr = parse_encryption_info_header(&info_bytes).expect("header");
        let info =
            super::parse_standard_encryption_info(&info_bytes, &hdr).expect("parse standard");

        let mut encrypted_package = Vec::new();
        encrypted_package
            .extend_from_slice(&(crate::MAX_ENCRYPTED_PACKAGE_ORIGINAL_SIZE + 1).to_le_bytes());
        let err = super::decrypt_standard_encrypted_package(&info, &encrypted_package, "Password")
            .expect_err("expected size limit error");

        assert!(
            matches!(err, OfficeCryptoError::SizeLimitExceededU64 { .. }),
            "err={err:?}"
        );
    }

    fn derive_key_ref(
        hash_alg: HashAlgorithm,
        key_bits: u32,
        salt: &[u8],
        password: &str,
        block_index: u32,
        derivation: StandardKeyDerivation,
    ) -> Vec<u8> {
        // Matches `StandardKeyDeriver` (50k spin) + MS-OFFCRYPTO Standard derivation:
        // - RC4: key = H_final[..key_len] (truncation)
        // - AES: CryptoAPI `CryptDeriveKey` (ipad/opad expansion) even when key_len <= hash_len
        // - AES (non-standard): some producers truncate H_final directly, same as RC4
        let pw = password_to_utf16le(password);
        let pw_hash = hash_password(hash_alg, salt, &pw, 50_000);

        let mut buf = Vec::with_capacity(pw_hash.len() + 4);
        buf.extend_from_slice(&pw_hash);
        buf.extend_from_slice(&block_index.to_le_bytes());
        let h = hash_alg.digest(&buf);

        // MS-OFFCRYPTO: for Standard/CryptoAPI RC4, a keyBits of 0 must be interpreted as 40.
        let key_bits = match derivation {
            StandardKeyDerivation::Rc4 if key_bits == 0 => 40,
            _ => key_bits,
        };

        let key_len = (key_bits as usize) / 8;
        match derivation {
            StandardKeyDerivation::Rc4 => h[..key_len].to_vec(),
            StandardKeyDerivation::AesTruncateHash => {
                if h.len() >= key_len {
                    h[..key_len].to_vec()
                } else {
                    let mut out = vec![0x36u8; key_len];
                    out[..h.len()].copy_from_slice(&h);
                    out
                }
            }
            StandardKeyDerivation::Aes => {
                // CryptoAPI key expansion used by MS-OFFCRYPTO Standard encryption: XOR the hash
                // into 0x36/0x5C padded 64-byte blocks (HMAC-like), hash each, and concatenate.
                let mut buf1 = vec![0x36u8; 64];
                let mut buf2 = vec![0x5cu8; 64];
                for (i, &b) in h.iter().take(64).enumerate() {
                    buf1[i] ^= b;
                    buf2[i] ^= b;
                }
                let mut out = hash_alg.digest(&buf1);
                if key_len > out.len() {
                    out.extend_from_slice(&hash_alg.digest(&buf2));
                }
                out.truncate(key_len);
                out
            }
        }
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
            for key_bits in [0u32, 40u32, 56u32, 128u32] {
                let effective_key_bits = if key_bits == 0 { 40 } else { key_bits };
                // Derive the expected key for block 0 and validate truncation.
                let key_ref = derive_key_ref(
                    hash_alg,
                    key_bits,
                    &salt,
                    password,
                    0,
                    StandardKeyDerivation::Rc4,
                );
                let deriver = StandardKeyDeriver::new(
                    hash_alg,
                    key_bits,
                    &salt,
                    password,
                    StandardKeyDerivation::Rc4,
                );
                let key0 = deriver.derive_key_for_block(0).expect("derive key");
                assert_eq!(
                    key0.as_slice(),
                    key_ref.as_slice(),
                    "StandardKeyDeriver should match the reference derivation"
                );
                assert_eq!(
                    key_ref.len(),
                    (effective_key_bits / 8) as usize,
                    "derived key length should be keySize/8"
                );
                assert!(
                    key_ref.len() <= hash_alg.digest_len(),
                    "for RC4, keyLen should be <= hashLen (truncation case)"
                );

                // Encrypt verifier + verifierHash with RC4 using a single continuous stream (the
                // CryptoAPI/Office behavior).
                let verifier_hash = hash_alg.digest(&verifier_plain);
                let mut verifier_buf =
                    Vec::with_capacity(verifier_plain.len() + verifier_hash.len());
                verifier_buf.extend_from_slice(&verifier_plain);
                verifier_buf.extend_from_slice(&verifier_hash);
                rc4_xor_in_place(&key_ref, &mut verifier_buf).expect("rc4 encrypt verifier buf");
                let encrypted_verifier = verifier_buf[..16].to_vec();
                let encrypted_verifier_hash = verifier_buf[16..].to_vec();

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
        let key_ref = derive_key_ref(
            hash_alg,
            key_bits,
            &salt,
            password,
            0,
            StandardKeyDerivation::Aes,
        );
        let deriver = StandardKeyDeriver::new(
            hash_alg,
            key_bits,
            &salt,
            password,
            StandardKeyDerivation::Aes,
        );
        let key0 = deriver.derive_key_for_block(0).expect("derive key");
        assert_eq!(key0.as_slice(), key_ref.as_slice());
        assert_eq!(key_ref.len(), 32);
        assert!(
            key_ref.len() > hash_alg.digest_len(),
            "AES-256+SHA1 should require CryptDeriveKey expansion"
        );

        // Encrypt verifier and verifierHash with AES-ECB (no IV).
        let encrypted_verifier =
            aes_ecb_encrypt(&key_ref, &verifier_plain).expect("encrypt verifier");

        let verifier_hash = hash_alg.digest(&verifier_plain);
        let mut verifier_hash_padded = verifier_hash.clone();
        verifier_hash_padded.resize(32, 0);
        let encrypted_verifier_hash =
            aes_ecb_encrypt(&key_ref, &verifier_hash_padded).expect("encrypt verifier hash");

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

        verify_password_standard(&header, &verifier, password)
            .expect("correct password should verify");
        let err = verify_password_standard(&header, &verifier, wrong_password)
            .expect_err("wrong password");
        assert!(
            matches!(err, OfficeCryptoError::InvalidPassword),
            "unexpected error: {err:?}"
        );
    }

    #[test]
    fn cryptderivekey_is_applied_for_aes128_sha1_even_though_keylen_le_digest_len() {
        let password = "correct horse battery staple";
        let salt: [u8; 16] = [0x24u8; 16];
        let hash_alg = HashAlgorithm::Sha1;
        let key_bits = 128u32;

        let key_ref = derive_key_ref(
            hash_alg,
            key_bits,
            &salt,
            password,
            0,
            StandardKeyDerivation::Aes,
        );
        assert_eq!(key_ref.len(), 16);

        // Ensure the AES derivation is *not* the RC4-style truncation.
        let pw = password_to_utf16le(password);
        let pw_hash = hash_password(hash_alg, &salt, &pw, 50_000);
        let mut buf = Vec::with_capacity(pw_hash.len() + 4);
        buf.extend_from_slice(&pw_hash);
        buf.extend_from_slice(&0u32.to_le_bytes());
        let h_final = hash_alg.digest(&buf);
        assert_ne!(key_ref.as_slice(), &h_final[..16]);

        let deriver = StandardKeyDeriver::new(
            hash_alg,
            key_bits,
            &salt,
            password,
            StandardKeyDerivation::Aes,
        );
        let key0 = deriver.derive_key_for_block(0).expect("derive key");
        assert_eq!(key0.as_slice(), key_ref.as_slice());
    }

    fn parsed_info() -> super::StandardEncryptionInfo {
        let info_bytes = standard_encryption_info_fixture();
        let header = parse_encryption_info_header(&info_bytes).expect("parse header");
        super::parse_standard_encryption_info(&info_bytes, &header).expect("parse standard")
    }

    #[test]
    fn decrypt_standard_rejects_u64_max_encrypted_package_size() {
        let info = parsed_info();

        // `u64::MAX` should be rejected as an absurd EncryptedPackage size before any allocation.
        let encrypted_package = u64::MAX.to_le_bytes().to_vec();

        let err = super::decrypt_standard_encrypted_package(&info, &encrypted_package, "Password")
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
    fn decrypt_standard_rejects_size_larger_than_ciphertext() {
        let info = parsed_info();

        let mut encrypted_package = Vec::new();
        encrypted_package.extend_from_slice(&100u64.to_le_bytes());
        encrypted_package.extend_from_slice(&[0u8; 16]);

        let err = super::decrypt_standard_encrypted_package(&info, &encrypted_package, "Password")
            .unwrap_err();
        assert!(matches!(err, OfficeCryptoError::InvalidFormat(_)));
    }

    #[test]
    fn decrypt_standard_rc4_rejects_size_larger_than_ciphertext() {
        // Regression test: ciphertext truncation should be rejected as InvalidFormat (not
        // InvalidPassword), and should happen before running the expensive password KDF.
        let info = StandardEncryptionInfo {
            version_major: 0,
            version_minor: 0,
            flags: 0,
            header: EncryptionHeader {
                alg_id: CALG_RC4,
                alg_id_hash: 0x0000_8004u32, // CALG_SHA1
                key_bits: 40,
                provider_type: 0,
                csp_name: String::new(),
            },
            verifier: EncryptionVerifier {
                salt: vec![0u8; 16],
                encrypted_verifier: vec![0u8; 16],
                verifier_hash_size: 20,
                encrypted_verifier_hash: vec![0u8; 20],
            },
        };

        let mut encrypted_package = Vec::new();
        encrypted_package.extend_from_slice(&100u64.to_le_bytes());
        encrypted_package.extend_from_slice(&[0u8; 16]);

        let err = super::decrypt_standard_encrypted_package(&info, &encrypted_package, "Password")
            .unwrap_err();
        assert!(matches!(err, OfficeCryptoError::InvalidFormat(_)));
    }
}
