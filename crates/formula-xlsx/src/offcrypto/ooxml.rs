use super::{
    decrypt_aes_cbc_no_padding, decrypt_agile_encrypted_package, derive_iv, derive_key,
    hash_password, HashAlgorithm, OffCryptoError, Result, AES_BLOCK_SIZE,
};

use std::io::Cursor;

use aes::cipher::{generic_array::GenericArray, BlockDecrypt, KeyInit};
use aes::{Aes128, Aes192, Aes256};
use sha1::Digest as _;

const ENCRYPTED_PACKAGE_SIZE_PREFIX_LEN: usize = 8;
const ENCRYPTED_PACKAGE_SEGMENT_LEN: usize = 0x1000;

// CryptoAPI algorithm identifiers (ALG_ID).
//
// Reference: WinCrypt `CALG_*` constants.
const CALG_RC4: u32 = 0x0000_6801;
const CALG_AES_128: u32 = 0x0000_660E;
const CALG_AES_192: u32 = 0x0000_660F;
const CALG_AES_256: u32 = 0x0000_6610;
const CALG_SHA1: u32 = 0x0000_8004;

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

    // MS-OFFCRYPTO identifies "Standard" encryption by `versionMinor == 2`, but real-world files
    // vary the major version across Office generations (2/3/4). Keep this aligned with our
    // detection logic (`formula-io` / `ooxml-encryption-info`) so we can decrypt 2.2/3.2/4.2
    // Standard-encrypted workbooks.
    match (major, minor) {
        (4, 4) => decrypt_agile_encrypted_package(
            encryption_info_stream,
            encrypted_package_stream,
            password,
        ),
        // MS-OFFCRYPTO / ECMA-376 identifies Standard (CryptoAPI) encryption via `versionMinor == 2`,
        // but `versionMajor` varies across Office generations (2/3/4 are observed in the wild).
        (major, 2) if (2..=4).contains(&major) => {
            decrypt_standard(encryption_info_stream, encrypted_package_stream, password)
        }
        _ => Err(OffCryptoError::UnsupportedEncryptionVersion { major, minor }),
    }
}

#[derive(Debug)]
struct StandardEncryptionInfo {
    alg_id: u32,
    alg_id_hash: u32,
    salt: Vec<u8>,
    key_size_bits: usize,
    key_len: usize,
    verifier_hash_size: usize,
    encrypted_verifier: Vec<u8>,
    encrypted_verifier_hash: Vec<u8>,
}

// Standard (CryptoAPI) password hashing uses a fixed spin count (MS-OFFCRYPTO / ECMA-376).
const STANDARD_SPIN_COUNT: u32 = 50_000;

// Legacy compatibility:
//
// We keep a fallback decryption path using a reduced spin count + CBC-based verifier decryption.
// This matches historical fixtures (and some third-party producers) that are not strictly
// spec-compliant.
const STANDARD_COMPAT_SPIN_COUNT: u32 = 1_000;

fn looks_like_ooxml_zip(bytes: &[u8]) -> bool {
    // XLSX/XLSM/etc are OPC ZIP archives. Avoid spending time on `zip` parsing when the signature
    // doesn't match.
    if !bytes.starts_with(b"PK") {
        return false;
    }

    // Validate via ZIP central directory parsing. This is a stronger check than `PK` prefix alone
    // and prevents false positives when trying multiple Standard `EncryptedPackage` schemes.
    //
    // We also require the OPC `[Content_Types].xml` part to exist, which is mandatory for valid
    // OOXML packages.
    let Ok(archive) = zip::ZipArchive::new(Cursor::new(bytes)) else {
        return false;
    };
    let ok = archive.file_names().any(|name| {
        let normalized = name.trim_start_matches(|c| c == '/' || c == '\\');
        normalized.eq_ignore_ascii_case("[Content_Types].xml")
    });
    ok
}

fn decrypt_standard(
    encryption_info_stream: &[u8],
    encrypted_package_stream: &[u8],
    password: &str,
) -> Result<Vec<u8>> {
    let info = parse_standard_encryption_info(encryption_info_stream)?;

    // Standard encryption is defined by CryptoAPI `AlgID`s inside the binary EncryptionHeader.
    //
    // `formula-xlsx` currently implements the AES subset used by Excel. Other algorithms (e.g. RC4)
    // are surfaced as "unsupported encryption".
    match info.alg_id {
        CALG_AES_128 | CALG_AES_192 | CALG_AES_256 => {
            // Try the spec-compliant CryptoAPI AES path first.
            match decrypt_standard_cryptoapi_aes(&info, encrypted_package_stream, password) {
                Ok(bytes) => Ok(bytes),
                Err(OffCryptoError::WrongPassword) => {
                    // Backwards-compatible fallback for non-standard producers / legacy fixtures.
                    decrypt_standard_compat_aes(&info, encrypted_package_stream, password)
                }
                Err(err) => Err(err),
            }
        }
        CALG_RC4 => Err(OffCryptoError::UnsupportedCipherAlgorithm {
            cipher: "RC4 (CryptoAPI Standard encryption)".to_string(),
        }),
        other => Err(OffCryptoError::UnsupportedCipherAlgorithm {
            cipher: format!("CryptoAPI algId=0x{other:08x}"),
        }),
    }
}

fn decrypt_standard_cryptoapi_aes(
    info: &StandardEncryptionInfo,
    encrypted_package_stream: &[u8],
    password: &str,
) -> Result<Vec<u8>> {
    if !matches!(info.key_len, 16 | 24 | 32) {
        return Err(OffCryptoError::UnsupportedCipherAlgorithm {
            cipher: format!("AES (keySize {} bits)", info.key_size_bits),
        });
    }
    if info.alg_id_hash != CALG_SHA1 {
        return Err(OffCryptoError::UnsupportedHashAlgorithm {
            hash: format!("CryptoAPI algIdHash=0x{:08x}", info.alg_id_hash),
        });
    }

    // 1) Compute the iterated password hash H (fixed spin count = 50,000 for Standard).
    let h = hash_password(password, &info.salt, STANDARD_SPIN_COUNT, HashAlgorithm::Sha1).map_err(
        |err| OffCryptoError::StandardEncryptionInfoMalformed {
            reason: err.to_string(),
        },
    )?;

    // 2) Derive a per-block hash for block=0: H_block = SHA1(H || LE32(0)).
    let mut hasher = sha1::Sha1::new();
    hasher.update(&h);
    hasher.update(&0u32.to_le_bytes());
    let h_block0 = hasher.finalize();

    // 3) Derive the AES file key via CryptoAPI `CryptDeriveKey`.
    let key0 = cryptoapi_crypt_derive_key_sha1(&h_block0, info.key_len)?;

    // 4) Verify the password by decrypting the verifier fields with AES-ECB.
    verify_standard_verifier_aes_ecb(
        &key0,
        &info.encrypted_verifier,
        &info.encrypted_verifier_hash,
        info.verifier_hash_size,
    )?;

    // 5) Decrypt the `EncryptedPackage` stream.
    //
    // Standard/CryptoAPI encryption has multiple variants in the wild. In practice, the main
    // interoperability differences are in how the package ciphertext is chunked and how the AES-CBC
    // IV (and sometimes key) are varied per chunk.
    //
    // Try a small set of schemes and pick the first result that looks like an OOXML ZIP payload.
    //
    // NOTE: We validate by parsing the decrypted ZIP, not just checking `PK` prefix. `PK` can occur
    // by chance with wrong IVs/keys, and we want decryption to reliably surface WrongPassword
    // instead of returning garbage bytes that fail later.
    //
    // 1) The spec-like CBC-per-segment-IV scheme.
    if let Ok(out) = decrypt_encrypted_package_stream(
        encrypted_package_stream,
        &key0,
        &info.salt,
        HashAlgorithm::Sha1,
        AES_BLOCK_SIZE,
    ) {
        if looks_like_ooxml_zip(&out) {
            return Ok(out);
        }
    }

    // 2) Some producers encrypt the package with AES-ECB (no IV).
    if let Ok(out) = decrypt_encrypted_package_stream_aes_ecb(encrypted_package_stream, &key0) {
        if looks_like_ooxml_zip(&out) {
            return Ok(out);
        }
    }

    // 3) Segment into 4096-byte chunks; for chunk N use key(block=N) and IV=0.
    if let Ok(out) =
        decrypt_standard_encrypted_package_per_block_key_iv_zero(encrypted_package_stream, &h, info.key_len)
    {
        if looks_like_ooxml_zip(&out) {
            return Ok(out);
        }
    }

    // 4) Constant-key CBC with IV=0 (per segment).
    if let Ok(out) = decrypt_encrypted_package_stream_constant_iv_zero(encrypted_package_stream, &key0) {
        if looks_like_ooxml_zip(&out) {
            return Ok(out);
        }
    }

    Err(OffCryptoError::WrongPassword)
}

fn decrypt_standard_compat_aes(
    info: &StandardEncryptionInfo,
    encrypted_package_stream: &[u8],
    password: &str,
) -> Result<Vec<u8>> {
    // Derive the file key (AES key) from the password using the legacy/compat KDF.
    let pw_hash = hash_password(
        password,
        &info.salt,
        STANDARD_COMPAT_SPIN_COUNT,
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

    // Package decryption scheme varies; prefer the historical (hash-based IV) scheme, but fall back
    // to IV=0 if needed.
    if let Ok(out) = decrypt_encrypted_package_stream(
        encrypted_package_stream,
        &key,
        &info.salt,
        HashAlgorithm::Sha1,
        AES_BLOCK_SIZE,
    ) {
        if looks_like_ooxml_zip(&out) {
            return Ok(out);
        }
    }

    // Some producers encrypt the package with AES-ECB (no IV).
    if let Ok(out) = decrypt_encrypted_package_stream_aes_ecb(encrypted_package_stream, &key) {
        if looks_like_ooxml_zip(&out) {
            return Ok(out);
        }
    }

    if let Ok(out) = decrypt_encrypted_package_stream_constant_iv_zero(encrypted_package_stream, &key) {
        if looks_like_ooxml_zip(&out) {
            return Ok(out);
        }
    }

    Err(OffCryptoError::WrongPassword)
}

fn decrypt_standard_encrypted_package_per_block_key_iv_zero(
    encrypted_package_stream: &[u8],
    iterated_hash: &[u8],
    key_len: usize,
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
    let iv = [0u8; AES_BLOCK_SIZE];
    while offset < ciphertext.len() && out.len() < orig_size_usize {
        let remaining = ciphertext.len() - offset;
        let seg_len = remaining.min(ENCRYPTED_PACKAGE_SEGMENT_LEN);

        if seg_len % AES_BLOCK_SIZE != 0 {
            return Err(OffCryptoError::CiphertextNotBlockAligned {
                field: "EncryptedPackage segment",
                len: seg_len,
            });
        }

        // key(block=N) = CryptDeriveKey(SHA1(Hash(H || LE32(N))), keyLen)
        let mut hasher = sha1::Sha1::new();
        hasher.update(iterated_hash);
        hasher.update(&segment_index.to_le_bytes());
        let h_block = hasher.finalize();
        let key = cryptoapi_crypt_derive_key_sha1(&h_block, key_len)?;

        let decrypted =
            decrypt_aes_cbc_no_padding(&key, &iv, &ciphertext[offset..offset + seg_len])?;

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

fn decrypt_encrypted_package_stream_aes_ecb(
    encrypted_package_stream: &[u8],
    key: &[u8],
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

    let mut out = decrypt_aes_ecb_no_padding(key, ciphertext)?;
    if out.len() < orig_size_usize {
        return Err(OffCryptoError::DecryptedLengthShorterThanHeader {
            declared_len: orig_size_usize,
            available_len: out.len(),
        });
    }
    out.truncate(orig_size_usize);
    Ok(out)
}

fn decrypt_encrypted_package_stream_constant_iv_zero(
    encrypted_package_stream: &[u8],
    key: &[u8],
) -> Result<Vec<u8>> {
    let iv = [0u8; AES_BLOCK_SIZE];

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
    while offset < ciphertext.len() && out.len() < orig_size_usize {
        let remaining = ciphertext.len() - offset;
        let seg_len = remaining.min(ENCRYPTED_PACKAGE_SEGMENT_LEN);

        if seg_len % AES_BLOCK_SIZE != 0 {
            return Err(OffCryptoError::CiphertextNotBlockAligned {
                field: "EncryptedPackage segment",
                len: seg_len,
            });
        }

        let decrypted = decrypt_aes_cbc_no_padding(key, &iv, &ciphertext[offset..offset + seg_len])?;

        let remaining_needed = orig_size_usize - out.len();
        if decrypted.len() > remaining_needed {
            out.extend_from_slice(&decrypted[..remaining_needed]);
            break;
        }
        out.extend_from_slice(&decrypted);

        offset += seg_len;
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

fn cryptoapi_crypt_derive_key_sha1(h_block: &[u8], key_len: usize) -> Result<Vec<u8>> {
    // CryptoAPI's `CryptDeriveKey` for SHA1 uses an ipad/opad expansion (similar to HMAC) and then
    // truncates to the requested key length:
    //
    // derived = SHA1((D XOR 0x36*64)) || SHA1((D XOR 0x5c*64))
    // where D is the digest padded to 64 bytes with zeros.
    if key_len > 40 {
        return Err(OffCryptoError::StandardEncryptionInfoMalformed {
            reason: "CryptoAPI-derived key length exceeds SHA1 expansion length".to_string(),
        });
    }
    if h_block.len() != 20 {
        return Err(OffCryptoError::StandardEncryptionInfoMalformed {
            reason: "CryptoAPI H_block must be a SHA1 digest (20 bytes)".to_string(),
        });
    }

    let mut d = [0u8; 64];
    d[..20].copy_from_slice(h_block);

    let mut inner = [0x36u8; 64];
    let mut outer = [0x5cu8; 64];
    for i in 0..64 {
        inner[i] ^= d[i];
        outer[i] ^= d[i];
    }

    let x1 = sha1::Sha1::digest(inner);
    let x2 = sha1::Sha1::digest(outer);
    let mut derived = [0u8; 40];
    derived[..20].copy_from_slice(&x1);
    derived[20..].copy_from_slice(&x2);

    Ok(derived[..key_len].to_vec())
}

fn decrypt_aes_ecb_no_padding(key: &[u8], ciphertext: &[u8]) -> Result<Vec<u8>> {
    if ciphertext.is_empty() {
        return Ok(Vec::new());
    }
    if ciphertext.len() % AES_BLOCK_SIZE != 0 {
        return Err(OffCryptoError::CiphertextNotBlockAligned {
            field: "AES-ECB ciphertext",
            len: ciphertext.len(),
        });
    }

    let mut out = ciphertext.to_vec();
    match key.len() {
        16 => {
            let dec = Aes128::new_from_slice(key).map_err(|_| {
                OffCryptoError::UnsupportedCipherAlgorithm {
                    cipher: format!("AES (key length {} bytes)", key.len()),
                }
            })?;
            for chunk in out.chunks_exact_mut(AES_BLOCK_SIZE) {
                dec.decrypt_block(GenericArray::from_mut_slice(chunk));
            }
        }
        24 => {
            let dec = Aes192::new_from_slice(key).map_err(|_| {
                OffCryptoError::UnsupportedCipherAlgorithm {
                    cipher: format!("AES (key length {} bytes)", key.len()),
                }
            })?;
            for chunk in out.chunks_exact_mut(AES_BLOCK_SIZE) {
                dec.decrypt_block(GenericArray::from_mut_slice(chunk));
            }
        }
        32 => {
            let dec = Aes256::new_from_slice(key).map_err(|_| {
                OffCryptoError::UnsupportedCipherAlgorithm {
                    cipher: format!("AES (key length {} bytes)", key.len()),
                }
            })?;
            for chunk in out.chunks_exact_mut(AES_BLOCK_SIZE) {
                dec.decrypt_block(GenericArray::from_mut_slice(chunk));
            }
        }
        other => {
            return Err(OffCryptoError::UnsupportedCipherAlgorithm {
                cipher: format!("AES (key length {other} bytes)"),
            })
        }
    }
    Ok(out)
}

fn verify_standard_verifier_aes_ecb(
    key: &[u8],
    encrypted_verifier: &[u8],
    encrypted_verifier_hash: &[u8],
    verifier_hash_size: usize,
) -> Result<()> {
    let mut ciphertext = Vec::with_capacity(encrypted_verifier.len() + encrypted_verifier_hash.len());
    ciphertext.extend_from_slice(encrypted_verifier);
    ciphertext.extend_from_slice(encrypted_verifier_hash);

    let plaintext = decrypt_aes_ecb_no_padding(key, &ciphertext)?;

    let verifier = plaintext
        .get(..16)
        .ok_or_else(|| OffCryptoError::StandardEncryptionInfoMalformed {
            reason: "decrypted verifier is truncated".to_string(),
        })?;
    let verifier_hash = plaintext
        .get(16..16 + verifier_hash_size)
        .ok_or_else(|| OffCryptoError::StandardEncryptionInfoMalformed {
            reason: "decrypted verifierHash is truncated".to_string(),
        })?;

    let expected = HashAlgorithm::Sha1.hash(verifier);
    let expected = expected.get(..verifier_hash_size).ok_or_else(|| {
        OffCryptoError::StandardEncryptionInfoMalformed {
            reason: "verifierHashSize larger than SHA1 digest length".to_string(),
        }
    })?;
    if expected != verifier_hash {
        return Err(OffCryptoError::WrongPassword);
    }
    Ok(())
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
    let header_end = offset.checked_add(header_size).ok_or_else(|| {
        OffCryptoError::StandardEncryptionInfoMalformed {
            reason: "EncryptionHeader.size causes overflow".to_string(),
        }
    })?;
    if bytes.len() < header_end {
        return Err(OffCryptoError::StandardEncryptionInfoMalformed {
            reason: "truncated EncryptionHeader".to_string(),
        });
    }
    let header_bytes = &bytes[offset..header_end];
    offset = header_end;

    if header_bytes.len() < 8 * 4 {
        return Err(OffCryptoError::StandardEncryptionInfoMalformed {
            reason: "EncryptionHeader missing fixed fields".to_string(),
        });
    }

    let alg_id =
        u32::from_le_bytes(header_bytes[8..12].try_into().expect("slice is 4 bytes"));
    let alg_id_hash =
        u32::from_le_bytes(header_bytes[12..16].try_into().expect("slice is 4 bytes"));

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
    let salt_end = offset.checked_add(salt_size).ok_or_else(|| {
        OffCryptoError::StandardEncryptionInfoMalformed {
            reason: "EncryptionVerifier.saltSize causes overflow".to_string(),
        }
    })?;
    if bytes.len() < salt_end {
        return Err(OffCryptoError::StandardEncryptionInfoMalformed {
            reason: "truncated EncryptionVerifier.salt".to_string(),
        });
    }
    let salt = bytes[offset..salt_end].to_vec();
    offset = salt_end;

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
        alg_id,
        alg_id_hash,
        salt,
        key_size_bits,
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

    // --- Guardrails for malicious `orig_size` ---
    //
    // `EncryptedPackage` stores the unencrypted package size (`orig_size`) separately from the
    // ciphertext bytes. A corrupt/malicious size can otherwise induce large allocations (OOM) or
    // panics in `Vec::with_capacity` on 64-bit targets.
    //
    // Keep checks conservative to avoid rejecting valid-but-unusual files.
    let plausible_max =
        (ciphertext.len() as u64).saturating_add(ENCRYPTED_PACKAGE_SEGMENT_LEN as u64);
    if orig_size > plausible_max {
        return Err(OffCryptoError::InvalidAttribute {
            element: "EncryptedPackage".to_string(),
            attr: "origSize".to_string(),
            reason: format!(
                "origSize {orig_size} is implausibly large for ciphertext length {}",
                ciphertext.len()
            ),
        });
    }

    // Guardrail: ciphertext must be long enough to possibly contain `orig_size` bytes (accounting
    // for AES block padding).
    let expected_min_ciphertext_len = orig_size
        .checked_add((AES_BLOCK_SIZE - 1) as u64)
        .and_then(|v| v.checked_div(AES_BLOCK_SIZE as u64))
        .and_then(|blocks| blocks.checked_mul(AES_BLOCK_SIZE as u64))
        .ok_or_else(|| OffCryptoError::InvalidAttribute {
            element: "EncryptedPackage".to_string(),
            attr: "origSize".to_string(),
            reason: format!(
                "origSize {orig_size} is implausibly large for ciphertext length {}",
                ciphertext.len()
            ),
        })?;
    if (ciphertext.len() as u64) < expected_min_ciphertext_len {
        return Err(OffCryptoError::InvalidAttribute {
            element: "EncryptedPackage".to_string(),
            attr: "origSize".to_string(),
            reason: format!(
                "ciphertext length {} is too short for declared origSize {orig_size}",
                ciphertext.len()
            ),
        });
    }

    let orig_size_usize =
        usize::try_from(orig_size).map_err(|_| OffCryptoError::InvalidAttribute {
            element: "EncryptedPackage".to_string(),
            attr: "origSize".to_string(),
            reason: "origSize does not fit into usize".to_string(),
        })?;

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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn rejects_implausible_orig_size_without_panic() {
        let key = [0u8; 16];
        let salt = [0u8; 16];

        let mut stream = Vec::new();
        stream.extend_from_slice(&u64::MAX.to_le_bytes());
        stream.extend_from_slice(&[0u8; AES_BLOCK_SIZE]); // 1 AES block of ciphertext

        let err = decrypt_encrypted_package_stream(
            &stream,
            &key,
            &salt,
            HashAlgorithm::Sha1,
            AES_BLOCK_SIZE,
        )
        .expect_err("expected error");

        assert!(
            matches!(
                err,
                OffCryptoError::InvalidAttribute { ref element, ref attr, .. }
                    if element == "EncryptedPackage" && attr == "origSize"
            ),
            "unexpected error: {err:?}"
        );
    }

    #[test]
    fn rejects_orig_size_near_u64_max_without_overflow() {
        let key = [0u8; 16];
        let salt = [0u8; 16];

        let mut stream = Vec::new();
        stream.extend_from_slice(&(u64::MAX - 4094).to_le_bytes());
        stream.extend_from_slice(&[0u8; AES_BLOCK_SIZE]);

        let err = decrypt_encrypted_package_stream(
            &stream,
            &key,
            &salt,
            HashAlgorithm::Sha1,
            AES_BLOCK_SIZE,
        )
        .expect_err("expected error");

        assert!(
            matches!(
                err,
                OffCryptoError::InvalidAttribute { ref element, ref attr, .. }
                    if element == "EncryptedPackage" && attr == "origSize"
            ),
            "unexpected error: {err:?}"
        );
    }
}
