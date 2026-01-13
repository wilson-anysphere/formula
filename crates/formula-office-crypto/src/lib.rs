//! Decryption of the standard Office OOXML encryption wrapper (`EncryptionInfo` + `EncryptedPackage`)
//! stored inside an OLE/CFB container.
//!
//! This crate supports:
//! - MS-OFFCRYPTO "Agile Encryption" (XML descriptor, Office 2010+)
//! - MS-OFFCRYPTO / ECMA-376 "Standard Encryption" (binary descriptor, Office 2007-era)
//!
//! The decrypted output is the raw OOXML ZIP/OPC bytes (should start with `PK`).

mod agile;
mod crypto;
mod error;
mod ole;
mod standard;
mod util;

use std::io::{Cursor, Read};

pub use crate::crypto::HashAlgorithm;
pub use crate::error::OfficeCryptoError;
pub use crate::ole::{extract_ole_entries, OleEntries, OleEntry, OleStream};

const OLE_MAGIC: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

/// Default maximum `spinCount` accepted during Agile password-based decryption.
///
/// `spinCount` is attacker-controlled in Agile-encrypted files; bounding it avoids CPU DoS from
/// maliciously large values (Excel commonly uses 100,000).
pub const DEFAULT_MAX_SPIN_COUNT: u32 = 1_000_000;

/// Maximum allowed size for the Agile (XML) `EncryptionInfo` payload.
///
/// The XML descriptor is typically a few KB. This cap prevents parsing or base64-decoding
/// arbitrarily large XML documents from untrusted files.
pub const MAX_AGILE_ENCRYPTION_INFO_XML_BYTES: usize = 1024 * 1024; // 1MiB

/// Maximum allowed size for the Standard (binary) `EncryptionHeader` section of `EncryptionInfo`.
pub const MAX_STANDARD_ENCRYPTION_HEADER_BYTES: usize = 16 * 1024;

/// Maximum allowed size for the Standard `EncryptionHeader.cspName` field.
pub const MAX_STANDARD_CSPNAME_BYTES: usize = 8 * 1024;

/// Maximum allowed verifier hash size in bytes for Standard encryption.
///
/// Supported hashes (SHA-1..SHA-512) are <= 64 bytes.
pub const MAX_STANDARD_VERIFIER_HASH_SIZE_BYTES: usize = 64;

/// Maximum allowed declared decrypted size for the `EncryptedPackage` stream.
///
/// This size prefix is untrusted; callers should reject absurd values before allocating output.
pub const MAX_ENCRYPTED_PACKAGE_ORIGINAL_SIZE: u64 = 512 * 1024 * 1024; // 512MiB

#[derive(Debug, Copy, Clone, PartialEq, Eq)]
pub enum EncryptionScheme {
    Agile,
    Standard,
}

#[derive(Debug, Clone)]
pub struct EncryptOptions {
    pub scheme: EncryptionScheme,
    pub key_bits: usize,
    pub hash_algorithm: HashAlgorithm,
    pub spin_count: u32,
}

impl Default for EncryptOptions {
    fn default() -> Self {
        Self {
            scheme: EncryptionScheme::Agile,
            key_bits: 256,
            hash_algorithm: HashAlgorithm::Sha512,
            spin_count: 100_000,
        }
    }
}

/// Options controlling decryption resource limits.
#[derive(Debug, Clone)]
pub struct DecryptOptions {
    /// Maximum allowed Agile `spinCount` before rejecting the file as unsafe/too expensive to
    /// process.
    pub max_spin_count: u32,
}

impl Default for DecryptOptions {
    fn default() -> Self {
        Self {
            max_spin_count: DEFAULT_MAX_SPIN_COUNT,
        }
    }
}

/// Returns true if the provided bytes look like an OLE/CFB container holding an Office-encrypted
/// OOXML package (streams `EncryptionInfo` and `EncryptedPackage`).
pub fn is_encrypted_ooxml_ole(bytes: &[u8]) -> bool {
    if bytes.len() < OLE_MAGIC.len() || bytes[..OLE_MAGIC.len()] != OLE_MAGIC {
        return false;
    }

    let cursor = Cursor::new(bytes);
    let Ok(mut ole) = cfb::CompoundFile::open(cursor) else {
        return false;
    };

    stream_exists_case_tolerant(&mut ole, "EncryptionInfo")
        && stream_exists_case_tolerant(&mut ole, "EncryptedPackage")
}

/// Decrypt an Office-encrypted OOXML OLE/CFB wrapper and return the decrypted raw ZIP bytes.
pub fn decrypt_encrypted_package_ole(
    bytes: &[u8],
    password: &str,
) -> Result<Vec<u8>, OfficeCryptoError> {
    decrypt_encrypted_package_ole_with_options(bytes, password, &DecryptOptions::default())
}

/// Like [`decrypt_encrypted_package_ole`], but allows overriding resource limits.
pub fn decrypt_encrypted_package_ole_with_options(
    bytes: &[u8],
    password: &str,
    opts: &DecryptOptions,
) -> Result<Vec<u8>, OfficeCryptoError> {
    let cursor = Cursor::new(bytes);
    let mut ole = cfb::CompoundFile::open(cursor)?;

    let mut encryption_info = Vec::new();
    open_stream_case_tolerant(&mut ole, "EncryptionInfo")?.read_to_end(&mut encryption_info)?;

    let mut encrypted_package = Vec::new();
    open_stream_case_tolerant(&mut ole, "EncryptedPackage")?.read_to_end(&mut encrypted_package)?;

    decrypt_encrypted_package_streams_with_options(
        &encryption_info,
        &encrypted_package,
        password,
        opts,
    )
}

/// Decrypt an Office-encrypted OOXML OLE/CFB wrapper and return the decrypted raw ZIP bytes.
///
/// This is a convenience wrapper around [`decrypt_encrypted_package_ole`] that matches the
/// call shape used by `formula-io` and other consumers.
pub fn decrypt_encrypted_package(
    ole_bytes: &[u8],
    password: &str,
) -> Result<Vec<u8>, OfficeCryptoError> {
    decrypt_encrypted_package_with_options(ole_bytes, password, &DecryptOptions::default())
}

/// Like [`decrypt_encrypted_package`], but allows overriding resource limits.
pub fn decrypt_encrypted_package_with_options(
    ole_bytes: &[u8],
    password: &str,
    opts: &DecryptOptions,
) -> Result<Vec<u8>, OfficeCryptoError> {
    decrypt_encrypted_package_ole_with_options(ole_bytes, password, opts)
}

/// Decrypt a Standard (CryptoAPI) encrypted OOXML package given the raw `EncryptionInfo` and
/// `EncryptedPackage` stream bytes.
///
/// This is a lower-level helper that avoids requiring the full OLE/CFB container. It is primarily
/// intended for callers (like `formula-io`) that already have access to the streams.
pub fn decrypt_standard_encrypted_package(
    encryption_info: &[u8],
    encrypted_package: &[u8],
    password: &str,
) -> Result<Vec<u8>, OfficeCryptoError> {
    let header = util::parse_encryption_info_header(encryption_info)?;
    if header.kind != util::EncryptionInfoKind::Standard {
        return Err(OfficeCryptoError::UnsupportedEncryption(format!(
            "expected Standard EncryptionInfo, got version {}.{}",
            header.version_major, header.version_minor
        )));
    }

    let info = standard::parse_standard_encryption_info(encryption_info, &header)?;
    let out = standard::decrypt_standard_encrypted_package(&info, encrypted_package, password)?;
    validate_decrypted_package(&out)?;
    Ok(out)
}

/// Encrypt a raw OOXML ZIP package into an Office `EncryptedPackage` OLE/CFB wrapper.
///
/// The returned bytes are an OLE/CFB container containing:
/// - `EncryptionInfo` stream (Agile XML descriptor, by default)
/// - `EncryptedPackage` stream (8-byte decrypted size prefix + encrypted payload)
pub fn encrypt_package_to_ole(
    zip_bytes: &[u8],
    password: &str,
    opts: EncryptOptions,
) -> Result<Vec<u8>, OfficeCryptoError> {
    encrypt_package_to_ole_with_entries(zip_bytes, password, opts, None)
}

/// Encrypt a raw OOXML ZIP package into an Office `EncryptedPackage` OLE/CFB wrapper, optionally
/// preserving extra OLE streams/storages from a source container.
///
/// `preserve` should typically contain the non-encryption streams/storages extracted from the
/// original encrypted file (e.g. `\u{0005}SummaryInformation`). The `EncryptionInfo` and
/// `EncryptedPackage` streams are **always** replaced with freshly generated values.
pub fn encrypt_package_to_ole_with_entries(
    zip_bytes: &[u8],
    password: &str,
    opts: EncryptOptions,
    preserve: Option<&OleEntries>,
) -> Result<Vec<u8>, OfficeCryptoError> {
    use std::io::Write as _;

    let (encryption_info, encrypted_package) = match opts.scheme {
        EncryptionScheme::Agile => {
            agile::encrypt_agile_encrypted_package(zip_bytes, password, &opts)?
        }
        EncryptionScheme::Standard => {
            standard::encrypt_standard_encrypted_package(zip_bytes, password, &opts)?
        }
    };

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor)?;

    if let Some(entries) = preserve {
        ole::copy_entries_into_ole(&mut ole, entries)?;
    }

    ole.create_stream("EncryptionInfo")?
        .write_all(&encryption_info)?;
    ole.create_stream("EncryptedPackage")?
        .write_all(&encrypted_package)?;

    Ok(ole.into_inner().into_inner())
}

fn decrypt_encrypted_package_streams_with_options(
    encryption_info: &[u8],
    encrypted_package: &[u8],
    password: &str,
    opts: &DecryptOptions,
) -> Result<Vec<u8>, OfficeCryptoError> {
    let header = util::parse_encryption_info_header(encryption_info)?;
    match header.kind {
        util::EncryptionInfoKind::Agile => {
            let info = agile::parse_agile_encryption_info(encryption_info, &header)?;
            let out =
                agile::decrypt_agile_encrypted_package(&info, encrypted_package, password, opts)?;
            validate_decrypted_package(&out)?;
            Ok(out)
        }
        util::EncryptionInfoKind::Standard => {
            let info = standard::parse_standard_encryption_info(encryption_info, &header)?;
            let out =
                standard::decrypt_standard_encrypted_package(&info, encrypted_package, password)?;
            validate_decrypted_package(&out)?;
            Ok(out)
        }
    }
}

fn open_stream_case_tolerant<R: Read + std::io::Seek>(
    ole: &mut cfb::CompoundFile<R>,
    name: &str,
) -> Result<cfb::Stream<R>, OfficeCryptoError> {
    // Some OLE writers expose root streams with a leading `/` (e.g. `/EncryptionInfo`) even though
    // most Office-produced containers use `EncryptionInfo`. Additionally, some producers vary
    // casing for these stream names. Be tolerant and try:
    // - `name` and `/{name}`
    // - case-insensitive match via `walk()` for implementations that treat `open_stream` as
    //   case-sensitive.
    let want = name.trim_start_matches('/');
    let mut all_not_found = true;
    let mut first_err: Option<std::io::Error> = None;

    fn record_err(all_not_found: &mut bool, first_err: &mut Option<std::io::Error>, err: std::io::Error) {
        if err.kind() != std::io::ErrorKind::NotFound {
            *all_not_found = false;
        }
        if first_err.is_none() {
            *first_err = Some(err);
        }
    }

    match ole.open_stream(want) {
        Ok(s) => return Ok(s),
        Err(err) => record_err(&mut all_not_found, &mut first_err, err),
    }

    let with_leading_slash = format!("/{want}");
    match ole.open_stream(&with_leading_slash) {
        Ok(s) => return Ok(s),
        Err(err) => record_err(&mut all_not_found, &mut first_err, err),
    }

    // Case-insensitive fallback: walk the directory tree and match stream paths.
    let mut found_path: Option<String> = None;
    for entry in ole.walk() {
        if !entry.is_stream() {
            continue;
        }
        let path = entry.path().to_string_lossy();
        let normalized = path.as_ref().strip_prefix('/').unwrap_or(path.as_ref());
        if normalized.eq_ignore_ascii_case(want) {
            found_path = Some(path.into_owned());
            break;
        }
    }

    if let Some(found_path) = found_path {
        match ole.open_stream(&found_path) {
            Ok(s) => return Ok(s),
            Err(err) => record_err(&mut all_not_found, &mut first_err, err),
        }

        // Some implementations accept the walk()-returned path but reject a leading slash
        // (or vice versa); try again stripped.
        let stripped = found_path.strip_prefix('/').unwrap_or(found_path.as_str());
        if stripped != found_path {
            match ole.open_stream(stripped) {
                Ok(s) => return Ok(s),
                Err(err) => record_err(&mut all_not_found, &mut first_err, err),
            }
            let with_slash = format!("/{stripped}");
            match ole.open_stream(&with_slash) {
                Ok(s) => return Ok(s),
                Err(err) => record_err(&mut all_not_found, &mut first_err, err),
            }
        }
    }

    if all_not_found {
        return Err(OfficeCryptoError::InvalidFormat(format!(
            "missing OLE stream {want}"
        )));
    }

    let suffix = first_err
        .map(|e| format!(": {e}"))
        .unwrap_or_default();
    Err(OfficeCryptoError::InvalidFormat(format!(
        "failed to open OLE stream {want}{suffix}"
    )))
}

fn stream_exists_case_tolerant<R: Read + std::io::Seek>(
    ole: &mut cfb::CompoundFile<R>,
    name: &str,
) -> bool {
    open_stream_case_tolerant(ole, name).is_ok()
}

fn validate_decrypted_package(bytes: &[u8]) -> Result<(), OfficeCryptoError> {
    if !util::looks_like_zip(bytes) {
        return Err(OfficeCryptoError::InvalidFormat(
            "decrypted package does not look like a valid ZIP archive".to_string(),
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::crypto::{HashAlgorithm, StandardKeyDerivation, StandardKeyDeriver};
    use crate::test_alloc::MAX_ALLOC;
    use std::path::{Path, PathBuf};
    use std::sync::atomic::Ordering;

    #[test]
    fn detects_encrypted_ooxml_ole_container() {
        let cursor = Cursor::new(Vec::new());
        let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
        ole.create_stream("EncryptionInfo")
            .expect("create EncryptionInfo stream");
        ole.create_stream("EncryptedPackage")
            .expect("create EncryptedPackage stream");
        let bytes = ole.into_inner().into_inner();
        assert!(is_encrypted_ooxml_ole(&bytes));
    }

    #[test]
    fn detects_encrypted_ooxml_ole_container_with_leading_slash_stream_names() {
        let cursor = Cursor::new(Vec::new());
        let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
        ole.create_stream("/EncryptionInfo")
            .expect("create /EncryptionInfo stream");
        ole.create_stream("/EncryptedPackage")
            .expect("create /EncryptedPackage stream");
        let bytes = ole.into_inner().into_inner();

        // Ensure the encrypted OOXML detection is tolerant to the leading `/` naming quirk.
        assert!(is_encrypted_ooxml_ole(&bytes));

        // Ensure we reach a format-related error (corrupt/missing content) instead of failing to
        // open the streams with NotFound.
        let err = decrypt_encrypted_package_ole(&bytes, "password").expect_err("expected error");
        assert!(
            matches!(err, OfficeCryptoError::InvalidFormat(_)),
            "unexpected error: {err:?}"
        );
    }

    #[test]
    fn detects_encrypted_ooxml_ole_container_with_case_variant_stream_names() {
        let cursor = Cursor::new(Vec::new());
        let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
        ole.create_stream("/encryptioninfo")
            .expect("create /encryptioninfo stream");
        ole.create_stream("/ENCRYPTEDPACKAGE")
            .expect("create /ENCRYPTEDPACKAGE stream");
        let bytes = ole.into_inner().into_inner();

        // Ensure detection is tolerant to casing (some producers vary stream case).
        assert!(is_encrypted_ooxml_ole(&bytes));

        // Ensure we reach a format-related error instead of failing to open streams.
        let err = decrypt_encrypted_package_ole(&bytes, "password").expect_err("expected error");
        assert!(
            matches!(err, OfficeCryptoError::InvalidFormat(_)),
            "unexpected error: {err:?}"
        );
    }

    #[test]
    fn parses_standard_encryption_info_minimal() {
        let info_bytes = standard::tests::standard_encryption_info_fixture();
        let header = util::parse_encryption_info_header(&info_bytes).expect("parse header");
        assert_eq!(header.kind, util::EncryptionInfoKind::Standard);
        let parsed =
            standard::parse_standard_encryption_info(&info_bytes, &header).expect("parse standard");
        assert_eq!(parsed.version_major, 4);
        assert_eq!(parsed.version_minor, 2);
        assert_eq!(parsed.header.key_bits, 128);
        assert_eq!(parsed.verifier.salt.len(), 16);
    }

    #[test]
    fn parses_standard_encryption_info_accepts_version_2_2() {
        // Standard encryption is identified by `versionMinor == 2` and `versionMajor ∈ {2,3,4}` in
        // the wild; ensure we don't over-gate on the common `3.2`/`4.2` pairs.
        let mut info_bytes = standard::tests::standard_encryption_info_fixture();
        info_bytes[..2].copy_from_slice(&2u16.to_le_bytes()); // versionMajor
        info_bytes[2..4].copy_from_slice(&2u16.to_le_bytes()); // versionMinor

        let header = util::parse_encryption_info_header(&info_bytes).expect("parse header");
        assert_eq!(header.kind, util::EncryptionInfoKind::Standard);

        let parsed =
            standard::parse_standard_encryption_info(&info_bytes, &header).expect("parse standard");
        assert_eq!(parsed.version_major, 2);
        assert_eq!(parsed.version_minor, 2);
    }

    #[test]
    fn parses_agile_encryption_info_minimal() {
        let info_bytes = agile::tests::agile_encryption_info_fixture();
        let header = util::parse_encryption_info_header(&info_bytes).expect("parse header");
        assert_eq!(header.kind, util::EncryptionInfoKind::Agile);
        let parsed = agile::parse_agile_encryption_info(&info_bytes, &header).expect("parse agile");
        assert_eq!(parsed.version_major, 4);
        assert_eq!(parsed.version_minor, 4);
        assert_eq!(parsed.key_data.key_bits, 256);
        assert_eq!(parsed.password_key_encryptor.spin_count, 100_000);
    }

    #[test]
    fn rejects_agile_spin_count_above_default_max() {
        // Construct a minimal Agile EncryptionInfo stream with an oversized spinCount. The
        // EncryptedPackage body can be empty; we only want to assert we reject before running the
        // expensive password KDF.
        let xml = r#"
            <encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
                        xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
              <keyData saltValue="AAAAAAAAAAAAAAAAAAAAAA==" hashAlgorithm="SHA1"
                       cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                       keyBits="128" blockSize="16" />
              <dataIntegrity encryptedHmacKey="AA==" encryptedHmacValue="AA==" />
              <keyEncryptors>
                <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
                  <p:encryptedKey saltValue="AAAAAAAAAAAAAAAAAAAAAA==" spinCount="4294967295" hashAlgorithm="SHA1"
                                  cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                                  keyBits="128" blockSize="16"
                                  encryptedVerifierHashInput="AA=="
                                  encryptedVerifierHashValue="AA=="
                                  encryptedKeyValue="AA=="/>
                </keyEncryptor>
              </keyEncryptors>
            </encryption>
        "#;

        let mut encryption_info = Vec::new();
        encryption_info.extend_from_slice(&4u16.to_le_bytes()); // versionMajor
        encryption_info.extend_from_slice(&4u16.to_le_bytes()); // versionMinor
        encryption_info.extend_from_slice(&0x40u32.to_le_bytes()); // flags (arbitrary)
        encryption_info.extend_from_slice(xml.as_bytes());

        let encrypted_package = 0u64.to_le_bytes().to_vec();

        let err = decrypt_encrypted_package_streams_with_options(
            &encryption_info,
            &encrypted_package,
            "password",
            &DecryptOptions::default(),
        )
        .expect_err("expected error");
        assert!(
            matches!(err, OfficeCryptoError::SpinCountTooLarge { .. }),
            "unexpected error: {err:?}"
        );
    }

    #[test]
    fn standard_key_derivation_matches_vector() {
        // Deterministic vector to catch regressions in key derivation.
        let password = "Password";
        let salt: [u8; 16] = [
            0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D,
            0x0E, 0x0F,
        ];
        // AES-128 Standard/CryptoAPI key derivation (CryptDeriveKey-style).
        let deriver = StandardKeyDeriver::new(
            HashAlgorithm::Sha1,
            128,
            &salt,
            password,
            StandardKeyDerivation::Aes,
        );
        let key_block0 = deriver.derive_key_for_block(0).expect("derive key");
        assert_eq!(
            key_block0.as_slice(),
            &[
                0x1B, 0xA0, 0x05, 0x26, 0x1A, 0xAE, 0xE4, 0x68, 0x6A, 0x99, 0x39, 0x43, 0x70, 0x75,
                0xE6, 0xC4,
            ]
        );
    }

    #[test]
    fn standard_key_derivation_rc4_cryptoapi_sha1_vectors() {
        // Deterministic MS-OFFCRYPTO Standard/CryptoAPI RC4 key derivation vectors.
        //
        // These lock in:
        // - password UTF-16LE encoding
        // - initial hash input order: salt || password_utf16le
        // - spin loop: H = SHA1(LE32(i) || H), i in 0..50000
        // - per-block key: key(b) = SHA1(H || LE32(b))[0..keySizeBytes]
        //
        // Using multiple non-zero block indices catches mistakes in block-index encoding/order.
        let password = "password";
        let salt: [u8; 16] = [
            0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D,
            0x0E, 0x0F,
        ];

        let deriver = StandardKeyDeriver::new(
            HashAlgorithm::Sha1,
            128,
            &salt,
            password,
            StandardKeyDerivation::Rc4,
        );
        let expected = [
            (
                0u32,
                [
                    0x6a, 0xd7, 0xde, 0xdf, 0x2d, 0xa3, 0x51, 0x4b, 0x1d, 0x85, 0xea, 0xbe, 0xe0,
                    0x69, 0xd4, 0x7d,
                ],
            ),
            (
                1u32,
                [
                    0x2e, 0xd4, 0xe8, 0x82, 0x5c, 0xd4, 0x8a, 0xa4, 0xa4, 0x79, 0x94, 0xcd, 0xa7,
                    0x41, 0x5b, 0x4a,
                ],
            ),
            (
                2u32,
                [
                    0x9c, 0xe5, 0x7d, 0x06, 0x99, 0xbe, 0x39, 0x38, 0x95, 0x1f, 0x47, 0xfa, 0x94,
                    0x93, 0x61, 0xdb,
                ],
            ),
            (
                3u32,
                [
                    0xe6, 0x5b, 0x26, 0x43, 0xea, 0xba, 0x38, 0x15, 0xa3, 0x7a, 0x61, 0x15, 0x9f,
                    0x13, 0x78, 0x40,
                ],
            ),
        ];

        for (block, expected_key) in expected {
            let key = deriver.derive_key_for_block(block).expect("derive key");
            assert_eq!(key.as_slice(), expected_key.as_slice(), "block={block}");
        }

        // 40-bit key size => 5-byte key truncation.
        let deriver_40 = StandardKeyDeriver::new(
            HashAlgorithm::Sha1,
            40,
            &salt,
            password,
            StandardKeyDerivation::Rc4,
        );
        let key_40 = deriver_40
            .derive_key_for_block(0)
            .expect("derive 40-bit key");
        assert_eq!(key_40.as_slice(), &[0x6a, 0xd7, 0xde, 0xdf, 0x2d]);
    }

    #[test]
    fn oversized_encrypted_package_size_errors_without_large_allocation() {
        let total_size: u64 = if usize::BITS < 64 {
            (usize::MAX as u64) + 1
        } else {
            u64::MAX
        };

        let mut encrypted_package = Vec::new();
        encrypted_package.extend_from_slice(&total_size.to_le_bytes());

        let dummy_standard = standard::StandardEncryptionInfo {
            version_major: 0,
            version_minor: 0,
            flags: 0,
            header: standard::EncryptionHeader {
                alg_id: 0,
                alg_id_hash: 0,
                key_bits: 0,
                provider_type: 0,
                csp_name: String::new(),
            },
            verifier: standard::EncryptionVerifier {
                salt: Vec::new(),
                encrypted_verifier: Vec::new(),
                verifier_hash_size: 0,
                encrypted_verifier_hash: Vec::new(),
            },
        };

        let dummy_agile = agile::AgileEncryptionInfo {
            version_major: 0,
            version_minor: 0,
            flags: 0,
            key_data: agile::AgileKeyData {
                salt: Vec::new(),
                block_size: 16,
                key_bits: 128,
                hash_algorithm: HashAlgorithm::Sha256,
                hash_size: HashAlgorithm::Sha256.digest_len(),
                cipher_algorithm: String::new(),
                cipher_chaining: String::new(),
            },
            data_integrity: Some(agile::AgileDataIntegrity {
                encrypted_hmac_key: Vec::new(),
                encrypted_hmac_value: Vec::new(),
            }),
            password_key_encryptor: agile::AgilePasswordKeyEncryptor {
                salt: Vec::new(),
                block_size: 16,
                key_bits: 128,
                spin_count: 0,
                hash_algorithm: HashAlgorithm::Sha256,
                hash_size: HashAlgorithm::Sha256.digest_len(),
                cipher_algorithm: String::new(),
                cipher_chaining: String::new(),
                encrypted_verifier_hash_input: Vec::new(),
                encrypted_verifier_hash_value: Vec::new(),
                encrypted_key_value: Vec::new(),
            },
        };

        MAX_ALLOC.store(0, Ordering::Relaxed);

        let err =
            standard::decrypt_standard_encrypted_package(&dummy_standard, &encrypted_package, "")
                .expect_err("expected size overflow");
        assert!(
            matches!(
                err,
                OfficeCryptoError::SizeLimitExceededU64 {
                    context: "EncryptedPackage.originalSize",
                    limit
                } if limit == crate::MAX_ENCRYPTED_PACKAGE_ORIGINAL_SIZE
            ),
            "expected SizeLimitExceededU64(EncryptedPackage.originalSize), got {err:?}"
        );

        let err = agile::decrypt_agile_encrypted_package(
            &dummy_agile,
            &encrypted_package,
            "",
            &DecryptOptions::default(),
        )
        .expect_err("expected size overflow");
        assert!(
            matches!(
                err,
                OfficeCryptoError::SizeLimitExceededU64 {
                    context: "EncryptedPackage.originalSize",
                    limit
                } if limit == crate::MAX_ENCRYPTED_PACKAGE_ORIGINAL_SIZE
            ),
            "expected SizeLimitExceededU64(EncryptedPackage.originalSize), got {err:?}"
        );

        let max_alloc = MAX_ALLOC.load(Ordering::Relaxed);
        assert!(
            max_alloc < 16 * 1024 * 1024,
            "expected no large allocation attempts, observed max allocation request: {max_alloc} bytes"
        );
    }

    #[test]
    #[cfg(not(target_arch = "wasm32"))]
    fn decrypts_repo_standard_basic_xlsm_fixture() {
        fn fixture_path(rel: &str) -> PathBuf {
            Path::new(concat!(env!("CARGO_MANIFEST_DIR"), "/../../fixtures/encrypted/ooxml/"))
                .join(rel)
        }

        let encrypted =
            std::fs::read(fixture_path("standard-basic.xlsm")).expect("read standard-basic.xlsm");
        let expected = std::fs::read(fixture_path("plaintext-basic.xlsm"))
            .expect("read plaintext-basic.xlsm");

        let decrypted =
            decrypt_encrypted_package_ole(&encrypted, "password").expect("decrypt standard-basic");
        assert!(decrypted.starts_with(b"PK"));
        assert_eq!(decrypted, expected);
    }
}

#[cfg(all(test, not(target_arch = "wasm32")))]
mod fuzz_tests;

#[cfg(test)]
mod test_alloc {
    use std::alloc::{GlobalAlloc, Layout, System};
    use std::sync::atomic::{AtomicUsize, Ordering};

    pub static MAX_ALLOC: AtomicUsize = AtomicUsize::new(0);

    pub struct TrackingAllocator;

    unsafe impl GlobalAlloc for TrackingAllocator {
        unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
            record(layout.size());
            System.alloc(layout)
        }

        unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
            record(layout.size());
            System.alloc_zeroed(layout)
        }

        unsafe fn realloc(&self, ptr: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
            record(new_size);
            System.realloc(ptr, layout, new_size)
        }

        unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
            System.dealloc(ptr, layout)
        }
    }

    #[inline]
    fn record(size: usize) {
        let mut prev = MAX_ALLOC.load(Ordering::Relaxed);
        while size > prev {
            match MAX_ALLOC.compare_exchange_weak(prev, size, Ordering::Relaxed, Ordering::Relaxed)
            {
                Ok(_) => break,
                Err(next) => prev = next,
            }
        }
    }

    // Ensure tests can assert that huge `total_size` values are rejected *before*
    // attempting allocations.
    #[global_allocator]
    static GLOBAL: TrackingAllocator = TrackingAllocator;
}
