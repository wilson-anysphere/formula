//! MS-OFFCRYPTO parsing and crypto utilities.
//!
//! This crate currently supports:
//! - Parsing the *Standard* (CryptoAPI) `EncryptionInfo` stream header (`versionMinor == 2`;
//!   commonly 3.2, but `versionMajor ∈ {2,3,4}` is observed in the wild)
//! - Parsing the *Agile* `EncryptionInfo` stream (version 4.4) (password key-encryptor subset)
//! - Parsing the `EncryptedPackage` stream header
//! - ECMA-376 Standard password→key derivation + verifier checks
//! - Decrypting Standard-encrypted OOXML packages via [`decrypt_standard_ooxml_from_bytes`]
//!
//! Verifier digests are compared in constant time to reduce timing side channels.

mod util;

pub mod agile;
pub mod standard;

use core::fmt;
use std::io::{Cursor, Read};

use aes::cipher::{generic_array::GenericArray, BlockDecrypt, KeyInit};
use aes::{Aes128, Aes192, Aes256};

use base64::engine::general_purpose::{STANDARD, STANDARD_NO_PAD};
use base64::Engine;
use cbc::Decryptor;
use cipher::{block_padding::NoPadding, BlockDecryptMut, KeyIvInit};
use quick_xml::events::Event as XmlEvent;
use quick_xml::Reader as XmlReader;
use sha1::{Digest as _, Sha1};
use zeroize::Zeroizing;

const ITER_COUNT: u32 = 50_000;
const SHA1_LEN: usize = 20;
const MAX_DIGEST_LEN: usize = 64; // SHA-512
const AES_BLOCK_SIZE: usize = 16;

// Agile Encryption guardrails (MS-OFFCRYPTO uses 16-byte salts and AES block alignment).
const AGILE_SALT_LEN: usize = 16;
const AGILE_MAX_ENCRYPTED_LEN: usize = 64;

/// Recommended default upper bound for Agile `spinCount`.
///
/// Excel commonly uses `100_000` for Agile encryption. Allowing up to `1_000_000`
/// preserves compatibility while preventing pathological attacker-controlled values
/// (e.g. `u32::MAX`) from hanging the process.
pub const DEFAULT_MAX_SPIN_COUNT: u32 = 1_000_000;

/// Limits to apply during decryption to prevent resource exhaustion (DoS).
#[derive(Debug, Clone)]
pub struct DecryptLimits {
    /// Maximum allowed Agile `spinCount` value.
    ///
    /// `None` disables the limit.
    pub max_spin_count: Option<u32>,
}

impl Default for DecryptLimits {
    fn default() -> Self {
        Self {
            max_spin_count: Some(DEFAULT_MAX_SPIN_COUNT),
        }
    }
}

/// Options controlling decryption behavior.
#[derive(Debug, Clone)]
pub struct DecryptOptions {
    pub limits: DecryptLimits,
}

impl Default for DecryptOptions {
    fn default() -> Self {
        Self {
            limits: DecryptLimits::default(),
        }
    }
}

fn check_spin_count(spin_count: u32, limits: &DecryptLimits) -> Result<(), OffcryptoError> {
    if let Some(max) = limits.max_spin_count {
        if spin_count > max {
            return Err(OffcryptoError::SpinCountTooLarge { spin_count, max });
        }
    }
    Ok(())
}

pub mod encrypted_package;
pub use encrypted_package::{agile_decrypt_package, decrypt_encrypted_package};

const PASSWORD_KEY_ENCRYPTOR_NS: &str =
    "http://schemas.microsoft.com/office/2006/keyEncryptor/password";

// CryptoAPI algorithm identifiers used by Standard encryption.
const CALG_AES_128: u32 = 0x0000_660E;
const CALG_AES_192: u32 = 0x0000_660F;
const CALG_AES_256: u32 = 0x0000_6610;
const CALG_SHA1: u32 = 0x0000_8004;

/// Parsed `EncryptionVersionInfo` (MS-OFFCRYPTO).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct EncryptionVersionInfo {
    pub major: u16,
    pub minor: u16,
    pub flags: u32,
}

impl EncryptionVersionInfo {
    /// Parse the MS-OFFCRYPTO `EncryptionVersionInfo` header from an `EncryptionInfo` stream.
    pub fn parse(encryption_info_stream: &[u8]) -> Result<Self, OffcryptoError> {
        let mut r = Reader::new(encryption_info_stream);
        let major = r.read_u16_le("EncryptionVersionInfo.major")?;
        let minor = r.read_u16_le("EncryptionVersionInfo.minor")?;
        let flags = r.read_u32_le("EncryptionVersionInfo.flags")?;
        Ok(Self { major, minor, flags })
    }
}

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
            match MAX_ALLOC.compare_exchange_weak(
                prev,
                size,
                Ordering::Relaxed,
                Ordering::Relaxed,
            ) {
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

/// Parsed Standard (CryptoAPI) `EncryptionHeader`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StandardEncryptionHeader {
    pub flags: u32,
    pub size_extra: u32,
    pub alg_id: u32,
    pub alg_id_hash: u32,
    pub key_size_bits: u32,
    pub provider_type: u32,
    pub reserved1: u32,
    pub reserved2: u32,
    pub csp_name: String,
}

/// Parsed Standard (CryptoAPI) `EncryptionVerifier`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StandardEncryptionVerifier {
    pub salt: Vec<u8>,
    pub encrypted_verifier: [u8; 16],
    pub verifier_hash_size: u32,
    pub encrypted_verifier_hash: Vec<u8>,
}

/// Parsed Standard (CryptoAPI) `EncryptionInfo`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StandardEncryptionInfo {
    pub header: StandardEncryptionHeader,
    pub verifier: StandardEncryptionVerifier,
}

fn validate_standard_encryption_info(info: &StandardEncryptionInfo) -> Result<(), OffcryptoError> {
    let expected_key_size = match info.header.alg_id {
        CALG_AES_128 => 128,
        CALG_AES_192 => 192,
        CALG_AES_256 => 256,
        other => {
            return Err(OffcryptoError::UnsupportedAlgorithm(format!(
                "algId=0x{other:08x}"
            )))
        }
    };

    if info.header.key_size_bits != expected_key_size {
        // Mirror the parsing behaviour: mismatch means we don't support the declared algorithm
        // parameters.
        return Err(OffcryptoError::UnsupportedAlgorithm(format!(
            "keySize={} mismatch for algId=0x{:08x}",
            info.header.key_size_bits, info.header.alg_id
        )));
    }

    if info.header.alg_id_hash != CALG_SHA1 {
        return Err(OffcryptoError::UnsupportedAlgorithm(format!(
            "algIdHash=0x{:08x}",
            info.header.alg_id_hash
        )));
    }

    if info.verifier.salt.len() != 16 {
        return Err(OffcryptoError::InvalidEncryptionInfo {
            context: "EncryptionVerifier.saltSize must be 16 for Standard encryption",
        });
    }

    if info.verifier.verifier_hash_size != 20 {
        return Err(OffcryptoError::InvalidEncryptionInfo {
            context: "EncryptionVerifier.verifierHashSize must be 20 (SHA1) for Standard encryption",
        });
    }

    Ok(())
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HashAlgorithm {
    Sha1,
    Sha256,
    Sha384,
    Sha512,
}

impl HashAlgorithm {
    fn parse_offcrypto_name(name: &str) -> Result<Self, OffcryptoError> {
        match name.trim().to_ascii_uppercase().as_str() {
            "SHA1" | "SHA-1" => Ok(HashAlgorithm::Sha1),
            "SHA256" | "SHA-256" => Ok(HashAlgorithm::Sha256),
            "SHA384" | "SHA-384" => Ok(HashAlgorithm::Sha384),
            "SHA512" | "SHA-512" => Ok(HashAlgorithm::Sha512),
            _ => Err(OffcryptoError::InvalidEncryptionInfo {
                context: "unsupported hashAlgorithm",
            }),
        }
    }

    pub(crate) fn digest_len(self) -> usize {
        match self {
            HashAlgorithm::Sha1 => 20,
            HashAlgorithm::Sha256 => 32,
            HashAlgorithm::Sha384 => 48,
            HashAlgorithm::Sha512 => 64,
        }
    }

    pub(crate) fn digest_into(self, data: &[u8], out: &mut [u8]) {
        debug_assert!(out.len() >= self.digest_len());
        match self {
            HashAlgorithm::Sha1 => {
                let mut hasher = Sha1::new();
                hasher.update(data);
                out[..20].copy_from_slice(&hasher.finalize());
            }
            HashAlgorithm::Sha256 => {
                let mut hasher = sha2::Sha256::new();
                hasher.update(data);
                out[..32].copy_from_slice(&hasher.finalize());
            }
            HashAlgorithm::Sha384 => {
                let mut hasher = sha2::Sha384::new();
                hasher.update(data);
                out[..48].copy_from_slice(&hasher.finalize());
            }
            HashAlgorithm::Sha512 => {
                let mut hasher = sha2::Sha512::new();
                hasher.update(data);
                out[..64].copy_from_slice(&hasher.finalize());
            }
        }
    }

    pub(crate) fn digest_two_into(self, a: &[u8], b: &[u8], out: &mut [u8]) {
        debug_assert!(out.len() >= self.digest_len());
        match self {
            HashAlgorithm::Sha1 => {
                let mut hasher = Sha1::new();
                hasher.update(a);
                hasher.update(b);
                out[..20].copy_from_slice(&hasher.finalize());
            }
            HashAlgorithm::Sha256 => {
                let mut hasher = sha2::Sha256::new();
                hasher.update(a);
                hasher.update(b);
                out[..32].copy_from_slice(&hasher.finalize());
            }
            HashAlgorithm::Sha384 => {
                let mut hasher = sha2::Sha384::new();
                hasher.update(a);
                hasher.update(b);
                out[..48].copy_from_slice(&hasher.finalize());
            }
            HashAlgorithm::Sha512 => {
                let mut hasher = sha2::Sha512::new();
                hasher.update(a);
                hasher.update(b);
                out[..64].copy_from_slice(&hasher.finalize());
            }
        }
    }

    pub(crate) fn digest(self, data: &[u8]) -> Vec<u8> {
        let mut out = vec![0u8; self.digest_len()];
        self.digest_into(data, &mut out);
        out
    }
}

/// Parsed contents of an Agile (XML) `EncryptionInfo` stream, restricted to the subset required
/// for password-based decryption.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AgileEncryptionInfo {
    pub key_data_salt: Vec<u8>,
    pub key_data_hash_algorithm: HashAlgorithm,
    pub key_data_block_size: usize,

    pub encrypted_hmac_key: Vec<u8>,
    pub encrypted_hmac_value: Vec<u8>,

    // Password key encryptor fields (`p:encryptedKey`).
    pub spin_count: u32,
    pub password_salt: Vec<u8>,
    pub password_hash_algorithm: HashAlgorithm,
    pub password_key_bits: usize,
    pub encrypted_key_value: Vec<u8>,
    pub encrypted_verifier_hash_input: Vec<u8>,
    pub encrypted_verifier_hash_value: Vec<u8>,
}

/// Parsed `EncryptionInfo`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum EncryptionInfo {
    /// Standard (CryptoAPI) encryption (MS-OFFCRYPTO `versionMinor == 2`; commonly 3.2).
    Standard {
        version: EncryptionVersionInfo,
        header: StandardEncryptionHeader,
        verifier: StandardEncryptionVerifier,
    },
    /// Agile (XML) encryption (MS-OFFCRYPTO version 4.4).
    Agile {
        version: EncryptionVersionInfo,
        info: AgileEncryptionInfo,
    },
    /// A version we do not yet support.
    Unsupported { version: EncryptionVersionInfo },
}

/// Header for the `EncryptedPackage` stream.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct EncryptedPackageHeader {
    /// Original unencrypted package size in bytes.
    pub original_size: u64,
}

/// Errors returned by this crate.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum OffcryptoError {
    /// Not enough bytes to parse the requested structure.
    Truncated { context: &'static str },
    /// Input bytes were structurally invalid.
    InvalidFormat { context: &'static str },
    /// CSPName was not valid UTF-16LE.
    InvalidCspNameUtf16,
    /// The encrypted streams are structurally invalid.
    InvalidStructure(String),
    /// Standard encryption uses an algorithm not supported by the current implementation.
    UnsupportedAlgorithm(String),
    /// The stream contents are structurally invalid (e.g. missing required attributes).
    InvalidEncryptionInfo { context: &'static str },
    /// The decrypted package size from the `EncryptedPackage` header does not fit into a `Vec<u8>`.
    EncryptedPackageSizeOverflow { total_size: u64 },
    /// Failed to reserve memory for the decrypted output buffer.
    EncryptedPackageAllocationFailed { total_size: u64 },
    /// `EncryptedPackage` declared plaintext size is not plausible for the available ciphertext.
    ///
    /// For AES-based Office encryption, ciphertext is padded to the AES block size (16 bytes), so
    /// the ciphertext length must be at least `ceil(total_size / 16) * 16`.
    EncryptedPackageSizeMismatch { total_size: u64, ciphertext_len: usize },
    /// The `EncryptionInfo` version is not supported by the current parser.
    UnsupportedVersion { major: u16, minor: u16 },
    /// The encryption schema is known but not supported by the selected decryption mode.
    ///
    /// For example: attempting to decrypt an Agile-encrypted OOXML package using a Standard-only
    /// decryptor.
    UnsupportedEncryption { encryption_type: EncryptionType },
    /// Ciphertext length must be a multiple of 16 bytes for AES-ECB.
    InvalidCiphertextLength { len: usize },
    /// Invalid AES key length (expected 16, 24, or 32 bytes).
    InvalidKeyLength { len: usize },
    /// Standard encryption keySize must be a multiple of 8 bits.
    InvalidKeySizeBits { key_size_bits: u32 },
    /// The requested key size is larger than the 40-byte derivation output.
    DerivedKeyTooLong {
        key_size_bits: u32,
        required_bytes: usize,
        available_bytes: usize,
    },
    /// Decrypted verifier hash is too short.
    InvalidVerifierHashLength { len: usize },
    /// Password/key did not pass verifier check.
    InvalidPassword,
    /// Agile `spinCount` is larger than allowed by the configured decryption limits.
    SpinCountTooLarge { spin_count: u32, max: u32 },
}

impl fmt::Display for OffcryptoError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            OffcryptoError::Truncated { context } => {
                write!(f, "truncated data while reading {context}")
            }
            OffcryptoError::InvalidFormat { context } => write!(f, "invalid format: {context}"),
            OffcryptoError::InvalidCspNameUtf16 => write!(f, "invalid UTF-16LE CSPName"),
            OffcryptoError::InvalidStructure(msg) => write!(f, "invalid structure: {msg}"),
            OffcryptoError::UnsupportedAlgorithm(msg) => write!(f, "unsupported algorithm: {msg}"),
            OffcryptoError::InvalidEncryptionInfo { context } => {
                write!(f, "invalid EncryptionInfo: {context}")
            }
            OffcryptoError::EncryptedPackageSizeOverflow { total_size } => write!(
                f,
                "EncryptedPackage reported invalid original size {total_size}"
            ),
            OffcryptoError::EncryptedPackageAllocationFailed { total_size } => {
                write!(f, "failed to allocate decrypted package buffer of size {total_size}")
            }
            OffcryptoError::EncryptedPackageSizeMismatch {
                total_size,
                ciphertext_len,
            } => write!(
                f,
                "EncryptedPackage declared original size {total_size} exceeds ciphertext length {ciphertext_len}"
            ),
            OffcryptoError::UnsupportedVersion { major, minor } => {
                write!(f, "unsupported EncryptionInfo version {major}.{minor}")
            }
            OffcryptoError::UnsupportedEncryption { encryption_type } => {
                write!(f, "unsupported encryption type {encryption_type:?}")
            }
            OffcryptoError::InvalidCiphertextLength { len } => write!(
                f,
                "ciphertext length must be a multiple of 16 bytes for AES-ECB, got {len}"
            ),
            OffcryptoError::InvalidKeyLength { len } => write!(
                f,
                "invalid AES key length {len}; expected 16, 24, or 32 bytes"
            ),
            OffcryptoError::InvalidKeySizeBits { key_size_bits } => write!(
                f,
                "standard encryption keySize must be a multiple of 8 bits, got {key_size_bits}"
            ),
            OffcryptoError::DerivedKeyTooLong {
                key_size_bits,
                required_bytes,
                available_bytes,
            } => write!(
                f,
                "keySize ({key_size_bits} bits) requires {required_bytes} bytes, but the SHA1-based derivation output is only {available_bytes} bytes"
            ),
            OffcryptoError::InvalidVerifierHashLength { len } => write!(
                f,
                "encrypted verifier hash must be at least 20 bytes after decryption, got {len}"
            ),
            OffcryptoError::InvalidPassword => write!(f, "invalid password or key"),
            OffcryptoError::SpinCountTooLarge { spin_count, max } => {
                write!(f, "Agile spinCount too large: {spin_count} (max {max})")
            }
        }
    }
}

impl std::error::Error for OffcryptoError {}

struct Reader<'a> {
    bytes: &'a [u8],
    pos: usize,
}

impl<'a> Reader<'a> {
    fn new(bytes: &'a [u8]) -> Self {
        Self { bytes, pos: 0 }
    }

    fn remaining(&self) -> &'a [u8] {
        &self.bytes[self.pos..]
    }

    fn take(&mut self, n: usize, context: &'static str) -> Result<&'a [u8], OffcryptoError> {
        let end = self.pos.saturating_add(n);
        if end > self.bytes.len() {
            return Err(OffcryptoError::Truncated { context });
        }
        let out = &self.bytes[self.pos..end];
        self.pos = end;
        Ok(out)
    }

    fn read_u16_le(&mut self, context: &'static str) -> Result<u16, OffcryptoError> {
        let b = self.take(2, context)?;
        Ok(u16::from_le_bytes([b[0], b[1]]))
    }

    fn read_u32_le(&mut self, context: &'static str) -> Result<u32, OffcryptoError> {
        let b = self.take(4, context)?;
        Ok(u32::from_le_bytes([b[0], b[1], b[2], b[3]]))
    }

    fn read_u64_le(&mut self, context: &'static str) -> Result<u64, OffcryptoError> {
        let b = self.take(8, context)?;
        Ok(u64::from_le_bytes([
            b[0], b[1], b[2], b[3], b[4], b[5], b[6], b[7],
        ]))
    }
}

fn decode_csp_name_utf16le(bytes: &[u8]) -> Result<String, OffcryptoError> {
    if bytes.is_empty() {
        return Ok(String::new());
    }
    if bytes.len() % 2 != 0 {
        return Err(OffcryptoError::Truncated {
            context: "cspName UTF-16LE",
        });
    }

    let mut code_units: Vec<u16> = Vec::with_capacity(bytes.len() / 2);
    for chunk in bytes.chunks_exact(2) {
        code_units.push(u16::from_le_bytes([chunk[0], chunk[1]]));
    }

    let end = if let Some(nul_pos) = code_units.iter().position(|u| *u == 0) {
        nul_pos
    } else {
        // Be tolerant of a missing terminator: trim trailing NULs but otherwise use
        // the full remaining buffer.
        let mut end = code_units.len();
        while end > 0 && code_units[end - 1] == 0 {
            end -= 1;
        }
        end
    };

    String::from_utf16(&code_units[..end]).map_err(|_| OffcryptoError::InvalidCspNameUtf16)
}

/// Parse an MS-OFFCRYPTO `EncryptionInfo` stream header.
pub fn parse_encryption_info(bytes: &[u8]) -> Result<EncryptionInfo, OffcryptoError> {
    let mut r = Reader::new(bytes);
    let major = r.read_u16_le("EncryptionVersionInfo.major")?;
    let minor = r.read_u16_le("EncryptionVersionInfo.minor")?;
    let flags = r.read_u32_le("EncryptionVersionInfo.flags")?;
    let version = EncryptionVersionInfo { major, minor, flags };

    if (major, minor) == (4, 4) {
        // Agile EncryptionInfo payload is an UTF-8 XML document beginning at byte offset 8.
        let info = parse_agile_encryption_info_xml(r.remaining())?;
        return Ok(EncryptionInfo::Agile { version, info });
    }

    // MS-OFFCRYPTO / ECMA-376 identifies "Standard" encryption via `versionMinor == 2`, but
    // real-world files vary `versionMajor` across 2/3/4 (see nolze/msoffcrypto-tool).
    //
    // Treat everything else (including "Extensible" encryption, versionMinor == 3) as unsupported
    // for now so callers can surface an actionable error.
    let is_standard = minor == 2 && matches!(major, 2 | 3 | 4);
    if !is_standard {
        return Ok(EncryptionInfo::Unsupported { version });
    }

    let header_size = r.read_u32_le("EncryptionInfo.header_size")? as usize;
    // Standard `EncryptionHeader` has a fixed 8-DWORD prefix (32 bytes). Reject header sizes that
    // are clearly invalid (as opposed to merely truncated inputs).
    const MIN_STANDARD_HEADER_SIZE: usize = 8 * 4;
    const MAX_STANDARD_HEADER_SIZE: usize = 1024 * 1024; // 1MiB: far larger than any real CSP name.
    if header_size < MIN_STANDARD_HEADER_SIZE || header_size > MAX_STANDARD_HEADER_SIZE {
        return Err(OffcryptoError::InvalidEncryptionInfo {
            context: "EncryptionInfo.header_size is out of bounds",
        });
    }

    let header_bytes = r.take(header_size, "EncryptionHeader")?;

    let mut hr = Reader::new(header_bytes);
    let header = StandardEncryptionHeader {
        flags: hr.read_u32_le("EncryptionHeader.flags")?,
        size_extra: hr.read_u32_le("EncryptionHeader.sizeExtra")?,
        alg_id: hr.read_u32_le("EncryptionHeader.algId")?,
        alg_id_hash: hr.read_u32_le("EncryptionHeader.algIdHash")?,
        key_size_bits: hr.read_u32_le("EncryptionHeader.keySize")?,
        provider_type: hr.read_u32_le("EncryptionHeader.providerType")?,
        reserved1: hr.read_u32_le("EncryptionHeader.reserved1")?,
        reserved2: hr.read_u32_le("EncryptionHeader.reserved2")?,
        csp_name: decode_csp_name_utf16le(hr.remaining())?,
    };

    // Algorithm/parameter validation.
    //
    // Standard encryption produced by Excel uses AES + SHA1. Restrict the parser to this subset
    // so downstream crypto code can rely on the parameters being consistent.
    let expected_key_size = match header.alg_id {
        CALG_AES_128 => 128,
        CALG_AES_192 => 192,
        CALG_AES_256 => 256,
        other => {
            return Err(OffcryptoError::UnsupportedAlgorithm(format!(
                "algId=0x{other:08x}"
            )))
        }
    };
    if header.key_size_bits != expected_key_size {
        return Err(OffcryptoError::UnsupportedAlgorithm(format!(
            "keySize={} mismatch for algId=0x{:08x}",
            header.key_size_bits, header.alg_id
        )));
    }
    if header.alg_id_hash != CALG_SHA1 {
        return Err(OffcryptoError::UnsupportedAlgorithm(format!(
            "algIdHash=0x{:08x}",
            header.alg_id_hash
        )));
    }

    // EncryptionVerifier occupies the remaining bytes after the header.
    let salt_size = r.read_u32_le("EncryptionVerifier.saltSize")? as usize;
    if salt_size != 16 {
        return Err(OffcryptoError::InvalidEncryptionInfo {
            context: "EncryptionVerifier.saltSize must be 16 for Standard encryption",
        });
    }
    let salt = r.take(16, "EncryptionVerifier.salt")?.to_vec();

    let enc_ver = r.take(16, "EncryptionVerifier.encryptedVerifier")?;
    let mut encrypted_verifier = [0u8; 16];
    encrypted_verifier.copy_from_slice(enc_ver);

    let verifier_hash_size = r.read_u32_le("EncryptionVerifier.verifierHashSize")?;
    if verifier_hash_size != 20 {
        return Err(OffcryptoError::InvalidEncryptionInfo {
            context: "EncryptionVerifier.verifierHashSize must be 20 (SHA1) for Standard encryption",
        });
    }
    // SHA1 hashes are 20 bytes, padded to an AES block boundary (16) => 32 bytes.
    let encrypted_verifier_hash = r
        .take(32, "EncryptionVerifier.encryptedVerifierHash")?
        .to_vec();

    let verifier = StandardEncryptionVerifier {
        salt,
        encrypted_verifier,
        verifier_hash_size,
        encrypted_verifier_hash,
    };

    Ok(EncryptionInfo::Standard {
        version,
        header,
        verifier,
    })
}

/// Decrypt an `EncryptedPackage` stream in **Standard-only** mode.
///
/// This helper is intended for callers that only implement ECMA-376 Standard encryption.
/// If the provided `EncryptionInfo` stream describes Agile encryption, this returns
/// [`OffcryptoError::UnsupportedEncryptionType`] (even if the password is correct).
///
/// Inputs are the raw `EncryptionInfo` and `EncryptedPackage` *stream bytes* extracted from the
/// OLE/CFB wrapper.
pub fn decrypt_standard_only(
    encryption_info: &[u8],
    encrypted_package: &[u8],
    password: &str,
) -> Result<Vec<u8>, OffcryptoError> {
    let info = parse_encryption_info(encryption_info)?;
    match info {
        EncryptionInfo::Standard {
            header, verifier, ..
        } => {
            let info = StandardEncryptionInfo { header, verifier };
            // Derived keys are sensitive; keep them in a `Zeroizing` buffer so failed password
            // attempts don't leave key material lingering in heap allocations.
            let key = Zeroizing::new(standard_derive_key(&info, password)?);
            standard_verify_key(&info, key.as_slice())?;

            encrypted_package::decrypt_encrypted_package(encrypted_package, |_idx, ct, pt| {
                pt.copy_from_slice(ct);
                aes_ecb_decrypt_in_place(key.as_slice(), pt)
            })
        }
        EncryptionInfo::Agile { .. } => Err(OffcryptoError::UnsupportedEncryption {
            encryption_type: EncryptionType::Agile,
        }),
        EncryptionInfo::Unsupported { version } => {
            if version.minor == 3 && matches!(version.major, 3 | 4) {
                // "Extensible" encryption (MS-OFFCRYPTO): known scheme, but not supported by the
                // Standard-only decrypt entrypoint.
                Err(OffcryptoError::UnsupportedEncryption {
                    encryption_type: EncryptionType::Extensible,
                })
            } else {
                Err(OffcryptoError::UnsupportedVersion {
                    major: version.major,
                    minor: version.minor,
                })
            }
        }
    }
}
#[derive(Debug, Clone)]
struct NamespaceFrame {
    decls: Vec<(Vec<u8> /* prefix */, Vec<u8> /* uri */)>,
}

fn push_namespace_frame<'a>(
    stack: &mut Vec<NamespaceFrame>,
    elem: &quick_xml::events::BytesStart<'a>,
) -> Result<(), OffcryptoError> {
    let mut frame = NamespaceFrame { decls: Vec::new() };

    for attr in elem.attributes().with_checks(false) {
        let attr = attr.map_err(|_| OffcryptoError::InvalidEncryptionInfo {
            context: "invalid XML attribute",
        })?;
        let key = attr.key.as_ref();
        let value = attr.value.as_ref();

        if key == b"xmlns" {
            frame.decls.push((Vec::new(), value.to_vec()));
        } else if let Some(prefix) = key.strip_prefix(b"xmlns:") {
            frame.decls.push((prefix.to_vec(), value.to_vec()));
        }
    }

    stack.push(frame);
    Ok(())
}

fn pop_namespace_frame(stack: &mut Vec<NamespaceFrame>) {
    stack.pop();
}

fn resolve_namespace_uri<'a>(stack: &'a [NamespaceFrame], prefix: &[u8]) -> Option<&'a [u8]> {
    for frame in stack.iter().rev() {
        for (p, uri) in &frame.decls {
            if p.as_slice() == prefix {
                return Some(uri.as_slice());
            }
        }
    }
    None
}

fn element_prefix(name: &[u8]) -> &[u8] {
    name.iter()
        .rposition(|b| *b == b':')
        .map(|idx| &name[..idx])
        .unwrap_or(&[])
}

fn local_name(name: &[u8]) -> &[u8] {
    name.iter()
        .rposition(|b| *b == b':')
        .map(|idx| &name[idx + 1..])
        .unwrap_or(name)
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

fn strip_utf8_bom(bytes: &[u8]) -> &[u8] {
    bytes.strip_prefix(&[0xEF, 0xBB, 0xBF]).unwrap_or(bytes)
}

fn is_nul_heavy(bytes: &[u8]) -> bool {
    if bytes.is_empty() {
        return false;
    }
    let zeros = bytes.iter().filter(|&&b| b == 0).count();
    zeros > bytes.len() / 8
}

fn decode_utf16le_xml(bytes: &[u8]) -> Result<String, OffcryptoError> {
    let mut bytes = bytes;
    // Trim trailing UTF-16LE NUL terminators / padding.
    while bytes.len() >= 2 {
        let n = bytes.len();
        if bytes[n - 2] == 0 && bytes[n - 1] == 0 {
            bytes = &bytes[..n - 2];
        } else {
            break;
        }
    }

    if bytes.starts_with(&[0xFF, 0xFE]) {
        bytes = &bytes[2..];
    }

    // UTF-16 requires an even number of bytes; ignore a trailing odd byte.
    bytes = &bytes[..bytes.len().saturating_sub(bytes.len() % 2)];

    let mut units = Vec::with_capacity(bytes.len() / 2);
    for chunk in bytes.chunks_exact(2) {
        units.push(u16::from_le_bytes([chunk[0], chunk[1]]));
    }

    let mut xml = String::from_utf16(&units).map_err(|_| OffcryptoError::InvalidEncryptionInfo {
        context: "agile EncryptionInfo XML is not valid UTF-16LE",
    })?;

    // Be tolerant of a BOM encoded as U+FEFF.
    if let Some(stripped) = xml.strip_prefix('\u{FEFF}') {
        xml = stripped.to_string();
    }
    while xml.ends_with('\0') {
        xml.pop();
    }
    Ok(xml)
}

fn length_prefixed_slice(payload: &[u8]) -> Option<&[u8]> {
    let len_bytes: [u8; 4] = payload.get(0..4)?.try_into().ok()?;
    let len = u32::from_le_bytes(len_bytes) as usize;
    if len == 0 || len > payload.len().saturating_sub(4) {
        return None;
    }
    let candidate = payload.get(4..4 + len)?;

    // Ensure the candidate *looks* like XML to avoid false positives on arbitrary binary data.
    let candidate_trimmed = trim_start_ascii_whitespace(candidate);
    let candidate_trimmed = strip_utf8_bom(candidate_trimmed);
    let looks_like_utf8 = candidate_trimmed.first() == Some(&b'<');
    let looks_like_utf16le = candidate_trimmed.starts_with(&[0xFF, 0xFE])
        || (candidate_trimmed.len() >= 2 && candidate_trimmed[0] == b'<' && candidate_trimmed[1] == 0);
    if !(looks_like_utf8 || looks_like_utf16le) {
        return None;
    }

    Some(candidate)
}

fn scan_to_first_xml_tag(payload: &[u8]) -> Option<&[u8]> {
    // Only scan when we see the expected root tag later; this keeps the heuristic conservative.
    const NEEDLE: &[u8] = b"<encryption";
    if !payload
        .windows(NEEDLE.len())
        .any(|w| w.eq_ignore_ascii_case(NEEDLE))
    {
        return None;
    }

    let payload = strip_utf8_bom(payload);
    let trimmed = trim_start_ascii_whitespace(payload);
    if trimmed.first() == Some(&b'<') {
        return None;
    }

    let idx = payload.iter().position(|&b| b == b'<')?;
    Some(&payload[idx..])
}

fn try_parse_agile_xml_utf8(bytes: &[u8]) -> Result<AgileEncryptionInfo, OffcryptoError> {
    let bytes = trim_trailing_nul_bytes(bytes);
    let bytes = strip_utf8_bom(bytes);
    let xml = std::str::from_utf8(bytes).map_err(|_| OffcryptoError::InvalidEncryptionInfo {
        context: "agile EncryptionInfo XML is not valid UTF-8",
    })?;
    let xml = xml.strip_prefix('\u{FEFF}').unwrap_or(xml);
    parse_agile_encryption_info_xml_str(xml)
}

fn try_parse_agile_xml_utf16le(bytes: &[u8]) -> Result<AgileEncryptionInfo, OffcryptoError> {
    let xml = decode_utf16le_xml(bytes)?;
    parse_agile_encryption_info_xml_str(&xml)
}

fn parse_agile_encryption_info_xml(xml_bytes: &[u8]) -> Result<AgileEncryptionInfo, OffcryptoError> {
    // Primary: treat remainder as UTF-8 XML (trim UTF-8 BOM, trim trailing NULs).
    //
    // Fallbacks:
    // - NUL-heavy => UTF-16LE decode
    // - scan forward to `<` when `<encryption` appears later
    // - optional 4-byte length prefix
    let mut last_err = match try_parse_agile_xml_utf8(xml_bytes) {
        Ok(info) => return Ok(info),
        Err(err) => err,
    };

    if is_nul_heavy(xml_bytes) {
        match try_parse_agile_xml_utf16le(xml_bytes) {
            Ok(info) => return Ok(info),
            Err(err) => last_err = err,
        }
    }

    if let Some(scanned) = scan_to_first_xml_tag(xml_bytes) {
        match try_parse_agile_xml_utf8(scanned) {
            Ok(info) => return Ok(info),
            Err(err) => last_err = err,
        }
        if is_nul_heavy(scanned) {
            match try_parse_agile_xml_utf16le(scanned) {
                Ok(info) => return Ok(info),
                Err(err) => last_err = err,
            }
        }
    }

    if let Some(len_slice) = length_prefixed_slice(xml_bytes) {
        match try_parse_agile_xml_utf8(len_slice) {
            Ok(info) => return Ok(info),
            Err(err) => last_err = err,
        }
        if is_nul_heavy(len_slice) {
            match try_parse_agile_xml_utf16le(len_slice) {
                Ok(info) => return Ok(info),
                Err(err) => last_err = err,
            }
        }
    }

    Err(last_err)
}

fn parse_agile_encryption_info_xml_str(xml: &str) -> Result<AgileEncryptionInfo, OffcryptoError> {

    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut ns_stack: Vec<NamespaceFrame> = Vec::new();

    let mut key_data_salt: Option<Vec<u8>> = None;
    let mut key_data_hash_algorithm: Option<HashAlgorithm> = None;
    let mut key_data_block_size: Option<usize> = None;

    let mut encrypted_hmac_key: Option<Vec<u8>> = None;
    let mut encrypted_hmac_value: Option<Vec<u8>> = None;

    let mut spin_count: Option<u32> = None;
    let mut password_salt: Option<Vec<u8>> = None;
    let mut password_hash_algorithm: Option<HashAlgorithm> = None;
    let mut password_key_bits: Option<usize> = None;
    let mut encrypted_key_value: Option<Vec<u8>> = None;
    let mut encrypted_verifier_hash_input: Option<Vec<u8>> = None;
    let mut encrypted_verifier_hash_value: Option<Vec<u8>> = None;

    loop {
        let event = reader.read_event_into(&mut buf).map_err(|_| {
            OffcryptoError::InvalidEncryptionInfo {
                context: "agile EncryptionInfo XML parse error",
            }
        })?;

        match event {
            XmlEvent::Start(e) => {
                push_namespace_frame(&mut ns_stack, &e)?;
                parse_agile_element(
                    &mut ns_stack,
                    &e,
                    &mut key_data_salt,
                    &mut key_data_hash_algorithm,
                    &mut key_data_block_size,
                    &mut encrypted_hmac_key,
                    &mut encrypted_hmac_value,
                    &mut spin_count,
                    &mut password_salt,
                    &mut password_hash_algorithm,
                    &mut password_key_bits,
                    &mut encrypted_key_value,
                    &mut encrypted_verifier_hash_input,
                    &mut encrypted_verifier_hash_value,
                )?;
            }
            XmlEvent::Empty(e) => {
                push_namespace_frame(&mut ns_stack, &e)?;
                parse_agile_element(
                    &mut ns_stack,
                    &e,
                    &mut key_data_salt,
                    &mut key_data_hash_algorithm,
                    &mut key_data_block_size,
                    &mut encrypted_hmac_key,
                    &mut encrypted_hmac_value,
                    &mut spin_count,
                    &mut password_salt,
                    &mut password_hash_algorithm,
                    &mut password_key_bits,
                    &mut encrypted_key_value,
                    &mut encrypted_verifier_hash_input,
                    &mut encrypted_verifier_hash_value,
                )?;
                pop_namespace_frame(&mut ns_stack);
            }
            XmlEvent::End(_) => pop_namespace_frame(&mut ns_stack),
            XmlEvent::Eof => break,
            _ => {}
        }

        if key_data_salt.is_some()
            && key_data_hash_algorithm.is_some()
            && key_data_block_size.is_some()
            && encrypted_hmac_key.is_some()
            && encrypted_hmac_value.is_some()
            && spin_count.is_some()
            && password_salt.is_some()
            && password_hash_algorithm.is_some()
            && password_key_bits.is_some()
            && encrypted_key_value.is_some()
            && encrypted_verifier_hash_input.is_some()
            && encrypted_verifier_hash_value.is_some()
        {
            break;
        }

        buf.clear();
    }

    Ok(AgileEncryptionInfo {
        key_data_salt: key_data_salt.ok_or(OffcryptoError::InvalidEncryptionInfo {
            context: "missing <keyData> element",
        })?,
        key_data_hash_algorithm: key_data_hash_algorithm.ok_or(OffcryptoError::InvalidEncryptionInfo {
            context: "missing <keyData> element",
        })?,
        key_data_block_size: key_data_block_size.ok_or(OffcryptoError::InvalidEncryptionInfo {
            context: "missing <keyData> element",
        })?,
        encrypted_hmac_key: encrypted_hmac_key.ok_or(OffcryptoError::InvalidEncryptionInfo {
            context: "missing <dataIntegrity> element",
        })?,
        encrypted_hmac_value: encrypted_hmac_value.ok_or(OffcryptoError::InvalidEncryptionInfo {
            context: "missing <dataIntegrity> element",
        })?,
        spin_count: spin_count.ok_or(OffcryptoError::InvalidEncryptionInfo {
            context: "missing password <encryptedKey> element",
        })?,
        password_salt: password_salt.ok_or(OffcryptoError::InvalidEncryptionInfo {
            context: "missing password <encryptedKey> element",
        })?,
        password_hash_algorithm: password_hash_algorithm.ok_or(OffcryptoError::InvalidEncryptionInfo {
            context: "missing password <encryptedKey> element",
        })?,
        password_key_bits: password_key_bits.ok_or(OffcryptoError::InvalidEncryptionInfo {
            context: "missing password <encryptedKey> element",
        })?,
        encrypted_key_value: encrypted_key_value.ok_or(OffcryptoError::InvalidEncryptionInfo {
            context: "missing password <encryptedKey> element",
        })?,
        encrypted_verifier_hash_input: encrypted_verifier_hash_input.ok_or(
            OffcryptoError::InvalidEncryptionInfo {
                context: "missing password <encryptedKey> element",
            },
        )?,
        encrypted_verifier_hash_value: encrypted_verifier_hash_value.ok_or(
            OffcryptoError::InvalidEncryptionInfo {
                context: "missing password <encryptedKey> element",
            },
        )?,
    })
}

#[allow(clippy::too_many_arguments)]
fn parse_agile_element<'a>(
    ns_stack: &mut Vec<NamespaceFrame>,
    e: &quick_xml::events::BytesStart<'a>,
    key_data_salt: &mut Option<Vec<u8>>,
    key_data_hash_algorithm: &mut Option<HashAlgorithm>,
    key_data_block_size: &mut Option<usize>,
    encrypted_hmac_key: &mut Option<Vec<u8>>,
    encrypted_hmac_value: &mut Option<Vec<u8>>,
    spin_count: &mut Option<u32>,
    password_salt: &mut Option<Vec<u8>>,
    password_hash_algorithm: &mut Option<HashAlgorithm>,
    password_key_bits: &mut Option<usize>,
    encrypted_key_value: &mut Option<Vec<u8>>,
    encrypted_verifier_hash_input: &mut Option<Vec<u8>>,
    encrypted_verifier_hash_value: &mut Option<Vec<u8>>,
) -> Result<(), OffcryptoError> {
    match e.local_name().as_ref() {
        b"keyData" => {
            let (salt, alg, block_size) = parse_key_data_attrs(e)?;
            *key_data_salt = Some(salt);
            *key_data_hash_algorithm = Some(alg);
            *key_data_block_size = Some(block_size);
        }
        b"dataIntegrity" => {
            let (key, value) = parse_data_integrity_attrs(e)?;
            *encrypted_hmac_key = Some(key);
            *encrypted_hmac_value = Some(value);
        }
        b"encryptedKey" => {
            let name = e.name();
            let prefix = element_prefix(name.as_ref());
            let ns_uri = resolve_namespace_uri(ns_stack, prefix);
            if ns_uri == Some(PASSWORD_KEY_ENCRYPTOR_NS.as_bytes()) {
                let (
                    sc,
                    salt,
                    alg,
                    bits,
                    key_value,
                    vhi,
                    vhv,
                ) = parse_password_encrypted_key_attrs(e)?;
                *spin_count = Some(sc);
                *password_salt = Some(salt);
                *password_hash_algorithm = Some(alg);
                *password_key_bits = Some(bits);
                *encrypted_key_value = Some(key_value);
                *encrypted_verifier_hash_input = Some(vhi);
                *encrypted_verifier_hash_value = Some(vhv);
            }
        }
        _ => {}
    }
    Ok(())
}

fn parse_key_data_attrs<'a>(
    e: &quick_xml::events::BytesStart<'a>,
) -> Result<(Vec<u8>, HashAlgorithm, usize), OffcryptoError> {
    let mut salt_value: Option<Vec<u8>> = None;
    let mut hash_algorithm: Option<HashAlgorithm> = None;
    let mut block_size: Option<usize> = None;

    for attr in e.attributes().with_checks(false) {
        let attr = attr.map_err(|_| OffcryptoError::InvalidEncryptionInfo {
            context: "invalid XML attribute",
        })?;
        let key = local_name(attr.key.as_ref());
        let value = attr.value.as_ref();
        match key {
            b"cipherAlgorithm" => {
                let s = std::str::from_utf8(value).map_err(|_| OffcryptoError::InvalidEncryptionInfo {
                    context: "invalid UTF-8 attribute value",
                })?;
                if !s.trim().eq_ignore_ascii_case("AES") {
                    return Err(OffcryptoError::UnsupportedAlgorithm(
                        "keyData.cipherAlgorithm must be AES".to_string(),
                    ));
                }
            }
            b"cipherChaining" => {
                let s = std::str::from_utf8(value).map_err(|_| OffcryptoError::InvalidEncryptionInfo {
                    context: "invalid UTF-8 attribute value",
                })?;
                if !s.trim().eq_ignore_ascii_case("ChainingModeCBC") {
                    return Err(OffcryptoError::UnsupportedAlgorithm(
                        "keyData.cipherChaining must be ChainingModeCBC".to_string(),
                    ));
                }
            }
            b"saltValue" => {
                salt_value = Some(decode_base64_bounded(
                    value,
                    AGILE_SALT_LEN,
                    "keyData.saltValue too large",
                )?);
            }
            b"hashAlgorithm" => {
                let s = std::str::from_utf8(value).map_err(|_| OffcryptoError::InvalidEncryptionInfo {
                    context: "invalid UTF-8 attribute value",
                })?;
                hash_algorithm = Some(HashAlgorithm::parse_offcrypto_name(s)?);
            }
            b"blockSize" => {
                let parsed = parse_decimal_usize(value, "blockSize")?;
                // AES-CBC requires a 16-byte block size (and IV length).
                if parsed != 16 {
                    return Err(OffcryptoError::UnsupportedAlgorithm(
                        "keyData.blockSize must be 16 for AES".to_string(),
                    ));
                }
                block_size = Some(parsed);
            }
            _ => {}
        }
    }

    let salt_value = salt_value.ok_or(OffcryptoError::InvalidEncryptionInfo {
        context: "missing keyData.saltValue",
    })?;
    validate_agile_salt_len(&salt_value, "keyData.saltValue must be 16 bytes")?;

    let hash_algorithm = hash_algorithm.ok_or(OffcryptoError::InvalidEncryptionInfo {
        context: "missing keyData.hashAlgorithm",
    })?;
    let block_size = block_size.ok_or(OffcryptoError::InvalidEncryptionInfo {
        context: "missing keyData.blockSize",
    })?;

    Ok((salt_value, hash_algorithm, block_size))
}

fn parse_data_integrity_attrs<'a>(
    e: &quick_xml::events::BytesStart<'a>,
) -> Result<(Vec<u8>, Vec<u8>), OffcryptoError> {
    let mut encrypted_hmac_key: Option<Vec<u8>> = None;
    let mut encrypted_hmac_value: Option<Vec<u8>> = None;

    for attr in e.attributes().with_checks(false) {
        let attr = attr.map_err(|_| OffcryptoError::InvalidEncryptionInfo {
            context: "invalid XML attribute",
        })?;
        let key = local_name(attr.key.as_ref());
        let value = attr.value.as_ref();
        match key {
            b"encryptedHmacKey" => {
                encrypted_hmac_key = Some(decode_base64_bounded(
                    value,
                    AGILE_MAX_ENCRYPTED_LEN,
                    "dataIntegrity.encryptedHmacKey too large",
                )?);
            }
            b"encryptedHmacValue" => {
                encrypted_hmac_value = Some(decode_base64_bounded(
                    value,
                    AGILE_MAX_ENCRYPTED_LEN,
                    "dataIntegrity.encryptedHmacValue too large",
                )?);
            }
            _ => {}
        }
    }

    let encrypted_hmac_key = encrypted_hmac_key.ok_or(OffcryptoError::InvalidEncryptionInfo {
        context: "missing dataIntegrity.encryptedHmacKey",
    })?;
    validate_agile_encrypted_len(
        &encrypted_hmac_key,
        "dataIntegrity.encryptedHmacKey must be AES-block aligned",
    )?;

    let encrypted_hmac_value = encrypted_hmac_value.ok_or(OffcryptoError::InvalidEncryptionInfo {
        context: "missing dataIntegrity.encryptedHmacValue",
    })?;
    validate_agile_encrypted_len(
        &encrypted_hmac_value,
        "dataIntegrity.encryptedHmacValue must be AES-block aligned",
    )?;

    Ok((encrypted_hmac_key, encrypted_hmac_value))
}

fn parse_password_encrypted_key_attrs<'a>(
    e: &quick_xml::events::BytesStart<'a>,
) -> Result<
    (
        u32,
        Vec<u8>,
        HashAlgorithm,
        usize,
        Vec<u8>,
        Vec<u8>,
        Vec<u8>,
    ),
    OffcryptoError,
> {
    let mut spin_count: Option<u32> = None;
    let mut salt_value: Option<Vec<u8>> = None;
    let mut hash_algorithm: Option<HashAlgorithm> = None;
    let mut key_bits: Option<usize> = None;

    let mut encrypted_key_value: Option<Vec<u8>> = None;
    let mut encrypted_verifier_hash_input: Option<Vec<u8>> = None;
    let mut encrypted_verifier_hash_value: Option<Vec<u8>> = None;

    for attr in e.attributes().with_checks(false) {
        let attr = attr.map_err(|_| OffcryptoError::InvalidEncryptionInfo {
            context: "invalid XML attribute",
        })?;
        let key = local_name(attr.key.as_ref());
        let value = attr.value.as_ref();
        match key {
            b"cipherAlgorithm" => {
                let s = std::str::from_utf8(value).map_err(|_| OffcryptoError::InvalidEncryptionInfo {
                    context: "invalid UTF-8 attribute value",
                })?;
                if !s.trim().eq_ignore_ascii_case("AES") {
                    return Err(OffcryptoError::UnsupportedAlgorithm(
                        "encryptedKey.cipherAlgorithm must be AES".to_string(),
                    ));
                }
            }
            b"cipherChaining" => {
                let s = std::str::from_utf8(value).map_err(|_| OffcryptoError::InvalidEncryptionInfo {
                    context: "invalid UTF-8 attribute value",
                })?;
                if !s.trim().eq_ignore_ascii_case("ChainingModeCBC") {
                    return Err(OffcryptoError::UnsupportedAlgorithm(
                        "encryptedKey.cipherChaining must be ChainingModeCBC".to_string(),
                    ));
                }
            }
            b"spinCount" => spin_count = Some(parse_decimal_u32(value, "spinCount")?),
            b"saltValue" => {
                salt_value = Some(decode_base64_bounded(
                    value,
                    AGILE_SALT_LEN,
                    "encryptedKey.saltValue too large",
                )?)
            }
            b"hashAlgorithm" => {
                let s = std::str::from_utf8(value).map_err(|_| OffcryptoError::InvalidEncryptionInfo {
                    context: "invalid UTF-8 attribute value",
                })?;
                hash_algorithm = Some(HashAlgorithm::parse_offcrypto_name(s)?);
            }
            b"keyBits" => key_bits = Some(parse_decimal_usize(value, "keyBits")?),
            b"encryptedKeyValue" => {
                encrypted_key_value = Some(decode_base64_bounded(
                    value,
                    AGILE_MAX_ENCRYPTED_LEN,
                    "encryptedKey.encryptedKeyValue too large",
                )?)
            }
            b"encryptedVerifierHashInput" => {
                encrypted_verifier_hash_input = Some(decode_base64_bounded(
                    value,
                    AGILE_MAX_ENCRYPTED_LEN,
                    "encryptedKey.encryptedVerifierHashInput too large",
                )?)
            }
            b"encryptedVerifierHashValue" => {
                encrypted_verifier_hash_value = Some(decode_base64_bounded(
                    value,
                    AGILE_MAX_ENCRYPTED_LEN,
                    "encryptedKey.encryptedVerifierHashValue too large",
                )?)
            }
            _ => {}
        }
    }

    let spin_count = spin_count.ok_or(OffcryptoError::InvalidEncryptionInfo {
        context: "missing encryptedKey.spinCount",
    })?;
    let salt_value = salt_value.ok_or(OffcryptoError::InvalidEncryptionInfo {
        context: "missing encryptedKey.saltValue",
    })?;
    validate_agile_salt_len(&salt_value, "encryptedKey.saltValue must be 16 bytes")?;

    let hash_algorithm = hash_algorithm.ok_or(OffcryptoError::InvalidEncryptionInfo {
        context: "missing encryptedKey.hashAlgorithm",
    })?;
    let key_bits = key_bits.ok_or(OffcryptoError::InvalidEncryptionInfo {
        context: "missing encryptedKey.keyBits",
    })?;

    let encrypted_key_value = encrypted_key_value.ok_or(OffcryptoError::InvalidEncryptionInfo {
        context: "missing encryptedKey.encryptedKeyValue",
    })?;
    validate_agile_encrypted_len(
        &encrypted_key_value,
        "encryptedKey.encryptedKeyValue must be AES-block aligned",
    )?;

    let required_bytes = key_bits
        .checked_add(7)
        .ok_or(OffcryptoError::InvalidFormat {
            context: "encryptedKey.keyBits too large",
        })?
        / 8;
    if required_bytes == 0 || required_bytes > AGILE_MAX_ENCRYPTED_LEN {
        return Err(OffcryptoError::InvalidFormat {
            context: "encryptedKey.keyBits out of range",
        });
    }
    if encrypted_key_value.len() < required_bytes {
        return Err(OffcryptoError::InvalidFormat {
            context: "encryptedKey.encryptedKeyValue too short for keyBits",
        });
    }

    let encrypted_verifier_hash_input =
        encrypted_verifier_hash_input.ok_or(OffcryptoError::InvalidEncryptionInfo {
            context: "missing encryptedKey.encryptedVerifierHashInput",
        })?;
    validate_agile_encrypted_len(
        &encrypted_verifier_hash_input,
        "encryptedKey.encryptedVerifierHashInput must be AES-block aligned",
    )?;

    let encrypted_verifier_hash_value =
        encrypted_verifier_hash_value.ok_or(OffcryptoError::InvalidEncryptionInfo {
            context: "missing encryptedKey.encryptedVerifierHashValue",
        })?;
    validate_agile_encrypted_len(
        &encrypted_verifier_hash_value,
        "encryptedKey.encryptedVerifierHashValue must be AES-block aligned",
    )?;

    Ok((
        spin_count,
        salt_value,
        hash_algorithm,
        key_bits,
        encrypted_key_value,
        encrypted_verifier_hash_input,
        encrypted_verifier_hash_value,
    ))
}

fn validate_agile_salt_len(salt: &[u8], context: &'static str) -> Result<(), OffcryptoError> {
    if salt.len() != AGILE_SALT_LEN {
        return Err(OffcryptoError::InvalidFormat { context });
    }
    Ok(())
}

fn validate_agile_encrypted_len(buf: &[u8], context: &'static str) -> Result<(), OffcryptoError> {
    if buf.is_empty() || buf.len() > AGILE_MAX_ENCRYPTED_LEN || buf.len() % AES_BLOCK_SIZE != 0 {
        return Err(OffcryptoError::InvalidFormat { context });
    }
    Ok(())
}

fn decode_base64_bounded(
    value: &[u8],
    max_len: usize,
    context: &'static str,
) -> Result<Vec<u8>, OffcryptoError> {
    // Some producers pretty-print the `EncryptionInfo` XML and may insert whitespace into long
    // base64 attribute values. Additionally, some omit `=` padding. Be permissive, but avoid
    // decoding unreasonably large values.
    let mut non_ws_len: usize = 0;
    let mut has_ws = false;
    let mut last: Option<u8> = None;
    let mut second_last: Option<u8> = None;
    for &b in value {
        if matches!(b, b'\r' | b'\n' | b'\t' | b' ') {
            has_ws = true;
            continue;
        }
        second_last = last;
        last = Some(b);
        non_ws_len += 1;
    }

    if non_ws_len == 0 {
        return Ok(Vec::new());
    }

    let rem = non_ws_len % 4;
    if rem == 1 {
        return Err(OffcryptoError::InvalidEncryptionInfo {
            context: "invalid base64 value",
        });
    }

    let quads = non_ws_len / 4;
    let mut max_decoded = quads
        .checked_mul(3)
        .ok_or(OffcryptoError::InvalidFormat { context })?;
    match rem {
        0 => {
            let pad = match (second_last, last) {
                (Some(b'='), Some(b'=')) => 2,
                (_, Some(b'=')) => 1,
                _ => 0,
            };
            max_decoded = max_decoded.saturating_sub(pad);
        }
        2 => {
            max_decoded = max_decoded
                .checked_add(1)
                .ok_or(OffcryptoError::InvalidFormat { context })?;
        }
        3 => {
            max_decoded = max_decoded
                .checked_add(2)
                .ok_or(OffcryptoError::InvalidFormat { context })?;
        }
        _ => {}
    }

    if max_decoded > max_len {
        return Err(OffcryptoError::InvalidFormat { context });
    }

    let cleaned = if has_ws {
        let mut out = Vec::with_capacity(non_ws_len);
        for &b in value {
            if !matches!(b, b'\r' | b'\n' | b'\t' | b' ') {
                out.push(b);
            }
        }
        Some(out)
    } else {
        None
    };

    let input = cleaned.as_deref().unwrap_or(value);
    let decoded = STANDARD
        .decode(input)
        .or_else(|_| STANDARD_NO_PAD.decode(input))
        .map_err(|_| OffcryptoError::InvalidEncryptionInfo {
            context: "invalid base64 value",
        })?;

    if decoded.len() > max_len {
        return Err(OffcryptoError::InvalidFormat { context });
    }

    Ok(decoded)
}

#[cfg(test)]
fn decode_b64_attr(value: &str) -> Result<Vec<u8>, OffcryptoError> {
    // Some producers pretty-print the `EncryptionInfo` XML and may insert whitespace into long
    // base64 attribute values. Additionally, some omit `=` padding. Be permissive.
    let bytes = value.as_bytes();

    // Avoid allocating in the common case where there is no whitespace.
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
    STANDARD
        .decode(input)
        .or_else(|_| STANDARD_NO_PAD.decode(input))
        .map_err(|_| OffcryptoError::InvalidEncryptionInfo {
            context: "invalid base64 value",
        })
}

fn parse_decimal_u32(value: &[u8], _name: &'static str) -> Result<u32, OffcryptoError> {
    let s = std::str::from_utf8(value).map_err(|_| OffcryptoError::InvalidEncryptionInfo {
        context: "invalid UTF-8 numeric attribute",
    })?;
    s.trim().parse::<u32>().map_err(|_| OffcryptoError::InvalidEncryptionInfo {
        context: "invalid numeric attribute",
    })
}

fn parse_decimal_usize(value: &[u8], _name: &'static str) -> Result<usize, OffcryptoError> {
    let s = std::str::from_utf8(value).map_err(|_| OffcryptoError::InvalidEncryptionInfo {
        context: "invalid UTF-8 numeric attribute",
    })?;
    s.trim()
        .parse::<usize>()
        .map_err(|_| OffcryptoError::InvalidEncryptionInfo {
            context: "invalid numeric attribute",
        })
}

/// Parse the 8-byte header at the start of an MS-OFFCRYPTO `EncryptedPackage` stream.
pub fn parse_encrypted_package_header(
    bytes: &[u8],
) -> Result<EncryptedPackageHeader, OffcryptoError> {
    let mut r = Reader::new(bytes);
    let original_size = r.read_u64_le("EncryptedPackageHeader.original_size")?;
    Ok(EncryptedPackageHeader { original_size })
}

/// Which encryption schema the `EncryptionInfo` stream uses.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EncryptionType {
    Agile,
    Standard,
    Extensible,
}

/// A best-effort summary of an `EncryptionInfo` stream.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EncryptionInfoSummary {
    pub encryption_type: EncryptionType,
    pub agile: Option<AgileEncryptionInfoSummary>,
    pub standard: Option<StandardEncryptionInfoSummary>,
}

/// Minimal information useful for prompting users about Agile encryption.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AgileEncryptionInfoSummary {
    pub hash_algorithm: HashAlgorithm,
    pub spin_count: u32,
    pub key_bits: u32,
}

/// Minimal information useful for prompting users about Standard encryption.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StandardEncryptionInfoSummary {
    pub alg_id: StandardAlgId,
    pub key_size: u32,
}

/// Subset of CryptoAPI `ALG_ID` values used for Standard Office encryption.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StandardAlgId {
    Aes128,
    Aes192,
    Aes256,
    Unknown(u32),
}

impl StandardAlgId {
    fn from_raw(raw: u32) -> Self {
        match raw {
            // CryptoAPI constants:
            // https://learn.microsoft.com/en-us/windows/win32/seccrypto/alg-id
            0x0000_660E => Self::Aes128,
            0x0000_660F => Self::Aes192,
            0x0000_6610 => Self::Aes256,
            other => Self::Unknown(other),
        }
    }
}

/// Inspect an `EncryptionInfo` stream without requiring a password.
///
/// Supported schemas:
/// - Standard (`*.2` / versionMinor == 2): extracts `EncryptionHeader.algId` and
///   `EncryptionHeader.keySize` (real-world files vary `versionMajor` across 2/3/4)
/// - Agile (`4.4`): extracts `hashAlgorithm`, `spinCount`, and `keyBits` from the password
///   `encryptedKey` element in the XML payload
pub fn inspect_encryption_info(
    encryption_info: &[u8],
) -> Result<EncryptionInfoSummary, OffcryptoError> {
    // `parse_encryption_info` is intentionally strict (it validates algorithms and required field
    // sizes). For user prompting / preflight checks, we want a best-effort summary that can be
    // extracted from *partially-formed* EncryptionInfo buffers.
    //
    // For Standard (versionMinor == 2), only the fixed EncryptionHeader fields are needed
    // (algId/keySize).
    // For Agile (4.4), we reuse the existing XML parser (it already produces actionable errors).
    let mut r = Reader::new(encryption_info);
    let major = r.read_u16_le("EncryptionVersionInfo.major")?;
    let minor = r.read_u16_le("EncryptionVersionInfo.minor")?;
    let _flags = r.read_u32_le("EncryptionVersionInfo.flags")?;
    if (major, minor) == (4, 4) {
        let info = parse_agile_encryption_info_xml(r.remaining())?;
        let key_bits = u32::try_from(info.password_key_bits).map_err(|_| {
            OffcryptoError::InvalidEncryptionInfo {
                context: "encryptedKey.keyBits too large",
            }
        })?;
        return Ok(EncryptionInfoSummary {
            encryption_type: EncryptionType::Agile,
            agile: Some(AgileEncryptionInfoSummary {
                hash_algorithm: info.password_hash_algorithm,
                spin_count: info.spin_count,
                key_bits,
            }),
            standard: None,
        });
    }

    // MS-OFFCRYPTO identifies Standard (CryptoAPI) encryption via `versionMinor == 2`, but
    // real-world files vary `versionMajor` (commonly 3/4; 2 is also seen).
    if minor != 2 || !matches!(major, 2 | 3 | 4) {
        return Err(OffcryptoError::UnsupportedVersion { major, minor });
    }

    let header_size = r.read_u32_le("EncryptionInfo.header_size")? as usize;
    const MIN_STANDARD_HEADER_SIZE: usize = 8 * 4;
    const MAX_STANDARD_HEADER_SIZE: usize = 1024 * 1024;
    if header_size < MIN_STANDARD_HEADER_SIZE || header_size > MAX_STANDARD_HEADER_SIZE {
        return Err(OffcryptoError::InvalidEncryptionInfo {
            context: "EncryptionInfo.header_size is out of bounds",
        });
    }

    let header_bytes = r.take(header_size, "EncryptionHeader")?;
    let mut hr = Reader::new(header_bytes);
    let _flags = hr.read_u32_le("EncryptionHeader.flags")?;
    let _size_extra = hr.read_u32_le("EncryptionHeader.sizeExtra")?;
    let alg_id = hr.read_u32_le("EncryptionHeader.algId")?;
    let _alg_id_hash = hr.read_u32_le("EncryptionHeader.algIdHash")?;
    let key_size = hr.read_u32_le("EncryptionHeader.keySize")?;

    Ok(EncryptionInfoSummary {
        encryption_type: EncryptionType::Standard,
        agile: None,
        standard: Some(StandardEncryptionInfoSummary {
            alg_id: StandardAlgId::from_raw(alg_id),
            key_size,
        }),
    })
}

fn round_up_to_multiple(n: usize, multiple: usize) -> Option<usize> {
    if multiple == 0 {
        return None;
    }
    let rem = n % multiple;
    if rem == 0 {
        return Some(n);
    }
    n.checked_add(multiple - rem)
}

/// Validate an MS-OFFCRYPTO `EncryptedPackage` stream for Standard (CryptoAPI) encryption.
///
/// This is intentionally lightweight and only checks framing invariants:
/// - the 8-byte `original_size` prefix is present
/// - ciphertext length is AES block-aligned (multiple of 16)
pub fn validate_standard_encrypted_package_stream(
    encrypted_package_stream: &[u8],
) -> Result<(), OffcryptoError> {
    parse_encrypted_package_header(encrypted_package_stream)?;
    let ciphertext_len = encrypted_package_stream.len().saturating_sub(8);
    if ciphertext_len % 16 != 0 {
        return Err(OffcryptoError::InvalidCiphertextLength {
            len: ciphertext_len,
        });
    }
    Ok(())
}

/// Validate an Agile-encrypted segment decrypt call (IV/ciphertext lengths).
///
/// Agile encryption processes the package stream in 4096-byte plaintext segments, encrypted
/// independently using AES-CBC with a per-segment IV.
///
/// This helper is intended for robustness tests and defensive callers: it ensures we return an
/// `OffcryptoError` on malformed inputs rather than panicking.
pub fn validate_agile_segment_decrypt_inputs(
    iv: &[u8],
    ciphertext: &[u8],
    expected_plaintext_len: usize,
) -> Result<(), OffcryptoError> {
    if iv.len() != 16 {
        return Err(OffcryptoError::InvalidEncryptionInfo {
            context: "Agile segment IV length must be 16 bytes",
        });
    }
    if ciphertext.len() % 16 != 0 {
        return Err(OffcryptoError::InvalidCiphertextLength {
            len: ciphertext.len(),
        });
    }
    let min_cipher_len = round_up_to_multiple(expected_plaintext_len, 16).ok_or(
        OffcryptoError::InvalidEncryptionInfo {
            context: "Agile segment expected_plaintext_len overflow",
        },
    )?;
    if ciphertext.len() < min_cipher_len {
        return Err(OffcryptoError::InvalidEncryptionInfo {
            context: "Agile segment ciphertext is too short for expected_plaintext_len",
        });
    }
    Ok(())
}

fn password_to_utf16le_bytes(password: &str) -> Vec<u8> {
    let mut out = Vec::with_capacity(password.len().saturating_mul(2));
    for unit in password.encode_utf16() {
        out.extend_from_slice(&unit.to_le_bytes());
    }
    out
}

fn sha1(data: &[u8]) -> [u8; SHA1_LEN] {
    Sha1::digest(data).into()
}

fn aes_ecb_decrypt_in_place(key: &[u8], buf: &mut [u8]) -> Result<(), OffcryptoError> {
    if buf.len() % 16 != 0 {
        return Err(OffcryptoError::InvalidCiphertextLength { len: buf.len() });
    }

    fn decrypt_with<C>(key: &[u8], buf: &mut [u8]) -> Result<(), OffcryptoError>
    where
        C: BlockDecrypt + KeyInit,
    {
        let cipher =
            C::new_from_slice(key).map_err(|_| OffcryptoError::InvalidKeyLength { len: key.len() })?;
        for block in buf.chunks_mut(16) {
            cipher.decrypt_block(GenericArray::from_mut_slice(block));
        }
        Ok(())
    }

    match key.len() {
        16 => decrypt_with::<Aes128>(key, buf),
        24 => decrypt_with::<Aes192>(key, buf),
        32 => decrypt_with::<Aes256>(key, buf),
        _ => Err(OffcryptoError::InvalidKeyLength { len: key.len() }),
    }
}

/// Standard/ECMA-376 key derivation used by CryptoAPI-based encryption (SHA1 + AES).
///
/// This matches the ECMA-376 "Standard" (CryptoAPI) algorithm used by `msoffcrypto-tool`.
///
/// - `salt` comes from `EncryptionVerifier.salt`
/// - `key_size_bits` comes from `EncryptionHeader.keySize` and must be 128/192/256
pub fn make_key_from_password(
    password: &str,
    salt: &[u8],
    key_size_bits: u32,
) -> Result<Vec<u8>, OffcryptoError> {
    let key_len = match key_size_bits {
        128 | 192 | 256 => (key_size_bits / 8) as usize,
        other => {
            return Err(OffcryptoError::UnsupportedAlgorithm(format!(
                "keySize={other} bits"
            )))
        }
    };

    let password_utf16 = password_to_utf16le_bytes(password);

    // h = sha1(salt || password_utf16le)
    let mut hasher = Sha1::new();
    hasher.update(salt);
    hasher.update(&password_utf16);
    let mut h: [u8; SHA1_LEN] = hasher.finalize().into();

    // for i in 0..ITER_COUNT: h = sha1(u32le(i) || h)
    let mut buf = [0u8; 4 + SHA1_LEN];
    for i in 0..ITER_COUNT {
        buf[..4].copy_from_slice(&(i as u32).to_le_bytes());
        buf[4..].copy_from_slice(&h);
        h = sha1(&buf);
    }

    // h_final = sha1(h || u32le(0))
    let mut buf0 = [0u8; SHA1_LEN + 4];
    buf0[..SHA1_LEN].copy_from_slice(&h);
    buf0[SHA1_LEN..].copy_from_slice(&0u32.to_le_bytes());
    let h_final = sha1(&buf0);

    // key = (sha1((0x36*64) ^ h_final) || sha1((0x5c*64) ^ h_final))[..key_len]
    let mut buf1 = [0x36u8; 64];
    let mut buf2 = [0x5cu8; 64];
    for i in 0..SHA1_LEN {
        buf1[i] ^= h_final[i];
        buf2[i] ^= h_final[i];
    }
    let x1 = sha1(&buf1);
    let x2 = sha1(&buf2);

    let mut out = [0u8; SHA1_LEN * 2];
    out[..SHA1_LEN].copy_from_slice(&x1);
    out[SHA1_LEN..].copy_from_slice(&x2);

    debug_assert!(key_len <= out.len());
    Ok(out[..key_len].to_vec())
}

/// Verify a Standard encryption password/key using the encrypted verifier fields.
pub fn verify_password(
    key: &[u8],
    encrypted_verifier: &[u8],
    encrypted_verifier_hash: &[u8],
) -> Result<(), OffcryptoError> {
    if encrypted_verifier.len() != 16 {
        return Err(OffcryptoError::InvalidStructure(format!(
            "`encryptedVerifier` must be 16 bytes, got {}",
            encrypted_verifier.len()
        )));
    }
    if encrypted_verifier_hash.len() % 16 != 0 {
        return Err(OffcryptoError::InvalidStructure(format!(
            "`encryptedVerifierHash` length must be a multiple of 16 bytes, got {}",
            encrypted_verifier_hash.len()
        )));
    }

    let mut verifier = [0u8; 16];
    verifier.copy_from_slice(encrypted_verifier);
    aes_ecb_decrypt_in_place(key, &mut verifier)?;
    let expected_hash: [u8; SHA1_LEN] = sha1(&verifier);

    let mut verifier_hash = encrypted_verifier_hash.to_vec();
    aes_ecb_decrypt_in_place(key, &mut verifier_hash)?;
    if verifier_hash.len() < SHA1_LEN {
        return Err(OffcryptoError::InvalidStructure(format!(
            "decrypted verifier hash must be at least 20 bytes, got {}",
            verifier_hash.len()
        )));
    }

    if util::ct_eq(&expected_hash[..], &verifier_hash[..SHA1_LEN]) {
        Ok(())
    } else {
        Err(OffcryptoError::InvalidPassword)
    }
}

/// Decrypt the `EncryptedPackage` stream for Standard (CryptoAPI / AES) encryption using AES-ECB.
///
/// The stream format is:
/// - `u64` original size (LE)
/// - ciphertext bytes (AES-ECB, block-aligned, no padding)
pub fn decrypt_encrypted_package_ecb(
    key: &[u8],
    encrypted_package: &[u8],
) -> Result<Vec<u8>, OffcryptoError> {
    if encrypted_package.len() < 8 {
        return Err(OffcryptoError::InvalidStructure(format!(
            "`EncryptedPackage` must be at least 8 bytes, got {}",
            encrypted_package.len()
        )));
    }

    let original_size = u64::from_le_bytes(
        encrypted_package[..8]
            .try_into()
            .expect("slice length checked above"),
    );
    let original_size: usize = original_size.try_into().map_err(|_| {
        OffcryptoError::InvalidStructure("original size does not fit into usize".to_string())
    })?;

    let ciphertext = &encrypted_package[8..];
    if ciphertext.len() % 16 != 0 {
        return Err(OffcryptoError::InvalidStructure(format!(
            "ciphertext length must be a multiple of 16 bytes, got {}",
            ciphertext.len()
        )));
    }

    let mut plaintext = ciphertext.to_vec();
    aes_ecb_decrypt_in_place(key, &mut plaintext)?;

    if original_size > plaintext.len() {
        return Err(OffcryptoError::InvalidStructure(format!(
            "original size {original_size} exceeds plaintext length {}",
            plaintext.len()
        )));
    }

    plaintext.truncate(original_size);
    Ok(plaintext)
}

/// Decrypt a password-protected ECMA-376 (OOXML) file which uses MS-OFFCRYPTO "Standard"
/// (CryptoAPI / AES) encryption.
///
/// `data` must be an OLE Compound File containing the `EncryptionInfo` and `EncryptedPackage`
/// streams.
pub fn decrypt_from_bytes(data: &[u8], password: &str) -> Result<Vec<u8>, OffcryptoError> {
    let cursor = Cursor::new(data);
    let mut ole = cfb::CompoundFile::open(cursor)
        .map_err(|e| OffcryptoError::InvalidStructure(format!("failed to open OLE compound file: {e}")))?;

    let mut encryption_info = Vec::new();
    {
        let mut stream = match ole.open_stream("EncryptionInfo") {
            Ok(stream) => stream,
            Err(err) if err.kind() == std::io::ErrorKind::NotFound => {
                // Some CFB implementations treat stream paths as absolute and require a leading `/`.
                // Be permissive and try both.
                match ole.open_stream("/EncryptionInfo") {
                    Ok(stream) => stream,
                    Err(err) if err.kind() == std::io::ErrorKind::NotFound => {
                        return Err(OffcryptoError::InvalidStructure(
                            "missing `EncryptionInfo` stream".to_string(),
                        ));
                    }
                    Err(err) => {
                        return Err(OffcryptoError::InvalidStructure(format!(
                            "failed to open `EncryptionInfo`: {err}"
                        )));
                    }
                }
            }
            Err(err) => {
                return Err(OffcryptoError::InvalidStructure(format!(
                    "failed to open `EncryptionInfo`: {err}"
                )));
            }
        };
        stream
            .read_to_end(&mut encryption_info)
            .map_err(|e| OffcryptoError::InvalidStructure(format!("failed to read `EncryptionInfo`: {e}")))?;
    }

    let (header, verifier) = match parse_encryption_info(&encryption_info)? {
        EncryptionInfo::Standard { header, verifier, .. } => (header, verifier),
        EncryptionInfo::Agile { .. } => {
            return Err(OffcryptoError::UnsupportedEncryption {
                encryption_type: EncryptionType::Agile,
            })
        }
        EncryptionInfo::Unsupported { version } => {
            if version.minor == 3 && matches!(version.major, 3 | 4) {
                // MS-OFFCRYPTO "Extensible" encryption: known scheme, but not supported by this
                // Standard-only decryptor.
                return Err(OffcryptoError::UnsupportedEncryption {
                    encryption_type: EncryptionType::Extensible,
                });
            }
            return Err(OffcryptoError::UnsupportedVersion {
                major: version.major,
                minor: version.minor,
            });
        }
    };

    let key = make_key_from_password(password, &verifier.salt, header.key_size_bits)?;
    verify_password(&key, &verifier.encrypted_verifier, &verifier.encrypted_verifier_hash)?;

    let mut encrypted_package = Vec::new();
    {
        let mut stream = match ole.open_stream("EncryptedPackage") {
            Ok(stream) => stream,
            Err(err) if err.kind() == std::io::ErrorKind::NotFound => {
                match ole.open_stream("/EncryptedPackage") {
                    Ok(stream) => stream,
                    Err(err) if err.kind() == std::io::ErrorKind::NotFound => {
                        return Err(OffcryptoError::InvalidStructure(
                            "missing `EncryptedPackage` stream".to_string(),
                        ));
                    }
                    Err(err) => {
                        return Err(OffcryptoError::InvalidStructure(format!(
                            "failed to open `EncryptedPackage`: {err}"
                        )));
                    }
                }
            }
            Err(err) => {
                return Err(OffcryptoError::InvalidStructure(format!(
                    "failed to open `EncryptedPackage`: {err}"
                )));
            }
        };
        stream
            .read_to_end(&mut encrypted_package)
            .map_err(|e| OffcryptoError::InvalidStructure(format!("failed to read `EncryptedPackage`: {e}")))?;
    }

    decrypt_encrypted_package_ecb(&key, &encrypted_package)
}

/// ECMA-376 Standard Encryption password→key derivation.
///
/// Reference algorithm: `msoffcrypto` `ECMA376Standard.makekey_from_password`.
pub fn standard_derive_key(
    info: &StandardEncryptionInfo,
    password: &str,
) -> Result<Vec<u8>, OffcryptoError> {
    validate_standard_encryption_info(info)?;

    let key_len = match info.header.key_size_bits.checked_div(8) {
        Some(v) if info.header.key_size_bits % 8 == 0 => v as usize,
        _ => {
            return Err(OffcryptoError::InvalidKeySizeBits {
                key_size_bits: info.header.key_size_bits,
            })
        }
    };

    // Password-derived material should not linger in heap buffers longer than needed.
    let password_utf16 = Zeroizing::new(password_to_utf16le_bytes(password));

    // h = sha1(salt || password_utf16)
    let mut hasher = Sha1::new();
    hasher.update(&info.verifier.salt);
    hasher.update(password_utf16.as_slice());
    let mut h: Zeroizing<[u8; SHA1_LEN]> = Zeroizing::new(hasher.finalize().into());

    // for i in 0..ITER_COUNT-1: h = sha1(u32le(i) || h)
    let mut buf: Zeroizing<[u8; 4 + SHA1_LEN]> = Zeroizing::new([0u8; 4 + SHA1_LEN]);
    for i in 0..ITER_COUNT {
        buf[..4].copy_from_slice(&(i as u32).to_le_bytes());
        buf[4..].copy_from_slice(&h[..]);
        *h = sha1(&buf[..]);
    }

    // hfinal = sha1(h || u32le(0))
    let mut buf0: Zeroizing<[u8; SHA1_LEN + 4]> = Zeroizing::new([0u8; SHA1_LEN + 4]);
    buf0[..SHA1_LEN].copy_from_slice(&h[..]);
    buf0[SHA1_LEN..].copy_from_slice(&0u32.to_le_bytes());
    let hfinal: Zeroizing<[u8; SHA1_LEN]> = Zeroizing::new(sha1(&buf0[..]));

    // key = (sha1((0x36*64) ^ hfinal) || sha1((0x5c*64) ^ hfinal))[..key_len]
    let mut buf1: Zeroizing<[u8; 64]> = Zeroizing::new([0x36u8; 64]);
    let mut buf2: Zeroizing<[u8; 64]> = Zeroizing::new([0x5cu8; 64]);
    for i in 0..SHA1_LEN {
        buf1[i] ^= hfinal[i];
        buf2[i] ^= hfinal[i];
    }
    let x1: Zeroizing<[u8; SHA1_LEN]> = Zeroizing::new(sha1(&buf1[..]));
    let x2: Zeroizing<[u8; SHA1_LEN]> = Zeroizing::new(sha1(&buf2[..]));

    let mut out: Zeroizing<[u8; SHA1_LEN * 2]> = Zeroizing::new([0u8; SHA1_LEN * 2]);
    out[..SHA1_LEN].copy_from_slice(&x1[..]);
    out[SHA1_LEN..].copy_from_slice(&x2[..]);

    if key_len > out.len() {
        return Err(OffcryptoError::DerivedKeyTooLong {
            key_size_bits: info.header.key_size_bits,
            required_bytes: key_len,
            available_bytes: out.len(),
        });
    }

    Ok(out[..key_len].to_vec())
}

/// [`standard_derive_key`] variant that returns a [`Zeroizing<Vec<u8>>`], ensuring the derived
/// key bytes are wiped from memory when dropped.
pub fn standard_derive_key_zeroizing(
    info: &StandardEncryptionInfo,
    password: &str,
) -> Result<Zeroizing<Vec<u8>>, OffcryptoError> {
    Ok(Zeroizing::new(standard_derive_key(info, password)?))
}

/// ECMA-376 Standard Encryption key verifier check.
///
/// Reference algorithm: `msoffcrypto` `ECMA376Standard.verifykey`.
pub fn standard_verify_key(info: &StandardEncryptionInfo, key: &[u8]) -> Result<(), OffcryptoError> {
    validate_standard_encryption_info(info)?;

    let mut verifier: Zeroizing<[u8; 16]> = Zeroizing::new(info.verifier.encrypted_verifier);
    aes_ecb_decrypt_in_place(key, &mut verifier[..])?;
    let expected_hash: Zeroizing<[u8; SHA1_LEN]> = Zeroizing::new(sha1(&verifier[..]));

    let mut verifier_hash: Zeroizing<Vec<u8>> = Zeroizing::new(info.verifier.encrypted_verifier_hash.clone());
    aes_ecb_decrypt_in_place(key, &mut verifier_hash[..])?;
    if verifier_hash.len() < SHA1_LEN {
        return Err(OffcryptoError::InvalidVerifierHashLength {
            len: verifier_hash.len(),
        });
    }

    if util::ct_eq(&expected_hash[..], &verifier_hash[..SHA1_LEN]) {
        Ok(())
    } else {
        Err(OffcryptoError::InvalidPassword)
    }
}

const BLK_KEY_VERIFIER_HASH_INPUT: [u8; 8] = [0xFE, 0xA7, 0xD2, 0x76, 0x3B, 0x4B, 0x9E, 0x79];
const BLK_KEY_VERIFIER_HASH_VALUE: [u8; 8] = [0xD7, 0xAA, 0x0F, 0x6D, 0x30, 0x61, 0x34, 0x4E];
const BLK_KEY_ENCRYPTED_KEY_VALUE: [u8; 8] = [0x14, 0x6E, 0x0B, 0xE7, 0xAB, 0xAC, 0xD0, 0xD6];

fn normalize_key_material(bytes: &[u8], out_len: usize) -> Vec<u8> {
    if bytes.len() >= out_len {
        return bytes[..out_len].to_vec();
    }

    // MS-OFFCRYPTO `TruncateHash` expansion: append 0x36 bytes (matches `msoffcrypto-tool`).
    let mut out = vec![0x36u8; out_len];
    out[..bytes.len()].copy_from_slice(bytes);
    out
}

fn derive_iterated_hash_from_password(
    password: &str,
    salt_value: &[u8],
    hash_algorithm: HashAlgorithm,
    spin_count: u32,
    limits: &DecryptLimits,
    mut on_iteration: Option<&mut dyn FnMut(u32)>,
) -> Result<Zeroizing<Vec<u8>>, OffcryptoError> {
    // `spinCount` is attacker-controlled; enforce limits up front to avoid CPU DoS.
    check_spin_count(spin_count, limits)?;

    let password_utf16 = Zeroizing::new(password_to_utf16le_bytes(password));
    let digest_len = hash_algorithm.digest_len();
    debug_assert!(digest_len <= MAX_DIGEST_LEN);

    // Avoid per-iteration allocations (spinCount is often 100k):
    // keep the current digest in a fixed buffer and overwrite it each round.
    let mut h_buf: Zeroizing<[u8; MAX_DIGEST_LEN]> = Zeroizing::new([0u8; MAX_DIGEST_LEN]);
    hash_algorithm.digest_two_into(
        salt_value,
        password_utf16.as_slice(),
        &mut h_buf[..digest_len],
    );

    match hash_algorithm {
        HashAlgorithm::Sha1 => {
            for i in 0..spin_count {
                if let Some(cb) = on_iteration.as_mut() {
                    cb(i);
                }
                let mut hasher = Sha1::new();
                hasher.update(i.to_le_bytes());
                hasher.update(&h_buf[..20]);
                h_buf[..20].copy_from_slice(&hasher.finalize());
            }
        }
        HashAlgorithm::Sha256 => {
            for i in 0..spin_count {
                if let Some(cb) = on_iteration.as_mut() {
                    cb(i);
                }
                let mut hasher = sha2::Sha256::new();
                hasher.update(i.to_le_bytes());
                hasher.update(&h_buf[..32]);
                h_buf[..32].copy_from_slice(&hasher.finalize());
            }
        }
        HashAlgorithm::Sha384 => {
            for i in 0..spin_count {
                if let Some(cb) = on_iteration.as_mut() {
                    cb(i);
                }
                let mut hasher = sha2::Sha384::new();
                hasher.update(i.to_le_bytes());
                hasher.update(&h_buf[..48]);
                h_buf[..48].copy_from_slice(&hasher.finalize());
            }
        }
        HashAlgorithm::Sha512 => {
            for i in 0..spin_count {
                if let Some(cb) = on_iteration.as_mut() {
                    cb(i);
                }
                let mut hasher = sha2::Sha512::new();
                hasher.update(i.to_le_bytes());
                hasher.update(&h_buf[..64]);
                h_buf[..64].copy_from_slice(&hasher.finalize());
            }
        }
    }

    Ok(Zeroizing::new(h_buf[..digest_len].to_vec()))
}

fn derive_encryption_key(
    h: &[u8],
    block_key: &[u8],
    hash_algorithm: HashAlgorithm,
    key_bits: usize,
) -> Result<Zeroizing<Vec<u8>>, OffcryptoError> {
    if key_bits % 8 != 0 {
        return Err(OffcryptoError::InvalidEncryptionInfo {
            context: "keyBits is not divisible by 8",
        });
    }
    let key_len = key_bits / 8;

    // Avoid allocating a temporary `H || blockKey` buffer: hash with two updates into a fixed
    // stack buffer.
    let digest_len = hash_algorithm.digest_len();
    debug_assert!(digest_len <= MAX_DIGEST_LEN);
    let mut digest: Zeroizing<[u8; MAX_DIGEST_LEN]> = Zeroizing::new([0u8; MAX_DIGEST_LEN]);
    hash_algorithm.digest_two_into(h, block_key, &mut digest[..digest_len]);

    Ok(Zeroizing::new(normalize_key_material(
        &digest[..digest_len],
        key_len,
    )))
}

fn derive_iv_from_salt(
    salt: &[u8],
    block_key: &[u8],
    hash_algorithm: HashAlgorithm,
) -> Result<[u8; 16], OffcryptoError> {
    // Avoid allocating a temporary `salt || blockKey` buffer: hash with two updates into a fixed
    // stack buffer.
    let digest_len = hash_algorithm.digest_len();
    debug_assert!(digest_len <= MAX_DIGEST_LEN);
    if digest_len < 16 {
        return Err(OffcryptoError::InvalidEncryptionInfo {
            context: "hash output shorter than AES block size",
        });
    }
    let mut digest: Zeroizing<[u8; MAX_DIGEST_LEN]> = Zeroizing::new([0u8; MAX_DIGEST_LEN]);
    hash_algorithm.digest_two_into(salt, block_key, &mut digest[..digest_len]);

    let mut iv = [0u8; 16];
    iv.copy_from_slice(&digest[..16]);
    Ok(iv)
}

fn salt_iv(salt: &[u8]) -> Result<[u8; 16], OffcryptoError> {
    let mut iv = [0u8; 16];
    iv.copy_from_slice(
        salt.get(..16).ok_or(OffcryptoError::InvalidEncryptionInfo {
            context: "password salt is shorter than AES block size",
        })?,
    );
    Ok(iv)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum AgilePasswordIvMode {
    /// Use `iv = passwordSalt` (most common / matches `msoffcrypto-tool` docstring vectors).
    Salt,
    /// Use `iv = HASH(passwordSalt || blockKey)[:16]` (observed in some fixtures/tooling).
    Derived,
}

fn aes_cbc_decrypt(
    ciphertext: &[u8],
    key: &[u8],
    iv: &[u8],
) -> Result<Zeroizing<Vec<u8>>, OffcryptoError> {
    if iv.len() != 16 {
        return Err(OffcryptoError::InvalidEncryptionInfo {
            context: "AES-CBC IV must be 16 bytes",
        });
    }
    if ciphertext.len() % 16 != 0 {
        return Err(OffcryptoError::InvalidEncryptionInfo {
            context: "AES-CBC ciphertext length must be a multiple of 16 bytes",
        });
    }

    let mut buf = Zeroizing::new(ciphertext.to_vec());

    match key.len() {
        16 => {
            let decryptor =
                Decryptor::<Aes128>::new_from_slices(key, iv).map_err(|_| OffcryptoError::InvalidKeyLength {
                    len: key.len(),
                })?;
            decryptor
                .decrypt_padded_mut::<NoPadding>(&mut buf[..])
                .map_err(|_| OffcryptoError::InvalidEncryptionInfo {
                    context: "AES-CBC decrypt failed",
                })?;
        }
        24 => {
            let decryptor =
                Decryptor::<Aes192>::new_from_slices(key, iv).map_err(|_| OffcryptoError::InvalidKeyLength {
                    len: key.len(),
                })?;
            decryptor
                .decrypt_padded_mut::<NoPadding>(&mut buf[..])
                .map_err(|_| OffcryptoError::InvalidEncryptionInfo {
                    context: "AES-CBC decrypt failed",
                })?;
        }
        32 => {
            let decryptor =
                Decryptor::<Aes256>::new_from_slices(key, iv).map_err(|_| OffcryptoError::InvalidKeyLength {
                    len: key.len(),
                })?;
            decryptor
                .decrypt_padded_mut::<NoPadding>(&mut buf[..])
                .map_err(|_| OffcryptoError::InvalidEncryptionInfo {
                    context: "AES-CBC decrypt failed",
                })?;
        }
        _ => return Err(OffcryptoError::InvalidKeyLength { len: key.len() }),
    };

    Ok(buf)
}

/// Verify an OOXML Agile (ECMA-376) password using the same algorithm as `msoffcrypto-tool`.
///
/// This matches `ECMA376Agile.verify_password` in `msoffcrypto-tool`.
pub fn agile_verify_password(
    info: &AgileEncryptionInfo,
    password: &str,
) -> Result<(), OffcryptoError> {
    let options = DecryptOptions::default();
    agile_verify_password_with_options(info, password, &options)
}

/// Like [`agile_verify_password`], but allows overriding resource limits.
pub fn agile_verify_password_with_options(
    info: &AgileEncryptionInfo,
    password: &str,
    options: &DecryptOptions,
) -> Result<(), OffcryptoError> {
    let hfinal = derive_iterated_hash_from_password(
        password,
        &info.password_salt,
        info.password_hash_algorithm,
        info.spin_count,
        &options.limits,
        None,
    )?;

    let key1 = derive_encryption_key(
        &hfinal[..],
        &BLK_KEY_VERIFIER_HASH_INPUT,
        info.password_hash_algorithm,
        info.password_key_bits,
    )?;
    let key2 = derive_encryption_key(
        &hfinal[..],
        &BLK_KEY_VERIFIER_HASH_VALUE,
        info.password_hash_algorithm,
        info.password_key_bits,
    )?;

    let hash_len = match info.password_hash_algorithm {
        HashAlgorithm::Sha1 => 20,
        HashAlgorithm::Sha256 => 32,
        HashAlgorithm::Sha384 => 48,
        HashAlgorithm::Sha512 => 64,
    };

    let try_mode = |mode: AgilePasswordIvMode| -> Result<(), OffcryptoError> {
        let iv1 = match mode {
            AgilePasswordIvMode::Salt => salt_iv(&info.password_salt)?,
            AgilePasswordIvMode::Derived => derive_iv_from_salt(
                &info.password_salt,
                &BLK_KEY_VERIFIER_HASH_INPUT,
                info.password_hash_algorithm,
            )?,
        };
        let iv2 = match mode {
            AgilePasswordIvMode::Salt => salt_iv(&info.password_salt)?,
            AgilePasswordIvMode::Derived => derive_iv_from_salt(
                &info.password_salt,
                &BLK_KEY_VERIFIER_HASH_VALUE,
                info.password_hash_algorithm,
            )?,
        };

        let verifier_hash_input = aes_cbc_decrypt(&info.encrypted_verifier_hash_input, &key1, &iv1)?;
        let verifier_hash_value_full =
            aes_cbc_decrypt(&info.encrypted_verifier_hash_value, &key2, &iv2)?;

        let verifier_hash_value = verifier_hash_value_full.get(..hash_len).ok_or(
            OffcryptoError::InvalidEncryptionInfo {
                context: "decrypted verifierHashValue shorter than hash output",
            },
        )?;

        agile::verify_password(&verifier_hash_input, verifier_hash_value, info.password_hash_algorithm)
    };

    match try_mode(AgilePasswordIvMode::Salt) {
        Ok(()) => Ok(()),
        Err(OffcryptoError::InvalidPassword) => try_mode(AgilePasswordIvMode::Derived),
        Err(other) => Err(other),
    }
}

/// Extract the Agile "secret key" by decrypting `encryptedKeyValue`.
///
/// This matches the algorithm used by `msoffcrypto`'s
/// `ECMA376Agile.makekey_from_password` implementation:
///
/// 1) Compute an iterated hash from `password`, `passwordSalt`, `spinCount`, and `passwordHashAlgorithm`.
/// 2) Derive `encryption_key = HASH(h || block3).digest()[..keyBits/8]` where
///    `block3 = 14 6E 0B E7 AB AC D0 D6`.
/// 3) Decrypt `encryptedKeyValue` using AES-CBC/NoPadding.
///
/// Most files use `iv = passwordSalt` (as in `msoffcrypto-tool`), but some toolchains derive
/// `iv = HASH(passwordSalt || block3)[:16]`. When verifier fields are present in `info`, this
/// function attempts both schemes to ensure interoperability.
pub fn agile_secret_key(
    info: &AgileEncryptionInfo,
    password: &str,
) -> Result<Zeroizing<Vec<u8>>, OffcryptoError> {
    let options = DecryptOptions::default();
    agile_secret_key_with_options(info, password, &options)
}

/// Like [`agile_secret_key`], but allows overriding resource limits.
pub fn agile_secret_key_with_options(
    info: &AgileEncryptionInfo,
    password: &str,
    options: &DecryptOptions,
) -> Result<Zeroizing<Vec<u8>>, OffcryptoError> {
    let hfinal = derive_iterated_hash_from_password(
        password,
        &info.password_salt,
        info.password_hash_algorithm,
        info.spin_count,
        &options.limits,
        None,
    )?;

    // Select the IV scheme. If verifier fields are missing, fall back to the common `iv = salt`
    // behavior (used by `msoffcrypto-tool`'s `makekey_from_password` vector).
    let iv_mode = if info.encrypted_verifier_hash_input.is_empty() || info.encrypted_verifier_hash_value.is_empty()
    {
        AgilePasswordIvMode::Salt
    } else {
        // Reuse the same verifier logic as `agile_verify_password`, but return which mode succeeded.
        let key1 = derive_encryption_key(
            &hfinal,
            &BLK_KEY_VERIFIER_HASH_INPUT,
            info.password_hash_algorithm,
            info.password_key_bits,
        )?;
        let key2 = derive_encryption_key(
            &hfinal,
            &BLK_KEY_VERIFIER_HASH_VALUE,
            info.password_hash_algorithm,
            info.password_key_bits,
        )?;

        let hash_len = match info.password_hash_algorithm {
            HashAlgorithm::Sha1 => 20,
            HashAlgorithm::Sha256 => 32,
            HashAlgorithm::Sha384 => 48,
            HashAlgorithm::Sha512 => 64,
        };

        let verify_with_mode = |mode: AgilePasswordIvMode| -> Result<(), OffcryptoError> {
            let iv1 = match mode {
                AgilePasswordIvMode::Salt => salt_iv(&info.password_salt)?,
                AgilePasswordIvMode::Derived => derive_iv_from_salt(
                    &info.password_salt,
                    &BLK_KEY_VERIFIER_HASH_INPUT,
                    info.password_hash_algorithm,
                )?,
            };
            let iv2 = match mode {
                AgilePasswordIvMode::Salt => salt_iv(&info.password_salt)?,
                AgilePasswordIvMode::Derived => derive_iv_from_salt(
                    &info.password_salt,
                    &BLK_KEY_VERIFIER_HASH_VALUE,
                    info.password_hash_algorithm,
                )?,
            };

            let verifier_hash_input =
                aes_cbc_decrypt(&info.encrypted_verifier_hash_input, &key1, &iv1)?;
            let verifier_hash_value_full =
                aes_cbc_decrypt(&info.encrypted_verifier_hash_value, &key2, &iv2)?;

            let verifier_hash_value = verifier_hash_value_full.get(..hash_len).ok_or(
                OffcryptoError::InvalidEncryptionInfo {
                    context: "decrypted verifierHashValue shorter than hash output",
                },
            )?;

            agile::verify_password(
                &verifier_hash_input,
                verifier_hash_value,
                info.password_hash_algorithm,
            )
        };

        match verify_with_mode(AgilePasswordIvMode::Salt) {
            Ok(()) => AgilePasswordIvMode::Salt,
            Err(OffcryptoError::InvalidPassword) => match verify_with_mode(AgilePasswordIvMode::Derived) {
                Ok(()) => AgilePasswordIvMode::Derived,
                Err(err) => return Err(err),
            },
            Err(err) => return Err(err),
        }
    };

    let encryption_key = derive_encryption_key(
        &hfinal[..],
        &BLK_KEY_ENCRYPTED_KEY_VALUE,
        info.password_hash_algorithm,
        info.password_key_bits,
    )?;

    let iv = match iv_mode {
        AgilePasswordIvMode::Salt => salt_iv(&info.password_salt)?,
        AgilePasswordIvMode::Derived => derive_iv_from_salt(
            &info.password_salt,
            &BLK_KEY_ENCRYPTED_KEY_VALUE,
            info.password_hash_algorithm,
        )?,
    };
    let mut secret_key = aes_cbc_decrypt(&info.encrypted_key_value, &encryption_key[..], &iv)?;

    // The decrypted blob may include trailing zero padding; only the first `keyBits/8` bytes are
    // the actual package key.
    if info.password_key_bits % 8 != 0 {
        return Err(OffcryptoError::InvalidEncryptionInfo {
            context: "keyBits is not divisible by 8",
        });
    }
    let key_len = info.password_key_bits / 8;
    if key_len == 0 || key_len > secret_key.len() {
        return Err(OffcryptoError::InvalidEncryptionInfo {
            context: "decrypted encryptedKeyValue shorter than keyBits/8",
        });
    }
    secret_key.truncate(key_len);
    Ok(secret_key)
}

/// Decrypt a Standard-encrypted OOXML package (e.g. `.docx`, `.xlsx`) from a raw OLE/CFB wrapper.
///
/// This performs native MS-OFFCRYPTO Standard (CryptoAPI / AES) password verification and
/// decryption and returns the decrypted OOXML ZIP bytes.
pub fn decrypt_standard_ooxml_from_bytes(
    raw_ole: Vec<u8>,
    password: &str,
) -> Result<Vec<u8>, OffcryptoError> {
    decrypt_from_bytes(&raw_ole, password)
}

/// Decrypt an Agile-encrypted OOXML package (e.g. `.xlsx`, `.docx`) from a raw OLE/CFB wrapper.
///
/// `raw_ole` must be an OLE Compound File containing the `EncryptionInfo` and `EncryptedPackage`
/// streams.
pub fn decrypt_agile_ooxml_from_bytes(
    raw_ole: Vec<u8>,
    password: &str,
) -> Result<Vec<u8>, OffcryptoError> {
    // 1) Parse and validate `EncryptionInfo` (must be Agile 4.4).
    let encryption_info = read_ole_stream(&raw_ole, "EncryptionInfo")?;
    let info = match parse_encryption_info(&encryption_info)? {
        EncryptionInfo::Agile { info, .. } => info,
        EncryptionInfo::Standard { .. } => {
            return Err(OffcryptoError::UnsupportedEncryption {
                encryption_type: EncryptionType::Standard,
            })
        }
        EncryptionInfo::Unsupported { version } => {
            if version.minor == 3 && matches!(version.major, 3 | 4) {
                // MS-OFFCRYPTO "Extensible" encryption: known scheme, but not supported by this
                // Agile-only decryptor.
                return Err(OffcryptoError::UnsupportedEncryption {
                    encryption_type: EncryptionType::Extensible,
                });
            }
            return Err(OffcryptoError::UnsupportedVersion {
                major: version.major,
                minor: version.minor,
            });
        }
    };

    // 2) Derive secret key (also validates the password via verifier hashes).
    let secret_key = agile_secret_key(&info, password)?;

    // 3) Decrypt `EncryptedPackage`.
    let encrypted_package = read_ole_stream(&raw_ole, "EncryptedPackage")?;
    let decrypted = agile_decrypt_package(&info, &secret_key, &encrypted_package)?;

    // Sanity check: decrypted OOXML packages are ZIP/OPC containers.
    if decrypted.len() < 2 || &decrypted[..2] != b"PK" {
        return Err(OffcryptoError::InvalidStructure(
            "decrypted package does not look like a ZIP (missing PK signature)".to_string(),
        ));
    }

    Ok(decrypted)
}

fn read_ole_stream(raw_ole: &[u8], stream: &'static str) -> Result<Vec<u8>, OffcryptoError> {
    let cursor = Cursor::new(raw_ole);
    let mut ole = cfb::CompoundFile::open(cursor).map_err(|e| {
        OffcryptoError::InvalidStructure(format!("failed to open OLE compound file: {e}"))
    })?;

    let mut s = match ole.open_stream(stream) {
        Ok(s) => s,
        Err(err) if err.kind() == std::io::ErrorKind::NotFound => {
            // Some writers store streams with a leading `/` even at the root; be permissive.
            let with_slash = format!("/{stream}");
            match ole.open_stream(&with_slash) {
                Ok(s) => s,
                Err(err) if err.kind() == std::io::ErrorKind::NotFound => {
                    return Err(OffcryptoError::InvalidStructure(format!(
                        "missing `{stream}` stream"
                    )));
                }
                Err(err) => {
                    return Err(OffcryptoError::InvalidStructure(format!(
                        "failed to open `{stream}`: {err}"
                    )));
                }
            }
        }
        Err(err) => {
            return Err(OffcryptoError::InvalidStructure(format!(
                "failed to open `{stream}`: {err}"
            )))
        }
    };

    let mut buf = Vec::new();
    s.read_to_end(&mut buf).map_err(|e| {
        OffcryptoError::InvalidStructure(format!("failed to read `{stream}`: {e}"))
    })?;
    Ok(buf)
}

#[cfg(test)]
mod tests {
    use super::*;
    use aes::cipher::BlockEncrypt;

    #[test]
    fn decode_b64_attr_padded() {
        let decoded = decode_b64_attr("AQIDBA==").expect("decode");
        assert_eq!(decoded, vec![1, 2, 3, 4]);
    }

    #[test]
    fn decode_b64_attr_unpadded() {
        let decoded = decode_b64_attr("AQIDBA").expect("decode");
        assert_eq!(decoded, vec![1, 2, 3, 4]);
    }

    #[test]
    fn decode_b64_attr_whitespace() {
        let decoded = decode_b64_attr("A QID\r\nBA==\t").expect("decode");
        assert_eq!(decoded, vec![1, 2, 3, 4]);
    }

    fn minimal_agile_xml() -> String {
        fn b64_no_pad(bytes: &[u8]) -> String {
            let mut s = STANDARD.encode(bytes);
            while s.ends_with('=') {
                s.pop();
            }
            s
        }

        fn with_spaces(s: &str) -> String {
            let mut out = String::with_capacity(s.len() + s.len() / 5);
            for (idx, ch) in s.chars().enumerate() {
                if idx != 0 && idx % 5 == 0 {
                    out.push(' ');
                }
                out.push(ch);
            }
            out
        }

        // Include unpadded base64 and embedded whitespace to match the tolerant
        // decoding behavior required for pretty-printed EncryptionInfo XML.
        let key_data_salt: Vec<u8> = (0u8..16).collect();
        let encrypted_hmac_key: Vec<u8> = (0x10u8..0x30).collect();
        let encrypted_hmac_value: Vec<u8> = (0xA0u8..0xC0).collect();
        let password_salt: Vec<u8> = (1u8..17).collect();
        let encrypted_key_value: Vec<u8> = (0x20u8..0x40).collect();
        let encrypted_verifier_hash_input: Vec<u8> = (0x30u8..0x50).collect();
        let encrypted_verifier_hash_value: Vec<u8> = (0x40u8..0x60).collect();

        let key_data_salt = with_spaces(&b64_no_pad(&key_data_salt));
        let encrypted_hmac_key = with_spaces(&b64_no_pad(&encrypted_hmac_key));
        let encrypted_hmac_value = with_spaces(&b64_no_pad(&encrypted_hmac_value));
        let password_salt = with_spaces(&b64_no_pad(&password_salt));
        let encrypted_key_value = with_spaces(&b64_no_pad(&encrypted_key_value));
        let encrypted_verifier_hash_input =
            with_spaces(&b64_no_pad(&encrypted_verifier_hash_input));
        let encrypted_verifier_hash_value =
            with_spaces(&b64_no_pad(&encrypted_verifier_hash_value));

        format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
    xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
  <keyData cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" saltValue="{key_data_salt}" hashAlgorithm="sha256" blockSize="16"/>
  <dataIntegrity encryptedHmacKey="{encrypted_hmac_key}" encryptedHmacValue="{encrypted_hmac_value}"/>
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" spinCount="100000" saltValue="{password_salt}" hashAlgorithm="sha512" keyBits="256"
        encryptedKeyValue="{encrypted_key_value}"
        encryptedVerifierHashInput="{encrypted_verifier_hash_input}"
        encryptedVerifierHashValue="{encrypted_verifier_hash_value}"/>
    </keyEncryptor>
  </keyEncryptors>
</encryption>
"#
        )
    }

    fn build_agile_encryption_info_stream(payload: &[u8]) -> Vec<u8> {
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&4u16.to_le_bytes());
        bytes.extend_from_slice(&4u16.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(payload);
        bytes
    }

    fn parse_agile_info(payload: &[u8]) -> AgileEncryptionInfo {
        let bytes = build_agile_encryption_info_stream(payload);
        let parsed = parse_encryption_info(&bytes).expect("parse");
        let EncryptionInfo::Agile { info, .. } = parsed else {
            panic!("expected Agile EncryptionInfo");
        };
        info
    }

    #[test]
    fn parses_minimal_agile_encryption_info() {
        let xml = minimal_agile_xml();
        let info = parse_agile_info(xml.as_bytes());

        assert_eq!(info.key_data_salt, (0u8..16).collect::<Vec<_>>());
        assert_eq!(info.key_data_hash_algorithm, HashAlgorithm::Sha256);
        assert_eq!(info.key_data_block_size, 16);

        assert_eq!(info.encrypted_hmac_key, (0x10u8..0x30).collect::<Vec<_>>());
        assert_eq!(info.encrypted_hmac_value, (0xA0u8..0xC0).collect::<Vec<_>>());

        assert_eq!(info.spin_count, 100_000);
        assert_eq!(info.password_salt, (1u8..17).collect::<Vec<_>>());
        assert_eq!(info.password_hash_algorithm, HashAlgorithm::Sha512);
        assert_eq!(info.password_key_bits, 256);
        assert_eq!(info.encrypted_key_value, (0x20u8..0x40).collect::<Vec<_>>());
        assert_eq!(
            info.encrypted_verifier_hash_input,
            (0x30u8..0x50).collect::<Vec<_>>()
        );
        assert_eq!(
            info.encrypted_verifier_hash_value,
            (0x40u8..0x60).collect::<Vec<_>>()
        );
    }

    #[test]
    fn parses_agile_encryption_info_with_utf8_bom_and_trailing_nuls() {
        let xml = minimal_agile_xml();
        let expected = parse_agile_info(xml.as_bytes());

        let mut payload = Vec::new();
        payload.extend_from_slice(&[0xEF, 0xBB, 0xBF]); // UTF-8 BOM
        payload.extend_from_slice(xml.as_bytes());
        payload.extend_from_slice(&[0, 0, 0]); // common padding

        let parsed = parse_agile_info(&payload);
        assert_eq!(parsed, expected);
    }

    #[test]
    fn parses_agile_encryption_info_with_utf16le_xml() {
        let xml = minimal_agile_xml();
        let expected = parse_agile_info(xml.as_bytes());

        let mut payload = Vec::new();
        // UTF-16LE BOM
        payload.extend_from_slice(&[0xFF, 0xFE]);
        for unit in xml.encode_utf16() {
            payload.extend_from_slice(&unit.to_le_bytes());
        }
        // UTF-16 NUL terminator
        payload.extend_from_slice(&[0x00, 0x00]);

        let parsed = parse_agile_info(&payload);
        assert_eq!(parsed, expected);
    }

    #[test]
    fn parses_agile_encryption_info_with_length_prefix() {
        let xml = minimal_agile_xml();
        let expected = parse_agile_info(xml.as_bytes());

        let xml_bytes = xml.as_bytes();
        let mut payload = Vec::new();
        payload.extend_from_slice(&(xml_bytes.len() as u32).to_le_bytes());
        payload.extend_from_slice(xml_bytes);
        payload.extend_from_slice(b"GARBAGE"); // force length-prefix slicing

        let parsed = parse_agile_info(&payload);
        assert_eq!(parsed, expected);
    }

    #[test]
    fn parses_agile_encryption_info_with_leading_bytes_before_xml() {
        let xml = minimal_agile_xml();
        let expected = parse_agile_info(xml.as_bytes());

        let mut payload = Vec::new();
        payload.extend_from_slice(b"JUNK");
        payload.extend_from_slice(xml.as_bytes());

        let parsed = parse_agile_info(&payload);
        assert_eq!(parsed, expected);
    }

    #[test]
    fn rejects_clearly_invalid_agile_payloads() {
        let bytes = build_agile_encryption_info_stream(b"not xml at all");
        let err = parse_encryption_info(&bytes).expect_err("should error");
        assert!(matches!(err, OffcryptoError::InvalidEncryptionInfo { .. }));
    }

    #[test]
    fn agile_verify_password_matches_msoffcrypto_tool_vectors() {
        // Test vectors from `msoffcrypto-tool`:
        // https://github.com/nolze/msoffcrypto-tool/blob/master/msoffcrypto/method/ecma376_agile.py
        // (docstring in `ECMA376Agile.verify_password`).
        let info = AgileEncryptionInfo {
            key_data_salt: Vec::new(),
            key_data_hash_algorithm: HashAlgorithm::Sha512,
            key_data_block_size: 16,
            encrypted_hmac_key: Vec::new(),
            encrypted_hmac_value: Vec::new(),
            spin_count: 100_000,
            password_salt: vec![
                0xCB, 0xCA, 0x1C, 0x99, 0x93, 0x43, 0xFB, 0xAD, 0x92, 0x07, 0x56, 0x34, 0x15,
                0x00, 0x34, 0xB0,
            ],
            password_hash_algorithm: HashAlgorithm::Sha512,
            password_key_bits: 256,
            encrypted_key_value: Vec::new(),
            encrypted_verifier_hash_input: vec![
                0x39, 0xEE, 0xA5, 0x4E, 0x26, 0xE5, 0x14, 0x79, 0x8C, 0x28, 0x4B, 0xC7, 0x71,
                0x4D, 0x38, 0xAC,
            ],
            encrypted_verifier_hash_value: vec![
                0x14, 0x37, 0x6D, 0x6D, 0x81, 0x73, 0x34, 0xE6, 0xB0, 0xFF, 0x4F, 0xD8, 0x22,
                0x1A, 0x7C, 0x67, 0x8E, 0x5D, 0x8A, 0x78, 0x4E, 0x8F, 0x99, 0x9F, 0x4C, 0x18,
                0x89, 0x30, 0xC3, 0x6A, 0x4B, 0x29, 0xC5, 0xB3, 0x33, 0x60, 0x5B, 0x5C, 0xD4,
                0x03, 0xB0, 0x50, 0x03, 0xAD, 0xCF, 0x18, 0xCC, 0xA8, 0xCB, 0xAB, 0x8D, 0xEB,
                0xE3, 0x73, 0xC6, 0x56, 0x04, 0xA0, 0xBE, 0xCF, 0xAE, 0x5C, 0x0A, 0xD0,
            ],
        };

        agile_verify_password(&info, "Password1234_").expect("expected password to verify");
    }

    #[test]
    fn agile_rejects_unsupported_cipher_algorithm() {
        // Start from a valid Agile XML payload so we fail specifically on the algorithm check.
        let xml = minimal_agile_xml().replacen(r#"cipherAlgorithm="AES""#, r#"cipherAlgorithm="DES""#, 1);

        let bytes = build_agile_encryption_info_stream(xml.as_bytes());
        let err = parse_encryption_info(&bytes).expect_err("expected unsupported algorithm");
        assert!(matches!(err, OffcryptoError::UnsupportedAlgorithm(_)));
    }

    #[test]
    fn agile_rejects_unsupported_cipher_chaining() {
        // Start from a valid Agile XML payload so we fail specifically on the chaining mode check.
        let xml = minimal_agile_xml().replacen(
            r#"<p:encryptedKey cipherAlgorithm="AES" cipherChaining="ChainingModeCBC""#,
            r#"<p:encryptedKey cipherAlgorithm="AES" cipherChaining="ChainingModeCFB""#,
            1,
        );

        let bytes = build_agile_encryption_info_stream(xml.as_bytes());
        let err = parse_encryption_info(&bytes).expect_err("expected unsupported algorithm");
        assert!(matches!(err, OffcryptoError::UnsupportedAlgorithm(_)));
    }

    #[test]
    fn inspects_minimal_agile_encryption_info() {
        let xml = minimal_agile_xml();
        let bytes = build_agile_encryption_info_stream(xml.as_bytes());
        let summary = inspect_encryption_info(&bytes).expect("inspect");
        assert_eq!(summary.encryption_type, EncryptionType::Agile);
        assert_eq!(
            summary.agile,
            Some(AgileEncryptionInfoSummary {
                hash_algorithm: HashAlgorithm::Sha512,
                spin_count: 100_000,
                key_bits: 256,
            })
        );
        assert!(summary.standard.is_none());
    }

    #[test]
    fn parses_agile_encryption_info_with_utf8_bom_and_padding() {
        let xml = minimal_agile_xml();

        let mut bytes = Vec::new();
        bytes.extend_from_slice(&4u16.to_le_bytes());
        bytes.extend_from_slice(&4u16.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(&[0xEF, 0xBB, 0xBF]); // UTF-8 BOM
        bytes.extend_from_slice(xml.as_bytes());
        bytes.extend_from_slice(&[0, 0, 0]); // padding

        let parsed = parse_encryption_info(&bytes).expect("parse");
        let EncryptionInfo::Agile { info, .. } = parsed else {
            panic!("expected Agile EncryptionInfo");
        };

        assert_eq!(info.key_data_hash_algorithm, HashAlgorithm::Sha256);
        assert_eq!(info.spin_count, 100_000);
    }

    #[test]
    fn inspects_minimal_standard_encryption_info() {
        // Minimal Standard EncryptionInfo buffer sufficient for `inspect_encryption_info`:
        // - version (major varies; minor=2)
        // - header size + header (AES-256 + SHA1, keySize matches algId)
        // - verifier with saltSize=16, verifierHashSize=20 (SHA1) and a 32-byte encrypted hash
        for major in [2u16, 3u16, 4u16] {
            let mut bytes = Vec::new();
            bytes.extend_from_slice(&major.to_le_bytes());
            bytes.extend_from_slice(&2u16.to_le_bytes());
            bytes.extend_from_slice(&0u32.to_le_bytes());

            let mut header = Vec::new();
            header.extend_from_slice(&0u32.to_le_bytes()); // flags
            header.extend_from_slice(&0u32.to_le_bytes()); // sizeExtra
            header.extend_from_slice(&CALG_AES_256.to_le_bytes()); // algId = CALG_AES_256
            header.extend_from_slice(&CALG_SHA1.to_le_bytes()); // algIdHash = CALG_SHA1
            header.extend_from_slice(&256u32.to_le_bytes()); // keySize
            header.extend_from_slice(&0u32.to_le_bytes()); // providerType
            header.extend_from_slice(&0u32.to_le_bytes()); // reserved1
            header.extend_from_slice(&0u32.to_le_bytes()); // reserved2

            bytes.extend_from_slice(&(header.len() as u32).to_le_bytes());
            bytes.extend_from_slice(&header);

            // EncryptionVerifier
            bytes.extend_from_slice(&16u32.to_le_bytes()); // saltSize
            bytes.extend_from_slice(&[0u8; 16]); // salt
            bytes.extend_from_slice(&[0u8; 16]); // encryptedVerifier
            bytes.extend_from_slice(&20u32.to_le_bytes()); // verifierHashSize (SHA1)
            bytes.extend_from_slice(&[0u8; 32]); // encryptedVerifierHash (SHA1 padded to AES block size)

            let summary = inspect_encryption_info(&bytes).expect("inspect");
            assert_eq!(summary.encryption_type, EncryptionType::Standard);
            assert_eq!(
                summary.standard,
                Some(StandardEncryptionInfoSummary {
                    alg_id: StandardAlgId::Aes256,
                    key_size: 256,
                })
            );
            assert!(summary.agile.is_none());
        }
    }

    #[test]
    fn inspects_minimal_standard_encryption_info_version_4_2() {
        // Same as `inspects_minimal_standard_encryption_info`, but with version 4.2. Some Office
        // producers emit Standard EncryptionInfo with versionMajor=4, versionMinor=2.
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&4u16.to_le_bytes());
        bytes.extend_from_slice(&2u16.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());

        let mut header = Vec::new();
        header.extend_from_slice(&0u32.to_le_bytes()); // flags
        header.extend_from_slice(&0u32.to_le_bytes()); // sizeExtra
        header.extend_from_slice(&CALG_AES_128.to_le_bytes()); // algId = CALG_AES_128
        header.extend_from_slice(&CALG_SHA1.to_le_bytes()); // algIdHash = CALG_SHA1
        header.extend_from_slice(&128u32.to_le_bytes()); // keySize
        header.extend_from_slice(&0u32.to_le_bytes()); // providerType
        header.extend_from_slice(&0u32.to_le_bytes()); // reserved1
        header.extend_from_slice(&0u32.to_le_bytes()); // reserved2

        bytes.extend_from_slice(&(header.len() as u32).to_le_bytes());
        bytes.extend_from_slice(&header);

        bytes.extend_from_slice(&16u32.to_le_bytes()); // saltSize
        bytes.extend_from_slice(&[0u8; 16]); // salt
        bytes.extend_from_slice(&[0u8; 16]); // encryptedVerifier
        bytes.extend_from_slice(&20u32.to_le_bytes()); // verifierHashSize
        bytes.extend_from_slice(&[0u8; 32]); // encryptedVerifierHash

        let summary = inspect_encryption_info(&bytes).expect("inspect");
        assert_eq!(summary.encryption_type, EncryptionType::Standard);
        assert_eq!(
            summary.standard,
            Some(StandardEncryptionInfoSummary {
                alg_id: StandardAlgId::Aes128,
                key_size: 128,
            })
        );
        assert!(summary.agile.is_none());
    }

    #[test]
    fn standard_verify_key_mismatch_uses_constant_time_compare() {
        // Ensure the Standard verifier hash comparison uses the shared constant-time helper.
        util::reset_ct_eq_calls();

        let key = [0u8; 16];

        fn aes128_ecb_encrypt_in_place(key: &[u8; 16], buf: &mut [u8]) {
            let cipher = Aes128::new_from_slice(key).expect("valid AES-128 key");
            for block in buf.chunks_mut(16) {
                cipher.encrypt_block(GenericArray::from_mut_slice(block));
            }
        }

        // Choose plaintext verifier and verifierHash that are guaranteed to mismatch after hashing.
        let mut encrypted_verifier = [0u8; 16]; // plaintext verifier = all zeros
        let mut encrypted_verifier_hash = vec![0u8; 32]; // plaintext verifierHash = all zeros

        aes128_ecb_encrypt_in_place(&key, &mut encrypted_verifier);
        aes128_ecb_encrypt_in_place(&key, &mut encrypted_verifier_hash);

        let info = StandardEncryptionInfo {
            header: StandardEncryptionHeader {
                flags: 0,
                size_extra: 0,
                alg_id: CALG_AES_128,
                alg_id_hash: CALG_SHA1,
                key_size_bits: 128,
                provider_type: 0,
                reserved1: 0,
                reserved2: 0,
                csp_name: String::new(),
            },
            verifier: StandardEncryptionVerifier {
                salt: vec![0u8; 16],
                encrypted_verifier,
                verifier_hash_size: 20,
                encrypted_verifier_hash,
            },
        };

        let err = standard_verify_key(&info, &key).expect_err("expected verifier mismatch");
        assert!(matches!(err, OffcryptoError::InvalidPassword));

        assert!(
            util::ct_eq_call_count() >= 1,
            "expected ct_eq helper to be invoked"
        );
    }

    #[test]
    fn agile_spin_count_just_below_limit_succeeds() {
        let limits = DecryptLimits {
            max_spin_count: Some(10),
        };

        let spin_count = 9;
        let out = derive_iterated_hash_from_password(
            "password",
            b"01234567",
            HashAlgorithm::Sha256,
            spin_count,
            &limits,
            None,
        )
        .expect("spinCount below limit should succeed");

        assert_eq!(out.len(), 32);
        assert!(out.iter().any(|b| *b != 0));
    }

    #[test]
    fn agile_spin_count_above_limit_errors_without_iterating() {
        let limits = DecryptLimits {
            max_spin_count: Some(10),
        };

        // A huge spinCount that would be a CPU DoS without an up-front check.
        let spin_count = u32::MAX;

        let mut iter_hook = |_i: u32| -> () {
            panic!("spinCount loop should not run when over the limit");
        };

        let err = derive_iterated_hash_from_password(
            "password",
            b"01234567",
            HashAlgorithm::Sha256,
            spin_count,
            &limits,
            Some(&mut iter_hook),
        )
        .expect_err("spinCount above limit should error");

        assert_eq!(
            err,
            OffcryptoError::SpinCountTooLarge {
                spin_count,
                max: 10
            }
        );
    }
}
