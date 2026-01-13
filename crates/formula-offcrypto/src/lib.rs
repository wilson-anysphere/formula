//! MS-OFFCRYPTO parsing and crypto utilities.
//!
//! This crate currently supports:
//! - Parsing the *Standard* (CryptoAPI) `EncryptionInfo` stream header (version 3.2)
//! - Parsing the *Agile* `EncryptionInfo` stream (version 4.4) (password key-encryptor subset)
//! - Parsing the `EncryptedPackage` stream header
//! - ECMA-376 Standard password→key derivation + verifier checks

use core::fmt;
use aes::cipher::{generic_array::GenericArray, BlockDecrypt, KeyInit};
use aes::{Aes128, Aes192, Aes256};

use base64::engine::general_purpose::{STANDARD, STANDARD_NO_PAD};
use base64::Engine;
use quick_xml::events::Event as XmlEvent;
use quick_xml::Reader as XmlReader;
use sha1::{Digest as _, Sha1};

const ITER_COUNT: u32 = 50_000;
const SHA1_LEN: usize = 20;

pub mod encrypted_package;
pub use encrypted_package::decrypt_encrypted_package;

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
    /// Standard (CryptoAPI) encryption (MS-OFFCRYPTO version 3.2).
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
    /// CSPName was not valid UTF-16LE.
    InvalidCspNameUtf16,
    /// Standard encryption uses an algorithm not supported by the current implementation.
    UnsupportedAlgorithm(u32),
    /// The stream contents are structurally invalid (e.g. missing required attributes).
    InvalidEncryptionInfo { context: &'static str },
    /// The decrypted package size from the `EncryptedPackage` header does not fit into a `Vec<u8>`.
    EncryptedPackageSizeOverflow { total_size: u64 },
    /// Failed to reserve memory for the decrypted output buffer.
    EncryptedPackageAllocationFailed { total_size: u64 },
    /// The `EncryptionInfo` version is not supported by the current parser.
    UnsupportedVersion { major: u16, minor: u16 },
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
}

impl fmt::Display for OffcryptoError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            OffcryptoError::Truncated { context } => {
                write!(f, "truncated data while reading {context}")
            }
            OffcryptoError::InvalidCspNameUtf16 => write!(f, "invalid UTF-16LE CSPName"),
            OffcryptoError::UnsupportedAlgorithm(id) => {
                write!(f, "unsupported encryption algorithm id 0x{id:08X}")
            }
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
            OffcryptoError::UnsupportedVersion { major, minor } => {
                write!(f, "unsupported EncryptionInfo version {major}.{minor}")
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

    if (major, minor) != (3, 2) {
        return Ok(EncryptionInfo::Unsupported { version });
    }

    let header_size = r.read_u32_le("EncryptionInfo.header_size")? as usize;
    let header_bytes = r.take(header_size, "EncryptionHeader")?;

    if header_bytes.len() < 8 * 4 {
        return Err(OffcryptoError::Truncated {
            context: "EncryptionHeader (missing fixed fields)",
        });
    }

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
        other => return Err(OffcryptoError::UnsupportedAlgorithm(other)),
    };
    if header.key_size_bits != expected_key_size {
        return Err(OffcryptoError::UnsupportedAlgorithm(header.alg_id));
    }
    if header.alg_id_hash != CALG_SHA1 {
        return Err(OffcryptoError::UnsupportedAlgorithm(header.alg_id_hash));
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

fn parse_agile_encryption_info_xml(xml_bytes: &[u8]) -> Result<AgileEncryptionInfo, OffcryptoError> {
    let xml = std::str::from_utf8(xml_bytes).map_err(|_| OffcryptoError::InvalidEncryptionInfo {
        context: "agile EncryptionInfo XML is not valid UTF-8",
    })?;

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
            b"saltValue" => {
                salt_value = Some(decode_base64(value)?);
            }
            b"hashAlgorithm" => {
                let s = std::str::from_utf8(value).map_err(|_| OffcryptoError::InvalidEncryptionInfo {
                    context: "invalid UTF-8 attribute value",
                })?;
                hash_algorithm = Some(HashAlgorithm::parse_offcrypto_name(s)?);
            }
            b"blockSize" => {
                block_size = Some(parse_decimal_usize(value, "blockSize")?);
            }
            _ => {}
        }
    }

    Ok((
        salt_value.ok_or(OffcryptoError::InvalidEncryptionInfo {
            context: "missing keyData.saltValue",
        })?,
        hash_algorithm.ok_or(OffcryptoError::InvalidEncryptionInfo {
            context: "missing keyData.hashAlgorithm",
        })?,
        block_size.ok_or(OffcryptoError::InvalidEncryptionInfo {
            context: "missing keyData.blockSize",
        })?,
    ))
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
                encrypted_hmac_key = Some(decode_base64(value)?);
            }
            b"encryptedHmacValue" => {
                encrypted_hmac_value = Some(decode_base64(value)?);
            }
            _ => {}
        }
    }

    Ok((
        encrypted_hmac_key.ok_or(OffcryptoError::InvalidEncryptionInfo {
            context: "missing dataIntegrity.encryptedHmacKey",
        })?,
        encrypted_hmac_value.ok_or(OffcryptoError::InvalidEncryptionInfo {
            context: "missing dataIntegrity.encryptedHmacValue",
        })?,
    ))
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
            b"spinCount" => spin_count = Some(parse_decimal_u32(value, "spinCount")?),
            b"saltValue" => salt_value = Some(decode_base64(value)?),
            b"hashAlgorithm" => {
                let s = std::str::from_utf8(value).map_err(|_| OffcryptoError::InvalidEncryptionInfo {
                    context: "invalid UTF-8 attribute value",
                })?;
                hash_algorithm = Some(HashAlgorithm::parse_offcrypto_name(s)?);
            }
            b"keyBits" => key_bits = Some(parse_decimal_usize(value, "keyBits")?),
            b"encryptedKeyValue" => encrypted_key_value = Some(decode_base64(value)?),
            b"encryptedVerifierHashInput" => {
                encrypted_verifier_hash_input = Some(decode_base64(value)?)
            }
            b"encryptedVerifierHashValue" => {
                encrypted_verifier_hash_value = Some(decode_base64(value)?)
            }
            _ => {}
        }
    }

    Ok((
        spin_count.ok_or(OffcryptoError::InvalidEncryptionInfo {
            context: "missing encryptedKey.spinCount",
        })?,
        salt_value.ok_or(OffcryptoError::InvalidEncryptionInfo {
            context: "missing encryptedKey.saltValue",
        })?,
        hash_algorithm.ok_or(OffcryptoError::InvalidEncryptionInfo {
            context: "missing encryptedKey.hashAlgorithm",
        })?,
        key_bits.ok_or(OffcryptoError::InvalidEncryptionInfo {
            context: "missing encryptedKey.keyBits",
        })?,
        encrypted_key_value.ok_or(OffcryptoError::InvalidEncryptionInfo {
            context: "missing encryptedKey.encryptedKeyValue",
        })?,
        encrypted_verifier_hash_input.ok_or(OffcryptoError::InvalidEncryptionInfo {
            context: "missing encryptedKey.encryptedVerifierHashInput",
        })?,
        encrypted_verifier_hash_value.ok_or(OffcryptoError::InvalidEncryptionInfo {
            context: "missing encryptedKey.encryptedVerifierHashValue",
        })?,
    ))
}

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

fn decode_base64(value: &[u8]) -> Result<Vec<u8>, OffcryptoError> {
    let s = std::str::from_utf8(value).map_err(|_| OffcryptoError::InvalidEncryptionInfo {
        context: "invalid UTF-8 base64 value",
    })?;
    decode_b64_attr(s)
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

fn password_to_utf16le_bytes(password: &str) -> Vec<u8> {
    let mut out = Vec::with_capacity(password.len() * 2);
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

/// ECMA-376 Standard Encryption password→key derivation.
///
/// Reference algorithm: `msoffcrypto` `ECMA376Standard.makekey_from_password`.
pub fn standard_derive_key(
    info: &StandardEncryptionInfo,
    password: &str,
) -> Result<Vec<u8>, OffcryptoError> {
    let key_len = match info.header.key_size_bits.checked_div(8) {
        Some(v) if info.header.key_size_bits % 8 == 0 => v as usize,
        _ => {
            return Err(OffcryptoError::InvalidKeySizeBits {
                key_size_bits: info.header.key_size_bits,
            })
        }
    };

    let password_utf16 = password_to_utf16le_bytes(password);

    // h = sha1(salt || password_utf16)
    let mut hasher = Sha1::new();
    hasher.update(&info.verifier.salt);
    hasher.update(&password_utf16);
    let mut h: [u8; SHA1_LEN] = hasher.finalize().into();

    // for i in 0..ITER_COUNT-1: h = sha1(u32le(i) || h)
    let mut buf = [0u8; 4 + SHA1_LEN];
    for i in 0..ITER_COUNT {
        buf[..4].copy_from_slice(&(i as u32).to_le_bytes());
        buf[4..].copy_from_slice(&h);
        h = sha1(&buf);
    }

    // hfinal = sha1(h || u32le(0))
    let mut buf0 = [0u8; SHA1_LEN + 4];
    buf0[..SHA1_LEN].copy_from_slice(&h);
    buf0[SHA1_LEN..].copy_from_slice(&0u32.to_le_bytes());
    let hfinal = sha1(&buf0);

    // key = (sha1((0x36*64) ^ hfinal) || sha1((0x5c*64) ^ hfinal))[..key_len]
    let mut buf1 = [0x36u8; 64];
    let mut buf2 = [0x5cu8; 64];
    for i in 0..SHA1_LEN {
        buf1[i] ^= hfinal[i];
        buf2[i] ^= hfinal[i];
    }
    let x1 = sha1(&buf1);
    let x2 = sha1(&buf2);

    let mut out = [0u8; SHA1_LEN * 2];
    out[..SHA1_LEN].copy_from_slice(&x1);
    out[SHA1_LEN..].copy_from_slice(&x2);

    if key_len > out.len() {
        return Err(OffcryptoError::DerivedKeyTooLong {
            key_size_bits: info.header.key_size_bits,
            required_bytes: key_len,
            available_bytes: out.len(),
        });
    }

    Ok(out[..key_len].to_vec())
}

/// ECMA-376 Standard Encryption key verifier check.
///
/// Reference algorithm: `msoffcrypto` `ECMA376Standard.verifykey`.
pub fn standard_verify_key(info: &StandardEncryptionInfo, key: &[u8]) -> Result<(), OffcryptoError> {
    let mut verifier = info.verifier.encrypted_verifier;
    aes_ecb_decrypt_in_place(key, &mut verifier)?;
    let expected_hash: [u8; SHA1_LEN] = sha1(&verifier);

    let mut verifier_hash = info.verifier.encrypted_verifier_hash.clone();
    aes_ecb_decrypt_in_place(key, &mut verifier_hash)?;
    if verifier_hash.len() < SHA1_LEN {
        return Err(OffcryptoError::InvalidVerifierHashLength {
            len: verifier_hash.len(),
        });
    }

    if &expected_hash[..] == &verifier_hash[..SHA1_LEN] {
        Ok(())
    } else {
        Err(OffcryptoError::InvalidPassword)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

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

    #[test]
    fn parses_minimal_agile_encryption_info() {
        // Include unpadded base64 and embedded whitespace to match the tolerant
        // decoding behavior required for pretty-printed EncryptionInfo XML.
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
    xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
  <keyData saltValue="AAECAwQF BgcICQoLDA0ODw" hashAlgorithm="SHA256" blockSize="16"/>
  <dataIntegrity encryptedHmacKey="EBE SEw" encryptedHmacValue="q rvM"/>
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey spinCount="100000" saltValue="AQID BA" hashAlgorithm="SHA512" keyBits="256"
        encryptedKeyValue="BQY HCA"
        encryptedVerifierHashInput="CQoL DA"
        encryptedVerifierHashValue="DQ4P EA"/>
    </keyEncryptor>
  </keyEncryptors>
</encryption>
"#;

        let mut bytes = Vec::new();
        bytes.extend_from_slice(&4u16.to_le_bytes());
        bytes.extend_from_slice(&4u16.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(xml.as_bytes());

        let parsed = parse_encryption_info(&bytes).expect("parse");
        let EncryptionInfo::Agile { info, .. } = parsed else {
            panic!("expected Agile EncryptionInfo");
        };

        assert_eq!(info.key_data_salt, (0u8..16).collect::<Vec<_>>());
        assert_eq!(info.key_data_hash_algorithm, HashAlgorithm::Sha256);
        assert_eq!(info.key_data_block_size, 16);

        assert_eq!(info.encrypted_hmac_key, vec![0x10, 0x11, 0x12, 0x13]);
        assert_eq!(info.encrypted_hmac_value, vec![0xaa, 0xbb, 0xcc]);

        assert_eq!(info.spin_count, 100_000);
        assert_eq!(info.password_salt, vec![1, 2, 3, 4]);
        assert_eq!(info.password_hash_algorithm, HashAlgorithm::Sha512);
        assert_eq!(info.password_key_bits, 256);
        assert_eq!(info.encrypted_key_value, vec![5, 6, 7, 8]);
        assert_eq!(info.encrypted_verifier_hash_input, vec![9, 10, 11, 12]);
        assert_eq!(info.encrypted_verifier_hash_value, vec![13, 14, 15, 16]);
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
