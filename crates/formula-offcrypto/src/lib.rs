//! MS-OFFCRYPTO parsing utilities.
//!
//! This crate currently supports parsing the *Standard* (CryptoAPI) `EncryptionInfo`
//! stream header (version 3.2) as well as the `EncryptedPackage` stream header.

use core::fmt;

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

/// Parsed `EncryptionInfo`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum EncryptionInfo {
    /// Standard (CryptoAPI) encryption (MS-OFFCRYPTO version 3.2).
    Standard {
        version: EncryptionVersionInfo,
        header: StandardEncryptionHeader,
        verifier: StandardEncryptionVerifier,
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
    /// The `EncryptionInfo` version is not supported by the current parser.
    UnsupportedVersion { major: u16, minor: u16 },
}

impl fmt::Display for OffcryptoError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            OffcryptoError::Truncated { context } => {
                write!(f, "truncated data while reading {context}")
            }
            OffcryptoError::InvalidCspNameUtf16 => write!(f, "invalid UTF-16LE CSPName"),
            OffcryptoError::UnsupportedVersion { major, minor } => {
                write!(f, "unsupported EncryptionInfo version {major}.{minor}")
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

    // EncryptionVerifier occupies the remaining bytes after the header.
    let salt_size = r.read_u32_le("EncryptionVerifier.saltSize")? as usize;
    let salt = r
        .take(salt_size, "EncryptionVerifier.salt")?
        .to_vec();

    let enc_ver = r.take(16, "EncryptionVerifier.encryptedVerifier")?;
    let mut encrypted_verifier = [0u8; 16];
    encrypted_verifier.copy_from_slice(enc_ver);

    let verifier_hash_size = r.read_u32_le("EncryptionVerifier.verifierHashSize")?;
    let encrypted_verifier_hash = r.remaining().to_vec();

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

/// Parse the 8-byte header at the start of an MS-OFFCRYPTO `EncryptedPackage` stream.
pub fn parse_encrypted_package_header(
    bytes: &[u8],
) -> Result<EncryptedPackageHeader, OffcryptoError> {
    let mut r = Reader::new(bytes);
    let original_size = r.read_u64_le("EncryptedPackageHeader.original_size")?;
    Ok(EncryptedPackageHeader { original_size })
}

