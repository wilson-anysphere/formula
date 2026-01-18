use super::aes_cbc::AesCbcDecryptError;
use super::crypto::CryptoError;
use thiserror::Error;

/// Result type for `[MS-OFFCRYPTO]` operations.
pub type Result<T> = std::result::Result<T, OffCryptoError>;

/// Errors returned while parsing or decrypting `[MS-OFFCRYPTO]` encrypted OOXML packages.
///
/// The variants are designed to be actionable for end users ("wrong password" vs "unsupported
/// algorithm") while avoiding accidental exposure of sensitive data (passwords, derived keys).
#[derive(Debug, Error)]
pub enum OffCryptoError {
    // --- High-level capability / compatibility -------------------------------------------------
    #[error(
        "unsupported OOXML encryption version {major}.{minor}; supported versions are Agile Encryption (4.4) and Standard Encryption (minor=2; major=2/3/4)"
    )]
    UnsupportedEncryptionVersion { major: u16, minor: u16 },

    #[error("unsupported OOXML encryption cipher algorithm `{cipher}`")]
    UnsupportedCipherAlgorithm { cipher: String },

    #[error("unsupported OOXML encryption chaining mode `{chaining}`")]
    UnsupportedChainingMode { chaining: String },

    #[error(
        "unsupported OOXML encryption cipher chaining mode `{chaining}`; only `ChainingModeCBC` is supported"
    )]
    UnsupportedCipherChaining { chaining: String },

    #[error("unsupported OOXML encryption hash algorithm `{hash}`")]
    UnsupportedHashAlgorithm { hash: String },

    #[error("invalid AES block size {block_size} bytes (expected 16)")]
    InvalidBlockSize { block_size: usize },

    #[error("invalid OOXML Agile encryption parameter: {param}")]
    InvalidAgileParameter { param: &'static str },

    #[error("allocation failure: {0}")]
    AllocationFailure(&'static str),

    // --- OLE/CFB container helpers --------------------------------------------------------------
    #[error("missing required OLE stream `{stream}`")]
    MissingRequiredStream { stream: String },

    // --- I/O -----------------------------------------------------------------------------------
    #[error("I/O error while {context}: {source}")]
    Io {
        context: &'static str,
        #[source]
        source: std::io::Error,
    },

    // --- Size / DoS limits --------------------------------------------------------------------
    #[error("EncryptionInfo XML is too large ({len} bytes; max {max} bytes)")]
    EncryptionInfoTooLarge { len: usize, max: usize },

    #[error("EncryptionInfo field `{field}` is too large ({len} bytes; max {max} bytes)")]
    FieldTooLarge {
        field: &'static str,
        len: usize,
        max: usize,
    },

    // --- EncryptionInfo XML parsing ------------------------------------------------------------
    #[error("EncryptionInfo XML is not valid UTF-8: {source}")]
    EncryptionInfoXmlNotUtf8 {
        #[source]
        source: std::str::Utf8Error,
    },

    #[error("EncryptionInfo XML is not valid UTF-16LE: {source}")]
    EncryptionInfoXmlNotUtf16 {
        #[source]
        source: std::string::FromUtf16Error,
    },

    #[error("failed to parse EncryptionInfo XML: {source}")]
    EncryptionInfoXmlMalformed {
        #[source]
        source: roxmltree::Error,
    },

    #[error("EncryptionInfo XML missing required element `{element}`")]
    MissingRequiredElement { element: String },

    #[error("EncryptionInfo XML missing required attribute `{attr}` on element `{element}`")]
    MissingRequiredAttribute { element: String, attr: String },

    #[error("EncryptionInfo XML invalid attribute `{attr}` on element `{element}`: {reason}")]
    InvalidAttribute {
        element: String,
        attr: String,
        reason: String,
    },

    #[error("{message}")]
    UnsupportedKeyEncryptor {
        /// The set of `<keyEncryptor uri="...">` values present in the file.
        ///
        /// Office can emit multiple encryptors (e.g. password + certificate). Formula currently
        /// supports only the password key encryptor.
        available_uris: Vec<String>,
        /// User-facing error message (pre-formatted so it can include helpful context/hints).
        message: String,
    },

    #[error(
         "EncryptionInfo XML invalid base64 value for attribute `{attr}` on element `{element}`: {source}"
    )]
    Base64Decode {
        element: String,
        attr: String,
        #[source]
        source: base64::DecodeError,
    },

    #[error(
        "EncryptionInfo XML spinCount {spin_count} exceeds maximum allowed {max} (refusing to run expensive password KDF)"
    )]
    SpinCountTooLarge { spin_count: u32, max: u32 },

    // --- Cryptographic verification ------------------------------------------------------------
    #[error("wrong password for encrypted workbook (verifier mismatch)")]
    WrongPassword,

    #[error(
        "encrypted workbook integrity check failed (HMAC mismatch); the file may be corrupted or the password is incorrect"
    )]
    IntegrityMismatch,

    // --- Structural errors ---------------------------------------------------------------------
    #[error("EncryptionInfo stream is too short ({len} bytes)")]
    EncryptionInfoTooShort { len: usize },

    #[error("EncryptedPackage stream is too short ({len} bytes)")]
    EncryptedPackageTooShort { len: usize },

    #[error("{field} ciphertext length {len} is not a multiple of the AES block size (16 bytes)")]
    CiphertextNotBlockAligned { field: &'static str, len: usize },

    #[error(
        "decrypted EncryptedPackage is truncated: header declares {declared_len} bytes but only {available_len} bytes are available"
    )]
    DecryptedLengthShorterThanHeader {
        declared_len: usize,
        available_len: usize,
    },

    // --- Standard (CryptoAPI) parsing ----------------------------------------------------------
    #[error("Standard EncryptionInfo is malformed: {reason}")]
    StandardEncryptionInfoMalformed { reason: String },
}

impl From<std::str::Utf8Error> for OffCryptoError {
    fn from(source: std::str::Utf8Error) -> Self {
        Self::EncryptionInfoXmlNotUtf8 { source }
    }
}

impl From<std::string::FromUtf16Error> for OffCryptoError {
    fn from(source: std::string::FromUtf16Error) -> Self {
        Self::EncryptionInfoXmlNotUtf16 { source }
    }
}

impl From<roxmltree::Error> for OffCryptoError {
    fn from(source: roxmltree::Error) -> Self {
        Self::EncryptionInfoXmlMalformed { source }
    }
}

impl From<CryptoError> for OffCryptoError {
    fn from(source: CryptoError) -> Self {
        match source {
            CryptoError::UnsupportedHashAlgorithm(hash) => Self::UnsupportedHashAlgorithm { hash },
            CryptoError::InvalidParameter(param) => Self::InvalidAgileParameter { param },
            CryptoError::AllocationFailure(ctx) => Self::AllocationFailure(ctx),
        }
    }
}

impl From<AesCbcDecryptError> for OffCryptoError {
    fn from(source: AesCbcDecryptError) -> Self {
        match source {
            AesCbcDecryptError::UnsupportedKeyLength(len) => Self::UnsupportedCipherAlgorithm {
                cipher: format!("AES (key length {len} bytes)"),
            },
            AesCbcDecryptError::InvalidIvLength(_len) => {
                // The caller is expected to derive the IV from the encryption parameters.
                // Treat an unexpected IV length as a format/parameter issue (not a crypto failure).
                Self::InvalidAgileParameter {
                    param: "AES-CBC IV length",
                }
            }
            AesCbcDecryptError::InvalidCiphertextLength(len) => Self::CiphertextNotBlockAligned {
                // The caller is expected to validate ciphertext lengths and surface a more specific
                // `field` string. Keep a deterministic fallback for any unexpected propagation.
                field: "ciphertext",
                len,
            },
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::offcrypto::crypto::{hash_password, HashAlgorithm};
    use crate::offcrypto::AesCbcDecryptError;

    #[test]
    fn maps_crypto_error_unsupported_hash_algorithm() {
        let err = HashAlgorithm::parse_offcrypto_name("md5").expect_err("md5 not supported");
        let off: OffCryptoError = err.into();
        assert!(
            matches!(off, OffCryptoError::UnsupportedHashAlgorithm { ref hash } if hash == "md5"),
            "expected UnsupportedHashAlgorithm, got {off:?}"
        );
    }

    #[test]
    fn maps_crypto_error_invalid_parameter() {
        let err = hash_password("pw", &[], 0, HashAlgorithm::Sha1).expect_err("empty salt");
        let off: OffCryptoError = err.into();
        assert!(
            matches!(off, OffCryptoError::InvalidAgileParameter { param } if param.contains("salt")),
            "expected InvalidAgileParameter, got {off:?}"
        );
    }

    #[test]
    fn maps_aes_cbc_ciphertext_length_error_to_block_alignment() {
        let err = AesCbcDecryptError::InvalidCiphertextLength(15);
        let off: OffCryptoError = err.into();
        assert!(
            matches!(
                off,
                OffCryptoError::CiphertextNotBlockAligned {
                    field: "ciphertext",
                    len: 15
                }
            ),
            "expected CiphertextNotBlockAligned, got {off:?}"
        );
        assert!(
            off.to_string().to_lowercase().contains("ciphertext length"),
            "message should mention ciphertext length context: {}",
            off
        );
    }
}
