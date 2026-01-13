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
        "unsupported OOXML encryption version {major}.{minor}; only Agile Encryption (4.4) is supported"
    )]
    UnsupportedEncryptionVersion { major: u16, minor: u16 },

    #[error("unsupported OOXML encryption cipher algorithm `{cipher}`")]
    UnsupportedCipherAlgorithm { cipher: String },

    #[error("unsupported OOXML encryption chaining mode `{chaining}`")]
    UnsupportedChainingMode { chaining: String },

    #[error("unsupported OOXML encryption hash algorithm `{hash}`")]
    UnsupportedHashAlgorithm { hash: String },

    #[error("invalid OOXML Agile encryption parameter: {param}")]
    InvalidAgileParameter { param: &'static str },

    // --- EncryptionInfo XML parsing ------------------------------------------------------------
    #[error("EncryptionInfo XML is not valid UTF-8: {source}")]
    EncryptionInfoXmlNotUtf8 {
        #[source]
        source: std::str::Utf8Error,
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

    // --- Cryptographic verification ------------------------------------------------------------
    #[error("wrong password for encrypted workbook (verifier mismatch)")]
    WrongPassword,

    #[error(
        "encrypted workbook integrity check failed (HMAC mismatch); the file may be corrupted or the password is incorrect"
    )]
    IntegrityMismatch,

    // --- Structural errors ---------------------------------------------------------------------
    #[error("EncryptedPackage stream is too short ({len} bytes)")]
    EncryptedPackageTooShort { len: usize },

    #[error(
        "EncryptedPackage ciphertext length {ciphertext_len} is not a multiple of the block size ({block_size} bytes)"
    )]
    CiphertextNotBlockAligned {
        ciphertext_len: usize,
        block_size: usize,
    },

    #[error(
        "decrypted EncryptedPackage is truncated: header declares {declared_len} bytes but only {available_len} bytes are available"
    )]
    DecryptedLengthShorterThanHeader {
        declared_len: usize,
        available_len: usize,
    },
}

impl From<std::str::Utf8Error> for OffCryptoError {
    fn from(source: std::str::Utf8Error) -> Self {
        Self::EncryptionInfoXmlNotUtf8 { source }
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
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::offcrypto::crypto::{hash_password, HashAlgorithm};

    #[test]
    fn maps_crypto_error_unsupported_hash_algorithm() {
        let err = HashAlgorithm::parse_offcrypto_name("md5").expect_err("md5 not supported");
        let off: OffCryptoError = err.into();
        assert!(
            matches!(off, OffCryptoError::UnsupportedHashAlgorithm { ref hash } if hash == "md5"),
            "expected UnsupportedHashAlgorithm, got {off:?}"
        );
        assert!(
            off.to_string().to_lowercase().contains("encryption"),
            "message should mention encryption context: {}",
            off
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
        assert!(
            off.to_string().to_lowercase().contains("agile"),
            "message should mention Agile encryption context: {}",
            off
        );
    }
}
