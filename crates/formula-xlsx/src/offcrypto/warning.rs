use super::crypto::HashAlgorithm;

/// Non-fatal anomalies detected while parsing/decrypting MS-OFFCRYPTO containers.
///
/// Warnings are intended for diagnostics/telemetry and must **never** include sensitive data
/// (passwords, derived keys, decrypted bytes).
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum OffCryptoWarning {
    /// Multiple password `<keyEncryptor>` entries were present.
    ///
    /// Decryption is deterministic: the first password key encryptor wins.
    MultiplePasswordKeyEncryptors { count: usize },

    /// The document declared a `hashSize` that differs from the typical default for the selected
    /// hash algorithm (the full digest length), but is still usable for decryption.
    NonStandardHashSize {
        element: &'static str,
        hash_algorithm: HashAlgorithm,
        hash_size: usize,
        expected_size: usize,
    },

    /// The document declared a `saltSize` that differs from typical Excel defaults (16 bytes), but
    /// is still usable for decryption.
    NonStandardSaltSize {
        element: &'static str,
        salt_size: usize,
        expected_size: usize,
    },

    /// The document declared a `saltSize` that does not match the decoded `saltValue` length.
    ///
    /// Decryption uses the actual `saltValue` bytes, so this is treated as a warning rather than an
    /// error.
    SaltSizeMismatch {
        element: &'static str,
        declared_salt_size: usize,
        salt_value_len: usize,
    },

    /// An XML element was present that Formula does not recognize.
    UnrecognizedXmlElement { element: String },

    /// An XML attribute was present on a recognized element that Formula does not recognize.
    UnrecognizedXmlAttribute { element: String, attr: String },
}

