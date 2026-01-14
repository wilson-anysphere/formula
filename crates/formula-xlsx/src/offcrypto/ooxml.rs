use super::{decrypt_agile_encrypted_package, OffCryptoError, Result};

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
    // vary the major version across Office generations (2/3/4).
    match (major, minor) {
        (4, 4) => decrypt_agile_encrypted_package(
            encryption_info_stream,
            encrypted_package_stream,
            password,
        ),
        (major, 2) if (2..=4).contains(&major) => {
            decrypt_standard(encryption_info_stream, encrypted_package_stream, password)
        }
        _ => Err(OffCryptoError::UnsupportedEncryptionVersion { major, minor }),
    }
}

fn decrypt_standard(
    encryption_info_stream: &[u8],
    encrypted_package_stream: &[u8],
    password: &str,
) -> Result<Vec<u8>> {
    formula_offcrypto::decrypt_ooxml_standard(
        encryption_info_stream,
        encrypted_package_stream,
        password,
    )
    .map_err(|err| {
        map_standard_error(
            err,
            encryption_info_stream.len(),
            encrypted_package_stream.len(),
        )
    })
}

fn map_standard_error(
    err: formula_offcrypto::OffcryptoError,
    encryption_info_len: usize,
    encrypted_package_len: usize,
) -> OffCryptoError {
    use formula_offcrypto::OffcryptoError as OE;
    match err {
        OE::InvalidPassword => OffCryptoError::WrongPassword,
        OE::IntegrityCheckFailed => OffCryptoError::IntegrityMismatch,
        OE::SpinCountTooLarge { spin_count, max } => {
            OffCryptoError::SpinCountTooLarge { spin_count, max }
        }
        OE::UnsupportedVersion { major, minor } => {
            OffCryptoError::UnsupportedEncryptionVersion { major, minor }
        }
        OE::UnsupportedAlgorithm(reason) => {
            if reason.starts_with("algIdHash=") {
                OffCryptoError::UnsupportedHashAlgorithm { hash: reason }
            } else {
                OffCryptoError::UnsupportedCipherAlgorithm { cipher: reason }
            }
        }
        OE::InvalidCiphertextLength { len } => OffCryptoError::CiphertextNotBlockAligned {
            field: "EncryptedPackage",
            len,
        },
        OE::Truncated { context } => {
            // Preserve the higher-level framing errors where possible; otherwise fall back to a
            // structured Standard error with the underlying reason.
            if context.starts_with("Encryption") {
                OffCryptoError::EncryptionInfoTooShort {
                    len: encryption_info_len,
                }
            } else if context.starts_with("EncryptedPackage") {
                OffCryptoError::EncryptedPackageTooShort {
                    len: encrypted_package_len,
                }
            } else {
                OffCryptoError::StandardEncryptionInfoMalformed {
                    reason: err.to_string(),
                }
            }
        }
        // Keep the mapping conservative: for now we treat other Standard errors as
        // "malformed/unsupported" rather than attempting to mirror every `formula-offcrypto` error
        // variant.
        _ => OffCryptoError::StandardEncryptionInfoMalformed {
            reason: err.to_string(),
        },
    }
}
