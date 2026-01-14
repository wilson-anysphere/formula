use std::borrow::Cow;
use std::io::Cursor;
use std::io::Read;

use thiserror::Error;

const OLE_MAGIC: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

#[derive(Debug, Error)]
pub(crate) enum EncryptedOoxmlError {
    #[error("invalid password or corrupted encrypted workbook (decrypted payload is not a valid ZIP archive)")]
    InvalidPassword,
    #[error("unsupported encryption: {0}")]
    UnsupportedEncryption(String),
    #[error("invalid encrypted workbook: {0}")]
    InvalidEncryptedWorkbook(String),
    #[error("io error: {0}")]
    Io(#[from] std::io::Error),
}

fn looks_like_zip(bytes: &[u8]) -> bool {
    bytes.starts_with(b"PK")
}

fn looks_like_ole(bytes: &[u8]) -> bool {
    bytes.starts_with(&OLE_MAGIC)
}

fn open_stream<R: Read + std::io::Seek>(
    ole: &mut cfb::CompoundFile<R>,
    name: &str,
) -> Result<cfb::Stream<R>, std::io::Error> {
    match ole.open_stream(name) {
        Ok(s) => Ok(s),
        Err(err1) => {
            let with_slash = format!("/{name}");
            ole.open_stream(&with_slash).map_err(|_err2| err1)
        }
    }
}

fn stream_exists<R: Read + std::io::Seek>(ole: &mut cfb::CompoundFile<R>, name: &str) -> bool {
    open_stream(ole, name).is_ok()
}

fn map_offcrypto_error(err: crate::offcrypto::OffCryptoError) -> EncryptedOoxmlError {
    use crate::offcrypto::OffCryptoError;
    match err {
        OffCryptoError::WrongPassword | OffCryptoError::IntegrityMismatch => {
            EncryptedOoxmlError::InvalidPassword
        }
        OffCryptoError::UnsupportedEncryptionVersion { .. }
        | OffCryptoError::UnsupportedCipherAlgorithm { .. }
        | OffCryptoError::UnsupportedCipherChaining { .. }
        | OffCryptoError::UnsupportedChainingMode { .. }
        | OffCryptoError::UnsupportedHashAlgorithm { .. } => {
            EncryptedOoxmlError::UnsupportedEncryption(err.to_string())
        }
        OffCryptoError::UnsupportedKeyEncryptor { message, .. } => {
            EncryptedOoxmlError::UnsupportedEncryption(message)
        }
        other => EncryptedOoxmlError::InvalidEncryptedWorkbook(other.to_string()),
    }
}

pub(crate) fn maybe_decrypt_office_encrypted_package<'a>(
    bytes: &'a [u8],
    password: &str,
) -> Result<Cow<'a, [u8]>, EncryptedOoxmlError> {
    // Common fast path: ordinary XLSX/XLSM are ZIP-based OPC archives.
    if looks_like_zip(bytes) {
        return Ok(Cow::Borrowed(bytes));
    }
    if !looks_like_ole(bytes) {
        return Ok(Cow::Borrowed(bytes));
    }

    // Encrypted OOXML workbooks are OLE containers holding `EncryptionInfo` + `EncryptedPackage`
    // streams (MS-OFFCRYPTO).
    let cursor = Cursor::new(bytes);
    let Ok(mut ole) = cfb::CompoundFile::open(cursor) else {
        return Ok(Cow::Borrowed(bytes));
    };
    if !(stream_exists(&mut ole, "EncryptionInfo") && stream_exists(&mut ole, "EncryptedPackage")) {
        return Ok(Cow::Borrowed(bytes));
    }

    let mut encryption_info = Vec::new();
    {
        let mut stream = open_stream(&mut ole, "EncryptionInfo")?;
        stream.read_to_end(&mut encryption_info)?;
    }

    let mut encrypted_package = Vec::new();
    {
        let mut stream = open_stream(&mut ole, "EncryptedPackage")?;
        stream.read_to_end(&mut encrypted_package)?;
    }

    let decrypted = crate::offcrypto::decrypt_ooxml_encrypted_package(
        &encryption_info,
        &encrypted_package,
        password,
    )
    .map_err(map_offcrypto_error)?;

    // The decrypted content should be the underlying ZIP package (`.xlsx`/`.xlsm`/`.xlsb`). Sanity check
    // with ZIP parsing so callers get a clearer error than "unexpected EOF" later.
    if zip::ZipArchive::new(Cursor::new(decrypted.as_slice())).is_err() {
        return Err(EncryptedOoxmlError::InvalidEncryptedWorkbook(
            "decrypted payload is not a valid ZIP archive".to_string(),
        ));
    }

    Ok(Cow::Owned(decrypted))
}
