use std::io::{Cursor, Read, Seek, SeekFrom};

use crate::{XlsxError, XlsxPackage};

fn zip_entry_name_matches(candidate: &str, target: &str) -> bool {
    let target = target.trim_start_matches(|c| c == '/' || c == '\\');
    let target = if target.contains('\\') {
        std::borrow::Cow::Owned(target.replace('\\', "/"))
    } else {
        std::borrow::Cow::Borrowed(target)
    };

    let mut normalized = candidate.trim_start_matches(|c| c == '/' || c == '\\');
    let replaced;
    if normalized.contains('\\') {
        replaced = normalized.replace('\\', "/");
        normalized = replaced.trim_start_matches('/');
    }

    normalized.eq_ignore_ascii_case(target.as_ref())
}

fn zip_has_entry<R: Read + Seek>(archive: &zip::ZipArchive<R>, name: &str) -> bool {
    archive
        .file_names()
        .any(|candidate| zip_entry_name_matches(candidate, name))
}

fn validate_excel_zip(decrypted: &[u8]) -> Result<(), XlsxError> {
    let cursor = Cursor::new(decrypted);
    let archive = zip::ZipArchive::new(cursor).map_err(|e| {
        XlsxError::InvalidEncryptedWorkbook(format!(
            "decrypted payload is not a valid OOXML ZIP package: {e}"
        ))
    })?;

    let has_workbook_xml = zip_has_entry(&archive, "xl/workbook.xml");
    let has_workbook_bin = zip_has_entry(&archive, "xl/workbook.bin");
    if has_workbook_xml || has_workbook_bin {
        return Ok(());
    }

    let looks_like_word = zip_has_entry(&archive, "word/document.xml");
    let looks_like_ppt = zip_has_entry(&archive, "ppt/presentation.xml");

    if looks_like_word {
        return Err(XlsxError::InvalidEncryptedWorkbook(
            "decrypted payload looks like a Word document, not an Excel workbook".to_string(),
        ));
    }
    if looks_like_ppt {
        return Err(XlsxError::InvalidEncryptedWorkbook(
            "decrypted payload looks like a PowerPoint document, not an Excel workbook".to_string(),
        ));
    }

    Err(XlsxError::InvalidEncryptedWorkbook(
        "decrypted payload does not appear to be an Excel workbook (missing xl/workbook.xml or xl/workbook.bin)".to_string(),
    ))
}

fn decrypt_encrypted_ole_bytes(bytes: &[u8], password: &str) -> Result<Vec<u8>, XlsxError> {
    match formula_office_crypto::decrypt_encrypted_package_ole(bytes, password) {
        Ok(bytes) => Ok(bytes),
        Err(formula_office_crypto::OfficeCryptoError::InvalidPassword)
        | Err(formula_office_crypto::OfficeCryptoError::IntegrityCheckFailed) => {
            Err(XlsxError::InvalidPassword)
        }
        Err(formula_office_crypto::OfficeCryptoError::UnsupportedEncryption(msg)) => {
            Err(XlsxError::UnsupportedEncryption(msg))
        }
        Err(formula_office_crypto::OfficeCryptoError::Io(err)) => Err(XlsxError::Io(err)),
        Err(err) => Err(XlsxError::InvalidEncryptedWorkbook(format!(
            "failed to decrypt Office-encrypted workbook: {err}"
        ))),
    }
}

/// Load an [`XlsxPackage`] from an Office-encrypted OLE wrapper (`EncryptionInfo` + `EncryptedPackage`).
///
/// The decrypted payload is validated to ensure it looks like an Excel workbook before attempting
/// to parse it as an OPC ZIP container.
pub fn load_from_encrypted_ole_bytes(
    bytes: &[u8],
    password: &str,
) -> Result<XlsxPackage, XlsxError> {
    let decrypted = decrypt_encrypted_ole_bytes(bytes, password)?;

    validate_excel_zip(&decrypted)?;
    XlsxPackage::from_bytes(&decrypted)
}

/// Read an Excel workbook model from an Office-encrypted OLE wrapper (`EncryptionInfo` + `EncryptedPackage`).
///
/// The decrypted payload is validated to ensure it looks like an Excel workbook before parsing.
pub fn read_workbook_from_encrypted_reader<R: Read + Seek>(
    mut reader: R,
    password: &str,
) -> Result<formula_model::Workbook, XlsxError> {
    reader.seek(SeekFrom::Start(0))?;
    let mut bytes = Vec::new();
    reader.read_to_end(&mut bytes)?;

    let decrypted = decrypt_encrypted_ole_bytes(&bytes, password)?;

    validate_excel_zip(&decrypted)?;
    crate::read_workbook_from_reader(Cursor::new(decrypted))
}
