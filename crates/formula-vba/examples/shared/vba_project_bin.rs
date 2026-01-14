use std::io::Cursor;
use std::path::Path;

#[path = "zip_util.rs"]
mod zip_util;

/// Load `xl/vbaProject.bin` bytes from a workbook path.
///
/// Supports:
/// - Raw `vbaProject.bin` OLE files (returned as-is).
/// - Unencrypted ZIP-based workbooks (`.xlsx`/`.xlsm`/`.xlsb`), extracting `xl/vbaProject.bin`.
/// - Office-encrypted OOXML workbooks (OLE wrapper containing `EncryptionInfo` + `EncryptedPackage`)
///   by decrypting with the provided password and then extracting `xl/vbaProject.bin` from the
///   decrypted ZIP payload.
pub(crate) fn load_vba_project_bin(
    path: &Path,
    password: Option<&str>,
) -> Result<(Vec<u8>, String), String> {
    match try_extract_vba_project_bin_from_zip(path)? {
        Some(bytes) => {
            return Ok((
                bytes,
                format!("{} (zip entry xl/vbaProject.bin)", path.display()),
            ));
        }
        None => {}
    }

    // Not a zip workbook; it could be:
    // - an Office-encrypted OOXML wrapper (OLE container holding EncryptionInfo + EncryptedPackage), or
    // - a raw `vbaProject.bin` OLE container (or other binary input).
    let bytes =
        std::fs::read(path).map_err(|e| format!("failed to read {}: {e}", path.display()))?;

    if formula_office_crypto::is_encrypted_ooxml_ole(&bytes) {
        let password =
            password.ok_or_else(|| "password required for encrypted workbook".to_owned())?;

        let decrypted = formula_office_crypto::decrypt_encrypted_package_ole(&bytes, password)
            .map_err(|e| {
                format!(
                    "failed to decrypt encrypted workbook {}: {e}",
                    path.display()
                )
            })?;

        let bytes = extract_vba_project_bin_from_zip_bytes(&decrypted).map_err(|e| {
            format!(
                "failed to extract xl/vbaProject.bin from decrypted workbook {}: {e}",
                path.display()
            )
        })?;

        return Ok((
            bytes,
            format!(
                "{} (encrypted workbook decrypted; zip entry xl/vbaProject.bin)",
                path.display()
            ),
        ));
    }

    Ok((bytes, path.display().to_string()))
}

fn try_extract_vba_project_bin_from_zip(path: &Path) -> Result<Option<Vec<u8>>, String> {
    let file = match std::fs::File::open(path) {
        Ok(f) => f,
        Err(e) => return Err(format!("failed to open {}: {e}", path.display())),
    };

    let mut archive = match zip::ZipArchive::new(file) {
        Ok(a) => a,
        Err(_) => return Ok(None),
    };

    let Some(buf) = zip_util::read_zip_entry_bytes(&mut archive, "xl/vbaProject.bin")
        .map_err(|e| format!("failed to read zip {}: {e}", path.display()))?
    else {
        return Err(format!(
            "{} is a zip, but does not contain xl/vbaProject.bin",
            path.display()
        ));
    };
    Ok(Some(buf))
}

fn extract_vba_project_bin_from_zip_bytes(zip_bytes: &[u8]) -> Result<Vec<u8>, String> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive =
        zip::ZipArchive::new(cursor).map_err(|e| format!("failed to open zip: {e}"))?;

    let Some(buf) = zip_util::read_zip_entry_bytes(&mut archive, "xl/vbaProject.bin")
        .map_err(|e| format!("failed to read zip: {e}"))?
    else {
        return Err("zip does not contain xl/vbaProject.bin".to_owned());
    };
    Ok(buf)
}
