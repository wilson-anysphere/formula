use std::io::{Cursor, Read, Seek};
use std::path::Path;

#[path = "zip_util.rs"]
mod zip_util;

const OLE_MAGIC: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

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

    // If the file is an encrypted OOXML wrapper (OLE + EncryptionInfo/EncryptedPackage), require a
    // password and decrypt.
    if looks_like_encrypted_ooxml(path)? {
        let password =
            password.ok_or_else(|| "password required for encrypted workbook".to_owned())?;

        let decrypted = office_crypto::decrypt_from_file(path, password).map_err(|e| {
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

    // Not a zip workbook; treat as a raw vbaProject.bin OLE file.
    let bytes =
        std::fs::read(path).map_err(|e| format!("failed to read {}: {e}", path.display()))?;
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

fn looks_like_encrypted_ooxml(path: &Path) -> Result<bool, String> {
    use std::io::SeekFrom;

    let mut file =
        std::fs::File::open(path).map_err(|e| format!("failed to open {}: {e}", path.display()))?;

    let mut magic = [0u8; 8];
    match file.read_exact(&mut magic) {
        Ok(()) => {}
        Err(err) if err.kind() == std::io::ErrorKind::UnexpectedEof => return Ok(false),
        Err(err) => return Err(format!("failed to read {}: {err}", path.display())),
    }
    if magic != OLE_MAGIC {
        return Ok(false);
    }

    file.seek(SeekFrom::Start(0))
        .map_err(|e| format!("failed to rewind {}: {e}", path.display()))?;

    let mut ole = match cfb::CompoundFile::open(file) {
        Ok(ole) => ole,
        Err(_) => return Ok(false),
    };

    Ok(stream_exists(&mut ole, "EncryptionInfo") && stream_exists(&mut ole, "EncryptedPackage"))
}

fn stream_exists<R: std::io::Read + std::io::Write + std::io::Seek>(
    ole: &mut cfb::CompoundFile<R>,
    name: &str,
) -> bool {
    if ole.open_stream(name).is_ok() {
        return true;
    }
    let with_leading_slash = format!("/{name}");
    ole.open_stream(&with_leading_slash).is_ok()
}
