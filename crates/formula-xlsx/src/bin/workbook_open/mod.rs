//! Helper for developer CLI binaries to open `.xlsx/.xlsm` files that may be wrapped in
//! Office encryption (OLE `EncryptionInfo` + `EncryptedPackage` streams).
//!
//! These CLIs are intended for maintainers and fixture triage, so we keep the API minimal
//! and rely on `office-crypto` for MS-OFFCRYPTO decryption.

#![cfg(not(target_arch = "wasm32"))]

use std::error::Error;
use std::io::Cursor;
use std::path::Path;

use formula_xlsx::XlsxPackage;

const OLE_MAGIC: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

pub fn open_xlsx_package(path: &Path, password: Option<&str>) -> Result<XlsxPackage, Box<dyn Error>> {
    let bytes = std::fs::read(path)?;

    let bytes = if is_encrypted_ooxml_ole(&bytes) {
        let Some(password) = password else {
            return Err(format!(
                "password required: workbook `{}` is Office-encrypted; pass --password <pw>",
                path.display()
            )
            .into());
        };
        office_crypto::decrypt_from_bytes(bytes, password).map_err(|err| {
            format!(
                "failed to decrypt workbook `{}` with provided password: {err}",
                path.display()
            )
        })?
    } else {
        bytes
    };

    Ok(XlsxPackage::from_bytes(&bytes)?)
}

fn is_encrypted_ooxml_ole(bytes: &[u8]) -> bool {
    if bytes.len() < OLE_MAGIC.len() {
        return false;
    }
    if bytes[..OLE_MAGIC.len()] != OLE_MAGIC {
        return false;
    }

    let cursor = Cursor::new(bytes);
    let mut ole = match cfb::CompoundFile::open(cursor) {
        Ok(ole) => ole,
        Err(_) => return false,
    };

    cfb_stream_exists(&mut ole, "EncryptionInfo") && cfb_stream_exists(&mut ole, "EncryptedPackage")
}

fn cfb_stream_exists<R: std::io::Read + std::io::Seek>(ole: &mut cfb::CompoundFile<R>, name: &str) -> bool {
    if ole.open_stream(name).is_ok() {
        return true;
    }
    let with_leading_slash = format!("/{name}");
    ole.open_stream(&with_leading_slash).is_ok()
}

