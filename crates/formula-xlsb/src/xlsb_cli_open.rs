use std::fs::File;
use std::io::{self, Read};
use std::path::Path;

use formula_office_crypto as office_crypto;
use formula_xlsb::XlsbWorkbook;

/// OLE/CFB file signature.
///
/// See: <https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/>
const OLE_MAGIC: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

/// Open an `.xlsb` workbook, optionally using a password for Office-encrypted OLE wrappers.
pub fn open_xlsb_workbook(
    path: &Path,
    password: Option<&str>,
) -> Result<XlsbWorkbook, Box<dyn std::error::Error>> {
    // Fast path: ZIP-based `.xlsb`.
    if !looks_like_ole_compound_file(path)? {
        return Ok(XlsbWorkbook::open(path)?);
    }

    // OLE-based: could be legacy `.xls` or Office-encrypted OOXML.
    let bytes = std::fs::read(path)?;

    if office_crypto::is_encrypted_ooxml_ole(&bytes) {
        let password = password.ok_or_else(|| {
            io::Error::new(
                io::ErrorKind::InvalidInput,
                "encrypted workbook requires a password; pass --password <pw>",
            )
        })?;

        let zip_bytes = office_crypto::decrypt_encrypted_package_ole(&bytes, password)?;
        return Ok(XlsbWorkbook::open_from_vec(zip_bytes)?);
    }

    // Fall back to the normal ZIP open path so the caller gets a sensible parse error.
    Ok(XlsbWorkbook::open(path)?)
}

fn looks_like_ole_compound_file(path: &Path) -> Result<bool, io::Error> {
    let mut file = File::open(path)?;
    let mut header = [0u8; OLE_MAGIC.len()];
    let n = file.read(&mut header)?;
    Ok(n == OLE_MAGIC.len() && header == OLE_MAGIC)
}
