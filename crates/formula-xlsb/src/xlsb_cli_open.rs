use std::fs::File;
use std::io::{self, Read, Seek};
use std::path::Path;

use aes::{Aes128, Aes192, Aes256};
use cbc::cipher::{block_padding::NoPadding, BlockDecryptMut, KeyIvInit};
use cbc::Decryptor;
use formula_offcrypto::{
    agile_decrypt_package, agile_secret_key, decrypt_standard_encrypted_package, EncryptionInfo,
    OffcryptoError, StandardEncryptionInfo,
};
use formula_xlsb::XlsbWorkbook;
use sha1::{Digest as _, Sha1};

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
    let file = File::open(path)?;
    let mut ole = cfb::CompoundFile::open(file)?;

    if stream_exists(&mut ole, "EncryptionInfo") && stream_exists(&mut ole, "EncryptedPackage") {
        let password = password.ok_or_else(|| {
            io::Error::new(
                io::ErrorKind::InvalidInput,
                "encrypted workbook requires a password; pass --password <pw>",
            )
        })?;

        let zip_bytes = decrypt_ooxml_encrypted_package(&mut ole, password)?;
        return Ok(XlsbWorkbook::open_from_bytes(&zip_bytes)?);
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

fn stream_exists<R: Read + Seek + std::io::Write>(ole: &mut cfb::CompoundFile<R>, name: &str) -> bool {
    if ole.open_stream(name).is_ok() {
        return true;
    }
    let with_leading_slash = format!("/{name}");
    ole.open_stream(&with_leading_slash).is_ok()
}

fn open_stream<R: Read + Seek + std::io::Write>(
    ole: &mut cfb::CompoundFile<R>,
    name: &str,
) -> Result<cfb::Stream<R>, Box<dyn std::error::Error>> {
    match ole.open_stream(name) {
        Ok(s) => Ok(s),
        Err(_) => {
            let with_leading_slash = format!("/{name}");
            Ok(ole.open_stream(&with_leading_slash)?)
        }
    }
}

fn read_stream<R: Read + Seek + std::io::Write>(
    ole: &mut cfb::CompoundFile<R>,
    name: &str,
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let mut stream = open_stream(ole, name)?;
    let mut bytes = Vec::new();
    stream.read_to_end(&mut bytes)?;
    Ok(bytes)
}

fn decrypt_ooxml_encrypted_package<R: Read + Seek + std::io::Write>(
    ole: &mut cfb::CompoundFile<R>,
    password: &str,
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let encryption_info_bytes = read_stream(ole, "EncryptionInfo")?;
    let encrypted_package_bytes = read_stream(ole, "EncryptedPackage")?;

    let info = formula_offcrypto::parse_encryption_info(&encryption_info_bytes)?;
    match info {
        EncryptionInfo::Standard { header, verifier, .. } => {
            let info = StandardEncryptionInfo { header, verifier };
            let key = formula_offcrypto::standard_derive_key_zeroizing(&info, password)?;
            formula_offcrypto::standard_verify_key(&info, key.as_slice())?;

            // Prefer the spec-friendly AES-ECB decryptor used by `formula-offcrypto`.
            if let Ok(decrypted) =
                decrypt_standard_encrypted_package(key.as_slice(), &encrypted_package_bytes)
            {
                if decrypted.starts_with(b"PK") {
                    return Ok(decrypted);
                }
            }

            // Fall back to the per-segment IV AES-CBC scheme observed in some files/toolchains.
            let salt = info.verifier.salt.clone();
            let decrypted = formula_offcrypto::encrypted_package::decrypt_encrypted_package(
                &encrypted_package_bytes,
                |block: u32, ct: &[u8], pt: &mut [u8]| {
                    pt.copy_from_slice(ct);
                    let iv = standard_segment_iv(&salt, block);
                    aes_cbc_decrypt_in_place(key.as_slice(), &iv, pt)
                },
            )?;
            Ok(decrypted)
        }
        EncryptionInfo::Agile { info, .. } => {
            let secret_key = agile_secret_key(&info, password)?;
            Ok(agile_decrypt_package(
                &info,
                secret_key.as_slice(),
                &encrypted_package_bytes,
            )?)
        }
        EncryptionInfo::Unsupported { version } => Err(Box::new(io::Error::new(
            io::ErrorKind::InvalidInput,
            format!(
                "unsupported EncryptionInfo version {}.{} (supported: Standard *.2 with major=2/3/4, and Agile 4.4)",
                version.major, version.minor
            ),
        ))),
    }
}

fn standard_segment_iv(salt: &[u8], segment_index: u32) -> [u8; 16] {
    let mut hasher = Sha1::new();
    hasher.update(salt);
    hasher.update(segment_index.to_le_bytes());
    let digest = hasher.finalize();

    let mut iv = [0u8; 16];
    iv.copy_from_slice(&digest[..16]);
    iv
}

fn aes_cbc_decrypt_in_place(
    key: &[u8],
    iv: &[u8; 16],
    buf: &mut [u8],
) -> Result<(), OffcryptoError> {
    if buf.len() % 16 != 0 {
        return Err(OffcryptoError::InvalidEncryptionInfo {
            context: "AES-CBC ciphertext length must be a multiple of 16 bytes",
        });
    }

    match key.len() {
        16 => {
            let decryptor = Decryptor::<Aes128>::new_from_slices(key, iv)
                .map_err(|_| OffcryptoError::InvalidKeyLength { len: key.len() })?;
            decryptor
                .decrypt_padded_mut::<NoPadding>(buf)
                .map_err(|_| OffcryptoError::InvalidEncryptionInfo {
                    context: "AES-CBC decrypt failed",
                })?;
        }
        24 => {
            let decryptor = Decryptor::<Aes192>::new_from_slices(key, iv)
                .map_err(|_| OffcryptoError::InvalidKeyLength { len: key.len() })?;
            decryptor
                .decrypt_padded_mut::<NoPadding>(buf)
                .map_err(|_| OffcryptoError::InvalidEncryptionInfo {
                    context: "AES-CBC decrypt failed",
                })?;
        }
        32 => {
            let decryptor = Decryptor::<Aes256>::new_from_slices(key, iv)
                .map_err(|_| OffcryptoError::InvalidKeyLength { len: key.len() })?;
            decryptor
                .decrypt_padded_mut::<NoPadding>(buf)
                .map_err(|_| OffcryptoError::InvalidEncryptionInfo {
                    context: "AES-CBC decrypt failed",
                })?;
        }
        _ => return Err(OffcryptoError::InvalidKeyLength { len: key.len() }),
    }

    Ok(())
}
