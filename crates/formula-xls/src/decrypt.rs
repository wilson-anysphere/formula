use thiserror::Error;

/// Errors returned while decrypting password-protected `.xls` BIFF workbooks.
#[derive(Debug, Error, Clone, PartialEq, Eq)]
pub enum DecryptError {
    #[error("unsupported encryption scheme: {0}")]
    UnsupportedEncryption(String),
    #[error("wrong password")]
    WrongPassword,
    #[error("invalid encryption info: {0}")]
    InvalidFormat(String),
}

fn map_biff_decrypt_error(err: crate::biff::encryption::DecryptError) -> DecryptError {
    match err {
        crate::biff::encryption::DecryptError::WrongPassword => DecryptError::WrongPassword,
        crate::biff::encryption::DecryptError::UnsupportedEncryption(scheme) => {
            DecryptError::UnsupportedEncryption(scheme)
        }
        err @ crate::biff::encryption::DecryptError::SizeLimitExceeded { .. } => {
            DecryptError::InvalidFormat(err.to_string())
        }
        crate::biff::encryption::DecryptError::InvalidFilePass(message) => {
            DecryptError::InvalidFormat(message)
        }
        crate::biff::encryption::DecryptError::NoFilePass => {
            DecryptError::InvalidFormat("missing FILEPASS record".to_string())
        }
        crate::biff::encryption::DecryptError::PasswordRequired => DecryptError::WrongPassword,
    }
}

/// Decrypt an in-memory BIFF workbook stream for any supported `FILEPASS` scheme.
///
/// The workbook stream is decrypted **in place** and the `FILEPASS` record id is *masked*
/// (replaced with `0xFFFF`) so downstream BIFF parsers (and `calamine`) treat the stream as
/// plaintext without shifting any record offsets (e.g. `BoundSheet8.lbPlyPos`).
pub(crate) fn decrypt_biff_workbook_stream(
    workbook_stream: &mut Vec<u8>,
    password: &str,
) -> Result<(), DecryptError> {
    crate::biff::encryption::decrypt_workbook_stream(workbook_stream, password)
        .map_err(map_biff_decrypt_error)?;
    crate::biff::records::mask_workbook_globals_filepass_record_id_in_place(workbook_stream);
    Ok(())
}

/// Allocating convenience wrapper around [`decrypt_biff_workbook_stream`].
///
/// This is retained for test/ergonomics: callers that already own the workbook-stream `Vec<u8>`
/// should prefer the in-place API to avoid temporarily doubling memory usage for large `.xls`
/// files.
#[allow(dead_code)]
pub(crate) fn decrypt_biff_workbook_stream_allocating(
    workbook_stream: &[u8],
    password: &str,
) -> Result<Vec<u8>, DecryptError> {
    let mut out = workbook_stream.to_vec();
    decrypt_biff_workbook_stream(&mut out, password)?;
    Ok(out)
}
