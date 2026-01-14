//! Regression test for `open_workbook_with_options` error mapping.
//!
//! Malformed OOXML encrypted containers should surface as an OOXML-related error (not the legacy
//! `.xls`/FILEPASS `EncryptedWorkbook` error).
#![cfg(all(feature = "encrypted-workbooks", not(target_arch = "wasm32")))]

use std::io::{Cursor, Write as _};

use formula_io::{open_workbook_with_options, Error, OpenOptions};

#[test]
fn maps_invalid_encryption_info_to_unsupported_ooxml_encryption() {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");

    // Minimal Agile (4.4) EncryptionInfo header, but *no* XML payload.
    // Additionally, keep the `EncryptedPackage` stream truncated (size prefix only) so the decrypt
    // path fails fast.
    {
        let mut stream = ole
            .create_stream("EncryptionInfo")
            .expect("create EncryptionInfo stream");
        stream
            .write_all(&[4, 0, 4, 0, 0, 0, 0, 0])
            .expect("write EncryptionInfo header");
    }
    {
        let mut stream = ole
            .create_stream("EncryptedPackage")
            .expect("create EncryptedPackage stream");
        stream
            .write_all(&0u64.to_le_bytes())
            .expect("write plaintext length prefix");
        // Ensure the stream is not trivially short so the password-aware open path reaches the
        // EncryptionInfo parsing logic (instead of failing early on a missing ciphertext payload).
        stream
            .write_all(&[0u8; 16])
            .expect("write dummy ciphertext bytes");
    }

    let bytes = ole.into_inner().into_inner();
    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("invalid-encryption-info.xlsx");
    std::fs::write(&path, &bytes).expect("write fixture to disk");

    let err = open_workbook_with_options(
        &path,
        OpenOptions {
            password: Some("password".to_string()),
        },
    )
    .expect_err("expected invalid EncryptionInfo to error");

    assert!(
        matches!(err, Error::UnsupportedOoxmlEncryption { .. }),
        "expected UnsupportedOoxmlEncryption for malformed EncryptionInfo/EncryptedPackage, got {err:?}"
    );
}
