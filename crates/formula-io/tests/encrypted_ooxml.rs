use std::io::{Cursor, Write};
use std::path::PathBuf;

use formula_io::{
    detect_workbook_encryption, detect_workbook_format, open_workbook, open_workbook_model,
    open_workbook_model_with_password, open_workbook_with_password, Error, WorkbookEncryptionKind,
};

fn fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../fixtures").join(rel)
}

fn encrypted_ooxml_bytes() -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole
            .create_stream("EncryptionInfo")
            .expect("create EncryptionInfo stream");
        // Minimal EncryptionInfo header:
        // - VersionMajor = 4
        // - VersionMinor = 4 (Agile encryption)
        // - Flags = 0
        stream
            .write_all(&[4, 0, 4, 0, 0, 0, 0, 0])
            .expect("write EncryptionInfo header");
    }
    ole.create_stream("EncryptedPackage")
        .expect("create EncryptedPackage stream");
    ole.into_inner().into_inner()
}

#[test]
fn detects_encrypted_ooxml_xlsx_container() {
    let tmp = tempfile::tempdir().expect("tempdir");

    let bytes = encrypted_ooxml_bytes();

    // Test both correct and incorrect extensions to ensure content sniffing detects encryption
    // before attempting to open as legacy BIFF.
    for filename in ["encrypted.xlsx", "encrypted.xls", "encrypted.xlsb"] {
        let path = tmp.path().join(filename);
        std::fs::write(&path, &bytes).expect("write encrypted fixture");

        let info = detect_workbook_encryption(&path)
            .expect("detect encryption")
            .expect("expected encrypted workbook to be detected");
        assert_eq!(info.kind, WorkbookEncryptionKind::OoxmlOleEncryptedPackage);

        let err = detect_workbook_format(&path).expect_err("expected encrypted workbook to error");
        assert!(
            matches!(err, Error::PasswordRequired { .. }),
            "expected Error::PasswordRequired, got {err:?}"
        );

        let err = open_workbook(&path).expect_err("expected encrypted workbook to error");
        assert!(
            matches!(err, Error::PasswordRequired { .. }),
            "expected Error::PasswordRequired, got {err:?}"
        );
        let msg = err.to_string().to_lowercase();
        assert!(
            msg.contains("password") || msg.contains("encrypt"),
            "expected error message to mention encryption/password protection, got: {msg}"
        );

        let err = open_workbook_model(&path).expect_err("expected encrypted workbook to error");
        assert!(
            matches!(err, Error::PasswordRequired { .. }),
            "expected Error::PasswordRequired, got {err:?}"
        );

        // Providing a password should either surface a distinct "invalid password" error (when
        // decryption support is not enabled) or a more specific decryption/compatibility error (when
        // `encrypted-workbooks` is enabled, because this fixture does not contain a valid encrypted
        // payload).
        let err = open_workbook_with_password(&path, Some("wrong"))
            .expect_err("expected password-protected open to error");
        if cfg!(feature = "encrypted-workbooks") {
            assert!(
                matches!(
                    err,
                    Error::UnsupportedOoxmlEncryption {
                        version_major: 4,
                        version_minor: 4,
                        ..
                    }
                ),
                "expected UnsupportedOoxmlEncryption(4,4), got {err:?}"
            );
        } else {
            assert!(
                matches!(err, Error::InvalidPassword { .. }),
                "expected Error::InvalidPassword, got {err:?}"
            );
        }

        let err = open_workbook_model_with_password(&path, Some("wrong"))
            .expect_err("expected password-protected open to error");
        if cfg!(feature = "encrypted-workbooks") {
            assert!(
                matches!(
                    err,
                    Error::UnsupportedOoxmlEncryption {
                        version_major: 4,
                        version_minor: 4,
                        ..
                    }
                ),
                "expected UnsupportedOoxmlEncryption(4,4), got {err:?}"
            );
        } else {
            assert!(
                matches!(err, Error::InvalidPassword { .. }),
                "expected Error::InvalidPassword, got {err:?}"
            );
        }
    }
}

#[test]
fn detects_encrypted_ooxml_xlsx_container_for_model_loader() {
    let tmp = tempfile::tempdir().expect("tempdir");
    let bytes = encrypted_ooxml_bytes();

    for filename in ["encrypted.xlsx", "encrypted.xls"] {
        let path = tmp.path().join(filename);
        std::fs::write(&path, &bytes).expect("write encrypted fixture");

        let err = open_workbook_model(&path).expect_err("expected encrypted workbook to error");
        assert!(
            matches!(err, Error::PasswordRequired { .. }),
            "expected Error::PasswordRequired, got {err:?}"
        );
    }
}

#[test]
fn encrypted_ooxml_fixtures_require_password() {
    for rel in [
        "encrypted/ooxml/agile.xlsx",
        "encrypted/ooxml/agile-empty-password.xlsx",
        "encrypted/ooxml/standard.xlsx",
    ] {
        let path = fixture_path(rel);

        let err = open_workbook(&path).expect_err("expected encrypted workbook to error");
        let msg = err.to_string().to_lowercase();
        assert!(
            msg.contains("password") || msg.contains("encrypt"),
            "expected error message to mention encryption/password protection, got: {msg}"
        );

        let err = open_workbook_model(&path).expect_err("expected encrypted workbook to error");
        let msg = err.to_string().to_lowercase();
        assert!(
            msg.contains("password") || msg.contains("encrypt"),
            "expected error message to mention encryption/password protection, got: {msg}"
        );
    }
}

#[test]
fn errors_on_unsupported_encryption_version() {
    let tmp = tempfile::tempdir().expect("tempdir");

    // Same OLE container structure, but with an unsupported EncryptionInfo version.
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole
            .create_stream("EncryptionInfo")
            .expect("create EncryptionInfo stream");
        // VersionMajor = 9, VersionMinor = 9 (nonsense, but exercises error reporting).
        stream
            .write_all(&[9, 0, 9, 0, 0, 0, 0, 0])
            .expect("write EncryptionInfo header");
    }
    ole.create_stream("EncryptedPackage")
        .expect("create EncryptedPackage stream");
    let bytes = ole.into_inner().into_inner();

    let path = tmp.path().join("unsupported.xlsx");
    std::fs::write(&path, &bytes).expect("write unsupported encrypted fixture");

    let err = open_workbook(&path).expect_err("expected unsupported encryption to error");
    assert!(
        matches!(
            err,
            Error::UnsupportedOoxmlEncryption {
                version_major: 9,
                version_minor: 9,
                ..
            }
        ),
        "expected Error::UnsupportedOoxmlEncryption(9,9), got {err:?}"
    );
    let msg = err.to_string();
    assert!(
        msg.contains("9.9"),
        "expected error message to include encryption version, got: {msg}"
    );
}
