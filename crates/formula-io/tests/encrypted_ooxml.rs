use std::io::{Cursor, Write};
use std::path::PathBuf;

use formula_io::{
    detect_workbook_encryption, detect_workbook_format, open_workbook, open_workbook_model,
    open_workbook_model_with_password, open_workbook_with_password, Error, Workbook,
    OoxmlEncryptedPackageScheme, WorkbookEncryption,
};
use formula_model::{CellRef, CellValue};

fn fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../fixtures").join(rel)
}

fn encrypted_ooxml_bytes_with_stream_names_and_encrypted_package(
    encryption_info: &str,
    encrypted_package: &str,
    encrypted_package_bytes: &[u8],
) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole
            .create_stream(encryption_info)
            .unwrap_or_else(|_| panic!("create {encryption_info} stream"));
        // Minimal EncryptionInfo header:
        // - VersionMajor = 4
        // - VersionMinor = 4 (Agile encryption)
        // - Flags = 0
        stream
            .write_all(&[4, 0, 4, 0, 0, 0, 0, 0])
            .expect("write EncryptionInfo header");
    }
    {
        let mut stream = ole
            .create_stream(encrypted_package)
            .unwrap_or_else(|_| panic!("create {encrypted_package} stream"));
        stream
            .write_all(encrypted_package_bytes)
            .expect("write EncryptedPackage bytes");
    }
    ole.into_inner().into_inner()
}

fn encrypted_ooxml_bytes_with_stream_names(encryption_info: &str, encrypted_package: &str) -> Vec<u8> {
    encrypted_ooxml_bytes_with_stream_names_and_encrypted_package(encryption_info, encrypted_package, &[])
}

fn encrypted_ooxml_bytes_with_encrypted_package(encrypted_package: &[u8]) -> Vec<u8> {
    encrypted_ooxml_bytes_with_stream_names_and_encrypted_package(
        "EncryptionInfo",
        "EncryptedPackage",
        encrypted_package,
    )
}

#[test]
fn detects_encrypted_ooxml_xlsx_container() {
    let tmp = tempfile::tempdir().expect("tempdir");

    for (info_stream, package_stream) in [
        ("EncryptionInfo", "EncryptedPackage"),
        ("encryptioninfo", "encryptedpackage"),
        ("/encryptioninfo", "/encryptedpackage"),
    ] {
        let bytes = encrypted_ooxml_bytes_with_stream_names(info_stream, package_stream);

        // Test both correct and incorrect extensions to ensure content sniffing detects encryption
        // before attempting to open as legacy BIFF.
        for filename in ["encrypted.xlsx", "encrypted.xls", "encrypted.xlsb"] {
            let path = tmp.path().join(filename);
            std::fs::write(&path, &bytes).expect("write encrypted fixture");

            let encryption = detect_workbook_encryption(&path).expect("detect encryption");
            assert!(
                matches!(
                    encryption,
                    WorkbookEncryption::OoxmlEncryptedPackage {
                        scheme: Some(OoxmlEncryptedPackageScheme::Agile)
                    }
                ),
                "expected OOXML EncryptedPackage, got {encryption:?}"
            );

            let err =
                detect_workbook_format(&path).expect_err("expected encrypted workbook to error");
            if cfg!(feature = "encrypted-workbooks") {
                assert!(
                    matches!(err, Error::PasswordRequired { .. }),
                    "expected Error::PasswordRequired, got {err:?}"
                );
            } else {
                assert!(
                    matches!(err, Error::UnsupportedEncryption { .. }),
                    "expected Error::UnsupportedEncryption, got {err:?}"
                );
            }

            let err = open_workbook(&path).expect_err("expected encrypted workbook to error");
            if cfg!(feature = "encrypted-workbooks") {
                assert!(
                    matches!(err, Error::PasswordRequired { .. }),
                    "expected Error::PasswordRequired, got {err:?}"
                );
            } else {
                assert!(
                    matches!(err, Error::UnsupportedEncryption { .. }),
                    "expected Error::UnsupportedEncryption, got {err:?}"
                );
            }
            let msg = err.to_string().to_lowercase();
            if cfg!(feature = "encrypted-workbooks") {
                assert!(
                    msg.contains("password") || msg.contains("encrypt"),
                    "expected error message to mention encryption/password protection, got: {msg}"
                );
            } else {
                assert!(
                    msg.contains("unsupported") || msg.contains("not supported"),
                    "expected error message to mention that encryption is unsupported, got: {msg}"
                );
            }

            let err = open_workbook_model(&path).expect_err("expected encrypted workbook to error");
            if cfg!(feature = "encrypted-workbooks") {
                assert!(
                    matches!(err, Error::PasswordRequired { .. }),
                    "expected Error::PasswordRequired, got {err:?}"
                );
            } else {
                assert!(
                    matches!(err, Error::UnsupportedEncryption { .. }),
                    "expected Error::UnsupportedEncryption, got {err:?}"
                );
            }

            // Providing a password should attempt to decrypt and surface a distinct error from
            // `PasswordRequired`. Because this fixture is intentionally malformed (it does not
            // contain a valid `EncryptedPackage` payload), different parsing/decryption paths may
            // surface either `InvalidPassword` or an "unsupported/malformed encryption container"
            // error.
            let err = open_workbook_with_password(&path, Some("wrong"))
                .expect_err("expected password-protected open to error");
            if cfg!(feature = "encrypted-workbooks") {
                assert!(
                    matches!(
                        err,
                        Error::InvalidPassword { .. }
                            | Error::UnsupportedOoxmlEncryption { .. }
                            | Error::DecryptOoxml { .. }
                    ),
                    "expected InvalidPassword, UnsupportedOoxmlEncryption, or DecryptOoxml, got {err:?}"
                );
            } else {
                assert!(
                    matches!(err, Error::UnsupportedEncryption { .. }),
                    "expected Error::UnsupportedEncryption, got {err:?}"
                );
            }

            let err = open_workbook_model_with_password(&path, Some("wrong"))
                .expect_err("expected password-protected open to error");
            if cfg!(feature = "encrypted-workbooks") {
                assert!(
                    matches!(
                        err,
                        Error::InvalidPassword { .. }
                            | Error::UnsupportedOoxmlEncryption { .. }
                            | Error::DecryptOoxml { .. }
                    ),
                    "expected InvalidPassword, UnsupportedOoxmlEncryption, or DecryptOoxml, got {err:?}"
                );
            } else {
                assert!(
                    matches!(err, Error::UnsupportedEncryption { .. }),
                    "expected Error::UnsupportedEncryption, got {err:?}"
                );
            }
        }
    }
}

#[test]
fn detects_encrypted_ooxml_xlsx_container_for_model_loader() {
    let tmp = tempfile::tempdir().expect("tempdir");

    for (info_stream, package_stream) in [
        ("EncryptionInfo", "EncryptedPackage"),
        ("encryptioninfo", "encryptedpackage"),
        ("/encryptioninfo", "/encryptedpackage"),
    ] {
        let bytes = encrypted_ooxml_bytes_with_stream_names(info_stream, package_stream);

        for filename in ["encrypted.xlsx", "encrypted.xls"] {
            let path = tmp.path().join(filename);
            std::fs::write(&path, &bytes).expect("write encrypted fixture");

            let err =
                open_workbook_model(&path).expect_err("expected encrypted workbook to error");
            if cfg!(feature = "encrypted-workbooks") {
                assert!(
                    matches!(err, Error::PasswordRequired { .. }),
                    "expected Error::PasswordRequired, got {err:?}"
                );
            } else {
                assert!(
                    matches!(err, Error::UnsupportedEncryption { .. }),
                    "expected Error::UnsupportedEncryption, got {err:?}"
                );
            }
        }
    }
}

#[test]
fn encrypted_ooxml_fixtures_require_password() {
    for rel in [
        "encrypted/ooxml/agile.xlsx",
        "encrypted/ooxml/standard.xlsx",
        "encrypted/ooxml/standard-4.2.xlsx",
    ] {
        let path = fixture_path(rel);

        let err = open_workbook(&path).expect_err("expected encrypted workbook to error");
        let msg = err.to_string().to_lowercase();
        assert!(
            msg.contains("encrypt") || msg.contains("password"),
            "expected error message to mention encryption/password protection, got: {msg}"
        );

        let err = open_workbook_model(&path).expect_err("expected encrypted workbook to error");
        let msg = err.to_string().to_lowercase();
        assert!(
            msg.contains("encrypt") || msg.contains("password"),
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

#[test]
fn encrypted_ooxml_plaintext_xlsb_payload_opens() {
    // Wrap a real `.xlsb` OPC/ZIP payload in a synthetic OLE `EncryptedPackage` container (where the
    // payload is already plaintext). This exercises the "already-decrypted" pipeline path.
    let zip_bytes = std::fs::read(PathBuf::from(env!("CARGO_MANIFEST_DIR")).join(
        "../formula-xlsb/tests/fixtures/simple.xlsb",
    ))
    .expect("read xlsb fixture");

    let bytes = encrypted_ooxml_bytes_with_encrypted_package(&zip_bytes);

    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("encrypted.xlsb");
    std::fs::write(&path, &bytes).expect("write encrypted fixture");

    let wb = open_workbook_with_password(&path, Some("dummy")).expect("open xlsb via password API");
    assert!(
        matches!(wb, Workbook::Xlsb(_)),
        "expected Workbook::Xlsb, got {wb:?}"
    );

    let model =
        open_workbook_model_with_password(&path, Some("dummy")).expect("open xlsb model workbook");
    let sheet = model.sheet_by_name("Sheet1").expect("Sheet1 missing");
    assert_eq!(
        sheet.value(CellRef::from_a1("A1").unwrap()),
        CellValue::String("Hello".to_string())
    );
}

#[test]
fn encrypted_package_size_prefix_starting_with_pk_is_not_treated_as_plaintext_zip() {
    // Regression test: `EncryptedPackage` always begins with a u64le plaintext size prefix.
    // If those low bytes happen to be `PK` (0x4b50), we must not misclassify the stream as a
    // plaintext ZIP payload and attempt to open it directly.
    //
    // `0x0000_0000_0000_4b50` => little-endian prefix bytes: `PK\0\0...`.
    let mut encrypted_package = Vec::new();
    encrypted_package.extend_from_slice(&0x4b50u64.to_le_bytes());
    encrypted_package.extend_from_slice(&[0u8; 16]);

    let bytes = encrypted_ooxml_bytes_with_encrypted_package(&encrypted_package);

    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("pkprefix.xlsx");
    std::fs::write(&path, &bytes).expect("write encrypted fixture");

    let err = open_workbook_with_password(&path, Some("password")).expect_err("expected error");
    assert!(
        !matches!(err, Error::OpenXlsx { .. }),
        "expected encryption-related error (not OpenXlsx from misclassified ZIP), got {err:?}"
    );

    if cfg!(feature = "encrypted-workbooks") {
        assert!(
            matches!(
                err,
                Error::InvalidPassword { .. }
                    | Error::UnsupportedOoxmlEncryption { .. }
                    | Error::DecryptOoxml { .. }
            ),
            "expected InvalidPassword, UnsupportedOoxmlEncryption, or DecryptOoxml, got {err:?}"
        );
    } else {
        assert!(
            matches!(err, Error::UnsupportedEncryption { .. }),
            "expected UnsupportedEncryption when encrypted-workbooks feature is disabled, got {err:?}"
        );
    }
}

#[test]
fn opens_real_encrypted_ooxml_fixtures_with_password() {
    let fixtures = [
        ("encryption/encrypted_agile.xlsx", "password"),
        ("encryption/encrypted_standard.xlsx", "password"),
        ("encryption/encrypted_agile_unicode.xlsx", "pässwörd"),
        ("encryption/encrypted_agile.xlsm", "password"),
    ];

    for (rel, password) in fixtures {
        let path = fixture_path(rel);

        let err = open_workbook(&path).expect_err("expected encrypted workbook to error");
        if cfg!(feature = "encrypted-workbooks") {
            assert!(
                matches!(err, Error::PasswordRequired { .. }),
                "expected Error::PasswordRequired, got {err:?}"
            );
        }

        if cfg!(feature = "encrypted-workbooks") {
            let wb = open_workbook_with_password(&path, Some(password)).expect("open encrypted workbook");
            // For OOXML workbooks, `formula-io` opens both `.xlsx` and `.xlsm` as an `Xlsx` package.
            assert!(
                matches!(wb, Workbook::Xlsx(_)),
                "expected Workbook::Xlsx(..), got {wb:?}"
            );

            let model = open_workbook_model_with_password(&path, Some(password))
                .expect("open encrypted workbook model");
            assert!(
                !model.sheets.is_empty(),
                "expected decrypted model workbook to contain at least one sheet"
            );
        }
    }
}
