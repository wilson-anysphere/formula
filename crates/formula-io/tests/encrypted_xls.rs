use std::io::{Cursor, Read, Write};
use std::path::PathBuf;

use formula_io::{
    detect_workbook_encryption, detect_workbook_format, open_workbook, open_workbook_model, Error,
    open_workbook_model_with_options, open_workbook_with_options, open_workbook_with_password,
    OpenOptions, Workbook, WorkbookEncryption,
};

use formula_model::{CellRef, CellValue};

const PASSWORD: &str = "correct horse battery staple";

fn fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures")
        .join(rel)
}

fn record(record_id: u16, payload: &[u8]) -> Vec<u8> {
    let mut out = Vec::with_capacity(4 + payload.len());
    out.extend_from_slice(&record_id.to_le_bytes());
    out.extend_from_slice(&(payload.len() as u16).to_le_bytes());
    out.extend_from_slice(payload);
    out
}

fn encrypted_biff_xls_bytes_filepass() -> Vec<u8> {
    // Minimal BIFF stream:
    // - BOF (BIFF8) with dummy payload
    // - FILEPASS (0x002F) indicates workbook encryption/password protection
    // - EOF
    const RECORD_BOF_BIFF8: u16 = 0x0809;
    const RECORD_FILEPASS: u16 = 0x002F;
    const RECORD_EOF: u16 = 0x000A;

    let workbook_stream = [
        record(RECORD_BOF_BIFF8, &[0u8; 16]),
        record(RECORD_FILEPASS, &[]),
        record(RECORD_EOF, &[]),
    ]
    .concat();

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream bytes");
    }
    ole.into_inner().into_inner()
}

#[test]
fn errors_on_encrypted_xls_filepass() {
    let tmp = tempfile::tempdir().expect("tempdir");
    let bytes = encrypted_biff_xls_bytes_filepass();

    // Test both correct and incorrect extensions to ensure content sniffing routes through the
    // `.xls` importer and surfaces an encryption error.
    for filename in ["encrypted.xls", "encrypted.xlsx", "encrypted.xlsb"] {
        let path = tmp.path().join(filename);
        std::fs::write(&path, &bytes).expect("write encrypted xls fixture");

        let encryption = detect_workbook_encryption(&path).expect("detect encryption");
        assert!(
            matches!(
                encryption,
                WorkbookEncryption::LegacyXlsFilePass { scheme: None }
            ),
            "expected LegacyXlsFilePass, got {encryption:?}"
        );

        let err = detect_workbook_format(&path).expect_err("expected encrypted workbook to error");
        assert!(
            matches!(err, Error::EncryptedWorkbook { .. }),
            "expected Error::EncryptedWorkbook, got {err:?}"
        );

        let err = open_workbook(&path).expect_err("expected encrypted workbook to error");
        assert!(
            matches!(err, Error::EncryptedWorkbook { .. }),
            "expected Error::EncryptedWorkbook, got {err:?}"
        );
        let msg = err.to_string().to_lowercase();
        assert!(
            msg.contains("encrypted") || msg.contains("password"),
            "expected error message to mention encryption/password protection, got: {msg}"
        );

        let err = open_workbook_model(&path).expect_err("expected encrypted workbook to error");
        assert!(
            matches!(err, Error::EncryptedWorkbook { .. }),
            "expected Error::EncryptedWorkbook, got {err:?}"
        );
    }
}

fn encrypted_rc4_cryptoapi_fixture_path() -> std::path::PathBuf {
    std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../formula-xls/tests/fixtures/encrypted")
        .join("biff8_rc4_cryptoapi_pw_open.xls")
}

#[test]
fn opens_encrypted_xls_with_options_password() {
    let tmp = tempfile::tempdir().expect("tempdir");
    let bytes =
        std::fs::read(encrypted_rc4_cryptoapi_fixture_path()).expect("read encrypted xls fixture");

    let path = tmp.path().join("encrypted.xls");
    std::fs::write(&path, &bytes).expect("write encrypted xls fixture");

    // No password: options-based open APIs should surface PasswordRequired for encrypted legacy `.xls`.
    let err = open_workbook_model_with_options(&path, OpenOptions { password: None })
        .expect_err("expected password required");
    assert!(
        matches!(err, Error::PasswordRequired { .. }),
        "expected Error::PasswordRequired, got {err:?}"
    );
    let err = open_workbook_with_options(&path, OpenOptions { password: None })
        .expect_err("expected password required");
    assert!(
        matches!(err, Error::PasswordRequired { .. }),
        "expected Error::PasswordRequired, got {err:?}"
    );

    // Correct password: both model and package loaders succeed.
    let model = open_workbook_model_with_options(
        &path,
        OpenOptions {
            password: Some(PASSWORD.to_string()),
        },
    )
    .expect("open encrypted xls as model");

    let sheet1 = model.sheet_by_name("Sheet1").expect("Sheet1 missing");
    assert_eq!(
        sheet1.value(CellRef::from_a1("A1").unwrap()),
        CellValue::Number(42.0)
    );

    let wb = open_workbook_with_options(
        &path,
        OpenOptions {
            password: Some(PASSWORD.to_string()),
        },
    )
    .expect("open encrypted xls");
    let Workbook::Xls(xls_res) = wb else {
        panic!("expected Workbook::Xls for encrypted xls file");
    };
    let sheet1 = xls_res.workbook.sheet_by_name("Sheet1").expect("Sheet1 missing");
    assert_eq!(
        sheet1.value(CellRef::from_a1("A1").unwrap()),
        CellValue::Number(42.0)
    );

    // Wrong password: deterministic error variant.
    let err = open_workbook_model_with_options(
        &path,
        OpenOptions {
            password: Some("wrong".to_string()),
        },
    )
    .expect_err("expected invalid password to error");
    assert!(
        matches!(err, Error::InvalidPassword { .. }),
        "expected Error::InvalidPassword, got {err:?}"
    );

    let err = open_workbook_with_options(
        &path,
        OpenOptions {
            password: Some("wrong".to_string()),
        },
    )
    .expect_err("expected invalid password to error");
    assert!(
        matches!(err, Error::InvalidPassword { .. }),
        "expected Error::InvalidPassword, got {err:?}"
    );
}

fn read_workbook_stream_from_xls_bytes(data: &[u8]) -> Vec<u8> {
    let cursor = Cursor::new(data.to_vec());
    let mut ole = cfb::CompoundFile::open(cursor).expect("open xls cfb");

    for candidate in ["/Workbook", "/Book", "Workbook", "Book"] {
        if let Ok(mut stream) = ole.open_stream(candidate) {
            let mut buf = Vec::new();
            stream.read_to_end(&mut buf).expect("read workbook stream");
            return buf;
        }
    }

    panic!("fixture missing Workbook/Book stream");
}

fn build_xls_from_workbook_stream(workbook_stream: &[u8]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut ole =
        cfb::CompoundFile::create_with_version(cfb::Version::V3, cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

fn patch_filepass_cryptoapi_alg_id(workbook_stream: &mut [u8], new_alg_id: u32) {
    const RECORD_FILEPASS: u16 = 0x002F;

    let mut offset = 0usize;
    while offset + 4 <= workbook_stream.len() {
        let record_id = u16::from_le_bytes([workbook_stream[offset], workbook_stream[offset + 1]]);
        let len =
            u16::from_le_bytes([workbook_stream[offset + 2], workbook_stream[offset + 3]]) as usize;
        let data_start = offset + 4;
        let data_end = data_start + len;
        assert!(
            data_end <= workbook_stream.len(),
            "truncated record while scanning"
        );

        if record_id == RECORD_FILEPASS {
            // FILEPASS payload:
            //   u16 wEncryptionType
            //   u16 wEncryptionSubType
            //   u32 dwEncryptionInfoLen
            //   EncryptionInfo bytes...
            // EncryptionInfo:
            //   u16 MajorVersion
            //   u16 MinorVersion
            //   u32 Flags
            //   u32 HeaderSize
            //   EncryptionHeader (HeaderSize bytes) where AlgID lives at offset 8.
            let payload = workbook_stream
                .get_mut(data_start..data_end)
                .expect("FILEPASS payload in range");
            assert!(
                payload.len() >= 32,
                "expected FILEPASS payload to contain CryptoAPI EncryptionInfo"
            );

            let header_size =
                u32::from_le_bytes([payload[16], payload[17], payload[18], payload[19]]) as usize;
            let header_start = 20usize;
            assert!(
                header_start + header_size <= payload.len(),
                "EncryptionHeader out of range (header_size={header_size}, payload_len={})",
                payload.len()
            );

            let alg_id_off = header_start + 8;
            payload[alg_id_off..alg_id_off + 4].copy_from_slice(&new_alg_id.to_le_bytes());
            return;
        }

        offset = data_end;
    }

    panic!("FILEPASS record not found");
}

#[test]
fn errors_on_unsupported_encryption_for_encrypted_xls() {
    // Patch the fixture FILEPASS header to claim AES-128 instead of RC4.
    const CALG_AES_128: u32 = 0x0000_660E;

    let tmp = tempfile::tempdir().expect("tempdir");
    let bytes =
        std::fs::read(encrypted_rc4_cryptoapi_fixture_path()).expect("read encrypted xls fixture");
    let mut workbook_stream = read_workbook_stream_from_xls_bytes(&bytes);
    patch_filepass_cryptoapi_alg_id(&mut workbook_stream, CALG_AES_128);
    let patched_xls = build_xls_from_workbook_stream(&workbook_stream);

    let path = tmp.path().join("unsupported.xls");
    std::fs::write(&path, &patched_xls).expect("write xls bytes");

    let err = open_workbook_model_with_options(
        &path,
        OpenOptions {
            password: Some(PASSWORD.to_string()),
        },
    )
    .expect_err("expected unsupported encryption error");
    assert!(
        matches!(err, Error::UnsupportedEncryption { .. }),
        "expected Error::UnsupportedEncryption, got {err:?}"
    );
    let msg = err.to_string().to_lowercase();
    assert!(
        msg.contains("unsupported"),
        "expected unsupported-encryption error message to mention unsupported encryption, got: {msg}"
    );
    assert!(
        !msg.contains("wrong password") && !msg.contains("invalid password"),
        "unsupported-encryption error message should not imply the password is wrong, got: {msg}"
    );

    let err = open_workbook_with_options(
        &path,
        OpenOptions {
            password: Some(PASSWORD.to_string()),
        },
    )
    .expect_err("expected unsupported encryption error");
    assert!(
        matches!(err, Error::UnsupportedEncryption { .. }),
        "expected Error::UnsupportedEncryption, got {err:?}"
    );
    let msg = err.to_string().to_lowercase();
    assert!(
        msg.contains("unsupported"),
        "expected unsupported-encryption error message to mention unsupported encryption, got: {msg}"
    );
    assert!(
        !msg.contains("wrong password") && !msg.contains("invalid password"),
        "unsupported-encryption error message should not imply the password is wrong, got: {msg}"
    );
}

#[test]
fn opens_real_encrypted_xls_fixtures_with_password() {
    let fixtures = [
        ("encryption/encrypted_rc4_cryptoapi.xls", "password"),
        ("encryption/encrypted_rc4_cryptoapi_unicode.xls", "pässwörd"),
    ];

    for (rel, password) in fixtures {
        let path = fixture_path(rel);

        // The non-password open API should return a generic encrypted-workbook error.
        let err = open_workbook(&path).expect_err("expected encrypted workbook to error");
        assert!(
            matches!(err, Error::EncryptedWorkbook { .. }),
            "expected Error::EncryptedWorkbook, got {err:?}"
        );

        // The password-aware API should surface a password-required error when no password is
        // provided.
        let err =
            open_workbook_with_password(&path, None).expect_err("expected password required error");
        assert!(
            matches!(err, Error::PasswordRequired { .. }),
            "expected Error::PasswordRequired, got {err:?}"
        );

        // With the correct password, the workbook should open.
        let wb = open_workbook_with_password(&path, Some(password))
            .expect("open encrypted xls with password");
        assert!(
            matches!(wb, Workbook::Xls(_)),
            "expected Workbook::Xls(..), got {wb:?}"
        );
    }
}

