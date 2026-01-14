use std::io::{Cursor, Read as _};
use std::path::PathBuf;

use formula_io::{
    open_workbook, open_workbook_model_with_password, open_workbook_with_password,
    save_workbook_with_options, Error, SaveEncryptionScheme, SaveOptions, Workbook,
};
use formula_model::CellValue;

fn fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../fixtures").join(rel)
}

const OLE_MAGIC: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

fn encryption_info_version(bytes: &[u8]) -> Option<(u16, u16)> {
    let cursor = Cursor::new(bytes);
    let mut ole = cfb::CompoundFile::open(cursor).ok()?;
    let mut stream = ole.open_stream("EncryptionInfo").ok()?;
    let mut header = [0u8; 4];
    stream.read_exact(&mut header).ok()?;
    let major = u16::from_le_bytes([header[0], header[1]]);
    let minor = u16::from_le_bytes([header[2], header[3]]);
    Some((major, minor))
}

#[test]
fn saves_password_protected_xlsx_and_reopens_with_password() {
    let tmp = tempfile::tempdir().expect("temp dir");
    let password = "Password1234_";
    let src = fixture_path("xlsx/basic/basic.xlsx");
    let wb = open_workbook(&src).expect("open plaintext fixture");

    let out_path = tmp.path().join("protected.xlsx");
    let res = save_workbook_with_options(
        &wb,
        &out_path,
        SaveOptions {
            password: Some(password.to_string()),
            ..Default::default()
        },
    );

    if !cfg!(feature = "encrypted-workbooks") {
        let err = res.expect_err("expected save to fail without encrypted-workbooks support");
        assert!(
            matches!(err, Error::UnsupportedEncryption { .. }),
            "expected Error::UnsupportedEncryption, got {err:?}"
        );
        return;
    }
    res.expect("save password-protected workbook");

    // Encrypted OOXML files are OLE compound files, not ZIP files.
    let bytes = std::fs::read(&out_path).expect("read saved workbook");
    assert!(
        bytes.starts_with(&OLE_MAGIC),
        "expected encrypted output to be an OLE compound file"
    );
    assert_eq!(
        encryption_info_version(&bytes),
        Some((4, 4)),
        "expected default encrypted save to use Agile (4.4) EncryptionInfo"
    );

    if cfg!(feature = "encrypted-workbooks") {
        let model = open_workbook_model_with_password(&out_path, Some(password))
            .expect("re-open workbook with password");
        let sheet = model.sheet_by_name("Sheet1").expect("Sheet1 missing");
        assert_eq!(sheet.value_a1("A1").unwrap(), CellValue::Number(1.0));
        assert_eq!(
            sheet.value_a1("B1").unwrap(),
            CellValue::String("Hello".to_string())
        );
    } else {
        let err = open_workbook_model_with_password(&out_path, Some(password))
            .expect_err("expected open to fail without encrypted-workbooks support");
        assert!(
            matches!(err, Error::UnsupportedEncryption { .. }),
            "expected Error::UnsupportedEncryption, got {err:?}"
        );
    }
}

#[test]
fn opening_with_wrong_password_fails_cleanly() {
    let tmp = tempfile::tempdir().expect("temp dir");
    let src = fixture_path("xlsx/basic/basic.xlsx");
    let wb = open_workbook(&src).expect("open plaintext fixture");

    let out_path = tmp.path().join("protected.xlsx");
    let res = save_workbook_with_options(
        &wb,
        &out_path,
        SaveOptions {
            password: Some("CorrectPassword".to_string()),
            ..Default::default()
        },
    );

    if !cfg!(feature = "encrypted-workbooks") {
        let err = res.expect_err("expected save to fail without encrypted-workbooks support");
        assert!(
            matches!(err, Error::UnsupportedEncryption { .. }),
            "expected Error::UnsupportedEncryption, got {err:?}"
        );
        return;
    }
    res.expect("save password-protected workbook");

    let err = open_workbook_model_with_password(&out_path, Some("WrongPassword"))
        .expect_err("expected wrong password to fail");

    if cfg!(feature = "encrypted-workbooks") {
        assert!(
            matches!(err, Error::InvalidPassword { .. }),
            "expected Error::InvalidPassword, got {err:?}"
        );
    } else {
        assert!(
            matches!(err, Error::UnsupportedEncryption { .. }),
            "expected Error::UnsupportedEncryption, got {err:?}"
        );
    }
}

#[test]
fn saving_legacy_xls_is_still_unsupported_even_with_password() {
    let src = fixture_path("xlsx/basic/basic.xlsx");
    let wb = open_workbook(&src).expect("open fixture");

    let tmp = tempfile::tempdir().expect("temp dir");
    let out_path = tmp.path().join("out.xls");
    let err = save_workbook_with_options(
        &wb,
        &out_path,
        SaveOptions {
            password: Some("pw".to_string()),
            ..Default::default()
        },
    )
    .expect_err("expected `.xls` output to be unsupported");

    assert!(
        matches!(err, Error::UnsupportedExtension { .. }),
        "expected Error::UnsupportedExtension, got {err:?}"
    );
}

#[test]
fn saves_password_protected_xlsx_with_standard_encryption() {
    let tmp = tempfile::tempdir().expect("temp dir");
    let password = "Password1234_";
    let src = fixture_path("xlsx/basic/basic.xlsx");
    let wb = open_workbook(&src).expect("open plaintext fixture");

    let out_path = tmp.path().join("protected-standard.xlsx");
    let res = save_workbook_with_options(
        &wb,
        &out_path,
        SaveOptions {
            password: Some(password.to_string()),
            encryption_scheme: SaveEncryptionScheme::Standard,
        },
    );

    if !cfg!(feature = "encrypted-workbooks") {
        let err = res.expect_err("expected save to fail without encrypted-workbooks support");
        assert!(
            matches!(err, Error::UnsupportedEncryption { .. }),
            "expected Error::UnsupportedEncryption, got {err:?}"
        );
        return;
    }
    res.expect("save password-protected workbook (Standard)");

    let bytes = std::fs::read(&out_path).expect("read saved workbook");
    assert!(
        bytes.starts_with(&OLE_MAGIC),
        "expected encrypted output to be an OLE compound file"
    );
    assert_eq!(
        encryption_info_version(&bytes),
        Some((3, 2)),
        "expected Standard encrypted save to use EncryptionInfo version 3.2"
    );

    let model = open_workbook_model_with_password(&out_path, Some(password))
        .expect("re-open workbook with password");
    let sheet = model.sheet_by_name("Sheet1").expect("Sheet1 missing");
    assert_eq!(sheet.value_a1("A1").unwrap(), CellValue::Number(1.0));
    assert_eq!(
        sheet.value_a1("B1").unwrap(),
        CellValue::String("Hello".to_string())
    );
}

#[test]
fn saves_password_protected_xlsm_preserves_vba_part() {
    let tmp = tempfile::tempdir().expect("temp dir");
    let password = "Password1234_";
    let src = fixture_path("xlsx/macros/basic.xlsm");
    let wb = open_workbook(&src).expect("open plaintext fixture");

    let out_path = tmp.path().join("protected.xlsm");
    let res = save_workbook_with_options(
        &wb,
        &out_path,
        SaveOptions {
            password: Some(password.to_string()),
            encryption_scheme: SaveEncryptionScheme::Standard,
        },
    );

    if !cfg!(feature = "encrypted-workbooks") {
        let err = res.expect_err("expected save to fail without encrypted-workbooks support");
        assert!(
            matches!(err, Error::UnsupportedEncryption { .. }),
            "expected Error::UnsupportedEncryption, got {err:?}"
        );
        return;
    }
    res.expect("save password-protected workbook (XLSM)");

    let bytes = std::fs::read(&out_path).expect("read saved workbook");
    assert!(
        bytes.starts_with(&OLE_MAGIC),
        "expected encrypted output to be an OLE compound file"
    );
    assert_eq!(
        encryption_info_version(&bytes),
        Some((3, 2)),
        "expected Standard encrypted save to use EncryptionInfo version 3.2"
    );

    let reopened = open_workbook_with_password(&out_path, Some(password))
        .expect("re-open encrypted workbook with password");
    match reopened {
        Workbook::Xlsx(package) => {
            let presence = package.macro_presence();
            assert!(
                presence.has_vba,
                "expected encrypted `.xlsm` output to preserve xl/vbaProject.bin"
            );
        }
        other => panic!("expected Workbook::Xlsx, got {other:?}"),
    }
}

#[test]
fn saves_password_protected_xlsb_and_reopens_with_password() {
    let tmp = tempfile::tempdir().expect("temp dir");
    let password = "Password1234_";
    let src = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../crates/formula-xlsb/tests/fixtures/simple.xlsb");
    let wb = open_workbook(&src).expect("open plaintext XLSB fixture");

    let out_path = tmp.path().join("protected.xlsb");
    let res = save_workbook_with_options(
        &wb,
        &out_path,
        SaveOptions {
            password: Some(password.to_string()),
            encryption_scheme: SaveEncryptionScheme::Standard,
        },
    );

    if !cfg!(feature = "encrypted-workbooks") {
        let err = res.expect_err("expected save to fail without encrypted-workbooks support");
        assert!(
            matches!(err, Error::UnsupportedEncryption { .. }),
            "expected Error::UnsupportedEncryption, got {err:?}"
        );
        return;
    }
    res.expect("save password-protected workbook (XLSB)");

    let bytes = std::fs::read(&out_path).expect("read saved workbook");
    assert!(
        bytes.starts_with(&OLE_MAGIC),
        "expected encrypted output to be an OLE compound file"
    );
    assert_eq!(
        encryption_info_version(&bytes),
        Some((3, 2)),
        "expected Standard encrypted save to use EncryptionInfo version 3.2"
    );

    let reopened = open_workbook_with_password(&out_path, Some(password))
        .expect("re-open encrypted workbook with password");
    assert!(
        matches!(reopened, Workbook::Xlsb(_)),
        "expected Workbook::Xlsb, got {reopened:?}"
    );

    let model = open_workbook_model_with_password(&out_path, Some(password))
        .expect("re-open model workbook with password");
    let sheet = model.sheet_by_name("Sheet1").expect("Sheet1 missing");
    assert_eq!(
        sheet.value_a1("A1").unwrap(),
        CellValue::String("Hello".to_string())
    );
    assert_eq!(sheet.value_a1("B1").unwrap(), CellValue::Number(42.5));
}
