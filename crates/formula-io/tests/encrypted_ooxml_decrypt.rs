//! End-to-end decryption tests for Office-encrypted OOXML workbooks (Agile + Standard encryption).
//!
//! These are gated behind the `encrypted-workbooks` feature because decryption is optional.
#![cfg(all(feature = "encrypted-workbooks", not(target_arch = "wasm32")))]

use std::io::{Cursor, Read as _, Write as _};
use std::path::{Path, PathBuf};

use ms_offcrypto_writer::Ecma376AgileWriter;
use rand::{rngs::StdRng, SeedableRng as _};

use formula_io::{
    detect_workbook_format, open_workbook_model, open_workbook_model_with_password,
    open_workbook_with_options, open_workbook_with_password, Error, OpenOptions, Workbook,
    WorkbookFormat,
};
use formula_model::{CellRef, CellValue};

fn build_tiny_xlsx() -> Vec<u8> {
    let mut workbook = formula_model::Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1").expect("add sheet");
    let sheet = workbook.sheet_mut(sheet_id).expect("sheet exists");
    sheet.set_value(CellRef::from_a1("A1").unwrap(), CellValue::Number(1.0));
    sheet.set_value(
        CellRef::from_a1("B1").unwrap(),
        CellValue::String("Hello".to_string()),
    );

    let mut cursor = Cursor::new(Vec::new());
    formula_io::xlsx::write_workbook_to_writer(&workbook, &mut cursor)
        .expect("write xlsx to bytes");
    cursor.into_inner()
}

fn encrypt_zip_with_password(plain_zip: &[u8], password: &str) -> Vec<u8> {
    let mut cursor = Cursor::new(Vec::new());
    let mut rng = StdRng::from_seed([0u8; 32]);
    let mut agile =
        Ecma376AgileWriter::create(&mut rng, password, &mut cursor).expect("create agile");
    agile
        .write_all(plain_zip)
        .expect("write plaintext zip to agile writer");
    agile.finalize().expect("finalize agile writer");
    cursor.into_inner()
}

fn strip_data_integrity_from_encryption_info(encryption_info: &[u8]) -> Vec<u8> {
    assert!(
        encryption_info.len() >= 8,
        "EncryptionInfo must include 8-byte header"
    );
    let header = &encryption_info[..8];
    let xml_bytes = &encryption_info[8..];
    let xml = std::str::from_utf8(xml_bytes).expect("EncryptionInfo XML should be UTF-8");
    let mut stripped = xml.to_string();

    // Remove the first `<dataIntegrity .../>` (self-closing) or `<dataIntegrity>...</dataIntegrity>`
    // element when present. This is only used to synthesize fixtures missing the element.
    if let Some(start) = stripped.find("<dataIntegrity") {
        if let Some(end_rel) = stripped[start..].find("/>") {
            stripped.replace_range(start..start + end_rel + 2, "");
        } else if let Some(end_rel) = stripped[start..].find("</dataIntegrity>") {
            stripped.replace_range(start..start + end_rel + "</dataIntegrity>".len(), "");
        }
    }

    let mut out = Vec::new();
    out.extend_from_slice(header);
    out.extend_from_slice(stripped.as_bytes());
    out
}

fn strip_data_integrity_from_encrypted_cfb(encrypted_cfb: &[u8]) -> Vec<u8> {
    let mut ole = cfb::CompoundFile::open(Cursor::new(encrypted_cfb)).expect("parse cfb");
    let mut encryption_info = Vec::new();
    ole.open_stream("EncryptionInfo")
        .expect("open EncryptionInfo stream")
        .read_to_end(&mut encryption_info)
        .expect("read EncryptionInfo");
    let mut encrypted_package = Vec::new();
    ole.open_stream("EncryptedPackage")
        .expect("open EncryptedPackage stream")
        .read_to_end(&mut encrypted_package)
        .expect("read EncryptedPackage");

    let encryption_info = strip_data_integrity_from_encryption_info(&encryption_info);
    assert!(
        !std::str::from_utf8(&encryption_info[8..])
            .expect("utf-8")
            .contains("dataIntegrity"),
        "expected synthesized EncryptionInfo to omit <dataIntegrity>"
    );

    // Rebuild the OLE container with the modified EncryptionInfo, preserving ciphertext.
    let cursor = Cursor::new(Vec::new());
    let mut out = cfb::CompoundFile::create(cursor).expect("create cfb");
    out.create_stream("EncryptionInfo")
        .expect("create EncryptionInfo stream")
        .write_all(&encryption_info)
        .expect("write EncryptionInfo");
    out.create_stream("EncryptedPackage")
        .expect("create EncryptedPackage stream")
        .write_all(&encrypted_package)
        .expect("write EncryptedPackage");
    out.into_inner().into_inner()
}

fn xlsb_fixture_bytes() -> Vec<u8> {
    let path = Path::new(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../formula-xlsb/tests/fixtures/simple.xlsb"
    ));
    std::fs::read(path).expect("read xlsb fixture bytes")
}

#[test]
fn open_workbook_with_password_decrypts_agile_encrypted_package() {
    let password = "correct horse battery staple";
    let plain_xlsx = build_tiny_xlsx();
    let encrypted_cfb = encrypt_zip_with_password(&plain_xlsx, password);

    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("encrypted.xlsx");
    std::fs::write(&path, &encrypted_cfb).expect("write encrypted file");

    // Missing password => prompt.
    let err = open_workbook_with_password(&path, None).expect_err("expected PasswordRequired");
    assert!(
        matches!(err, Error::PasswordRequired { .. }),
        "expected Error::PasswordRequired, got {err:?}"
    );

    // Wrong password => invalid password.
    let err =
        open_workbook_with_password(&path, Some("wrong-password")).expect_err("expected error");
    assert!(
        matches!(err, Error::InvalidPassword { .. }),
        "expected Error::InvalidPassword, got {err:?}"
    );

    // Correct password => decrypted ZIP is routed into the lazy/streaming OPC package wrapper
    // (`Workbook::Xlsx` / `XlsxLazyPackage`).
    let wb = open_workbook_with_password(&path, Some(password)).expect("open decrypted workbook");
    match wb {
        Workbook::Xlsx(package) => {
            let workbook_xml = package
                .read_part("xl/workbook.xml")
                .expect("read xl/workbook.xml")
                .expect("xl/workbook.xml missing in zip");
            let workbook_xml_str =
                std::str::from_utf8(&workbook_xml).expect("xl/workbook.xml must be valid UTF-8");
            assert!(
                workbook_xml_str.contains("Sheet1"),
                "expected xl/workbook.xml to mention Sheet1, got:\n{workbook_xml_str}"
            );
        }
        other => panic!("expected Workbook::Xlsx, got {other:?}"),
    }
}

#[test]
fn open_workbook_with_password_decrypts_agile_encrypted_package_without_data_integrity() {
    let password = "correct horse battery staple";
    let plain_xlsx = build_tiny_xlsx();
    let encrypted_cfb = encrypt_zip_with_password(&plain_xlsx, password);
    let encrypted_cfb = strip_data_integrity_from_encrypted_cfb(&encrypted_cfb);

    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("encrypted-no-integrity.xlsx");
    std::fs::write(&path, &encrypted_cfb).expect("write encrypted file");

    let wb = open_workbook_with_password(&path, Some(password)).expect("open decrypted workbook");
    match wb {
        Workbook::Xlsx(package) => {
            let workbook_xml = package
                .read_part("xl/workbook.xml")
                .expect("read xl/workbook.xml")
                .expect("xl/workbook.xml missing in zip");
            let workbook_xml_str =
                std::str::from_utf8(&workbook_xml).expect("xl/workbook.xml must be valid UTF-8");
            assert!(
                workbook_xml_str.contains("Sheet1"),
                "expected xl/workbook.xml to mention Sheet1, got:\n{workbook_xml_str}"
            );
        }
        other => panic!("expected Workbook::Xlsx, got {other:?}"),
    }
}

#[test]
fn open_workbook_model_with_password_decrypts_agile_encrypted_xlsx() {
    let password = "password";
    let plain_xlsx = build_tiny_xlsx();
    let encrypted_cfb = encrypt_zip_with_password(&plain_xlsx, password);

    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("encrypted.xlsx");
    std::fs::write(&path, &encrypted_cfb).expect("write encrypted file");

    let model =
        open_workbook_model_with_password(&path, Some(password)).expect("open decrypted model");
    let sheet = model.sheet_by_name("Sheet1").expect("Sheet1 missing");
    assert_eq!(
        sheet.value(CellRef::from_a1("A1").unwrap()),
        CellValue::Number(1.0)
    );
    assert_eq!(
        sheet.value(CellRef::from_a1("B1").unwrap()),
        CellValue::String("Hello".to_string())
    );
}

#[test]
fn open_workbook_with_password_decrypts_agile_encrypted_xlsb() {
    let password = "password";
    let plain_xlsb = xlsb_fixture_bytes();
    let encrypted_cfb = encrypt_zip_with_password(&plain_xlsb, password);

    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("encrypted.xlsb");
    std::fs::write(&path, &encrypted_cfb).expect("write encrypted file");

    let wb = open_workbook_with_password(&path, Some(password)).expect("open decrypted workbook");
    assert!(
        matches!(wb, Workbook::Xlsb(_)),
        "expected Workbook::Xlsb, got {wb:?}"
    );
}

#[test]
fn open_workbook_model_with_password_decrypts_agile_encrypted_xlsb() {
    let password = "password";
    let plain_xlsb = xlsb_fixture_bytes();
    let encrypted_cfb = encrypt_zip_with_password(&plain_xlsb, password);

    let tmp = tempfile::tempdir().expect("tempdir");
    let encrypted_path = tmp.path().join("encrypted.xlsb");
    std::fs::write(&encrypted_path, &encrypted_cfb).expect("write encrypted file");

    let workbook = open_workbook_model_with_password(&encrypted_path, Some(password))
        .expect("open encrypted xlsb as model");
    let sheet = workbook.sheet_by_name("Sheet1").expect("Sheet1 missing");
    assert_eq!(
        sheet.value(CellRef::from_a1("A1").unwrap()),
        CellValue::String("Hello".to_string())
    );
    assert_eq!(
        sheet.value(CellRef::from_a1("B1").unwrap()),
        CellValue::Number(42.5)
    );
    assert_eq!(sheet.formula(CellRef::from_a1("C1").unwrap()), Some("B1*2"));
}

fn fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/encrypted/ooxml")
        .join(rel)
}

fn assert_expected_contents(workbook: &formula_model::Workbook) {
    assert_eq!(workbook.sheets.len(), 1, "expected exactly one sheet");
    assert_eq!(workbook.sheets[0].name, "Sheet1");

    let sheet = workbook.sheet_by_name("Sheet1").expect("Sheet1 missing");
    assert_eq!(
        sheet.value(CellRef::from_a1("A1").unwrap()),
        CellValue::Number(1.0)
    );
    assert_eq!(
        sheet.value(CellRef::from_a1("B1").unwrap()),
        CellValue::String("Hello".to_string())
    );
}

fn assert_expected_excel_contents(workbook: &formula_model::Workbook) {
    assert_eq!(workbook.sheets.len(), 1, "expected exactly one sheet");
    assert_eq!(workbook.sheets[0].name, "Sheet1");

    let sheet = workbook.sheet_by_name("Sheet1").expect("Sheet1 missing");
    assert_eq!(
        sheet.value(CellRef::from_a1("A1").unwrap()),
        CellValue::String("lorem".to_string())
    );
    assert_eq!(
        sheet.value(CellRef::from_a1("B1").unwrap()),
        CellValue::String("ipsum".to_string())
    );
}

fn open_model_with_password(path: &Path, password: &str) -> formula_model::Workbook {
    open_workbook_model_with_password(path, Some(password))
        .unwrap_or_else(|err| panic!("open encrypted workbook {path:?} failed: {err:?}"))
}

fn open_decrypted_package_bytes_with_password(path: &Path, password: &str) -> Vec<u8> {
    let wb = open_workbook_with_password(path, Some(password))
        .unwrap_or_else(|err| panic!("open encrypted workbook {path:?} failed: {err:?}"));
    match wb {
        Workbook::Xlsx(package) => package
            .write_to_bytes()
            .expect("serialize decrypted workbook package to bytes"),
        other => panic!("expected Workbook::Xlsx, got {other:?}"),
    }
}

fn assert_has_vba_project(decrypted: &[u8]) {
    let archive = zip::ZipArchive::new(Cursor::new(decrypted)).expect("open decrypted ZIP");
    let mut found = false;
    for name in archive.file_names() {
        if name.eq_ignore_ascii_case("xl/vbaProject.bin") {
            found = true;
            break;
        }
    }
    assert!(
        found,
        "expected decrypted package to contain xl/vbaProject.bin"
    );
}

fn assert_detects_xlsm(decrypted: &[u8]) {
    let tmp = tempfile::tempdir().expect("temp dir");
    let path = tmp.path().join("book.xlsm");
    std::fs::write(&path, decrypted).expect("write decrypted workbook bytes");
    assert_eq!(
        detect_workbook_format(&path).expect("detect workbook format"),
        WorkbookFormat::Xlsm
    );
}

#[test]
fn decrypts_agile_and_standard_fixtures_with_correct_password() {
    let plaintext_path = fixture_path("plaintext.xlsx");
    let agile_path = fixture_path("agile.xlsx");
    let agile_empty_password_path = fixture_path("agile-empty-password.xlsx");
    let standard_path = fixture_path("standard.xlsx");
    let standard_4_2_path = fixture_path("standard-4.2.xlsx");

    let plaintext = open_workbook_model(&plaintext_path).expect("open plaintext.xlsx");
    assert_expected_contents(&plaintext);

    let agile = open_model_with_password(&agile_path, "password");
    assert_expected_contents(&agile);

    let agile_empty = open_model_with_password(&agile_empty_password_path, "");
    assert_expected_contents(&agile_empty);

    let standard = open_model_with_password(&standard_path, "password");
    assert_expected_contents(&standard);

    let standard_4_2 = open_model_with_password(&standard_4_2_path, "password");
    assert_expected_contents(&standard_4_2);
}

#[test]
fn open_workbook_with_options_decrypts_agile_and_standard_fixtures() {
    for (fixture, password) in [
        ("agile.xlsx", "password"),
        ("agile-empty-password.xlsx", ""),
        ("agile-unicode.xlsx", "pässwörd"),
        ("standard.xlsx", "password"),
        ("standard-4.2.xlsx", "password"),
    ] {
        let path = fixture_path(fixture);
        let wb = open_workbook_with_options(
            &path,
            OpenOptions {
                password: Some(password.to_string()),
            },
        )
        .unwrap_or_else(|err| panic!("open_workbook_with_options({fixture}) failed: {err:?}"));

        match wb {
            Workbook::Xlsx(package) => {
                let bytes = package
                    .write_to_bytes()
                    .expect("serialize decrypted workbook package to bytes");
                let model = formula_io::xlsx::read_workbook_from_reader(Cursor::new(bytes))
                    .expect("parse decrypted workbook bytes");
                assert_expected_contents(&model);
            }
            other => panic!("expected Workbook::Xlsx for {fixture}, got {other:?}"),
        }
    }
}

#[test]
fn decrypts_agile_macro_enabled_xlsm_fixture_with_correct_password() {
    let plaintext_basic_path = fixture_path("plaintext-basic.xlsm");
    let agile_basic_path = fixture_path("agile-basic.xlsm");

    assert_eq!(
        detect_workbook_format(&plaintext_basic_path).expect("detect plaintext-basic.xlsm"),
        WorkbookFormat::Xlsm
    );
    let plaintext_basic_bytes =
        std::fs::read(&plaintext_basic_path).expect("read plaintext-basic.xlsm");
    assert_has_vba_project(&plaintext_basic_bytes);

    let agile_basic = open_model_with_password(&agile_basic_path, "password");
    assert!(
        !agile_basic.sheets.is_empty(),
        "expected decrypted macro workbook to have at least one sheet"
    );
    let agile_basic_bytes =
        open_decrypted_package_bytes_with_password(&agile_basic_path, "password");
    assert_has_vba_project(&agile_basic_bytes);
    assert_detects_xlsm(&agile_basic_bytes);
}

#[test]
fn decrypts_standard_macro_enabled_xlsm_fixture_with_correct_password() {
    let plaintext_basic_path = fixture_path("plaintext-basic.xlsm");
    let standard_basic_path = fixture_path("standard-basic.xlsm");

    assert_eq!(
        detect_workbook_format(&plaintext_basic_path).expect("detect plaintext-basic.xlsm"),
        WorkbookFormat::Xlsm
    );
    let plaintext_basic_bytes =
        std::fs::read(&plaintext_basic_path).expect("read plaintext-basic.xlsm");
    assert_has_vba_project(&plaintext_basic_bytes);

    let standard_basic = open_model_with_password(&standard_basic_path, "password");
    assert!(
        !standard_basic.sheets.is_empty(),
        "expected decrypted macro workbook to have at least one sheet"
    );
    let standard_basic_bytes =
        open_decrypted_package_bytes_with_password(&standard_basic_path, "password");
    assert_has_vba_project(&standard_basic_bytes);
    assert_detects_xlsm(&standard_basic_bytes);
}

#[test]
fn decrypts_standard_fixture_with_correct_password() {
    let standard_path = fixture_path("standard.xlsx");
    let wb = open_model_with_password(&standard_path, "password");
    assert_expected_contents(&wb);
}

#[test]
fn decrypts_standard_large_fixture_with_correct_password() {
    let plaintext_large_path = fixture_path("plaintext-large.xlsx");
    let standard_large_path = fixture_path("standard-large.xlsx");

    let plaintext = open_workbook_model(&plaintext_large_path).expect("open plaintext-large.xlsx");
    let decrypted = open_model_with_password(&standard_large_path, "password");

    let a1 = CellRef::from_a1("A1").unwrap();
    let b2 = CellRef::from_a1("B2").unwrap();

    let sheet_plain = plaintext.sheet_by_name("Sheet1").expect("Sheet1 missing");
    let sheet_decrypted = decrypted.sheet_by_name("Sheet1").expect("Sheet1 missing");

    assert_eq!(sheet_decrypted.value(a1), CellValue::Number(1.0));
    assert_eq!(sheet_decrypted.value(b2), CellValue::Number(2.0));

    // The encrypted fixture should decrypt to the same workbook contents as the plaintext fixture.
    assert_eq!(sheet_decrypted.value(a1), sheet_plain.value(a1));
    assert_eq!(sheet_decrypted.value(b2), sheet_plain.value(b2));
}

#[test]
fn decrypts_standard_rc4_fixture_with_correct_password() {
    let path = fixture_path("standard-rc4.xlsx");
    let wb = open_model_with_password(&path, "password");
    assert_expected_contents(&wb);
}

#[test]
fn open_workbook_with_password_decrypts_standard_fixture() {
    let path = fixture_path("standard.xlsx");
    let decrypted = open_decrypted_package_bytes_with_password(&path, "password");
    let model = formula_io::xlsx::read_workbook_from_reader(Cursor::new(decrypted))
        .expect("parse decrypted standard.xlsx bytes");
    assert_expected_contents(&model);
}

#[test]
fn open_workbook_with_password_decrypts_standard_large_fixture() {
    let plaintext_large_path = fixture_path("plaintext-large.xlsx");
    let standard_large_path = fixture_path("standard-large.xlsx");

    let plaintext = open_workbook_model(&plaintext_large_path).expect("open plaintext-large.xlsx");

    let decrypted_bytes = open_decrypted_package_bytes_with_password(&standard_large_path, "password");
    let decrypted =
        formula_io::xlsx::read_workbook_from_reader(Cursor::new(decrypted_bytes))
            .expect("parse decrypted standard-large.xlsx bytes");

    let a1 = CellRef::from_a1("A1").unwrap();
    let b2 = CellRef::from_a1("B2").unwrap();
    let sheet_plain = plaintext.sheet_by_name("Sheet1").expect("Sheet1 missing");
    let sheet_decrypted = decrypted.sheet_by_name("Sheet1").expect("Sheet1 missing");
    assert_eq!(sheet_decrypted.value(a1), sheet_plain.value(a1));
    assert_eq!(sheet_decrypted.value(b2), sheet_plain.value(b2));
}

#[test]
fn errors_on_missing_password_for_empty_password_fixture() {
    let agile_empty_password_path = fixture_path("agile-empty-password.xlsx");

    let err = open_workbook_model_with_password(&agile_empty_password_path, None)
        .expect_err("expected missing password to error");
    assert!(
        matches!(err, Error::PasswordRequired { .. }),
        "expected Error::PasswordRequired, got {err:?}"
    );
}

#[test]
fn errors_on_missing_password_for_standard_fixture() {
    let standard_path = fixture_path("standard.xlsx");

    let err = open_workbook_model_with_password(&standard_path, None)
        .expect_err("expected missing password to error");
    assert!(
        matches!(err, Error::PasswordRequired { .. }),
        "expected Error::PasswordRequired, got {err:?}"
    );
}

#[test]
fn errors_on_wrong_password_fixtures() {
    let agile_path = fixture_path("agile.xlsx");
    let agile_empty_password_path = fixture_path("agile-empty-password.xlsx");
    let standard_path = fixture_path("standard.xlsx");
    let standard_4_2_path = fixture_path("standard-4.2.xlsx");
    let standard_unicode_path = fixture_path("standard-unicode.xlsx");
    let agile_unicode_path = fixture_path("agile-unicode.xlsx");
    let agile_unicode_excel_path = fixture_path("agile-unicode-excel.xlsx");
    let agile_basic_path = fixture_path("agile-basic.xlsm");
    let standard_basic_path = fixture_path("standard-basic.xlsm");

    for path in [
        &agile_path,
        &agile_empty_password_path,
        &standard_path,
        &standard_4_2_path,
        &standard_unicode_path,
        &agile_unicode_path,
        &agile_unicode_excel_path,
        &agile_basic_path,
        &standard_basic_path,
    ] {
        assert!(
            matches!(
                open_workbook_model_with_password(path, Some("wrong-password")),
                Err(Error::InvalidPassword { .. })
            ),
            "expected InvalidPassword error for {path:?}"
        );
    }
}

#[test]
fn decrypts_agile_unicode_password() {
    let path = fixture_path("agile-unicode.xlsx");
    let wb = open_model_with_password(&path, "pässwörd");
    assert_expected_contents(&wb);
}

#[test]
fn decrypts_agile_unicode_excel_password() {
    let plaintext_path = fixture_path("plaintext-excel.xlsx");
    let encrypted_path = fixture_path("agile-unicode-excel.xlsx");

    let plaintext = open_workbook_model(&plaintext_path).expect("open plaintext-excel.xlsx");
    assert_expected_excel_contents(&plaintext);

    let wb = open_model_with_password(&encrypted_path, "pässwörd🔒");
    assert_expected_excel_contents(&wb);
}

#[test]
fn decrypts_standard_unicode_password() {
    let plaintext_path = fixture_path("plaintext.xlsx");
    let encrypted_path = fixture_path("standard-unicode.xlsx");

    let plaintext = open_workbook_model(&plaintext_path).expect("open plaintext.xlsx");
    assert_expected_contents(&plaintext);

    let wb = open_model_with_password(&encrypted_path, "pässwörd🔒");
    assert_expected_contents(&wb);
}

#[test]
fn agile_unicode_password_different_normalization_fails() {
    // NFC password is "pässwörd" (U+00E4, U+00F6). NFD decomposes those into combining marks.
    let nfd = "pa\u{0308}sswo\u{0308}rd";
    assert_ne!(
        nfd, "pässwörd",
        "strings should differ before UTF-16 encoding"
    );

    let path = fixture_path("agile-unicode.xlsx");
    assert!(
        matches!(
            open_workbook_model_with_password(&path, Some(nfd)),
            Err(Error::InvalidPassword { .. })
        ),
        "expected InvalidPassword error for NFD-normalized password"
    );
}

#[test]
fn standard_unicode_password_different_normalization_fails() {
    // NFC password is "pässwörd🔒" (U+00E4, U+00F6). NFD decomposes those into combining marks, but
    // leaves the non-BMP emoji alone.
    let nfd = "pa\u{0308}sswo\u{0308}rd🔒";
    assert_ne!(
        nfd, "pässwörd🔒",
        "strings should differ before UTF-16 encoding"
    );

    let path = fixture_path("standard-unicode.xlsx");
    assert!(
        matches!(
            open_workbook_model_with_password(&path, Some(nfd)),
            Err(Error::InvalidPassword { .. })
        ),
        "expected InvalidPassword error for NFD-normalized password"
    );
}

#[test]
fn agile_unicode_excel_password_different_normalization_fails() {
    // NFC password is "pässwörd🔒" (U+00E4, U+00F6). NFD decomposes those into combining marks, but
    // leaves the non-BMP emoji alone.
    let nfd = "pa\u{0308}sswo\u{0308}rd🔒";
    assert_ne!(
        nfd, "pässwörd🔒",
        "strings should differ before UTF-16 encoding"
    );

    let path = fixture_path("agile-unicode-excel.xlsx");
    assert!(
        matches!(
            open_workbook_model_with_password(&path, Some(nfd)),
            Err(Error::InvalidPassword { .. })
        ),
        "expected InvalidPassword error for NFD-normalized password"
    );
}
