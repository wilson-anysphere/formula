use std::io::{Cursor, Write};

use formula_model::CellRef;
use formula_model::CellValue;
use formula_xlsx::{
    load_from_bytes_with_password, read_workbook_model_from_bytes_with_password, ReadError,
    XlsxError, XlsxPackage,
};
use ms_offcrypto_writer::Ecma376AgileWriter;
use rand::rngs::StdRng;
use rand::SeedableRng;
use rust_xlsxwriter::Workbook;

fn build_plain_xlsx() -> Vec<u8> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .write_string(0, 0, "Secret")
        .expect("write_string");
    workbook.save_to_buffer().expect("save_to_buffer")
}

fn encrypt_ooxml_agile(plain: &[u8], password: &str) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut rng = StdRng::seed_from_u64(0);
    let mut writer = Ecma376AgileWriter::create(&mut rng, password, cursor).expect("create writer");
    writer.write_all(plain).expect("write plaintext");
    let cursor = writer.into_inner().expect("finalize");
    cursor.into_inner()
}

#[test]
fn read_workbook_model_from_bytes_with_password_decrypts_encrypted_package() {
    let plain = build_plain_xlsx();
    let encrypted = encrypt_ooxml_agile(&plain, "passw0rd");

    let workbook =
        read_workbook_model_from_bytes_with_password(&encrypted, "passw0rd").expect("decrypt");
    assert_eq!(workbook.sheets.len(), 1);

    let sheet = &workbook.sheets[0];
    let value = sheet.value(CellRef::from_a1("A1").unwrap());
    assert_eq!(value, CellValue::String("Secret".to_string()));
}

#[test]
fn load_from_bytes_with_password_decrypts_encrypted_package() {
    let plain = build_plain_xlsx();
    let encrypted = encrypt_ooxml_agile(&plain, "passw0rd");

    let doc = load_from_bytes_with_password(&encrypted, "passw0rd").expect("decrypt");
    let sheet = &doc.workbook.sheets[0];
    let value = sheet.value(CellRef::from_a1("A1").unwrap());
    assert_eq!(value, CellValue::String("Secret".to_string()));
}

#[test]
fn invalid_password_errors_are_exposed() {
    let plain = build_plain_xlsx();
    let encrypted = encrypt_ooxml_agile(&plain, "passw0rd");

    let err = load_from_bytes_with_password(&encrypted, "wrong").expect_err("expected failure");
    assert!(
        matches!(err, ReadError::InvalidPassword),
        "expected ReadError::InvalidPassword, got {err:?}"
    );
}

#[test]
fn xlsx_package_from_bytes_with_password_decrypts_encrypted_package() {
    let plain = build_plain_xlsx();
    let encrypted = encrypt_ooxml_agile(&plain, "passw0rd");

    let pkg = XlsxPackage::from_bytes_with_password(&encrypted, "passw0rd").expect("decrypt");
    assert!(
        pkg.part("xl/workbook.xml").is_some(),
        "expected decrypted package to contain xl/workbook.xml"
    );

    let err = XlsxPackage::from_bytes_with_password(&encrypted, "wrong").expect_err("bad password");
    assert!(
        matches!(err, XlsxError::InvalidPassword),
        "expected XlsxError::InvalidPassword, got {err:?}"
    );
}

