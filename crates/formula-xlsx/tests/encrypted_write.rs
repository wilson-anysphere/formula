#![cfg(not(target_arch = "wasm32"))]

use std::io::{Cursor, Read};

use formula_model::{Cell, CellRef, CellValue, Workbook};
use formula_xlsx::{
    write_workbook_to_writer_encrypted, write_workbook_to_writer_with_kind, WorkbookKind,
};

fn noisy_ascii_string(len: usize) -> String {
    // Deterministic (non-cryptographic) "random-ish" string to ensure the resulting XLSX package
    // exceeds 4096 bytes. This exercises multi-segment (4096-byte) MS-OFFCRYPTO encryption.
    let mut out = String::with_capacity(len);
    let mut state: u32 = 0x1234_5678;
    for _ in 0..len {
        state = state.wrapping_mul(1664525).wrapping_add(1013904223);
        let idx = (state % 62) as u8;
        let c = match idx {
            0..=9 => b'0' + idx,
            10..=35 => b'A' + (idx - 10),
            _ => b'a' + (idx - 36),
        };
        out.push(char::from(c));
    }
    out
}

fn build_simple_workbook() -> Workbook {
    let mut workbook = Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1").expect("add sheet");
    let sheet = workbook.sheet_mut(sheet_id).expect("sheet");
    sheet.set_cell(
        CellRef::from_a1("A1").expect("a1"),
        Cell::new(CellValue::String(noisy_ascii_string(16 * 1024))),
    );
    workbook
}

#[test]
fn encrypted_write_roundtrips_to_valid_zip() {
    let workbook = build_simple_workbook();

    let password = "password123";

    // Ensure the underlying package bytes are large enough for our decrypt oracle.
    let mut plain_cursor = Cursor::new(Vec::new());
    write_workbook_to_writer_with_kind(&workbook, &mut plain_cursor, WorkbookKind::Workbook)
        .expect("write plaintext workbook");
    let plain_zip_bytes = plain_cursor.into_inner();
    assert!(
        plain_zip_bytes.len() > 4096,
        "expected plaintext zip to be >4096 bytes (got {})",
        plain_zip_bytes.len()
    );

    let mut ole_bytes = Vec::new();
    write_workbook_to_writer_encrypted(
        &workbook,
        &mut ole_bytes,
        WorkbookKind::Workbook,
        password,
    )
    .expect("write encrypted workbook");

    // Decrypt the OLE wrapper back into the underlying package bytes.
    let zip_bytes = formula_office_crypto::decrypt_encrypted_package_ole(&ole_bytes, password)
        .expect("decrypt workbook");
    assert_eq!(
        zip_bytes, plain_zip_bytes,
        "decrypted package bytes should match plaintext writer output"
    );

    // Also validate the `XlsxPackage` encryption helper.
    let package = formula_xlsx::XlsxPackage::from_bytes(&plain_zip_bytes).expect("parse package");
    let package_zip_bytes = package.write_to_bytes().expect("rewrite package");
    let ole_from_package = package
        .write_to_encrypted_ole_bytes(password)
        .expect("encrypt package");
    let decrypted_package =
        formula_office_crypto::decrypt_encrypted_package_ole(&ole_from_package, password)
            .expect("decrypt package");
    assert_eq!(
        decrypted_package, package_zip_bytes,
        "package encryption should roundtrip (XlsxPackage::write_to_bytes())"
    );

    let cursor = Cursor::new(zip_bytes);
    let mut zip = zip::ZipArchive::new(cursor).expect("open decrypted zip");

    let mut workbook_xml = String::new();
    zip.by_name("xl/workbook.xml")
        .expect("xl/workbook.xml present")
        .read_to_string(&mut workbook_xml)
        .expect("read workbook.xml");
    assert!(
        workbook_xml.contains("Sheet1"),
        "expected workbook.xml to include Sheet1, got: {workbook_xml:?}"
    );

    zip.by_name("[Content_Types].xml")
        .expect("[Content_Types].xml present");
    zip.by_name("xl/worksheets/sheet1.xml")
        .expect("sheet1.xml present");
}
