use std::io::Read;
use std::path::PathBuf;

use formula_model::CellValue;

fn encrypted_payload_len_after_filepass(workbook_stream: &[u8]) -> Option<usize> {
    const RECORD_FILEPASS: u16 = 0x002F;

    let mut offset = 0usize;
    let mut saw_filepass = false;
    let mut total = 0usize;

    while offset + 4 <= workbook_stream.len() {
        let record_id = u16::from_le_bytes([workbook_stream[offset], workbook_stream[offset + 1]]);
        let len =
            u16::from_le_bytes([workbook_stream[offset + 2], workbook_stream[offset + 3]]) as usize;
        offset += 4;
        let data_end = offset.checked_add(len)?;
        if data_end > workbook_stream.len() {
            break;
        }

        if saw_filepass {
            total += len;
        }
        if record_id == RECORD_FILEPASS {
            saw_filepass = true;
        }

        offset = data_end;
    }

    saw_filepass.then_some(total)
}

#[test]
fn imports_encrypted_xls_rc4_cryptoapi_across_1024_byte_boundary() {
    // Fixture password: "password"
    // Use the top-level fixture so other crates/tests can share the same encrypted workbook corpus.
    let fixture_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/encryption/biff8_rc4_cryptoapi_boundary_pw_open.xls");

    // Best-effort: verify the encrypted payload after FILEPASS crosses the 1024-byte RC4 rekey
    // boundary so decryption must re-key mid-stream.
    let file = std::fs::File::open(&fixture_path).expect("open encrypted fixture");
    let mut comp = cfb::CompoundFile::open(file).expect("open cfb");
    let mut stream = comp
        .open_stream("Workbook")
        .or_else(|_| comp.open_stream("Book"))
        .expect("Workbook stream");
    let mut workbook_stream = Vec::new();
    stream
        .read_to_end(&mut workbook_stream)
        .expect("read Workbook stream");

    let encrypted_len = encrypted_payload_len_after_filepass(&workbook_stream)
        .expect("expected FILEPASS record in workbook globals");
    assert!(
        encrypted_len > 1024,
        "expected >1024 bytes of encrypted payload after FILEPASS, got {encrypted_len}"
    );

    let result = formula_xls::import_xls_path_with_password(&fixture_path, Some("password"))
        .expect("import encrypted xls with password");
    let sheet = result
        .workbook
        .sheet_by_name("Sheet1")
        .expect("Sheet1 missing");

    // Validate values near the end of the stream/sheet so the test fails if RC4 rekey logic is
    // broken past the 1024-byte boundary.
    assert_eq!(
        sheet.value_a1("A400").unwrap(),
        CellValue::Number(399.0),
        "expected numeric value in A400"
    );
    assert_eq!(
        sheet.value_a1("B400").unwrap(),
        CellValue::String("RC4_BOUNDARY_OK".to_string()),
        "expected decrypted marker string in B400"
    );
}
