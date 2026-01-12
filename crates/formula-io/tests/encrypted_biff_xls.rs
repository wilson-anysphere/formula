use std::io::{Cursor, Write};

use formula_io::{open_workbook, open_workbook_model, Error};

fn record(record_id: u16, payload: &[u8]) -> Vec<u8> {
    let mut out = Vec::with_capacity(4 + payload.len());
    out.extend_from_slice(&record_id.to_le_bytes());
    out.extend_from_slice(&(payload.len() as u16).to_le_bytes());
    out.extend_from_slice(payload);
    out
}

fn encrypted_biff_xls_bytes_filepass_after_large_prefix() -> Vec<u8> {
    // Some callers (like `formula-io` format detection) use a capped scan budget when sniffing OLE
    // workbooks for BIFF encryption markers. Ensure we still surface a clean encryption error even
    // if FILEPASS occurs after that scan budget and the file is (mis)classified as a plain `.xls`.
    const RECORD_BOF_BIFF8: u16 = 0x0809;
    const RECORD_FILEPASS: u16 = 0x002F;
    const RECORD_EOF: u16 = 0x000A;

    // Arbitrary non-BOF/EOF record id with a large payload to exceed common sniffing limits.
    const RECORD_DUMMY: u16 = 0x00FC;

    let mut workbook_stream = Vec::new();
    workbook_stream.extend_from_slice(&record(RECORD_BOF_BIFF8, &[0u8; 16]));

    // Add a >4MiB prefix without introducing another BOF/EOF; this ensures sniffers with a scan
    // budget will bail out before reaching FILEPASS.
    let dummy_payload = vec![0u8; u16::MAX as usize];
    let dummy_record = record(RECORD_DUMMY, &dummy_payload);
    for _ in 0..64 {
        workbook_stream.extend_from_slice(&dummy_record);
    }

    workbook_stream.extend_from_slice(&record(RECORD_FILEPASS, &[]));
    workbook_stream.extend_from_slice(&record(RECORD_EOF, &[]));

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole.create_stream("Workbook").expect("Workbook stream");
        stream
            .write_all(&workbook_stream)
            .expect("write Workbook stream");
    }
    ole.into_inner().into_inner()
}

#[test]
fn errors_on_encrypted_biff_xls_filepass_even_if_sniffer_misses_it() {
    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("encrypted_biff_large.xls");
    std::fs::write(&path, encrypted_biff_xls_bytes_filepass_after_large_prefix())
        .expect("write encrypted fixture");

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
