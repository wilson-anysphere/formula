use std::io::Read;
use std::path::Path;
use std::path::PathBuf;

fn read_workbook_stream(path: &Path) -> Vec<u8> {
    let bytes = std::fs::read(path).expect("read xls fixture");
    let cursor = std::io::Cursor::new(bytes);
    let mut ole = cfb::CompoundFile::open(cursor).expect("open cfb");
    let mut stream = ole.open_stream("Workbook").expect("Workbook stream");
    let mut out = Vec::new();
    stream
        .read_to_end(&mut out)
        .expect("read Workbook stream bytes");
    out
}

fn read_record(stream: &[u8], offset: usize) -> Option<(u16, &[u8], usize)> {
    if offset + 4 > stream.len() {
        return None;
    }
    let record_id = u16::from_le_bytes([stream[offset], stream[offset + 1]]);
    let len = u16::from_le_bytes([stream[offset + 2], stream[offset + 3]]) as usize;
    let data_start = offset + 4;
    let data_end = data_start.checked_add(len)?;
    let data = stream.get(data_start..data_end)?;
    Some((record_id, data, data_end))
}

#[test]
fn errors_on_encrypted_xls_fixtures() {
    let fixtures_dir = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("encrypted");

    let fixtures = [
        ("biff8_xor_pw_open.xls", &[0x00, 0x00, 0x34, 0x12, 0x78, 0x56][..]),
        (
            "biff8_rc4_standard_pw_open.xls",
            &[0x01, 0x00, 0x01, 0x00][..],
        ),
        (
            "biff8_rc4_cryptoapi_pw_open.xls",
            &[0x01, 0x00, 0x02, 0x00][..],
        ),
    ];

    for (filename, expected_filepass_prefix) in fixtures {
        let path = fixtures_dir.join(filename);
        let err = formula_xls::import_xls_path(&path)
            .expect_err(&format!("expected encrypted workbook error for {path:?}"));
        assert!(
            matches!(err, formula_xls::ImportError::EncryptedWorkbook),
            "expected ImportError::EncryptedWorkbook for {path:?}, got {err:?}"
        );

        let msg = err.to_string().to_lowercase();
        assert!(
            msg.contains("encrypted"),
            "expected error message to mention encryption; got: {msg}"
        );
        assert!(
            msg.contains("password"),
            "expected error message to mention password protection; got: {msg}"
        );

        // Assert the underlying fixture stream matches its documented encryption scheme.
        let workbook_stream = read_workbook_stream(&path);
        let (record_id, bof_payload, next) =
            read_record(&workbook_stream, 0).expect("read BOF record");
        assert_eq!(
            record_id, 0x0809,
            "expected first record to be BIFF8 BOF in {path:?}"
        );
        assert!(
            bof_payload.len() >= 4
                && bof_payload[0..2] == [0x00, 0x06]
                && bof_payload[2..4] == [0x05, 0x00],
            "expected BOF payload to indicate BIFF8 workbook globals in {path:?}"
        );

        let (record_id, filepass_payload, next) =
            read_record(&workbook_stream, next).expect("read FILEPASS record");
        assert_eq!(
            record_id, 0x002F,
            "expected second record to be FILEPASS in {path:?}"
        );
        assert!(
            filepass_payload.starts_with(expected_filepass_prefix),
            "unexpected FILEPASS payload prefix for {path:?}; expected {:02X?}, got {:02X?}",
            expected_filepass_prefix,
            &filepass_payload[..filepass_payload.len().min(expected_filepass_prefix.len())]
        );

        let (record_id, _, _) = read_record(&workbook_stream, next).expect("read EOF record");
        assert_eq!(
            record_id, 0x000A,
            "expected third record to be EOF in {path:?}"
        );
    }
}
