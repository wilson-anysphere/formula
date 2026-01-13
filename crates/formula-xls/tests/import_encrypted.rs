use std::path::Path;
use std::path::PathBuf;
use std::{io::Read, ops::Range as ByteRange};

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

fn workbook_globals_filepass_payload_range(workbook_stream: &[u8]) -> Option<ByteRange<usize>> {
    // Scan the workbook globals substream for FILEPASS. Stop at EOF or a subsequent BOF.
    const RECORD_BOF_BIFF8: u16 = 0x0809;
    const RECORD_EOF: u16 = 0x000A;
    const RECORD_FILEPASS: u16 = 0x002F;

    let (record_id, bof_payload, mut offset) = read_record(workbook_stream, 0)?;
    if record_id != RECORD_BOF_BIFF8 {
        return None;
    }
    if bof_payload.len() < 4 || bof_payload[0..2] != [0x00, 0x06] || bof_payload[2..4] != [0x05, 0x00]
    {
        return None;
    }

    while let Some((record_id, data, next)) = read_record(workbook_stream, offset) {
        if record_id == RECORD_EOF {
            break;
        }
        // A subsequent BOF indicates the next substream (worksheet, etc); FILEPASS must be in
        // workbook globals.
        if record_id == RECORD_BOF_BIFF8 {
            break;
        }
        if record_id == RECORD_FILEPASS {
            // Return the payload range within the workbook stream buffer.
            let data_start = offset + 4;
            let data_end = data_start + data.len();
            return Some(data_start..data_end);
        }
        offset = next;
    }
    None
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
        let Some(filepass_range) = workbook_globals_filepass_payload_range(&workbook_stream) else {
            panic!("expected to find FILEPASS in workbook globals substream for {path:?}");
        };
        let filepass_payload = &workbook_stream[filepass_range];
        assert!(
            filepass_payload.starts_with(expected_filepass_prefix),
            "unexpected FILEPASS payload prefix for {path:?}; expected {:02X?}, got {:02X?}",
            expected_filepass_prefix,
            &filepass_payload[..filepass_payload.len().min(expected_filepass_prefix.len())]
        );
    }
}
