use std::io::{Cursor, Write};

use formula_vba::{compress_container, project_normalized_data_v3, DirParseError, ParseError};

fn push_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(data.len() as u32).to_le_bytes());
    out.extend_from_slice(data);
}

fn build_vba_bin_with_dir_decompressed(dir_decompressed: &[u8]) -> Vec<u8> {
    let dir_container = compress_container(dir_decompressed);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    ole.create_storage("VBA").expect("VBA storage");
    {
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }
    ole.into_inner().into_inner()
}

#[test]
fn project_normalized_data_v3_missing_vba_dir_stream() {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    ole.create_storage("VBA").expect("VBA storage");

    let vba_bin = ole.into_inner().into_inner();
    let err = project_normalized_data_v3(&vba_bin).expect_err("expected MissingStream");
    match err {
        ParseError::MissingStream("VBA/dir") => {}
        other => panic!("expected MissingStream(\"VBA/dir\"), got {other:?}"),
    }
}

#[test]
fn project_normalized_data_v3_dir_truncated_record_header() {
    // One valid record followed by <6 leftover bytes so the next record header is truncated.
    let dir_decompressed = {
        let mut out = Vec::new();
        // REFERENCEREGISTERED
        push_record(&mut out, 0x000D, b"X");
        out.extend_from_slice(&[0xAA, 0xBB, 0xCC, 0xDD, 0xEE]); // 5 bytes (truncated header)
        out
    };

    let vba_bin = build_vba_bin_with_dir_decompressed(&dir_decompressed);
    let err = project_normalized_data_v3(&vba_bin).expect_err("expected dir parse error");
    match err {
        ParseError::Dir(DirParseError::Truncated) => {}
        other => panic!("expected Dir(Truncated), got {other:?}"),
    }
}

#[test]
fn project_normalized_data_v3_dir_bad_record_length_beyond_buffer() {
    // Header claims `len=10`, but only 1 payload byte is present.
    let dir_decompressed = {
        let mut out = Vec::new();
        out.extend_from_slice(&0x000Du16.to_le_bytes()); // REFERENCEREGISTERED
        out.extend_from_slice(&10u32.to_le_bytes());
        out.extend_from_slice(b"X"); // insufficient payload
        out
    };

    let vba_bin = build_vba_bin_with_dir_decompressed(&dir_decompressed);
    let err = project_normalized_data_v3(&vba_bin).expect_err("expected dir parse error");
    match err {
        ParseError::Dir(DirParseError::BadRecordLength { id, len }) => {
            assert_eq!(id, 0x000D);
            assert_eq!(len, 10);
        }
        other => panic!("expected Dir(BadRecordLength), got {other:?}"),
    }
}

