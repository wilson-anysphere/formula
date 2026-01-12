use std::io::{Cursor, Write};

use formula_vba::{
    compress_container, project_normalized_data, project_normalized_data_v3, DirParseError,
    ParseError,
};

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

fn utf16le_bytes(s: &str) -> Vec<u8> {
    let mut out = Vec::new();
    for unit in s.encode_utf16() {
        out.extend_from_slice(&unit.to_le_bytes());
    }
    out
}

#[test]
fn project_normalized_data_includes_expected_dir_records_and_prefers_unicode_variants() {
    // Build a synthetic decompressed `VBA/dir` stream with:
    // - multiple included project-info records
    // - one excluded record
    // - ANSI + UNICODE pairs where the algorithm must prefer the UNICODE record.
    let dir_decompressed = {
        let mut out = Vec::new();

        // Included: PROJECTSYSKIND
        push_record(&mut out, 0x0001, &1u32.to_le_bytes());
        // Included: PROJECTLCID
        push_record(&mut out, 0x0002, &0x0409u32.to_le_bytes());
        // Included: PROJECTCODEPAGE
        push_record(&mut out, 0x0003, &1252u16.to_le_bytes());
        // Included: PROJECTNAME
        push_record(&mut out, 0x0004, b"MyProject");

        // Included (ANSI), but followed by UNICODE -> should be skipped in favor of UNICODE.
        push_record(&mut out, 0x0005, b"Doc");
        // Included: PROJECTDOCSTRINGUNICODE (paired with 0x0005 above).
        push_record(&mut out, 0x0040, &utf16le_bytes("Doc"));

        // Excluded: REFERENCEREGISTERED (0x000D)
        push_record(&mut out, 0x000D, b"{EXCLUDED}");

        // Included (ANSI), but followed by UNICODE -> should be skipped in favor of UNICODE.
        push_record(&mut out, 0x000C, b"Const=1");
        // Included: PROJECTCONSTANTSUNICODE (paired with 0x000C above).
        push_record(&mut out, 0x003C, &utf16le_bytes("Const=1"));

        out
    };

    let vba_bin = build_vba_bin_with_dir_decompressed(&dir_decompressed);
    let normalized = project_normalized_data(&vba_bin).expect("ProjectNormalizedData");

    let expected = [
        1u32.to_le_bytes().as_slice(),
        0x0409u32.to_le_bytes().as_slice(),
        1252u16.to_le_bytes().as_slice(),
        b"MyProject".as_slice(),
        utf16le_bytes("Doc").as_slice(),
        utf16le_bytes("Const=1").as_slice(),
    ]
    .concat();

    assert_eq!(normalized, expected);
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

