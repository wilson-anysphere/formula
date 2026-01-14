use std::ffi::OsStr;
use std::fs::File;
use std::io::{Cursor, Read};
use std::path::Path;

use formula_xlsb::biff12_varint::{
    read_record_id, read_record_len, write_record_id, write_record_len,
};
use pretty_assertions::assert_eq;
use proptest::prelude::*;

fn is_encodable_record_id(id: u32) -> bool {
    id <= 0x0FFF_FFFF
}

fn valid_record_id() -> impl Strategy<Value = u32> {
    0u32..=0x0FFF_FFFF
}

#[test]
fn record_id_vectors_lock_in_encoding() {
    let vectors: &[(u32, &[u8])] = &[
        (0x00, &[0x00]),
        (0x01, &[0x01]),
        (0x7F, &[0x7F]),
        (0x80, &[0x80, 0x01]),
        (0x81, &[0x81, 0x01]),
        (0xFF, &[0xFF, 0x01]),
        // Taken from `simple.xlsb` (`xl/workbook.bin` begins with BrtBeginBook = 0x0083).
        (0x0083, &[0x83, 0x01]),
        (0x009C, &[0x9C, 0x01]),
        (0x3FFF, &[0xFF, 0x7F]),
        (0x4000, &[0x80, 0x80, 0x01]),
        (0x0FFF_FFFF, &[0xFF, 0xFF, 0xFF, 0x7F]),
    ];

    for (id, expected) in vectors {
        let mut encoded = Vec::new();
        write_record_id(&mut encoded, *id).expect("encode record id");
        assert_eq!(
            encoded, *expected,
            "record id encoding mismatch for id={id:#x}"
        );

        let mut cursor = Cursor::new(&encoded);
        let decoded = read_record_id(&mut cursor)
            .expect("decode record id")
            .expect("some id");
        assert_eq!(decoded, *id, "record id round-trip mismatch for id={id:#x}");
        assert_eq!(cursor.position() as usize, encoded.len());
    }
}

#[test]
fn record_len_vectors_lock_in_encoding() {
    let vectors: &[(u32, &[u8])] = &[
        (0x00, &[0x00]),
        (0x01, &[0x01]),
        (0x7F, &[0x7F]),
        (0x80, &[0x80, 0x01]),
        (0x3FFF, &[0xFF, 0x7F]),
        (0x4000, &[0x80, 0x80, 0x01]),
        (0x0FFF_FFFF, &[0xFF, 0xFF, 0xFF, 0x7F]),
    ];

    for (len, expected) in vectors {
        let mut encoded = Vec::new();
        write_record_len(&mut encoded, *len).expect("encode record len");
        assert_eq!(
            encoded, *expected,
            "record length encoding mismatch for len={len:#x}"
        );

        let mut cursor = Cursor::new(&encoded);
        let decoded = read_record_len(&mut cursor)
            .expect("decode record len")
            .expect("some len");
        assert_eq!(
            decoded, *len,
            "record length round-trip mismatch for len={len:#x}"
        );
        assert_eq!(cursor.position() as usize, encoded.len());
    }
}

proptest! {
    #![proptest_config(ProptestConfig {
        failure_persistence: None,
        .. ProptestConfig::default()
    })]

    #[test]
    fn record_id_roundtrips(id in valid_record_id()) {
        let mut encoded = Vec::new();
        write_record_id(&mut encoded, id).unwrap();

        let mut cursor = Cursor::new(&encoded);
        let decoded = read_record_id(&mut cursor).unwrap().unwrap();

        prop_assert_eq!(decoded, id);
        prop_assert_eq!(cursor.position() as usize, encoded.len());
    }

    #[test]
    fn record_id_writer_accepts_exactly_encodable_values(id in 0u32..=0x1FFF_FFFF) {
        let mut encoded = Vec::new();
        let res = write_record_id(&mut encoded, id);

        if is_encodable_record_id(id) {
            prop_assert!(res.is_ok());
            let mut cursor = Cursor::new(&encoded);
            let decoded = read_record_id(&mut cursor).unwrap().unwrap();
            prop_assert_eq!(decoded, id);
            prop_assert_eq!(cursor.position() as usize, encoded.len());
        } else {
            prop_assert!(res.is_err());
        }
    }

    #[test]
    fn record_len_roundtrips(len in 0u32..=0x0FFF_FFFF) {
        let mut encoded = Vec::new();
        write_record_len(&mut encoded, len).unwrap();

        let mut cursor = Cursor::new(&encoded);
        let decoded = read_record_len(&mut cursor).unwrap().unwrap();

        prop_assert_eq!(decoded, len);
        prop_assert_eq!(cursor.position() as usize, encoded.len());
    }
}

#[test]
fn record_len_rejects_values_above_28_bits() {
    let mut encoded = Vec::new();
    let err = write_record_len(&mut encoded, 0x1000_0000).expect_err("len too large");
    assert_eq!(err.kind(), std::io::ErrorKind::InvalidInput);
}

#[test]
fn fixture_headers_reserialize_byte_for_byte() {
    let fixtures_dir = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures");
    for entry in std::fs::read_dir(&fixtures_dir).expect("read fixtures dir") {
        let path = entry.expect("fixture entry").path();
        if path.extension() != Some(OsStr::new("xlsb")) {
            continue;
        }

        let file = File::open(&path).expect("open xlsb fixture");
        let mut zip = zip::ZipArchive::new(file).expect("read xlsb zip");

        for i in 0..zip.len() {
            let mut entry = zip.by_index(i).expect("read zip entry");
            let name = entry.name().to_string();
            if !is_biff12_record_stream_part(&name) {
                continue;
            }

            // Do not trust `ZipFile::size()` for allocation; ZIP metadata is untrusted and can
            // advertise enormous uncompressed sizes (zip-bomb style OOM).
            let mut bytes = Vec::new();
            entry.read_to_end(&mut bytes).expect("read part bytes");

            let mut cursor = Cursor::new(&bytes);
            let mut record_index = 0usize;

            loop {
                let start = cursor.position() as usize;
                let Some(id) = read_record_id(&mut cursor).expect("read record id") else {
                    break;
                };
                let Some(len) = read_record_len(&mut cursor).expect("read record len") else {
                    panic!(
                        "unexpected EOF while reading record len in {} ({name}) after id={id:#x}",
                        path.display()
                    );
                };
                let end = cursor.position() as usize;

                let mut encoded_header = Vec::new();
                write_record_id(&mut encoded_header, id).expect("re-encode id");
                write_record_len(&mut encoded_header, len).expect("re-encode len");

                assert_eq!(
                    encoded_header,
                    bytes[start..end],
                    "header mismatch in {} ({name}) record {record_index} (id={id:#x}, len={len:#x})",
                    path.display()
                );

                let next_pos = cursor.position() + len as u64;
                assert!(
                    (next_pos as usize) <= bytes.len(),
                    "record payload out of bounds in {} ({name}) record {record_index} (id={id:#x}, len={len:#x})",
                    path.display()
                );
                cursor.set_position(next_pos);
                record_index += 1;
            }

            assert_eq!(
                cursor.position() as usize,
                bytes.len(),
                "did not consume entire part stream for {} ({name})",
                path.display()
            );
        }
    }
}

fn is_biff12_record_stream_part(name: &str) -> bool {
    matches!(
        name,
        "xl/workbook.bin" | "xl/sharedStrings.bin" | "xl/styles.bin" | "xl/calcChain.bin"
    ) || name.starts_with("xl/worksheets/")
}
