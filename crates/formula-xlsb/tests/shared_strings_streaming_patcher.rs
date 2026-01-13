use std::io::Cursor;

use formula_xlsb::{SharedStringsWriter, SharedStringsWriterStreaming};

fn varint_u32_noncanonical_2bytes(v: u8) -> [u8; 2] {
    // Non-canonical BIFF12 varint encoding for values < 128.
    // Example: 8 -> [0x88, 0x00]
    [v | 0x80, 0x00]
}

fn record_id_sst_noncanonical() -> [u8; 3] {
    // BrtSST id = 0x009F, encoded with an extra zero continuation byte:
    // canonical [0x9F, 0x01] -> non-canonical [0x9F, 0x81, 0x00]
    [0x9F, 0x81, 0x00]
}

fn record_id_si_noncanonical() -> [u8; 2] {
    // BrtSI id = 0x0013 (< 128), non-canonical 2-byte varint.
    varint_u32_noncanonical_2bytes(0x13)
}

fn record_id_sst_end_noncanonical() -> [u8; 3] {
    // BrtSSTEnd id = 0x00A0, non-canonical 3-byte varint.
    [0xA0, 0x81, 0x00]
}

fn build_plain_si_payload(s: &str) -> Vec<u8> {
    // BrtSI payload: [flags:u8=0][cch:u32][utf16 units...]
    let mut out = Vec::new();
    out.push(0u8);
    let units: Vec<u16> = s.encode_utf16().collect();
    out.extend_from_slice(&(units.len() as u32).to_le_bytes());
    for u in units {
        out.extend_from_slice(&u.to_le_bytes());
    }
    out
}

fn write_record_raw(out: &mut Vec<u8>, id_raw: &[u8], len_raw: &[u8], payload: &[u8]) {
    out.extend_from_slice(id_raw);
    out.extend_from_slice(len_raw);
    out.extend_from_slice(payload);
}

fn build_shared_strings_bin_noncanonical_headers(unique_count_in_header: u32) -> Vec<u8> {
    let mut out = Vec::new();

    // BrtSST: [totalCount:u32][uniqueCount:u32]
    let total_count: u32 = 2;
    let mut sst_payload = Vec::new();
    sst_payload.extend_from_slice(&total_count.to_le_bytes());
    sst_payload.extend_from_slice(&unique_count_in_header.to_le_bytes());
    write_record_raw(
        &mut out,
        &record_id_sst_noncanonical(),
        &varint_u32_noncanonical_2bytes(8),
        &sst_payload,
    );

    // BrtSI: "Hello"
    let si1 = build_plain_si_payload("Hello");
    write_record_raw(
        &mut out,
        &record_id_si_noncanonical(),
        &varint_u32_noncanonical_2bytes(si1.len() as u8),
        &si1,
    );

    // BrtSI: "World" (use canonical id/len bytes to ensure mixed encodings are preserved)
    let si2 = build_plain_si_payload("World");
    // Canonical BIFF12 varint for 0x0013 is [0x93, 0x00]? No: canonical would be [0x13].
    // But we want the raw bytes to differ from the first SI to ensure we copy both as-is.
    write_record_raw(&mut out, &[0x13], &[si2.len() as u8], &si2);

    // BrtSSTEnd (len=0 encoded non-canonically)
    write_record_raw(
        &mut out,
        &record_id_sst_end_noncanonical(),
        &varint_u32_noncanonical_2bytes(0),
        &[],
    );

    out
}

#[test]
fn streaming_shared_strings_patcher_matches_in_memory_writer_with_noncanonical_headers() {
    let input = build_shared_strings_bin_noncanonical_headers(2);

    // In-memory writer (existing behavior)
    let mut w = SharedStringsWriter::new(input.clone()).expect("SharedStringsWriter::new");
    w.intern_plain("New").expect("intern_plain");
    w.note_total_ref_delta(1).expect("note_total_ref_delta");
    let expected = w.into_bytes().expect("into_bytes");

    // Streaming patcher
    let mut actual = Vec::new();
    SharedStringsWriterStreaming::patch(
        Cursor::new(input),
        &mut actual,
        &[String::from("New")],
        2, // base SI count ("Hello", "World")
        1, // total ref delta
    )
    .expect("streaming patch");

    assert_eq!(actual, expected);
}

#[test]
fn streaming_shared_strings_patcher_repairs_unique_count_like_in_memory_writer() {
    // Header claims uniqueCount=100, but there are only 2 SI records.
    let input = build_shared_strings_bin_noncanonical_headers(100);

    // In-memory writer repairs uniqueCount whenever it patches the header.
    let mut w = SharedStringsWriter::new(input.clone()).expect("SharedStringsWriter::new");
    w.note_total_ref_delta(-1).expect("note_total_ref_delta");
    let expected = w.into_bytes().expect("into_bytes");

    let mut actual = Vec::new();
    SharedStringsWriterStreaming::patch(
        Cursor::new(input),
        &mut actual,
        &[],
        2,  // base SI count
        -1, // total ref delta
    )
    .expect("streaming patch");

    assert_eq!(actual, expected);
}
