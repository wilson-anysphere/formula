use formula_xlsb::biff12_varint;
use formula_xlsb::SharedStringsWriter;
use pretty_assertions::assert_eq;

const BRT_SST: u32 = 0x009F;
const BRT_SI: u32 = 0x0013;
const BRT_SST_END: u32 = 0x00A0;

fn push_record(out: &mut Vec<u8>, id: u32, payload: &[u8]) {
    biff12_varint::write_record_id(out, id).expect("write record id");
    biff12_varint::write_record_len(out, payload.len() as u32).expect("write record len");
    out.extend_from_slice(payload);
}

#[test]
fn shared_strings_writer_reuses_plain_si_with_trailing_bytes() {
    let text = "Hello";

    let mut shared_strings_bin = Vec::new();

    // BrtSST payload: [cstTotal: u32][cstUnique: u32]
    let mut sst_payload = Vec::new();
    sst_payload.extend_from_slice(&1u32.to_le_bytes());
    sst_payload.extend_from_slice(&1u32.to_le_bytes());
    push_record(&mut shared_strings_bin, BRT_SST, &sst_payload);

    // BrtSI payload:
    //   [flags: u8][text: XLWideString]
    // XLWideString: [cch: u32][utf16 chars...]
    let utf16: Vec<u16> = text.encode_utf16().collect();
    let cch = utf16.len() as u32;
    let mut si_payload = Vec::new();
    si_payload.push(0u8); // flags == 0 (plain)
    si_payload.extend_from_slice(&cch.to_le_bytes());
    for unit in utf16 {
        si_payload.extend_from_slice(&unit.to_le_bytes());
    }
    // Some writers include benign trailing bytes after the UTF-16 text even with flags==0.
    si_payload.extend_from_slice(&[0xAA, 0xBB, 0xCC]);
    push_record(&mut shared_strings_bin, BRT_SI, &si_payload);

    // BrtSSTEnd payload is empty.
    push_record(&mut shared_strings_bin, BRT_SST_END, &[]);

    let original = shared_strings_bin.clone();
    let mut writer = SharedStringsWriter::new(shared_strings_bin).expect("parse shared strings");

    let idx = writer
        .intern_plain(text)
        .expect("intern plain shared string");
    assert_eq!(idx, 0, "expected plain SI entry to be reused");

    let out = writer.into_bytes().expect("serialize shared strings");
    assert_eq!(out, original, "expected byte-identical stream for pure reuse");
}

