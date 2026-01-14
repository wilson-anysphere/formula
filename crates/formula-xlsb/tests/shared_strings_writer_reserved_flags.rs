use formula_xlsb::{biff12_varint, SharedStringsWriter};
use pretty_assertions::assert_eq;

const SST: u32 = 0x009F;
const SI: u32 = 0x0013;
const SST_END: u32 = 0x00A0;

fn write_record(out: &mut Vec<u8>, id: u32, payload: &[u8]) {
    biff12_varint::write_record_id(out, id).expect("write record id");
    biff12_varint::write_record_len(out, payload.len() as u32).expect("write record len");
    out.extend_from_slice(payload);
}

fn build_minimal_shared_strings_bin_with_si(flags: u8, text: &str, trailing: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();

    // BrtSST: [totalCount:u32][uniqueCount:u32]
    let mut sst = Vec::new();
    sst.extend_from_slice(&0u32.to_le_bytes()); // totalCount
    sst.extend_from_slice(&1u32.to_le_bytes()); // uniqueCount
    write_record(&mut out, SST, &sst);

    // BrtSI: [flags:u8][cch:u32][utf16 chars...][optional trailing bytes...]
    let units: Vec<u16> = text.encode_utf16().collect();
    let mut si = Vec::new();
    si.push(flags);
    si.extend_from_slice(&(units.len() as u32).to_le_bytes());
    for unit in units {
        si.extend_from_slice(&unit.to_le_bytes());
    }
    si.extend_from_slice(trailing);
    write_record(&mut out, SI, &si);

    write_record(&mut out, SST_END, &[]);

    out
}

#[test]
fn shared_strings_writer_treats_reserved_flag_si_as_reusable_plain_string() {
    let shared_strings_bin =
        build_minimal_shared_strings_bin_with_si(0x80, "Hello", &[0xAA, 0xBB, 0xCC]);

    let mut writer = SharedStringsWriter::new(shared_strings_bin.clone()).expect("new writer");
    let idx = writer.intern_plain("Hello").expect("intern plain");
    assert_eq!(idx, 0, "expected reserved-flag SI to be indexed for reuse");

    // If the string was reused, `into_bytes` should return the original bytes unchanged.
    let out = writer.into_bytes().expect("into_bytes");
    assert_eq!(
        out, shared_strings_bin,
        "expected SharedStringsWriter to avoid appending a duplicate SI when the original SI has only reserved flag bits set"
    );
}

