use formula_xlsb::biff12_varint::{write_record_id, write_record_len};
use formula_xlsb::Styles;
use formula_format::BUILTIN_NUM_FMT_ID_PLACEHOLDER_PREFIX;

fn push_record(buf: &mut Vec<u8>, id: u32, data: &[u8]) {
    write_record_id(buf, id).expect("write record id");
    write_record_len(buf, data.len() as u32).expect("write record len");
    buf.extend_from_slice(data);
}

#[test]
fn preserves_unknown_reserved_num_fmt_ids_as_placeholders() {
    // Minimal `xl/styles.bin` payload containing one XF record whose numFmtId is a reserved built-in
    // outside the 0–49 mapping.
    //
    // XLSB allows these ids to appear without an explicit format code, so we preserve the id via
    // `__builtin_numFmtId:<id>` so downstream converters can round-trip it.
    const BEGIN_CELL_XFS: u32 = 0x0122;
    const END_CELL_XFS: u32 = 0x0123;
    const BRT_XF: u32 = 0x002F;

    let mut bytes = Vec::new();
    push_record(&mut bytes, BEGIN_CELL_XFS, &[]);
    push_record(&mut bytes, BRT_XF, &50u16.to_le_bytes());
    push_record(&mut bytes, END_CELL_XFS, &[]);

    let styles = Styles::parse(&bytes).expect("parse styles");
    let style = styles.get(0).expect("xf 0");
    assert_eq!(style.num_fmt_id, 50);
    let expected = format!("{BUILTIN_NUM_FMT_ID_PLACEHOLDER_PREFIX}50");
    assert_eq!(style.number_format.as_deref(), Some(expected.as_str()));
    assert!(
        style.is_date_time,
        "reserved ids in 50..=58 should be treated as datetime"
    );
}
