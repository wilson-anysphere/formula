use std::io::{Cursor, Read};
use std::io::{self};

use formula_biff::encode_rgce;
use formula_xlsb::{
    biff12_varint, patch_sheet_bin, patch_sheet_bin_streaming, CellEdit, CellValue, Error,
};

mod fixture_builder;
use fixture_builder::XlsbFixtureBuilder;

// Record IDs follow the conventions used by `formula-xlsb`'s BIFF12 reader.
const SHEETDATA: u32 = 0x0091;
const SHEETDATA_END: u32 = 0x0092;
const ROW: u32 = 0x0000;
const BLANK: u32 = 0x0001;
const NUM: u32 = 0x0002;
const BOOLERR: u32 = 0x0003;
const BOOL: u32 = 0x0004;
const FLOAT: u32 = 0x0005;

fn read_sheet1_bin_from_fixture(bytes: &[u8]) -> Vec<u8> {
    let mut zip = zip::ZipArchive::new(Cursor::new(bytes)).expect("open xlsb zip");
    let mut entry = zip
        .by_name("xl/worksheets/sheet1.bin")
        .expect("read sheet1.bin");
    // Do not trust `ZipFile::size()` for allocation; ZIP metadata is untrusted and can
    // advertise enormous uncompressed sizes (zip-bomb style OOM).
    let mut out = Vec::new();
    entry.read_to_end(&mut out).expect("read sheet bytes");
    out
}

fn assert_invalid_input_contains(err: Error, needle: &str) {
    match err {
        Error::Io(io_err) => {
            assert_eq!(io_err.kind(), io::ErrorKind::InvalidInput);
            assert!(
                io_err.to_string().contains(needle),
                "expected error message to contain {needle:?}, got: {io_err}"
            );
        }
        other => panic!("expected Error::Io, got {other:?}"),
    }
}

fn append_trailing_bytes_to_cell_payload(
    sheet_bin: &[u8],
    target_row: u32,
    target_col: u32,
    record_id: u32,
    extra: &[u8],
) -> Vec<u8> {
    let mut cursor = Cursor::new(sheet_bin);
    let mut out = Vec::with_capacity(sheet_bin.len() + extra.len());
    let mut in_sheet_data = false;
    let mut current_row = 0u32;
    let mut replaced = false;

    loop {
        let record_start = cursor.position() as usize;
        let Some(id) = biff12_varint::read_record_id(&mut cursor)
            .ok()
            .flatten()
        else {
            break;
        };
        let id_end = cursor.position() as usize;
        let Some(len) = biff12_varint::read_record_len(&mut cursor)
            .ok()
            .flatten()
        else {
            break;
        };
        let len_end = cursor.position() as usize;

        let payload_start = len_end;
        let payload_end = payload_start + len as usize;
        let payload = &sheet_bin[payload_start..payload_end];
        cursor.set_position(payload_end as u64);

        match id {
            SHEETDATA => in_sheet_data = true,
            SHEETDATA_END => in_sheet_data = false,
            ROW if in_sheet_data => {
                if payload.len() >= 4 {
                    current_row = u32::from_le_bytes(payload[0..4].try_into().unwrap());
                }
            }
            _ => {}
        }

        let mut tweak = false;
        if !replaced && in_sheet_data && id == record_id && current_row == target_row && payload.len() >= 4
        {
            let col = u32::from_le_bytes(payload[0..4].try_into().unwrap());
            if col == target_col {
                tweak = true;
            }
        }

        if tweak {
            replaced = true;

            out.extend_from_slice(&sheet_bin[record_start..id_end]); // id varint bytes
            let new_len = len.saturating_add(extra.len() as u32);
            biff12_varint::write_record_len(&mut out, new_len).expect("write record len");
            out.extend_from_slice(payload);
            out.extend_from_slice(extra);
        } else {
            out.extend_from_slice(&sheet_bin[record_start..payload_end]);
        }
    }

    assert!(
        replaced,
        "expected to find and rewrite the cell record 0x{record_id:04X}"
    );
    out
}

fn cell_payload_for_id(
    sheet_bin: &[u8],
    target_row: u32,
    target_col: u32,
    record_id: u32,
) -> Option<(u32, Vec<u8>)> {
    let mut cursor = Cursor::new(sheet_bin);
    let mut in_sheet_data = false;
    let mut current_row = 0u32;
    loop {
        let id = biff12_varint::read_record_id(&mut cursor).ok().flatten()?;
        let len = biff12_varint::read_record_len(&mut cursor).ok().flatten()? as usize;
        let payload_start = cursor.position() as usize;
        let payload_end = payload_start.checked_add(len)?;
        let payload = sheet_bin.get(payload_start..payload_end)?;
        cursor.set_position(payload_end as u64);

        match id {
            SHEETDATA => in_sheet_data = true,
            SHEETDATA_END => in_sheet_data = false,
            ROW if in_sheet_data => {
                if payload.len() >= 4 {
                    current_row = u32::from_le_bytes(payload[0..4].try_into().unwrap());
                }
            }
            _ if id == record_id && in_sheet_data && current_row == target_row && payload.len() >= 4 => {
                let col = u32::from_le_bytes(payload[0..4].try_into().unwrap());
                if col == target_col {
                    return Some((len as u32, payload.to_vec()));
                }
            }
            _ => {}
        }
    }
}

fn decode_rk_number(raw: u32) -> f64 {
    let raw_i = raw as i32;
    let mut v = if raw_i & 0x02 != 0 {
        (raw_i >> 2) as f64
    } else {
        let shifted = raw & 0xFFFFFFFC;
        f64::from_bits((shifted as u64) << 32)
    };
    if raw_i & 0x01 != 0 {
        v /= 100.0;
    }
    v
}

#[test]
fn converting_value_cell_with_trailing_bytes_to_formula_requires_explicit_new_rgcb() {
    // Seed a sheet with a single FLOAT cell, then introduce malformed trailing bytes in that
    // record's payload by increasing the record length. Converting such a record to a formula
    // should require the caller to explicitly provide `CellEdit.new_rgcb` (even empty) so we don't
    // silently drop unknown bytes.
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number(0, 0, 1.0);
    let sheet_bin = read_sheet1_bin_from_fixture(&builder.build_bytes());
    let tweaked = append_trailing_bytes_to_cell_payload(&sheet_bin, 0, 0, FLOAT, &[0xAB]);

    let rgce = encode_rgce("=1+1").expect("encode formula");
    let edits_missing_rgcb = [CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Number(2.0),
        new_style: None,
        clear_formula: false,
        new_formula: Some(rgce.clone()),
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
    }];

    let err = patch_sheet_bin(&tweaked, &edits_missing_rgcb)
        .expect_err("expected InvalidInput when converting record with trailing bytes");
    assert_invalid_input_contains(err, "provide CellEdit.new_rgcb");

    let mut out = Vec::new();
    let err = patch_sheet_bin_streaming(Cursor::new(&tweaked), &mut out, &edits_missing_rgcb)
        .expect_err("expected InvalidInput when streaming convert record with trailing bytes");
    assert_invalid_input_contains(err, "provide CellEdit.new_rgcb");

    // Providing `new_rgcb` (even empty) should allow the conversion to proceed.
    let edits_with_rgcb = [CellEdit {
        new_rgcb: Some(Vec::new()),
        ..edits_missing_rgcb[0].clone()
    }];
    let patched_in_mem = patch_sheet_bin(&tweaked, &edits_with_rgcb).expect("patch_sheet_bin");

    let mut patched_stream = Vec::new();
    let changed =
        patch_sheet_bin_streaming(Cursor::new(&tweaked), &mut patched_stream, &edits_with_rgcb)
            .expect("patch_sheet_bin_streaming");
    assert!(changed);

    assert_eq!(patched_stream, patched_in_mem);
}

#[test]
fn patching_value_cell_with_trailing_bytes_preserves_unknown_payload_bytes() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number(0, 0, 1.0);
    let sheet_bin = read_sheet1_bin_from_fixture(&builder.build_bytes());
    let tweaked = append_trailing_bytes_to_cell_payload(&sheet_bin, 0, 0, FLOAT, &[0xAB]);

    let edits = [CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Number(2.0),
        new_style: None,
        clear_formula: false,
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
    }];

    let patched_in_mem = patch_sheet_bin(&tweaked, &edits).expect("patch_sheet_bin");

    let mut patched_stream = Vec::new();
    patch_sheet_bin_streaming(Cursor::new(&tweaked), &mut patched_stream, &edits)
        .expect("patch_sheet_bin_streaming");
    assert_eq!(patched_stream, patched_in_mem);

    let (len, payload) = cell_payload_for_id(&patched_in_mem, 0, 0, FLOAT)
        .expect("find patched FLOAT cell");
    assert_eq!(
        len,
        17,
        "expected patched FLOAT record to preserve trailing bytes length"
    );
    assert_eq!(payload.last().copied(), Some(0xAB));
    assert_eq!(
        f64::from_le_bytes(payload[8..16].try_into().unwrap()),
        2.0,
        "expected patched FLOAT value to be updated in place"
    );
}

#[test]
fn patching_bool_cell_with_trailing_bytes_preserves_unknown_payload_bytes() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_bool(0, 0, true);
    let sheet_bin = read_sheet1_bin_from_fixture(&builder.build_bytes());
    let tweaked = append_trailing_bytes_to_cell_payload(&sheet_bin, 0, 0, BOOL, &[0xAB]);

    let edits = [CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Bool(false),
        new_style: None,
        clear_formula: false,
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
    }];

    let patched_in_mem = patch_sheet_bin(&tweaked, &edits).expect("patch_sheet_bin");
    let mut patched_stream = Vec::new();
    patch_sheet_bin_streaming(Cursor::new(&tweaked), &mut patched_stream, &edits)
        .expect("patch_sheet_bin_streaming");
    assert_eq!(patched_stream, patched_in_mem);

    let (len, payload) =
        cell_payload_for_id(&patched_in_mem, 0, 0, BOOL).expect("find patched BOOL cell");
    assert_eq!(len, 10, "expected patched BOOL record to preserve trailing bytes length");
    assert_eq!(payload.last().copied(), Some(0xAB));
    assert_eq!(payload[8], 0, "expected patched BOOL payload");
}

#[test]
fn patching_error_cell_with_trailing_bytes_preserves_unknown_payload_bytes() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_error(0, 0, 0x07); // #DIV/0!
    let sheet_bin = read_sheet1_bin_from_fixture(&builder.build_bytes());
    let tweaked = append_trailing_bytes_to_cell_payload(&sheet_bin, 0, 0, BOOLERR, &[0xAB]);

    let edits = [CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Error(0x2A), // #N/A
        new_style: None,
        clear_formula: false,
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
    }];

    let patched_in_mem = patch_sheet_bin(&tweaked, &edits).expect("patch_sheet_bin");
    let mut patched_stream = Vec::new();
    patch_sheet_bin_streaming(Cursor::new(&tweaked), &mut patched_stream, &edits)
        .expect("patch_sheet_bin_streaming");
    assert_eq!(patched_stream, patched_in_mem);

    let (len, payload) =
        cell_payload_for_id(&patched_in_mem, 0, 0, BOOLERR).expect("find patched BOOLERR cell");
    assert_eq!(
        len, 10,
        "expected patched BOOLERR record to preserve trailing bytes length"
    );
    assert_eq!(payload.last().copied(), Some(0xAB));
    assert_eq!(payload[8], 0x2A, "expected patched BOOLERR payload");
}

#[test]
fn patching_blank_cell_with_trailing_bytes_preserves_unknown_payload_bytes() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_blank(0, 0);
    let sheet_bin = read_sheet1_bin_from_fixture(&builder.build_bytes());
    let tweaked = append_trailing_bytes_to_cell_payload(&sheet_bin, 0, 0, BLANK, &[0xAB]);

    let edits = [CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Blank,
        new_style: Some(1),
        clear_formula: false,
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
    }];

    let patched_in_mem = patch_sheet_bin(&tweaked, &edits).expect("patch_sheet_bin");
    let mut patched_stream = Vec::new();
    patch_sheet_bin_streaming(Cursor::new(&tweaked), &mut patched_stream, &edits)
        .expect("patch_sheet_bin_streaming");
    assert_eq!(patched_stream, patched_in_mem);

    let (len, payload) =
        cell_payload_for_id(&patched_in_mem, 0, 0, BLANK).expect("find patched BLANK cell");
    assert_eq!(
        len, 9,
        "expected patched BLANK record to preserve trailing bytes length"
    );
    assert_eq!(payload.last().copied(), Some(0xAB));
    assert_eq!(
        u32::from_le_bytes(payload[4..8].try_into().unwrap()),
        1
    );
}

#[test]
fn patching_rk_cell_with_trailing_bytes_preserves_unknown_payload_bytes() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number_rk(0, 0, 42.0);
    let sheet_bin = read_sheet1_bin_from_fixture(&builder.build_bytes());
    let tweaked = append_trailing_bytes_to_cell_payload(&sheet_bin, 0, 0, NUM, &[0xAB]);

    let edits = [CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Number(43.0),
        new_style: None,
        clear_formula: false,
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
    }];

    let patched_in_mem = patch_sheet_bin(&tweaked, &edits).expect("patch_sheet_bin");
    let mut patched_stream = Vec::new();
    patch_sheet_bin_streaming(Cursor::new(&tweaked), &mut patched_stream, &edits)
        .expect("patch_sheet_bin_streaming");
    assert_eq!(patched_stream, patched_in_mem);

    let (len, payload) =
        cell_payload_for_id(&patched_in_mem, 0, 0, NUM).expect("find patched RK/NUM cell");
    assert_eq!(
        len, 13,
        "expected patched RK record to preserve trailing bytes length"
    );
    assert_eq!(payload.last().copied(), Some(0xAB));
    let rk_raw = u32::from_le_bytes(payload[8..12].try_into().unwrap());
    assert_eq!(decode_rk_number(rk_raw), 43.0);
}
