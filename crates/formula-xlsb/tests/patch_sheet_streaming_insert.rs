use std::io::{Cursor, Read};

use formula_xlsb::{
    biff12_varint, patch_sheet_bin, patch_sheet_bin_streaming, CellEdit, CellValue,
};

mod fixture_builder;
use fixture_builder::XlsbFixtureBuilder;

fn read_sheet_bin(xlsb_bytes: Vec<u8>) -> Vec<u8> {
    let mut zip = zip::ZipArchive::new(Cursor::new(xlsb_bytes)).expect("open in-memory xlsb zip");
    let mut entry = zip
        .by_name("xl/worksheets/sheet1.bin")
        .expect("find sheet1.bin");
    let mut sheet_bin = Vec::with_capacity(entry.size() as usize);
    entry.read_to_end(&mut sheet_bin).expect("read sheet bytes");
    sheet_bin
}

fn read_dimension_bounds(sheet_bin: &[u8]) -> Option<(u32, u32, u32, u32)> {
    const DIMENSION: u32 = 0x0194;

    let mut cursor = Cursor::new(sheet_bin);
    loop {
        let id = biff12_varint::read_record_id(&mut cursor).ok().flatten()?;
        let len = biff12_varint::read_record_len(&mut cursor).ok().flatten()? as usize;
        let mut payload = vec![0u8; len];
        cursor.read_exact(&mut payload).ok()?;
        if id == DIMENSION && payload.len() >= 16 {
            let r1 = u32::from_le_bytes(payload[0..4].try_into().unwrap());
            let r2 = u32::from_le_bytes(payload[4..8].try_into().unwrap());
            let c1 = u32::from_le_bytes(payload[8..12].try_into().unwrap());
            let c2 = u32::from_le_bytes(payload[12..16].try_into().unwrap());
            return Some((r1, r2, c1, c2));
        }
    }
}

#[test]
fn patch_sheet_bin_streaming_insert_matches_in_memory_insert() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number(0, 0, 1.0); // A1 only
    let sheet_bin = read_sheet_bin(builder.build_bytes());

    let edits = [
        // Insert B1 and C5 (zero-based coords).
        CellEdit {
            row: 0,
            col: 1,
            new_value: CellValue::Number(42.0),
            new_formula: None,
            new_rgcb: None,
            shared_string_index: None,
        },
        CellEdit {
            row: 4,
            col: 2,
            new_value: CellValue::Number(99.0),
            new_formula: None,
            new_rgcb: None,
            shared_string_index: None,
        },
    ];

    let patched_in_mem = patch_sheet_bin(&sheet_bin, &edits).expect("patch_sheet_bin");

    let mut patched_stream = Vec::new();
    let changed = patch_sheet_bin_streaming(Cursor::new(&sheet_bin), &mut patched_stream, &edits)
        .expect("patch_sheet_bin_streaming");

    assert!(changed, "expected streaming patcher to report changes");
    assert_eq!(patched_stream, patched_in_mem);
}

#[test]
fn patch_sheet_bin_streaming_expands_dimension_for_inserts() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number(0, 0, 1.0);
    let sheet_bin = read_sheet_bin(builder.build_bytes());

    let edits = [
        CellEdit {
            row: 0,
            col: 1,
            new_value: CellValue::Number(42.0),
            new_formula: None,
            new_rgcb: None,
            shared_string_index: None,
        },
        CellEdit {
            row: 4,
            col: 2,
            new_value: CellValue::Number(99.0),
            new_formula: None,
            new_rgcb: None,
            shared_string_index: None,
        },
    ];

    let mut patched_stream = Vec::new();
    patch_sheet_bin_streaming(Cursor::new(&sheet_bin), &mut patched_stream, &edits)
        .expect("patch_sheet_bin_streaming");

    assert_eq!(read_dimension_bounds(&patched_stream), Some((0, 4, 0, 2)));
}

#[test]
fn patch_sheet_bin_streaming_noop_insertion_is_byte_identical() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number(0, 0, 1.0);
    let sheet_bin = read_sheet_bin(builder.build_bytes());

    let edits = [CellEdit {
        row: 4,
        col: 2,
        new_value: CellValue::Blank,
        new_formula: None,
        new_rgcb: None,
        shared_string_index: None,
    }];

    let mut patched_stream = Vec::new();
    let changed = patch_sheet_bin_streaming(Cursor::new(&sheet_bin), &mut patched_stream, &edits)
        .expect("patch_sheet_bin_streaming");

    assert!(!changed, "expected no-op insertion to report unchanged");
    assert_eq!(patched_stream, sheet_bin);
}
