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
    // Record IDs follow the conventions used by `formula-xlsb`'s BIFF12 reader.
    // See `crates/formula-xlsb/src/parser.rs` (`biff12` module).
    const DIMENSION: u32 = 0x0094;

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

fn move_dimension_record_to_end(sheet_bin: &[u8]) -> Vec<u8> {
    // Record IDs follow the conventions used by `formula-xlsb`'s BIFF12 reader.
    const DIMENSION: u32 = 0x0094;
    const WORKSHEET_END: u32 = 0x0082;

    let mut cursor = Cursor::new(sheet_bin);
    let mut ranges: Vec<(u32, usize, usize)> = Vec::new();
    loop {
        let start = cursor.position() as usize;
        let Some(id) = biff12_varint::read_record_id(&mut cursor).ok().flatten() else {
            break;
        };
        let Some(len) = biff12_varint::read_record_len(&mut cursor).ok().flatten() else {
            break;
        };
        let payload_start = cursor.position() as usize;
        let payload_end = payload_start + len as usize;
        cursor.set_position(payload_end as u64);
        ranges.push((id, start, payload_end));
    }

    let mut out = Vec::with_capacity(sheet_bin.len());
    let mut dims: Vec<&[u8]> = Vec::new();

    for (id, start, end) in ranges {
        let bytes = &sheet_bin[start..end];
        if id == DIMENSION {
            dims.push(bytes);
            continue;
        }
        if id == WORKSHEET_END {
            for dim in &dims {
                out.extend_from_slice(dim);
            }
            dims.clear();
        }
        out.extend_from_slice(bytes);
    }

    for dim in dims {
        out.extend_from_slice(dim);
    }

    out
}

fn cell_coords_in_stream_order(sheet_bin: &[u8]) -> Vec<(u32, u32)> {
    // Record IDs follow the conventions used by `formula-xlsb`'s BIFF12 reader.
    const SHEETDATA: u32 = 0x0091;
    const SHEETDATA_END: u32 = 0x0092;
    const ROW: u32 = 0x0000;

    let mut cursor = Cursor::new(sheet_bin);
    let mut in_sheet_data = false;
    let mut current_row = 0u32;
    let mut coords = Vec::new();

    loop {
        let id = match biff12_varint::read_record_id(&mut cursor).ok().flatten() {
            Some(id) => id,
            None => break,
        };
        let len = match biff12_varint::read_record_len(&mut cursor).ok().flatten() {
            Some(len) => len as usize,
            None => break,
        };
        let mut payload = vec![0u8; len];
        if cursor.read_exact(&mut payload).is_err() {
            break;
        }

        match id {
            SHEETDATA => in_sheet_data = true,
            SHEETDATA_END => in_sheet_data = false,
            ROW if in_sheet_data => {
                if payload.len() >= 4 {
                    current_row = u32::from_le_bytes(payload[0..4].try_into().unwrap());
                }
            }
            _ if in_sheet_data => {
                if payload.len() >= 4 {
                    let col = u32::from_le_bytes(payload[0..4].try_into().unwrap());
                    coords.push((current_row, col));
                }
            }
            _ => {}
        }
    }

    coords
}

fn find_cell_record(sheet_bin: &[u8], target_row: u32, target_col: u32) -> Option<(u32, Vec<u8>)> {
    // Record IDs follow the conventions used by `formula-xlsb`'s BIFF12 reader.
    const SHEETDATA: u32 = 0x0091;
    const SHEETDATA_END: u32 = 0x0092;
    const ROW: u32 = 0x0000;

    let mut cursor = Cursor::new(sheet_bin);
    let mut in_sheet_data = false;
    let mut current_row = 0u32;

    loop {
        let id = match biff12_varint::read_record_id(&mut cursor).ok().flatten() {
            Some(id) => id,
            None => break,
        };
        let len = match biff12_varint::read_record_len(&mut cursor).ok().flatten() {
            Some(len) => len as usize,
            None => break,
        };
        let mut payload = vec![0u8; len];
        if cursor.read_exact(&mut payload).is_err() {
            break;
        }

        match id {
            SHEETDATA => in_sheet_data = true,
            SHEETDATA_END => in_sheet_data = false,
            ROW if in_sheet_data => {
                if payload.len() >= 4 {
                    current_row = u32::from_le_bytes(payload[0..4].try_into().unwrap());
                }
            }
            _ if in_sheet_data => {
                if payload.len() >= 4 {
                    let col = u32::from_le_bytes(payload[0..4].try_into().unwrap());
                    if current_row == target_row && col == target_col {
                        return Some((id, payload));
                    }
                }
            }
            _ => {}
        }
    }

    None
}

fn rewrite_dimension_len_as_two_byte_varint(sheet_bin: &[u8]) -> Vec<u8> {
    const DIMENSION: u32 = 0x0094;

    let mut cursor = Cursor::new(sheet_bin);
    let mut out = Vec::with_capacity(sheet_bin.len() + 4);

    loop {
        let record_start = cursor.position() as usize;
        let Some(id) = biff12_varint::read_record_id(&mut cursor).ok().flatten() else {
            break;
        };
        let id_end = cursor.position() as usize;
        let Some(len) = biff12_varint::read_record_len(&mut cursor).ok().flatten() else {
            break;
        };
        let len_start = id_end;
        let len_end = cursor.position() as usize;

        let payload_start = len_end;
        let payload_end = payload_start + len as usize;
        cursor.set_position(payload_end as u64);

        out.extend_from_slice(&sheet_bin[record_start..id_end]); // id varint bytes
        if id == DIMENSION && len == 16 {
            // Non-canonical, but valid, 2-byte LEB128 encoding for length 16.
            out.extend_from_slice(&[0x90, 0x00]);
        } else {
            out.extend_from_slice(&sheet_bin[len_start..len_end]); // original len varint bytes
        }
        out.extend_from_slice(&sheet_bin[payload_start..payload_end]);
    }

    out
}

fn rewrite_cell_isst_header_as_two_byte_varints(
    sheet_bin: &[u8],
    target_row: u32,
    target_col: u32,
) -> Vec<u8> {
    // Record IDs follow the conventions used by `formula-xlsb`'s BIFF12 reader.
    const SHEETDATA: u32 = 0x0091;
    const SHEETDATA_END: u32 = 0x0092;
    const ROW: u32 = 0x0000;
    const STRING: u32 = 0x0007;

    let mut cursor = Cursor::new(sheet_bin);
    let mut out = Vec::with_capacity(sheet_bin.len() + 4);
    let mut in_sheet_data = false;
    let mut current_row = 0u32;

    loop {
        let record_start = cursor.position() as usize;
        let Some(id) = biff12_varint::read_record_id(&mut cursor).ok().flatten() else {
            break;
        };
        let id_end = cursor.position() as usize;
        let Some(len) = biff12_varint::read_record_len(&mut cursor).ok().flatten() else {
            break;
        };
        let len_start = id_end;
        let len_end = cursor.position() as usize;

        let payload_start = len_end;
        let payload_end = payload_start + len as usize;
        cursor.set_position(payload_end as u64);
        let payload = &sheet_bin[payload_start..payload_end];

        let mut tweak = false;
        match id {
            SHEETDATA => in_sheet_data = true,
            SHEETDATA_END => in_sheet_data = false,
            ROW if in_sheet_data => {
                if payload.len() >= 4 {
                    current_row = u32::from_le_bytes(payload[0..4].try_into().unwrap());
                }
            }
            STRING if in_sheet_data => {
                if payload.len() >= 4 {
                    let col = u32::from_le_bytes(payload[0..4].try_into().unwrap());
                    if current_row == target_row && col == target_col {
                        tweak = true;
                    }
                }
            }
            _ => {}
        }

        if tweak {
            // Non-canonical, but valid, 2-byte varint encodings for:
            // - record id = 7 (`BrtCellIsst`)
            // - payload length = 12 (`[col][style][isst]`)
            out.extend_from_slice(&[0x87, 0x00]);
            if len == 12 {
                out.extend_from_slice(&[0x8C, 0x00]);
            } else {
                out.extend_from_slice(&sheet_bin[len_start..len_end]);
            }
        } else {
            out.extend_from_slice(&sheet_bin[record_start..id_end]); // id varint bytes
            out.extend_from_slice(&sheet_bin[len_start..len_end]); // original len varint bytes
        }
        out.extend_from_slice(payload);
    }

    out
}

fn rewrite_cell_header_as_two_byte_varints(sheet_bin: &[u8], target_row: u32, target_col: u32) -> Vec<u8> {
    // Record IDs follow the conventions used by `formula-xlsb`'s BIFF12 reader.
    const SHEETDATA: u32 = 0x0091;
    const SHEETDATA_END: u32 = 0x0092;
    const ROW: u32 = 0x0000;

    let mut cursor = Cursor::new(sheet_bin);
    let mut out = Vec::with_capacity(sheet_bin.len() + 4);
    let mut in_sheet_data = false;
    let mut current_row = 0u32;

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
        let len_start = id_end;
        let len_end = cursor.position() as usize;

        let payload_start = len_end;
        let payload_end = payload_start + len as usize;
        cursor.set_position(payload_end as u64);
        let payload = &sheet_bin[payload_start..payload_end];

        let mut tweak = false;
        match id {
            SHEETDATA => in_sheet_data = true,
            SHEETDATA_END => in_sheet_data = false,
            ROW if in_sheet_data => {
                if payload.len() >= 4 {
                    current_row = u32::from_le_bytes(payload[0..4].try_into().unwrap());
                }
            }
            _ if in_sheet_data => {
                if payload.len() >= 4 {
                    let col = u32::from_le_bytes(payload[0..4].try_into().unwrap());
                    if current_row == target_row && col == target_col {
                        tweak = true;
                    }
                }
            }
            _ => {}
        }

        if tweak {
            assert!(
                id < 0x80 && len < 0x80,
                "test helper only supports rewriting 1-byte varints"
            );
            // Non-canonical, but valid, 2-byte LEB128 varints for values < 128.
            out.extend_from_slice(&[(id as u8) | 0x80, 0x00]);
            out.extend_from_slice(&[(len as u8) | 0x80, 0x00]);
        } else {
            out.extend_from_slice(&sheet_bin[record_start..id_end]); // id varint bytes
            out.extend_from_slice(&sheet_bin[len_start..len_end]); // original len varint bytes
        }
        out.extend_from_slice(payload);
    }

    out
}

fn dimension_header_raw(sheet_bin: &[u8]) -> Option<(Vec<u8>, Vec<u8>)> {
    const DIMENSION: u32 = 0x0094;

    let mut cursor = Cursor::new(sheet_bin);
    loop {
        let record_start = cursor.position() as usize;
        let id = biff12_varint::read_record_id(&mut cursor).ok().flatten()?;
        let id_end = cursor.position() as usize;
        let len = biff12_varint::read_record_len(&mut cursor).ok().flatten()? as usize;
        let len_end = cursor.position() as usize;
        let payload_start = len_end;
        let payload_end = payload_start + len;
        cursor.set_position(payload_end as u64);

        if id == DIMENSION {
            return Some((
                sheet_bin[record_start..id_end].to_vec(),
                sheet_bin[id_end..len_end].to_vec(),
            ));
        }
    }
}

fn cell_header_raw(
    sheet_bin: &[u8],
    target_row: u32,
    target_col: u32,
) -> Option<(Vec<u8>, Vec<u8>)> {
    // Record IDs follow the conventions used by `formula-xlsb`'s BIFF12 reader.
    const SHEETDATA: u32 = 0x0091;
    const SHEETDATA_END: u32 = 0x0092;
    const ROW: u32 = 0x0000;

    let mut cursor = Cursor::new(sheet_bin);
    let mut in_sheet_data = false;
    let mut current_row = 0u32;

    loop {
        let record_start = cursor.position() as usize;
        let id = biff12_varint::read_record_id(&mut cursor).ok().flatten()?;
        let id_end = cursor.position() as usize;
        let len = biff12_varint::read_record_len(&mut cursor).ok().flatten()? as usize;
        let len_end = cursor.position() as usize;

        let payload_start = len_end;
        let payload_end = payload_start + len;
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
            _ if in_sheet_data => {
                if payload.len() >= 4 {
                    let col = u32::from_le_bytes(payload[0..4].try_into().unwrap());
                    if current_row == target_row && col == target_col {
                        return Some((
                            sheet_bin[record_start..id_end].to_vec(),
                            sheet_bin[id_end..len_end].to_vec(),
                        ));
                    }
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
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
        },
        CellEdit {
            row: 4,
            col: 2,
            new_value: CellValue::Number(99.0),
            new_formula: None,
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
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
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
        },
        CellEdit {
            row: 4,
            col: 2,
            new_value: CellValue::Number(99.0),
            new_formula: None,
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
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
        new_formula_flags: None,
        shared_string_index: None,
        new_style: None,
    }];

    let mut patched_stream = Vec::new();
    let changed = patch_sheet_bin_streaming(Cursor::new(&sheet_bin), &mut patched_stream, &edits)
        .expect("patch_sheet_bin_streaming");

    assert!(!changed, "expected no-op insertion to report unchanged");
    assert_eq!(patched_stream, sheet_bin);
}

#[test]
fn patch_sheet_bin_streaming_handles_dimension_after_sheetdata() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number(0, 0, 1.0);
    let sheet_bin = read_sheet_bin(builder.build_bytes());
    let moved = move_dimension_record_to_end(&sheet_bin);

    let edits = [
        CellEdit {
            row: 0,
            col: 1,
            new_value: CellValue::Number(42.0),
            new_formula: None,
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
        },
        CellEdit {
            row: 4,
            col: 2,
            new_value: CellValue::Number(99.0),
            new_formula: None,
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
        },
    ];

    let patched_in_mem = patch_sheet_bin(&moved, &edits).expect("patch_sheet_bin");

    let mut patched_stream = Vec::new();
    let changed = patch_sheet_bin_streaming(Cursor::new(&moved), &mut patched_stream, &edits)
        .expect("patch_sheet_bin_streaming");

    assert!(changed, "expected streaming patcher to report changes");
    assert_eq!(patched_stream, patched_in_mem);
    assert_eq!(read_dimension_bounds(&patched_stream), Some((0, 4, 0, 2)));
}

#[test]
fn patch_sheet_bin_streaming_inserts_cells_in_column_order() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number(0, 5, 1.0); // F1 only
    let sheet_bin = read_sheet_bin(builder.build_bytes());

    let edits = [
        // Insert B1 (before the existing F1 cell).
        CellEdit {
            row: 0,
            col: 1,
            new_value: CellValue::Number(42.0),
            new_formula: None,
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
        },
        // Insert H1 (after the existing F1 cell).
        CellEdit {
            row: 0,
            col: 7,
            new_value: CellValue::Number(100.0),
            new_formula: None,
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
        },
    ];

    let mut patched_stream = Vec::new();
    patch_sheet_bin_streaming(Cursor::new(&sheet_bin), &mut patched_stream, &edits)
        .expect("patch_sheet_bin_streaming");

    assert_eq!(
        cell_coords_in_stream_order(&patched_stream),
        vec![(0, 1), (0, 5), (0, 7)]
    );
}

#[test]
fn patch_sheet_bin_streaming_is_lossless_for_noop_value_edit() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number(0, 0, 1.0);
    let sheet_bin = read_sheet_bin(builder.build_bytes());

    let edits = [CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Number(1.0),
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
        new_style: None,
    }];

    let mut patched_stream = Vec::new();
    let changed = patch_sheet_bin_streaming(Cursor::new(&sheet_bin), &mut patched_stream, &edits)
        .expect("patch_sheet_bin_streaming");

    assert!(!changed, "expected no-op value edit to report unchanged");
    assert_eq!(patched_stream, sheet_bin);
}

#[test]
fn patch_sheet_bin_streaming_inserts_text_cell_as_shared_string_when_isst_provided() {
    const STRING: u32 = 0x0007;

    let mut builder = XlsbFixtureBuilder::new();
    builder.add_shared_string("Hello");
    builder.set_cell_number(0, 0, 1.0);
    let sheet_bin = read_sheet_bin(builder.build_bytes());

    let edits = [CellEdit {
        row: 4,
        col: 2,
        new_value: CellValue::Text("Hello".to_string()),
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: Some(0),
        new_style: None,
    }];

    let patched_in_mem = patch_sheet_bin(&sheet_bin, &edits).expect("patch_sheet_bin");

    let mut patched_stream = Vec::new();
    patch_sheet_bin_streaming(Cursor::new(&sheet_bin), &mut patched_stream, &edits)
        .expect("patch_sheet_bin_streaming");

    assert_eq!(patched_stream, patched_in_mem);

    let (id, payload) = find_cell_record(&patched_stream, 4, 2).expect("find inserted cell");
    assert_eq!(id, STRING, "expected BrtCellIsst/STRING record id");
    assert_eq!(
        u32::from_le_bytes(payload[8..12].try_into().unwrap()),
        0,
        "expected inserted cell to reference shared string index 0"
    );
}

#[test]
fn patch_sheet_bin_streaming_inserts_text_cell_as_inline_string_when_isst_missing() {
    const CELL_ST: u32 = 0x0006;

    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number(0, 0, 1.0);
    let sheet_bin = read_sheet_bin(builder.build_bytes());

    let edits = [CellEdit {
        row: 4,
        col: 2,
        new_value: CellValue::Text("Hello".to_string()),
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
        new_style: None,
    }];

    let patched_in_mem = patch_sheet_bin(&sheet_bin, &edits).expect("patch_sheet_bin");

    let mut patched_stream = Vec::new();
    patch_sheet_bin_streaming(Cursor::new(&sheet_bin), &mut patched_stream, &edits)
        .expect("patch_sheet_bin_streaming");

    assert_eq!(patched_stream, patched_in_mem);

    let (id, _payload) = find_cell_record(&patched_stream, 4, 2).expect("find inserted cell");
    assert_eq!(id, CELL_ST, "expected BrtCellSt/CELL_ST record id");
}

#[test]
fn patch_sheet_bin_streaming_is_lossless_for_noop_formula_edit() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_sheet_name("NoopFormula");

    // PtgInt 1 (formula token stream for literal `1`).
    let rgce = vec![0x1E, 0x01, 0x00];
    let extra = vec![0xDE, 0xAD, 0xBE, 0xEF];
    builder.set_cell_formula_num(0, 0, 1.0, rgce, extra);

    let sheet_bin = read_sheet_bin(builder.build_bytes());
    let edits = [CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Number(1.0),
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
        new_style: None,
    }];

    let mut patched_stream = Vec::new();
    let changed = patch_sheet_bin_streaming(Cursor::new(&sheet_bin), &mut patched_stream, &edits)
        .expect("patch_sheet_bin_streaming");

    assert!(!changed, "expected no-op formula edit to report unchanged");
    assert_eq!(patched_stream, sheet_bin);
    assert_eq!(
        patch_sheet_bin(&sheet_bin, &edits).expect("patch_sheet_bin"),
        sheet_bin
    );
}

#[test]
fn patch_sheet_bin_streaming_can_insert_into_empty_sheet() {
    let builder = XlsbFixtureBuilder::new();
    let sheet_bin = read_sheet_bin(builder.build_bytes());

    let edits = [CellEdit {
        row: 5,
        col: 3,
        new_value: CellValue::Number(123.0),
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
        new_style: None,
    }];

    let patched_in_mem = patch_sheet_bin(&sheet_bin, &edits).expect("patch_sheet_bin");

    let mut patched_stream = Vec::new();
    let changed = patch_sheet_bin_streaming(Cursor::new(&sheet_bin), &mut patched_stream, &edits)
        .expect("patch_sheet_bin_streaming");

    assert!(changed);
    assert_eq!(patched_stream, patched_in_mem);
    assert_eq!(cell_coords_in_stream_order(&patched_stream), vec![(5, 3)]);
    assert_eq!(read_dimension_bounds(&patched_stream), Some((0, 5, 0, 3)));
}

#[test]
fn patch_sheet_bin_streaming_noop_insertions_in_empty_sheet_are_lossless() {
    let builder = XlsbFixtureBuilder::new();
    let sheet_bin = read_sheet_bin(builder.build_bytes());

    let edits = [
        CellEdit {
            row: 5,
            col: 3,
            new_value: CellValue::Blank,
            new_formula: None,
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
        },
        CellEdit {
            row: 5,
            col: 4,
            new_value: CellValue::Blank,
            new_formula: None,
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
        },
    ];

    let patched_in_mem = patch_sheet_bin(&sheet_bin, &edits).expect("patch_sheet_bin");
    assert_eq!(patched_in_mem, sheet_bin);

    let mut patched_stream = Vec::new();
    let changed = patch_sheet_bin_streaming(Cursor::new(&sheet_bin), &mut patched_stream, &edits)
        .expect("patch_sheet_bin_streaming");

    assert!(!changed);
    assert_eq!(patched_stream, sheet_bin);
}

#[test]
fn patch_sheet_bin_streaming_inserts_missing_rows_before_first_row() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number(5, 0, 1.0);
    let sheet_bin = read_sheet_bin(builder.build_bytes());

    let edits = [CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Number(2.0),
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
        new_style: None,
    }];

    let patched_in_mem = patch_sheet_bin(&sheet_bin, &edits).expect("patch_sheet_bin");

    let mut patched_stream = Vec::new();
    patch_sheet_bin_streaming(Cursor::new(&sheet_bin), &mut patched_stream, &edits)
        .expect("patch_sheet_bin_streaming");

    assert_eq!(patched_stream, patched_in_mem);
    assert_eq!(
        cell_coords_in_stream_order(&patched_stream),
        vec![(0, 0), (5, 0)]
    );
    assert_eq!(read_dimension_bounds(&patched_stream), Some((0, 5, 0, 0)));
}

#[test]
fn patch_sheet_bin_streaming_inserts_missing_rows_between_existing_rows() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number(0, 0, 1.0);
    builder.set_cell_number(5, 0, 3.0);
    let sheet_bin = read_sheet_bin(builder.build_bytes());

    let edits = [CellEdit {
        row: 3,
        col: 0,
        new_value: CellValue::Number(2.0),
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
        new_style: None,
    }];

    let patched_in_mem = patch_sheet_bin(&sheet_bin, &edits).expect("patch_sheet_bin");

    let mut patched_stream = Vec::new();
    patch_sheet_bin_streaming(Cursor::new(&sheet_bin), &mut patched_stream, &edits)
        .expect("patch_sheet_bin_streaming");

    assert_eq!(patched_stream, patched_in_mem);
    assert_eq!(
        cell_coords_in_stream_order(&patched_stream),
        vec![(0, 0), (3, 0), (5, 0)]
    );
    assert_eq!(read_dimension_bounds(&patched_stream), Some((0, 5, 0, 0)));
}

#[test]
fn patch_sheet_bin_streaming_preserves_dimension_header_varint_bytes() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number(0, 0, 1.0);
    let sheet_bin = read_sheet_bin(builder.build_bytes());
    let tweaked = rewrite_dimension_len_as_two_byte_varint(&sheet_bin);

    let Some((id_raw, len_raw)) = dimension_header_raw(&tweaked) else {
        panic!("expected DIMENSION record");
    };
    assert_eq!(
        len_raw,
        vec![0x90, 0x00],
        "expected non-canonical len varint"
    );

    let edits = [CellEdit {
        row: 5,
        col: 3,
        new_value: CellValue::Number(123.0),
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
        new_style: None,
    }];

    let patched_in_mem = patch_sheet_bin(&tweaked, &edits).expect("patch_sheet_bin");

    let mut patched_stream = Vec::new();
    patch_sheet_bin_streaming(Cursor::new(&tweaked), &mut patched_stream, &edits)
        .expect("patch_sheet_bin_streaming");

    assert_eq!(patched_stream, patched_in_mem);

    let Some((patched_id_raw, patched_len_raw)) = dimension_header_raw(&patched_stream) else {
        panic!("expected DIMENSION record");
    };
    assert_eq!(patched_id_raw, id_raw);
    assert_eq!(patched_len_raw, len_raw);
}

#[test]
fn patch_sheet_bin_streaming_preserves_cell_isst_header_varint_bytes_when_patching_isst() {
    const STRING: u32 = 0x0007;

    let mut builder = XlsbFixtureBuilder::new();
    builder.add_shared_string("Hello");
    builder.add_shared_string("World");
    builder.set_cell_sst(0, 0, 0);
    let sheet_bin = read_sheet_bin(builder.build_bytes());
    let tweaked = rewrite_cell_isst_header_as_two_byte_varints(&sheet_bin, 0, 0);

    let Some((id_raw, len_raw)) = cell_header_raw(&tweaked, 0, 0) else {
        panic!("expected cell record");
    };
    assert_eq!(id_raw, vec![0x87, 0x00], "expected non-canonical id varint");
    assert_eq!(
        len_raw,
        vec![0x8C, 0x00],
        "expected non-canonical len varint"
    );

    let edits = [CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Text("World".to_string()),
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: Some(1),
        new_style: None,
    }];

    let patched_in_mem = patch_sheet_bin(&tweaked, &edits).expect("patch_sheet_bin");

    let mut patched_stream = Vec::new();
    let changed = patch_sheet_bin_streaming(Cursor::new(&tweaked), &mut patched_stream, &edits)
        .expect("patch_sheet_bin_streaming");
    assert!(changed);

    assert_eq!(patched_stream, patched_in_mem);

    let Some((patched_id_raw, patched_len_raw)) = cell_header_raw(&patched_stream, 0, 0) else {
        panic!("expected cell record");
    };
    assert_eq!(patched_id_raw, id_raw);
    assert_eq!(patched_len_raw, len_raw);

    let (id, payload) = find_cell_record(&patched_stream, 0, 0).expect("find patched cell");
    assert_eq!(id, STRING, "expected BrtCellIsst/STRING record id");
    assert_eq!(
        u32::from_le_bytes(payload[8..12].try_into().unwrap()),
        1,
        "expected patched cell to reference shared string index 1"
    );
}

#[test]
fn patch_sheet_bin_streaming_preserves_cell_header_varint_bytes_for_fixed_size_value_edits() {
    const FLOAT: u32 = 0x0005;

    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number(0, 0, 1.0);
    let sheet_bin = read_sheet_bin(builder.build_bytes());
    let tweaked = rewrite_cell_header_as_two_byte_varints(&sheet_bin, 0, 0);

    let Some((id_raw, len_raw)) = cell_header_raw(&tweaked, 0, 0) else {
        panic!("expected cell record");
    };
    assert_eq!(id_raw, vec![0x85, 0x00], "expected non-canonical id varint for FLOAT");
    assert_eq!(len_raw, vec![0x90, 0x00], "expected non-canonical len varint for FLOAT payload");

    let edits = [CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Number(2.0),
        new_style: None,
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
    }];

    let patched_in_mem = patch_sheet_bin(&tweaked, &edits).expect("patch_sheet_bin");

    let mut patched_stream = Vec::new();
    let changed = patch_sheet_bin_streaming(Cursor::new(&tweaked), &mut patched_stream, &edits)
        .expect("patch_sheet_bin_streaming");
    assert!(changed);
    assert_eq!(patched_stream, patched_in_mem);

    let Some((patched_id_raw, patched_len_raw)) = cell_header_raw(&patched_in_mem, 0, 0) else {
        panic!("expected cell record");
    };
    assert_eq!(patched_id_raw, id_raw);
    assert_eq!(patched_len_raw, len_raw);

    let (id, payload) = find_cell_record(&patched_in_mem, 0, 0).expect("find patched cell");
    assert_eq!(id, FLOAT);
    assert_eq!(
        f64::from_le_bytes(payload[8..16].try_into().unwrap()),
        2.0,
        "expected patched FLOAT payload"
    );
}

#[test]
fn patch_sheet_bin_streaming_preserves_formula_header_varint_bytes_when_payload_len_unchanged() {
    const FORMULA_FLOAT: u32 = 0x0009;

    let mut builder = XlsbFixtureBuilder::new();
    let mut rgce = Vec::new();
    fixture_builder::rgce::push_int(&mut rgce, 1);
    builder.set_cell_formula_num(0, 0, 1.0, rgce, vec![]);
    let sheet_bin = read_sheet_bin(builder.build_bytes());
    let tweaked = rewrite_cell_header_as_two_byte_varints(&sheet_bin, 0, 0);

    let Some((id_raw, len_raw)) = cell_header_raw(&tweaked, 0, 0) else {
        panic!("expected cell record");
    };
    assert_eq!(
        id_raw,
        vec![0x89, 0x00],
        "expected non-canonical id varint for FORMULA_FLOAT"
    );
    // Payload length is 22 + cce (PtgInt: 3 bytes) = 25 (0x19) => 0x99 0x00.
    assert_eq!(
        len_raw,
        vec![0x99, 0x00],
        "expected non-canonical len varint for formula payload"
    );

    let edits = [CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Number(2.0),
        new_style: None,
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
    }];

    let patched_in_mem = patch_sheet_bin(&tweaked, &edits).expect("patch_sheet_bin");

    let mut patched_stream = Vec::new();
    let changed = patch_sheet_bin_streaming(Cursor::new(&tweaked), &mut patched_stream, &edits)
        .expect("patch_sheet_bin_streaming");
    assert!(changed);
    assert_eq!(patched_stream, patched_in_mem);

    let Some((patched_id_raw, patched_len_raw)) = cell_header_raw(&patched_in_mem, 0, 0) else {
        panic!("expected cell record");
    };
    assert_eq!(patched_id_raw, id_raw);
    assert_eq!(patched_len_raw, len_raw);

    let (id, payload) = find_cell_record(&patched_in_mem, 0, 0).expect("find patched cell");
    assert_eq!(id, FORMULA_FLOAT);
    assert_eq!(
        f64::from_le_bytes(payload[8..16].try_into().unwrap()),
        2.0,
        "expected patched cached value in FORMULA_FLOAT payload"
    );
}

#[test]
fn patch_sheet_bin_streaming_insert_formula_with_rgcb_matches_in_memory() {
    const FORMULA_FLOAT: u32 = 0x0009;

    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number(0, 0, 1.0);
    let sheet_bin = read_sheet_bin(builder.build_bytes());

    let rgce = vec![0x20, 0, 0, 0, 0, 0, 0, 0]; // PtgArray placeholder
    let rgcb = vec![0xDE, 0xAD, 0xBE, 0xEF];
    let edits = [CellEdit {
        row: 4,
        col: 2,
        new_value: CellValue::Number(6.0),
        new_formula: Some(rgce.clone()),
        new_rgcb: Some(rgcb.clone()),
        new_formula_flags: None,
        shared_string_index: None,
        new_style: None,
    }];

    let patched_in_mem = patch_sheet_bin(&sheet_bin, &edits).expect("patch_sheet_bin");

    let mut patched_stream = Vec::new();
    patch_sheet_bin_streaming(Cursor::new(&sheet_bin), &mut patched_stream, &edits)
        .expect("patch_sheet_bin_streaming");

    assert_eq!(patched_stream, patched_in_mem);

    let (id, payload) = find_cell_record(&patched_stream, 4, 2).expect("find inserted cell");
    assert_eq!(
        id, FORMULA_FLOAT,
        "expected BrtFmlaNum/FORMULA_FLOAT record id"
    );

    let cce = u32::from_le_bytes(payload[18..22].try_into().unwrap()) as usize;
    assert_eq!(payload[22..22 + cce], rgce);
    assert_eq!(payload[22 + cce..], rgcb);
}

#[test]
fn patch_sheet_bin_streaming_can_patch_formula_rgcb_bytes() {
    const FORMULA_FLOAT: u32 = 0x0009;

    let mut builder = XlsbFixtureBuilder::new();
    builder.set_sheet_name("PatchRgcb");
    let rgce = vec![0x20, 0, 0, 0, 0, 0, 0, 0]; // PtgArray placeholder
    let rgcb = vec![0xAA, 0xBB, 0xCC];
    builder.set_cell_formula_num(0, 0, 6.0, rgce.clone(), rgcb);
    let sheet_bin = read_sheet_bin(builder.build_bytes());

    let new_rgcb = vec![0x11, 0x22, 0x33, 0x44];
    let edits = [CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Number(6.0),
        new_formula: None,
        new_rgcb: Some(new_rgcb.clone()),
        new_formula_flags: None,
        shared_string_index: None,
        new_style: None,
    }];

    let patched_in_mem = patch_sheet_bin(&sheet_bin, &edits).expect("patch_sheet_bin");

    let mut patched_stream = Vec::new();
    patch_sheet_bin_streaming(Cursor::new(&sheet_bin), &mut patched_stream, &edits)
        .expect("patch_sheet_bin_streaming");

    assert_eq!(patched_stream, patched_in_mem);

    let (id, payload) = find_cell_record(&patched_stream, 0, 0).expect("find patched cell");
    assert_eq!(
        id, FORMULA_FLOAT,
        "expected BrtFmlaNum/FORMULA_FLOAT record id"
    );

    let cce = u32::from_le_bytes(payload[18..22].try_into().unwrap()) as usize;
    assert_eq!(payload[22..22 + cce], rgce);
    assert_eq!(payload[22 + cce..], new_rgcb);
}

#[test]
fn patch_sheet_bin_streaming_inserts_bool_and_error_cells_matches_in_memory() {
    const BOOLERR: u32 = 0x0003;
    const BOOL: u32 = 0x0004;

    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number(0, 0, 1.0);
    let sheet_bin = read_sheet_bin(builder.build_bytes());

    let edits = [
        CellEdit {
            row: 0,
            col: 1,
            new_value: CellValue::Bool(true),
            new_formula: None,
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
        },
        CellEdit {
            row: 4,
            col: 2,
            new_value: CellValue::Error(0x07),
            new_formula: None,
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
        },
    ];

    let patched_in_mem = patch_sheet_bin(&sheet_bin, &edits).expect("patch_sheet_bin");

    let mut patched_stream = Vec::new();
    patch_sheet_bin_streaming(Cursor::new(&sheet_bin), &mut patched_stream, &edits)
        .expect("patch_sheet_bin_streaming");

    assert_eq!(patched_stream, patched_in_mem);

    let (id, payload) = find_cell_record(&patched_stream, 0, 1).expect("find inserted bool cell");
    assert_eq!(id, BOOL, "expected BrtCellBool/BOOL record id");
    assert_eq!(payload[8], 1);

    let (id, payload) = find_cell_record(&patched_stream, 4, 2).expect("find inserted error cell");
    assert_eq!(id, BOOLERR, "expected BrtCellBoolErr/BOOLERR record id");
    assert_eq!(payload[8], 0x07);
}

#[test]
fn patch_sheet_bin_streaming_inserts_formula_bool_and_error_cells_matches_in_memory() {
    const FORMULA_BOOL: u32 = 0x000A;
    const FORMULA_BOOLERR: u32 = 0x000B;

    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number(0, 0, 1.0);
    let sheet_bin = read_sheet_bin(builder.build_bytes());

    let rgce_bool_true = vec![0x1D, 0x01]; // PtgBool TRUE
    let rgce_err_div0 = vec![0x1C, 0x07]; // PtgErr #DIV/0!

    let edits = [
        CellEdit {
            row: 0,
            col: 1,
            new_value: CellValue::Bool(true),
            new_formula: Some(rgce_bool_true.clone()),
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
        },
        CellEdit {
            row: 4,
            col: 2,
            new_value: CellValue::Error(0x07),
            new_formula: Some(rgce_err_div0.clone()),
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
        },
    ];

    let patched_in_mem = patch_sheet_bin(&sheet_bin, &edits).expect("patch_sheet_bin");

    let mut patched_stream = Vec::new();
    patch_sheet_bin_streaming(Cursor::new(&sheet_bin), &mut patched_stream, &edits)
        .expect("patch_sheet_bin_streaming");

    assert_eq!(patched_stream, patched_in_mem);

    let (id, payload) =
        find_cell_record(&patched_stream, 0, 1).expect("find inserted formula bool cell");
    assert_eq!(
        id, FORMULA_BOOL,
        "expected BrtFmlaBool/FORMULA_BOOL record id"
    );
    assert_eq!(payload[8], 1);
    let cce = u32::from_le_bytes(payload[11..15].try_into().unwrap()) as usize;
    assert_eq!(payload[15..15 + cce], rgce_bool_true);

    let (id, payload) =
        find_cell_record(&patched_stream, 4, 2).expect("find inserted formula error cell");
    assert_eq!(
        id, FORMULA_BOOLERR,
        "expected BrtFmlaError/FORMULA_BOOLERR record id"
    );
    assert_eq!(payload[8], 0x07);
    let cce = u32::from_le_bytes(payload[11..15].try_into().unwrap()) as usize;
    assert_eq!(payload[15..15 + cce], rgce_err_div0);
}

#[test]
fn patch_sheet_bin_streaming_inserts_formula_string_cell_matches_in_memory() {
    const FORMULA_STRING: u32 = 0x0008;

    fn ptg_str(s: &str) -> Vec<u8> {
        let mut out = vec![0x17]; // PtgStr
        let units: Vec<u16> = s.encode_utf16().collect();
        out.extend_from_slice(&(units.len() as u16).to_le_bytes());
        for unit in units {
            out.extend_from_slice(&unit.to_le_bytes());
        }
        out
    }

    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number(0, 0, 1.0);
    let sheet_bin = read_sheet_bin(builder.build_bytes());

    let rgce = ptg_str("Hello");
    let edits = [CellEdit {
        row: 0,
        col: 1,
        new_value: CellValue::Text("Hello".to_string()),
        new_formula: Some(rgce.clone()),
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
        new_style: None,
    }];

    let patched_in_mem = patch_sheet_bin(&sheet_bin, &edits).expect("patch_sheet_bin");

    let mut patched_stream = Vec::new();
    patch_sheet_bin_streaming(Cursor::new(&sheet_bin), &mut patched_stream, &edits)
        .expect("patch_sheet_bin_streaming");

    assert_eq!(patched_stream, patched_in_mem);

    let (id, payload) =
        find_cell_record(&patched_stream, 0, 1).expect("find inserted formula string cell");
    assert_eq!(
        id, FORMULA_STRING,
        "expected BrtFmlaString/FORMULA_STRING record id"
    );

    let cch = u32::from_le_bytes(payload[8..12].try_into().unwrap()) as usize;
    let flags = u16::from_le_bytes(payload[12..14].try_into().unwrap());
    assert_eq!(flags, 0);

    let utf16_start = 14usize;
    let utf16_end = utf16_start + cch * 2;
    let raw = &payload[utf16_start..utf16_end];
    let mut units = Vec::with_capacity(cch);
    for chunk in raw.chunks_exact(2) {
        units.push(u16::from_le_bytes([chunk[0], chunk[1]]));
    }
    assert_eq!(String::from_utf16_lossy(&units), "Hello");

    let cce = u32::from_le_bytes(payload[utf16_end..utf16_end + 4].try_into().unwrap()) as usize;
    assert_eq!(payload[utf16_end + 4..utf16_end + 4 + cce], rgce);
    assert!(payload[utf16_end + 4 + cce..].is_empty());
}

#[test]
fn patch_sheet_bin_streaming_patches_rk_cell_preserving_rk_record_when_possible() {
    const NUM: u32 = 0x0002;

    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number_rk(0, 1, 0.0);
    let sheet_bin = read_sheet_bin(builder.build_bytes());

    let edits = [CellEdit {
        row: 0,
        col: 1,
        new_value: CellValue::Number(0.125),
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
        new_style: None,
    }];

    let patched_in_mem = patch_sheet_bin(&sheet_bin, &edits).expect("patch_sheet_bin");

    let mut patched_stream = Vec::new();
    let changed = patch_sheet_bin_streaming(Cursor::new(&sheet_bin), &mut patched_stream, &edits)
        .expect("patch_sheet_bin_streaming");
    assert!(changed);

    assert_eq!(patched_stream, patched_in_mem);

    let (id, payload) = find_cell_record(&patched_stream, 0, 1).expect("find patched cell");
    assert_eq!(id, NUM, "expected RK NUM record id");
    let rk = u32::from_le_bytes(payload[8..12].try_into().unwrap());
    assert_eq!(decode_rk_number(rk).to_bits(), 0.125f64.to_bits());
}

#[test]
fn patch_sheet_bin_streaming_converts_rk_cell_to_float_when_needed() {
    const FLOAT: u32 = 0x0005;

    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number_rk(0, 1, 0.0);
    let sheet_bin = read_sheet_bin(builder.build_bytes());

    let edits = [CellEdit {
        row: 0,
        col: 1,
        new_value: CellValue::Number(0.1234),
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
        new_style: None,
    }];

    let patched_in_mem = patch_sheet_bin(&sheet_bin, &edits).expect("patch_sheet_bin");

    let mut patched_stream = Vec::new();
    let changed = patch_sheet_bin_streaming(Cursor::new(&sheet_bin), &mut patched_stream, &edits)
        .expect("patch_sheet_bin_streaming");
    assert!(changed);

    assert_eq!(patched_stream, patched_in_mem);

    let (id, payload) = find_cell_record(&patched_stream, 0, 1).expect("find patched cell");
    assert_eq!(id, FLOAT, "expected FLOAT record id");
    let v = f64::from_le_bytes(payload[8..16].try_into().unwrap());
    assert_eq!(v.to_bits(), 0.1234f64.to_bits());
}

#[test]
fn patch_sheet_bin_streaming_patches_shared_string_cell_when_isst_provided() {
    const STRING: u32 = 0x0007;

    let mut builder = XlsbFixtureBuilder::new();
    builder.add_shared_string("Hello");
    builder.add_shared_string("World");
    builder.set_cell_sst(0, 0, 0);
    let sheet_bin = read_sheet_bin(builder.build_bytes());

    let edits = [CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Text("World".to_string()),
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: Some(1),
        new_style: None,
    }];

    let patched_in_mem = patch_sheet_bin(&sheet_bin, &edits).expect("patch_sheet_bin");

    let mut patched_stream = Vec::new();
    let changed = patch_sheet_bin_streaming(Cursor::new(&sheet_bin), &mut patched_stream, &edits)
        .expect("patch_sheet_bin_streaming");
    assert!(changed);

    assert_eq!(patched_stream, patched_in_mem);

    let (id, payload) = find_cell_record(&patched_stream, 0, 0).expect("find patched cell");
    assert_eq!(id, STRING, "expected BrtCellIsst/STRING record id");
    assert_eq!(
        u32::from_le_bytes(payload[8..12].try_into().unwrap()),
        1,
        "expected patched cell to reference shared string index 1"
    );
}

#[test]
fn patch_sheet_bin_streaming_converts_shared_string_cell_to_inline_string_when_isst_missing() {
    const CELL_ST: u32 = 0x0006;

    let mut builder = XlsbFixtureBuilder::new();
    builder.add_shared_string("Hello");
    builder.add_shared_string("World");
    builder.set_cell_sst(0, 0, 0);
    let sheet_bin = read_sheet_bin(builder.build_bytes());

    let edits = [CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Text("World".to_string()),
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
        new_style: None,
    }];

    let patched_in_mem = patch_sheet_bin(&sheet_bin, &edits).expect("patch_sheet_bin");

    let mut patched_stream = Vec::new();
    let changed = patch_sheet_bin_streaming(Cursor::new(&sheet_bin), &mut patched_stream, &edits)
        .expect("patch_sheet_bin_streaming");
    assert!(changed);

    assert_eq!(patched_stream, patched_in_mem);

    let (id, payload) = find_cell_record(&patched_stream, 0, 0).expect("find patched cell");
    assert_eq!(id, CELL_ST, "expected BrtCellSt/CELL_ST record id");

    let cch = u32::from_le_bytes(payload[8..12].try_into().unwrap()) as usize;
    let raw = &payload[12..12 + cch * 2];
    let mut units = Vec::with_capacity(cch);
    for chunk in raw.chunks_exact(2) {
        units.push(u16::from_le_bytes([chunk[0], chunk[1]]));
    }
    assert_eq!(String::from_utf16_lossy(&units), "World");
}
