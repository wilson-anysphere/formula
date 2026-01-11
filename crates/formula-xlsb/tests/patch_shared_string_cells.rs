use std::fs::File;
use std::io::{Cursor, Read};

use formula_xlsb::{biff12_varint, CellEdit, CellValue, XlsbWorkbook};
use pretty_assertions::assert_eq;
use tempfile::tempdir;

mod fixture_builder;
use fixture_builder::XlsbFixtureBuilder;

fn read_zip_part(path: &str, part_path: &str) -> Vec<u8> {
    let file = File::open(path).expect("open xlsb");
    let mut zip = zip::ZipArchive::new(file).expect("open zip");
    let mut entry = zip.by_name(part_path).expect("find part");
    let mut bytes = Vec::with_capacity(entry.size() as usize);
    entry.read_to_end(&mut bytes).expect("read part bytes");
    bytes
}

fn find_cell_record(sheet_bin: &[u8], target_row: u32, target_col: u32) -> Option<(u32, Vec<u8>)> {
    const SHEETDATA: u32 = 0x0191;
    const SHEETDATA_END: u32 = 0x0192;
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
            None => return None,
        };
        let mut payload = vec![0u8; len];
        cursor.read_exact(&mut payload).ok()?;

        match id {
            SHEETDATA => in_sheet_data = true,
            SHEETDATA_END => in_sheet_data = false,
            ROW if in_sheet_data => {
                if payload.len() >= 4 {
                    current_row = u32::from_le_bytes(payload[0..4].try_into().unwrap());
                }
            }
            _ if in_sheet_data => {
                if payload.len() < 8 {
                    continue;
                }
                let col = u32::from_le_bytes(payload[0..4].try_into().unwrap());
                if current_row == target_row && col == target_col {
                    return Some((id, payload));
                }
            }
            _ => {}
        }
    }
    None
}

struct SharedStringsInfo {
    unique_count: Option<u32>,
    strings: Vec<String>,
}

fn read_shared_strings_info(shared_strings_bin: &[u8]) -> SharedStringsInfo {
    const SST: u32 = 0x019F;
    const SI: u32 = 0x0013;
    const SST_END: u32 = 0x01A0;

    let mut cursor = Cursor::new(shared_strings_bin);
    let mut unique_count = None;
    let mut strings = Vec::new();

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
        cursor
            .read_exact(&mut payload)
            .expect("read record payload");

        match id {
            SST if payload.len() >= 8 => {
                unique_count = Some(u32::from_le_bytes(payload[0..4].try_into().unwrap()));
            }
            SI if payload.len() >= 5 => {
                let flags = payload[0];
                let cch = u32::from_le_bytes(payload[1..5].try_into().unwrap()) as usize;
                let byte_len = cch.saturating_mul(2);
                let raw = payload.get(5..5 + byte_len).unwrap_or(&[]);
                let mut units = Vec::with_capacity(cch);
                for chunk in raw.chunks_exact(2) {
                    units.push(u16::from_le_bytes([chunk[0], chunk[1]]));
                }
                let text = String::from_utf16_lossy(&units);
                if flags == 0 {
                    strings.push(text);
                }
            }
            SST_END => break,
            _ => {}
        }
    }

    SharedStringsInfo {
        unique_count,
        strings,
    }
}

#[test]
fn patching_shared_string_cell_keeps_it_as_string_record() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.add_shared_string("Hello");
    builder.add_shared_string("World");
    builder.set_cell_sst(0, 0, 0);

    let bytes = builder.build_bytes();

    let tmpdir = tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("input.xlsb");
    let output_path = tmpdir.path().join("output.xlsb");
    std::fs::write(&input_path, bytes).expect("write input workbook");

    let wb = XlsbWorkbook::open(&input_path).expect("open input workbook");
    wb.save_with_cell_edits_shared_strings(
        &output_path,
        0,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Text("World".to_string()),
            new_formula: None,
            shared_string_index: None,
        }],
    )
    .expect("save_with_cell_edits_shared_strings");

    let wb2 = XlsbWorkbook::open(&output_path).expect("open output workbook");
    let sheet = wb2.read_sheet(0).expect("read sheet");
    let a1 = sheet
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 0)
        .expect("A1 exists");
    assert_eq!(a1.value, CellValue::Text("World".to_string()));

    let sheet_bin = read_zip_part(output_path.to_str().unwrap(), "xl/worksheets/sheet1.bin");
    let (id, payload) = find_cell_record(&sheet_bin, 0, 0).expect("find A1 record");
    assert_eq!(id, 0x0007, "expected BrtCellIsst/STRING record id");
    assert_eq!(
        u32::from_le_bytes(payload[8..12].try_into().unwrap()),
        1,
        "expected A1 to reference shared string index 1"
    );
}

#[test]
fn patching_shared_string_cell_appends_to_shared_strings_bin() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.add_shared_string("Hello");
    builder.add_shared_string("World");
    builder.set_cell_sst(0, 0, 0);

    let bytes = builder.build_bytes();

    let tmpdir = tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("input.xlsb");
    let output_path = tmpdir.path().join("output.xlsb");
    std::fs::write(&input_path, bytes).expect("write input workbook");

    let wb = XlsbWorkbook::open(&input_path).expect("open input workbook");
    wb.save_with_cell_edits_shared_strings(
        &output_path,
        0,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Text("New".to_string()),
            new_formula: None,
            shared_string_index: None,
        }],
    )
    .expect("save_with_cell_edits_shared_strings");

    let wb2 = XlsbWorkbook::open(&output_path).expect("open output workbook");
    let sheet = wb2.read_sheet(0).expect("read sheet");
    let a1 = sheet
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 0)
        .expect("A1 exists");
    assert_eq!(a1.value, CellValue::Text("New".to_string()));

    let sheet_bin = read_zip_part(output_path.to_str().unwrap(), "xl/worksheets/sheet1.bin");
    let (id, payload) = find_cell_record(&sheet_bin, 0, 0).expect("find A1 record");
    assert_eq!(id, 0x0007, "expected BrtCellIsst/STRING record id");
    assert_eq!(
        u32::from_le_bytes(payload[8..12].try_into().unwrap()),
        2,
        "expected A1 to reference appended shared string index 2"
    );

    let shared_strings_bin = read_zip_part(output_path.to_str().unwrap(), "xl/sharedStrings.bin");
    let info = read_shared_strings_info(&shared_strings_bin);
    assert_eq!(info.unique_count, Some(3));
    assert_eq!(info.strings.len(), 3);
    assert_eq!(info.strings[2], "New");
    assert!(
        info.strings.contains(&"New".to_string()),
        "expected sharedStrings.bin to contain 'New'"
    );
}

#[test]
fn inserting_new_text_cell_uses_shared_string_record_and_updates_shared_strings_bin() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.add_shared_string("Hello");
    builder.set_cell_sst(0, 0, 0);

    let bytes = builder.build_bytes();

    let tmpdir = tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("input.xlsb");
    let output_path = tmpdir.path().join("output.xlsb");
    std::fs::write(&input_path, bytes).expect("write input workbook");

    let wb = XlsbWorkbook::open(&input_path).expect("open input workbook");
    wb.save_with_cell_edits_shared_strings(
        &output_path,
        0,
        &[CellEdit {
            row: 0,
            col: 1,
            new_value: CellValue::Text("New".to_string()),
            new_formula: None,
            shared_string_index: None,
        }],
    )
    .expect("save_with_cell_edits_shared_strings");

    let wb2 = XlsbWorkbook::open(&output_path).expect("open output workbook");
    let sheet = wb2.read_sheet(0).expect("read sheet");
    let b1 = sheet
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 1)
        .expect("B1 exists");
    assert_eq!(b1.value, CellValue::Text("New".to_string()));

    let sheet_bin = read_zip_part(output_path.to_str().unwrap(), "xl/worksheets/sheet1.bin");
    let (id, payload) = find_cell_record(&sheet_bin, 0, 1).expect("find B1 record");
    assert_eq!(id, 0x0007, "expected BrtCellIsst/STRING record id");
    assert_eq!(
        u32::from_le_bytes(payload[8..12].try_into().unwrap()),
        1,
        "expected B1 to reference appended shared string index 1"
    );

    let shared_strings_bin = read_zip_part(output_path.to_str().unwrap(), "xl/sharedStrings.bin");
    let info = read_shared_strings_info(&shared_strings_bin);
    assert_eq!(info.unique_count, Some(2));
    assert_eq!(info.strings.len(), 2);
    assert_eq!(info.strings[1], "New");
 }
