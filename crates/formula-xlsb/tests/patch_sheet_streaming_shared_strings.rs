use std::fs::File;
use std::io::{Cursor, Read};
use std::path::Path;

use formula_xlsb::{biff12_varint, CellEdit, CellValue, XlsbWorkbook};

mod fixture_builder;
use fixture_builder::XlsbFixtureBuilder;

fn format_report(report: &xlsx_diff::DiffReport) -> String {
    report
        .differences
        .iter()
        .map(|d| d.to_string())
        .collect::<Vec<_>>()
        .join("\n")
}

fn read_zip_part(path: &Path, part_path: &str) -> Vec<u8> {
    let file = File::open(path).expect("open xlsb");
    let mut zip = zip::ZipArchive::new(file).expect("open zip");
    let mut entry = zip.by_name(part_path).expect("find part");
    let mut bytes = Vec::with_capacity(entry.size() as usize);
    entry.read_to_end(&mut bytes).expect("read part bytes");
    bytes
}

fn find_cell_record(sheet_bin: &[u8], target_row: u32, target_col: u32) -> Option<(u32, Vec<u8>)> {
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
            0x0191 => in_sheet_data = true,  // BrtSheetData
            0x0192 => in_sheet_data = false, // BrtSheetDataEnd
            0x0000 if in_sheet_data => {
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

struct SharedStringsStats {
    total_count: Option<u32>,
    unique_count: Option<u32>,
    si_count: u32,
    strings: Vec<String>,
}

fn read_shared_strings_stats(shared_strings_bin: &[u8]) -> SharedStringsStats {
    let mut cursor = Cursor::new(shared_strings_bin);
    let mut total_count = None;
    let mut unique_count = None;
    let mut si_count = 0u32;
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
            0x019F if payload.len() >= 8 => {
                // BrtSST: [totalCount][uniqueCount]
                total_count = Some(u32::from_le_bytes(payload[0..4].try_into().unwrap()));
                unique_count = Some(u32::from_le_bytes(payload[4..8].try_into().unwrap()));
            }
            0x0013 if payload.len() >= 5 => {
                // BrtSI payload: [flags: u8][text: XLWideString]
                si_count = si_count.saturating_add(1);
                let flags = payload[0];
                let cch = u32::from_le_bytes(payload[1..5].try_into().unwrap()) as usize;
                let byte_len = cch.saturating_mul(2);
                let raw = payload.get(5..5 + byte_len).unwrap_or(&[]);
                let mut units = Vec::with_capacity(cch);
                for chunk in raw.chunks_exact(2) {
                    units.push(u16::from_le_bytes([chunk[0], chunk[1]]));
                }
                if flags == 0 {
                    strings.push(String::from_utf16_lossy(&units));
                }
            }
            0x01A0 => break, // BrtSSTEnd
            _ => {}
        }
    }

    SharedStringsStats {
        total_count,
        unique_count,
        si_count,
        strings,
    }
}

fn build_fixture_bytes() -> Vec<u8> {
    let mut builder = XlsbFixtureBuilder::new();
    builder.add_shared_string("Hello");
    builder.add_shared_string("World");
    builder.set_cell_sst(0, 0, 0); // A1 = "Hello"
    builder.set_cell_number(0, 1, 42.0); // B1 = 42.0

    builder.build_bytes()
}

#[test]
fn streaming_shared_string_edit_updates_isst_and_preserves_counts() {
    let tmpdir = tempfile::tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("input.xlsb");
    std::fs::write(&input_path, build_fixture_bytes()).expect("write input workbook");
    let wb = XlsbWorkbook::open(&input_path).expect("open workbook");
    let out_path = tmpdir.path().join("out.xlsb");

    wb.save_with_cell_edits_streaming_shared_strings(
        &out_path,
        0,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Text("World".to_string()),
            new_formula: None,
            shared_string_index: None,
        }],
    )
    .expect("save_with_cell_edits_streaming_shared_strings");

    let sheet_bin = read_zip_part(&out_path, "xl/worksheets/sheet1.bin");
    let (id, payload) = find_cell_record(&sheet_bin, 0, 0).expect("find A1 record");
    assert_eq!(id, 0x0007, "expected BrtCellIsst/STRING record id");
    assert_eq!(
        u32::from_le_bytes(payload[8..12].try_into().unwrap()),
        1,
        "expected A1 to reference shared string index 1"
    );

    let shared_strings_bin = read_zip_part(&out_path, "xl/sharedStrings.bin");
    let stats = read_shared_strings_stats(&shared_strings_bin);
    assert_eq!(stats.total_count, Some(1));
    assert_eq!(stats.unique_count, Some(2));
    assert_eq!(stats.si_count, 2);
}

#[test]
fn streaming_shared_string_edit_appends_new_si_and_updates_unique_count() {
    let tmpdir = tempfile::tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("input.xlsb");
    std::fs::write(&input_path, build_fixture_bytes()).expect("write input workbook");
    let wb = XlsbWorkbook::open(&input_path).expect("open workbook");
    let out_path = tmpdir.path().join("out.xlsb");

    wb.save_with_cell_edits_streaming_shared_strings(
        &out_path,
        0,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Text("New".to_string()),
            new_formula: None,
            shared_string_index: None,
        }],
    )
    .expect("save_with_cell_edits_streaming_shared_strings");

    let sheet_bin = read_zip_part(&out_path, "xl/worksheets/sheet1.bin");
    let (id, payload) = find_cell_record(&sheet_bin, 0, 0).expect("find A1 record");
    assert_eq!(id, 0x0007, "expected BrtCellIsst/STRING record id");
    assert_eq!(
        u32::from_le_bytes(payload[8..12].try_into().unwrap()),
        2,
        "expected A1 to reference appended shared string index 2"
    );

    let shared_strings_bin = read_zip_part(&out_path, "xl/sharedStrings.bin");
    let stats = read_shared_strings_stats(&shared_strings_bin);
    assert_eq!(stats.total_count, Some(1));
    assert_eq!(stats.unique_count, Some(3));
    assert_eq!(stats.si_count, 3);
    assert!(stats.strings.contains(&"New".to_string()));
}

#[test]
fn streaming_shared_string_noop_is_lossless() {
    let tmpdir = tempfile::tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("input.xlsb");
    let bytes = build_fixture_bytes();
    std::fs::write(&input_path, &bytes).expect("write input workbook");
    let wb = XlsbWorkbook::open(&input_path).expect("open workbook");
    let out_path = tmpdir.path().join("out.xlsb");

    wb.save_with_cell_edits_streaming_shared_strings(
        &out_path,
        0,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Text("Hello".to_string()),
            new_formula: None,
            shared_string_index: None,
        }],
    )
    .expect("save_with_cell_edits_streaming_shared_strings");

    let report = xlsx_diff::diff_workbooks(&input_path, &out_path).expect("diff workbooks");
    assert!(
        report.is_empty(),
        "expected no OPC part diffs for no-op streaming shared-string edit, got:\n{}",
        format_report(&report)
    );

    // Also ensure the serialized bytes are the same for the key parts we touched.
    let out_sheet = read_zip_part(&out_path, "xl/worksheets/sheet1.bin");
    let out_sst = read_zip_part(&out_path, "xl/sharedStrings.bin");
    assert_eq!(
        out_sheet,
        read_zip_part(&input_path, "xl/worksheets/sheet1.bin"),
        "expected worksheet part bytes to be identical for a no-op edit"
    );
    assert_eq!(
        out_sst,
        read_zip_part(&input_path, "xl/sharedStrings.bin"),
        "expected sharedStrings.bin bytes to be identical for a no-op edit"
    );

    // Ensure we didn't accidentally mutate the in-memory bytes either.
    assert_eq!(bytes, std::fs::read(&input_path).expect("read input workbook"));
}

#[test]
fn streaming_shared_string_total_count_updates_when_cell_leaves_sst() {
    let tmpdir = tempfile::tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("input.xlsb");
    std::fs::write(&input_path, build_fixture_bytes()).expect("write input workbook");
    let wb = XlsbWorkbook::open(&input_path).expect("open workbook");
    let out_path = tmpdir.path().join("out.xlsb");

    // A1 is an SST cell in the fixture. Changing it to a number should decrement totalCount.
    wb.save_with_cell_edits_streaming_shared_strings(
        &out_path,
        0,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Number(123.0),
            new_formula: None,
            shared_string_index: None,
        }],
    )
    .expect("save_with_cell_edits_streaming_shared_strings");

    let shared_strings_bin = read_zip_part(&out_path, "xl/sharedStrings.bin");
    let stats = read_shared_strings_stats(&shared_strings_bin);
    assert_eq!(stats.total_count, Some(0));
    assert_eq!(stats.unique_count, Some(2));
    assert_eq!(stats.si_count, 2);

    let sheet_bin = read_zip_part(&out_path, "xl/worksheets/sheet1.bin");
    let (id, _payload) = find_cell_record(&sheet_bin, 0, 0).expect("find A1 record");
    assert_eq!(id, 0x0005, "expected BrtCellReal/FLOAT record id");
}

#[test]
fn streaming_shared_string_total_count_updates_when_cell_enters_sst() {
    let tmpdir = tempfile::tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("input.xlsb");
    std::fs::write(&input_path, build_fixture_bytes()).expect("write input workbook");
    let wb = XlsbWorkbook::open(&input_path).expect("open workbook");
    let out_path = tmpdir.path().join("out.xlsb");

    // B1 is numeric in the fixture. Changing it to text should convert it into an SST cell.
    wb.save_with_cell_edits_streaming_shared_strings(
        &out_path,
        0,
        &[CellEdit {
            row: 0,
            col: 1,
            new_value: CellValue::Text("World".to_string()),
            new_formula: None,
            shared_string_index: None,
        }],
    )
    .expect("save_with_cell_edits_streaming_shared_strings");

    let shared_strings_bin = read_zip_part(&out_path, "xl/sharedStrings.bin");
    let stats = read_shared_strings_stats(&shared_strings_bin);
    assert_eq!(stats.total_count, Some(2));
    assert_eq!(stats.unique_count, Some(2));
    assert_eq!(stats.si_count, 2);

    let sheet_bin = read_zip_part(&out_path, "xl/worksheets/sheet1.bin");
    let (id, payload) = find_cell_record(&sheet_bin, 0, 1).expect("find B1 record");
    assert_eq!(id, 0x0007, "expected BrtCellIsst/STRING record id");
    assert_eq!(
        u32::from_le_bytes(payload[8..12].try_into().unwrap()),
        1,
        "expected B1 to reference shared string index 1"
    );
}
