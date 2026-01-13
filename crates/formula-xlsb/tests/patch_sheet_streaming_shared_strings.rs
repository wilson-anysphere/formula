use std::fs::File;
use std::io::{Cursor, Read, Write};
use std::path::{Path, PathBuf};

use formula_xlsb::{biff12_varint, CellEdit, CellValue, XlsbWorkbook};
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

mod fixture_builder;
use fixture_builder::XlsbFixtureBuilder;

fn rich_shared_strings_fixture_path() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/rich_shared_strings.xlsb")
}

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
            0x0091 => in_sheet_data = true,  // BrtSheetData
            0x0092 => in_sheet_data = false, // BrtSheetDataEnd
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
            0x009F if payload.len() >= 8 => {
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
            0x00A0 => break, // BrtSSTEnd
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

fn build_fixture_bytes_with_inline_string() -> Vec<u8> {
    let mut builder = XlsbFixtureBuilder::new();
    builder.add_shared_string("Hello");
    builder.set_cell_sst(0, 0, 0); // A1 = "Hello" via SST
    builder.set_cell_inline_string(0, 1, "Hello"); // B1 = "Hello" inline
    builder.build_bytes()
}

fn build_single_sst_fixture_bytes() -> Vec<u8> {
    let mut builder = XlsbFixtureBuilder::new();
    builder.add_shared_string("Hello");
    builder.set_cell_sst(0, 0, 0); // A1 = "Hello" via SST
    builder.build_bytes()
}

#[test]
fn streaming_shared_strings_does_not_touch_sst_for_inserted_formula_string_cells() {
    // This workbook has a shared string table, but no cells that reference it (totalCount=0).
    // Inserting a formula string cell should *not* intern the cached value into the SST or bump
    // `totalCount` because formula cached strings are stored inline.

    fn ptg_str(s: &str) -> Vec<u8> {
        // PtgStr (0x17): [cch:u16][utf16 chars...]
        let mut out = vec![0x17];
        let units: Vec<u16> = s.encode_utf16().collect();
        out.extend_from_slice(&(units.len() as u16).to_le_bytes());
        for unit in units {
            out.extend_from_slice(&unit.to_le_bytes());
        }
        out
    }

    let mut builder = XlsbFixtureBuilder::new();
    builder.add_shared_string("Hello");
    let input_bytes = builder.build_bytes();

    let tmpdir = tempfile::tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("input.xlsb");
    std::fs::write(&input_path, &input_bytes).expect("write input workbook");
    let wb = XlsbWorkbook::open(&input_path).expect("open workbook");
    let out_path = tmpdir.path().join("out.xlsb");

    wb.save_with_cell_edits_streaming_shared_strings(
        &out_path,
        0,
        &[CellEdit {
            row: 0,
            col: 1,
            new_value: CellValue::Text("New".to_string()),
            new_formula: Some(ptg_str("New")),
            new_rgcb: None,
            new_formula_flags: None,
            // Even if the caller supplies an `isst`, formula cached strings are stored inline and
            // should not affect the shared string table counts.
            shared_string_index: Some(0),
            new_style: None,
            clear_formula: false,
        }],
    )
    .expect("save_with_cell_edits_streaming_shared_strings");

    let sheet_bin = read_zip_part(&out_path, "xl/worksheets/sheet1.bin");
    let (id, _payload) = find_cell_record(&sheet_bin, 0, 1).expect("find B1 record");
    assert_eq!(
        id, 0x0008,
        "expected BrtFmlaString/FORMULA_STRING record id"
    );

    let shared_strings_in = read_zip_part(&input_path, "xl/sharedStrings.bin");
    let shared_strings_out = read_zip_part(&out_path, "xl/sharedStrings.bin");
    assert_eq!(
        shared_strings_out, shared_strings_in,
        "expected sharedStrings.bin to be byte-identical for formula string inserts"
    );

    let stats = read_shared_strings_stats(&shared_strings_out);
    assert_eq!(stats.total_count, Some(0));
    assert_eq!(stats.unique_count, Some(1));
    assert_eq!(stats.si_count, 1);
    assert!(!stats.strings.contains(&"New".to_string()));
}

#[test]
fn streaming_shared_strings_converts_formula_string_cell_to_shared_string_value_cell() {
    // Regression test: when an edit clears an existing *formula* string cell to a plain text
    // value, the resulting value cell should be stored as `BrtCellIsst` and must bump
    // `BrtSST.totalCount` (+1).

    fn ptg_str(s: &str) -> Vec<u8> {
        // PtgStr (0x17): [cch:u16][utf16 chars...]
        let mut out = vec![0x17];
        let units: Vec<u16> = s.encode_utf16().collect();
        out.extend_from_slice(&(units.len() as u16).to_le_bytes());
        for unit in units {
            out.extend_from_slice(&unit.to_le_bytes());
        }
        out
    }

    let mut builder = XlsbFixtureBuilder::new();
    builder.add_shared_string("Hello");
    builder.set_cell_formula_str(0, 0, "Hello", ptg_str("Hello"));
    let input_bytes = builder.build_bytes();

    let tmpdir = tempfile::tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("input.xlsb");
    std::fs::write(&input_path, &input_bytes).expect("write input workbook");
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
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
            clear_formula: true,
        }],
    )
    .expect("save_with_cell_edits_streaming_shared_strings");

    let patched = XlsbWorkbook::open(&out_path).expect("re-open patched workbook");
    let sheet = patched.read_sheet(0).expect("read sheet");
    let a1 = sheet
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 0)
        .expect("A1 exists");
    assert_eq!(a1.value, CellValue::Text("Hello".to_string()));
    assert!(a1.formula.is_none(), "expected cleared formula");

    let sheet_bin = read_zip_part(&out_path, "xl/worksheets/sheet1.bin");
    let (id, payload) = find_cell_record(&sheet_bin, 0, 0).expect("find A1 record");
    assert_eq!(id, 0x0007, "expected BrtCellIsst");
    assert_eq!(
        u32::from_le_bytes(payload[8..12].try_into().unwrap()),
        0,
        "expected A1 to reference existing isst=0 ('Hello')"
    );

    let shared_strings_bin = read_zip_part(&out_path, "xl/sharedStrings.bin");
    let stats = read_shared_strings_stats(&shared_strings_bin);
    assert_eq!(stats.total_count, Some(1));
    assert_eq!(stats.unique_count, Some(1));
    assert_eq!(stats.si_count, 1);
}

fn with_corrupt_sst_unique_count(input: &[u8], bad_unique_count: u32) -> Vec<u8> {
    let mut zip = ZipArchive::new(Cursor::new(input)).expect("open xlsb zip");

    let mut parts: Vec<(String, Vec<u8>)> = Vec::new();
    for i in 0..zip.len() {
        let mut entry = zip.by_index(i).expect("zip entry");
        if entry.is_dir() {
            continue;
        }
        let name = entry.name().to_string();
        let mut bytes = Vec::with_capacity(entry.size() as usize);
        entry.read_to_end(&mut bytes).expect("read zip entry");
        parts.push((name, bytes));
    }

    for (name, bytes) in &mut parts {
        if name == "xl/sharedStrings.bin" {
            let mut cursor = Cursor::new(bytes.as_slice());
            let id = biff12_varint::read_record_id(&mut cursor)
                .expect("read record id")
                .expect("record id");
            let len = biff12_varint::read_record_len(&mut cursor)
                .expect("read record len")
                .expect("record len");
            assert_eq!(id, 0x009F, "expected BrtSST header record");
            assert!(
                len >= 8,
                "expected BrtSST payload to contain [totalCount][uniqueCount]"
            );
            let payload_start = cursor.position() as usize;
            bytes[payload_start + 4..payload_start + 8]
                .copy_from_slice(&bad_unique_count.to_le_bytes());
            break;
        }
    }

    let mut out = ZipWriter::new(Cursor::new(Vec::<u8>::new()));
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Stored);
    for (name, bytes) in parts {
        out.start_file(name, options.clone()).expect("start file");
        out.write_all(&bytes).expect("write bytes");
    }
    out.finish().expect("finish zip").into_inner()
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
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
            clear_formula: false,
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
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
            clear_formula: false,
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
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
            clear_formula: false,
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
    assert_eq!(
        bytes,
        std::fs::read(&input_path).expect("read input workbook")
    );
}

#[test]
fn streaming_shared_string_noop_rich_sst_is_lossless() {
    let fixture_path = rich_shared_strings_fixture_path();
    let wb = XlsbWorkbook::open(&fixture_path).expect("open workbook");

    let tmpdir = tempfile::tempdir().expect("create temp dir");
    let out_path = tmpdir.path().join("out.xlsb");

    // A1 in `rich_shared_strings.xlsb` is a shared-string cell whose `BrtSI` has the rich-text
    // flag set. A no-op edit should keep the original `isst` and preserve both the worksheet and
    // sharedStrings parts byte-for-byte.
    wb.save_with_cell_edits_streaming_shared_strings(
        &out_path,
        0,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Text("Hello Bold".to_string()),
            new_formula: None,
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
            clear_formula: false,
        }],
    )
    .expect("save_with_cell_edits_streaming_shared_strings");

    let report = xlsx_diff::diff_workbooks(&fixture_path, &out_path).expect("diff workbooks");
    assert!(
        report.is_empty(),
        "expected no OPC part diffs for no-op rich-SST edit, got:\n{}",
        format_report(&report)
    );

    let sheet_bin = read_zip_part(&out_path, "xl/worksheets/sheet1.bin");
    let (id, payload) = find_cell_record(&sheet_bin, 0, 0).expect("find A1 record");
    assert_eq!(id, 0x0007, "expected BrtCellIsst/STRING record id");
    assert_eq!(
        u32::from_le_bytes(payload[8..12].try_into().unwrap()),
        0,
        "expected A1 to continue referencing shared string index 0"
    );
}

#[test]
fn streaming_shared_string_noop_inline_string_does_not_touch_shared_strings() {
    let tmpdir = tempfile::tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("input.xlsb");
    std::fs::write(&input_path, build_fixture_bytes_with_inline_string())
        .expect("write input workbook");
    let wb = XlsbWorkbook::open(&input_path).expect("open workbook");
    let out_path = tmpdir.path().join("out.xlsb");

    // B1 is an inline string cell. Editing it to the same text should be a no-op, and should not
    // mutate `xl/sharedStrings.bin` counts (since the cell still uses inline storage).
    wb.save_with_cell_edits_streaming_shared_strings(
        &out_path,
        0,
        &[CellEdit {
            row: 0,
            col: 1,
            new_value: CellValue::Text("Hello".to_string()),
            new_formula: None,
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
            clear_formula: false,
        }],
    )
    .expect("save_with_cell_edits_streaming_shared_strings");

    let report = xlsx_diff::diff_workbooks(&input_path, &out_path).expect("diff workbooks");
    assert!(
        report.is_empty(),
        "expected no OPC part diffs for no-op inline-string edit, got:\n{}",
        format_report(&report)
    );

    let shared_strings_bin = read_zip_part(&out_path, "xl/sharedStrings.bin");
    let stats = read_shared_strings_stats(&shared_strings_bin);
    assert_eq!(stats.total_count, Some(1));
    assert_eq!(stats.unique_count, Some(1));
    assert_eq!(stats.si_count, 1);

    let sheet_bin = read_zip_part(&out_path, "xl/worksheets/sheet1.bin");
    let (id, _payload) = find_cell_record(&sheet_bin, 0, 1).expect("find B1 record");
    assert_eq!(id, 0x0006, "expected BrtCellSt/CELL_ST record id");
}

#[test]
fn streaming_shared_strings_repairs_unique_count_when_header_is_incorrect() {
    let tmpdir = tempfile::tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("input.xlsb");

    // Construct a workbook whose sharedStrings.bin has an invalid uniqueCount in the BrtSST
    // header. Our writer should emit a consistent uniqueCount after appending a new BrtSI.
    let bytes = build_single_sst_fixture_bytes();
    let bytes = with_corrupt_sst_unique_count(&bytes, 100);
    std::fs::write(&input_path, &bytes).expect("write input workbook");

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
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
            clear_formula: false,
        }],
    )
    .expect("save_with_cell_edits_streaming_shared_strings");

    let shared_strings_bin = read_zip_part(&out_path, "xl/sharedStrings.bin");
    let stats = read_shared_strings_stats(&shared_strings_bin);
    assert_eq!(stats.si_count, 2, "expected 2 BrtSI records after append");
    assert_eq!(stats.unique_count, Some(2));
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
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
            clear_formula: false,
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
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
            clear_formula: false,
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
