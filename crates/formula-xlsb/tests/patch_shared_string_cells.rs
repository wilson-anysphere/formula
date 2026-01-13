use std::fs::File;
use std::io::{Cursor, Read};
use std::path::{Path, PathBuf};

use formula_xlsb::{biff12_varint, CellEdit, CellValue, XlsbWorkbook};
use pretty_assertions::assert_eq;
use tempfile::tempdir;

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

fn read_zip_part(path: &str, part_path: &str) -> Vec<u8> {
    let file = File::open(path).expect("open xlsb");
    let mut zip = zip::ZipArchive::new(file).expect("open zip");
    let mut entry = zip.by_name(part_path).expect("find part");
    let mut bytes = Vec::with_capacity(entry.size() as usize);
    entry.read_to_end(&mut bytes).expect("read part bytes");
    bytes
}

fn find_cell_record(sheet_bin: &[u8], target_row: u32, target_col: u32) -> Option<(u32, Vec<u8>)> {
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
    total_count: Option<u32>,
    unique_count: Option<u32>,
    strings: Vec<String>,
}

fn read_shared_strings_info(shared_strings_bin: &[u8]) -> SharedStringsInfo {
    const SST: u32 = 0x009F;
    const SI: u32 = 0x0013;
    const SST_END: u32 = 0x00A0;

    let mut cursor = Cursor::new(shared_strings_bin);
    let mut total_count = None;
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
                total_count = Some(u32::from_le_bytes(payload[0..4].try_into().unwrap()));
                unique_count = Some(u32::from_le_bytes(payload[4..8].try_into().unwrap()));
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
        total_count,
        unique_count,
        strings,
    }
}

#[test]
fn shared_strings_save_does_not_touch_sst_for_inserted_formula_string_cells() {
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
    let bytes = builder.build_bytes();

    let tmpdir = tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("input.xlsb");
    let output_path = tmpdir.path().join("output.xlsb");
    std::fs::write(&input_path, &bytes).expect("write input workbook");

    let wb = XlsbWorkbook::open(&input_path).expect("open input workbook");
    wb.save_with_cell_edits_shared_strings(
        &output_path,
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
    .expect("save_with_cell_edits_shared_strings");

    let sheet_bin = read_zip_part(output_path.to_str().unwrap(), "xl/worksheets/sheet1.bin");
    let (id, _payload) = find_cell_record(&sheet_bin, 0, 1).expect("find B1 record");
    assert_eq!(
        id, 0x0008,
        "expected BrtFmlaString/FORMULA_STRING record id"
    );

    let shared_strings_in = read_zip_part(input_path.to_str().unwrap(), "xl/sharedStrings.bin");
    let shared_strings_out = read_zip_part(output_path.to_str().unwrap(), "xl/sharedStrings.bin");
    assert_eq!(
        shared_strings_out, shared_strings_in,
        "expected sharedStrings.bin to be byte-identical for formula string inserts"
    );

    let info = read_shared_strings_info(&shared_strings_out);
    assert_eq!(info.total_count, Some(0));
    assert_eq!(info.unique_count, Some(1));
    assert!(!info.strings.contains(&"New".to_string()));
}

#[test]
fn shared_strings_save_converts_formula_string_cell_to_shared_string_value_cell() {
    // Regression test: when an edit clears an existing *formula* string cell to a plain text value,
    // the resulting value cell should be eligible for shared-string storage (`BrtCellIsst`) and
    // must bump `BrtSST.cstTotal` (+1). Formula cached strings are inline and do not count toward
    // `cstTotal`.

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
    // Workbook has an SST, but no cells that reference it yet (totalCount=0).
    builder.add_shared_string("Hello");
    // A1 is a formula string cell with cached value "Hello" (stored inline, not via SST).
    builder.set_cell_formula_str(0, 0, "Hello", ptg_str("Hello"));
    let bytes = builder.build_bytes();

    let tmpdir = tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("input.xlsb");
    let output_path = tmpdir.path().join("output.xlsb");
    std::fs::write(&input_path, &bytes).expect("write input workbook");

    let wb = XlsbWorkbook::open(&input_path).expect("open input workbook");
    wb.save_with_cell_edits_shared_strings(
        &output_path,
        0,
        &[CellEdit {
            row: 0,
            col: 0,
            // Paste-values style: clear the formula, keep the cached string value.
            new_value: CellValue::Text("Hello".to_string()),
            new_formula: None,
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
            clear_formula: true,
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
    assert_eq!(a1.value, CellValue::Text("Hello".to_string()));
    assert!(a1.formula.is_none(), "expected cleared formula");

    let sheet_bin = read_zip_part(output_path.to_str().unwrap(), "xl/worksheets/sheet1.bin");
    let (id, payload) = find_cell_record(&sheet_bin, 0, 0).expect("find A1 record");
    assert_eq!(id, 0x0007, "expected BrtCellIsst/STRING record id");
    assert_eq!(
        u32::from_le_bytes(payload[8..12].try_into().unwrap()),
        0,
        "expected A1 to reference existing shared string index 0"
    );

    let shared_strings_bin = read_zip_part(output_path.to_str().unwrap(), "xl/sharedStrings.bin");
    let info = read_shared_strings_info(&shared_strings_bin);
    assert_eq!(info.total_count, Some(1), "expected cstTotal to increment (+1)");
    assert_eq!(info.unique_count, Some(1), "expected no new unique strings");
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
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
            clear_formula: false,
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

    let shared_strings_bin = read_zip_part(output_path.to_str().unwrap(), "xl/sharedStrings.bin");
    let info = read_shared_strings_info(&shared_strings_bin);
    assert_eq!(info.total_count, Some(1));
    assert_eq!(info.unique_count, Some(2));
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
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
            clear_formula: false,
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
    assert_eq!(info.total_count, Some(1));
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
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
            clear_formula: false,
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
    assert_eq!(info.total_count, Some(2));
    assert_eq!(info.unique_count, Some(2));
    assert_eq!(info.strings.len(), 2);
    assert_eq!(info.strings[1], "New");
}

#[test]
fn patching_existing_numeric_cell_to_text_uses_shared_string_record_and_updates_total_count() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.add_shared_string("Hello");
    builder.add_shared_string("World");
    builder.set_cell_sst(0, 0, 0); // A1 = "Hello" via SST
    builder.set_cell_number(0, 1, 42.0); // B1 = 42.0

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
            new_value: CellValue::Text("World".to_string()),
            new_formula: None,
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
            clear_formula: false,
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
    assert_eq!(b1.value, CellValue::Text("World".to_string()));

    let sheet_bin = read_zip_part(output_path.to_str().unwrap(), "xl/worksheets/sheet1.bin");
    let (id, payload) = find_cell_record(&sheet_bin, 0, 1).expect("find B1 record");
    assert_eq!(id, 0x0007, "expected BrtCellIsst/STRING record id");
    assert_eq!(
        u32::from_le_bytes(payload[8..12].try_into().unwrap()),
        1,
        "expected B1 to reference shared string index 1"
    );

    let shared_strings_bin = read_zip_part(output_path.to_str().unwrap(), "xl/sharedStrings.bin");
    let info = read_shared_strings_info(&shared_strings_bin);
    assert_eq!(info.total_count, Some(2));
    assert_eq!(info.unique_count, Some(2));
}

#[test]
fn patching_inline_string_noop_is_lossless_and_does_not_touch_shared_strings() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.add_shared_string("Hello");
    builder.set_cell_sst(0, 0, 0); // A1 = "Hello" via SST
    builder.set_cell_inline_string(0, 1, "Hello"); // B1 = "Hello" inline

    let bytes = builder.build_bytes();

    let tmpdir = tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("input.xlsb");
    let output_path = tmpdir.path().join("output.xlsb");
    std::fs::write(&input_path, &bytes).expect("write input workbook");

    let wb = XlsbWorkbook::open(&input_path).expect("open input workbook");
    wb.save_with_cell_edits_shared_strings(
        &output_path,
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
    .expect("save_with_cell_edits_shared_strings");

    let report = xlsx_diff::diff_workbooks(&input_path, &output_path).expect("diff workbooks");
    assert!(
        report.is_empty(),
        "expected no OPC part diffs for no-op inline-string edit, got:\n{}",
        format_report(&report)
    );

    let out_sheet = read_zip_part(output_path.to_str().unwrap(), "xl/worksheets/sheet1.bin");
    let out_sst = read_zip_part(output_path.to_str().unwrap(), "xl/sharedStrings.bin");
    assert_eq!(
        out_sheet,
        read_zip_part(input_path.to_str().unwrap(), "xl/worksheets/sheet1.bin"),
        "expected worksheet part bytes to be identical for a no-op inline-string edit"
    );
    assert_eq!(
        out_sst,
        read_zip_part(input_path.to_str().unwrap(), "xl/sharedStrings.bin"),
        "expected sharedStrings.bin bytes to be identical for a no-op inline-string edit"
    );

    let (id, _payload) = find_cell_record(&out_sheet, 0, 1).expect("find B1 record");
    assert_eq!(id, 0x0006, "expected BrtCellSt/CELL_ST record id");
}

#[test]
fn patching_rich_shared_string_noop_is_lossless() {
    let fixture_path = rich_shared_strings_fixture_path();
    let wb = XlsbWorkbook::open(&fixture_path).expect("open xlsb fixture");

    let tmpdir = tempdir().expect("create temp dir");
    let output_path = tmpdir.path().join("output.xlsb");

    wb.save_with_cell_edits_shared_strings(
        &output_path,
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
    .expect("save_with_cell_edits_shared_strings");

    let report = xlsx_diff::diff_workbooks(&fixture_path, &output_path).expect("diff workbooks");
    assert!(
        report.is_empty(),
        "expected no OPC part diffs for no-op rich-SST edit, got:\n{}",
        format_report(&report)
    );

    let sheet_bin = read_zip_part(output_path.to_str().unwrap(), "xl/worksheets/sheet1.bin");
    let (id, payload) = find_cell_record(&sheet_bin, 0, 0).expect("find A1 record");
    assert_eq!(id, 0x0007, "expected BrtCellIsst/STRING record id");
    assert_eq!(
        u32::from_le_bytes(payload[8..12].try_into().unwrap()),
        0,
        "expected A1 to continue referencing shared string index 0"
    );

    let shared_strings_out = read_zip_part(output_path.to_str().unwrap(), "xl/sharedStrings.bin");
    let shared_strings_in = read_zip_part(fixture_path.to_str().unwrap(), "xl/sharedStrings.bin");
    assert_eq!(
        shared_strings_out, shared_strings_in,
        "expected sharedStrings.bin bytes to be identical for a no-op rich shared-string edit"
    );
}
