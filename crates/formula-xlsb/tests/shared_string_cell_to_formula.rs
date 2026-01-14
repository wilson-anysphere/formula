use std::fs::File;
use std::io::{Cursor, Read};
use std::path::Path;

use formula_biff::encode_rgce;
use formula_xlsb::{biff12_varint, CellEdit, CellValue, XlsbWorkbook};
use pretty_assertions::assert_eq;
use tempfile::tempdir;

mod fixture_builder;
use fixture_builder::XlsbFixtureBuilder;

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

fn read_zip_part(path: &Path, part_path: &str) -> Vec<u8> {
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

#[derive(Debug)]
struct SharedStringsStats {
    total_count: Option<u32>,
    unique_count: Option<u32>,
    si_count: u32,
    strings: Vec<String>,
}

fn read_shared_strings_stats(shared_strings_bin: &[u8]) -> SharedStringsStats {
    const SST: u32 = 0x009F;
    const SI: u32 = 0x0013;
    const SST_END: u32 = 0x00A0;

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
            .expect("read shared string record payload");

        match id {
            SST if payload.len() >= 8 => {
                total_count = Some(u32::from_le_bytes(payload[0..4].try_into().unwrap()));
                unique_count = Some(u32::from_le_bytes(payload[4..8].try_into().unwrap()));
            }
            SI if payload.len() >= 5 => {
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
            SST_END => break,
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

fn build_sst_fixture_bytes() -> Vec<u8> {
    let mut builder = XlsbFixtureBuilder::new();
    builder.add_shared_string("Hello");
    builder.add_shared_string("World");
    builder.set_cell_sst(0, 0, 0); // A1 = "Hello" via SST (BrtCellIsst)
    builder.build_bytes()
}

#[test]
fn shared_strings_save_converting_shared_string_cell_to_formula_decrements_total_count() {
    let input_bytes = build_sst_fixture_bytes();

    let tmpdir = tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("input.xlsb");
    let output_path = tmpdir.path().join("output.xlsb");
    std::fs::write(&input_path, &input_bytes).expect("write input workbook");

    let sheet_in = read_zip_part(&input_path, "xl/worksheets/sheet1.bin");
    let (id_in, _) = find_cell_record(&sheet_in, 0, 0).expect("find A1 record in input");
    assert_eq!(id_in, 0x0007, "expected input A1 to be BrtCellIsst");

    let sst_in = read_zip_part(&input_path, "xl/sharedStrings.bin");
    let stats_in = read_shared_strings_stats(&sst_in);
    assert_eq!(stats_in.total_count, Some(1));
    assert_eq!(stats_in.unique_count, Some(2));
    assert_eq!(stats_in.si_count, 2);

    let wb = XlsbWorkbook::open(&input_path).expect("open workbook");
    let rgce = encode_rgce("=1").expect("encode formula");
    wb.save_with_cell_edits_shared_strings(
        &output_path,
        0,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Number(1.0),
            new_style: None,
            clear_formula: false,
            new_formula: Some(rgce.clone()),
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            clear_formula: false,
        }],
    )
    .expect("save_with_cell_edits_shared_strings");

    let wb_out = XlsbWorkbook::open(&output_path).expect("open output workbook");
    let sheet = wb_out.read_sheet(0).expect("read output sheet");
    let a1 = sheet
        .cells
        .iter()
        .find(|c| (c.row, c.col) == (0, 0))
        .expect("A1 exists");
    assert_eq!(a1.value, CellValue::Number(1.0));
    let formula = a1.formula.as_ref().expect("A1 should now be a formula cell");
    assert_eq!(formula.rgce, rgce);

    let sheet_out = read_zip_part(&output_path, "xl/worksheets/sheet1.bin");
    let (id_out, _) = find_cell_record(&sheet_out, 0, 0).expect("find A1 record in output");
    assert_eq!(id_out, 0x0009, "expected output A1 to be BrtFmlaNum");

    let sst_out = read_zip_part(&output_path, "xl/sharedStrings.bin");
    let stats_out = read_shared_strings_stats(&sst_out);
    assert_eq!(
        stats_out.total_count,
        Some(stats_in.total_count.unwrap().saturating_sub(1))
    );
    assert_eq!(stats_out.unique_count, stats_in.unique_count);
    assert_eq!(stats_out.si_count, stats_in.si_count);
    assert_eq!(stats_out.strings, stats_in.strings);
}

#[test]
fn streaming_shared_strings_save_converting_shared_string_cell_to_formula_decrements_total_count() {
    let input_bytes = build_sst_fixture_bytes();

    let tmpdir = tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("input.xlsb");
    let output_path = tmpdir.path().join("output.xlsb");
    std::fs::write(&input_path, &input_bytes).expect("write input workbook");

    let sst_in = read_zip_part(&input_path, "xl/sharedStrings.bin");
    let stats_in = read_shared_strings_stats(&sst_in);
    assert_eq!(stats_in.total_count, Some(1));
    assert_eq!(stats_in.unique_count, Some(2));
    assert_eq!(stats_in.si_count, 2);

    let wb = XlsbWorkbook::open(&input_path).expect("open workbook");
    let rgce = encode_rgce("=1").expect("encode formula");
    wb.save_with_cell_edits_streaming_shared_strings(
        &output_path,
        0,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Number(1.0),
            new_style: None,
            clear_formula: false,
            new_formula: Some(rgce.clone()),
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            clear_formula: false,
        }],
    )
    .expect("save_with_cell_edits_streaming_shared_strings");

    let wb_out = XlsbWorkbook::open(&output_path).expect("open output workbook");
    let sheet = wb_out.read_sheet(0).expect("read output sheet");
    let a1 = sheet
        .cells
        .iter()
        .find(|c| (c.row, c.col) == (0, 0))
        .expect("A1 exists");
    assert_eq!(a1.value, CellValue::Number(1.0));
    let formula = a1.formula.as_ref().expect("A1 should now be a formula cell");
    assert_eq!(formula.rgce, rgce);

    let sheet_out = read_zip_part(&output_path, "xl/worksheets/sheet1.bin");
    let (id_out, _) = find_cell_record(&sheet_out, 0, 0).expect("find A1 record in output");
    assert_eq!(id_out, 0x0009, "expected output A1 to be BrtFmlaNum");

    let sst_out = read_zip_part(&output_path, "xl/sharedStrings.bin");
    let stats_out = read_shared_strings_stats(&sst_out);
    assert_eq!(
        stats_out.total_count,
        Some(stats_in.total_count.unwrap().saturating_sub(1))
    );
    assert_eq!(stats_out.unique_count, stats_in.unique_count);
    assert_eq!(stats_out.si_count, stats_in.si_count);
    assert_eq!(stats_out.strings, stats_in.strings);
}

#[test]
fn shared_strings_save_converting_shared_string_cell_to_formula_string_decrements_total_count() {
    let input_bytes = build_sst_fixture_bytes();

    let tmpdir = tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("input.xlsb");
    let output_path = tmpdir.path().join("output.xlsb");
    std::fs::write(&input_path, &input_bytes).expect("write input workbook");

    let sst_in = read_zip_part(&input_path, "xl/sharedStrings.bin");
    let stats_in = read_shared_strings_stats(&sst_in);
    assert_eq!(stats_in.total_count, Some(1));
    assert_eq!(stats_in.unique_count, Some(2));
    assert_eq!(stats_in.si_count, 2);

    let wb = XlsbWorkbook::open(&input_path).expect("open workbook");
    let rgce = ptg_str("New");
    wb.save_with_cell_edits_shared_strings(
        &output_path,
        0,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Text("New".to_string()),
            new_style: None,
            clear_formula: false,
            new_formula: Some(rgce.clone()),
            new_rgcb: None,
            new_formula_flags: None,
            // Even if the caller passes an `isst`, formula cached strings are stored inline and
            // should not reference the shared string table.
            shared_string_index: Some(0),
        }],
    )
    .expect("save_with_cell_edits_shared_strings");

    let wb_out = XlsbWorkbook::open(&output_path).expect("open output workbook");
    let sheet = wb_out.read_sheet(0).expect("read output sheet");
    let a1 = sheet
        .cells
        .iter()
        .find(|c| (c.row, c.col) == (0, 0))
        .expect("A1 exists");
    assert_eq!(a1.value, CellValue::Text("New".to_string()));
    let formula = a1.formula.as_ref().expect("A1 should now be a formula cell");
    assert_eq!(formula.rgce, rgce);

    let sheet_out = read_zip_part(&output_path, "xl/worksheets/sheet1.bin");
    let (id_out, _) = find_cell_record(&sheet_out, 0, 0).expect("find A1 record in output");
    assert_eq!(id_out, 0x0008, "expected output A1 to be BrtFmlaString");

    let sst_out = read_zip_part(&output_path, "xl/sharedStrings.bin");
    let stats_out = read_shared_strings_stats(&sst_out);
    assert_eq!(
        stats_out.total_count,
        Some(stats_in.total_count.unwrap().saturating_sub(1))
    );
    assert_eq!(stats_out.unique_count, stats_in.unique_count);
    assert_eq!(stats_out.si_count, stats_in.si_count);
    assert_eq!(stats_out.strings, stats_in.strings);
}

#[test]
fn streaming_shared_strings_save_converting_shared_string_cell_to_formula_string_decrements_total_count() {
    let input_bytes = build_sst_fixture_bytes();

    let tmpdir = tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("input.xlsb");
    let output_path = tmpdir.path().join("output.xlsb");
    std::fs::write(&input_path, &input_bytes).expect("write input workbook");

    let sst_in = read_zip_part(&input_path, "xl/sharedStrings.bin");
    let stats_in = read_shared_strings_stats(&sst_in);
    assert_eq!(stats_in.total_count, Some(1));
    assert_eq!(stats_in.unique_count, Some(2));
    assert_eq!(stats_in.si_count, 2);

    let wb = XlsbWorkbook::open(&input_path).expect("open workbook");
    let rgce = ptg_str("New");
    wb.save_with_cell_edits_streaming_shared_strings(
        &output_path,
        0,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Text("New".to_string()),
            new_style: None,
            clear_formula: false,
            new_formula: Some(rgce.clone()),
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: Some(0),
        }],
    )
    .expect("save_with_cell_edits_streaming_shared_strings");

    let wb_out = XlsbWorkbook::open(&output_path).expect("open output workbook");
    let sheet = wb_out.read_sheet(0).expect("read output sheet");
    let a1 = sheet
        .cells
        .iter()
        .find(|c| (c.row, c.col) == (0, 0))
        .expect("A1 exists");
    assert_eq!(a1.value, CellValue::Text("New".to_string()));
    let formula = a1.formula.as_ref().expect("A1 should now be a formula cell");
    assert_eq!(formula.rgce, rgce);

    let sheet_out = read_zip_part(&output_path, "xl/worksheets/sheet1.bin");
    let (id_out, _) = find_cell_record(&sheet_out, 0, 0).expect("find A1 record in output");
    assert_eq!(id_out, 0x0008, "expected output A1 to be BrtFmlaString");

    let sst_out = read_zip_part(&output_path, "xl/sharedStrings.bin");
    let stats_out = read_shared_strings_stats(&sst_out);
    assert_eq!(
        stats_out.total_count,
        Some(stats_in.total_count.unwrap().saturating_sub(1))
    );
    assert_eq!(stats_out.unique_count, stats_in.unique_count);
    assert_eq!(stats_out.si_count, stats_in.si_count);
    assert_eq!(stats_out.strings, stats_in.strings);
}
