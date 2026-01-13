use std::collections::HashMap;
use std::fs::File;
use std::io::{Cursor, Read};
use std::path::Path;

use formula_xlsb::{biff12_varint, patch_sheet_bin, CellEdit, CellValue, XlsbWorkbook};
use tempfile::tempdir;

mod fixture_builder;
use fixture_builder::XlsbFixtureBuilder;

fn read_zip_part(path: &Path, part_path: &str) -> Vec<u8> {
    let file = File::open(path).expect("open xlsb");
    let mut zip = zip::ZipArchive::new(file).expect("open zip");
    let mut entry = zip.by_name(part_path).expect("find part");
    let mut bytes = Vec::with_capacity(entry.size() as usize);
    entry.read_to_end(&mut bytes).expect("read part bytes");
    bytes
}

fn find_row_record_payload(sheet_bin: &[u8], target_row: u32) -> Option<Vec<u8>> {
    const SHEETDATA: u32 = 0x0091;
    const SHEETDATA_END: u32 = 0x0092;
    const ROW: u32 = 0x0000;

    let mut cursor = Cursor::new(sheet_bin);
    let mut in_sheet_data = false;

    loop {
        let id = match biff12_varint::read_record_id(&mut cursor).ok().flatten() {
            Some(id) => id,
            None => break,
        };
        let len = biff12_varint::read_record_len(&mut cursor).ok().flatten()? as usize;
        let mut payload = vec![0u8; len];
        cursor.read_exact(&mut payload).ok()?;

        match id {
            SHEETDATA => in_sheet_data = true,
            SHEETDATA_END => in_sheet_data = false,
            ROW if in_sheet_data => {
                if payload.len() < 4 {
                    continue;
                }
                let row = u32::from_le_bytes(payload[0..4].try_into().unwrap());
                if row == target_row {
                    return Some(payload);
                }
            }
            _ => {}
        }
    }

    None
}

fn count_row_records_in_sheetdata(sheet_bin: &[u8]) -> usize {
    const SHEETDATA: u32 = 0x0091;
    const SHEETDATA_END: u32 = 0x0092;
    const ROW: u32 = 0x0000;

    let mut cursor = Cursor::new(sheet_bin);
    let mut in_sheet_data = false;
    let mut count = 0usize;

    loop {
        let id = match biff12_varint::read_record_id(&mut cursor).ok().flatten() {
            Some(id) => id,
            None => break,
        };
        let len = match biff12_varint::read_record_len(&mut cursor).ok().flatten() {
            Some(len) => len as usize,
            None => break,
        };
        cursor.set_position(cursor.position() + len as u64);

        match id {
            SHEETDATA => in_sheet_data = true,
            SHEETDATA_END => in_sheet_data = false,
            ROW if in_sheet_data => count += 1,
            _ => {}
        }
    }

    count
}

#[test]
fn patch_sheet_bin_inserts_rows_using_source_row_template_payload() {
    let row_trailing = vec![0xBA, 0xAD, 0xF0, 0x0D, 0x12, 0x34, 0x56, 0x78];

    // Only row 5 exists in the input; row 0/12 are missing and should be inserted by patching.
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_row_record_trailing_bytes(row_trailing.clone());
    builder.set_cell_number(5, 0, 1.0);

    let bytes = builder.build_bytes();
    let tmpdir = tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("row-template-input.xlsb");
    let output_path = tmpdir.path().join("row-template-output.xlsb");
    std::fs::write(&input_path, bytes).expect("write input workbook");

    let wb = XlsbWorkbook::open(&input_path).expect("open workbook");
    let sheet_part = wb.sheet_metas()[0].part_path.clone();
    let sheet_bin = read_zip_part(&input_path, &sheet_part);

    let template_payload = find_row_record_payload(&sheet_bin, 5).expect("existing row 5 record");
    assert_eq!(template_payload.len(), 4 + row_trailing.len());
    assert_eq!(template_payload[4..], row_trailing);

    let patched_sheet_bin = patch_sheet_bin(
        &sheet_bin,
        &[
            CellEdit {
                row: 0,
                col: 3,
                new_value: CellValue::Number(99.0),
                new_formula: None,
                new_rgcb: None,
                new_formula_flags: None,
                shared_string_index: None,
                new_style: None,
                clear_formula: false,
            },
            CellEdit {
                row: 12,
                col: 1,
                new_value: CellValue::Number(42.0),
                new_formula: None,
                new_rgcb: None,
                new_formula_flags: None,
                shared_string_index: None,
                new_style: None,
                clear_formula: false,
            },
        ],
    )
    .expect("patch sheet bin");

    let inserted_before = find_row_record_payload(&patched_sheet_bin, 0).expect("inserted row 0");
    assert_eq!(inserted_before.len(), template_payload.len());
    assert_eq!(inserted_before[4..], template_payload[4..]);
    assert_eq!(inserted_before[0..4], 0u32.to_le_bytes());

    let inserted_after = find_row_record_payload(&patched_sheet_bin, 12).expect("inserted row 12");
    assert_eq!(inserted_after.len(), template_payload.len());
    assert_eq!(inserted_after[4..], template_payload[4..]);
    assert_eq!(inserted_after[0..4], 12u32.to_le_bytes());

    wb.save_with_part_overrides(
        &output_path,
        &HashMap::from([(sheet_part.clone(), patched_sheet_bin)]),
    )
    .expect("write patched workbook");

    let wb2 = XlsbWorkbook::open(&output_path).expect("open patched workbook");
    let sheet = wb2.read_sheet(0).expect("read patched sheet");

    let inserted = sheet
        .cells
        .iter()
        .find(|c| (c.row, c.col) == (0, 3))
        .expect("inserted cell (0, 3) exists");
    assert_eq!(inserted.value, CellValue::Number(99.0));

    let inserted2 = sheet
        .cells
        .iter()
        .find(|c| (c.row, c.col) == (12, 1))
        .expect("inserted cell (12, 1) exists");
    assert_eq!(inserted2.value, CellValue::Number(42.0));
}

#[test]
fn patch_sheet_bin_inserts_minimal_row_record_when_no_template_exists() {
    // SheetData is empty, so no BrtRow records exist to use as a template.
    let builder = XlsbFixtureBuilder::new();
    let bytes = builder.build_bytes();

    let tmpdir = tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("row-no-template-input.xlsb");
    let output_path = tmpdir.path().join("row-no-template-output.xlsb");
    std::fs::write(&input_path, bytes).expect("write input workbook");

    let wb = XlsbWorkbook::open(&input_path).expect("open workbook");
    let sheet_part = wb.sheet_metas()[0].part_path.clone();
    let sheet_bin = read_zip_part(&input_path, &sheet_part);
    assert_eq!(count_row_records_in_sheetdata(&sheet_bin), 0);

    let patched_sheet_bin = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 5,
            col: 3,
            new_value: CellValue::Number(99.0),
            new_formula: None,
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
            clear_formula: false,
        }],
    )
    .expect("patch sheet bin");

    let inserted_row = find_row_record_payload(&patched_sheet_bin, 5).expect("inserted row 5");
    assert_eq!(inserted_row.as_slice(), 5u32.to_le_bytes());

    wb.save_with_part_overrides(
        &output_path,
        &HashMap::from([(sheet_part.clone(), patched_sheet_bin)]),
    )
    .expect("write patched workbook");

    let wb2 = XlsbWorkbook::open(&output_path).expect("open patched workbook");
    let sheet = wb2.read_sheet(0).expect("read patched sheet");
    let inserted_cell = sheet
        .cells
        .iter()
        .find(|c| (c.row, c.col) == (5, 3))
        .expect("inserted cell exists");
    assert_eq!(inserted_cell.value, CellValue::Number(99.0));
}
