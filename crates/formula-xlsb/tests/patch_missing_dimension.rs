use std::collections::HashMap;
use std::io::{Cursor, Read};

use formula_biff::encode_rgce_with_rgcb;
use formula_xlsb::{
    biff12_varint, patch_sheet_bin, patch_sheet_bin_streaming, CellEdit, CellValue, XlsbWorkbook,
};
use tempfile::tempdir;

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

fn remove_dimension_record(sheet_bin: &[u8]) -> Vec<u8> {
    const DIMENSION: u32 = 0x0094;
    let mut cursor = Cursor::new(sheet_bin);
    let mut out = Vec::with_capacity(sheet_bin.len());

    loop {
        let record_start = cursor.position() as usize;
        let Some(id) = biff12_varint::read_record_id(&mut cursor).ok().flatten() else {
            break;
        };
        let Some(len) = biff12_varint::read_record_len(&mut cursor).ok().flatten() else {
            break;
        };
        let payload_start = cursor.position() as usize;
        let payload_end = payload_start + len as usize;
        cursor.set_position(payload_end as u64);
        if payload_end > sheet_bin.len() {
            break;
        }
        if id == DIMENSION {
            continue;
        }
        out.extend_from_slice(&sheet_bin[record_start..payload_end]);
    }

    out
}

fn read_dimension_bounds(sheet_bin: &[u8]) -> Option<(u32, u32, u32, u32)> {
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

#[test]
fn patch_sheet_bin_inserts_missing_dimension_before_sheetdata() {
    let builder = XlsbFixtureBuilder::new();
    let xlsb_bytes = builder.build_bytes();
    let sheet_bin = read_sheet_bin(xlsb_bytes.clone());
    let sheet_no_dim = remove_dimension_record(&sheet_bin);
    assert_eq!(read_dimension_bounds(&sheet_no_dim), None);

    let edits = [CellEdit {
        row: 4,
        col: 2,
        new_value: CellValue::Number(123.0),
        new_style: None,
        clear_formula: false,
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
    }];
    let patched_sheet_bin = patch_sheet_bin(&sheet_no_dim, &edits).expect("patch_sheet_bin");
    assert_eq!(
        read_dimension_bounds(&patched_sheet_bin),
        Some((4, 4, 2, 2))
    );

    let tmpdir = tempdir().expect("tempdir");
    let input_path = tmpdir.path().join("input.xlsb");
    let output_path = tmpdir.path().join("output.xlsb");
    std::fs::write(&input_path, xlsb_bytes).expect("write input");

    let wb = XlsbWorkbook::open(&input_path).expect("open workbook");
    let sheet_part = wb.sheet_metas()[0].part_path.clone();
    wb.save_with_part_overrides(
        &output_path,
        &HashMap::from([(sheet_part, patched_sheet_bin)]),
    )
    .expect("write patched workbook");

    let wb2 = XlsbWorkbook::open(&output_path).expect("open patched");
    let sheet = wb2.read_sheet(0).expect("read patched sheet");
    let dim = sheet.dimension.expect("expected synthesized DIMENSION");
    assert_eq!(dim.start_row, 4);
    assert_eq!(dim.start_col, 2);
    assert_eq!(dim.height, 1);
    assert_eq!(dim.width, 1);
}

#[test]
fn patch_sheet_bin_synthesized_dimension_includes_existing_cells() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number(0, 0, 1.0);
    let xlsb_bytes = builder.build_bytes();
    let sheet_bin = read_sheet_bin(xlsb_bytes.clone());
    let sheet_no_dim = remove_dimension_record(&sheet_bin);
    assert_eq!(read_dimension_bounds(&sheet_no_dim), None);

    let edits = [CellEdit {
        row: 4,
        col: 2,
        new_value: CellValue::Number(123.0),
        new_style: None,
        clear_formula: false,
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
    }];
    let patched_sheet_bin = patch_sheet_bin(&sheet_no_dim, &edits).expect("patch_sheet_bin");
    assert_eq!(
        read_dimension_bounds(&patched_sheet_bin),
        Some((0, 4, 0, 2))
    );

    let tmpdir = tempdir().expect("tempdir");
    let input_path = tmpdir.path().join("input.xlsb");
    let output_path = tmpdir.path().join("output.xlsb");
    std::fs::write(&input_path, xlsb_bytes).expect("write input");

    let wb = XlsbWorkbook::open(&input_path).expect("open workbook");
    let sheet_part = wb.sheet_metas()[0].part_path.clone();
    wb.save_with_part_overrides(
        &output_path,
        &HashMap::from([(sheet_part, patched_sheet_bin)]),
    )
    .expect("write patched workbook");

    let wb2 = XlsbWorkbook::open(&output_path).expect("open patched");
    let sheet = wb2.read_sheet(0).expect("read patched sheet");
    let dim = sheet.dimension.expect("expected synthesized DIMENSION");
    assert_eq!(dim.start_row, 0);
    assert_eq!(dim.start_col, 0);
    assert_eq!(dim.height, 5);
    assert_eq!(dim.width, 3);
}

#[test]
fn patch_sheet_bin_streaming_inserts_missing_dimension_before_sheetdata() {
    let builder = XlsbFixtureBuilder::new();
    let xlsb_bytes = builder.build_bytes();
    let sheet_bin = read_sheet_bin(xlsb_bytes.clone());
    let sheet_no_dim = remove_dimension_record(&sheet_bin);
    assert_eq!(read_dimension_bounds(&sheet_no_dim), None);

    let edits = [CellEdit {
        row: 4,
        col: 2,
        new_value: CellValue::Number(123.0),
        new_style: None,
        clear_formula: false,
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
    }];

    let mut patched_stream = Vec::new();
    let changed =
        patch_sheet_bin_streaming(Cursor::new(&sheet_no_dim), &mut patched_stream, &edits)
            .expect("patch_sheet_bin_streaming");
    assert!(changed, "expected streaming patcher to report changes");
    assert_eq!(read_dimension_bounds(&patched_stream), Some((4, 4, 2, 2)));

    let tmpdir = tempdir().expect("tempdir");
    let input_path = tmpdir.path().join("input.xlsb");
    let output_path = tmpdir.path().join("output.xlsb");
    std::fs::write(&input_path, xlsb_bytes).expect("write input");

    let wb = XlsbWorkbook::open(&input_path).expect("open workbook");
    let sheet_part = wb.sheet_metas()[0].part_path.clone();
    wb.save_with_part_overrides(&output_path, &HashMap::from([(sheet_part, patched_stream)]))
        .expect("write patched workbook");

    let wb2 = XlsbWorkbook::open(&output_path).expect("open patched");
    let sheet = wb2.read_sheet(0).expect("read patched sheet");
    let dim = sheet.dimension.expect("expected synthesized DIMENSION");
    assert_eq!(dim.start_row, 4);
    assert_eq!(dim.start_col, 2);
    assert_eq!(dim.height, 1);
    assert_eq!(dim.width, 1);
}

#[test]
fn patch_sheet_bin_streaming_synthesized_dimension_includes_existing_cells() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number(0, 0, 1.0);
    let xlsb_bytes = builder.build_bytes();
    let sheet_bin = read_sheet_bin(xlsb_bytes.clone());
    let sheet_no_dim = remove_dimension_record(&sheet_bin);
    assert_eq!(read_dimension_bounds(&sheet_no_dim), None);

    let edits = [CellEdit {
        row: 4,
        col: 2,
        new_value: CellValue::Number(123.0),
        new_style: None,
        clear_formula: false,
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
    }];

    let mut patched_stream = Vec::new();
    let changed =
        patch_sheet_bin_streaming(Cursor::new(&sheet_no_dim), &mut patched_stream, &edits)
            .expect("patch_sheet_bin_streaming");
    assert!(changed, "expected streaming patcher to report changes");
    assert_eq!(read_dimension_bounds(&patched_stream), Some((0, 4, 0, 2)));

    let tmpdir = tempdir().expect("tempdir");
    let input_path = tmpdir.path().join("input.xlsb");
    let output_path = tmpdir.path().join("output.xlsb");
    std::fs::write(&input_path, xlsb_bytes).expect("write input");

    let wb = XlsbWorkbook::open(&input_path).expect("open workbook");
    let sheet_part = wb.sheet_metas()[0].part_path.clone();
    wb.save_with_part_overrides(&output_path, &HashMap::from([(sheet_part, patched_stream)]))
        .expect("write patched workbook");

    let wb2 = XlsbWorkbook::open(&output_path).expect("open patched");
    let sheet = wb2.read_sheet(0).expect("read patched sheet");
    let dim = sheet.dimension.expect("expected synthesized DIMENSION");
    assert_eq!(dim.start_row, 0);
    assert_eq!(dim.start_col, 0);
    assert_eq!(dim.height, 5);
    assert_eq!(dim.width, 3);
}

#[test]
fn patch_sheet_bin_inserts_missing_dimension_when_converting_value_cell_to_formula() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number(0, 0, 1.0);
    let xlsb_bytes = builder.build_bytes();
    let sheet_bin = read_sheet_bin(xlsb_bytes.clone());
    let sheet_no_dim = remove_dimension_record(&sheet_bin);
    assert_eq!(read_dimension_bounds(&sheet_no_dim), None);

    let encoded = encode_rgce_with_rgcb("=1+1").expect("encode formula");
    let edits = [CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Number(2.0),
        new_style: None,
        clear_formula: false,
        new_formula: Some(encoded.rgce),
        new_rgcb: Some(encoded.rgcb),
        new_formula_flags: None,
        shared_string_index: None,
    }];
    let patched_sheet_bin = patch_sheet_bin(&sheet_no_dim, &edits).expect("patch_sheet_bin");
    assert_eq!(
        read_dimension_bounds(&patched_sheet_bin),
        Some((0, 0, 0, 0))
    );

    let tmpdir = tempdir().expect("tempdir");
    let input_path = tmpdir.path().join("input.xlsb");
    let output_path = tmpdir.path().join("output.xlsb");
    std::fs::write(&input_path, xlsb_bytes).expect("write input");

    let wb = XlsbWorkbook::open(&input_path).expect("open workbook");
    let sheet_part = wb.sheet_metas()[0].part_path.clone();
    wb.save_with_part_overrides(
        &output_path,
        &HashMap::from([(sheet_part, patched_sheet_bin)]),
    )
    .expect("write patched workbook");

    let wb2 = XlsbWorkbook::open(&output_path).expect("open patched");
    let sheet = wb2.read_sheet(0).expect("read patched sheet");
    let dim = sheet.dimension.expect("expected synthesized DIMENSION");
    assert_eq!(dim.start_row, 0);
    assert_eq!(dim.start_col, 0);
    assert_eq!(dim.height, 1);
    assert_eq!(dim.width, 1);
}

#[test]
fn patch_sheet_bin_streaming_inserts_missing_dimension_when_converting_value_cell_to_formula() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number(0, 0, 1.0);
    let xlsb_bytes = builder.build_bytes();
    let sheet_bin = read_sheet_bin(xlsb_bytes.clone());
    let sheet_no_dim = remove_dimension_record(&sheet_bin);
    assert_eq!(read_dimension_bounds(&sheet_no_dim), None);

    let encoded = encode_rgce_with_rgcb("=1+1").expect("encode formula");
    let edits = [CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Number(2.0),
        new_style: None,
        clear_formula: false,
        new_formula: Some(encoded.rgce),
        new_rgcb: Some(encoded.rgcb),
        new_formula_flags: None,
        shared_string_index: None,
    }];

    let mut patched_stream = Vec::new();
    let changed =
        patch_sheet_bin_streaming(Cursor::new(&sheet_no_dim), &mut patched_stream, &edits)
            .expect("patch_sheet_bin_streaming");
    assert!(changed, "expected streaming patcher to report changes");
    assert_eq!(read_dimension_bounds(&patched_stream), Some((0, 0, 0, 0)));

    let tmpdir = tempdir().expect("tempdir");
    let input_path = tmpdir.path().join("input.xlsb");
    let output_path = tmpdir.path().join("output.xlsb");
    std::fs::write(&input_path, xlsb_bytes).expect("write input");

    let wb = XlsbWorkbook::open(&input_path).expect("open workbook");
    let sheet_part = wb.sheet_metas()[0].part_path.clone();
    wb.save_with_part_overrides(&output_path, &HashMap::from([(sheet_part, patched_stream)]))
        .expect("write patched workbook");

    let wb2 = XlsbWorkbook::open(&output_path).expect("open patched");
    let sheet = wb2.read_sheet(0).expect("read patched sheet");
    let dim = sheet.dimension.expect("expected synthesized DIMENSION");
    assert_eq!(dim.start_row, 0);
    assert_eq!(dim.start_col, 0);
    assert_eq!(dim.height, 1);
    assert_eq!(dim.width, 1);
}
