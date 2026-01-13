use std::collections::HashMap;
use std::fs::File;
use std::io::{Cursor, Read};
use std::path::Path;

use formula_xlsb::rgce::{encode_rgce_with_context, CellCoord};
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

fn dim_end_row_col(dim: &formula_xlsb::Dimension) -> (u32, u32) {
    (
        dim.start_row + dim.height.saturating_sub(1),
        dim.start_col + dim.width.saturating_sub(1),
    )
}

fn move_dimension_record_to_end(sheet_bin: &[u8]) -> Vec<u8> {
    const DIMENSION: u32 = 0x0094;
    const WORKSHEET_END: u32 = 0x0082;

    let mut cursor = Cursor::new(sheet_bin);
    let mut ranges: Vec<(u32, usize, usize)> = Vec::new();
    loop {
        let start = cursor.position() as usize;
        let Some(id) = biff12_varint::read_record_id(&mut cursor).expect("read record id") else {
            break;
        };
        let Some(len) = biff12_varint::read_record_len(&mut cursor).expect("read record len")
        else {
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

fn sheet_has_cell(sheet_bin: &[u8], target_row: u32, target_col: u32) -> bool {
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
                        return true;
                    }
                }
            }
            _ => {}
        }
    }

    false
}

#[test]
fn patch_sheet_bin_can_insert_into_missing_row_and_expand_dimension() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number(0, 0, 1.0);

    let bytes = builder.build_bytes();
    let tmpdir = tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("insert-missing-row-input.xlsb");
    let output_path = tmpdir.path().join("insert-missing-row-output.xlsb");
    std::fs::write(&input_path, bytes).expect("write input workbook");

    let wb = XlsbWorkbook::open(&input_path).expect("open workbook");
    let sheet_part = wb.sheet_metas()[0].part_path.clone();
    let sheet_bin = read_zip_part(&input_path, &sheet_part);

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

    wb.save_with_part_overrides(
        &output_path,
        &HashMap::from([(sheet_part.clone(), patched_sheet_bin)]),
    )
    .expect("write patched workbook");

    let wb2 = XlsbWorkbook::open(&output_path).expect("open patched workbook");
    let sheet = wb2.read_sheet(0).expect("read patched sheet");

    let coords: Vec<(u32, u32)> = sheet.cells.iter().map(|c| (c.row, c.col)).collect();
    assert_eq!(coords, vec![(0, 0), (5, 3)]);

    let original = sheet
        .cells
        .iter()
        .find(|c| (c.row, c.col) == (0, 0))
        .expect("original cell exists");
    assert_eq!(original.value, CellValue::Number(1.0));

    let inserted = sheet
        .cells
        .iter()
        .find(|c| (c.row, c.col) == (5, 3))
        .expect("inserted cell exists");
    assert_eq!(inserted.value, CellValue::Number(99.0));

    let dim = sheet.dimension.expect("dimension exists");
    let (end_row, end_col) = dim_end_row_col(&dim);
    assert_eq!(dim.start_row, 0);
    assert_eq!(dim.start_col, 0);
    assert_eq!(end_row, 5);
    assert_eq!(end_col, 3);
}

#[test]
fn patch_sheet_bin_can_insert_formula_with_rgcb_and_expand_dimension() {
    let ctx = formula_xlsb::workbook_context::WorkbookContext::default();

    let encoded =
        encode_rgce_with_context("=SUM({4,5})", &ctx, CellCoord::new(5, 3)).expect("encode rgce");
    assert!(
        !encoded.rgcb.is_empty(),
        "expected array formula encoding to produce rgcb bytes"
    );

    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number(0, 0, 1.0);

    let bytes = builder.build_bytes();
    let tmpdir = tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("insert-formula-rgcb-input.xlsb");
    let output_path = tmpdir.path().join("insert-formula-rgcb-output.xlsb");
    std::fs::write(&input_path, bytes).expect("write input workbook");

    let wb = XlsbWorkbook::open(&input_path).expect("open workbook");
    let sheet_part = wb.sheet_metas()[0].part_path.clone();
    let sheet_bin = read_zip_part(&input_path, &sheet_part);

    let patched_sheet_bin = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 5,
            col: 3,
            new_value: CellValue::Number(9.0),
            new_formula: Some(encoded.rgce.clone()),
            new_rgcb: Some(encoded.rgcb.clone()),
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
            clear_formula: false,
        }],
    )
    .expect("patch sheet bin");

    wb.save_with_part_overrides(
        &output_path,
        &HashMap::from([(sheet_part.clone(), patched_sheet_bin)]),
    )
    .expect("write patched workbook");

    let wb2 = XlsbWorkbook::open(&output_path).expect("open patched workbook");
    let sheet = wb2.read_sheet(0).expect("read patched sheet");

    let coords: Vec<(u32, u32)> = sheet.cells.iter().map(|c| (c.row, c.col)).collect();
    assert_eq!(coords, vec![(0, 0), (5, 3)]);

    let inserted = sheet
        .cells
        .iter()
        .find(|c| (c.row, c.col) == (5, 3))
        .expect("inserted cell exists");
    assert_eq!(inserted.value, CellValue::Number(9.0));
    let formula = inserted.formula.as_ref().expect("inserted formula");
    assert_eq!(formula.extra, encoded.rgcb);
    assert_eq!(formula.text.as_deref(), Some("SUM({4,5})"));

    let dim = sheet.dimension.expect("dimension exists");
    let (end_row, end_col) = dim_end_row_col(&dim);
    assert_eq!(dim.start_row, 0);
    assert_eq!(dim.start_col, 0);
    assert_eq!(end_row, 5);
    assert_eq!(end_col, 3);
}

#[test]
fn patch_sheet_bin_rejects_inserting_formula_that_requires_rgcb_without_rgcb() {
    let ctx = formula_xlsb::workbook_context::WorkbookContext::default();

    let encoded =
        encode_rgce_with_context("=SUM({4,5})", &ctx, CellCoord::new(5, 3)).expect("encode rgce");
    assert!(
        !encoded.rgcb.is_empty(),
        "expected array formula encoding to produce rgcb bytes"
    );

    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number(0, 0, 1.0);
    let bytes = builder.build_bytes();

    let tmpdir = tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("insert-formula-missing-rgcb-input.xlsb");
    std::fs::write(&input_path, bytes).expect("write input workbook");

    let wb = XlsbWorkbook::open(&input_path).expect("open workbook");
    let sheet_part = wb.sheet_metas()[0].part_path.clone();
    let sheet_bin = read_zip_part(&input_path, &sheet_part);

    let err = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 5,
            col: 3,
            new_value: CellValue::Number(9.0),
            new_formula: Some(encoded.rgce.clone()),
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
        }],
    )
    .expect_err("expected InvalidInput when inserting formula without rgcb");

    match err {
        formula_xlsb::Error::Io(io_err) => {
            assert_eq!(io_err.kind(), std::io::ErrorKind::InvalidInput);
            assert!(
                io_err.to_string().contains("set CellEdit.new_rgcb"),
                "expected error to instruct caller to set CellEdit.new_rgcb, got: {io_err}"
            );
        }
        other => panic!("expected InvalidInput, got {other:?}"),
    }
}

#[test]
fn patch_sheet_bin_can_insert_into_existing_row_in_column_order() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number(0, 0, 1.0);

    let bytes = builder.build_bytes();
    let tmpdir = tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("insert-existing-row-input.xlsb");
    let output_path = tmpdir.path().join("insert-existing-row-output.xlsb");
    std::fs::write(&input_path, bytes).expect("write input workbook");

    let wb = XlsbWorkbook::open(&input_path).expect("open workbook");
    let sheet_part = wb.sheet_metas()[0].part_path.clone();
    let sheet_bin = read_zip_part(&input_path, &sheet_part);

    let patched_sheet_bin = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 0,
            col: 10,
            new_value: CellValue::Number(42.0),
            new_formula: None,
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
            clear_formula: false,
        }],
    )
    .expect("patch sheet bin");

    wb.save_with_part_overrides(
        &output_path,
        &HashMap::from([(sheet_part.clone(), patched_sheet_bin)]),
    )
    .expect("write patched workbook");

    let wb2 = XlsbWorkbook::open(&output_path).expect("open patched workbook");
    let sheet = wb2.read_sheet(0).expect("read patched sheet");

    let coords: Vec<(u32, u32)> = sheet.cells.iter().map(|c| (c.row, c.col)).collect();
    assert_eq!(coords, vec![(0, 0), (0, 10)]);

    let inserted = sheet
        .cells
        .iter()
        .find(|c| (c.row, c.col) == (0, 10))
        .expect("inserted cell exists");
    assert_eq!(inserted.value, CellValue::Number(42.0));

    let dim = sheet.dimension.expect("dimension exists");
    let (end_row, end_col) = dim_end_row_col(&dim);
    assert_eq!(dim.start_row, 0);
    assert_eq!(dim.start_col, 0);
    assert_eq!(end_row, 0);
    assert_eq!(end_col, 10);
}

#[test]
fn patch_sheet_bin_can_expand_dimension_when_brtwsdim_is_after_sheetdata() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number(0, 0, 1.0);

    let bytes = builder.build_bytes();
    let tmpdir = tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("dimension-after-sheetdata-input.xlsb");
    std::fs::write(&input_path, bytes).expect("write input workbook");

    let wb = XlsbWorkbook::open(&input_path).expect("open workbook");
    let sheet_part = wb.sheet_metas()[0].part_path.clone();
    let sheet_bin = read_zip_part(&input_path, &sheet_part);

    let moved = move_dimension_record_to_end(&sheet_bin);
    let patched = patch_sheet_bin(
        &moved,
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

    assert!(sheet_has_cell(&patched, 5, 3));
    assert_eq!(read_dimension_bounds(&patched), Some((0, 5, 0, 3)));
}

#[test]
fn patch_sheet_bin_does_not_materialize_missing_blank_cells() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number(0, 0, 1.0);

    let bytes = builder.build_bytes();
    let tmpdir = tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("insert-blank-noop-input.xlsb");
    std::fs::write(&input_path, bytes).expect("write input workbook");

    let wb = XlsbWorkbook::open(&input_path).expect("open workbook");
    let sheet_part = wb.sheet_metas()[0].part_path.clone();
    let sheet_bin = read_zip_part(&input_path, &sheet_part);

    let patched_sheet_bin = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 5,
            col: 3,
            new_value: CellValue::Blank,
            new_formula: None,
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
            clear_formula: false,
        }],
    )
    .expect("patch sheet bin");

    assert_eq!(patched_sheet_bin, sheet_bin);
}

#[test]
fn patch_sheet_bin_can_patch_style_without_changing_value() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number(0, 0, 1.0);

    let bytes = builder.build_bytes();
    let tmpdir = tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("style-only-update-input.xlsb");
    let output_path = tmpdir.path().join("style-only-update-output.xlsb");
    std::fs::write(&input_path, bytes).expect("write input workbook");

    let wb = XlsbWorkbook::open(&input_path).expect("open workbook");
    let sheet_part = wb.sheet_metas()[0].part_path.clone();
    let sheet_bin = read_zip_part(&input_path, &sheet_part);

    let patched_sheet_bin = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Number(1.0),
            new_formula: None,
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: Some(7),
            clear_formula: false,
        }],
    )
    .expect("patch sheet bin");

    wb.save_with_part_overrides(
        &output_path,
        &HashMap::from([(sheet_part.clone(), patched_sheet_bin)]),
    )
    .expect("write patched workbook");

    let wb2 = XlsbWorkbook::open(&output_path).expect("open patched workbook");
    let sheet = wb2.read_sheet(0).expect("read patched sheet");

    let cell = sheet
        .cells
        .iter()
        .find(|c| (c.row, c.col) == (0, 0))
        .expect("expected A1 cell");
    assert_eq!(cell.value, CellValue::Number(1.0));
    assert_eq!(cell.style, 7);
}

#[test]
fn patch_sheet_bin_can_insert_cell_with_style_override() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number(0, 0, 1.0);

    let bytes = builder.build_bytes();
    let tmpdir = tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("style-insert-input.xlsb");
    let output_path = tmpdir.path().join("style-insert-output.xlsb");
    std::fs::write(&input_path, bytes).expect("write input workbook");

    let wb = XlsbWorkbook::open(&input_path).expect("open workbook");
    let sheet_part = wb.sheet_metas()[0].part_path.clone();
    let sheet_bin = read_zip_part(&input_path, &sheet_part);

    let patched_sheet_bin = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 1,
            col: 1,
            new_value: CellValue::Number(2.0),
            new_formula: None,
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: Some(5),
            clear_formula: false,
        }],
    )
    .expect("patch sheet bin");

    wb.save_with_part_overrides(
        &output_path,
        &HashMap::from([(sheet_part.clone(), patched_sheet_bin)]),
    )
    .expect("write patched workbook");

    let wb2 = XlsbWorkbook::open(&output_path).expect("open patched workbook");
    let sheet = wb2.read_sheet(0).expect("read patched sheet");

    let cell = sheet
        .cells
        .iter()
        .find(|c| (c.row, c.col) == (1, 1))
        .expect("expected inserted B2 cell");
    assert_eq!(cell.value, CellValue::Number(2.0));
    assert_eq!(cell.style, 5);
}
