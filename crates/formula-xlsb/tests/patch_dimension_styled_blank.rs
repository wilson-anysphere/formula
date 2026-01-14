use std::collections::HashMap;
use std::fs::File;
use std::io::{Cursor, Read};
use std::path::Path;

use formula_xlsb::{
    patch_sheet_bin, patch_sheet_bin_streaming, CellEdit, CellValue, Dimension, XlsbWorkbook,
};
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

fn dim_end_row_col(dim: &Dimension) -> (u32, u32) {
    (
        dim.start_row + dim.height.saturating_sub(1),
        dim.start_col + dim.width.saturating_sub(1),
    )
}

#[test]
fn patch_sheet_bin_inserts_styled_blank_cell_and_expands_dimension() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number(0, 0, 1.0); // A1 only -> dimension A1

    let bytes = builder.build_bytes();
    let tmpdir = tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("input.xlsb");
    let output_path = tmpdir.path().join("output.xlsb");
    std::fs::write(&input_path, bytes).expect("write input workbook");

    let wb = XlsbWorkbook::open(&input_path).expect("open workbook");
    let sheet_part = wb.sheet_metas()[0].part_path.clone();
    let sheet_bin = read_zip_part(&input_path, &sheet_part);

    let patched_sheet_bin = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 10,
            col: 10,
            new_value: CellValue::Blank,
            new_style: Some(5),
            clear_formula: false,
            new_formula: None,
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
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

    let inserted = sheet
        .cells
        .iter()
        .find(|c| (c.row, c.col) == (10, 10))
        .expect("inserted cell exists");
    assert_eq!(inserted.value, CellValue::Blank);
    assert_eq!(inserted.style, 5);

    let dim = sheet.dimension.expect("dimension exists");
    let (end_row, end_col) = dim_end_row_col(&dim);
    assert_eq!(dim.start_row, 0);
    assert_eq!(dim.start_col, 0);
    assert_eq!((end_row, end_col), (10, 10));
}

#[test]
fn patch_sheet_bin_streaming_inserts_styled_blank_cell_and_expands_dimension() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number(0, 0, 1.0); // A1 only -> dimension A1

    let bytes = builder.build_bytes();
    let tmpdir = tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("input.xlsb");
    let output_path = tmpdir.path().join("output.xlsb");
    std::fs::write(&input_path, bytes).expect("write input workbook");

    let wb = XlsbWorkbook::open(&input_path).expect("open workbook");
    let sheet_part = wb.sheet_metas()[0].part_path.clone();
    let sheet_bin = read_zip_part(&input_path, &sheet_part);

    let edits = [CellEdit {
        row: 10,
        col: 10,
        new_value: CellValue::Blank,
        new_style: Some(5),
        clear_formula: false,
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
    }];

    let mut patched_sheet_bin = Vec::new();
    let changed = patch_sheet_bin_streaming(Cursor::new(sheet_bin), &mut patched_sheet_bin, &edits)
        .expect("patch sheet bin streaming");
    assert!(changed, "expected streaming patcher to report changes");

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
        .find(|c| (c.row, c.col) == (10, 10))
        .expect("inserted cell exists");
    assert_eq!(inserted.value, CellValue::Blank);
    assert_eq!(inserted.style, 5);

    let dim = sheet.dimension.expect("dimension exists");
    let (end_row, end_col) = dim_end_row_col(&dim);
    assert_eq!(dim.start_row, 0);
    assert_eq!(dim.start_col, 0);
    assert_eq!((end_row, end_col), (10, 10));
}
