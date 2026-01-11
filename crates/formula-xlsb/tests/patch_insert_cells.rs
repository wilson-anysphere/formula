use std::collections::HashMap;
use std::fs::File;
use std::io::Read;
use std::path::Path;

use formula_xlsb::{patch_sheet_bin, CellEdit, CellValue, XlsbWorkbook};
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
            shared_string_index: None,
        }],
    )
    .expect("patch sheet bin");

    assert_eq!(patched_sheet_bin, sheet_bin);
}
