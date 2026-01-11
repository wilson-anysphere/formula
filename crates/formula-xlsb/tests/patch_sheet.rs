use std::collections::HashMap;
use std::fs::File;
use std::io::{Cursor, Read};

use formula_biff::encode_rgce;
use formula_xlsb::{patch_sheet_bin, CellEdit, CellValue, XlsbWorkbook};
use pretty_assertions::assert_eq;
use tempfile::tempdir;

mod fixture_builder;
use fixture_builder::XlsbFixtureBuilder;

fn read_zip_part(path: &str, part_path: &str) -> Vec<u8> {
    let file = File::open(path).expect("open xlsb fixture");
    let mut zip = zip::ZipArchive::new(file).expect("open zip");
    let mut entry = zip.by_name(part_path).expect("find part");
    let mut bytes = Vec::with_capacity(entry.size() as usize);
    entry.read_to_end(&mut bytes).expect("read part bytes");
    bytes
}

#[test]
fn patch_sheet_bin_is_byte_identical_for_noop_numeric_edit() {
    let fixture = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/simple.xlsb");
    let wb = XlsbWorkbook::open(fixture).expect("open xlsb");
    let sheet_part = wb.sheet_metas()[0].part_path.clone();

    let sheet_bin = read_zip_part(fixture, &sheet_part);
    let patched = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 0,
            col: 1,
            new_value: CellValue::Number(42.5),
            new_formula: None,
        }],
    )
    .expect("patch sheet bin");

    assert_eq!(patched, sheet_bin);
}

#[test]
fn patch_sheet_bin_round_trips_numeric_cell_preserving_formula() {
    let fixture = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/simple.xlsb");
    let wb = XlsbWorkbook::open(fixture).expect("open xlsb");
    let sheet_part = wb.sheet_metas()[0].part_path.clone();

    let original_sheet = wb.read_sheet(0).expect("read original sheet");
    let original_formula_rgce = original_sheet
        .cells
        .iter()
        .find(|c| (c.row, c.col) == (0, 2))
        .and_then(|c| c.formula.as_ref())
        .expect("original formula present")
        .rgce
        .clone();

    let sheet_bin = read_zip_part(fixture, &sheet_part);
    let patched_sheet_bin = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Text("World".to_string()),
            new_formula: None,
        },
        CellEdit {
            row: 0,
            col: 1,
            new_value: CellValue::Number(100.0),
            new_formula: None,
        }],
    )
    .expect("patch sheet bin");

    let tmpdir = tempdir().expect("create temp dir");
    let out_path = tmpdir.path().join("patched.xlsb");

    wb.save_with_part_overrides(
        &out_path,
        &HashMap::from([(sheet_part.clone(), patched_sheet_bin)]),
    )
    .expect("write patched workbook");

    let patched = XlsbWorkbook::open(&out_path).expect("open patched workbook");
    let patched_sheet = patched.read_sheet(0).expect("read patched sheet");

    let mut cells = patched_sheet
        .cells
        .iter()
        .map(|c| ((c.row, c.col), c))
        .collect::<HashMap<_, _>>();

    assert_eq!(
        cells.remove(&(0, 0)).expect("patched A1").value,
        CellValue::Text("World".to_string())
    );
    assert_eq!(
        cells.remove(&(0, 1)).expect("patched B1").value,
        CellValue::Number(100.0)
    );

    let formula_cell = cells.remove(&(0, 2)).expect("formula cell still present");
    assert_eq!(formula_cell.value, CellValue::Number(85.0));
    let formula = formula_cell.formula.as_ref().expect("formula metadata preserved");
    assert_eq!(formula.text.as_deref(), Some("B1*2"));
    assert_eq!(formula.rgce.as_slice(), original_formula_rgce.as_slice());
}

#[test]
fn patch_sheet_bin_can_update_formula_cached_result() {
    let fixture = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/simple.xlsb");
    let wb = XlsbWorkbook::open(fixture).expect("open xlsb");
    let sheet_part = wb.sheet_metas()[0].part_path.clone();

    let original_sheet = wb.read_sheet(0).expect("read original sheet");
    let original_formula_rgce = original_sheet
        .cells
        .iter()
        .find(|c| (c.row, c.col) == (0, 2))
        .and_then(|c| c.formula.as_ref())
        .expect("original formula present")
        .rgce
        .clone();

    let sheet_bin = read_zip_part(fixture, &sheet_part);
    let patched_sheet_bin = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 0,
            col: 2,
            new_value: CellValue::Number(200.0),
            new_formula: None,
        }],
    )
    .expect("patch sheet bin");

    let tmpdir = tempdir().expect("create temp dir");
    let out_path = tmpdir.path().join("patched-formula.xlsb");

    wb.save_with_part_overrides(
        &out_path,
        &HashMap::from([(sheet_part.clone(), patched_sheet_bin)]),
    )
    .expect("write patched workbook");

    let patched = XlsbWorkbook::open(&out_path).expect("open patched workbook");
    let patched_sheet = patched.read_sheet(0).expect("read patched sheet");

    let formula_cell = patched_sheet
        .cells
        .iter()
        .find(|c| (c.row, c.col) == (0, 2))
        .expect("formula cell still present");

    assert_eq!(formula_cell.value, CellValue::Number(200.0));
    let formula = formula_cell.formula.as_ref().expect("formula metadata preserved");
    assert_eq!(formula.text.as_deref(), Some("B1*2"));
    assert_eq!(formula.rgce.as_slice(), original_formula_rgce.as_slice());
}

#[test]
fn patch_sheet_bin_can_update_formula_rgce_bytes() {
    let fixture = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/simple.xlsb");
    let wb = XlsbWorkbook::open(fixture).expect("open xlsb");
    let sheet_part = wb.sheet_metas()[0].part_path.clone();

    let original_sheet = wb.read_sheet(0).expect("read original sheet");
    let original_formula_rgce = original_sheet
        .cells
        .iter()
        .find(|c| (c.row, c.col) == (0, 2))
        .and_then(|c| c.formula.as_ref())
        .expect("original formula present")
        .rgce
        .clone();

    // The fixture formula rgce is `B1*2`. Rewrite the `PtgInt` literal from 2 to 3.
    let mut new_rgce = original_formula_rgce.clone();
    assert_eq!(new_rgce.len(), 11);
    assert_eq!(new_rgce[7], 0x1E, "expected PtgInt at offset 7");
    new_rgce[8] = 3;

    let sheet_bin = read_zip_part(fixture, &sheet_part);
    let patched_sheet_bin = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 0,
            col: 2,
            new_value: CellValue::Number(127.5),
            new_formula: Some(new_rgce.clone()),
        }],
    )
    .expect("patch sheet bin");

    let tmpdir = tempdir().expect("create temp dir");
    let out_path = tmpdir.path().join("patched-formula-rgce.xlsb");

    wb.save_with_part_overrides(
        &out_path,
        &HashMap::from([(sheet_part.clone(), patched_sheet_bin)]),
    )
    .expect("write patched workbook");

    let patched = XlsbWorkbook::open(&out_path).expect("open patched workbook");
    let patched_sheet = patched.read_sheet(0).expect("read patched sheet");

    let b1 = patched_sheet
        .cells
        .iter()
        .find(|c| (c.row, c.col) == (0, 1))
        .expect("B1 exists");
    assert_eq!(b1.value, CellValue::Number(42.5));

    let formula_cell = patched_sheet
        .cells
        .iter()
        .find(|c| (c.row, c.col) == (0, 2))
        .expect("formula cell still present");

    assert_eq!(formula_cell.value, CellValue::Number(127.5));
    let formula = formula_cell.formula.as_ref().expect("formula metadata preserved");
    assert_eq!(formula.text.as_deref(), Some("B1*3"));
    assert_eq!(formula.rgce.as_slice(), new_rgce.as_slice());
}

#[test]
fn patch_sheet_bin_can_update_formula_from_text() {
    let fixture = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/simple.xlsb");
    let wb = XlsbWorkbook::open(fixture).expect("open xlsb");
    let sheet_part = wb.sheet_metas()[0].part_path.clone();

    let rgce = encode_rgce("=IF(B1>0,B1*3,0)").expect("encode formula");

    let sheet_bin = read_zip_part(fixture, &sheet_part);
    let patched_sheet_bin = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 0,
            col: 2,
            new_value: CellValue::Number(1.0),
            new_formula: Some(rgce),
        }],
    )
    .expect("patch sheet bin");

    let tmpdir = tempdir().expect("create temp dir");
    let out_path = tmpdir.path().join("patched-formula-text.xlsb");

    wb.save_with_part_overrides(
        &out_path,
        &HashMap::from([(sheet_part.clone(), patched_sheet_bin)]),
    )
    .expect("write patched workbook");

    let patched = XlsbWorkbook::open(&out_path).expect("open patched workbook");
    let patched_sheet = patched.read_sheet(0).expect("read patched sheet");

    let formula_cell = patched_sheet
        .cells
        .iter()
        .find(|c| (c.row, c.col) == (0, 2))
        .expect("formula cell still present");

    assert_eq!(formula_cell.value, CellValue::Number(1.0));
    let formula = formula_cell.formula.as_ref().expect("formula metadata preserved");
    assert_eq!(formula.text.as_deref(), Some("IF(B1>0,B1*3,0)"));
}

#[test]
fn patch_sheet_bin_preserves_formula_trailing_bytes() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_sheet_name("Arrayy");

    // `PtgArray` placeholder token, followed by extra bytes (rgcb) describing the array constant.
    let rgce = vec![0x20, 0, 0, 0, 0, 0, 0, 0];
    let extra = vec![0xDE, 0xAD, 0xBE, 0xEF];
    builder.set_cell_formula_num(0, 0, 1.0, rgce, extra);

    let xlsb_bytes = builder.build_bytes();
    let mut zip = zip::ZipArchive::new(Cursor::new(xlsb_bytes)).expect("open in-memory xlsb zip");
    let mut entry = zip
        .by_name("xl/worksheets/sheet1.bin")
        .expect("find sheet1.bin");
    let mut sheet_bin = Vec::with_capacity(entry.size() as usize);
    entry.read_to_end(&mut sheet_bin).expect("read sheet bytes");

    // No-op patch (cached value unchanged): must reserialize byte-for-byte, including the extra
    // bytes after `rgce`.
    let patched = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Number(1.0),
            new_formula: None,
        }],
    )
    .expect("patch sheet bin");

    assert_eq!(patched, sheet_bin);
}

#[test]
fn save_with_edits_can_patch_rk_number_cells() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_sheet_name("Sheet1");
    builder.set_cell_number_rk(0, 1, 42.0);

    let bytes = builder.build_bytes();

    let tmpdir = tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("rk_input.xlsb");
    let output_path = tmpdir.path().join("rk_output.xlsb");
    std::fs::write(&input_path, bytes).expect("write input workbook");

    let wb = XlsbWorkbook::open(&input_path).expect("open generated xlsb");
    wb.save_with_edits(&output_path, 0, 0, 1, 100.0)
        .expect("save_with_edits");

    let wb2 = XlsbWorkbook::open(&output_path).expect("open patched workbook");
    let sheet = wb2.read_sheet(0).expect("read sheet");
    let b1 = sheet
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 1)
        .expect("B1 exists");
    assert_eq!(b1.value, CellValue::Number(100.0));
}

#[test]
fn patch_sheet_bin_is_byte_identical_for_noop_bool_error_blank_edits() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_sheet_name("Types");
    builder.set_cell_bool(0, 0, true);
    builder.set_cell_error(0, 1, 0x07);
    builder.set_cell_blank(0, 2);

    let xlsb_bytes = builder.build_bytes();
    let mut zip = zip::ZipArchive::new(Cursor::new(xlsb_bytes)).expect("open in-memory xlsb zip");
    let mut entry = zip
        .by_name("xl/worksheets/sheet1.bin")
        .expect("find sheet1.bin");
    let mut sheet_bin = Vec::with_capacity(entry.size() as usize);
    entry.read_to_end(&mut sheet_bin).expect("read sheet bytes");

    let patched = patch_sheet_bin(
        &sheet_bin,
        &[
            CellEdit {
                row: 0,
                col: 0,
                new_value: CellValue::Bool(true),
                new_formula: None,
            },
            CellEdit {
                row: 0,
                col: 1,
                new_value: CellValue::Error(0x07),
                new_formula: None,
            },
            CellEdit {
                row: 0,
                col: 2,
                new_value: CellValue::Blank,
                new_formula: None,
            },
        ],
    )
    .expect("patch sheet bin");

    assert_eq!(patched, sheet_bin);
}

#[test]
fn save_with_cell_edits_can_patch_bool_error_blank_cells() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_sheet_name("Types");
    builder.set_cell_bool(0, 0, true);
    builder.set_cell_error(0, 1, 0x07);
    builder.set_cell_blank(0, 2);

    let bytes = builder.build_bytes();
    let tmpdir = tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("types_input.xlsb");
    let output_path = tmpdir.path().join("types_output.xlsb");
    std::fs::write(&input_path, bytes).expect("write input workbook");

    let wb = XlsbWorkbook::open(&input_path).expect("open generated xlsb");
    wb.save_with_cell_edits(
        &output_path,
        0,
        &[
            CellEdit {
                row: 0,
                col: 0,
                new_value: CellValue::Bool(false),
                new_formula: None,
            },
            CellEdit {
                row: 0,
                col: 1,
                new_value: CellValue::Error(0x2A),
                new_formula: None,
            },
            // Patch the blank cell into a numeric value.
            CellEdit {
                row: 0,
                col: 2,
                new_value: CellValue::Number(123.0),
                new_formula: None,
            },
        ],
    )
    .expect("save_with_cell_edits");

    let wb2 = XlsbWorkbook::open(&output_path).expect("open patched workbook");
    let sheet = wb2.read_sheet(0).expect("read sheet");
    let mut cells = sheet
        .cells
        .iter()
        .map(|c| ((c.row, c.col), c.value.clone()))
        .collect::<HashMap<_, _>>();

    assert_eq!(cells.remove(&(0, 0)).expect("A1 exists"), CellValue::Bool(false));
    assert_eq!(
        cells.remove(&(0, 1)).expect("B1 exists"),
        CellValue::Error(0x2A)
    );
    assert_eq!(
        cells.remove(&(0, 2)).expect("C1 exists"),
        CellValue::Number(123.0)
    );
}
