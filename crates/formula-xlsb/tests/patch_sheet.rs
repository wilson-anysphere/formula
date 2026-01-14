use std::collections::HashMap;
use std::fs::File;
use std::io::{Cursor, Read};

use formula_biff::encode_rgce;
use formula_xlsb::rgce::{encode_rgce_with_context, CellCoord};
use formula_xlsb::{biff12_varint, patch_sheet_bin, CellEdit, CellValue, XlsbWorkbook};
use pretty_assertions::assert_eq;
use tempfile::tempdir;

mod fixture_builder;
use fixture_builder::{rgce, XlsbFixtureBuilder};

fn read_zip_part(path: &str, part_path: &str) -> Vec<u8> {
    let file = File::open(path).expect("open xlsb fixture");
    let mut zip = zip::ZipArchive::new(file).expect("open zip");
    let mut entry = zip.by_name(part_path).expect("find part");
    // Do not trust `ZipFile::size()` for allocation; ZIP metadata is untrusted and can
    // advertise enormous uncompressed sizes (zip-bomb style OOM).
    let mut bytes = Vec::new();
    entry.read_to_end(&mut bytes).expect("read part bytes");
    bytes
}

#[test]
fn patch_sheet_bin_rejects_out_of_range_cell_edit_row() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number(0, 0, 1.0);

    let xlsb_bytes = builder.build_bytes();
    let mut zip = zip::ZipArchive::new(Cursor::new(xlsb_bytes)).expect("open in-memory xlsb zip");
    let mut entry = zip
        .by_name("xl/worksheets/sheet1.bin")
        .expect("find sheet1.bin");
    let mut sheet_bin = Vec::with_capacity(entry.size() as usize);
    entry.read_to_end(&mut sheet_bin).expect("read sheet bytes");

    let err = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 1_048_576,
            col: 0,
            new_value: CellValue::Number(2.0),
            new_style: None,
            clear_formula: false,
            new_formula: None,
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
        }],
    )
    .expect_err("expected InvalidInput for out-of-range row");

    match err {
        formula_xlsb::Error::Io(io_err) => {
            assert_eq!(io_err.kind(), std::io::ErrorKind::InvalidInput);
            let msg = io_err.to_string();
            assert!(msg.contains("row=1048576"), "unexpected message: {msg}");
            assert!(msg.contains("col=0"), "unexpected message: {msg}");
            assert!(msg.contains("max row=1048575"), "unexpected message: {msg}");
            assert!(msg.contains("max col=16383"), "unexpected message: {msg}");
        }
        other => panic!("expected InvalidInput, got {other:?}"),
    }
}

#[test]
fn patch_sheet_bin_rejects_out_of_range_cell_edit_col() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number(0, 0, 1.0);

    let xlsb_bytes = builder.build_bytes();
    let mut zip = zip::ZipArchive::new(Cursor::new(xlsb_bytes)).expect("open in-memory xlsb zip");
    let mut entry = zip
        .by_name("xl/worksheets/sheet1.bin")
        .expect("find sheet1.bin");
    let mut sheet_bin = Vec::with_capacity(entry.size() as usize);
    entry.read_to_end(&mut sheet_bin).expect("read sheet bytes");

    let err = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 0,
            col: 16_384,
            new_value: CellValue::Number(2.0),
            new_style: None,
            clear_formula: false,
            new_formula: None,
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
        }],
    )
    .expect_err("expected InvalidInput for out-of-range col");

    match err {
        formula_xlsb::Error::Io(io_err) => {
            assert_eq!(io_err.kind(), std::io::ErrorKind::InvalidInput);
            let msg = io_err.to_string();
            assert!(msg.contains("row=0"), "unexpected message: {msg}");
            assert!(msg.contains("col=16384"), "unexpected message: {msg}");
            assert!(msg.contains("max row=1048575"), "unexpected message: {msg}");
            assert!(msg.contains("max col=16383"), "unexpected message: {msg}");
        }
        other => panic!("expected InvalidInput, got {other:?}"),
    }
}

#[test]
fn patch_sheet_bin_allows_cell_edit_at_excel_grid_limit() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number(0, 0, 1.0);

    let xlsb_bytes = builder.build_bytes();
    let mut zip = zip::ZipArchive::new(Cursor::new(xlsb_bytes)).expect("open in-memory xlsb zip");
    let mut entry = zip
        .by_name("xl/worksheets/sheet1.bin")
        .expect("find sheet1.bin");
    let mut sheet_bin = Vec::with_capacity(entry.size() as usize);
    entry.read_to_end(&mut sheet_bin).expect("read sheet bytes");

    patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 1_048_575,
            col: 16_383,
            new_value: CellValue::Number(2.0),
            new_style: None,
            clear_formula: false,
            new_formula: None,
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
        }],
    )
    .expect("expected max row/col to be accepted");
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
            clear_formula: false,
            new_formula: None,
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
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
        &[
            CellEdit {
                row: 0,
                col: 0,
                new_value: CellValue::Text("World".to_string()),
                clear_formula: false,
                new_formula: None,
                new_rgcb: None,
                new_formula_flags: None,
                shared_string_index: None,
                new_style: None,
            },
            CellEdit {
                row: 0,
                col: 1,
                new_value: CellValue::Number(100.0),
                clear_formula: false,
                new_formula: None,
                new_rgcb: None,
                new_formula_flags: None,
                shared_string_index: None,
                new_style: None,
            },
        ],
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
    let formula = formula_cell
        .formula
        .as_ref()
        .expect("formula metadata preserved");
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
            clear_formula: false,
            new_formula: None,
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
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
    let formula = formula_cell
        .formula
        .as_ref()
        .expect("formula metadata preserved");
    assert_eq!(formula.text.as_deref(), Some("B1*2"));
    assert_eq!(formula.rgce.as_slice(), original_formula_rgce.as_slice());
}

#[test]
fn save_with_cell_edits_can_clear_formula_to_plain_value_cell() {
    let fixture = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/simple.xlsb");
    let wb = XlsbWorkbook::open(fixture).expect("open xlsb");

    // C1 is a numeric formula cell in the fixture (`B1*2`). Capture its current cached value so the
    // test remains robust if the fixture changes.
    let original_sheet = wb.read_sheet(0).expect("read original sheet");
    let original_c1 = original_sheet
        .cells
        .iter()
        .find(|c| (c.row, c.col) == (0, 2))
        .expect("C1 exists");
    assert!(
        original_c1.formula.is_some(),
        "expected C1 to be a formula cell"
    );
    let cached = match original_c1.value {
        CellValue::Number(v) => v,
        ref other => panic!("expected numeric cached value in C1, got {other:?}"),
    };

    let tmpdir = tempdir().expect("create temp dir");
    let out_path = tmpdir.path().join("paste-values.xlsb");

    wb.save_with_cell_edits(
        &out_path,
        0,
        &[CellEdit {
            row: 0,
            col: 2,
            new_value: CellValue::Number(cached),
            new_style: None,
            clear_formula: true,
            new_formula: None,
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
        }],
    )
    .expect("save_with_cell_edits");

    let patched = XlsbWorkbook::open(&out_path).expect("open patched workbook");
    let patched_sheet = patched.read_sheet(0).expect("read patched sheet");
    let patched_c1 = patched_sheet
        .cells
        .iter()
        .find(|c| (c.row, c.col) == (0, 2))
        .expect("patched C1 exists");

    assert_eq!(patched_c1.value, CellValue::Number(cached));
    assert!(
        patched_c1.formula.is_none(),
        "expected formula metadata to be removed when clear_formula=true"
    );
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
            clear_formula: false,
            new_formula: Some(new_rgce.clone()),
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
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
    let formula = formula_cell
        .formula
        .as_ref()
        .expect("formula metadata preserved");
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
            clear_formula: false,
            new_formula: Some(rgce),
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
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
    let formula = formula_cell
        .formula
        .as_ref()
        .expect("formula metadata preserved");
    assert_eq!(formula.text.as_deref(), Some("IF(B1>0,B1*3,0)"));
}

#[test]
fn patch_sheet_bin_can_update_udf_formula_using_workbook_context() {
    let fixture = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/udf.xlsb");
    let wb = XlsbWorkbook::open(fixture).expect("open xlsb");
    let sheet_part = wb.sheet_metas()[0].part_path.clone();

    let ctx = wb.workbook_context();
    let encoded = encode_rgce_with_context("=MyAddinFunc(1,2,3)", ctx, CellCoord::new(0, 3))
        .expect("encode UDF via NameX");
    assert!(
        encoded.rgcb.is_empty(),
        "UDF encoding should not require rgcb"
    );

    let sheet_bin = read_zip_part(fixture, &sheet_part);
    let patched_sheet_bin = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 0,
            col: 3,
            new_value: CellValue::Number(0.0),
            clear_formula: false,
            new_formula: Some(encoded.rgce),
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
        }],
    )
    .expect("patch sheet bin");

    let tmpdir = tempdir().expect("create temp dir");
    let out_path = tmpdir.path().join("patched-udf.xlsb");

    wb.save_with_part_overrides(
        &out_path,
        &HashMap::from([(sheet_part.clone(), patched_sheet_bin)]),
    )
    .expect("write patched workbook");

    let patched = XlsbWorkbook::open(&out_path).expect("open patched workbook");
    let patched_sheet = patched.read_sheet(0).expect("read patched sheet");

    let udf_cell = patched_sheet
        .cells
        .iter()
        .find(|c| (c.row, c.col) == (0, 3))
        .expect("D1 cell");
    assert_eq!(udf_cell.value, CellValue::Number(0.0));
    let formula = udf_cell
        .formula
        .as_ref()
        .expect("formula metadata preserved");
    assert_eq!(formula.text.as_deref(), Some("MyAddinFunc(1,2,3)"));
}

#[test]
fn patch_sheet_bin_preserves_formula_trailing_bytes() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_sheet_name("Arrayy");

    // `PtgArray` placeholder token, followed by extra bytes (rgcb) describing the array constant.
    let rgce = rgce::array_placeholder();
    let extra = vec![0xDE, 0xAD, 0xBE, 0xEF];
    builder.set_cell_formula_num(0, 0, 1.0, rgce, extra);

    let xlsb_bytes = builder.build_bytes();
    let mut zip = zip::ZipArchive::new(Cursor::new(xlsb_bytes)).expect("open in-memory xlsb zip");
    let mut entry = zip
        .by_name("xl/worksheets/sheet1.bin")
        .expect("find sheet1.bin");
    // Do not trust `ZipFile::size()` for allocation; ZIP metadata is untrusted and can
    // advertise enormous uncompressed sizes (zip-bomb style OOM).
    let mut sheet_bin = Vec::new();
    entry.read_to_end(&mut sheet_bin).expect("read sheet bytes");

    // No-op patch (cached value unchanged): must reserialize byte-for-byte, including the extra
    // bytes after `rgce`.
    let patched = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Number(1.0),
            clear_formula: false,
            new_formula: None,
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
        }],
    )
    .expect("patch sheet bin");

    assert_eq!(patched, sheet_bin);
}

#[test]
fn patch_sheet_bin_can_update_formula_rgcb_bytes() {
    let ctx = formula_xlsb::workbook_context::WorkbookContext::default();

    let encoded_123 =
        encode_rgce_with_context("=SUM({1,2,3})", &ctx, CellCoord::new(0, 0)).expect("encode rgce");
    assert!(
        !encoded_123.rgcb.is_empty(),
        "expected array formula encoding to produce rgcb bytes"
    );

    let mut builder = XlsbFixtureBuilder::new();
    builder.set_sheet_name("ArrayRgcb");
    builder.set_cell_formula_num(
        0,
        0,
        6.0,
        encoded_123.rgce.clone(),
        encoded_123.rgcb.clone(),
    );

    let xlsb_bytes = builder.build_bytes();
    let mut zip = zip::ZipArchive::new(Cursor::new(xlsb_bytes)).expect("open in-memory xlsb zip");
    let mut entry = zip
        .by_name("xl/worksheets/sheet1.bin")
        .expect("find sheet1.bin");
    // Do not trust `ZipFile::size()` for allocation; ZIP metadata is untrusted and can
    // advertise enormous uncompressed sizes (zip-bomb style OOM).
    let mut sheet_bin = Vec::new();
    entry.read_to_end(&mut sheet_bin).expect("read sheet bytes");

    let encoded_45 =
        encode_rgce_with_context("=SUM({4,5})", &ctx, CellCoord::new(0, 0)).expect("encode rgce");
    assert_eq!(
        encoded_45.rgce, encoded_123.rgce,
        "expected SUM(array) formulas to share the same rgce stream so only rgcb changes"
    );

    let patched = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Number(9.0),
            new_style: None,
            clear_formula: false,
            new_formula: None,
            new_rgcb: Some(encoded_45.rgcb.clone()),
            new_formula_flags: None,
            shared_string_index: None,
        }],
    )
    .expect("patch sheet bin");

    let parsed = formula_xlsb::parse_sheet_bin_with_context(&mut Cursor::new(&patched), &[], &ctx)
        .expect("parse patched sheet");
    let cell = parsed
        .cells
        .iter()
        .find(|c| (c.row, c.col) == (0, 0))
        .expect("A1 exists");
    assert_eq!(cell.value, CellValue::Number(9.0));

    let formula = cell.formula.as_ref().expect("formula metadata");
    assert_eq!(formula.extra, encoded_45.rgcb);
    assert_eq!(formula.text.as_deref(), Some("SUM({4,5})"));
}

#[test]
fn cell_edit_with_formula_text_with_context_can_patch_rgcb_formulas() {
    let ctx = formula_xlsb::workbook_context::WorkbookContext::default();

    let encoded_123 =
        encode_rgce_with_context("=SUM({1,2,3})", &ctx, CellCoord::new(0, 0)).expect("encode rgce");
    assert!(
        !encoded_123.rgcb.is_empty(),
        "expected array formula encoding to produce rgcb bytes"
    );

    let mut builder = XlsbFixtureBuilder::new();
    builder.set_sheet_name("ArrayRgcb");
    builder.set_cell_formula_num(
        0,
        0,
        6.0,
        encoded_123.rgce.clone(),
        encoded_123.rgcb.clone(),
    );

    let xlsb_bytes = builder.build_bytes();
    let mut zip = zip::ZipArchive::new(Cursor::new(xlsb_bytes)).expect("open in-memory xlsb zip");
    let mut entry = zip
        .by_name("xl/worksheets/sheet1.bin")
        .expect("find sheet1.bin");
    // Do not trust `ZipFile::size()` for allocation; ZIP metadata is untrusted and can
    // advertise enormous uncompressed sizes (zip-bomb style OOM).
    let mut sheet_bin = Vec::new();
    entry.read_to_end(&mut sheet_bin).expect("read sheet bytes");

    let edit =
        CellEdit::with_formula_text_with_context(0, 0, CellValue::Number(9.0), "=SUM({4,5})", &ctx)
            .expect("encode formula with context");
    let patched = patch_sheet_bin(&sheet_bin, &[edit]).expect("patch sheet bin");

    let parsed = formula_xlsb::parse_sheet_bin_with_context(&mut Cursor::new(&patched), &[], &ctx)
        .expect("parse patched sheet");
    let cell = parsed
        .cells
        .iter()
        .find(|c| (c.row, c.col) == (0, 0))
        .expect("A1 exists");
    assert_eq!(cell.value, CellValue::Number(9.0));

    let formula = cell.formula.as_ref().expect("formula metadata");
    assert_eq!(formula.text.as_deref(), Some("SUM({4,5})"));

    let encoded_45 =
        encode_rgce_with_context("=SUM({4,5})", &ctx, CellCoord::new(0, 0)).expect("encode rgce");
    assert_eq!(formula.rgce, encoded_45.rgce);
    assert_eq!(formula.extra, encoded_45.rgcb);
}

#[test]
fn patch_sheet_bin_preserves_formula_rgcb_when_updating_cached_value_only() {
    let ctx = formula_xlsb::workbook_context::WorkbookContext::default();

    let encoded =
        encode_rgce_with_context("=SUM({1,2,3})", &ctx, CellCoord::new(0, 0)).expect("encode rgce");
    assert!(
        !encoded.rgcb.is_empty(),
        "expected array formula encoding to produce rgcb bytes"
    );

    let mut builder = XlsbFixtureBuilder::new();
    builder.set_sheet_name("ArrayRgcb");
    builder.set_cell_formula_num(0, 0, 6.0, encoded.rgce.clone(), encoded.rgcb.clone());

    let xlsb_bytes = builder.build_bytes();
    let mut zip = zip::ZipArchive::new(Cursor::new(xlsb_bytes)).expect("open in-memory xlsb zip");
    let mut entry = zip
        .by_name("xl/worksheets/sheet1.bin")
        .expect("find sheet1.bin");
    // Do not trust `ZipFile::size()` for allocation; ZIP metadata is untrusted and can
    // advertise enormous uncompressed sizes (zip-bomb style OOM).
    let mut sheet_bin = Vec::new();
    entry.read_to_end(&mut sheet_bin).expect("read sheet bytes");

    let patched = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Number(10.0),
            clear_formula: false,
            new_formula: None,
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
        }],
    )
    .expect("patch sheet bin");

    let parsed = formula_xlsb::parse_sheet_bin_with_context(&mut Cursor::new(&patched), &[], &ctx)
        .expect("parse patched sheet");
    let cell = parsed
        .cells
        .iter()
        .find(|c| (c.row, c.col) == (0, 0))
        .expect("A1 exists");
    assert_eq!(cell.value, CellValue::Number(10.0));

    let formula = cell.formula.as_ref().expect("formula metadata");
    assert_eq!(formula.rgce, encoded.rgce);
    assert_eq!(formula.extra, encoded.rgcb);
    assert_eq!(formula.text.as_deref(), Some("SUM({1,2,3})"));
}

#[test]
fn patch_sheet_bin_requires_new_rgcb_when_replacing_rgce_for_formula_with_existing_rgcb() {
    let ctx = formula_xlsb::workbook_context::WorkbookContext::default();

    let encoded_sum =
        encode_rgce_with_context("=SUM({1,2,3})", &ctx, CellCoord::new(0, 0)).expect("encode rgce");
    assert!(
        !encoded_sum.rgcb.is_empty(),
        "expected rgcb for array formula"
    );

    let mut builder = XlsbFixtureBuilder::new();
    builder.set_sheet_name("ArrayRgcb");
    builder.set_cell_formula_num(
        0,
        0,
        6.0,
        encoded_sum.rgce.clone(),
        encoded_sum.rgcb.clone(),
    );

    let xlsb_bytes = builder.build_bytes();
    let mut zip = zip::ZipArchive::new(Cursor::new(xlsb_bytes)).expect("open in-memory xlsb zip");
    let mut entry = zip
        .by_name("xl/worksheets/sheet1.bin")
        .expect("find sheet1.bin");
    // Do not trust `ZipFile::size()` for allocation; ZIP metadata is untrusted and can
    // advertise enormous uncompressed sizes (zip-bomb style OOM).
    let mut sheet_bin = Vec::new();
    entry.read_to_end(&mut sheet_bin).expect("read sheet bytes");

    let encoded_max =
        encode_rgce_with_context("=MAX({1,2,3})", &ctx, CellCoord::new(0, 0)).expect("encode rgce");
    assert_ne!(
        encoded_max.rgce, encoded_sum.rgce,
        "expected MAX(array) to change rgce"
    );

    let err = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Number(3.0),
            new_style: None,
            clear_formula: false,
            new_formula: Some(encoded_max.rgce.clone()),
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
        }],
    )
    .expect_err("expected InvalidInput when changing rgce without supplying new_rgcb");
    match err {
        formula_xlsb::Error::Io(io_err) => {
            assert_eq!(io_err.kind(), std::io::ErrorKind::InvalidInput);
        }
        other => panic!("expected InvalidInput, got {other:?}"),
    }

    let patched = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Number(3.0),
            new_style: None,
            clear_formula: false,
            new_formula: Some(encoded_max.rgce.clone()),
            new_rgcb: Some(encoded_sum.rgcb.clone()),
            new_formula_flags: None,
            shared_string_index: None,
        }],
    )
    .expect("patch sheet bin with explicit rgcb");

    let parsed = formula_xlsb::parse_sheet_bin_with_context(&mut Cursor::new(&patched), &[], &ctx)
        .expect("parse patched sheet");
    let cell = parsed
        .cells
        .iter()
        .find(|c| (c.row, c.col) == (0, 0))
        .expect("A1 exists");
    assert_eq!(cell.value, CellValue::Number(3.0));

    let formula = cell.formula.as_ref().expect("formula metadata");
    assert_eq!(formula.extra, encoded_sum.rgcb);
    assert_eq!(formula.text.as_deref(), Some("MAX({1,2,3})"));
}

#[test]
fn patch_sheet_bin_can_clear_rgcb_when_replacing_rgce_for_formula_with_existing_rgcb() {
    let ctx = formula_xlsb::workbook_context::WorkbookContext::default();

    let encoded_sum =
        encode_rgce_with_context("=SUM({1,2,3})", &ctx, CellCoord::new(0, 0)).expect("encode rgce");
    assert!(
        !encoded_sum.rgcb.is_empty(),
        "expected rgcb for array formula"
    );

    let mut builder = XlsbFixtureBuilder::new();
    builder.set_sheet_name("ArrayRgcb");
    builder.set_cell_formula_num(
        0,
        0,
        6.0,
        encoded_sum.rgce.clone(),
        encoded_sum.rgcb.clone(),
    );

    let xlsb_bytes = builder.build_bytes();
    let mut zip = zip::ZipArchive::new(Cursor::new(xlsb_bytes)).expect("open in-memory xlsb zip");
    let mut entry = zip
        .by_name("xl/worksheets/sheet1.bin")
        .expect("find sheet1.bin");
    // Do not trust `ZipFile::size()` for allocation; ZIP metadata is untrusted and can
    // advertise enormous uncompressed sizes (zip-bomb style OOM).
    let mut sheet_bin = Vec::new();
    entry.read_to_end(&mut sheet_bin).expect("read sheet bytes");

    // Update the formula to a version that does not require any trailing rgcb bytes and explicitly
    // clear the original rgcb payload by providing `Some(empty)`.
    let encoded_no_rgcb =
        encode_rgce_with_context("=SUM(1,2,3)", &ctx, CellCoord::new(0, 0)).expect("encode rgce");
    assert!(
        encoded_no_rgcb.rgcb.is_empty(),
        "expected SUM(1,2,3) encoding to not require rgcb bytes"
    );

    let patched = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Number(6.0),
            clear_formula: false,
            new_formula: Some(encoded_no_rgcb.rgce.clone()),
            new_rgcb: Some(encoded_no_rgcb.rgcb.clone()),
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
        }],
    )
    .expect("patch sheet bin");

    let parsed = formula_xlsb::parse_sheet_bin_with_context(&mut Cursor::new(&patched), &[], &ctx)
        .expect("parse patched sheet");
    let cell = parsed
        .cells
        .iter()
        .find(|c| (c.row, c.col) == (0, 0))
        .expect("A1 exists");
    assert_eq!(cell.value, CellValue::Number(6.0));

    let formula = cell.formula.as_ref().expect("formula metadata");
    assert_eq!(formula.rgce, encoded_no_rgcb.rgce);
    assert!(
        formula.extra.is_empty(),
        "expected rgcb bytes to be cleared"
    );
    assert_eq!(formula.text.as_deref(), Some("SUM(1,2,3)"));
}

#[test]
fn patch_sheet_bin_errors_when_replacing_rgce_for_formula_with_existing_rgcb_without_new_rgcb() {
    let ctx = formula_xlsb::workbook_context::WorkbookContext::default();

    let encoded_sum =
        encode_rgce_with_context("=SUM({1,2,3})", &ctx, CellCoord::new(0, 0)).expect("encode rgce");
    assert!(
        !encoded_sum.rgcb.is_empty(),
        "expected rgcb for array formula"
    );

    let mut builder = XlsbFixtureBuilder::new();
    builder.set_sheet_name("ArrayRgcb");
    builder.set_cell_formula_num(
        0,
        0,
        6.0,
        encoded_sum.rgce.clone(),
        encoded_sum.rgcb.clone(),
    );

    let xlsb_bytes = builder.build_bytes();
    let mut zip = zip::ZipArchive::new(Cursor::new(xlsb_bytes)).expect("open in-memory xlsb zip");
    let mut entry = zip
        .by_name("xl/worksheets/sheet1.bin")
        .expect("find sheet1.bin");
    // Do not trust `ZipFile::size()` for allocation; ZIP metadata is untrusted and can
    // advertise enormous uncompressed sizes (zip-bomb style OOM).
    let mut sheet_bin = Vec::new();
    entry.read_to_end(&mut sheet_bin).expect("read sheet bytes");

    let encoded_no_rgcb =
        encode_rgce_with_context("=SUM(1,2,3)", &ctx, CellCoord::new(0, 0)).expect("encode rgce");
    assert!(
        encoded_no_rgcb.rgcb.is_empty(),
        "expected SUM(1,2,3) encoding to not require rgcb bytes"
    );

    let err = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Number(6.0),
            clear_formula: false,
            new_formula: Some(encoded_no_rgcb.rgce.clone()),
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
            new_style: None,
        }],
    )
    .expect_err("expected InvalidInput when replacing rgce without new_rgcb");

    match err {
        formula_xlsb::Error::Io(io_err) => {
            assert_eq!(io_err.kind(), std::io::ErrorKind::InvalidInput);
            assert!(
                io_err.to_string().contains("provide CellEdit.new_rgcb"),
                "expected error to instruct caller to provide CellEdit.new_rgcb, got: {io_err}"
            );
        }
        other => panic!("expected InvalidInput, got {other:?}"),
    }
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
    // Do not trust `ZipFile::size()` for allocation; ZIP metadata is untrusted and can
    // advertise enormous uncompressed sizes (zip-bomb style OOM).
    let mut sheet_bin = Vec::new();
    entry.read_to_end(&mut sheet_bin).expect("read sheet bytes");

    let patched = patch_sheet_bin(
        &sheet_bin,
        &[
            CellEdit {
                row: 0,
                col: 0,
                new_value: CellValue::Bool(true),
                clear_formula: false,
                new_formula: None,
                new_rgcb: None,
                new_formula_flags: None,
                shared_string_index: None,
                new_style: None,
            },
            CellEdit {
                row: 0,
                col: 1,
                new_value: CellValue::Error(0x07),
                clear_formula: false,
                new_formula: None,
                new_rgcb: None,
                new_formula_flags: None,
                shared_string_index: None,
                new_style: None,
            },
            CellEdit {
                row: 0,
                col: 2,
                new_value: CellValue::Blank,
                clear_formula: false,
                new_formula: None,
                new_rgcb: None,
                new_formula_flags: None,
                shared_string_index: None,
                new_style: None,
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
                clear_formula: false,
                new_formula: None,
                new_rgcb: None,
                new_formula_flags: None,
                shared_string_index: None,
                new_style: None,
            },
            CellEdit {
                row: 0,
                col: 1,
                new_value: CellValue::Error(0x2A),
                clear_formula: false,
                new_formula: None,
                new_rgcb: None,
                new_formula_flags: None,
                shared_string_index: None,
                new_style: None,
            },
            // Patch the blank cell into a numeric value.
            CellEdit {
                row: 0,
                col: 2,
                new_value: CellValue::Number(123.0),
                clear_formula: false,
                new_formula: None,
                new_rgcb: None,
                new_formula_flags: None,
                shared_string_index: None,
                new_style: None,
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

    assert_eq!(
        cells.remove(&(0, 0)).expect("A1 exists"),
        CellValue::Bool(false)
    );
    assert_eq!(
        cells.remove(&(0, 1)).expect("B1 exists"),
        CellValue::Error(0x2A)
    );
    assert_eq!(
        cells.remove(&(0, 2)).expect("C1 exists"),
        CellValue::Number(123.0)
    );
}

#[test]
fn save_with_cell_edits_can_patch_formula_bool_string_error_cells() {
    fn ptg_str(s: &str) -> Vec<u8> {
        let mut out = vec![0x17]; // PtgStr
        let units: Vec<u16> = s.encode_utf16().collect();
        out.extend_from_slice(&(units.len() as u16).to_le_bytes());
        for u in units {
            out.extend_from_slice(&u.to_le_bytes());
        }
        out
    }

    let mut builder = XlsbFixtureBuilder::new();
    builder.set_sheet_name("FormulaTypes");
    builder.set_cell_formula_bool(0, 0, true, vec![0x1D, 0x01]); // TRUE
    builder.set_cell_formula_str(0, 1, "Hello", ptg_str("Hello"));
    builder.set_cell_formula_err(0, 2, 0x07, vec![0x1C, 0x07]); // #DIV/0!

    let bytes = builder.build_bytes();
    let tmpdir = tempdir().expect("create temp dir");
    let input_path = tmpdir.path().join("formula_types_input.xlsb");
    let output_path = tmpdir.path().join("formula_types_output.xlsb");
    std::fs::write(&input_path, bytes).expect("write input workbook");

    let wb = XlsbWorkbook::open(&input_path).expect("open generated xlsb");

    let edits = [
        CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Bool(false),
            new_style: None,
            clear_formula: false,
            new_formula: Some(vec![0x1D, 0x00]), // FALSE
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
        },
        CellEdit {
            row: 0,
            col: 1,
            new_value: CellValue::Text("World".to_string()),
            new_style: None,
            clear_formula: false,
            new_formula: Some(ptg_str("World")),
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
        },
        CellEdit {
            row: 0,
            col: 2,
            new_value: CellValue::Error(0x2A), // #N/A
            new_style: None,
            clear_formula: false,
            new_formula: Some(vec![0x1C, 0x2A]),
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
        },
    ];

    wb.save_with_cell_edits(&output_path, 0, &edits)
        .expect("save_with_cell_edits");

    let wb2 = XlsbWorkbook::open(&output_path).expect("open patched workbook");
    let sheet = wb2.read_sheet(0).expect("read sheet");
    let mut cells = sheet
        .cells
        .iter()
        .map(|c| ((c.row, c.col), c))
        .collect::<HashMap<_, _>>();

    let a1 = cells.remove(&(0, 0)).expect("A1 exists");
    assert_eq!(a1.value, CellValue::Bool(false));
    let a1_formula = a1.formula.as_ref().expect("A1 formula");
    assert_eq!(a1_formula.rgce, vec![0x1D, 0x00]);

    let b1 = cells.remove(&(0, 1)).expect("B1 exists");
    assert_eq!(b1.value, CellValue::Text("World".to_string()));
    let b1_formula = b1.formula.as_ref().expect("B1 formula");
    assert_eq!(b1_formula.rgce, ptg_str("World"));

    let c1 = cells.remove(&(0, 2)).expect("C1 exists");
    assert_eq!(c1.value, CellValue::Error(0x2A));
    let c1_formula = c1.formula.as_ref().expect("C1 formula");
    assert_eq!(c1_formula.rgce, vec![0x1C, 0x2A]);
}

fn decode_rk_number(raw: u32) -> f64 {
    let raw_i = raw as i32;
    let mut v = if raw_i & 0x02 != 0 {
        (raw_i >> 2) as f64
    } else {
        let shifted = raw & 0xFFFFFFFC;
        f64::from_bits((shifted as u64) << 32)
    };
    if raw_i & 0x01 != 0 {
        v /= 100.0;
    }
    v
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

#[test]
fn patch_sheet_bin_is_byte_identical_for_noop_rk_float_edit() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_sheet_name("RKFloat");
    builder.set_cell_number_rk(0, 1, 0.125);

    let xlsb_bytes = builder.build_bytes();
    let mut zip = zip::ZipArchive::new(Cursor::new(xlsb_bytes)).expect("open in-memory xlsb zip");
    let mut entry = zip
        .by_name("xl/worksheets/sheet1.bin")
        .expect("find sheet1.bin");
    // Do not trust `ZipFile::size()` for allocation; ZIP metadata is untrusted and can
    // advertise enormous uncompressed sizes (zip-bomb style OOM).
    let mut sheet_bin = Vec::new();
    entry.read_to_end(&mut sheet_bin).expect("read sheet bytes");

    let patched = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 0,
            col: 1,
            new_value: CellValue::Number(0.125),
            new_style: None,
            clear_formula: false,
            new_formula: None,
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
        }],
    )
    .expect("patch sheet bin");

    assert_eq!(patched, sheet_bin);
}

#[test]
fn patch_sheet_bin_keeps_rk_record_for_float_rk_values() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_sheet_name("RKFloat");
    builder.set_cell_number_rk(0, 1, 0.0);

    let xlsb_bytes = builder.build_bytes();
    let mut zip = zip::ZipArchive::new(Cursor::new(xlsb_bytes)).expect("open in-memory xlsb zip");
    let mut entry = zip
        .by_name("xl/worksheets/sheet1.bin")
        .expect("find sheet1.bin");
    // Do not trust `ZipFile::size()` for allocation; ZIP metadata is untrusted and can
    // advertise enormous uncompressed sizes (zip-bomb style OOM).
    let mut sheet_bin = Vec::new();
    entry.read_to_end(&mut sheet_bin).expect("read sheet bytes");

    let patched = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 0,
            col: 1,
            new_value: CellValue::Number(0.125),
            new_style: None,
            clear_formula: false,
            new_formula: None,
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
        }],
    )
    .expect("patch sheet bin");

    let (id, payload) = find_cell_record(&patched, 0, 1).expect("patched cell record");
    assert_eq!(id, 0x0002, "expected RK NUM record");
    let rk = u32::from_le_bytes(payload[8..12].try_into().unwrap());
    assert_eq!(decode_rk_number(rk), 0.125);
}
