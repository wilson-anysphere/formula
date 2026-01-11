use formula_xlsb::{CellValue, XlsbWorkbook};
use pretty_assertions::assert_eq;
use std::io::Write;
use tempfile::NamedTempFile;

mod fixture_builder;

use fixture_builder::{rgce, XlsbFixtureBuilder};

fn write_temp_xlsb(bytes: &[u8]) -> NamedTempFile {
    let mut file = tempfile::Builder::new()
        .prefix("formula_xlsb_fixture_")
        .suffix(".xlsb")
        .tempfile()
        .expect("create temp xlsb");
    file.write_all(bytes).expect("write temp xlsb");
    file.flush().expect("flush temp xlsb");
    file
}

#[test]
fn opens_fixture_and_reads_cells() {
    let path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/simple.xlsb");
    let wb = XlsbWorkbook::open(path).expect("open xlsb");

    assert_eq!(wb.sheet_metas().len(), 1);
    assert_eq!(wb.sheet_metas()[0].name, "Sheet1");

    let sheet = wb.read_sheet(0).expect("read sheet1");

    let mut cells = sheet
        .cells
        .iter()
        .map(|c| ((c.row, c.col), c))
        .collect::<std::collections::HashMap<_, _>>();

    assert_eq!(
        cells.remove(&(0, 0)).unwrap().value,
        CellValue::Text("Hello".to_string())
    );
    assert_eq!(
        cells.remove(&(0, 1)).unwrap().value,
        CellValue::Number(42.5)
    );

    let formula_cell = cells.remove(&(0, 2)).unwrap();
    assert_eq!(formula_cell.value, CellValue::Number(85.0));
    let formula = formula_cell
        .formula
        .as_ref()
        .expect("formula payload preserved");
    assert_eq!(formula.text.as_deref(), Some("B1*2"));
    assert_eq!(formula.rgce.len(), 11);
    assert!(formula.extra.is_empty());
}

#[test]
fn generated_fixture_reads_number_and_formula_cells() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_sheet_name("Sheet1");
    builder.set_cell_number(0, 1, 42.5);

    let mut rgce_bytes = Vec::new();
    rgce::push_ref(&mut rgce_bytes, 0, 1, false, false); // B1
    rgce::push_int(&mut rgce_bytes, 2);
    rgce::push_mul(&mut rgce_bytes);

    builder.set_cell_formula_num(
        0,
        2,
        85.0,
        rgce_bytes.clone(),
        Vec::new(),
    );

    let bytes = builder.build_bytes();
    let tmp = write_temp_xlsb(&bytes);
    let wb = XlsbWorkbook::open(tmp.path()).expect("open generated xlsb");

    assert_eq!(wb.sheet_metas().len(), 1);
    assert_eq!(wb.sheet_metas()[0].name, "Sheet1");

    let sheet = wb.read_sheet(0).expect("read sheet1");
    let mut cells = sheet
        .cells
        .iter()
        .map(|c| ((c.row, c.col), c))
        .collect::<std::collections::HashMap<_, _>>();

    assert_eq!(
        cells.remove(&(0, 1)).unwrap().value,
        CellValue::Number(42.5)
    );

    let formula_cell = cells.remove(&(0, 2)).unwrap();
    assert_eq!(formula_cell.value, CellValue::Number(85.0));
    let formula = formula_cell.formula.as_ref().expect("formula payload preserved");
    assert_eq!(formula.text.as_deref(), Some("B1*2"));
    assert_eq!(
        formula.rgce,
        vec![0x24, 0, 0, 0, 0, 0x01, 0xC0, 0x1E, 0x02, 0x00, 0x05]
    );
    assert!(formula.extra.is_empty());

}

#[test]
fn generated_fixture_supports_shared_strings_and_absolute_refs() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_sheet_name("Data");

    let sst_hello = builder.add_shared_string("Hello");
    builder.set_cell_sst(1, 0, sst_hello);

    builder.set_cell_number(0, 0, 1.0);
    builder.set_cell_number(0, 1, 2.0);

    let mut rgce_bytes = Vec::new();
    rgce::push_ref(&mut rgce_bytes, 0, 0, true, true); // $A$1
    rgce::push_ref(&mut rgce_bytes, 0, 1, false, true); // $B1
    rgce::push_add(&mut rgce_bytes);

    builder.set_cell_formula_num(
        0,
        2,
        3.0,
        rgce_bytes,
        Vec::new(),
    );

    let bytes = builder.build_bytes();
    let tmp = write_temp_xlsb(&bytes);
    let wb = XlsbWorkbook::open(tmp.path()).expect("open generated xlsb");

    assert_eq!(wb.sheet_metas().len(), 1);
    assert_eq!(wb.sheet_metas()[0].name, "Data");
    assert_eq!(wb.shared_strings(), &["Hello".to_string()]);

    let sheet = wb.read_sheet(0).expect("read sheet1");
    let mut cells = sheet
        .cells
        .iter()
        .map(|c| ((c.row, c.col), c))
        .collect::<std::collections::HashMap<_, _>>();

    assert_eq!(
        cells.remove(&(1, 0)).unwrap().value,
        CellValue::Text("Hello".to_string())
    );

    let formula_cell = cells.remove(&(0, 2)).unwrap();
    assert_eq!(formula_cell.value, CellValue::Number(3.0));
    let formula = formula_cell.formula.as_ref().expect("formula payload preserved");
    assert_eq!(formula.text.as_deref(), Some("$A$1+$B1"));
    assert!(formula.extra.is_empty());

}

#[test]
fn generated_fixture_preserves_formula_extra_bytes_for_array_constants() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_sheet_name("Arrayy");

    // `PtgArray` (0x20) is a placeholder token that typically requires extra bytes
    // after `rgce` (the `rgcb` stream) to describe the constant array. Our decoder
    // requires a valid payload, but the reader should still preserve `rgce` and
    // the unknown bytes for round-tripping.
    let rgce = rgce::array_placeholder();
    let extra = vec![0xDE, 0xAD, 0xBE, 0xEF];

    builder.set_cell_formula_num(0, 0, 1.0, rgce.clone(), extra.clone());

    let bytes = builder.build_bytes();
    let tmp = write_temp_xlsb(&bytes);
    let wb = XlsbWorkbook::open(tmp.path()).expect("open generated xlsb");
    let sheet = wb.read_sheet(0).expect("read sheet1");

    let formula_cell = sheet
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 0)
        .expect("formula cell");
    assert_eq!(formula_cell.value, CellValue::Number(1.0));
    let formula = formula_cell.formula.as_ref().expect("formula payload preserved");
    assert_eq!(formula.text.as_deref(), None);
    assert_eq!(formula.rgce, rgce);
    assert_eq!(formula.extra, extra);

}

#[test]
fn generated_fixture_supports_inline_strings_without_shared_strings_part() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_sheet_name("Inline");
    builder.set_cell_inline_string(0, 0, "Hello");

    let bytes = builder.build_bytes();
    let tmp = write_temp_xlsb(&bytes);
    let wb = XlsbWorkbook::open(tmp.path()).expect("open generated xlsb");

    assert_eq!(wb.sheet_metas().len(), 1);
    assert_eq!(wb.sheet_metas()[0].name, "Inline");
    assert!(wb.shared_strings().is_empty());

    let sheet = wb.read_sheet(0).expect("read sheet1");
    let cell = sheet
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 0)
        .expect("cell A1");
    assert_eq!(cell.value, CellValue::Text("Hello".to_string()));
}

#[test]
fn decodes_addin_udf_calls_via_namex() {
    let path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/udf.xlsb");
    let wb = XlsbWorkbook::open(path).expect("open xlsb");

    let sheet = wb.read_sheet(0).expect("read sheet1");
    let mut cells = sheet
        .cells
        .iter()
        .map(|c| ((c.row, c.col), c))
        .collect::<std::collections::HashMap<_, _>>();

    let udf_cell = cells.remove(&(0, 3)).expect("D1 cell");
    assert_eq!(udf_cell.value, CellValue::Number(0.0));
    let formula = udf_cell.formula.as_ref().expect("formula payload preserved");
    assert_eq!(formula.text.as_deref(), Some("MyAddinFunc(1,2)"));
}
