use formula_xlsb::{CellValue, XlsbWorkbook};
use pretty_assertions::assert_eq;

mod fixture_builder;

use fixture_builder::XlsbFixtureBuilder;
use std::path::PathBuf;
use std::sync::atomic::{AtomicUsize, Ordering};

fn write_temp_xlsb(bytes: &[u8]) -> PathBuf {
    static COUNTER: AtomicUsize = AtomicUsize::new(0);

    let n = COUNTER.fetch_add(1, Ordering::Relaxed);
    let pid = std::process::id();
    let ts = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .expect("clock")
        .as_nanos();

    let mut path = std::env::temp_dir();
    path.push(format!("formula_xlsb_fixture_{pid}_{ts}_{n}.xlsb"));
    std::fs::write(&path, bytes).expect("write temp xlsb");
    path
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
    let formula = formula_cell.formula.as_ref().expect("formula payload preserved");
    assert_eq!(formula.text.as_deref(), Some("B1*2"));
    assert_eq!(formula.rgce.len(), 11);
}

#[test]
fn generated_fixture_reads_number_and_formula_cells() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_sheet_name("Sheet1");
    builder.set_cell_number(0, 1, 42.5);
    builder.set_cell_formula_num(
        0,
        2,
        85.0,
        // Same token stream as `tests/fixtures/simple.xlsb` (`B1*2`).
        vec![0x24, 0, 0, 0, 0, 0x01, 0xC0, 0x1E, 0x02, 0x00, 0x05],
        Vec::new(),
    );

    let bytes = builder.build_bytes();
    let path = write_temp_xlsb(&bytes);
    let wb = XlsbWorkbook::open(&path).expect("open generated xlsb");

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

    let _ = std::fs::remove_file(&path);
}

#[test]
fn generated_fixture_supports_shared_strings_and_absolute_refs() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_sheet_name("Data");

    let sst_hello = builder.add_shared_string("Hello");
    builder.set_cell_sst(1, 0, sst_hello);

    builder.set_cell_number(0, 0, 1.0);
    builder.set_cell_number(0, 1, 2.0);
    builder.set_cell_formula_num(
        0,
        2,
        3.0,
        // `$A$1+$B1`
        vec![
            0x24, 0, 0, 0, 0, 0x00, 0x00, // $A$1
            0x24, 0, 0, 0, 0, 0x01, 0x40, // $B1
            0x03, // +
        ],
        Vec::new(),
    );

    let bytes = builder.build_bytes();
    let path = write_temp_xlsb(&bytes);
    let wb = XlsbWorkbook::open(&path).expect("open generated xlsb");

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

    let _ = std::fs::remove_file(&path);
}

#[test]
fn generated_fixture_preserves_formula_extra_bytes_for_array_constants() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_sheet_name("Arrayy");

    // `PtgArray` (0x20) is a placeholder token that typically requires extra bytes
    // after `rgce` (the `rgcb` stream) to describe the constant array. Our decoder
    // doesn't understand it yet, but the reader should still preserve `rgce` and
    // successfully load the sheet.
    let rgce = vec![0x20, 0, 0, 0, 0, 0, 0, 0];
    let extra = vec![0xDE, 0xAD, 0xBE, 0xEF];

    builder.set_cell_formula_num(0, 0, 1.0, rgce.clone(), extra);

    let bytes = builder.build_bytes();
    let path = write_temp_xlsb(&bytes);
    let wb = XlsbWorkbook::open(&path).expect("open generated xlsb");
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

    let _ = std::fs::remove_file(&path);
}
