use formula_xlsb::{CellValue, XlsbWorkbook};
use pretty_assertions::assert_eq;

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
    assert_eq!(formula.text.as_deref(), Some("A1+B1"));
    assert_eq!(formula.rgce.len(), 15);
}
