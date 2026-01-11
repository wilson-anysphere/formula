use formula_xlsb::XlsbWorkbook;
use pretty_assertions::assert_eq;

#[test]
fn roundtrips_save_as_losslessly() {
    let path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/simple.xlsb");
    let wb = XlsbWorkbook::open(path).expect("open xlsb fixture");

    let tmpdir = tempfile::tempdir().expect("create temp dir");
    let saved_path = tmpdir.path().join("roundtrip.xlsb");
    wb.save_as(&saved_path).expect("save_as");

    let wb2 = XlsbWorkbook::open(&saved_path).expect("re-open saved xlsb");

    assert_eq!(wb.sheet_metas(), wb2.sheet_metas());

    let sheet1 = wb.read_sheet(0).expect("read original sheet");
    let sheet2 = wb2.read_sheet(0).expect("read round-tripped sheet");

    assert_eq!(sheet1.dimension, sheet2.dimension);
    assert_eq!(sheet1.cells.len(), sheet2.cells.len());

    let mut cells1 = sheet1
        .cells
        .iter()
        .map(|c| ((c.row, c.col), c))
        .collect::<std::collections::HashMap<_, _>>();
    let mut cells2 = sheet2
        .cells
        .iter()
        .map(|c| ((c.row, c.col), c))
        .collect::<std::collections::HashMap<_, _>>();

    for (pos, cell1) in cells1.drain() {
        let cell2 = cells2.remove(&pos).expect("cell exists in saved sheet");
        assert_eq!(cell1.value, cell2.value, "cell value mismatch at {:?}", pos);
    }
    assert_eq!(cells2.len(), 0, "unexpected extra cells in saved sheet");

    let formula_pos = (0, 2);
    let formula1 = sheet1
        .cells
        .iter()
        .find(|c| (c.row, c.col) == formula_pos)
        .and_then(|c| c.formula.as_ref())
        .expect("original formula cell has formula payload");
    let formula2 = sheet2
        .cells
        .iter()
        .find(|c| (c.row, c.col) == formula_pos)
        .and_then(|c| c.formula.as_ref())
        .expect("saved formula cell has formula payload");

    assert_eq!(formula1.rgce, formula2.rgce);
}
