use formula_xlsb::{OpenOptions, XlsbWorkbook};

#[test]
fn open_options_decode_formulas_false_skips_formula_text_decoding() {
    let path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/simple.xlsb");
    let wb = XlsbWorkbook::open_with_options(
        path,
        OpenOptions {
            decode_formulas: false,
            ..OpenOptions::default()
        },
    )
    .expect("open xlsb");

    let sheet = wb.read_sheet(0).expect("read sheet");
    let formula_cell = sheet
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 2)
        .expect("expected formula cell at C1");
    let formula = formula_cell.formula.as_ref().expect("formula payload preserved");

    assert!(
        !formula.rgce.is_empty(),
        "expected formula rgce bytes to be preserved"
    );
    assert_eq!(
        formula.text, None,
        "expected formula text decoding to be skipped"
    );
    assert!(
        formula.warnings.is_empty(),
        "expected formula warnings to be empty when decoding is skipped"
    );
}

