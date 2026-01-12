use formula_xlsb::{CellValue, OpenOptions, XlsbWorkbook};
use std::io::Write;
use std::ops::ControlFlow;
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
fn open_options_decode_formulas_skips_formula_text_decoding() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_sheet_name("Sheet1");
    builder.set_cell_number(0, 1, 42.5);

    let mut rgce_bytes = Vec::new();
    rgce::push_ref(&mut rgce_bytes, 0, 1, false, false); // B1
    rgce::push_int(&mut rgce_bytes, 2);
    rgce::push_mul(&mut rgce_bytes);

    builder.set_cell_formula_num(0, 2, 85.0, rgce_bytes.clone(), Vec::new());

    let bytes = builder.build_bytes();
    let tmp = write_temp_xlsb(&bytes);

    let wb = XlsbWorkbook::open_with_options(
        tmp.path(),
        OpenOptions {
            decode_formulas: false,
            ..Default::default()
        },
    )
    .expect("open xlsb with decode_formulas=false");

    let sheet = wb.read_sheet(0).expect("read sheet");
    let cell = sheet
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 2)
        .expect("expected formula cell at C1");
    assert_eq!(cell.value, CellValue::Number(85.0));
    let formula = cell.formula.as_ref().expect("formula payload preserved");
    assert!(
        !formula.rgce.is_empty(),
        "expected formula rgce bytes to be preserved"
    );
    assert!(formula.text.is_none(), "expected formula.text to be None");
    assert_eq!(formula.rgce, rgce_bytes);
    assert!(formula.warnings.is_empty(), "expected no decode warnings");

    // Also ensure the streaming worksheet path honors the option.
    let mut streamed = None;
    wb.for_each_cell_control_flow(0, |cell| {
        if cell.row == 0 && cell.col == 2 {
            streamed = Some(cell);
            ControlFlow::Break(())
        } else {
            ControlFlow::Continue(())
        }
    })
    .expect("stream cells");
    let streamed = streamed.expect("expected to find formula cell at C1 via streaming");
    let streamed_formula = streamed.formula.as_ref().expect("formula payload preserved");
    assert!(
        !streamed_formula.rgce.is_empty(),
        "expected streamed formula rgce bytes to be preserved"
    );
    assert_eq!(
        streamed_formula.text, None,
        "expected streamed formula text decoding to be skipped"
    );
    assert!(
        streamed_formula.warnings.is_empty(),
        "expected streamed formula warnings to be empty when decoding is skipped"
    );

    // Default behavior still decodes formula text.
    let wb_default = XlsbWorkbook::open(tmp.path()).expect("open xlsb with default options");
    let sheet_default = wb_default.read_sheet(0).expect("read sheet");
    let cell_default = sheet_default
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 2)
        .expect("expected formula cell C1");
    let formula_default = cell_default.formula.as_ref().expect("formula payload preserved");
    assert_eq!(formula_default.text.as_deref(), Some("B1*2"));
}

