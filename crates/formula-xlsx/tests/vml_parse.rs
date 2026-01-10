use std::fs;

use formula_model::CellRef;
use formula_xlsx::comments::parse_vml_drawing_cells;
use formula_xlsx::XlsxPackage;

#[test]
fn parses_vml_note_cell_refs() {
    let fixture_path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/comments.xlsx");
    let bytes = fs::read(fixture_path).expect("fixture workbook should be readable");
    let pkg = XlsxPackage::from_bytes(&bytes).expect("fixture should parse as xlsx package");

    let vml = pkg
        .part("xl/drawings/vmlDrawing1.vml")
        .expect("fixture should contain vml drawing part");
    let cells = parse_vml_drawing_cells(vml).expect("vml should parse");
    assert!(cells.contains(&CellRef::new(0, 0)));
}
