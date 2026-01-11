use formula_xlsb::{CellValue, XlsbWorkbook};

#[test]
fn resolves_xf_number_formats_for_date_cells() {
    let path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/tests/fixtures_styles/date.xlsb"
    );
    let wb = XlsbWorkbook::open(path).expect("open xlsb");

    let sheet = wb.read_sheet(0).expect("read sheet");
    let cell = sheet
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 0)
        .expect("A1 missing");

    assert_eq!(cell.value, CellValue::Number(44927.0));
    assert_eq!(cell.style, 1);

    let style = wb.styles().get(cell.style).expect("style mapping for XF");
    assert!(style.is_date_time);
    assert_eq!(style.number_format.as_deref(), Some("m/d/yyyy"));
}
