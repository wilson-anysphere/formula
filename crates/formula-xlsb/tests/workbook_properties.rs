use formula_format::{format_value, DateSystem, FormatOptions, Value};
use formula_xlsb::{CellValue, SheetVisibility, XlsbWorkbook};
use pretty_assertions::{assert_eq, assert_ne};

#[test]
fn opens_1904_fixture_and_formats_date_serials() {
    let path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/date1904.xlsb");
    let wb = XlsbWorkbook::open(path).expect("open xlsb");

    assert!(wb.workbook_properties().date_system_1904);
    assert_eq!(wb.sheet_metas().len(), 1);
    assert_eq!(wb.sheet_metas()[0].visibility, SheetVisibility::Visible);

    // In the fixture, B1 contains the numeric serial 0.0. For 1904 workbooks this is 1904-01-01,
    // while for 1900 workbooks it is 1899-12-31.
    let sheet = wb.read_sheet(0).expect("read sheet1");
    let b1 = sheet
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 1)
        .expect("B1 present");
    let CellValue::Number(serial) = b1.value else {
        panic!("expected B1 to be numeric");
    };

    let mut opts_1904 = FormatOptions::default();
    opts_1904.date_system = DateSystem::Excel1904;
    let text_1904 = format_value(Value::Number(serial), Some("m/d/yyyy"), &opts_1904).text;

    let mut opts_1900 = FormatOptions::default();
    opts_1900.date_system = DateSystem::Excel1900;
    let text_1900 = format_value(Value::Number(serial), Some("m/d/yyyy"), &opts_1900).text;

    assert_ne!(text_1900, text_1904);
    assert_eq!(text_1904, "1/1/1904");
    assert_eq!(text_1900, "12/31/1899");
}

