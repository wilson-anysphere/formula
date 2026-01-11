use formula_model::CellValue;
use formula_xlsx::read_workbook;

#[test]
fn read_workbook_rich_text_shared_strings_plain_text() {
    let fixture_path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/styles/rich-text-shared-strings.xlsx"
    );
    let wb = read_workbook(fixture_path).expect("read workbook");
    let sheet = wb.sheets.first().expect("workbook has at least one sheet");

    let value = sheet.value_a1("A1").expect("parse A1");
    let text = match value {
        CellValue::String(s) => s,
        CellValue::RichText(r) => r.text,
        other => panic!("expected string cell value, got {other:?}"),
    };

    assert_eq!(text, "Hello Bold Italic");
}

