use formula_xlsb::{CellValue, XlsbWorkbook};
use pretty_assertions::assert_eq;

#[test]
fn saves_xlsb_as_encrypted_and_reopens_with_password() {
    let path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/simple.xlsb");
    let wb = XlsbWorkbook::open(path).expect("open xlsb");

    let tmp = tempfile::tempdir().expect("tempdir");
    let encrypted_path = tmp.path().join("encrypted.xlsb");

    wb.save_as_encrypted(&encrypted_path, "password123")
        .expect("save_as_encrypted");

    // OLE magic bytes.
    let encrypted_bytes = std::fs::read(&encrypted_path).expect("read encrypted output");
    assert!(
        encrypted_bytes.starts_with(&[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]),
        "expected OLE header"
    );

    let reopened = XlsbWorkbook::open_with_password(&encrypted_path, "password123")
        .expect("open_with_password");

    assert_eq!(reopened.sheet_metas().len(), 1);
    assert_eq!(reopened.sheet_metas()[0].name, "Sheet1");

    let sheet = reopened.read_sheet(0).expect("read sheet1");
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
}
