use formula_xlsb::XlsbWorkbook;
use std::io::Cursor;

#[test]
fn opens_fixture_from_reader_bytes() {
    let path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/simple.xlsb");
    let bytes = std::fs::read(path).expect("read fixture bytes");

    let wb = XlsbWorkbook::open_from_reader(Cursor::new(bytes)).expect("open xlsb from reader");
    assert_eq!(wb.sheet_metas().len(), 1);
    assert!(wb.preserved_parts().contains_key("[Content_Types].xml"));

    // Ensure worksheet streaming works without a filesystem path.
    let mut saw_a1 = false;
    wb.for_each_cell(0, |cell| {
        if cell.row == 0 && cell.col == 0 {
            saw_a1 = true;
        }
    })
    .expect("stream sheet cells");
    assert!(saw_a1, "expected to see cell A1 in streamed iteration");
}

#[test]
fn opens_fixture_from_bytes() {
    let path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/simple.xlsb");
    let bytes = std::fs::read(path).expect("read fixture bytes");

    let wb = XlsbWorkbook::open_from_bytes(&bytes).expect("open xlsb from bytes");
    assert_eq!(wb.sheet_metas().len(), 1);
    assert!(wb.preserved_parts().contains_key("[Content_Types].xml"));

    let sheet = wb.read_sheet(0).expect("read sheet");
    assert!(
        sheet.cells.iter().any(|c| c.row == 0 && c.col == 0),
        "expected cell A1 in parsed sheet"
    );
}

#[test]
fn opens_fixture_from_vec_without_copy() {
    let path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/simple.xlsb");
    let bytes = std::fs::read(path).expect("read fixture bytes");

    // `open_from_vec` consumes the Vec, avoiding an extra copy for callers that already have an
    // owned buffer (e.g. decrypted EncryptedPackage bytes).
    let wb = XlsbWorkbook::open_from_vec(bytes).expect("open xlsb from vec");
    assert_eq!(wb.sheet_metas().len(), 1);
    let sheet = wb.read_sheet(0).expect("read sheet");
    assert!(sheet.cells.iter().any(|c| c.row == 0 && c.col == 0));
}
