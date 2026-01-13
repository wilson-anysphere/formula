use formula_xlsb::{CellValue, OpenOptions, XlsbWorkbook};
use pretty_assertions::assert_eq;
use std::io::Cursor;
use std::sync::Arc;

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

    let wb = XlsbWorkbook::from_bytes(Arc::from(bytes), OpenOptions::default())
        .expect("open xlsb from bytes");
    assert_eq!(wb.sheet_metas().len(), 1);
    assert!(wb.preserved_parts().contains_key("[Content_Types].xml"));

    let sheet = wb.read_sheet(0).expect("read sheet");
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

#[test]
fn in_memory_save_as_to_writer_is_lossless_at_opc_part_level() {
    let path = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/simple.xlsb");
    let bytes = std::fs::read(path).expect("read fixture bytes");

    let wb = XlsbWorkbook::open_from_vec(bytes.clone()).expect("open xlsb from vec");

    let mut out = Cursor::new(Vec::new());
    wb.save_as_to_writer(&mut out)
        .expect("save_as_to_writer should succeed");
    let out_bytes = out.into_inner();

    let expected = xlsx_diff::WorkbookArchive::from_bytes(&bytes).expect("read expected archive");
    let actual = xlsx_diff::WorkbookArchive::from_bytes(&out_bytes).expect("read actual archive");
    let report = xlsx_diff::diff_archives(&expected, &actual);
    assert!(
        report.is_empty(),
        "expected no OPC part diffs, got:\n{}",
        report
            .differences
            .iter()
            .map(|d| d.to_string())
            .collect::<Vec<_>>()
            .join("\n")
    );
}
