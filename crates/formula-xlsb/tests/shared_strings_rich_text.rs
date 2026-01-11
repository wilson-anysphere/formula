use formula_xlsb::{CellValue, XlsbWorkbook};
use pretty_assertions::assert_eq;

#[test]
fn parses_rich_shared_strings_and_preserves_runs() {
    let path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/tests/fixtures/rich_shared_strings.xlsb"
    );
    let wb = XlsbWorkbook::open(path).expect("open xlsb");

    assert_eq!(wb.shared_strings().len(), 1);
    assert_eq!(wb.shared_strings()[0], "Hello Bold");

    let si = &wb.shared_strings_table()[0];
    assert_eq!(si.plain_text(), "Hello Bold");
    assert_eq!(si.rich_text.runs.len(), 2);
    assert_eq!(si.rich_text.slice_run_text(&si.rich_text.runs[0]), "Hello ");
    assert_eq!(si.rich_text.slice_run_text(&si.rich_text.runs[1]), "Bold");

    assert_eq!(si.run_formats.len(), 2);
    assert_eq!(si.run_formats[0], vec![0, 0, 0, 0]);
    assert_eq!(si.run_formats[1], vec![1, 0, 0, 0]);
    assert_eq!(si.phonetic, None);
    assert!(si.raw_si.is_some(), "rich shared strings should preserve raw SI bytes");

    // Ensure worksheet string lookup uses the plain-text projection.
    let sheet = wb.read_sheet(0).expect("read sheet1");
    let a1 = sheet
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 0)
        .expect("A1 exists");
    assert_eq!(a1.value, CellValue::Text("Hello Bold".to_string()));
}

