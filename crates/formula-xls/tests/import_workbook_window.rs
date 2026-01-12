use std::io::Write;

use formula_model::WorkbookWindowState;

mod common;

use common::xls_fixture_builder;

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

#[test]
fn imports_workbook_window_from_biff_window1() {
    let bytes = xls_fixture_builder::build_workbook_window_fixture_xls();
    let result = import_fixture(&bytes);

    let window = result
        .workbook
        .view
        .window
        .as_ref()
        .expect("workbook window missing");

    assert_eq!(window.x, Some(120));
    assert_eq!(window.y, Some(240));
    assert_eq!(window.width, Some(800));
    assert_eq!(window.height, Some(600));
    assert_eq!(window.state, Some(WorkbookWindowState::Minimized));
}

