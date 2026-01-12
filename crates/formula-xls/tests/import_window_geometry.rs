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
fn imports_workbook_window_geometry_from_window1() {
    let bytes = xls_fixture_builder::build_window_geometry_fixture_xls();
    let result = import_fixture(&bytes);

    let window = result
        .workbook
        .view
        .window
        .as_ref()
        .expect("expected Workbook.view.window");

    assert_eq!(window.x, Some(100));
    assert_eq!(window.y, Some(200));
    assert_eq!(window.width, Some(300));
    assert_eq!(window.height, Some(400));
    assert_eq!(window.state, Some(WorkbookWindowState::Maximized));
}

