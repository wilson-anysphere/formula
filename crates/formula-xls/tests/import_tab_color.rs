use std::io::Write;

mod common;

use common::xls_fixture_builder;

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

#[test]
fn imports_sheet_tab_color_from_sheetext() {
    let bytes = xls_fixture_builder::build_tab_color_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("TabColor")
        .expect("TabColor missing");

    let color = sheet.tab_color.as_ref().expect("tab_color missing");
    assert_eq!(color.rgb.as_deref(), Some("FF112233"));
    assert!(color.indexed.is_none(), "expected rgb-only tab color");
}

