use std::io::Write;

use formula_model::SheetVisibility;

mod common;

use common::xls_fixture_builder;

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

#[test]
fn imports_sheet_visibility_from_biff_boundsheet() {
    let bytes = xls_fixture_builder::build_sheet_visibility_fixture_xls();
    let result = import_fixture(&bytes);

    let visible = result
        .workbook
        .sheet_by_name("Visible")
        .expect("Visible sheet missing");
    assert_eq!(visible.visibility, SheetVisibility::Visible);

    let hidden = result
        .workbook
        .sheet_by_name("Hidden")
        .expect("Hidden sheet missing");
    assert_eq!(hidden.visibility, SheetVisibility::Hidden);

    let very_hidden = result
        .workbook
        .sheet_by_name("VeryHidden")
        .expect("VeryHidden sheet missing");
    assert_eq!(very_hidden.visibility, SheetVisibility::VeryHidden);
}

