use std::io::Write;

use formula_model::CellRef;

mod common;

use common::xls_fixture_builder;

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

#[test]
fn imports_phonetic_guides_from_biff8_sst_extrst() {
    const PHONETIC_MARKER: &str = "PHO_MARKER_123";
    let bytes = xls_fixture_builder::build_sst_phonetic_fixture_xls(PHONETIC_MARKER);
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("Sheet1")
        .expect("Sheet1 missing");
    let cell = sheet
        .cell(CellRef::from_a1("A1").unwrap())
        .expect("A1 missing");

    assert_eq!(cell.phonetic.as_deref(), Some(PHONETIC_MARKER));
}

