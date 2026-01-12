use std::io::Write;

use formula_model::Range;

mod common;

use common::xls_fixture_builder;

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

#[test]
fn warns_on_filtermode_and_preserves_autofilter_dropdown_range() {
    let bytes = xls_fixture_builder::build_autofilter_filtermode_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("Filtered")
        .expect("Filtered missing");
    let af = sheet
        .auto_filter
        .as_ref()
        .expect("expected sheet.auto_filter to be set");
    assert_eq!(af.range, Range::from_a1("A1:B3").expect("valid range"));

    let warning_substr = "sheet `Filtered` has FILTERMODE (filtered rows); filter criteria/hidden rows are not preserved on import";
    let matching: Vec<&formula_xls::ImportWarning> = result
        .warnings
        .iter()
        .filter(|w| w.message.contains(warning_substr))
        .collect();
    assert_eq!(
        matching.len(),
        1,
        "expected exactly one FILTERMODE warning, got warnings={:?}",
        result.warnings
    );
}
