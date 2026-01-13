use std::io::Write;

mod common;

use common::xls_fixture_builder;

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

#[test]
fn imports_row_and_col_default_styles_from_ixfe() {
    let bytes = xls_fixture_builder::build_row_col_style_fixture_xls();
    let result = import_fixture(&bytes);

    assert!(
        result.warnings.is_empty(),
        "expected no warnings, got {:?}",
        result.warnings
    );

    let sheet = result
        .workbook
        .sheet_by_name("RowColStyles")
        .expect("RowColStyles sheet missing");

    // Row 2 (1-based) => index 1.
    let row_style_id = sheet
        .row_properties
        .get(&1)
        .and_then(|p| p.style_id)
        .expect("expected row style id for row 2");
    assert_ne!(row_style_id, 0, "expected a non-default row style id");
    let row_style = result
        .workbook
        .styles
        .get(row_style_id)
        .expect("row style missing");
    assert_eq!(row_style.number_format.as_deref(), Some("0.00%"));

    // Column C => index 2.
    let col_style_id = sheet
        .col_properties
        .get(&2)
        .and_then(|p| p.style_id)
        .expect("expected col style id for column C");
    assert_ne!(col_style_id, 0, "expected a non-default col style id");
    let col_style = result
        .workbook
        .styles
        .get(col_style_id)
        .expect("col style missing");
    assert_eq!(col_style.number_format.as_deref(), Some("[h]:mm:ss"));

    assert_ne!(
        row_style_id, col_style_id,
        "expected distinct row/col style ids"
    );
}
