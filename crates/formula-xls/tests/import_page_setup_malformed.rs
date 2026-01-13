use std::io::Write;

use formula_model::{CellRef, PageSetup};

mod common;

use common::xls_fixture_builder;

fn import_fixture_no_panic(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");

    let result = std::panic::catch_unwind(|| formula_xls::import_xls_path(tmp.path()));
    match result {
        Ok(Ok(result)) => result,
        Ok(Err(err)) => panic!("expected import to succeed, got error: {err:?}"),
        Err(_) => panic!("expected import not to panic"),
    }
}

#[test]
fn ignores_truncated_page_setup_records_without_panicking() {
    let bytes = xls_fixture_builder::build_page_setup_malformed_fixture_xls();
    let result = import_fixture_no_panic(&bytes);

    // Workbook still loads the sheet and at least one cell.
    let sheet = result
        .workbook
        .sheet_by_name("PageSetupMalformed")
        .expect("expected sheet to exist");
    assert!(
        sheet.cell(CellRef::new(0, 0)).is_some(),
        "expected A1 cell to be present"
    );

    // Print settings parsing should fail best-effort, leaving defaults/empty.
    let settings = result
        .workbook
        .sheet_print_settings_by_name("PageSetupMalformed");
    assert_eq!(settings.page_setup, PageSetup::default());
    assert!(
        settings.manual_page_breaks.is_empty(),
        "expected manual page breaks to be empty, got {:?}",
        settings.manual_page_breaks
    );

    // Warnings surfaced for truncated payloads.
    let warnings: Vec<&str> = result.warnings.iter().map(|w| w.message.as_str()).collect();
    assert!(
        warnings
            .iter()
            .any(|w| w.contains("truncated SETUP record")),
        "expected truncated SETUP warning, got: {warnings:?}"
    );
    assert!(
        warnings
            .iter()
            .any(|w| w.contains("truncated LEFTMARGIN record")),
        "expected truncated LEFTMARGIN warning, got: {warnings:?}"
    );
    assert!(
        warnings
            .iter()
            .any(|w| w.contains("HorizontalPageBreaks")),
        "expected HorizontalPageBreaks warning, got: {warnings:?}"
    );
}
