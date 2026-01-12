use std::io::Write;

mod common;

use common::xls_fixture_builder;

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

#[test]
fn imports_biff_workbook_and_sheet_protection() {
    let bytes = xls_fixture_builder::build_protection_fixture_xls();
    let result = import_fixture(&bytes);

    assert_eq!(result.workbook.workbook_protection.lock_structure, true);
    assert_eq!(result.workbook.workbook_protection.lock_windows, true);
    assert_eq!(
        result.workbook.workbook_protection.password_hash,
        Some(0x83AF)
    );

    let sheet = result
        .workbook
        .sheets
        .first()
        .expect("fixture should contain one sheet");
    assert_eq!(sheet.sheet_protection.enabled, true);
    assert_eq!(sheet.sheet_protection.password_hash, Some(0xCBEB));
    assert_eq!(sheet.sheet_protection.edit_objects, true);
    assert_eq!(sheet.sheet_protection.edit_scenarios, true);
}

#[test]
fn warns_on_truncated_biff_protection_records_but_continues() {
    let bytes = xls_fixture_builder::build_protection_truncated_fixture_xls();
    let result = import_fixture(&bytes);

    // Final values still imported.
    assert_eq!(result.workbook.workbook_protection.lock_structure, true);
    assert_eq!(result.workbook.workbook_protection.lock_windows, true);
    assert_eq!(
        result.workbook.workbook_protection.password_hash,
        Some(0x83AF)
    );

    let sheet = result
        .workbook
        .sheets
        .first()
        .expect("fixture should contain one sheet");
    assert_eq!(sheet.sheet_protection.enabled, true);
    assert_eq!(sheet.sheet_protection.password_hash, Some(0xCBEB));
    assert_eq!(sheet.sheet_protection.edit_objects, true);
    assert_eq!(sheet.sheet_protection.edit_scenarios, true);

    // Warnings surfaced for truncated payloads.
    let warnings: Vec<&str> = result.warnings.iter().map(|w| w.message.as_str()).collect();
    assert!(
        warnings.iter().any(|w| w.contains("truncated PROTECT record")),
        "expected truncated PROTECT warning, got: {warnings:?}"
    );
    assert!(
        warnings.iter().any(|w| w.contains("truncated WINDOWPROTECT record")),
        "expected truncated WINDOWPROTECT warning, got: {warnings:?}"
    );
    assert!(
        warnings.iter().any(|w| w.contains("truncated PASSWORD record")),
        "expected truncated PASSWORD warning, got: {warnings:?}"
    );
    assert!(
        warnings.iter().any(|w| w.contains("truncated OBJPROTECT record")),
        "expected truncated OBJPROTECT warning, got: {warnings:?}"
    );
    assert!(
        warnings.iter().any(|w| w.contains("truncated SCENPROTECT record")),
        "expected truncated SCENPROTECT warning, got: {warnings:?}"
    );
}
