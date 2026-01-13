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

#[test]
fn imports_biff_sheet_protection_allow_flags() {
    let bytes = xls_fixture_builder::build_sheet_protection_allow_flags_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheets
        .first()
        .expect("fixture should contain one sheet");
    let p = &sheet.sheet_protection;

    assert_eq!(p.enabled, true);
    assert_eq!(p.password_hash, Some(0xCBEB));
    assert_eq!(p.edit_objects, true);
    assert_eq!(p.edit_scenarios, true);

    // Enhanced allow flags from FEAT/FEATHEADR records.
    assert_eq!(p.select_locked_cells, false);
    assert_eq!(p.select_unlocked_cells, true);
    assert_eq!(p.format_cells, true);
    assert_eq!(p.format_columns, true);
    assert_eq!(p.format_rows, false);
    assert_eq!(p.insert_columns, true);
    assert_eq!(p.insert_rows, false);
    assert_eq!(p.insert_hyperlinks, true);
    assert_eq!(p.delete_columns, false);
    assert_eq!(p.delete_rows, true);
    assert_eq!(p.sort, true);
    assert_eq!(p.auto_filter, true);
    assert_eq!(p.pivot_tables, false);
}

#[test]
fn warns_on_malformed_feat_protection_record_but_continues() {
    let bytes = xls_fixture_builder::build_sheet_protection_allow_flags_malformed_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheets
        .first()
        .expect("fixture should contain one sheet");
    let p = &sheet.sheet_protection;

    // Allow flags still imported from the valid record.
    assert_eq!(p.enabled, true);
    assert_eq!(p.select_unlocked_cells, true);
    assert_eq!(p.format_cells, true);
    assert_eq!(p.sort, true);
    assert_eq!(p.auto_filter, true);

    // Malformed FEAT record surfaces a warning but does not abort import.
    let warnings: Vec<&str> = result.warnings.iter().map(|w| w.message.as_str()).collect();
    assert!(
        warnings.iter().any(|w| w.contains("failed to parse FEAT record")),
        "expected FEAT warning, got: {warnings:?}"
    );
}
