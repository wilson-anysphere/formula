use std::io::Write;

use formula_engine::{parse_formula, ParseOptions};
use formula_model::Range;
use formula_model::{DefinedNameScope, XLNM_FILTER_DATABASE};

mod common;

use common::xls_fixture_builder;

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

fn import_fixture_without_biff(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path_without_biff(tmp.path()).expect("import xls")
}

fn assert_parseable(expr: &str) {
    let expr = expr.trim();
    assert!(!expr.is_empty(), "expected expression to be non-empty");
    parse_formula(expr, ParseOptions::default()).unwrap_or_else(|e| {
        panic!("expected expression to be parseable, expr={expr:?}, err={e:?}")
    });
}

#[test]
fn imports_autofilter_range_from_filterdatabase_defined_name() {
    let bytes = xls_fixture_builder::build_defined_names_builtins_fixture_xls();
    let result = import_fixture(&bytes);
    let sheet = result
        .workbook
        .sheet_by_name("Sheet1")
        .expect("Sheet1 missing");

    let auto_filter = sheet
        .auto_filter
        .as_ref()
        .expect("expected Sheet1.auto_filter to be set");

    assert_eq!(auto_filter.range, Range::from_a1("A1:C10").unwrap());

    let filter_db = result
        .workbook
        .get_defined_name(DefinedNameScope::Sheet(sheet.id), XLNM_FILTER_DATABASE)
        .expect("expected _FilterDatabase defined name");
    assert_parseable(&filter_db.refers_to);
}

#[test]
fn imports_autofilter_fixture_range_and_empty_state() {
    let bytes = xls_fixture_builder::build_autofilter_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("Filter")
        .expect("Filter missing");
    let auto_filter = sheet.auto_filter.as_ref().expect("auto_filter missing");

    assert_eq!(auto_filter.range, Range::from_a1("A1:C5").unwrap());
    assert!(auto_filter.filter_columns.is_empty());
    assert!(auto_filter.sort_state.is_none());
    assert!(auto_filter.raw_xml.is_empty());

    let filter_db = result
        .workbook
        .get_defined_name(DefinedNameScope::Sheet(sheet.id), XLNM_FILTER_DATABASE)
        .expect("expected _FilterDatabase defined name");
    assert_parseable(&filter_db.refers_to);
}

#[test]
fn imports_autofilter_range_from_filterdatabase_defined_name_without_biff() {
    let bytes = xls_fixture_builder::build_autofilter_fixture_xls();
    let result = import_fixture_without_biff(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("Filter")
        .expect("Filter missing");

    let auto_filter = sheet.auto_filter.as_ref().expect("auto_filter missing");
    assert_eq!(auto_filter.range, Range::from_a1("A1:C5").unwrap());
    assert!(auto_filter.filter_columns.is_empty());
    assert!(auto_filter.sort_state.is_none());
    assert!(auto_filter.raw_xml.is_empty());
}

#[test]
fn imports_autofilter_range_via_calamine_defined_name_fallback_when_biff_unavailable() {
    let bytes = xls_fixture_builder::build_autofilter_calamine_fixture_xls();
    let result = import_fixture_without_biff(&bytes);

    let sheet = result.workbook.sheet_by_name("Filter").expect("Filter missing");

    let auto_filter = sheet.auto_filter.as_ref().unwrap_or_else(|| {
        panic!(
            "auto_filter missing; defined_names={:?}; warnings={:?}",
            result.workbook.defined_names, result.warnings
        )
    });
    assert_eq!(auto_filter.range, Range::from_a1("A1:C5").unwrap());
    assert!(auto_filter.filter_columns.is_empty());
    assert!(auto_filter.sort_state.is_none());
    assert!(auto_filter.raw_xml.is_empty());

    // Calamine fallback does not reliably surface built-in defined names like `_FilterDatabase`.
    // If the name was imported anyway (e.g. due to future calamine improvements), it should remain
    // parseable.
    if let Some(filter_db) = result
        .workbook
        .get_defined_name(DefinedNameScope::Sheet(sheet.id), XLNM_FILTER_DATABASE)
    {
        assert_parseable(&filter_db.refers_to);
    }
}

#[test]
fn imports_autofilter_range_from_workbook_scope_filterdatabase_name_via_externsheet() {
    let bytes = xls_fixture_builder::build_autofilter_workbook_scope_externsheet_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("AutoFilter")
        .expect("AutoFilter sheet missing");

    let auto_filter = sheet.auto_filter.as_ref().expect("auto_filter missing");
    assert_eq!(auto_filter.range, Range::from_a1("A1:C5").unwrap());
    assert!(auto_filter.filter_columns.is_empty());
    assert!(auto_filter.sort_state.is_none());
    assert!(auto_filter.raw_xml.is_empty());

    let filter_db = result
        .workbook
        .get_defined_name(DefinedNameScope::Workbook, XLNM_FILTER_DATABASE)
        .expect("expected workbook-scoped _FilterDatabase defined name");
    assert_parseable(&filter_db.refers_to);
}

#[test]
fn imports_autofilter_range_from_workbook_scope_filterdatabase_name_without_biff() {
    let bytes = xls_fixture_builder::build_autofilter_workbook_scope_externsheet_fixture_xls();
    let result = import_fixture_without_biff(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("AutoFilter")
        .expect("AutoFilter sheet missing");

    let af = sheet.auto_filter.as_ref().unwrap_or_else(|| {
        panic!(
            "auto_filter missing; defined_names={:?}; warnings={:?}",
            result.workbook.defined_names, result.warnings
        )
    });

    assert_eq!(af.range, Range::from_a1("A1:C5").unwrap());
    assert!(af.filter_columns.is_empty());
    assert!(af.sort_state.is_none());
    assert!(af.raw_xml.is_empty());
}

#[test]
fn imports_autofilter_range_from_filterdatabase_alias_defined_name_without_biff() {
    // Some `.xls` files (or decoders like calamine) surface the AutoFilter built-in name without
    // the `_xlnm.` prefix (i.e. `_FilterDatabase`). Ensure we still import the AutoFilter range
    // from the calamine defined-name fallback path.
    let bytes = xls_fixture_builder::build_autofilter_calamine_filterdatabase_alias_fixture_xls();
    let result = import_fixture_without_biff(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("AutoFilter")
        .expect("AutoFilter sheet missing");

    let af = sheet.auto_filter.as_ref().unwrap_or_else(|| {
        panic!(
            "auto_filter missing; defined_names={:?}; warnings={:?}",
            result.workbook.defined_names, result.warnings
        )
    });
    assert_eq!(af.range, Range::from_a1("A1:C5").unwrap());
    assert!(af.filter_columns.is_empty());
    assert!(af.sort_state.is_none());
    assert!(af.raw_xml.is_empty());
}

#[test]
fn recovers_missing_autofilter_ranges_when_calamine_import_is_partial() {
    let bytes = xls_fixture_builder::build_autofilter_mixed_calamine_and_builtin_fixture_xls();
    let result = import_fixture_without_biff(&bytes);

    let calamine_sheet = result
        .workbook
        .sheet_by_name("Calamine")
        .expect("Calamine sheet missing");
    let builtin_sheet = result
        .workbook
        .sheet_by_name("Builtin")
        .expect("Builtin sheet missing");

    let calamine_af = calamine_sheet
        .auto_filter
        .as_ref()
        .expect("expected Calamine.auto_filter to be set");
    assert_eq!(calamine_af.range, Range::from_a1("A1:C5").unwrap());

    let builtin_af = builtin_sheet.auto_filter.as_ref().unwrap_or_else(|| {
        panic!(
            "expected Builtin.auto_filter to be set; defined_names={:?}; warnings={:?}",
            result.workbook.defined_names, result.warnings
        )
    });
    assert_eq!(builtin_af.range, Range::from_a1("A1:B3").unwrap());
}
