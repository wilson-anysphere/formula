use std::io::Write;

use formula_model::Range;
use formula_model::{DefinedNameScope, XLNM_FILTER_DATABASE};

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

#[test]
fn warns_on_filtermode_and_sets_autofilter_from_sheet_stream_when_filterdatabase_missing() {
    let bytes = xls_fixture_builder::build_autofilter_filtermode_no_filterdatabase_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("FilteredNoDb")
        .expect("FilteredNoDb missing");
    let af = sheet
        .auto_filter
        .as_ref()
        .expect("expected sheet.auto_filter to be set");
    // Best-effort inference from DIMENSIONS + AUTOFILTERINFO: A1:B5.
    assert_eq!(af.range, Range::from_a1("A1:B5").expect("valid range"));

    assert!(
        result
            .workbook
            .get_defined_name(DefinedNameScope::Sheet(sheet.id), XLNM_FILTER_DATABASE)
            .is_none(),
        "unexpected _FilterDatabase defined name; defined_names={:?}",
        result.workbook.defined_names
    );

    let warning_substr = "sheet `FilteredNoDb` has FILTERMODE (filtered rows); filter criteria/hidden rows are not preserved on import";
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

#[test]
fn warns_on_filtermode_and_sets_autofilter_from_dimensions_when_autofilterinfo_missing() {
    let bytes = xls_fixture_builder::build_filtermode_only_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("FilterModeOnly")
        .expect("FilterModeOnly missing");
    let af = sheet
        .auto_filter
        .as_ref()
        .expect("expected sheet.auto_filter to be set");
    // Best-effort inference from DIMENSIONS (since AUTOFILTERINFO is missing): A1:C4.
    assert_eq!(af.range, Range::from_a1("A1:C4").expect("valid range"));

    assert!(
        result
            .workbook
            .get_defined_name(DefinedNameScope::Sheet(sheet.id), XLNM_FILTER_DATABASE)
            .is_none(),
        "unexpected _FilterDatabase defined name; defined_names={:?}",
        result.workbook.defined_names
    );

    let warning_substr = "sheet `FilterModeOnly` has FILTERMODE (filtered rows); filter criteria/hidden rows are not preserved on import";
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
