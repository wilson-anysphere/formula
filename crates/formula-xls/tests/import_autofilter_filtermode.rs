use formula_model::autofilter::{FilterCriterion, FilterValue};
use formula_model::Range;
use formula_model::{DefinedNameScope, XLNM_FILTER_DATABASE};

mod common;

use common::xls_fixture_builder;

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    formula_xls::import_xls_bytes(bytes).expect("import xls bytes")
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

    let warning_substr = "sheet `Filtered` has FILTERMODE record at offset";
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

    assert!(
        result.warnings.iter().any(|w| w
            .message
            .contains("FILTERMODE is set but no AutoFilter criteria records were found")),
        "expected missing-criteria warning when FILTERMODE is present without AUTOFILTER records; warnings={:?}",
        result.warnings
    );
}

#[test]
fn filtermode_preserves_filtered_rows_as_filter_hidden() {
    let bytes = xls_fixture_builder::build_autofilter_filtermode_hidden_rows_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("FilteredHiddenRows")
        .expect("FilteredHiddenRows missing");

    let af = sheet
        .auto_filter
        .as_ref()
        .expect("expected sheet.auto_filter to be set");
    assert_eq!(af.range, Range::from_a1("A1:B3").expect("valid range"));

    assert!(
        !af.filter_columns.is_empty(),
        "expected filter criteria to be imported; auto_filter={af:?}; warnings={:?}",
        result.warnings
    );
    assert_eq!(af.filter_columns[0].col_id, 0);
    assert_eq!(
        af.filter_columns[0].criteria,
        vec![FilterCriterion::Equals(FilterValue::Text("X".to_string()))]
    );

    // Row 2 (1-based) is hidden in the BIFF row metadata, but when FILTERMODE is present we do not
    // preserve filtered-row visibility as user-hidden rows. Instead it is mapped to
    // `OutlineEntry.hidden.filter`.
    assert!(
        sheet.is_row_hidden_effective(2),
        "expected row 2 to be hidden by filter; row_props={:?}; outline={:?}; warnings={:?}",
        sheet.row_properties(1),
        sheet.row_outline_entry(2),
        result.warnings
    );

    let outline = sheet.row_outline_entry(2);
    assert!(
        outline.hidden.filter,
        "expected row 2 to be filter-hidden; outline={outline:?}; warnings={:?}",
        result.warnings
    );
    assert!(
        !outline.hidden.user,
        "expected row 2 to not be user-hidden; outline={outline:?}; warnings={:?}",
        result.warnings
    );
}

#[test]
fn filtermode_sets_filter_hidden_flag_on_outline_entry() {
    // Explicit regression test for `OutlineEntry.hidden.filter` when a BIFF ROW record is hidden
    // inside an active FILTERMODE range.
    let bytes = xls_fixture_builder::build_autofilter_filtermode_hidden_rows_fixture_xls();
    let result = import_fixture(&bytes);
    let sheet = result
        .workbook
        .sheet_by_name("FilteredHiddenRows")
        .expect("FilteredHiddenRows missing");

    assert!(
        sheet.row_outline_entry(2).hidden.filter,
        "expected row 2 to be filter-hidden; outline={:?}; warnings={:?}",
        sheet.row_outline_entry(2),
        result.warnings
    );
}

#[test]
fn filtermode_does_not_reclassify_user_hidden_rows_outside_final_autofilter_range() {
    let bytes = xls_fixture_builder::build_autofilter_workbook_scope_unqualified_multisheet_filtermode_hidden_row_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("Unqualified")
        .expect("Unqualified missing");

    let af = sheet
        .auto_filter
        .as_ref()
        .expect("expected sheet.auto_filter to be set");
    assert_eq!(af.range, Range::from_a1("A1:B3").expect("valid range"));

    // Row 4 (1-based) is hidden in the BIFF row metadata but lies outside the final AutoFilter
    // range (A1:B3), so it should remain user-hidden rather than being treated as filter-hidden.
    let entry = sheet.row_outline_entry(4);
    assert!(
        entry.hidden.user,
        "expected row 4 to remain user-hidden; entry={:?}; warnings={:?}",
        entry,
        result.warnings
    );
    assert!(
        !entry.hidden.filter,
        "expected row 4 to not be filter-hidden; entry={:?}; warnings={:?}",
        entry,
        result.warnings
    );
}

#[test]
fn filtermode_does_not_override_outline_hidden_rows() {
    let bytes = xls_fixture_builder::build_autofilter_filtermode_outline_hidden_rows_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("FilteredOutlineHiddenRows")
        .expect("FilteredOutlineHiddenRows missing");

    let af = sheet
        .auto_filter
        .as_ref()
        .expect("expected sheet.auto_filter to be set");
    assert_eq!(af.range, Range::from_a1("A1:B4").expect("valid range"));

    // Row 2 (1-based) is hidden by a collapsed outline group. FILTERMODE should not reclassify it
    // as filter-hidden, and it should not be surfaced as user-hidden.
    let entry = sheet.row_outline_entry(2);
    assert!(
        entry.hidden.outline,
        "expected row 2 to be outline-hidden; entry={entry:?}; warnings={:?}",
        result.warnings
    );
    assert!(
        !entry.hidden.filter,
        "expected row 2 to not be filter-hidden; entry={entry:?}; warnings={:?}",
        result.warnings
    );
    assert!(
        !entry.hidden.user,
        "expected row 2 to not be user-hidden; entry={entry:?}; warnings={:?}",
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

    let warning_substr = "sheet `FilteredNoDb` has FILTERMODE record at offset";
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

    let warning_substr = "sheet `FilterModeOnly` has FILTERMODE record at offset";
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
fn sets_autofilter_from_sheet_stream_when_autofilterinfo_present_without_filtermode() {
    let bytes = xls_fixture_builder::build_autofilterinfo_only_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("AutoFilterInfoOnly")
        .expect("AutoFilterInfoOnly missing");
    let af = sheet
        .auto_filter
        .as_ref()
        .expect("expected sheet.auto_filter to be set");
    // DIMENSIONS is A1:D4 but AUTOFILTERINFO says only 2 columns, so clamp to A1:B4.
    assert_eq!(af.range, Range::from_a1("A1:B4").expect("valid range"));

    assert!(
        result
            .workbook
            .get_defined_name(DefinedNameScope::Sheet(sheet.id), XLNM_FILTER_DATABASE)
            .is_none(),
        "unexpected _FilterDatabase defined name; defined_names={:?}",
        result.workbook.defined_names
    );

    assert!(
        !result
            .warnings
            .iter()
            .any(|w| w.message.contains("has FILTERMODE record")),
        "unexpected FILTERMODE warning(s): {:?}",
        result.warnings
    );
}
