use std::io::Write;
use std::collections::BTreeSet;

use formula_model::{ColRange, Orientation, Range, RowRange, Scaling};
use formula_model::{DefinedNameScope, XLNM_PRINT_AREA, XLNM_PRINT_TITLES};

mod common;

use common::{assert_parseable_formula, xls_fixture_builder};

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

#[test]
fn imports_print_settings_from_biff_builtin_defined_names() {
    let bytes = xls_fixture_builder::build_defined_names_builtins_fixture_xls();
    let result = import_fixture(&bytes);
    let workbook = result.workbook;

    let sheet1_settings = workbook.sheet_print_settings_by_name("Sheet1");
    assert_eq!(
        sheet1_settings.print_area,
        Some(vec![
            Range::from_a1("A1:A2").unwrap(),
            Range::from_a1("C1:C2").unwrap()
        ])
    );
    let sheet1_id = workbook.sheet_by_name("Sheet1").expect("Sheet1 missing").id;
    let print_area = workbook
        .get_defined_name(DefinedNameScope::Sheet(sheet1_id), XLNM_PRINT_AREA)
        .expect("missing Print_Area defined name");
    assert_parseable_formula(&print_area.refers_to);

    let sheet2_settings = workbook.sheet_print_settings_by_name("Sheet2");
    let titles = sheet2_settings
        .print_titles
        .expect("expected print_titles for Sheet2");
    assert_eq!(titles.repeat_rows, Some(RowRange { start: 0, end: 0 }));
    assert_eq!(titles.repeat_cols, Some(ColRange { start: 0, end: 0 }));
    let sheet2_id = workbook.sheet_by_name("Sheet2").expect("Sheet2 missing").id;
    let print_titles = workbook
        .get_defined_name(DefinedNameScope::Sheet(sheet2_id), XLNM_PRINT_TITLES)
        .expect("missing Print_Titles defined name");
    assert_parseable_formula(&print_titles.refers_to);
}

#[test]
fn imports_print_settings_via_calamine_defined_name_fallback_when_biff_unavailable() {
    let bytes = xls_fixture_builder::build_print_settings_calamine_fixture_xls();
    let result = import_fixture_without_biff(&bytes);
    let workbook = &result.workbook;

    let sheet1_settings = workbook.sheet_print_settings_by_name("Sheet1");
    assert_eq!(
        sheet1_settings.print_area,
        Some(vec![Range::from_a1("A1:A2").unwrap()])
    );
    let print_area = workbook
        .get_defined_name(DefinedNameScope::Workbook, XLNM_PRINT_AREA)
        .expect("missing Print_Area defined name");
    assert_parseable_formula(&print_area.refers_to);

    let sheet2_settings = workbook.sheet_print_settings_by_name("Sheet2");
    let titles = sheet2_settings.print_titles.unwrap_or_else(|| {
        panic!(
            "expected print_titles for Sheet2; defined_names={:?}; warnings={:?}",
            result.workbook.defined_names, result.warnings
        )
    });
    assert_eq!(titles.repeat_rows, Some(RowRange { start: 0, end: 0 }));
    assert_eq!(titles.repeat_cols, None);
    let print_titles = workbook
        .get_defined_name(DefinedNameScope::Workbook, XLNM_PRINT_TITLES)
        .expect("missing Print_Titles defined name");
    assert_parseable_formula(&print_titles.refers_to);
}

#[test]
fn imports_print_area_from_builtin_defined_name_with_unicode_quoted_sheet_name() {
    let bytes = xls_fixture_builder::build_print_settings_unicode_sheet_name_fixture_xls();
    let result = import_fixture(&bytes);
    let workbook = result.workbook;

    let settings = workbook.sheet_print_settings_by_name("Ünicode Name");
    assert_eq!(
        settings.print_area,
        Some(vec![
            Range::from_a1("A1:A2").unwrap(),
            Range::from_a1("C1:C2").unwrap()
        ])
    );

    let sheet_id = workbook
        .sheet_by_name("Ünicode Name")
        .expect("Ünicode Name sheet missing")
        .id;
    let print_area = workbook
        .get_defined_name(DefinedNameScope::Sheet(sheet_id), XLNM_PRINT_AREA)
        .expect("missing Print_Area defined name");
    assert_parseable_formula(&print_area.refers_to);
}

#[test]
fn imports_sheet_page_setup_and_manual_page_breaks_from_biff() {
    let bytes = xls_fixture_builder::build_sheet_print_settings_fixture_xls();
    let result = import_fixture(&bytes);
    let workbook = result.workbook;

    let settings = workbook.sheet_print_settings_by_name("Print");
    assert_eq!(settings.page_setup.orientation, Orientation::Landscape);
    assert_eq!(settings.page_setup.paper_size.code, 9);
    assert_eq!(
        settings.page_setup.scaling,
        Scaling::FitTo {
            width: 2,
            height: 3
        }
    );
    assert_eq!(settings.page_setup.margins.left, 1.1);
    assert_eq!(settings.page_setup.margins.right, 1.2);
    assert_eq!(settings.page_setup.margins.top, 1.3);
    assert_eq!(settings.page_setup.margins.bottom, 1.4);
    assert_eq!(settings.page_setup.margins.header, 0.5);
    assert_eq!(settings.page_setup.margins.footer, 0.6);

    assert_eq!(
        settings.manual_page_breaks.row_breaks_after,
        BTreeSet::from([2u32, 4u32])
    );
    assert_eq!(
        settings.manual_page_breaks.col_breaks_after,
        BTreeSet::from([1u32])
    );
}
