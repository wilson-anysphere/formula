use std::io::Write;

use formula_engine::{parse_formula, ParseOptions};
use formula_model::{DefinedNameScope, XLNM_FILTER_DATABASE, XLNM_PRINT_AREA, XLNM_PRINT_TITLES};

mod common;

use common::xls_fixture_builder;

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

fn assert_parseable(expr: &str) {
    parse_formula(&format!("={expr}"), ParseOptions::default())
        .unwrap_or_else(|e| panic!("expected expression to be parseable, expr={expr:?}, err={e:?}"));
}

#[test]
fn imports_biff8_builtin_defined_names_with_scope_and_hidden() {
    let bytes = xls_fixture_builder::build_defined_names_builtins_fixture_xls();
    let result = import_fixture(&bytes);
    let workbook = result.workbook;

    let sheet1_id = workbook
        .sheet_by_name("Sheet1")
        .expect("Sheet1 missing")
        .id;
    let sheet2_id = workbook
        .sheet_by_name("Sheet2")
        .expect("Sheet2 missing")
        .id;

    let print_area = workbook
        .get_defined_name(DefinedNameScope::Sheet(sheet1_id), XLNM_PRINT_AREA)
        .expect("missing Print_Area defined name");
    assert_eq!(print_area.name, XLNM_PRINT_AREA);
    assert_eq!(print_area.scope, DefinedNameScope::Sheet(sheet1_id));
    assert!(print_area.hidden, "Print_Area should be hidden");
    assert_eq!(print_area.xlsx_local_sheet_id, Some(0));
    assert_eq!(
        print_area.refers_to,
        "Sheet1!$A$1:$A$2,Sheet1!$C$1:$C$2"
    );
    assert_parseable(&print_area.refers_to);

    let print_titles = workbook
        .get_defined_name(DefinedNameScope::Sheet(sheet2_id), XLNM_PRINT_TITLES)
        .expect("missing Print_Titles defined name");
    assert_eq!(print_titles.name, XLNM_PRINT_TITLES);
    assert_eq!(print_titles.scope, DefinedNameScope::Sheet(sheet2_id));
    assert!(!print_titles.hidden, "Print_Titles should not be hidden");
    assert_eq!(print_titles.xlsx_local_sheet_id, Some(1));
    assert_eq!(print_titles.refers_to, "Sheet2!$1:$1,Sheet2!$A:$A");
    assert_parseable(&print_titles.refers_to);

    let filter_db = workbook
        .get_defined_name(DefinedNameScope::Sheet(sheet1_id), XLNM_FILTER_DATABASE)
        .expect("missing _FilterDatabase defined name");
    assert_eq!(filter_db.name, XLNM_FILTER_DATABASE);
    assert_eq!(filter_db.scope, DefinedNameScope::Sheet(sheet1_id));
    assert!(filter_db.hidden, "_FilterDatabase should be hidden");
    assert_eq!(filter_db.xlsx_local_sheet_id, Some(0));
    assert_eq!(filter_db.refers_to, "Sheet1!$A$1:$C$10");
    assert_parseable(&filter_db.refers_to);
}

#[test]
fn builtin_defined_name_prefers_rgchname_when_chkey_mismatched() {
    let bytes = xls_fixture_builder::build_defined_names_builtins_chkey_mismatch_fixture_xls();
    let result = import_fixture(&bytes);
    let workbook = result.workbook;

    let sheet1_id = workbook
        .sheet_by_name("Sheet1")
        .expect("Sheet1 missing")
        .id;
    let sheet2_id = workbook
        .sheet_by_name("Sheet2")
        .expect("Sheet2 missing")
        .id;

    // The fixture stores a mismatched built-in id in `NAME.chKey` for the Sheet1 Print_Area name.
    // Empirically Excel prefers the built-in id from `rgchName` and treats `chKey` as a keyboard
    // shortcut; the importer should do the same.
    let print_area = workbook.get_defined_name(DefinedNameScope::Sheet(sheet1_id), XLNM_PRINT_AREA);
    assert!(
        print_area.is_none(),
        "unexpected Print_Area defined name on Sheet1"
    );
    let print_titles_sheet1 =
        workbook.get_defined_name(DefinedNameScope::Sheet(sheet1_id), XLNM_PRINT_TITLES);
    assert!(
        print_titles_sheet1.is_some(),
        "missing Print_Titles defined name on Sheet1"
    );
    assert_parseable(&print_titles_sheet1.unwrap().refers_to);

    let print_titles_sheet2 =
        workbook.get_defined_name(DefinedNameScope::Sheet(sheet2_id), XLNM_PRINT_TITLES);
    assert!(
        print_titles_sheet2.is_some(),
        "missing Print_Titles defined name on Sheet2"
    );
    assert_parseable(&print_titles_sheet2.unwrap().refers_to);
}
