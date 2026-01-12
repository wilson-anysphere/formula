use std::io::Write;

use calamine::{open_workbook, Reader, Xls};
use formula_model::{DefinedNameScope, XLNM_PRINT_AREA};

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

#[test]
fn imports_biff_defined_names_with_scope_and_3d_refs() {
    let bytes = xls_fixture_builder::build_defined_names_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet1_id = result
        .workbook
        .sheet_by_name("Sheet1")
        .expect("Sheet1 missing")
        .id;

    let zed = result
        .workbook
        .defined_names
        .iter()
        .find(|n| n.name == "ZedName")
        .expect("ZedName missing");
    assert_eq!(zed.scope, DefinedNameScope::Workbook);
    assert!(!zed.hidden);
    assert_eq!(zed.refers_to, "Sheet1!$B$1");
    assert_eq!(zed.xlsx_local_sheet_id, None);

    let local = result
        .workbook
        .defined_names
        .iter()
        .find(|n| n.name == "LocalName")
        .expect("LocalName missing");
    assert_eq!(local.scope, DefinedNameScope::Sheet(sheet1_id));
    assert!(!local.hidden);
    assert_eq!(local.refers_to, "Sheet1!$A$1");
    assert_eq!(local.comment.as_deref(), Some("Local description"));
    assert_eq!(local.xlsx_local_sheet_id, Some(0));

    let hidden = result
        .workbook
        .defined_names
        .iter()
        .find(|n| n.name == "HiddenName")
        .expect("HiddenName missing");
    assert_eq!(hidden.scope, DefinedNameScope::Workbook);
    assert!(hidden.hidden);
    assert_eq!(hidden.refers_to, "Sheet1!$A$1:$B$2");

    let union = result
        .workbook
        .defined_names
        .iter()
        .find(|n| n.name == "UnionName")
        .expect("UnionName missing");
    assert_eq!(union.scope, DefinedNameScope::Workbook);
    assert_eq!(union.refers_to, "Sheet1!$A$1,Sheet1!$B$1");

    let my_name = result
        .workbook
        .defined_names
        .iter()
        .find(|n| n.name == "MyName")
        .expect("MyName missing");
    assert_eq!(my_name.scope, DefinedNameScope::Workbook);
    assert!(!my_name.hidden);
    assert_eq!(my_name.refers_to, "SUM(Sheet1!$A$1:$A$3)");

    let abs = result
        .workbook
        .defined_names
        .iter()
        .find(|n| n.name == "AbsName")
        .expect("AbsName missing");
    assert_eq!(abs.scope, DefinedNameScope::Workbook);
    assert_eq!(abs.refers_to, "ABS(1)");

    let union_func = result
        .workbook
        .defined_names
        .iter()
        .find(|n| n.name == "UnionFunc")
        .expect("UnionFunc missing");
    assert_eq!(union_func.scope, DefinedNameScope::Workbook);
    assert_eq!(union_func.refers_to, "SUM((Sheet1!$A$1,Sheet1!$B$1))");

    let miss = result
        .workbook
        .defined_names
        .iter()
        .find(|n| n.name == "MissingArgName")
        .expect("MissingArgName missing");
    assert_eq!(miss.scope, DefinedNameScope::Workbook);
    assert_eq!(miss.refers_to, "IF(,1,2)");

    let print_area = result
        .workbook
        .defined_names
        .iter()
        .find(|n| n.name == XLNM_PRINT_AREA)
        .expect("Print_Area missing");
    assert_eq!(print_area.scope, DefinedNameScope::Sheet(sheet1_id));
    assert!(print_area.hidden);
    assert_eq!(
        print_area.refers_to,
        "Sheet1!$A$1:$B$2,Sheet1!$D$4:$E$5"
    );
    assert_eq!(print_area.xlsx_local_sheet_id, Some(0));
}

#[test]
fn defined_name_formulas_quote_sheet_names() {
    let bytes = xls_fixture_builder::build_defined_names_quoting_fixture_xls();
    let result = import_fixture(&bytes);

    let cases = [
        ("SpaceRef", "'Sheet One'!$A$1"),
        ("QuoteRef", "'O''Brien'!$B$2"),
        ("ReservedRef", "'TRUE'!$C$3"),
        ("SpanRef", "'Sheet One:O''Brien'!$D$4"),
    ];

    for (name, expected_refers_to) in cases {
        let dn = result
            .workbook
            .defined_names
            .iter()
            .find(|n| n.name == name)
            .unwrap_or_else(|| panic!("{name} missing"));
        assert_eq!(dn.refers_to, expected_refers_to);
    }
}

#[test]
fn imports_workbook_defined_names_via_calamine_fallback_when_biff_unavailable() {
    let bytes = xls_fixture_builder::build_defined_name_calamine_fixture_xls();
    let result = import_fixture_without_biff(&bytes);

    let name = result
        .workbook
        .defined_names
        .iter()
        .find(|n| n.name == "TestName")
        .unwrap_or_else(|| {
            panic!(
                "TestName missing; defined_names={:?}; warnings={:?}",
                result.workbook.defined_names, result.warnings
            )
        });
    assert_eq!(name.scope, DefinedNameScope::Workbook);
    assert_eq!(name.refers_to, "Sheet1!$A$1:$A$1");
}

#[test]
fn rewrites_calamine_defined_name_formulas_to_sanitized_sheet_names() {
    let bytes = xls_fixture_builder::build_defined_name_sheet_name_sanitization_fixture_xls();
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(&bytes).expect("write xls bytes");

    // Sanity check: calamine should surface the original invalid sheet name in the defined name.
    let calamine_refers_to = {
        let wb: Xls<_> = open_workbook(tmp.path()).expect("open xls fixture via calamine");
        wb.defined_names()
            .iter()
            .find(|(name, _)| name.replace('\0', "") == "TestName")
            .map(|(_, refers_to)| refers_to.as_str())
            .unwrap_or("<missing>")
            .to_string()
    };
    assert!(
        calamine_refers_to.contains("Bad:Name"),
        "expected calamine refers_to to reference original sheet name; refers_to={calamine_refers_to:?}"
    );

    let result = formula_xls::import_xls_path_without_biff(tmp.path()).expect("import xls");

    assert!(
        result.workbook.sheet_by_name("Bad:Name").is_none(),
        "expected invalid sheet name to be sanitized"
    );
    assert!(
        result.workbook.sheet_by_name("Bad_Name").is_some(),
        "expected sanitized sheet to be present"
    );

    let name = result
        .workbook
        .defined_names
        .iter()
        .find(|n| n.name == "TestName")
        .unwrap_or_else(|| {
            panic!(
                "TestName missing; defined_names={:?}; warnings={:?}",
                result.workbook.defined_names, result.warnings
            )
        });
    assert_eq!(name.scope, DefinedNameScope::Workbook);
    assert_eq!(name.refers_to, "Bad_Name!$A$1:$A$1");

    assert!(
        result
            .warnings
            .iter()
            .any(|w| w.message.contains("sanitized sheet name `Bad:Name` -> `Bad_Name`")),
        "expected import warnings to mention sheet name sanitization; warnings={:?}",
        result.warnings
    );
}
