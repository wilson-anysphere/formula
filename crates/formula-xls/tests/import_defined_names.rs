use std::io::Write;

use calamine::{open_workbook, Reader, Xls};
use formula_engine::{parse_formula, ParseOptions};
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

fn assert_parseable_refers_to(expr: &str) {
    parse_formula(&format!("={expr}"), ParseOptions::default())
        .unwrap_or_else(|e| panic!("expected refers_to to be parseable, expr={expr:?}, err={e:?}"));
}

#[test]
fn imports_defined_names_from_biff_name_records() {
    let bytes = xls_fixture_builder::build_defined_names_fixture_xls();
    let result = import_fixture(&bytes);

    let wb = &result.workbook;
    let sheet1_id = wb.sheet_by_name("Sheet1").expect("Sheet1 missing").id;

    let global = wb
        .get_defined_name(DefinedNameScope::Workbook, "GlobalName")
        .expect("GlobalName missing");
    assert_eq!(global.hidden, false);
    assert_eq!(global.refers_to, "Sheet1!$A$1");
    assert_eq!(global.xlsx_local_sheet_id, None);

    let zed = wb
        .get_defined_name(DefinedNameScope::Workbook, "ZedName")
        .expect("ZedName missing");
    assert_eq!(zed.hidden, false);
    assert_eq!(zed.refers_to, "Sheet1!$B$1");
    assert_eq!(zed.xlsx_local_sheet_id, None);
    assert_parseable_refers_to(&zed.refers_to);

    let local = wb
        .get_defined_name(DefinedNameScope::Sheet(sheet1_id), "LocalName")
        .expect("LocalName missing");
    assert_eq!(local.hidden, false);
    assert_eq!(local.refers_to, "Sheet1!$A$1");
    assert_eq!(local.comment.as_deref(), Some("Local description"));
    assert_eq!(local.xlsx_local_sheet_id, Some(0));
    assert_parseable_refers_to(&local.refers_to);

    let hidden = wb
        .get_defined_name(DefinedNameScope::Workbook, "HiddenName")
        .expect("HiddenName missing");
    assert!(hidden.hidden);
    assert_eq!(hidden.refers_to, "Sheet1!$A$1:$B$2");
    assert_parseable_refers_to(&hidden.refers_to);

    let union = wb
        .get_defined_name(DefinedNameScope::Workbook, "UnionName")
        .expect("UnionName missing");
    assert_eq!(union.refers_to, "Sheet1!$A$1,Sheet1!$B$1");
    assert_parseable_refers_to(&union.refers_to);

    let my_name = wb
        .get_defined_name(DefinedNameScope::Workbook, "MyName")
        .expect("MyName missing");
    assert!(!my_name.hidden);
    assert_eq!(my_name.refers_to, "SUM(Sheet1!$A$1:$A$3)");
    assert_parseable_refers_to(&my_name.refers_to);

    let abs = wb
        .get_defined_name(DefinedNameScope::Workbook, "AbsName")
        .expect("AbsName missing");
    assert_eq!(abs.refers_to, "ABS(1)");
    assert_parseable_refers_to(&abs.refers_to);

    let union_func = wb
        .get_defined_name(DefinedNameScope::Workbook, "UnionFunc")
        .expect("UnionFunc missing");
    assert_eq!(union_func.refers_to, "SUM((Sheet1!$A$1,Sheet1!$B$1))");
    assert_parseable_refers_to(&union_func.refers_to);

    let miss = wb
        .get_defined_name(DefinedNameScope::Workbook, "MissingArgName")
        .expect("MissingArgName missing");
    assert_eq!(miss.refers_to, "IF(,1,2)");
    assert_parseable_refers_to(&miss.refers_to);

    let print_area = wb
        .get_defined_name(DefinedNameScope::Sheet(sheet1_id), XLNM_PRINT_AREA)
        .expect("Print_Area missing");
    assert!(print_area.hidden, "expected Print_Area to be hidden");
    assert_eq!(print_area.xlsx_local_sheet_id, Some(0));
    assert_eq!(
        print_area.refers_to,
        "Sheet1!$A$1:$B$2,Sheet1!$D$4:$E$5"
    );
    assert_parseable_refers_to(&print_area.refers_to);

    // All imported name formulas should be parseable by our formula parser (stored without `=`).
    for name in &wb.defined_names {
        let f = format!("={}", name.refers_to);
        parse_formula(&f, ParseOptions::default()).unwrap_or_else(|err| {
            panic!(
                "failed to parse refers_to for defined name `{}` (scope={:?}): {err}; formula={f}",
                name.name, name.scope
            )
        });
    }
}

#[test]
fn defined_name_formulas_quote_sheet_names_and_are_parseable() {
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
        assert_parseable_refers_to(&dn.refers_to);
    }

    for name in &result.workbook.defined_names {
        let f = format!("={}", name.refers_to);
        parse_formula(&f, ParseOptions::default()).unwrap_or_else(|err| {
            panic!(
                "failed to parse refers_to for defined name `{}` (scope={:?}): {err}; formula={f}",
                name.name, name.scope
            )
        });
    }
}

#[test]
fn imports_defined_names_with_external_workbook_3d_refs() {
    let bytes = xls_fixture_builder::build_defined_names_external_workbook_refs_fixture_xls();
    let result = import_fixture(&bytes);

    let ext_single = result
        .workbook
        .defined_names
        .iter()
        .find(|n| n.name == "ExtSingle")
        .expect("ExtSingle missing");
    assert_eq!(ext_single.refers_to, "'[Book1.xlsx]SheetA'!$A$1");

    let ext_span = result
        .workbook
        .defined_names
        .iter()
        .find(|n| n.name == "ExtSpan")
        .expect("ExtSpan missing");
    assert_eq!(ext_span.refers_to, "'[Book1.xlsx]SheetA:SheetC'!$A$1");
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
    assert_parseable_refers_to(&name.refers_to);
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
    assert_parseable_refers_to(&name.refers_to);

    assert!(
        result
            .warnings
            .iter()
            .any(|w| w.message.contains("sanitized sheet name `Bad:Name` -> `Bad_Name`")),
        "expected import warnings to mention sheet name sanitization; warnings={:?}",
        result.warnings
    );
}
