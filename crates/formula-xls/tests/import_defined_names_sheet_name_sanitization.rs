use std::io::Write;

use formula_engine::{parse_formula, ParseOptions};
use formula_model::DefinedNameScope;

mod common;

use common::xls_fixture_builder;

fn assert_parseable(expr: &str) {
    let expr = expr.trim();
    assert!(!expr.is_empty(), "expected expression to be non-empty");
    parse_formula(expr, ParseOptions::default())
        .unwrap_or_else(|e| panic!("expected expression to be parseable, expr={expr:?}, err={e:?}"));
}

#[test]
fn rewrites_defined_name_sheet_refs_to_sanitized_sheet_names() {
    let bytes = xls_fixture_builder::build_sanitized_sheet_name_defined_name_fixture_xls();
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(&bytes).expect("write xls bytes");

    let result = formula_xls::import_xls_path(tmp.path()).expect("import xls");

    assert!(
        result.workbook.sheet_by_name("Bad/Name").is_none(),
        "expected invalid sheet name to be sanitized"
    );
    assert!(
        result.workbook.sheet_by_name("Bad_Name").is_some(),
        "expected sanitized sheet to exist"
    );

    let dn = result
        .workbook
        .defined_names
        .iter()
        .find(|n| n.name == "MyRange")
        .expect("MyRange missing");
    assert_eq!(dn.scope, DefinedNameScope::Workbook);
    assert_eq!(dn.refers_to, "Bad_Name!$A$1");
    assert_parseable(&dn.refers_to);
}

#[test]
fn does_not_cascade_defined_name_rewrites_when_sanitization_collides_with_another_sheet_name() {
    let bytes = xls_fixture_builder::build_sanitized_sheet_name_defined_name_collision_fixture_xls();
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(&bytes).expect("write xls bytes");

    let result = formula_xls::import_xls_path(tmp.path()).expect("import xls");

    // Sheet 0: invalid -> sanitized base name.
    assert!(result.workbook.sheet_by_name("Bad_Name").is_some());
    // Sheet 1: original name collides with sanitized base name and is deduped.
    assert!(result.workbook.sheet_by_name("Bad_Name (2)").is_some());

    let dn = result
        .workbook
        .defined_names
        .iter()
        .find(|n| n.name == "MyRange")
        .expect("MyRange missing");

    // The defined name refers to sheet 0 (invalid `Bad/Name`), which sanitizes to `Bad_Name`.
    // Ensure rewriting did not mistakenly redirect it to `Bad_Name (2)`.
    assert_eq!(dn.refers_to, "Bad_Name!$A$1");
    assert_parseable(&dn.refers_to);
}
