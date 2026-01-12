use std::io::Write;

use calamine::{open_workbook, Reader, Xls};
use formula_engine::{parse_formula, ParseOptions};
use formula_model::{validate_defined_name, DefinedNameScope};

mod common;

use common::xls_fixture_builder;

fn assert_parseable(expr: &str) {
    let expr = expr.trim();
    assert!(!expr.is_empty(), "expected expression to be non-empty");
    parse_formula(expr, ParseOptions::default())
        .unwrap_or_else(|e| panic!("expected expression to be parseable, expr={expr:?}, err={e:?}"));
}

#[test]
fn imports_defined_names_via_calamine_fallback_when_biff_is_unavailable() {
    let bytes = xls_fixture_builder::build_defined_names_fixture_xls();
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(&bytes).expect("write xls bytes");

    // Sanity check: calamine should see at least one defined name in the fixture.
    let workbook_defined_names = {
        let wb: Xls<_> = open_workbook(tmp.path()).expect("open xls fixture via calamine");
        wb.defined_names().to_vec()
    };
    assert!(
        !workbook_defined_names.is_empty(),
        "expected fixture to contain at least one defined name"
    );

    // Pick a defined name that the importer would accept (some BIFF fixtures may include invalid
    // names like `A1`, which the model rejects).
    let (expected_name, expected_refers_to) = workbook_defined_names
        .iter()
        .map(|(name, refers_to)| {
            let name = name.replace('\0', "");
            let refers_to = refers_to.trim();
            let refers_to = refers_to.strip_prefix('=').unwrap_or(refers_to).to_string();
            (name, refers_to)
        })
        .find(|(name, _)| validate_defined_name(name).is_ok())
        .expect("expected at least one valid defined name from calamine");
    assert_parseable(&expected_refers_to);

    // Force BIFF workbook-stream parsing to be unavailable so the importer has to use the
    // calamine fallback path.
    let result = formula_xls::import_xls_path_without_biff(tmp.path()).expect("import xls");

    let imported = result
        .workbook
        .defined_names
        .iter()
        .find(|n| {
            n.scope == DefinedNameScope::Workbook
                && n.name == expected_name
                && n.refers_to == expected_refers_to
                && !n.hidden
                && n.comment.is_none()
                && n.xlsx_local_sheet_id.is_none()
        })
        .unwrap_or_else(|| {
            panic!(
                "expected imported workbook to contain defined name {expected_name:?}; defined_names={:?}",
                result.workbook.defined_names
            )
        });
    assert_parseable(&imported.refers_to);

    assert!(
        result.warnings.iter().any(|w| w
            .message
            .contains("imported `.xls` defined names via calamine fallback")),
        "missing calamine fallback warning; warnings={:?}",
        result.warnings
    );
}
