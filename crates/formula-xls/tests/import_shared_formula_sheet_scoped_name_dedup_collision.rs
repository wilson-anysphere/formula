use std::io::Write;

use formula_model::{CellRef, CellValue};

mod common;

use common::{assert_parseable_formula, xls_fixture_builder};

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

fn quote_sheet_name_if_needed(name: &str) -> String {
    // Mirror the minimal quoting behavior used by the BIFF rgce decoder: quote when the name
    // contains characters that require quoting in an Excel sheet reference.
    //
    // For this test we only need to handle the importer dedupe form (`Bad_Name (2)`), which
    // requires quoting due to the space.
    if name.contains(' ') {
        format!("'{name}'")
    } else {
        name.to_string()
    }
}

fn assert_decoded_formula_resolves_to_scoped_sheet(bytes: &[u8]) {
    let result = import_fixture(bytes);

    assert!(
        result.workbook.sheet_by_name("Bad:Name").is_none(),
        "expected invalid sheet name to be sanitized"
    );

    // Find the final sheet name for the invalid source sheet by looking for the unique marker
    // value (A1=111) inserted by the fixture.
    let marker_cell = CellRef::from_a1("A1").unwrap();
    let mut invalid_sheet_name: Option<String> = None;
    let mut collision_sheet_name: Option<String> = None;
    for sheet in &result.workbook.sheets {
        match sheet.value(marker_cell) {
            CellValue::Number(n) if (n - 111.0).abs() < f64::EPSILON => {
                invalid_sheet_name = Some(sheet.name.clone());
            }
            CellValue::Number(n) if (n - 222.0).abs() < f64::EPSILON => {
                collision_sheet_name = Some(sheet.name.clone());
            }
            _ => {}
        }
    }
    let invalid_sheet_name = invalid_sheet_name.expect("expected marker sheet (111) to exist");
    let collision_sheet_name =
        collision_sheet_name.expect("expected collision sheet (222) to exist");

    let sheet = result.workbook.sheet_by_name("Ref").expect("Ref missing");
    let formula = sheet
        .formula(CellRef::from_a1("A2").unwrap())
        .expect("expected formula in Ref!A2");

    let (sheet_prefix, name_ref) = formula
        .split_once('!')
        .expect("expected sheet-qualified name reference");
    assert_eq!(name_ref, "LocalName");

    let invalid_prefix = quote_sheet_name_if_needed(&invalid_sheet_name);
    assert_eq!(
        sheet_prefix, invalid_prefix,
        "expected formula to reference the sheet scoped by NAME.itab, got {formula:?}"
    );
    let collision_prefix = quote_sheet_name_if_needed(&collision_sheet_name);
    assert_ne!(
        sheet_prefix, collision_prefix,
        "expected formula not to reference the colliding sheet, got {formula:?}"
    );

    let expected = format!(
        "{}!LocalName",
        quote_sheet_name_if_needed(&invalid_sheet_name)
    );
    assert_eq!(formula, expected);
    assert_parseable_formula(formula);
}

#[test]
fn decodes_sheet_scoped_ptgname_when_sanitization_collides_and_sheet_is_deduped() {
    let bytes =
        xls_fixture_builder::build_shared_formula_sheet_scoped_name_dedup_collision_fixture_xls();
    assert_decoded_formula_resolves_to_scoped_sheet(&bytes);
}

#[test]
fn decodes_sheet_scoped_ptgname_when_invalid_sheet_sanitization_is_deduped() {
    let bytes = xls_fixture_builder::
        build_shared_formula_sheet_scoped_name_dedup_collision_invalid_second_fixture_xls();
    assert_decoded_formula_resolves_to_scoped_sheet(&bytes);
}
