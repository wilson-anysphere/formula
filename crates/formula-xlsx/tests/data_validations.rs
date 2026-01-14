use std::path::Path;

use formula_model::{DataValidationKind, Range};

fn fixture_path(rel: &str) -> std::path::PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../")
        .join(rel)
}

#[test]
fn reads_data_validation_list_rules_into_workbook_model() {
    let bytes = std::fs::read(fixture_path(
        "fixtures/xlsx/metadata/data-validation-list.xlsx",
    ))
    .expect("read fixture");
    let workbook = formula_xlsx::read_workbook_model_from_bytes(&bytes).expect("read workbook");

    assert_eq!(workbook.sheets.len(), 1);
    let sheet = &workbook.sheets[0];

    assert_eq!(sheet.data_validations.len(), 1);
    let assignment = &sheet.data_validations[0];

    assert_eq!(assignment.ranges, vec![Range::from_a1("A1").unwrap()]);
    assert_eq!(assignment.validation.kind, DataValidationKind::List);
    assert_eq!(assignment.validation.allow_blank, true);
    assert_eq!(assignment.validation.show_input_message, true);
    assert_eq!(assignment.validation.show_error_message, true);
    assert_eq!(assignment.validation.formula1, "\"Yes,No\"");
}
