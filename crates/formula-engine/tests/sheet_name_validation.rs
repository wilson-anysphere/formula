use formula_engine::Engine;
use formula_model::EXCEL_MAX_SHEET_NAME_LEN;

fn assert_rename_rejected_without_side_effects(new_name: &str) {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine
        .set_cell_formula("Sheet2", "A1", "=Sheet1!A1")
        .unwrap();

    let sheet_id = engine.sheet_id("Sheet1").unwrap();
    let old_sheet_name = engine.sheet_name(sheet_id).unwrap().to_string();
    let formula_before = engine.get_cell_formula("Sheet2", "A1").unwrap().to_string();

    assert!(
        !engine.rename_sheet("Sheet1", new_name),
        "expected rename to {new_name:?} to fail"
    );

    // Sheet name/id mappings should remain unchanged.
    assert_eq!(engine.sheet_name(sheet_id), Some(old_sheet_name.as_str()));
    assert_eq!(engine.sheet_id("Sheet1"), Some(sheet_id));
    assert_eq!(engine.sheet_id(new_name), None);

    // Stored formulas should not be rewritten on failure.
    assert_eq!(
        engine.get_cell_formula("Sheet2", "A1"),
        Some(formula_before.as_str())
    );
}

#[test]
fn rename_sheet_rejects_empty_name() {
    assert_rename_rejected_without_side_effects("");
}

#[test]
fn rename_sheet_rejects_whitespace_only_name() {
    assert_rename_rejected_without_side_effects("   ");
}

#[test]
fn rename_sheet_rejects_colon_in_name() {
    // Important for 3D span parsing ambiguity like `Sheet1:Sheet3!A1`.
    assert_rename_rejected_without_side_effects("Bad:Name");
}

#[test]
fn rename_sheet_rejects_brackets_in_name() {
    assert_rename_rejected_without_side_effects("Bad[Name]");
}

#[test]
fn rename_sheet_rejects_leading_or_trailing_apostrophe() {
    assert_rename_rejected_without_side_effects("'Bad");
    assert_rename_rejected_without_side_effects("Bad'");
}

#[test]
fn rename_sheet_rejects_names_longer_than_31_utf16_units() {
    let too_long = "a".repeat(EXCEL_MAX_SHEET_NAME_LEN + 1);
    assert_rename_rejected_without_side_effects(&too_long);
}
