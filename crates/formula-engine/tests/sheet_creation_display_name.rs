use formula_engine::{Engine, SheetLifecycleError, Value};
use formula_model::{validate_sheet_name, SheetNameError, EXCEL_MAX_SHEET_NAME_LEN};

#[test]
fn long_stable_sheet_key_gets_valid_default_display_name_and_remains_addressable() {
    let mut engine = Engine::new();
    let sheet_key = "550e8400-e29b-41d4-a716-446655440000-this-is-way-too-long-for-excel";

    // Creating/using the sheet by its stable key should always work.
    engine.set_cell_value(sheet_key, "A1", 123.0).unwrap();
    assert_eq!(engine.get_cell_value(sheet_key, "A1"), Value::Number(123.0));

    // The generated display name must be Excel-valid.
    let sheet_id = engine.sheet_id(sheet_key).expect("sheet id exists");
    let display_name = engine.sheet_name(sheet_id).expect("sheet name exists");
    validate_sheet_name(display_name).expect("display name must be Excel-valid");
    assert!(
        display_name.encode_utf16().count() <= EXCEL_MAX_SHEET_NAME_LEN,
        "display name too long: {display_name:?}"
    );
}

#[test]
fn sheet_key_with_invalid_excel_chars_does_not_leak_into_display_name() {
    let mut engine = Engine::new();
    let sheet_key = "collab:sheet:uuid:with:colons";

    engine.set_cell_value(sheet_key, "A1", 1.0).unwrap();
    assert_eq!(engine.get_cell_value(sheet_key, "A1"), Value::Number(1.0));

    let sheet_id = engine.sheet_id(sheet_key).expect("sheet id exists");
    let display_name = engine.sheet_name(sheet_id).expect("sheet name exists");
    validate_sheet_name(display_name).expect("display name must be Excel-valid");
    assert!(
        !display_name.contains(':'),
        "display name must not contain ':'; got {display_name:?}"
    );
}

#[test]
fn ensure_sheet_with_display_name_validates_display_name() {
    let mut engine = Engine::new();

    // Invalid Excel sheet name (contains ':') should be rejected and must not create a sheet.
    let err = engine
        .ensure_sheet_with_display_name("sheet-key-1", "Bad:Name")
        .unwrap_err();
    assert_eq!(
        err,
        SheetLifecycleError::InvalidName(SheetNameError::InvalidCharacter(':'))
    );
    assert!(engine.sheet_id("sheet-key-1").is_none());

    // Valid display names should succeed.
    let sheet_id = engine
        .ensure_sheet_with_display_name("sheet-key-2", "GoodName")
        .unwrap();
    assert_eq!(engine.sheet_id("sheet-key-2"), Some(sheet_id));
    assert_eq!(engine.sheet_name(sheet_id), Some("GoodName"));
}
