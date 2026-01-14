use formula_engine::{Engine, SheetLifecycleError};
use formula_model::SheetNameError;

#[test]
fn sheet_lifecycle_by_id_rename_reorder_delete() {
    let mut engine = Engine::new();
    engine.ensure_sheet("Sheet1");
    engine.ensure_sheet("Sheet2");
    engine.ensure_sheet("Sheet3");

    let sheet1_id = engine.sheet_id("Sheet1").expect("Sheet1 id");
    let sheet2_id = engine.sheet_id("Sheet2").expect("Sheet2 id");
    let sheet3_id = engine.sheet_id("Sheet3").expect("Sheet3 id");

    // Tab order should be based on the explicit tab ordering vector, not stable ids.
    assert_eq!(engine.sheet_ids_in_order(), vec![sheet1_id, sheet2_id, sheet3_id]);

    // Reorder by id should update tab order but keep ids stable.
    engine.reorder_sheet_by_id(sheet3_id, 0).unwrap();
    assert_eq!(engine.sheet_ids_in_order(), vec![sheet3_id, sheet1_id, sheet2_id]);
    assert_eq!(engine.sheet_id("Sheet3"), Some(sheet3_id));

    // Rename by id should update the name mapping for future lookups.
    engine.rename_sheet_by_id(sheet2_id, "Renamed").unwrap();
    assert_eq!(engine.sheet_id("Renamed"), Some(sheet2_id));
    assert_eq!(engine.sheet_name(sheet2_id), Some("Renamed"));
    assert_eq!(engine.sheet_id("Sheet2"), None);

    // Delete by id should tombstone the sheet and invalidate lookups by id and name.
    engine.delete_sheet_by_id(sheet1_id).unwrap();
    assert_eq!(engine.sheet_name(sheet1_id), None);
    assert_eq!(engine.sheet_id("Sheet1"), None);
    assert!(!engine.sheet_ids_in_order().contains(&sheet1_id));

    assert_eq!(
        engine.rename_sheet_by_id(sheet1_id, "DoesNotExist").unwrap_err(),
        SheetLifecycleError::SheetNotFound
    );
    assert_eq!(
        engine.reorder_sheet_by_id(sheet1_id, 0).unwrap_err(),
        SheetLifecycleError::SheetNotFound
    );
    assert_eq!(
        engine.delete_sheet_by_id(sheet1_id).unwrap_err(),
        SheetLifecycleError::SheetNotFound
    );

    // Re-creating a sheet with the same name should not reuse the old id.
    engine.ensure_sheet("Sheet1");
    let recreated_id = engine.sheet_id("Sheet1").expect("recreated sheet id");
    assert_ne!(recreated_id, sheet1_id);
}

#[test]
fn sheet_delete_by_id_cannot_delete_last_sheet() {
    let mut engine = Engine::new();
    engine.ensure_sheet("Sheet1");
    let sheet1_id = engine.sheet_id("Sheet1").expect("sheet id");

    assert_eq!(
        engine.delete_sheet_by_id(sheet1_id).unwrap_err(),
        SheetLifecycleError::CannotDeleteLastSheet
    );
}

#[test]
fn sheet_lifecycle_by_id_validates_names_and_indices() {
    let mut engine = Engine::new();
    engine.ensure_sheet("Sheet1");
    engine.ensure_sheet("Sheet2");

    let sheet1_id = engine.sheet_id("Sheet1").expect("Sheet1 id");
    let sheet2_id = engine.sheet_id("Sheet2").expect("Sheet2 id");

    assert_eq!(
        engine.rename_sheet_by_id(sheet2_id, "Sheet1").unwrap_err(),
        SheetLifecycleError::InvalidName(SheetNameError::DuplicateName)
    );
    assert_eq!(
        engine.rename_sheet_by_id(sheet2_id, "").unwrap_err(),
        SheetLifecycleError::InvalidName(SheetNameError::EmptyName)
    );
    assert_eq!(
        engine.rename_sheet_by_id(sheet2_id, "Bad:Name").unwrap_err(),
        SheetLifecycleError::InvalidName(SheetNameError::InvalidCharacter(':'))
    );

    assert_eq!(
        engine.reorder_sheet_by_id(sheet1_id, 2).unwrap_err(),
        SheetLifecycleError::IndexOutOfRange
    );
}
