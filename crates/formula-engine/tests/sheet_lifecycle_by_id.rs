use formula_engine::{Engine, SheetLifecycleError};
use formula_model::SheetNameError;
use pretty_assertions::assert_eq;

#[test]
fn sheet_lifecycle_by_id_rename_reorder_delete() {
    let mut engine = Engine::new();
    // Use stable sheet keys that differ from display names so renames invalidate lookups by the
    // old display name (Excel semantics) without changing stable keys.
    engine.ensure_sheet("sheet1_key");
    engine.ensure_sheet("sheet2_key");
    engine.ensure_sheet("sheet3_key");
    engine.set_sheet_display_name("sheet1_key", "Sheet1");
    engine.set_sheet_display_name("sheet2_key", "Sheet2");
    engine.set_sheet_display_name("sheet3_key", "Sheet3");

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
    // Old display-name lookups should no longer resolve, but stable sheet keys remain usable.
    assert_eq!(engine.sheet_id("Sheet2"), None);
    assert_eq!(engine.sheet_id("sheet2_key"), Some(sheet2_id));

    // Delete by id should tombstone the sheet and invalidate lookups by id and name.
    engine.delete_sheet_by_id(sheet1_id).unwrap();
    assert_eq!(engine.sheet_name(sheet1_id), None);
    assert_eq!(engine.sheet_id("Sheet1"), None);
    assert!(!engine.sheet_ids_in_order().contains(&sheet1_id));

    // Invalid / tombstoned ids should be handled gracefully (no-op, workbook unchanged).
    let order_before = engine.sheet_ids_in_order();
    engine.rename_sheet_by_id(sheet1_id, "DoesNotExist").unwrap();
    engine.reorder_sheet_by_id(sheet1_id, 0).unwrap();
    engine.delete_sheet_by_id(sheet1_id).unwrap();
    assert_eq!(engine.sheet_ids_in_order(), order_before);
    assert_eq!(engine.sheet_name(sheet1_id), None);

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

#[test]
fn sheet_lifecycle_by_id_rewrites_formulas_and_preserves_external_refs() {
    let mut engine = Engine::new();
    engine.ensure_sheet("Sheet1");
    engine.ensure_sheet("Sheet2");
    engine.ensure_sheet("Sheet3");

    // Make sure we have a stored formula that references Sheet1 both locally and via an external
    // workbook reference. Local references should be rewritten on rename/delete, but external
    // workbook refs must remain intact.
    engine
        .set_cell_formula("Sheet2", "A1", "=[Book.xlsx]Sheet1!A1+Sheet1!A1")
        .unwrap();
    // Include a 3D sheet span so deletion must shift boundaries inward.
    engine
        .set_cell_formula("Sheet2", "B1", "=SUM(Sheet1:Sheet3!A1)")
        .unwrap();

    let sheet1_id = engine.sheet_id("Sheet1").unwrap();
    let sheet2_id = engine.sheet_id("Sheet2").unwrap();
    let sheet3_id = engine.sheet_id("Sheet3").unwrap();

    engine.rename_sheet_by_id(sheet1_id, "Renamed").unwrap();
    assert_eq!(engine.sheet_name(sheet1_id), Some("Renamed"));
    assert_eq!(engine.sheet_id("Renamed"), Some(sheet1_id));
    assert_eq!(engine.sheet_id("Sheet1"), None);

    assert_eq!(
        engine.get_cell_formula("Sheet2", "A1"),
        Some("=[Book.xlsx]Sheet1!A1+Renamed!A1")
    );
    assert_eq!(
        engine.get_cell_formula("Sheet2", "B1"),
        Some("=SUM(Renamed:Sheet3!A1)")
    );

    engine.delete_sheet_by_id(sheet1_id).unwrap();
    assert_eq!(engine.sheet_name(sheet1_id), None);
    assert_eq!(engine.sheet_id("Renamed"), None);
    assert!(!engine.sheet_ids_in_order().contains(&sheet1_id));

    assert_eq!(
        engine.get_cell_formula("Sheet2", "A1"),
        Some("=[Book.xlsx]Sheet1!A1+#REF!")
    );
    assert_eq!(
        engine.get_cell_formula("Sheet2", "B1"),
        Some("=SUM(Sheet2:Sheet3!A1)")
    );

    // Reorder by id (remaining sheets).
    engine.reorder_sheet_by_id(sheet3_id, 0).unwrap();
    assert_eq!(engine.sheet_ids_in_order(), vec![sheet3_id, sheet2_id]);
    assert_eq!(engine.sheet_id("Sheet2"), Some(sheet2_id));
    assert_eq!(engine.sheet_id("Sheet3"), Some(sheet3_id));

    // Out-of-range ids are handled gracefully.
    let order_before = engine.sheet_ids_in_order();
    engine.rename_sheet_by_id(123_456, "Nope").unwrap();
    engine.reorder_sheet_by_id(123_456, 0).unwrap();
    engine.delete_sheet_by_id(123_456).unwrap();
    assert_eq!(engine.sheet_ids_in_order(), order_before);
}
