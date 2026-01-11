use formula_model::{CellRef, Range, SheetSelection, Workbook, Worksheet};

#[test]
fn workbook_active_sheet_and_selection_roundtrip() {
    let mut wb = Workbook::new();
    let sheet1 = wb.add_sheet("Sheet1").unwrap();
    let sheet2 = wb.add_sheet("Sheet2").unwrap();

    assert_eq!(wb.active_sheet_id(), Some(sheet1));
    assert!(wb.set_active_sheet(sheet2));

    let selection = SheetSelection::new(
        CellRef::new(1, 1), // B2
        vec![
            Range::from_a1("B2:C3").unwrap(),
            Range::from_a1("E5").unwrap(),
        ],
    );
    wb.sheet_mut(sheet2)
        .unwrap()
        .set_selection(selection.clone());

    let json = serde_json::to_value(&wb).unwrap();
    let roundtrip: Workbook = serde_json::from_value(json).unwrap();

    assert_eq!(roundtrip.active_sheet_id(), Some(sheet2));
    let sheet = roundtrip.sheet(sheet2).unwrap();
    assert_eq!(sheet.selection(), Some(&selection));
}

#[test]
fn legacy_freeze_and_zoom_migrates_into_view() {
    let json = serde_json::json!({
        "id": 1,
        "name": "Sheet1",
        "frozen_rows": 2,
        "frozen_cols": 3,
        "zoom": 1.25
    });

    let sheet: Worksheet = serde_json::from_value(json).unwrap();
    assert_eq!(sheet.frozen_rows, 2);
    assert_eq!(sheet.frozen_cols, 3);
    assert!((sheet.zoom - 1.25).abs() < f32::EPSILON);
    assert_eq!(sheet.view.pane.frozen_rows, 2);
    assert_eq!(sheet.view.pane.frozen_cols, 3);
    assert!((sheet.view.zoom - 1.25).abs() < f32::EPSILON);
}

#[test]
fn view_overrides_legacy_freeze_and_zoom_on_deserialize() {
    let json = serde_json::json!({
        "id": 1,
        "name": "Sheet1",
        "frozen_rows": 1,
        "frozen_cols": 1,
        "zoom": 1.0,
        "view": {
            "pane": { "frozen_rows": 4, "frozen_cols": 5 },
            "zoom": 2.0
        }
    });

    let sheet: Worksheet = serde_json::from_value(json).unwrap();
    assert_eq!(sheet.frozen_rows, 4);
    assert_eq!(sheet.frozen_cols, 5);
    assert!((sheet.zoom - 2.0).abs() < f32::EPSILON);
    assert_eq!(sheet.view.pane.frozen_rows, 4);
    assert_eq!(sheet.view.pane.frozen_cols, 5);
    assert!((sheet.view.zoom - 2.0).abs() < f32::EPSILON);
}
