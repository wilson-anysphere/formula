use formula_engine::Engine;

#[test]
fn sheet_list_and_mapping_apis_reflect_tab_order() {
    let mut engine = Engine::new();
    engine.ensure_sheet("Sheet1");
    engine.ensure_sheet("Sheet2");
    engine.ensure_sheet("Sheet3");

    let id1 = engine.sheet_id("Sheet1").unwrap();
    let id2 = engine.sheet_id("sheet2").unwrap(); // case-insensitive
    let id3 = engine.sheet_id("Sheet3").unwrap();

    assert_eq!(engine.sheet_name(id1), Some("Sheet1"));
    assert_eq!(engine.sheet_name(id2), Some("Sheet2"));
    assert_eq!(engine.sheet_name(id3), Some("Sheet3"));

    assert_eq!(engine.sheet_ids_in_order(), vec![id1, id2, id3]);
    assert_eq!(
        engine.sheet_names_in_order(),
        vec![
            "Sheet1".to_string(),
            "Sheet2".to_string(),
            "Sheet3".to_string()
        ]
    );
    assert_eq!(
        engine.sheets_in_order(),
        vec![
            (id1, "Sheet1".to_string()),
            (id2, "Sheet2".to_string()),
            (id3, "Sheet3".to_string())
        ]
    );

    // Unicode/NFKC-aware lookup: composed vs decomposed forms should match.
    let mut engine = Engine::new();
    engine.ensure_sheet("Cafe\u{0301}"); // "Café" (decomposed)
    let id = engine.sheet_id("Café").unwrap();
    assert_eq!(engine.sheet_name(id), Some("Cafe\u{0301}"));
}

#[test]
fn sheet_order_can_change_without_changing_sheet_ids() {
    let mut engine = Engine::new();
    engine.ensure_sheet("Sheet1");
    engine.ensure_sheet("Sheet2");
    engine.ensure_sheet("Sheet3");

    let id1 = engine.sheet_id("Sheet1").unwrap();
    let id2 = engine.sheet_id("Sheet2").unwrap();
    let id3 = engine.sheet_id("Sheet3").unwrap();

    assert!(engine.reorder_sheet("Sheet3", 0));
    assert_eq!(engine.sheet_ids_in_order(), vec![id3, id1, id2]);

    // Name lookup should still return stable ids.
    assert_eq!(engine.sheet_id("Sheet1"), Some(id1));
    assert_eq!(engine.sheet_id("Sheet2"), Some(id2));
    assert_eq!(engine.sheet_id("Sheet3"), Some(id3));
}

#[test]
fn sheet_rename_updates_name_mapping_and_formula_text() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine
        .set_cell_formula("Sheet2", "A1", "=Sheet1!A1")
        .unwrap();

    let id1 = engine.sheet_id("Sheet1").unwrap();
    assert!(engine.rename_sheet("Sheet1", "Renamed"));

    assert_eq!(engine.sheet_name(id1), Some("Renamed"));
    assert_eq!(engine.sheet_id("Renamed"), Some(id1));
    assert_eq!(engine.sheet_id("Sheet1"), None);

    // Stored formula text should be rewritten so future recompilation keeps working.
    assert_eq!(engine.get_cell_formula("Sheet2", "A1"), Some("=Renamed!A1"));
}

#[test]
fn deleting_a_sheet_removes_it_from_order_and_mapping() {
    let mut engine = Engine::new();
    engine.ensure_sheet("Sheet1");
    engine.ensure_sheet("Sheet2");
    engine.ensure_sheet("Sheet3");

    let id2 = engine.sheet_id("Sheet2").unwrap();
    engine.delete_sheet("Sheet2").unwrap();

    assert_eq!(engine.sheet_name(id2), None);
    assert_eq!(engine.sheet_id("Sheet2"), None);
    assert!(!engine.sheet_ids_in_order().contains(&id2));

    // Re-adding the same name should allocate a fresh id (stable ids are never reused).
    engine.ensure_sheet("Sheet2");
    let new_id2 = engine.sheet_id("Sheet2").unwrap();
    assert_ne!(new_id2, id2);
}
