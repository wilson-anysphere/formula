use formula_engine::{Engine, Value};

#[test]
fn info_origin_reflects_host_sheet_origin() {
    let mut engine = Engine::new();

    // The INFO key is trimmed/case-insensitive; ensure dependency detection matches runtime parsing.
    engine
        .set_cell_formula("Sheet1", "A1", r#"=INFO(" origin ")"#)
        .unwrap();
    engine.recalculate();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text("$A$1".to_string())
    );

    engine.set_sheet_origin("Sheet1", Some("C5")).unwrap();
    assert!(engine.has_dirty_cells());

    engine.recalculate();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text("$C$5".to_string())
    );
}

#[test]
fn info_origin_is_scoped_per_sheet() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", r#"=INFO("origin")"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet2", "A1", r#"=INFO("origin")"#)
        .unwrap();

    engine.set_sheet_origin("Sheet1", Some("C5")).unwrap();
    engine.set_sheet_origin("Sheet2", Some("D6")).unwrap();

    engine.recalculate();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text("$C$5".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet2", "A1"),
        Value::Text("$D$6".to_string())
    );
}

#[test]
fn info_origin_dynamic_key_marks_dependents_dirty() {
    let mut engine = Engine::new();

    // Use a runtime-resolved key (cell reference) so dependency analysis must be conservative.
    engine.set_cell_value("Sheet1", "A1", "origin").unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", r#"=INFO(A1)"#)
        .unwrap();

    engine.recalculate();
    assert!(!engine.has_dirty_cells());
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("$A$1".to_string())
    );

    engine.set_sheet_origin("Sheet1", Some("C5")).unwrap();
    assert!(engine.has_dirty_cells());

    engine.recalculate();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("$C$5".to_string())
    );
}

#[test]
fn info_origin_falls_back_to_legacy_engine_info_override() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", r#"=INFO("origin")"#)
        .unwrap();

    // Backward compatibility: if the host uses the legacy `EngineInfo.origin*` plumbing,
    // `INFO("origin")` should still resolve that A1 reference when no sheet view origin is set.
    engine.set_info_origin_for_sheet("Sheet1", Some("C5"));

    engine.recalculate();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text("$C$5".to_string())
    );
}

#[test]
fn info_origin_falls_back_to_legacy_engine_info_origin() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", r#"=INFO("origin")"#)
        .unwrap();
    engine.recalculate();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text("$A$1".to_string())
    );

    // Legacy workbook-level origin metadata should be observed when the per-sheet origin is unset.
    engine.set_info_origin(Some("D6"));
    engine.recalculate();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text("$D$6".to_string())
    );

    // The new per-sheet `set_sheet_origin` API takes precedence.
    engine.set_sheet_origin("Sheet1", Some("C5")).unwrap();
    engine.recalculate();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text("$C$5".to_string())
    );
}

#[test]
fn info_origin_legacy_per_sheet_overrides_workbook_level_fallback() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", r#"=INFO("origin")"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet2", "A1", r#"=INFO("origin")"#)
        .unwrap();

    engine.set_info_origin(Some("D6"));
    engine.set_info_origin_for_sheet("Sheet1", Some("C5"));

    engine.recalculate();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text("$C$5".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet2", "A1"),
        Value::Text("$D$6".to_string())
    );
}
