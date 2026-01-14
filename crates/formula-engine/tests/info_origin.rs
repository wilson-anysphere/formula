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

