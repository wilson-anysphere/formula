use formula_engine::{Engine, Value};

#[test]
fn set_cell_value_blank_preserves_style_only_cells_and_keeps_sparse_storage() {
    let mut engine = Engine::new();

    // A formatted cell should keep its formatting when clearing contents.
    engine.set_cell_value("Sheet1", "A1", 123.0).unwrap();
    engine.set_cell_style_id("Sheet1", "A1", 42).unwrap();
    engine.set_cell_value("Sheet1", "A1", Value::Blank).unwrap();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Blank);
    assert_eq!(engine.get_cell_formula("Sheet1", "A1"), None);
    assert_eq!(engine.get_cell_style_id("Sheet1", "A1").unwrap(), Some(42));

    // An unformatted cell should be removed from the sparse map when clearing contents.
    engine.set_cell_value("Sheet1", "B1", 5.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", Value::Blank).unwrap();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Blank);
    assert_eq!(engine.get_cell_formula("Sheet1", "B1"), None);
    assert_eq!(engine.get_cell_style_id("Sheet1", "B1").unwrap(), None);
}
