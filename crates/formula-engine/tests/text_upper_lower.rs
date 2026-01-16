use formula_engine::{Engine, Value};

#[test]
fn upper_lower_ascii() {
  let mut engine = Engine::new();
  engine
    .set_cell_formula("Sheet1", "A1", r#"=UPPER("Abc")"#)
    .unwrap();
  engine
    .set_cell_formula("Sheet1", "A2", r#"=LOWER("AbC")"#)
    .unwrap();
  engine.recalculate();

  assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Text("ABC".to_string()));
  assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Text("abc".to_string()));
}

#[test]
fn upper_lower_unicode() {
  let mut engine = Engine::new();
  engine
    .set_cell_formula("Sheet1", "A1", r#"=UPPER("straße")"#)
    .unwrap();
  engine
    .set_cell_formula("Sheet1", "A2", r#"=LOWER("Ö")"#)
    .unwrap();
  engine.recalculate();

  assert_eq!(
    engine.get_cell_value("Sheet1", "A1"),
    Value::Text("STRASSE".to_string())
  );
  assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Text("ö".to_string()));
}

