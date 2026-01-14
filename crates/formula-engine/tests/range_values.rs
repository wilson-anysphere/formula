use formula_engine::{Engine, ErrorKind, Value};
use formula_model::Range;

#[test]
fn get_range_values_includes_blanks_for_unset_cells() {
    let mut engine = Engine::new();
    engine
        .set_cell_value("Sheet1", "A1", 1.0)
        .expect("set A1");
    engine
        .set_cell_value("Sheet1", "C2", "hello")
        .expect("set C2");

    let range = Range::from_a1("A1:C2").expect("range");
    let values = engine
        .get_range_values("Sheet1", range)
        .expect("get_range_values");

    assert_eq!(
        values,
        vec![
            vec![Value::Number(1.0), Value::Blank, Value::Blank],
            vec![
                Value::Blank,
                Value::Blank,
                Value::Text("hello".to_string())
            ],
        ]
    );
}

#[test]
fn get_range_values_returns_ref_for_out_of_bounds_cells() {
    let mut engine = Engine::new();
    engine.set_sheet_dimensions("Sheet1", 2, 2).unwrap(); // A1:B2
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 2.0).unwrap();

    let range = Range::from_a1("A1:C3").unwrap();
    let values = engine.get_range_values("Sheet1", range).unwrap();

    assert_eq!(
        values,
        vec![
            vec![
                Value::Number(1.0),
                Value::Blank,
                Value::Error(ErrorKind::Ref)
            ],
            vec![
                Value::Blank,
                Value::Number(2.0),
                Value::Error(ErrorKind::Ref)
            ],
            vec![
                Value::Error(ErrorKind::Ref),
                Value::Error(ErrorKind::Ref),
                Value::Error(ErrorKind::Ref)
            ],
        ]
    );
}

#[test]
fn get_range_values_includes_spilled_array_outputs() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=SEQUENCE(2,2)")
        .unwrap();
    engine.recalculate();

    let range = Range::from_a1("A1:B2").unwrap();
    let values = engine.get_range_values("Sheet1", range).unwrap();

    assert_eq!(
        values,
        vec![
            vec![Value::Number(1.0), Value::Number(2.0)],
            vec![Value::Number(3.0), Value::Number(4.0)],
        ]
    );
}
