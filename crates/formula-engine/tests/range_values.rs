use formula_engine::{Engine, Value};
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

