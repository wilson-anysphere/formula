use formula_engine::eval::CellAddr;
use formula_engine::value::RecordValue;
use formula_engine::{Engine, PrecedentNode, Value};
use std::collections::HashMap;

#[test]
fn field_access_tracks_base_reference_dependencies() {
    let mut engine = Engine::new();

    engine
        .set_cell_value(
            "Sheet1",
            "A1",
            Value::Record(RecordValue::with_fields(
                "Record",
                HashMap::from([("Price".to_string(), Value::Number(10.0))]),
            )),
        )
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=A1.Price")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(10.0));
    assert_eq!(
        engine.precedents("Sheet1", "B1").unwrap(),
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 0 }, // A1
        }]
    );
    assert_eq!(
        engine.dependents("Sheet1", "A1").unwrap(),
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 1 }, // B1
        }]
    );

    engine
        .set_cell_value(
            "Sheet1",
            "A1",
            Value::Record(RecordValue::with_fields(
                "Record",
                HashMap::from([("Price".to_string(), Value::Number(25.0))]),
            )),
        )
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(25.0));
}
