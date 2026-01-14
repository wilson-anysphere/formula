use formula_engine::eval::parse_a1;
use formula_engine::value::RecordValue;
use formula_engine::{Engine, ErrorKind, Value};

#[test]
fn sortby_sorts_rows_by_single_key_descending() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 30.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "B3", 20.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "D1", "=SORTBY(A1:B3,B1:B3,-1)")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "D1").expect("spill range");
    assert_eq!(start, parse_a1("D1").unwrap());
    assert_eq!(end, parse_a1("E3").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(30.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E2"), Value::Number(20.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D3"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E3"), Value::Number(10.0));
}

#[test]
fn sortby_multi_key_sort_is_stable_for_ties() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "C1", "r1").unwrap();

    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "C2", "r2").unwrap();

    engine.set_cell_value("Sheet1", "A3", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B3", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "C3", "r3").unwrap();

    engine.set_cell_value("Sheet1", "A4", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B4", 0.0).unwrap();
    engine.set_cell_value("Sheet1", "C4", "r4").unwrap();

    engine.set_cell_value("Sheet1", "A5", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "B5", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "C5", "r5").unwrap();

    engine
        .set_cell_formula("Sheet1", "E1", "=SORTBY(A1:C5,A1:A5,1,B1:B5,1)")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "E1").expect("spill range");
    assert_eq!(start, parse_a1("E1").unwrap());
    assert_eq!(end, parse_a1("G5").unwrap());

    assert_eq!(
        engine.get_cell_value("Sheet1", "G1"),
        Value::Text("r4".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "G2"),
        Value::Text("r1".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "G3"),
        Value::Text("r3".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "G4"),
        Value::Text("r2".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "G5"),
        Value::Text("r5".to_string())
    );
}

#[test]
fn sortby_rejects_length_mismatch() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 4.0).unwrap();
    engine.set_cell_value("Sheet1", "B3", 5.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "D1", "=SORTBY(A1:A2,B1:B3)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "D1"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn sortby_rejects_invalid_sort_order() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "C1", "=SORTBY(A1:A2,A1:A2,0)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "C1"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn sortby_uses_record_display_field_for_keying() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", "first").unwrap();
    engine.set_cell_value("Sheet1", "A2", "second").unwrap();

    let mut key_a = RecordValue::new("zzz").field("Name", "A");
    // Use a different case from the stored field key to ensure display_field resolution is
    // case-insensitive.
    key_a.display_field = Some("name".to_string());
    let mut key_b = RecordValue::new("aaa").field("Name", "B");
    key_b.display_field = Some("NAME".to_string());

    engine
        .set_cell_value("Sheet1", "B1", Value::Record(key_a))
        .unwrap();
    engine
        .set_cell_value("Sheet1", "B2", Value::Record(key_b))
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "D1", "=SORTBY(A1:A2,B1:B2)")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "D1").expect("spill range");
    assert_eq!(start, parse_a1("D1").unwrap());
    assert_eq!(end, parse_a1("D2").unwrap());

    // Sort keys should be derived from the record display_field ("A"/"B"), not from fallback
    // display values ("zzz"/"aaa").
    assert_eq!(
        engine.get_cell_value("Sheet1", "D1"),
        Value::Text("first".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "D2"),
        Value::Text("second".to_string())
    );
}
