use formula_engine::{Engine, ErrorKind, Value};

// Notes on Excel behavior:
// Microsoft documents PHONETIC's argument as:
// "Text string or a reference to a single cell or a range of cells that contain a furigana text
// string."
//
// https://support.microsoft.com/en-us/office/phonetic-function-9a329dac-0c0f-42f8-9a55-639086988554

#[test]
fn phonetic_accepts_scalar_string_literals() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", r#"=PHONETIC("abc")"#)
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Text("abc".to_string()));
}

#[test]
fn phonetic_coerces_scalar_numbers_to_text() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=PHONETIC(1.5)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text("1.5".to_string())
    );
}

#[test]
fn phonetic_coerces_scalar_booleans_to_text() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=PHONETIC(TRUE)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text("TRUE".to_string())
    );
}

#[test]
fn phonetic_propagates_scalar_errors() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=PHONETIC(NA())")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Error(ErrorKind::NA));
}

#[test]
fn phonetic_lifts_over_array_literals() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", r#"=PHONETIC({"a","b"})"#)
        .unwrap();
    engine.recalculate_single_threaded();

    // Dynamic arrays should spill across the input array shape.
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Text("a".to_string()));
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Text("b".to_string()));
}

