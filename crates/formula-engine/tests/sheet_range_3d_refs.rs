use formula_engine::{parse_formula, Engine, ParseOptions, SerializeOptions, Value};
use pretty_assertions::assert_eq;

#[test]
fn parse_and_roundtrip_sheet_range_ref() {
    let ast = parse_formula("=SUM(Sheet1:Sheet3!A1)", ParseOptions::default()).unwrap();
    let roundtrip = ast.to_string(SerializeOptions::default()).unwrap();
    assert_eq!(roundtrip, "=SUM(Sheet1:Sheet3!A1)");
}

#[test]
fn evaluates_sum_over_sheet_range_cell_ref() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();

    engine
        .set_cell_formula("Summary", "A1", "=SUM(Sheet1:Sheet3!A1)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Summary", "A1"), Value::Number(6.0));
}

#[test]
fn evaluates_sum_over_sheet_range_area_ref() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 10.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet2", "A2", 20.0).unwrap();
    engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();
    engine.set_cell_value("Sheet3", "A2", 30.0).unwrap();

    engine
        .set_cell_formula("Summary", "A1", "=SUM(Sheet1:Sheet3!A1:A2)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Summary", "A1"), Value::Number(66.0));
}

#[test]
fn recalculates_when_intermediate_sheet_changes() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();

    engine
        .set_cell_formula("Summary", "A1", "=SUM(Sheet1:Sheet3!A1)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Summary", "A1"), Value::Number(6.0));

    engine.set_cell_value("Sheet2", "A1", 5.0).unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Summary", "A1"), Value::Number(9.0));
}

