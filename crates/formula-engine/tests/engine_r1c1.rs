use formula_engine::eval::CellAddr;
use formula_engine::value::RecordValue;
use formula_engine::{eval, Engine, PrecedentNode, Value};
use std::collections::HashMap;

#[test]
fn engine_evaluates_r1c1_relative_cell_reference_equivalent_to_a1() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 42.0).unwrap();

    engine.set_cell_formula("Sheet1", "C5", "=A1").unwrap();
    engine.recalculate();
    let a1_value = engine.get_cell_value("Sheet1", "C5");

    engine
        .set_cell_formula_r1c1("Sheet1", "C5", "=R[-4]C[-2]")
        .unwrap();
    engine.recalculate();
    let r1c1_value = engine.get_cell_value("Sheet1", "C5");

    assert_eq!(a1_value, r1c1_value);
    assert_eq!(r1c1_value, Value::Number(42.0));
}

#[test]
fn engine_evaluates_r1c1_ranges_equivalent_to_a1() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();

    engine.set_cell_formula("Sheet1", "C5", "=A1:A3").unwrap();
    engine.recalculate();
    let a1_value = engine.get_cell_value("Sheet1", "C5");
    let a1_spill_range = engine.spill_range("Sheet1", "C5");

    assert_eq!(a1_value, Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C6"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C7"), Value::Number(3.0));
    assert_eq!(
        a1_spill_range,
        Some((eval::parse_a1("C5").unwrap(), eval::parse_a1("C7").unwrap()))
    );

    engine
        .set_cell_formula_r1c1("Sheet1", "C5", "=R1C1:R3C1")
        .unwrap();
    engine.recalculate();
    let r1c1_value = engine.get_cell_value("Sheet1", "C5");
    let r1c1_spill_range = engine.spill_range("Sheet1", "C5");

    assert_eq!(a1_value, r1c1_value);
    assert_eq!(engine.get_cell_value("Sheet1", "C6"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C7"), Value::Number(3.0));
    assert_eq!(
        r1c1_spill_range,
        Some((eval::parse_a1("C5").unwrap(), eval::parse_a1("C7").unwrap()))
    );
    assert_eq!(a1_spill_range, r1c1_spill_range);
}

#[test]
fn engine_renders_stored_a1_formula_as_r1c1_for_cell() {
    let mut engine = Engine::new();
    engine.set_cell_formula("Sheet1", "C5", "=A1").unwrap();

    assert_eq!(
        engine.get_cell_formula_r1c1("Sheet1", "C5"),
        Some("=R[-4]C[-2]".to_string())
    );

    engine.set_cell_formula("Sheet1", "C5", "=$A$1").unwrap();
    assert_eq!(
        engine.get_cell_formula_r1c1("Sheet1", "C5"),
        Some("=R1C1".to_string())
    );
}

#[test]
fn engine_evaluates_r1c1_sheet_range_refs() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();
    engine
        .set_cell_formula_r1c1("Summary", "A1", "=SUM(Sheet1:Sheet3!R1C1)")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Summary", "A1"), Value::Number(6.0));
}

#[test]
fn engine_evaluates_r1c1_sheet_range_refs_with_quoted_span() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet 1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet 2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet 3", "A1", 3.0).unwrap();
    engine
        .set_cell_formula_r1c1("Summary", "A1", "=SUM('Sheet 1:Sheet 3'!R1C1)")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Summary", "A1"), Value::Number(6.0));
}

#[test]
fn engine_evaluates_r1c1_field_access_and_tracks_dependencies() {
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

    // From B1, `RC[-1]` refers to A1. This regression test ensures the lexer treats `.Price` as
    // field access rather than part of an identifier like `RC[-1].Price`.
    engine
        .set_cell_formula_r1c1("Sheet1", "B1", "=RC[-1].Price")
        .unwrap();

    // Formulas are stored in canonical A1 style but can be rendered back to R1C1.
    assert_eq!(engine.get_cell_formula("Sheet1", "B1"), Some("=A1.Price"));
    assert_eq!(
        engine.get_cell_formula_r1c1("Sheet1", "B1"),
        Some("=RC[-1].Price".to_string())
    );

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

#[test]
fn engine_evaluates_r1c1_bracket_field_access_and_roundtrips() {
    let mut engine = Engine::new();

    engine
        .set_cell_value(
            "Sheet1",
            "A1",
            Value::Record(RecordValue::with_fields(
                "Record",
                HashMap::from([("Change%".to_string(), Value::Number(5.0))]),
            )),
        )
        .unwrap();

    // Field names that aren't valid identifier segments (spaces, punctuation, etc) use Excel's
    // bracketed selector syntax.
    engine
        .set_cell_formula_r1c1("Sheet1", "B1", r#"=RC[-1].["Change%"]"#)
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(5.0));
    assert_eq!(
        engine.get_cell_formula("Sheet1", "B1"),
        Some(r#"=A1.["Change%"]"#)
    );
    assert_eq!(
        engine.get_cell_formula_r1c1("Sheet1", "B1"),
        Some(r#"=RC[-1].["Change%"]"#.to_string())
    );
}
