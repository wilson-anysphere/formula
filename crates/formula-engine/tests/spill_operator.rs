use formula_engine::{parse_formula, Engine, ParseOptions, SerializeOptions, Value};
use pretty_assertions::assert_eq;

#[test]
fn spill_operator_roundtrips_in_canonical_ast() {
    let opts = ParseOptions::default();
    let ser = SerializeOptions::default();

    let ast = parse_formula("=A1#", opts.clone()).unwrap();
    let rendered = ast.to_string(ser).unwrap();
    assert_eq!(rendered, "=A1#");

    let ast2 = parse_formula(&rendered, opts).unwrap();
    assert_eq!(ast, ast2);
}

#[test]
fn engine_evaluates_spill_operator_against_spilled_ranges() {
    let mut engine = Engine::new();

    // Seed a 2x2 source range away from the spill targets.
    engine.set_cell_value("Sheet1", "D1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "E1", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "D2", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "E2", 4.0).unwrap();

    // A1 spills into A1:B2 by referencing D1:E2.
    engine.set_cell_formula("Sheet1", "A1", "=D1:E2").unwrap();
    engine.set_cell_formula("Sheet1", "G1", "=A1#").unwrap();
    engine.set_cell_formula("Sheet1", "I1", "=B1#").unwrap();
    engine.set_cell_formula("Sheet1", "C4", "=SUM(A1#)").unwrap();

    engine.recalculate();

    // Validate the original spill.
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(4.0));

    // Validate referencing the spill via `A1#` in a different region.
    assert_eq!(engine.get_cell_value("Sheet1", "G1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "H1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G2"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "H2"), Value::Number(4.0));

    // `B1` is a spill child; `B1#` should resolve to the same range as `A1#`.
    assert_eq!(engine.get_cell_value("Sheet1", "I1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "J1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "I2"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "J2"), Value::Number(4.0));

    assert_eq!(engine.get_cell_value("Sheet1", "C4"), Value::Number(10.0));

    // Change the spill shape from 2x2 to 2x1.
    engine.set_cell_formula("Sheet1", "A1", "=D1:D2").unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Blank);
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Blank);

    // `A1#` now spills a single column.
    assert_eq!(engine.get_cell_value("Sheet1", "G1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G2"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "H1"), Value::Blank);
    assert_eq!(engine.get_cell_value("Sheet1", "H2"), Value::Blank);

    // `B1` is no longer part of a spill; `B1#` should return `#REF!`.
    assert_eq!(engine.get_cell_value("Sheet1", "I1"), Value::Error(formula_engine::ErrorKind::Ref));
    assert_eq!(engine.get_cell_value("Sheet1", "J1"), Value::Blank);
    assert_eq!(engine.get_cell_value("Sheet1", "I2"), Value::Blank);
    assert_eq!(engine.get_cell_value("Sheet1", "J2"), Value::Blank);

    assert_eq!(engine.get_cell_value("Sheet1", "C4"), Value::Number(4.0));
}
