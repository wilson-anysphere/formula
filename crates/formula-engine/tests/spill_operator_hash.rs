use formula_engine::eval::parse_a1;
use formula_engine::{Engine, ErrorKind, Value};

#[test]
fn spill_operator_hash_expands_spill_range_in_functions() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();

    // C1 spills down (C1:C3).
    engine.set_cell_formula("Sheet1", "C1", "=A1:A3").unwrap();

    // SUM should see the full spill range (C1:C3) when using the spill operator.
    engine
        .set_cell_formula("Sheet1", "D1", "=SUM(C1#)")
        .unwrap();
    // Also ensure references from within a spill range resolve to the spill origin.
    engine
        .set_cell_formula("Sheet1", "D2", "=SUM(C2#)")
        .unwrap();

    // Ensure spill-range formulas compile to bytecode (no AST fallback).
    let report = engine.bytecode_compile_report(usize::MAX);
    for addr in ["C1", "D1", "D2"] {
        let cell = parse_a1(addr).unwrap();
        assert!(
            !report.iter().any(|e| e.addr == cell),
            "expected {addr} to compile to bytecode (report={report:?})"
        );
    }

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(6.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(6.0));
}

#[test]
fn spill_operator_hash_spills_again_when_used_as_formula_result() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();
    engine.set_cell_formula("Sheet1", "C1", "=A1:A3").unwrap();

    // Referencing a spill range as a top-level formula returns an array, which should spill.
    engine.set_cell_formula("Sheet1", "E1", "=C1#").unwrap();

    // Ensure the spill-range reference compiles to bytecode.
    let report = engine.bytecode_compile_report(usize::MAX);
    for addr in ["C1", "E1"] {
        let cell = parse_a1(addr).unwrap();
        assert!(
            !report.iter().any(|e| e.addr == cell),
            "expected {addr} to compile to bytecode (report={report:?})"
        );
    }

    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "E1").expect("spill range");
    assert_eq!(start, parse_a1("E1").unwrap());
    assert_eq!(end, parse_a1("E3").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E2"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E3"), Value::Number(3.0));
}

#[test]
fn spill_operator_hash_on_non_spill_origin_returns_ref() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_formula("Sheet1", "B1", "=A1#").unwrap();

    // `A1#` should be bytecode-eligible even though it returns `#REF!` at runtime.
    let report = engine.bytecode_compile_report(usize::MAX);
    let b1 = parse_a1("B1").unwrap();
    assert!(
        !report.iter().any(|e| e.addr == b1),
        "expected B1 to compile to bytecode (report={report:?})"
    );

    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Error(ErrorKind::Ref)
    );
}

#[test]
fn spill_operator_hash_preserves_reference_semantics_for_let_bindings() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();

    // C1 spills down (C1:C3).
    engine.set_cell_formula("Sheet1", "C1", "=A1:A3").unwrap();

    // LET bindings that capture a reference to a spill origin should be usable with the spill
    // operator (`x#`) without implicit intersection turning the reference into a scalar.
    engine
        .set_cell_formula("Sheet1", "D1", "=LET(x,C1,SUM(x#))")
        .unwrap();

    // Ensure both formulas compile to bytecode (no AST fallback).
    assert_eq!(
        engine.bytecode_program_count(),
        2,
        "expected spill-range LET formula to compile to bytecode"
    );

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(6.0));
}
