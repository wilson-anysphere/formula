#![cfg(not(target_arch = "wasm32"))]

use formula_engine::eval::CellAddr;
use formula_engine::{BytecodeCompileReason, Engine, PrecedentNode, Value};

#[test]
fn bytecode_dynamic_deps_offset_compiles_and_traces_dynamic_dependents() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 20.0).unwrap();

    // OFFSET(A1,0,0) should preserve reference semantics for its first argument (A1 is treated as
    // a reference, not a scalar), so the result is the value in A1.
    engine
        .set_cell_formula("Sheet1", "B1", "=OFFSET(A1,0,0)")
        .unwrap();

    // This OFFSET returns a dynamically-determined reference (A2). The engine should trace that
    // dependency at runtime and update the dependency graph accordingly.
    engine
        .set_cell_formula("Sheet1", "C1", "=OFFSET(A1,1,0)")
        .unwrap();

    let report = engine.bytecode_compile_report(10);
    assert!(
        report.is_empty(),
        "expected OFFSET formulas to compile to bytecode, got: {report:?}"
    );
    assert!(
        !report
            .iter()
            .any(|e| matches!(e.reason, BytecodeCompileReason::DynamicDependencies)),
        "unexpected DynamicDependencies fallback: {report:?}"
    );

    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(10.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(20.0));

    // The OFFSET target cell should be visible as a dependent after evaluation updates the graph.
    let dependents = engine.dependents("Sheet1", "A2").unwrap();
    assert_eq!(
        dependents,
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 2 } // C1
        }]
    );
}

#[test]
fn bytecode_dynamic_deps_indirect_compiles_and_traces_dynamic_dependents_across_sheets() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 10.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 42.0).unwrap();

    // Constant-string INDIRECT has no static cell precedents; runtime dependency tracing is
    // required for the engine to build correct graph edges.
    engine
        .set_cell_formula("Sheet1", "B1", r#"=INDIRECT("A1")"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", r#"=INDIRECT("Sheet2!A1")"#)
        .unwrap();

    let report = engine.bytecode_compile_report(10);
    assert!(
        report.is_empty(),
        "expected INDIRECT formulas to compile to bytecode, got: {report:?}"
    );

    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(10.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(42.0));

    let dependents_a1 = engine.dependents("Sheet1", "A1").unwrap();
    assert_eq!(
        dependents_a1,
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 1 } // B1
        }]
    );

    let dependents_sheet2_a1 = engine.dependents("Sheet2", "A1").unwrap();
    assert_eq!(
        dependents_sheet2_a1,
        vec![PrecedentNode::Cell {
            sheet: 0,
            addr: CellAddr { row: 0, col: 2 } // C1
        }]
    );
}
