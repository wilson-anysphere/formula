#![cfg(not(target_arch = "wasm32"))]

use formula_engine::{BytecodeCompileReason, Engine, Value};

#[test]
fn bytecode_custom_sheet_dims_whole_row_and_column_refs() {
    let mut engine = Engine::new();
    engine
        .set_sheet_dimensions("Sheet1", 10, 5)
        .expect("set sheet dimensions");

    // Populate a few values in column A (leave implicit blanks elsewhere).
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A10", 3.0).unwrap();

    // Avoid circular references by keeping these formulas out of the referenced row/column.
    engine.set_cell_formula("Sheet1", "B1", "=SUM(A:A)").unwrap();
    engine.set_cell_formula("Sheet1", "B2", "=ROWS(A:A)").unwrap();
    engine
        .set_cell_formula("Sheet1", "B3", "=COLUMNS(1:1)")
        .unwrap();

    // These formulas should compile to bytecode (no AST fallback reasons).
    let report = engine.bytecode_compile_report(10);
    assert!(
        report.is_empty(),
        "expected formulas to compile to bytecode on custom sheet dims; got: {report:?}"
    );
    assert!(
        !report.iter().any(|e| e.reason == BytecodeCompileReason::NonDefaultSheetDimensions),
        "NonDefaultSheetDimensions should not be reported after dimension-aware lowering"
    );

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(6.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(10.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B3"), Value::Number(5.0));
}
