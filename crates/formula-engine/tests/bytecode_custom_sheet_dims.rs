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

#[test]
fn bytecode_custom_sheet_dims_use_referenced_sheet_for_sheet_prefixed_whole_row_col_refs() {
    let mut engine = Engine::new();
    engine
        .set_sheet_dimensions("Sheet1", 10, 5)
        .expect("set Sheet1 dimensions");
    engine
        .set_sheet_dimensions("Sheet2", 20, 7)
        .expect("set Sheet2 dimensions");

    // Put values on Sheet2 in rows/cols that would be out of bounds if the lowerer accidentally
    // used Sheet1's dimensions.
    engine.set_cell_value("Sheet2", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet2", "A20", 3.0).unwrap();

    // Evaluate on Sheet1 but reference Sheet2's whole column/row.
    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(Sheet2!A:A)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=ROWS(Sheet2!A:A)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B3", "=COLUMNS(Sheet2!1:1)")
        .unwrap();

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
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(20.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B3"), Value::Number(7.0));
}

#[test]
fn bytecode_custom_sheet_dims_expand_3d_whole_column_per_sheet() {
    let mut engine = Engine::new();
    engine
        .set_sheet_dimensions("Sheet1", 10, 5)
        .expect("set Sheet1 dimensions");
    engine
        .set_sheet_dimensions("Sheet2", 20, 5)
        .expect("set Sheet2 dimensions");

    engine.set_cell_value("Sheet1", "A10", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A20", 2.0).unwrap();

    // Put the formula on Sheet1 but avoid circular refs by keeping it out of column A.
    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(Sheet1:Sheet2!A:A)")
        .unwrap();

    let report = engine.bytecode_compile_report(10);
    assert!(
        report.is_empty(),
        "expected formulas to compile to bytecode on custom sheet dims; got: {report:?}"
    );

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(3.0));
}

#[test]
fn bytecode_custom_sheet_dims_expand_3d_whole_row_per_sheet() {
    let mut engine = Engine::new();
    engine
        .set_sheet_dimensions("Sheet1", 10, 5)
        .expect("set Sheet1 dimensions");
    engine
        .set_sheet_dimensions("Sheet2", 10, 7)
        .expect("set Sheet2 dimensions");

    // Row 1 spans different column counts on each sheet.
    engine.set_cell_value("Sheet1", "E1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "G1", 2.0).unwrap();

    // Put the formula on row 2 to avoid a circular reference (since `1:1` includes row 1).
    engine
        .set_cell_formula("Sheet1", "B2", "=SUM(Sheet1:Sheet2!1:1)")
        .unwrap();

    let report = engine.bytecode_compile_report(10);
    assert!(
        report.is_empty(),
        "expected formulas to compile to bytecode on custom sheet dims; got: {report:?}"
    );

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(3.0));
}
