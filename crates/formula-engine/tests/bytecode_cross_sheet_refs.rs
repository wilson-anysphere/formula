use formula_engine::{Engine, Value};

#[test]
fn bytecode_cross_sheet_refs_compile_and_eval() {
    let mut engine = Engine::new();
    engine.ensure_sheet("Sheet1");
    engine.ensure_sheet("Sheet2");

    engine.set_cell_value("Sheet2", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet2", "A3", 3.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "A1", "=Sheet2!A1")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=Sheet2!A1+1")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A3", "=SUM(Sheet2!A1:A3)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A4", "=ROW(Sheet2!A3)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A5", "=COLUMN(Sheet2!B1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A6", "=ROWS(Sheet2!A1:A3)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A7", "=COLUMNS(Sheet2!A1:C1)")
        .unwrap();

    // Cross-sheet spill range operator (Sheet2 spills an array; Sheet1 references it).
    engine.set_cell_formula("Sheet2", "A5", "={1;2;3}").unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=Sheet2!A5#")
        .unwrap();

    // Ensure the formulas compiled to bytecode (and did not fall back to AST).
    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 9);
    assert_eq!(
        stats.compiled, 9,
        "expected all formulas to compile to bytecode"
    );
    assert_eq!(stats.fallback, 0);
    assert!(
        engine.bytecode_compile_report(usize::MAX).is_empty(),
        "expected no bytecode fallback report entries"
    );

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A3"), Value::Number(6.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A4"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A5"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A6"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A7"), Value::Number(3.0));

    // The spill range reference should spill the array onto Sheet1.
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B3"), Value::Number(3.0));
}
