use formula_engine::{Engine, ErrorKind, NameDefinition, NameScope, Value};

#[test]
fn bytecode_inlines_named_range_into_sum() {
    let mut engine = Engine::new();

    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();

    engine
        .define_name(
            "MyRange",
            NameScope::Workbook,
            NameDefinition::Reference("Sheet1!$A$1:$A$3".to_string()),
        )
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(MyRange)")
        .unwrap();

    // Ensure we're exercising the bytecode path (name is inlined into a supported RangeRef).
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(6.0));

    // Ensure the AST evaluator agrees with the bytecode backend.
    let debug = engine.debug_evaluate("Sheet1", "B1").unwrap();
    assert_eq!(debug.value, engine.get_cell_value("Sheet1", "B1"));
}

#[test]
fn bytecode_inlines_name_formula_into_scalar_math() {
    let mut engine = Engine::new();

    engine
        .define_name(
            "MyExpr",
            NameScope::Workbook,
            NameDefinition::Formula("1+2".to_string()),
        )
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "A1", "=MyExpr*10")
        .unwrap();

    // Ensure we're exercising the bytecode path (name is inlined into scalar math).
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(30.0));

    // Ensure the AST evaluator agrees with the bytecode backend.
    let debug = engine.debug_evaluate("Sheet1", "A1").unwrap();
    assert_eq!(debug.value, engine.get_cell_value("Sheet1", "A1"));
}

#[test]
fn bytecode_defined_name_formula_cycles_fall_back_without_recursing_forever() {
    let mut engine = Engine::new();

    engine
        .define_name(
            "SelfRef",
            NameScope::Workbook,
            NameDefinition::Formula("=SelfRef+1".to_string()),
        )
        .unwrap();

    // If name inlining doesn't guard against cycles, this will recurse indefinitely during
    // bytecode compilation.
    engine
        .set_cell_formula("Sheet1", "A1", "=SelfRef")
        .unwrap();

    // The bytecode backend should cleanly fall back to AST evaluation.
    assert_eq!(engine.bytecode_program_count(), 0);

    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::Name)
    );
}
