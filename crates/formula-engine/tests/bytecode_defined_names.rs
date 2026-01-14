use formula_engine::{value::Array, Engine, ErrorKind, NameDefinition, NameScope, Value};

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
fn bytecode_inlines_structured_ref_defined_name_references() {
    use formula_model::table::TableColumn;
    use formula_model::{Range, Table};

    fn table_fixture(range: &str) -> Table {
        Table {
            id: 1,
            name: "Table1".into(),
            display_name: "Table1".into(),
            range: Range::from_a1(range).unwrap(),
            header_row_count: 1,
            totals_row_count: 0,
            columns: vec![TableColumn {
                id: 1,
                name: "Col1".into(),
                formula: None,
                totals_formula: None,
            }],
            style: None,
            auto_filter: None,
            relationship_id: None,
            part_path: None,
        }
    }

    let mut engine = Engine::new();
    engine.ensure_sheet("Sheet1");
    engine.set_sheet_tables("Sheet1", vec![table_fixture("A1:A3")]);

    engine.set_cell_value("Sheet1", "A2", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 2.0).unwrap();

    engine
        .define_name(
            "MyCol",
            NameScope::Workbook,
            NameDefinition::Reference("Table1[Col1]".to_string()),
        )
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(MyCol)")
        .unwrap();

    // Ensure we're exercising the bytecode path (the structured ref is inlined into the formula
    // and rewritten to a supported RangeRef).
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(3.0));

    // Ensure the AST evaluator agrees with the bytecode backend.
    let debug = engine.debug_evaluate("Sheet1", "B1").unwrap();
    assert_eq!(debug.value, engine.get_cell_value("Sheet1", "B1"));
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
    engine.set_cell_formula("Sheet1", "A1", "=SelfRef").unwrap();

    // The bytecode backend should cleanly fall back to AST evaluation.
    assert_eq!(engine.bytecode_program_count(), 0);

    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::Name)
    );
}

#[test]
fn bytecode_inlines_defined_name_constants_inside_let_when_unshadowed() {
    let mut engine = Engine::new();

    engine
        .define_name(
            "X",
            NameScope::Workbook,
            NameDefinition::Constant(Value::Number(5.0)),
        )
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "A1", "=LET(y,1,X+1)")
        .unwrap();

    // Ensure we're exercising the bytecode path (name constant is inlined; lowering does not
    // support NameRef directly).
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(6.0));
}

#[test]
fn bytecode_does_not_inline_defined_name_constants_shadowed_by_let() {
    let mut engine = Engine::new();

    engine
        .define_name(
            "X",
            NameScope::Workbook,
            NameDefinition::Constant(Value::Number(5.0)),
        )
        .unwrap();

    // LET-bound variables shadow workbook defined names of the same identifier.
    engine
        .set_cell_formula("Sheet1", "A1", "=LET(X,1,X+1)")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(2.0));
}

#[test]
fn bytecode_inlines_defined_name_constants_under_percent_postfix() {
    let mut engine = Engine::new();

    engine
        .define_name(
            "Rate",
            NameScope::Workbook,
            NameDefinition::Constant(Value::Number(5.0)),
        )
        .unwrap();

    engine.set_cell_formula("Sheet1", "A1", "=Rate%").unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(0.05));
}

#[test]
fn bytecode_inlines_defined_name_constants_inside_array_literals() {
    let mut engine = Engine::new();

    engine
        .define_name(
            "X",
            NameScope::Workbook,
            NameDefinition::Constant(Value::Number(1.0)),
        )
        .unwrap();

    engine.set_cell_formula("Sheet1", "A1", "={X,2}").unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(2.0));
}

#[test]
fn bytecode_inlines_defined_name_constants_inside_concat_binary_op() {
    let mut engine = Engine::new();

    engine
        .define_name(
            "X",
            NameScope::Workbook,
            NameDefinition::Constant(Value::Number(1.0)),
        )
        .unwrap();

    engine.set_cell_formula("Sheet1", "A1", "=X&\"a\"").unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text("1a".to_string())
    );
}

#[test]
fn bytecode_inlines_defined_name_constant_arrays_as_array_literals() {
    let mut engine = Engine::new();

    engine
        .define_name(
            "Arr",
            NameScope::Workbook,
            NameDefinition::Constant(Value::Array(Array::new(
                2,
                2,
                vec![
                    Value::Number(1.0),
                    Value::Number(2.0),
                    Value::Number(3.0),
                    Value::Number(4.0),
                ],
            ))),
        )
        .unwrap();

    engine.set_cell_formula("Sheet1", "A1", "=Arr").unwrap();

    // Ensure we're exercising the bytecode path (the array constant is inlined into an array literal).
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(4.0));
}
