use formula_engine::{
    EditError, EditOp, Engine, ErrorKind, NameDefinition, NameScope, PrecedentNode, Value,
};

#[test]
fn workbook_scoped_name_can_be_used_in_other_cells() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 10.0).unwrap();
    engine
        .define_name(
            "MyX",
            NameScope::Workbook,
            NameDefinition::Reference("Sheet1!A1".to_string()),
        )
        .unwrap();

    engine.set_cell_formula("Sheet2", "B1", "=MyX*2").unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet2", "B1"), Value::Number(20.0));
}

#[test]
fn unicode_defined_names_are_case_insensitive() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 10.0).unwrap();
    engine
        .define_name(
            "ÜBER",
            NameScope::Workbook,
            NameDefinition::Reference("Sheet1!A1".to_string()),
        )
        .unwrap();

    engine.set_cell_formula("Sheet1", "B1", "=über*2").unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(20.0));
}

#[test]
fn sheet_scoped_names_shadow_workbook_names() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 10.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 5.0).unwrap();

    engine
        .define_name(
            "Foo",
            NameScope::Workbook,
            NameDefinition::Reference("Sheet1!A1".to_string()),
        )
        .unwrap();
    engine
        .define_name(
            "Foo",
            NameScope::Sheet("Sheet2"),
            NameDefinition::Reference("Sheet2!A1".to_string()),
        )
        .unwrap();

    engine.set_cell_formula("Sheet1", "B1", "=Foo").unwrap();
    engine.set_cell_formula("Sheet2", "B1", "=Foo").unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(10.0));
    assert_eq!(engine.get_cell_value("Sheet2", "B1"), Value::Number(5.0));
}

#[test]
fn name_defined_as_formula_registers_cell_precedents() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine
        .define_name(
            "MySum",
            NameScope::Workbook,
            NameDefinition::Formula("=Sheet1!A1+Sheet1!A2".to_string()),
        )
        .unwrap();
    engine.set_cell_formula("Sheet1", "B1", "=MySum*2").unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(6.0));

    // Updating a precedent referenced from the name should dirty the dependent cell.
    engine.set_cell_value("Sheet1", "A1", 2.0).unwrap();
    assert!(engine.is_dirty("Sheet1", "B1"));
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(8.0));
}

#[test]
fn changing_name_definition_marks_dependents_dirty_and_updates_dependencies() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 99.0).unwrap();
    engine
        .define_name(
            "MyX",
            NameScope::Workbook,
            NameDefinition::Reference("Sheet1!A1".to_string()),
        )
        .unwrap();
    engine.set_cell_formula("Sheet1", "B1", "=MyX").unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(10.0));

    // Repoint the name at a different cell; the dependent should become dirty.
    engine
        .define_name(
            "MyX",
            NameScope::Workbook,
            NameDefinition::Reference("Sheet1!A2".to_string()),
        )
        .unwrap();
    assert!(engine.is_dirty("Sheet1", "B1"));
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(99.0));

    // Dependencies should now follow A2 (not A1).
    engine.set_cell_value("Sheet1", "A1", 11.0).unwrap();
    assert!(!engine.is_dirty("Sheet1", "B1"));

    engine.set_cell_value("Sheet1", "A2", 100.0).unwrap();
    assert!(engine.is_dirty("Sheet1", "B1"));
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(100.0));
}

#[test]
fn name_definitions_do_not_implicitly_create_missing_sheets() {
    let mut engine = Engine::new();
    engine
        .define_name(
            "BadRef",
            NameScope::Workbook,
            NameDefinition::Reference("NoSuchSheet!A1".to_string()),
        )
        .unwrap();
    engine.set_cell_formula("Sheet1", "A1", "=BadRef").unwrap();
    engine.recalculate();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::Ref)
    );

    // Ensure the missing sheet wasn't created as a side effect of defining the name.
    let err = engine
        .apply_operation(EditOp::InsertRows {
            sheet: "NoSuchSheet".to_string(),
            row: 0,
            count: 1,
        })
        .unwrap_err();
    assert_eq!(err, EditError::SheetNotFound("NoSuchSheet".to_string()));
}

#[test]
fn name_defined_as_array_literal_spills_when_used_as_formula_result() {
    let mut engine = Engine::new();
    engine
        .define_name(
            "MyArr",
            NameScope::Workbook,
            NameDefinition::Formula("={1,2;3,4}".to_string()),
        )
        .unwrap();
    engine.set_cell_formula("Sheet1", "A1", "=MyArr").unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(4.0));
}

#[test]
fn defined_name_lambdas_called_like_functions_register_dependencies() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 99.0).unwrap();

    engine
        .define_name(
            "AddX",
            NameScope::Workbook,
            NameDefinition::Formula("=LAMBDA(x,Sheet1!A1+x)".to_string()),
        )
        .unwrap();

    engine.set_cell_formula("Sheet2", "B1", "=AddX(2)").unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet2", "B1"), Value::Number(12.0));

    // Updating a precedent referenced from the named lambda should dirty the dependent cell.
    engine.set_cell_value("Sheet1", "A1", 20.0).unwrap();
    assert!(engine.is_dirty("Sheet2", "B1"));
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet2", "B1"), Value::Number(22.0));

    // Repoint the name at a different cell; the dependent should become dirty and follow the new precedent.
    engine
        .define_name(
            "AddX",
            NameScope::Workbook,
            NameDefinition::Formula("=LAMBDA(x,Sheet1!A2+x)".to_string()),
        )
        .unwrap();
    assert!(engine.is_dirty("Sheet2", "B1"));
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet2", "B1"), Value::Number(101.0));

    engine.set_cell_value("Sheet1", "A1", 30.0).unwrap();
    assert!(!engine.is_dirty("Sheet2", "B1"));

    engine.set_cell_value("Sheet1", "A2", 100.0).unwrap();
    assert!(engine.is_dirty("Sheet2", "B1"));
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet2", "B1"), Value::Number(102.0));
}

#[test]
fn defining_name_after_function_call_recalculates_dependents() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=NewLambda(2)")
        .unwrap();
    engine.recalculate();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::Name)
    );

    engine
        .define_name(
            "NewLambda",
            NameScope::Workbook,
            NameDefinition::Formula("=LAMBDA(x,x+1)".to_string()),
        )
        .unwrap();

    assert!(engine.is_dirty("Sheet1", "A1"));
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(3.0));
}

#[test]
fn let_bound_lambda_calls_do_not_depend_on_same_named_defined_names() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 10.0).unwrap();
    engine
        .define_name(
            "F",
            NameScope::Workbook,
            NameDefinition::Reference("Sheet1!A1".to_string()),
        )
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", "=LET(f,LAMBDA(x,x+1),f(2))")
        .unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(3.0));

    // Changing the defined-name precedent should not dirty the LET-bound lambda call.
    engine.set_cell_value("Sheet1", "A1", 20.0).unwrap();
    assert!(!engine.is_dirty("Sheet1", "B1"));
}

#[test]
fn let_variable_names_do_not_register_defined_name_dependencies() {
    let mut engine = Engine::new();
    engine
        .define_name(
            "X",
            NameScope::Workbook,
            NameDefinition::Constant(Value::Number(1.0)),
        )
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "A1", "=LET(X,1,X+1)")
        .unwrap();
    engine.recalculate();
    assert!(
        !engine.is_dirty("Sheet1", "A1"),
        "cell should be clean after recalc"
    );

    // Updating a workbook-scoped name "X" should not dirty the cell because all references to "X"
    // in the LET expression are local bindings.
    engine
        .define_name(
            "X",
            NameScope::Workbook,
            NameDefinition::Constant(Value::Number(2.0)),
        )
        .unwrap();
    assert!(
        !engine.is_dirty("Sheet1", "A1"),
        "LET-bound variable should not create a defined-name dependency"
    );
}

#[test]
fn lambda_parameter_names_do_not_register_defined_name_dependencies() {
    let mut engine = Engine::new();
    engine
        .define_name(
            "X",
            NameScope::Workbook,
            NameDefinition::Constant(Value::Number(1.0)),
        )
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "A1", "=LAMBDA(X,X+1)(1)")
        .unwrap();
    engine.recalculate();
    assert!(
        !engine.is_dirty("Sheet1", "A1"),
        "cell should be clean after recalc"
    );

    // Updating a workbook-scoped name "X" should not dirty the cell because the LAMBDA parameter
    // shadows the defined name within the body.
    engine
        .define_name(
            "X",
            NameScope::Workbook,
            NameDefinition::Constant(Value::Number(2.0)),
        )
        .unwrap();
    assert!(
        !engine.is_dirty("Sheet1", "A1"),
        "LAMBDA parameters should not create a defined-name dependency"
    );
}

#[test]
fn let_locals_do_not_shadow_sheet_qualified_defined_name_dependencies() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "B1", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 20.0).unwrap();

    engine
        .define_name(
            "X",
            NameScope::Workbook,
            NameDefinition::Reference("Sheet1!B1".to_string()),
        )
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "A1", "=LET(X,1,Sheet1!X)")
        .unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(10.0));

    // Updating a precedent referenced from the defined name should dirty the cell even though a LET
    // binding shadows the unqualified identifier.
    engine.set_cell_value("Sheet1", "B1", 11.0).unwrap();
    assert!(
        engine.is_dirty("Sheet1", "A1"),
        "sheet-qualified name references should still track defined-name precedents"
    );
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(11.0));

    // Updating the defined name should dirty the cell and repoint dependencies.
    engine
        .define_name(
            "X",
            NameScope::Workbook,
            NameDefinition::Reference("Sheet1!B2".to_string()),
        )
        .unwrap();
    assert!(
        engine.is_dirty("Sheet1", "A1"),
        "changing a defined name should dirty dependents even if a LET binding shadows the identifier"
    );
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(20.0));

    // Old precedents should no longer dirty the cell...
    engine.set_cell_value("Sheet1", "B1", 12.0).unwrap();
    assert!(
        !engine.is_dirty("Sheet1", "A1"),
        "dependencies should update after the defined name is repointed"
    );

    // ...but the new precedent should.
    engine.set_cell_value("Sheet1", "B2", 21.0).unwrap();
    assert!(engine.is_dirty("Sheet1", "A1"));
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(21.0));
}

#[test]
fn let_and_lambda_locals_do_not_surface_external_precedents() {
    let mut engine = Engine::new();

    // Define a workbook-scoped name `X` that points at an external workbook cell.
    // LET/LAMBDA locals named `X` should *not* treat that as a defined-name reference.
    engine
        .define_name(
            "X",
            NameScope::Workbook,
            NameDefinition::Reference("[Book.xlsx]Sheet1!A1".to_string()),
        )
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "A1", "=LET(X,1,X+1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=LAMBDA(X,X+1)(1)")
        .unwrap();

    for addr in ["A1", "A2"] {
        let precedents = engine.precedents("Sheet1", addr).unwrap();
        assert!(
            precedents.iter().all(|p| {
                !matches!(
                    p,
                    PrecedentNode::ExternalCell { .. } | PrecedentNode::ExternalRange { .. }
                )
            }),
            "expected no external precedents for {addr}, got {precedents:?}"
        );
    }
}

#[test]
fn let_locals_do_not_shadow_sheet_qualified_external_precedents() {
    let mut engine = Engine::new();

    engine
        .define_name(
            "X",
            NameScope::Workbook,
            NameDefinition::Reference("[Book.xlsx]Sheet1!A1".to_string()),
        )
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "A1", "=LET(X,1,Sheet1!X)")
        .unwrap();

    let precedents = engine.precedents("Sheet1", "A1").unwrap();
    assert!(
        precedents.iter().any(|p| matches!(
            p,
            PrecedentNode::ExternalCell { sheet, addr } if sheet == "[Book.xlsx]Sheet1" && addr.row == 0 && addr.col == 0
        )),
        "expected external precedent for [Book.xlsx]Sheet1!A1, got {precedents:?}"
    );
}

#[test]
fn named_lambda_calls_surface_external_precedents() {
    let mut engine = Engine::new();
    engine
        .define_name(
            "ADD_EXT",
            NameScope::Workbook,
            NameDefinition::Formula("=LAMBDA(n,[Book.xlsx]Sheet1!A1+n)".to_string()),
        )
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "A1", "=ADD_EXT(1)")
        .unwrap();

    let precedents = engine.precedents("Sheet1", "A1").unwrap();
    assert!(
        precedents.iter().any(|p| matches!(
            p,
            PrecedentNode::ExternalCell { sheet, addr } if sheet == "[Book.xlsx]Sheet1" && addr.row == 0 && addr.col == 0
        )),
        "expected external precedent for [Book.xlsx]Sheet1!A1, got {precedents:?}"
    );
}
