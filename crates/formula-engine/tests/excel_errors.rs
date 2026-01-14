use formula_engine::{Engine, ErrorKind, Value};

fn all_errors() -> &'static [(ErrorKind, &'static str)] {
    &[
        (ErrorKind::Null, "#NULL!"),
        (ErrorKind::Div0, "#DIV/0!"),
        (ErrorKind::Value, "#VALUE!"),
        (ErrorKind::Ref, "#REF!"),
        (ErrorKind::Name, "#NAME?"),
        (ErrorKind::Num, "#NUM!"),
        (ErrorKind::NA, "#N/A"),
        (ErrorKind::GettingData, "#GETTING_DATA"),
        (ErrorKind::Spill, "#SPILL!"),
        (ErrorKind::Calc, "#CALC!"),
        (ErrorKind::Field, "#FIELD!"),
        (ErrorKind::Connect, "#CONNECT!"),
        (ErrorKind::Blocked, "#BLOCKED!"),
        (ErrorKind::Unknown, "#UNKNOWN!"),
    ]
}

#[test]
fn engine_evaluates_all_error_literals_ast_and_bytecode() {
    for bytecode_enabled in [false, true] {
        let mut engine = Engine::new();
        engine.set_bytecode_enabled(bytecode_enabled);

        for (idx, (kind, lit)) in all_errors().iter().enumerate() {
            let addr = format!("A{}", idx + 1);
            let formula = format!("={}", lit.to_ascii_lowercase());
            engine.set_cell_formula("Sheet1", &addr, &formula).unwrap();

            // Ensure the literal survives round-trip/display.
            assert_eq!(kind.as_code(), *lit);
        }

        if bytecode_enabled {
            // Error literals should be eligible for bytecode compilation.
            assert_eq!(engine.bytecode_program_count(), all_errors().len());
        } else {
            assert_eq!(engine.bytecode_program_count(), 0);
        }

        engine.recalculate_single_threaded();

        for (idx, (kind, _lit)) in all_errors().iter().enumerate() {
            let addr = format!("A{}", idx + 1);
            assert_eq!(engine.get_cell_value("Sheet1", &addr), Value::Error(*kind));
        }
    }
}

#[test]
fn error_type_returns_excel_codes_for_all_errors() {
    let mut engine = Engine::new();
    // ERROR.TYPE is not supported by the bytecode evaluator today; force AST evaluation so this
    // test stays focused on the mapping itself.
    engine.set_bytecode_enabled(false);

    for (idx, (kind, lit)) in all_errors().iter().enumerate() {
        let addr = format!("A{}", idx + 1);
        let formula = format!("=ERROR.TYPE({lit})");
        engine.set_cell_formula("Sheet1", &addr, &formula).unwrap();

        assert!(kind.code() > 0);
    }

    engine.recalculate_single_threaded();

    for (idx, (kind, _lit)) in all_errors().iter().enumerate() {
        let addr = format!("A{}", idx + 1);
        assert_eq!(
            engine.get_cell_value("Sheet1", &addr),
            Value::Number(kind.code() as f64)
        );
    }
}

#[test]
fn error_kind_from_code_accepts_na_exclamation_alias() {
    assert_eq!(ErrorKind::from_code("#N/A!"), Some(ErrorKind::NA));
    assert_eq!(ErrorKind::from_code("  #n/a!  "), Some(ErrorKind::NA));
}

#[test]
fn engine_evaluates_na_exclamation_error_literal_ast_and_bytecode() {
    for bytecode_enabled in [false, true] {
        let mut engine = Engine::new();
        engine.set_bytecode_enabled(bytecode_enabled);

        engine.set_cell_formula("Sheet1", "A1", "=#N/A!").unwrap();

        if bytecode_enabled {
            assert_eq!(engine.bytecode_program_count(), 1);
        } else {
            assert_eq!(engine.bytecode_program_count(), 0);
        }

        engine.recalculate_single_threaded();
        assert_eq!(
            engine.get_cell_value("Sheet1", "A1"),
            Value::Error(ErrorKind::NA)
        );
    }
}

#[test]
fn non_classic_errors_propagate_through_arithmetic_in_bytecode() {
    let mut engine = Engine::new();
    engine.set_bytecode_enabled(true);

    engine
        .set_cell_formula("Sheet1", "A1", "=#field!+1")
        .unwrap();

    // Ensure the formula was actually compiled to bytecode (regression guard).
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::Field)
    );
}
