use formula_engine::eval::parse_a1;
use formula_engine::{Engine, ErrorKind, Value};

fn setup_base_engine(bytecode_enabled: bool) -> Engine {
    let mut engine = Engine::new();
    engine.set_bytecode_enabled(bytecode_enabled);
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();
    engine
}

#[test]
fn bytecode_spills_match_ast_for_range_reference_and_elementwise_ops() {
    for (formula, expected) in [
        ("=A1:A3", vec![Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]),
        ("=A1:A3+1", vec![Value::Number(2.0), Value::Number(3.0), Value::Number(4.0)]),
    ] {
        let mut ast = setup_base_engine(false);
        ast.set_cell_formula("Sheet1", "C1", formula).unwrap();
        ast.recalculate_single_threaded();

        let mut bytecode = setup_base_engine(true);
        bytecode.set_cell_formula("Sheet1", "C1", formula).unwrap();
        assert_eq!(
            bytecode.bytecode_program_count(),
            1,
            "expected formula {formula} to compile to bytecode"
        );
        bytecode.recalculate_single_threaded();

        assert_eq!(
            bytecode.spill_range("Sheet1", "C1"),
            Some((parse_a1("C1").unwrap(), parse_a1("C3").unwrap())),
            "expected spill range for bytecode formula {formula}"
        );
        assert_eq!(
            ast.spill_range("Sheet1", "C1"),
            Some((parse_a1("C1").unwrap(), parse_a1("C3").unwrap())),
            "expected spill range for AST formula {formula}"
        );

        for (addr, expected_value) in ["C1", "C2", "C3"].into_iter().zip(expected) {
            assert_eq!(
                bytecode.get_cell_value("Sheet1", addr),
                expected_value,
                "bytecode mismatch at {addr} for {formula}"
            );
            assert_eq!(
                ast.get_cell_value("Sheet1", addr),
                expected_value,
                "AST mismatch at {addr} for {formula}"
            );
        }
    }
}

#[test]
fn bytecode_blocked_spills_match_ast() {
    fn run(bytecode_enabled: bool) -> Engine {
        let mut engine = setup_base_engine(bytecode_enabled);
        engine.set_cell_formula("Sheet1", "C1", "=A1:A3").unwrap();
        engine.recalculate_single_threaded();

        // Block the middle spill cell with a user value.
        engine.set_cell_value("Sheet1", "C2", 99.0).unwrap();
        engine.recalculate_single_threaded();
        engine
    }

    let ast = run(false);
    let bytecode = run(true);
    assert_eq!(
        bytecode.bytecode_program_count(),
        1,
        "expected spill formula to compile to bytecode"
    );

    assert_eq!(bytecode.get_cell_value("Sheet1", "C1"), Value::Error(ErrorKind::Spill));
    assert_eq!(bytecode.get_cell_value("Sheet1", "C2"), Value::Number(99.0));
    assert_eq!(bytecode.get_cell_value("Sheet1", "C3"), Value::Blank);
    assert!(bytecode.spill_range("Sheet1", "C1").is_none());

    assert_eq!(bytecode.get_cell_value("Sheet1", "C1"), ast.get_cell_value("Sheet1", "C1"));
    assert_eq!(bytecode.get_cell_value("Sheet1", "C2"), ast.get_cell_value("Sheet1", "C2"));
    assert_eq!(bytecode.get_cell_value("Sheet1", "C3"), ast.get_cell_value("Sheet1", "C3"));
    assert_eq!(bytecode.spill_range("Sheet1", "C1"), ast.spill_range("Sheet1", "C1"));
}

