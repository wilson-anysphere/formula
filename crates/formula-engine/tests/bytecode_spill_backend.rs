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

fn setup_info_engine(bytecode_enabled: bool) -> Engine {
    let mut engine = Engine::new();
    engine.set_bytecode_enabled(bytecode_enabled);
    // A1 left blank.
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", "x").unwrap();
    engine
}

fn setup_error_engine(bytecode_enabled: bool) -> Engine {
    let mut engine = Engine::new();
    engine.set_bytecode_enabled(bytecode_enabled);
    engine
        .set_cell_value("Sheet1", "A1", Value::Error(ErrorKind::Div0))
        .unwrap();
    engine
        .set_cell_value("Sheet1", "A2", Value::Error(ErrorKind::NA))
        .unwrap();
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

#[test]
fn bytecode_spills_match_ast_for_information_functions_over_ranges() {
    for (formula, expected) in [
        (
            "=ISBLANK(A1:A3)",
            vec![Value::Bool(true), Value::Bool(false), Value::Bool(false)],
        ),
        (
            "=ISNUMBER(A1:A3)",
            vec![Value::Bool(false), Value::Bool(true), Value::Bool(false)],
        ),
        (
            "=ISTEXT(A1:A3)",
            vec![Value::Bool(false), Value::Bool(false), Value::Bool(true)],
        ),
        (
            "=N(A1:A3)",
            vec![Value::Number(0.0), Value::Number(2.0), Value::Number(0.0)],
        ),
        (
            "=T(A1:A3)",
            vec![
                Value::Text(String::new()),
                Value::Text(String::new()),
                Value::Text("x".to_string()),
            ],
        ),
    ] {
        let mut ast = setup_info_engine(false);
        ast.set_cell_formula("Sheet1", "C1", formula).unwrap();
        ast.recalculate_single_threaded();

        let mut bytecode = setup_info_engine(true);
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
fn bytecode_spills_match_ast_for_iserror_and_isna_over_ranges() {
    for (formula, expected) in [
        (
            "=ISERROR(A1:A3)",
            vec![Value::Bool(true), Value::Bool(true), Value::Bool(false)],
        ),
        (
            "=ISNA(A1:A3)",
            vec![Value::Bool(false), Value::Bool(true), Value::Bool(false)],
        ),
    ] {
        let mut ast = setup_error_engine(false);
        ast.set_cell_formula("Sheet1", "C1", formula).unwrap();
        ast.recalculate_single_threaded();

        let mut bytecode = setup_error_engine(true);
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
fn bytecode_spills_match_ast_for_row_and_column_reference_functions() {
    fn run(bytecode_enabled: bool) -> Engine {
        let mut engine = Engine::new();
        engine.set_bytecode_enabled(bytecode_enabled);
        engine.set_cell_formula("Sheet1", "A1", "=ROW(D4:F5)").unwrap();
        engine
            .set_cell_formula("Sheet1", "E1", "=COLUMN(D4:F5)")
            .unwrap();
        engine.set_cell_formula("Sheet1", "J1", "=ROW(5:7)").unwrap();
        engine.set_cell_formula("Sheet1", "A10", "=COLUMN(D:F)").unwrap();
        engine.recalculate_single_threaded();
        engine
    }

    let ast = run(false);
    let bytecode = run(true);
    assert_eq!(
        bytecode.bytecode_program_count(),
        4,
        "expected row/column formulas to compile to bytecode"
    );

    // ROW(D4:F5) -> {4,4,4;5,5,5} spills to A1:C2.
    assert_eq!(
        bytecode.spill_range("Sheet1", "A1"),
        Some((parse_a1("A1").unwrap(), parse_a1("C2").unwrap()))
    );
    assert_eq!(bytecode.get_cell_value("Sheet1", "A1"), Value::Number(4.0));
    assert_eq!(bytecode.get_cell_value("Sheet1", "B1"), Value::Number(4.0));
    assert_eq!(bytecode.get_cell_value("Sheet1", "C1"), Value::Number(4.0));
    assert_eq!(bytecode.get_cell_value("Sheet1", "A2"), Value::Number(5.0));
    assert_eq!(bytecode.get_cell_value("Sheet1", "B2"), Value::Number(5.0));
    assert_eq!(bytecode.get_cell_value("Sheet1", "C2"), Value::Number(5.0));

    // COLUMN(D4:F5) -> {4,5,6;4,5,6} spills to E1:G2.
    assert_eq!(
        bytecode.spill_range("Sheet1", "E1"),
        Some((parse_a1("E1").unwrap(), parse_a1("G2").unwrap()))
    );
    assert_eq!(bytecode.get_cell_value("Sheet1", "E1"), Value::Number(4.0));
    assert_eq!(bytecode.get_cell_value("Sheet1", "F1"), Value::Number(5.0));
    assert_eq!(bytecode.get_cell_value("Sheet1", "G1"), Value::Number(6.0));
    assert_eq!(bytecode.get_cell_value("Sheet1", "E2"), Value::Number(4.0));
    assert_eq!(bytecode.get_cell_value("Sheet1", "F2"), Value::Number(5.0));
    assert_eq!(bytecode.get_cell_value("Sheet1", "G2"), Value::Number(6.0));

    // ROW(5:7) -> {5;6;7} spills to J1:J3.
    assert_eq!(
        bytecode.spill_range("Sheet1", "J1"),
        Some((parse_a1("J1").unwrap(), parse_a1("J3").unwrap()))
    );
    assert_eq!(bytecode.get_cell_value("Sheet1", "J1"), Value::Number(5.0));
    assert_eq!(bytecode.get_cell_value("Sheet1", "J2"), Value::Number(6.0));
    assert_eq!(bytecode.get_cell_value("Sheet1", "J3"), Value::Number(7.0));

    // COLUMN(D:F) -> {4,5,6} spills to A10:C10.
    assert_eq!(
        bytecode.spill_range("Sheet1", "A10"),
        Some((parse_a1("A10").unwrap(), parse_a1("C10").unwrap()))
    );
    assert_eq!(bytecode.get_cell_value("Sheet1", "A10"), Value::Number(4.0));
    assert_eq!(bytecode.get_cell_value("Sheet1", "B10"), Value::Number(5.0));
    assert_eq!(bytecode.get_cell_value("Sheet1", "C10"), Value::Number(6.0));

    // Sanity check: compare a few representative cells against the AST backend.
    for addr in ["A1", "B2", "F1", "J3", "B10"] {
        assert_eq!(
            bytecode.get_cell_value("Sheet1", addr),
            ast.get_cell_value("Sheet1", addr),
            "bytecode mismatch at {addr}"
        );
    }
    for origin in ["A1", "E1", "J1", "A10"] {
        assert_eq!(
            bytecode.spill_range("Sheet1", origin),
            ast.spill_range("Sheet1", origin),
            "spill range mismatch at {origin}"
        );
    }
}
