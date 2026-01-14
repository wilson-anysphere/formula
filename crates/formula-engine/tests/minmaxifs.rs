use formula_engine::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};
use formula_engine::{Engine, ErrorKind, Value};

fn eval(engine: &mut Engine, formula: &str) -> Value {
    engine
        .set_cell_formula("Sheet1", "Z1", formula)
        .expect("set formula");
    engine.recalculate();
    engine.get_cell_value("Sheet1", "Z1")
}

#[test]
fn minmaxifs_basic_filtering() {
    let mut engine = Engine::new();

    for (addr, v) in [
        ("A1", 5.0),
        ("A2", 3.0),
        ("A3", 7.0),
        ("A4", 2.0),
        ("A5", 9.0),
    ] {
        engine.set_cell_value("Sheet1", addr, v).expect("set value");
    }

    for (addr, v) in [
        ("B1", "A"),
        ("B2", "B"),
        ("B3", "A"),
        ("B4", "B"),
        ("B5", "A"),
    ] {
        engine.set_cell_value("Sheet1", addr, v).expect("set value");
    }

    assert_eq!(
        eval(&mut engine, "=MINIFS(A1:A5,B1:B5,\"A\")"),
        Value::Number(5.0)
    );
    assert_eq!(
        eval(&mut engine, "=MAXIFS(A1:A5,B1:B5,\"B\")"),
        Value::Number(3.0)
    );
}

#[test]
fn minifs_numeric_criteria_does_not_treat_text_as_zero() {
    let mut engine = Engine::new();

    engine.set_cell_value("Sheet1", "A1", "x").unwrap();
    engine.set_cell_value("Sheet1", "A2", 0.0).unwrap();
    // A3 left unset (blank) -> treated as 0 for numeric criteria.

    engine.set_cell_value("Sheet1", "B1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "B3", 20.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "Z1", "=MINIFS(B1:B3,A1:A3,0)")
        .unwrap();
    assert!(
        engine.bytecode_program_count() > 0,
        "expected MINIFS formula to compile to bytecode for this test"
    );
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "Z1"), Value::Number(10.0));
}

#[test]
fn maxifs_numeric_criteria_does_not_treat_text_as_zero() {
    let mut engine = Engine::new();

    engine.set_cell_value("Sheet1", "A1", "x").unwrap();
    engine.set_cell_value("Sheet1", "A2", 0.0).unwrap();
    // A3 left unset (blank) -> treated as 0 for numeric criteria.

    engine.set_cell_value("Sheet1", "B1", 100.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "B3", 20.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "Z1", "=MAXIFS(B1:B3,A1:A3,0)")
        .unwrap();
    assert!(
        engine.bytecode_program_count() > 0,
        "expected MAXIFS formula to compile to bytecode for this test"
    );
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "Z1"), Value::Number(20.0));
}

#[test]
fn minmaxifs_multiple_criteria() {
    let mut engine = Engine::new();

    for (addr, v) in [
        ("A1", 5.0),
        ("A2", 3.0),
        ("A3", 7.0),
        ("A4", 2.0),
        ("A5", 9.0),
    ] {
        engine.set_cell_value("Sheet1", addr, v).expect("set value");
    }
    for (addr, v) in [
        ("B1", "A"),
        ("B2", "B"),
        ("B3", "A"),
        ("B4", "B"),
        ("B5", "A"),
    ] {
        engine.set_cell_value("Sheet1", addr, v).expect("set value");
    }
    for (addr, v) in [
        ("C1", 1.0),
        ("C2", 2.0),
        ("C3", 3.0),
        ("C4", 4.0),
        ("C5", 5.0),
    ] {
        engine.set_cell_value("Sheet1", addr, v).expect("set value");
    }

    assert_eq!(
        eval(&mut engine, "=MINIFS(A1:A5,B1:B5,\"A\",C1:C5,\">2\")"),
        Value::Number(7.0)
    );
    assert_eq!(
        eval(&mut engine, "=MAXIFS(A1:A5,B1:B5,\"A\",C1:C5,\"<5\")"),
        Value::Number(7.0)
    );
}

#[test]
fn minmaxifs_ignore_non_numeric_targets() {
    let mut engine = Engine::new();

    engine
        .set_cell_value("Sheet1", "A1", 10.0)
        .expect("set value");
    engine
        .set_cell_value("Sheet1", "A2", "x")
        .expect("set value");
    engine
        .set_cell_value("Sheet1", "A3", 5.0)
        .expect("set value");
    engine
        .set_cell_value("Sheet1", "A4", true)
        .expect("set value");

    for addr in ["B1", "B2", "B3", "B4"] {
        engine
            .set_cell_value("Sheet1", addr, "A")
            .expect("set value");
    }

    assert_eq!(
        eval(&mut engine, "=MINIFS(A1:A4,B1:B4,\"A\")"),
        Value::Number(5.0)
    );
    assert_eq!(
        eval(&mut engine, "=MAXIFS(A1:A4,B1:B4,\"A\")"),
        Value::Number(10.0)
    );
}

#[test]
fn minmaxifs_error_propagation_only_when_included() {
    let mut engine = Engine::new();

    engine
        .set_cell_value("Sheet1", "A1", 10.0)
        .expect("set value");
    engine
        .set_cell_value("Sheet1", "A2", Value::Error(ErrorKind::Div0))
        .expect("set value");
    engine
        .set_cell_value("Sheet1", "A3", 5.0)
        .expect("set value");

    engine
        .set_cell_value("Sheet1", "B1", 1.0)
        .expect("set value");
    engine
        .set_cell_value("Sheet1", "B2", 2.0)
        .expect("set value");
    engine
        .set_cell_value("Sheet1", "B3", 3.0)
        .expect("set value");

    // Excludes the error row, so it should not propagate.
    assert_eq!(
        eval(&mut engine, "=MINIFS(A1:A3,B1:B3,\">2\")"),
        Value::Number(5.0)
    );

    // Includes the error row, so it should propagate.
    assert_eq!(
        eval(&mut engine, "=MAXIFS(A1:A3,B1:B3,\">1\")"),
        Value::Error(ErrorKind::Div0)
    );
}

#[test]
fn minmaxifs_empty_result_returns_zero() {
    let mut engine = Engine::new();

    engine
        .set_cell_value("Sheet1", "A1", 10.0)
        .expect("set value");
    engine
        .set_cell_value("Sheet1", "A2", 20.0)
        .expect("set value");
    engine
        .set_cell_value("Sheet1", "B1", "A")
        .expect("set value");
    engine
        .set_cell_value("Sheet1", "B2", "B")
        .expect("set value");

    assert_eq!(
        eval(&mut engine, "=MINIFS(A1:A2,B1:B2,\"C\")"),
        Value::Number(0.0)
    );
    assert_eq!(
        eval(&mut engine, "=MAXIFS(A1:A2,B1:B2,\"C\")"),
        Value::Number(0.0)
    );
}

#[test]
fn minmaxifs_wildcard_criteria() {
    let mut engine = Engine::new();

    engine
        .set_cell_value("Sheet1", "A1", 1.0)
        .expect("set value");
    engine
        .set_cell_value("Sheet1", "A2", 2.0)
        .expect("set value");
    engine
        .set_cell_value("Sheet1", "A3", 3.0)
        .expect("set value");

    engine
        .set_cell_value("Sheet1", "B1", "apple")
        .expect("set value");
    engine
        .set_cell_value("Sheet1", "B2", "apricot")
        .expect("set value");
    engine
        .set_cell_value("Sheet1", "B3", "banana")
        .expect("set value");

    assert_eq!(
        eval(&mut engine, "=MINIFS(A1:A3,B1:B3,\"ap*\")"),
        Value::Number(1.0)
    );
    assert_eq!(
        eval(&mut engine, "=MAXIFS(A1:A3,B1:B3,\"ap*\")"),
        Value::Number(2.0)
    );
}

#[test]
fn minmaxifs_date_criteria_is_parsed_with_date_system() {
    let mut engine = Engine::new();

    let system = ExcelDateSystem::EXCEL_1900;
    let jan1 = ymd_to_serial(ExcelDate::new(2020, 1, 1), system).expect("date serial") as f64;
    let jan2 = ymd_to_serial(ExcelDate::new(2020, 1, 2), system).expect("date serial") as f64;
    let jan3 = ymd_to_serial(ExcelDate::new(2020, 1, 3), system).expect("date serial") as f64;

    engine
        .set_cell_value("Sheet1", "A1", jan1)
        .expect("set value");
    engine
        .set_cell_value("Sheet1", "A2", jan2)
        .expect("set value");
    engine
        .set_cell_value("Sheet1", "A3", jan3)
        .expect("set value");

    engine
        .set_cell_value("Sheet1", "B1", 10.0)
        .expect("set value");
    engine
        .set_cell_value("Sheet1", "B2", 5.0)
        .expect("set value");
    engine
        .set_cell_value("Sheet1", "B3", 7.0)
        .expect("set value");

    assert_eq!(
        eval(&mut engine, "=MINIFS(B1:B3,A1:A3,\">=1/2/2020\")"),
        Value::Number(5.0)
    );
    assert_eq!(
        eval(&mut engine, "=MAXIFS(B1:B3,A1:A3,\">=1/2/2020\")"),
        Value::Number(7.0)
    );
}

#[test]
fn minmaxifs_validates_arg_shape_and_pairs() {
    let mut engine = Engine::new();

    engine
        .set_cell_value("Sheet1", "A1", 1.0)
        .expect("set value");
    engine
        .set_cell_value("Sheet1", "A2", 2.0)
        .expect("set value");
    engine
        .set_cell_value("Sheet1", "B1", "A")
        .expect("set value");
    engine
        .set_cell_value("Sheet1", "B2", "B")
        .expect("set value");

    // Missing criteria for the second criteria-range argument.
    assert_eq!(
        eval(&mut engine, "=MINIFS(A1:A2,B1:B2,\"A\",C1:C2)"),
        Value::Error(ErrorKind::Value)
    );

    // Same number of cells but different shapes (2x1 vs 1x2).
    assert_eq!(
        eval(&mut engine, "=MAXIFS(A1:A2,B1:C1,\"A\")"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn minmaxifs_accepts_reference_returning_functions() {
    let mut engine = Engine::new();

    for (addr, v) in [("A1", 5.0), ("A2", 3.0), ("A3", 7.0)] {
        engine.set_cell_value("Sheet1", addr, v).expect("set value");
    }

    for (addr, v) in [("B1", "A"), ("B2", "B"), ("B3", "A")] {
        engine.set_cell_value("Sheet1", addr, v).expect("set value");
    }

    // OFFSET returns a reference value; MINIFS/MAXIFS should accept those wherever a range is expected.
    assert_eq!(
        eval(
            &mut engine,
            "=MINIFS(OFFSET(A1,0,0,3,1),OFFSET(B1,0,0,3,1),\"A\")"
        ),
        Value::Number(5.0)
    );

    // Errors returned from reference-returning functions should propagate as-is.
    assert_eq!(
        eval(&mut engine, "=MAXIFS(OFFSET(A1,1048576,0,1,1),B1:B3,\"A\")"),
        Value::Error(ErrorKind::Ref)
    );
}

#[test]
fn minmaxifs_error_precedence_is_row_major_for_sparse_iteration() {
    let mut engine = Engine::new();

    // Criteria range: include everything.
    for (addr, v) in [("D1", 1.0), ("E1", 1.0), ("D2", 1.0), ("E2", 1.0)] {
        engine.set_cell_value("Sheet1", addr, v).expect("set value");
    }

    // Target range has multiple errors. We expect the *first* error in row-major order
    // (B1, C1, B2, C2) to win, even if the sparse iterator yields cells in hash order.
    engine
        .set_cell_value("Sheet1", "B1", 10.0)
        .expect("set value");
    engine
        .set_cell_value("Sheet1", "C1", Value::Error(ErrorKind::Ref))
        .expect("set value");
    engine
        .set_cell_value("Sheet1", "B2", Value::Error(ErrorKind::Div0))
        .expect("set value");
    engine
        .set_cell_value("Sheet1", "C2", 5.0)
        .expect("set value");

    assert_eq!(
        eval(&mut engine, r#"=MINIFS(B1:C2,D1:E2,">0")"#),
        Value::Error(ErrorKind::Ref)
    );
    assert_eq!(
        eval(&mut engine, r#"=MAXIFS(B1:C2,D1:E2,">0")"#),
        Value::Error(ErrorKind::Ref)
    );
}

#[test]
fn minmaxifs_accept_xlfn_prefix() {
    let mut engine = Engine::new();

    engine
        .set_cell_value("Sheet1", "A1", 5.0)
        .expect("set value");
    engine
        .set_cell_value("Sheet1", "A2", 3.0)
        .expect("set value");
    engine
        .set_cell_value("Sheet1", "A3", 7.0)
        .expect("set value");

    engine
        .set_cell_value("Sheet1", "B1", "A")
        .expect("set value");
    engine
        .set_cell_value("Sheet1", "B2", "B")
        .expect("set value");
    engine
        .set_cell_value("Sheet1", "B3", "A")
        .expect("set value");

    assert_eq!(
        eval(&mut engine, r#"=_xlfn.MINIFS(A1:A3,B1:B3,"A")"#),
        Value::Number(5.0)
    );
    assert_eq!(
        eval(&mut engine, r#"=_xlfn.MAXIFS(A1:A3,B1:B3,"A")"#),
        Value::Number(7.0)
    );
}
