use formula_engine::{Engine, ErrorKind, Value};
use pretty_assertions::assert_eq;

#[test]
fn error_propagation_and_short_circuiting() {
    let mut engine = Engine::new();

    engine
        .set_cell_formula("Sheet1", "A1", "=1/0")
        .unwrap();
    engine.recalculate();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::Div0)
    );

    engine
        .set_cell_formula("Sheet1", "B1", "=A1+1")
        .unwrap();
    engine.recalculate();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Error(ErrorKind::Div0)
    );

    // IF must short-circuit non-selected branches.
    engine
        .set_cell_formula("Sheet1", "C1", "=IF(TRUE, 1, 1/0)")
        .unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(1.0));

    // IFERROR must also short-circuit the fallback.
    engine
        .set_cell_formula("Sheet1", "D1", "=IFERROR(1/0, 5)")
        .unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(5.0));

    // ISERROR swallows errors and returns TRUE/FALSE.
    engine
        .set_cell_formula("Sheet1", "E1", "=ISERROR(1/0)")
        .unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Bool(true));
}

#[test]
fn sum_coercion_scalar_vs_reference_args() {
    let mut engine = Engine::new();

    engine.set_cell_value("Sheet1", "A1", "5").unwrap();
    engine.set_cell_value("Sheet1", "A2", true).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", r#"=SUM("5", TRUE, 3)"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=SUM(A1:A3)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B3", "=SUM(A2)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B4", "=SUM(TRUE)")
        .unwrap();

    engine.recalculate();

    // Scalar args: "5" -> 5, TRUE -> 1.
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(9.0));

    // Reference args: text and logicals are ignored, only numeric cells are included.
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B3"), Value::Number(0.0));

    // Literal TRUE is treated as 1.
    assert_eq!(engine.get_cell_value("Sheet1", "B4"), Value::Number(1.0));
}

#[test]
fn sheet_and_quoted_sheet_references() {
    let mut engine = Engine::new();

    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("My Sheet", "A1", 10.0).unwrap();

    engine
        .set_cell_formula("Sheet2", "B1", "=Sheet1!A1+1")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=Sheet2!B1*2")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "D1", "='My Sheet'!A1+1")
        .unwrap();

    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet2", "B1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(11.0));
}

#[test]
fn incremental_recalc_updates_only_affected_cells() {
    let mut engine = Engine::new();

    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=A1+A2")
        .unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(3.0));

    engine.set_cell_value("Sheet1", "A1", 10.0).unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(12.0));

    // Change the formula; dependency graph should update.
    engine
        .set_cell_formula("Sheet1", "B1", "=A1*2")
        .unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(20.0));

    // A2 is no longer a precedent of B1.
    engine.set_cell_value("Sheet1", "A2", 100.0).unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(20.0));
}

#[test]
fn multithreaded_recalc_matches_single_threaded() {
    fn setup(engine: &mut Engine) {
        engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
        engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
        engine
            .set_cell_formula("Sheet1", "B1", "=A1+A2")
            .unwrap();
        engine
            .set_cell_formula("Sheet1", "B2", "=A1*A2")
            .unwrap();
        engine
            .set_cell_formula("Sheet1", "C1", "=B1+B2")
            .unwrap();
    }

    let mut single = Engine::new();
    setup(&mut single);
    single.recalculate_single_threaded();

    let mut multi = Engine::new();
    setup(&mut multi);
    multi.recalculate_multi_threaded();

    assert_eq!(single.get_cell_value("Sheet1", "B1"), Value::Number(3.0));
    assert_eq!(single.get_cell_value("Sheet1", "B2"), Value::Number(2.0));
    assert_eq!(single.get_cell_value("Sheet1", "C1"), Value::Number(5.0));

    assert_eq!(multi.get_cell_value("Sheet1", "B1"), single.get_cell_value("Sheet1", "B1"));
    assert_eq!(multi.get_cell_value("Sheet1", "B2"), single.get_cell_value("Sheet1", "B2"));
    assert_eq!(multi.get_cell_value("Sheet1", "C1"), single.get_cell_value("Sheet1", "C1"));
}

