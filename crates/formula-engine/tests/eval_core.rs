use formula_engine::eval::parse_a1;
use formula_engine::value::EntityValue;
use formula_engine::{Engine, ErrorKind, Value};
use pretty_assertions::assert_eq;

#[test]
fn compares_rich_values_as_text_case_insensitive() {
    fn assert_mode(bytecode: bool) {
        let mut engine = Engine::new();
        engine.set_bytecode_enabled(bytecode);
        engine
            .set_cell_value("Sheet1", "A1", Value::Entity(EntityValue::new("Apple")))
            .unwrap();
        engine
            .set_cell_formula("Sheet1", "B1", "=A1=\"apple\"")
            .unwrap();
        engine.recalculate();
        assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Bool(true));
    }

    // Cover both evaluator and bytecode runtime compare paths.
    assert_mode(false);
    assert_mode(true);
}

#[test]
fn error_propagation_and_short_circuiting() {
    let mut engine = Engine::new();

    engine.set_cell_formula("Sheet1", "A1", "=1/0").unwrap();
    engine.recalculate();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::Div0)
    );

    engine.set_cell_formula("Sheet1", "B1", "=A1+1").unwrap();
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
    engine.set_cell_formula("Sheet1", "B3", "=SUM(A2)").unwrap();
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
    engine.set_cell_formula("Sheet1", "B1", "=A1+A2").unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(3.0));

    engine.set_cell_value("Sheet1", "A1", 10.0).unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(12.0));

    // Change the formula; dependency graph should update.
    engine.set_cell_formula("Sheet1", "B1", "=A1*2").unwrap();
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
        engine.set_cell_formula("Sheet1", "B1", "=A1+A2").unwrap();
        engine.set_cell_formula("Sheet1", "B2", "=A1*A2").unwrap();
        engine.set_cell_formula("Sheet1", "C1", "=B1+B2").unwrap();
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

    assert_eq!(
        multi.get_cell_value("Sheet1", "B1"),
        single.get_cell_value("Sheet1", "B1")
    );
    assert_eq!(
        multi.get_cell_value("Sheet1", "B2"),
        single.get_cell_value("Sheet1", "B2")
    );
    assert_eq!(
        multi.get_cell_value("Sheet1", "C1"),
        single.get_cell_value("Sheet1", "C1")
    );
}

#[test]
fn evaluates_selected_financial_functions() {
    let mut engine = Engine::new();

    engine
        .set_cell_formula("Sheet1", "A1", "=PV(0, 3, -10)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=PMT(0, 2, 10)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A3", "=SLN(30, 0, 3)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A4", "=SLN(30, 0, 0)")
        .unwrap();

    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(30.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(-5.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A3"), Value::Number(10.0));
    assert_eq!(
        engine.get_cell_value("Sheet1", "A4"),
        Value::Error(ErrorKind::Div0)
    );
}

#[test]
fn evaluates_cashflow_functions_with_ranges() {
    let mut engine = Engine::new();

    // NPV matches the scalar-vs-reference coercion quirk used by SUM.
    engine.set_cell_value("Sheet1", "A1", "5").unwrap();
    engine.set_cell_value("Sheet1", "A2", true).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", r#"=NPV(0, "5", TRUE, 3)"#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=NPV(0, A1:A3)")
        .unwrap();

    // IRR/XNPV/XIRR examples mirror the existing numeric unit tests in
    // `tests/functions/financial_cashflows.rs`.
    let cashflows = [-70_000.0, 12_000.0, 15_000.0, 18_000.0, 21_000.0, 26_000.0];
    for (i, v) in cashflows.iter().enumerate() {
        let addr = format!("C{}", i + 1);
        engine.set_cell_value("Sheet1", &addr, *v).unwrap();
    }
    engine
        .set_cell_formula("Sheet1", "D1", "=IRR(C1:C6)")
        .unwrap();

    let values = [-10_000.0, 2_750.0, 4_250.0, 3_250.0, 2_750.0];
    let dates = [39448.0, 39508.0, 39751.0, 39859.0, 39904.0];
    for (i, (v, d)) in values.iter().zip(dates.iter()).enumerate() {
        let v_addr = format!("E{}", i + 1);
        let d_addr = format!("F{}", i + 1);
        engine.set_cell_value("Sheet1", &v_addr, *v).unwrap();
        engine.set_cell_value("Sheet1", &d_addr, *d).unwrap();
    }
    engine
        .set_cell_formula("Sheet1", "G1", "=XNPV(0.09, E1:E5, F1:F5)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "G2", "=XIRR(E1:E5, F1:F5)")
        .unwrap();

    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(9.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(3.0));

    let irr = match engine.get_cell_value("Sheet1", "D1") {
        Value::Number(n) => n,
        other => panic!("expected IRR numeric result, got {other:?}"),
    };
    assert!((irr - 0.08663094803653162).abs() <= 1e-12);

    let xnpv = match engine.get_cell_value("Sheet1", "G1") {
        Value::Number(n) => n,
        other => panic!("expected XNPV numeric result, got {other:?}"),
    };
    assert!((xnpv - 2_086.6476020315354).abs() <= 1e-10);

    let xirr = match engine.get_cell_value("Sheet1", "G2") {
        Value::Number(n) => n,
        other => panic!("expected XIRR numeric result, got {other:?}"),
    };
    assert!((xirr - 0.3733625335188314).abs() <= 1e-12);
}

#[test]
fn irr_treats_non_numeric_cells_as_zero_in_ranges() {
    let mut engine = Engine::new();

    // Text/logical/blank entries are ignored by IRR when supplied via references,
    // but their positions still count as periods (i.e. they contribute 0 at that
    // period).
    engine.set_cell_value("Sheet1", "A1", -1000.0).unwrap();
    engine
        .set_cell_value("Sheet1", "A2", "not a number")
        .unwrap();
    engine.set_cell_value("Sheet1", "A3", true).unwrap();
    engine.set_cell_value("Sheet1", "A4", 1100.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", "=IRR(A1:A4)")
        .unwrap();

    engine.recalculate();

    let irr = match engine.get_cell_value("Sheet1", "B1") {
        Value::Number(n) => n,
        other => panic!("expected IRR numeric result, got {other:?}"),
    };

    // Equivalent cashflows are [-1000, 0, 0, 1100]:
    // -1000 + 1100 / (1 + r)^3 = 0 => r = 1.1^(1/3) - 1.
    let expected = 1.1_f64.powf(1.0 / 3.0) - 1.0;
    assert!(
        (irr - expected).abs() <= 1e-12,
        "expected {expected}, got {irr}"
    );
}

#[test]
fn npv_preserves_period_index_for_blank_cells_in_ranges() {
    let mut engine = Engine::new();

    // Leaving a cell blank inside the values range should act like a 0 cashflow for
    // that period (it should not "compress" subsequent cashflows earlier).
    engine.set_cell_value("Sheet1", "A1", 100.0).unwrap();
    // A2 intentionally left blank.
    engine.set_cell_value("Sheet1", "A3", 100.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", "=NPV(0.1, A1:A3)")
        .unwrap();

    engine.recalculate();

    let result = match engine.get_cell_value("Sheet1", "B1") {
        Value::Number(n) => n,
        other => panic!("expected NPV numeric result, got {other:?}"),
    };

    let expected = 100.0 / 1.1 + 0.0 / 1.1_f64.powi(2) + 100.0 / 1.1_f64.powi(3);
    assert!(
        (result - expected).abs() <= 1e-12,
        "expected {expected}, got {result}"
    );
}

#[test]
fn exponentiation_operator_matches_excel_precedence_and_associativity() {
    let mut engine = Engine::new();

    engine.set_cell_formula("Sheet1", "A1", "=2^3").unwrap();
    engine.set_cell_formula("Sheet1", "A2", "=-2^2").unwrap();
    engine.set_cell_formula("Sheet1", "A3", "=(-2)^2").unwrap();
    engine.set_cell_formula("Sheet1", "A4", "=2^3^2").unwrap();

    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(8.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(-4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A3"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A4"), Value::Number(512.0));
}

#[test]
fn array_functions_spill_and_respect_spill_blocking() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "C1", 3.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "D1", "=TRANSPOSE(A1:C1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "D2", "=_xlfn.SEQUENCE(2,2,1,1)")
        .unwrap();
    // A second, non-overlapping transpose should spill successfully even when another
    // array formula is present elsewhere on the sheet.
    engine
        .set_cell_formula("Sheet1", "F1", "=TRANSPOSE(A1:C1)")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.spill_range("Sheet1", "D1"), None,);
    assert_eq!(
        engine.get_cell_value("Sheet1", "D1"),
        Value::Error(ErrorKind::Spill)
    );

    // The D2 formula blocks D1's spill attempt, but D2's own dynamic array spills successfully.
    // SEQUENCE spills a 2x2 matrix starting at D2.
    assert_eq!(
        engine.spill_range("Sheet1", "D2"),
        Some((parse_a1("D2").unwrap(), parse_a1("E3").unwrap()))
    );
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E2"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D3"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E3"), Value::Number(4.0));

    // TRANSPOSE spills down when unblocked.
    assert_eq!(
        engine.spill_range("Sheet1", "F1"),
        Some((parse_a1("F1").unwrap(), parse_a1("F3").unwrap()))
    );
    assert_eq!(engine.get_cell_value("Sheet1", "F1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F2"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F3"), Value::Number(3.0));
}
