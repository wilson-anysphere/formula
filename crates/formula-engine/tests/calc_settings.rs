use formula_engine::calc_settings::{CalcSettings, CalculationMode, IterativeCalculationSettings};
use formula_engine::{Engine, Value};

#[test]
fn manual_mode_marks_dirty_but_defers_recalc_until_user_triggers() {
    let mut engine = Engine::new();
    let mut settings = CalcSettings::default();
    settings.calculation_mode = CalculationMode::Manual;
    engine.set_calc_settings(settings);

    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_formula("Sheet1", "B1", "=A1+1").unwrap();

    // Manual mode: the formula should not be recalculated automatically.
    assert!(engine.has_dirty_cells());
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Blank);

    engine.recalculate();
    assert!(!engine.has_dirty_cells());
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(2.0));

    // Mutating a precedent should dirty dependents but not recalc.
    engine.set_cell_value("Sheet1", "A1", 10.0).unwrap();
    assert!(engine.has_dirty_cells());
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(2.0));

    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(11.0));
}

#[test]
fn iterative_calculation_converges_for_simple_cycle() {
    let mut engine = Engine::new();
    let mut settings = CalcSettings::default();
    settings.calculation_mode = CalculationMode::Manual;
    settings.iterative = IterativeCalculationSettings {
        enabled: true,
        max_iterations: 1000,
        max_change: 1e-9,
    };
    engine.set_calc_settings(settings);

    engine
        .set_cell_formula("Sheet1", "A1", "=(B1+1)/2")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=(A1+1)/2")
        .unwrap();

    engine.recalculate();

    let a = match engine.get_cell_value("Sheet1", "A1") {
        Value::Number(n) => n,
        other => panic!("expected numeric A1, got {other:?}"),
    };
    let b = match engine.get_cell_value("Sheet1", "B1") {
        Value::Number(n) => n,
        other => panic!("expected numeric B1, got {other:?}"),
    };

    assert!((a - 1.0).abs() < 1e-6, "A1={a}");
    assert!((b - 1.0).abs() < 1e-6, "B1={b}");
}

#[test]
fn non_iterative_cycles_set_zero_and_surface_circular_reference_warning() {
    let mut engine = Engine::new();
    let mut settings = CalcSettings::default();
    settings.calculation_mode = CalculationMode::Manual;
    settings.iterative.enabled = false;
    engine.set_calc_settings(settings);

    engine.set_cell_formula("Sheet1", "A1", "=B1").unwrap();
    engine.set_cell_formula("Sheet1", "B1", "=A1").unwrap();

    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(0.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(0.0));
    assert_eq!(engine.circular_reference_count(), 2);
}
