use formula_engine::calc_settings::CalcSettings;
use formula_engine::{Engine, Value};

#[test]
fn precision_as_displayed_rounds_numeric_literals_fixed_decimals() {
    let mut engine = Engine::new();
    let mut settings: CalcSettings = engine.calc_settings().clone();
    settings.full_precision = false;
    engine.set_calc_settings(settings);

    engine
        .set_cell_number_format("Sheet1", "A1", Some("0.00".to_string()))
        .unwrap();
    engine.set_cell_value("Sheet1", "A1", 1.239).unwrap();

    // The stored value should be rounded to match the displayed precision.
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.24));

    // Downstream formulas should observe the rounded stored value.
    engine.set_cell_formula("Sheet1", "B1", "=A1").unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.24));
}

#[test]
fn precision_as_displayed_rounds_numeric_literals_percent() {
    let mut engine = Engine::new();
    let mut settings: CalcSettings = engine.calc_settings().clone();
    settings.full_precision = false;
    engine.set_calc_settings(settings);

    engine
        .set_cell_number_format("Sheet1", "A1", Some("0%".to_string()))
        .unwrap();
    engine.set_cell_value("Sheet1", "A1", 0.1234).unwrap();

    // "0%" displays 12% for 0.1234, so the stored value should be 0.12.
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(0.12));

    engine.set_cell_formula("Sheet1", "B1", "=A1*100").unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(12.0));
}

#[test]
fn full_precision_does_not_round_numeric_literals() {
    let mut engine = Engine::new();
    let mut settings: CalcSettings = engine.calc_settings().clone();
    settings.full_precision = true;
    engine.set_calc_settings(settings);

    engine
        .set_cell_number_format("Sheet1", "A1", Some("0.00".to_string()))
        .unwrap();
    engine.set_cell_value("Sheet1", "A1", 1.239).unwrap();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.239));

    engine.set_cell_formula("Sheet1", "B1", "=A1").unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.239));
}

