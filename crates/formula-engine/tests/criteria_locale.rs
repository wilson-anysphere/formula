use formula_engine::{Engine, LocaleConfig, Value};

#[test]
fn countif_parses_locale_decimal_in_criteria_strings() {
    let mut engine = Engine::new();
    engine.set_locale_config(LocaleConfig::de_de());

    engine.set_cell_value("Sheet1", "A1", 1.4).unwrap();
    engine.set_cell_value("Sheet1", "A2", 1.6).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", r#"=COUNTIF(A1:A2,">1,5")"#)
        .unwrap();

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));
}

#[test]
fn countif_accepts_canonical_decimal_in_non_en_us_locale() {
    let mut engine = Engine::new();
    engine.set_locale_config(LocaleConfig::de_de());

    engine.set_cell_value("Sheet1", "A1", 1.4).unwrap();
    engine.set_cell_value("Sheet1", "A2", 1.6).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", r#"=COUNTIF(A1:A2,">1.5")"#)
        .unwrap();

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));
}

#[test]
fn sumif_parses_locale_decimal_in_criteria_strings() {
    let mut engine = Engine::new();
    engine.set_locale_config(LocaleConfig::de_de());

    engine.set_cell_value("Sheet1", "A1", 1.4).unwrap();
    engine.set_cell_value("Sheet1", "A2", 1.6).unwrap();
    engine.set_cell_value("Sheet1", "B1", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 20.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", r#"=SUMIF(A1:A2,">1,5",B1:B2)"#)
        .unwrap();

    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(20.0));
}
