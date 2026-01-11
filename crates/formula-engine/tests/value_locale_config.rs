use formula_engine::calc_settings::{CalcSettings, CalculationMode};
use formula_engine::locale::ValueLocaleConfig;
use formula_engine::Engine;

#[test]
fn set_value_locale_marks_compiled_formulas_dirty() {
    let mut engine = Engine::new();
    let mut settings = CalcSettings::default();
    settings.calculation_mode = CalculationMode::Manual;
    engine.set_calc_settings(settings);

    engine.set_cell_formula("Sheet1", "A1", "=1+1").unwrap();
    engine.recalculate();
    assert!(!engine.has_dirty_cells(), "engine should be clean after recalc");

    engine.set_value_locale(ValueLocaleConfig::de_de());
    assert!(engine.has_dirty_cells(), "locale change should dirty compiled formulas");
    assert_eq!(engine.value_locale(), ValueLocaleConfig::de_de());
}

