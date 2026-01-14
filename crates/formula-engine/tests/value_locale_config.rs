use formula_engine::calc_settings::{CalcSettings, CalculationMode};
use formula_engine::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};
use formula_engine::locale::ValueLocaleConfig;
use formula_engine::value::Value;
use formula_engine::Engine;

fn engine_manual() -> Engine {
    let mut engine = Engine::new();
    let mut settings = CalcSettings::default();
    settings.calculation_mode = CalculationMode::Manual;
    engine.set_calc_settings(settings);
    engine
}

#[test]
fn set_value_locale_marks_compiled_formulas_dirty() {
    let mut engine = engine_manual();

    engine.set_cell_formula("Sheet1", "A1", "=1+1").unwrap();
    engine.recalculate();
    assert!(!engine.has_dirty_cells(), "engine should be clean after recalc");

    engine.set_value_locale(ValueLocaleConfig::de_de());
    assert!(engine.has_dirty_cells(), "locale change should dirty compiled formulas");
    assert_eq!(engine.value_locale(), ValueLocaleConfig::de_de());
}

#[test]
fn bytecode_respects_value_locale_for_numeric_coercion() {
    let mut engine = engine_manual();
    engine.set_value_locale(ValueLocaleConfig::de_de());

    engine.set_cell_value("Sheet1", "A1", "1,5").unwrap();
    engine.set_cell_formula("Sheet1", "B1", "=A1+1").unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(2.5));
}

#[test]
fn bytecode_countif_respects_value_locale_for_numeric_criteria_strings() {
    let mut engine = engine_manual();
    engine.set_value_locale(ValueLocaleConfig::de_de());

    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", r#"=COUNTIF(A1:A3,">1,5")"#)
        .unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(2.0));
}

#[test]
fn bytecode_countif_respects_workbook_locale_for_non_numeric_criteria_strings() {
    let mut engine = engine_manual();
    engine.set_value_locale(ValueLocaleConfig::de_de());

    engine.set_cell_value("Sheet1", "A1", "1,2,3").unwrap();
    engine.set_cell_value("Sheet1", "A2", "x").unwrap();
    engine.set_cell_value("Sheet1", "A3", "1,2,3").unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", r#"=COUNTIF(A1:A3,"1,2,3")"#)
        .unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(2.0));
}

#[test]
fn bytecode_countif_respects_workbook_locale_for_dot_separated_date_criteria_strings() {
    let mut engine = engine_manual();
    engine.set_value_locale(ValueLocaleConfig::de_de());

    let system = ExcelDateSystem::EXCEL_1900;
    let feb_1_2020 = ymd_to_serial(ExcelDate::new(2020, 2, 1), system).unwrap() as f64;
    let jan_2_2020 = ymd_to_serial(ExcelDate::new(2020, 1, 2), system).unwrap() as f64;
    engine.set_cell_value("Sheet1", "A1", feb_1_2020).unwrap();
    engine.set_cell_value("Sheet1", "A2", jan_2_2020).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", r#"=COUNTIF(A1:A2,"1.2.2020")"#)
        .unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));
}

#[test]
fn set_value_locale_id_accepts_common_en_dmy_locales_and_uses_dmy_date_order() {
    let system = ExcelDateSystem::EXCEL_1900;
    let expected_serial =
        ymd_to_serial(ExcelDate::new(2020, 2, 1), system).unwrap() as f64;

    // Many English-speaking regions use DMY date order (Excel-compatible parsing).
    // Note: formula parsing locale still resolves to `en-US`; this only affects value parsing.
    for locale_id in ["en-GB", "en-AU", "en-NZ", "en-IE", "en-ZA"] {
        let mut engine = engine_manual();
        assert!(
            engine.set_value_locale_id(locale_id),
            "expected {locale_id} to be accepted as a value locale id"
        );

        engine
            .set_cell_formula("Sheet1", "A1", r#"=DATEVALUE("1/2/2020")"#)
            .unwrap();
        engine.recalculate();

        assert_eq!(
            engine.get_cell_value("Sheet1", "A1"),
            Value::Number(expected_serial),
            "unexpected DATEVALUE result for {locale_id}"
        );
    }
}
