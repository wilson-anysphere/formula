use formula_model::calc_settings::CalculationMode;

#[test]
fn read_workbook_populates_calc_settings() {
    let fixture_path =
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/calc_settings.xlsx");
    let wb = formula_xlsx::read_workbook(fixture_path).unwrap();

    let settings = &wb.calc_settings;
    assert_eq!(settings.calculation_mode, CalculationMode::Manual);
    assert!(settings.calculate_before_save);
    assert!(settings.iterative.enabled);
    assert_eq!(settings.iterative.max_iterations, 10);
    assert!((settings.iterative.max_change - 0.0001).abs() < 1e-12);
    assert!(settings.full_precision);
    assert!(
        !settings.full_calc_on_load,
        "fixture does not set fullCalcOnLoad, default should be false"
    );
}

#[test]
fn load_from_bytes_populates_calc_settings() {
    let bytes = include_bytes!("fixtures/calc_settings.xlsx");
    let doc = formula_xlsx::load_from_bytes(bytes).unwrap();

    let settings = &doc.workbook.calc_settings;
    assert_eq!(settings.calculation_mode, CalculationMode::Manual);
    assert!(settings.calculate_before_save);
    assert!(settings.iterative.enabled);
    assert_eq!(settings.iterative.max_iterations, 10);
    assert!((settings.iterative.max_change - 0.0001).abs() < 1e-12);
    assert!(settings.full_precision);
    assert!(
        !settings.full_calc_on_load,
        "fixture does not set fullCalcOnLoad, default should be false"
    );
}

