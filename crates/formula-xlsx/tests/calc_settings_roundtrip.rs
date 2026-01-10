use formula_model::calc_settings::{CalcSettings, CalculationMode, IterativeCalculationSettings};
use formula_xlsx::calc_settings::read_calc_settings_from_workbook_xml;
use formula_xlsx::XlsxPackage;

#[test]
fn round_trip_preserves_calc_chain_and_calc_settings() {
    let bytes = include_bytes!("fixtures/calc_settings.xlsx");
    let mut pkg = XlsxPackage::from_bytes(bytes).unwrap();

    let settings = pkg.calc_settings().unwrap();
    assert_eq!(settings.calculation_mode, CalculationMode::Manual);
    assert!(settings.calculate_before_save);
    assert!(settings.iterative.enabled);
    assert_eq!(settings.iterative.max_iterations, 10);
    assert!((settings.iterative.max_change - 0.0001).abs() < 1e-12);
    assert!(settings.full_precision);

    let original_calc_chain = pkg.part("xl/calcChain.xml").unwrap().to_vec();

    // Modify and round-trip settings; ensure calcChain.xml is preserved.
    let mut new_settings = CalcSettings::default();
    new_settings.calculation_mode = CalculationMode::Automatic;
    new_settings.iterative = IterativeCalculationSettings {
        enabled: false,
        max_iterations: 25,
        max_change: 0.01,
    };
    new_settings.calculate_before_save = false;
    new_settings.full_precision = true;

    pkg.set_calc_settings(&new_settings).unwrap();

    let out_bytes = pkg.write_to_bytes().unwrap();
    let out_pkg = XlsxPackage::from_bytes(&out_bytes).unwrap();
    assert_eq!(
        out_pkg.part("xl/calcChain.xml").unwrap(),
        original_calc_chain.as_slice()
    );

    let round_tripped = out_pkg.calc_settings().unwrap();
    assert_eq!(round_tripped, new_settings);

    // Also verify workbook.xml contains a calcPr node with our values.
    let workbook_xml = out_pkg.part("xl/workbook.xml").unwrap();
    let parsed_again = read_calc_settings_from_workbook_xml(workbook_xml).unwrap();
    assert_eq!(parsed_again, new_settings);
}
