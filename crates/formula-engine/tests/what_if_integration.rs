use formula_engine::what_if::goal_seek::{GoalSeek, GoalSeekParams};
use formula_engine::what_if::monte_carlo::{
    Distribution, InputDistribution, MonteCarloEngine, SimulationConfig,
};
use formula_engine::what_if::scenario_manager::ScenarioManager;
use formula_engine::what_if::{CellRef, CellValue, EngineWhatIfModel, WhatIfModel};
use formula_engine::{Engine, RecalcMode};

#[test]
fn goal_seek_operates_over_engine_formulas() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_formula("Sheet1", "B1", "=A1*A1").unwrap();

    let mut model =
        EngineWhatIfModel::new(&mut engine, "Sheet1").with_recalc_mode(RecalcMode::SingleThreaded);

    let mut params = GoalSeekParams::new("B1", 25.0, "A1");
    params.tolerance = 1e-9;

    let result = GoalSeek::solve(&mut model, params).unwrap();
    assert!(result.success(), "{result:?}");
    assert!(
        (result.solution - 5.0).abs() < 1e-6,
        "solution = {}",
        result.solution
    );
}

#[test]
fn monte_carlo_drives_engine_recalc_and_collects_outputs() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 0.0).unwrap();
    engine.set_cell_formula("Sheet1", "B1", "=A1+1").unwrap();

    let mut model =
        EngineWhatIfModel::new(&mut engine, "Sheet1").with_recalc_mode(RecalcMode::SingleThreaded);

    let mut config = SimulationConfig::new(5_000);
    config.seed = 123;
    config.input_distributions = vec![InputDistribution {
        cell: CellRef::from("A1"),
        distribution: Distribution::Uniform { min: 0.0, max: 1.0 },
    }];
    config.output_cells = vec![CellRef::from("B1")];

    let result = MonteCarloEngine::run_simulation(&mut model, config).unwrap();
    let stats = result.output_stats.get(&CellRef::from("B1")).unwrap();

    // For A1 ~ U(0,1), B1 = A1 + 1 has mean 1.5.
    assert!((stats.mean - 1.5).abs() < 0.02, "mean = {}", stats.mean);
}

#[test]
fn scenario_manager_applies_scenarios_and_generates_summary_report_over_engine() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 2.0).unwrap();
    engine.set_cell_formula("Sheet1", "B1", "=A1*2").unwrap();

    let mut model =
        EngineWhatIfModel::new(&mut engine, "Sheet1").with_recalc_mode(RecalcMode::SingleThreaded);

    let mut manager = ScenarioManager::new();
    let low = manager
        .create_scenario(
            "Low",
            vec![CellRef::from("A1")],
            vec![CellValue::Number(3.0)],
            "tester",
            None,
        )
        .unwrap();
    let high = manager
        .create_scenario(
            "High",
            vec![CellRef::from("A1")],
            vec![CellValue::Number(8.0)],
            "tester",
            None,
        )
        .unwrap();

    manager.apply_scenario(&mut model, low).unwrap();
    let b1 = model.get_cell_value(&CellRef::from("B1")).unwrap();
    assert!((b1.as_number().unwrap() - 6.0).abs() < 1e-9);

    manager.restore_base(&mut model).unwrap();
    let a1 = model.get_cell_value(&CellRef::from("A1")).unwrap();
    let b1 = model.get_cell_value(&CellRef::from("B1")).unwrap();
    assert!((a1.as_number().unwrap() - 2.0).abs() < 1e-9);
    assert!((b1.as_number().unwrap() - 4.0).abs() < 1e-9);

    let report = manager
        .generate_summary_report(&mut model, vec![CellRef::from("B1")], vec![low, high])
        .unwrap();

    let base = report.results.get("Base").unwrap();
    assert!((base.get(&CellRef::from("B1")).unwrap().as_number().unwrap() - 4.0).abs() < 1e-9);

    let low_row = report.results.get("Low").unwrap();
    assert!(
        (low_row
            .get(&CellRef::from("B1"))
            .unwrap()
            .as_number()
            .unwrap()
            - 6.0)
            .abs()
            < 1e-9
    );

    let high_row = report.results.get("High").unwrap();
    assert!(
        (high_row
            .get(&CellRef::from("B1"))
            .unwrap()
            .as_number()
            .unwrap()
            - 16.0)
            .abs()
            < 1e-9
    );
}
