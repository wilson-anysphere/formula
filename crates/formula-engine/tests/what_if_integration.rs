use formula_engine::what_if::goal_seek::{GoalSeek, GoalSeekParams};
use formula_engine::what_if::monte_carlo::{
    Distribution, InputDistribution, MonteCarloEngine, SimulationConfig,
};
use formula_engine::what_if::{CellRef, EngineWhatIfModel};
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
