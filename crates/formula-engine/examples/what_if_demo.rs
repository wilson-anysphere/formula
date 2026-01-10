use formula_engine::what_if::goal_seek::{GoalSeek, GoalSeekParams};
use formula_engine::what_if::monte_carlo::{
    Distribution, InputDistribution, MonteCarloEngine, SimulationConfig,
};
use formula_engine::what_if::{CellRef, EngineWhatIfModel};
use formula_engine::{Engine, RecalcMode};

fn main() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_formula("Sheet1", "B1", "=A1*A1").unwrap();

    // Goal Seek: find A1 such that B1 = 9 (i.e. A1 = 3).
    {
        let mut model = EngineWhatIfModel::new(&mut engine, "Sheet1")
            .with_recalc_mode(RecalcMode::SingleThreaded);
        let mut params = GoalSeekParams::new("B1", 9.0, "A1");
        params.tolerance = 1e-9;

        let result = GoalSeek::solve(&mut model, params).expect("goal seek should run");
        println!("Goal Seek status: {:?}", result.status);
        println!("A1 ≈ {}", result.solution);
        println!("B1 ≈ {}", result.final_output);
    }

    // Monte Carlo: simulate A1 ~ Normal(0, 1) and observe B1 (=A1^2).
    {
        let mut model = EngineWhatIfModel::new(&mut engine, "Sheet1")
            .with_recalc_mode(RecalcMode::SingleThreaded);

        let mut config = SimulationConfig::new(1_000);
        config.seed = 42;
        config.input_distributions = vec![InputDistribution {
            cell: CellRef::from("A1"),
            distribution: Distribution::Normal {
                mean: 0.0,
                std_dev: 1.0,
            },
        }];
        config.output_cells = vec![CellRef::from("B1")];

        let sim =
            MonteCarloEngine::run_simulation(&mut model, config).expect("simulation should run");
        let stats = sim.output_stats.get(&CellRef::from("B1")).unwrap();
        println!("Monte Carlo mean (A1^2) ≈ {}", stats.mean);
        println!("Monte Carlo std_dev (A1^2) ≈ {}", stats.std_dev);
    }
}
