use formula_engine::what_if::goal_seek::{GoalSeek, GoalSeekParams};
use formula_engine::what_if::monte_carlo::{
    Distribution, InputDistribution, MonteCarloEngine, SimulationConfig,
};
use formula_engine::what_if::{CellRef, CellValue, WhatIfModel};

/// Minimal model: `B1 = A1^2`.
struct SquareModel {
    input: f64,
    output: f64,
}

impl SquareModel {
    fn new(input: f64) -> Self {
        Self {
            input,
            output: input * input,
        }
    }
}

impl WhatIfModel for SquareModel {
    type Error = &'static str;

    fn get_cell_value(&self, cell: &CellRef) -> Result<CellValue, Self::Error> {
        match cell.as_str() {
            "A1" => Ok(CellValue::Number(self.input)),
            "B1" => Ok(CellValue::Number(self.output)),
            _ => Ok(CellValue::Blank),
        }
    }

    fn set_cell_value(&mut self, cell: &CellRef, value: CellValue) -> Result<(), Self::Error> {
        if cell.as_str() == "A1" {
            self.input = value.as_number().ok_or("A1 must be numeric")?;
        }
        Ok(())
    }

    fn recalculate(&mut self) -> Result<(), Self::Error> {
        self.output = self.input * self.input;
        Ok(())
    }
}

fn main() {
    // Goal Seek: find A1 such that B1 = 9 (i.e. A1 = 3).
    let mut model = SquareModel::new(1.0);
    let mut params = GoalSeekParams::new("B1", 9.0, "A1");
    params.tolerance = 1e-9;

    let result = GoalSeek::solve(&mut model, params).expect("goal seek should run");
    println!("Goal Seek status: {:?}", result.status);
    println!("A1 ≈ {}", result.solution);
    println!("B1 ≈ {}", result.final_output);

    // Monte Carlo: simulate A1 ~ Normal(100, 10) and observe A1 itself.
    let mut mem_model = formula_engine::what_if::InMemoryModel::new();
    let mut config = SimulationConfig::new(1_000);
    config.seed = 42;
    config.input_distributions = vec![InputDistribution {
        cell: CellRef::from("A1"),
        distribution: Distribution::Normal {
            mean: 100.0,
            std_dev: 10.0,
        },
    }];
    config.output_cells = vec![CellRef::from("A1")];

    let sim =
        MonteCarloEngine::run_simulation(&mut mem_model, config).expect("simulation should run");
    let stats = sim.output_stats.get(&CellRef::from("A1")).unwrap();
    println!("Monte Carlo mean ≈ {}", stats.mean);
    println!("Monte Carlo std_dev ≈ {}", stats.std_dev);
}
