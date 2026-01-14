use formula_engine::what_if::monte_carlo::{
    CorrelationMatrix, Distribution, InputDistribution, MonteCarloEngine, SimulationConfig,
};
use formula_engine::what_if::{CellRef, CellValue, InMemoryModel, WhatIfError, WhatIfModel};

#[test]
fn monte_carlo_correlated_uniforms_approximate_requested_correlation() {
    let mut model = InMemoryModel::new();
    model
        .set_cell_value(&CellRef::from("A1"), CellValue::Number(0.0))
        .unwrap();
    model
        .set_cell_value(&CellRef::from("B1"), CellValue::Number(0.0))
        .unwrap();

    let rho = 0.8;
    let mut config = SimulationConfig::new(20_000);
    config.seed = 1337;
    config.input_distributions = vec![
        InputDistribution {
            cell: CellRef::from("A1"),
            distribution: Distribution::Uniform { min: 0.0, max: 1.0 },
        },
        InputDistribution {
            cell: CellRef::from("B1"),
            distribution: Distribution::Uniform { min: 0.0, max: 1.0 },
        },
    ];
    config.output_cells = vec![CellRef::from("A1"), CellRef::from("B1")];
    config.correlations = Some(CorrelationMatrix::new(vec![vec![1.0, rho], vec![rho, 1.0]]));

    let result = MonteCarloEngine::run_simulation(&mut model, config).unwrap();
    let a = result.output_samples.get(&CellRef::from("A1")).unwrap();
    let b = result.output_samples.get(&CellRef::from("B1")).unwrap();

    let corr = sample_correlation(a, b);
    assert!(
        (corr - rho).abs() < 0.1,
        "expected corr ≈ {rho}, got {corr}"
    );
}

#[test]
fn monte_carlo_correlated_normal_and_lognormal_is_supported_and_finite() {
    let mut model = InMemoryModel::new();
    model
        .set_cell_value(&CellRef::from("A1"), CellValue::Number(0.0))
        .unwrap();
    model
        .set_cell_value(&CellRef::from("B1"), CellValue::Number(0.0))
        .unwrap();

    let mut config = SimulationConfig::new(5_000);
    config.seed = 42;
    config.input_distributions = vec![
        InputDistribution {
            cell: CellRef::from("A1"),
            distribution: Distribution::Normal {
                mean: 0.0,
                std_dev: 1.0,
            },
        },
        InputDistribution {
            cell: CellRef::from("B1"),
            distribution: Distribution::Lognormal {
                mean: 0.0,
                std_dev: 0.5,
            },
        },
    ];
    config.output_cells = vec![CellRef::from("A1"), CellRef::from("B1")];
    config.correlations = Some(CorrelationMatrix::new(vec![vec![1.0, 0.6], vec![0.6, 1.0]]));

    let result = MonteCarloEngine::run_simulation(&mut model, config).unwrap();
    for cell in [CellRef::from("A1"), CellRef::from("B1")] {
        let samples = result.output_samples.get(&cell).unwrap();
        assert!(
            samples.iter().all(|v| v.is_finite()),
            "non-finite sample found for {cell}"
        );
    }
}

#[test]
fn monte_carlo_correlated_sampling_rejects_unsupported_distributions() {
    let mut model = InMemoryModel::new();
    model
        .set_cell_value(&CellRef::from("A1"), CellValue::Number(0.0))
        .unwrap();
    model
        .set_cell_value(&CellRef::from("B1"), CellValue::Number(0.0))
        .unwrap();

    let mut config = SimulationConfig::new(10);
    config.seed = 0;
    config.input_distributions = vec![
        InputDistribution {
            cell: CellRef::from("A1"),
            distribution: Distribution::Normal {
                mean: 0.0,
                std_dev: 1.0,
            },
        },
        InputDistribution {
            cell: CellRef::from("B1"),
            distribution: Distribution::Discrete {
                values: vec![0.0, 1.0],
                probabilities: vec![0.5, 0.5],
            },
        },
    ];
    config.output_cells = vec![CellRef::from("A1")];
    config.correlations = Some(CorrelationMatrix::new(vec![vec![1.0, 0.0], vec![0.0, 1.0]]));

    let err = MonteCarloEngine::run_simulation(&mut model, config).unwrap_err();
    match err {
        WhatIfError::InvalidParams(msg) => assert!(
            msg.contains("discrete"),
            "expected error message mentioning discrete distributions, got {msg}"
        ),
        other => panic!("expected InvalidParams error, got {other:?}"),
    }
}

fn sample_correlation(x: &[f64], y: &[f64]) -> f64 {
    assert_eq!(x.len(), y.len());
    let n = x.len();
    let mean_x = x.iter().sum::<f64>() / n as f64;
    let mean_y = y.iter().sum::<f64>() / n as f64;

    let mut cov = 0.0;
    let mut var_x = 0.0;
    let mut var_y = 0.0;
    for (&xi, &yi) in x.iter().zip(y.iter()) {
        let dx = xi - mean_x;
        let dy = yi - mean_y;
        cov += dx * dy;
        var_x += dx * dx;
        var_y += dy * dy;
    }

    if n > 1 {
        cov /= n as f64 - 1.0;
        var_x /= n as f64 - 1.0;
        var_y /= n as f64 - 1.0;
    }

    cov / (var_x.sqrt() * var_y.sqrt())
}
