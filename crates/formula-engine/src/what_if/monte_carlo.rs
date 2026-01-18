use std::collections::{BTreeMap, HashMap};

use crate::what_if::{CellRef, CellValue, WhatIfError, WhatIfModel};
use serde::{Deserialize, Serialize};
use statrs::distribution::{Beta as StatrsBeta, ContinuousCDF};

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SimulationConfig {
    pub iterations: usize,
    pub input_distributions: Vec<InputDistribution>,
    pub output_cells: Vec<CellRef>,
    pub seed: u64,
    /// Optional correlation matrix between input distributions.
    ///
    /// When supplied, correlated sampling uses a Gaussian copula:
    ///
    /// 1. Generate correlated standard normals using the provided correlation matrix.
    /// 2. Convert each normal sample `z` into a uniform sample `u = Φ(z)`.
    /// 3. Transform `u` through the inverse CDF of each input distribution.
    ///
    /// Correlated sampling is currently supported for the following distributions:
    /// - [`Distribution::Normal`]
    /// - [`Distribution::Uniform`]
    /// - [`Distribution::Triangular`]
    /// - [`Distribution::Lognormal`]
    /// - [`Distribution::Exponential`]
    /// - [`Distribution::Beta`]
    ///
    /// Discrete distributions (e.g. [`Distribution::Discrete`], [`Distribution::Poisson`])
    /// are not yet supported with correlations.
    pub correlations: Option<CorrelationMatrix>,
    pub histogram_bins: usize,
}

impl SimulationConfig {
    pub fn new(iterations: usize) -> Self {
        Self {
            iterations,
            input_distributions: Vec::new(),
            output_cells: Vec::new(),
            seed: 0,
            correlations: None,
            histogram_bins: 50,
        }
    }
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct InputDistribution {
    pub cell: CellRef,
    pub distribution: Distribution,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(tag = "type", rename_all = "snake_case")]
pub enum Distribution {
    Normal {
        mean: f64,
        #[serde(rename = "stdDev")]
        std_dev: f64,
    },
    Uniform {
        min: f64,
        max: f64,
    },
    Triangular {
        min: f64,
        mode: f64,
        max: f64,
    },
    Lognormal {
        mean: f64,
        #[serde(rename = "stdDev")]
        std_dev: f64,
    },
    Discrete {
        values: Vec<f64>,
        probabilities: Vec<f64>,
    },
    Beta {
        alpha: f64,
        beta: f64,
        min: Option<f64>,
        max: Option<f64>,
    },
    Exponential {
        rate: f64,
    },
    Poisson {
        lambda: f64,
    },
}

impl Distribution {
    fn validate(&self) -> Result<(), &'static str> {
        match self {
            Distribution::Normal { mean, std_dev } => {
                if !mean.is_finite() {
                    return Err("normal mean must be finite");
                }
                if !std_dev.is_finite() {
                    return Err("normal std_dev must be finite");
                }
                if !(*std_dev >= 0.0) {
                    return Err("normal std_dev must be >= 0");
                }
                Ok(())
            }
            Distribution::Uniform { min, max } => {
                if !min.is_finite() || !max.is_finite() {
                    return Err("uniform min and max must be finite");
                }
                if !(*min <= *max) {
                    return Err("uniform min must be <= max");
                }
                Ok(())
            }
            Distribution::Triangular { min, mode, max } => {
                if !min.is_finite() || !mode.is_finite() || !max.is_finite() {
                    return Err("triangular min, mode, and max must be finite");
                }
                if !(*min <= *mode && *mode <= *max) {
                    return Err("triangular requires min <= mode <= max");
                }
                Ok(())
            }
            Distribution::Lognormal { mean, std_dev } => {
                if !mean.is_finite() {
                    return Err("lognormal mean must be finite");
                }
                if !std_dev.is_finite() {
                    return Err("lognormal std_dev must be finite");
                }
                if !(*std_dev >= 0.0) {
                    return Err("lognormal std_dev must be >= 0");
                }
                Ok(())
            }
            Distribution::Discrete {
                values,
                probabilities,
            } => {
                if values.is_empty() {
                    return Err("discrete distribution requires at least one value");
                }
                if values.len() != probabilities.len() {
                    return Err("discrete values and probabilities must have equal length");
                }
                if values.iter().any(|v| !v.is_finite()) {
                    return Err("discrete values must be finite");
                }
                if probabilities.iter().any(|p| !p.is_finite()) {
                    return Err("discrete probabilities must be finite");
                }
                if probabilities.iter().any(|p| *p < 0.0) {
                    return Err("discrete probabilities must be >= 0");
                }
                let sum: f64 = probabilities.iter().sum();
                if !(sum > 0.0) {
                    return Err("discrete probabilities must sum to > 0");
                }
                Ok(())
            }
            Distribution::Beta {
                alpha,
                beta,
                min,
                max,
            } => {
                if !alpha.is_finite() || !beta.is_finite() {
                    return Err("beta alpha and beta must be finite");
                }
                if !(*alpha > 0.0 && *beta > 0.0) {
                    return Err("beta alpha and beta must be > 0");
                }
                if let (Some(min), Some(max)) = (min, max) {
                    if !min.is_finite() || !max.is_finite() {
                        return Err("beta min and max must be finite");
                    }
                    if !(*min <= *max) {
                        return Err("beta min must be <= max");
                    }
                }
                Ok(())
            }
            Distribution::Exponential { rate } => {
                if !rate.is_finite() {
                    return Err("exponential rate must be finite");
                }
                if !(*rate > 0.0) {
                    return Err("exponential rate must be > 0");
                }
                Ok(())
            }
            Distribution::Poisson { lambda } => {
                if !lambda.is_finite() {
                    return Err("poisson lambda must be finite");
                }
                if !(*lambda >= 0.0) {
                    return Err("poisson lambda must be >= 0");
                }
                Ok(())
            }
        }
    }

    fn correlated_sampling_validation_error(&self) -> Option<&'static str> {
        match self {
            Distribution::Discrete { .. } => {
                Some("correlated sampling is not supported for discrete distributions")
            }
            Distribution::Poisson { .. } => {
                Some("correlated sampling is not supported for poisson distributions")
            }
            _ => None,
        }
    }

    fn from_standard_normal(&self, z: f64) -> f64 {
        match self {
            Distribution::Normal { mean, std_dev } => {
                if *std_dev == 0.0 {
                    *mean
                } else {
                    *mean + *std_dev * z
                }
            }
            Distribution::Uniform { min, max } => {
                if min == max {
                    *min
                } else {
                    let u = standard_normal_cdf(z);
                    min + (max - min) * u
                }
            }
            Distribution::Triangular { min, mode, max } => {
                if min == max {
                    *min
                } else {
                    let u = standard_normal_cdf(z);
                    let f = (mode - min) / (max - min);
                    if u < f {
                        min + (u * (max - min) * (mode - min)).sqrt()
                    } else {
                        max - ((1.0 - u) * (max - min) * (max - mode)).sqrt()
                    }
                }
            }
            Distribution::Lognormal { mean, std_dev } => {
                if *std_dev == 0.0 {
                    mean.exp()
                } else {
                    (mean + std_dev * z).exp()
                }
            }
            Distribution::Discrete { .. } => {
                debug_assert!(
                    false,
                    "discrete distributions are rejected for correlated sampling"
                );
                f64::NAN
            }
            Distribution::Beta {
                alpha,
                beta,
                min,
                max,
            } => {
                let u = standard_normal_cdf(z).clamp(f64::MIN_POSITIVE, 1.0 - f64::EPSILON);
                let raw = if *alpha == 1.0 && *beta == 1.0 {
                    // Uniform(0, 1) shortcut.
                    u
                } else {
                    // `StatrsBeta::new` should not fail after `validate`, but we avoid panics
                    // here since this pathway is only used when correlations are requested.
                    match StatrsBeta::new(*alpha, *beta) {
                        Ok(dist) => dist.inverse_cdf(u),
                        Err(_) => {
                            debug_assert!(false, "validated beta parameters failed to construct");
                            u
                        }
                    }
                };
                scale_unit_interval(raw, *min, *max)
            }
            Distribution::Exponential { rate } => {
                let u = standard_normal_cdf(z);
                let tail = (1.0 - u).max(f64::MIN_POSITIVE);
                -tail.ln() / rate
            }
            Distribution::Poisson { .. } => {
                debug_assert!(
                    false,
                    "poisson distributions are rejected for correlated sampling"
                );
                f64::NAN
            }
        }
    }

    fn sample(&self, rng: &mut SeededRng) -> f64 {
        match self {
            Distribution::Normal { mean, std_dev } => {
                if *std_dev == 0.0 {
                    return *mean;
                }
                *mean + *std_dev * standard_normal(rng)
            }
            Distribution::Uniform { min, max } => {
                if min == max {
                    return *min;
                }
                min + (max - min) * rng.next_f64()
            }
            Distribution::Triangular { min, mode, max } => {
                if min == max {
                    return *min;
                }
                let u = rng.next_f64();
                let f = (mode - min) / (max - min);
                if u < f {
                    min + (u * (max - min) * (mode - min)).sqrt()
                } else {
                    max - ((1.0 - u) * (max - min) * (max - mode)).sqrt()
                }
            }
            Distribution::Lognormal { mean, std_dev } => {
                if *std_dev == 0.0 {
                    return mean.exp();
                }
                (mean + std_dev * standard_normal(rng)).exp()
            }
            Distribution::Discrete {
                values,
                probabilities,
            } => {
                let total: f64 = probabilities.iter().sum();
                let mut threshold = rng.next_f64() * total;
                for (value, p) in values.iter().zip(probabilities.iter()) {
                    threshold -= p;
                    if threshold <= 0.0 {
                        return *value;
                    }
                }
                // Due to floating-point rounding, fall back to the last entry.
                let Some(last) = values.last() else {
                    debug_assert!(
                        false,
                        "Discrete distribution should have been validated as non-empty"
                    );
                    return f64::NAN;
                };
                *last
            }
            Distribution::Beta {
                alpha,
                beta,
                min,
                max,
            } => {
                if *alpha == 1.0 && *beta == 1.0 {
                    // Uniform(0, 1) shortcut.
                    let raw = rng.next_f64();
                    return scale_unit_interval(raw, *min, *max);
                }
                let x = sample_gamma(rng, *alpha);
                let y = sample_gamma(rng, *beta);
                let raw = if x == 0.0 && y == 0.0 {
                    0.5
                } else {
                    x / (x + y)
                };
                scale_unit_interval(raw, *min, *max)
            }
            Distribution::Exponential { rate } => {
                let u = rng.next_f64().max(f64::MIN_POSITIVE);
                -u.ln() / rate
            }
            Distribution::Poisson { lambda } => {
                if *lambda == 0.0 {
                    return 0.0;
                }
                if *lambda < 30.0 {
                    let l = (-lambda).exp();
                    let mut k: u64 = 0;
                    let mut p = 1.0;
                    loop {
                        k += 1;
                        p *= rng.next_f64();
                        if p <= l {
                            return (k - 1) as f64;
                        }
                    }
                } else {
                    // Normal approximation for large lambda.
                    let z = standard_normal(rng);
                    let sample = lambda + lambda.sqrt() * z;
                    sample.max(0.0).round()
                }
            }
        }
    }
}

fn scale_unit_interval(raw: f64, min: Option<f64>, max: Option<f64>) -> f64 {
    let min = min.unwrap_or(0.0);
    let max = max.unwrap_or(1.0);
    if min == 0.0 && max == 1.0 {
        raw
    } else {
        min + raw * (max - min)
    }
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CorrelationMatrix {
    pub matrix: Vec<Vec<f64>>,
}

impl CorrelationMatrix {
    pub fn new(matrix: Vec<Vec<f64>>) -> Self {
        Self { matrix }
    }

    fn validate(&self, expected_size: usize) -> Result<(), &'static str> {
        let n = self.matrix.len();
        if n == 0 {
            return Err("correlation matrix must not be empty");
        }
        if n != expected_size {
            return Err("correlation matrix size must match input_distributions length");
        }

        for (i, row) in self.matrix.iter().enumerate() {
            if row.len() != n {
                return Err("correlation matrix must be square");
            }

            for (j, value) in row.iter().enumerate() {
                if !value.is_finite() {
                    return Err("correlation matrix contains non-finite value");
                }

                if i == j {
                    if (value - 1.0).abs() > 1e-9 {
                        return Err("correlation matrix diagonal entries must be 1");
                    }
                } else if *value < -1.0 || *value > 1.0 {
                    return Err("correlation matrix entries must be within [-1, 1]");
                } else if (value - self.matrix[j][i]).abs() > 1e-9 {
                    return Err("correlation matrix must be symmetric");
                }
            }
        }

        Ok(())
    }
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct HistogramBin {
    pub start: f64,
    pub end: f64,
    pub count: usize,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Histogram {
    pub bins: Vec<HistogramBin>,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct OutputStatistics {
    pub mean: f64,
    pub median: f64,
    pub std_dev: f64,
    pub min: f64,
    pub max: f64,
    /// Percentile (0-100) -> value.
    pub percentiles: BTreeMap<u8, f64>,
    pub histogram: Histogram,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SimulationResult {
    pub iterations: usize,
    pub output_stats: HashMap<CellRef, OutputStatistics>,
    /// Raw output samples for charting/inspection.
    pub output_samples: HashMap<CellRef, Vec<f64>>,
}

#[derive(Clone, Copy, Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SimulationProgress {
    pub completed_iterations: usize,
    pub total_iterations: usize,
}

pub struct MonteCarloEngine;

impl MonteCarloEngine {
    pub fn run_simulation<M: WhatIfModel>(
        model: &mut M,
        config: SimulationConfig,
    ) -> Result<SimulationResult, WhatIfError<M::Error>> {
        Self::run_simulation_with_progress(model, config, |_| {})
    }

    pub fn run_simulation_with_progress<M: WhatIfModel, F: FnMut(SimulationProgress)>(
        model: &mut M,
        config: SimulationConfig,
        mut progress: F,
    ) -> Result<SimulationResult, WhatIfError<M::Error>> {
        if config.iterations == 0 {
            return Err(WhatIfError::InvalidParams("iterations must be > 0"));
        }
        if config.histogram_bins == 0 {
            return Err(WhatIfError::InvalidParams("histogram_bins must be > 0"));
        }
        if config.output_cells.is_empty() {
            return Err(WhatIfError::InvalidParams("output_cells must not be empty"));
        }
        for input in &config.input_distributions {
            input
                .distribution
                .validate()
                .map_err(WhatIfError::InvalidParams)?;
        }

        let mut rng = SeededRng::new(config.seed);

        let correlated = if let Some(corr) = &config.correlations {
            corr.validate(config.input_distributions.len())
                .map_err(WhatIfError::InvalidParams)?;

            for input in &config.input_distributions {
                if let Some(msg) = input.distribution.correlated_sampling_validation_error() {
                    return Err(WhatIfError::InvalidParams(msg));
                }
            }

            let l = cholesky_decomposition(&corr.matrix).map_err(WhatIfError::InvalidParams)?;
            Some(l)
        } else {
            None
        };

        let mut output_samples: HashMap<CellRef, Vec<f64>> = HashMap::new();
        if output_samples
            .try_reserve(config.output_cells.len())
            .is_err()
        {
            debug_assert!(false, "monte carlo allocation failed (output_cells)");
            return Err(WhatIfError::NumericalFailure(
                "allocation failed (output cells)",
            ));
        }
        for cell in &config.output_cells {
            let mut samples: Vec<f64> = Vec::new();
            if samples.try_reserve_exact(config.iterations).is_err() {
                debug_assert!(
                    false,
                    "monte carlo allocation failed (iterations={})",
                    config.iterations
                );
                return Err(WhatIfError::NumericalFailure(
                    "allocation failed (output samples)",
                ));
            }
            output_samples.insert(cell.clone(), samples);
        }

        for i in 0..config.iterations {
            if let Some(l) = &correlated {
                let z = generate_correlated_normals(&mut rng, l).map_err(|_| {
                    WhatIfError::NumericalFailure("allocation failed (correlated normals)")
                })?;
                for (input, zi) in config.input_distributions.iter().zip(z.into_iter()) {
                    let value = input.distribution.from_standard_normal(zi);
                    model.set_cell_value(&input.cell, CellValue::Number(value))?;
                }
            } else {
                for input in &config.input_distributions {
                    let value = input.distribution.sample(&mut rng);
                    model.set_cell_value(&input.cell, CellValue::Number(value))?;
                }
            }

            model.recalculate()?;

            for cell in &config.output_cells {
                let value = model.get_cell_value(cell)?;
                let number = value
                    .as_number()
                    .ok_or_else(|| WhatIfError::NonNumericCell {
                        cell: cell.clone(),
                        value,
                    })?;

                let Some(samples) = output_samples.get_mut(cell) else {
                    debug_assert!(
                        false,
                        "Missing output samples buffer for configured output cell {cell}"
                    );
                    return Err(WhatIfError::NumericalFailure(
                        "missing output samples buffer",
                    ));
                };
                samples.push(number);
            }

            // Report progress roughly every 1% (and always on the last iteration).
            let step = (config.iterations / 100).max(1);
            if i % step == 0 || i + 1 == config.iterations {
                progress(SimulationProgress {
                    completed_iterations: i + 1,
                    total_iterations: config.iterations,
                });
            }
        }

        let mut output_stats = HashMap::new();
        for cell in &config.output_cells {
            let Some(samples) = output_samples.get(cell) else {
                debug_assert!(
                    false,
                    "Missing output samples buffer for configured output cell {cell}"
                );
                return Err(WhatIfError::NumericalFailure(
                    "missing output samples buffer",
                ));
            };
            output_stats.insert(
                cell.clone(),
                analyze_samples(samples, config.histogram_bins),
            );
        }

        Ok(SimulationResult {
            iterations: config.iterations,
            output_stats,
            output_samples,
        })
    }
}

fn cholesky_decomposition(matrix: &[Vec<f64>]) -> Result<Vec<Vec<f64>>, &'static str> {
    let n = matrix.len();
    let mut l: Vec<Vec<f64>> = Vec::new();
    if l.try_reserve_exact(n).is_err() {
        debug_assert!(false, "cholesky allocation failed (n={n})");
        return Err("allocation failed");
    }
    for _ in 0..n {
        let mut row: Vec<f64> = Vec::new();
        if row.try_reserve_exact(n).is_err() {
            debug_assert!(false, "cholesky allocation failed (n={n})");
            return Err("allocation failed");
        }
        row.resize(n, 0.0);
        l.push(row);
    }

    for i in 0..n {
        for j in 0..=i {
            let mut sum = 0.0;
            for k in 0..j {
                sum += l[i][k] * l[j][k];
            }

            if i == j {
                let diag = matrix[i][i] - sum;
                if diag <= 0.0 {
                    return Err("correlation matrix is not positive definite");
                }
                l[i][j] = diag.sqrt();
            } else {
                if l[j][j] == 0.0 {
                    return Err("correlation matrix is not positive definite");
                }
                l[i][j] = (matrix[i][j] - sum) / l[j][j];
            }
        }
    }

    Ok(l)
}

fn generate_correlated_normals(
    rng: &mut SeededRng,
    l: &[Vec<f64>],
) -> Result<Vec<f64>, &'static str> {
    let n = l.len();
    let mut z: Vec<f64> = Vec::new();
    if z.try_reserve_exact(n).is_err() {
        debug_assert!(false, "correlated normals allocation failed (n={n})");
        return Err("allocation failed");
    }
    z.resize(n, 0.0);
    for zi in &mut z {
        *zi = standard_normal(rng);
    }

    let mut out: Vec<f64> = Vec::new();
    if out.try_reserve_exact(n).is_err() {
        debug_assert!(false, "correlated normals allocation failed (n={n})");
        return Err("allocation failed");
    }
    out.resize(n, 0.0);
    for i in 0..n {
        let mut sum = 0.0;
        for j in 0..=i {
            sum += l[i][j] * z[j];
        }
        out[i] = sum;
    }
    Ok(out)
}

fn analyze_samples(samples: &[f64], histogram_bins: usize) -> OutputStatistics {
    let n = samples.len().max(1) as f64;

    let min = samples.iter().copied().fold(f64::INFINITY, |a, b| a.min(b));
    let max = samples
        .iter()
        .copied()
        .fold(f64::NEG_INFINITY, |a, b| a.max(b));

    let mean = samples.iter().sum::<f64>() / n;

    let mut variance_sum = 0.0;
    for v in samples {
        let d = *v - mean;
        variance_sum += d * d;
    }
    let std_dev = if samples.len() > 1 {
        (variance_sum / (samples.len() as f64 - 1.0)).sqrt()
    } else {
        0.0
    };

    let mut sorted = samples.to_vec();
    sorted.sort_by(|a, b| a.total_cmp(b));

    let median = percentile_sorted(&sorted, 50.0);

    let mut percentiles = BTreeMap::new();
    for p in [5_u8, 10, 25, 75, 90, 95] {
        percentiles.insert(p, percentile_sorted(&sorted, p as f64));
    }

    let histogram = build_histogram(samples, min, max, histogram_bins);

    OutputStatistics {
        mean,
        median,
        std_dev,
        min,
        max,
        percentiles,
        histogram,
    }
}

fn percentile_sorted(sorted: &[f64], percentile: f64) -> f64 {
    if sorted.is_empty() {
        return f64::NAN;
    }
    if percentile <= 0.0 {
        return sorted[0];
    }
    if percentile >= 100.0 {
        return sorted[sorted.len() - 1];
    }

    let rank = (percentile / 100.0) * (sorted.len() as f64 - 1.0);
    let lo = rank.floor() as usize;
    let hi = rank.ceil() as usize;
    if lo == hi {
        sorted[lo]
    } else {
        let w = rank - lo as f64;
        sorted[lo] + (sorted[hi] - sorted[lo]) * w
    }
}

fn build_histogram(samples: &[f64], min: f64, max: f64, bins: usize) -> Histogram {
    if bins == 0 {
        return Histogram { bins: Vec::new() };
    }

    if !min.is_finite() || !max.is_finite() || samples.is_empty() {
        return Histogram { bins: Vec::new() };
    }

    if min == max {
        let mut out: Vec<HistogramBin> = Vec::new();
        if out.try_reserve_exact(1).is_err() {
            debug_assert!(false, "histogram allocation failed (bins=1)");
            return Histogram { bins: Vec::new() };
        }
        out.push(HistogramBin {
            start: min,
            end: max,
            count: samples.len(),
        });
        return Histogram { bins: out };
    }

    let width = (max - min) / bins as f64;
    let mut counts: Vec<usize> = Vec::new();
    if counts.try_reserve_exact(bins).is_err() {
        debug_assert!(false, "histogram allocation failed (bins={bins})");
        return Histogram { bins: Vec::new() };
    }
    counts.resize(bins, 0);
    for v in samples {
        let mut idx = ((*v - min) / width) as isize;
        if idx < 0 {
            idx = 0;
        }
        if idx as usize >= bins {
            idx = bins as isize - 1;
        }
        counts[idx as usize] += 1;
    }

    let mut bin_defs: Vec<HistogramBin> = Vec::new();
    if bin_defs.try_reserve_exact(bins).is_err() {
        debug_assert!(false, "histogram allocation failed (bins={bins})");
        return Histogram { bins: Vec::new() };
    }
    for (i, count) in counts.into_iter().enumerate() {
        let start = min + i as f64 * width;
        let end = if i + 1 == bins {
            max
        } else {
            min + (i + 1) as f64 * width
        };
        bin_defs.push(HistogramBin { start, end, count });
    }

    Histogram { bins: bin_defs }
}

/// A small deterministic RNG (SplitMix64). This keeps the crate dependency-free.
#[derive(Clone, Debug)]
struct SeededRng {
    state: u64,
}

impl SeededRng {
    fn new(seed: u64) -> Self {
        Self { state: seed }
    }

    fn next_u64(&mut self) -> u64 {
        self.state = self.state.wrapping_add(0x9E3779B97F4A7C15);
        let mut z = self.state;
        z = (z ^ (z >> 30)).wrapping_mul(0xBF58476D1CE4E5B9);
        z = (z ^ (z >> 27)).wrapping_mul(0x94D049BB133111EB);
        z ^ (z >> 31)
    }

    fn next_f64(&mut self) -> f64 {
        // Use the top 53 bits to create a float in [0, 1).
        let bits = self.next_u64() >> 11;
        (bits as f64) / ((1_u64 << 53) as f64)
    }
}

fn standard_normal(rng: &mut SeededRng) -> f64 {
    // Box–Muller transform.
    let u1 = rng.next_f64().max(f64::MIN_POSITIVE);
    let u2 = rng.next_f64();
    let r = (-2.0 * u1.ln()).sqrt();
    let theta = 2.0 * std::f64::consts::PI * u2;
    r * theta.cos()
}

fn standard_normal_cdf(z: f64) -> f64 {
    // Φ(z) = 0.5 * (1 + erf(z / sqrt(2))).
    0.5 * (1.0 + libm::erf(z / std::f64::consts::SQRT_2))
}

fn sample_gamma(rng: &mut SeededRng, shape: f64) -> f64 {
    // Marsaglia and Tsang method (with k<1 boost).
    if shape <= 0.0 {
        return 0.0;
    }

    if shape < 1.0 {
        let u = rng.next_f64().max(f64::MIN_POSITIVE);
        return sample_gamma(rng, shape + 1.0) * u.powf(1.0 / shape);
    }

    let d = shape - 1.0 / 3.0;
    let c = (1.0 / (3.0 * d)).sqrt();

    loop {
        let x = standard_normal(rng);
        let v = 1.0 + c * x;
        if v <= 0.0 {
            continue;
        }
        let v = v * v * v;
        let u = rng.next_f64();

        // Squeeze test.
        if u < 1.0 - 0.0331 * x * x * x * x {
            return d * v;
        }

        if u.ln() < 0.5 * x * x + d * (1.0 - v + v.ln()) {
            return d * v;
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::what_if::{InMemoryModel, WhatIfModel};

    #[test]
    fn monte_carlo_normal_distribution_mean_and_std_dev_are_reasonable() {
        let mut model = InMemoryModel::new();
        model
            .set_cell_value(&CellRef::from("A1"), CellValue::Number(0.0))
            .unwrap();

        let mut config = SimulationConfig::new(10_000);
        config.seed = 1234;
        config.input_distributions = vec![InputDistribution {
            cell: CellRef::from("A1"),
            distribution: Distribution::Normal {
                mean: 100.0,
                std_dev: 10.0,
            },
        }];
        config.output_cells = vec![CellRef::from("A1")];

        let result = MonteCarloEngine::run_simulation(&mut model, config).unwrap();
        let stats = result.output_stats.get(&CellRef::from("A1")).unwrap();

        // With a deterministic seed and 10k samples, these should be very stable
        // while still allowing reasonable tolerance for implementation changes.
        assert!((stats.mean - 100.0).abs() < 0.5, "mean = {}", stats.mean);
        assert!(
            (stats.std_dev - 10.0).abs() < 0.5,
            "std_dev = {}",
            stats.std_dev
        );
        assert!(stats.min.is_finite());
        assert!(stats.max.is_finite());
        assert!(stats.histogram.bins.len() > 1);
    }

    #[test]
    fn monte_carlo_correlated_normals_approximate_requested_correlation() {
        let mut model = InMemoryModel::new();
        model
            .set_cell_value(&CellRef::from("A1"), CellValue::Number(0.0))
            .unwrap();
        model
            .set_cell_value(&CellRef::from("B1"), CellValue::Number(0.0))
            .unwrap();

        let rho = 0.8;
        let mut config = SimulationConfig::new(10_000);
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
                distribution: Distribution::Normal {
                    mean: 0.0,
                    std_dev: 1.0,
                },
            },
        ];
        config.output_cells = vec![CellRef::from("A1"), CellRef::from("B1")];
        config.correlations = Some(CorrelationMatrix::new(vec![vec![1.0, rho], vec![rho, 1.0]]));

        let result = MonteCarloEngine::run_simulation(&mut model, config).unwrap();
        let a = result.output_samples.get(&CellRef::from("A1")).unwrap();
        let b = result.output_samples.get(&CellRef::from("B1")).unwrap();

        let corr = sample_correlation(a, b);
        assert!(
            (corr - rho).abs() < 0.05,
            "expected corr ≈ {rho}, got {corr}"
        );
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
}
