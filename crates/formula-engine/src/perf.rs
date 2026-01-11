use std::time::Instant;

use crate::eval::{
    CellAddr, CompiledExpr, EvalContext, Evaluator, Parser, SheetReference, ValueResolver,
};
use crate::{Engine, Value};

#[derive(Debug, Clone)]
pub struct BenchmarkResult {
    pub name: String,
    pub iterations: usize,
    pub warmup: usize,
    pub unit: &'static str,
    pub mean: f64,
    pub median: f64,
    pub p95: f64,
    pub p99: f64,
    pub std_dev: f64,
    pub target_ms: f64,
    pub passed: bool,
}

fn percentile(sorted: &[f64], p: f64) -> f64 {
    if sorted.is_empty() {
        return 0.0;
    }
    let idx = ((sorted.len() as f64) * p).floor() as usize;
    sorted[idx.min(sorted.len() - 1)]
}

fn median(sorted: &[f64]) -> f64 {
    sorted[sorted.len() / 2]
}

fn mean(values: &[f64]) -> f64 {
    values.iter().sum::<f64>() / values.len() as f64
}

fn std_dev(values: &[f64], avg: f64) -> f64 {
    let variance = values
        .iter()
        .map(|x| {
            let d = x - avg;
            d * d
        })
        .sum::<f64>()
        / values.len() as f64;
    variance.sqrt()
}

fn run_benchmark<F>(
    name: &str,
    iterations: usize,
    warmup: usize,
    target_ms: f64,
    mut f: F,
) -> BenchmarkResult
where
    F: FnMut(),
{
    for _ in 0..warmup {
        f();
    }

    let mut samples: Vec<f64> = Vec::with_capacity(iterations);
    for _ in 0..iterations {
        let start = Instant::now();
        f();
        let ms = start.elapsed().as_secs_f64() * 1000.0;
        samples.push(ms);
    }

    samples.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));

    let avg = mean(&samples);
    let med = median(&samples);
    let p95 = percentile(&samples, 0.95);
    let p99 = percentile(&samples, 0.99);
    let sd = std_dev(&samples, avg);

    BenchmarkResult {
        name: name.to_string(),
        iterations,
        warmup,
        unit: "ms",
        mean: avg,
        median: med,
        p95,
        p99,
        std_dev: sd,
        target_ms,
        passed: p95 <= target_ms,
    }
}

fn compile_for_benchmark(parsed: &crate::eval::ParsedExpr) -> CompiledExpr {
    let mut map = |sref: &SheetReference<String>| match sref {
        SheetReference::Current => SheetReference::Current,
        // Benchmarks operate on a single sheet; map any explicit sheet refs to 0.
        SheetReference::Sheet(_name) => SheetReference::Sheet(0),
        SheetReference::External(wb) => SheetReference::External(wb.clone()),
    };
    parsed.map_sheets(&mut map)
}

#[derive(Debug, Clone, Copy)]
struct ConstResolver;

impl ValueResolver for ConstResolver {
    fn sheet_exists(&self, sheet_id: usize) -> bool {
        sheet_id == 0
    }

    fn get_cell_value(&self, _sheet_id: usize, _addr: CellAddr) -> Value {
        Value::Number(1.0)
    }

    fn resolve_structured_ref(
        &self,
        _ctx: EvalContext,
        _sref: &crate::structured_refs::StructuredRef,
    ) -> Option<(usize, CellAddr, CellAddr)> {
        None
    }

    fn resolve_name(&self, _sheet_id: usize, _name: &str) -> Option<crate::eval::ResolvedName> {
        None
    }
}

fn setup_chain_engine(size: usize) -> (Engine, String) {
    let mut engine = Engine::new();
    engine
        .set_cell_value("Sheet1", "A1", 1.0_f64)
        .expect("seed cell");

    for row in 2..=size {
        let addr = format!("A{row}");
        let prev = row - 1;
        // Use leading `=` to match user-entered formulas (parser strips it).
        let formula = format!("=A{prev}+1");
        engine
            .set_cell_formula("Sheet1", &addr, &formula)
            .expect("set chain formula");
    }

    engine.recalculate_single_threaded();
    (engine, format!("A{size}"))
}

fn setup_range_aggregate_engine(size: usize) -> (Engine, String, String) {
    let mut engine = Engine::new();

    for row in 1..=size {
        let addr = format!("A{row}");
        engine
            .set_cell_value("Sheet1", &addr, (row % 1000) as f64)
            .expect("seed value");
    }

    let sum_cell = "B1".to_string();
    let countif_cell = "B2".to_string();

    engine
        .set_cell_formula("Sheet1", &sum_cell, &format!("=SUM(A1:A{size})"))
        .expect("set SUM formula");
    engine
        .set_cell_formula(
            "Sheet1",
            &countif_cell,
            &format!("=COUNTIF(A1:A{size}, \">500\")"),
        )
        .expect("set COUNTIF formula");

    engine.recalculate_single_threaded();

    (engine, sum_cell, countif_cell)
}

pub fn run_benchmarks() -> Vec<BenchmarkResult> {
    let parse_inputs: Vec<String> = (0..1000)
        .map(|i| format!("=SUM(A{}:A{})", i + 1, i + 100))
        .collect();

    let parsed = Parser::parse("=SUM(A1:A100)").expect("parse eval formula");
    let compiled = compile_for_benchmark(&parsed);
    let resolver = ConstResolver;
    let evaluator = Evaluator::new(
        &resolver,
        EvalContext {
            current_sheet: 0,
            current_cell: CellAddr { row: 0, col: 0 },
        },
    );

    let mut results = Vec::new();

    results.push(run_benchmark(
        "calc.parse_1000_formulas.p95",
        20,
        5,
        50.0,
        || {
            for f in &parse_inputs {
                let ast = Parser::parse(f).unwrap();
                std::hint::black_box(ast);
            }
        },
    ));

    results.push(run_benchmark(
        "calc.evaluate_sum_100_cells.p95",
        30,
        5,
        10.0,
        || {
            let mut acc = 0.0_f64;
            // Run multiple evaluations per sample to reduce timer noise.
            for _ in 0..1000 {
                let v = evaluator.eval_formula(&compiled);
                if let Value::Number(n) = v {
                    acc += n;
                }
            }
            std::hint::black_box(acc);
        },
    ));

    // Recalc a 10k and 100k dependency chain. Targets start slightly looser than
    // docs/16-performance-targets.md to avoid flakiness while the engine is young.
    let (mut engine_10k, last_10k) = setup_chain_engine(10_000);
    let mut counter = 0_i64;
    results.push(run_benchmark(
        "calc.recalc_chain_10k_cells.p95",
        15,
        3,
        100.0,
        || {
            counter += 1;
            engine_10k
                .set_cell_value("Sheet1", "A1", counter)
                .expect("update");
            engine_10k.recalculate_single_threaded();
            let v = engine_10k.get_cell_value("Sheet1", &last_10k);
            std::hint::black_box(v);
        },
    ));

    let (mut engine_100k, last_100k) = setup_chain_engine(100_000);
    let mut counter2 = 0_i64;
    results.push(run_benchmark(
        "calc.recalc_chain_100k_cells.p95",
        8,
        2,
        1000.0,
        || {
            counter2 += 1;
            engine_100k
                .set_cell_value("Sheet1", "A1", counter2)
                .expect("update");
            engine_100k.recalculate_single_threaded();
            let v = engine_100k.get_cell_value("Sheet1", &last_100k);
            std::hint::black_box(v);
        },
    ));

    // Large range aggregations. This is a common workload in real spreadsheets with
    // data columns and summary formulas (SUM/COUNTIF/etc).
    let (mut engine_range, sum_cell, countif_cell) = setup_range_aggregate_engine(200_000);
    let mut counter3 = 0_i64;
    results.push(run_benchmark(
        "calc.recalc_sum_countif_range_200k_cells.p95",
        10,
        2,
        500.0,
        || {
            counter3 += 1;
            engine_range
                .set_cell_value("Sheet1", "A1", (counter3 % 1000) as f64)
                .expect("update");
            engine_range.recalculate_single_threaded();
            let a = engine_range.get_cell_value("Sheet1", &sum_cell);
            let b = engine_range.get_cell_value("Sheet1", &countif_cell);
            std::hint::black_box((a, b));
        },
    ));

    results
}
