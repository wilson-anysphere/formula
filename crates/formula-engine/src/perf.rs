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
        SheetReference::SheetRange(_start, _end) => SheetReference::SheetRange(0, 0),
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
    ) -> Result<Vec<(usize, CellAddr, CellAddr)>, crate::ErrorKind> {
        Err(crate::ErrorKind::Name)
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

fn setup_sparse_huge_range_engine() -> (Engine, String, String) {
    let mut engine = Engine::new();

    // Sparse column with a couple of values spread far apart. This models sheets where users
    // write `=SUM(A:A)`/`=COUNTIF(A:A, ...)` over mostly-empty columns.
    engine
        .set_cell_value("Sheet1", "A1", 1.0_f64)
        .expect("seed value");
    engine
        .set_cell_value("Sheet1", "A500000", 2.0_f64)
        .expect("seed value");
    engine
        .set_cell_value("Sheet1", "A1048576", 3.0_f64)
        .expect("seed value");

    let sum_cell = "B1".to_string();
    let countif_cell = "B2".to_string();

    engine
        .set_cell_formula("Sheet1", &sum_cell, "=SUM(A:A)")
        .expect("set SUM formula");
    engine
        .set_cell_formula("Sheet1", &countif_cell, "=COUNTIF(A:A, 0)")
        .expect("set COUNTIF formula");

    engine.recalculate_single_threaded();

    (engine, sum_cell, countif_cell)
}

fn setup_range_aggregate_engine_ast(size: usize) -> (Engine, String, String) {
    let mut engine = Engine::new();
    // This benchmark exists specifically to cover the AST evaluator (and the scalar/array
    // aggregation paths it exercises). Disable the bytecode backend so improvements/regressions in
    // the AST path remain visible even as bytecode coverage expands over time.
    engine.set_bytecode_enabled(false);

    for row in 1..=size {
        let addr = format!("A{row}");
        engine
            .set_cell_value("Sheet1", &addr, (row % 1000) as f64)
            .expect("seed value");
    }

    let sum_cell = "C1".to_string();
    let countif_cell = "C2".to_string();

    // Wrap in LET to model real-world non-bytecode sheets (LET/LAMBDA heavy formulas). Bytecode is
    // explicitly disabled above, so this stays on the AST evaluator.
    engine
        .set_cell_formula("Sheet1", &sum_cell, &format!("=LET(r,A1:A{size},SUM(r))"))
        .expect("set SUM/LET formula");
    engine
        .set_cell_formula(
            "Sheet1",
            &countif_cell,
            &format!("=LET(r,A1:A{size},COUNTIF(r, \">500\"))"),
        )
        .expect("set COUNTIF/LET formula");

    engine.recalculate_single_threaded();

    (engine, sum_cell, countif_cell)
}

fn setup_bytecode_array_aggregate_engine(size: usize) -> (Engine, String) {
    let mut engine = Engine::new();
    engine
        .set_cell_value("Sheet1", "A1", 0.0_f64)
        .expect("seed cell");

    // Important: keep the output cell *outside* the referenced row range (`1:{size}`).
    //
    // Range nodes participate in calc ordering, and placing the formula inside the range would
    // create a trivial circular reference (cell -> range node -> cell) which the engine resolves
    // to `0` without evaluating the formula (since iterative calculation is disabled by default).
    let out_cell = format!("B{}", size + 1);
    let half = size / 2;
    // Use ROW over a row-range to produce a large in-memory array on the bytecode backend, then
    // aggregate it. This exercises bytecode `Value::Array` aggregate fast paths (SUM/COUNTIF).
    let formula = format!("=LET(x,A1,r,ROW(1:{size}),SUM(r)+COUNTIF(r, \">{half}\")+x)");
    engine
        .set_cell_formula("Sheet1", &out_cell, &formula)
        .expect("set bytecode array aggregate formula");

    // This benchmark is intended to cover the bytecode backend. If the formula ever stops being
    // bytecode-eligible, fail loudly so we don't silently lose perf coverage.
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    (engine, out_cell)
}

fn setup_filled_formula_recalc_engine(size: usize) -> (Engine, String) {
    let mut engine = Engine::new();
    engine
        .set_cell_value("Sheet1", "A1", 0.0_f64)
        .expect("seed shared precedent");

    // Create a large filled-down block of formulas that share a single bytecode program via the
    // normalized-key cache.
    //
    // Each formula references:
    // - a row-relative cell (`A{row}`), which gives us a fill-down pattern (normalized to offsets),
    // - plus a shared absolute input (`$A$1`), which lets us dirty all formulas each iteration by
    //   editing a single cell.
    for row in 1..=size {
        let addr = format!("B{row}");
        let formula = format!("=A{row}+$A$1");
        engine
            .set_cell_formula("Sheet1", &addr, &formula)
            .expect("set filled formula");
    }

    // Ensure the "filled formula" block is actually benefiting from bytecode program interning;
    // otherwise this benchmark won't catch regressions in shared-formula caching behavior.
    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, size);
    assert_eq!(stats.compiled, size);
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();
    (engine, format!("B{size}"))
}

pub fn run_benchmarks() -> Vec<BenchmarkResult> {
    let parse_inputs: Vec<String> = (0..1000)
        .map(|i| format!("=SUM(A{}:A{})", i + 1, i + 100))
        .collect();

    let parsed = Parser::parse("=SUM(A1:A100)").expect("parse eval formula");
    let compiled = compile_for_benchmark(&parsed);
    let resolver = ConstResolver;
    let recalc_ctx = crate::eval::RecalcContext::new(0);
    let evaluator = Evaluator::new(
        &resolver,
        EvalContext {
            current_sheet: 0,
            current_cell: CellAddr { row: 0, col: 0 },
        },
        &recalc_ctx,
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

    // Bytecode large array aggregation: this covers in-memory array aggregates (not reference
    // aggregates), which happen for array-producing functions and array literals.
    let (mut engine_bytecode_array, out_cell) = setup_bytecode_array_aggregate_engine(50_000);
    let mut counter_array = 0_i64;
    results.push(run_benchmark(
        "calc.recalc_sum_countif_row_array_50k_cells_bytecode.p95",
        8,
        2,
        250.0,
        || {
            counter_array += 1;
            engine_bytecode_array
                .set_cell_value("Sheet1", "A1", (counter_array % 1000) as f64)
                .expect("update");
            engine_bytecode_array.recalculate_single_threaded();
            let v = engine_bytecode_array.get_cell_value("Sheet1", &out_cell);
            std::hint::black_box(v);
        },
    ));

    // Many filled formulas: this guards against regressions where recalculating a large block of
    // shared-formula cells reintroduces per-cell deep clones / allocations (e.g. cloning a full AST
    // for every dirty cell even when bytecode evaluation is used).
    let (mut engine_filled, filled_out_cell) = setup_filled_formula_recalc_engine(50_000);
    let mut counter_filled = 0_i64;
    results.push(run_benchmark(
        "calc.recalc_filled_50k_formulas_shared_program.p95",
        8,
        2,
        1000.0,
        || {
            counter_filled += 1;
            engine_filled
                .set_cell_value("Sheet1", "A1", counter_filled)
                .expect("update shared precedent");
            engine_filled.recalculate_single_threaded();
            let v = engine_filled.get_cell_value("Sheet1", &filled_out_cell);
            std::hint::black_box(v);
        },
    ));

    // Opt-in benchmark for huge sparse ranges (`A:A`-style). Enable via:
    //   FORMULA_ENGINE_BENCH_SPARSE_HUGE_RANGES=1 cargo run -p formula-engine --bin perf_bench
    if std::env::var("FORMULA_ENGINE_BENCH_SPARSE_HUGE_RANGES").is_ok() {
        let (mut engine_sparse, sum_cell, countif_cell) = setup_sparse_huge_range_engine();
        let mut counter_sparse = 0_i64;
        results.push(run_benchmark(
            "calc.recalc_sparse_sum_countif_full_column.p95",
            20,
            5,
            50.0,
            || {
                counter_sparse += 1;
                engine_sparse
                    .set_cell_value("Sheet1", "A1", (counter_sparse % 10) as f64)
                    .expect("update");
                engine_sparse.recalculate_single_threaded();
                let a = engine_sparse.get_cell_value("Sheet1", &sum_cell);
                let b = engine_sparse.get_cell_value("Sheet1", &countif_cell);
                std::hint::black_box((a, b));
            },
        ));
    }

    // Same workload as the range aggregation benchmark, but forced through the AST evaluator so we
    // can catch regressions in non-bytecode aggregation performance (e.g. LET/LAMBDA-heavy sheets).
    let (mut engine_range_ast, sum_cell_ast, countif_cell_ast) =
        setup_range_aggregate_engine_ast(100_000);
    let mut counter_ast = 0_i64;
    results.push(run_benchmark(
        "calc.recalc_sum_countif_range_100k_cells_ast.p95",
        4,
        1,
        1000.0,
        || {
            counter_ast += 1;
            engine_range_ast
                .set_cell_value("Sheet1", "A1", (counter_ast % 1000) as f64)
                .expect("update");
            engine_range_ast.recalculate_single_threaded();
            let a = engine_range_ast.get_cell_value("Sheet1", &sum_cell_ast);
            let b = engine_range_ast.get_cell_value("Sheet1", &countif_cell_ast);
            std::hint::black_box((a, b));
        },
    ));

    // Dynamic array aggregation: SUM(SEQUENCE(n)) exercises array aggregation in the AST evaluator
    // (SEQUENCE is not currently bytecode-eligible).
    let parsed_seq = Parser::parse("=SUM(SEQUENCE(50000))").expect("parse SEQUENCE benchmark");
    let compiled_seq = compile_for_benchmark(&parsed_seq);
    results.push(run_benchmark(
        "calc.evaluate_sum_sequence_50k_cells.p95",
        6,
        1,
        250.0,
        || {
            let v = evaluator.eval_formula(&compiled_seq);
            std::hint::black_box(v);
        },
    ));

    results
}
