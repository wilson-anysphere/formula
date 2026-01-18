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

fn failed_benchmark(name: &str, iterations: usize, warmup: usize, target_ms: f64) -> BenchmarkResult {
    BenchmarkResult {
        name: name.to_string(),
        iterations,
        warmup,
        unit: "ms",
        mean: 0.0,
        median: 0.0,
        p95: 0.0,
        p99: 0.0,
        std_dev: 0.0,
        target_ms,
        passed: false,
    }
}

fn failed_benchmark_set(include_sparse_huge_ranges: bool) -> Vec<BenchmarkResult> {
    let mut out = Vec::new();
    out.push(failed_benchmark("calc.parse_1000_formulas.p95", 20, 5, 50.0));
    out.push(failed_benchmark("calc.evaluate_sum_100_cells.p95", 30, 5, 10.0));
    out.push(failed_benchmark("calc.recalc_chain_10k_cells.p95", 15, 3, 100.0));
    out.push(failed_benchmark(
        "calc.recalc_chain_100k_cells.p95",
        8,
        2,
        1000.0,
    ));
    out.push(failed_benchmark(
        "calc.recalc_sum_countif_range_200k_cells.p95",
        10,
        2,
        500.0,
    ));
    out.push(failed_benchmark(
        "calc.recalc_sum_countif_row_array_50k_cells_bytecode.p95",
        8,
        2,
        250.0,
    ));
    out.push(failed_benchmark(
        "calc.recalc_filled_50k_formulas_shared_program.p95",
        8,
        2,
        1000.0,
    ));
    if include_sparse_huge_ranges {
        out.push(failed_benchmark(
            "calc.recalc_sparse_sum_countif_full_column.p95",
            20,
            5,
            50.0,
        ));
    }
    out.push(failed_benchmark(
        "calc.recalc_sum_countif_range_100k_cells_ast.p95",
        4,
        1,
        1000.0,
    ));
    out.push(failed_benchmark(
        "calc.evaluate_sum_sequence_50k_cells.p95",
        6,
        1,
        250.0,
    ));
    out
}

fn percentile(sorted: &[f64], p: f64) -> f64 {
    if sorted.is_empty() {
        return 0.0;
    }
    let idx = ((sorted.len() as f64) * p).floor() as usize;
    sorted[idx.min(sorted.len() - 1)]
}

fn median(sorted: &[f64]) -> f64 {
    if sorted.is_empty() {
        return 0.0;
    }
    sorted[sorted.len() / 2]
}

fn mean(values: &[f64]) -> f64 {
    if values.is_empty() {
        return 0.0;
    }
    values.iter().sum::<f64>() / values.len() as f64
}

fn std_dev(values: &[f64], avg: f64) -> f64 {
    if values.is_empty() {
        return 0.0;
    }
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
    if iterations == 0 {
        debug_assert!(false, "benchmark iterations must be > 0");
    return failed_benchmark(name, iterations, warmup, target_ms);
    }

    for _ in 0..warmup {
        f();
    }

    let mut samples: Vec<f64> = Vec::new();
    if samples.try_reserve_exact(iterations).is_err() {
    debug_assert!(
      false,
      "benchmark sample buffer allocation failed (iterations={iterations})"
    );
        return failed_benchmark(name, iterations, warmup, target_ms);
    }
    for _ in 0..iterations {
        let start = Instant::now();
        f();
        let ms = start.elapsed().as_secs_f64() * 1000.0;
        samples.push(ms);
    }

    samples.sort_by(|a, b| a.total_cmp(b));

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

fn setup_chain_engine(size: usize) -> Option<(Engine, String)> {
    let mut engine = Engine::new();
    if engine
        .set_cell_value("Sheet1", "A1", 1.0_f64)
        .is_err()
    {
        debug_assert!(false, "seed cell");
        return None;
    }

    for row in 2..=size {
        let addr = format!("A{row}");
        let prev = row - 1;
        // Use leading `=` to match user-entered formulas (parser strips it).
        let formula = format!("=A{prev}+1");
        if engine.set_cell_formula("Sheet1", &addr, &formula).is_err() {
            debug_assert!(false, "set chain formula");
            return None;
        }
    }

    engine.recalculate_single_threaded();
    Some((engine, format!("A{size}")))
}

fn setup_range_aggregate_engine(size: usize) -> Option<(Engine, String, String)> {
    let mut engine = Engine::new();

    for row in 1..=size {
        let addr = format!("A{row}");
        if engine
            .set_cell_value("Sheet1", &addr, (row % 1000) as f64)
            .is_err()
        {
            debug_assert!(false, "seed value");
            return None;
        }
    }

    let sum_cell = "B1".to_string();
    let countif_cell = "B2".to_string();

    if engine
        .set_cell_formula("Sheet1", &sum_cell, &format!("=SUM(A1:A{size})"))
        .is_err()
    {
        debug_assert!(false, "set SUM formula");
        return None;
    }
    if engine
        .set_cell_formula(
            "Sheet1",
            &countif_cell,
            &format!("=COUNTIF(A1:A{size}, \">500\")"),
        )
        .is_err()
    {
        debug_assert!(false, "set COUNTIF formula");
        return None;
    }

    engine.recalculate_single_threaded();

    Some((engine, sum_cell, countif_cell))
}

fn setup_sparse_huge_range_engine() -> Option<(Engine, String, String)> {
    let mut engine = Engine::new();

    // Sparse column with a couple of values spread far apart. This models sheets where users
    // write `=SUM(A:A)`/`=COUNTIF(A:A, ...)` over mostly-empty columns.
    if engine.set_cell_value("Sheet1", "A1", 1.0_f64).is_err() {
        debug_assert!(false, "seed value");
        return None;
    }
    if engine.set_cell_value("Sheet1", "A500000", 2.0_f64).is_err() {
        debug_assert!(false, "seed value");
        return None;
    }
    if engine
        .set_cell_value("Sheet1", "A1048576", 3.0_f64)
        .is_err()
    {
        debug_assert!(false, "seed value");
        return None;
    }

    let sum_cell = "B1".to_string();
    let countif_cell = "B2".to_string();

    if engine
        .set_cell_formula("Sheet1", &sum_cell, "=SUM(A:A)")
        .is_err()
    {
        debug_assert!(false, "set SUM formula");
        return None;
    }
    if engine
        .set_cell_formula("Sheet1", &countif_cell, "=COUNTIF(A:A, 0)")
        .is_err()
    {
        debug_assert!(false, "set COUNTIF formula");
        return None;
    }

    engine.recalculate_single_threaded();

    Some((engine, sum_cell, countif_cell))
}

fn setup_range_aggregate_engine_ast(size: usize) -> Option<(Engine, String, String)> {
    let mut engine = Engine::new();
    // This benchmark exists specifically to cover the AST evaluator (and the scalar/array
    // aggregation paths it exercises). Disable the bytecode backend so improvements/regressions in
    // the AST path remain visible even as bytecode coverage expands over time.
    engine.set_bytecode_enabled(false);

    for row in 1..=size {
        let addr = format!("A{row}");
        if engine
            .set_cell_value("Sheet1", &addr, (row % 1000) as f64)
            .is_err()
        {
            debug_assert!(false, "seed value");
            return None;
        }
    }

    let sum_cell = "C1".to_string();
    let countif_cell = "C2".to_string();

    // Wrap in LET to model real-world non-bytecode sheets (LET/LAMBDA heavy formulas). Bytecode is
    // explicitly disabled above, so this stays on the AST evaluator.
    if engine
        .set_cell_formula("Sheet1", &sum_cell, &format!("=LET(r,A1:A{size},SUM(r))"))
        .is_err()
    {
        debug_assert!(false, "set SUM/LET formula");
        return None;
    }
    if engine
        .set_cell_formula(
            "Sheet1",
            &countif_cell,
            &format!("=LET(r,A1:A{size},COUNTIF(r, \">500\"))"),
        )
        .is_err()
    {
        debug_assert!(false, "set COUNTIF/LET formula");
        return None;
    }

    engine.recalculate_single_threaded();

    Some((engine, sum_cell, countif_cell))
}

fn setup_bytecode_array_aggregate_engine(size: usize) -> Option<(Engine, String)> {
    let mut engine = Engine::new();
    if engine.set_cell_value("Sheet1", "A1", 0.0_f64).is_err() {
        debug_assert!(false, "seed cell");
        return None;
    }

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
    if engine
        .set_cell_formula("Sheet1", &out_cell, &formula)
        .is_err()
    {
        debug_assert!(false, "set bytecode array aggregate formula");
        return None;
    }

    // This benchmark is intended to cover the bytecode backend. If the formula ever stops being
    // bytecode-eligible, fail loudly so we don't silently lose perf coverage.
    if engine.bytecode_program_count() != 1 {
        debug_assert!(false, "expected bytecode program count to be 1");
        return None;
    }

    engine.recalculate_single_threaded();
    Some((engine, out_cell))
}

fn setup_filled_formula_recalc_engine(size: usize) -> Option<(Engine, String)> {
    let mut engine = Engine::new();
    if engine.set_cell_value("Sheet1", "A1", 0.0_f64).is_err() {
        debug_assert!(false, "seed shared precedent");
        return None;
    }

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
        if engine.set_cell_formula("Sheet1", &addr, &formula).is_err() {
            debug_assert!(false, "set filled formula");
            return None;
        }
    }

    // Ensure the "filled formula" block is actually benefiting from bytecode program interning;
    // otherwise this benchmark won't catch regressions in shared-formula caching behavior.
    let stats = engine.bytecode_compile_stats();
    if stats.total_formula_cells != size || stats.compiled != size || engine.bytecode_program_count() != 1 {
        debug_assert!(false, "filled formula benchmark expected 1 shared bytecode program");
        return None;
    }

    engine.recalculate_single_threaded();
    Some((engine, format!("B{size}")))
}

pub fn run_benchmarks() -> Vec<BenchmarkResult> {
    let include_sparse_huge_ranges =
        std::env::var("FORMULA_ENGINE_BENCH_SPARSE_HUGE_RANGES").is_ok();

  let mut parse_inputs: Vec<String> = Vec::new();
  if parse_inputs.try_reserve_exact(1000).is_err() {
    debug_assert!(false, "benchmark parse input allocation failed (count=1000)");
    return failed_benchmark_set(include_sparse_huge_ranges);
  }
  for i in 0..1000 {
    parse_inputs.push(format!("=SUM(A{}:A{})", i + 1, i + 100));
  }

    let parsed = match Parser::parse("=SUM(A1:A100)") {
        Ok(v) => v,
        Err(err) => {
            debug_assert!(false, "parse eval formula: {err:?}");
            return failed_benchmark_set(include_sparse_huge_ranges);
        }
    };
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
  let expected_results = if include_sparse_huge_ranges { 10 } else { 9 };
  if results.try_reserve_exact(expected_results).is_err() {
    debug_assert!(false, "benchmark results allocation failed (count={expected_results})");
    return failed_benchmark_set(include_sparse_huge_ranges);
  }

    results.push(run_benchmark(
        "calc.parse_1000_formulas.p95",
        20,
        5,
        50.0,
        || {
            for f in &parse_inputs {
                if let Ok(ast) = Parser::parse(f) {
                    std::hint::black_box(ast);
                } else {
                    debug_assert!(false, "parse failed for benchmark input");
                }
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
    if let Some((mut engine_10k, last_10k)) = setup_chain_engine(10_000) {
        let mut counter = 0_i64;
        results.push(run_benchmark(
            "calc.recalc_chain_10k_cells.p95",
            15,
            3,
            100.0,
            || {
                counter += 1;
                if engine_10k.set_cell_value("Sheet1", "A1", counter).is_err() {
                    debug_assert!(false, "update");
                    return;
                }
                engine_10k.recalculate_single_threaded();
                let v = engine_10k.get_cell_value("Sheet1", &last_10k);
                std::hint::black_box(v);
            },
        ));
    } else {
        results.push(failed_benchmark("calc.recalc_chain_10k_cells.p95", 15, 3, 100.0));
    }

    if let Some((mut engine_100k, last_100k)) = setup_chain_engine(100_000) {
        let mut counter2 = 0_i64;
        results.push(run_benchmark(
            "calc.recalc_chain_100k_cells.p95",
            8,
            2,
            1000.0,
            || {
                counter2 += 1;
                if engine_100k.set_cell_value("Sheet1", "A1", counter2).is_err() {
                    debug_assert!(false, "update");
                    return;
                }
                engine_100k.recalculate_single_threaded();
                let v = engine_100k.get_cell_value("Sheet1", &last_100k);
                std::hint::black_box(v);
            },
        ));
    } else {
        results.push(failed_benchmark("calc.recalc_chain_100k_cells.p95", 8, 2, 1000.0));
    }

    // Large range aggregations. This is a common workload in real spreadsheets with
    // data columns and summary formulas (SUM/COUNTIF/etc).
    if let Some((mut engine_range, sum_cell, countif_cell)) = setup_range_aggregate_engine(200_000) {
        let mut counter3 = 0_i64;
        results.push(run_benchmark(
            "calc.recalc_sum_countif_range_200k_cells.p95",
            10,
            2,
            500.0,
            || {
                counter3 += 1;
                if engine_range
                    .set_cell_value("Sheet1", "A1", (counter3 % 1000) as f64)
                    .is_err()
                {
                    debug_assert!(false, "update");
                    return;
                }
                engine_range.recalculate_single_threaded();
                let a = engine_range.get_cell_value("Sheet1", &sum_cell);
                let b = engine_range.get_cell_value("Sheet1", &countif_cell);
                std::hint::black_box((a, b));
            },
        ));
    } else {
        results.push(failed_benchmark(
            "calc.recalc_sum_countif_range_200k_cells.p95",
            10,
            2,
            500.0,
        ));
    }

    // Bytecode large array aggregation: this covers in-memory array aggregates (not reference
    // aggregates), which happen for array-producing functions and array literals.
    if let Some((mut engine_bytecode_array, out_cell)) = setup_bytecode_array_aggregate_engine(50_000) {
        let mut counter_array = 0_i64;
        results.push(run_benchmark(
            "calc.recalc_sum_countif_row_array_50k_cells_bytecode.p95",
            8,
            2,
            250.0,
            || {
                counter_array += 1;
                if engine_bytecode_array
                    .set_cell_value("Sheet1", "A1", (counter_array % 1000) as f64)
                    .is_err()
                {
                    debug_assert!(false, "update");
                    return;
                }
                engine_bytecode_array.recalculate_single_threaded();
                let v = engine_bytecode_array.get_cell_value("Sheet1", &out_cell);
                std::hint::black_box(v);
            },
        ));
    } else {
        results.push(failed_benchmark(
            "calc.recalc_sum_countif_row_array_50k_cells_bytecode.p95",
            8,
            2,
            250.0,
        ));
    }

    // Many filled formulas: this guards against regressions where recalculating a large block of
    // shared-formula cells reintroduces per-cell deep clones / allocations (e.g. cloning a full AST
    // for every dirty cell even when bytecode evaluation is used).
    if let Some((mut engine_filled, filled_out_cell)) = setup_filled_formula_recalc_engine(50_000) {
        let mut counter_filled = 0_i64;
        results.push(run_benchmark(
            "calc.recalc_filled_50k_formulas_shared_program.p95",
            8,
            2,
            1000.0,
            || {
                counter_filled += 1;
                if engine_filled
                    .set_cell_value("Sheet1", "A1", counter_filled)
                    .is_err()
                {
                    debug_assert!(false, "update shared precedent");
                    return;
                }
                engine_filled.recalculate_single_threaded();
                let v = engine_filled.get_cell_value("Sheet1", &filled_out_cell);
                std::hint::black_box(v);
            },
        ));
    } else {
        results.push(failed_benchmark(
            "calc.recalc_filled_50k_formulas_shared_program.p95",
            8,
            2,
            1000.0,
        ));
    }

    // Opt-in benchmark for huge sparse ranges (`A:A`-style). Enable via:
    //   FORMULA_ENGINE_BENCH_SPARSE_HUGE_RANGES=1 cargo run -p formula-engine --bin perf_bench
    if include_sparse_huge_ranges {
        if let Some((mut engine_sparse, sum_cell, countif_cell)) = setup_sparse_huge_range_engine() {
            let mut counter_sparse = 0_i64;
            results.push(run_benchmark(
                "calc.recalc_sparse_sum_countif_full_column.p95",
                20,
                5,
                50.0,
                || {
                    counter_sparse += 1;
                    if engine_sparse
                        .set_cell_value("Sheet1", "A1", (counter_sparse % 10) as f64)
                        .is_err()
                    {
                        debug_assert!(false, "update");
                        return;
                    }
                    engine_sparse.recalculate_single_threaded();
                    let a = engine_sparse.get_cell_value("Sheet1", &sum_cell);
                    let b = engine_sparse.get_cell_value("Sheet1", &countif_cell);
                    std::hint::black_box((a, b));
                },
            ));
        } else {
            results.push(failed_benchmark(
                "calc.recalc_sparse_sum_countif_full_column.p95",
                20,
                5,
                50.0,
            ));
        }
    }

    // Same workload as the range aggregation benchmark, but forced through the AST evaluator so we
    // can catch regressions in non-bytecode aggregation performance (e.g. LET/LAMBDA-heavy sheets).
    if let Some((mut engine_range_ast, sum_cell_ast, countif_cell_ast)) =
        setup_range_aggregate_engine_ast(100_000)
    {
        let mut counter_ast = 0_i64;
        results.push(run_benchmark(
            "calc.recalc_sum_countif_range_100k_cells_ast.p95",
            4,
            1,
            1000.0,
            || {
                counter_ast += 1;
                if engine_range_ast
                    .set_cell_value("Sheet1", "A1", (counter_ast % 1000) as f64)
                    .is_err()
                {
                    debug_assert!(false, "update");
                    return;
                }
                engine_range_ast.recalculate_single_threaded();
                let a = engine_range_ast.get_cell_value("Sheet1", &sum_cell_ast);
                let b = engine_range_ast.get_cell_value("Sheet1", &countif_cell_ast);
                std::hint::black_box((a, b));
            },
        ));
    } else {
        results.push(failed_benchmark(
            "calc.recalc_sum_countif_range_100k_cells_ast.p95",
            4,
            1,
            1000.0,
        ));
    }

    // Dynamic array aggregation: SUM(SEQUENCE(n)) exercises array aggregation in the AST evaluator
    // (SEQUENCE is not currently bytecode-eligible).
    match Parser::parse("=SUM(SEQUENCE(50000))") {
        Ok(parsed_seq) => {
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
        }
        Err(err) => {
            debug_assert!(false, "parse SEQUENCE benchmark: {err:?}");
            results.push(failed_benchmark(
                "calc.evaluate_sum_sequence_50k_cells.p95",
                6,
                1,
                250.0,
            ));
        }
    }

    results
}

#[cfg(test)]
mod tests {
  use super::{mean, median, percentile, std_dev};

  #[test]
  fn percentile_uses_floor_indexing() {
    let mut sorted: Vec<f64> = Vec::new();
    if sorted.try_reserve_exact(10).is_err() {
      panic!("allocation failed (percentile test sorted)");
    }
    for n in 0..10 {
      sorted.push(n as f64);
    }
    assert_eq!(percentile(&sorted, 0.0), 0.0);
    assert_eq!(percentile(&sorted, 0.5), 5.0);
    assert_eq!(percentile(&sorted, 0.95), 9.0);
    assert_eq!(percentile(&sorted, 0.99), 9.0);
  }

  #[test]
  fn median_selects_upper_middle_for_even_lengths() {
    let mut sorted: Vec<f64> = vec![1.0, 2.0, 3.0, 4.0];
    sorted.sort_by(|a, b| a.total_cmp(b));
    assert_eq!(median(&sorted), 3.0);
  }

  #[test]
  fn mean_and_std_dev_are_sane_for_constant_input() {
    let mut values: Vec<f64> = Vec::new();
    if values.try_reserve_exact(10).is_err() {
      panic!("allocation failed (mean/std_dev constant values)");
    }
    values.resize(10, 2.0_f64);
    let avg = mean(&values);
    assert_eq!(avg, 2.0);
    assert_eq!(std_dev(&values, avg), 0.0);
  }
}
