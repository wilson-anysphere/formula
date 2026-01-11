// Criterion + rand rely on OS functionality that is not available on `wasm32-unknown-unknown`.
// Provide a no-op main so `cargo check -p formula-engine --target wasm32-unknown-unknown --benches`
// succeeds (useful for CI sanity checks), while keeping the native benchmark intact.
#[cfg(target_arch = "wasm32")]
fn main() {}

#[cfg(not(target_arch = "wasm32"))]
use criterion::{criterion_group, criterion_main, BatchSize, BenchmarkId, Criterion};
#[cfg(not(target_arch = "wasm32"))]
use formula_engine::bytecode::{
    eval_ast, parse_formula, BytecodeCache, CalcGraph, CellCoord, ColumnarGrid, FormulaCell,
    RecalcEngine, Vm,
};
#[cfg(not(target_arch = "wasm32"))]
use rand::{rngs::StdRng, Rng, SeedableRng};

#[cfg(not(target_arch = "wasm32"))]
fn bench_parse(c: &mut Criterion) {
    let mut group = c.benchmark_group("parse");
    let origin = CellCoord::new(10, 3);
    let formula = "=SUM(A1:A100)+B1*C1-42/7";
    group.bench_function("simple", |b| {
        b.iter(|| parse_formula(formula, origin).unwrap())
    });
    group.finish();
}

#[cfg(not(target_arch = "wasm32"))]
fn bench_compile(c: &mut Criterion) {
    let origin = CellCoord::new(10, 3);
    let formula = "=SUM(A1:A100)+B1*C1-42/7";
    let expr = parse_formula(formula, origin).unwrap();
    c.bench_function("compile_cold", |b| {
        b.iter_batched(
            BytecodeCache::new,
            |cache| {
                cache.get_or_compile(&expr);
            },
            BatchSize::SmallInput,
        )
    });
    let cache = BytecodeCache::new();
    cache.get_or_compile(&expr);
    c.bench_function("compile_cache_hit", |b| {
        b.iter(|| cache.get_or_compile(&expr))
    });
}

#[cfg(not(target_arch = "wasm32"))]
fn bench_eval_single(c: &mut Criterion) {
    let mut grid = ColumnarGrid::new(200, 10);
    for row in 0..100 {
        grid.set_number(CellCoord::new(row, 0), row as f64);
    }
    grid.set_number(CellCoord::new(0, 1), 2.0);
    grid.set_number(CellCoord::new(0, 2), 3.0);

    let origin = CellCoord::new(0, 3);
    let expr = parse_formula("=SUM(A1:A100)+B1*C1", origin).unwrap();
    let cache = BytecodeCache::new();
    let program = cache.get_or_compile(&expr);

    let mut vm = Vm::with_capacity(32);
    c.bench_function("eval_bytecode_single", |b| {
        b.iter(|| vm.eval(&program, &grid, origin))
    });
    c.bench_function("eval_ast_single", |b| {
        b.iter(|| eval_ast(&expr, &grid, origin))
    });
}

#[cfg(not(target_arch = "wasm32"))]
fn bench_recalc(c: &mut Criterion) {
    let mut group = c.benchmark_group("recalc");

    for &n in &[10_000usize, 100_000usize] {
        group.bench_with_input(BenchmarkId::new("independent", n), &n, |b, &n| {
            b.iter_batched(
                || build_independent_workbook(n),
                |(engine, graph, mut grid)| {
                    engine.recalc(&graph, &mut grid);
                },
                BatchSize::SmallInput,
            );
        });
    }

    group.finish();
}

#[cfg(not(target_arch = "wasm32"))]
fn build_independent_workbook(n: usize) -> (RecalcEngine, CalcGraph, ColumnarGrid) {
    let engine = RecalcEngine::new();
    let rows = n as i32;
    let cols = 4;
    let mut grid = ColumnarGrid::new(rows, cols);

    // Fill inputs.
    let mut rng = StdRng::seed_from_u64(123);
    for row in 0..rows {
        grid.set_number(CellCoord::new(row, 0), rng.gen_range(0.0..1000.0));
        grid.set_number(CellCoord::new(row, 1), rng.gen_range(0.0..1000.0));
    }

    // Column C is formula: =A{row}+B{row}
    let mut cells = Vec::with_capacity(n);
    let template_origin = CellCoord::new(0, 2);
    let template = parse_formula("=A1+B1", template_origin).unwrap();
    for row in 0..rows {
        let coord = CellCoord::new(row, 2);
        cells.push(FormulaCell {
            coord,
            expr: template.clone(),
        });
    }
    let graph = engine.build_graph(cells);
    (engine, graph, grid)
}

#[cfg(not(target_arch = "wasm32"))]
criterion_group!(
    benches,
    bench_parse,
    bench_compile,
    bench_eval_single,
    bench_recalc
);
#[cfg(not(target_arch = "wasm32"))]
criterion_main!(benches);
