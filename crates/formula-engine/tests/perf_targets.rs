//! Optional performance target assertions.
//!
//! These checks are intentionally **opt-in** to avoid flakiness across CI / developer machines.
//! Enable by setting `FORMULA_ENGINE_ENFORCE_PERF_TARGETS=1`.
//!
//! Targets are derived from `docs/16-performance-targets.md`.

use formula_engine::bytecode::{
    parse_formula, CalcGraph, CellCoord, ColumnarGrid, FormulaCell, RecalcEngine,
};
use std::time::{Duration, Instant};

const ENV: &str = "FORMULA_ENGINE_ENFORCE_PERF_TARGETS";

#[test]
fn perf_100k_recalc_under_100ms_reference() {
    if std::env::var_os(ENV).is_none() {
        return;
    }

    let n = 100_000usize;
    let (engine, graph, mut grid) = build_independent_workbook(n);

    // Warm up.
    engine.recalc(&graph, &mut grid);

    let start = Instant::now();
    engine.recalc(&graph, &mut grid);
    let elapsed = start.elapsed();

    assert!(
        elapsed < Duration::from_millis(100),
        "100k independent recalc took {:?}, expected < 100ms (see docs/16-performance-targets.md)",
        elapsed
    );
}

fn build_independent_workbook(n: usize) -> (RecalcEngine, CalcGraph, ColumnarGrid) {
    let engine = RecalcEngine::new();
    let rows = n as i32;
    let cols = 4;
    let mut grid = ColumnarGrid::new(rows, cols);

    for row in 0..rows {
        grid.set_number(CellCoord::new(row, 0), row as f64);
        grid.set_number(CellCoord::new(row, 1), (row as f64) * 2.0);
    }

    // Formula column: =A1+B1 (normalized and shared across rows).
    let template_origin = CellCoord::new(0, 2);
    let template = parse_formula("=A1+B1", template_origin).unwrap();

    let mut cells = Vec::with_capacity(n);
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
