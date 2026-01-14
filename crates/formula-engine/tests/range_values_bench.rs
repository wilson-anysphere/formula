//! Optional micro-benchmark for `Engine::get_range_values`.
//!
//! This is intentionally **opt-in** to avoid flakiness / long runtimes in CI.
//! Enable by setting `FORMULA_ENGINE_BENCH_RANGE_READ=1`.

use formula_engine::{Engine, Value};
use formula_model::{CellRef, Range};
use std::hint::black_box;
use std::time::Instant;

const ENV: &str = "FORMULA_ENGINE_BENCH_RANGE_READ";

#[test]
fn bench_get_range_values_vs_get_cell_value() {
    if std::env::var_os(ENV).is_none() {
        return;
    }

    let mut engine = Engine::new();
    // Seed a small amount of data so the benchmark isn't trivially "all blanks".
    engine.set_cell_value("Sheet1", "A1", "Header1").unwrap();
    engine.set_cell_value("Sheet1", "B1", "Header2").unwrap();
    engine.set_cell_value("Sheet1", "A2", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 2.0).unwrap();

    // 200x200 = 40k cells: large enough to show meaningful differences but still manageable.
    let height = 200u32;
    let width = 200u32;
    let range = Range::new(CellRef::new(0, 0), CellRef::new(height - 1, width - 1));

    let start = Instant::now();
    let bulk = engine.get_range_values("Sheet1", range).unwrap();
    black_box(&bulk);
    let bulk_elapsed = start.elapsed();

    // Per-cell reads via `get_cell_value` (includes per-cell A1 parsing + sheet lookup).
    let start = Instant::now();
    let mut per_cell = Vec::with_capacity(height as usize);
    for row in 0..height {
        let mut row_out = Vec::with_capacity(width as usize);
        for col in 0..width {
            let addr = CellRef::new(row, col).to_a1();
            row_out.push(engine.get_cell_value("Sheet1", &addr));
        }
        per_cell.push(row_out);
    }
    black_box(&per_cell);
    let per_cell_elapsed = start.elapsed();

    // Defensive sanity check: both methods must agree on a few cells.
    assert_eq!(bulk[0][0], Value::Text("Header1".to_string()));
    assert_eq!(bulk[1][1], Value::Number(2.0));
    assert_eq!(per_cell[0][0], bulk[0][0]);
    assert_eq!(per_cell[1][1], bulk[1][1]);

    eprintln!(
        "range read benchmark ({}x{}): bulk={:?} per_cell={:?} (env {ENV}=1)",
        height, width, bulk_elapsed, per_cell_elapsed
    );
}
