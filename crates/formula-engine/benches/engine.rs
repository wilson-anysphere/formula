use formula_engine::run_benchmarks;

fn main() {
    // `cargo bench -p formula-engine` will execute this binary.
    //
    // The CI harness uses `cargo run --bin perf_bench` (JSON output). This file is
    // primarily for local developer iteration with a familiar `cargo bench` entrypoint.
    for r in run_benchmarks() {
        println!(
            "{:<32} p95={:>8.3}ms  target={:>8.3}ms  {}",
            r.name,
            r.p95,
            r.target_ms,
            if r.passed { "PASS" } else { "FAIL" }
        );
    }
}

