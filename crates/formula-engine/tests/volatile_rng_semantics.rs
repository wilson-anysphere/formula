use formula_engine::eval::parse_a1;
use formula_engine::{Engine, Value};

fn setup_rng_sheet(engine: &mut Engine) {
    // Simple volatile RNG cells.
    engine.set_cell_formula("Sheet1", "A1", "=RAND()").unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=RANDBETWEEN(1, 1000000)")
        .unwrap();

    // Nested volatile calls within a single cell evaluation. The engine should treat each call as
    // a distinct deterministic draw, scoped to the cell evaluation.
    engine
        .set_cell_formula("Sheet1", "A3", "=RAND()+RAND()")
        .unwrap();
    // Stronger check for "distinct draws": if RAND() were incorrectly cached within a single cell
    // evaluation, this would always be zero.
    engine
        .set_cell_formula("Sheet1", "A5", "=RAND()-RAND()")
        .unwrap();

    // LET should evaluate its bound expression exactly once and reuse the value for each reference
    // of the bound name (Excel semantics).
    engine
        .set_cell_formula("Sheet1", "A4", "=LET(x,RAND(),x=x)")
        .unwrap();

    // A dependent cell to ensure volatile cell results are stable when referenced.
    engine.set_cell_formula("Sheet1", "B1", "=A1+A2").unwrap();
    // Repeated references to the same volatile cell should observe a consistent value within a
    // single recalc pass.
    engine.set_cell_formula("Sheet1", "B2", "=A1-A1").unwrap();

    // A spilled volatile RNG result: 2x2, integer output.
    engine
        .set_cell_formula("Sheet1", "C1", "=RANDARRAY(2,2,1,1000000,TRUE)")
        .unwrap();
    // Force a consumer of the spilled range to ensure spill scheduling/dependencies are exercised.
    engine
        .set_cell_formula("Sheet1", "E1", "=SUM(C1#)")
        .unwrap();
}

fn assert_rand_unit_interval(value: &Value, label: &str) {
    match value {
        Value::Number(n) => {
            assert!(
                *n >= 0.0 && *n < 1.0,
                "expected {label} to be in [0,1), got {n}"
            );
        }
        other => panic!("expected {label} to be a number, got {other:?}"),
    }
}

fn assert_randbetween_bounds(value: &Value, label: &str, low: f64, high: f64) {
    match value {
        Value::Number(n) => {
            assert!(n.is_finite(), "expected {label} to be finite, got {n}");
            assert!(
                (n.fract()).abs() < 1e-9,
                "expected {label} to be an integer, got {n}"
            );
            assert!(
                *n >= low && *n <= high,
                "expected {label} in [{low},{high}], got {n}"
            );
        }
        other => panic!("expected {label} to be a number, got {other:?}"),
    }
}

fn assert_rand_sum_bounds(value: &Value, label: &str) {
    match value {
        Value::Number(n) => {
            assert!(n.is_finite(), "expected {label} to be finite, got {n}");
            assert!(
                *n >= 0.0 && *n < 2.0,
                "expected {label} to be in [0,2), got {n}"
            );
        }
        other => panic!("expected {label} to be a number, got {other:?}"),
    }
}

fn snapshot(engine: &Engine) -> Vec<Value> {
    // Include the spill cells explicitly so equality checks catch spill scheduling differences.
    [
        "A1", "A2", "A3", "A4", "A5", "B1", "B2", "C1", "D1", "C2", "D2", "E1",
    ]
    .into_iter()
    .map(|addr| engine.get_cell_value("Sheet1", addr))
    .collect()
}

#[test]
fn volatile_rng_semantics_are_stable_within_recalc_and_order_independent() {
    let mut single = Engine::new();
    setup_rng_sheet(&mut single);
    single.recalculate_single_threaded();

    // Structural spill footprint should match the RANDARRAY shape.
    assert_eq!(
        single.spill_range("Sheet1", "C1"),
        Some((parse_a1("C1").unwrap(), parse_a1("D2").unwrap()))
    );
    assert_eq!(
        single.spill_range("Sheet1", "D2"),
        Some((parse_a1("C1").unwrap(), parse_a1("D2").unwrap()))
    );

    // Within a single recalc pass, LET bindings should be stable (Excel semantics).
    assert_eq!(single.get_cell_value("Sheet1", "A4"), Value::Bool(true));
    // Repeated references to the same volatile cell are stable within a recalc pass.
    assert_eq!(single.get_cell_value("Sheet1", "B2"), Value::Number(0.0));
    // Multiple RAND() calls within a single cell evaluation should produce distinct draws.
    // We avoid asserting this is *always* non-zero because collisions are theoretically possible,
    // but we do assert basic invariants and later require observing a non-zero result across a few
    // recalcs.
    match single.get_cell_value("Sheet1", "A5") {
        Value::Number(n) => {
            assert!(
                n.is_finite(),
                "expected RAND()-RAND() to be finite, got {n}"
            );
            assert!(
                n > -1.0 && n < 1.0,
                "expected RAND()-RAND() in (-1,1), got {n}"
            );
        }
        other => panic!("expected RAND()-RAND() to be a number, got {other:?}"),
    }

    // Basic bounds / type invariants for each RNG function.
    assert_rand_unit_interval(&single.get_cell_value("Sheet1", "A1"), "RAND()");
    assert_randbetween_bounds(
        &single.get_cell_value("Sheet1", "A2"),
        "RANDBETWEEN()",
        1.0,
        1_000_000.0,
    );
    assert_rand_sum_bounds(&single.get_cell_value("Sheet1", "A3"), "RAND()+RAND()");
    for addr in ["C1", "D1", "C2", "D2"] {
        assert_randbetween_bounds(
            &single.get_cell_value("Sheet1", addr),
            "RANDARRAY()",
            1.0,
            1_000_000.0,
        );
    }

    // Multi-threaded recalc must match single-threaded for the same initial state. This is the
    // key scheduling/ordering invariant for deterministic volatile RNG.
    let single_snapshot = snapshot(&single);
    let mut multi = Engine::new();
    setup_rng_sheet(&mut multi);
    multi.recalculate_multi_threaded();
    assert_eq!(snapshot(&multi), single_snapshot);
    assert_eq!(
        multi.spill_range("Sheet1", "C1"),
        Some((parse_a1("C1").unwrap(), parse_a1("D2").unwrap()))
    );

    // Recalculation without mutations must advance volatile RNG results. We allow a few attempts
    // to avoid pathological collisions.
    let first_a1 = single.get_cell_value("Sheet1", "A1");
    let first_a2 = single.get_cell_value("Sheet1", "A2");
    let first_a3 = single.get_cell_value("Sheet1", "A3");
    let first_array: Vec<Value> = ["C1", "D1", "C2", "D2"]
        .into_iter()
        .map(|addr| single.get_cell_value("Sheet1", addr))
        .collect();

    let mut changed_a1 = false;
    let mut changed_a2 = false;
    let mut changed_a3 = false;
    let mut changed_array = false;
    let mut observed_distinct_draws_in_cell = false;

    for _ in 0..10 {
        single.recalculate_single_threaded();

        // Invariants should continue to hold after every recalc.
        assert_eq!(single.get_cell_value("Sheet1", "A4"), Value::Bool(true));
        assert_eq!(single.get_cell_value("Sheet1", "B2"), Value::Number(0.0));
        match single.get_cell_value("Sheet1", "A5") {
            Value::Number(n) => {
                assert!(
                    n.is_finite(),
                    "expected RAND()-RAND() to be finite, got {n}"
                );
                assert!(
                    n > -1.0 && n < 1.0,
                    "expected RAND()-RAND() in (-1,1), got {n}"
                );
                if n != 0.0 {
                    observed_distinct_draws_in_cell = true;
                }
            }
            other => panic!("expected RAND()-RAND() to be a number, got {other:?}"),
        }
        assert_rand_unit_interval(&single.get_cell_value("Sheet1", "A1"), "RAND()");
        assert_randbetween_bounds(
            &single.get_cell_value("Sheet1", "A2"),
            "RANDBETWEEN()",
            1.0,
            1_000_000.0,
        );
        assert_rand_sum_bounds(&single.get_cell_value("Sheet1", "A3"), "RAND()+RAND()");
        for addr in ["C1", "D1", "C2", "D2"] {
            assert_randbetween_bounds(
                &single.get_cell_value("Sheet1", addr),
                "RANDARRAY()",
                1.0,
                1_000_000.0,
            );
        }

        if single.get_cell_value("Sheet1", "A1") != first_a1 {
            changed_a1 = true;
        }
        if single.get_cell_value("Sheet1", "A2") != first_a2 {
            changed_a2 = true;
        }
        if single.get_cell_value("Sheet1", "A3") != first_a3 {
            changed_a3 = true;
        }
        let next_array: Vec<Value> = ["C1", "D1", "C2", "D2"]
            .into_iter()
            .map(|addr| single.get_cell_value("Sheet1", addr))
            .collect();
        if next_array != first_array {
            changed_array = true;
        }

        if changed_a1 && changed_a2 && changed_a3 && changed_array {
            break;
        }
    }

    assert!(
        changed_a1,
        "expected RAND() to change across recalculations"
    );
    assert!(
        changed_a2,
        "expected RANDBETWEEN() to change across recalculations"
    );
    assert!(
        changed_a3,
        "expected RAND()+RAND() to change across recalculations"
    );
    assert!(
        changed_array,
        "expected RANDARRAY() spill values to change across recalculations"
    );
    assert!(
        observed_distinct_draws_in_cell,
        "expected to observe distinct RAND() draws within a single cell evaluation (RAND()-RAND() != 0) across a few recalcs"
    );
}
