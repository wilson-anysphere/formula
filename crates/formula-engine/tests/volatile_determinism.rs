use formula_engine::{Engine, Value};

#[test]
fn now_is_frozen_within_single_recalc() {
    let mut engine = Engine::new();
    engine.set_cell_formula("Sheet1", "A1", "=NOW()").unwrap();
    engine.set_cell_formula("Sheet1", "B1", "=NOW()").unwrap();

    engine.recalculate();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        engine.get_cell_value("Sheet1", "B1")
    );
}

#[test]
fn rand_changes_across_recalcs_but_is_stable_within_one() {
    let mut engine = Engine::new();
    engine.set_cell_formula("Sheet1", "A1", "=RAND()").unwrap();
    assert_eq!(
        engine.bytecode_program_count(),
        1,
        "expected RAND() to be bytecode-eligible"
    );

    engine.recalculate();
    let first = engine.get_cell_value("Sheet1", "A1");

    let first_num = match first {
        Value::Number(n) => n,
        other => panic!("expected RAND() to return a number, got {other:?}"),
    };
    assert!(first_num >= 0.0);
    assert!(first_num < 1.0);

    // Volatile RNG should update on each recalc call. We allow a few attempts to avoid
    // pathological collisions.
    let mut changed = false;
    for _ in 0..5 {
        engine.recalculate();
        if engine.get_cell_value("Sheet1", "A1") != first {
            changed = true;
            break;
        }
    }
    assert!(changed, "expected RAND() to change across recalculations");
}

#[test]
fn multithreaded_and_singlethreaded_recalc_match_for_rng() {
    fn setup(engine: &mut Engine) {
        engine.set_cell_formula("Sheet1", "A1", "=RAND()").unwrap();
        engine
            .set_cell_formula("Sheet1", "A2", "=RANDBETWEEN(1, 1000000)")
            .unwrap();
        engine.set_cell_formula("Sheet1", "B1", "=A1+A2").unwrap();
        engine
            .set_cell_formula("Sheet1", "C1", "=RANDARRAY(2,2,1,10,TRUE)")
            .unwrap();
    }

    let mut single = Engine::new();
    setup(&mut single);
    assert_eq!(
        single.bytecode_program_count(),
        3,
        "expected RAND/RANDBETWEEN and scalar arithmetic to compile to bytecode"
    );
    single.recalculate_single_threaded();

    let mut multi = Engine::new();
    setup(&mut multi);
    assert_eq!(
        multi.bytecode_program_count(),
        3,
        "expected RAND/RANDBETWEEN and scalar arithmetic to compile to bytecode"
    );
    multi.recalculate_multi_threaded();

    assert_eq!(
        multi.get_cell_value("Sheet1", "A1"),
        single.get_cell_value("Sheet1", "A1")
    );
    assert_eq!(
        multi.get_cell_value("Sheet1", "A2"),
        single.get_cell_value("Sheet1", "A2")
    );
    assert_eq!(
        multi.get_cell_value("Sheet1", "B1"),
        single.get_cell_value("Sheet1", "B1")
    );
    assert_eq!(
        multi.get_cell_value("Sheet1", "C1"),
        single.get_cell_value("Sheet1", "C1")
    );
    assert_eq!(
        multi.get_cell_value("Sheet1", "D1"),
        single.get_cell_value("Sheet1", "D1")
    );
    assert_eq!(
        multi.get_cell_value("Sheet1", "C2"),
        single.get_cell_value("Sheet1", "C2")
    );
    assert_eq!(
        multi.get_cell_value("Sheet1", "D2"),
        single.get_cell_value("Sheet1", "D2")
    );
}

#[test]
fn randarray_changes_across_recalcs_and_respects_bounds() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=RANDARRAY(2,2,1,3,TRUE)")
        .unwrap();

    engine.recalculate();

    let first: Vec<Value> = ["A1", "B1", "A2", "B2"]
        .into_iter()
        .map(|addr| engine.get_cell_value("Sheet1", addr))
        .collect();

    // Validate basic spill footprint + invariants (integers within [1,3]).
    for addr in ["A1", "B1", "A2", "B2"] {
        let v = engine.get_cell_value("Sheet1", addr);
        let n = match v {
            Value::Number(n) => n,
            other => panic!("expected RANDARRAY cell {addr} to be a number, got {other:?}"),
        };
        assert!(
            n >= 1.0 && n <= 3.0,
            "expected {addr} within [1,3], got {n}"
        );
        assert!(
            (n.fract()).abs() < 1e-9,
            "expected {addr} to be an integer, got {n}"
        );
    }

    // RANDARRAY should change across recalcs; allow a few attempts to avoid collisions.
    let mut changed = false;
    for _ in 0..5 {
        engine.recalculate();
        let next: Vec<Value> = ["A1", "B1", "A2", "B2"]
            .into_iter()
            .map(|addr| engine.get_cell_value("Sheet1", addr))
            .collect();
        if next != first {
            changed = true;
            break;
        }
    }
    assert!(
        changed,
        "expected RANDARRAY() to change across recalculations"
    );
}

#[test]
fn map_with_rand_is_volatile_and_deterministic_within_one_recalc() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=MAP(SEQUENCE(3),LAMBDA(x,RAND()))")
        .unwrap();

    engine.recalculate();
    let first: Vec<Value> = ["A1", "A2", "A3"]
        .into_iter()
        .map(|addr| engine.get_cell_value("Sheet1", addr))
        .collect();

    // Validate all cells are numbers in [0,1).
    for (idx, value) in first.iter().enumerate() {
        let n = match value {
            Value::Number(n) => *n,
            other => panic!("expected MAP RAND result to be a number, got {other:?}"),
        };
        assert!(
            n >= 0.0 && n < 1.0,
            "expected value {idx} in [0,1), got {n}"
        );
    }

    // MAP should be treated as volatile because the supplied LAMBDA calls RAND().
    // Allow a few attempts to avoid pathological collisions.
    let mut changed = false;
    for _ in 0..5 {
        engine.recalculate();
        let next: Vec<Value> = ["A1", "A2", "A3"]
            .into_iter()
            .map(|addr| engine.get_cell_value("Sheet1", addr))
            .collect();
        if next != first {
            changed = true;
            break;
        }
    }
    assert!(
        changed,
        "expected MAP(...,LAMBDA(...,RAND())) to change across recalculations"
    );
}
