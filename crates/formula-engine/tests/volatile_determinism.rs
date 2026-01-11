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
    }

    let mut single = Engine::new();
    setup(&mut single);
    single.recalculate_single_threaded();

    let mut multi = Engine::new();
    setup(&mut multi);
    multi.recalculate_multi_threaded();

    assert_eq!(multi.get_cell_value("Sheet1", "A1"), single.get_cell_value("Sheet1", "A1"));
    assert_eq!(multi.get_cell_value("Sheet1", "A2"), single.get_cell_value("Sheet1", "A2"));
    assert_eq!(multi.get_cell_value("Sheet1", "B1"), single.get_cell_value("Sheet1", "B1"));
}

