use formula_engine::{Engine, Value};

const N: u32 = 50;

fn set_formulas(engine: &mut Engine, reverse: bool) {
    let mut rows: Vec<u32> = (1..=N).collect();
    if reverse {
        rows.reverse();
    }

    // Many independent volatile cells. If RNG were global / order-dependent, different insertion
    // orders could yield different results even within the same recalc tick.
    for row in &rows {
        engine
            .set_cell_formula("Sheet1", &format!("A{row}"), "=RAND()")
            .unwrap();
        engine
            .set_cell_formula("Sheet1", &format!("B{row}"), "=RAND()+RAND()")
            .unwrap();
        engine
            .set_cell_formula("Sheet1", &format!("C{row}"), "=RANDBETWEEN(1, 1000000)")
            .unwrap();
    }
}

fn snapshot(engine: &Engine) -> Vec<Value> {
    let mut out = Vec::with_capacity((N as usize) * 3);
    for row in 1..=N {
        out.push(engine.get_cell_value("Sheet1", &format!("A{row}")));
        out.push(engine.get_cell_value("Sheet1", &format!("B{row}")));
        out.push(engine.get_cell_value("Sheet1", &format!("C{row}")));
    }
    out
}

fn assert_rng_bounds(snapshot: &[Value]) {
    for (idx, value) in snapshot.iter().enumerate() {
        match idx % 3 {
            0 => match value {
                Value::Number(n) => {
                    assert!(
                        *n >= 0.0 && *n < 1.0,
                        "expected RAND() to be in [0,1), got {n}"
                    );
                }
                other => panic!("expected RAND() to return a number, got {other:?}"),
            },
            1 => match value {
                Value::Number(n) => {
                    assert!(n.is_finite(), "expected RAND()+RAND() finite, got {n}");
                    assert!(
                        *n >= 0.0 && *n < 2.0,
                        "expected RAND()+RAND() to be in [0,2), got {n}"
                    );
                }
                other => panic!("expected RAND()+RAND() to return a number, got {other:?}"),
            },
            _ => match value {
                Value::Number(n) => {
                    assert!(n.is_finite(), "expected RANDBETWEEN() finite, got {n}");
                    assert!(
                        (n.fract()).abs() < 1e-9,
                        "expected RANDBETWEEN() integer, got {n}"
                    );
                    assert!(
                        *n >= 1.0 && *n <= 1_000_000.0,
                        "expected RANDBETWEEN() in [1,1000000], got {n}"
                    );
                }
                other => panic!("expected RANDBETWEEN() to return a number, got {other:?}"),
            },
        }
    }
}

#[test]
fn volatile_rng_is_deterministic_wrt_insertion_order() {
    let mut forward = Engine::new();
    set_formulas(&mut forward, false);
    forward.recalculate_single_threaded();
    let snap1 = snapshot(&forward);
    assert_rng_bounds(&snap1);

    let mut reverse = Engine::new();
    set_formulas(&mut reverse, true);
    reverse.recalculate_single_threaded();
    let snap1_reverse = snapshot(&reverse);
    assert_rng_bounds(&snap1_reverse);

    // Determinism invariant: the same workbook state should produce the same results regardless
    // of how formulas were inserted (and regardless of underlying evaluation order).
    assert_eq!(snap1_reverse, snap1);

    // Volatility invariant: a second recalc without mutations should advance RNG results, but
    // determinism should still hold between identical workbooks.
    forward.recalculate_single_threaded();
    reverse.recalculate_single_threaded();
    let snap2 = snapshot(&forward);
    let snap2_reverse = snapshot(&reverse);
    assert_rng_bounds(&snap2);
    assert_eq!(snap2_reverse, snap2);

    // Extremely unlikely to collide for all cells; this is a sanity check that recalc_id advances.
    assert_ne!(snap2, snap1);
}
